using NPOI.HSSF.UserModel;  // 記得要 Using
using NPOI.POIFS.Crypt;
using NPOI.SS.UserModel;      // 記得要 Using
using NPOI.XSSF.UserModel;  // 記得要 Using
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

//1.先安裝 NPOI 套件
//2. using NPOI.SS.UserModel; using NPOI.XSSF.UserModel; using NPOI.HSSF.UserModel;

namespace FrameMVC_NPOI_Excel
{
    public class ExcelProcess
    {
        /// <summary>
        /// 預設的主表分頁名稱與對應欄位
        /// </summary>
        private static Dictionary<string, List<string>> Main_SheetNameColumnDict
            = new Dictionary<string, List<string>>()
            {
                 {"買賣", new List<string>(){ "Key", "SubKey", "Price", "行政區", "時間" } },
                 {"預售屋", new List<string>(){ "Key", "SubKey", "Price", "行政區", "時間" } },
                  {"租賃", new List<string>(){ "Key", "SubKey", "Price", "行政區", "時間" } }
            };

        /// <summary>
        /// 預設的交易表分頁名稱與對應欄位
        /// </summary>
        private static Dictionary<string, List<string>> Sub_SheetNameColumnDict
           = new Dictionary<string, List<string>>()
           {
                 {"買賣", new List<string>(){ "Key", "SubKey", "Price", "行政區", "時間" } },
                 {"預售屋", new List<string>(){ "Key", "SubKey", "Price", "行政區", "時間" } },
                  {"租賃", new List<string>(){ "Key", "SubKey", "Price", "行政區", "時間" } }
           };


        /// <summary>
        ///  取得預設檢查項目的清單內容
        /// </summary>
        /// <param name="SheetNameColumnDict">預設表單之檢查格式字典</param>
        /// <returns> CheckSheetHeaderDto 清單 List</returns>
        public static List<CheckSheetHeaderDto> LoadDefaultCheckDtos(Dictionary<string, List<string>> SheetNameColumnDict)
        {
            List<CheckSheetHeaderDto> DefaultCheckDtos = new List<CheckSheetHeaderDto>();
            foreach (var DictItem in SheetNameColumnDict)
            {
                CheckSheetHeaderDto tempDto = new CheckSheetHeaderDto();
                tempDto.SheetName = DictItem.Key;
                tempDto.TitleColumns = DictItem.Value;
                DefaultCheckDtos.Add(tempDto);
            }
            return DefaultCheckDtos;
        }

        /// <summary>
        ///  檢查檔案之 EXCLE 新舊版本
        /// </summary>
        /// <param name="File">檔案資料</param>
        /// <returns>IsNewOrOldVersion DTO 物件</returns>
        public static IsNewOrOldVersion CheckExcelVersion(HttpPostedFileBase File)  // Todo: 可能還有其他的 Mimi Type： To Be Check
        {
            IsNewOrOldVersion VersionDto = new IsNewOrOldVersion();
            if (File.ContentType == "application/vnd.ms-excel")  // 舊版本 xls  
            {
                VersionDto.IsOldVersion = true;
            }
            else if (File.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")   // 新版本 xlsx
            {
                VersionDto.IsNewVersion = true;
            }
            return VersionDto;
        }

        /// <summary>
        ///  針對新舊 EXCEL 版本檔案，進行 NPOI 實體建立
        /// </summary>
        /// <param name="VersionDto">檔案版本 DTO 物件</param>
        /// <param name="ms">檔案資料</param>
        /// <returns>NPOI 物件</returns>
        public static IWorkbook GetNpoiObj(IsNewOrOldVersion VersionDto, MemoryStream ms)
        {
            IWorkbook workbook = null;
            if (VersionDto.IsOldVersion)
            {
                workbook = new HSSFWorkbook(ms); // 取得 舊版 的實體
            }
            else if (VersionDto.IsNewVersion)
            {
                workbook = new XSSFWorkbook(ms);  // 取得 新版 的實體
            }
            return workbook;
        }


        /// <summary>
        ///  進行檔案格式的檢查
        /// </summary>
        /// <param name="File">檔案資料</param>
        /// <param name="IsMainFile">此檔案是否為主檔</param>
        /// <param name="IsSubFile">此檔案是否為副檔</param>
        /// <returns>ResultReplyViewModel 類型結果物件</returns>
        public static ResultReplyViewModel CheckInputFilesValidation(HttpPostedFileBase File, bool IsMainFile, bool IsSubFile)
        {
            ResultReplyViewModel result_Error = new ResultReplyViewModel();
            result_Error.Status = "Error";

            string FileTypeName = IsMainFile ? "主檔" : IsSubFile ? "交易檔" : "";

            if (File == null)
            {
                result_Error.Msg = $"{FileTypeName}為空，請先選擇{FileTypeName}";
                return result_Error;
            }
            if (File.ContentLength <= 0)
            {
                result_Error.Msg = $"{FileTypeName}為空，請先選擇{FileTypeName}";
                return result_Error;
            }
            if (!File.FileName.EndsWith(".xls") && !File.FileName.EndsWith(".xlsx"))
            {
                result_Error.Msg = $"{FileTypeName}非 excel 相關檔案";
                return result_Error;
            }
            var VersionDto = CheckExcelVersion(File);
            if (VersionDto.IsNewVersion == false && VersionDto.IsOldVersion == false)
            {
                result_Error.Msg = $"{FileTypeName}非 excel 相關檔案";
                return result_Error;
            }

            using (var ms = new MemoryStream())  // 將檔案 轉入 MemoryStream
            {
                File.InputStream.CopyTo(ms);      // 將檔案 轉入 MemoryStream
                ms.Position = 0;                           // 重置流的位置到開頭 (沒有這行會出錯)
                byte[] fileBytes = ms.ToArray();   // 將檔案 轉成 Byte[]
                // 進行 Magic Number 確認
                if (VersionDto.IsOldVersion)
                {
                    // 讀取文件頭部的前 8 個字節
                    if (!fileBytes.Take(8).SequenceEqual(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }))
                    {
                        result_Error.Msg = $"{FileTypeName}格式非 excel 相關格式";
                        return result_Error;
                    }
                }
                else if (VersionDto.IsNewVersion)
                {
                    if (!fileBytes.Take(2).SequenceEqual(new byte[] { 0x50, 0x4B }))  // 讀取文件頭部的前 2 個字節
                    {
                        result_Error.Msg = $"{FileTypeName}格式非 excel 相關格式";
                        return result_Error;
                    }
                }

                // 進行 NPOI 實體建立
                IWorkbook workbook = GetNpoiObj(VersionDto, ms);

                // 進行 分頁和欄位確認
                try
                {
                    // 給定檢查的預設檢查項目與內容
                    var SheetNameColumnDict = IsMainFile ? Main_SheetNameColumnDict : IsSubFile ? Sub_SheetNameColumnDict : null;
                    var DetaultCheckDtos = LoadDefaultCheckDtos(SheetNameColumnDict);

                    int workBookSheetNumber = workbook.NumberOfSheets;  // 上傳檔案的分頁數
                    if (workBookSheetNumber != DetaultCheckDtos.Count)
                    {
                        result_Error.Msg = $"{FileTypeName}格式之分頁數量與預設格式之分頁數量不一致";
                        return result_Error;
                    }

                    for (int i = 0; i < workbook.NumberOfSheets; i++) // Loop 上傳檔案 的每張分頁
                    {
                        var sheet = workbook.GetSheetAt(i);         // 取得 分頁 物件
                        string sheetName = sheet.SheetName;     // 取得 分頁 名稱
                        for (int j = 0; j < DetaultCheckDtos.Count; j++)   // Loop 預設檔案 的每張分頁
                        {
                            if (DetaultCheckDtos[j].SheetName == sheetName && DetaultCheckDtos[j].HasCompared == false)
                            {   // 分頁名稱相同，且物件尚未被檢查過
                                var row = sheet.GetRow(0); // 取得 上傳檔案 第一行: Header 標題
                                foreach (var TitleColumn in DetaultCheckDtos[j].TitleColumns)  // Loop 預設檔案 該分頁 的 標題欄位
                                {
                                    bool containsTitle = false;
                                    foreach (var cell in row.Cells)    // Loop上傳檔案 該分頁 的 標題欄位
                                    {
                                        if (cell.StringCellValue.Trim() == TitleColumn)
                                        {
                                            containsTitle = true;
                                            break;
                                        }
                                    }
                                    if (!containsTitle)
                                    {
                                        result_Error.Msg = $"{FileTypeName}{sheetName}分頁的缺少{TitleColumn}欄位";
                                        return result_Error;
                                    }

                                }
                                DetaultCheckDtos[j].HasCompared = true;
                            }
                        }
                    }

                    foreach (var checkDto in DetaultCheckDtos)  // Loop 預設檔案 的每個分頁是否都完成檢查
                    {
                        if (checkDto.HasCompared == false)
                        {
                            result_Error.Msg = $"{FileTypeName}缺乏{checkDto.SheetName}分頁";
                            return result_Error;
                        }
                    }
                }
                catch (Exception ex)
                {
                    result_Error.Msg = $"Error_Issue:{ex.Message}";
                    return result_Error;
                }
            }

            // 通過檔案檢查
            ResultReplyViewModel result_Success = new ResultReplyViewModel();
            result_Success.Status = "Success";
            return result_Success;
        }


        /// <summary>
        ///  進行 EXCEL 資料的擷取(包含整合主表和交易表資料)
        /// </summary>
        /// <param name="uploadMainFile">主檔資料格式</param>
        /// <param name="uploadSubFile">交易檔資料格式</param>
        /// <returns>UploadAndInsertResultDto 物件，包含要加入 SLQ的資料清單</returns>

        public static UploadAndInsertResultDto GetFilesUploadAndInsert(HttpPostedFileBase uploadMainFile, HttpPostedFileBase uploadSubFile)
        {
            UploadAndInsertResultDto resultDto = new UploadAndInsertResultDto();
            try
            {
                var DefaultCheckList_Main = LoadDefaultCheckDtos(Main_SheetNameColumnDict);
                var DefaultCheckList_Sub = LoadDefaultCheckDtos(Sub_SheetNameColumnDict);

                using (MemoryStream ms_Main = new MemoryStream())
                {
                    uploadMainFile.InputStream.CopyTo(ms_Main);      // 將檔案 轉入 MemoryStream
                    ms_Main.Position = 0;                                               // 重置流的位置到開頭 (沒有這行會出錯)
                    var MainVersionDto = CheckExcelVersion(uploadMainFile);
                    IWorkbook workbook_Main = GetNpoiObj(MainVersionDto, ms_Main);

                    using (MemoryStream ms_Sub = new MemoryStream())
                    {
                        uploadSubFile.InputStream.CopyTo(ms_Sub);      // 將檔案 轉入 MemoryStream
                        ms_Sub.Position = 0;                                             // 重置流的位置到開頭 (沒有這行會出錯)
                        var SubVersionDto = CheckExcelVersion(uploadSubFile);
                        IWorkbook workbook_Sub = GetNpoiObj(SubVersionDto, ms_Sub);

                        // 建立 預計加入資料庫的清單物件
                        List<SqlInsertLandBuildDto> LandTypeList = new List<SqlInsertLandBuildDto>();
                        List<SqlInsertLandBuildDto> BuildTypeList = new List<SqlInsertLandBuildDto>();

                        foreach (var CheckSheet in DefaultCheckList_Main) // Loop 預設 主表 分頁
                        {

                            // 取得 上傳檔案的主表和交易表 特定分頁 相關資料
                            var Sheet_Main = workbook_Main.GetSheet(CheckSheet.SheetName);
                            var Sheet_Sub = workbook_Sub.GetSheet(CheckSheet.SheetName);
                            var Sheet_Main_RowCount = Sheet_Main.LastRowNum;
                            var Sheet_Sub_RowCount = Sheet_Sub.LastRowNum;
                            // 取得 上傳檔案的主表和交易表 要搜尋的欄位名稱和 Index 對應字典物件
                            var Main_ColumnIndexDict = LoadColumnIndex(Sheet_Main.GetRow(0), CheckSheet.TitleColumns);
                            var Sub_ColumnIndexDict = LoadColumnIndex(Sheet_Main.GetRow(0), CheckSheet.TitleColumns);

                            for (int i = 1; i < Sheet_Main_RowCount; i++) // i =0 是 title Row，不需要跑迴圈  // Loop 上傳檔案 主要表單的每一Row
                            {
                                var row_Main = Sheet_Main.GetRow(i);

                                // 取得主表更新後的資料
                                var sqlInsertLandBuildDto = LoadMainTableModel(Main_ColumnIndexDict, row_Main);
                                // 取得交易表更新後的資料
                                sqlInsertLandBuildDto = LoadSubTableModel(sqlInsertLandBuildDto, Sub_ColumnIndexDict, Sheet_Sub, Sheet_Sub_RowCount);
                                // 進行 SQL Table 分類，分類加入 Land 或者 build 的 List
                                if (sqlInsertLandBuildDto.IsLandType)
                                {
                                    LandTypeList.Add(sqlInsertLandBuildDto);
                                }
                                if (sqlInsertLandBuildDto.IsBuildType)
                                {
                                    BuildTypeList.Add(sqlInsertLandBuildDto);
                                }
                            }
                        }
                        // 擷取資料成功
                        resultDto.LandTypeList = LandTypeList;
                        resultDto.BuildTypeList = BuildTypeList;
                        return resultDto;
                    }
                }

            }
            catch (Exception ex)
            {
                resultDto.ErrorMessage = ex.Message;
                return resultDto;
            }
        }

        /// <summary>
        ///  取得所需要的標題欄位名稱和其對應的 Index 字典
        /// </summary>
        /// <param name="row">NPOI Row 物件，為第一個 Row: Header Row</param>
        /// <param name="titleColumns">預設格式的欄位清單 List 資料</param>
        /// <returns>具有 欄位名稱 和 欄位 Index 的 字典物件</returns>
        private static Dictionary<string, int> LoadColumnIndex(IRow rowHeader, List<string> titleColumns)
        {
            Dictionary<string, int> ColumnDicList = new Dictionary<string, int>();

            foreach (var columnName in titleColumns)  // Loop 預設格式 的欄位清單
            {
                for (int i = 0; i < rowHeader.Cells.Count; i++) // Loop 上傳檔案的 的第一Row 的每個 Cell
                {
                    if (rowHeader.GetCell(i).StringCellValue.Trim() == columnName)
                    {
                        ColumnDicList.Add(columnName, i);
                        break;
                    }
                }
            }
            return ColumnDicList;
        }


        /// <summary>
        ///  進行主表資料Mapping，取得主要表單要匯入 SQL 的資料
        /// </summary>
        /// <param name="Main_ColumnIndexDict">預設表單(主表)欄位字典清單</param>
        /// <param name="row_Main">主要表單的特定一Rows</param>
        /// <returns>SqlInsertLandBuildDto 物件</returns>
        private static SqlInsertLandBuildDto LoadMainTableModel(Dictionary<string, int> Main_ColumnIndexDict, IRow row_Main)
        {
            SqlInsertLandBuildDto SqlInsertDto = new SqlInsertLandBuildDto();

            foreach (var columnIndex in Main_ColumnIndexDict) // Loop 每一個字典清單(代表 each 欄位名稱和 Index)
            {
                /// 可能需要進行修改相關 case 
                switch (columnIndex.Key)
                {
                    case "Key":
                        SqlInsertDto.Key = row_Main.GetCell(columnIndex.Value).StringCellValue.Trim();
                        break;
                    case "SubKey":
                        SqlInsertDto.SubKey = row_Main.GetCell(columnIndex.Value).StringCellValue.Trim();
                        break;
                    case "LandType":
                        SqlInsertDto.LandType = row_Main.GetCell(columnIndex.Value).StringCellValue.Trim();
                        break;
                    case "BuildType":
                        SqlInsertDto.BuildType = row_Main.GetCell(columnIndex.Value).StringCellValue.Trim();
                        break;
                    case "Price":
                        SqlInsertDto.Price = row_Main.GetCell(columnIndex.Value).NumericCellValue;
                        break;
                    case "行政區":
                        SqlInsertDto.Area = row_Main.GetCell(columnIndex.Value).StringCellValue.Trim();
                        break;
                    case "時間":
                        SqlInsertDto.DateTime = row_Main.GetCell(columnIndex.Value).DateCellValue.GetValueOrDefault();  // Todo
                        break;
                    default:
                        break;
                }
            }

            if (SqlInsertDto.LandType == "土地")
            {
                SqlInsertDto.IsLandType = true;
            }
            else if (!string.IsNullOrEmpty(SqlInsertDto.LandType) && SqlInsertDto.LandType != "車位")
            {
                SqlInsertDto.IsBuildType = true;
            }

            return SqlInsertDto;
        }



        /// <summary>
        ///  進行交易表資料Mapping，並針對 SqlInsertDto 物件，進行交易表單的資料更新
        /// </summary>
        /// <param name="Dto">SqlInsertLandBuildDto 物件(更新前)</param>
        /// <param name="sub_ColumnIndexDict">預設表單(交易表)欄位字典清單</param>
        /// <param name="sheet_Sub">上傳檔案(交易表) 的分頁</param>
        /// <param name="Sheet_Sub_RowCount">上傳檔案(交易表) 的分頁 Rows 數量</param>
        /// <returns>SqlInsertLandBuildDto 物件(更新後)</returns>
        private static SqlInsertLandBuildDto LoadSubTableModel(SqlInsertLandBuildDto Dto, Dictionary<string, int> sub_ColumnIndexDict, ISheet sheet_Sub, int Sheet_Sub_RowCount)
        {
            List<IRow> Rows_Match_MainTable_KeyID_List = new List<IRow>();
            for (int i = 1; i < Sheet_Sub_RowCount; i++)  // i =0 是 title Row，不需要跑迴圈  // Loop 上傳檔案 交易表單的每一Row
            {
                // 取得該 Rows 的 Key-Pairs                 
                /// 可能需要進行修改 "Key"   "SubKey"
                string sheet_Sub_Row_KeyCell_String = sheet_Sub.GetRow(i).GetCell(sub_ColumnIndexDict["Key"]).StringCellValue.Trim();
                string sheet_Sub_Row_SubKeyCell_String = sheet_Sub.GetRow(i).GetCell(sub_ColumnIndexDict["SubKey"]).StringCellValue.Trim();

                // 如果 Key-Pairs 等於 Dto 對應的欄位值
                if (sheet_Sub_Row_KeyCell_String == Dto.Key && sheet_Sub_Row_SubKeyCell_String == Dto.SubKey)
                {
                    Rows_Match_MainTable_KeyID_List.Add(sheet_Sub.GetRow(i));
                }
            }

            foreach (var Match_Row in Rows_Match_MainTable_KeyID_List)
            {
                // todo:這邊就要處理交易表單的欄位邏輯問題，並且update MainModel
                // 可以事情況擴充 Dto Props
            }

            return Dto;
        }

    }


    public class ResultReplyViewModel
    {
        public string Status { get; set; } // Success or Error

        public string Msg { get; set; }
    }

    public class CheckSheetHeaderDto
    {
        public string SheetName { get; set; }

        public List<string> TitleColumns { get; set; } = new List<string>();

        public bool HasCompared { get; set; } = false;
    }

    public class IsNewOrOldVersion
    {
        public bool IsNewVersion { get; set; } = false;

        public bool IsOldVersion { get; set; } = false;
    }


    public class MainTableViewModel
    {
        public string Key { get; set; }
        public string SubKey { get; set; }
        public string Type { get; set; }
        public double Price { get; set; }
        public string Area { get; set; }
        public DateTime DateTime { get; set; }
        public bool IsLandType { get; set; } = false;
        public bool IsBuildType { get; set; } = false;
    }

    public class SubTableViewModel
    {
        public string Key { get; set; }
        public string SubKey { get; set; }
        public double Price { get; set; }

        public string Area { get; set; }
        public DateTime DateTime { get; set; }
    }


    public class SqlInsertLandBuildDto // 取兩者最大交集
    {
        public string Key { get; set; } = "";
        public string SubKey { get; set; } = "";
        public string LandType { get; set; } = "";
        public string BuildType { get; set; } = "";
        public double Price { get; set; } = 0.0;
        public string Area { get; set; } = "";
        public DateTime DateTime { get; set; }
        public bool IsLandType { get; set; } = false;
        public bool IsBuildType { get; set; } = false;
    }

    public class UploadAndInsertResultDto
    {
        public List<SqlInsertLandBuildDto> LandTypeList { get; set; } = new List<SqlInsertLandBuildDto>();
        public List<SqlInsertLandBuildDto> BuildTypeList { get; set; } = new List<SqlInsertLandBuildDto>();
        public string ErrorMessage { get; set; } = "";
    }
}