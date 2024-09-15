using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace FrameMVC_NPOI_Excel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        // FileUpload Reference：https://kevintsengtw.blogspot.com/2013/03/aspnet-mvc.html
        /// <summary>
        ///  進行主表和交易表的檔案格式確認 API 
        /// </summary>
        /// <param name="checkMainFile">主表檔案資料</param>
        /// <param name="checkSubFile">交易表檔案資料</param>
        /// <returns>ResultReplyViewModel 檢查結果，轉成 JSON 物件</returns>
        [HttpPost]
        public ActionResult CheckExcelFormat(HttpPostedFileBase checkMainFile, HttpPostedFileBase checkSubFile)  
        {
            ResultReplyViewModel Result_Main = ExcelProcess.CheckInputFilesValidation(checkMainFile,true,false);
            if (Result_Main.Status != "Success")
            {
                return Json(Result_Main, JsonRequestBehavior.AllowGet);
            }

            ResultReplyViewModel Result_Sub = ExcelProcess.CheckInputFilesValidation(checkSubFile, false, true);
            if (Result_Sub.Status != "Success")
            {
                return Json(Result_Sub, JsonRequestBehavior.AllowGet);
            }

            var Result_CheckDone= new ResultReplyViewModel();
            Result_CheckDone.Status = "Success";
            Result_CheckDone.Msg = "檔案檢查成功，請點選上傳檔案按鈕";
            return Json(Result_CheckDone, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        ///  進行主表和交易表的檔案格式確認 API 
        /// </summary>
        /// <param name="uploadMainFile">主表檔案資料</param>
        /// <param name="uploadSubFile">交易表檔案資料</param>
        /// <returns>ResultReplyViewModel 上傳結果，轉成 JSON 物件</returns>
        [HttpPost]
        public ActionResult UploadExcel(HttpPostedFileBase uploadMainFile, HttpPostedFileBase uploadSubFile)
        {
            ResultReplyViewModel result_Error = new ResultReplyViewModel();
            result_Error.Status = "Error";

            // 1. 驗證資料 ( 再次進行驗證)
            ResultReplyViewModel Result_Main = ExcelProcess.CheckInputFilesValidation(uploadMainFile, true, false);
            if (Result_Main.Status != "Success")
            {
                return Json(Result_Main, JsonRequestBehavior.AllowGet);
            }
            ResultReplyViewModel Result_Sub = ExcelProcess.CheckInputFilesValidation(uploadSubFile, false, true);
            if (Result_Sub.Status != "Success")
            {
                return Json(Result_Sub, JsonRequestBehavior.AllowGet);
            }

            // 2. 擷取資料
            UploadAndInsertResultDto UploadResultDto = ExcelProcess.GetFilesUploadAndInsert(uploadMainFile, uploadSubFile);
            if (UploadResultDto.ErrorMessage != "")
            {
                result_Error.Msg = $"擷取資料失敗:{UploadResultDto.ErrorMessage}";
                return Json(result_Error, JsonRequestBehavior.AllowGet);
            }

            // 3. 儲存資料 ( 儲存進去 Land 和 Build SQL Table)

            // 4. 計算資料 ( 針對 Land 和 Build SQL Table 的資料進行彙整計算，儲存進去 Calculation Table)



            var Result_UploadDone = new ResultReplyViewModel();
            Result_UploadDone.Status = "Success";
            Result_UploadDone.Msg = "檔案上傳成功";
            return Json(Result_UploadDone, JsonRequestBehavior.AllowGet);
        }






        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

    }
}