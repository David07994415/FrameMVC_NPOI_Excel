﻿using System.Web;
using System.Web.Mvc;

namespace FrameMVC_NPOI_Excel
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
