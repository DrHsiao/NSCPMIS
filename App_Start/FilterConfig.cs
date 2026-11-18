using System.Web.Mvc;

namespace WCals.App_Start
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            // 報表匯出後端例外 -> 導回 StateN 顯示「查無資料」，取代黃頁（須先於 HandleError）
            filters.Add(new WCals.Filters.ReportPrintExceptionFilter());

            filters.Add(new HandleErrorAttribute());

            // 整個網站預設都要登入（全域Authorize）
            filters.Add(new AuthorizeAttribute());
        }
    }
}