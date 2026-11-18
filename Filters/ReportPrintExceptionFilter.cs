using System;
using System.Web.Mvc;

namespace WCals.Filters
{
    /// <summary>
    /// 報表匯出（Report / RainReport 的 ReportPrintN 系列）若因資料不足而在後端拋出例外
    /// （例如所選計畫在該年月無對應資料，導致彙總物件為 null 或 DataTable 無列），
    /// 於此攔截並導回原本的報表設定頁，以 TempData 帶出「查無資料」訊息，
    /// 由頁面顯示紅字說明，取代預設的例外（黃）頁。不需修改各報表 Controller。
    /// </summary>
    public class ReportPrintExceptionFilter : IExceptionFilter
    {
        public void OnException(ExceptionContext filterContext)
        {
            if (filterContext == null || filterContext.ExceptionHandled)
            {
                return;
            }

            var controller = filterContext.RouteData.Values["controller"] as string;
            var action = filterContext.RouteData.Values["action"] as string;

            // 僅處理報表匯出類動作，且限報表兩支 Controller
            bool isReportController =
                string.Equals(controller, "Report", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(controller, "RainReport", StringComparison.OrdinalIgnoreCase);

            // 報表匯出動作：名稱含 ReportPrint，或 RainReport/Index 以 POST 匯出報表
            bool isPost = string.Equals(
                filterContext.HttpContext.Request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase);
            bool isPrintAction =
                !string.IsNullOrEmpty(action) &&
                (action.IndexOf("ReportPrint", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 (isPost && string.Equals(action, "Index", StringComparison.OrdinalIgnoreCase)));

            if (!isReportController || !isPrintAction)
            {
                return;
            }

            var url = ResolveReturnUrl(filterContext, controller, action);
            if (string.IsNullOrEmpty(url))
            {
                return; // 無法判斷返回頁時不攔截，維持原行為
            }

            // 記錄實際例外，否則報表錯誤會被本 filter 完全吞掉而無從診斷
            try
            {
                var logDir = filterContext.HttpContext.Server.MapPath("~/App_Data");
                if (!System.IO.Directory.Exists(logDir)) System.IO.Directory.CreateDirectory(logDir);
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(logDir, "ReportPrintError.log"),
                    string.Format("[{0:yyyy-MM-dd HH:mm:ss}] {1}/{2}{3}{4}{3}{5}{3}",
                        DateTime.Now, controller, action, Environment.NewLine,
                        filterContext.Exception, new string('-', 80)));
            }
            catch { /* 記錄失敗不可影響使用者流程 */ }

            // 以 Session 傳遞訊息（例外階段設定 TempData 的儲存時機不可靠），由目標頁讀取後清除
            if (filterContext.HttpContext.Session != null)
            {
                filterContext.HttpContext.Session["ReportMessage"] = "查無資料";
            }
            filterContext.Result = new RedirectResult(url);
            filterContext.ExceptionHandled = true;
        }

        /// <summary>
        /// 決定導回的設定頁：優先回到表單來源頁（Referrer），
        /// 若取不到再以 Report 的「ReportPrintN → StateN」規則推導。
        /// </summary>
        private string ResolveReturnUrl(ExceptionContext filterContext, string controller, string action)
        {
            var referrer = filterContext.HttpContext.Request.UrlReferrer;
            if (referrer != null && !string.IsNullOrEmpty(referrer.PathAndQuery))
            {
                return referrer.PathAndQuery;
            }

            // Fallback：僅 Report 的規則性命名（ReportPrint1 -> State1）
            if (string.Equals(controller, "Report", StringComparison.OrdinalIgnoreCase))
            {
                var suffix = action.Substring("ReportPrint".Length);
                var target = "State" + suffix;
                return new UrlHelper(filterContext.RequestContext).Action(target, "Report");
            }

            return null;
        }
    }
}
