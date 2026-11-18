using System.Web.Mvc;

namespace WCals.Controllers
{
    public class ErrorController : Controller
    {
        /// <summary>
        /// 錯誤頁面顯示(根據不同的Http狀態碼顯示不同的錯誤訊息)
        /// </summary>
        /// <param name="errorCode"> Http狀態碼 </param>
        /// <param name="message"> 自訂錯誤訊息 </param>
        /// <returns></returns>
        public ActionResult ErrorPage(int errorCode = 500, string message = null)
        {
            Response.StatusCode = errorCode;
            string title;
            switch (errorCode)
            {
                case 400:
                    title = "400 - 請求錯誤";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "參數錯誤，請確認輸入是否正確。";
                    }
                    break;
                case 401:
                    title = "401 - 未授權";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "請先登入才能存取此頁面。";
                    }
                    break;
                case 403:
                    title = "403 - 禁止存取";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "您沒有權限存取該資源。";
                    }
                    break;
                case 404:
                    title = "404 - 找不到頁面";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "您所查詢的頁面不存在。";
                    }
                    break;
                case 405:
                    title = "405 - 方法不被允許";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "請求的方法不被允許。";
                    }
                    break;
                case 422:
                    title = "422 - 請求語意錯誤";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "請求的內容格式正確但資料有誤。";
                    }
                    break;
                case 429:
                    title = "429 - 請求過多";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "您請求過於頻繁，請稍後再試。";
                    }
                    break;
                case 500:
                    title = "500 - 系統錯誤";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "系統發生錯誤，請稍後再試。";
                    }
                    break;
                case 503:
                    title = "503 - 服務不可用";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "服務暫時無法使用，請稍後再試。";
                    }
                    break;
                case 999: // 自訂錯誤碼
                    title = "發生錯誤";
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "處理您的要求時發生錯誤，請稍後再試。";
                    }
                    break;
                default:
                    title = errorCode + "- 錯誤";
                    break;
            }
            ViewBag.ErrorTitle = title;
            ViewBag.Message = message;
            return View("Error");
        }
    }
}