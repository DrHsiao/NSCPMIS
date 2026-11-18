using System;
using System.Security.Principal;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Routing;
using System.Web.Security;
using WCals.App_Start;

namespace WCals
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters); // 整個網站預設都要登入（全域Authorize）
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            Aspose.Words.License licenseWords = new Aspose.Words.License();
            licenseWords.SetLicense("Aspose.Total.lic");

            Aspose.Cells.License licenseExcels = new Aspose.Cells.License();
            licenseExcels.SetLicense("Aspose.Total.lic");

        }

        protected void Application_AuthenticateRequest(Object sender, EventArgs e)
        {

            bool hasUser = HttpContext.Current.User != null;
            bool isAuthenticated = hasUser && HttpContext.Current.User.Identity.IsAuthenticated;
            bool isIdentity = isAuthenticated && HttpContext.Current.User.Identity is FormsIdentity;
            if (isIdentity)
            {
                //取得表單認證目前這位使用者身份
                FormsIdentity id = (FormsIdentity)HttpContext.Current.User.Identity;

                //取得 FormAuthenticationTicket 物件
                FormsAuthenticationTicket ticket = id.Ticket;

                //取得 UserData 欄位資料 (這裡儲存的是角色), 如果有多個角色可以用逗號分隔 先預留暫無處理
                string[] roles = ticket.UserData.Split(',');

                //賦予該使用者新的身分 (含角色資訊)
                HttpContext.Current.User = new GenericPrincipal(id, roles);
            }

        }

        protected void Application_PreSendRequestHeaders(Object sender, EventArgs e)
        {
            if (HttpContext.Current != null)
            {
                HttpResponse curResponse = HttpContext.Current.Response;
                curResponse.Headers.Remove("X-Powered-By");
                curResponse.Headers.Remove("X-AspNet-Version");
                curResponse.Headers.Remove("X-AspNetMvc-Version");
                curResponse.Headers.Remove("Server");
                //curResponse.Headers.Set("Server", "Apache/2.4.7 (Unix) OpenSSL/1.0.1e");
            }
        }
    }


}
