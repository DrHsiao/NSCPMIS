using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WCals.Controllers
{
    public class LogActionAttribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            var request = filterContext.HttpContext.Request;

            var controller = filterContext.RequestContext.RouteData.Values["controller"].ToString();
            var action = filterContext.RequestContext.RouteData.Values["Action"].ToString();
            var usrId = (request.IsAuthenticated) ? filterContext.HttpContext.User.Identity.Name : "Anonymous";
            Helpers.WebSiteHelper.logFunctionInfo(usrId, controller, action);

            base.OnActionExecuting(filterContext);
        }
    }
}