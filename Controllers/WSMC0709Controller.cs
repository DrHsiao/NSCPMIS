using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;
using WCals.ViewModels;

namespace WCals.Controllers
{
    public class WSMC0709Controller : Controller
    {
        //private readonly Models.WCalsEntities db = new Models.WCalsEntities();
        //private int pageSize = 15;

        //// GET: WSMC0709
        //private ViewModels.ManageViewModel bindData(int pageNumber)
        //{
        //    ViewModels.ManageViewModel vmodel = new ViewModels.ManageViewModel();
        //    vmodel.SMCLOGSList = db.SMCLOGS.Where(f => f.ACTION_DESC != "login").OrderByDescending(f => f.ADD_DT).ToPagedList(pageNumber, pageSize);
        //    return vmodel;
        //}
        //public ActionResult Index(int? page)
        //{
        //    var pageNumber = page ?? 1;

        //    return View(bindData(pageNumber));
        //}
        //[HttpPost]
        //[LogAction]
        //public ActionResult Index(ViewModels.ManageViewModel model)
        //{
        //    int pageIndex = model.PageNumber < 1 ? 1 : model.PageNumber;

        //    return View(bindData(pageIndex));
        //}

        private readonly WCals.Models.CMBASEModels.CMBASEEntities db = new WCals.Models.CMBASEModels.CMBASEEntities();
        private int pageSize = 15;

        private IQueryable<WCals.Models.CMBASEModels.SMCLOGS> getQueryData(LogSearchViewModel SearchParameter)
        {
            var query = db.SMCLOGS.AsQueryable();

            if (null != SearchParameter)
            {
                if (!string.IsNullOrWhiteSpace(SearchParameter.ACTION_TYPE))
                {
                    SearchParameter.ACTION_TYPE = SearchParameter.ACTION_TYPE.Trim();
                    query = query.Where(
                     f => f.ACTION_TYPE.Equals(SearchParameter.ACTION_TYPE));
                }

                if (!string.IsNullOrWhiteSpace(SearchParameter.USR_ID))
                {
                    SearchParameter.USR_ID = SearchParameter.USR_ID.Trim();
                    query = query.Where(
                     f => f.USR_ID.Contains(SearchParameter.USR_ID));
                }

                if (!string.IsNullOrWhiteSpace(SearchParameter.LOGIN_IP))
                {
                    SearchParameter.LOGIN_IP = SearchParameter.LOGIN_IP.Trim();
                    query = query.Where(
                        f => f.LOGIN_IP.Contains(SearchParameter.LOGIN_IP));
                }

                if (SearchParameter.LOGIN_DT_START.HasValue)
                {
                    query = query.Where(
                        f => DbFunctions.TruncateTime(f.LOGIN_DT) >= DbFunctions.TruncateTime(SearchParameter.LOGIN_DT_START));
                }

                if (SearchParameter.LOGIN_DT_END.HasValue)
                {
                    query = query.Where(
                        f => DbFunctions.TruncateTime(f.LOGIN_DT) <= DbFunctions.TruncateTime(SearchParameter.LOGIN_DT_END));
                }
            }

            return query = query.OrderByDescending(f => f.LOGIN_DT);
        }

        // GET: WSMC0709
        [LogAction]
        public ActionResult Index(int? page)
        {
            TempData.Remove(Constants.SEARCH_PARAMETER_KEY);
            var pageNumber = page ?? 1;
            TempData[Constants.PAGE_NUMBER_KEY] = pageNumber;

            var searchParameters = (LogSearchViewModel)TempData.Peek(Constants.SEARCH_PARAMETER_KEY);
            var query = getQueryData(searchParameters);

            ViewBag.ACTION_TYPE = Helpers.TenderHelper.getItemActionTypeOptions();
            var result = new ManageViewModel()
            {
                SearchParameter = searchParameters ?? new LogSearchViewModel(),
                SMCLOGSList = query.ToPagedList(pageNumber, pageSize)
            };

            return View(result);
        }

        public ActionResult PagedIndex(int? page)
        {
            var pageNumeber = page ?? 1;
            var searchParameters = (LogSearchViewModel)TempData.Peek(Constants.SEARCH_PARAMETER_KEY);
            TempData[Constants.PAGE_NUMBER_KEY] = page;

            var query = getQueryData(searchParameters);

            ViewBag.ACTION_TYPE = Helpers.TenderHelper.getItemActionTypeOptions();
            ManageViewModel result = new ManageViewModel()
            {
                SearchParameter = searchParameters ?? new LogSearchViewModel(),
                SMCLOGSList = query.ToPagedList(pageNumeber, pageSize)
            };

            return View("Index", result);
        }

        [HttpPost]
        public ActionResult Index(ViewModels.ManageViewModel Model)
        {
            TempData[Constants.SEARCH_PARAMETER_KEY] = Model.SearchParameter;
            TempData[Constants.PAGE_NUMBER_KEY] = 1;

            var query = getQueryData(Model.SearchParameter);

            ViewBag.ACTION_TYPE = Helpers.TenderHelper.getItemActionTypeOptions();
            var result = new ManageViewModel()
            {
                SearchParameter = Model.SearchParameter,
                SMCLOGSList = query.ToPagedList(1, pageSize)
            };

            return View(result);
        }

    }
}