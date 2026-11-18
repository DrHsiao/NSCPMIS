using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;
using Syncfusion.XlsIO;
using WCals.Helpers;
using WCals.Models;
using WCals.ViewModels;

namespace WCals.Controllers
{
    [Authorize]
    public class UserComponentController : Controller
    {
        private WCals.Models.CMBASEModels.CMBASEEntities db = new WCals.Models.CMBASEModels.CMBASEEntities();
        private const string SEARCH_PARAMETERS_KEY = "UsrComponentSP";
        private int pageSize = 10;

        private IQueryable<OGCUSERItemViewModel> getQueryData(UsrComponentSearchViewModel SearchParameter, bool isAvailiable)
        {

            var query = (from f in db.OGCUSER
                         join o in db.OGCDPUR.Where(r => r.DATA_TP == "D") on f.USR_ID equals o.USR_ID into os
                         from o in os.DefaultIfEmpty()
                         join k in db.OGCDEPT on o.DEPT_NO equals k.DEPT_NO
                         select new OGCUSERItemViewModel()
                         {
                             //OGCUSER = f,
                             DEPT_NAME = k.DEPT_NAME
                         });

            if (isAvailiable)
            {
                query.Where(f => f.OGCUSER.USR_STAT ==true);
            }
            else
            {
                query.Where(f => f.OGCUSER.USR_STAT == false);
            }

            if (null != SearchParameter)
            {
                if (!string.IsNullOrWhiteSpace(SearchParameter.USR_ID))
                {
                    SearchParameter.USR_ID = SearchParameter.USR_ID.Trim();
                    query = query.Where(
                        f => f.OGCUSER.USR_ID.Contains(SearchParameter.USR_ID));
                }
                if (!string.IsNullOrWhiteSpace(SearchParameter.AD_USR_ID))
                {
                    SearchParameter.AD_USR_ID = SearchParameter.AD_USR_ID.Trim();
                    query = query.Where(
                        f => f.OGCUSER.USR_ID.Contains(SearchParameter.AD_USR_ID));
                }
                if (!string.IsNullOrWhiteSpace(SearchParameter.USR_NAME))
                {
                    SearchParameter.USR_NAME = SearchParameter.USR_NAME.Trim();
                    query = query.Where(
                        f => f.OGCUSER.USR_NAME.Contains(SearchParameter.USR_NAME));
                }
                if (!string.IsNullOrWhiteSpace(SearchParameter.DEPT_NAME))
                {
                    SearchParameter.DEPT_NAME = SearchParameter.DEPT_NAME.Trim();
                    query = query.Where(
                        f => f.DEPT_NAME.Contains(SearchParameter.DEPT_NAME));
                }
            }


            return query = query.OrderBy(f => f.OGCUSER.USR_ID);
        }

        [HttpGet]
        public ActionResult QueryUsr(string userid, string aduserid, string username, string deptNo)
        {
            TempData.Remove(SEARCH_PARAMETERS_KEY);

            var model = new UsrComponentListViewModel();
            var SearchParameter = new UsrComponentSearchViewModel();
            SearchParameter.USR_ID = userid;
            SearchParameter.AD_USR_ID = aduserid;
            SearchParameter.USR_NAME = username;
            SearchParameter.DEPT_NAME = deptNo;
            model.SearchParameter = SearchParameter;
            model.OGCUSER = getQueryData(SearchParameter, true).ToPagedList(1, pageSize);
            
            return PartialView("_UserQueryComponent", model);
        }

        [HttpGet]
        public ActionResult QueryUsrPage(string USR_ID, string AD_USR_ID, string USR_NAME, string DEPT_NAME, int? page)
        {
            int pageNumber = page ?? 1;
            var searchParameters = new UsrComponentSearchViewModel();
            searchParameters.USR_ID = USR_ID;
            searchParameters.AD_USR_ID = AD_USR_ID;
            searchParameters.USR_NAME = USR_NAME;
            searchParameters.DEPT_NAME = DEPT_NAME;


            var model = new UsrComponentListViewModel()
            {
                SearchParameter = searchParameters,
                OGCUSER = getQueryData(searchParameters, true).ToPagedList(pageNumber, pageSize)
            };
            return PartialView("_UserQueryComponent", model);
        }

        //[HttpPost]
        //public ActionResult QueryUsr(OGCDEPTListViewModel Model)
        //{
        //    TempData[SEARCH_PARAMETERS_OGCDEPT_MODAL] = Model.SearchParameter;

        //    var queryForAvailable = getQueryForOGCDEPTData(Model.SearchParameter, false);

        //    var result = new OGCDEPTListViewModel()
        //    {
        //        SearchParameter = Model.SearchParameter,
        //        OGCDEPT_AVAILABLE = queryForAvailable.ToPagedList(1, pageSizeForModal)
        //    };

        //    return PartialView("_OGCDEPTQueryDept", result);
        //}
    }
}