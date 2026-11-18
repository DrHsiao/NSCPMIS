using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;
using PagedList;
using WCals.ViewModels;

namespace WCals.Controllers
{
    [Authorize]
    public class WSMC0702Controller : Controller
    {
        //private readonly Models.WCalsEntities db = new Models.WCalsEntities();
        private readonly Models.CMBASEModels.CMBASEEntities db = new Models.CMBASEModels.CMBASEEntities();
        //public static string SEARCH_PARAMETERS_KEY = "0702";
        //public static string PAGE_QUERY_STRING_KEY = "page";
        private int pageSize = 15;

        private IQueryable<Models.CMBASEModels.OGCUSER> getQueryData(OGCUSERSearchViewModel SearchParameter)
        {
            var query = db.OGCUSER.AsQueryable();

            if (null != SearchParameter)
            {
                if (!string.IsNullOrWhiteSpace(SearchParameter.USR_ID))
                {
                    SearchParameter.USR_ID = SearchParameter.USR_ID.Trim();
                    query = query.Where(
                     f => f.USR_ID.Contains(SearchParameter.USR_ID));
                }

                if (!string.IsNullOrWhiteSpace(SearchParameter.AD_USR_ID))
                {
                    SearchParameter.AD_USR_ID = SearchParameter.AD_USR_ID.Trim();
                    query = query.Where(
                        f => f.LDAP_ID.Contains(SearchParameter.AD_USR_ID));
                }

                if (!string.IsNullOrWhiteSpace(SearchParameter.USR_NAME))
                {
                    SearchParameter.USR_NAME = SearchParameter.USR_NAME.Trim();
                    query = query.Where(
                        f => f.USR_NAME.Contains(SearchParameter.USR_NAME));
                }
            }

            return query = query.OrderBy(f => f.USR_ID);
        }

        // GET: WSMC0702
        [LogAction]
        public ActionResult Index(int? page)
        {
            TempData.Remove(Constants.SEARCH_PARAMETER_KEY);
            var pageNumber = page ?? 1;
            TempData[Constants.PAGE_NUMBER_KEY] = pageNumber;

            var searchParameters = (OGCUSERSearchViewModel)TempData.Peek(Constants.SEARCH_PARAMETER_KEY);
            var query = getQueryData(searchParameters);

            var result = new OGCUSERListViewModel()
            {
                SearchParameter = searchParameters ?? new OGCUSERSearchViewModel(),
                OGCUSER = (IPagedList<Models.CMBASEModels.OGCUSER>)query.ToPagedList(pageNumber, pageSize)
            };

            return View(result);
        }

        public ActionResult PagedIndex(int? page)
        {
            var pageNumeber = page ?? 1;
            var searchParameters = (OGCUSERSearchViewModel)TempData.Peek(Constants.SEARCH_PARAMETER_KEY);
            TempData[Constants.PAGE_NUMBER_KEY] = page;

            var query = getQueryData(searchParameters);
            OGCUSERListViewModel result = new OGCUSERListViewModel()
            {
                SearchParameter = searchParameters ?? new OGCUSERSearchViewModel(),
                OGCUSER = (IPagedList<Models.CMBASEModels.OGCUSER>)query.ToPagedList(pageNumeber, pageSize)
            };


            return View("Index", result);
        }

        [HttpPost]
        public ActionResult Index(OGCUSERListViewModel Model)
        {
            TempData[Constants.SEARCH_PARAMETER_KEY] = Model.SearchParameter;
            TempData[Constants.PAGE_NUMBER_KEY] = 1;

            var query = getQueryData(Model.SearchParameter);

            var result = new OGCUSERListViewModel()
            {
                SearchParameter = Model.SearchParameter,
                OGCUSER = (IPagedList<Models.CMBASEModels.OGCUSER>)query.ToPagedList(1, pageSize)
            };

            return View(result);
        }
        [Authorize]
        public ActionResult New()
        {
            var ogcuser = new Models.CMBASEModels.OGCUSER();
            var model = new OGCUSEREditViewModel()
            {
                OGCUSER = ogcuser
            };

            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult New(OGCUSEREditViewModel model)
        {
            if (ModelState.IsValid)
            {
                var ogcuser = new Models.CMBASEModels.OGCUSER();

                ogcuser.USR_ID = model.VALIDATE_USR_ID;
                //ogcuser.AD_USR_ID = model.VALIDATE_AD_USR_ID;
                ogcuser.USR_NAME = model.OGCUSER.USR_NAME;
                //ogcuser.USR_PSW = model.USR_PSW;
                ogcuser.EMAIL = model.OGCUSER.EMAIL;
                ogcuser.TEL_O = model.OGCUSER.TEL_O;
                ogcuser.TEL_M = model.OGCUSER.TEL_M;
                ogcuser.TEL_F = model.OGCUSER.TEL_F;
                db.OGCUSER.Add(ogcuser);
                db.SaveChanges();

                return RedirectToAction("Index", "WSMC0702");
            }

            return View(model);
        }


        public ActionResult Edit(String id)
        {
            var ogcuser = (from f in db.OGCUSER
                           where f.USR_ID == id
                           select f).FirstOrDefault();

            var ogcdpur_D = (from f in db.OGCDPUR
                             join g in db.OGCDEPT on f.DEPT_NO equals g.DEPT_NO
                             where f.USR_ID == id && f.DATA_TP == "D"
                             select new { f.DEPT_NO, f.USR_ID, g.DEPT_NAME, METHOD = "original" }).ToList();

            var ogcdpur_G = (from f in db.OGCDPUR
                             join g in db.OGCDEPT on f.DEPT_NO equals g.DEPT_NO
                             where f.USR_ID == id && f.DATA_TP == "G"
                             select new { f.DEPT_NO, f.USR_ID, g.DEPT_NAME, METHOD = "original" }).ToList();

            string jsonOGCDPUR_D = JsonConvert.SerializeObject(ogcdpur_D);
            string jsonOGCDPUR_G = JsonConvert.SerializeObject(ogcdpur_G);

            var ogcdept_G = (from f in db.OGCDEPT
                             where f.DATA_TP == "G"
                             select f).ToList();
            ViewBag.DEPT_NO_G = new SelectList(ogcdept_G, "DEPT_NO", "DEPT_NAME");

            var model = new OGCUSEREditViewModel()
            {
                OGCUSER = ogcuser,
                jsonOGCDPUR_D = jsonOGCDPUR_D,
                jsonOGCDPUR_G = jsonOGCDPUR_G,
                page = (int)TempData.Peek(Constants.PAGE_NUMBER_KEY),
                VALIDATE_USR_ID = ogcuser.USR_ID,
                //VALIDATE_AD_USR_ID = ogcuser.AD_USR_ID
            };

            return View(model);
        }

        [HttpPost]
        public ActionResult Edit(OGCUSEREditViewModel model)
        {
            var ogcuser = (from f in db.OGCUSER
                           where f.USR_ID == model.VALIDATE_USR_ID
                           select f).FirstOrDefault();

            String jsonOGCDPUR_D = model.jsonOGCDPUR_D;
            String jsonOGCDPUR_G = model.jsonOGCDPUR_G;

            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    //ogcuser.AD_USR_ID = model.VALIDATE_AD_USR_ID;
                    ogcuser.USR_NAME = model.OGCUSER.USR_NAME;
                    //ogcuser.USR_PSW = model.USR_PSW;
                    ogcuser.EMAIL = model.OGCUSER.EMAIL;
                    ogcuser.TEL_O = model.OGCUSER.TEL_O;
                    ogcuser.TEL_M = model.OGCUSER.TEL_M;
                    ogcuser.TEL_F = model.OGCUSER.TEL_F;
                    db.SaveChanges();

                    proceedEditOGCDPUR(model.OGCUSER.USR_ID, jsonOGCDPUR_D, "D");
                    proceedEditOGCDPUR(model.OGCUSER.USR_ID, jsonOGCDPUR_G, "G");

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    throw;
                }

            }

            return RedirectToAction("PagedIndex", "WSMC0702", new { page = model.page });
        }

        public ActionResult Delete(string id)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var item = db.OGCUSER.Find(id);
                    db.OGCUSER.Remove(item);
                    db.SaveChanges();

                    var ogcdpurList = (from f in db.OGCDPUR
                                   where f.USR_ID == id
                                   select f).ToList();
                    if (ogcdpurList.Count > 0)
                    {
                        foreach (Models.CMBASEModels.OGCDPUR ogcdpur in ogcdpurList)
                        {
                            db.OGCDPUR.Remove(ogcdpur);
                            db.SaveChanges();
                        }
                    }

                    transaction.Commit();

                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    throw;
                }
            }
            return RedirectToAction("Index", "WSMC0702");
        }

        private void proceedEditOGCDPUR(string USR_ID, string json, string type)
        {
            if (String.IsNullOrEmpty(json) == false)
            {
                dynamic x = JsonConvert.DeserializeObject(json);
                for (int i = 0; i < x.Count; i++)
                {
                    if (x[i].METHOD == "new")
                    {
                        Models.CMBASEModels.OGCDPUR table = new Models.CMBASEModels.OGCDPUR();
                        table.USR_ID = USR_ID;
                        table.DEPT_NO = x[i].DEPT_NO;
                        table.DATA_TP = type;
                        //table.ADD_DT = DateTime.Now;
                        table.UPD_UID = User.Identity.Name;
                        db.OGCDPUR.Add(table);
                        db.SaveChanges();
                    }
                    else if (x[i].METHOD == "edit")
                    {
                        if (type == "D")
                        {
                            String deptNo = x[i].DEPT_NO;
                            String usrId = x[i].USR_ID;
                            var ogcdpur = (from f in db.OGCDPUR
                                           where f.DATA_TP == type && f.USR_ID == usrId
                                           select f).FirstOrDefault();
                            if (null != ogcdpur)
                            {
                                if (ogcdpur.DEPT_NO != deptNo)
                                {
                                    // 採取delete insert
                                    var ADD_DT_BK = ogcdpur.ADD_DT;
                                    db.OGCDPUR.Remove(ogcdpur);

                                    Models.CMBASEModels.OGCDPUR ogcdpur_new = new Models.CMBASEModels.OGCDPUR();
                                    ogcdpur_new.ADD_DT = ADD_DT_BK;
                                    ogcdpur_new.DATA_TP = "D";
                                    ogcdpur_new.DEPT_NO = deptNo;
                                    ogcdpur_new.USR_ID = usrId;
                                    ogcdpur_new.UPD_UID = User.Identity.Name;
                                    ogcdpur_new.UPD_DT = DateTime.Now;

                                    db.OGCDPUR.Add(ogcdpur_new);
                                    db.SaveChanges();
                                }
                            }
                            else
                            {
                                Models.CMBASEModels.OGCDPUR table = new Models.CMBASEModels.OGCDPUR();
                                table.USR_ID = USR_ID;
                                table.DEPT_NO = x[i].DEPT_NO;
                                table.DATA_TP = type;
                                table.ADD_DT = DateTime.Now;
                                table.UPD_UID = User.Identity.Name;
                                db.OGCDPUR.Add(table);
                                db.SaveChanges();
                            }
                        }
                    }
                    else if (x[i].METHOD == "delete")
                    {
                        if (x[i].DEPT_NO != null && x[i].USR_ID != null)
                        {
                            String deptNo = x[i].DEPT_NO;
                            String usrId = x[i].USR_ID;
                            var ogcdpur = db.OGCDPUR.Where(f => f.DEPT_NO == deptNo).Where(f => f.USR_ID == usrId).FirstOrDefault();
                            db.OGCDPUR.Remove(ogcdpur);
                            db.SaveChanges();
                        }
                    }
                }
            }
        }

    }
}