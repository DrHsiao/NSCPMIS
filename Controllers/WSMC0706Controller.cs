using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WCals.ViewModels;

namespace WCals.Controllers
{
    [Authorize]
    public class WSMC0706Controller : Controller
    {
        private readonly WCals.Models.CMBASEModels.CMBASEEntities db = new WCals.Models.CMBASEModels.CMBASEEntities();
        // GET: WSMC0706
        [LogAction]
        public ActionResult Index()
        {
            String ADV_DP = Helpers.WebSiteHelper.getUserDept();
            var smcnews = db.SMCNEWS.OrderByDescending(f => f.ADD_DT);
            var model = new AnnounceListViewModel()
            {
                SearchParameter = new AnnounceSearchViewModel(),
                SMCNEWS = smcnews.ToList()
            };

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(AnnounceListViewModel Model)
        {
            String ADV_DP = Model.ADV_DP ?? Helpers.WebSiteHelper.getUserDept();
            var query = db.SMCNEWS.AsQueryable();
            if (null != Model.SearchParameter)
            {
                if (!string.IsNullOrWhiteSpace(Model.SearchParameter.SUBJECT))
                {
                    query = query.Where(
                        f => f.HEADLINE.Contains(Model.SearchParameter.SUBJECT));
                }
                if (!string.IsNullOrWhiteSpace(Model.SearchParameter.CONTENT))
                {
                    query = query.Where(
                        f => f.SUMMARY.Contains(Model.SearchParameter.CONTENT));
                }
                if (Model.SearchParameter.POST_DT_START.HasValue)
                {
                    query = query.Where(
                        f => DbFunctions.TruncateTime(f.ADD_DT) >= DbFunctions.TruncateTime(Model.SearchParameter.POST_DT_START));
                }
                if (Model.SearchParameter.POST_DT_END.HasValue)
                {
                    query = query.Where(
                        f => DbFunctions.TruncateTime(f.ADD_DT) <= DbFunctions.TruncateTime(Model.SearchParameter.POST_DT_END));
                }
            }

            query = query.OrderByDescending(f => f.ADD_DT);

            var result = new AnnounceListViewModel()
            {
                SearchParameter = Model.SearchParameter,
                SMCNEWS = query.ToList()
            };

            return View(result);
        }

        public ActionResult Edit(String id)
        {
            if (String.IsNullOrEmpty(id) == false)
            {
                decimal did = Convert.ToDecimal(id);
                var smcnews = db.SMCNEWS.Where(f => f.NEWSID == did).FirstOrDefault();

                ViewData["MainDept"] = false;

                /*if (Helpers.WebSiteHelper.getUserDepartmentType() == Helpers.WebSiteHelper.deptType.MainDept)
                {
                    var ogcdept = db.SMV_OGCSCDP.ToList(); //只要取用總處與區處單位清單資料
                    ViewBag.DEPT_NO = new SelectList(ogcdept, "DEPT_SNO", "DEPT_NAME", smcnews.TO_DP);
                    ViewData["MainDept"] = true;
                }*/

                return View(smcnews);
            }
            else
            {
                return View("/Views/Shared/Error.cshtml");
            }
        }

        [HttpPost]
        public ActionResult Edit(WCals.Models.CMBASEModels.SMCNEWS model)
        {
            var smcnews = db.SMCNEWS.Where(f => f.NEWSID == model.NEWSID).FirstOrDefault();
            if (smcnews != null)
            {
                /*
                smcnews. = model.SUBJECT;
                smcnews.DESP = model.DESP;
                smcnews.CONTENT = model.CONTENT;
                smcnews.FILENAME = model.FILENAME;
                smcnews.START_DT = model.START_DT;
                smcnews.END_DT = model.END_DT;

                if (Helpers.WebSiteHelper.getUserDepartmentType() == Helpers.WebSiteHelper.deptType.MainDept)
                {
                    if (model.TO_DP != null) { smcnews.TO_DP = model.TO_DP; }
                }
                else
                {
                    smcnews.TO_DP = Helpers.WebSiteHelper.getUserDept().Substring(0, 2);
                }
                */
                smcnews.UPD_DT = DateTime.Now;
                smcnews.UPD_UID = User.Identity.Name;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            else
            {
                return View("/Views/Shared/Error.cshtml");
            }
        }

        public ActionResult UploadFile()
        {
            try
            {
                HttpFileCollection files = System.Web.HttpContext.Current.Request.Files;
                if (files.Count > 0)
                {
                    String UploadPath = System.Web.Configuration.WebConfigurationManager.AppSettings["uploadPath"];
                    String oriFileName = System.IO.Path.GetFileName(files[0].FileName);
                    String afilename = Helpers.WebSiteHelper.replaceFileName(oriFileName, "");
                    files[0].SaveAs(System.Web.HttpContext.Current.Server.MapPath("~/" + UploadPath) + System.IO.Path.GetFileName(afilename));
                    return Json(new { Msg = files[0].FileName + "檔案上傳成功", FileName = oriFileName, HashedFileName = afilename }, "text/html", JsonRequestBehavior.AllowGet);
                }
                else
                {
                    //檔案不存在
                    return Json(new { Msg = "Error", FileName = "" }, "text/html", JsonRequestBehavior.AllowGet);
                }

            }
            catch (Exception e)
            {

            }

            return new JsonResult();
        }

        public ActionResult Delete(String id)
        {
            decimal did = Convert.ToDecimal(id);
            var smcnews = db.SMCNEWS.Where(f => f.NEWSID == did).FirstOrDefault();
            db.SMCNEWS.Remove(smcnews);
            db.SaveChanges();
            return RedirectToAction("Index");
        }


        public ActionResult New()
        {
            ViewData["MainDept"] = false;
            /*if (Helpers.WebSiteHelper.getUserDepartmentType() == Helpers.WebSiteHelper.deptType.MainDept)
            {
                var ogcdept = db.SMV_OGCSCDP.ToList(); //只要取用總處與區處單位清單資料;
                ViewData["MainDept"] = true; //總處使用者
                ViewBag.DEPT_NO = new SelectList(ogcdept, "DEPT_SNO", "DEPT_NAME");
            }*/
            return View(new Models.CMBASEModels.SMCNEWS());
        }

        [HttpPost]
        public ActionResult New(Models.CMBASEModels.SMCNEWS model)
        {
            /*DateTime currentTime = DateTime.Now;
            var smcnews = new WCals.Models.CMBASEModels.SMCNEWS();
            smcnews.SUBJECT = model.SUBJECT;
            smcnews.DESP = model.DESP;
            smcnews.CONTENT = model.CONTENT;

            if (Helpers.WebSiteHelper.getUserDepartmentType() == Helpers.WebSiteHelper.deptType.MainDept)
            {
                if (model.TO_DP != null) { smcnews.TO_DP = model.TO_DP; }
            }
            else
            {
                smcnews.TO_DP = Helpers.WebSiteHelper.getUserDept().Substring(0, 2);
            }


            smcnews.POST_DT = currentTime;
            smcnews.START_DT = model.START_DT;
            smcnews.END_DT = model.END_DT;
            smcnews.FILENAME = model.FILENAME;
            smcnews.ADV_DP = Helpers.WebSiteHelper.getUserDept();
            smcnews.ADV_UID = User.Identity.Name;
            smcnews.ADD_DT = currentTime;
            db.SMCNEWS.Add(smcnews);
            db.SaveChanges();*/
            return RedirectToAction("Index");
        }

        public ActionResult DownloadFile(String hrefFileName, String fileName)
        {
            string fileNameValue = HttpUtility.UrlDecode(fileName) ?? hrefFileName;
            return File(Helpers.WebSiteHelper.getFileFromPath(hrefFileName), System.Net.Mime.MediaTypeNames.Application.Octet, fileNameValue);
        }
    }
}