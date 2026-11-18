using Newtonsoft.Json;
using System;
using System.Data;
using System.Linq;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using WCals.Models;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEPMSModels;

namespace WCals.Controllers
{
    public class WorkPlanMaintenController : Controller
    {
        CMBASEEntities db1 = new CMBASEEntities();
        CMEPMSEntities db3 = new CMEPMSEntities();
        // GET: WorkPlanMainten
        public ActionResult Index()
        {
            var PJ_NAME = db3.PJV_PJMFORM.DistinctBy(m => m.PJ_NAME).Where(m => m.PJ_TYPE == "W").ToList();
            ViewBag.PJ_NAME = PJ_NAME;
            return View();
        }


        public ActionResult BranchPlan(string PJ_NAME, string PJ_SYR)
        {
            var PJ_PID = db3.PJV_PJMFORM.Where(m => m.PJ_NAME == PJ_NAME && m.PJ_SYR == PJ_SYR && m.PJ_TYPE == "W").FirstOrDefault().PID;
            var PJ = db3.PJMFORM.Where(m => m.AN_PID == PJ_PID && m.PJ_TYPE == "B").ToList();
            ViewBag.PJ_NAME = PJ_NAME;
            ViewBag.PJ_SYR = PJ_SYR;
            ViewBag.PJ = PJ.DistinctBy(m=> m.PJ_NAME);
            ViewBag.PID = PJ_PID;
            return View();
        }

        public JsonResult Choose_PJ_NAME(string PJ_NAME)
        {
            var PJ_SYR = db3.PJV_PJMFORM.Where(m => m.PJ_NAME == PJ_NAME && m.PJ_TYPE == "W")/*.Select(m=>m.PJ_SYR)*/.ToList();
            string json = JsonConvert.SerializeObject(PJ_SYR);
            return Json(json);
        }

        public JsonResult Choose_SYR(string PJ_SYR , string PJ_NAME)
        {
            var PJ = db3.PJV_PJMFORM.Where(m=> m.PJ_NAME == PJ_NAME && m.PJ_SYR == PJ_SYR && m.PJ_TYPE == "W").FirstOrDefault();
            string json = JsonConvert.SerializeObject(PJ);
            return Json(json);
        }

        public JsonResult Choose_Branch_PJ_NAME(string PJ_NAME, decimal PID)
        {
            var SYR = db3.PJMFORM.Where(m => m.AN_PID == PID && m.PJ_TYPE == "B" && m.PJ_NAME == PJ_NAME).ToList();
            string json = JsonConvert.SerializeObject(SYR); 
            return Json(json);
        }

        public JsonResult Choose_Branch_SYR(string PJ_NAME, string SYR)
        {
            var PID = db3.PJMFORM.Where(m=>m.PJ_NAME == PJ_NAME && m.SYR == SYR).FirstOrDefault().PID;
            var WATER_TYPE = db3.PJMFORM.Where(m => m.PJ_NAME == PJ_NAME && m.SYR == SYR).FirstOrDefault().Water_Type;

            var PJ = db3.PJV_PJMFORM.Where(m => m.PID == PID).FirstOrDefault();
            if(PJ.FUND_CATY == "S")
            {
                PJ.FUND_CATY = "特別預算";
            }else if (PJ.FUND_CATY == "Y")
            {
                PJ.FUND_CATY = "公務預算";
            }
            PJ.Water_Type = WATER_TYPE;
            string json = JsonConvert.SerializeObject(PJ);
            return Json(json); 
        }

        public ActionResult CRUD_SORT(string CRUD, PJMFORM_CMEPMS PJ)
        {
            string serializedModel = JsonConvert.SerializeObject(PJ);

            switch (CRUD)
            {
                case "新增": 
                    return RedirectToAction("Master_Plan_Create");
                case "修改":
                    return RedirectToAction("Master_Plan_Edit", new { model = serializedModel});                    
                case "刪除":
                    return RedirectToAction("Master_Plan_Delete", new { model = serializedModel });
                default:
                    return View();
            }
        }

        public ActionResult BRANCHPLAN_CRUD_SORT(string CRUD, PJMFORM_CMEPMS PJ, string Master_Plan_Name, string Master_Plan_SYR)
        {
            var PID = db3.PJV_PJMFORM.Where(m=>m.PJ_NAME == Master_Plan_Name && m.PJ_SYR == Master_Plan_SYR).FirstOrDefault().PID;
            string serializedModel = JsonConvert.SerializeObject(PJ);
            string WK_CODE = db3.PJMFORM.Where(m=>m.PID == PID).FirstOrDefault().WK_CODE;

            switch (CRUD)
            {
                case "新增":
                    return RedirectToAction("Branch_Plan_Create", new { Master_Plan_Name = Master_Plan_Name , Master_Plan_SYR = Master_Plan_SYR , PID = PID, WK_CODE = WK_CODE });
                case "修改":
                    return RedirectToAction("Branch_Plan_Edit", new { model = serializedModel, Master_Plan_Name = Master_Plan_Name, Master_Plan_SYR = Master_Plan_SYR });
                case "刪除":
                    return RedirectToAction("Branch_Plan_Delete", new { model = serializedModel});
                default:
                    return View();
            }
        }

        public ActionResult Master_Plan_Create()
        {
            string pageUserId = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            string pageDeptNo = Helpers.WebSiteHelper.getUserDept(pageUserId); //(string)Session["DeptId"];
            ViewBag.USER = Helpers.WebSiteHelper.getUserName(pageUserId); //(string)Session["UsrName"];
            ViewBag.DEPT = Helpers.WebSiteHelper.getDeptName(pageDeptNo); //(string)Session["Dept"];
            ViewBag.UserID = pageUserId;
            ViewBag.DeptID = pageDeptNo;
            ViewBag.SYR = DateTime.Now.Year - 1911;
            return View();
        }


        ///////////// 主計畫新增 ///////////////
        public ActionResult Master_Plan_Create_Function(PJMFORM_CMEPMS model, string ADV_DP, string ADV_UID)
        {
            model.SLOCK_FG = false;
            model.PJ_TYPE = "W";
            model.FLOCK_FG = false;
            model.ADD_DT = DateTime.Now;
            model.ADV_DP = ADV_DP;
            model.ADV_UID = ADV_UID;

            db3.PJMFORM.Add(model);
            db3.SaveChanges();

            return RedirectToAction("Index");
        }

        public ActionResult Branch_Plan_Create(string Master_Plan_Name, string Master_Plan_SYR, string PID , string WK_CODE)
        {
            string pageUserId = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            string pageDeptNo = Helpers.WebSiteHelper.getUserDept(pageUserId); //(string)Session["DeptId"];
            ViewBag.Master_Plan_Name = Master_Plan_Name;
            ViewBag.Master_Plan_SYR = Master_Plan_SYR;
            ViewBag.Master_Plan_PID = PID;
            ViewBag.Master_Plan_WK_CODE = WK_CODE;
            ViewBag.USER = Helpers.WebSiteHelper.getUserName(pageUserId); //(string)Session["UsrName"];
            ViewBag.DEPT = Helpers.WebSiteHelper.getDeptName(pageDeptNo); //(string)Session["Dept"];
            ViewBag.UserID = pageUserId;
            ViewBag.DeptID = pageDeptNo;
            ViewBag.SYR = DateTime.Now.Year - 1911;
            var DataSource = db1.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            return View(DataSource);
        }

        public ActionResult Branch_Plan_Create_Function(PJMFORM_CMEPMS model)
        {


            model.END_SYR = model.SYR;
            model.SLOCK_FG = false;
            model.PJ_TYPE = "B";
            model.FLOCK_FG = false;
            model.ADD_DT = DateTime.Now;
            model.ADV_DP = Helpers.WebSiteHelper.getUserDept(); //(string)Session["DeptId"];
            model.ADV_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            var FUND_CATY = model.FUND_CATY;
            if (FUND_CATY == "特別預算")
            {
                model.FUND_CATY = "S";
            }
            else if (FUND_CATY == "公務預算")
            {
                model.FUND_CATY = "Y";
            }



            db3.PJMFORM.Add(model);
            db3.SaveChanges();

            return RedirectToAction("Index");
        }


        public ActionResult Master_Plan_Edit(string model)
        {
            PJMFORM_CMEPMS PJ = JsonConvert.DeserializeObject<PJMFORM_CMEPMS>(model);
            var UserID = PJ.ADV_UID;
            var DeptID = PJ.ADV_DP;
            ViewBag.UserID = UserID;
            ViewBag.DeptID = DeptID;
            ViewBag.ADV_DP_NAME = db1.OGCDEPT.Where(m=>m.DEPT_NO == DeptID).FirstOrDefault().DEPT_NAME;
            ViewBag.APV_UID_NAME = db1.OGCUSER.Where(m=>m.USR_ID == UserID).FirstOrDefault().USR_NAME;
            return View(PJ);
        }

        public ActionResult Branch_Plan_Edit(string model , string Master_Plan_Name, string Master_Plan_SYR)
        {
            ViewBag.Master_Plan_Name = Master_Plan_Name;
            ViewBag.Master_Plan_SYR = Master_Plan_SYR;
            PJMFORM_CMEPMS PJ = JsonConvert.DeserializeObject<PJMFORM_CMEPMS>(model);
            ViewBag.PID = db3.PJMFORM.Where(m => m.PJ_NAME == PJ.PJ_NAME && m.SYR == PJ.SYR).FirstOrDefault().PID;
            var UserName = PJ.ADV_UID;
            var DeptName = PJ.ADV_DP;
            ViewBag.UserID = db1.OGCUSER.Where(m=>m.USR_NAME == UserName).FirstOrDefault().USR_ID;
            ViewBag.DeptID = db1.OGCDEPT.Where(m=>m.DEPT_NAME == DeptName).FirstOrDefault().DEPT_NO;
            ViewBag.ADV_DP_NAME = DeptName;
            ViewBag.APV_UID_NAME = UserName;
            ViewBag.DEPT = db1.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            return View(PJ);
        }


        public ActionResult Master_Plan_Edit_Function(PJMFORM_CMEPMS model, string ADV_UID)
        {
            var item = db3.PJMFORM.Where(m => m.PID == model.PID).FirstOrDefault();
            item.PJ_NAME = model.PJ_NAME;
            item.END_SYR = model.END_SYR;
            item.WK_CODE = model.WK_CODE;
            item.PJ_FUND = model.PJ_FUND;
            item.UPD_DT = DateTime.Now;
            item.UPD_UID = ADV_UID;
            db3.SaveChanges();
            return RedirectToAction("Index");
        }

        public ActionResult Branch_Plan_Edit_Function(PJMFORM_CMEPMS model)
        {
            var item = db3.PJMFORM.Where(m => m.PID == model.PID).FirstOrDefault();
            item.PJ_NAME = model.PJ_NAME;
            item.SYR = model.SYR;
            item.END_SYR = model.SYR;
            item.SIP_W_CODE = model.SIP_W_CODE;
            item.PJ_CODE = model.PJ_CODE;
            item.PJ_FUND = model.PJ_FUND;
            item.W_CODE_N = model.W_CODE_N;
            item.E_CODE_N = model.E_CODE_N;
            item.W_CODE = model.W_CODE;
            item.E_CODE = model.E_CODE;

            var ADV_DP_ID = db1.OGCDEPT.Where(m=>m.DEPT_NAME == model.ADV_DP).FirstOrDefault().DEPT_NO;
            item.ADV_DP = ADV_DP_ID;
            item.ADV_UID = model.ADV_UID;

            var FUND_CATY = model.FUND_CATY;
            if (FUND_CATY == "特別預算")
            {
                item.FUND_CATY = "S";
            }
            else if (FUND_CATY == "公務預算")
            {
                item.FUND_CATY = "Y";
            }

            item.GP_ORDER = model.GP_ORDER;
            item.Water_Type = model.Water_Type;
            item.CONT_DESC = model.CONT_DESC;

            item.UPD_DT = DateTime.Now;
            item.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            db3.SaveChanges();
            return RedirectToAction("Index");
        }


        public ActionResult Master_Plan_Delete(string model)
        {
            PJMFORM_CMEPMS PJ = JsonConvert.DeserializeObject<PJMFORM_CMEPMS>(model);
            var item = db3.PJMFORM.Where(m => m.PID == PJ.PID).FirstOrDefault();
            db3.PJMFORM.Remove(item);
            db3.SaveChanges();
            return RedirectToAction("Index");
        }

        public ActionResult Branch_Plan_Delete(string model)
        {
            PJMFORM_CMEPMS PJ = JsonConvert.DeserializeObject<PJMFORM_CMEPMS>(model);
            var item = db3.PJMFORM.Where(m => m.AN_PID == PJ.PID && m.PJ_NAME == PJ.PJ_NAME).FirstOrDefault();
            db3.PJMFORM.Remove(item);
            db3.SaveChanges();
            return RedirectToAction("Index");
        }



        public JsonResult Branch_Plan_DEPTNAME_AJAX(string DEPT_NAME)
        {
            var DEPT_NO = db1.OGCDEPT.Where(m=>m.DEPT_NAME == DEPT_NAME).FirstOrDefault().DEPT_NO;
            var OGCDPUR = db1.OGCDPUR.ToList();
            var OGCUSER = db1.OGCUSER.ToList();
            var query = (from a in OGCDPUR
                        join b in OGCUSER on a.USR_ID equals b.USR_ID
                        where a.USR_ID != null && a.DEPT_NO == DEPT_NO
                        select new
                        {
                            USER_ID = a.USR_ID,
                            USER_NAME = b.USR_NAME
                        }).ToList();
            var result = JsonConvert.SerializeObject(query);
            return Json(result);
        }
    }
}