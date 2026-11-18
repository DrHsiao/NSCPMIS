using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.DynamicData;
using System.Web.Mvc;
using Aspose.Cells.Drawing;
using Newtonsoft.Json;
using WCals.Models;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEPMSModels;

namespace WCals.Controllers._1._1._1_1._1._4
{
    public class BranchPlanMaintenController : Controller
    {
        CMBASEEntities db1 = new CMBASEEntities();
        CMEPMSEntities db3 = new CMEPMSEntities();
        // GET: BranchPlanMainten
        public ActionResult Index()
        {
            var PJMFORM = db3.PJMFORM.ToList();
            return View(PJMFORM);
        }

        public JsonResult Choose_SYR(string SYR, string PJ_NAME)
        {
            var PID = db3.PJMFORM.Where(m => m.PJ_NAME == PJ_NAME && m.SYR == SYR && m.PJ_TYPE == "W").FirstOrDefault().PID;
            var PJ = db3.PJMFORM.Where(m => m.AN_PID == PID).ToList();
            string json = JsonConvert.SerializeObject(PJ);
            return Json(json);
        }

        public JsonResult Choose_Branch_PJ_NAME(string PJ_NAME)
        {
            var SYR = db3.PJMFORM.Where(m => m.PJ_TYPE == "B" && m.PJ_NAME == PJ_NAME).ToList();
            string json = JsonConvert.SerializeObject(SYR);
            return Json(json);
        }

        public ActionResult BRANCHPLAN_CRUD_SORT(string CRUD, PJMFORM_CMEPMS PJ, string Master_Plan_Name, string Master_Plan_SYR)
        {

            var PID = db3.PJMFORM.Where(m => m.PJ_NAME == Master_Plan_Name && m.SYR == Master_Plan_SYR).FirstOrDefault().PID;
            string serializedModel = JsonConvert.SerializeObject(PJ);
            string WK_CODE = db3.PJMFORM.Where(m => m.PID == PID).FirstOrDefault().WK_CODE;

            switch (CRUD)
            {
                case "新增":
                    return RedirectToAction("Branch_Plan_Create", new { Master_Plan_Name = Master_Plan_Name, Master_Plan_SYR = Master_Plan_SYR, PID = PID, WK_CODE = WK_CODE });
                case "修改":
                    return RedirectToAction("Branch_Plan_Edit", new { model = serializedModel, Master_Plan_Name = Master_Plan_Name, Master_Plan_SYR = Master_Plan_SYR });
                case "刪除":
                    return RedirectToAction("Branch_Plan_Delete", new { model = serializedModel });
                default:
                    return View();
            }
        }

        public ActionResult Branch_Plan_Create(string Master_Plan_Name, string Master_Plan_SYR, string PID, string WK_CODE)
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

        public ActionResult Branch_Plan_Edit(string model, string Master_Plan_Name, string Master_Plan_SYR)
        {
            ViewBag.Master_Plan_Name = Master_Plan_Name;
            ViewBag.Master_Plan_SYR = Master_Plan_SYR;
            PJMFORM_CMEPMS PJ = JsonConvert.DeserializeObject<PJMFORM_CMEPMS>(model);
            ViewBag.PID = db3.PJMFORM.Where(m => m.PJ_NAME == PJ.PJ_NAME && m.SYR == PJ.SYR).FirstOrDefault().PID;
            var UserName = PJ.ADV_UID;
            var DeptName = PJ.ADV_DP;
            ViewBag.UserID = db1.OGCUSER.Where(m => m.USR_NAME == UserName).FirstOrDefault().USR_ID;
            ViewBag.DeptID = db1.OGCDEPT.Where(m => m.DEPT_NAME == DeptName).FirstOrDefault().DEPT_NO;
            ViewBag.ADV_DP_NAME = DeptName;
            ViewBag.APV_UID_NAME = UserName;
            ViewBag.DEPT = db1.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            return View(PJ);
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

            var ADV_DP_ID = db1.OGCDEPT.Where(m => m.DEPT_NAME == model.ADV_DP).FirstOrDefault().DEPT_NO;
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


        public ActionResult Branch_Plan_Delete(string model)
        {
            PJMFORM_CMEPMS PJ = JsonConvert.DeserializeObject<PJMFORM_CMEPMS>(model);
            var item = db3.PJMFORM.Where(m => m.SYR == PJ.SYR && m.PJ_NAME == PJ.PJ_NAME).FirstOrDefault();
            db3.PJMFORM.Remove(item);
            db3.SaveChanges();
            return RedirectToAction("Index");
        }

        public ActionResult Sourcemode(int i, string branch_plan, string branch_SYR)
        {
            var url = "";
            switch (i)
            {
                case 1:
                    url = "Item_content?branch_plan="+branch_plan+"&branch_SYR="+branch_SYR;
                    break;
                case 2:
                    url = "annual_distribution?branch_plan="+branch_plan+"&branch_SYR="+branch_SYR;
                    break;
                case 3:
                    url = "centralfunds_proportion?branch_plan="+branch_plan+"&branch_SYR="+branch_SYR;
                    break;
            }
            return Json(url);
        }

        public ActionResult Item_content(string branch_plan, string branch_SYR , string message)
        {
            ViewBag.message = message;
            ViewBag.branch_plan = branch_plan;
            ViewBag.branch_SYR = branch_SYR;
            var PID = db3.PJMFORM.Where(m=>m.PJ_NAME == branch_plan && m.SYR ==branch_SYR).FirstOrDefault().PID;
            ViewBag.PID = PID;
            var PJDCATY = db3.PJDCATY.Where(m=>m.PID == PID).ToList();
            var DEPT = db1.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            ViewBag.DEPT = DEPT;
            var DEPT_NAME = "";
            foreach(var item in DEPT)
            {
                DEPT_NAME += item.DEPT_NAME + ",";
            }
            DEPT_NAME =DEPT_NAME.TrimEnd(',');
            ViewBag.DEPT_NAME = DEPT_NAME;
            var DEPT_ID = "";
            foreach(var item in DEPT)
            {
                DEPT_ID += item.DEPT_NO + ",";
            }
            DEPT_ID = DEPT_ID.TrimEnd(',');
            ViewBag.DEPT_ID = DEPT_ID;

            PJDCATYOGCDEPTOGCUSERVIEWMODEL Model = new PJDCATYOGCDEPTOGCUSERVIEWMODEL();
            Model.OGCUSER = db1.OGCUSER.ToList();
            Model.OGCDEPT = db1.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            Model.PJDCATY = PJDCATY;
            return View(Model);
        }

        public ActionResult Item_content_CRUD(PJDCATY_CMEPMS model, string type, string branch_plan, string branch_SYR)
        {
            var counts = db3.PJDCATY.Where(m => m.PID == model.PID).ToList();

            if(type == "新增")
            {
                model.CID = DateTime.Now.Ticks; /////將CID用成日期流水號

                var DEPT_ID = db1.OGCDEPT.Where(m => m.DEPT_NAME == model.ADV_DP).FirstOrDefault().DEPT_NO;
                var USER_ID = db1.OGCUSER.Where(m => m.USR_NAME == model.ADV_UID).FirstOrDefault().USR_ID;  
                model.ADV_DP = DEPT_ID;
                model.ADV_UID = USER_ID;
                model.ADD_DT = DateTime.Now;
                if(db3.PJDCATY.Where(m=>m.CATY_CODE == model.CATY_CODE && m.PID == model.PID).ToList().Count() > 0)
                {
                    var message = "編號不能重複";
                    return RedirectToAction("Item_content", new { branch_plan = branch_plan, branch_SYR = branch_SYR , message = message});
                }
                else
                {
                    db3.PJDCATY.Add(model);
                    db3.SaveChanges();
                    return RedirectToAction("Item_content" , new { branch_plan = branch_plan, branch_SYR = branch_SYR });
                }
                
            }
            else if (type == "刪除")
            {
                var item = db3.PJDCATY.Where(m=>m.PID == model.PID && m.CATY_CODE == model.CATY_CODE).FirstOrDefault();
                db3.PJDCATY.Remove(item);
                db3.SaveChanges();
                return RedirectToAction("Item_content", new { branch_plan = branch_plan, branch_SYR = branch_SYR });
            }
            else ////修改
            {
                var item = db3.PJDCATY.Where(m => m.PID == model.PID && m.CATY_CODE == model.CATY_CODE).FirstOrDefault();
                item.FCATY_CODE = model.FCATY_CODE;
                item.CATY_NAME = model.CATY_NAME;
                item.ADV_DP = db1.OGCDEPT.Where(m => m.DEPT_NAME == model.ADV_DP).FirstOrDefault().DEPT_NO;
                item.ADV_UID = db1.OGCUSER.Where(m => m.USR_NAME == model.ADV_UID).FirstOrDefault().USR_ID;
                item.CATY_FUND = model.CATY_FUND;
                item.ACC_ID = model.ACC_ID;
                item.CATY_TYPE = model.CATY_TYPE;
                item.UPD_UID = item.ADV_UID;
                item.UPD_DT = DateTime.Now;
                db3.SaveChanges();

                return RedirectToAction("Item_content", new { branch_plan = branch_plan, branch_SYR = branch_SYR });
            }
        }


        public ActionResult annual_distribution(string branch_plan, string branch_SYR)
        {
            ViewBag.branch_plan = branch_plan;
            ViewBag.branch_SYR = branch_SYR;
            var PID = db3.PJMFORM.Where(m => m.PJ_NAME == branch_plan && m.SYR == branch_SYR).FirstOrDefault().PID;
            var PJDFDDS = db3.PJDFDDS.Where(m=>m.PID ==PID).FirstOrDefault();
            if(PJDFDDS != null)
            {
                for (int i = 1; i <= 12; i++)
                {
                    string month = $"M_{i:D2}";

                    var value = (decimal?)typeof(PJDFDDS).GetProperty(month)?.GetValue(PJDFDDS);
                    if (value == null)
                    {
                        typeof(PJDFDDS).GetProperty(month)?.SetValue(PJDFDDS, (decimal)0);
                    }
                }
            }
            else { }
           
            return View(PJDFDDS);
        }

        public ActionResult annual_distribution_Save(decimal[] month, string branch_plan, string branch_SYR , string button)
        {
            if(button == "返回")
            {
                return RedirectToAction("annual_distribution", new { branch_plan = branch_plan, branch_SYR = branch_SYR });
            }
            var PID = db3.PJMFORM.Where(m => m.PJ_NAME == branch_plan && m.SYR == branch_SYR).FirstOrDefault().PID;
            var UserID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            var oldModel = db3.PJDFDDS.Where(m => m.PID == PID && m.SYR == branch_SYR).FirstOrDefault();
            if (oldModel != null)
            {
                oldModel.PID = PID;
                oldModel.SYR = branch_SYR;
                oldModel.M_01 = month[0];
                oldModel.M_02 = month[1];
                oldModel.M_03 = month[2];
                oldModel.M_04 = month[3];
                oldModel.M_05 = month[4];
                oldModel.M_06 = month[5];
                oldModel.M_07 = month[6];
                oldModel.M_08 = month[7];
                oldModel.M_09 = month[8];
                oldModel.M_10 = month[9];
                oldModel.M_11 = month[10];
                oldModel.M_12 = month[11];
                oldModel.ADD_DT = DateTime.Now;
                oldModel.UPD_UID = UserID;
                db3.SaveChanges();
            }
            else
            {
                PJDFDDS model = new PJDFDDS();
                model.PID = PID;
                model.SYR = branch_SYR;
                model.M_01 = month[0];
                model.M_02 = month[1];
                model.M_03 = month[2];
                model.M_04 = month[3];
                model.M_05 = month[4];
                model.M_06 = month[5];
                model.M_07 = month[6];
                model.M_08 = month[7];
                model.M_09 = month[8];
                model.M_10 = month[9];
                model.M_11 = month[10];
                model.M_12 = month[11];
                model.UPD_DT = DateTime.Now;
                model.UPD_UID = UserID;
                db3.PJDFDDS.Add(model);
                db3.SaveChanges();
            }

            return RedirectToAction("annual_distribution", new { branch_plan = branch_plan , branch_SYR  = branch_SYR });
        }

        public ActionResult centralfunds_proportion(string branch_plan, string branch_SYR)
        {
            ViewBag.branch_plan = branch_plan;
            ViewBag.branch_SYR = branch_SYR;
            var PID = db3.PJMFORM.Where(m => m.PJ_NAME == branch_plan && m.SYR == branch_SYR).FirstOrDefault().PID;
            var PJDCTSP = db3.PJDCTSP.Where(m=>m.PID ==PID).ToList();
            ViewBag.CITY = db3.PJDCTSP.DistinctBy(m=>m.CITY).ToList();
            return View(PJDCTSP);

        }

        public ActionResult centralfunds_proportion_Save(decimal[] VALUE, string branch_plan, string branch_SYR, string button)
        {
            if(button == "返回")
            {
                return RedirectToAction("centralfunds_proportion", new { branch_plan = branch_plan, branch_SYR = branch_SYR });
            }
            var PID = db3.PJMFORM.Where(m => m.PJ_NAME == branch_plan && m.SYR == branch_SYR).FirstOrDefault().PID;
            var CITY = db3.PJDCTSP.DistinctBy(m => m.CITY).ToList();
            var oldmodel = db3.PJDCTSP.Where(m=>m.PID == PID).ToList();
            if(oldmodel.Count > 0)
            {
                foreach(var m in oldmodel)
                db3.PJDCTSP.Remove(m);
            }
            for(int i = 0; i < VALUE.Length; i++)
            {
                PJDCTSP model = new PJDCTSP();
                model.PID = PID;
                model.CITY = CITY[i].CITY;
                model.SUP_RATIO = VALUE[i];
                db3.PJDCTSP.Add(model);
                db3.SaveChanges();
            }

            return RedirectToAction("centralfunds_proportion", new { branch_plan = branch_plan, branch_SYR = branch_SYR });
        }
    }
}