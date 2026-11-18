using System;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WCals.Models;
using System.IO;
using WCals.Models.CMBASEModels;
using WCals.Models.CMPSGSModels;

namespace WCals.Controllers
{
    public class PlanController : Controller
    {
        CMBASEEntities db1 = new CMBASEEntities();
        CMPSGSEntities db4 = new CMPSGSEntities();
        // GET: Plan

        public ActionResult Test()
        {
            return View();
        }
        public ActionResult Index(decimal? PPID, string Message)
        {
            ViewBag.TreeView = db4.PJMFORM.ToList();
            ViewBag.Plan = db4.PJMFORM.Where(m => m.PPID == PPID).FirstOrDefault();
            ViewBag.PPID = PPID;
            ViewBag.Source1 = db4.PJDCATY.Where(m => m.PPID == PPID).ToList();
            ViewBag.Source2 = db4.PJDCMCT.Where(m => m.PPID == PPID).ToList();
            ViewBag.Source6 = db4.PJDSYRC.Where(m => m.PPID == PPID).ToList();
            ViewBag.Message = Message;
            var PJDFUND = db4.PJDFUND.Where(m => m.PPID == PPID).ToList();
            var SMCCODE = db4.SMCCODE.Where(m => m.SYS_ID == "FD" && m.CATY_ID == "05").ToList();
            ViewBag.Source3 = (from a in PJDFUND
                               join b in SMCCODE on a.IPJ_CODE equals b.CATY_COD
                               select new Medium_Long_PlanViewModel
                               {
                                   CATY_COD = b.CATY_COD,
                                   SYR = a.SYR,
                                   IPJ_CODE = a.IPJ_CODE,
                                   CODE_DESC = b.CODE_DESC,
                                   FUND = a.FUND,
                                   FUND_DESC = a.FUND_DESC,
                                   YBG_FUND = a.YBG_FUND
                               }).ToList();
            ViewBag.Source3_1 = SMCCODE;
            var ALL_DEPT = db1.OGCDEPT.Where(m => m.DEPT_NO.Contains("A27") || m.AN_DEPT_NO == "000000" || m.DEPT_NAME == "副主管室").ToList();
            var CHOOSE_DEPT = db4.PJDSUPD.Where(m => m.PPID == PPID).ToList();
            ViewBag.NotChoose = ALL_DEPT.Where(a => !CHOOSE_DEPT.Any(b => a.DEPT_NO == b.ADV_DP)).ToList();
            ViewBag.Choose = ALL_DEPT.Where(a => CHOOSE_DEPT.Any(b => a.DEPT_NO == b.ADV_DP)).ToList();
            ViewBag.Source5 = db4.PJDDOCS.Where(m => m.PPID == PPID).ToList();
            return View();
        }

        //[HttpPost]
        public ActionResult CRUD(HttpPostedFileBase file, PJMFORM_CMPSGS model, PJDDOCS file_model, string btn, decimal? TreeView_PPID, string FILE_DELETE_NAME, decimal? FILE_DELETE_PPID, decimal? PPID, string[] DEPT_NO)
        {
            if (DEPT_NO != null)
            {
                var message = "";
                if (PPID == null)
                {
                    message = "請選擇左側計畫";
                    return Json(message, JsonRequestBehavior.AllowGet);
                }

                var PJDSUPD = db4.PJDSUPD.Where(m => m.PPID == PPID).ToList();
                db4.PJDSUPD.RemoveRange(PJDSUPD);
                foreach (var DEPT in DEPT_NO)
                {
                    PJDSUPD Dept_model = new PJDSUPD();
                    Dept_model.PPID = (decimal)PPID;
                    Dept_model.ADV_DP = DEPT;
                    Dept_model.UPD_DT = DateTime.Now;
                    Dept_model.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                    db4.PJDSUPD.Add(Dept_model);
                    db4.SaveChanges();
                }
                message = "儲存成功";
                return Json(message, JsonRequestBehavior.AllowGet);
            }

            if (file != null && file.ContentLength > 0)
            {
                // 處理檔案，例如儲存到伺服器上
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/File"), fileName);
                file.SaveAs(path);
                long fileSizeInBytes = file.ContentLength;
                double fileSizeInKB = fileSizeInBytes / 1024.0;

                ViewBag.Message = "上傳成功";
                file_model.FILE_PATH = path;
                file_model.FILE_NAME = fileName;
                file_model.FILE_SIZE = (decimal?)fileSizeInKB;
                file_model.ADD_DT = DateTime.Now;
                if (file_model.PPID == 0)
                {
                    ViewBag.Message = "請先點選左側的主計畫";
                    return RedirectToAction("Index", "Plan", new { PPID = TreeView_PPID, Message = ViewBag.Message });
                }
                db4.PJDDOCS.Add(file_model);
                db4.SaveChanges();

            }
            else
            {
                ViewBag.Message = "";
            }

            var UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];

            if (FILE_DELETE_NAME != null)
            {
                var FILE = db4.PJDDOCS.Where(m => m.FILE_NAME == FILE_DELETE_NAME && m.PPID == FILE_DELETE_PPID).FirstOrDefault();
                db4.PJDDOCS.Remove(FILE);
                db4.SaveChanges();
            }

            if (btn == "0")  ////清除按鈕
            {
                return RedirectToAction("Index", "Plan");
            }
            if (model.PJ_NAME == null && btn == "1")
            {
                return RedirectToAction("Index", "Plan");
            }
            if (btn == "1" && db4.PJMFORM.Where(m => m.PJ_NAME == model.PJ_NAME).FirstOrDefault() == null) /////新增///// 
            {
                model.UPD_UID = UID;
                model.ADD_DT = DateTime.Now;
                db4.PJMFORM.Add(model);
                db4.SaveChanges();
            }
            if (btn == "2")  ////刪除按鈕
            {
                var Plan = db4.PJMFORM.Where(m => m.PJ_NAME == model.PJ_NAME).FirstOrDefault();
                db4.PJMFORM.Remove(Plan);
                db4.SaveChanges();
            }
            if (btn == "3")  ////修改按鈕
            {
                var Column = db4.PJMFORM.Select(item => item);
                var SamePlan = db4.PJMFORM.Where(m => m.PJ_NAME == model.PJ_NAME).FirstOrDefault();
                if (SamePlan != null)
                {
                    SamePlan.PJ_NAME = model.PJ_NAME;
                    SamePlan.PJ_CODE = model.PJ_CODE;
                    SamePlan.PJ_CATY = model.PJ_CATY;
                    SamePlan.ADV_DP = model.ADV_DP;
                    SamePlan.ADV_UID = model.ADV_UID;
                    SamePlan.PJ_CONT = model.PJ_CONT;
                    SamePlan.STT_DT = model.STT_DT;
                    SamePlan.END_DT = model.END_DT;
                    SamePlan.LEV1_DT = model.LEV1_DT;
                    SamePlan.LEV1_FN = model.LEV1_FN;
                    SamePlan.LEV2_DT = model.LEV2_DT;
                    SamePlan.LEV2_FN = model.LEV2_FN;
                    SamePlan.LEV3_DT = model.LEV3_DT;
                    SamePlan.LEV3_FN = model.LEV3_FN;
                    SamePlan.UPD_DT = DateTime.Now;

                }
                db4.SaveChanges();
            }
            /////之後要存入ADD_DT,UPD_DT UPD_UID 新增日期 更新日期 更改所登入的帳號
            return RedirectToAction("Index", "Plan", new { PPID = TreeView_PPID, Message = ViewBag.Message });
        }

        [HttpPost]
        public ActionResult Create_Sys_type(string cattyname, bool evrffg, DateTime otdsgdt, string otdsgbuss, DateTime dsgrldt, DateTime pccjdgdt, string pccjdgfund, DateTime indsgdt, string ppid)
        {
            var list = new Models.CMPSGSModels.PJDCATY_CMPSGS
            {
                PPID = Source.Tools.StringToDecimal(ppid),
                CATY_NAME = cattyname,
                EVRF_FG = evrffg,
                //OTDSG_DT = otdsgdt,
                OTDSG_BUSS = otdsgbuss,
                //DSGRL_DT = dsgrldt,
                //PCCJDG_DT = pccjdgdt,
                PCCJDG_FUND = Source.Tools.StringToDecimal(pccjdgfund),
                //INIDSG_DT = indsgdt,
                ADD_DT = DateTime.Now,
                UPD_DT = DateTime.Now,
                UPD_UID = Helpers.WebSiteHelper.getUserId() //(string)Session["UsrId"]
        };
            db4.PJDCATY.Add(list);
            db4.SaveChanges();
            return Json(new { success = true });
        }
        [HttpPost]
        public ActionResult Update_Sys_type(string pjcatycode, string cattyname, bool evrffg, DateTime otdsgdt, string otdsgbuss, DateTime dsgrldt, DateTime pccjdgdt, string pccjdgfund, DateTime indsgdt, string ppid)
        {
            decimal ppidDecimal = Source.Tools.StringToDecimal(ppid);
            decimal CIDDecimal = Source.Tools.StringToDecimal(pjcatycode);
            var list = db4.PJDCATY.Where(m => m.PPID == ppidDecimal && m.CID == CIDDecimal).FirstOrDefault();
            if (list != null)
            {
                list.PJ_CATY_CODE = pjcatycode;
                list.CATY_NAME = cattyname;
                list.EVRF_FG = evrffg;
                //list.OTDSG_DT = otdsgdt;
                list.OTDSG_BUSS = otdsgbuss;
                //list.DSGRL_DT = dsgrldt;
                //list.PCCJDG_DT = pccjdgdt;
                list.PCCJDG_FUND = Source.Tools.StringToDecimal(pccjdgfund);
                //list.INIDSG_DT = indsgdt;
                list.UPD_DT = DateTime.Now;
                list.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            };

            db4.SaveChanges();
            return Json(new { success = true });
        }
        [HttpPost]
        public ActionResult Delete_Sys_type(string ppid, string pjcatycode)
        {
            decimal ppidDecimal = Source.Tools.StringToDecimal(ppid);
            decimal CIDDecimal = Source.Tools.StringToDecimal(pjcatycode);
            var itemToDelete = db4.PJDCATY.FirstOrDefault(m => m.PPID == ppidDecimal && m.CID == CIDDecimal);
            if (itemToDelete != null)
            {
                db4.PJDCATY.Remove(itemToDelete);
                db4.SaveChanges();

                // 返回成功消息或重定向到其他页面
                return Json(new { success = true });
            }
            else
            {
                // 如果找不到要删除的对象，则返回错误消息
                return View();
            }
        }
        [HttpPost]
        public ActionResult Create_apply_execution(string cmscode, string cmsname, string cmsunit, string cmsnote, string ppid)
        {
            decimal ppidDecimal = Source.Tools.StringToDecimal(ppid);
            var list = db4.PJDCMCT.Where(m => m.PPID == ppidDecimal && m.CMS_CODE == cmscode).FirstOrDefault();
            if (list != null)
            {
                return Json(new { success = false, error = "項目代碼不能重複" });
            }
            else
            {
                try
                {
                    // 创建新记录并添加到数据库
                    var newItem = new PJDCMCT
                    {
                        PPID = ppidDecimal,
                        CMS_CODE = cmscode,
                        CMS_NAME = cmsname,
                        CMS_UNIT = cmsunit,
                        CMS_NOTE = cmsnote,
                        ADD_DT = DateTime.Now,
                        UPD_DT = DateTime.Now,
                        UPD_UID = Helpers.WebSiteHelper.getUserId() //(string)Session["UsrId"]
                    };

                    db4.PJDCMCT.Add(newItem);
                    db4.SaveChanges();

                    // 返回成功响应
                    return Json(new { success = true });
                }
                catch (Exception ex)
                {
                    // 返回错误响应
                    return Json(new { success = false, error = ex.Message });
                }
            }

        }
        [HttpPost]
        public ActionResult Update_apply_execution(string cmscode, string cmsname, string cmsunit, string cmsnote, string ppid, string cmscode_key)
        {
            decimal ppidDecimal = Source.Tools.StringToDecimal(ppid);
            var delete_Item = db4.PJDCMCT.Where(m => m.PPID == ppidDecimal && m.CMS_CODE == cmscode_key).FirstOrDefault();
            db4.PJDCMCT.Remove(delete_Item);
            var list = db4.PJDCMCT.Where(m => m.PPID == ppidDecimal && m.CMS_CODE == cmscode).FirstOrDefault();
            if (list != null)
            {
                return Json(new { success = false, error = "項目代碼不能重複" });
            }
            else
            {
                try
                {
                    var create_newItem = new PJDCMCT
                    {
                        PPID = ppidDecimal,
                        CMS_CODE = cmscode,
                        CMS_NAME = cmsname,
                        CMS_UNIT = cmsunit,
                        CMS_NOTE = cmsnote,
                        UPD_DT = DateTime.Now,
                        UPD_UID = Helpers.WebSiteHelper.getUserId() //(string)Session["UsrId"]
                    };

                    db4.PJDCMCT.Add(create_newItem);
                    db4.SaveChanges();

                    return Json(new { success = true });
                }
                catch (Exception ex)
                {
                    return Json(new { success = false, error = ex.Message });
                }
            }
        }

        [HttpPost]
        public ActionResult Delete_apply_execution(string cmscode, string ppid)
        {
            decimal ppidDecimal = Source.Tools.StringToDecimal(ppid);
            var itemToDelete = db4.PJDCMCT.FirstOrDefault(m => m.PPID == ppidDecimal && m.CMS_CODE == cmscode);
            if (itemToDelete != null)
            {
                db4.PJDCMCT.Remove(itemToDelete);
                db4.SaveChanges();
                return Json(new { success = true });
            }
            else
            {
                return View();
            }
        }
        [HttpPost]
        public ActionResult Create_pj_fund(string syr, string ipjcode, string fund, string funddesc, string ybgfund, string ppid,string catycod,string codedesc)
        {
            var PJDFUND = db4.PJDFUND.Where(m => m.PPID == Source.Tools.StringToDecimal(ppid)).ToList();
            decimal ppidDecimal = Source.Tools.StringToDecimal(ppid);
            try
            {
                var newItemPJDFUND = new PJDFUND
                {
                    PPID = ppidDecimal,
                    SYR = syr,
                    IPJ_CODE = ipjcode,
                    FUND = Source.Tools.StringToDecimal(fund),
                    FUND_DESC = funddesc,
                    YBG_FUND = Source.Tools.StringToDecimal(ybgfund),
                    ADD_DT = DateTime.Now,
                    UPD_DT = DateTime.Now,
                    UPD_UID = Helpers.WebSiteHelper.getUserId() //(string)Session["UsrId"]
                };
                var newItemSMCCODE = new Models.CMPSGSModels.SMCCODE_CMPSGS
                {
                    SYS_ID = "FD",
                    CATY_ID = "05",
                    CATY_COD = catycod,
                    CODE_DESC = codedesc
                };

                db4.PJDFUND.Add(newItemPJDFUND);
                db4.SMCCODE.Add(newItemSMCCODE);
                db4.SaveChanges();

                // 返回成功响应
                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                // 返回错误响应
                return Json(new { success = false, error = ex.Message });
            }
        }
        [HttpPost]
        public ActionResult Update_pj_fund(string syr, string ipjcode, string fund, string funddesc, string ybgfund, string ppid)
        {
            decimal ppidDecimal = Source.Tools.StringToDecimal(ppid);
            decimal fundDecimal = Source.Tools.StringToDecimal(fund);
            decimal ybgfundDecimal = Source.Tools.StringToDecimal(ybgfund);

            var PJDFUND = db4.PJDFUND.Where(m => m.PPID == ppidDecimal && m.SYR == syr && m.IPJ_CODE == ipjcode).FirstOrDefault();
            if (PJDFUND != null)
            {
                PJDFUND.PPID = ppidDecimal;
                PJDFUND.SYR = syr;
                PJDFUND.IPJ_CODE = ipjcode;
                PJDFUND.FUND = fundDecimal;
                PJDFUND.FUND_DESC = funddesc;
                PJDFUND.YBG_FUND = ybgfundDecimal;
                PJDFUND.UPD_DT = DateTime.Now;
                PJDFUND.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            }

            db4.SaveChanges();

            return Json(new { success = true });
        }
        [HttpPost]
        public ActionResult Delete_pj_fund(string ipjcode, string syr, string ppid)
        {
            decimal ppidDecimal = Source.Tools.StringToDecimal(ppid);
            var itemToDelete = db4.PJDFUND.FirstOrDefault(m => m.PPID == ppidDecimal && m.IPJ_CODE == ipjcode && m.SYR == syr);
            if (itemToDelete != null)
            {
                db4.PJDFUND.Remove(itemToDelete);
                db4.SaveChanges();
                return Json(new { success = true });
            }
            else
            {
                return View();
            }
        }
        [HttpPost]
        public ActionResult Create_apply_controller(string psgname, string bgnsyr, string endsyr, bool openfg, string ppid)
        {
            try
            {
                // 创建新记录并添加到数据库
                var newItem = new PJDSYRC // 请根据实际情况更改模型名称
                {
                    PSG_NAME = psgname,
                    BGN_SYR = bgnsyr,
                    END_SYR = endsyr,
                    OPEN_FG = openfg,
                    PPID = Source.Tools.StringToDecimal(ppid),
                    ADD_DT = DateTime.Now,
                    UPD_DT = DateTime.Now,
                    UPD_UID = Helpers.WebSiteHelper.getUserId() //(string)Session["UsrId"]
                };

                db4.PJDSYRC.Add(newItem);
                db4.SaveChanges();

                // 返回成功响应
                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                // 返回错误响应
                return Json(new { success = false, error = ex.Message });
            }
        }
        [HttpPost]
        public ActionResult Update_apply_controller(string yrcid, string psgname, string bgnsyr, string endsyr, bool openfg, string ppid)
        {
            decimal ppidDecimal = Source.Tools.StringToDecimal(ppid);
            decimal yrcidDecimal = Source.Tools.StringToDecimal(yrcid);
            var PJDSYRC = db4.PJDSYRC.Where(m => m.YRC_ID == yrcidDecimal).FirstOrDefault();
            if (PJDSYRC != null)
            {
                PJDSYRC.PPID = ppidDecimal;
                PJDSYRC.PSG_NAME = psgname;
                PJDSYRC.BGN_SYR = bgnsyr;
                PJDSYRC.END_SYR = endsyr;
                PJDSYRC.OPEN_FG = openfg;
                PJDSYRC.UPD_DT = DateTime.Now;
                PJDSYRC.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            }
            db4.SaveChanges();

            return Json(new { success = true });
        }
        [HttpPost]
        public ActionResult Delete_apply_controller(string yrcId)
        {
            decimal yrcIdDecimal = Source.Tools.StringToDecimal(yrcId);
            var itemToDelete = db4.PJDSYRC.FirstOrDefault(m => m.YRC_ID == yrcIdDecimal);
            if (itemToDelete != null)
            {
                db4.PJDSYRC.Remove(itemToDelete);
                db4.SaveChanges();
                return Json(new { success = true });
            }
            else
            {
                return View();
            }

        }
        //[HttpPost]
        //public ActionResult Upload()
        //{

        //    try
        //    {
        //        HttpPostedFileBase file = Request.Files[0]; // 获取上传的文件

        //        if (file != null && file.ContentLength > 0)
        //        {
        //            // 在此处处理文件，例如保存到服务器上的特定位置
        //            string fileName = Path.GetFileName(file.FileName);
        //            string filePath = Path.Combine(Server.MapPath("~/Uploads"), fileName);
        //            file.SaveAs(filePath);

        //            // 可以在这里执行其他操作，例如将文件信息保存到数据库

        //            return Content("文件上传成功");
        //        }
        //        else
        //        {
        //            return Content("未选择文件或文件为空");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        return Content("文件上传失败：" + ex.Message);
        //    }
        //}
    }
}