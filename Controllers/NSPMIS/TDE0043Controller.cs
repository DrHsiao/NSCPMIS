using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMETMSModels;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS {
    public class TDE0043Controller : Controller {

        private readonly CMETMSEntities db = new CMETMSEntities();


        // GET: PAGE3
        public ActionResult PAGE3(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform
            };

            GetSelectList();
            return View("PAGE3", model);
        }

        public void GetSelectList() {
            SqlCommand cmd = new SqlCommand("SELECT SYR, PID, PJ_NAME FROM PJMFORM WHERE ADV_UID='" + WebSiteHelper.getUserId() + "'ORDER BY SYR DESC");
            //cmd.Parameters.Add(new SqlParameter("PID", PID));
            DataTable pids = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.PIDS = tools.ToSelectList(pids, "PID", "PJ_NAME", null);

            cmd = new SqlCommand("SELECT SYR FROM PJMFORM WHERE ADV_UID='" + WebSiteHelper.getUserId() + "'GROUP BY SYR ORDER BY SYR DESC");
            //cmd.Parameters.Add(new SqlParameter("PID", PID));
            DataTable syrs = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.SYRS = tools.ToSelectList(syrs, "SYR", "SYR", null);






            DataTable dt;
            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
            ViewBag.towns = tools.ToSelectList(util.GetTownNew(null), "TOWN", "TOWN", null);
            //類型/工作類別
            //ViewBag.TD_CATY = tools.ToSelectList(util.GetCMEPMSCODES("FD", "00"), "id", "name", null);

            //執行方式
            dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");
            DataRow dr = dt.NewRow();
            dr["id"] = "01";
            dr["name"] = "01-非BOT";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["id"] = "02";
            dr["name"] = "02-BOT";
            dt.Rows.Add(dr);
            ViewBag.EXE_TYPE = tools.ToSelectList(dt, "id", "name", null);

            //經費層級
            ViewBag.IPJ_CODE = tools.ToSelectList(util.GetCMEPMSCODES("SM", "02"), "id", "name", null);


            //辦理層級
            ViewBag.N_IPJ_CODE = tools.ToSelectList(util.GetCMETMSCODES("SM", "02"), "id", "name", null);

            //分類(雨水用)
            ViewBag.EMG_CATY = tools.ToSelectList(util.GetCMEPMSCODES("TT", "05"), "id", "name", null);

            //會計科目
            ViewBag.ACC_ID = tools.ToSelectList(util.GetCMEPMSCODES("FD", "02"), "id", "name", null);
            //計畫分項
            //ViewBag.PJ_CATY_CODE = tools.ToSelectList(BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT CATY_CODE as id, CATY_NAME as name FROM PJDCATY WHERE PID=" + PID), "id", "name", null);
            //工程類別
            ViewBag.YR_WORK = tools.ToSelectList(util.GetCMETMSCODES("TT", "04"), "id", "name", null);

            //工作類別
            //ViewBag.TD_CATY = tools.ToSelectList(util.GetCMETMSCODES("SM", "09"), "id", "name", null);
            ViewBag.TD_CATY = tools.ToSelectList(util.GetCMETMSCODES("SM", "09"), "id", "name", null);

            //單位
            string sql = "SELECT o.DEPT_NO,\r\n    MAX(o.DEPT_NAME) AS DEPT_NAME, \r\n COUNT(r.USR_ID) AS USER_COUNT \r\nFROM OGCDPUR r\r\nJOIN OGCDEPT o ON o.DEPT_NO = r.DEPT_NO\r\nJOIN OGCUSER u ON u.USR_ID = r.USR_ID\r\nGROUP BY o.DEPT_NO;";
            ViewBag.depts = tools.ToSelectList(SqlDao.CMBASEQuery4List(sql), "DEPT_NO", "DEPT_NAME", "請選擇單位");

            //案件類型
            ViewBag.CATY_COD = tools.ToSelectList(util.GetCMETMSCODES("TT", "03"), "id", "name", null);

            //經費來源
            ViewBag.FIT_CODE = tools.ToSelectList(util.GetCMEPMSCODES("FD", "05"), "id", "name", null);
            //人員
            ViewBag.users = tools.ToSelectList(SqlDao.CMBASEQuery4List("SELECT USR_ID, USR_NAME FROM OGCUSER"), "USR_ID", "USR_NAME", "請選擇");

            // 汙水系統下拉單
            try {
                var sysList = db.PJCSYST.Select(s => new { s.SYS_ID, s.SYS_NAME }).ToList();
                ViewBag.SYS_ID = new SelectList(sysList, "SYS_ID", "SYS_NAME");
            } catch {
                ViewBag.SYS_ID = new SelectList(new List<SelectListItem>(), "Value", "Text");
            }

            // 實施計畫下拉單
            try {
                var pjList = db.PJCFORM.Select(p => new { p.PJ_ID, p.PJ_NAME }).ToList();
                ViewBag.PJ_ID = new SelectList(pjList, "PJ_ID", "PJ_NAME");
            } catch {
                ViewBag.PJ_ID = new SelectList(new List<SelectListItem>(), "Value", "Text");
            }
        }


        private DataTable GetPage3_1List(string TD_CODE, string REC_TP) {

            String sql = @"select 
                BGFID, 
                TD_CODE, 
                REC_TP, 
                case REC_TP
                when 'B' then '預算資料'
                when 'C' then '契約資料'
                when 'S' then '結算資料'
                when 'F' then '決算資料'
                end as REC_TP_NAME,
                SEQ_NO,
                case   
                when SEQ_NO = 0 then 
                case REC_TP 
                when 'B' then '原預算'
                when 'C' then '原契約'
                when 'S' then '結算'
                when 'F' then '決算'
                end
                else 
                case REC_TP 
                when 'B' then '第' + convert(nvarchar(2), SEQ_NO) +'次修正預算'
                when 'C' then '第' + convert(nvarchar(2), SEQ_NO) +'次契約變更'
                when 'S' then '結算'
                when 'F' then '決算'
                end
                end BGRC_NAME,　--期別　
                SVRF_DT,
                SVRF_FNO,
                CVRF_DT,
                CVRF_FNO
                from CMETMS.dbo.TDDBGRC T
                where  TD_CODE=@TD_CODE 
                And REC_TP = @REC_TP ORDER BY SEQ_NO DESC
                ";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("REC_TP", REC_TP));
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            return BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
        }

        public ActionResult PAGE3_1(string TD_CODE, string REC_TP) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform,
                TD_CODE = TD_CODE,
                REC_TP = REC_TP
            };

            GetSelectList();
            ViewBag.List1 = GetPage3_1List(TD_CODE, REC_TP);            
            //DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List();

            return View("PAGE3_1", model);
        }

        //修改
        public ActionResult PAGE3_1_EDIT(string TD_CODE, string SEQ_NO, string REC_TP) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            short seqNo = Convert.ToInt16(SEQ_NO); ;
            //short.TryParse(SEQ_NO, out seqNo);

            // 使用 Entity Framework 查詢 TDDBGRC
            var tddbgrc = db.TDDBGRC.FirstOrDefault(t => t.TD_CODE == TD_CODE && t.SEQ_NO == seqNo && t.REC_TP == REC_TP);

            if(tddbgrc == null) {
                TempData["Error"] = "找不到指定的預算/契約資料";
                return RedirectToAction("PAGE3_1", new { TD_CODE = TD_CODE, REC_TP = REC_TP });
            }

            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform,
                TDDBGRC = tddbgrc
            };

            GetSelectList();
            return View("PAGE3_1_EDIT", model);
        }

        // POST: PAGE3_1_EDIT - 儲存 TDDBGRC 資料
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE3_1_EDIT_Save(TDE0040Model model) {
            if(string.IsNullOrEmpty(model.TDDBGRC.TD_CODE)) {
                TempData["Error"] = "TD_CODE 不可為空";
                GetSelectList();
                return View("PAGE3_1_EDIT", model);
            }

            try {
                var tddbgrc = db.TDDBGRC.FirstOrDefault(t =>
                   t.TD_CODE == model.TDDBGRC.TD_CODE &&
                         t.SEQ_NO == model.TDDBGRC.SEQ_NO &&
                      t.REC_TP == model.TDDBGRC.REC_TP);

                if(tddbgrc == null) {
                    TempData["Error"] = "找不到指定的預算/契約資料";
                    GetSelectList();
                    return View("PAGE3_1_EDIT", model);
                }

                // 更新欄位
                tddbgrc.SVRF_DT = model.TDDBGRC.SVRF_DT;
                tddbgrc.SVRF_FNO = model.TDDBGRC.SVRF_FNO;
                tddbgrc.CVRF_DT = model.TDDBGRC.CVRF_DT;
                tddbgrc.CVRF_FNO = model.TDDBGRC.CVRF_FNO;
                tddbgrc.BGRC_DT = model.TDDBGRC.BGRC_DT;
                tddbgrc.PIROR_FEE = model.TDDBGRC.PIROR_FEE;
                tddbgrc.CUR_FEE = model.TDDBGRC.CUR_FEE;
                tddbgrc.CUR_CHGFEE = model.TDDBGRC.CUR_CHGFEE;
                tddbgrc.UPD_DT = DateTime.Now;
                tddbgrc.UPD_UID = WebSiteHelper.getUserId();

                db.SaveChanges();
                TempData["Success"] = "資料更新成功";

                // 重新載入資料
                var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == model.TDDBGRC.TD_CODE);
                model.TDMFORM = tdmform;
                model.TDDBGRC = tddbgrc;

                GetSelectList();
                return View("PAGE3_1_EDIT", model);
            } catch(Exception ex) {
                TempData["Error"] = "資料更新失敗：" + ex.Message;

                // 重新載入 TDMFORM
                var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == model.TDDBGRC.TD_CODE);
                model.TDMFORM = tdmform;

                GetSelectList();
                return View("PAGE3_1_EDIT", model);
            }
        }

        // 修改C
        public ActionResult PAGE3_1_EDIT_C(string TD_CODE, string SEQ_NO, string REC_TP) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            short seqNo = 0;
            short.TryParse(SEQ_NO, out seqNo);

            // 使用 Entity Framework 查詢 TDDBDRC
            var tddbdrc = db.TDDBDRC.FirstOrDefault(t => t.TD_CODE == TD_CODE && t.SEQ_NO == seqNo);

            // 如果找不到資料，建立新的空白物件供新增使用
            bool isNew = false;
            if(tddbdrc == null) {
                isNew = true;
                tddbdrc = new TDDBDRC {
                    TD_CODE = TD_CODE,
                    SEQ_NO = seqNo
                };
            }

            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform,
                TDDBDRC = tddbdrc
            };

            ViewBag.IsNew = isNew;
            GetSelectList();
            return View("PAGE3_1_EDIT_C", model);
        }

        // POST: PAGE3_1_EDIT_C - 儲存 TDDBDRC 資料 (新增或修改)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE3_1_EDIT_C_Save(TDE0040Model model, short SEQ_NO, bool IsNew = false) {
            if(model.TDDBDRC == null || string.IsNullOrEmpty(model.TDDBDRC.TD_CODE)) {
                TempData["Error"] = "TD_CODE 不可為空";
                GetSelectList();
                ViewBag.IsNew = IsNew;
                return View("PAGE3_1_EDIT_C", model);
            }

            try {
                if(IsNew) {
                    // 新增模式
                    var newTddbdrc = new TDDBDRC {
                        TD_CODE = model.TDDBDRC.TD_CODE,
                        SEQ_NO = SEQ_NO,
                        INET_DT = model.TDDBDRC.INET_DT ?? "",
                        PUB_DT = model.TDDBDRC.PUB_DT ?? "",
                        PUB_FN = model.TDDBDRC.PUB_FN ?? "",
                        RCVFM_DT = model.TDDBDRC.RCVFM_DT ?? "",
                        OPBD_DT = model.TDDBDRC.OPBD_DT ?? "",
                        ASGMAST_DT = model.TDDBDRC.ASGMAST_DT ?? "",
                        MASER = model.TDDBDRC.MASER ?? "",
                        SENFM_DT = model.TDDBDRC.SENFM_DT ?? "",
                        SUG_HARDPN = model.TDDBDRC.SUG_HARDPN,
                        DES_HARDPN = model.TDDBDRC.DES_HARDPN,
                        EXE_UID = model.TDDBDRC.EXE_UID ?? "",
                        SVEY_MEMO = model.TDDBDRC.SVEY_MEMO ?? "",
                        BID_RES = model.TDDBDRC.BID_RES ?? "",
                        FAIL_RESN = model.TDDBDRC.FAIL_RESN ?? "",
                        NOTE = model.TDDBDRC.NOTE ?? "",
                        ADD_DT = DateTime.Now.ToString("yyyyMMdd"),
                        UPD_DT = tools.GetDateByTWDate(DateTime.Now),
                        UPD_TM = DateTime.Now.ToString("HHmmss"),
                        UPD_UID = WebSiteHelper.getUserId() ?? ""
                    };

                    db.TDDBDRC.Add(newTddbdrc);
                    db.SaveChanges();
                    TempData["Success"] = "資料新增成功";

                    // 重新載入資料
                    model.TDDBDRC = newTddbdrc;
                } else {
                    // 修改模式
                    var tddbdrc = db.TDDBDRC.FirstOrDefault(t =>
                      t.TD_CODE == model.TDDBDRC.TD_CODE &&
                          t.SEQ_NO == SEQ_NO);

                    if(tddbdrc == null) {
                        TempData["Error"] = "找不到指定的招標資料";
                        GetSelectList();
                        ViewBag.IsNew = false;
                        return View("PAGE3_1_EDIT_C", model);
                    }

                    // 更新欄位
                    tddbdrc.INET_DT = model.TDDBDRC.INET_DT ?? "";
                    tddbdrc.PUB_DT = model.TDDBDRC.PUB_DT ?? "";
                    tddbdrc.PUB_FN = model.TDDBDRC.PUB_FN ?? "";
                    tddbdrc.RCVFM_DT = model.TDDBDRC.RCVFM_DT ?? "";
                    tddbdrc.OPBD_DT = model.TDDBDRC.OPBD_DT ?? "";
                    tddbdrc.ASGMAST_DT = model.TDDBDRC.ASGMAST_DT ?? "";
                    tddbdrc.MASER = model.TDDBDRC.MASER ?? "";
                    tddbdrc.SENFM_DT = model.TDDBDRC.SENFM_DT ?? "";
                    tddbdrc.SUG_HARDPN = model.TDDBDRC.SUG_HARDPN;
                    tddbdrc.DES_HARDPN = model.TDDBDRC.DES_HARDPN;
                    tddbdrc.EXE_UID = model.TDDBDRC.EXE_UID ?? "";
                    tddbdrc.SVEY_MEMO = model.TDDBDRC.SVEY_MEMO ?? "";
                    tddbdrc.BID_RES = model.TDDBDRC.BID_RES ?? "";
                    tddbdrc.FAIL_RESN = model.TDDBDRC.FAIL_RESN ?? "";
                    tddbdrc.NOTE = model.TDDBDRC.NOTE ?? "";
                    tddbdrc.UPD_DT = DateTime.Now.ToString("yyyyMMdd");
                    tddbdrc.UPD_TM = DateTime.Now.ToString("HHmmss");
                    tddbdrc.UPD_UID = WebSiteHelper.getUserId() ?? "";

                    db.SaveChanges();
                    TempData["Success"] = "資料更新成功";

                    // 重新載入 TDMFORM
                    var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == model.TDDBDRC.TD_CODE);
                    model.TDMFORM = tdmform;
                }

                GetSelectList();
                ViewBag.IsNew = false; // 儲存後變為編輯模式
                return View("PAGE3_1_EDIT_C", model);
            } catch(System.Data.Entity.Validation.DbEntityValidationException dbEx) {
                // 處理 Entity Framework 驗證錯誤
                var errorMessages = new StringBuilder();
                foreach(var validationErrors in dbEx.EntityValidationErrors) {
                    foreach(var validationError in validationErrors.ValidationErrors) {
                        errorMessages.AppendLine($"欄位: {validationError.PropertyName}, 錯誤: {validationError.ErrorMessage}");
                    }
                }
                TempData["Error"] = (IsNew ? "資料新增失敗：" : "資料更新失敗：") + errorMessages.ToString();

                // 重新載入 TDMFORM
                var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == model.TDDBDRC.TD_CODE);
                model.TDMFORM = tdmform;

                GetSelectList();
                ViewBag.IsNew = IsNew;
                return View("PAGE3_1_EDIT_C", model);
            } catch(Exception ex) {
                // 取得完整的錯誤訊息，包括內部例外
                var errorMsg = new StringBuilder();
                errorMsg.AppendLine(ex.Message);
                var inner = ex.InnerException;
                while(inner != null) {
                    errorMsg.AppendLine(inner.Message);
                    inner = inner.InnerException;
                }
                TempData["Error"] = (IsNew ? "資料新增失敗。" : "資料更新失敗：") + errorMsg.ToString();

                // 重新載入 TDMFORM
                var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == model.TDDBDRC.TD_CODE);
                model.TDMFORM = tdmform;

                GetSelectList();
                ViewBag.IsNew = IsNew;
                return View("PAGE3_1_EDIT_C", model);
            }
        }

        // 經費款項
        public ActionResult PAGE3_1_EDIT_D(string TD_CODE, string SEQ_NO, string REC_TP) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            // 使用參數化查詢取得 TDDBGLS 及 SMCCODE 關聯資料（多筆）
            string sql = @"SELECT B.BGFID, B.TD_CODE, B.REC_TP, B.SEQ_NO, B.IT_CODE, B.IT_ORDER, 
     B.AC_ITEM_CODE, S.CODE_DESC AS AC_ITEM_NAME, 
           B.BUDGET_FEE, B.EXE_FEE, B.CLEAR_FEE, B.CLOSE_FEE, B.OUTSRC_FG, B.DEL_FG
         FROM TDDBGLS B 
         LEFT JOIN SMCCODE S ON B.AC_ITEM_CODE = S.CATY_COD AND S.SYS_ID = 'SM' AND S.CATY_ID = '03'
  WHERE B.TD_CODE = @TD_CODE AND B.REC_TP = @REC_TP
     ORDER BY B.IT_ORDER, B.IT_CODE";

            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            cmd.Parameters.Add(new SqlParameter("REC_TP", REC_TP ?? ""));

            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
            ViewBag.List1 = dt;
            ViewBag.REC_TP = REC_TP;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            // 取得會計科目下拉選單
            ViewBag.AC_ITEM_CODE = tools.ToSelectList(util.GetCMETMSCODES("SM", "03"), "id", "name", null);

            GetSelectList();
            return View("PAGE3_1_EDIT_D", model);
        }

        // POST: PAGE3_1_EDIT_D - 儲存 TDDBGLS 資料（多筆）
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE3_1_EDIT_D_Save(string TD_CODE, string REC_TP, string[] BGFID, string[] AC_ITEM_CODE, decimal?[] BUDGET_FEE, bool[] OUTSRC_FG) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                TempData["Error"] = "TD_CODE 不可為空";
                return RedirectToAction("PAGE3_1_EDIT_D", new { TD_CODE = TD_CODE, REC_TP = REC_TP });
            }

            try {
                if(BGFID != null && BGFID.Length > 0) {
                    for(int i = 0; i < BGFID.Length; i++) {
                        decimal bgfidValue = 0;
                        if(decimal.TryParse(BGFID[i], out bgfidValue)) {
                            var tddbgls = db.TDDBGLS.FirstOrDefault(t => t.BGFID == bgfidValue);
                            if(tddbgls != null) {
                                // 更新欄位
                                if(AC_ITEM_CODE != null && i < AC_ITEM_CODE.Length) {
                                    tddbgls.AC_ITEM_CODE = AC_ITEM_CODE[i];
                                }
                                if(BUDGET_FEE != null && i < BUDGET_FEE.Length) {
                                    tddbgls.BUDGET_FEE = BUDGET_FEE[i];
                                }
                                if(OUTSRC_FG != null && i < OUTSRC_FG.Length) {
                                    tddbgls.OUTSRC_FG = OUTSRC_FG[i];
                                }
                                tddbgls.UPD_DT = DateTime.Now;
                                tddbgls.UPD_UID = WebSiteHelper.getUserId();
                            }
                        }
                    }
                    db.SaveChanges();
                    TempData["Success"] = "資料更新成功";
                }
            } catch(Exception ex) {
                TempData["Error"] = "資料更新失敗：" + ex.Message;
            }

            return RedirectToAction("PAGE3_1_EDIT_D", new { TD_CODE = TD_CODE, REC_TP = REC_TP });
        }

        // POST: PAGE3_1_EDIT_D_Add - 新增 TDDBGLS 資料 (AJAX)
        [HttpPost]
        public ActionResult PAGE3_1_EDIT_D_Add(string TD_CODE, string REC_TP, string AC_ITEM_CODE, decimal? BUDGET_FEE, bool OUTSRC_FG) {
            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    return Json(new { success = false, message = "TD_CODE 不可為空" });
                }

                // 取得新的 IT_CODE (流水號)
                string newItCode = "01";
                var maxItCode = db.TDDBGLS
              .Where(t => t.TD_CODE == TD_CODE && t.REC_TP == REC_TP)
       .OrderByDescending(t => t.IT_CODE)
            .Select(t => t.IT_CODE)
        .FirstOrDefault();

                if(!string.IsNullOrEmpty(maxItCode)) {
                    int maxCode = 0;
                    if(int.TryParse(maxItCode, out maxCode)) {
                        newItCode = (maxCode + 1).ToString("00");
                    }
                }

                // 新增資料
                var newTddbgls = new TDDBGLS {
                    TD_CODE = TD_CODE,
                    REC_TP = REC_TP,
                    SEQ_NO = 0,
                    IT_CODE = newItCode,
                    AC_ITEM_CODE = AC_ITEM_CODE,
                    BUDGET_FEE = BUDGET_FEE,
                    OUTSRC_FG = OUTSRC_FG,
                    DEL_FG = "N",
                    ADD_DT = DateTime.Now,
                    UPD_DT = DateTime.Now,
                    UPD_UID = WebSiteHelper.getUserId()
                };

                db.TDDBGLS.Add(newTddbgls);
                db.SaveChanges();

                return Json(new {
                    success = true,
                    message = "新增成功",
                    data = new {
                        BGFID = newTddbgls.BGFID,
                        IT_CODE = newTddbgls.IT_CODE,
                        AC_ITEM_CODE = newTddbgls.AC_ITEM_CODE,
                        BUDGET_FEE = newTddbgls.BUDGET_FEE,
                        OUTSRC_FG = newTddbgls.OUTSRC_FG
                    }
                });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增失敗：" + ex.Message });
            }
        }

        // POST: PAGE3_1_EDIT_D_Update - 更新單筆 TDDBGLS 資料 (AJAX)
        [HttpPost]
        public ActionResult PAGE3_1_EDIT_D_Update(decimal BGFID, string AC_ITEM_CODE, decimal? BUDGET_FEE, bool OUTSRC_FG) {
            try {
                var tddbgls = db.TDDBGLS.FirstOrDefault(t => t.BGFID == BGFID);
                if(tddbgls == null) {
                    return Json(new { success = false, message = "找不到指定的資料" });
                }

                tddbgls.AC_ITEM_CODE = AC_ITEM_CODE;
                tddbgls.BUDGET_FEE = BUDGET_FEE;
                tddbgls.OUTSRC_FG = OUTSRC_FG;
                tddbgls.UPD_DT = DateTime.Now;
                tddbgls.UPD_UID = WebSiteHelper.getUserId();

                db.SaveChanges();

                return Json(new { success = true, message = "更新成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        // POST: PAGE3_1_EDIT_D_Delete - 刪除 TDDBGLS 資料 (AJAX)
        [HttpPost]
        public ActionResult PAGE3_1_EDIT_D_Delete(decimal BGFID) {
            try {
                var tddbgls = db.TDDBGLS.FirstOrDefault(t => t.BGFID == BGFID);
                if(tddbgls == null) {
                    return Json(new { success = false, message = "找不到指定的資料" });
                }

                db.TDDBGLS.Remove(tddbgls);
                db.SaveChanges();

                return Json(new { success = true, message = "刪除成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }


        #region PAGE3_1_EDIT_E - XML 標單匯入功能
        /*
        =====================================================================================
        XML 標單匯入功能說明 (PAGE3_1_EDIT_E)
        =====================================================================================

        【功能概述】
        此功能用於匯入工程會電子標單 (ETenderSheet) XML 檔案，解析後儲存至資料庫。

        【資料流程】
        1. 使用者上傳 XML 檔案
        2. 系統使用 xmlutil.ParseETenderSheetXmlContent() 解析 XML
        3. 解析結果暫存於 Session，供預覽使用
        4. 使用者確認後，將資料儲存至：
           - TDDBGAC：樹狀結構 (左側樹狀選單)
           - TDDWQTY：工項明細 (右側 Grid)

        【資料表對應】

        XML DetailList 欄位          ->  TDDBGAC 欄位
        ─────────────────────────────────────────────
        ItemKey                      ->  (用於建立父子關係)
        ParentItemKey                ->  (用於建立父子關係)
        ItemNo (壹,一,a,1,2...)      ->  AC_ITEM_CODE (項次編號)
        RefItemCode                  ->  IT_CODE (參考代碼)
        Description                  ->  AC_ITEM_DESC (項目說明)
        ItemKind                     ->  ITEM_TYPE (M=主項/G=一般/V=變動/S=小計)
        Price                        ->  PRICE (單價)
        Amount                       ->  BUDGET_FEE (複價)
        PriceFixed                   ->  FIXED (1=固定/0=非固定)
        Remark 含 [發包]             ->  OUTSRC_FG (發包標記)
        Remark                       ->  NOTE (備註)

        XML DetailList 欄位          ->  TDDWQTY 欄位
        ─────────────────────────────────────────────
        (僅限 ItemKind=general/variablePrice 的底層項目)
        ItemNo                       ->  WK_ITEM_NO (工項編號)
        Description                  ->  WK_ITEM (品名)
        Unit                         ->  WK_UNIT (單位)
        Quantity                     ->  WK_QTY (數量)
        Price                        ->  WK_PRICE (單價)
        Amount                       ->  BD_PRICE (複價)
        ItemKind                     ->  WK_TYPE (G=一般/V=變動)
        PriceFixed                   ->  FIXED (固定標記)

        【ItemKind 對應 ITEM_TYPE】
        mainItem      -> M (主項目：壹、一、a 等層級)
        general       -> G (一般項目：底層工項)
        variablePrice -> V (變動單價項目)
        subtotal      -> S (小計)

        【未來改進方向】
        1. 支援契約書 XML 匯入 (REC_TP='C')
        2. 支援差異比對功能 (新舊版本比較)
        3. 支援批次更新特定欄位
        4. 增加匯入前的資料驗證
        5. 支援 CostBreakdownList 單價分析資料匯入

        =====================================================================================
        */

        /// <summary>
        /// GET: PAGE3_1_EDIT_E - 標單匯入頁面
        /// 進入時會檢查是否已有匯入資料，若有則顯示現有結構
        /// </summary>
        public ActionResult PAGE3_1_EDIT_E(string TD_CODE, short SEQ_NO, string REC_TP) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            // 查詢已匯入的 TDDBGAC 資料 (樹狀結構)
            var existingTreeData = db.TDDBGAC
                .Where(t => t.TD_CODE == TD_CODE && t.REC_TP == REC_TP && t.SEQ_NO == SEQ_NO)
                .OrderBy(t => t.AC_ITEM_CODE)
                .ToList();

            // 查詢已匯入的 TDDWQTY 資料 (Grid 明細)
            var existingGridData = db.TDDWQTY
                .Where(t => t.TD_CODE == TD_CODE && t.REC_TP == REC_TP && t.SEQ_NO == SEQ_NO)
                .OrderBy(t => t.WK_ITEM_NO)
                .ToList();

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE,
                SEQ_NO = SEQ_NO,
                REC_TP = REC_TP,
                TDMFORM = tdmform
            };

            // 傳遞已匯入的資料到 View
            ViewBag.ExistingTreeData = existingTreeData;
            ViewBag.ExistingGridData = existingGridData;
            ViewBag.HasExistingData = existingTreeData.Any() || existingGridData.Any();
            ViewBag.TreeCount = existingTreeData.Count;
            ViewBag.GridCount = existingGridData.Count;

            GetSelectList();
            return View("PAGE3_1_EDIT_E", model);
        }

        /// <summary>
        /// POST: PAGE3_1_EDIT_E_Upload - 上傳並解析 XML 標單檔案
        /// 解析後暫存於 Session，供預覽和確認使用
        /// </summary>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE3_1_EDIT_E_Upload(string TD_CODE, short SEQ_NO, string REC_TP, HttpPostedFileBase xmlFile) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Json(new { success = false, message = "TD_CODE 不可為空" });
            }

            if(xmlFile == null || xmlFile.ContentLength == 0) {
                return Json(new { success = false, message = "請選擇要上傳的 XML 檔案" });
            }

            // 檢查副檔名
            string extension = System.IO.Path.GetExtension(xmlFile.FileName)?.ToLower();
            if(extension != ".xml") {
                return Json(new { success = false, message = "只允許上傳 XML 檔案" });
            }

            try {
                // 讀取 XML 內容
                string xmlContent;
                using(var reader = new System.IO.StreamReader(xmlFile.InputStream, System.Text.Encoding.UTF8)) {
                    xmlContent = reader.ReadToEnd();
                }

                // 使用 xmlutil 解析 XML (位於 classes/john/xmlutil.cs)
                // 解析結果包含：TenderInfo, GrandTotal, DetailList, CostBreakdownList, ResourceList
                DataSet ds = xmlutil.ParseETenderSheetXmlContent(xmlContent);

                // 儲存解析結果到 Session 供後續使用
                // Key 格式：ParsedXmlDataSet_{TD_CODE}_{REC_TP}_{SEQ_NO}
                string sessionKey = $"ParsedXmlDataSet_{TD_CODE}_{REC_TP}_{SEQ_NO}";
                Session[sessionKey] = ds;

                // 準備回傳資料供前端預覽
                var tables = ds.Tables.Cast<DataTable>().Select(t => new {
                    name = t.TableName,
                    rowCount = t.Rows.Count,
                    columns = t.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList()
                }).ToList();

                object tenderInfo = null;

                // 取得標案基本資訊
                if(ds.Tables.Contains("TenderInfo") && ds.Tables["TenderInfo"].Rows.Count > 0) {
                    var tenderInfoRow = ds.Tables["TenderInfo"].Rows[0];
                    tenderInfo = new {
                        contractNo = tenderInfoRow["ContractNo"]?.ToString(),
                        contractTitle = tenderInfoRow["ContractTitle"]?.ToString(),
                        procuringEntity = tenderInfoRow["ProcuringEntity"]?.ToString(),
                        documentType = tenderInfoRow["DocumentType"]?.ToString() // budget=預算書, contract=契約書
                    };
                }

                // 統計 DetailList 中各類型項目數量
                int mainItemCount = 0, generalCount = 0, variablePriceCount = 0, subtotalCount = 0;
                if(ds.Tables.Contains("DetailList")) {
                    var detailList = ds.Tables["DetailList"];
                    foreach(DataRow row in detailList.Rows) {
                        string itemKind = row["ItemKind"]?.ToString();
                        switch(itemKind) {
                            case "mainItem": mainItemCount++; break;
                            case "general": generalCount++; break;
                            case "variablePrice": variablePriceCount++; break;
                            case "subtotal": subtotalCount++; break;
                        }
                    }
                }

                return Json(new {
                    success = true,
                    message = "XML 檔案解析成功",
                    fileName = xmlFile.FileName,
                    tables = tables,
                    tenderInfo = tenderInfo,
                    itemStats = new {
                        mainItemCount,      // 主項目數 (會存入 TDDBGAC)
                        generalCount,       // 一般項目數 (會存入 TDDWQTY)
                        variablePriceCount, // 變動單價項目數 (會存入 TDDWQTY)
                        subtotalCount       // 小計項目數 (會存入 TDDBGAC)
                    }
                });
            } catch(Exception ex) {
                return Json(new { success = false, message = "XML 解析失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// GET: PAGE3_1_EDIT_E_GetTableData - 取得解析後的特定資料表內容 (分頁)
        /// 用於前端預覽解析結果
        /// </summary>
        public ActionResult PAGE3_1_EDIT_E_GetTableData(string TD_CODE, string REC_TP, short SEQ_NO, string tableName, int page = 1, int pageSize = 50) {
            string sessionKey = $"ParsedXmlDataSet_{TD_CODE}_{REC_TP}_{SEQ_NO}";
            var ds = Session[sessionKey] as DataSet;

            if(ds == null) {
                return Json(new { success = false, message = "找不到解析資料，請重新上傳 XML 檔案" }, JsonRequestBehavior.AllowGet);
            }

            if(!ds.Tables.Contains(tableName)) {
                return Json(new { success = false, message = $"找不到資料表 {tableName}" }, JsonRequestBehavior.AllowGet);
            }

            DataTable dt = ds.Tables[tableName];
            int totalRows = dt.Rows.Count;
            int totalPages = (int)Math.Ceiling((double)totalRows / pageSize);

            var rows = dt.AsEnumerable()
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .Select(r => dt.Columns.Cast<DataColumn>().ToDictionary(c => c.ColumnName, c => r[c] == DBNull.Value ? null : r[c]))
                .ToList();

            return Json(new {
                success = true,
                tableName = tableName,
                totalRows = totalRows,
                totalPages = totalPages,
                currentPage = page,
                pageSize = pageSize,
                columns = dt.Columns.Cast<DataColumn>().Select(c => new { name = c.ColumnName, type = c.DataType.Name }).ToList(),
                data = rows
            }, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// POST: PAGE3_1_EDIT_E_SaveToDatabase - 將解析資料儲存至資料庫
        ///
        /// 儲存邏輯：
        /// 1. TDDBGAC (樹狀結構)：儲存所有 DetailList 項目，建立階層關係
        /// 2. TDDWQTY (Grid 明細)：儲存 ItemKind=general/variablePrice 的底層工項
        ///
        /// 注意事項：
        /// - 會先清除該 TD_CODE + REC_TP + SEQ_NO 的舊資料再重新匯入
        /// - AC_ITEM_CODE 儲存 ItemKey (唯一識別碼)
        /// - AN_AC_ITEM_CODE 儲存 ParentItemKey (父項目識別碼)
        /// - IT_CODE 儲存 ItemNo (項次編號如 壹、一、1 等)
        /// </summary>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE3_1_EDIT_E_SaveToDatabase(string TD_CODE, short SEQ_NO, string REC_TP) {
            return ExecuteSaveToDatabase(TD_CODE, SEQ_NO, REC_TP);
        }

        /// <summary>
        /// 執行實際的儲存邏輯 (供 SaveToDatabase 和 QuickUpdate 共用)
        /// </summary>
        private ActionResult ExecuteSaveToDatabase(string TD_CODE, short SEQ_NO, string REC_TP) {
            string sessionKey = $"ParsedXmlDataSet_{TD_CODE}_{REC_TP}_{SEQ_NO}";
       var ds = Session[sessionKey] as DataSet;

            if(ds == null) {
  return Json(new { success = false, message = "找不到解析資料，請重新上傳 XML 檔案" });
            }

       try {
   int treeCount = 0;
    int gridCount = 0;
        string userId = WebSiteHelper.getUserId();
   DateTime now = DateTime.Now;

     // 清除舊資料
                string deleteTddbgacSql = "DELETE FROM TDDBGAC WHERE TD_CODE = @p0 AND REC_TP = @p1 AND SEQ_NO = @p2";
       db.Database.ExecuteSqlCommand(deleteTddbgacSql, TD_CODE, REC_TP, SEQ_NO);

     var oldGridData = db.TDDWQTY.Where(t => t.TD_CODE == TD_CODE && t.REC_TP == REC_TP && t.SEQ_NO == SEQ_NO);
      db.TDDWQTY.RemoveRange(oldGridData);
             db.SaveChanges();

         // 處理 DetailList
             if(ds.Tables.Contains("DetailList")) {
          DataTable dtDetail = ds.Tables["DetailList"];

   foreach(DataRow row in dtDetail.Rows) {
            string itemKey = row["ItemKey"]?.ToString();
      string parentItemKey = row["ParentItemKey"]?.ToString();
      string itemNo = row["ItemNo"]?.ToString()?.Trim();
      string refItemCode = row["RefItemCode"]?.ToString()?.Trim();
    string itemKind = row["ItemKind"]?.ToString();
      string description = row["Description"]?.ToString()?.Trim();
   string unit = row["Unit"]?.ToString()?.Trim();
      string remark = row["Remark"]?.ToString()?.Trim();
             decimal? price = row["Price"] != DBNull.Value ? (decimal?)row["Price"] : null;
      decimal? amount = row["Amount"] != DBNull.Value ? (decimal?)row["Amount"] : null;
decimal? quantity = row["Quantity"] != DBNull.Value ? (decimal?)row["Quantity"] : null;
   bool priceFixed = row["PriceFixed"] != DBNull.Value && (bool)row["PriceFixed"];

       string itemType = ConvertItemKindToType(itemKind);
   bool outsrcFg = !string.IsNullOrEmpty(remark) && remark.Contains("[發包]");

            // 儲存到 TDDBGAC
              string insertTddbgacSql = @"
        INSERT INTO TDDBGAC
     (TD_CODE, REC_TP, SEQ_NO, IT_CODE, AC_ITEM_CODE, AC_ITEM_DESC, AN_AC_ITEM_CODE,
           BUDGET_FEE, OUTSRC_FG, ITEM_TYPE, FIXED, DEL_FG, NOTE, ADD_DT, UPD_DT, UPD_UID, PRICE)
            VALUES
  (@p0, @p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10, @p11, @p12, @p13, @p14, @p15, @p16)";

 db.Database.ExecuteSqlCommand(insertTddbgacSql,
       TD_CODE, REC_TP, SEQ_NO,
        itemNo ?? "",
     itemKey ?? "",
       description ?? "",
      parentItemKey ?? "",
              (object)amount ?? DBNull.Value,
        outsrcFg,
         itemType,
         priceFixed ? (short)1 : (short)0,
     false,
          remark ?? "",
                now, now,
       userId ?? "",
         (object)price ?? DBNull.Value
       );
            treeCount++;

     // 儲存到 TDDWQTY
   if(itemKind == "general" || itemKind == "variablePrice") {
     var newGridItem = new TDDWQTY {
        TD_CODE = TD_CODE,
 REC_TP = REC_TP,
           SEQ_NO = SEQ_NO,
                IT_CODE = refItemCode ?? "",
     IT_NAME = description ?? "",
     AC_ITEM_CODE = parentItemKey ?? "",
          WK_ITEM_NO = itemNo ?? "",
          WK_ITEM = description ?? "",
    WK_ITEM_SPEC = "",
    WK_UNIT = unit ?? "",
            WK_QTY = quantity,
        WK_PRICE = price,
            WK_TYPE = itemKind == "variablePrice" ? "V" : "G",
            FIXED = priceFixed ? (short)1 : (short)0,
       BD_PRICE = amount,
           ADD_DT = now,
     UPD_DT = now,
          UPD_UID = userId ?? ""
            };
        db.TDDWQTY.Add(newGridItem);
     gridCount++;
        }
 }

     db.SaveChanges();
}

      Session.Remove(sessionKey);

       return Json(new {
     success = true,
                    message = $"匯入完成！樹狀結構 {treeCount} 筆，工項明細 {gridCount} 筆",
        treeCount = treeCount,
 gridCount = gridCount
    });
  } catch(Exception ex) {
                var errorMsg = new StringBuilder();
              errorMsg.AppendLine(ex.Message);
      var inner = ex.InnerException;
           while(inner != null) {
                errorMsg.AppendLine(inner.Message);
             inner = inner.InnerException;
           }
           return Json(new { success = false, message = "儲存失敗：" + errorMsg.ToString() });
 }
        }

      /// <summary>
        /// 將 XML 的 ItemKind 轉換為資料庫的 ITEM_TYPE
    /// </summary>
        private string ConvertItemKindToType(string itemKind) {
   switch(itemKind) {
         case "mainItem": return "M";      // 主項目
   case "general": return "G";       // 一般項目
    case "variablePrice": return "V"; // 變動單價
      case "subtotal": return "S";      // 小計
 default: return "O";   // 其他
     }
     }

        /// <summary>
      /// POST: PAGE3_1_EDIT_E_QuickUpdate - 簡易更新功能
        /// 重新上傳 XML 並更新資料，不需要預覽確認
     /// </summary>
        [HttpPost]
        [ValidateAntiForgeryToken]
      public ActionResult PAGE3_1_EDIT_E_QuickUpdate(string TD_CODE, short SEQ_NO, string REC_TP, HttpPostedFileBase xmlFile) {
            if(string.IsNullOrEmpty(TD_CODE)) {
   return Json(new { success = false, message = "TD_CODE 不可為空" });
    }

  if(xmlFile == null || xmlFile.ContentLength == 0) {
               return Json(new { success = false, message = "請選擇要上傳的 XML 檔案" });
  }

    string extension = System.IO.Path.GetExtension(xmlFile.FileName)?.ToLower();
          if(extension != ".xml") {
   return Json(new { success = false, message = "只允許上傳 XML 檔案" });
 }

try {
      // 讀取並解析 XML
       string xmlContent;
    using(var reader = new System.IO.StreamReader(xmlFile.InputStream, System.Text.Encoding.UTF8)) {
           xmlContent = reader.ReadToEnd();
         }

  DataSet ds = xmlutil.ParseETenderSheetXmlContent(xmlContent);

     // 暫存到 Session
        string sessionKey = $"ParsedXmlDataSet_{TD_CODE}_{REC_TP}_{SEQ_NO}";
            Session[sessionKey] = ds;

    // 直接執行儲存邏輯
      return ExecuteSaveToDatabase(TD_CODE, SEQ_NO, REC_TP);
  } catch(Exception ex) {
     return Json(new { success = false, message = "更新失敗：" + ex.Message });
    }
  }

  /// <summary>
      /// GET: PAGE3_1_EDIT_E_GetExistingData - 取得已匯入的資料 (AJAX)
   /// 用於前端顯示已匯入的樹狀結構和 Grid 明細
        /// </summary>
        public ActionResult PAGE3_1_EDIT_E_GetExistingData(string TD_CODE, short SEQ_NO, string REC_TP) {
            try {
         // 查詢 TDDBGAC 樹狀結構
       string treeSql = @"SELECT BGFID, TD_CODE, REC_TP, SEQ_NO, IT_CODE, AC_ITEM_CODE, AC_ITEM_DESC,
   AN_AC_ITEM_CODE, BUDGET_FEE, OUTSRC_FG, ITEM_TYPE, FIXED, NOTE, PRICE, DEL_FG
   FROM TDDBGAC
                 WHERE TD_CODE = @p0 AND REC_TP = @p1 AND SEQ_NO = @p2
            ORDER BY AC_ITEM_CODE";
  var treeData = db.Database.SqlQuery<TDDBGAC>(treeSql, TD_CODE, REC_TP, SEQ_NO).ToList();

 // 查詢 TDDWQTY Grid 明細
         var gridData = db.TDDWQTY
   .Where(t => t.TD_CODE == TD_CODE && t.REC_TP == REC_TP && t.SEQ_NO == SEQ_NO)
     .OrderBy(t => t.WK_ITEM_NO)
           .ToList();

     return Json(new {
      success = true,
       treeData = treeData.Select(t => new {
       t.BGFID,
   t.IT_CODE,
     t.AC_ITEM_CODE,
       t.AC_ITEM_DESC,
             t.AN_AC_ITEM_CODE,
       t.BUDGET_FEE,
     t.OUTSRC_FG,
   t.ITEM_TYPE,
                t.FIXED,
    t.NOTE,
             t.PRICE,
        t.DEL_FG
       }),
       gridData = gridData.Select(g => new {
         g.ID,
                g.IT_CODE,
 g.WK_ITEM_NO,
          g.WK_ITEM,
         g.WK_UNIT,
           g.WK_QTY,
  g.WK_PRICE,
     g.WK_TYPE,
    g.BD_PRICE,
   g.AC_ITEM_CODE
         }),
        treeCount = treeData.Count,
          gridCount = gridData.Count
  }, JsonRequestBehavior.AllowGet);
 } catch(Exception ex) {
          return Json(new { success = false, message = ex.Message }, JsonRequestBehavior.AllowGet);
     }
        }

        #endregion

        // PAGE3_1_EDIT_F - 複製 TDDBGRC 資料並新增一筆 SEQ_NO+1 的記錄
        [HttpPost]
        [ValidateAntiForgeryToken]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "PAGE3_1_EDIT_F")]
        public ActionResult PAGE3_1_EDIT_F(string TD_CODE, string REC_TP) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return PAGE3_1(TD_CODE, REC_TP);
            }

            try {
                // 找到 SEQ_NO 最大的那筆 TDDBGRC 資料
                var maxSeqRecord = db.TDDBGRC
           .Where(t => t.TD_CODE == TD_CODE && t.REC_TP == REC_TP)
  .OrderByDescending(t => t.SEQ_NO)
       .FirstOrDefault();

                short newSeqNo = 1; // 預設從 1 開始

                if(maxSeqRecord != null) {
                    // 有找到資料，複製並 SEQ_NO + 1
                    newSeqNo = (short)(maxSeqRecord.SEQ_NO + 1);

                    var newTddbgrc = new TDDBGRC {
                        TD_CODE = maxSeqRecord.TD_CODE,
                        REC_TP = maxSeqRecord.REC_TP,
                        SEQ_NO = newSeqNo,
                        // 複製其他欄位
                        SVRF_DT = maxSeqRecord.SVRF_DT,
                        SVRF_FNO = maxSeqRecord.SVRF_FNO,
                        CVRF_DT = maxSeqRecord.CVRF_DT,
                        CVRF_FNO = maxSeqRecord.CVRF_FNO,
                        BGRC_DT = maxSeqRecord.BGRC_DT,
                        PIROR_FEE = maxSeqRecord.PIROR_FEE,
                        CUR_FEE = maxSeqRecord.CUR_FEE,
                        CUR_CHGFEE = maxSeqRecord.CUR_CHGFEE,
                        // 設定新增/更新資訊
                        ADD_DT = DateTime.Now,
                        UPD_DT = DateTime.Now,
                        UPD_UID = WebSiteHelper.getUserId()
                    };

                    db.TDDBGRC.Add(newTddbgrc);
                    db.SaveChanges();

                    TempData["Success"] = $"已成功新增第 {newSeqNo} 次修正預算/契約變更";
                } else {
                    // 沒有找到資料，建立第一筆 (SEQ_NO = 0 為原始資料，這裡從 1 開始為變更)
                    var newTddbgrc = new TDDBGRC {
                        TD_CODE = TD_CODE,
                        REC_TP = REC_TP,
                        SEQ_NO = 0, // 第一筆為原預算/原契約
                        ADD_DT = DateTime.Now,
                        UPD_DT = DateTime.Now,
                        UPD_UID = WebSiteHelper.getUserId()
                    };

                    db.TDDBGRC.Add(newTddbgrc);
                    db.SaveChanges();

                    newSeqNo = 0;
                    TempData["Success"] = "已成功新增原預算/原契約資料";
                }

                // 導向 PAGE3_1_EDIT 進行編輯
                return RedirectToAction("PAGE3_1_EDIT", new {
                    TD_CODE = TD_CODE,
                    SEQ_NO = newSeqNo.ToString(),
                    REC_TP = REC_TP
                });

            } catch(System.Data.Entity.Validation.DbEntityValidationException dbEx) {
                var errorMessages = new StringBuilder();
                foreach(var validationErrors in dbEx.EntityValidationErrors) {
                    foreach(var validationError in validationErrors.ValidationErrors) {
                        errorMessages.AppendLine($"欄位: {validationError.PropertyName}, 錯誤: {validationError.ErrorMessage}");
                    }
                }
                TempData["Error"] = "新增失敗：" + errorMessages.ToString();
                return RedirectToAction("PAGE3_1", new { TD_CODE = TD_CODE, REC_TP = REC_TP });
            } catch(Exception ex) {
                TempData["Error"] = "新增失敗：" + GetFullErrorMessage(ex);
                return RedirectToAction("PAGE3_1", new { TD_CODE = TD_CODE, REC_TP = REC_TP });
            }
        }

        // 取得完整錯誤訊息的輔助方法
        private string GetFullErrorMessage(Exception ex) {
            var sb = new StringBuilder();
            sb.AppendLine(ex.Message);

            var inner = ex.InnerException;
            int level = 1;
            while(inner != null && level <= 5) {
                sb.AppendLine($"[內部錯誤 {level}]: {inner.Message}");
                inner = inner.InnerException;
                level++;
            }

            return sb.ToString();
        }
    }
}