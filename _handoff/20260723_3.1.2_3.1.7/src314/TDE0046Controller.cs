using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMETMSModels;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS {
    public class TDE0046Controller : Controller {

        private readonly CMETMSEntities db = new CMETMSEntities();

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


        // GET: PAGE6 - 施工資料
        public ActionResult PAGE6(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE,
                TDMFORM = tdmform
            };

            GetSelectList();
            return View("PAGE6", model);
        }

        

        // GET: PAGE6_2 - 用地取得
        public ActionResult PAGE6_2(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            ViewBag.TD_CODE = TD_CODE;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_2", model);
        }

        



        #region PAGE6_6 - 接管戶數統計 (TDDSEVL)

        // GET: PAGE6_6 - 接管戶數統計
        public ActionResult PAGE6_6(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢接管戶數統計紀錄
            string sql = @"SELECT EID, TD_CODE, SYR, SMN, EVL_TYPE,
                YEAR_PRE_CON_COUNT_01, PRE_CON_COUNT_01, REL_CON_COUNT_01, CON_DIFF,
                SELF_REMOVE_COUNT_01, FOCUS_REMOVE_COUNT_01
                FROM TDDSEVL
                WHERE TD_CODE = @TD_CODE
                ORDER BY SYR, SMN";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List1 = dt;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_6", model);
        }

        // POST: PAGE6_6_Add - 新增接管戶數統計 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_6_Add(string TD_CODE, string SYR, string SMN, string EVL_TYPE,
            decimal? YEAR_PRE_CON_COUNT_01, decimal? PRE_CON_COUNT_01, decimal? REL_CON_COUNT_01,
            decimal? CON_DIFF, decimal? SELF_REMOVE_COUNT_01, decimal? FOCUS_REMOVE_COUNT_01) {
            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    return Json(new { success = false, message = "TD_CODE 不可為空" });
                }
                if(string.IsNullOrEmpty(SYR)) {
                    return Json(new { success = false, message = "請填寫年度" });
                }

                // 產生 EID
                string eid = DateTime.Now.ToString("yyyyMMddHHmmssfff");

                string actualEvlType = string.IsNullOrEmpty(EVL_TYPE) ? "01" : EVL_TYPE.Trim();

                string insertSql = @"INSERT INTO TDDSEVL
                    (EID, TD_CODE, SYR, SMN, EVL_TYPE, YEAR_PRE_CON_COUNT_01, PRE_CON_COUNT_01,
                    REL_CON_COUNT_01, CON_DIFF, SELF_REMOVE_COUNT_01, FOCUS_REMOVE_COUNT_01, ADD_DT)
                    VALUES (@EID, @TD_CODE, @SYR, @SMN, @EVL_TYPE, @YEAR_PRE_CON_COUNT_01, @PRE_CON_COUNT_01,
                    @REL_CON_COUNT_01, @CON_DIFF, @SELF_REMOVE_COUNT_01, @FOCUS_REMOVE_COUNT_01, @ADD_DT)";

                SqlCommand cmd = new SqlCommand(insertSql);
                cmd.Parameters.AddWithValue("@EID", eid);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@SYR", SYR);
                cmd.Parameters.AddWithValue("@SMN", (object)SMN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@EVL_TYPE", actualEvlType);
                cmd.Parameters.AddWithValue("@YEAR_PRE_CON_COUNT_01", (object)YEAR_PRE_CON_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PRE_CON_COUNT_01", (object)PRE_CON_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REL_CON_COUNT_01", (object)REL_CON_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CON_DIFF", (object)CON_DIFF ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@SELF_REMOVE_COUNT_01", (object)SELF_REMOVE_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@FOCUS_REMOVE_COUNT_01", (object)FOCUS_REMOVE_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ADD_DT", DateTime.Now);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "新增成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_6_Update - 更新接管戶數統計 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_6_Update(string EID, string SYR, string SMN, string EVL_TYPE,
            decimal? YEAR_PRE_CON_COUNT_01, decimal? PRE_CON_COUNT_01, decimal? REL_CON_COUNT_01,
            decimal? CON_DIFF, decimal? SELF_REMOVE_COUNT_01, decimal? FOCUS_REMOVE_COUNT_01) {
            try {
                if(string.IsNullOrEmpty(EID)) {
                    return Json(new { success = false, message = "EID 不可為空" });
                }
                if(string.IsNullOrEmpty(SYR)) {
                    return Json(new { success = false, message = "請填寫年度" });
                }

                string updateSql = @"UPDATE TDDSEVL SET
                    SYR = @SYR, SMN = @SMN, EVL_TYPE = @EVL_TYPE,
                    YEAR_PRE_CON_COUNT_01 = @YEAR_PRE_CON_COUNT_01,
                    PRE_CON_COUNT_01 = @PRE_CON_COUNT_01,
                    REL_CON_COUNT_01 = @REL_CON_COUNT_01,
                    CON_DIFF = @CON_DIFF,
                    SELF_REMOVE_COUNT_01 = @SELF_REMOVE_COUNT_01,
                    FOCUS_REMOVE_COUNT_01 = @FOCUS_REMOVE_COUNT_01,
                    UPD_DT = @UPD_DT
                    WHERE EID = @EID";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.AddWithValue("@EID", EID);
                cmd.Parameters.AddWithValue("@SYR", SYR);
                cmd.Parameters.AddWithValue("@SMN", (object)SMN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@EVL_TYPE", (object)EVL_TYPE ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@YEAR_PRE_CON_COUNT_01", (object)YEAR_PRE_CON_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PRE_CON_COUNT_01", (object)PRE_CON_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REL_CON_COUNT_01", (object)REL_CON_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CON_DIFF", (object)CON_DIFF ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@SELF_REMOVE_COUNT_01", (object)SELF_REMOVE_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@FOCUS_REMOVE_COUNT_01", (object)FOCUS_REMOVE_COUNT_01 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@UPD_DT", DateTime.Now);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_6_Delete - 刪除接管戶數統計 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_6_Delete(string EID) {
            try {
                if(string.IsNullOrEmpty(EID)) {
                    return Json(new { success = false, message = "EID 不可為空" });
                }

                string deleteSql = "DELETE FROM TDDSEVL WHERE EID = @EID";
                SqlCommand cmd = new SqlCommand(deleteSql);
                cmd.Parameters.AddWithValue("@EID", EID);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        #endregion

        #region PAGE6_7 - 化糞池填/拆除統計 (TDDSEVL)

        // GET: PAGE6_7 - 化糞池填/拆除統計
        public ActionResult PAGE6_7(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢化糞池填/拆除統計紀錄
            string sql = @"SELECT EID, TD_CODE, SYR, SMN,
                YEAR_PRE_REMOVE_HOLDER_02, PRE_REMOVE_HOLDER_02, REL_REMOVE_HOLDER_02, NO_REMOVE_HOLDER_02,
                YEAR_PRE_REMOVE_PLACE_02, PRE_REMOVE_PLACE_02, REL_REMOVE_PLACE_02, NO_REMOVE_PLACE_02
                FROM TDDSEVL
                WHERE TD_CODE = @TD_CODE
                ORDER BY SYR, SMN";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List1 = dt;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_7", model);
        }

        // POST: PAGE6_7_Add - 新增化糞池填/拆除統計 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_7_Add(string TD_CODE, string SYR, string SMN,
            decimal? YEAR_PRE_REMOVE_HOLDER_02, decimal? PRE_REMOVE_HOLDER_02, decimal? REL_REMOVE_HOLDER_02, decimal? NO_REMOVE_HOLDER_02,
            decimal? YEAR_PRE_REMOVE_PLACE_02, decimal? PRE_REMOVE_PLACE_02, decimal? REL_REMOVE_PLACE_02, decimal? NO_REMOVE_PLACE_02) {
            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    return Json(new { success = false, message = "TD_CODE 不可為空" });
                }
                if(string.IsNullOrEmpty(SYR)) {
                    return Json(new { success = false, message = "請填寫年度" });
                }

                // 產生 EID
                string eid = DateTime.Now.ToString("yyyyMMddHHmmssfff");

                string insertSql = @"INSERT INTO TDDSEVL
                    (EID, TD_CODE, SYR, SMN, EVL_TYPE, YEAR_PRE_REMOVE_HOLDER_02, PRE_REMOVE_HOLDER_02, REL_REMOVE_HOLDER_02, NO_REMOVE_HOLDER_02,
                    YEAR_PRE_REMOVE_PLACE_02, PRE_REMOVE_PLACE_02, REL_REMOVE_PLACE_02, NO_REMOVE_PLACE_02, ADD_DT)
                    VALUES (@EID, @TD_CODE, @SYR, @SMN, @EVL_TYPE, @YEAR_PRE_REMOVE_HOLDER_02, @PRE_REMOVE_HOLDER_02, @REL_REMOVE_HOLDER_02, @NO_REMOVE_HOLDER_02,
                    @YEAR_PRE_REMOVE_PLACE_02, @PRE_REMOVE_PLACE_02, @REL_REMOVE_PLACE_02, @NO_REMOVE_PLACE_02, @ADD_DT)";

                SqlCommand cmd = new SqlCommand(insertSql);
                cmd.Parameters.AddWithValue("@EID", eid);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@SYR", SYR);
                cmd.Parameters.AddWithValue("@SMN", (object)SMN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@EVL_TYPE", "02");
                cmd.Parameters.AddWithValue("@YEAR_PRE_REMOVE_HOLDER_02", (object)YEAR_PRE_REMOVE_HOLDER_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PRE_REMOVE_HOLDER_02", (object)PRE_REMOVE_HOLDER_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REL_REMOVE_HOLDER_02", (object)REL_REMOVE_HOLDER_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@NO_REMOVE_HOLDER_02", (object)NO_REMOVE_HOLDER_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@YEAR_PRE_REMOVE_PLACE_02", (object)YEAR_PRE_REMOVE_PLACE_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PRE_REMOVE_PLACE_02", (object)PRE_REMOVE_PLACE_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REL_REMOVE_PLACE_02", (object)REL_REMOVE_PLACE_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@NO_REMOVE_PLACE_02", (object)NO_REMOVE_PLACE_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ADD_DT", DateTime.Now);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "新增成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_7_Update - 更新化糞池填/拆除統計 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_7_Update(string EID, string SYR, string SMN,
            decimal? YEAR_PRE_REMOVE_HOLDER_02, decimal? PRE_REMOVE_HOLDER_02, decimal? REL_REMOVE_HOLDER_02, decimal? NO_REMOVE_HOLDER_02,
            decimal? YEAR_PRE_REMOVE_PLACE_02, decimal? PRE_REMOVE_PLACE_02, decimal? REL_REMOVE_PLACE_02, decimal? NO_REMOVE_PLACE_02) {
            try {
                if(string.IsNullOrEmpty(EID)) {
                    return Json(new { success = false, message = "EID 不可為空" });
                }
                if(string.IsNullOrEmpty(SYR)) {
                    return Json(new { success = false, message = "請填寫年度" });
                }

                string updateSql = @"UPDATE TDDSEVL SET
                    SYR = @SYR, SMN = @SMN,
                    YEAR_PRE_REMOVE_HOLDER_02 = @YEAR_PRE_REMOVE_HOLDER_02,
                    PRE_REMOVE_HOLDER_02 = @PRE_REMOVE_HOLDER_02,
                    REL_REMOVE_HOLDER_02 = @REL_REMOVE_HOLDER_02,
                    NO_REMOVE_HOLDER_02 = @NO_REMOVE_HOLDER_02,
                    YEAR_PRE_REMOVE_PLACE_02 = @YEAR_PRE_REMOVE_PLACE_02,
                    PRE_REMOVE_PLACE_02 = @PRE_REMOVE_PLACE_02,
                    REL_REMOVE_PLACE_02 = @REL_REMOVE_PLACE_02,
                    NO_REMOVE_PLACE_02 = @NO_REMOVE_PLACE_02,
                    UPD_DT = @UPD_DT
                    WHERE EID = @EID";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.AddWithValue("@EID", EID);
                cmd.Parameters.AddWithValue("@SYR", SYR);
                cmd.Parameters.AddWithValue("@SMN", (object)SMN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@YEAR_PRE_REMOVE_HOLDER_02", (object)YEAR_PRE_REMOVE_HOLDER_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PRE_REMOVE_HOLDER_02", (object)PRE_REMOVE_HOLDER_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REL_REMOVE_HOLDER_02", (object)REL_REMOVE_HOLDER_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@NO_REMOVE_HOLDER_02", (object)NO_REMOVE_HOLDER_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@YEAR_PRE_REMOVE_PLACE_02", (object)YEAR_PRE_REMOVE_PLACE_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PRE_REMOVE_PLACE_02", (object)PRE_REMOVE_PLACE_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REL_REMOVE_PLACE_02", (object)REL_REMOVE_PLACE_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@NO_REMOVE_PLACE_02", (object)NO_REMOVE_PLACE_02 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@UPD_DT", DateTime.Now);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_7_Delete - 刪除化糞池填/拆除統計 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_7_Delete(string EID) {
            try {
                if(string.IsNullOrEmpty(EID)) {
                    return Json(new { success = false, message = "EID 不可為空" });
                }

                string deleteSql = "DELETE FROM TDDSEVL WHERE EID = @EID";
                SqlCommand cmd = new SqlCommand(deleteSql);
                cmd.Parameters.AddWithValue("@EID", EID);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        #endregion
        // GET: PAGE6_2_1 - 估驗日期列表
        public ActionResult PAGE6_2_1(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            string sql = @"SELECT SEQ_NO, ECC_DT, ECC_FEE, ECCPRV_FEE, ECCDIS_FEE, PPAYDIS_FEE, ADJ_FEE, NOTE 
        FROM CMETMS.dbo.TDDECCR 
          WHERE TD_CODE = @TD_CODE 
            ORDER BY SEQ_NO";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List1 = dt;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_2_1", model);
        }

        // GET: PAGE6_2_2 - 應付未付
        public ActionResult PAGE6_2_2(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            string sql = @"SELECT SEQ_NO, ECC_DT, PPAYDIS_FEE, NOTE 
   FROM CMETMS.dbo.TDDECCR 
                WHERE TD_CODE = @TD_CODE 
      ORDER BY SEQ_NO";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List2 = dt;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_2_2", model);
        }

        // GET: PAGE6_2_3 - 扣款金額
        public ActionResult PAGE6_2_3(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            string sql = @"SELECT SEQ_NO, ECC_DT, ECCDIS_FEE, NOTE 
   FROM CMETMS.dbo.TDDECCR 
             WHERE TD_CODE = @TD_CODE 
         ORDER BY SEQ_NO";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List3 = dt;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_2_3", model);
        }

        // GET: PAGE6_2_Edit - 新增/編輯估驗期別
        public ActionResult PAGE6_2_Edit(string TD_CODE, string SEQ_NO) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            DataTable dt = null;
            bool isNew = string.IsNullOrEmpty(SEQ_NO);

            if(!isNew) {
                string sql = @"SELECT * FROM CMETMS.dbo.TDDECCR 
           WHERE TD_CODE = @TD_CODE AND SEQ_NO = @SEQ_NO";
                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmd.Parameters.Add(new SqlParameter("SEQ_NO", SEQ_NO));
                dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
            }

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.IsNew = isNew;
            ViewBag.Data = (!isNew && dt != null && dt.Rows.Count > 0) ? dt.Rows[0] : null;

            // 3.1.4rev：新增時自動帶出下一個可用期別，使用者不必自己猜、也不會撞到重複主鍵
            if(isNew) {
                string nextSql = @"SELECT ISNULL(MAX(SEQ_NO), 0) + 1 FROM CMETMS.dbo.TDDECCR WHERE TD_CODE = @TD_CODE";
                SqlCommand nextCmd = new SqlCommand(nextSql);
                nextCmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                DataTable dtNext = BASE_DAO.SqlDao.CMETMSQuery4List(nextCmd);
                ViewBag.NextSeqNo = (dtNext != null && dtNext.Rows.Count > 0) ? dtNext.Rows[0][0].ToString() : "1";
            }

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_2_Edit", model);
        }

        // POST: PAGE6_2_Save - 儲存估驗期別
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE6_2_Save(string TD_CODE, int SEQ_NO,
       string ECC_DT, decimal? ECC_FEE, decimal? ECCPRV_FEE,
    decimal? ECCDIS_FEE, decimal? PPAYDIS_FEE, decimal? ADJ_FEE, string NOTE,
              bool IsNew = false) {

            int savedSeqNo = SEQ_NO;

            // 除錯記錄
            System.Diagnostics.Debug.WriteLine($"PAGE6_2_Save called: TD_CODE={TD_CODE}, SEQ_NO={SEQ_NO}, ECC_DT={ECC_DT}, IsNew={IsNew}");

            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    TempData["Error"] = "TD_CODE 不可為空";
                    return RedirectToAction("PAGE6_2_Edit", new { TD_CODE = TD_CODE });
                }

                // 處理日期格式 - 將 "114/06/01" 或 "2025/06/01" 轉換為 "1140601" (7碼民國年格式)
                string formattedEccDt = null;
                if(!string.IsNullOrEmpty(ECC_DT)) {
                    // 移除所有斜線和空白
                    string cleanDate = ECC_DT.Replace("/", "").Replace("-", "").Replace(" ", "").Trim();

                    // 如果是8碼西元年格式 (如 20250601)，轉換為7碼民國年
                    if(cleanDate.Length == 8 && cleanDate.StartsWith("20")) {
                        int year = int.Parse(cleanDate.Substring(0, 4)) - 1911;
                        formattedEccDt = year.ToString("000") + cleanDate.Substring(4, 4);
                    }
                           // 如果已經是7碼民國年格式
                           else if(cleanDate.Length == 7) {
                        formattedEccDt = cleanDate;
                    }
                         // 如果是其他格式，嘗試解析
                         else {
                        // 嘗試解析日期
                        DateTime parsedDate;
                        if(DateTime.TryParse(ECC_DT, out parsedDate)) {
                            int year = parsedDate.Year - 1911;
                            formattedEccDt = year.ToString("000") + parsedDate.ToString("MMdd");
                        } else {
                            // 無法解析，取前7個字元
                            formattedEccDt = cleanDate.Length > 7 ? cleanDate.Substring(0, 7) : cleanDate;
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Formatted ECC_DT: {formattedEccDt}");

                if(IsNew) {
                    // 新增模式
                    string insertSql = @"INSERT INTO TDDECCR 
          (TD_CODE, SEQ_NO, ECC_DT, ECC_FEE, ECCPRV_FEE, ECCDIS_FEE, PPAYDIS_FEE, ADJ_FEE, NOTE, ADD_DT, UPD_DT, UPD_UID)
          VALUES (@TD_CODE, @SEQ_NO, @ECC_DT, @ECC_FEE, @ECCPRV_FEE, @ECCDIS_FEE, @PPAYDIS_FEE, @ADJ_FEE, @NOTE, @ADD_DT, @UPD_DT, @UPD_UID)";

                    SqlCommand cmd = new SqlCommand(insertSql);
                    cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                    cmd.Parameters.AddWithValue("@SEQ_NO", SEQ_NO);
                    cmd.Parameters.AddWithValue("@ECC_DT", (object)formattedEccDt ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ECC_FEE", (object)ECC_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ECCPRV_FEE", (object)ECCPRV_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ECCDIS_FEE", (object)ECCDIS_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@PPAYDIS_FEE", (object)PPAYDIS_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ADJ_FEE", (object)ADJ_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@NOTE", (object)NOTE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ADD_DT", DateTime.Now);
                    cmd.Parameters.AddWithValue("@UPD_DT", DateTime.Now);
                    cmd.Parameters.AddWithValue("@UPD_UID", WebSiteHelper.getUserId() ?? "");

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                    TempData["Success"] = "估驗期別新增成功";
                    TempData["SaveSuccess"] = true;
                } else {
                    // 更新模式
                    string updateSql = @"UPDATE TDDECCR SET 
    ECC_DT = @ECC_DT, ECC_FEE = @ECC_FEE, ECCPRV_FEE = @ECCPRV_FEE, 
         ECCDIS_FEE = @ECCDIS_FEE, PPAYDIS_FEE = @PPAYDIS_FEE, ADJ_FEE = @ADJ_FEE, 
              NOTE = @NOTE, UPD_DT = @UPD_DT, UPD_UID = @UPD_UID
    WHERE TD_CODE = @TD_CODE AND SEQ_NO = @SEQ_NO";

                    SqlCommand cmd = new SqlCommand(updateSql);
                    cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                    cmd.Parameters.AddWithValue("@SEQ_NO", SEQ_NO);
                    cmd.Parameters.AddWithValue("@ECC_DT", (object)formattedEccDt ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ECC_FEE", (object)ECC_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ECCPRV_FEE", (object)ECCPRV_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ECCDIS_FEE", (object)ECCDIS_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@PPAYDIS_FEE", (object)PPAYDIS_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ADJ_FEE", (object)ADJ_FEE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@NOTE", (object)NOTE ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@UPD_DT", DateTime.Now);
                    cmd.Parameters.AddWithValue("@UPD_UID", WebSiteHelper.getUserId() ?? "");

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                    TempData["Success"] = "估驗期別更新成功";
                    TempData["SaveSuccess"] = true;
                }
            } catch(Exception ex) {
                System.Diagnostics.Debug.WriteLine($"PAGE6_2_Save error: {ex.Message}");
                // 取得完整的錯誤訊息
                string fullError = ex.Message;
                if(ex.InnerException != null) {
                    fullError += " | 內部錯誤: " + ex.InnerException.Message;
                }
                // 3.1.4rev：主鍵重複原本會把整段 SQL 錯誤原文丟給使用者(看不懂)，改成看得懂的說明
                if(fullError.Contains("PRIMARY KEY") || fullError.Contains("重複的索引鍵")) {
                    TempData["Error"] = $"期別「{SEQ_NO}」已經存在，請改用其他期別；若要修改既有資料，請關閉本視窗後在列表點該筆的「編輯」。";
                } else {
                    TempData["Error"] = (IsNew ? "新增失敗：" : "更新失敗：") + fullError;
                }
            }

            // 返回編輯頁面
            return RedirectToAction("PAGE6_2_Edit", new { TD_CODE = TD_CODE, SEQ_NO = savedSeqNo.ToString() });
        }

        #region PAGE6_1 - 停復工紀錄 (TDDSTWR)

        // GET: PAGE6_1 - 停復工紀錄列表
        public ActionResult PAGE6_1(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢停復工紀錄
            string sql = @"SELECT TD_CODE, SEQ_NO, STOP_DT, STOP_VRF_DT, STOP_VRF_FN,
                PRE_REWK_DT, REWK_DT, REWK_VRF_DT, REWK_VRF_FN,
                BREW_PRE_EDT, AREW_PRE_EDT, STOP_DAYS, STOP_RESN, RELV_DESC
                FROM TDDSTWR
                WHERE TD_CODE = @TD_CODE
                ORDER BY SEQ_NO";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List1 = dt;

            // 取得下一個 SEQ_NO
            short nextSeqNo = 1;
            if(dt.Rows.Count > 0) {
                var maxSeq = dt.AsEnumerable().Max(r => r.Field<short>("SEQ_NO"));
                nextSeqNo = (short)(maxSeq + 1);
            }
            ViewBag.NextSeqNo = nextSeqNo;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_1", model);
        }

        // POST: PAGE6_1_Add - 新增停復工紀錄 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_1_Add(string TD_CODE, short SEQ_NO,
            string STOP_DT, string STOP_VRF_DT, string STOP_VRF_FN,
            string PRE_REWK_DT, string REWK_DT, string REWK_VRF_DT, string REWK_VRF_FN,
            string BREW_PRE_EDT, string AREW_PRE_EDT, int? STOP_DAYS,
            string STOP_RESN, string RELV_DESC) {
            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    return Json(new { success = false, message = "TD_CODE 不可為空" });
                }

                string insertSql = @"INSERT INTO TDDSTWR
                    (TD_CODE, SEQ_NO, STOP_DT, STOP_VRF_DT, STOP_VRF_FN,
                    PRE_REWK_DT, REWK_DT, REWK_VRF_DT, REWK_VRF_FN,
                    BREW_PRE_EDT, AREW_PRE_EDT, STOP_DAYS, STOP_RESN, RELV_DESC,
                    ADD_DT, UPD_DT, UPD_UID)
                    VALUES (@TD_CODE, @SEQ_NO, @STOP_DT, @STOP_VRF_DT, @STOP_VRF_FN,
                    @PRE_REWK_DT, @REWK_DT, @REWK_VRF_DT, @REWK_VRF_FN,
                    @BREW_PRE_EDT, @AREW_PRE_EDT, @STOP_DAYS, @STOP_RESN, @RELV_DESC,
                    @ADD_DT, @UPD_DT, @UPD_UID)";

                SqlCommand cmd = new SqlCommand(insertSql);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@SEQ_NO", SEQ_NO);
                cmd.Parameters.AddWithValue("@STOP_DT", ParseDateTimeOrNull(STOP_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@STOP_VRF_DT", ParseDateTimeOrNull(STOP_VRF_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@STOP_VRF_FN", (object)STOP_VRF_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PRE_REWK_DT", ParseDateTimeOrNull(PRE_REWK_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@REWK_DT", ParseDateTimeOrNull(REWK_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@REWK_VRF_DT", ParseDateTimeOrNull(REWK_VRF_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@REWK_VRF_FN", (object)REWK_VRF_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@BREW_PRE_EDT", ParseDateTimeOrNull(BREW_PRE_EDT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@AREW_PRE_EDT", ParseDateTimeOrNull(AREW_PRE_EDT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@STOP_DAYS", (object)STOP_DAYS ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@STOP_RESN", (object)STOP_RESN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@RELV_DESC", (object)RELV_DESC ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ADD_DT", DateTime.Now);
                cmd.Parameters.AddWithValue("@UPD_DT", DateTime.Now);
                cmd.Parameters.AddWithValue("@UPD_UID", WebSiteHelper.getUserId() ?? "");

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "新增成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_1_Update - 更新停復工紀錄 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_1_Update(string TD_CODE, short SEQ_NO,
            string STOP_DT, string STOP_VRF_DT, string STOP_VRF_FN,
            string PRE_REWK_DT, string REWK_DT, string REWK_VRF_DT, string REWK_VRF_FN,
            string BREW_PRE_EDT, string AREW_PRE_EDT, int? STOP_DAYS,
            string STOP_RESN, string RELV_DESC) {
            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    return Json(new { success = false, message = "TD_CODE 不可為空" });
                }

                string updateSql = @"UPDATE TDDSTWR SET
                    STOP_DT = @STOP_DT, STOP_VRF_DT = @STOP_VRF_DT, STOP_VRF_FN = @STOP_VRF_FN,
                    PRE_REWK_DT = @PRE_REWK_DT, REWK_DT = @REWK_DT,
                    REWK_VRF_DT = @REWK_VRF_DT, REWK_VRF_FN = @REWK_VRF_FN,
                    BREW_PRE_EDT = @BREW_PRE_EDT, AREW_PRE_EDT = @AREW_PRE_EDT,
                    STOP_DAYS = @STOP_DAYS, STOP_RESN = @STOP_RESN, RELV_DESC = @RELV_DESC,
                    UPD_DT = @UPD_DT, UPD_UID = @UPD_UID
                    WHERE TD_CODE = @TD_CODE AND SEQ_NO = @SEQ_NO";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@SEQ_NO", SEQ_NO);
                cmd.Parameters.AddWithValue("@STOP_DT", ParseDateTimeOrNull(STOP_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@STOP_VRF_DT", ParseDateTimeOrNull(STOP_VRF_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@STOP_VRF_FN", (object)STOP_VRF_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PRE_REWK_DT", ParseDateTimeOrNull(PRE_REWK_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@REWK_DT", ParseDateTimeOrNull(REWK_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@REWK_VRF_DT", ParseDateTimeOrNull(REWK_VRF_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@REWK_VRF_FN", (object)REWK_VRF_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@BREW_PRE_EDT", ParseDateTimeOrNull(BREW_PRE_EDT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@AREW_PRE_EDT", ParseDateTimeOrNull(AREW_PRE_EDT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@STOP_DAYS", (object)STOP_DAYS ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@STOP_RESN", (object)STOP_RESN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@RELV_DESC", (object)RELV_DESC ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@UPD_DT", DateTime.Now);
                cmd.Parameters.AddWithValue("@UPD_UID", WebSiteHelper.getUserId() ?? "");

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_1_Delete - 刪除停復工紀錄 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_1_Delete(string TD_CODE, short SEQ_NO) {
            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    return Json(new { success = false, message = "TD_CODE 不可為空" });
                }

                string deleteSql = @"DELETE FROM TDDSTWR WHERE TD_CODE = @TD_CODE AND SEQ_NO = @SEQ_NO";

                SqlCommand cmd = new SqlCommand(deleteSql);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@SEQ_NO", SEQ_NO);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        // 日期解析輔助方法 (支援民國年格式 yymmdd 如 1140101)
        private DateTime? ParseDateTimeOrNull(string dateStr) {
            if(string.IsNullOrEmpty(dateStr)) {
                return null;
            }

            // 移除可能的空白
            dateStr = dateStr.Trim();

            // 嘗試解析民國年格式 (7碼: 1140101 或 6碼: 140101)
            if(dateStr.Length >= 6 && dateStr.Length <= 7) {
                try {
                    int year, month, day;
                    if(dateStr.Length == 7) {
                        // 1140101 格式
                        year = int.Parse(dateStr.Substring(0, 3)) + 1911;
                        month = int.Parse(dateStr.Substring(3, 2));
                        day = int.Parse(dateStr.Substring(5, 2));
                    } else {
                        // 140101 格式
                        year = int.Parse(dateStr.Substring(0, 2)) + 1911;
                        month = int.Parse(dateStr.Substring(2, 2));
                        day = int.Parse(dateStr.Substring(4, 2));
                    }
                    return new DateTime(year, month, day);
                } catch {
                    // 解析失敗，繼續嘗試其他格式
                }
            }

            // 嘗試標準日期格式
            DateTime result;
            if(DateTime.TryParse(dateStr, out result)) {
                return result;
            }
            return null;
        }

        #endregion

        #region PAGE6_3 - 驗收紀錄 (TDDACPR)

        // GET: PAGE6_3 - 驗收紀錄列表
        public ActionResult PAGE6_3(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢驗收紀錄 (使用 View 取得 REC_TP_DESC)
            string sql = @"SELECT TD_CODE, REC_TP, REC_TP_DESC, SEQ_NO, ACPCHK_DT, ACP_USR, JON_USR, CHK_USR,
                ACPPASS_DT, ACPPASS_FN, ASK1_DT, ASK1_FN, ASK2_DT, ASK2_FN,
                VRF1_DT, VRF1_FN, VRF2_DT, VRF2_FN, ACP_NOTE, ACP_RESULT, UNOK_RESN, OTHE_RESN
                FROM TDV_TDDACPR
                WHERE TD_CODE = @TD_CODE
                ORDER BY REC_TP, SEQ_NO";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List1 = dt;

            // 取得 REC_TP 類別選項 (SMCCODE TD:05)
            string sqlRecTp = @"SELECT CATY_COD, CODE_DESC FROM SMCCODE
                WHERE SYS_ID = 'TD' AND CATY_ID = '05' ORDER BY CATY_COD";
            DataTable dtRecTp = BASE_DAO.SqlDao.CMETMSQuery4List(new SqlCommand(sqlRecTp));
            ViewBag.RecTpList = dtRecTp;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_3", model);
        }

        // POST: PAGE6_3_Add - 新增驗收紀錄 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_3_Add(string TD_CODE, string REC_TP, int SEQ_NO,
            string ACPCHK_DT, string ACP_USR, string JON_USR, string CHK_USR,
            string ACPPASS_DT, string ACPPASS_FN, string ASK1_DT, string ASK1_FN,
            string ASK2_DT, string ASK2_FN, string VRF1_DT, string VRF1_FN,
            string VRF2_DT, string VRF2_FN, string ACP_NOTE, string ACP_RESULT,
            string UNOK_RESN, string OTHE_RESN) {
            try {
                if(string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(REC_TP)) {
                    return Json(new { success = false, message = "TD_CODE 和 REC_TP 不可為空" });
                }

                string insertSql = @"INSERT INTO TDDACPR
                    (TD_CODE, REC_TP, SEQ_NO, ACPCHK_DT, ACP_USR, JON_USR, CHK_USR,
                    ACPPASS_DT, ACPPASS_FN, ASK1_DT, ASK1_FN, ASK2_DT, ASK2_FN,
                    VRF1_DT, VRF1_FN, VRF2_DT, VRF2_FN, ACP_NOTE, ACP_RESULT,
                    UNOK_RESN, OTHE_RESN, ADD_DT, UPD_DT, UPD_UID)
                    VALUES (@TD_CODE, @REC_TP, @SEQ_NO, @ACPCHK_DT, @ACP_USR, @JON_USR, @CHK_USR,
                    @ACPPASS_DT, @ACPPASS_FN, @ASK1_DT, @ASK1_FN, @ASK2_DT, @ASK2_FN,
                    @VRF1_DT, @VRF1_FN, @VRF2_DT, @VRF2_FN, @ACP_NOTE, @ACP_RESULT,
                    @UNOK_RESN, @OTHE_RESN, @ADD_DT, @UPD_DT, @UPD_UID)";

                // [防呆校驗] 僅檢查格式，不改變字串型態（民國年如 1150706）
                if (!string.IsNullOrEmpty(ACPCHK_DT) && ParseDateTimeOrNull(ACPCHK_DT) == null)
                {
                    return Json(new { success = false, message = "停工檢查日期格式錯誤（請輸入民國年如 1150706）" });
                }
                if (!string.IsNullOrEmpty(ACPPASS_DT) && ParseDateTimeOrNull(ACPPASS_DT) == null)
                {
                    return Json(new { success = false, message = "核准日期格式錯誤（請輸入民國年如 1150706）" });
                }

                SqlCommand cmd = new SqlCommand(insertSql);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@REC_TP", REC_TP);
                cmd.Parameters.AddWithValue("@SEQ_NO", SEQ_NO);
                cmd.Parameters.AddWithValue("@ACPCHK_DT", (object)ACPCHK_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ACP_USR", (object)ACP_USR ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@JON_USR", (object)JON_USR ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CHK_USR", (object)CHK_USR ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ACPPASS_DT", (object)ACPPASS_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ACPPASS_FN", (object)ACPPASS_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ASK1_DT", (object)ASK1_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ASK1_FN", (object)ASK1_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ASK2_DT", (object)ASK2_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ASK2_FN", (object)ASK2_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@VRF1_DT", (object)VRF1_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@VRF1_FN", (object)VRF1_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@VRF2_DT", (object)VRF2_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@VRF2_FN", (object)VRF2_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ACP_NOTE", (object)ACP_NOTE ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ACP_RESULT", (object)ACP_RESULT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@UNOK_RESN", (object)UNOK_RESN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@OTHE_RESN", (object)OTHE_RESN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ADD_DT", DateTime.Now);
                cmd.Parameters.AddWithValue("@UPD_DT", DateTime.Now);
                cmd.Parameters.AddWithValue("@UPD_UID", WebSiteHelper.getUserId() ?? "");

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "新增成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_3_Update - 更新驗收紀錄 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_3_Update(string TD_CODE, string REC_TP, int SEQ_NO,
            string ACPCHK_DT, string ACP_USR, string JON_USR, string CHK_USR,
            string ACPPASS_DT, string ACPPASS_FN, string ASK1_DT, string ASK1_FN,
            string ASK2_DT, string ASK2_FN, string VRF1_DT, string VRF1_FN,
            string VRF2_DT, string VRF2_FN, string ACP_NOTE, string ACP_RESULT,
            string UNOK_RESN, string OTHE_RESN) {
            try {
                if(string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(REC_TP)) {
                    return Json(new { success = false, message = "TD_CODE 和 REC_TP 不可為空" });
                }

                string updateSql = @"UPDATE TDDACPR SET
                    ACPCHK_DT = @ACPCHK_DT, ACP_USR = @ACP_USR, JON_USR = @JON_USR, CHK_USR = @CHK_USR,
                    ACPPASS_DT = @ACPPASS_DT, ACPPASS_FN = @ACPPASS_FN,
                    ASK1_DT = @ASK1_DT, ASK1_FN = @ASK1_FN, ASK2_DT = @ASK2_DT, ASK2_FN = @ASK2_FN,
                    VRF1_DT = @VRF1_DT, VRF1_FN = @VRF1_FN, VRF2_DT = @VRF2_DT, VRF2_FN = @VRF2_FN,
                    ACP_NOTE = @ACP_NOTE, ACP_RESULT = @ACP_RESULT,
                    UNOK_RESN = @UNOK_RESN, OTHE_RESN = @OTHE_RESN,
                    UPD_DT = @UPD_DT, UPD_UID = @UPD_UID
                    WHERE TD_CODE = @TD_CODE AND REC_TP = @REC_TP AND SEQ_NO = @SEQ_NO";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@REC_TP", REC_TP);
                cmd.Parameters.AddWithValue("@SEQ_NO", SEQ_NO);
                cmd.Parameters.AddWithValue("@ACPCHK_DT", ParseDateTimeOrNull(ACPCHK_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@ACP_USR", (object)ACP_USR ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@JON_USR", (object)JON_USR ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CHK_USR", (object)CHK_USR ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ACPPASS_DT", ParseDateTimeOrNull(ACPPASS_DT) ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@ACPPASS_FN", (object)ACPPASS_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ASK1_DT", (object)ASK1_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ASK1_FN", (object)ASK1_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ASK2_DT", (object)ASK2_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ASK2_FN", (object)ASK2_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@VRF1_DT", (object)VRF1_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@VRF1_FN", (object)VRF1_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@VRF2_DT", (object)VRF2_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@VRF2_FN", (object)VRF2_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ACP_NOTE", (object)ACP_NOTE ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ACP_RESULT", (object)ACP_RESULT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@UNOK_RESN", (object)UNOK_RESN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@OTHE_RESN", (object)OTHE_RESN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@UPD_DT", DateTime.Now);
                cmd.Parameters.AddWithValue("@UPD_UID", WebSiteHelper.getUserId() ?? "");

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_3_Delete - 刪除驗收紀錄 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_3_Delete(string TD_CODE, string REC_TP, int SEQ_NO) {
            try {
                if(string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(REC_TP)) {
                    return Json(new { success = false, message = "TD_CODE 和 REC_TP 不可為空" });
                }

                string deleteSql = @"DELETE FROM TDDACPR WHERE TD_CODE = @TD_CODE AND REC_TP = @REC_TP AND SEQ_NO = @SEQ_NO";

                SqlCommand cmd = new SqlCommand(deleteSql);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@REC_TP", REC_TP);
                cmd.Parameters.AddWithValue("@SEQ_NO", SEQ_NO);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        // 取得指定類別的下一個 SEQ_NO (AJAX)
        [HttpPost]
        public ActionResult PAGE6_3_GetNextSeqNo(string TD_CODE, string REC_TP) {
            try {
                string sql = @"SELECT ISNULL(MAX(SEQ_NO), 0) + 1 AS NextSeqNo
                    FROM TDDACPR WHERE TD_CODE = @TD_CODE AND REC_TP = @REC_TP";
                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@REC_TP", REC_TP);
                DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
                int nextSeqNo = dt.Rows.Count > 0 ? Convert.ToInt32(dt.Rows[0]["NextSeqNo"]) : 1;
                return Json(new { success = true, nextSeqNo = nextSeqNo });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        #endregion

        #region PAGE6_4 - 技師證照 (MECHANIC)

        // GET: PAGE6_4 - 技師證照列表
        public ActionResult PAGE6_4(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢技師證照紀錄
            string sql = @"SELECT ID, TD_CODE, LICENSE, CARDID, NAME, COMPANY_NO, COMPANY_NAME,
                EMP_SDT, EMP_EDT, CASE_TP, LICENSE_PAPER
                FROM MECHANIC
                WHERE TD_CODE = @TD_CODE
                ORDER BY ID";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List1 = dt;

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_4", model);
        }

        // POST: PAGE6_4_Add - 新增技師證照 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_4_Add(string TD_CODE, string LICENSE, string CARDID,
            string NAME, string COMPANY_NO, string COMPANY_NAME,
            string EMP_SDT, string EMP_EDT, string CASE_TP, string LICENSE_PAPER) {
            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    return Json(new { success = false, message = "TD_CODE 不可為空" });
                }
                if(string.IsNullOrEmpty(CARDID)) {
                    return Json(new { success = false, message = "身分證字號不可為空" });
                }

                // 產生新 ID (使用時間戳記)
                string newId = DateTime.Now.ToString("yyyyMMddHHmmssfff");

                string insertSql = @"INSERT INTO MECHANIC
                    (ID, TD_CODE, LICENSE, CARDID, NAME, COMPANY_NO, COMPANY_NAME,
                    EMP_SDT, EMP_EDT, CASE_TP, LICENSE_PAPER)
                    VALUES (@ID, @TD_CODE, @LICENSE, @CARDID, @NAME, @COMPANY_NO, @COMPANY_NAME,
                    @EMP_SDT, @EMP_EDT, @CASE_TP, @LICENSE_PAPER)";

                SqlCommand cmd = new SqlCommand(insertSql);
                cmd.Parameters.AddWithValue("@ID", newId);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@LICENSE", (object)LICENSE ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CARDID", CARDID);
                cmd.Parameters.AddWithValue("@NAME", (object)NAME ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@COMPANY_NO", (object)COMPANY_NO ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@COMPANY_NAME", (object)COMPANY_NAME ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@EMP_SDT", (object)EMP_SDT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@EMP_EDT", (object)EMP_EDT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CASE_TP", (object)CASE_TP ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@LICENSE_PAPER", (object)LICENSE_PAPER ?? DBNull.Value);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "新增成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_4_Update - 更新技師證照 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_4_Update(string ID, string LICENSE, string CARDID,
            string NAME, string COMPANY_NO, string COMPANY_NAME,
            string EMP_SDT, string EMP_EDT, string CASE_TP, string LICENSE_PAPER) {
            try {
                if(string.IsNullOrEmpty(ID)) {
                    return Json(new { success = false, message = "ID 不可為空" });
                }
                if(string.IsNullOrEmpty(CARDID)) {
                    return Json(new { success = false, message = "身分證字號不可為空" });
                }

                string updateSql = @"UPDATE MECHANIC SET
                    LICENSE = @LICENSE, CARDID = @CARDID, NAME = @NAME,
                    COMPANY_NO = @COMPANY_NO, COMPANY_NAME = @COMPANY_NAME,
                    EMP_SDT = @EMP_SDT, EMP_EDT = @EMP_EDT,
                    CASE_TP = @CASE_TP, LICENSE_PAPER = @LICENSE_PAPER
                    WHERE ID = @ID";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.AddWithValue("@ID", ID);
                cmd.Parameters.AddWithValue("@LICENSE", (object)LICENSE ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CARDID", CARDID);
                cmd.Parameters.AddWithValue("@NAME", (object)NAME ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@COMPANY_NO", (object)COMPANY_NO ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@COMPANY_NAME", (object)COMPANY_NAME ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@EMP_SDT", (object)EMP_SDT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@EMP_EDT", (object)EMP_EDT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CASE_TP", (object)CASE_TP ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@LICENSE_PAPER", (object)LICENSE_PAPER ?? DBNull.Value);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_4_Delete - 刪除技師證照 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_4_Delete(string ID) {
            try {
                if(string.IsNullOrEmpty(ID)) {
                    return Json(new { success = false, message = "ID 不可為空" });
                }

                string deleteSql = @"DELETE FROM MECHANIC WHERE ID = @ID";

                SqlCommand cmd = new SqlCommand(deleteSql);
                cmd.Parameters.AddWithValue("@ID", ID);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        #endregion

        #region PAGE6_5 - 專業人員 (TDDPMUSR)

        // GET: PAGE6_5 - 專業人員列表
        public ActionResult PAGE6_5(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢專業人員紀錄
            string sql = @"SELECT UID, TD_CODE, PRO_CATY, LISN_ID, USR_ID, USR_NAME,
                USR_TEL, USR_EMAIL, LISN_DEADLINE, REM_TP, REM_DT, REM_FN
                FROM TDDPMUSR
                WHERE TD_CODE = @TD_CODE
                ORDER BY UID";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List1 = dt;

            // 取得專業類別選項
            ViewBag.ProCatyList = util.GetCMETMSCODES("TD", "03");

            // 取得解除原因選項
            ViewBag.RemTpList = util.GetCMETMSCODES("TD", "10");

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_5", model);
        }

        // POST: PAGE6_5_Add - 新增專業人員 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_5_Add(string TD_CODE, string PRO_CATY, string LISN_ID,
            string USR_ID, string USR_NAME, string USR_TEL, string USR_EMAIL,
            string LISN_DEADLINE, string REM_TP, string REM_DT, string REM_FN) {
            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    return Json(new { success = false, message = "TD_CODE 不可為空" });
                }

                string insertSql = @"INSERT INTO TDDPMUSR
                    (TD_CODE, PRO_CATY, LISN_ID, USR_ID, USR_NAME, USR_TEL, USR_EMAIL,
                    LISN_DEADLINE, REM_TP, REM_DT, REM_FN, ADD_DT)
                    VALUES (@TD_CODE, @PRO_CATY, @LISN_ID, @USR_ID, @USR_NAME, @USR_TEL, @USR_EMAIL,
                    @LISN_DEADLINE, @REM_TP, @REM_DT, @REM_FN, @ADD_DT)";

                SqlCommand cmd = new SqlCommand(insertSql);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@PRO_CATY", (object)PRO_CATY ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@LISN_ID", (object)LISN_ID ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_ID", (object)USR_ID ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_NAME", (object)USR_NAME ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_TEL", (object)USR_TEL ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_EMAIL", (object)USR_EMAIL ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@LISN_DEADLINE", (object)LISN_DEADLINE ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_TP", (object)REM_TP ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_DT", (object)REM_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_FN", (object)REM_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ADD_DT", DateTime.Now);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "新增成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_5_Update - 更新專業人員 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_5_Update(decimal UID, string PRO_CATY, string LISN_ID,
            string USR_ID, string USR_NAME, string USR_TEL, string USR_EMAIL,
            string LISN_DEADLINE, string REM_TP, string REM_DT, string REM_FN) {
            try {
                string updateSql = @"UPDATE TDDPMUSR SET
                    PRO_CATY = @PRO_CATY, LISN_ID = @LISN_ID, USR_ID = @USR_ID,
                    USR_NAME = @USR_NAME, USR_TEL = @USR_TEL, USR_EMAIL = @USR_EMAIL,
                    LISN_DEADLINE = @LISN_DEADLINE, REM_TP = @REM_TP, REM_DT = @REM_DT, REM_FN = @REM_FN
                    WHERE UID = @UID";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.AddWithValue("@UID", UID);
                cmd.Parameters.AddWithValue("@PRO_CATY", (object)PRO_CATY ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@LISN_ID", (object)LISN_ID ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_ID", (object)USR_ID ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_NAME", (object)USR_NAME ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_TEL", (object)USR_TEL ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_EMAIL", (object)USR_EMAIL ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@LISN_DEADLINE", (object)LISN_DEADLINE ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_TP", (object)REM_TP ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_DT", (object)REM_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_FN", (object)REM_FN ?? DBNull.Value);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_5_Delete - 刪除專業人員 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_5_Delete(decimal UID) {
            try {
                string deleteSql = @"DELETE FROM TDDPMUSR WHERE UID = @UID";

                SqlCommand cmd = new SqlCommand(deleteSql);
                cmd.Parameters.AddWithValue("@UID", UID);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        #endregion

        #region PAGE6_8 - 監造廠商人員 (TDDPMUSR)

        // GET: PAGE6_8 - 監造廠商人員列表
        public ActionResult PAGE6_8(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢監造廠商人員紀錄
            string sql = @"SELECT UID, TD_CODE, PRO_CATY, LISN_ID, USR_ID, USR_NAME,
                USR_TEL, REM_TP, REM_DT, REM_FN
                FROM TDDPMUSR
                WHERE TD_CODE = @TD_CODE
                ORDER BY UID";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List1 = dt;

            // 取得服務類型選項
            ViewBag.ProCatyList = util.GetCMETMSCODES("TD", "03");

            // 取得解除原因選項
            ViewBag.RemTpList = util.GetCMETMSCODES("TD", "10");

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE6_8", model);
        }

        // POST: PAGE6_8_Add - 新增監造廠商人員 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_8_Add(string TD_CODE, string PRO_CATY, string LISN_ID,
            string USR_ID, string USR_NAME, string USR_TEL,
            string REM_TP, string REM_DT, string REM_FN) {
            try {
                if(string.IsNullOrEmpty(TD_CODE)) {
                    return Json(new { success = false, message = "TD_CODE 不可為空" });
                }

                string insertSql = @"INSERT INTO TDDPMUSR
                    (TD_CODE, PRO_CATY, LISN_ID, USR_ID, USR_NAME, USR_TEL,
                    REM_TP, REM_DT, REM_FN, ADD_DT)
                    VALUES (@TD_CODE, @PRO_CATY, @LISN_ID, @USR_ID, @USR_NAME, @USR_TEL,
                    @REM_TP, @REM_DT, @REM_FN, @ADD_DT)";

                SqlCommand cmd = new SqlCommand(insertSql);
                cmd.Parameters.AddWithValue("@TD_CODE", TD_CODE);
                cmd.Parameters.AddWithValue("@PRO_CATY", (object)PRO_CATY ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@LISN_ID", (object)LISN_ID ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_ID", (object)USR_ID ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_NAME", (object)USR_NAME ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_TEL", (object)USR_TEL ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_TP", (object)REM_TP ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_DT", (object)REM_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_FN", (object)REM_FN ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ADD_DT", DateTime.Now);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "新增成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_8_Update - 更新監造廠商人員 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_8_Update(decimal UID, string PRO_CATY, string LISN_ID,
            string USR_ID, string USR_NAME, string USR_TEL,
            string REM_TP, string REM_DT, string REM_FN) {
            try {
                string updateSql = @"UPDATE TDDPMUSR SET
                    PRO_CATY = @PRO_CATY, LISN_ID = @LISN_ID, USR_ID = @USR_ID,
                    USR_NAME = @USR_NAME, USR_TEL = @USR_TEL,
                    REM_TP = @REM_TP, REM_DT = @REM_DT, REM_FN = @REM_FN
                    WHERE UID = @UID";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.AddWithValue("@UID", UID);
                cmd.Parameters.AddWithValue("@PRO_CATY", (object)PRO_CATY ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@LISN_ID", (object)LISN_ID ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_ID", (object)USR_ID ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_NAME", (object)USR_NAME ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@USR_TEL", (object)USR_TEL ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_TP", (object)REM_TP ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_DT", (object)REM_DT ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@REM_FN", (object)REM_FN ?? DBNull.Value);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        // POST: PAGE6_8_Delete - 刪除監造廠商人員 (AJAX)
        [HttpPost]
        public ActionResult PAGE6_8_Delete(decimal UID) {
            try {
                string deleteSql = @"DELETE FROM TDDPMUSR WHERE UID = @UID";

                SqlCommand cmd = new SqlCommand(deleteSql);
                cmd.Parameters.AddWithValue("@UID", UID);

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        #endregion

    }
}