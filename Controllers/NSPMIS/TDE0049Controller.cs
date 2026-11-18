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
using static System.Runtime.CompilerServices.RuntimeHelpers;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS {


    public class TDE0049Controller : Controller {

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


        // GET: PAGE9 - 工程填報(污)
        public ActionResult PAGE9(string TD_CODE) {
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
            return View("PAGE9", model);
        }

        // GET: PAGE9_1 - 管線工程（污水）
        public ActionResult PAGE9_1(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢 TDMSWNOTE 資料
            string sql = @"SELECT * FROM CMETMS.dbo.TDMSWNOTE WHERE TD_CODE = @TD_CODE";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            DataTable dt = null;
            try {
                dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
            } catch(Exception ex) {
                ViewBag.TableError = "查詢 TDMSWNOTE 失敗：" + ex.Message;
            }

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.IsNew = (dt == null || dt.Rows.Count == 0);

            if(dt != null && dt.Rows.Count > 0) {
                ViewBag.Data = dt.Rows[0];
            } else {
                ViewBag.Data = null;
            }

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE9_1", model);
        }

        // POST: PAGE9_1 - 儲存管線工程資料（污水）
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE9_1_Save(string TD_CODE,
         string P_S_C_1, string P_E_1, string P_P_I_1,
         bool IsNew = false) {

            try {
                if(IsNew) {
                    // 新增模式
                    string insertSql = @"INSERT INTO CMETMS.dbo.TDMSWNOTE 
 (TD_CODE, P_S_C_1, P_E_1, P_P_I_1)
  VALUES (@TD_CODE, @P_S_C_1, @P_E_1, @P_P_I_1)";

                    SqlCommand cmd = new SqlCommand(insertSql);
                    cmd.Parameters.Add(new SqlParameter("TD_CODE", (object)TD_CODE ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_S_C_1", (object)P_S_C_1 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_E_1", (object)P_E_1 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_P_I_1", (object)P_P_I_1 ?? DBNull.Value));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                    TempData["Success"] = "資料新增成功";
                } else {
                    // 更新模式
                    string updateSql = @"UPDATE CMETMS.dbo.TDMSWNOTE SET 
 P_S_C_1 = @P_S_C_1,
           P_E_1 = @P_E_1,
    P_P_I_1 = @P_P_I_1
   WHERE TD_CODE = @TD_CODE";

                    SqlCommand cmd = new SqlCommand(updateSql);
                    cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                    cmd.Parameters.Add(new SqlParameter("P_S_C_1", (object)P_S_C_1 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_E_1", (object)P_E_1 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_P_I_1", (object)P_P_I_1 ?? DBNull.Value));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                    TempData["Success"] = "資料更新成功";
                }
            } catch(Exception ex) {
                TempData["Error"] = (IsNew ? "資料新增失敗：" : "資料更新失敗：") + ex.Message;
            }

            return RedirectToAction("PAGE9_1", new { TD_CODE = TD_CODE });
        }

        // GET: PAGE9_2 - 污水處理廠工程
        public ActionResult PAGE9_2(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢 TDMSWNOTE 資料
            string sql = @"SELECT * FROM CMETMS.dbo.TDMSWNOTE WHERE TD_CODE = @TD_CODE";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            DataTable dt = null;
            try {
                dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
            } catch(Exception ex) {
                ViewBag.TableError = "查詢 TDMSWNOTE 失敗：" + ex.Message;
            }

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.IsNew = (dt == null || dt.Rows.Count == 0);

            if(dt != null && dt.Rows.Count > 0) {
                ViewBag.Data = dt.Rows[0];
            } else {
                ViewBag.Data = null;
            }

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE9_2", model);
        }

        // POST: PAGE9_2 - 儲存污水處理廠工程資料
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE9_2_Save(string TD_CODE,
    string P_S_C_2, string P_E_2, string P_P_I_2,
      bool IsNew = false) {

            try {
                if(IsNew) {
                    // 新增模式
                    string insertSql = @"INSERT INTO CMETMS.dbo.TDMSWNOTE 
          (TD_CODE, P_S_C_2, P_E_2, P_P_I_2)
      VALUES (@TD_CODE, @P_S_C_2, @P_E_2, @P_P_I_2)";

                    SqlCommand cmd = new SqlCommand(insertSql);
                    cmd.Parameters.Add(new SqlParameter("TD_CODE", (object)TD_CODE ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_S_C_2", (object)P_S_C_2 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_E_2", (object)P_E_2 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_P_I_2", (object)P_P_I_2 ?? DBNull.Value));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                    TempData["Success"] = "資料新增成功";
                } else {
                    // 更新模式
                    string updateSql = @"UPDATE CMETMS.dbo.TDMSWNOTE SET 
    P_S_C_2 = @P_S_C_2,
      P_E_2 = @P_E_2,
          P_P_I_2 = @P_P_I_2
    WHERE TD_CODE = @TD_CODE";

                    SqlCommand cmd = new SqlCommand(updateSql);
                    cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                    cmd.Parameters.Add(new SqlParameter("P_S_C_2", (object)P_S_C_2 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_E_2", (object)P_E_2 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_P_I_2", (object)P_P_I_2 ?? DBNull.Value));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                    TempData["Success"] = "資料更新成功";
                }
            } catch(Exception ex) {
                TempData["Error"] = (IsNew ? "資料新增失敗：" : "資料更新失敗：") + ex.Message;
            }

            return RedirectToAction("PAGE9_2", new { TD_CODE = TD_CODE });
        }

        // GET: PAGE9_3 - 抽水站工程（污水）
        public ActionResult PAGE9_3(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 查詢 TDMSWNOTE 資料
            string sql = @"SELECT * FROM CMETMS.dbo.TDMSWNOTE WHERE TD_CODE = @TD_CODE";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            DataTable dt = null;
            try {
                dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
            } catch(Exception ex) {
                ViewBag.TableError = "查詢 TDMSWNOTE 失敗：" + ex.Message;
            }

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.IsNew = (dt == null || dt.Rows.Count == 0);

            if(dt != null && dt.Rows.Count > 0) {
                ViewBag.Data = dt.Rows[0];
            } else {
                ViewBag.Data = null;
            }

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE9_3", model);
        }

        // POST: PAGE9_3 - 儲存抽水站工程資料（污水）
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE9_3_Save(string TD_CODE,
            string P_S_C_3, string P_E_3, string P_P_I_3,
            bool IsNew = false) {

            try {
                if(IsNew) {
                    // 新增模式
                    string insertSql = @"INSERT INTO CMETMS.dbo.TDMSWNOTE 
     (TD_CODE, P_S_C_3, P_E_3, P_P_I_3)
               VALUES (@TD_CODE, @P_S_C_3, @P_E_3, @P_P_I_3)";

                    SqlCommand cmd = new SqlCommand(insertSql);
                    cmd.Parameters.Add(new SqlParameter("TD_CODE", (object)TD_CODE ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_S_C_3", (object)P_S_C_3 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_E_3", (object)P_E_3 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_P_I_3", (object)P_P_I_3 ?? DBNull.Value));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                    TempData["Success"] = "資料新增成功";
                } else {
                    // 更新模式
                    string updateSql = @"UPDATE CMETMS.dbo.TDMSWNOTE SET 
           P_S_C_3 = @P_S_C_3,
          P_E_3 = @P_E_3,
      P_P_I_3 = @P_P_I_3
 WHERE TD_CODE = @TD_CODE";

                    SqlCommand cmd = new SqlCommand(updateSql);
                    cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                    cmd.Parameters.Add(new SqlParameter("P_S_C_3", (object)P_S_C_3 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_E_3", (object)P_E_3 ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("P_P_I_3", (object)P_P_I_3 ?? DBNull.Value));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                    TempData["Success"] = "資料更新成功";
                }
            } catch(Exception ex) {
                TempData["Error"] = (IsNew ? "資料新增失敗：" : "資料更新失敗：") + ex.Message;
            }

            return RedirectToAction("PAGE9_3", new { TD_CODE = TD_CODE });
        }


    }

} 

