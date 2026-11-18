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
    public class TDE0045Controller : Controller {

        private readonly CMETMSEntities db = new CMETMSEntities();

        // GET: PAGE5
        public ActionResult PAGE5(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);




            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform,
                TD_CODE = TD_CODE
            };

            GetSelectList();
            return View("PAGE5", model);
        }

        public ActionResult PAGE5_1(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            // 查詢 IT_TYPE='S' 的單筆資料（計畫期程摘要）
            string sqlSummary = @"SELECT * FROM TDDSDWB WHERE TD_CODE=@TD_CODE AND IT_TYPE='S'";
            SqlCommand cmdSummary = new SqlCommand(sqlSummary);
            cmdSummary.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dtSummary = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSummary);
            ViewBag.SummaryData = (dtSummary != null && dtSummary.Rows.Count > 0) ? dtSummary.Rows[0] : null;

            // 查詢 IT_TYPE!='S' 的列表資料（工作控制點）
            string sqlList = @"SELECT * FROM TDDSDWB WHERE TD_CODE=@TD_CODE AND IT_TYPE!='S' ORDER BY GP_ORDER";
            SqlCommand cmdList = new SqlCommand(sqlList);
            cmdList.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dtList = BASE_DAO.SqlDao.CMETMSQuery4List(cmdList);
            ViewBag.ListData = dtList;

            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform,
                TD_CODE = TD_CODE
            };

            GetSelectList();
            return View("PAGE5_1", model);
        }

        // POST: PAGE5_1_SaveSummary - 儲存計畫期程摘要(IT_TYPE='S')
        [HttpPost]
        public ActionResult PAGE5_1_SaveSummary(string TD_CODE, string PRE_SDT, string PRE_EDT,
            string REL_SDT, string REL_EDT, string NOTE) {
            try {
                // 檢查是否已有資料
                string checkSql = @"SELECT COUNT(*) FROM TDDSDWB WHERE TD_CODE=@TD_CODE AND IT_TYPE='S'";
                SqlCommand checkCmd = new SqlCommand(checkSql);
                checkCmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                DataTable dtCheck = BASE_DAO.SqlDao.CMETMSQuery4List(checkCmd);
                int count = Convert.ToInt32(dtCheck.Rows[0][0]);

                if(count > 0) {
                    // 更新
                    string updateSql = @"UPDATE TDDSDWB SET 
       PRE_SDT=@PRE_SDT, PRE_EDT=@PRE_EDT, 
           REL_SDT=@REL_SDT, REL_EDT=@REL_EDT, 
     NOTE=@NOTE, UPD_DT=@UPD_DT, UPD_UID=@UPD_UID 
 WHERE TD_CODE=@TD_CODE AND IT_TYPE='S'";
                    SqlCommand updateCmd = new SqlCommand(updateSql);
                    updateCmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                    updateCmd.Parameters.Add(new SqlParameter("PRE_SDT", (object)PRE_SDT ?? DBNull.Value));
                    updateCmd.Parameters.Add(new SqlParameter("PRE_EDT", (object)PRE_EDT ?? DBNull.Value));
                    updateCmd.Parameters.Add(new SqlParameter("REL_SDT", (object)REL_SDT ?? DBNull.Value));
                    updateCmd.Parameters.Add(new SqlParameter("REL_EDT", (object)REL_EDT ?? DBNull.Value));
                    updateCmd.Parameters.Add(new SqlParameter("NOTE", (object)NOTE ?? DBNull.Value));
                    updateCmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                    updateCmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));
                    BASE_DAO.SqlDao.CMETMSExecuteSql(updateCmd);
                } else {
                    // 新增
                    string insertSql = @"INSERT INTO TDDSDWB 
    (TD_CODE, IT_CODE, IT_TYPE, PRE_SDT, PRE_EDT, REL_SDT, REL_EDT, NOTE, ADD_DT, UPD_DT, UPD_UID) 
              VALUES (@TD_CODE, '01', 'S', @PRE_SDT, @PRE_EDT, @REL_SDT, @REL_EDT, @NOTE, @ADD_DT, @UPD_DT, @UPD_UID)";
                    SqlCommand insertCmd = new SqlCommand(insertSql);
                    insertCmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                    //insertCmd.Parameters.Add(new SqlParameter("IT_CODE", "01"));
                    insertCmd.Parameters.Add(new SqlParameter("PRE_SDT", (object)PRE_SDT ?? DBNull.Value));
                    insertCmd.Parameters.Add(new SqlParameter("PRE_EDT", (object)PRE_EDT ?? DBNull.Value));
                    insertCmd.Parameters.Add(new SqlParameter("REL_SDT", (object)REL_SDT ?? DBNull.Value));
                    insertCmd.Parameters.Add(new SqlParameter("REL_EDT", (object)REL_EDT ?? DBNull.Value));
                    insertCmd.Parameters.Add(new SqlParameter("NOTE", (object)NOTE ?? DBNull.Value));
                    insertCmd.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now));
                    insertCmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                    insertCmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));
                    BASE_DAO.SqlDao.CMETMSExecuteSql(insertCmd);
                }

                return Json(new { success = true, message = "儲存成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // POST: PAGE5_1_SaveItem - 儲存工作控制點(IT_TYPE!='S')
        [HttpPost]
        public ActionResult PAGE5_1_SaveItem(string IT_CODE, string TD_CODE, string IT_NAME,
  string PRE_EDT, string REL_EDT, int? GP_ORDER, decimal? IT_WGT) {
            try {
                if(string.IsNullOrEmpty(IT_CODE)) {
                    return Json(new { success = false, message = "IT_CODE 不可為空" });
                }

                string updateSql = @"UPDATE TDDSDWB SET 
    IT_NAME=@IT_NAME, PRE_EDT=@PRE_EDT, REL_EDT=@REL_EDT, 
  GP_ORDER=@GP_ORDER, IT_WGT=@IT_WGT, UPD_DT=@UPD_DT, UPD_UID=@UPD_UID 
            WHERE TD_CODE=@TD_CODE AND IT_CODE=@IT_CODE AND IT_TYPE!='S'";
                SqlCommand updateCmd = new SqlCommand(updateSql);
                updateCmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                updateCmd.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));
                updateCmd.Parameters.Add(new SqlParameter("IT_NAME", (object)IT_NAME ?? DBNull.Value));
                updateCmd.Parameters.Add(new SqlParameter("PRE_EDT", (object)PRE_EDT ?? DBNull.Value));
                updateCmd.Parameters.Add(new SqlParameter("REL_EDT", (object)REL_EDT ?? DBNull.Value));
                updateCmd.Parameters.Add(new SqlParameter("GP_ORDER", (object)GP_ORDER ?? DBNull.Value));
                updateCmd.Parameters.Add(new SqlParameter("IT_WGT", (object)IT_WGT ?? DBNull.Value));
                updateCmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                updateCmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));
                BASE_DAO.SqlDao.CMETMSExecuteSql(updateCmd);

                return Json(new { success = true, message = "儲存成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // 3.1.4rev：工作控制點的「分月進度」原本只有 alert 佔位（業主回報「尚未提供功能」）。
        // 舊系統 App_Data\CMETMS\TDE0020B041.Xml 定義其為 年度(SYR) + 一月~十二月，
        // 實體資料表為 CMETMS.dbo.TDDWBMS（鍵：TD_CODE + IT_CODE + SYR，月份欄位 M01~M12，無底線）。

        // GET: 取得某工作控制點的分月進度
        public ActionResult PAGE5_1_GetMonth(string TD_CODE, string IT_CODE) {
            try {
                if(string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(IT_CODE)) {
                    return Json(new { success = false, message = "TD_CODE / IT_CODE 不可為空" }, JsonRequestBehavior.AllowGet);
                }

                string sql = @"SELECT SYR, M01, M02, M03, M04, M05, M06,
                                      M07, M08, M09, M10, M11, M12, NOTE
                               FROM TDDWBMS WHERE TD_CODE=@TD_CODE AND IT_CODE=@IT_CODE ORDER BY SYR";
                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmd.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));
                DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

                var rows = new List<object>();
                if(dt != null) {
                    foreach(DataRow r in dt.Rows) {
                        var months = new List<object>();
                        for(int m = 1; m <= 12; m++) {
                            object v = r["M" + m.ToString("00")];
                            months.Add(v == DBNull.Value ? null : (object)Convert.ToDecimal(v));
                        }
                        rows.Add(new {
                            SYR = r["SYR"] == DBNull.Value ? "" : r["SYR"].ToString().Trim(),
                            NOTE = r["NOTE"] == DBNull.Value ? "" : r["NOTE"].ToString(),
                            M = months
                        });
                    }
                }
                return Json(new { success = true, rows = rows }, JsonRequestBehavior.AllowGet);
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        // POST: 儲存某工作控制點某年度的分月進度（無資料則新增）
        [HttpPost]
        public ActionResult PAGE5_1_SaveMonth(string TD_CODE, string IT_CODE, string SYR, decimal?[] M, string NOTE) {
            try {
                if(string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(IT_CODE) || string.IsNullOrEmpty(SYR)) {
                    return Json(new { success = false, message = "工程編號 / 控制點 / 年度不可為空" });
                }
                if(M == null || M.Length != 12) {
                    return Json(new { success = false, message = "月份資料需為 12 個月" });
                }

                string checkSql = @"SELECT COUNT(*) FROM TDDWBMS WHERE TD_CODE=@TD_CODE AND IT_CODE=@IT_CODE AND SYR=@SYR";
                SqlCommand checkCmd = new SqlCommand(checkSql);
                checkCmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                checkCmd.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));
                checkCmd.Parameters.Add(new SqlParameter("SYR", SYR));
                DataTable dtCheck = BASE_DAO.SqlDao.CMETMSQuery4List(checkCmd);
                bool exists = Convert.ToInt32(dtCheck.Rows[0][0]) > 0;

                string sql = exists
                    ? @"UPDATE TDDWBMS SET M01=@M01,M02=@M02,M03=@M03,M04=@M04,M05=@M05,M06=@M06,
                                           M07=@M07,M08=@M08,M09=@M09,M10=@M10,M11=@M11,M12=@M12,
                                           NOTE=@NOTE, UPD_DT=@UPD_DT, UPD_UID=@UPD_UID
                        WHERE TD_CODE=@TD_CODE AND IT_CODE=@IT_CODE AND SYR=@SYR"
                    : @"INSERT INTO TDDWBMS (TD_CODE, IT_CODE, SYR,
                            M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,
                            NOTE, ADD_DT, UPD_DT, UPD_UID)
                        VALUES (@TD_CODE, @IT_CODE, @SYR,
                            @M01,@M02,@M03,@M04,@M05,@M06,@M07,@M08,@M09,@M10,@M11,@M12,
                            @NOTE, @ADD_DT, @UPD_DT, @UPD_UID)";

                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmd.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));
                cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                for(int i = 0; i < 12; i++) {
                    cmd.Parameters.Add(new SqlParameter("M" + (i + 1).ToString("00"), (object)M[i] ?? DBNull.Value));
                }
                cmd.Parameters.Add(new SqlParameter("NOTE", (object)NOTE ?? DBNull.Value));
                if(!exists) { cmd.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now)); }
                cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                return Json(new { success = true, message = exists ? "分月進度更新成功" : "分月進度新增成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        public ActionResult PAGE5_2(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            // 查詢 CTDMNDS 列表資料
            string sql = @"SELECT * FROM CMEPMS.dbo.CTDMNDS WHERE TD_CODE=@TD_CODE";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dtList = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
            ViewBag.ListData = dtList;

            // 查詢 TDDMSCD 年度列表（用於下拉選單）
            string sqlSyr = @"SELECT DISTINCT SYR FROM CMETMS.dbo.TDDMSCD WHERE TD_CODE=@TD_CODE ORDER BY SYR DESC";
            SqlCommand cmdSyr = new SqlCommand(sqlSyr);
            cmdSyr.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dtSyr = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSyr);
            ViewBag.SyrList = dtSyr;

            // 預設查詢第一個年度的 TDDMSCD 資料
            if(dtSyr != null && dtSyr.Rows.Count > 0) {
                string defaultSyr = dtSyr.Rows[0]["SYR"].ToString();
                string sqlMscd = @"SELECT * FROM CMETMS.dbo.TDDMSCD WHERE TD_CODE=@TD_CODE AND SYR=@SYR ORDER BY SMN";
                SqlCommand cmdMscd = new SqlCommand(sqlMscd);
                cmdMscd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmdMscd.Parameters.Add(new SqlParameter("SYR", defaultSyr));
                DataTable dtMscd = BASE_DAO.SqlDao.CMETMSQuery4List(cmdMscd);
                ViewBag.MscdData = dtMscd;
                ViewBag.SelectedSyr = defaultSyr;
            }

            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform,
                TD_CODE = TD_CODE
            };

            GetSelectList();
            return View("PAGE5_2", model);
        }

        // AJAX: 根據年度查詢 TDDMSCD 資料
        [HttpPost]
        public ActionResult PAGE5_2_GetMscdBySyr(string TD_CODE, string SYR) {
            try {
                string sql = @"SELECT * FROM CMETMS.dbo.TDDMSCD WHERE TD_CODE=@TD_CODE AND SYR=@SYR ORDER BY SMN";
                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

                var result = new List<Dictionary<string, object>>();
                foreach(DataRow row in dt.Rows) {
                    var dict = new Dictionary<string, object>();
                    foreach(DataColumn col in dt.Columns) {
                        dict[col.ColumnName] = row[col] == DBNull.Value ? "" : row[col];
                    }
                    result.Add(dict);
                }
                return Json(new { success = true, data = result });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // PAGE5_2_Edit - 編輯分月進度 (GET)
        public ActionResult PAGE5_2_Edit(string ID) {
            if(string.IsNullOrEmpty(ID)) {
                return Content("ID 不可為空");
            }

            string sql = @"SELECT * FROM CMETMS.dbo.TDDMSCD WHERE ID = @ID";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("ID", ID));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.ID = ID;
            ViewBag.IsNew = false;

            if(dt != null && dt.Rows.Count > 0) {
                ViewBag.Data = dt.Rows[0];
                ViewBag.TD_CODE = dt.Rows[0]["TD_CODE"]?.ToString();
            } else {
                ViewBag.Data = null;
            }

            return View("PAGE5_2_Edit");
        }

        // PAGE5_2_Edit - 編輯分月進度 (POST)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE5_2_Edit_Save(string ID, string TD_CODE,
            decimal? ECC_UPAY_FEE, decimal? ORG_SUM_PRE_SCHD,
            decimal? PRE_SCHD, decimal? REL_SCHD, decimal? PRE_COST, decimal? REL_COST,
            decimal? CT_PRE_COST, decimal? RV_PRE_COST,
            decimal? L_PLN_COST, decimal? L_REL_COST, decimal? L_UNPAY_COST, decimal? SAVE_FEE,
            decimal? PRE_CON_LEN, decimal? REL_CON_LEN, decimal? PRE_CON_FLOOD, decimal? REL_CON_FLOOD,
            string EXECT_DESC, string DELAY_RESN, string RESLV_MTHD, string RESLV_LMT_DT) {

            try {
                string updateSql = @"UPDATE CMETMS.dbo.TDDMSCD SET
                    ECC_UPAY_FEE = @ECC_UPAY_FEE,
                    ORG_SUM_PRE_SCHD = @ORG_SUM_PRE_SCHD,
                    PRE_SCHD = @PRE_SCHD,
                    REL_SCHD = @REL_SCHD,
                    PRE_COST = @PRE_COST,
                    REL_COST = @REL_COST,
                    CT_PRE_COST = @CT_PRE_COST,
                    RV_PRE_COST = @RV_PRE_COST,
                    L_PLN_COST = @L_PLN_COST,
                    L_REL_COST = @L_REL_COST,
                    L_UNPAY_COST = @L_UNPAY_COST,
                    SAVE_FEE = @SAVE_FEE,
                    PRE_CON_LEN = @PRE_CON_LEN,
                    REL_CON_LEN = @REL_CON_LEN,
                    PRE_CON_FLOOD = @PRE_CON_FLOOD,
                    REL_CON_FLOOD = @REL_CON_FLOOD,
                    EXECT_DESC = @EXECT_DESC,
                    DELAY_RESN = @DELAY_RESN,
                    RESLV_MTHD = @RESLV_MTHD,
                    RESLV_LMT_DT = @RESLV_LMT_DT,
                    UPD_DT = @UPD_DT,
                    UPD_UID = @UPD_UID
                    WHERE ID = @ID";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.Add(new SqlParameter("ID", ID));
                cmd.Parameters.Add(new SqlParameter("ECC_UPAY_FEE", (object)ECC_UPAY_FEE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("ORG_SUM_PRE_SCHD", (object)ORG_SUM_PRE_SCHD ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("PRE_SCHD", (object)PRE_SCHD ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("REL_SCHD", (object)REL_SCHD ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("PRE_COST", (object)PRE_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("REL_COST", (object)REL_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("CT_PRE_COST", (object)CT_PRE_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("RV_PRE_COST", (object)RV_PRE_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("L_PLN_COST", (object)L_PLN_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("L_REL_COST", (object)L_REL_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("L_UNPAY_COST", (object)L_UNPAY_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("SAVE_FEE", (object)SAVE_FEE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("PRE_CON_LEN", (object)PRE_CON_LEN ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("REL_CON_LEN", (object)REL_CON_LEN ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("PRE_CON_FLOOD", (object)PRE_CON_FLOOD ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("REL_CON_FLOOD", (object)REL_CON_FLOOD ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("EXECT_DESC", (object)EXECT_DESC ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("DELAY_RESN", (object)DELAY_RESN ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("RESLV_MTHD", (object)RESLV_MTHD ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("RESLV_LMT_DT", (object)RESLV_LMT_DT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                TempData["Success"] = "儲存成功";
                TempData["SaveSuccess"] = true;
            } catch(Exception ex) {
                TempData["Error"] = "儲存失敗：" + ex.Message;
            }

            return RedirectToAction("PAGE5_2_Edit", new { ID = ID });
        }

        public ActionResult PAGE5_3(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            // 查詢經費來源列表
          string sql = @"SELECT P.PID, T.TID, T.FID, P.SYR, P.PJ_NAME
     FROM CMEPMS.dbo.CTDMNDS C
        LEFT JOIN CMEPMS.dbo.CTDITFD T ON C.TID = T.TID
     LEFT JOIN CMEPMS.dbo.PJMFORM P ON C.PID = P.PID
    WHERE T.FIT_CODE='11' AND C.TD_CODE=@TD_CODE
 ORDER BY P.SYR DESC";
            SqlCommand cmd = new SqlCommand(sql);
   cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
    DataTable dtList = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
         ViewBag.ListData = dtList;

            // 查詢 TDDMSCD 年度列表（用於下拉選單）
          string sqlSyr = @"SELECT DISTINCT SYR FROM TDDMSCD WHERE TD_CODE=@TD_CODE ORDER BY SYR DESC";
            SqlCommand cmdSyr = new SqlCommand(sqlSyr);
         cmdSyr.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dtSyr = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSyr);
       ViewBag.SyrList = dtSyr;

          // 預設查詢第一個年度的 TDDMSCD 資料
       if(dtSyr != null && dtSyr.Rows.Count > 0) {
   string defaultSyr = dtSyr.Rows[0]["SYR"].ToString();
     string sqlMscd = @"SELECT * FROM TDDMSCD WHERE TD_CODE=@TD_CODE AND SYR=@SYR ORDER BY SMN";
        SqlCommand cmdMscd = new SqlCommand(sqlMscd);
  cmdMscd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
   cmdMscd.Parameters.Add(new SqlParameter("SYR", defaultSyr));
 DataTable dtMscd = BASE_DAO.SqlDao.CMETMSQuery4List(cmdMscd);
   ViewBag.MscdData = dtMscd;
          ViewBag.SelectedSyr = defaultSyr;
            }

         TDE0040Model model = new TDE0040Model {
           TDMFORM = tdmform,
        TD_CODE = TD_CODE
     };

        GetSelectList();
            return View("PAGE5_3", model);
        }

        // AJAX: PAGE5_3 根據年度查詢 TDDMSCD 資料
        [HttpPost]
  public ActionResult PAGE5_3_GetMscdBySyr(string TD_CODE, string SYR) {
            try {
            string sql = @"SELECT * FROM TDDMSCD WHERE TD_CODE=@TD_CODE AND SYR=@SYR ORDER BY SMN";
      SqlCommand cmd = new SqlCommand(sql);
cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmd.Parameters.Add(new SqlParameter("SYR", SYR));
   DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

                var result = new List<Dictionary<string, object>>();
       foreach(DataRow row in dt.Rows) {
         var dict = new Dictionary<string, object>();
            foreach(DataColumn col in dt.Columns) {
  dict[col.ColumnName] = row[col] == DBNull.Value ? "" : row[col];
       }
      result.Add(dict);
      }
   return Json(new { success = true, data = result });
            } catch(Exception ex) {
         return Json(new { success = false, message = ex.Message });
     }
        }

        // PAGE5_3_Edit - 編輯經費來源 (GET)
        public ActionResult PAGE5_3_Edit(string TID, string FID, string TD_CODE) {
            if (string.IsNullOrEmpty(TID)) {
                return Content("TID 不可為空");
            }

            // 查詢 CTDITFD 資料
            string sql = @"SELECT * FROM CMEPMS.dbo.CTDITFD WHERE TID = @TID";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TID", TID));

            // 如果有 FID，則加入條件
            if (!string.IsNullOrEmpty(FID)) {
                sql = @"SELECT * FROM CMEPMS.dbo.CTDITFD WHERE TID = @TID AND FID = @FID";
                cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TID", TID));
                cmd.Parameters.Add(new SqlParameter("FID", FID));
            }

            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.TID = TID;
            ViewBag.TD_CODE = TD_CODE;

            if (dt != null && dt.Rows.Count > 0) {
                ViewBag.Data = dt.Rows[0];
                ViewBag.FID = dt.Rows[0]["FID"]?.ToString();
            } else {
                ViewBag.Data = null;
                ViewBag.FID = FID;
            }

            // 查詢關聯的 PJMFORM 資料 (透過 CTDMNDS)
            string sqlPj = @"SELECT P.PID, P.SYR, P.PJ_NAME
                FROM CMEPMS.dbo.CTDMNDS C
                LEFT JOIN CMEPMS.dbo.PJMFORM P ON C.PID = P.PID
                WHERE C.TID = @TID";
            SqlCommand cmdPj = new SqlCommand(sqlPj);
            cmdPj.Parameters.Add(new SqlParameter("TID", TID));
            DataTable dtPj = BASE_DAO.SqlDao.CMETMSQuery4List(cmdPj);
            ViewBag.PjData = (dtPj != null && dtPj.Rows.Count > 0) ? dtPj.Rows[0] : null;

            // 經費來源類型下拉選單
            ViewBag.FIT_CODE = tools.ToSelectList(util.GetCMEPMSCODES("FD", "05"), "id", "name", null);

            return View("PAGE5_3_Edit");
        }

        // PAGE5_3_Edit - 編輯經費來源 (POST)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE5_3_Edit_Save(string TID, string FID, string TD_CODE,
            string SYR, string FIT_CODE, decimal? FUND, string FUND_DESC, decimal? EXP_FUND,
            decimal? M_01, decimal? M_02, decimal? M_03, decimal? M_04,
            decimal? M_05, decimal? M_06, decimal? M_07, decimal? M_08,
            decimal? M_09, decimal? M_10, decimal? M_11, decimal? M_12) {

            try {
                string updateSql = @"UPDATE CMEPMS.dbo.CTDITFD SET
                    SYR = @SYR,
                    FIT_CODE = @FIT_CODE,
                    FUND = @FUND,
                    FUND_DESC = @FUND_DESC,
                    EXP_FUND = @EXP_FUND,
                    M_01 = @M_01, M_02 = @M_02, M_03 = @M_03, M_04 = @M_04,
                    M_05 = @M_05, M_06 = @M_06, M_07 = @M_07, M_08 = @M_08,
                    M_09 = @M_09, M_10 = @M_10, M_11 = @M_11, M_12 = @M_12,
                    UPD_DT = @UPD_DT,
                    UPD_UID = @UPD_UID
                    WHERE TID = @TID AND FID = @FID";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.Add(new SqlParameter("TID", TID));
                cmd.Parameters.Add(new SqlParameter("FID", FID));
                cmd.Parameters.Add(new SqlParameter("SYR", (object)SYR ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("FIT_CODE", (object)FIT_CODE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("FUND", (object)FUND ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("FUND_DESC", (object)FUND_DESC ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("EXP_FUND", (object)EXP_FUND ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_01", (object)M_01 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_02", (object)M_02 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_03", (object)M_03 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_04", (object)M_04 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_05", (object)M_05 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_06", (object)M_06 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_07", (object)M_07 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_08", (object)M_08 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_09", (object)M_09 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_10", (object)M_10 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_11", (object)M_11 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("M_12", (object)M_12 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                TempData["Success"] = "儲存成功";
                TempData["SaveSuccess"] = true;
            } catch (Exception ex) {
                TempData["Error"] = "儲存失敗：" + ex.Message;
            }

            return RedirectToAction("PAGE5_3_Edit", new { TID = TID, FID = FID, TD_CODE = TD_CODE });
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


    }
}