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
    public class TDE0047Controller : Controller {
        
        private readonly CMETMSEntities db = new CMETMSEntities();

      // GET: PAGE7
        public ActionResult PAGE7(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
   return RedirectToAction("Index");
         }

     var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

     if(tdmform == null) {
          TempData["Error"] = "找不到指定的工程資料";
             return RedirectToAction("Index");
            }

   // 查詢 TDDSPRS 資料
   string sql = @"SELECT * FROM TDDSPRS WHERE TD_CODE = @TD_CODE ORDER BY SP_SEQ DESC, SP_DT DESC";
            SqlCommand cmd = new SqlCommand(sql);
      cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
        DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
            ViewBag.List1 = dt;

          TDE0040Model model = new TDE0040Model {
            TD_CODE = TD_CODE,
       TDMFORM = tdmform
            };

            GetSelectList();
    return View("PAGE7", model);
        }

        #region PAGE7 AJAX 操作

   // POST: PAGE7_Add - 新增 TDDSPRS 資料 (AJAX)
     [HttpPost]
 public ActionResult PAGE7_Add(string TD_CODE, DateTime? SP_DT, string SP_DEPT, decimal? SP_SEQ,
     string SP_S, string SP_W, decimal? SP_PRE_SCHD, decimal? SP_REL_SCHD,
            string SP_RESULT, string SP_ITEM, decimal? SP_SCORE, decimal? SP_ALL_SCORE, string SP_LEVEL,
          string SP_TEST_ITEM, string SP_ITEM_RESULT, string SP_NOTE,
        decimal? EXEC_SCORE, decimal? CUS_SCORE, decimal? ADM_SCORE, decimal? WK_SCORE,
          decimal? MTEST_SCORE, decimal? SCUR_SCORE, decimal? QC_SCORE, decimal? QT_SCORE, decimal? WKSCHD_SCORE,
   string AC_TEST, string UP_MTHD,
            string CATEGORY, string IMP_DT, string FIN_DT, string HIATUS, string HIATUS2,
        decimal? CONSTRUCTION, decimal? SUPERVISION, decimal? HIATUS_PAY,
    string EXTEND_DT, string EXTEND, string CON_FN, string SUP_FN,
            string CON_DT, decimal? CON_FEE, string SUP_DT, decimal? SUP_FEE, string EXA_UNIT) {
  try {
   if(string.IsNullOrEmpty(TD_CODE)) {
    return Json(new { success = false, message = "TD_CODE 不可為空" });
          }

          string insertSql = @"INSERT INTO TDDSPRS 
    (TD_CODE, SP_DT, SP_DEPT, SP_SEQ, SP_S, SP_W, SP_PRE_SCHD, SP_REL_SCHD,
            SP_RESULT, SP_ITEM, SP_SCORE, SP_ALL_SCORE, SP_LEVEL,
           SP_TEST_ITEM, SP_ITEM_RESULT, SP_NOTE,
       EXEC_SCORE, CUS_SCORE, ADM_SCORE, WK_SCORE, MTEST_SCORE, SCUR_SCORE,
        QC_SCORE, QT_SCORE, WKSCHD_SCORE, AC_TEST, UP_MTHD,
       CATEGORY, IMP_DT, FIN_DT, HIATUS, HIATUS2,
    CONSTRUCTION, SUPERVISION, HIATUS_PAY,
           EXTEND_DT, EXTEND, CON_FN, SUP_FN,
               CON_DT, CON_FEE, SUP_DT, SUP_FEE, EXA_UNIT,
     ADD_DT, UPD_DT, UPD_UID)
         VALUES 
 (@TD_CODE, @SP_DT, @SP_DEPT, @SP_SEQ, @SP_S, @SP_W, @SP_PRE_SCHD, @SP_REL_SCHD,
        @SP_RESULT, @SP_ITEM, @SP_SCORE, @SP_ALL_SCORE, @SP_LEVEL,
            @SP_TEST_ITEM, @SP_ITEM_RESULT, @SP_NOTE,
          @EXEC_SCORE, @CUS_SCORE, @ADM_SCORE, @WK_SCORE, @MTEST_SCORE, @SCUR_SCORE,
           @QC_SCORE, @QT_SCORE, @WKSCHD_SCORE, @AC_TEST, @UP_MTHD,
          @CATEGORY, @IMP_DT, @FIN_DT, @HIATUS, @HIATUS2,
        @CONSTRUCTION, @SUPERVISION, @HIATUS_PAY,
         @EXTEND_DT, @EXTEND, @CON_FN, @SUP_FN,
      @CON_DT, @CON_FEE, @SUP_DT, @SUP_FEE, @EXA_UNIT,
        @ADD_DT, @UPD_DT, @UPD_UID)";

      SqlCommand cmd = new SqlCommand(insertSql);
      cmd.Parameters.Add(new SqlParameter("TD_CODE", (object)TD_CODE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("SP_DT", (object)SP_DT ?? DBNull.Value));
     cmd.Parameters.Add(new SqlParameter("SP_DEPT", (object)SP_DEPT ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("SP_SEQ", (object)SP_SEQ ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("SP_S", (object)SP_S ?? DBNull.Value));
       cmd.Parameters.Add(new SqlParameter("SP_W", (object)SP_W ?? DBNull.Value));
             cmd.Parameters.Add(new SqlParameter("SP_PRE_SCHD", (object)SP_PRE_SCHD ?? DBNull.Value));
            cmd.Parameters.Add(new SqlParameter("SP_REL_SCHD", (object)SP_REL_SCHD ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("SP_RESULT", (object)SP_RESULT ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("SP_ITEM", (object)SP_ITEM ?? DBNull.Value));
         cmd.Parameters.Add(new SqlParameter("SP_SCORE", (object)SP_SCORE ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("SP_ALL_SCORE", (object)SP_ALL_SCORE ?? DBNull.Value));
         cmd.Parameters.Add(new SqlParameter("SP_LEVEL", (object)SP_LEVEL ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("SP_TEST_ITEM", (object)SP_TEST_ITEM ?? DBNull.Value));
         cmd.Parameters.Add(new SqlParameter("SP_ITEM_RESULT", (object)SP_ITEM_RESULT ?? DBNull.Value));
         cmd.Parameters.Add(new SqlParameter("SP_NOTE", (object)SP_NOTE ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("EXEC_SCORE", (object)EXEC_SCORE ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("CUS_SCORE", (object)CUS_SCORE ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("ADM_SCORE", (object)ADM_SCORE ?? DBNull.Value));
          cmd.Parameters.Add(new SqlParameter("WK_SCORE", (object)WK_SCORE ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("MTEST_SCORE", (object)MTEST_SCORE ?? DBNull.Value));
            cmd.Parameters.Add(new SqlParameter("SCUR_SCORE", (object)SCUR_SCORE ?? DBNull.Value));
   cmd.Parameters.Add(new SqlParameter("QC_SCORE", (object)QC_SCORE ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("QT_SCORE", (object)QT_SCORE ?? DBNull.Value));
       cmd.Parameters.Add(new SqlParameter("WKSCHD_SCORE", (object)WKSCHD_SCORE ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("AC_TEST", (object)AC_TEST ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("UP_MTHD", (object)UP_MTHD ?? DBNull.Value));
       cmd.Parameters.Add(new SqlParameter("CATEGORY", (object)CATEGORY ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("IMP_DT", (object)IMP_DT ?? DBNull.Value));
           cmd.Parameters.Add(new SqlParameter("FIN_DT", (object)FIN_DT ?? DBNull.Value));
            cmd.Parameters.Add(new SqlParameter("HIATUS", (object)HIATUS ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("HIATUS2", (object)HIATUS2 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("CONSTRUCTION", (object)CONSTRUCTION ?? DBNull.Value));
  cmd.Parameters.Add(new SqlParameter("SUPERVISION", (object)SUPERVISION ?? DBNull.Value));
  cmd.Parameters.Add(new SqlParameter("HIATUS_PAY", (object)HIATUS_PAY ?? DBNull.Value));
 cmd.Parameters.Add(new SqlParameter("EXTEND_DT", (object)EXTEND_DT ?? DBNull.Value));
              cmd.Parameters.Add(new SqlParameter("EXTEND", (object)EXTEND ?? DBNull.Value));
     cmd.Parameters.Add(new SqlParameter("CON_FN", (object)CON_FN ?? DBNull.Value));
 cmd.Parameters.Add(new SqlParameter("SUP_FN", (object)SUP_FN ?? DBNull.Value));
   cmd.Parameters.Add(new SqlParameter("CON_DT", (object)CON_DT ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("CON_FEE", (object)CON_FEE ?? DBNull.Value));
         cmd.Parameters.Add(new SqlParameter("SUP_DT", (object)SUP_DT ?? DBNull.Value));
           cmd.Parameters.Add(new SqlParameter("SUP_FEE", (object)SUP_FEE ?? DBNull.Value));
       cmd.Parameters.Add(new SqlParameter("EXA_UNIT", (object)EXA_UNIT ?? DBNull.Value));
   cmd.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now));
                cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
  cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

      return Json(new { success = true, message = "新增成功" });
            } catch(Exception ex) {
            return Json(new { success = false, message = "新增失敗：" + ex.Message });
            }
        }

 // POST: PAGE7_Update - 更新 TDDSPRS 資料 (AJAX)
        [HttpPost]
        public ActionResult PAGE7_Update(decimal SPID, DateTime? SP_DT, string SP_DEPT, decimal? SP_SEQ,
         string SP_S, string SP_W, decimal? SP_PRE_SCHD, decimal? SP_REL_SCHD,
       string SP_RESULT, string SP_ITEM, decimal? SP_SCORE, decimal? SP_ALL_SCORE, string SP_LEVEL,
            string SP_TEST_ITEM, string SP_ITEM_RESULT, string SP_NOTE,
 decimal? EXEC_SCORE, decimal? CUS_SCORE, decimal? ADM_SCORE, decimal? WK_SCORE,
            decimal? MTEST_SCORE, decimal? SCUR_SCORE, decimal? QC_SCORE, decimal? QT_SCORE, decimal? WKSCHD_SCORE,
   string AC_TEST, string UP_MTHD,
      string CATEGORY, string IMP_DT, string FIN_DT, string HIATUS, string HIATUS2,
     decimal? CONSTRUCTION, decimal? SUPERVISION, decimal? HIATUS_PAY,
            string EXTEND_DT, string EXTEND, string CON_FN, string SUP_FN,
string CON_DT, decimal? CON_FEE, string SUP_DT, decimal? SUP_FEE, string EXA_UNIT) {
          try {
          string updateSql = @"UPDATE TDDSPRS SET 
              SP_DT = @SP_DT, SP_DEPT = @SP_DEPT, SP_SEQ = @SP_SEQ,
   SP_S = @SP_S, SP_W = @SP_W, SP_PRE_SCHD = @SP_PRE_SCHD, SP_REL_SCHD = @SP_REL_SCHD,
    SP_RESULT = @SP_RESULT, SP_ITEM = @SP_ITEM, SP_SCORE = @SP_SCORE,
         SP_ALL_SCORE = @SP_ALL_SCORE, SP_LEVEL = @SP_LEVEL,
              SP_TEST_ITEM = @SP_TEST_ITEM, SP_ITEM_RESULT = @SP_ITEM_RESULT, SP_NOTE = @SP_NOTE,
        EXEC_SCORE = @EXEC_SCORE, CUS_SCORE = @CUS_SCORE, ADM_SCORE = @ADM_SCORE,
        WK_SCORE = @WK_SCORE, MTEST_SCORE = @MTEST_SCORE, SCUR_SCORE = @SCUR_SCORE,
  QC_SCORE = @QC_SCORE, QT_SCORE = @QT_SCORE, WKSCHD_SCORE = @WKSCHD_SCORE,
        AC_TEST = @AC_TEST, UP_MTHD = @UP_MTHD,
     CATEGORY = @CATEGORY, IMP_DT = @IMP_DT, FIN_DT = @FIN_DT,
    HIATUS = @HIATUS, HIATUS2 = @HIATUS2,
CONSTRUCTION = @CONSTRUCTION, SUPERVISION = @SUPERVISION, HIATUS_PAY = @HIATUS_PAY,
      EXTEND_DT = @EXTEND_DT, EXTEND = @EXTEND, CON_FN = @CON_FN, SUP_FN = @SUP_FN,
           CON_DT = @CON_DT, CON_FEE = @CON_FEE, SUP_DT = @SUP_DT, SUP_FEE = @SUP_FEE, EXA_UNIT = @EXA_UNIT,
        UPD_DT = @UPD_DT, UPD_UID = @UPD_UID
         WHERE SPID = @SPID";

   SqlCommand cmd = new SqlCommand(updateSql);
           cmd.Parameters.Add(new SqlParameter("SPID", SPID));
 cmd.Parameters.Add(new SqlParameter("SP_DT", (object)SP_DT ?? DBNull.Value));
     cmd.Parameters.Add(new SqlParameter("SP_DEPT", (object)SP_DEPT ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("SP_SEQ", (object)SP_SEQ ?? DBNull.Value));
 cmd.Parameters.Add(new SqlParameter("SP_S", (object)SP_S ?? DBNull.Value));
              cmd.Parameters.Add(new SqlParameter("SP_W", (object)SP_W ?? DBNull.Value));
   cmd.Parameters.Add(new SqlParameter("SP_PRE_SCHD", (object)SP_PRE_SCHD ?? DBNull.Value));
  cmd.Parameters.Add(new SqlParameter("SP_REL_SCHD", (object)SP_REL_SCHD ?? DBNull.Value));
     cmd.Parameters.Add(new SqlParameter("SP_RESULT", (object)SP_RESULT ?? DBNull.Value));
 cmd.Parameters.Add(new SqlParameter("SP_ITEM", (object)SP_ITEM ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("SP_SCORE", (object)SP_SCORE ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("SP_ALL_SCORE", (object)SP_ALL_SCORE ?? DBNull.Value));
  cmd.Parameters.Add(new SqlParameter("SP_LEVEL", (object)SP_LEVEL ?? DBNull.Value));
         cmd.Parameters.Add(new SqlParameter("SP_TEST_ITEM", (object)SP_TEST_ITEM ?? DBNull.Value));
           cmd.Parameters.Add(new SqlParameter("SP_ITEM_RESULT", (object)SP_ITEM_RESULT ?? DBNull.Value));
       cmd.Parameters.Add(new SqlParameter("SP_NOTE", (object)SP_NOTE ?? DBNull.Value));
 cmd.Parameters.Add(new SqlParameter("EXEC_SCORE", (object)EXEC_SCORE ?? DBNull.Value));
  cmd.Parameters.Add(new SqlParameter("CUS_SCORE", (object)CUS_SCORE ?? DBNull.Value));
  cmd.Parameters.Add(new SqlParameter("ADM_SCORE", (object)ADM_SCORE ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("WK_SCORE", (object)WK_SCORE ?? DBNull.Value));
     cmd.Parameters.Add(new SqlParameter("MTEST_SCORE", (object)MTEST_SCORE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("SCUR_SCORE", (object)SCUR_SCORE ?? DBNull.Value));
 cmd.Parameters.Add(new SqlParameter("QC_SCORE", (object)QC_SCORE ?? DBNull.Value));
         cmd.Parameters.Add(new SqlParameter("QT_SCORE", (object)QT_SCORE ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("WKSCHD_SCORE", (object)WKSCHD_SCORE ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("AC_TEST", (object)AC_TEST ?? DBNull.Value));
      cmd.Parameters.Add(new SqlParameter("UP_MTHD", (object)UP_MTHD ?? DBNull.Value));
         cmd.Parameters.Add(new SqlParameter("CATEGORY", (object)CATEGORY ?? DBNull.Value));
            cmd.Parameters.Add(new SqlParameter("IMP_DT", (object)IMP_DT ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("FIN_DT", (object)FIN_DT ?? DBNull.Value));
             cmd.Parameters.Add(new SqlParameter("HIATUS", (object)HIATUS ?? DBNull.Value));
     cmd.Parameters.Add(new SqlParameter("HIATUS2", (object)HIATUS2 ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("CONSTRUCTION", (object)CONSTRUCTION ?? DBNull.Value));
       cmd.Parameters.Add(new SqlParameter("SUPERVISION", (object)SUPERVISION ?? DBNull.Value));
             cmd.Parameters.Add(new SqlParameter("HIATUS_PAY", (object)HIATUS_PAY ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("EXTEND_DT", (object)EXTEND_DT ?? DBNull.Value));
       cmd.Parameters.Add(new SqlParameter("EXTEND", (object)EXTEND ?? DBNull.Value));
          cmd.Parameters.Add(new SqlParameter("CON_FN", (object)CON_FN ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("SUP_FN", (object)SUP_FN ?? DBNull.Value));
          cmd.Parameters.Add(new SqlParameter("CON_DT", (object)CON_DT ?? DBNull.Value));
       cmd.Parameters.Add(new SqlParameter("CON_FEE", (object)CON_FEE ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("SUP_DT", (object)SUP_DT ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("SUP_FEE", (object)SUP_FEE ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("EXA_UNIT", (object)EXA_UNIT ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
          cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

        return Json(new { success = true, message = "更新成功" });
  } catch(Exception ex) {
    return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        // POST: PAGE7_Delete - 刪除 TDDSPRS 資料 (AJAX)
        [HttpPost]
        public ActionResult PAGE7_Delete(decimal SPID) {
            try {
        string deleteSql = @"DELETE FROM TDDSPRS WHERE SPID = @SPID";

  SqlCommand cmd = new SqlCommand(deleteSql);
         cmd.Parameters.Add(new SqlParameter("SPID", SPID));

     BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

 return Json(new { success = true, message = "刪除成功" });
 } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
    }
        }

        #endregion

 public void GetSelectList() {
       SqlCommand cmd = new SqlCommand("SELECT SYR, PID, PJ_NAME FROM PJMFORM WHERE ADV_UID='" + WebSiteHelper.getUserId() + "'ORDER BY SYR DESC");
   DataTable pids = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.PIDS = tools.ToSelectList(pids, "PID", "PJ_NAME", null);

            cmd = new SqlCommand("SELECT SYR FROM PJMFORM WHERE ADV_UID='" + WebSiteHelper.getUserId() + "'GROUP BY SYR ORDER BY SYR DESC");
   DataTable syrs = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
   ViewBag.SYRS = tools.ToSelectList(syrs, "SYR", "SYR", null);

         DataTable dt;
   ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
          ViewBag.towns = tools.ToSelectList(util.GetTownNew(null), "TOWN", "TOWN", null);

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

        ViewBag.IPJ_CODE = tools.ToSelectList(util.GetCMEPMSCODES("SM", "02"), "id", "name", null);
        ViewBag.N_IPJ_CODE = tools.ToSelectList(util.GetCMETMSCODES("SM", "02"), "id", "name", null);
ViewBag.EMG_CATY = tools.ToSelectList(util.GetCMEPMSCODES("TT", "05"), "id", "name", null);
   ViewBag.ACC_ID = tools.ToSelectList(util.GetCMEPMSCODES("FD", "02"), "id", "name", null);
       ViewBag.YR_WORK = tools.ToSelectList(util.GetCMETMSCODES("TT", "04"), "id", "name", null);
  ViewBag.TD_CATY = tools.ToSelectList(util.GetCMETMSCODES("SM", "09"), "id", "name", null);

            string sql = "SELECT o.DEPT_NO, MAX(o.DEPT_NAME) AS DEPT_NAME, COUNT(r.USR_ID) AS USER_COUNT FROM OGCDPUR r JOIN OGCDEPT o ON o.DEPT_NO = r.DEPT_NO JOIN OGCUSER u ON u.USR_ID = r.USR_ID GROUP BY o.DEPT_NO;";
   ViewBag.depts = tools.ToSelectList(SqlDao.CMBASEQuery4List(sql), "DEPT_NO", "DEPT_NAME", "請選擇單位");

 ViewBag.CATY_COD = tools.ToSelectList(util.GetCMETMSCODES("TT", "03"), "id", "name", null);
     ViewBag.FIT_CODE = tools.ToSelectList(util.GetCMEPMSCODES("FD", "05"), "id", "name", null);
            ViewBag.users = tools.ToSelectList(SqlDao.CMBASEQuery4List("SELECT USR_ID, USR_NAME FROM OGCUSER"), "USR_ID", "USR_NAME", "請選擇");

            try {
      var sysList = db.PJCSYST.Select(s => new { s.SYS_ID, s.SYS_NAME }).ToList();
        ViewBag.SYS_ID = new SelectList(sysList, "SYS_ID", "SYS_NAME");
            } catch {
        ViewBag.SYS_ID = new SelectList(new List<SelectListItem>(), "Value", "Text");
            }

      try {
    var pjList = db.PJCFORM.Select(p => new { p.PJ_ID, p.PJ_NAME }).ToList();
     ViewBag.PJ_ID = new SelectList(pjList, "PJ_ID", "PJ_NAME");
       } catch {
   ViewBag.PJ_ID = new SelectList(new List<SelectListItem>(), "Value", "Text");
            }
        }
    }
}