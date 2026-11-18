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
    public class TDE0042Controller : Controller {

        private readonly CMETMSEntities db = new CMETMSEntities();

        // GET: PAGE2

        public ActionResult PAGE2(string TD_CODE) {
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
                TD_CODE = TD_CODE
            };

            GetSelectList();
            return View("PAGE2", model);
        }


        public ActionResult PAGE2_1(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            String sql = @"select 
            P.PID, --計畫序號key
            T.TID, --工程經費序號key
            T.TD_CODE, --工程編號
            P.SYR, --年度
            '(' + P.SYR + ')' + P.PJ_NAME as PJ_NAME, --計畫名稱
            P.FLOCK_FG, --經費已鎖 1:是 0:否
            P.ADV_DP_NAME, --計畫承辦單位
            -- 【修改重點】：改用 CASE WHEN 判斷 NULL 與小於等於 1 的情況
            CASE 
            WHEN SEQ_NO IS NULL OR SEQ_NO = 0 THEN '原預算'
            ELSE '第' + convert(nvarchar(3), SEQ_NO) + '期'
            END AS SEQ_NAME , --計畫經費期別
            T.YR_FUND, --年度預算金額(元)
            T.CEN_FUND --年度中央款金額(元)
            from
            CMEPMS.dbo.PJV_PJMFORM P left join CMEPMS.dbo.TDV_TDDCFTS T
            on P.PID = T.PID left join CMBASE.dbo.OGCUSER O
            on T.UPD_UID = O.USR_ID
            where P.PID in (select PID from PMS_CTDMNDS where TD_CODE = @TD_CODE)
            and T.TD_CODE = @TD_CODE
            order by P.SYR desc, P.PID desc
            ";

            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            //cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);


            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform,
                TD_CODE = TD_CODE
            };



            ViewBag.List1 = ps;





            return View("PAGE2_1", model);
        }

        public ActionResult PAGE2_2(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            // 直接查詢 CRDMNS Table （不使用 View，確保新增的資料能顯示）
            string sql = @"SELECT * FROM CRDMNDS WHERE TD_CODE=@TD_CODE";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.List1 = dt;

            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform,
                TD_CODE = TD_CODE
            };

            return View("PAGE2_2", model);
        }

        // PAGE2_1_1 - 用 TD_CODE 查詢，顯示計畫經費列表，期別下拉選單用 PID 查詢 PJDFDCR
        public ActionResult PAGE2_1_1(string TD_CODE, string PID, string SEQ_NO) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 取得第一筆的 PID（用於查詢期別下拉選單）
            string firstPID = PID ?? "";

            // 先查詢期別下拉選單，取得最大期別
            string maxSeqNo = "";
            DataTable dtSeq = null;

            // 如果沒有 PID，先用 TD_CODE 查詢取得 PID
            if(string.IsNullOrEmpty(firstPID)) {
                string pidSql = @"SELECT TOP 1 PID FROM CMEPMS.dbo.CTDMNDS WHERE TD_CODE = @TD_CODE";
                SqlCommand pidCmd = new SqlCommand(pidSql);
                pidCmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                DataTable dtPid = BASE_DAO.SqlDao.CMEPMSQuery4List(pidCmd);
                if(dtPid != null && dtPid.Rows.Count > 0) {
                    firstPID = dtPid.Rows[0]["PID"].ToString();
                }
            }

            // 查詢期別下拉選單
            if(!string.IsNullOrEmpty(firstPID)) {
                string seqSql = @"SELECT SEQ_NO, 
        /*CASE SEQ_NO 
           WHEN 0 THEN '原預算'
           ELSE '第' + convert(nvarchar(3), SEQ_NO) + '期'
           END SEQ_NAME*/
            CASE 
            WHEN SEQ_NO IS NULL OR SEQ_NO = 0 THEN '原預算'
            ELSE '第' + convert(nvarchar(3), SEQ_NO) + '期'
        END AS SEQ_NAME 
    FROM CMEPMS.dbo.PJDFDCR 
          WHERE PID = @PID 
         ORDER BY SEQ_NO DESC";

                SqlCommand seqCmd = new SqlCommand(seqSql);
                seqCmd.Parameters.Add(new SqlParameter("PID", firstPID));
                dtSeq = BASE_DAO.SqlDao.CMEPMSQuery4List(seqCmd);

                if(dtSeq != null && dtSeq.Rows.Count > 0) {
                    maxSeqNo = dtSeq.Rows[0]["SEQ_NO"].ToString();
                }
                ViewBag.SEQ_NO_List = tools.ToSelectList(dtSeq, "SEQ_NO", "SEQ_NAME", null);
            } else {
                dtSeq = new DataTable();
                dtSeq.Columns.Add("SEQ_NO");
                dtSeq.Columns.Add("SEQ_NAME");
                ViewBag.SEQ_NO_List = tools.ToSelectList(dtSeq, "SEQ_NO", "SEQ_NAME", null);
            }

            // 決定要查詢的期別（預設為最大期別）
            string querySeqNo = string.IsNullOrEmpty(SEQ_NO) ? maxSeqNo : SEQ_NO;
            bool isLatestSeq = (querySeqNo == maxSeqNo);

            // 根據期別選擇不同的查詢
            DataTable dt;
            if(isLatestSeq) {
                // 最新期：查詢 TDV_TDDCFTS（現有邏輯）
                String sql = @"SELECT 
           P.PID,
 T.TID,
          T.TD_CODE,
                    P.SYR,
        '('+ P.SYR + ')' + P.PJ_NAME as PJ_NAME,
    P.FLOCK_FG,
               P.ADV_DP_NAME,
         /*CASE T.SEQ_NO-1
            WHEN 0 THEN '原預算'
            ELSE '第' + convert(nvarchar(3), T.SEQ_NO-1) + '期'
        END SEQ_NAME,*/
                CASE 
            WHEN SEQ_NO IS NULL OR SEQ_NO = 0 THEN '原預算'
            ELSE '第' + convert(nvarchar(3), SEQ_NO) + '期'
            END AS SEQ_NAME , 
        T.YR_FUND,
           T.CEN_FUND
                FROM 
   CMEPMS.dbo.PJV_PJMFORM P 
      LEFT JOIN CMEPMS.dbo.TDV_TDDCFTS T ON P.PID = T.PID 
     LEFT JOIN CMBASE.dbo.OGCUSER O ON T.UPD_UID = O.USR_ID
    WHERE P.PID IN (SELECT PID FROM CMEPMS.dbo.CTDMNDS WHERE TD_CODE = @TD_CODE)
       AND T.TD_CODE = @TD_CODE
 ORDER BY P.SYR DESC, P.PID DESC";

                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            } else {
                // 舊期數：查詢 TDV_TDHCFTS（歷史資料表）
                String sql = @"SELECT 
  P.PID,
    T.TID,
           T.TD_CODE,
            P.SYR,
           '('+ P.SYR + ')' + P.PJ_NAME as PJ_NAME,
  P.FLOCK_FG,
           P.ADV_DP_NAME,
        CASE T.SEQ_NO-1
          WHEN 0 THEN '原預算'
   ELSE '第' + CONVERT(NVARCHAR(3), T.SEQ_NO-1) + '期' 
        END SEQ_NAME,
            T.YR_FUND,
    T.CEN_FUND
                FROM 
      CMEPMS.dbo.PJV_PJMFORM P 
            LEFT JOIN CMEPMS.dbo.TDV_TDHCFTS T ON P.PID = T.PID 
      LEFT JOIN CMBASE.dbo.OGCUSER O ON T.UPD_UID = O.USR_ID
    WHERE P.PID IN (SELECT PID FROM CMEPMS.dbo.CTDMNDS WHERE TD_CODE = @TD_CODE)
         AND T.TD_CODE = @TD_CODE
      AND T.SEQ_NO = @SEQ_NO
    ORDER BY P.SYR DESC, P.PID DESC";

                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmd.Parameters.Add(new SqlParameter("SEQ_NO", querySeqNo));
                dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            }

            // 如果查詢結果有資料，更新 PID
            if(dt != null && dt.Rows.Count > 0) {
                firstPID = dt.Rows[0]["PID"].ToString();
            }

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.List1 = dt;
            ViewBag.PID = firstPID;
            ViewBag.MaxSeqNo = maxSeqNo;
            ViewBag.CurrentSeqNo = querySeqNo;
            ViewBag.IsLatestSeq = isLatestSeq;

            TDE0040Model model = new TDE0040Model {
                PID = firstPID,
                TD_CODE = TD_CODE
            };

            return View("PAGE2_1_1", model);
        }

        // PAGE2_1_Add - 新增經費來源（GET）
        public ActionResult PAGE2_1_Add(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return Content("TD_CODE 不可為空");
            }

            // 取得工程資料
            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);
            if(tdmform == null) {
                return Content("找不到指定的工程資料");
            }

            // 取得可選擇的計畫列表（PJMFORM），排除已關聯的 PID
            string sql = @"SELECT P.PID, P.SYR, P.PJ_NAME,
                           '(' + P.SYR + ') ' + P.PJ_NAME AS PJ_DISPLAY
                           FROM CMEPMS.dbo.PJMFORM P
                           WHERE P.PID NOT IN (
                               SELECT DISTINCT PID FROM CMEPMS.dbo.CTDMNDS
                               WHERE TD_CODE = @TD_CODE AND PID IS NOT NULL
                           )
                           ORDER BY P.SYR DESC, P.PID DESC";
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dtPlans = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.TD_NAME = tdmform.TD_NAME;
            ViewBag.PlanList = tools.ToSelectList(dtPlans, "PID", "PJ_DISPLAY", "請選擇計畫");

            // 取得經費層級下拉選單
            ViewBag.FIT_CODE_List = tools.ToSelectList(util.GetCMEPMSCODES("FD", "05"), "id", "name", "請選擇");

            // 取得經費代碼下拉選單 (FD_CODE)
            ViewBag.FD_CODE_List = tools.ToSelectList(util.GetCMEPMSCODES("FD", "01"), "id", "name", "請選擇");

            return View("PAGE2_1_Add");
        }

        // PAGE2_1_Add - 新增經費來源（POST）
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE2_1_Add_Save(string TD_CODE, string PID, string FD_CODE,
            string SYR, string FIT_CODE, decimal? FUND, string FUND_DESC) {

            try {
                if(string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(PID)) {
                    TempData["Error"] = "請選擇計畫";
                    return RedirectToAction("PAGE2_1_Add", new { TD_CODE = TD_CODE });
                }

                if(string.IsNullOrEmpty(FD_CODE)) {
                    TempData["Error"] = "請選擇經費代碼";
                    return RedirectToAction("PAGE2_1_Add", new { TD_CODE = TD_CODE });
                }

                // 取得工程名稱
                var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);
                string tdName = tdmform?.TD_NAME ?? "";

                // 1. 先新增 CTDMNDS（產生 TID）
                string insertCtdmndsSql = @"INSERT INTO CMEPMS.dbo.CTDMNDS
                    (PID, TD_CODE, TD_NAME, FD_CODE, ADD_DT, UPD_DT, UPD_UID)
                    VALUES (@PID, @TD_CODE, @TD_NAME, @FD_CODE, @ADD_DT, @UPD_DT, @UPD_UID);
                    SELECT SCOPE_IDENTITY();";

                SqlCommand cmdCtdmnds = new SqlCommand(insertCtdmndsSql);
                cmdCtdmnds.Parameters.Add(new SqlParameter("PID", PID));
                cmdCtdmnds.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmdCtdmnds.Parameters.Add(new SqlParameter("TD_NAME", tdName));
                cmdCtdmnds.Parameters.Add(new SqlParameter("FD_CODE", FD_CODE));
                cmdCtdmnds.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now));
                cmdCtdmnds.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                cmdCtdmnds.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                DataTable dtResult = BASE_DAO.SqlDao.CMEPMSQuery4List(cmdCtdmnds);
                string newTID = "";
                if(dtResult.Rows.Count > 0 && dtResult.Rows[0][0] != DBNull.Value) {
                    newTID = dtResult.Rows[0][0].ToString();
                }

                // 2. 如果有填寫經費資料，新增 CTDITFD
                if(!string.IsNullOrEmpty(newTID) && !string.IsNullOrEmpty(FIT_CODE) && FUND.HasValue) {
                    string insertCtditfdSql = @"INSERT INTO CMEPMS.dbo.CTDITFD
                        (TID, FID, SYR, FIT_CODE, FUND, FUND_DESC, ADD_DT, UPD_DT, UPD_UID)
                        VALUES (@TID, 1, @SYR, @FIT_CODE, @FUND, @FUND_DESC, @ADD_DT, @UPD_DT, @UPD_UID)";

                    SqlCommand cmdCtditfd = new SqlCommand(insertCtditfdSql);
                    cmdCtditfd.Parameters.Add(new SqlParameter("TID", newTID));
                    cmdCtditfd.Parameters.Add(new SqlParameter("SYR", (object)SYR ?? DBNull.Value));
                    cmdCtditfd.Parameters.Add(new SqlParameter("FIT_CODE", FIT_CODE));
                    cmdCtditfd.Parameters.Add(new SqlParameter("FUND", FUND.Value));
                    cmdCtditfd.Parameters.Add(new SqlParameter("FUND_DESC", (object)FUND_DESC ?? DBNull.Value));
                    cmdCtditfd.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now));
                    cmdCtditfd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                    cmdCtditfd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                    BASE_DAO.SqlDao.CMEPMSExecuteSql(cmdCtditfd);
                }

                TempData["Success"] = "新增經費來源成功，TID=" + newTID;

                // 關閉視窗並重新整理父頁面
                return Content(@"<script>
                    if (window.opener && typeof window.opener.location !== 'undefined') {
                        window.opener.location.reload();
                    }
                    if (typeof top.close2_1_Add === 'function') {
                        top.close2_1_Add();
                    } else {
                        window.close();
                    }
                </script>", "text/html");

            } catch(Exception ex) {
                System.Diagnostics.Debug.WriteLine($"PAGE2_1_Add_Save error: {ex.Message}");
                TempData["Error"] = "新增失敗：" + ex.Message;
                return RedirectToAction("PAGE2_1_Add", new { TD_CODE = TD_CODE });
            }
        }

        // PAGE2_1_2 - 查詢 TDV_TDDCFTS 月份分配資料
        public ActionResult PAGE2_1_2(string TID, string PID, string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE) && string.IsNullOrEmpty(TID)) {
                return Content("TD_CODE 或 TID 不可為空");
            }

            DataTable dt = null;

            // 優先用 TID 查詢
            if(!string.IsNullOrEmpty(TID)) {
                string sql = @"SELECT 
   T.TID, 
    T.TD_CODE,
            T.PID,
         T.YR_FUND,
        T.CEN_FUND,
          T.CEN_SP_RATIO,
       T.CEN_SP_RATIO_S,
            T.CEN_SP_RATIO as CEN_SP_RATIO_D,
          T.UPD_DT,
      O.USR_NAME as UPD_USR_NAME,
          P.SYR
            FROM CMEPMS.dbo.TDV_TDDCFTS T 
        LEFT JOIN CMBASE.dbo.OGCUSER O ON T.UPD_UID = O.USR_ID
         LEFT JOIN CMEPMS.dbo.PJMFORM P ON T.PID = P.PID
     WHERE T.TID = @TID";

                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TID", TID));
                dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            }

            // 如果 TID 查不到，改用 TD_CODE 查詢
            if((dt == null || dt.Rows.Count == 0) && !string.IsNullOrEmpty(TD_CODE)) {
                string sql = @"SELECT 
        T.TID, 
         T.TD_CODE,
     T.PID,
T.YR_FUND,
         T.CEN_FUND,
         T.CEN_SP_RATIO,
              T.CEN_SP_RATIO_S,
          T.CEN_SP_RATIO as CEN_SP_RATIO_D,
        T.UPD_DT,
     O.USR_NAME as UPD_USR_NAME,
   P.SYR
     FROM CMEPMS.dbo.TDV_TDDCFTS T 
     LEFT JOIN CMBASE.dbo.OGCUSER O ON T.UPD_UID = O.USR_ID
           LEFT JOIN CMEPMS.dbo.PJMFORM P ON T.PID = P.PID
   WHERE T.TD_CODE = @TD_CODE";

                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            }

            string currentTID = (dt != null && dt.Rows.Count > 0) ? dt.Rows[0]["TID"].ToString() : TID;
            string currentTD_CODE = (dt != null && dt.Rows.Count > 0) ? dt.Rows[0]["TD_CODE"].ToString() : TD_CODE;
            ViewBag.TID = currentTID;
            ViewBag.PID = (dt != null && dt.Rows.Count > 0) ? dt.Rows[0]["PID"].ToString() : PID;
            ViewBag.TD_CODE = currentTD_CODE;
            ViewBag.IsNew = (dt == null || dt.Rows.Count == 0);

            // 取得 CITY（從 TDMFORM）
            string city = "";
            if(!string.IsNullOrEmpty(currentTD_CODE)) {
                string citySql = @"SELECT TD_CITY FROM CMETMS.dbo.TDMFORM WHERE TD_CODE = @TD_CODE";
                SqlCommand cityCmd = new SqlCommand(citySql);
                cityCmd.Parameters.Add(new SqlParameter("TD_CODE", currentTD_CODE));
                DataTable dtCity = BASE_DAO.SqlDao.CMETMSQuery4List(cityCmd);
                if(dtCity != null && dtCity.Rows.Count > 0 && dtCity.Rows[0]["TD_CITY"] != DBNull.Value) {
                    city = dtCity.Rows[0]["TD_CITY"].ToString();
                }
            }
            ViewBag.CITY = city;

            if(dt != null && dt.Rows.Count > 0) {
                ViewBag.Data = dt.Rows[0];
            } else {
                ViewBag.Data = null;
            }

            // 查詢經費明細列表（用 TID 查詢 TDV_CTDMNDS）
            DataTable dtFundList = null;
            if(!string.IsNullOrEmpty(currentTID)) {
                string fundSql = @"SELECT PID, TID, FID, FSYR as SYR, FIT_CODE, FIT_DESC, FUND, FUND_DESC 
         FROM CMEPMS.dbo.TDV_CTDMNDS WHERE TID = @TID";
                SqlCommand fundCmd = new SqlCommand(fundSql);
                fundCmd.Parameters.Add(new SqlParameter("TID", currentTID));
                dtFundList = BASE_DAO.SqlDao.CMEPMSQuery4List(fundCmd);
            }
            ViewBag.FundList = dtFundList;

            // 3.1.4rev：月分配數
            // 業主回報「月分配數功能無法使用」——本頁原本只做了預算分配與經費明細，
            // 缺少月分配。舊系統同畫面 App_Data\CMETMS\TDE0020B020.Xml 定義為
            // 每筆經費明細(FID) 各有 M_01~M_12，實體欄位在 CMEPMS.dbo.CTDITFD。
            DataTable dtMonthList = null;
            if(!string.IsNullOrEmpty(currentTID)) {
                string monthSql = @"SELECT FID, SYR, FUND,
                        M_01, M_02, M_03, M_04, M_05, M_06,
                        M_07, M_08, M_09, M_10, M_11, M_12
                    FROM CMEPMS.dbo.CTDITFD WHERE TID = @TID ORDER BY SYR, FID";
                SqlCommand monthCmd = new SqlCommand(monthSql);
                monthCmd.Parameters.Add(new SqlParameter("TID", currentTID));
                dtMonthList = BASE_DAO.SqlDao.CMEPMSQuery4List(monthCmd);
            }
            ViewBag.MonthList = dtMonthList;

            // 取得經費層級下拉選單
          ViewBag.FIT_CODE_List = tools.ToSelectList(util.GetCMEPMSCODES("FD", "05"), "id", "name", "請選擇");

            TDE0040Model model = new TDE0040Model {
                PID = ViewBag.PID,
                TD_CODE = ViewBag.TD_CODE
            };

            return View("PAGE2_1_2", model);
        }
        /*
        public ActionResult PAGE2_1_2(string TID, string PID, string TD_CODE) {
            if(string.IsNullOrEmpty(TID) && string.IsNullOrEmpty(PID)) {
                return Content("TID 或 PID 不可為空");
            }

            string sql = @"select C.TID, C.TD_CODE, C.PID, P.SYR,
           C.YR_FUND,
     C.M_01,
            C.M_02,
         C.M_03,
    C.M_04,
           C.M_05,
        C.M_06,
 C.M_07,
          C.M_08,
     C.M_09,
         C.M_10,
           C.M_11,
   C.M_12
              from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.PJMFORM P
         on C.PID = P.PID
   where C.TID = @TID";

            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("TID", TID ?? ""));
            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            ViewBag.TID = TID;
            ViewBag.PID = PID;
            ViewBag.TD_CODE = TD_CODE;
            ViewBag.IsNew = dt.Rows.Count == 0;

            if(dt.Rows.Count > 0) {
                // 有資料，帶出編輯
                ViewBag.Data = dt.Rows[0];
            } else {
                // 無資料，建立空白資料供新增
                ViewBag.Data = null;
            }

            TDE0040Model model = new TDE0040Model {
                PID = PID,
                TD_CODE = TD_CODE
            };

            return View("PAGE2_1_2", model);
        }
        */

        // POST: PAGE2_1_2 - 儲存中央款比例(標準) 到 PJDCTSP 表
        // 注意：YR_FUND, CEN_FUND 是由 CTDITFD 彙總計算的唯讀值，不需要儲存
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE2_1_2_Save(string TID, string PID, string TD_CODE, string CITY,
            decimal? CEN_SP_RATIO_S, string IsNew = "false") {

            try {
                if(string.IsNullOrEmpty(PID) || string.IsNullOrEmpty(CITY)) {
                    TempData["Error"] = "缺少必要參數（PID 或 CITY）";
                    return RedirectToAction("PAGE2_1_2", new { TID = TID, PID = PID, TD_CODE = TD_CODE });
                }

                // 檢查 PJDCTSP 是否已存在該筆資料
                string checkSql = @"SELECT COUNT(*) FROM CMEPMS.dbo.PJDCTSP WHERE PID = @PID AND CITY = @CITY";
                SqlCommand checkCmd = new SqlCommand(checkSql);
                checkCmd.Parameters.Add(new SqlParameter("PID", PID));
                checkCmd.Parameters.Add(new SqlParameter("CITY", CITY));
                DataTable dtCheck = BASE_DAO.SqlDao.CMEPMSQuery4List(checkCmd);
                bool exists = (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0);

                if(exists) {
                    // 更新模式 - 更新 PJDCTSP.SUP_RATIO
                    string updateSql = @"UPDATE CMEPMS.dbo.PJDCTSP SET
                        SUP_RATIO = @SUP_RATIO,
                        UPD_DT = @UPD_DT,
                        UPD_TM = @UPD_TM,
                        UPD_UID = @UPD_UID
                        WHERE PID = @PID AND CITY = @CITY";

                    SqlCommand cmd = new SqlCommand(updateSql);
                    cmd.Parameters.Add(new SqlParameter("PID", PID));
                    cmd.Parameters.Add(new SqlParameter("CITY", CITY));
                    cmd.Parameters.Add(new SqlParameter("SUP_RATIO", (object)CEN_SP_RATIO_S ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now.ToString("yyyyMMd")));
                    cmd.Parameters.Add(new SqlParameter("UPD_TM", DateTime.Now.ToString("HHmmss")));
                    cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                    BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
                    TempData["Success"] = "中央款比例(標準) 更新成功";
                } else {
                    // 新增模式 - 新增到 PJDCTSP
                    string insertSql = @"INSERT INTO CMEPMS.dbo.PJDCTSP
                        (PID, CITY, SUP_RATIO, ADD_DT, UPD_DT, UPD_TM, UPD_UID)
                        VALUES (@PID, @CITY, @SUP_RATIO, @ADD_DT, @UPD_DT, @UPD_TM, @UPD_UID)";

                    SqlCommand cmd = new SqlCommand(insertSql);
                    cmd.Parameters.Add(new SqlParameter("PID", PID));
                    cmd.Parameters.Add(new SqlParameter("CITY", CITY));
                    cmd.Parameters.Add(new SqlParameter("SUP_RATIO", (object)CEN_SP_RATIO_S ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now.ToString("yyyyMMdd")));
                    cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now.ToString("yyyyMMd")));
                    cmd.Parameters.Add(new SqlParameter("UPD_TM", DateTime.Now.ToString("HHmmss")));
                    cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                    BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
                    TempData["Success"] = "中央款比例(標準) 新增成功";
                }
            } catch(Exception ex) {
                System.Diagnostics.Debug.WriteLine($"PAGE2_1_2_Save error: {ex.Message}");
                string fullError = ex.Message;
                if(ex.InnerException != null) {
                    fullError += " | 內部錯誤: " + ex.InnerException.Message;
                }
                TempData["Error"] = "儲存失敗：" + fullError;
            }

            // 重新導向到 PAGE2_1_2
            return RedirectToAction("PAGE2_1_2", new { TID = TID, PID = PID, TD_CODE = TD_CODE });
        }

        // GET: PAGE2_1_2_Fund - 經費明細新增/修改視窗
        public ActionResult PAGE2_1_2_Fund(string FID, string TID, string PID, string TD_CODE, string mode) {
            ViewBag.TID = TID;
            ViewBag.PID = PID;
            ViewBag.TD_CODE = TD_CODE;
            ViewBag.IsNew = (mode == "add" || string.IsNullOrEmpty(FID));

        DataTable dt = null;

          // 修改模式：查詢現有資料（從 CTDITFD 表）
          if(!string.IsNullOrEmpty(FID) && mode == "edit") {
               string sql = @"SELECT F.FID, F.TID, F.SYR, F.FIT_CODE, F.FUND, F.FUND_DESC,
      (SELECT CODE_DESC FROM CMEPMS.dbo.SMCCODE WHERE SYS_ID = 'FD' AND CATY_ID = '05' AND CATY_COD = F.FIT_CODE) AS FIT_DESC
  FROM CMEPMS.dbo.CTDITFD F
         WHERE F.FID = @FID";
             SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("FID", FID));
     dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            }

          if(dt != null && dt.Rows.Count > 0) {
            ViewBag.Data = dt.Rows[0];
           ViewBag.FID = dt.Rows[0]["FID"].ToString();
  } else {
         ViewBag.Data = null;
             ViewBag.FID = "";
        }

       // 取得經費層級下拉選單
  ViewBag.FIT_CODE_List = tools.ToSelectList(util.GetCMEPMSCODES("FD", "05"), "id", "name", "請選擇");

            return View("PAGE2_1_2_Fund");
        }

   // POST: PAGE2_1_2_Fund_Save - 儲存經費明細
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE2_1_2_Fund_Save(string FID, string TID, string PID, string TD_CODE,
        string SYR, string FIT_CODE, decimal? FUND, string FUND_DESC, bool IsNew = false) {

    try {
          if(IsNew) {
     // 新增模式 - 新增到 CTDITFD
    string insertSql = @"INSERT INTO CMEPMS.dbo.CTDITFD 
   (TID, SYR, FIT_CODE, FUND, FUND_DESC, ADD_DT, UPD_DT, UPD_UID)
    VALUES (@TID, @SYR, @FIT_CODE, @FUND, @FUND_DESC, @ADD_DT, @UPD_DT, @UPD_UID)";

              SqlCommand cmd = new SqlCommand(insertSql);
      cmd.Parameters.Add(new SqlParameter("TID", (object)TID ?? DBNull.Value));
    cmd.Parameters.Add(new SqlParameter("SYR", (object)SYR ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("FIT_CODE", (object)FIT_CODE ?? DBNull.Value));
     cmd.Parameters.Add(new SqlParameter("FUND", (object)FUND ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("FUND_DESC", (object)FUND_DESC ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now));
             cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
        cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

          BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
           TempData["Success"] = "資料新增成功";
      } else {
              // 更新模式 - 更新 CTDITFD
        string updateSql = @"UPDATE CMEPMS.dbo.CTDITFD SET 
  SYR = @SYR, 
      FIT_CODE = @FIT_CODE, 
             FUND = @FUND, 
  FUND_DESC = @FUND_DESC,
    UPD_DT = @UPD_DT, 
               UPD_UID = @UPD_UID
 WHERE FID = @FID";

SqlCommand cmd = new SqlCommand(updateSql);
  cmd.Parameters.Add(new SqlParameter("FID", FID));
     cmd.Parameters.Add(new SqlParameter("SYR", (object)SYR ?? DBNull.Value));
             cmd.Parameters.Add(new SqlParameter("FIT_CODE", (object)FIT_CODE ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("FUND", (object)FUND ?? DBNull.Value));
        cmd.Parameters.Add(new SqlParameter("FUND_DESC", (object)FUND_DESC ?? DBNull.Value));
           cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
       cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

               BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
         TempData["Success"] = "資料更新成功";
 }
            } catch(Exception ex) {
            TempData["Error"] = (IsNew ? "資料新增失敗：" : "資料更新失敗：") + ex.Message;
       return RedirectToAction("PAGE2_1_2_Fund", new { FID = FID, TID = TID, PID = PID, TD_CODE = TD_CODE, mode = IsNew ? "add" : "edit" });
          }

            // 儲存成功，關閉視窗並重新整理父視窗
   return Content(@"<script>
                if (window.opener && typeof window.opener.refreshFundList === 'function') {
        window.opener.refreshFundList();
    }
     window.close();
  </script>", "text/html");
        }

        // POST: PAGE2_1_2_Fund_Ajax - AJAX 新增/更新/刪除經費來源明細 (CTDITFD)
        [HttpPost]
        public ActionResult PAGE2_1_2_Fund_Ajax(string action, string TID, string FID,
            string SYR, string FIT_CODE, decimal? FUND, string FUND_DESC, decimal? EXP_FUND) {
            try {
                if(action == "add") {
                    // 新增模式 - 新增到 CTDITFD
                    // FID 需要手動計算（取該 TID 下最大 FID + 1）
                    string maxFidSql = @"SELECT ISNULL(MAX(FID), 0) + 1 FROM CMEPMS.dbo.CTDITFD WHERE TID = @TID";
                    SqlCommand maxFidCmd = new SqlCommand(maxFidSql);
                    maxFidCmd.Parameters.Add(new SqlParameter("TID", TID));
                    DataTable dtMaxFid = BASE_DAO.SqlDao.CMEPMSQuery4List(maxFidCmd);
                    string newFID = dtMaxFid.Rows[0][0].ToString();

                    string insertSql = @"INSERT INTO CMEPMS.dbo.CTDITFD 
                        (TID, FID, SYR, FIT_CODE, FUND, FUND_DESC, EXP_FUND, ADD_DT, UPD_DT, UPD_UID)
                        VALUES (@TID, @FID, @SYR, @FIT_CODE, @FUND, @FUND_DESC, @EXP_FUND, @ADD_DT, @UPD_DT, @UPD_UID)";

                    SqlCommand cmd = new SqlCommand(insertSql);
                    cmd.Parameters.Add(new SqlParameter("TID", TID));
                    cmd.Parameters.Add(new SqlParameter("FID", newFID));
                    cmd.Parameters.Add(new SqlParameter("SYR", (object)SYR ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("FIT_CODE", (object)FIT_CODE ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("FUND", (object)FUND ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("FUND_DESC", (object)FUND_DESC ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("EXP_FUND", (object)EXP_FUND ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now));
                    cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                    cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                    BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
                    return Json(new { success = true, message = "新增成功", FID = newFID });

                } else if(action == "update") {
                    // 更新模式 - 更新 CTDITFD (複合主鍵: TID + FID)
                    string updateSql = @"UPDATE CMEPMS.dbo.CTDITFD SET 
                        SYR = @SYR, 
                        FIT_CODE = @FIT_CODE, 
                        FUND = @FUND, 
                        FUND_DESC = @FUND_DESC,
                        EXP_FUND = @EXP_FUND,
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
                    cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                    cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                    BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
                    return Json(new { success = true, message = "更新成功" });

                } else if(action == "delete") {
                    // 刪除模式 (複合主鍵: TID + FID)
                    string deleteSql = @"DELETE FROM CMEPMS.dbo.CTDITFD WHERE TID = @TID AND FID = @FID";
                    SqlCommand cmd = new SqlCommand(deleteSql);
                    cmd.Parameters.Add(new SqlParameter("TID", TID));
                    cmd.Parameters.Add(new SqlParameter("FID", FID));
                    BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
                    return Json(new { success = true, message = "刪除成功" });
                }

                return Json(new { success = false, message = "無效的操作" });

            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // POST: PAGE2_1_2_Month_Ajax - 3.1.4rev 儲存單筆經費明細的月分配數 (CMEPMS.dbo.CTDITFD)
        // 對應舊系統 App_Data\CMETMS\TDE0020B020.Xml 的 M_01~M_12 區塊
        [HttpPost]
        public ActionResult PAGE2_1_2_Month_Ajax(string TID, string FID, decimal?[] M) {
            try {
                if(string.IsNullOrEmpty(TID) || string.IsNullOrEmpty(FID)) {
                    return Json(new { success = false, message = "TID / FID 不可為空" });
                }
                if(M == null || M.Length != 12) {
                    return Json(new { success = false, message = "月份資料需為 12 個月" });
                }

                string updateSql = @"UPDATE CMEPMS.dbo.CTDITFD SET
                        M_01=@M_01, M_02=@M_02, M_03=@M_03, M_04=@M_04, M_05=@M_05, M_06=@M_06,
                        M_07=@M_07, M_08=@M_08, M_09=@M_09, M_10=@M_10, M_11=@M_11, M_12=@M_12,
                        UPD_DT=@UPD_DT, UPD_UID=@UPD_UID
                    WHERE TID = @TID AND FID = @FID";

                SqlCommand cmd = new SqlCommand(updateSql);
                cmd.Parameters.Add(new SqlParameter("TID", TID));
                cmd.Parameters.Add(new SqlParameter("FID", FID));
                for(int i = 0; i < 12; i++) {
                    cmd.Parameters.Add(new SqlParameter("M_" + (i + 1).ToString("00"), (object)M[i] ?? DBNull.Value));
                }
                cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);

                decimal total = 0;
                foreach(decimal? v in M) { total += v ?? 0; }
                return Json(new { success = true, message = "月分配數儲存成功", total = total });

            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }


   // GET: PAGE2_2_2 - 保留款月份分配（新增/編輯共用介面）
        public ActionResult PAGE2_2_2(string TD_CODE) {
            DataTable dt = null;

            if(!string.IsNullOrEmpty(TD_CODE)) {
                // 直接查詢 CRDMNS Table（不使用 View，避免 JOIN 條件導致查不到剛新增的資料）
                string sql = @"SELECT * FROM CMEPMS.dbo.CRDMNDS WHERE TD_CODE = @TD_CODE";
                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            }

            ViewBag.TD_CODE = TD_CODE;
            ViewBag.IsNew = (dt == null || dt.Rows.Count == 0);

            if(dt != null && dt.Rows.Count > 0) {
                // 有資料，帶出編輯
                ViewBag.Data = dt.Rows[0];
            } else {
                // 無資料，建立空白資料供新增
                ViewBag.Data = null;
            }

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE
            };

            return View("PAGE2_2_2", model);
        }

        // POST: PAGE2_2_2 - 儲存 CRDMDS 保留款月份分配資料
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE2_2_2_Save(string TD_CODE, string PID,
    decimal? REV_FUND, decimal? OCR_FUND, decimal? SPD_FUND, string REV_RESN,
       decimal? M_01, decimal? M_02, decimal? M_03, decimal? M_04,
            decimal? M_05, decimal? M_06, decimal? M_07, decimal? M_08,
            decimal? M_09, decimal? M_10, decimal? M_11, decimal? M_12,
            bool IsNew = false) {

            try {
                if(IsNew) {
                    // 新增模式 - TID 是 IDENTITY 自動遞增，不需要指定
                    string insertSql = @"INSERT INTO CMEPMS.dbo.CRDMNDS 
      (TD_CODE, REV_FUND, OCR_FUND, SPD_FUND, REV_RESN, 
             M_01, M_02, M_03, M_04, M_05, M_06, M_07, M_08, M_09, M_10, M_11, M_12, 
   ADD_DT, UPD_DT, UPD_UID)
           VALUES (@TD_CODE, @REV_FUND, @OCR_FUND, @SPD_FUND, @REV_RESN,
       @M_01, @M_02, @M_03, @M_04, @M_05, @M_06, @M_07, @M_08, @M_09, @M_10, @M_11, @M_12,
   @ADD_DT, @UPD_DT, @UPD_UID)";

                    SqlCommand cmd = new SqlCommand(insertSql);
                    cmd.Parameters.Add(new SqlParameter("TD_CODE", (object)TD_CODE ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("PID", (object)PID ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("REV_FUND", (object)REV_FUND ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("OCR_FUND", (object)OCR_FUND ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("SPD_FUND", (object)SPD_FUND ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("REV_RESN", (object)REV_RESN ?? DBNull.Value));
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
                    cmd.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now));
                    cmd.Parameters.Add(new SqlParameter("UPD_DT", DateTime.Now));
                    cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));

                    BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
                    TempData["Success"] = "資料新增成功";
                } else {
                    // 更新模式 - 用 TD_CODE 作為條件
                    string updateSql = @"UPDATE CMEPMS.dbo.CRDMNDS SET 
    REV_FUND = @REV_FUND,
              OCR_FUND = @OCR_FUND,
               SPD_FUND = @SPD_FUND,
      REV_RESN = @REV_RESN,
        M_01 = @M_01, M_02 = @M_02, M_03 = @M_03, M_04 = @M_04,
 M_05 = @M_05, M_06 = @M_06, M_07 = @M_07, M_08 = @M_08,
       M_09 = @M_09, M_10 = @M_10, M_11 = @M_11, M_12 = @M_12,
        UPD_DT = @UPD_DT, UPD_UID = @UPD_UID
           WHERE TD_CODE = @TD_CODE";

                    SqlCommand cmd = new SqlCommand(updateSql);
                    cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                    cmd.Parameters.Add(new SqlParameter("REV_FUND", (object)REV_FUND ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("OCR_FUND", (object)OCR_FUND ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("SPD_FUND", (object)SPD_FUND ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("REV_RESN", (object)REV_RESN ?? DBNull.Value));
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

                    BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
                    TempData["Success"] = "資料更新成功";
                }
            } catch(Exception ex) {
                TempData["Error"] = (IsNew ? "資料新增失敗：" : "資料更新失敗：") + ex.Message;
            }

            // 重新導向到 PAGE2_2，會重新查詢資料（新增後就有資料了，會變成編輯模式）
            return RedirectToAction("PAGE2_2_2", new { TD_CODE = TD_CODE });
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