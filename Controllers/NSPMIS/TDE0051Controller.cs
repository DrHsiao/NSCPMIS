using PagedList;

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

using WCals.ViewModels.NewModels;

using static WCals.classes.john.BASE_DAO;



namespace WCals.Controllers.NSPMIS {

    /// <summary>

    /// TDE0051 工程進度監控 Controller

    /// 5個頁籤：開/完工日填報不全、進度未填(前月)、進度異常(前月)、工程相關經費未填、連續三個月落後

    /// </summary>

    public class TDE0051Controller : Controller {



        private readonly Models.CMEPMSModels.CMEPMSEntities db_pms = new Models.CMEPMSModels.CMEPMSEntities();
        private readonly int pageSize = 20;



        // GET: TDE0051

        public ActionResult Index(TDE0051Model model) {

            if (model == null) {

                model = new TDE0051Model();

            }



            // 預設頁籤為1

            if (model.tab == 0) {

                model.tab = 1;

            }



            // 設定縣市鄉鎮下拉選單

            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);

            ViewBag.towns = tools.ToSelectList(util.GetTownNew(model.CITY), "TOWN", "TOWN", null);

            // 工作計畫-分支計畫
            ViewBag.pids = db_pms.PJV_PJDPCAT.OrderByDescending(w => w.SYR).Select(c => new SelectListItem {
                Value = c.PID.ToString(),
                Text = c.SYR_PJ_NAME
            }).ToList();

            // 預設年度計畫下拉選單為空，等使用者選擇年度後再動態載入
            ViewBag.emptyPIDs = new SelectList(new List<SelectListItem>(), "Value", "Text");

            // 取得所有年度計畫的年度
            SqlCommand syrCmd = new SqlCommand("SELECT PID_SYR as SYR FROM TDV_TDDCFLS_Lite GROUP BY PID_SYR ORDER BY PID_SYR DESC");
            DataTable syrs = BASE_DAO.SqlDao.CMEPMSQuery4List(syrCmd);
            ViewBag.SYRS = tools.ToSelectList(syrs, "SYR", "SYR", null);



            // 根據當前頁籤載入對應資料

            switch (model.tab) {

                case 1:

                    LoadTab1Data(model);

                    break;

                case 2:

                    LoadTab2Data(model);

                    break;

                case 3:

                    LoadTab3Data(model);

                    break;

                case 4:

                    LoadTab4Data(model);

                    break;

                case 5:

                    LoadTab5Data(model);

                    break;

            }



            return View("TDE0051Index", model);

        }



        /// <summary>

        /// 頁籤1: 開/完工日填報不全

        /// SQL: select SIP_TD_CODE as S1_HLK,* from TDV_TDDCNYM

        /// </summary>

        private void LoadTab1Data(TDE0051Model model) {

            if (model.page1 == 0) model.page1 = 1;



            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT SIP_TD_CODE as S1_HLK, * FROM TDV_TDDCNYM WHERE 1=1 ");



            SqlCommand cmd = new SqlCommand();

            AppendCommonFilters(sb, cmd, model);



            sb.Append("ORDER BY SIP_TD_CODE");

            cmd.CommandText = sb.ToString();



            DataTable dt = SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = dt.AsEnumerable().ToList();

            model.List_TDDCNYM = new PagedList<DataRow>(list, model.page1, pageSize);

        }



        /// <summary>

        /// 頁籤2: 進度未填(前月)

        /// SQL: select SIP_TD_CODE as S1_HLK,* from TDV_TDDCNSD

        /// </summary>

        private void LoadTab2Data(TDE0051Model model) {

            if (model.page2 == 0) model.page2 = 1;



            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT SIP_TD_CODE as S1_HLK, * FROM TDV_TDDCNSD WHERE 1=1 ");



            SqlCommand cmd = new SqlCommand();

            AppendCommonFilters(sb, cmd, model);



            sb.Append("ORDER BY SIP_TD_CODE");

            cmd.CommandText = sb.ToString();



            DataTable dt = SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = dt.AsEnumerable().ToList();

            model.List_TDDCNSD = new PagedList<DataRow>(list, model.page2, pageSize);

        }



        /// <summary>

        /// 頁籤3: 進度異常(前月)

        /// SQL: select SIP_TD_CODE as S1_HLK,* from TDV_TDDCNDL

        /// </summary>

        private void LoadTab3Data(TDE0051Model model) {

            if (model.page3 == 0) model.page3 = 1;



            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT SIP_TD_CODE as S1_HLK, * FROM TDV_TDDCNDL WHERE 1=1 ");



            SqlCommand cmd = new SqlCommand();

            AppendCommonFilters(sb, cmd, model);



            sb.Append("ORDER BY SIP_TD_CODE");

            cmd.CommandText = sb.ToString();



            DataTable dt = SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = dt.AsEnumerable().ToList();

            model.List_TDDCN = new PagedList<DataRow>(list, model.page3, pageSize);

        }



        /// <summary>

        /// 頁籤4: 工程相關經費未填

        /// SQL: select SIP_TD_CODE as S1_HLK,* from TDV_TDDCNFD

        /// </summary>

        private void LoadTab4Data(TDE0051Model model) {

            if (model.page4 == 0) model.page4 = 1;



            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT SIP_TD_CODE as S1_HLK, * FROM TDV_TDDCNFD WHERE 1=1 ");



            SqlCommand cmd = new SqlCommand();

            AppendCommonFilters(sb, cmd, model);



            sb.Append("ORDER BY SIP_TD_CODE");

            cmd.CommandText = sb.ToString();



            DataTable dt = SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = dt.AsEnumerable().ToList();

            model.List_TDDCNFD = new PagedList<DataRow>(list, model.page4, pageSize);

        }



        /// <summary>

        /// 頁籤5: 連續三個月落後

        /// SQL: select SIP_TD_CODE as S1_HLK,* from CMETMS.dbo.TDV_TDDMSCD

        /// </summary>

        private void LoadTab5Data(TDE0051Model model) {

            if (model.page5 == 0) model.page5 = 1;



            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT SIP_TD_CODE as S1_HLK, * FROM TDV_TDDMSCD WHERE 1=1 ");



            SqlCommand cmd = new SqlCommand();

            AppendCommonFilters(sb, cmd, model);



            sb.Append("ORDER BY SIP_TD_CODE");

            cmd.CommandText = sb.ToString();



            DataTable dt = SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = dt.AsEnumerable().ToList();

            model.List_TDDMSCD = new PagedList<DataRow>(list, model.page5, pageSize);

        }



        /// <summary>

        /// 共用篩選條件

        /// </summary>

        private void AppendCommonFilters(StringBuilder sb, SqlCommand cmd, TDE0051Model model) {
            // 年度計畫查詢條件：僅取有掛在該計畫下的工程
            if (model.PID > 0) {
                cmd.Parameters.Add(new SqlParameter("PID", model.PID));
                sb.Append("AND EXISTS (SELECT 1 FROM CMEPMS.dbo.CTDMNDS c WHERE c.TD_CODE = TD_CODE AND c.PID = @PID) ");
            }


            if (!string.IsNullOrEmpty(model.TD_CODE)) {

                cmd.Parameters.Add(new SqlParameter("TD_CODE", model.TD_CODE));

                sb.Append("AND TD_CODE = @TD_CODE ");

            }



            if (!string.IsNullOrEmpty(model.TD_NAME)) {

                cmd.Parameters.Add(new SqlParameter("TD_NAME", "%" + model.TD_NAME + "%"));

                sb.Append("AND TD_NAME LIKE @TD_NAME ");

            }



            if (!string.IsNullOrEmpty(model.CITY)) {

                cmd.Parameters.Add(new SqlParameter("CITY", model.CITY));

                sb.Append("AND TD_CITY = @CITY ");

            }



            /*if (!string.IsNullOrEmpty(model.TOWN)) {

                cmd.Parameters.Add(new SqlParameter("TOWN", model.TOWN));

                sb.Append("AND ADM_DIST = @TOWN ");

            }*/



            if (!string.IsNullOrEmpty(model.IPJ_CODE)) {

                cmd.Parameters.Add(new SqlParameter("IPJ_CODE", model.IPJ_CODE));

                sb.Append("AND IPJ_CODE = @IPJ_CODE ");

            }



            if (!string.IsNullOrEmpty(model.CATY_CODE)) {

                cmd.Parameters.Add(new SqlParameter("CATY_CODE", model.CATY_CODE));

                sb.Append("AND CATY_CODE = @CATY_CODE ");

            }

        }



        /// <summary>

        /// 切換頁籤

        /// </summary>

        public ActionResult ChangeTab(TDE0051Model model, int tab) {

            model.tab = tab;

            return Index(model);

        }



        /// <summary>

        /// 搜尋 (POST)

        /// </summary>

        [HttpPost]

        public ActionResult Search(TDE0051Model model) {

            return Index(model);

        }



        /// <summary>

        /// 搜尋 (GET) - 用於分頁連結

        /// </summary>

        public ActionResult Search(int tab = 1, string TD_CODE = null, string TD_NAME = null,

            string CITY = null, string TOWN = null, decimal PID = 0, string IPJ_CODE = null, string CATY_CODE = null,

            int page1 = 1, int page2 = 1, int page3 = 1, int page4 = 1, int page5 = 1) {



            TDE0051Model model = new TDE0051Model {

                PID = PID,

                tab = tab,

                TD_CODE = TD_CODE,

                TD_NAME = TD_NAME,

                CITY = CITY,

                TOWN = TOWN,

                IPJ_CODE = IPJ_CODE,

                CATY_CODE = CATY_CODE,

                page1 = page1,

                page2 = page2,

                page3 = page3,

                page4 = page4,

                page5 = page5

            };



            return Index(model);

        }



        /// <summary>

        /// Tab 頁面 (供 iframe 載入)

        /// </summary>

        public ActionResult Tab(int tab = 1, string TD_CODE = null, string TD_NAME = null,

            string CITY = null, string TOWN = null, decimal PID = 0, string IPJ_CODE = null, string CATY_CODE = null,

            int page1 = 1, int page2 = 1, int page3 = 1, int page4 = 1, int page5 = 1) {



            TDE0051Model model = new TDE0051Model {

                PID = PID,

                tab = tab,

                TD_CODE = TD_CODE,

                TD_NAME = TD_NAME,

                CITY = CITY,

                TOWN = TOWN,

                IPJ_CODE = IPJ_CODE,

                CATY_CODE = CATY_CODE,

                page1 = page1,

                page2 = page2,

                page3 = page3,

                page4 = page4,

                page5 = page5

            };



            // 根據當前頁籤載入對應資料

            switch (tab) {

                case 1:

                    LoadTab1Data(model);

                    break;

                case 2:

                    LoadTab2Data(model);

                    break;

                case 3:

                    LoadTab3Data(model);

                    break;

                case 4:

                    LoadTab4Data(model);

                    break;

                case 5:

                    LoadTab5Data(model);

                    break;

            }



            return View("TDE0051Tab", model);

        }



        /// <summary>

        /// 編輯頁面 - 開/完工日填報 (頁籤1)

        /// </summary>

        public ActionResult Edit(string TD_CODE) {

            if (string.IsNullOrEmpty(TD_CODE)) {

                return Content("TD_CODE 不可為空");

            }



            TDE0051Model model = new TDE0051Model {

                TD_CODE = TD_CODE

            };



            // 查詢工程名稱

            SqlCommand cmdName = new SqlCommand("SELECT TD_NAME FROM TDMFORM WHERE TD_CODE = @TD_CODE");

            cmdName.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            DataTable dtName = SqlDao.CMETMSQuery4List(cmdName);

            if (dtName.Rows.Count > 0) {

                model.TD_NAME = dtName.Rows[0]["TD_NAME"]?.ToString();

            }



            // 查詢 TDDSDWB 資料 (TK106, TK107)

            SqlCommand cmdList = new SqlCommand(@"SELECT IT_CODE, PRE_EDT, REL_EDT

                FROM TDDSDWB WHERE TD_CODE = @TD_CODE AND IT_TYPE = 'M' AND IT_CODE IN ('TK106', 'TK107')

                ORDER BY IT_CODE");

            cmdList.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            ViewBag.List1 = SqlDao.CMETMSQuery4List(cmdList);



            return View("TDE0051Edit", model);

        }



        /// <summary>

        /// AJAX 新增 TDDSDWB 資料

        /// </summary>

        [HttpPost]

        public JsonResult EditAdd(string TD_CODE, string IT_CODE, string PRE_EDT, string REL_EDT) {

            try {

                if (string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(IT_CODE)) {

                    return Json(new { success = false, message = "TD_CODE 和 IT_CODE 不可為空" });

                }



                // 檢查是否已存在

                SqlCommand cmdCheck = new SqlCommand("SELECT COUNT(*) as CNT FROM TDDSDWB WHERE TD_CODE = @TD_CODE AND IT_CODE = @IT_CODE AND IT_TYPE = 'M'");

                cmdCheck.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmdCheck.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));

                DataTable dtCheck = SqlDao.CMETMSQuery4List(cmdCheck);

                if (dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0]["CNT"]) > 0) {

                    return Json(new { success = false, message = "此項目代碼已存在" });

                }



                DateTime now = DateTime.Now;

                string userId = WebSiteHelper.getUserId();



                SqlCommand cmd = new SqlCommand(@"INSERT INTO TDDSDWB (TD_CODE, IT_CODE, IT_TYPE, PRE_EDT, REL_EDT, ADD_DT, UPD_DT, UPD_UID)

                    VALUES (@TD_CODE, @IT_CODE, 'M', @PRE_EDT, @REL_EDT, @ADD_DT, @UPD_DT, @UPD_UID)");

                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmd.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));

                cmd.Parameters.Add(new SqlParameter("PRE_EDT", string.IsNullOrEmpty(PRE_EDT) ? (object)DBNull.Value : PRE_EDT));

                cmd.Parameters.Add(new SqlParameter("REL_EDT", string.IsNullOrEmpty(REL_EDT) ? (object)DBNull.Value : REL_EDT));

                cmd.Parameters.Add(new SqlParameter("ADD_DT", now));

                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));

                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId ?? ""));



                SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "新增成功" });

            } catch (Exception ex) {

                return Json(new { success = false, message = "新增失敗：" + ex.Message });

            }

        }



        /// <summary>

        /// AJAX 更新 TDDSDWB 資料

        /// </summary>

        [HttpPost]

        public JsonResult EditUpdate(string TD_CODE, string IT_CODE, string PRE_EDT, string REL_EDT) {

            try {

                if (string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(IT_CODE)) {

                    return Json(new { success = false, message = "TD_CODE 和 IT_CODE 不可為空" });

                }



                DateTime now = DateTime.Now;

                string userId = WebSiteHelper.getUserId();



                SqlCommand cmd = new SqlCommand(@"UPDATE TDDSDWB SET

                    PRE_EDT = @PRE_EDT,

                    REL_EDT = @REL_EDT,

                    UPD_DT = @UPD_DT,

                    UPD_UID = @UPD_UID

                    WHERE TD_CODE = @TD_CODE AND IT_CODE = @IT_CODE AND IT_TYPE = 'M'");

                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmd.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));

                cmd.Parameters.Add(new SqlParameter("PRE_EDT", string.IsNullOrEmpty(PRE_EDT) ? (object)DBNull.Value : PRE_EDT));

                cmd.Parameters.Add(new SqlParameter("REL_EDT", string.IsNullOrEmpty(REL_EDT) ? (object)DBNull.Value : REL_EDT));

                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));

                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId ?? ""));



                SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });

            } catch (Exception ex) {

                return Json(new { success = false, message = "更新失敗：" + ex.Message });

            }

        }



        /// <summary>

        /// AJAX 刪除 TDDSDWB 資料

        /// </summary>

        [HttpPost]

        public JsonResult EditDelete(string TD_CODE, string IT_CODE) {

            try {

                if (string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(IT_CODE)) {

                    return Json(new { success = false, message = "TD_CODE 和 IT_CODE 不可為空" });

                }



                SqlCommand cmd = new SqlCommand("DELETE FROM TDDSDWB WHERE TD_CODE = @TD_CODE AND IT_CODE = @IT_CODE AND IT_TYPE = 'M'");

                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmd.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));



                SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });

            } catch (Exception ex) {

                return Json(new { success = false, message = "刪除失敗：" + ex.Message });

            }

        }



        #region 頁籤2 - 進度未填 (TDDMSCD)



        /// <summary>

        /// 編輯頁面2 - 進度填報 (頁籤2)

        /// </summary>

        public ActionResult Edit2(string TD_CODE) {

            if (string.IsNullOrEmpty(TD_CODE)) {

                return Content("TD_CODE 不可為空");

            }



            TDE0051Model model = new TDE0051Model {

                TD_CODE = TD_CODE

            };



            // 查詢工程名稱

            SqlCommand cmdName = new SqlCommand("SELECT TD_NAME FROM TDMFORM WHERE TD_CODE = @TD_CODE");

            cmdName.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            DataTable dtName = SqlDao.CMETMSQuery4List(cmdName);

            if (dtName.Rows.Count > 0) {

                model.TD_NAME = dtName.Rows[0]["TD_NAME"]?.ToString();

            }



            // 查詢 TDDMSCD 資料

            SqlCommand cmdList = new SqlCommand(@"SELECT SYR, SMN, PRE_SCHD, REL_SCHD, EXECT_DESC

                FROM TDDMSCD WHERE TD_CODE = @TD_CODE

                ORDER BY SYR DESC, SMN DESC");

            cmdList.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            ViewBag.List1 = SqlDao.CMETMSQuery4List(cmdList);



            return View("TDE0051Edit2", model);

        }



        /// <summary>

        /// AJAX 新增 TDDMSCD 資料

        /// </summary>

        [HttpPost]

        public JsonResult Edit2Add(string TD_CODE, string SYR, string SMN, decimal? PRE_SCHD, decimal? REL_SCHD, string EXECT_DESC) {

            try {

                if (string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN)) {

                    return Json(new { success = false, message = "TD_CODE、年度、月份不可為空" });

                }



                // 檢查是否已存在

                SqlCommand cmdCheck = new SqlCommand("SELECT COUNT(*) as CNT FROM TDDMSCD WHERE TD_CODE = @TD_CODE AND SYR = @SYR AND SMN = @SMN");

                cmdCheck.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmdCheck.Parameters.Add(new SqlParameter("SYR", SYR));

                cmdCheck.Parameters.Add(new SqlParameter("SMN", SMN));

                DataTable dtCheck = SqlDao.CMETMSQuery4List(cmdCheck);

                if (dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0]["CNT"]) > 0) {

                    return Json(new { success = false, message = "此年月資料已存在" });

                }



                DateTime now = DateTime.Now;

                string userId = WebSiteHelper.getUserId();



                SqlCommand cmd = new SqlCommand(@"INSERT INTO TDDMSCD (TD_CODE, SYR, SMN, PRE_SCHD, REL_SCHD, EXECT_DESC, ADD_DT, UPD_DT, UPD_UID)

                    VALUES (@TD_CODE, @SYR, @SMN, @PRE_SCHD, @REL_SCHD, @EXECT_DESC, @ADD_DT, @UPD_DT, @UPD_UID)");

                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmd.Parameters.Add(new SqlParameter("SYR", SYR));

                cmd.Parameters.Add(new SqlParameter("SMN", SMN));

                cmd.Parameters.Add(new SqlParameter("PRE_SCHD", PRE_SCHD.HasValue ? (object)PRE_SCHD.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("REL_SCHD", REL_SCHD.HasValue ? (object)REL_SCHD.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("EXECT_DESC", string.IsNullOrEmpty(EXECT_DESC) ? (object)DBNull.Value : EXECT_DESC));

                cmd.Parameters.Add(new SqlParameter("ADD_DT", now));

                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));

                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId ?? ""));



                SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "新增成功" });

            } catch (Exception ex) {

                return Json(new { success = false, message = "新增失敗：" + ex.Message });

            }

        }



        /// <summary>

        /// AJAX 更新 TDDMSCD 資料

        /// </summary>

        [HttpPost]

        public JsonResult Edit2Update(string TD_CODE, string SYR, string SMN, decimal? PRE_SCHD, decimal? REL_SCHD, string EXECT_DESC) {

            try {

                if (string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN)) {

                    return Json(new { success = false, message = "TD_CODE、年度、月份不可為空" });

                }



                DateTime now = DateTime.Now;

                string userId = WebSiteHelper.getUserId();



                SqlCommand cmd = new SqlCommand(@"UPDATE TDDMSCD SET

                    PRE_SCHD = @PRE_SCHD,

                    REL_SCHD = @REL_SCHD,

                    EXECT_DESC = @EXECT_DESC,

                    UPD_DT = @UPD_DT,

                    UPD_UID = @UPD_UID

                    WHERE TD_CODE = @TD_CODE AND SYR = @SYR AND SMN = @SMN");

                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmd.Parameters.Add(new SqlParameter("SYR", SYR));

                cmd.Parameters.Add(new SqlParameter("SMN", SMN));

                cmd.Parameters.Add(new SqlParameter("PRE_SCHD", PRE_SCHD.HasValue ? (object)PRE_SCHD.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("REL_SCHD", REL_SCHD.HasValue ? (object)REL_SCHD.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("EXECT_DESC", string.IsNullOrEmpty(EXECT_DESC) ? (object)DBNull.Value : EXECT_DESC));

                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));

                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId ?? ""));



                SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });

            } catch (Exception ex) {

                return Json(new { success = false, message = "更新失敗：" + ex.Message });

            }

        }



        /// <summary>

        /// AJAX 刪除 TDDMSCD 資料

        /// </summary>

        [HttpPost]

        public JsonResult Edit2Delete(string TD_CODE, string SYR, string SMN) {

            try {

                if (string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN)) {

                    return Json(new { success = false, message = "TD_CODE、年度、月份不可為空" });

                }



                SqlCommand cmd = new SqlCommand("DELETE FROM TDDMSCD WHERE TD_CODE = @TD_CODE AND SYR = @SYR AND SMN = @SMN");

                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmd.Parameters.Add(new SqlParameter("SYR", SYR));

                cmd.Parameters.Add(new SqlParameter("SMN", SMN));



                SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });

            } catch (Exception ex) {

                return Json(new { success = false, message = "刪除失敗：" + ex.Message });

            }

        }



        #endregion



        #region 頁籤3 - 進度異常 (TDDMSCD - 前月)



        /// <summary>

        /// 編輯頁面3 - 進度異常填報 (頁籤3)

        /// </summary>

        public ActionResult Edit3(string TD_CODE) {

            if (string.IsNullOrEmpty(TD_CODE)) {

                return Content("TD_CODE 不可為空");

            }



            TDE0051Model model = new TDE0051Model {

                TD_CODE = TD_CODE

            };



            // 查詢工程名稱

            SqlCommand cmdName = new SqlCommand("SELECT TD_NAME FROM TDMFORM WHERE TD_CODE = @TD_CODE");

            cmdName.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            DataTable dtName = SqlDao.CMETMSQuery4List(cmdName);

            if (dtName.Rows.Count > 0) {

                model.TD_NAME = dtName.Rows[0]["TD_NAME"]?.ToString();

            }



            // 查詢 TDDMSCD 前月資料

            SqlCommand cmdList = new SqlCommand(@"SELECT SYR, SMN, PRE_SCHD, REL_SCHD, DELAY_RESN, RESLV_MTHD

                FROM TDDMSCD

                WHERE TD_CODE = @TD_CODE

                AND SYR + SMN = SUBSTRING(dbo.getCHTDate(DATEADD(mm, -1, GETDATE()),''),1,5)");

            cmdList.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            ViewBag.List1 = SqlDao.CMETMSQuery4List(cmdList);



            return View("TDE0051Edit3", model);

        }



        /// <summary>

        /// AJAX 更新 TDDMSCD 進度異常資料

        /// </summary>

        [HttpPost]

        public JsonResult Edit3Update(string TD_CODE, string SYR, string SMN, decimal? PRE_SCHD, decimal? REL_SCHD, string DELAY_RESN, string RESLV_MTHD) {

            try {

                if (string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN)) {

                    return Json(new { success = false, message = "TD_CODE、年度、月份不可為空" });

                }



                DateTime now = DateTime.Now;

                string userId = WebSiteHelper.getUserId();



                SqlCommand cmd = new SqlCommand(@"UPDATE TDDMSCD SET

                    PRE_SCHD = @PRE_SCHD,

                    REL_SCHD = @REL_SCHD,

                    DELAY_RESN = @DELAY_RESN,

                    RESLV_MTHD = @RESLV_MTHD,

                    UPD_DT = @UPD_DT,

                    UPD_UID = @UPD_UID

                    WHERE TD_CODE = @TD_CODE AND SYR = @SYR AND SMN = @SMN");

                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmd.Parameters.Add(new SqlParameter("SYR", SYR));

                cmd.Parameters.Add(new SqlParameter("SMN", SMN));

                cmd.Parameters.Add(new SqlParameter("PRE_SCHD", PRE_SCHD.HasValue ? (object)PRE_SCHD.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("REL_SCHD", REL_SCHD.HasValue ? (object)REL_SCHD.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("DELAY_RESN", string.IsNullOrEmpty(DELAY_RESN) ? (object)DBNull.Value : DELAY_RESN));

                cmd.Parameters.Add(new SqlParameter("RESLV_MTHD", string.IsNullOrEmpty(RESLV_MTHD) ? (object)DBNull.Value : RESLV_MTHD));

                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));

                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId ?? ""));



                SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });

            } catch (Exception ex) {

                return Json(new { success = false, message = "更新失敗：" + ex.Message });

            }

        }



        #endregion



        #region 頁籤4 - 工程相關經費未填 (TDMFORM)



        /// <summary>

        /// 編輯頁面4 - 工程經費填報 (頁籤4)

        /// </summary>

        public ActionResult Edit4(string TD_CODE) {

            if (string.IsNullOrEmpty(TD_CODE)) {

                return Content("TD_CODE 不可為空");

            }



            TDE0051Model model = new TDE0051Model {

                TD_CODE = TD_CODE

            };



            // 查詢工程資料 (包含名稱、執行狀態、預算金額、發包金額)

            SqlCommand cmdList = new SqlCommand(@"SELECT TD_CODE, TD_NAME, TD_STAT_DESC, BUDGET_FUND, CTRCT_FUND

                FROM TDV_TDMFORM WHERE TD_CODE = @TD_CODE");

            cmdList.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            DataTable dtList = SqlDao.CMETMSQuery4List(cmdList);

            ViewBag.List1 = dtList;



            if (dtList.Rows.Count > 0) {

                model.TD_NAME = dtList.Rows[0]["TD_NAME"]?.ToString();

            }



            return View("TDE0051Edit4", model);

        }



        /// <summary>

        /// AJAX 更新 TDMFORM 經費資料

        /// </summary>

        [HttpPost]

        public JsonResult Edit4Update(string TD_CODE, decimal? BUDGET_FUND, decimal? CTRCT_FUND) {

            try {

                if (string.IsNullOrEmpty(TD_CODE)) {

                    return Json(new { success = false, message = "TD_CODE 不可為空" });

                }



                DateTime now = DateTime.Now;

                string userId = WebSiteHelper.getUserId();



                SqlCommand cmd = new SqlCommand(@"UPDATE TDMFORM SET

                    BUDGET_FUND = @BUDGET_FUND,

                    CTRCT_FUND = @CTRCT_FUND,

                    UPD_DT = @UPD_DT,

                    UPD_UID = @UPD_UID

                    WHERE TD_CODE = @TD_CODE");

                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmd.Parameters.Add(new SqlParameter("BUDGET_FUND", BUDGET_FUND.HasValue ? (object)BUDGET_FUND.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("CTRCT_FUND", CTRCT_FUND.HasValue ? (object)CTRCT_FUND.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));

                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId ?? ""));



                SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });

            } catch (Exception ex) {

                return Json(new { success = false, message = "更新失敗：" + ex.Message });

            }

        }



        #endregion



        #region 頁籤5 - 連續三個月落後 (TDDMSCD)



        /// <summary>

        /// 編輯頁面5 - 連續落後進度填報 (頁籤5)

        /// </summary>

        public ActionResult Edit5(string TD_CODE) {

            if (string.IsNullOrEmpty(TD_CODE)) {

                return Content("TD_CODE 不可為空");

            }



            TDE0051Model model = new TDE0051Model {

                TD_CODE = TD_CODE

            };



            // 查詢工程名稱

            SqlCommand cmdName = new SqlCommand("SELECT TD_NAME FROM TDMFORM WHERE TD_CODE = @TD_CODE");

            cmdName.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            DataTable dtName = SqlDao.CMETMSQuery4List(cmdName);

            if (dtName.Rows.Count > 0) {

                model.TD_NAME = dtName.Rows[0]["TD_NAME"]?.ToString();

            }



            // 查詢 TDDMSCD 連續落後資料 (最近3個月)

            SqlCommand cmdList = new SqlCommand(@"SELECT SYR, SMN, PRE_SCHD, REL_SCHD, DELAY_RESN, RESLV_MTHD

                FROM TDDMSCD

                WHERE TD_CODE = @TD_CODE

                ORDER BY SYR DESC, SMN DESC");

            cmdList.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

            ViewBag.List1 = SqlDao.CMETMSQuery4List(cmdList);



            return View("TDE0051Edit5", model);

        }



        /// <summary>

        /// AJAX 更新 TDDMSCD 連續落後進度資料

        /// </summary>

        [HttpPost]

        public JsonResult Edit5Update(string TD_CODE, string SYR, string SMN, decimal? PRE_SCHD, decimal? REL_SCHD, string DELAY_RESN, string RESLV_MTHD) {

            try {

                if (string.IsNullOrEmpty(TD_CODE) || string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN)) {

                    return Json(new { success = false, message = "TD_CODE、年度、月份不可為空" });

                }



                DateTime now = DateTime.Now;

                string userId = WebSiteHelper.getUserId();



                SqlCommand cmd = new SqlCommand(@"UPDATE TDDMSCD SET

                    PRE_SCHD = @PRE_SCHD,

                    REL_SCHD = @REL_SCHD,

                    DELAY_RESN = @DELAY_RESN,

                    RESLV_MTHD = @RESLV_MTHD,

                    UPD_DT = @UPD_DT,

                    UPD_UID = @UPD_UID

                    WHERE TD_CODE = @TD_CODE AND SYR = @SYR AND SMN = @SMN");

                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));

                cmd.Parameters.Add(new SqlParameter("SYR", SYR));

                cmd.Parameters.Add(new SqlParameter("SMN", SMN));

                cmd.Parameters.Add(new SqlParameter("PRE_SCHD", PRE_SCHD.HasValue ? (object)PRE_SCHD.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("REL_SCHD", REL_SCHD.HasValue ? (object)REL_SCHD.Value : DBNull.Value));

                cmd.Parameters.Add(new SqlParameter("DELAY_RESN", string.IsNullOrEmpty(DELAY_RESN) ? (object)DBNull.Value : DELAY_RESN));

                cmd.Parameters.Add(new SqlParameter("RESLV_MTHD", string.IsNullOrEmpty(RESLV_MTHD) ? (object)DBNull.Value : RESLV_MTHD));

                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));

                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId ?? ""));



                SqlDao.CMETMSExecuteSql(cmd);

                return Json(new { success = true, message = "更新成功" });

            } catch (Exception ex) {

                return Json(new { success = false, message = "更新失敗：" + ex.Message });

            }

        }



        #endregion



        /// <summary>

        /// 匯出 Excel

        /// </summary>

        public ActionResult ExportExcel(TDE0051Model model) {

            // TODO: 實作匯出 Excel 功能

            return Index(model);

        }

    }

}

