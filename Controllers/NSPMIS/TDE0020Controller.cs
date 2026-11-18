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

namespace WCals.Controllers.NSPMIS.TDE {
    public class TDE0020Controller : Controller {

        private readonly Models.CMETMSModels.CMETMSEntities db = new Models.CMETMSModels.CMETMSEntities();
        //private readonly Models.CMPSGSModels.CMPSGSEntities sgsdb = new Models.CMPSGSModels.CMPSGSEntities();
        private readonly Models.CMEPMSModels.CMEPMSEntities db_pms = new Models.CMEPMSModels.CMEPMSEntities();
        int pageSize = 5;

        // GET: TDE0020
        public ActionResult Index(ViewModels.NewModels.TDE0020Model model)
        {
            string uid = WebSiteHelper.getUserId();
            //model.pids = db_pms.PJMFORM.Where(w=>w.ADV_UID.Equals(uid)).OrderByDescending(w=>w.SYR).Select(c => new SelectListItem {
            //                       Value = c.PID.ToString(),
            //                       Text = "("+c.SYR+")"+c.PJ_NAME
            //                   }).ToList();

            // 1. 取得登入者縣市
            string loginCity = GetLoginUserCity();

            model.pids = db_pms.PJV_PJDPCAT.OrderByDescending(w => w.SYR).Select(c => new SelectListItem
            {
                Value = c.PID.ToString(),
                Text = c.SYR_PJ_NAME
            }).ToList();

            ViewBag.cities = tools.BuildCitySelectList(loginCity);

            // 2. 處理縣市下拉選單
            //if (!string.IsNullOrEmpty(loginCity))
            //{
            //    // 如果是縣市政府帳號，下拉選單只放入該縣市，並鎖定 model
            //    List<SelectListItem> singleCity = new List<SelectListItem>();
            //    singleCity.Add(new SelectListItem { Text = loginCity, Value = loginCity });
            //    ViewBag.cities = new SelectList(singleCity, "Value", "Text", loginCity);

            //    // 強制設定 model 預設值
            //    model.CITY = loginCity;
            //}
            //else
            //{
            //    // 中央帳號，顯示所有縣市
            //    ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);
            //}

            //ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);
            // 3.1.2 鄉鎮別與縣市別連動：縣市政府帳號的縣市下拉已鎖定且預選，onchange 不會觸發，
            // 故初始載入就依登入縣市取鄉鎮；中央帳號(loginCity 為 null)維持顯示全部鄉鎮，由 onchange 過濾。
            ViewBag.towns = tools.ToSelectList(util.GetTowns(loginCity), "TOWN", "TOWN", null);

            // 3.1.2 預設年度計畫下拉選單為空，等使用者選擇年度後再動態載入
            ViewBag.emptyPIDs = new SelectList(new List<SelectListItem>(), "Value", "Text");

            // 3.1.2 取得所有年度計畫的年度
            SqlCommand cmd = new SqlCommand("SELECT PID_SYR as SYR FROM TDV_TDDCFLS_Lite GROUP BY PID_SYR ORDER BY PID_SYR DESC");
            DataTable syrs = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.SYRS = tools.ToSelectList(syrs, "SYR", "SYR", null);

            return View("TDE0020Index", model);
        }

        public ActionResult Search(ViewModels.NewModels.TDE0020Model model) {
            

            return View("TDE0020Index", model);
        }

        /*
        url += "&town=" + 
            url += "&city=" + city
            url += "&td_code=" + td_code
            url += "&td_name=" + td_name
            url += "&pid=" + pid
            url += "&year=" + year
            url += "&month=" + month

        */

        DataTable GetPS(string town, string city, string td_code, string td_name, string pid, string year, string month, string type) {
            SqlCommand cmd = new SqlCommand();

            // [修正1] 把 SYR/SMN 條件移到 LEFT JOIN 的 ON 子句，避免笛卡爾積
            // [修正2] 使用子查詢處理 TDDMSCD 同一 TD_CODE+SYR+SMN 可能有多筆的問題，取 ID 最大的那筆
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT t.*, d.* FROM CMEPMS.dbo.TDV_TDDCFLS_Lite t ");
            sb.Append("LEFT OUTER JOIN ( ");
            sb.Append("    SELECT * FROM TDDMSCD d1 WHERE d1.ID = ( ");
            sb.Append("        SELECT MAX(d2.ID) FROM TDDMSCD d2 ");
            sb.Append("        WHERE d2.TD_CODE = d1.TD_CODE AND d2.SYR = d1.SYR AND d2.SMN = d1.SMN ");
            sb.Append("    ) ");
            if(!string.IsNullOrEmpty(year)) {
                cmd.Parameters.Add(new SqlParameter("SYR", year));
                sb.Append("    AND d1.SYR = @SYR ");
            }
            if(!string.IsNullOrEmpty(month)) {
                cmd.Parameters.Add(new SqlParameter("SMN", int.Parse(month).ToString("00")));
                sb.Append("    AND d1.SMN = @SMN ");
            }
            sb.Append(") d ON t.TD_CODE = d.TD_CODE ");
            sb.Append("WHERE t.CATY_CODE = @CATY_CODE AND t.TD_CODE IS NOT NULL ");

            // 限制為使用者所負責的 TD_CODE
            sb.Append("AND t.TD_CODE IN (SELECT TD_CODE FROM CMETMS.dbo.TDV_TDMADMC WHERE ADV_DP = @ADV_DP) ");

            cmd.Parameters.Add(new SqlParameter("ADV_DP", WebSiteHelper.getUserDept()));
            if(!string.IsNullOrEmpty(town)) {
                cmd.Parameters.Add(new SqlParameter("TOWN", town));
                sb.Append("AND t.TOWN=@TOWN ");
            }
            if(!string.IsNullOrEmpty(city)) {
                cmd.Parameters.Add(new SqlParameter("CITY", city));
                sb.Append("AND t.CITY=@CITY ");
            }
            if(!string.IsNullOrEmpty(td_name)) {
                // 修正：改為「包含」比對，原本 td_name+"%" 是前綴比對，使用者輸入名稱中間關鍵字會查無資料
                cmd.Parameters.Add(new SqlParameter("TD_NAME", "%" + td_name + "%"));
                sb.Append("AND t.TD_NAME LIKE @TD_NAME ");
            }
            if(!string.IsNullOrEmpty(td_code)) {
                cmd.Parameters.Add(new SqlParameter("TD_CODE", td_code));
                sb.Append("AND t.TD_CODE=@TD_CODE ");
            }
            if(!string.IsNullOrEmpty(pid)&&!pid.Equals("null")) {
                cmd.Parameters.Add(new SqlParameter("PID", pid));
                sb.Append("AND t.PID=@PID ");
            }
            //sb.Append(" ORDER BY t.ADD_DT DESC");
            cmd.Parameters.Add(new SqlParameter("CATY_CODE", type));
            cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            return ps;
        }
        string GenIncSql(string TD_CODE, string TD_TYPE) {
            return "SELECT * FROM TDDSDWB b LEFT OUTER JOIN SMWSDTP p ON p.IT_CODE=b.IT_CODE WHERE b.TD_CODE='" + TD_CODE + "'AND p.TD_TYPE='" + TD_TYPE + "' AND p.AN_IT_CODE IS NOT NULL ORDER BY p.GP_ORDER ";
        }

        // GET: TDE0020
        public ActionResult TAB1(string town, string city, string td_code, string td_name, string pid, string year, string month, string page, string type) {
            TDE0020Model model = new TDE0020Model();
            model.town = town;
            model.city = city;
            model.td_code = td_code;
            model.td_name = td_name;
            model.pid = pid;
            if(pid==null) {
                model.pid = null;
                pid = null;
            }
            model.year = year;
            model.month = month;
            model.page1 = Convert.ToInt32(page);
            model.type = type;

            int page1=1;
            if(!string.IsNullOrEmpty(page)) {
                page1= int.Parse(page);
            }

            //大表
            DataTable ps = GetPS(town, city, td_code, td_name, pid, year, month, type);
            DataTable drs;
            ViewBag.cols = BASE_DAO.SqlDao.CMETMSQuery4List("SELECT * FROM  SMWSDTP p WHERE p.TD_TYPE='"+type+"' AND p.AN_IT_CODE IS NOT NULL ORDER BY p.GP_ORDER;"); ;
            ps.Columns.Add("TDDSDWB", typeof(DataTable));

            // 預先建立所有工作控制點動態日期欄位(依 cols)，確保 view 逐一讀取每個 itCode 欄位時都存在(避免 Column 不存在例外)
            DataTable colsDt = (DataTable)ViewBag.cols;
            foreach(DataRow col in colsDt.Rows) {
                string it = col["IT_CODE"].ToString();
                if(!ps.Columns.Contains(it + "_PRE_SDT")) ps.Columns.Add(it + "_PRE_SDT", typeof(string));
                if(!ps.Columns.Contains(it + "_PRE_EDT")) ps.Columns.Add(it + "_PRE_EDT", typeof(string));
                if(!ps.Columns.Contains(it + "_REL_SDT")) ps.Columns.Add(it + "_REL_SDT", typeof(string));
                if(!ps.Columns.Contains(it + "_REL_EDT")) ps.Columns.Add(it + "_REL_EDT", typeof(string));
            }

            List<DataRow> list = ps.AsEnumerable().ToList();

            // [修正3] 過濾 TDV_TDDCFLS_Lite 的 TD_CODE 重複問題
            // 同一個 TD_CODE 可能有多個 PID(計畫)，只保留 PID_SYR(年度) 最大的那筆
            list = list
                .GroupBy(r => r["TD_CODE"]?.ToString())
                .Select(g => g.OrderByDescending(r => r["PID_SYR"]?.ToString()).First())
                .ToList();

            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page1, pageSize);

            // [效能修正] 只對「當前頁」的列查詢工作控制點日期(小表)。
            // 原本對 GetPS 全部結果逐列各發一次 GenIncSql，工程名稱查詢命中數千筆時會爆量查詢而逾時→查無結果。
            foreach(DataRow row in plist) {
                drs = BASE_DAO.SqlDao.CMETMSQuery4List(GenIncSql(row["TD_CODE"].ToString(), type));
                for(int j = 0; j < drs.Rows.Count; j++) {
                    string it = drs.Rows[j]["IT_CODE"].ToString();
                    if(!ps.Columns.Contains(it + "_PRE_SDT")) ps.Columns.Add(it + "_PRE_SDT", typeof(string));
                    if(!ps.Columns.Contains(it + "_PRE_EDT")) ps.Columns.Add(it + "_PRE_EDT", typeof(string));
                    if(!ps.Columns.Contains(it + "_REL_SDT")) ps.Columns.Add(it + "_REL_SDT", typeof(string));
                    if(!ps.Columns.Contains(it + "_REL_EDT")) ps.Columns.Add(it + "_REL_EDT", typeof(string));
                    row[it + "_PRE_SDT"] = drs.Rows[j]["PRE_SDT"].ToString();
                    row[it + "_PRE_EDT"] = drs.Rows[j]["PRE_EDT"].ToString();
                    row[it + "_REL_SDT"] = drs.Rows[j]["REL_SDT"].ToString();
                    row[it + "_REL_EDT"] = drs.Rows[j]["REL_EDT"].ToString();
                }
            }

            model.List1 = plist;
            return View("TDE0020IndexTab1", model);
        }

        public ActionResult TAB2(TDE0020Model model) {

            int page1 = (model.page1 == 0) ? 1 : model.page1;
            model.page1 = page1;
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List("SELECT*FROM TDMFORM t LEFT OUTER JOIN TDDMSCD d ON t.TD_CODE=d.TD_CODE WHERE t.ADV_UID='" + WebSiteHelper.getUserId() + "' ORDER BY t.ADD_DT DESC");
            DataTable drs;
            ViewBag.cols = BASE_DAO.SqlDao.CMETMSQuery4List("SELECT * FROM SMWSDTP p WHERE p.TD_TYPE='L' AND p.AN_IT_CODE IS NOT NULL ORDER BY p.GP_ORDER;"); ;
            ps.Columns.Add("TDDSDWB", typeof(DataTable));
            string sql;
            for(int i = 0; i < ps.Rows.Count; i++) {
                //sql = "SELECT b.* FROM TDDSDWB b WHERE b.TD_CODE='" + ps.Rows[i]["TD_CODE"] + "' AND AN_IT_CODE IS NOT NULL ORDER BY GP_ORDER";
                sql = "SELECT * FROM TDDSDWB b LEFT OUTER JOIN SMWSDTP p ON p.IT_CODE=b.IT_CODE WHERE b.TD_CODE='" + ps.Rows[i]["TD_CODE"] + "'AND p.TD_TYPE='L' AND p.AN_IT_CODE IS NOT NULL ORDER BY p.GP_ORDER ";
                drs = BASE_DAO.SqlDao.CMETMSQuery4List(sql);
                for(int j = 0; j < drs.Rows.Count; j++) {
                    if(!ps.Columns.Contains(drs.Rows[j]["IT_CODE"] + "_PRE_SDT")) {
                        ps.Columns.Add(drs.Rows[j]["IT_CODE"] + "_PRE_SDT", typeof(string));
                    }
                    if(!ps.Columns.Contains(drs.Rows[j]["IT_CODE"] + "_PRE_EDT")) {
                        ps.Columns.Add(drs.Rows[j]["IT_CODE"] + "_PRE_EDT", typeof(string));
                    }
                    ps.Rows[i][drs.Rows[j]["IT_CODE"] + "_PRE_SDT"] = drs.Rows[j]["PRE_SDT"].ToString();
                    ps.Rows[i][drs.Rows[j]["IT_CODE"] + "_PRE_EDT"] = drs.Rows[j]["PRE_EDT"].ToString();
                }
            }
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page1, pageSize);
            model.List1 = plist;
            return View("TDE0020IndexTab2", model);
        }

        public ActionResult TAB3() {


            return View("TDE0020IndexTab3");
        }

        public ActionResult TAB4() {


            return View("TDE0020IndexTab4");
        }

        public ActionResult TAB5() {


            return View("TDE0020IndexTab5");
        }

        public ActionResult TAB6() {


            return View("TDE0020IndexTab6");
        }

        public ActionResult TAB7() {


            return View("TDE0020IndexTab7");
        }

        /// <summary>
        /// AJAX 儲存日期欄位 (TDDSDWB 表)
        /// TD_CODE + IT_CODE 為識別鍵，REL_SDT/REL_EDT 為 string 格式 (民國年 YYYMMDD)
        /// </summary>
        [HttpPost]
        public JsonResult SaveDateField(string td_code, string syr, string smn, string it_code, string field, string value) {
            try {
                // 驗證欄位名稱，只允許 REL_SDT 或 REL_EDT
                if(field != "REL_SDT" && field != "REL_EDT") {
                    return Json(new { success = false, message = "無效的欄位名稱" });
                }

                string uid = WebSiteHelper.getUserId();
                DateTime now = DateTime.Now;
                string nowTime = now.ToString("HHmmss");

                // 先檢查記錄是否存在
                SqlCommand checkCmd = new SqlCommand("SELECT COUNT(*) FROM TDDSDWB WHERE TD_CODE = @td_code AND IT_CODE = @it_code");
                checkCmd.Parameters.Add(new SqlParameter("td_code", td_code));
                checkCmd.Parameters.Add(new SqlParameter("it_code", it_code));
                DataTable checkResult = BASE_DAO.SqlDao.CMETMSQuery4List(checkCmd);
                int count = Convert.ToInt32(checkResult.Rows[0][0]);

                if(count > 0) {
                    // UPDATE 現有記錄
                    StringBuilder sb = new StringBuilder();
                    sb.Append("UPDATE TDDSDWB SET ");
                    sb.Append(field + " = @value, ");
                    sb.Append("UPD_DT = @upd_dt, UPD_TM = @upd_tm, UPD_UID = @upd_uid ");
                    sb.Append("WHERE TD_CODE = @td_code AND IT_CODE = @it_code");

                    SqlCommand cmd = new SqlCommand(sb.ToString());
                    // REL_SDT/REL_EDT 是 string 格式，直接存民國年
                    cmd.Parameters.Add(new SqlParameter("value", string.IsNullOrEmpty(value) ? (object)DBNull.Value : value));
                    cmd.Parameters.Add(new SqlParameter("upd_dt", now));
                    cmd.Parameters.Add(new SqlParameter("upd_tm", nowTime));
                    cmd.Parameters.Add(new SqlParameter("upd_uid", uid));
                    cmd.Parameters.Add(new SqlParameter("td_code", td_code));
                    cmd.Parameters.Add(new SqlParameter("it_code", it_code));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                } else {
                    // INSERT 新記錄
                    StringBuilder sb = new StringBuilder();
                    sb.Append("INSERT INTO TDDSDWB (TD_CODE, IT_CODE, " + field + ", ADD_DT, UPD_DT, UPD_TM, UPD_UID) ");
                    sb.Append("VALUES (@td_code, @it_code, @value, @add_dt, @upd_dt, @upd_tm, @upd_uid)");

                    SqlCommand cmd = new SqlCommand(sb.ToString());
                    cmd.Parameters.Add(new SqlParameter("td_code", td_code));
                    cmd.Parameters.Add(new SqlParameter("it_code", it_code));
                    cmd.Parameters.Add(new SqlParameter("value", string.IsNullOrEmpty(value) ? (object)DBNull.Value : value));
                    cmd.Parameters.Add(new SqlParameter("add_dt", now));
                    cmd.Parameters.Add(new SqlParameter("upd_dt", now));
                    cmd.Parameters.Add(new SqlParameter("upd_tm", nowTime));
                    cmd.Parameters.Add(new SqlParameter("upd_uid", uid));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                }

                return Json(new { success = true });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// AJAX 儲存文字欄位 (TDDMSCD 表)
        /// TD_CODE + SYR + SMN 為識別鍵
        /// </summary>
        [HttpPost]
        public JsonResult SaveTextField(string td_code, string syr, string smn, string field, string value) {
            try {
                // 3.1.4rev3：年度補 3 碼、月份補 2 碼再寫入。
                //   存成 '1' 這種沒補零的月份，會讓舊系統的 SP_TDE0051_TDGMSD（分月進度清單）
                //   在 SMN = convert(int, SMN)-1 的子查詢同時命中 '01' 與 '1'，
                //   「子查詢傳回不只 1 個值」→ 整個辦理情形及施工進度畫面 500。
                syr = NormalizeSyr(syr);
                smn = NormalizeSmn(smn);

                // 驗證欄位名稱，只允許特定欄位
                string[] allowedFields = { "DELAY_RESN", "RESLV_MTHD", "EXECT_DESC" };
                if(!allowedFields.Contains(field)) {
                    return Json(new { success = false, message = "無效的欄位名稱" });
                }

                string uid = WebSiteHelper.getUserId();
                // UPD_DT 使用民國年格式 (YYYMMDD, 7碼)
                int rocYear = DateTime.Now.Year - 1911;
                string now = rocYear.ToString("000") + DateTime.Now.ToString("MMdd");
                string nowTime = DateTime.Now.ToString("HHmmss");

                // [修正4] 取得最新一筆記錄的 ID，避免更新多筆
                // 同一個 TD_CODE + SYR + SMN 可能有多筆記錄，只更新 ID 最大的那筆
                SqlCommand checkCmd = new SqlCommand("SELECT TOP 1 ID FROM TDDMSCD WHERE TD_CODE = @td_code AND SYR = @syr AND SMN = @smn ORDER BY ID DESC");
                checkCmd.Parameters.Add(new SqlParameter("td_code", td_code));
                checkCmd.Parameters.Add(new SqlParameter("syr", syr));
                checkCmd.Parameters.Add(new SqlParameter("smn", smn));
                DataTable checkResult = BASE_DAO.SqlDao.CMETMSQuery4List(checkCmd);

                if(checkResult.Rows.Count > 0) {
                    // UPDATE 指定 ID 的記錄
                    decimal recordId = Convert.ToDecimal(checkResult.Rows[0]["ID"]);
                    StringBuilder sb = new StringBuilder();
                    sb.Append("UPDATE TDDMSCD SET ");
                    sb.Append(field + " = @value, ");
                    sb.Append("UPD_DT = @upd_dt, UPD_TM = @upd_tm, UPD_UID = @upd_uid ");
                    sb.Append("WHERE ID = @id");

                    SqlCommand cmd = new SqlCommand(sb.ToString());
                    cmd.Parameters.Add(new SqlParameter("value", value ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("upd_dt", now));
                    cmd.Parameters.Add(new SqlParameter("upd_tm", nowTime));
                    cmd.Parameters.Add(new SqlParameter("upd_uid", uid));
                    cmd.Parameters.Add(new SqlParameter("id", recordId));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                } else {
                    // INSERT 新記錄
                    StringBuilder sb = new StringBuilder();
                    sb.Append("INSERT INTO TDDMSCD (TD_CODE, SYR, SMN, " + field + ", ADD_DT, UPD_DT, UPD_TM, UPD_UID) ");
                    sb.Append("VALUES (@td_code, @syr, @smn, @value, @add_dt, @upd_dt, @upd_tm, @upd_uid)");

                    SqlCommand cmd = new SqlCommand(sb.ToString());
                    cmd.Parameters.Add(new SqlParameter("td_code", td_code));
                    cmd.Parameters.Add(new SqlParameter("syr", syr));
                    cmd.Parameters.Add(new SqlParameter("smn", smn));
                    cmd.Parameters.Add(new SqlParameter("value", value ?? (object)DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("add_dt", now));
                    cmd.Parameters.Add(new SqlParameter("upd_dt", now));
                    cmd.Parameters.Add(new SqlParameter("upd_tm", nowTime));
                    cmd.Parameters.Add(new SqlParameter("upd_uid", uid));

                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                }

                return Json(new { success = true });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// 引用前月資料
        /// </summary>
        [HttpPost]
        // 3.1.4rev3：年度／月份一律正規化成 3 碼／2 碼再進出資料庫（原因見 SaveTextField）
        private static string NormalizeSyr(string syr) {
            int y;
            return (!string.IsNullOrWhiteSpace(syr) && int.TryParse(syr.Trim(), out y)) ? y.ToString("000") : syr;
        }

        private static string NormalizeSmn(string smn) {
            int m;
            return (!string.IsNullOrWhiteSpace(smn) && int.TryParse(smn.Trim(), out m)) ? m.ToString("00") : smn;
        }

        public JsonResult CopyFromPrevMonth(string td_code, string syr, string smn, string type) {
            try {
                syr = NormalizeSyr(syr);
                smn = NormalizeSmn(smn);

                // 計算前一個月
                int year = int.Parse(syr);
                int month = int.Parse(smn);
                month--;
                if(month < 1) {
                    month = 12;
                    year--;
                }
                string prevSyr = year.ToString("000");
                string prevSmn = month.ToString("00");

                // 查詢前月 TDDMSCD 資料
                SqlCommand cmdText = new SqlCommand("SELECT DELAY_RESN, RESLV_MTHD, EXECT_DESC FROM TDDMSCD WHERE TD_CODE = @td_code AND SYR = @syr AND SMN = @smn");
                cmdText.Parameters.Add(new SqlParameter("td_code", td_code));
                cmdText.Parameters.Add(new SqlParameter("syr", prevSyr));
                cmdText.Parameters.Add(new SqlParameter("smn", prevSmn));
                DataTable dtText = BASE_DAO.SqlDao.CMETMSQuery4List(cmdText);

                if(dtText.Rows.Count == 0) {
                    return Json(new { success = false, message = "找不到前月(" + prevSyr + "/" + prevSmn + ")資料" });
                }

                string delayResn = dtText.Rows[0]["DELAY_RESN"]?.ToString() ?? "";
                string reslvMthd = dtText.Rows[0]["RESLV_MTHD"]?.ToString() ?? "";
                string exectDesc = dtText.Rows[0]["EXECT_DESC"]?.ToString() ?? "";

                // 查詢前月 TDDSDWB 日期欄位
                string sql = GenIncSql(td_code, type);
                // 修改查詢為前月的資料 - 但 TDDSDWB 沒有年月欄位，日期欄位是跟著 TD_CODE
                // 所以日期欄位直接從現有資料複製即可
                SqlCommand cmdDate = new SqlCommand("SELECT b.IT_CODE, b.REL_SDT, b.REL_EDT FROM TDDSDWB b LEFT OUTER JOIN SMWSDTP p ON p.IT_CODE=b.IT_CODE WHERE b.TD_CODE = @td_code AND p.TD_TYPE = @type AND p.AN_IT_CODE IS NOT NULL ORDER BY p.GP_ORDER");
                cmdDate.Parameters.Add(new SqlParameter("td_code", td_code));
                cmdDate.Parameters.Add(new SqlParameter("type", type));
                DataTable dtDate = BASE_DAO.SqlDao.CMETMSQuery4List(cmdDate);

                var dateFields = new Dictionary<string, object>();
                foreach(DataRow dr in dtDate.Rows) {
                    string itCode = dr["IT_CODE"]?.ToString() ?? "";
                    if(!string.IsNullOrEmpty(itCode)) {
                        dateFields[itCode] = new {
                            REL_SDT = dr["REL_SDT"]?.ToString() ?? "",
                            REL_EDT = dr["REL_EDT"]?.ToString() ?? ""
                        };
                    }
                }

                // 將前月文字欄位寫入當月
                string uid = WebSiteHelper.getUserId();
                int rocYear = DateTime.Now.Year - 1911;
                string now = rocYear.ToString("000") + DateTime.Now.ToString("MMdd");
                string nowTime = DateTime.Now.ToString("HHmmss");

                // 檢查當月記錄是否存在
                SqlCommand checkCmd = new SqlCommand("SELECT COUNT(*) FROM TDDMSCD WHERE TD_CODE = @td_code AND SYR = @syr AND SMN = @smn");
                checkCmd.Parameters.Add(new SqlParameter("td_code", td_code));
                checkCmd.Parameters.Add(new SqlParameter("syr", syr));
                checkCmd.Parameters.Add(new SqlParameter("smn", smn));
                DataTable checkResult = BASE_DAO.SqlDao.CMETMSQuery4List(checkCmd);
                int count = Convert.ToInt32(checkResult.Rows[0][0]);

                if(count > 0) {
                    // UPDATE
                    SqlCommand updateCmd = new SqlCommand("UPDATE TDDMSCD SET DELAY_RESN = @delay_resn, RESLV_MTHD = @reslv_mthd, EXECT_DESC = @exect_desc, UPD_DT = @upd_dt, UPD_TM = @upd_tm, UPD_UID = @upd_uid WHERE TD_CODE = @td_code AND SYR = @syr AND SMN = @smn");
                    updateCmd.Parameters.Add(new SqlParameter("delay_resn", delayResn));
                    updateCmd.Parameters.Add(new SqlParameter("reslv_mthd", reslvMthd));
                    updateCmd.Parameters.Add(new SqlParameter("exect_desc", exectDesc));
                    updateCmd.Parameters.Add(new SqlParameter("upd_dt", now));
                    updateCmd.Parameters.Add(new SqlParameter("upd_tm", nowTime));
                    updateCmd.Parameters.Add(new SqlParameter("upd_uid", uid));
                    updateCmd.Parameters.Add(new SqlParameter("td_code", td_code));
                    updateCmd.Parameters.Add(new SqlParameter("syr", syr));
                    updateCmd.Parameters.Add(new SqlParameter("smn", smn));
                    BASE_DAO.SqlDao.CMETMSExecuteSql(updateCmd);
                } else {
                    // INSERT
                    SqlCommand insertCmd = new SqlCommand("INSERT INTO TDDMSCD (TD_CODE, SYR, SMN, DELAY_RESN, RESLV_MTHD, EXECT_DESC, ADD_DT, UPD_DT, UPD_TM, UPD_UID) VALUES (@td_code, @syr, @smn, @delay_resn, @reslv_mthd, @exect_desc, @add_dt, @upd_dt, @upd_tm, @upd_uid)");
                    insertCmd.Parameters.Add(new SqlParameter("td_code", td_code));
                    insertCmd.Parameters.Add(new SqlParameter("syr", syr));
                    insertCmd.Parameters.Add(new SqlParameter("smn", smn));
                    insertCmd.Parameters.Add(new SqlParameter("delay_resn", delayResn));
                    insertCmd.Parameters.Add(new SqlParameter("reslv_mthd", reslvMthd));
                    insertCmd.Parameters.Add(new SqlParameter("exect_desc", exectDesc));
                    insertCmd.Parameters.Add(new SqlParameter("add_dt", now));
                    insertCmd.Parameters.Add(new SqlParameter("upd_dt", now));
                    insertCmd.Parameters.Add(new SqlParameter("upd_tm", nowTime));
                    insertCmd.Parameters.Add(new SqlParameter("upd_uid", uid));
                    BASE_DAO.SqlDao.CMETMSExecuteSql(insertCmd);
                }

                return Json(new {
                    success = true,
                    data = new {
                        DELAY_RESN = delayResn,
                        RESLV_MTHD = reslvMthd,
                        EXECT_DESC = exectDesc,
                        dateFields = dateFields
                    }
                });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private string GetLoginUserCity()
        {
            // 1. 取得目前登入者的 ID (從 FormsAuthentication Cookie 解析出來的)
            string userId = User.Identity.Name;

            if (string.IsNullOrEmpty(userId))
            {
                return null;
            }

            // 2. 查詢該使用者的單位名稱
            // 邏輯：OGCUSER -> OGCDPUR -> OGCDEPT -> DEPT_NAME
            string sql = @"
        SELECT TOP 1 d.SMP_DEPT_NAME 
        FROM OGCDEPT d
        JOIN OGCDPUR r ON d.DEPT_NO = r.DEPT_NO
        WHERE r.USR_ID = @USR_ID";

            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("USR_ID", userId));

            DataTable dt = BASE_DAO.SqlDao.CMBASEQuery4List(cmd); // 注意：這裡是查 CMBASE

            if (dt.Rows.Count > 0)
            {
                string deptName = dt.Rows[0]["SMP_DEPT_NAME"].ToString();

                // 3. 判斷單位名稱是否包含「縣」或「市」，並回傳縣市名稱
                // 這裡需要根據貴單位的 DEPT_NAME 規則來寫
                // 例如：如果是 "宜蘭縣政府"，就回傳 "宜蘭縣"
                // 如果是 "營建署" (中央)，就回傳 null 或 ""

                if (deptName.Contains("縣政府"))
                {
                    return deptName.Replace("政府", ""); // "宜蘭縣政府" -> "宜蘭縣"
                }
                else if (deptName.EndsWith("市政府"))
                {
                    return deptName.Replace("政府", ""); // "高雄市政府" -> "高雄市"
                }
                else if (deptName.EndsWith("縣") || deptName.EndsWith("市"))
                {
                    // 如果單位名稱本身就是 "宜蘭縣" 或 "基隆市"
                    return deptName;
                }
            }

            // 如果是中央單位 (如營建署) 或查無資料，回傳 null
            return null;
        }

        /// <summary>
        /// 取得年度計畫(依年度)
        /// </summary>
        /// <param name="syr"> 年度 </param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult GetPidsBySyr(string syr)
        {
            if (string.IsNullOrEmpty(syr))
            {
                return Json(new List<object>());
            }

            var pids = db_pms.PJV_PJDPCAT
                .Where(s => s.SYR == syr)
                //.Select(c => new { c.PID, c.SYR_PJ_NAME })
                .Select(c => new { c.PID, c.AN_PJ_NAME, c.PJ_NAME })
                .AsEnumerable()
                .Select(c => new
                {
                    PID = c.PID.ToString(),
                    SYR_PJ_NAME = string.Format("({0}){1}", c.AN_PJ_NAME, c.PJ_NAME)
                })
                .ToList();

            return Json(pids);
        }

    }
}