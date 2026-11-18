using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Syncfusion.XlsIO;
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
    /// TDE0070 污水下水道普及率統計 Controller
    /// 頁籤1: 統計表 (CCV_CCDMCNT)
    /// 頁籤2: 資料設定 (CCEMCNT)
    /// </summary>
    public class TDE0070Controller : Controller {

        private readonly int pageSize = 50;

        // GET: TDE0070
        public ActionResult Index(TDE0070Model model) {
            if (model == null) {
                model = new TDE0070Model();
            }

            // 取得年度下拉選單
            var yearList = GetYearList();
            ViewBag.Years = yearList;

            // 預設年度為清單第一筆 (最新年度)
            if (string.IsNullOrEmpty(model.SYR) && yearList.Count > 0) {
                model.SYR = yearList[0].Value;
            }

            // 取得月份下拉選單 (依年度篩選)
            var monthList = GetMonthList(model.SYR);
            ViewBag.Months = monthList;

            // 預設月份為清單第一筆 (最新月份)
            if (string.IsNullOrEmpty(model.SMN) && monthList.Count > 0) {
                model.SMN = monthList[0].Value;
            }

            // 預設頁籤為1
            if (model.tab == 0) {
                model.tab = 1;
            }

            return View("TDE0070Index", model);
        }

        /// <summary>
        /// Tab 頁面 (供 iframe 載入)
        /// </summary>
        public ActionResult Tab(int tab = 1, string SYR = null, string SMN = null, int page1 = 1, int page2 = 1) {
            TDE0070Model model = new TDE0070Model {
                tab = tab,
                SYR = SYR,
                SMN = SMN,
                page1 = page1,
                page2 = page2
            };

            // 根據當前頁籤載入對應資料
            switch (tab) {
                case 1:
                    LoadTab1Data(model);
                    break;
                case 2:
                    LoadTab2Data(model);
                    break;
            }

            return View("TDE0070Tab", model);
        }

        /// <summary>
        /// 頁籤1: 統計表
        /// SQL: select * from CMETMS.dbo.CCV_CCDMCNT where SYR is not null
        /// </summary>
        private void LoadTab1Data(TDE0070Model model) {
            if (model.page1 == 0) model.page1 = 1;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM CCV_CCDMCNT WHERE SYR IS NOT NULL ");

            SqlCommand cmd = new SqlCommand();

            if (!string.IsNullOrEmpty(model.SYR)) {
                sb.Append("AND SYR = @SYR ");
                cmd.Parameters.Add(new SqlParameter("SYR", model.SYR));
            }

            if (!string.IsNullOrEmpty(model.SMN)) {
                sb.Append("AND SMN = @SMN ");
                cmd.Parameters.Add(new SqlParameter("SMN", model.SMN));
            }

            sb.Append("ORDER BY SYR DESC, SMN DESC, CITY");
            cmd.CommandText = sb.ToString();

            DataTable dt = SqlDao.CMETMSQuery4List(cmd);
            List<DataRow> list = dt.AsEnumerable().ToList();
            model.List_Tab1 = new PagedList<DataRow>(list, model.page1, pageSize);
        }

        /// <summary>
        /// 頁籤2: 資料設定
        /// SQL: select * from CMETMS.dbo.CCEMCNT where 1=1
        /// </summary>
        private void LoadTab2Data(TDE0070Model model) {
            if (model.page2 == 0) model.page2 = 1;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM CCEMCNT WHERE 1=1 ");

            SqlCommand cmd = new SqlCommand();

            if (!string.IsNullOrEmpty(model.SYR)) {
                sb.Append("AND SYR = @SYR ");
                cmd.Parameters.Add(new SqlParameter("SYR", model.SYR));
            }

            if (!string.IsNullOrEmpty(model.SMN)) {
                sb.Append("AND SMN = @SMN ");
                cmd.Parameters.Add(new SqlParameter("SMN", model.SMN));
            }

            sb.Append("ORDER BY SYR DESC, SMN DESC, CITY");
            cmd.CommandText = sb.ToString();

            DataTable dt = SqlDao.CMETMSQuery4List(cmd);
            List<DataRow> list = dt.AsEnumerable().ToList();
            model.List_Tab2 = new PagedList<DataRow>(list, model.page2, pageSize);
        }

        /// <summary>
        /// 取得年度下拉選單 (從資料庫動態取得)
        /// </summary>
        private List<SelectListItem> GetYearList() {
            List<SelectListItem> list = new List<SelectListItem>();
            try {
                SqlCommand cmd = new SqlCommand(@"SELECT DISTINCT SYR FROM CCV_CCDMCNT
                    WHERE SYR IS NOT NULL ORDER BY SYR DESC");
                DataTable dt = SqlDao.CMETMSQuery4List(cmd);
                foreach (DataRow row in dt.Rows) {
                    string syr = row["SYR"].ToString();
                    list.Add(new SelectListItem { Text = syr, Value = syr });
                }
            } catch {
                // 若查詢失敗，使用預設範圍
                int currentYear = DateTime.Now.Year - 1911;
                for (int i = currentYear; i >= 96; i--) {
                    list.Add(new SelectListItem { Text = i.ToString("000"), Value = i.ToString("000") });
                }
            }
            return list;
        }

        /// <summary>
        /// 取得月份下拉選單 (從資料庫動態取得)
        /// </summary>
        private List<SelectListItem> GetMonthList(string syr = null) {
            List<SelectListItem> list = new List<SelectListItem>();
            try {
                StringBuilder sb = new StringBuilder();
                sb.Append("SELECT DISTINCT SMN FROM CCV_CCDMCNT WHERE SMN IS NOT NULL ");
                SqlCommand cmd = new SqlCommand();
                if (!string.IsNullOrEmpty(syr)) {
                    sb.Append("AND SYR = @SYR ");
                    cmd.Parameters.Add(new SqlParameter("SYR", syr));
                }
                sb.Append("ORDER BY SMN DESC");
                cmd.CommandText = sb.ToString();
                DataTable dt = SqlDao.CMETMSQuery4List(cmd);
                foreach (DataRow row in dt.Rows) {
                    string smn = row["SMN"].ToString();
                    list.Add(new SelectListItem { Text = smn, Value = smn });
                }
            } catch {
                // 若查詢失敗，使用預設範圍
                for (int i = 12; i >= 1; i--) {
                    list.Add(new SelectListItem { Text = i.ToString("00"), Value = i.ToString("00") });
                }
            }
            return list;
        }

        /// <summary>
        /// AJAX 取得指定年度的月份清單
        /// </summary>
        [HttpPost]
        public JsonResult GetMonths(string SYR) {
            var list = GetMonthList(SYR);
            return Json(list.Select(x => new { Text = x.Text, Value = x.Value }));
        }

        #region 頁籤1 - 統計表編輯 (CCDMCNT)

        /// <summary>
        /// 編輯頁面 - 統計表 (頁籤1)
        /// </summary>
        public ActionResult Edit(string SYR, string SMN, string CITY) {
            if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN) || string.IsNullOrEmpty(CITY)) {
                return Content("參數不完整");
            }

            TDE0070Model model = new TDE0070Model {
                SYR = SYR,
                SMN = SMN,
                CITY = CITY
            };

            // 查詢 CCDMCNT 資料
            SqlCommand cmdList = new SqlCommand(@"SELECT * FROM CCDMCNT
                WHERE SYR = @SYR AND SMN = @SMN AND CITY = @CITY");
            cmdList.Parameters.Add(new SqlParameter("SYR", SYR));
            cmdList.Parameters.Add(new SqlParameter("SMN", SMN));
            cmdList.Parameters.Add(new SqlParameter("CITY", CITY));
            ViewBag.List1 = SqlDao.CMETMSQuery4List(cmdList);

            return View("TDE0070Edit", model);
        }

        /// <summary>
        /// AJAX 更新 CCDMCNT 資料
        /// </summary>
        [HttpPost]
        public JsonResult EditUpdate(string SYR, string SMN, string CITY,
            decimal? PC_CNT, decimal? AC_CNT, decimal? BC_CNT, decimal? CC_CNT,
            string CHANGE_NOTE, string UNUSUAL_NOTE) {
            try {
                if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN) || string.IsNullOrEmpty(CITY)) {
                    return Json(new { success = false, message = "參數不完整" });
                }

                // 檢查是否已鎖定
                SqlCommand cmdCheck = new SqlCommand("SELECT IS_LOCK FROM CCDMCNT WHERE SYR = @SYR AND SMN = @SMN AND CITY = @CITY");
                cmdCheck.Parameters.Add(new SqlParameter("SYR", SYR));
                cmdCheck.Parameters.Add(new SqlParameter("SMN", SMN));
                cmdCheck.Parameters.Add(new SqlParameter("CITY", CITY));
                DataTable dtCheck = SqlDao.CMETMSQuery4List(cmdCheck);
                if (dtCheck.Rows.Count > 0 && dtCheck.Rows[0]["IS_LOCK"]?.ToString() == "Y") {
                    return Json(new { success = false, message = "此資料已鎖定，無法編輯" });
                }

                // 檢查是否存在，不存在則新增
                if (dtCheck.Rows.Count == 0) {
                    SqlCommand cmdInsert = new SqlCommand(@"INSERT INTO CCDMCNT
                        (SYR, SMN, CITY, PC_CNT, AC_CNT, BC_CNT, CC_CNT, CHANGE_NOTE, UNUSUAL_NOTE, ADD_DT)
                        VALUES (@SYR, @SMN, @CITY, @PC_CNT, @AC_CNT, @BC_CNT, @CC_CNT, @CHANGE_NOTE, @UNUSUAL_NOTE, @ADD_DT)");
                    cmdInsert.Parameters.Add(new SqlParameter("SYR", SYR));
                    cmdInsert.Parameters.Add(new SqlParameter("SMN", SMN));
                    cmdInsert.Parameters.Add(new SqlParameter("CITY", CITY));
                    cmdInsert.Parameters.Add(new SqlParameter("PC_CNT", PC_CNT.HasValue ? (object)PC_CNT.Value : DBNull.Value));
                    cmdInsert.Parameters.Add(new SqlParameter("AC_CNT", AC_CNT.HasValue ? (object)AC_CNT.Value : DBNull.Value));
                    cmdInsert.Parameters.Add(new SqlParameter("BC_CNT", BC_CNT.HasValue ? (object)BC_CNT.Value : DBNull.Value));
                    cmdInsert.Parameters.Add(new SqlParameter("CC_CNT", CC_CNT.HasValue ? (object)CC_CNT.Value : DBNull.Value));
                    cmdInsert.Parameters.Add(new SqlParameter("CHANGE_NOTE", string.IsNullOrEmpty(CHANGE_NOTE) ? (object)DBNull.Value : CHANGE_NOTE));
                    cmdInsert.Parameters.Add(new SqlParameter("UNUSUAL_NOTE", string.IsNullOrEmpty(UNUSUAL_NOTE) ? (object)DBNull.Value : UNUSUAL_NOTE));
                    cmdInsert.Parameters.Add(new SqlParameter("ADD_DT", DateTime.Now));
                    SqlDao.CMETMSExecuteSql(cmdInsert);
                } else {
                    SqlCommand cmd = new SqlCommand(@"UPDATE CCDMCNT SET
                        PC_CNT = @PC_CNT,
                        AC_CNT = @AC_CNT,
                        BC_CNT = @BC_CNT,
                        CC_CNT = @CC_CNT,
                        CHANGE_NOTE = @CHANGE_NOTE,
                        UNUSUAL_NOTE = @UNUSUAL_NOTE
                        WHERE SYR = @SYR AND SMN = @SMN AND CITY = @CITY");
                    cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                    cmd.Parameters.Add(new SqlParameter("SMN", SMN));
                    cmd.Parameters.Add(new SqlParameter("CITY", CITY));
                    cmd.Parameters.Add(new SqlParameter("PC_CNT", PC_CNT.HasValue ? (object)PC_CNT.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("AC_CNT", AC_CNT.HasValue ? (object)AC_CNT.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("BC_CNT", BC_CNT.HasValue ? (object)BC_CNT.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("CC_CNT", CC_CNT.HasValue ? (object)CC_CNT.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("CHANGE_NOTE", string.IsNullOrEmpty(CHANGE_NOTE) ? (object)DBNull.Value : CHANGE_NOTE));
                    cmd.Parameters.Add(new SqlParameter("UNUSUAL_NOTE", string.IsNullOrEmpty(UNUSUAL_NOTE) ? (object)DBNull.Value : UNUSUAL_NOTE));
                    SqlDao.CMETMSExecuteSql(cmd);
                }

                return Json(new { success = true, message = "更新成功" });
            } catch (Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// AJAX 鎖定資料
        /// </summary>
        [HttpPost]
        public JsonResult Lock(string SYR, string SMN, string CITY) {
            try {
                if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN) || string.IsNullOrEmpty(CITY)) {
                    return Json(new { success = false, message = "參數不完整" });
                }

                DateTime now = DateTime.Now;
                string userId = WebSiteHelper.getUserId();

                // 檢查記錄是否存在
                SqlCommand cmdCheck = new SqlCommand("SELECT COUNT(*) as CNT FROM CCDMCNT WHERE SYR = @SYR AND SMN = @SMN AND CITY = @CITY");
                cmdCheck.Parameters.Add(new SqlParameter("SYR", SYR));
                cmdCheck.Parameters.Add(new SqlParameter("SMN", SMN));
                cmdCheck.Parameters.Add(new SqlParameter("CITY", CITY));
                DataTable dtCheck = SqlDao.CMETMSQuery4List(cmdCheck);

                if (dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0]["CNT"]) > 0) {
                    // 記錄存在，執行更新
                    SqlCommand cmd = new SqlCommand(@"UPDATE CCDMCNT SET
                        IS_LOCK = 'Y',
                        LOCK_UID = @LOCK_UID,
                        LOCK_DT = @LOCK_DT
                        WHERE SYR = @SYR AND SMN = @SMN AND CITY = @CITY");
                    cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                    cmd.Parameters.Add(new SqlParameter("SMN", SMN));
                    cmd.Parameters.Add(new SqlParameter("CITY", CITY));
                    cmd.Parameters.Add(new SqlParameter("LOCK_UID", userId ?? ""));
                    cmd.Parameters.Add(new SqlParameter("LOCK_DT", now));
                    SqlDao.CMETMSExecuteSql(cmd);
                } else {
                    // 記錄不存在，執行新增
                    SqlCommand cmdInsert = new SqlCommand(@"INSERT INTO CCDMCNT
                        (SYR, SMN, CITY, IS_LOCK, LOCK_UID, LOCK_DT, ADD_DT)
                        VALUES (@SYR, @SMN, @CITY, 'Y', @LOCK_UID, @LOCK_DT, @ADD_DT)");
                    cmdInsert.Parameters.Add(new SqlParameter("SYR", SYR));
                    cmdInsert.Parameters.Add(new SqlParameter("SMN", SMN));
                    cmdInsert.Parameters.Add(new SqlParameter("CITY", CITY));
                    cmdInsert.Parameters.Add(new SqlParameter("LOCK_UID", userId ?? ""));
                    cmdInsert.Parameters.Add(new SqlParameter("LOCK_DT", now));
                    cmdInsert.Parameters.Add(new SqlParameter("ADD_DT", now));
                    SqlDao.CMETMSExecuteSql(cmdInsert);
                }
                return Json(new { success = true, message = "鎖定成功" });
            } catch (Exception ex) {
                return Json(new { success = false, message = "鎖定失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// AJAX 解除鎖定
        /// </summary>
        [HttpPost]
        public JsonResult Unlock(string SYR, string SMN, string CITY) {
            try {
                if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN) || string.IsNullOrEmpty(CITY)) {
                    return Json(new { success = false, message = "參數不完整" });
                }

                DateTime now = DateTime.Now;
                string userId = WebSiteHelper.getUserId();

                SqlCommand cmd = new SqlCommand(@"UPDATE CCDMCNT SET
                    IS_LOCK = 'N',
                    ULOCK_UID = @ULOCK_UID,
                    ULOCK_DT = @ULOCK_DT
                    WHERE SYR = @SYR AND SMN = @SMN AND CITY = @CITY");
                cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                cmd.Parameters.Add(new SqlParameter("SMN", SMN));
                cmd.Parameters.Add(new SqlParameter("CITY", CITY));
                cmd.Parameters.Add(new SqlParameter("ULOCK_UID", userId ?? ""));
                cmd.Parameters.Add(new SqlParameter("ULOCK_DT", now));

                SqlDao.CMETMSExecuteSql(cmd);
                return Json(new { success = true, message = "解除鎖定成功" });
            } catch (Exception ex) {
                return Json(new { success = false, message = "解除鎖定失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// AJAX 批次鎖定 (整個年月)
        /// </summary>
        [HttpPost]
        public JsonResult LockAll(string SYR, string SMN) {
            try {
                if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN)) {
                    return Json(new { success = false, message = "請選擇年度和月份" });
                }

                DateTime now = DateTime.Now;
                string userId = WebSiteHelper.getUserId();

                // 先為 CCEMCNT 中存在但 CCDMCNT 不存在的記錄新增
                SqlCommand cmdInsert = new SqlCommand(@"INSERT INTO CCDMCNT (SYR, SMN, CITY, IS_LOCK, LOCK_UID, LOCK_DT, ADD_DT)
                    SELECT @SYR, @SMN, E.CITY, 'Y', @LOCK_UID, @LOCK_DT, @ADD_DT
                    FROM CCEMCNT E
                    WHERE E.SYR = @SYR AND E.SMN = @SMN
                    AND NOT EXISTS (SELECT 1 FROM CCDMCNT C WHERE C.SYR = E.SYR AND C.SMN = E.SMN AND C.CITY = E.CITY)");
                cmdInsert.Parameters.Add(new SqlParameter("SYR", SYR));
                cmdInsert.Parameters.Add(new SqlParameter("SMN", SMN));
                cmdInsert.Parameters.Add(new SqlParameter("LOCK_UID", userId ?? ""));
                cmdInsert.Parameters.Add(new SqlParameter("LOCK_DT", now));
                cmdInsert.Parameters.Add(new SqlParameter("ADD_DT", now));
                SqlDao.CMETMSExecuteSql(cmdInsert);

                // 更新已存在的記錄
                SqlCommand cmd = new SqlCommand(@"UPDATE CCDMCNT SET
                    IS_LOCK = 'Y',
                    LOCK_UID = @LOCK_UID,
                    LOCK_DT = @LOCK_DT
                    WHERE SYR = @SYR AND SMN = @SMN");
                cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                cmd.Parameters.Add(new SqlParameter("SMN", SMN));
                cmd.Parameters.Add(new SqlParameter("LOCK_UID", userId ?? ""));
                cmd.Parameters.Add(new SqlParameter("LOCK_DT", now));

                SqlDao.CMETMSExecuteSql(cmd);
                return Json(new { success = true, message = $"已鎖定 {SYR}年{SMN}月 所有縣市資料" });
            } catch (Exception ex) {
                return Json(new { success = false, message = "批次鎖定失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// AJAX 批次解除鎖定 (整個年月)
        /// </summary>
        [HttpPost]
        public JsonResult UnlockAll(string SYR, string SMN) {
            try {
                if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN)) {
                    return Json(new { success = false, message = "請選擇年度和月份" });
                }

                DateTime now = DateTime.Now;
                string userId = WebSiteHelper.getUserId();

                SqlCommand cmd = new SqlCommand(@"UPDATE CCDMCNT SET
                    IS_LOCK = 'N',
                    ULOCK_UID = @ULOCK_UID,
                    ULOCK_DT = @ULOCK_DT
                    WHERE SYR = @SYR AND SMN = @SMN");
                cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                cmd.Parameters.Add(new SqlParameter("SMN", SMN));
                cmd.Parameters.Add(new SqlParameter("ULOCK_UID", userId ?? ""));
                cmd.Parameters.Add(new SqlParameter("ULOCK_DT", now));

                SqlDao.CMETMSExecuteSql(cmd);
                return Json(new { success = true, message = $"已解除鎖定 {SYR}年{SMN}月 所有縣市資料" });
            } catch (Exception ex) {
                return Json(new { success = false, message = "批次解除鎖定失敗：" + ex.Message });
            }
        }

        #endregion

        #region 頁籤2 - 資料設定編輯 (CCEMCNT)

        /// <summary>
        /// 3.3.4：匯出 Excel 報表（用戶接管普及率及污水處理率統計一覽表）
        ///
        /// 完全比照舊系統 CMETMS\TDE0070.aspx.cs 的 GoPrint()：
        ///   - 範本：年度 &gt; 103 用 Template\TDE0070_104.xlsx，否則 Template\TDE0070.xlsx
        ///   - 資料：CCV_CCDMCNT（與畫面統計表同一個檢視，排序也相同）
        ///   - 自第 5 列起，A=CITY、T=CHANGE_NOTE、U=UNUSUAL_NOTE，
        ///     104 年制 B~K 對應 C1,C3~C11；舊制 B~K 對應 C1~C10
        ///   - Sheet2 第 6~10 列為附註，複製到資料下方後移除 Sheet2
        /// </summary>
        public ActionResult ExportExcel(string SYR, string SMN) {
            if(string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN)) {
                return Content("請先選擇年度與月份再匯出。");
            }

            // 與統計表相同的資料與排序
            SqlCommand cmd = new SqlCommand(
                "SELECT * FROM CCV_CCDMCNT WHERE SYR = @SYR AND SMN = @SMN ORDER BY CITY");
            cmd.Parameters.Add(new SqlParameter("SYR", SYR));
            cmd.Parameters.Add(new SqlParameter("SMN", SMN));
            DataTable dt = SqlDao.CMETMSQuery4List(cmd);

            if(dt == null || dt.Rows.Count == 0) {
                return Content($"{SYR} 年 {SMN} 月查無資料，無法匯出。");
            }

            bool isNew = false;
            int syrNum;
            if(int.TryParse(SYR, out syrNum)) { isNew = syrNum > 103; }

            string templateName = isNew ? "TDE0070_104.xlsx" : "TDE0070.xlsx";
            string templatePath = Server.MapPath("~/Template/" + templateName);
            if(!System.IO.File.Exists(templatePath)) {
                return Content("找不到報表範本：" + templateName);
            }

            using(ExcelEngine excelEngine = new ExcelEngine()) {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2007;

                IWorkbook workbook;
                using(FileStream readFile = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    workbook = application.Workbooks.Open(readFile);
                }

                IWorksheet sheet1 = workbook.Worksheets["Sheet1"];
                IWorksheet sheet2 = workbook.Worksheets["Sheet2"];

                const int line = 5;
                int lastCol = isNew ? 21 : 15;
                int noteCol = isNew ? 24 : 11;

                // 依資料筆數複製列格式
                for(int i = 1; i < dt.Rows.Count; i++) {
                    sheet1[line, 1, line, lastCol].CopyTo(sheet1[(line + i), 1]);
                    sheet1[(line + i), 1, (line + i), lastCol].RowHeight = sheet1[line, 1, line, lastCol].RowHeight;
                }

                sheet1["A2"].Text = "資料截止日期：" + GetYearMonthDay(SYR, SMN);

                // 104 年制多了 C2 以外的一欄，欄位對應與舊系統一致
                string[] cols = isNew
                    ? new[] { "C1", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11" }
                    : new[] { "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10" };

                for(int i = 0; i < dt.Rows.Count; i++) {
                    int r = line + i;
                    sheet1["A" + r].Text = SafeStr(dt.Rows[i], "CITY");
                    for(int c = 0; c < cols.Length; c++) {
                        // B..K
                        string cell = ((char)('B' + c)).ToString() + r;
                        sheet1[cell].Value = SafeStr(dt.Rows[i], cols[c]);
                    }
                    sheet1["T" + r].Value = SafeStr(dt.Rows[i], "CHANGE_NOTE");
                    sheet1["U" + r].Value = SafeStr(dt.Rows[i], "UNUSUAL_NOTE");
                }

                // Sheet2 第 6~10 列為附註區，接在資料下方
                for(int i = 6; i <= 10; i++) {
                    sheet2[i, 1, i, noteCol].CopyTo(sheet1[(i + dt.Rows.Count), 1]);
                    sheet1[(i + dt.Rows.Count), 1, (i + dt.Rows.Count), noteCol].RowHeight = sheet2[i, 1, i, noteCol].RowHeight;
                }
                sheet1["A" + (7 + dt.Rows.Count)].Text =
                    "1.上述各縣市總人口數及各縣市戶量係由內政部戶政司網站(http://www.ris.gov.tw)年度統計至" + SYR + SMN + "底資料。";

                sheet2.Remove();

                // 3.3.4：範本檔內還留著一張上次輸出結果的殘留工作表
                // （「用戶接管普及率及污水處理率統計一覽表 (2)」，內容是 103 年舊資料，且排在第一個），
                // 舊系統沒有清掉，使用者打開 Excel 會先看到錯的年份資料。
                // 這裡把資料表以外的工作表全部移除，並給它一個看得懂的名稱。
                for(int i = workbook.Worksheets.Count - 1; i >= 0; i--) {
                    if(!ReferenceEquals(workbook.Worksheets[i], sheet1)) {
                        workbook.Worksheets[i].Remove();
                    }
                }
                sheet1.Name = "用戶接管普及率及污水處理率統計一覽表";

                using(MemoryStream ms = new MemoryStream()) {
                    workbook.SaveAs(ms);
                    workbook.Close();
                    ms.Position = 0;
                    string fileName = $"用戶接管普及率及污水處理率統計一覽表_{SYR}{SMN}.xlsx";
                    return File(ms.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);
                }
            }
        }

        private static string SafeStr(DataRow row, string col) {
            if(!row.Table.Columns.Contains(col)) { return ""; }
            return row[col] == DBNull.Value ? "" : row[col].ToString();
        }

        /// <summary>民國年月 + 該月最後一天（比照舊系統 getYEARMONTHDAY）</summary>
        private static string GetYearMonthDay(string syr, string smn) {
            int y, m;
            if(!int.TryParse(syr, out y) || !int.TryParse(smn, out m) || m < 1 || m > 12) {
                return syr + smn;
            }
            int day = DateTime.DaysInMonth(y + 1911, m);
            return syr + smn + day.ToString("00");
        }

        /// <summary>
        /// 編輯頁面2 - 資料設定 (頁籤2)
        /// </summary>
        public ActionResult Edit2(string SYR, string SMN, string CITY) {
            if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN) || string.IsNullOrEmpty(CITY)) {
                return Content("參數不完整");
            }

            TDE0070Model model = new TDE0070Model {
                SYR = SYR,
                SMN = SMN,
                CITY = CITY
            };

            // 查詢 CCEMCNT 資料
            SqlCommand cmdList = new SqlCommand(@"SELECT * FROM CCEMCNT
                WHERE SYR = @SYR AND SMN = @SMN AND CITY = @CITY");
            cmdList.Parameters.Add(new SqlParameter("SYR", SYR));
            cmdList.Parameters.Add(new SqlParameter("SMN", SMN));
            cmdList.Parameters.Add(new SqlParameter("CITY", CITY));
            ViewBag.List1 = SqlDao.CMETMSQuery4List(cmdList);

            return View("TDE0070Edit2", model);
        }

        /// <summary>
        /// AJAX 更新 CCEMCNT 資料
        /// </summary>
        [HttpPost]
        public JsonResult Edit2Update(string SYR, string SMN, string CITY,
            decimal? POP_CNT, decimal? DOR_CNT, decimal? DOR_COUNT) {
            try {
                if (string.IsNullOrEmpty(SYR) || string.IsNullOrEmpty(SMN) || string.IsNullOrEmpty(CITY)) {
                    return Json(new { success = false, message = "參數不完整" });
                }

                // 檢查是否存在
                SqlCommand cmdCheck = new SqlCommand("SELECT COUNT(*) as CNT FROM CCEMCNT WHERE SYR = @SYR AND SMN = @SMN AND CITY = @CITY");
                cmdCheck.Parameters.Add(new SqlParameter("SYR", SYR));
                cmdCheck.Parameters.Add(new SqlParameter("SMN", SMN));
                cmdCheck.Parameters.Add(new SqlParameter("CITY", CITY));
                DataTable dtCheck = SqlDao.CMETMSQuery4List(cmdCheck);

                if (dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0]["CNT"]) > 0) {
                    // 更新
                    SqlCommand cmd = new SqlCommand(@"UPDATE CCEMCNT SET
                        POP_CNT = @POP_CNT,
                        DOR_CNT = @DOR_CNT,
                        DOR_COUNT = @DOR_COUNT
                        WHERE SYR = @SYR AND SMN = @SMN AND CITY = @CITY");
                    cmd.Parameters.Add(new SqlParameter("SYR", SYR));
                    cmd.Parameters.Add(new SqlParameter("SMN", SMN));
                    cmd.Parameters.Add(new SqlParameter("CITY", CITY));
                    cmd.Parameters.Add(new SqlParameter("POP_CNT", POP_CNT.HasValue ? (object)POP_CNT.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("DOR_CNT", DOR_CNT.HasValue ? (object)DOR_CNT.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("DOR_COUNT", DOR_COUNT.HasValue ? (object)DOR_COUNT.Value : DBNull.Value));
                    SqlDao.CMETMSExecuteSql(cmd);
                } else {
                    // 新增
                    SqlCommand cmdInsert = new SqlCommand(@"INSERT INTO CCEMCNT
                        (SYR, SMN, CITY, POP_CNT, DOR_CNT, DOR_COUNT)
                        VALUES (@SYR, @SMN, @CITY, @POP_CNT, @DOR_CNT, @DOR_COUNT)");
                    cmdInsert.Parameters.Add(new SqlParameter("SYR", SYR));
                    cmdInsert.Parameters.Add(new SqlParameter("SMN", SMN));
                    cmdInsert.Parameters.Add(new SqlParameter("CITY", CITY));
                    cmdInsert.Parameters.Add(new SqlParameter("POP_CNT", POP_CNT.HasValue ? (object)POP_CNT.Value : DBNull.Value));
                    cmdInsert.Parameters.Add(new SqlParameter("DOR_CNT", DOR_CNT.HasValue ? (object)DOR_CNT.Value : DBNull.Value));
                    cmdInsert.Parameters.Add(new SqlParameter("DOR_COUNT", DOR_COUNT.HasValue ? (object)DOR_COUNT.Value : DBNull.Value));
                    SqlDao.CMETMSExecuteSql(cmdInsert);
                }

                return Json(new { success = true, message = "更新成功" });
            } catch (Exception ex) {
                return Json(new { success = false, message = "更新失敗：" + ex.Message });
            }
        }

        #endregion
    }
}
