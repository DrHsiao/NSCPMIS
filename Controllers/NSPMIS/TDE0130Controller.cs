using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS {
    /// <summary>
    /// TDE0130 用戶接管及化糞池拆除統計 Controller
    /// 頁籤1: 用戶接管 (EVL_TYPE='01')
    /// 頁籤2: 化糞池拆除 (EVL_TYPE='02')
    /// </summary>
    public class TDE0130Controller : Controller {

        private readonly int pageSize = 50;

        // GET: TDE0130
        public ActionResult Index(TDE0130Model model) {
            if (model == null) {
                model = new TDE0130Model();
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

            // 取得縣市下拉選單
            ViewBag.Cities = GetCityList();

            // 預設頁籤為1
            if (model.tab == 0) {
                model.tab = 1;
            }

            return View("TDE0130Index", model);
        }

        /// <summary>
        /// Tab 頁面 (供 iframe 載入)
        /// </summary>
        public ActionResult Tab(int tab = 1, string SYR = null, string SMN = null, string TD_CITY = null, int page1 = 1, int page2 = 1) {
            TDE0130Model model = new TDE0130Model {
                tab = tab,
                SYR = SYR,
                SMN = SMN,
                TD_CITY = TD_CITY,
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

            return View("TDE0130Tab", model);
        }

        /// <summary>
        /// 頁籤1: 用戶接管 (EVL_TYPE='01')
        /// </summary>
        private void LoadTab1Data(TDE0130Model model) {
            if (model.page1 == 0) model.page1 = 1;

            StringBuilder sb = new StringBuilder();
            sb.Append(@"SELECT P.CITY AS TD_CITY,
                CASE @SYR WHEN '' THEN S.SYR ELSE @SYR END AS SYR,
                CASE @SMN WHEN '' THEN S.SMN ELSE @SMN END AS SMN,
                ISNULL(S.YEAR_PRE_CON_COUNT_01, 0) AS YEAR_PRE_CON_COUNT_01,
                ISNULL(S.REL_CON_COUNT_01, 0) AS REL_CON_COUNT_01,
                ISNULL(S.CON_DIFF, 0) AS CON_DIFF,
                ISNULL(S.SELF_REMOVE_COUNT_01, 0) AS SELF_REMOVE_COUNT_01,
                ISNULL(S.FOCUS_REMOVE_COUNT_01, 0) AS FOCUS_REMOVE_COUNT_01,
                ISNULL(S.REL_CON_COUNT_01_P, 0) AS REL_CON_COUNT_01_P,
                ISNULL(S.CON_DIFF_P, 0) AS CON_DIFF_P,
                ISNULL(S.SELF_REMOVE_COUNT_01_P, 0) AS SELF_REMOVE_COUNT_01_P,
                ISNULL(S.FOCUS_REMOVE_COUNT_01_P, 0) AS FOCUS_REMOVE_COUNT_01_P
            FROM CMBASE.dbo.SMCPOSTN P
            LEFT JOIN (
                SELECT CASE T.TD_CITY
                    WHEN '台北縣' THEN '新北市'
                    WHEN '台中縣' THEN '台中市'
                    WHEN '台南縣' THEN '台南市'
                    WHEN '高雄縣' THEN '高雄市'
                    WHEN '桃園縣' THEN '桃園市'
                    ELSE TD_CITY END AS TD_CITY,
                    S.SYR, S.SMN,
                    SUM(ISNULL(S.YEAR_PRE_CON_COUNT_01, 0)) AS YEAR_PRE_CON_COUNT_01,
                    SUM(ISNULL(S.REL_CON_COUNT_01, 0)) AS REL_CON_COUNT_01,
                    SUM(ISNULL(S.CON_DIFF, 0)) AS CON_DIFF,
                    SUM(ISNULL(S.SELF_REMOVE_COUNT_01, 0)) AS SELF_REMOVE_COUNT_01,
                    SUM(ISNULL(S.FOCUS_REMOVE_COUNT_01, 0)) AS FOCUS_REMOVE_COUNT_01,
                    CASE SUM(ISNULL(S.YEAR_PRE_CON_COUNT_01, 0)) WHEN 0 THEN 0
                        ELSE SUM(ISNULL(S.REL_CON_COUNT_01, 0)) * 100.00 / SUM(ISNULL(S.YEAR_PRE_CON_COUNT_01, 0)) END AS REL_CON_COUNT_01_P,
                    CASE SUM(ISNULL(S.YEAR_PRE_CON_COUNT_01, 0)) WHEN 0 THEN 0
                        ELSE SUM(ISNULL(S.CON_DIFF, 0)) * 100.00 / SUM(ISNULL(S.YEAR_PRE_CON_COUNT_01, 0)) END AS CON_DIFF_P,
                    CASE SUM(ISNULL(S.YEAR_PRE_CON_COUNT_01, 0)) WHEN 0 THEN 0
                        ELSE SUM(ISNULL(S.SELF_REMOVE_COUNT_01, 0)) * 100.00 / SUM(ISNULL(S.YEAR_PRE_CON_COUNT_01, 0)) END AS SELF_REMOVE_COUNT_01_P,
                    CASE SUM(ISNULL(S.YEAR_PRE_CON_COUNT_01, 0)) WHEN 0 THEN 0
                        ELSE SUM(ISNULL(S.FOCUS_REMOVE_COUNT_01, 0)) * 100.00 / SUM(ISNULL(S.YEAR_PRE_CON_COUNT_01, 0)) END AS FOCUS_REMOVE_COUNT_01_P
                FROM CMETMS.dbo.TDDSEVL S
                LEFT JOIN CMETMS.dbo.TDMFORM T ON T.TD_CODE = S.TD_CODE
                WHERE EVL_TYPE = '01' AND ISNULL(T.TD_CITY, '') NOT IN ('', '不分區') ");

            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add(new SqlParameter("SYR", string.IsNullOrEmpty(model.SYR) ? "" : model.SYR));
            cmd.Parameters.Add(new SqlParameter("SMN", string.IsNullOrEmpty(model.SMN) ? "" : model.SMN));

            if (!string.IsNullOrEmpty(model.SYR)) {
                sb.Append("AND S.SYR = @SYR2 ");
                cmd.Parameters.Add(new SqlParameter("SYR2", model.SYR));
            }
            if (!string.IsNullOrEmpty(model.SMN)) {
                sb.Append("AND S.SMN = @SMN2 ");
                cmd.Parameters.Add(new SqlParameter("SMN2", model.SMN));
            }

            sb.Append(@"GROUP BY CASE T.TD_CITY
                    WHEN '台北縣' THEN '新北市'
                    WHEN '台中縣' THEN '台中市'
                    WHEN '台南縣' THEN '台南市'
                    WHEN '高雄縣' THEN '高雄市'
                    WHEN '桃園縣' THEN '桃園市'
                    ELSE TD_CITY END, S.SYR, S.SMN
            ) S ON S.TD_CITY = P.CITY
            WHERE P.TOWN = '不分區' AND P.CITY <> '不分區' ");

            if (!string.IsNullOrEmpty(model.TD_CITY)) {
                sb.Append("AND P.CITY = @TD_CITY ");
                cmd.Parameters.Add(new SqlParameter("TD_CITY", model.TD_CITY));
            }

            sb.Append("ORDER BY P.CITY_ID, S.SYR, S.SMN");
            cmd.CommandText = sb.ToString();

            DataTable dt = SqlDao.CMBASEQuery4List(cmd);
            List<DataRow> list = dt.AsEnumerable().ToList();
            model.List_Tab1 = new PagedList<DataRow>(list, model.page1, pageSize);
        }

        /// <summary>
        /// 頁籤2: 化糞池拆除 (EVL_TYPE='02')
        /// </summary>
        private void LoadTab2Data(TDE0130Model model) {
            if (model.page2 == 0) model.page2 = 1;

            StringBuilder sb = new StringBuilder();
            sb.Append(@"SELECT P.CITY AS TD_CITY,
                CASE @SYR WHEN '' THEN S.SYR ELSE @SYR END AS SYR,
                CASE @SMN WHEN '' THEN S.SMN ELSE @SMN END AS SMN,
                ISNULL(S.YEAR_PRE_REMOVE_HOLDER_02, 0) AS YEAR_PRE_REMOVE_HOLDER_02,
                ISNULL(S.YEAR_PRE_REMOVE_PLACE_02, 0) AS YEAR_PRE_REMOVE_PLACE_02,
                ISNULL(S.PRE_REMOVE_HOLDER_02, 0) AS PRE_REMOVE_HOLDER_02,
                ISNULL(S.PRE_REMOVE_PLACE_02, 0) AS PRE_REMOVE_PLACE_02,
                ISNULL(S.REMOVE_HOLDER_RATE, 0) AS REMOVE_HOLDER_RATE,
                ISNULL(S.REMOVE_PLACE_RATE, 0) AS REMOVE_PLACE_RATE
            FROM CMBASE.dbo.SMCPOSTN P
            LEFT JOIN (
                SELECT CASE T.TD_CITY
                    WHEN '台北縣' THEN '新北市'
                    WHEN '台中縣' THEN '台中市'
                    WHEN '台南縣' THEN '台南市'
                    WHEN '高雄縣' THEN '高雄市'
                    WHEN '桃園縣' THEN '桃園市'
                    ELSE TD_CITY END AS TD_CITY,
                    S.SYR, S.SMN,
                    SUM(ISNULL(S.YEAR_PRE_REMOVE_HOLDER_02, 0)) AS YEAR_PRE_REMOVE_HOLDER_02,
                    SUM(ISNULL(S.YEAR_PRE_REMOVE_PLACE_02, 0)) AS YEAR_PRE_REMOVE_PLACE_02,
                    SUM(ISNULL(S.PRE_REMOVE_HOLDER_02, 0)) AS PRE_REMOVE_HOLDER_02,
                    SUM(ISNULL(S.PRE_REMOVE_PLACE_02, 0)) AS PRE_REMOVE_PLACE_02,
                    CASE SUM(ISNULL(S.YEAR_PRE_REMOVE_HOLDER_02, 0)) WHEN 0 THEN 0
                        ELSE SUM(ISNULL(S.PRE_REMOVE_HOLDER_02, 0)) * 100.00 / SUM(ISNULL(S.YEAR_PRE_REMOVE_HOLDER_02, 0)) END AS REMOVE_HOLDER_RATE,
                    CASE SUM(ISNULL(S.YEAR_PRE_REMOVE_PLACE_02, 0)) WHEN 0 THEN 0
                        ELSE SUM(ISNULL(S.PRE_REMOVE_PLACE_02, 0)) * 100.00 / SUM(ISNULL(S.YEAR_PRE_REMOVE_PLACE_02, 0)) END AS REMOVE_PLACE_RATE
                FROM CMETMS.dbo.TDDSEVL S
                LEFT JOIN CMETMS.dbo.TDMFORM T ON T.TD_CODE = S.TD_CODE
                WHERE EVL_TYPE = '02' AND ISNULL(T.TD_CITY, '') NOT IN ('', '不分區') ");

            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add(new SqlParameter("SYR", string.IsNullOrEmpty(model.SYR) ? "" : model.SYR));
            cmd.Parameters.Add(new SqlParameter("SMN", string.IsNullOrEmpty(model.SMN) ? "" : model.SMN));

            if (!string.IsNullOrEmpty(model.SYR)) {
                sb.Append("AND S.SYR = @SYR2 ");
                cmd.Parameters.Add(new SqlParameter("SYR2", model.SYR));
            }
            if (!string.IsNullOrEmpty(model.SMN)) {
                sb.Append("AND S.SMN = @SMN2 ");
                cmd.Parameters.Add(new SqlParameter("SMN2", model.SMN));
            }

            sb.Append(@"GROUP BY CASE T.TD_CITY
                    WHEN '台北縣' THEN '新北市'
                    WHEN '台中縣' THEN '台中市'
                    WHEN '台南縣' THEN '台南市'
                    WHEN '高雄縣' THEN '高雄市'
                    WHEN '桃園縣' THEN '桃園市'
                    ELSE TD_CITY END, S.SYR, S.SMN
            ) S ON S.TD_CITY = P.CITY
            WHERE P.TOWN = '不分區' AND P.CITY <> '不分區' ");

            if (!string.IsNullOrEmpty(model.TD_CITY)) {
                sb.Append("AND P.CITY = @TD_CITY ");
                cmd.Parameters.Add(new SqlParameter("TD_CITY", model.TD_CITY));
            }

            sb.Append("ORDER BY P.CITY_ID, S.SYR, S.SMN");
            cmd.CommandText = sb.ToString();

            DataTable dt = SqlDao.CMBASEQuery4List(cmd);
            List<DataRow> list = dt.AsEnumerable().ToList();
            model.List_Tab2 = new PagedList<DataRow>(list, model.page2, pageSize);
        }

        /// <summary>
        /// 取得年度下拉選單 (從資料庫動態取得)
        /// </summary>
        private List<SelectListItem> GetYearList() {
            List<SelectListItem> list = new List<SelectListItem>();
            try {
                SqlCommand cmd = new SqlCommand(@"SELECT DISTINCT SYR FROM CMETMS.dbo.TDDSEVL
                    WHERE SYR IS NOT NULL ORDER BY SYR DESC");
                DataTable dt = SqlDao.CMETMSQuery4List(cmd);
                foreach (DataRow row in dt.Rows) {
                    string syr = row["SYR"].ToString();
                    list.Add(new SelectListItem { Text = syr, Value = syr });
                }
            } catch {
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
                sb.Append("SELECT DISTINCT SMN FROM CMETMS.dbo.TDDSEVL WHERE SMN IS NOT NULL ");
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
                for (int i = 12; i >= 1; i--) {
                    list.Add(new SelectListItem { Text = i.ToString("00"), Value = i.ToString("00") });
                }
            }
            return list;
        }

        /// <summary>
        /// 取得縣市下拉選單
        /// </summary>
        private List<SelectListItem> GetCityList() {
            List<SelectListItem> list = new List<SelectListItem>();
            try {
                SqlCommand cmd = new SqlCommand(@"SELECT CITY FROM CMBASE.dbo.SMCPOSTN
                    WHERE TOWN = '不分區' AND CITY <> '不分區' ORDER BY CITY_ID");
                DataTable dt = SqlDao.CMBASEQuery4List(cmd);
                foreach (DataRow row in dt.Rows) {
                    string city = row["CITY"].ToString();
                    list.Add(new SelectListItem { Text = city, Value = city });
                }
            } catch {
                // 若查詢失敗，返回空列表
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
    }
}
