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

        // ======================================================================
        // 3.1.4rev3：分月進度「沒有資料時不知道怎麼產生」
        //   業主回報：進度管制的分月進度是空的，畫面上完全沒有新增按鈕；
        //             照提示去「工作控制點」填了預定開始日、預定結束日還是生不出來。
        //   查舊系統（C:\workspace\CPC\SEPMIS\App_Code\CMETMS\TDE0020.cs）後確認：
        //     ① 真正的產生入口在舊 TDE0020.aspx 主畫面的「新增年度 [YYY] [新增]」，
        //        按下去呼叫 NewTDDMSCD()，一次建立該年度 01~12 共 12 筆空白列。
        //        新版改寫成 MVC 時整個漏掉，所以永遠停在「無資料」。
        //     ② 舊 B042 子畫面那顆「新增」鈕只是把「現有最大年月 +1 個月」
        //        （Get_SYR_SMN_TDDMSCD），完全沒有讀預定開始／完成日；
        //        一筆資料都沒有時只會跳「無法取得年度與月份，請確認已填寫工程預定開始與完成日!!」
        //        —— 業主就是被這句話誤導才去填日期的，填了也不會有作用。
        //   本次作法：把 ① 補回來（PAGE5_2_AddYear），並依工程期程算出「建議建立的年度」，
        //   讓業主填在工作控制點的日期真的能決定要建立哪些年度，一鍵補齊。
        // ======================================================================

        // 期程推算的年度數上限，避免髒日期資料一次灌進上百個年度
        private const int MSCD_MAX_YEARS = 15;

        /// <summary>民國日期字串（YYYMMDD 或 YYMMDD）取出民國年；取不到回 null</summary>
        private static int? TwYearOf(object value) {
            if(value == null || value == DBNull.Value) { return null; }
            string s = new string(value.ToString().Where(char.IsDigit).ToArray());
            if(s.Length != 6 && s.Length != 7) { return null; }
            int y;
            if(!int.TryParse(s.Substring(0, s.Length - 4), out y)) { return null; }
            int nowTw = DateTime.Now.Year - 1911;
            // 民國 60 年以前與 10 年後的都當作髒資料
            return (y < 60 || y > nowTw + 10) ? (int?)null : y;
        }

        /// <summary>
        /// 依工程期程推算分月進度「建議建立」的年度（已建立的排除），供畫面預帶年度與一鍵補齊。
        /// 期程來源：TDDSDWB 的計畫期程摘要(IT_TYPE='S') 與各工作控制點的預定／實際完成日
        /// —— 也就是業主在「工作控制點」看得到、填得到的那些日期。
        /// </summary>
        private void SetMscdYearSuggestion(string TD_CODE) {
            int nowTw = DateTime.Now.Year - 1911;
            var suggest = new List<string>();
            var existing = new List<string>();

            try {
                SqlCommand cmdHas = new SqlCommand("SELECT DISTINCT SYR FROM TDDMSCD WHERE TD_CODE=@TD_CODE");
                cmdHas.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                DataTable dtHas = BASE_DAO.SqlDao.CMETMSQuery4List(cmdHas);
                if(dtHas != null) {
                    foreach(DataRow r in dtHas.Rows) {
                        int? y = TwYearOf((r["SYR"] ?? "").ToString() + "0101");
                        if(y.HasValue) { existing.Add(y.Value.ToString("000")); }
                    }
                }

                SqlCommand cmdDt = new SqlCommand(
                    "SELECT PRE_SDT, PRE_EDT, REL_SDT, REL_EDT FROM TDDSDWB WHERE TD_CODE=@TD_CODE");
                cmdDt.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                DataTable dtDate = BASE_DAO.SqlDao.CMETMSQuery4List(cmdDt);

                int min = int.MaxValue, max = int.MinValue;
                if(dtDate != null) {
                    foreach(DataRow r in dtDate.Rows) {
                        foreach(string col in new[] { "PRE_SDT", "PRE_EDT", "REL_SDT", "REL_EDT" }) {
                            int? y = TwYearOf(r[col]);
                            if(!y.HasValue) { continue; }
                            if(y.Value < min) { min = y.Value; }
                            if(y.Value > max) { max = y.Value; }
                        }
                    }
                }

                if(min != int.MaxValue) {
                    if(max - min + 1 > MSCD_MAX_YEARS) { max = min + MSCD_MAX_YEARS - 1; }
                    for(int y = min; y <= max; y++) {
                        string s = y.ToString("000");
                        if(!existing.Contains(s)) { suggest.Add(s); }
                    }
                }
            } catch(Exception ex) {
                // 推算失敗不能擋住畫面，退回「手動輸入年度」即可
                System.Diagnostics.Debug.WriteLine("SetMscdYearSuggestion: " + ex.Message);
            }

            // 給畫面用的都做成字串／數字，view 不必碰泛型集合
            ViewBag.MscdSuggestCount = suggest.Count;
            ViewBag.MscdSuggestCsv = string.Join(",", suggest);
            ViewBag.MscdSuggestLabel = suggest.Count > 0
                ? (suggest.Count == 1 ? suggest[0] : suggest[0] + "～" + suggest[suggest.Count - 1])
                : "";
            ViewBag.MscdDefaultYear = suggest.Count > 0 ? suggest[0] : nowTw.ToString("000");
        }

        /// <summary>
        /// 建立分月進度的年度資料：把該年度缺少的月份補成空白列（01~12），已存在的月份完全不動。
        /// SYR 可以是單一年度或以逗號分隔的多個年度（畫面的「依工程期程一次建立」用）。
        /// 對應舊系統 TDE0020.cs 的 NewTDDMSCD()，但刻意不沿用它「先刪掉整年再重建」的作法，
        /// 以免把業主已經填好的填報資料洗掉。
        /// </summary>
        [HttpPost]
        public ActionResult PAGE5_2_AddYear(string TD_CODE, string SYR) {
            try {
                if(string.IsNullOrWhiteSpace(TD_CODE)) {
                    return Json(new { success = false, message = "找不到工程編號，請關閉本頁後重新從工程清單進入。" });
                }

                int nowTw = DateTime.Now.Year - 1911;
                var years = new List<string>();
                foreach(string raw in (SYR ?? "").Split(',')) {
                    string t = raw.Trim();
                    if(t.Length == 0) { continue; }
                    int y;
                    if(!int.TryParse(t, out y) || y < 60 || y > nowTw + 10) {
                        return Json(new {
                            success = false,
                            message = "「" + t + "」不是有效的民國年度，請輸入 60～" + (nowTw + 10) + " 之間的數字（例如 " + nowTw + "）。"
                        });
                    }
                    string s = y.ToString("000");
                    if(!years.Contains(s)) { years.Add(s); }
                }
                if(years.Count == 0) {
                    return Json(new { success = false, message = "請先輸入要建立的年度（民國年，例如 " + nowTw + "）。" });
                }

                // TDDMSCD 的 ADD_DT/UPD_DT 是民國格式的字串欄位（varchar 8/7），不是 datetime
                string today = tools.GetDateByTWDate(DateTime.Now);
                string nowTm = DateTime.Now.ToString("HHmmss");
                string uid = WebSiteHelper.getUserId() ?? "";

                int created = 0;
                var done = new List<string>();
                var skipped = new List<string>();

                foreach(string y in years) {
                    SqlCommand cmdHas = new SqlCommand("SELECT SMN FROM TDDMSCD WHERE TD_CODE=@TD_CODE AND SYR=@SYR");
                    cmdHas.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                    cmdHas.Parameters.Add(new SqlParameter("SYR", y));
                    DataTable dtHas = BASE_DAO.SqlDao.CMETMSQuery4List(cmdHas);

                    var has = new List<string>();
                    if(dtHas != null) {
                        foreach(DataRow r in dtHas.Rows) {
                            has.Add((r["SMN"] ?? "").ToString().Trim().PadLeft(2, '0'));
                        }
                    }

                    var missing = new List<string>();
                    for(int m = 1; m <= 12; m++) {
                        string smn = m.ToString("00");
                        if(!has.Contains(smn)) { missing.Add(smn); }
                    }
                    if(missing.Count == 0) { skipped.Add(y); continue; }

                    SqlCommand ins = new SqlCommand();
                    ins.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                    ins.Parameters.Add(new SqlParameter("SYR", y));
                    ins.Parameters.Add(new SqlParameter("ADD_DT", today));
                    ins.Parameters.Add(new SqlParameter("UPD_DT", today));
                    ins.Parameters.Add(new SqlParameter("UPD_TM", nowTm));
                    ins.Parameters.Add(new SqlParameter("UPD_UID", uid));
                    var values = new List<string>();
                    for(int i = 0; i < missing.Count; i++) {
                        values.Add("(@TD_CODE,@SYR,@SMN" + i + ",@ADD_DT,@UPD_DT,@UPD_TM,@UPD_UID)");
                        ins.Parameters.Add(new SqlParameter("SMN" + i, missing[i]));
                    }
                    ins.CommandText = "INSERT INTO TDDMSCD (TD_CODE,SYR,SMN,ADD_DT,UPD_DT,UPD_TM,UPD_UID) VALUES "
                                    + string.Join(",", values);
                    BASE_DAO.SqlDao.CMETMSExecuteSql(ins);

                    created += missing.Count;
                    done.Add(y);
                }

                if(created == 0) {
                    return Json(new {
                        success = false,
                        message = "民國 " + string.Join("、", skipped) + " 年度的 1~12 月都已經建立過了，"
                                + "請直接在下方清單按「修改」填寫，資料不會被重複建立。"
                    });
                }

                string msg = "已建立民國 " + string.Join("、", done) + " 年度的分月進度，共 " + created
                           + " 個月，請在下方清單逐月按「修改」填寫。";
                if(skipped.Count > 0) {
                    msg += "（民國 " + string.Join("、", skipped) + " 年度原本就有資料，未做任何變更。）";
                }
                return Json(new { success = true, message = msg, syr = done[done.Count - 1] });
            } catch(Exception ex) {
                return Json(new { success = false, message = "建立年度資料失敗：" + ex.Message });
            }
        }

        // ======================================================================
        // 3.1.4rev3（辦理情形及施工進度）業主指示（含「補充說明」簡報）：
        //   ‧ Grid 的資料來源是 CMETMS.dbo.SP_TDE0051_TDGMSD（參數 TD_CODE、GET_SYR）
        //   ‧ Grid 開放寫入的欄位＝呼叫 CMETMS.dbo.SP_TDE0051_TDUMSD 的那些參數
        //   ‧ 年度下拉要多一個「所有年度」
        //   ‧ 分月進度上方要加一個「核定數／實支數／暫付款／實支+暫付」的經費彙總表
        //   ‧ 要有[新增]當年月份：當年度最後一個月 +1；當年度沒有值就從 01 加起
        // ======================================================================

        /// <summary>
        /// 分月進度清單：改由舊系統的 SP 取數（SYR 傳空字串＝所有年度）。
        /// SP 沒有回傳 TDDMSCD.ID，這裡把 ID 併回來，才能沿用既有的「修改」入口。
        /// </summary>
        private DataTable GetMscdFromSp(string TD_CODE, string SYR) {
            DataTable dt;
            try {
                SqlCommand cmd = new SqlCommand("SP_TDE0051_TDGMSD");
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 180;   // SP 內是 cursor 迴圈，舊系統也把 timeout 開得很大
                cmd.Parameters.Add(new SqlParameter("@TD_CODE", SqlDbType.VarChar, 20)).Value = TD_CODE ?? "";
                cmd.Parameters.Add(new SqlParameter("@GET_SYR", SqlDbType.VarChar, 3)).Value = SYR ?? "";
                SqlParameter vrf = cmd.Parameters.Add("@VRF_FG", SqlDbType.VarChar, 1);
                vrf.Direction = ParameterDirection.Output;
                dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
                ViewBag.MscdWarning = null;
            } catch(Exception ex) {
                // SP 內用 SMN = convert(int,SMN)-1 這種寫法，資料只要有沒補零的月份（'1' 與 '01' 並存）
                // 或同年月重複，子查詢就會傳回多筆而整支 SP 失敗。
                // 這種情況不能讓整個畫面 500，改為顯示原始資料並提示原因。
                System.Diagnostics.Debug.WriteLine("SP_TDE0051_TDGMSD: " + ex.Message);
                ViewBag.MscdWarning = BuildMscdDataWarning(TD_CODE);
                return GetMscdRaw(TD_CODE, SYR);
            }
            if(dt == null) { return null; }

            if(!dt.Columns.Contains("ID")) { dt.Columns.Add("ID", typeof(string)); }
            SqlCommand idCmd = new SqlCommand("SELECT ID, SYR, SMN FROM TDDMSCD WHERE TD_CODE=@TD_CODE");
            idCmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dtId = BASE_DAO.SqlDao.CMETMSQuery4List(idCmd);
            var map = new Dictionary<string, string>();
            if(dtId != null) {
                foreach(DataRow r in dtId.Rows) {
                    map[MscdKey(r["SYR"], r["SMN"])] = r["ID"].ToString();
                }
            }
            foreach(DataRow r in dt.Rows) {
                string id;
                r["ID"] = map.TryGetValue(MscdKey(r["SYR"], r["SMN"]), out id) ? id : "";
            }
            return dt;
        }

        private static string MscdKey(object syr, object smn) {
            return (syr ?? "").ToString().Trim() + "-" + (smn ?? "").ToString().Trim().PadLeft(2, '0');
        }

        /// <summary>SP 算不出來時的退路：直接讀 TDDMSCD（欄位另取 SUM_ 別名對齊畫面）</summary>
        private static DataTable GetMscdRaw(string TD_CODE, string SYR) {
            SqlCommand cmd = new SqlCommand(
                @"SELECT *, PRE_SCHD AS SUM_PRE_SCHD, REL_SCHD AS SUM_REL_SCHD
                  FROM TDDMSCD WHERE TD_CODE=@TD_CODE AND (@SYR='' OR SYR=@SYR) ORDER BY SYR, SMN");
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            cmd.Parameters.Add(new SqlParameter("SYR", SYR ?? ""));
            return BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
        }

        /// <summary>找出害 SP 算不出來的月份資料，翻成使用者看得懂的說明</summary>
        private static string BuildMscdDataWarning(string TD_CODE) {
            var bad = new List<string>();
            try {
                SqlCommand cmd = new SqlCommand(
                    @"SELECT SYR, SMN, COUNT(*) AS N, MIN(LEN(SMN)) AS MINLEN
                      FROM TDDMSCD WHERE TD_CODE=@TD_CODE
                      GROUP BY SYR, SMN HAVING COUNT(*) > 1 OR MIN(LEN(SMN)) < 2
                      ORDER BY SYR, SMN");
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
                if(dt != null) {
                    foreach(DataRow r in dt.Rows) {
                        bad.Add("民國 " + r["SYR"] + " 年 " + r["SMN"] + " 月");
                    }
                }
            } catch(Exception ex) {
                System.Diagnostics.Debug.WriteLine("BuildMscdDataWarning: " + ex.Message);
            }

            string msg = "這個工程的分月進度資料裡有月份格式不正確或重複的紀錄，"
                       + "舊系統的累計計算程式無法處理，因此下方改為顯示未經累計的原始資料"
                       + "（年度累計、總累計等欄位會是 0）。";
            if(bad.Count > 0) {
                msg += "問題資料：" + string.Join("、", bad.Take(10)) + (bad.Count > 10 ? " 等" : "") + "。";
            }
            return msg + "請聯絡系統管理員協助修正後即可恢復正常。";
        }

        /// <summary>CMEPMS 單一數值查詢，查不到或為 NULL 一律回 0</summary>
        private static decimal CmepmsScalar(string sql, params SqlParameter[] ps) {
            SqlCommand cmd = new SqlCommand(sql);
            foreach(SqlParameter p in ps) { cmd.Parameters.Add(p); }
            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            if(dt == null || dt.Rows.Count == 0 || dt.Rows[0][0] == DBNull.Value) { return 0m; }
            decimal v;
            return decimal.TryParse(dt.Rows[0][0].ToString(), out v) ? v : 0m;
        }

        /// <summary>
        /// 3.1.4rev3：分月進度上方的經費彙總（公務預算／振興經濟 × 核定數／實支數／暫付款）。
        /// SQL 完全照業主「補充說明」提供的舊系統程式碼：
        ///   ‧ 沒有選年度（所有年度）：核定數＝SUM(TDV_TDDCFTS.YR_FUND)，實支與暫付以 0 呈現
        ///   ‧ 有選年度：核定數＝TDV_TDDCFTS.CEN_FUND，實支＝MAX(TDDSTCF.COST_FEE)，暫付＝MAX(TDDSTCF.PPAY_FEE)
        /// </summary>
        private void SetFundSummary(string TD_CODE, string SYR) {
            bool allYears = string.IsNullOrWhiteSpace(SYR);
            var a = new decimal[2];   // 核定數
            var b = new decimal[2];   // 實支數
            var c = new decimal[2];   // 暫付款
            string[] pjNames = { "%公務預算%", "%振興經濟%" };

            for(int i = 0; i < 2; i++) {
                try {
                    if(allYears) {
                        a[i] = CmepmsScalar(
                            @"SELECT SUM(b.YR_FUND) FROM PJMFORM a, TDV_TDDCFTS b
                              WHERE a.PJ_NAME LIKE @PJ AND b.PID=a.PID AND b.TD_CODE=@TD_CODE",
                            new SqlParameter("PJ", pjNames[i]), new SqlParameter("TD_CODE", TD_CODE));
                    } else {
                        a[i] = CmepmsScalar(
                            @"SELECT b.CEN_FUND FROM PJMFORM a, TDV_TDDCFTS b
                              WHERE a.PJ_NAME LIKE @PJ AND a.SYR=@SYR AND b.PID=a.PID AND b.TD_CODE=@TD_CODE",
                            new SqlParameter("PJ", pjNames[i]), new SqlParameter("SYR", SYR),
                            new SqlParameter("TD_CODE", TD_CODE));
                        b[i] = CmepmsScalar(
                            @"SELECT MAX(c.COST_FEE) FROM PJMFORM a, CTDMNDS b, TDDSTCF c
                              WHERE a.PJ_NAME LIKE @PJ AND a.SYR=@SYR AND a.PID=b.PID AND a.SYR=c.SYR
                                AND b.TD_CODE=@TD_CODE AND b.TD_CODE=c.TD_CODE1 AND b.FD_CODE=c.FD_CODE",
                            new SqlParameter("PJ", pjNames[i]), new SqlParameter("SYR", SYR),
                            new SqlParameter("TD_CODE", TD_CODE));
                        c[i] = CmepmsScalar(
                            @"SELECT MAX(c.PPAY_FEE) FROM PJMFORM a, CTDMNDS b, TDDSTCF c
                              WHERE a.PJ_NAME LIKE @PJ AND a.SYR=@SYR AND a.PID=b.PID AND a.SYR=c.SYR
                                AND b.TD_CODE=@TD_CODE AND b.TD_CODE=c.TD_CODE1 AND b.FD_CODE=c.FD_CODE",
                            new SqlParameter("PJ", pjNames[i]), new SqlParameter("SYR", SYR),
                            new SqlParameter("TD_CODE", TD_CODE));
                    }
                } catch(Exception ex) {
                    // 彙總算不出來不能擋住整個畫面
                    System.Diagnostics.Debug.WriteLine("SetFundSummary: " + ex.Message);
                }
            }

            ViewBag.Fund1A = a[0]; ViewBag.Fund1B = b[0]; ViewBag.Fund1C = c[0];
            ViewBag.Fund2A = a[1]; ViewBag.Fund2B = b[1]; ViewBag.Fund2C = c[1];
        }

        public ActionResult PAGE5_2(string TD_CODE, string SYR) {
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

            // 查詢 TDDMSCD 年度列表（用於下拉選單；畫面另外會補一個「所有年度」）
            string sqlSyr = @"SELECT DISTINCT SYR FROM CMETMS.dbo.TDDMSCD WHERE TD_CODE=@TD_CODE ORDER BY SYR DESC";
            SqlCommand cmdSyr = new SqlCommand(sqlSyr);
            cmdSyr.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable dtSyr = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSyr);
            ViewBag.SyrList = dtSyr;

            // 預設顯示最新年度（SYR 傳空字串代表「所有年度」，是合法值，所以用 null 判斷）
            if(SYR == null && dtSyr != null && dtSyr.Rows.Count > 0) {
                SYR = dtSyr.Rows[0]["SYR"].ToString();
            }
            SYR = SYR ?? "";
            ViewBag.SelectedSyr = SYR;
            ViewBag.MscdData = GetMscdFromSp(TD_CODE, SYR);
            SetFundSummary(TD_CODE, SYR);

            // 3.1.4rev3：算出「建議建立的年度」，畫面才能預帶年度並提供一鍵補齊
            SetMscdYearSuggestion(TD_CODE);

            TDE0040Model model = new TDE0040Model {
                TDMFORM = tdmform,
                TD_CODE = TD_CODE
            };

            GetSelectList();
            return View("PAGE5_2", model);
        }

        // AJAX: 根據年度查詢分月進度（含上方經費彙總，兩者都要跟著年度變動）
        [HttpPost]
        public ActionResult PAGE5_2_GetMscdBySyr(string TD_CODE, string SYR) {
            try {
                DataTable dt = GetMscdFromSp(TD_CODE, SYR ?? "");

                var result = new List<Dictionary<string, object>>();
                if(dt != null) {
                    foreach(DataRow row in dt.Rows) {
                        var dict = new Dictionary<string, object>();
                        foreach(DataColumn col in dt.Columns) {
                            dict[col.ColumnName] = row[col] == DBNull.Value ? "" : row[col];
                        }
                        result.Add(dict);
                    }
                }

                SetFundSummary(TD_CODE, SYR ?? "");
                return Json(new {
                    success = true,
                    data = result,
                    warning = ViewBag.MscdWarning as string,
                    fund = new {
                        a1 = ViewBag.Fund1A, b1 = ViewBag.Fund1B, c1 = ViewBag.Fund1C,
                        a2 = ViewBag.Fund2A, b2 = ViewBag.Fund2B, c2 = ViewBag.Fund2C
                    }
                });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // POST: 3.1.4rev3 新增當年月份（TDDMSCD）
        [HttpPost]
        public ActionResult PAGE5_2_AddMonth(string TD_CODE, string SYR) {
            try {
                if(string.IsNullOrWhiteSpace(TD_CODE)) {
                    return Json(new { success = false, message = "找不到工程編號，請關閉本頁後重新從工程清單進入。" });
                }
                if(string.IsNullOrWhiteSpace(SYR)) {
                    return Json(new { success = false, message = "請先在左上角選擇一個年度（「所有年度」無法判斷要新增哪一年）。" });
                }

                SqlCommand cmdHas = new SqlCommand("SELECT SMN FROM TDDMSCD WHERE TD_CODE=@TD_CODE AND SYR=@SYR");
                cmdHas.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmdHas.Parameters.Add(new SqlParameter("SYR", SYR));
                string smn = NextSmnOf(BASE_DAO.SqlDao.CMETMSQuery4List(cmdHas));
                if(smn == null) {
                    return Json(new {
                        success = false,
                        message = "民國 " + SYR + " 年度的 1~12 月都已經建立完成，沒有月份可以再新增了。"
                    });
                }

                SqlCommand ins = new SqlCommand(
                    @"INSERT INTO TDDMSCD (TD_CODE, SYR, SMN, ADD_DT, UPD_DT, UPD_TM, UPD_UID)
                      VALUES (@TD_CODE, @SYR, @SMN, @ADD_DT, @UPD_DT, @UPD_TM, @UPD_UID)");
                ins.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                ins.Parameters.Add(new SqlParameter("SYR", SYR));
                ins.Parameters.Add(new SqlParameter("SMN", smn));
                ins.Parameters.Add(new SqlParameter("ADD_DT", tools.GetDateByTWDate(DateTime.Now)));
                ins.Parameters.Add(new SqlParameter("UPD_DT", tools.GetDateByTWDate(DateTime.Now)));
                ins.Parameters.Add(new SqlParameter("UPD_TM", DateTime.Now.ToString("HHmmss")));
                ins.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));
                BASE_DAO.SqlDao.CMETMSExecuteSql(ins);

                return Json(new {
                    success = true,
                    message = "已新增民國 " + SYR + " 年度 " + smn + " 月，請按該列的「修改」填寫。",
                    syr = SYR, smn = smn
                });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增月份失敗：" + ex.Message });
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
        // 3.1.4rev3：業主指示「Grid 開放寫入的資料欄位是呼叫 stored procedure
        //            CMETMS.dbo.SP_TDE0051_TDUMSD」，故改為呼叫該 SP，
        //            參數清單即為開放填報的欄位（原本這裡是自己組 UPDATE，
        //            欄位也和舊系統不一致）。SP 會回傳檢核碼，要翻成使用者看得懂的話。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE5_2_Edit_Save(string ID, string TD_CODE, string SYR, string SMN,
            decimal? ECC_UPAY_FEE, decimal? PRE_COST,
            decimal? L_PLN_COST, decimal? L_REL_COST, decimal? L_UNPAY_COST, decimal? SAVE_FEE,
            decimal? PRE_CON_LEN, decimal? REL_CON_LEN, decimal? PRE_CON_FLOOD, decimal? REL_CON_FLOOD,
            decimal? SUM_PRE_SCHD, decimal? SUM_REL_SCHD, decimal? ORG_SUM_PRE_SCHD,
            string EXECT_DESC, string DELAY_RESN, string RESLV_MTHD,
            string CTDELAY_RESN, string CTRESLV_MTHD, string RESLV_LMT_DT) {

            try {
                SqlCommand cmd = new SqlCommand("SP_TDE0051_TDUMSD");
                cmd.CommandType = CommandType.StoredProcedure;
                SqlParameter rtn = cmd.Parameters.Add("@RETURN_VALUE", SqlDbType.Int);
                rtn.Direction = ParameterDirection.ReturnValue;

                cmd.Parameters.Add(new SqlParameter("@TD_CODE", SqlDbType.VarChar, 20)).Value = TD_CODE ?? "";
                cmd.Parameters.Add(new SqlParameter("@SYR", SqlDbType.VarChar, 3)).Value = SYR ?? "";
                cmd.Parameters.Add(new SqlParameter("@SMN", SqlDbType.VarChar, 2)).Value = SMN ?? "";
                cmd.Parameters.AddWithValue("@ECC_UPAY_FEE", ECC_UPAY_FEE ?? 0m);
                cmd.Parameters.AddWithValue("@PRE_COST", PRE_COST ?? 0m);
                cmd.Parameters.AddWithValue("@L_PLN_COST", L_PLN_COST ?? 0m);
                cmd.Parameters.AddWithValue("@L_REL_COST", L_REL_COST ?? 0m);
                cmd.Parameters.AddWithValue("@L_UNPAY_COST", L_UNPAY_COST ?? 0m);
                cmd.Parameters.AddWithValue("@SAVE_FEE", SAVE_FEE ?? 0m);
                cmd.Parameters.AddWithValue("@PRE_CON_LEN", PRE_CON_LEN ?? 0m);
                cmd.Parameters.AddWithValue("@REL_CON_LEN", REL_CON_LEN ?? 0m);
                cmd.Parameters.AddWithValue("@PRE_CON_FLOOD", PRE_CON_FLOOD ?? 0m);
                cmd.Parameters.AddWithValue("@REL_CON_FLOOD", REL_CON_FLOOD ?? 0m);
                cmd.Parameters.AddWithValue("@SUM_PRE_SCHD", SUM_PRE_SCHD ?? 0m);
                cmd.Parameters.AddWithValue("@SUM_REL_SCHD", SUM_REL_SCHD ?? 0m);
                cmd.Parameters.AddWithValue("@ORG_SUM_PRE_SCHD", ORG_SUM_PRE_SCHD ?? 0m);
                cmd.Parameters.Add(new SqlParameter("@EXECT_DESC", SqlDbType.NVarChar, 2000)).Value = (object)EXECT_DESC ?? "";
                cmd.Parameters.Add(new SqlParameter("@DELAY_RESN", SqlDbType.NVarChar, 2000)).Value = (object)DELAY_RESN ?? "";
                cmd.Parameters.Add(new SqlParameter("@RESLV_MTHD", SqlDbType.NVarChar, 2000)).Value = (object)RESLV_MTHD ?? "";
                cmd.Parameters.Add(new SqlParameter("@CTDELAY_RESN", SqlDbType.NVarChar, 2000)).Value = (object)CTDELAY_RESN ?? "";
                cmd.Parameters.Add(new SqlParameter("@CTRESLV_MTHD", SqlDbType.NVarChar, 2000)).Value = (object)CTRESLV_MTHD ?? "";
                cmd.Parameters.Add(new SqlParameter("@RESLV_LMT_DT", SqlDbType.VarChar, 10)).Value = (object)RESLV_LMT_DT ?? "";
                cmd.Parameters.Add(new SqlParameter("@UPD_UID", SqlDbType.VarChar, 20)).Value = WebSiteHelper.getUserId() ?? "";

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                int rc = (rtn.Value == null || rtn.Value == DBNull.Value) ? -1 : Convert.ToInt32(rtn.Value);

                if(rc == 0) {
                    // SP 內是 UPD_DT = getdate()，但 TDDMSCD.UPD_DT 是民國格式的 varchar(7)，
                    // 直接寫 datetime 會被截成「07」這類無效值，所以補寫正確格式
                    SqlCommand fix = new SqlCommand(
                        @"UPDATE TDDMSCD SET UPD_DT=@UPD_DT, UPD_TM=@UPD_TM WHERE ID=@ID");
                    fix.Parameters.Add(new SqlParameter("ID", ID));
                    fix.Parameters.Add(new SqlParameter("UPD_DT", tools.GetDateByTWDate(DateTime.Now)));
                    fix.Parameters.Add(new SqlParameter("UPD_TM", DateTime.Now.ToString("HHmmss")));
                    BASE_DAO.SqlDao.CMETMSExecuteSql(fix);

                    TempData["Success"] = "儲存成功";
                    TempData["SaveSuccess"] = true;
                } else {
                    TempData["Error"] = MscdSaveMessage(rc, SYR, SMN);
                }
            } catch(Exception ex) {
                TempData["Error"] = "儲存失敗：" + ex.Message;
            }

            return RedirectToAction("PAGE5_2_Edit", new { ID = ID });
        }

        /// <summary>把 SP_TDE0051_TDUMSD 的檢核代碼翻成使用者看得懂的訊息</summary>
        private static string MscdSaveMessage(int rc, string SYR, string SMN) {
            string ym = "民國 " + SYR + " 年 " + SMN + " 月";
            switch(rc) {
                case 10: return "找不到這筆工程的基本資料，無法儲存，請聯絡系統管理員。";
                case 20: return "進度落後時必須填寫「落後原因或辦理情形說明」與「落後解決對策」，請補填後再儲存。";
                case 30: return ym + "的「總累計預定進度」不可以小於上一個月，請確認數字是否填錯。";
                case 40: return ym + "的「總累計實際進度」不可以小於上一個月，請確認數字是否填錯。";
                default: return "儲存失敗（檢核代碼 " + rc + "），請確認填寫內容或聯絡系統管理員。";
            }
        }

        // ======================================================================
        // 3.1.4rev3（地方預算執行數填報）：業主指示「此畫面與前一版本相同未修改到／資料來源 TDDPCOS」。
        //   rev2 誤把本頁的分月進度也接到 TDDMSCD，但舊系統 CMETMS\TDE0020B043.aspx.cs 用的是
        //   **CMETMS.dbo.TDDPCOS**（鍵＝TD_CODE + PID + SYR + SMN，另存 FD_CODE），
        //   而且是「先在上方選一個計畫(PID)，下方才顯示該計畫的分月執行數」。
        //   上方清單 SQL 也補回舊系統的 PJ_TYPE='B' 與年度升冪排序。
        // ======================================================================
        public ActionResult PAGE5_3(string TD_CODE, string PID, string SYR) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            // 上方計畫清單（舊系統 B043：PJ_TYPE='B'、依年度升冪）
            string sql = @"SELECT DISTINCT P.SYR, P.AN_PJ_NAME, P.PJ_CODE, P.PJ_NAME, P.PID
     FROM CMEPMS.dbo.PJV_PJDPCAT P
    WHERE P.PJ_TYPE='B' AND P.PID IN (SELECT PID FROM CMEPMS.dbo.CTDMNDS WHERE TD_CODE=@TD_CODE)
 ORDER BY P.SYR, P.PID";
            SqlCommand cmd = new SqlCommand(sql);
   cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
    DataTable dtList = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
         ViewBag.ListData = dtList;

            // 未指定計畫時，優先選「已經有填報資料」的最新計畫；都沒有才取最新年度那一筆
            if(string.IsNullOrWhiteSpace(PID) && dtList != null && dtList.Rows.Count > 0) {
                SqlCommand cmdHasPid = new SqlCommand(
                    "SELECT DISTINCT PID FROM TDDPCOS WHERE TD_CODE=@TD_CODE");
                cmdHasPid.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                DataTable dtHasPid = BASE_DAO.SqlDao.CMETMSQuery4List(cmdHasPid);
                var withData = new List<string>();
                if(dtHasPid != null) {
                    foreach(DataRow r in dtHasPid.Rows) { withData.Add(r["PID"].ToString()); }
                }
                for(int i = dtList.Rows.Count - 1; i >= 0; i--) {
                    string pid = dtList.Rows[i]["PID"].ToString();
                    if(withData.Contains(pid)) { PID = pid; break; }
                }
                if(string.IsNullOrWhiteSpace(PID)) {
                    PID = dtList.Rows[dtList.Rows.Count - 1]["PID"].ToString();
                }
            }
            ViewBag.SelectedPid = PID ?? "";

            // 該計畫的年度清單（來源 TDDPCOS）
            DataTable dtSyr = null;
            if(!string.IsNullOrWhiteSpace(PID)) {
                SqlCommand cmdSyr = new SqlCommand(
                    "SELECT DISTINCT SYR FROM TDDPCOS WHERE TD_CODE=@TD_CODE AND PID=@PID ORDER BY SYR DESC");
                cmdSyr.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmdSyr.Parameters.Add(new SqlParameter("PID", PID));
                dtSyr = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSyr);
            }
       ViewBag.SyrList = dtSyr;

            if(dtSyr != null && dtSyr.Rows.Count > 0) {
                if(string.IsNullOrWhiteSpace(SYR)) { SYR = dtSyr.Rows[0]["SYR"].ToString(); }
                ViewBag.MscdData = GetPcos(TD_CODE, PID, SYR);
                ViewBag.SelectedSyr = SYR;
            } else if(!string.IsNullOrWhiteSpace(PID) && dtList != null) {
                // 這個計畫還沒有任何分月資料：年度取計畫本身的年度（舊系統也是從 PJMFORM.SYR 取），
                // 這樣「新增當年月份」才有年度可用，會從 01 月開始建立
                foreach(DataRow r in dtList.Rows) {
                    if(r["PID"].ToString() == PID) { ViewBag.SelectedSyr = r["SYR"].ToString(); break; }
                }
            }

         TDE0040Model model = new TDE0040Model {
           TDMFORM = tdmform,
        TD_CODE = TD_CODE
     };

        GetSelectList();
            return View("PAGE5_3", model);
        }

        // 地方預算執行數（TDDPCOS）某計畫某年度的分月資料
        private static DataTable GetPcos(string TD_CODE, string PID, string SYR) {
            SqlCommand cmd = new SqlCommand(
                @"SELECT ID, SYR, SMN, ECC_UPAY_FEE, PRE_COST, REL_COST,
                         L_PLN_COST, L_REL_COST, L_UNPAY_COST, SAVE_FEE
                  FROM TDDPCOS WHERE TD_CODE=@TD_CODE AND PID=@PID AND SYR=@SYR ORDER BY SMN");
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            cmd.Parameters.Add(new SqlParameter("PID", PID));
            cmd.Parameters.Add(new SqlParameter("SYR", SYR));
            return BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
        }

        // AJAX: PAGE5_3 依計畫與年度查詢 TDDPCOS
        [HttpPost]
        public ActionResult PAGE5_3_GetPcos(string TD_CODE, string PID, string SYR) {
            try {
                DataTable dt = GetPcos(TD_CODE, PID, SYR);
                var result = new List<Dictionary<string, object>>();
                if(dt != null) {
                    foreach(DataRow row in dt.Rows) {
                        var dict = new Dictionary<string, object>();
                        foreach(DataColumn col in dt.Columns) {
                            dict[col.ColumnName] = row[col] == DBNull.Value ? "" : row[col];
                        }
                        result.Add(dict);
                    }
                }
                return Json(new { success = true, data = result });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// 3.1.4rev3：算出「新增當年月份」要建立的月份。
        /// 業主規定：預設為當年度最後一個月往後加一個月；當年度沒有任何資料就從 01 加起。
        /// 已經到 12 月則回 null（該年度沒有月份可再新增）。
        /// </summary>
        private static string NextSmnOf(DataTable existing) {
            int max = 0;
            if(existing != null) {
                foreach(DataRow r in existing.Rows) {
                    int v;
                    if(int.TryParse((r["SMN"] ?? "").ToString().Trim(), out v) && v > max) { max = v; }
                }
            }
            return max >= 12 ? null : (max + 1).ToString("00");
        }

        // POST: PAGE5_3 新增當年月份（TDDPCOS）
        [HttpPost]
        public ActionResult PAGE5_3_AddMonth(string TD_CODE, string PID, string SYR) {
            try {
                if(string.IsNullOrWhiteSpace(TD_CODE) || string.IsNullOrWhiteSpace(PID)) {
                    return Json(new { success = false, message = "請先在上方選擇一個計畫。" });
                }
                if(string.IsNullOrWhiteSpace(SYR)) {
                    return Json(new { success = false, message = "找不到年度，請確認這個計畫已有起始年度。" });
                }

                SqlCommand cmdHas = new SqlCommand(
                    "SELECT SMN FROM TDDPCOS WHERE TD_CODE=@TD_CODE AND PID=@PID AND SYR=@SYR");
                cmdHas.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmdHas.Parameters.Add(new SqlParameter("PID", PID));
                cmdHas.Parameters.Add(new SqlParameter("SYR", SYR));
                string smn = NextSmnOf(BASE_DAO.SqlDao.CMETMSQuery4List(cmdHas));
                if(smn == null) {
                    return Json(new {
                        success = false,
                        message = "民國 " + SYR + " 年度的 1~12 月都已經建立完成，沒有月份可以再新增了。"
                    });
                }

                // FD_CODE 比照舊系統由 CTDMNDS 取得（同一工程同一計畫的經費代碼）
                SqlCommand ins = new SqlCommand(
                    @"INSERT INTO TDDPCOS (PID, TD_CODE, FD_CODE, SYR, SMN,
                          ECC_UPAY_FEE, PRE_COST, REL_COST, DEL_FG, ADD_DT, UPD_DT, UPD_TM, UPD_UID,
                          L_PLN_COST, L_REL_COST, L_UNPAY_COST, NG_FG, SAVE_FEE)
                      VALUES (@PID, @TD_CODE,
                          (SELECT MAX(FD_CODE) FROM CMEPMS.dbo.CTDMNDS WHERE TD_CODE=@TD_CODE AND PID=@PID),
                          @SYR, @SMN, 0, 0, 0, '1', GETDATE(), GETDATE(), @UPD_TM, @UPD_UID, 0, 0, 0, 0, 0)");
                ins.Parameters.Add(new SqlParameter("PID", PID));
                ins.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                ins.Parameters.Add(new SqlParameter("SYR", SYR));
                ins.Parameters.Add(new SqlParameter("SMN", smn));
                ins.Parameters.Add(new SqlParameter("UPD_TM", DateTime.Now.ToString("HHmmss")));
                ins.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));
                BASE_DAO.SqlDao.CMETMSExecuteSql(ins);

                return Json(new {
                    success = true,
                    message = "已新增民國 " + SYR + " 年度 " + smn + " 月，請按該列的「修改」填寫金額。",
                    syr = SYR, smn = smn
                });
            } catch(Exception ex) {
                return Json(new { success = false, message = "新增月份失敗：" + ex.Message });
            }
        }

        // GET: 地方預算執行數 單月填報（只開放舊系統 B043 有開的欄位）
        public ActionResult PAGE5_3_PcosEdit(string ID) {
            if(string.IsNullOrEmpty(ID)) {
                return Content("ID 不可為空");
            }
            SqlCommand cmd = new SqlCommand("SELECT * FROM TDDPCOS WHERE ID=@ID");
            cmd.Parameters.Add(new SqlParameter("ID", ID));
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            ViewBag.ID = ID;
            ViewBag.Data = (dt != null && dt.Rows.Count > 0) ? dt.Rows[0] : null;
            return View("PAGE5_3_PcosEdit");
        }

        // POST: 地方預算執行數 單月填報儲存
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE5_3_PcosEdit_Save(string ID,
            decimal? ECC_UPAY_FEE, decimal? PRE_COST, decimal? REL_COST,
            decimal? L_PLN_COST, decimal? L_REL_COST, decimal? L_UNPAY_COST, decimal? SAVE_FEE) {
            try {
                SqlCommand cmd = new SqlCommand(
                    @"UPDATE TDDPCOS SET
                        ECC_UPAY_FEE=@ECC_UPAY_FEE, PRE_COST=@PRE_COST, REL_COST=@REL_COST,
                        L_PLN_COST=@L_PLN_COST, L_REL_COST=@L_REL_COST,
                        L_UNPAY_COST=@L_UNPAY_COST, SAVE_FEE=@SAVE_FEE,
                        UPD_DT=GETDATE(), UPD_TM=@UPD_TM, UPD_UID=@UPD_UID
                      WHERE ID=@ID");
                cmd.Parameters.Add(new SqlParameter("ID", ID));
                cmd.Parameters.Add(new SqlParameter("ECC_UPAY_FEE", (object)ECC_UPAY_FEE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("PRE_COST", (object)PRE_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("REL_COST", (object)REL_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("L_PLN_COST", (object)L_PLN_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("L_REL_COST", (object)L_REL_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("L_UNPAY_COST", (object)L_UNPAY_COST ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("SAVE_FEE", (object)SAVE_FEE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("UPD_TM", DateTime.Now.ToString("HHmmss")));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));
                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                TempData["Success"] = "儲存成功";
                TempData["SaveSuccess"] = true;
            } catch(Exception ex) {
                TempData["Error"] = "儲存失敗：" + ex.Message;
            }
            return RedirectToAction("PAGE5_3_PcosEdit", new { ID = ID });
        }

        // AJAX: PAGE5_3 根據年度查詢 TDDMSCD 資料（3.1.4rev3 起本頁改用 TDDPCOS，保留供相容）
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