using Newtonsoft.Json.Linq;
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
using WCals.Models.CMETMSModels;
using WCals.Models.CMPSGSModels;
using WCals.ViewModels.NewModels;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS {
    public class TDE0040Controller :Controller {
        // GET: TDE0040
        public ActionResult Index(TDE0040Model model) {

            GetSelectList();
            return View("TDE0040Index", model);
        }

        int pageSize = 10;

        [HttpPost]
        public ActionResult Test(TDE0040Model model) {

            System.Diagnostics.Debug.WriteLine(model.CITY);
            return Index(model);
        }

        [HttpPost]
        public ActionResult Search(TDE0040Model model)
        {
            if (model.page1 == 0)
            {
                model.page1 = 1;
            }

            string userId = WebSiteHelper.getUserId();

            string userDepts = GetAllChildDeptsStr(WebSiteHelper.getUserDept());

            // 1. 基底查詢：移除 INNER JOIN，確保沒綁計畫的工程不會消失
            StringBuilder sb = new StringBuilder(@"
        SELECT 
            t.PJ_NAME, 
            t.TD_CODE, 
            t.TD_NAME, 
            t.ADV_DP_DESC, 
            t.CATY_CODE_DESC, 
            t.TD_STAT_DESC, 
            t.SYS_ID_FG,
            t.PID_SYR,
            ps.SYS_NAME,
            pf.PJ_NAME PF_PJ_NAME
        FROM TDV_TDDCFLS_Lite t
        LEFT JOIN CMETMS.dbo.PJCSYST ps ON ps.SYS_ID = t.SYS_ID
        LEFT JOIN CMETMS.dbo.PJCFORM pf ON pf.PJ_ID = t.PJ_ID
        WHERE 1=1 ");

            SqlCommand cmd = new SqlCommand();

            bool hasPidAccess = false;
            if (!string.IsNullOrEmpty(model.PID))
            {
                string checkSql = $@"
        SELECT TOP 1 1 FROM TDV_TDDCFLS_Lite 
        WHERE PID = @PID AND (ADV_DP IN ({userDepts}))";
                SqlCommand checkCmd = new SqlCommand(checkSql);
                checkCmd.Parameters.Add(new SqlParameter("PID", model.PID));
                DataTable dtCheck = BASE_DAO.SqlDao.CMEPMSQuery4List(checkCmd);
                hasPidAccess = (dtCheck != null && dtCheck.Rows.Count > 0);
            }

            // 2. 組合主查詢 SQL
            sb.AppendFormat(@" AND (
    -- 條件 A：屬於我管轄單位的案子 (基本權限)
    (t.ADV_DP IN ({0}) AND t.ADV_UID = @UserId)
    OR t.TD_CODE IN (SELECT TD_CODE FROM CMETMS.dbo.TDV_TDMADMC WHERE ADV_DP IN ({0}))
    
    -- 條件 B：【安全穿透】我有參與這個計畫，所以我可以看到該計畫的所有子案
    OR (@HasPidAccess = 1 AND t.PID = @PID_Param)
) ", userDepts);

            cmd.Parameters.Add(new SqlParameter("UserId", userId));
            cmd.Parameters.Add(new SqlParameter("HasPidAccess", hasPidAccess ? 1 : 0));
            cmd.Parameters.Add(new SqlParameter("PID_Param", (object)model.PID ?? DBNull.Value));


            // 3. 經費歸零案過濾：【絕對不能加上 ISNULL】，利用 SQL 原生 Null 加法特性自動濾除
            sb.Append(" AND (t.YR_FUND + t.REV_FUND) <> 0 ");

            // =====================================================================
            // 以下為常規查詢條件
            // =====================================================================
            if (!string.IsNullOrEmpty(model.PID))
            {
                cmd.Parameters.Add(new SqlParameter("PID", model.PID)); // 恢復精確比對
                sb.Append("AND t.PID = @PID ");
            }

            if (!string.IsNullOrEmpty(model.TD_CODE))
            {
                cmd.Parameters.Add(new SqlParameter("TD_CODE", model.TD_CODE));
                sb.Append("AND t.TD_CODE = @TD_CODE ");
            }

            if (!string.IsNullOrEmpty(model.TD_NAME))
            {
                cmd.Parameters.Add(new SqlParameter("TD_NAME", "%" + model.TD_NAME + "%"));
                sb.Append("AND t.TD_NAME LIKE @TD_NAME ");
            }

            if (!string.IsNullOrEmpty(model.CITY))
            {
                // 舊版地名轉換邏輯
                if (model.CITY == "新北市") { sb.Append("AND t.CITY IN ('新北市', '台北縣') "); }
                else if (model.CITY == "台中市") { sb.Append("AND t.CITY IN ('台中市', '台中縣') "); }
                else if (model.CITY == "台南市") { sb.Append("AND t.CITY IN ('台南市', '台南縣') "); }
                else if (model.CITY == "高雄市") { sb.Append("AND t.CITY IN ('高雄市', '高雄縣') "); }
                else if (model.CITY == "桃園市") { sb.Append("AND t.CITY IN ('桃園市', '桃園縣') "); }
                else
                {
                    cmd.Parameters.Add(new SqlParameter("CITY", model.CITY));
                    sb.Append("AND t.CITY = @CITY ");
                }
            }

            if (!string.IsNullOrEmpty(model.TOWN))
            {
                string townPrefix = model.TOWN.Length >= 2 ? model.TOWN.Substring(0, 2) + "%" : model.TOWN;
                cmd.Parameters.Add(new SqlParameter("TOWN", townPrefix));
                sb.Append("AND t.TOWN LIKE @TOWN ");
            }

            if (!string.IsNullOrEmpty(model.CATY_CODE))
            {
                cmd.Parameters.Add(new SqlParameter("CATY_CODE", model.CATY_CODE));
                sb.Append("AND t.CATY_CODE = @CATY_CODE ");
            }

            if (!string.IsNullOrEmpty(model.IPJ_CODE))
            {
                cmd.Parameters.Add(new SqlParameter("IPJ_CODE", model.IPJ_CODE));
                sb.Append("AND t.IPJ_CODE = @IPJ_CODE ");
            }

            if (!string.IsNullOrEmpty(model.TD_STAT_DESC))
            {
                cmd.Parameters.Add(new SqlParameter("TD_STAT_DESC", model.TD_STAT_DESC + "%"));
                sb.Append("AND t.TD_STAT_DESC LIKE @TD_STAT_DESC ");
            }

            if (!string.IsNullOrEmpty(model.SYR))
            {
                cmd.Parameters.Add(new SqlParameter("SYR", model.SYR));
                sb.Append("AND t.PID_SYR = @SYR ");
            }

            // 執行查詢
            cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            // 4. C# 端資料整理：處理相同 TD_CODE 跨年度計畫的合併顯示
            ps.Columns.Add("PJ_NAME_ALL", typeof(string));
            List<DataRow> list = ps.AsEnumerable()
                .GroupBy(r => r["TD_CODE"]?.ToString())
                .Select(g =>
                {
                    // 取最新年度的那一筆作為主體顯示
                    var first = g.OrderByDescending(r => r["PID_SYR"]?.ToString()).First();

                    // 將所有關聯的計畫名稱組合起來 (例如: (114)都市綜合治水... \n (115)都市綜合治水...)
                    first["PJ_NAME_ALL"] = string.Join("\n", g.Where(r => !string.IsNullOrEmpty(r["PJ_NAME"]?.ToString()))
                                                .Select(r => "(" + r["PID_SYR"] + ")" + r["PJ_NAME"])
                                                .Distinct());
                    return first;
                }).ToList();

            // 分頁處理
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, model.page1, pageSize);
            model.List1 = plist;

            return Index(model);
        }

        public ActionResult Search(string PID, string TD_CODE, string TD_NAME, string CITY, string TOWN,
            // 3.1.4rev：page1 原為非可空 int，未帶此參數呼叫會直接 500
            // (The parameters dictionary contains a null entry for parameter 'page1')，改為預設第 1 頁
            string CATY_CODE, string IPJ_CODE, string TD_STAT_DESC, string SYR, string SYS_ID_FG, string SYS_PJ_FG, int page1 = 1) {

            TDE0040Model model = new TDE0040Model {
                PID = PID,
                TD_CODE = TD_CODE,
                TD_NAME = TD_NAME,
                CITY = CITY,
                TOWN = TOWN,
                CATY_CODE = CATY_CODE,
                IPJ_CODE = IPJ_CODE,
                TD_STAT_DESC = TD_STAT_DESC,
                SYR = SYR,
                SYS_ID_FG = SYS_ID_FG,
                SYS_PJ_FG = SYS_PJ_FG,
                page1 = page1
            };

            return Search(model);
        }

        public void GetSelectList()
        {
            //SqlCommand cmd = new SqlCommand("SELECT SYR, PID, PJ_NAME FROM PJMFORM ORDER BY SYR DESC");
            //cmd.Parameters.Add(new SqlParameter("PID", PID));
            //DataTable pids = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            //ViewBag.PIDS = tools.ToSelectList(pids, "PID", "PJ_NAME", null);
            ViewBag.PIDS = new SelectList(new List<SelectListItem>(), "Value", "Text");

            SqlCommand cmd = new SqlCommand("SELECT PID_SYR as SYR FROM TDV_TDDCFLS_Lite GROUP BY PID_SYR ORDER BY PID_SYR DESC");
            //cmd.Parameters.Add(new SqlParameter("PID", PID));
            DataTable syrs = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.SYRS = tools.ToSelectList(syrs, "SYR", "SYR", null);




            // 1. 縣市下拉依使用者部門對應 CMBASE.dbo.SECCITY 帶出 (3.1.4 需求，與 3.1.7 一致)
            //    查無資料 (中央/未設定帳號) 時退回全部縣市，避免下拉空白。
            string loginDept = WCals.Helpers.WebSiteHelper.getUserDept();
            DataTable dtDeptCities = util.GetCitiesByDept(loginDept);
            if (dtDeptCities == null || dtDeptCities.Rows.Count == 0)
            {
                dtDeptCities = util.GetCities();
            }
            ViewBag.cities = tools.ToSelectList(dtDeptCities, "CITY", "CITY", null);

            // 2. 處理縣市下拉選單
            //if (!string.IsNullOrEmpty(loginCity))
            //{
            //    // 如果是縣市政府帳號，下拉選單只放入該縣市，並鎖定 model
            //    List<SelectListItem> singleCity = new List<SelectListItem>();

            //    if (loginCity == "新北市")
            //    {
            //        singleCity.Add(new SelectListItem { Text = "新北市", Value = "新北市" });
            //        singleCity.Add(new SelectListItem { Text = "台北縣", Value = "台北縣" });
            //    }
            //    else if (loginCity == "台中市")
            //    {
            //        singleCity.Add(new SelectListItem { Text = "台中市", Value = "台中市" });
            //        singleCity.Add(new SelectListItem { Text = "台中縣", Value = "台中縣" });
            //    }
            //    else if (loginCity == "台南市")
            //    {
            //        singleCity.Add(new SelectListItem { Text = "台南市", Value = "台南市" });
            //        singleCity.Add(new SelectListItem { Text = "台南縣", Value = "台南縣" });
            //    }
            //    else if (loginCity == "高雄市")
            //    {
            //        singleCity.Add(new SelectListItem { Text = "高雄市", Value = "高雄市" });
            //        singleCity.Add(new SelectListItem { Text = "高雄縣", Value = "高雄縣" });
            //    }
            //    else if (loginCity == "桃園市")
            //    {
            //        singleCity.Add(new SelectListItem { Text = "桃園市", Value = "桃園市" });
            //        singleCity.Add(new SelectListItem { Text = "桃園縣", Value = "桃園縣" });
            //    }
            //    else
            //    {
            //        // 非五都之常規縣市，僅呈現登入縣市
            //        singleCity.Add(new SelectListItem { Text = loginCity, Value = loginCity });
            //    }

            //    //singleCity.Add(new SelectListItem { Text = loginCity, Value = loginCity });
            //    ViewBag.cities = new SelectList(singleCity, "Value", "Text", loginCity);

            //}
            //else
            //{
            //    // 中央帳號，顯示所有縣市
            //    ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);
            //}

            DataTable dt;
            //ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
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
            try
            {
                var sysList = db.PJCSYST.Select(s => new { s.SYS_ID, s.SYS_NAME, s.STAT_NO }).Where(s => s.STAT_NO == "S").ToList();
                ViewBag.SYS_ID = new SelectList(sysList, "SYS_ID", "SYS_NAME");
            }
            catch
            {
                ViewBag.SYS_ID = new SelectList(new List<SelectListItem>(), "Value", "Text");
            }

            // 實施計畫下拉單
            try
            {
                var pjList = db.PJCFORM.Select(p => new { p.PJ_ID, p.PJ_NAME }).ToList();
                ViewBag.PJ_ID = new SelectList(pjList, "PJ_ID", "PJ_NAME");
            }
            catch
            {
                ViewBag.PJ_ID = new SelectList(new List<SelectListItem>(), "Value", "Text");
            }
        }


        private readonly CMETMSEntities db = new CMETMSEntities();

        // GET: PAGE1
        public ActionResult PAGE1(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            TDE0040Model model = new TDE0040Model {
                TD_CODE = tdmform.TD_CODE,
                TD_NAME = tdmform.TD_NAME,
                CITY = tdmform.TD_CITY,
                TOWN = tdmform.TD_TOWN,
                CATY_CODE = tdmform.CATY_CODE,
                IPJ_CODE = tdmform.IPJ_CODE,
                PID = tdmform.PPID?.ToString(),
                SYR = tdmform.WRT_YR?.ToString(),
                TDMFORM = tdmform
            };

            GetSelectList();
            return View("PAGE1", model);
        }

        // POST: PAGE1
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE1(TDE0040Model model) {
            if(string.IsNullOrEmpty(model.TDMFORM.TD_CODE)) {
                TempData["Error"] = "TD_CODE 不可為空";
                GetSelectList();
                return View("PAGE1", model);
            }

            try {
                // Debug: 輸出接收到的 TD_FUND 值
                System.Diagnostics.Debug.WriteLine($"PAGE1 POST - TD_FUND 接收值: {model.TDMFORM?.TD_FUND}");

                var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == model.TDMFORM.TD_CODE);

                if(tdmform == null) {
                    TempData["Error"] = "找不到指定的工程資料";
                    GetSelectList();
                    return View("PAGE1", model);
                }

                // Debug: 輸出更新前的 TD_FUND 值
                System.Diagnostics.Debug.WriteLine($"PAGE1 POST - TD_FUND 更新前: {tdmform.TD_FUND}");

                if(model.TDMFORM != null) {
                    // 更新欄位
                    tdmform.TD_NAME = model.TDMFORM.TD_NAME;
                    tdmform.TD_CATY = model.TDMFORM.TD_CATY;
                    tdmform.TD_FUND = model.TDMFORM.TD_FUND;
                    tdmform.PRE_PRD = model.TDMFORM.PRE_PRD;
                    tdmform.PRD_TP = model.TDMFORM.PRD_TP;
                    tdmform.VRF_DT = model.TDMFORM.VRF_DT;
                    tdmform.VRF_FN = model.TDMFORM.VRF_FN;
                    tdmform.EXE_MCH = model.TDMFORM.EXE_MCH;
                    tdmform.ADM_CONT = model.TDMFORM.ADM_CONT;
                    tdmform.ADM_TEL = model.TDMFORM.ADM_TEL;
                    tdmform.ADM_EMAIL = model.TDMFORM.ADM_EMAIL;
                    tdmform.PCCTD_CODE = model.TDMFORM.PCCTD_CODE;
                    tdmform.DSG_BUSS_NAME = model.TDMFORM.DSG_BUSS_NAME;
                    tdmform.DSG_CONT = model.TDMFORM.DSG_CONT;
                    tdmform.DSG_CONTEL = model.TDMFORM.DSG_CONTEL;
                    tdmform.SVY_BUSS_NAME = model.TDMFORM.SVY_BUSS_NAME;
                    tdmform.SVY_CONT = model.TDMFORM.SVY_CONT;
                    tdmform.SVY_CONTEL = model.TDMFORM.SVY_CONTEL;
                    tdmform.ADV_CHIEF = model.TDMFORM.ADV_CHIEF;
                    tdmform.TD_NOTE = model.TDMFORM.TD_NOTE;
                    tdmform.TDM_TD_CODE = model.TDMFORM.TDM_TD_CODE;
                    tdmform.SAVE_FEE = model.TDMFORM.SAVE_FEE;
                    tdmform.SYS_ID = model.TDMFORM.SYS_ID;
                    tdmform.PJ_ID = model.TDMFORM.PJ_ID;
                    tdmform.TD_FINISH = model.TDMFORM.TD_FINISH;
                    tdmform.START_FUND = model.TDMFORM.START_FUND;
                    tdmform.HOUSEHOLD = model.TDMFORM.HOUSEHOLD;
                    tdmform.TOTAL_HOUSEHOLD = model.TDMFORM.TOTAL_HOUSEHOLD;
                    tdmform.WTR_PL_NEW_LEN = model.TDMFORM.WTR_PL_NEW_LEN;
                    tdmform.WTR_PL_REP_LEN = model.TDMFORM.WTR_PL_REP_LEN;
                    tdmform.WTR_FLOOD = model.TDMFORM.WTR_FLOOD;
                    tdmform.WTR_TOTAL_PUMP = model.TDMFORM.WTR_TOTAL_PUMP;
                    tdmform.WTR_GU_STA = model.TDMFORM.WTR_GU_STA;
                    tdmform.CRT_YR = model.TDMFORM.CRT_YR;
                    tdmform.TD_LINE = model.TDMFORM.TD_LINE;
                    tdmform.GIS_X = model.TDMFORM.GIS_X;
                    tdmform.GIS_Y = model.TDMFORM.GIS_Y;
                    tdmform.UPD_DT = DateTime.Now;
                    tdmform.UPD_UID = WebSiteHelper.getUserId();

                }

                db.SaveChanges();

                // Debug: 輸出儲存後的 TD_FUND 值
                System.Diagnostics.Debug.WriteLine($"PAGE1 POST - TD_FUND 儲存後: {tdmform.TD_FUND}");

                // 更新 model 使其顯示最新的資料
                model.TDMFORM = tdmform;

                TempData["Success"] = "資料更新成功";
                GetSelectList();
                return View("PAGE1", model);
            } catch(Exception ex) {
                TempData["Error"] = "資料更新失敗：" + ex.Message;
                GetSelectList();
                return View("PAGE1", model);
            }
        }

        

        
        

        // GET: PAGE4
        public ActionResult PAGE4(string TD_CODE) {
            if(string.IsNullOrEmpty(TD_CODE)) {
                return RedirectToAction("Index");
            }

            var tdmform = db.TDMFORM.FirstOrDefault(t => t.TD_CODE == TD_CODE);

            if(tdmform == null) {
                TempData["Error"] = "找不到指定的工程資料";
                return RedirectToAction("Index");
            }

            // 查詢該工程的 TDDDCOB 資料
            SqlCommand cmd = new SqlCommand("SELECT * FROM TDDDCOB WHERE TD_CODE = @TD_CODE ORDER BY IT_CODE");
            cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
            DataTable cobData = BASE_DAO.SqlDao.CMEDOCQuery4List(cmd);

            TDE0040Model model = new TDE0040Model {
                TD_CODE = TD_CODE,
                TD_NAME = tdmform.TD_NAME,
                TDMFORM = tdmform,
                List1 = new PagedList<DataRow>(cobData.AsEnumerable().ToList(), 1, 100)
            };

            GetSelectList();
            return View("PAGE4", model);
        }

        // POST: PAGE4 - 新增或修改 TDDDCOB
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult PAGE4(TDE0040Model model, string IT_CODE, string OBJ_TITLE, string OBJ_DESC) {

            if(string.IsNullOrEmpty(model.TD_CODE)) {
                TempData["Error"] = "TD_CODE 不可為空";
                return RedirectToAction("PAGE4", new { TD_CODE = model.TD_CODE });
            }

            try {
                DateTime now = DateTime.Now;
                string userId = WebSiteHelper.getUserId();

                // 處理文件上傳
                byte[] objBody = null;
                string objName = null;
                string objExtName = null;
                long? objSize = null;

                var files = Request.Files;
                if(files.Count > 0 && files[0] != null && files[0].ContentLength > 0) {
                    HttpPostedFileBase file = files[0];

                    // 讀取文件為二進制
                    using(var binaryReader = new System.IO.BinaryReader(file.InputStream)) {
                        objBody = binaryReader.ReadBytes(file.ContentLength);
                    }

                    objName = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                    objExtName = System.IO.Path.GetExtension(file.FileName);
                    objSize = file.ContentLength;
                }

                // 檢查是否已存在（編輯模式）
                SqlCommand cmdCheck = new SqlCommand("SELECT COUNT(*) as CNT FROM TDDDCOB WHERE TD_CODE = @TD_CODE AND IT_CODE = @IT_CODE");
                cmdCheck.Parameters.Add(new SqlParameter("TD_CODE", model.TD_CODE));
                cmdCheck.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));
                DataTable dtCheck = BASE_DAO.SqlDao.CMEDOCQuery4List(cmdCheck);
                bool isEdit = dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0]["CNT"]) > 0;

                SqlCommand cmd;
                if(isEdit) {
                    // 更新模式
                    StringBuilder sb = new StringBuilder("UPDATE TDDDCOB SET ");
                    sb.Append("OBJ_TITLE = @OBJ_TITLE, ");
                    sb.Append("OBJ_DESC = @OBJ_DESC, ");

                    if(objBody != null) {
                        sb.Append("OBJ_NAME = @OBJ_NAME, ");
                        sb.Append("OBJ_EXTNAME = @OBJ_EXTNAME, ");
                        sb.Append("OBJ_SIZE = @OBJ_SIZE, ");
                        sb.Append("OBJ_BODY = @OBJ_BODY, ");
                        sb.Append("OBJ_MOD_DM = @OBJ_MOD_DM, ");
                    }

                    sb.Append("UPD_DT = @UPD_DT, ");
                    sb.Append("UPD_UID = @UPD_UID ");
                    sb.Append("WHERE TD_CODE = @TD_CODE AND IT_CODE = @IT_CODE");

                    cmd = new SqlCommand(sb.ToString());
                } else {
                    // 新增模式
                    cmd = new SqlCommand(
                        "INSERT INTO TDDDCOB (TD_CODE, IT_CODE, OBJ_NAME, OBJ_EXTNAME, OBJ_TITLE, OBJ_DESC, " +
                        "OBJ_SIZE, OBJ_CRT_DM, OBJ_MOD_DM, OBJ_BODY, AUTHOR, ADD_DT, UPD_DT, UPD_UID) " +
                        "VALUES (@TD_CODE, @IT_CODE, @OBJ_NAME, @OBJ_EXTNAME, @OBJ_TITLE, @OBJ_DESC, " +
                        "@OBJ_SIZE, @OBJ_CRT_DM, @OBJ_MOD_DM, @OBJ_BODY, @AUTHOR, @ADD_DT, @UPD_DT, @UPD_UID)"
                    );
                    cmd.Parameters.Add(new SqlParameter("OBJ_CRT_DM", now));
                    cmd.Parameters.Add(new SqlParameter("AUTHOR", userId));
                    cmd.Parameters.Add(new SqlParameter("ADD_DT", now));
                }

                // 共同參數
                cmd.Parameters.Add(new SqlParameter("TD_CODE", model.TD_CODE));
                cmd.Parameters.Add(new SqlParameter("IT_CODE", IT_CODE));
                cmd.Parameters.Add(new SqlParameter("OBJ_TITLE", (object)OBJ_TITLE ?? DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("OBJ_DESC", (object)OBJ_DESC ?? DBNull.Value));

                if(objBody != null) {
                    cmd.Parameters.Add(new SqlParameter("OBJ_NAME", (object)objName ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("OBJ_EXTNAME", (object)objExtName ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("OBJ_SIZE", (object)objSize ?? DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("OBJ_BODY", objBody));
                    cmd.Parameters.Add(new SqlParameter("OBJ_MOD_DM", now));
                }

                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId));

                BASE_DAO.SqlDao.CMEDOCExecuteSql(cmd);

                TempData["Success"] = isEdit ? "資料更新成功" : "資料新增成功";
                return RedirectToAction("PAGE4", new { TD_CODE = model.TD_CODE });
            } catch(Exception ex) {
                TempData["Error"] = "操作失敗：" + ex.Message;
                return RedirectToAction("PAGE4", new { TD_CODE = model.TD_CODE });
            }
        }




        // 下載 TDDDCOB 檔案
        [HttpGet]
        public ActionResult DownloadTDDDCOBFile(string tdCode, string itCode) {
            try {
                SqlCommand cmd = new SqlCommand("SELECT OBJ_NAME, OBJ_EXTNAME, OBJ_BODY FROM TDDDCOB WHERE TD_CODE = @TD_CODE AND IT_CODE = @IT_CODE");
                cmd.Parameters.Add(new SqlParameter("TD_CODE", tdCode));
                cmd.Parameters.Add(new SqlParameter("IT_CODE", itCode));
                DataTable dt = BASE_DAO.SqlDao.CMEDOCQuery4List(cmd);

                if(dt.Rows.Count > 0) {
                    DataRow row = dt.Rows[0];
                    byte[] objBody = row["OBJ_BODY"] as byte[];
                    string objName = row["OBJ_NAME"].ToString();
                    string objExtName = row["OBJ_EXTNAME"].ToString();

                    if(objBody != null && objBody.Length > 0) {
                        string fileName = objName + objExtName;
                        var contentType = MimeMapping.GetMimeMapping(fileName);
                        return File(objBody, contentType, fileName);
                    }
                }

                TempData["Error"] = "檔案不存在";
                return RedirectToAction("Index");
            } catch(Exception ex) {
                TempData["Error"] = "下載失敗：" + ex.Message;
                return RedirectToAction("Index");
            }
        }


        // 取得單筆 TDDDCOB 資料 (用於編輯)
        [HttpPost]
        public ActionResult GetTDDDCOB(string tdCode, string itCode) {
            try {
                SqlCommand cmd = new SqlCommand("SELECT TD_CODE, IT_CODE, OBJ_NAME, OBJ_EXTNAME, OBJ_TITLE, OBJ_DESC, OBJ_SIZE FROM TDDDCOB WHERE TD_CODE = @TD_CODE AND IT_CODE = @IT_CODE");
                cmd.Parameters.Add(new SqlParameter("TD_CODE", tdCode));
                cmd.Parameters.Add(new SqlParameter("IT_CODE", itCode));
                DataTable dt = BASE_DAO.SqlDao.CMEDOCQuery4List(cmd);

                if(dt.Rows.Count > 0) {
                    DataRow row = dt.Rows[0];
                    return Json(new {
                        success = true,
                        data = new {
                            TD_CODE = row["TD_CODE"].ToString(),
                            IT_CODE = row["IT_CODE"].ToString(),
                            OBJ_TITLE = row["OBJ_TITLE"].ToString(),
                            OBJ_DESC = row["OBJ_DESC"].ToString(),
                            OBJ_NAME = row["OBJ_NAME"].ToString(),
                            OBJ_EXTNAME = row["OBJ_EXTNAME"].ToString()
                        }
                    });
                }

                return Json(new { success = false, message = "找不到資料" });
            } catch(Exception ex) {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // 刪除 TDDDCOB 資料
        [HttpPost]
        public ActionResult DeleteTDDDCOB(string tdCode, string itCode) {
            try {
                // 刪除資料庫記錄（文件內容存在 OBJ_BODY 中，會一起刪除）
                SqlCommand cmdDelete = new SqlCommand("DELETE FROM TDDDCOB WHERE TD_CODE = @TD_CODE AND IT_CODE = @IT_CODE");
                cmdDelete.Parameters.Add(new SqlParameter("TD_CODE", tdCode));
                cmdDelete.Parameters.Add(new SqlParameter("IT_CODE", itCode));
                BASE_DAO.SqlDao.CMEDOCExecuteSql(cmdDelete);

                return Json(new { success = true, message = "刪除成功" });
            } catch(Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
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

        // 寫一個小函式來取得包含子部門的字串
        private string GetAllChildDeptsStr(string mainDeptNo)
        {
            if (string.IsNullOrEmpty(mainDeptNo)) return "''";

            // 採用「跨體系根節點」自動展開邏輯
            string sql = @"
        WITH DeptTree AS (
            -- 1. 錨點：主部門 + 所有體系根節點 (對齊您提供的 A, B, F, 0 系列)
            SELECT DEPT_NO, AN_DEPT_NO, DATA_TP
            FROM CMBASE.dbo.OGCDEPT
            WHERE DEPT_NO = @MainDept
               OR (AN_DEPT_NO IN ('', '000000') AND LEFT(DEPT_NO, 1) IN ('0', '1', '2', 'A', 'B', 'F'))
            
            UNION ALL
            
            -- 2. 遞迴：向下展開
            SELECT d.DEPT_NO, d.AN_DEPT_NO, d.DATA_TP
            FROM CMBASE.dbo.OGCDEPT d
            INNER JOIN DeptTree t ON d.AN_DEPT_NO = t.DEPT_NO
            WHERE d.DEPT_NO <> d.AN_DEPT_NO -- 防死迴圈
        )
        SELECT DISTINCT DEPT_NO FROM DeptTree WHERE DATA_TP IN ('D', 'T')";

            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("@MainDept", mainDeptNo));
            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            if (dt == null || dt.Rows.Count == 0) return "'" + mainDeptNo + "'";

            var list = dt.AsEnumerable().Select(r => "'" + r["DEPT_NO"].ToString() + "'").ToList();
            return string.Join(",", list);
        }

        [HttpPost]
        public ActionResult GetPidsBySyr(string syr)
        {

            SqlCommand cmd = new SqlCommand("SELECT SYR, PID, PJ_NAME FROM PJMFORM WHERE SYR = @SYR ORDER BY SYR DESC");

            cmd.Parameters.Add(new SqlParameter("@SYR", syr));
            //cmd.Parameters.Add(new SqlParameter("PID", PID));
            DataTable pids = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            //ViewBag.PIDS = tools.ToSelectList(pids, "PID", "PJ_NAME", null);

            // 將 DataTable 轉為 JSON 回傳給前端
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(pids);
            return Json(json);
        }

    }
}