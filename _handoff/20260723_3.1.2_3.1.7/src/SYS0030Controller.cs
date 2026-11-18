using PagedList;
using Syncfusion.DocIO.DLS;
using Syncfusion.XlsIO;
using System.IO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using System.Web.UI;
using WCals.classes.john;
using WCals.Helpers;
using WCals.ViewModels.NewModels;
using static System.Runtime.CompilerServices.RuntimeHelpers;

namespace WCals.Controllers.NSPMIS {
    public class SYS0030Controller : Controller {



        // GET: SYS0030
        public ActionResult Index(SYS0030Model model) {

            ViewBag.cities = tools.ToSelectList(GetDeptCities(), "CITY", "CITY", null);
            ViewBag.towns = tools.ToSelectList(util.GetTowns(null), "TOWN", "TOWN", null);
            // 加入污水下水道系統下拉選單資料
            SqlCommand cmdSystems = new SqlCommand("SELECT SYS_ID, SYS_NAME FROM PJCSYST WHERE STAT_NO='S' ORDER BY SYS_NAME");
            DataTable dtSystems = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSystems);
            ViewBag.systems = tools.ToSelectList(dtSystems, "SYS_ID", "SYS_NAME", null);
            return View("SYS0030Index", model);
        }

        int pageSize = 5;
        int sysManagerPageSize = 20;   // 系統別管理每頁筆數 (3.1.7 需求)

        /// <summary>
        /// 依登入者部門取得授權縣市 (3.1.7 需求：對應 CMBASE.dbo.SECCITY)。
        /// 查無資料 (中央/未設定帳號) 時退回全部縣市，避免下拉空白。
        /// </summary>
        private DataTable GetDeptCities() {
            string dept = WCals.Helpers.WebSiteHelper.getUserDept();
            DataTable dt = util.GetCitiesByDept(dept);
            if(dt == null || dt.Rows.Count == 0) {
                dt = util.GetCities();
            }
            return dt;
        }

        /// <summary>
        /// 依縣市取得污水下水道系統別 (3.1.7 需求：系統別與縣市別連動)
        /// </summary>
        [HttpGet]
        public JsonResult GetSystemsByCity(string city) {
            StringBuilder sb = new StringBuilder("SELECT SYS_ID, SYS_NAME FROM PJCSYST WHERE STAT_NO='S' ");
            SqlCommand cmd = new SqlCommand();
            if(!string.IsNullOrEmpty(city)) {
                cmd.Parameters.Add(new SqlParameter("CITY", city));
                sb.Append("AND CITY=@CITY ");
            }
            sb.Append("ORDER BY SYS_NAME");
            cmd.CommandText = sb.ToString();
            DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
            var list = dt.AsEnumerable().Select(r => new {
                SYS_ID = Convert.ToString(r["SYS_ID"]),
                SYS_NAME = Convert.ToString(r["SYS_NAME"])
            }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Search(string sys, string city, string doc_name, string doc_type, int page) {
            SYS0030Model model = new SYS0030Model {
                sys = sys,
                CITY = city,
                doc_name = doc_name,
                doc_type = doc_type,
                page = page
            };

            return Search(model);
        }

        // POST: SYS0030首頁查詢
        [HttpPost]
        public ActionResult Search(SYS0030Model model) {

            if(model.page == 0) {
                model.page = 1;
            }

            StringBuilder sb = new StringBuilder("SELECT s.SYS_NAME, p.* FROM PJCFORM p, PJCSYST s WHERE p.SYS_ID in (select SYS_ID from PJCSYST where STAT_NO='S') AND p.SYS_ID=s.SYS_ID ");

            SqlCommand cmd = new SqlCommand();
            if(!string.IsNullOrEmpty(model.sys)) {
                // 修正：model.sys 是 SYS_ID，應該用 SYS_ID 來比對
                cmd.Parameters.Add(new SqlParameter("SYS_ID", model.sys));
                sb.Append("AND s.SYS_ID = @SYS_ID ");
            }

            if(!string.IsNullOrEmpty(model.CITY)) {
                cmd.Parameters.Add(new SqlParameter("SER_CITY", model.CITY));
                sb.Append("AND p.SER_CITY = @SER_CITY ");
            }

            if(!string.IsNullOrEmpty(model.doc_type)) {
                cmd.Parameters.Add(new SqlParameter("PJ_ID_TYPE", model.doc_type));
                sb.Append("AND p.PJ_ID_TYPE = @PJ_ID_TYPE ");
            }

            // 新增：文件名稱搜尋條件
            if(!string.IsNullOrEmpty(model.doc_name)) {
                cmd.Parameters.Add(new SqlParameter("PJ_NAME", "%" + model.doc_name + "%"));
                sb.Append("AND p.PJ_NAME LIKE @PJ_NAME ");
            }

            cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, model.page, pageSize);
            model.List1 = plist;
            return Index(model);
        }

        /// <summary>
        /// 系統別管理
        /// </summary>
        /// <returns></returns>
        public ActionResult SYSManager(SYS0030Model model) {

            ViewBag.cities = GetDeptCities();
            if(model.page == 0) {
                model.page = 1;
            }

            return View("SYSManager", model);
        }

        

        public ActionResult SYSManagerSearch(string CITY, string SYS_ID, string SYS_NAME, string CONST_TP, string IS_SUP, int page = 1) {

            SYS0030Model model = new SYS0030Model {
                CITY = CITY,
                SYS_ID = SYS_ID,
                SYS_NAME = SYS_NAME,
                CONST_TP = CONST_TP,
                IS_SUP = IS_SUP,
                page = page
            };

            return SYSManagerSearch(model);
        }

        /// <summary>
        /// 系統別管理
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SYSManagerSearchA")]
        public ActionResult SYSManagerSearchA(SYS0030Model model) {

            ViewBag.cities = GetDeptCities();
            if(model.page == 0) {
                model.page = 1;
            }

            StringBuilder sb = new StringBuilder("SELECT*FROM PJCSYST WHERE STAT_NO='S' ");

            SqlCommand cmd = new SqlCommand();
            if(!string.IsNullOrEmpty(model.CITY)) {
                cmd.Parameters.Add(new SqlParameter("CITY", model.CITY));
                sb.Append("AND CITY=@CITY ");
            }
            if(!string.IsNullOrEmpty(model.SYS_ID)) {
                cmd.Parameters.Add(new SqlParameter("SYS_ID", model.SYS_ID));
                sb.Append("AND SYS_ID=@SYS_ID ");
            }
            if(!string.IsNullOrEmpty(model.SYS_NAME)) {
                cmd.Parameters.Add(new SqlParameter("SYS_NAME", "%" + model.SYS_NAME + "%"));
                sb.Append("AND SYS_NAME LIKE @SYS_NAME ");
            }
            if(!string.IsNullOrEmpty(model.CONST_TP)) {
                cmd.Parameters.Add(new SqlParameter("CONST_TP", "%" + model.CONST_TP + "%"));
                sb.Append("AND CONST_TP LIKE @CONST_TP ");
            }
            if(!string.IsNullOrEmpty(model.IS_SUP)) {
                cmd.Parameters.Add(new SqlParameter("IS_SUP", model.IS_SUP));
                sb.Append("AND IS_SUP = @IS_SUP ");
            }
            cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, model.page, sysManagerPageSize);
            model.List1 = plist;

            ModelState.Clear(); // ★ 關鍵：讓 DropDownList 不吃舊值
            return SYSManager(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SYSManagerSearch")]
        public ActionResult SYSManagerSearch(SYS0030Model model) {

            ViewBag.cities = GetDeptCities();
            if(model.page == 0) {
                model.page = 1;
            }

            StringBuilder sb = new StringBuilder("SELECT*FROM PJCSYST WHERE STAT_NO='S' ");

            SqlCommand cmd = new SqlCommand();
            if(!string.IsNullOrEmpty(model.CITY)) {
                cmd.Parameters.Add(new SqlParameter("CITY", model.CITY));
                sb.Append("AND CITY=@CITY ");
            }
            if(!string.IsNullOrEmpty(model.SYS_ID)) {
                cmd.Parameters.Add(new SqlParameter("SYS_ID", model.SYS_ID));
                sb.Append("AND SYS_ID=@SYS_ID ");
            }
            if(!string.IsNullOrEmpty(model.SYS_NAME)) {
                cmd.Parameters.Add(new SqlParameter("SYS_NAME", "%" + model.SYS_NAME + "%"));
                sb.Append("AND SYS_NAME LIKE @SYS_NAME ");
            }
            if(!string.IsNullOrEmpty(model.CONST_TP)) {
                cmd.Parameters.Add(new SqlParameter("CONST_TP", "%" + model.CONST_TP + "%"));
                sb.Append("AND CONST_TP LIKE @CONST_TP ");
            }
            if(!string.IsNullOrEmpty(model.IS_SUP)) {
                cmd.Parameters.Add(new SqlParameter("IS_SUP", model.IS_SUP));
                sb.Append("AND IS_SUP = @IS_SUP ");
            }
            cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, model.page, sysManagerPageSize);
            model.List1 = plist;

            ModelState.Clear(); // ★ 關鍵：讓 DropDownList 不吃舊值
            return SYSManager(model);
        }

        /// <summary>
        /// 系統別管理
        /// </summary>
        /// <returns></returns>
        public ActionResult SYSManagerSave(SYS0030Model model) {

            ViewBag.cities = GetDeptCities();
            if(model.page == 0) {
                model.page = 1;
            }

            StringBuilder sb = new StringBuilder("SELECT*FROM PJCSYST WHERE STAT_NO='S' ");

            SqlCommand cmd = new SqlCommand();

            cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, model.page, sysManagerPageSize);
            model.List1 = plist;

            return SYSManager(model);
        }

        [HttpPost]
        public ActionResult SYSManagerAdd(SYS0030Model model) {
            try {
                // 驗證必填欄位
                if(string.IsNullOrEmpty(model.CITY)) {
                    ViewBag.Message = "請選擇縣市別";
                    return SYSManager(model);
                }
                if(string.IsNullOrEmpty(model.SYS_ID)) {
                    ViewBag.Message = "請輸入下水道系統編號";
                    return SYSManager(model);
                }
                if(string.IsNullOrEmpty(model.SYS_NAME)) {
                    ViewBag.Message = "請輸入下水道系統名稱";
                    return SYSManager(model);
                }

                // 檢查 SYS_ID 是否已存在
                SqlCommand checkCmd = new SqlCommand("SELECT COUNT(*) FROM PJCSYST WHERE SYS_ID = @SYS_ID");
                checkCmd.Parameters.Add(new SqlParameter("SYS_ID", model.SYS_ID));
                DataTable checkDt = BASE_DAO.SqlDao.CMETMSQuery4List(checkCmd);
                if(checkDt.Rows.Count > 0 && Convert.ToInt32(checkDt.Rows[0][0]) > 0) {
                    ViewBag.Message = "此系統編號已存在，請使用其他編號";
                    return SYSManager(model);
                }

                // 新增資料到 PJCSYST
                DateTime now = DateTime.Now;
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO PJCSYST (SYS_ID, CITY, SYS_NAME, CONST_TP, IS_SUP, ADD_DT, UPD_DT, UPD_UID, STAT_NO) " +
                    "VALUES (@SYS_ID, @CITY, @SYS_NAME, @CONST_TP, @IS_SUP, @ADD_DT, @UPD_DT, @UPD_UID, @STAT_NO)"
                );
                cmd.Parameters.Add(new SqlParameter("SYS_ID", model.SYS_ID));
                cmd.Parameters.Add(new SqlParameter("CITY", model.CITY));
                cmd.Parameters.Add(new SqlParameter("SYS_NAME", model.SYS_NAME ?? ""));
                cmd.Parameters.Add(new SqlParameter("CONST_TP", model.CONST_TP ?? ""));
                cmd.Parameters.Add(new SqlParameter("IS_SUP", model.IS_SUP ?? "N"));
                cmd.Parameters.Add(new SqlParameter("ADD_DT", now));
                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));
                cmd.Parameters.Add(new SqlParameter("STAT_NO", "S"));

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                ViewBag.Message = "新增成功";

                // 清空表單並重新查詢（顯示所有資料）
                model.CITY = "";
                model.SYS_ID = "";
                model.SYS_NAME = "";
                model.CONST_TP = "";
                model.IS_SUP = "";

                return SYSManagerSearch(model);
            } catch(Exception ex) {
                ViewBag.Message = "新增失敗：" + ex.Message;
                return SYSManager(model);
            }
        }

        public ActionResult SYSInfo(string pjId, string mode = null) {
            try {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();
                WCals.Models.CMETMSModels.PJCFORM pjcform = null;

                if(!string.IsNullOrEmpty(pjId)) {
                    // 編輯模式或複製模式：載入既有記錄
                    System.Diagnostics.Debug.WriteLine("SYSInfo - 查詢 pjId: " + pjId);
                    pjcform = db.PJCFORM.FirstOrDefault(x => x.PJ_ID == pjId);

                    if(pjcform != null) {
                        System.Diagnostics.Debug.WriteLine("SYSInfo - 找到記錄: " + pjcform.PJ_NAME);

                        // 複製模式：清空 PJ_ID，讓儲存時作為新增處理
                        if(mode == "copy") {
                            pjcform.PJ_ID = null;
                        }
                    } else {
                        System.Diagnostics.Debug.WriteLine("SYSInfo - 未找到記錄");
                    }
                }

                if(pjcform == null) {
                    // 新增模式：建立空物件
                    System.Diagnostics.Debug.WriteLine("SYSInfo - 創建新物件");
                    pjcform = new WCals.Models.CMETMSModels.PJCFORM();
                }

                // 傳遞 mode 給 View
                ViewBag.Mode = mode;

                // 加入縣市下拉選單資料
                ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);

                // 加入污水下水道系統下拉選單資料 (使用 SQL 查詢)
                SqlCommand cmdSystems = new SqlCommand("SELECT SYS_ID, SYS_NAME FROM PJCSYST WHERE STAT_NO='S' ORDER BY SYS_NAME");
                DataTable dtSystems = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSystems);
                ViewBag.systems = tools.ToSelectList(dtSystems, "SYS_ID", "SYS_NAME", null);

                return View("SYSInfo", pjcform);
            } catch(Exception ex) {
                System.Diagnostics.Debug.WriteLine("SYSInfo - 異常: " + ex.Message);
                // 即使發生異常也要返回新物件，避免 null reference
                ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);
                ViewBag.systems = new SelectList(new List<object>(), "SYS_ID", "SYS_NAME");
                return View("SYSInfo", new WCals.Models.CMETMSModels.PJCFORM());
            }
        }

        public ActionResult SYSStat() {
            return View("SYSStat");
        }

        /// <summary>
        /// 新增/修改規劃報告
        /// </summary>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SaveSYSInfo")]
        public ActionResult SaveSYSInfo(WCals.Models.CMETMSModels.PJCFORM model) {
            try {
                // 3.1.7 需求：縣市、污水下水道系統、文件類別為必填（後端防呆，取代 EF 原始驗證訊息）
                if(string.IsNullOrEmpty(model.SER_CITY) || string.IsNullOrEmpty(model.SYS_ID) || string.IsNullOrEmpty(model.PJ_ID_TYPE)) {
                    ViewBag.Message = "請填寫必填欄位：縣市、污水下水道系統、文件類別";
                    return SYSInfo(model.PJ_ID, ViewBag.Mode as string);
                }

                var db = new WCals.Models.CMETMSModels.CMETMSEntities();
                string pjId = model.PJ_ID;

                if(!string.IsNullOrEmpty(model.PJ_ID)) {
                    // 編輯模式：更新既有記錄
                    var existingPjcform = db.PJCFORM.FirstOrDefault(x => x.PJ_ID == model.PJ_ID);
                    if(existingPjcform != null) {
                        existingPjcform.SYS_ID = model.SYS_ID;
                        existingPjcform.PJ_ID_TYPE = model.PJ_ID_TYPE;
                        existingPjcform.SER_FN_1 = model.SER_FN_1;
                        existingPjcform.SER_FN_2 = model.SER_FN_2;
                        existingPjcform.SER_FN_3 = model.SER_FN_3;
                        existingPjcform.SER_FN_4 = model.SER_FN_4;
                        existingPjcform.SER_FN_5 = model.SER_FN_5;
                        existingPjcform.SER_FN_6 = model.SER_FN_6;
                        existingPjcform.SER_FN_7 = model.SER_FN_7;
                        existingPjcform.SER_CITY = model.SER_CITY;
                        existingPjcform.PJ_NAME = model.PJ_NAME;
                        existingPjcform.PJ_AREA_NAME = model.PJ_AREA_NAME;
                        existingPjcform.CHK_DT = model.CHK_DT;
                        existingPjcform.CHK_FN = model.CHK_FN;
                        existingPjcform.PJ_YEAR = model.PJ_YEAR;
                        existingPjcform.ADM_NAME = model.ADM_NAME;
                        existingPjcform.PJ_SEQ = model.PJ_SEQ;
                        existingPjcform.PJ_FIRST = model.PJ_FIRST;
                        existingPjcform.PJ_AREA = model.PJ_AREA;
                        existingPjcform.PJ_PEO = model.PJ_PEO;
                        existingPjcform.PJ_WATER = model.PJ_WATER;
                        existingPjcform.PJ_IN_BOD = model.PJ_IN_BOD;
                        existingPjcform.PJ_IN_SS = model.PJ_IN_SS;
                        existingPjcform.PJ_OUT_BOD = model.PJ_OUT_BOD;
                        existingPjcform.PJ_OUT_SS = model.PJ_OUT_SS;
                        existingPjcform.WATER_TYPE = model.WATER_TYPE;
                        existingPjcform.WATER_AREA = model.WATER_AREA;
                        existingPjcform.MUD_TYPE = model.MUD_TYPE;
                        existingPjcform.PJ_FUND = model.PJ_FUND;
                        existingPjcform.MAIN_PIPO = model.MAIN_PIPO;
                        existingPjcform.BRANCH_PIPO = model.BRANCH_PIPO;
                        existingPjcform.PIPO_BOT_FUND = model.PIPO_BOT_FUND;
                        existingPjcform.PIPO_SELF_FUND = model.PIPO_SELF_FUND;
                        existingPjcform.USER_BOT = model.USER_BOT;
                        existingPjcform.USER_SELF = model.USER_SELF;
                        existingPjcform.USER_BOT_FUND = model.USER_BOT_FUND;
                        existingPjcform.USER_SELF_FUND = model.USER_SELF_FUND;
                        existingPjcform.SEWAGE_NUM = model.SEWAGE_NUM;
                        existingPjcform.SEWAGE_AREA = model.SEWAGE_AREA;
                        existingPjcform.SEWAGE_CMD = model.SEWAGE_CMD;
                        existingPjcform.SEWAGE_LAND_FUND = model.SEWAGE_LAND_FUND;
                        existingPjcform.SEWAGE_BUILD_FUND = model.SEWAGE_BUILD_FUND;
                        existingPjcform.PUMPING_NUM = model.PUMPING_NUM;
                        existingPjcform.PUMPING_AREA = model.PUMPING_AREA;
                        existingPjcform.PUMPING_CMD = model.PUMPING_CMD;
                        existingPjcform.PUMPING_LAND_FUND = model.PUMPING_LAND_FUND;
                        existingPjcform.PUMPING_BUILD_FUND = model.PUMPING_BUILD_FUND;
                        existingPjcform.CUT_NUM = model.CUT_NUM;
                        existingPjcform.CUT_CMD = model.CUT_CMD;
                        existingPjcform.CUT_BUILD_FUND = model.CUT_BUILD_FUND;
                        existingPjcform.TD_ALL_FUND = model.TD_ALL_FUND;
                        existingPjcform.PUMPING_FUND = model.PUMPING_FUND;
                        existingPjcform.WORK_FUND = model.WORK_FUND;
                        existingPjcform.PRICES_YEAR = model.PRICES_YEAR;
                        existingPjcform.OVER_YEAR = model.OVER_YEAR;
                        existingPjcform.SER_NOTE = model.SER_NOTE;
                        existingPjcform.UPD_DT = DateTime.Now;

                        db.SaveChanges();
                        ViewBag.Message = "修改成功";
                    } else {
                        ViewBag.Message = "找不到該記錄";
                        return SYSInfo(pjId);
                    }
                } else {
                    // 新增模式：建立新記錄
                    var newPjcform = new WCals.Models.CMETMSModels.PJCFORM();
                    //pjId = Guid.NewGuid().ToString().Substring(0, 20);
                    pjId = model.SER_FN_1 + model.SER_FN_2 + model.SER_FN_3 + model.SER_FN_4 + model.SER_FN_5 + model.SER_FN_6 + model.SER_FN_7;
                    newPjcform.PJ_ID = pjId;
                    newPjcform.SYS_ID = model.SYS_ID;
                    newPjcform.PJ_ID_TYPE = model.PJ_ID_TYPE;
                    newPjcform.SER_FN_1 = model.SER_FN_1;
                    newPjcform.SER_FN_2 = model.SER_FN_2;
                    newPjcform.SER_FN_3 = model.SER_FN_3;
                    newPjcform.SER_FN_4 = model.SER_FN_4;
                    newPjcform.SER_FN_5 = model.SER_FN_5;
                    newPjcform.SER_FN_6 = model.SER_FN_6;
                    newPjcform.SER_FN_7 = model.SER_FN_7;
                    newPjcform.SER_CITY = model.SER_CITY;
                    newPjcform.PJ_NAME = model.PJ_NAME;
                    newPjcform.PJ_AREA_NAME = model.PJ_AREA_NAME;
                    newPjcform.CHK_DT = model.CHK_DT;
                    newPjcform.CHK_FN = model.CHK_FN;
                    newPjcform.PJ_YEAR = model.PJ_YEAR;
                    newPjcform.ADM_NAME = model.ADM_NAME;
                    newPjcform.PJ_SEQ = model.PJ_SEQ;
                    newPjcform.PJ_FIRST = model.PJ_FIRST;
                    newPjcform.PJ_AREA = model.PJ_AREA;
                    newPjcform.PJ_PEO = model.PJ_PEO;
                    newPjcform.PJ_WATER = model.PJ_WATER;
                    newPjcform.PJ_IN_BOD = model.PJ_IN_BOD;
                    newPjcform.PJ_IN_SS = model.PJ_IN_SS;
                    newPjcform.PJ_OUT_BOD = model.PJ_OUT_BOD;
                    newPjcform.PJ_OUT_SS = model.PJ_OUT_SS;
                    newPjcform.WATER_TYPE = model.WATER_TYPE;
                    newPjcform.WATER_AREA = model.WATER_AREA;
                    newPjcform.MUD_TYPE = model.MUD_TYPE;
                    newPjcform.PJ_FUND = model.PJ_FUND;
                    newPjcform.MAIN_PIPO = model.MAIN_PIPO;
                    newPjcform.BRANCH_PIPO = model.BRANCH_PIPO;
                    newPjcform.PIPO_BOT_FUND = model.PIPO_BOT_FUND;
                    newPjcform.PIPO_SELF_FUND = model.PIPO_SELF_FUND;
                    newPjcform.USER_BOT = model.USER_BOT;
                    newPjcform.USER_SELF = model.USER_SELF;
                    newPjcform.USER_BOT_FUND = model.USER_BOT_FUND;
                    newPjcform.USER_SELF_FUND = model.USER_SELF_FUND;
                    newPjcform.SEWAGE_NUM = model.SEWAGE_NUM;
                    newPjcform.SEWAGE_AREA = model.SEWAGE_AREA;
                    newPjcform.SEWAGE_CMD = model.SEWAGE_CMD;
                    newPjcform.SEWAGE_LAND_FUND = model.SEWAGE_LAND_FUND;
                    newPjcform.SEWAGE_BUILD_FUND = model.SEWAGE_BUILD_FUND;
                    newPjcform.PUMPING_NUM = model.PUMPING_NUM;
                    newPjcform.PUMPING_AREA = model.PUMPING_AREA;
                    newPjcform.PUMPING_CMD = model.PUMPING_CMD;
                    newPjcform.PUMPING_LAND_FUND = model.PUMPING_LAND_FUND;
                    newPjcform.PUMPING_BUILD_FUND = model.PUMPING_BUILD_FUND;
                    newPjcform.CUT_NUM = model.CUT_NUM;
                    newPjcform.CUT_CMD = model.CUT_CMD;
                    newPjcform.CUT_BUILD_FUND = model.CUT_BUILD_FUND;
                    newPjcform.TD_ALL_FUND = model.TD_ALL_FUND;
                    newPjcform.PUMPING_FUND = model.PUMPING_FUND;
                    newPjcform.WORK_FUND = model.WORK_FUND;
                    newPjcform.PRICES_YEAR = model.PRICES_YEAR;
                    newPjcform.OVER_YEAR = model.OVER_YEAR;
                    newPjcform.SER_NOTE = model.SER_NOTE;
                    newPjcform.UPD_DT = DateTime.Now;

                    db.PJCFORM.Add(newPjcform);
                    db.SaveChanges();
                    ViewBag.Message = "新增成功";
                }

                // 處理檔案上傳
                var files = Request.Files;
                if(files.Count > 0) {
                    int uploadCount = 0;
                    DateTime now = DateTime.Now;
                    int rocYear = now.Year - 1911;
                    string month = now.Month.ToString();
                    string day = now.Day.ToString();

                    string basePath = GetBaseFilePath();
                    string relativePath = $"{rocYear}\\{month}\\{day}\\{pjId}";
                    string fullPath = System.IO.Path.Combine(basePath, relativePath);

                    if(!System.IO.Directory.Exists(fullPath)) {
                        System.IO.Directory.CreateDirectory(fullPath);
                    }

                    for(int i = 0; i < files.Count; i++) {
                        HttpPostedFileBase file = files[i];

                        if(file != null && file.ContentLength > 0) {
                            string fileName = System.IO.Path.GetFileName(file.FileName);
                            string filePath = System.IO.Path.Combine(fullPath, fileName);

                            file.SaveAs(filePath);

                            decimal fileSizeKB = (decimal)file.ContentLength / 1024;
                            decimal fileId = decimal.Parse(DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                            string filePathValue = System.IO.Path.Combine(relativePath, fileName);

                            SqlCommand cmd = new SqlCommand(
                                "INSERT INTO PJCFILE (FILE_ID, PJ_ID, FILE_LINK, FILE_PATH, FILE_NAME, FILE_SIZE, ADD_DT, UPD_DT, UPD_UID) " +
                                "VALUES (@FILE_ID, @PJ_ID, @FILE_LINK, @FILE_PATH, @FILE_NAME, @FILE_SIZE, @ADD_DT, @UPD_DT, @UPD_UID)"
                            );
                            cmd.Parameters.Add(new SqlParameter("FILE_ID", fileId));
                            cmd.Parameters.Add(new SqlParameter("PJ_ID", pjId));
                            cmd.Parameters.Add(new SqlParameter("FILE_LINK", filePathValue));
                            cmd.Parameters.Add(new SqlParameter("FILE_PATH", filePathValue));
                            cmd.Parameters.Add(new SqlParameter("FILE_NAME", fileName));
                            cmd.Parameters.Add(new SqlParameter("FILE_SIZE", fileSizeKB));
                            cmd.Parameters.Add(new SqlParameter("ADD_DT", now));
                            cmd.Parameters.Add(new SqlParameter("UPD_DT", now));
                            cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId()));

                            BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                            uploadCount++;
                        }
                    }

                    if(uploadCount > 0) {
                        ViewBag.Message += $"，成功上傳 {uploadCount} 個檔案";
                    }
                }

                // 重新載入頁面
                return SYSInfo(pjId);
            } catch(Exception ex) {
                ViewBag.Message = "操作失敗：" + ex.Message;
                return SYSInfo(model.PJ_ID);
            }
        }

        /// <summary>
        /// 刪除規劃報告
        /// </summary>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Del")]
        public ActionResult DeleteSYSInfo(SYS0030Model m) {
            try {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();

                var pjcform = db.PJCFORM.FirstOrDefault(x => x.PJ_ID == m.PJ_ID);

                if(pjcform != null) {
                    db.PJCFORM.Remove(pjcform);
                    db.SaveChanges();
                    //return Json(new { success = true, message = "刪除成功" });
                    ViewBag.msg = "刪除成功";
                    return Search(m);
                } else {
                    //return Json(new { success = false, message = "找不到該記錄" });
                    ViewBag.msg = "找不到該記錄";
                    return Search(m);
                }
            } catch(Exception ex) {
                //return Json(new { success = false, message = "刪除失敗：" + ex.Message });
                ViewBag.msg = "刪除失敗";
                return Search(m);
            }
        }

        /// <summary>
        /// 取得檔案列表
        /// </summary>
        [HttpGet]
        public ActionResult GetFileList(string pjId) {
            try {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();
                var dbFiles = db.PJCFILE
                     .Where(x => x.PJ_ID == pjId)
                     .OrderByDescending(x => x.ADD_DT)
                     .ToList();

                var files = dbFiles.Select(x => new {
                    FILE_ID = x.FILE_ID.ToString(),
                    x.FILE_NAME,
                    x.FILE_SIZE,
                    x.ADD_DT
                }).ToList();

                return Json(new { success = true, files = files }, JsonRequestBehavior.AllowGet);
            } catch(Exception ex) {
                return Json(new { success = false, message = "取得檔案列表失敗：" + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 多檔案上傳
        /// </summary>
        [HttpPost]
        public ActionResult UploadFiles(string pjId) {
            try {
                if(string.IsNullOrEmpty(pjId)) {
                    return Json(new { success = false, message = "PJ_ID 不可為空" });
                }

                var files = Request.Files;
                if(files.Count == 0) {
                    return Json(new { success = false, message = "請選擇要上傳的檔案" });
                }

                var uploadedFiles = new List<object>();

                // 建立路徑: C:\backup\SEPMIS\CMETMS\Files\PJCFILE\民國年\月\日\PJ_ID\
                DateTime now = DateTime.Now;
                int rocYear = now.Year - 1911; // 轉換為民國年
                string month = now.Month.ToString();
                string day = now.Day.ToString();

                string basePath = GetBaseFilePath();
                string relativePath = $"{rocYear}\\{month}\\{day}\\{pjId}";
                string fullPath = System.IO.Path.Combine(basePath, relativePath);

                // 確保目錄存在
                if(!System.IO.Directory.Exists(fullPath)) {
                    System.IO.Directory.CreateDirectory(fullPath);
                }

                for(int i = 0; i < files.Count; i++) {
                    HttpPostedFileBase file = files[i];

                    if(file != null && file.ContentLength > 0) {
                        // 取得檔案名稱
                        string fileName = System.IO.Path.GetFileName(file.FileName);
                        string filePath = System.IO.Path.Combine(fullPath, fileName);

                        // 儲存檔案
                        file.SaveAs(filePath);

                        // 計算檔案大小 (KB)
                        decimal fileSizeKB = (decimal)file.ContentLength / 1024;

                        // 產生 FILE_ID - 使用時間戳記
                        decimal fileId = decimal.Parse(DateTime.Now.ToString("yyyyMMddHHmmssfff"));

                        // FILE_LINK 和 FILE_PATH 存相同的值
                        string filePathValue = System.IO.Path.Combine(relativePath, fileName);

                        // 使用 ADO.NET 插入資料
                        SqlCommand cmd = new SqlCommand(
                             "INSERT INTO PJCFILE (FILE_ID, PJ_ID, FILE_LINK, FILE_PATH, FILE_NAME, FILE_SIZE, ADD_DT, UPD_DT, UPD_UID) " +
                         "VALUES (@FILE_ID, @PJ_ID, @FILE_LINK, @FILE_PATH, @FILE_NAME, @FILE_SIZE, @ADD_DT, @UPD_DT, @UPD_UID)"
                  );
                        cmd.Parameters.Add(new SqlParameter("FILE_ID", fileId));
                        cmd.Parameters.Add(new SqlParameter("PJ_ID", pjId));
                        cmd.Parameters.Add(new SqlParameter("FILE_LINK", filePathValue));
                        cmd.Parameters.Add(new SqlParameter("FILE_PATH", filePathValue));
                        cmd.Parameters.Add(new SqlParameter("FILE_NAME", fileName));
                        cmd.Parameters.Add(new SqlParameter("FILE_SIZE", fileSizeKB));
                        cmd.Parameters.Add(new SqlParameter("ADD_DT", now));
                        cmd.Parameters.Add(new SqlParameter("UPD_DT", now));
                        cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId()));

                        BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                        uploadedFiles.Add(new {
                            fileId = fileId.ToString(),
                            fileName = fileName,
                            fileSize = fileSizeKB,
                            addDate = now
                        });
                    }
                }

                return Json(new {
                    success = true,
                    message = $"成功上傳 {uploadedFiles.Count} 個檔案",
                    files = uploadedFiles
                });
            } catch(Exception ex) {
                return Json(new { success = false, message = "上傳失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// 刪除檔案
        /// </summary>
                [HttpPost]                
                public ActionResult DeleteFile(string fileId) {
                    try {
                        //decimal id = decimal.Parse(fileId);
        
                        // Use direct SQL to find file path to avoid EF caching/mapping issues
                        string sql = "SELECT FILE_PATH FROM PJCFILE WHERE FILE_ID = @FILE_ID";
                        SqlCommand cmdSelect = new SqlCommand(sql);
                        cmdSelect.Parameters.Add(new SqlParameter("FILE_ID", fileId));
        
                        DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSelect);
        
                        if (dt.Rows.Count > 0) {
                            string filePath = dt.Rows[0]["FILE_PATH"].ToString();
                            string basePath = GetBaseFilePath();
                            string fullPath = System.IO.Path.Combine(basePath, filePath);
        
                            if (System.IO.File.Exists(fullPath)) {
                                System.IO.File.Delete(fullPath);
                            }
        
                            // Use direct SQL to delete record
                            SqlCommand cmdDelete = new SqlCommand("DELETE FROM PJCFILE WHERE FILE_ID = @FILE_ID");
                            cmdDelete.Parameters.Add(new SqlParameter("FILE_ID", fileId));
                            BASE_DAO.SqlDao.CMETMSExecuteSql(cmdDelete);
        
                            return Json(new { success = true, message = "刪除成功" });
                        } else {
                            // If record not found, treat as success per user request
                            return Json(new { success = true, message = "找不到該檔案記錄(已視為刪除)" });
                        }
                    } catch (Exception ex) {
                        return Json(new { success = false, message = "刪除失敗：" + ex.Message });
                    }
                }

        /// <summary>
        /// 下載檔案
        /// </summary>
        [HttpGet]
        public ActionResult DownloadFile(string fileId) {
            try {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();
                decimal id = decimal.Parse(fileId);
                var file = db.PJCFILE.FirstOrDefault(x => x.FILE_ID == id);

                if(file != null) {
                    string basePath = GetBaseFilePath();
                    string fullPath = System.IO.Path.Combine(basePath, file.FILE_PATH);

                    if(System.IO.File.Exists(fullPath)) {
                        var contentType = MimeMapping.GetMimeMapping(file.FILE_NAME);
                        return File(fullPath, contentType, file.FILE_NAME);
                    } else {
                        return HttpNotFound("檔案不存在");
                    }
                } else {
                    return HttpNotFound("找不到該檔案記錄");
                }
            } catch(Exception ex) {
                return Content("下載失敗：" + ex.Message);
            }
        }

        private string GetBaseFilePath() {
            string uploadPath = WebConfigurationManager.AppSettings["uploadPath"];
            if (string.IsNullOrEmpty(uploadPath)) {
                uploadPath = "File/Doc/";
            }
            if (!uploadPath.StartsWith("~") && !System.IO.Path.IsPathRooted(uploadPath)) {
                uploadPath = "~/" + uploadPath;
            }
            return Server.MapPath(uploadPath);
        }

        /// <summary>
        /// 經費統計表條件選取彈跳視窗首頁 (HttpGet)
        /// </summary>
        [HttpGet]
        public ActionResult PrtSysPopup(string STAT_NO)
        {
            ViewBag.STAT_NO = STAT_NO; // 🎯 放入 ViewBag 攜帶至 View 的隱藏欄位

            string sqlCity = "select distinct CITY from CMBASE.dbo.SMCPOSTN where isnull(CITY,'') not in ('','不分區')";
            //DataTable dtCity = GetT_SQL_DataTable(sqlCity, sConnStr);
            DataTable dtCity = BASE_DAO.SqlDao.CMETMSQuery4List(sqlCity);
            ViewBag.CityList = dtCity;

            string sqlYears = "select distinct SYR from CMEPMS.dbo.TDDSTCF union select distinct SYR from CMEPMS.dbo.PJMFORM where isnull(SYR,'')<>'' order by SYR";
            //DataTable dtYears = GetT_SQL_DataTable(sqlYears, sConnStr);
            DataTable dtYears = BASE_DAO.SqlDao.CMETMSQuery4List(sqlYears);
            ViewBag.StartYears = dtYears;

            DataTable dtEndYears = dtYears.Copy();
            dtEndYears.DefaultView.Sort = "SYR desc";
            ViewBag.EndYears = dtEndYears.DefaultView.ToTable();

            return View();
        }

        /// <summary>
        /// 處理彈窗表單提交，利用 Syncfusion XlsIO 載入範本並動態渲染匯出
        /// </summary>
        [HttpPost]
        public void ExportPrtSys(FormCollection form)
        {
            string STAT_NO = form["STAT_NO"] ?? ""; // 🎯 從 Hidden 隱藏欄位接收 "S" 參數
            string[] selectedCities = form.GetValues("SelectedCities");
            string sSYS_NAME_Txt = form["sSYS_NAME"] ?? "";
            string sSYR_Val = form["sSYR"] ?? "";
            string eSYR_Val = form["eSYR"] ?? "";
            string sIS_SUP = form["IS_SUP"] ?? "";
            string sIS_ALL_YR_FUND = form["IS_ALL_YR_FUND"] ?? "Y";
            string[] dataTypeValues = form.GetValues("DATA_TYPE");

            string ssCITY = selectedCities != null ? string.Join(",", selectedCities) : "";

            // 1. 取得下水道系統
            string sql = "select * from CMETMS.dbo.PJCSYST where CITY in ('" + ssCITY.Replace(",", "','") + "')";

            if (!string.IsNullOrEmpty(STAT_NO))
            {
                sql += " and STAT_NO = '" + STAT_NO.Replace("'", "''") + "'"; // 🎯 剛性加入系統別屬性（如 'S' 污水）
            }

            if (!string.IsNullOrEmpty(sSYS_NAME_Txt))
            {
                sql += " and SYS_NAME like '%" + sSYS_NAME_Txt + "%'";
            }
            if (!string.IsNullOrEmpty(sIS_SUP))
            {
                sql += " and isnull(IS_SUP,'N')='" + sIS_SUP + "'";
            }
            sql += " order by CITY,SYS_ID";
            //DataTable PJCSYST = GetT_SQL_DataTable(sql, sConnStr);
            DataTable PJCSYST = BASE_DAO.SqlDao.CMETMSQuery4List(sql);

            string sSYS_ID = "";
            for (int i = 0; i < PJCSYST.Rows.Count; i++)
            {
                sSYS_ID += "," + PJCSYST.Rows[i]["SYS_ID"].ToString();
            }
            if (sSYS_ID != "")
            {
                sSYS_ID = sSYS_ID.Substring(1);
            }

            int isSYR = Convert.ToInt32(sSYR_Val);
            int ieSYR = Convert.ToInt32(eSYR_Val);
            if (isSYR > ieSYR)
            {
                int temp = isSYR;
                isSYR = ieSYR;
                ieSYR = temp;
            }

            // 2. 取得年度工程經費核心聯集數據
            sql = "select T.PID_SYR,T.TD_CODE,T.SIP_TD_CODE,T.TD_NAME,T.CITY,T.TOWN,T.IPJ_DESC,T.YR_FUND,T.CEN_FUND,T.LOC_FUND,T.SYS_ID,P.CITY as SYS_CITY,P.SYS_NAME";
            sql += ",isnull(T1.COST_FEE,0) + isnull(T1.PPAY_FEE,0) as REL_CEN_FUND";
            sql += ",(isnull(T.YR_FUND,0) + isnull((select sum(isnull(REV_FUND,0)) as REV_FUND from CMEPMS.dbo.CRDMNDS where TD_CODE1=T.TD_CODE and SYR=T.PID_SYR),0)) as ALL_SYR_FUND ";
            sql += " from CMEPMS.dbo.TDV_TDDCFTS T ";
            sql += " left join CMETMS.dbo.PJCSYST P on T.SYS_ID=P.SYS_ID ";
            if (!string.IsNullOrEmpty(STAT_NO))
            {
                sql += " and P.STAT_NO = '" + STAT_NO.Replace("'", "''") + "'"; // 🎯 剛性加入系統別屬性（如 'S' 污水）
            }
            sql += " left join (select T1.TD_CODE1,T2.FD_CODE,T1.SYR,T1.SMN,T2.COST_FEE,T2.PPAY_FEE from (select TD_CODE1,FD_CODE,SYR,MAX(SMN) as SMN from CMEPMS.dbo.TDDSTCF group by TD_CODE1,FD_CODE,SYR) T1 left join CMEPMS.dbo.TDDSTCF T2 on T2.TD_CODE1=T1.TD_CODE1 and T2.FD_CODE=T1.FD_CODE and T2.SYR=T1.SYR and T2.SMN=T1.SMN) as T1 on T1.TD_CODE1=T.TD_CODE and T1.FD_CODE=T.FD_CODE and T1.SYR=T.PID_SYR";
            sql += " where T.SYS_ID in ('" + sSYS_ID.Replace(",", "','") + "')";
            if (!string.IsNullOrEmpty(sSYR_Val) && !string.IsNullOrEmpty(eSYR_Val))
            {
                sql += " and PID_SYR between '" + isSYR.ToString("000") + "' and '" + ieSYR.ToString("000") + "'";
            }
            sql += " order by T.PID_SYR,P.CITY,T.TD_CODE";
            DataTable dt1 = BASE_DAO.SqlDao.CMEPMSQuery4List(sql);

            DataTable dt = dt1.Clone();
            if (sIS_ALL_YR_FUND == "N")
            {
                DataRow[] rr = dt1.Select("ALL_SYR_FUND>0");
                for (int i = 0; i < rr.Length; i++)
                {
                    dt.ImportRow(rr[i]);
                }
            }
            else
            {
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    dt.ImportRow(dt1.Rows[i]);
                }
            }

            // 建立實支數暫存表
            string sTD_CODE_L = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sTD_CODE_L += "," + dt.Rows[i]["TD_CODE"].ToString();
            }
            if (sTD_CODE_L != "")
            {
                sTD_CODE_L = sTD_CODE_L.Substring(1);
            }

            sql = "select TD_CODE1 as TD_CODE,SYR,sum(COST_FEE+PPAY_FEE) as REL_CEN_FUND from CMEPMS.dbo.TDDSTCF where TD_CODE1 in ('" + sTD_CODE_L.Replace(",", "','") + "') and SMN='12' group by TD_CODE1,SYR";
            DataTable TDDSTCF = BASE_DAO.SqlDao.CMETMSQuery4List(sql);

            DataTable TDMFORM = dt.DefaultView.ToTable(true, new string[] { "SIP_TD_CODE", "TD_NAME", "TD_CODE", "SYS_CITY", "SYS_ID", "SYS_NAME", "IPJ_DESC" });
            TDMFORM.DefaultView.Sort = "SYS_CITY,TD_CODE";
            TDMFORM = TDMFORM.DefaultView.ToTable();

            DataTable TDMFORM_CODE = dt.DefaultView.ToTable(true, new string[] { "SIP_TD_CODE", "TD_CODE", "SYS_CITY", "SYS_ID" });
            TDMFORM_CODE.DefaultView.Sort = "SYS_CITY,TD_CODE";
            TDMFORM_CODE = TDMFORM_CODE.DefaultView.ToTable();

            // 開啟範本試算表
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2007;

            string sFullFile = Server.MapPath("~/Templates/SYS0030N_SYS.xlsx");
            FileStream readFile = new FileStream(sFullFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IWorkbook workbook = application.Workbooks.Open(readFile);
            readFile.Close();

            IWorksheet sheet1 = workbook.Worksheets["Sheet1"];
            IWorksheet sheet1_1 = workbook.Worksheets["Sheet1_1"];

            sheet1[1, 1, 1, 2].Merge();
            sheet1_1[2, 1, 5, 1].CopyTo(sheet1[2, 1]);
            sheet1_1[2, 2, 5, 2].CopyTo(sheet1[2, 2]);
            sheet1_1[2, 3, 5, 3].CopyTo(sheet1[2, 3]);
            sheet1_1[2, 4, 5, 4].CopyTo(sheet1[2, 4]);
            sheet1_1[2, 5, 5, 5].CopyTo(sheet1[2, 5]);

            int col = 6;
            int j = 0;
            int end_col = 5;

            for (int i = isSYR; i <= ieSYR; i++)
            {
                sheet1_1[2, col, 6, col].CopyTo(sheet1[2, (col + j * 4)]);
                sheet1_1[2, (col + 1), 5, (col + 1)].CopyTo(sheet1[2, (col + j * 4 + 1)]);
                sheet1_1[2, (col + 2), 5, (col + 2)].CopyTo(sheet1[2, (col + j * 4 + 2)]);
                sheet1_1[2, (col + 3), 5, (col + 3)].CopyTo(sheet1[2, (col + j * 4 + 3)]);

                // 🎯 修正：先解除可能殘留的範本合併，再重新合併該年度完整的 4 個欄位 (col 到 col + 3)
                var yearRange = sheet1[3, (col + j * 4), 3, (col + j * 4 + 3)];
                yearRange.UnMerge();
                yearRange.Merge();
                yearRange.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                yearRange.Text = i.ToString("000") + "年度";

                end_col += 4;
                j++;
            }

            sheet1_1[2, col, 6, col].CopyTo(sheet1[2, (col + j * 4)]);
            sheet1_1[2, (col + 1), 5, (col + 1)].CopyTo(sheet1[2, (col + j * 4 + 1)]);
            sheet1_1[2, (col + 2), 5, (col + 2)].CopyTo(sheet1[2, (col + j * 4 + 2)]);
            sheet1_1[2, (col + 3), 5, (col + 3)].CopyTo(sheet1[2, (col + j * 4 + 3)]);

            // 🎯 修正：「總計」欄位同樣橫跨最後一組完整的 4 個欄位
            var totalRange = sheet1[3, (col + j * 4), 3, (col + j * 4 + 3)];
            totalRange.UnMerge();
            totalRange.Merge();
            totalRange.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            totalRange.Text = "總計";
            end_col += 4;

            sheet1[2, 6, 2, end_col].Merge();
            sheet1[2, 6, 2, end_col].Text = "經費(元)";

            List<string> dTypes = dataTypeValues != null ? new List<string>(dataTypeValues) : new List<string>();

            if (dTypes.Contains("ALL"))
            {
                CreateSheet1(workbook, sheet1, dt, TDMFORM, TDMFORM_CODE, end_col, isSYR, ieSYR, "全部工程", "", "");
            }
            if (dTypes.Contains("CITY") && selectedCities != null)
            {
                foreach (var city in selectedCities)
                {
                    CreateSheet1(workbook, sheet1, dt, TDMFORM, TDMFORM_CODE, end_col, isSYR, ieSYR, city, city, "");
                }
            }
            if (dTypes.Contains("SYS_ID"))
            {
                for (int m = 0; m < PJCSYST.Rows.Count; m++)
                {
                    CreateSheet1(workbook, sheet1, dt, TDMFORM, TDMFORM_CODE, end_col, isSYR, ieSYR, PJCSYST.Rows[m]["SYS_NAME"].ToString(), "", PJCSYST.Rows[m]["SYS_ID"].ToString());
                }
            }

            sheet1.Remove();
            sheet1_1.Remove();

            string sFileName = HttpUtility.UrlEncode("經費統計表", System.Text.Encoding.UTF8);

            MemoryStream stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;

            workbook.Close();
            excelEngine.Dispose();

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + sFileName + ".xlsx");
            Response.BinaryWrite(stream.ToArray());
            Response.End();
        }

        private void CreateSheet1(IWorkbook workbook, IWorksheet sheet1, DataTable dt, DataTable TDMFORM, DataTable TDMFORM_CODE, int end_col, int isSYR, int ieSYR, string SheetNAME, string CITY, string SYS_ID)
        {
            IWorksheet sheet1tmp = workbook.Worksheets.AddCopyBefore(sheet1, sheet1);

            // =========================================================================
            // 🎯【防禦性命名優化】剛性隔離 Excel 分頁命名限制與重複衝突
            // =========================================================================
            string finalName = string.IsNullOrEmpty(SheetNAME) ? "未命名" : SheetNAME;

            // 1. 過濾 Excel 工作表名稱不允許的非法字元: \ / ? * [ ]
            char[] illegalChars = { '\\', '/', '?', '*', '[', ']' };
            foreach (char c in illegalChars)
            {
                finalName = finalName.Replace(c, '_');
            }

            // 2. 嚴格截斷至 Excel 規定的 31 字元上限
            if (finalName.Length > 31)
            {
                finalName = finalName.Substring(0, 31);
            }

            // 3. 自動動態查重：若發現名稱重複，自動附加序號如 (1), (2)
            string baseName = finalName;
            int counter = 1;
            bool exists = true;

            while (exists)
            {
                exists = false;
                foreach (IWorksheet ws in workbook.Worksheets)
                {
                    if (string.Equals(ws.Name, finalName, StringComparison.OrdinalIgnoreCase))
                    {
                        exists = true;
                        break;
                    }
                }

                if (exists)
                {
                    string suffix = "(" + counter + ")";
                    // 附加序號時亦須剛性符合 31 字元限制
                    if (baseName.Length + suffix.Length > 31)
                    {
                        finalName = baseName.Substring(0, 31 - suffix.Length) + suffix;
                    }
                    else
                    {
                        finalName = baseName + suffix;
                    }
                    counter++;
                }
            }

            sheet1tmp.Name = finalName;

            string sWhere = "isnull(TD_CODE,'')<>''";
            if (CITY != "")
            {
                sheet1tmp["A1"].Text = "下水道系統歸屬縣市別：" + SheetNAME;
                sWhere += " and SYS_CITY='" + CITY + "'";
            }
            if (SYS_ID != "")
            {
                sheet1tmp["A1"].Text = "下水道系統：" + SheetNAME;
                sWhere += " and SYS_ID='" + SYS_ID + "'";
            }

            DataRow[] rr = TDMFORM_CODE.Select(sWhere, "SYS_CITY,TD_CODE");
            int line = 6;
            for (int i = 0; i < rr.Length; i++)
            {
                sheet1tmp[5, 1, 5, end_col].CopyTo(sheet1tmp[(line + i), 1]);
                sheet1tmp[(line + i), 1, (line + i), end_col].RowHeight = sheet1tmp[5, 1, 5, end_col].RowHeight;
            }

            line = 5;
            for (int i = 0; i < rr.Length; i++)
            {
                DataRow[] rr1 = TDMFORM.Select("TD_CODE='" + rr[i]["TD_CODE"].ToString() + "'");
                sheet1tmp["A" + (line + i)].Text = rr[i]["SIP_TD_CODE"].ToString();
                if (rr1.Length > 0)
                {
                    sheet1tmp["B" + (line + i)].Text = rr1[0]["TD_NAME"].ToString();
                    sheet1tmp["C" + (line + i)].Text = rr1[0]["IPJ_DESC"].ToString();
                    sheet1tmp["D" + (line + i)].Text = rr1[0]["SYS_CITY"].ToString();
                    sheet1tmp["E" + (line + i)].Text = rr1[0]["SYS_NAME"].ToString();
                }

                double sCEN_FUND = 0, sLOC_FUND = 0, sYR_FUND = 0, sREL_CEN_FUND = 0;
                int j = 0;
                int dataCol = 6;
                for (int k = isSYR; k <= ieSYR; k++)
                {
                    sCEN_FUND += Convert.ToDouble(getFUND(dt, k.ToString("000"), rr[i]["TD_CODE"].ToString(), CITY, SYS_ID, "CEN_FUND"));
                    sLOC_FUND += Convert.ToDouble(getFUND(dt, k.ToString("000"), rr[i]["TD_CODE"].ToString(), CITY, SYS_ID, "LOC_FUND"));
                    sYR_FUND += Convert.ToDouble(getFUND(dt, k.ToString("000"), rr[i]["TD_CODE"].ToString(), CITY, SYS_ID, "YR_FUND"));
                    sREL_CEN_FUND += Convert.ToDouble(getFUND(dt, k.ToString("000"), rr[i]["TD_CODE"].ToString(), CITY, SYS_ID, "REL_CEN_FUND"));

                    sheet1tmp[(line + i), dataCol + j * 4].Value = getFUND(dt, k.ToString("000"), rr[i]["TD_CODE"].ToString(), CITY, SYS_ID, "CEN_FUND");
                    sheet1tmp[(line + i), dataCol + j * 4 + 1].Value = getFUND(dt, k.ToString("000"), rr[i]["TD_CODE"].ToString(), CITY, SYS_ID, "LOC_FUND");
                    sheet1tmp[(line + i), dataCol + j * 4 + 2].Value = getFUND(dt, k.ToString("000"), rr[i]["TD_CODE"].ToString(), CITY, SYS_ID, "YR_FUND");
                    sheet1tmp[(line + i), dataCol + j * 4 + 3].Value = getFUND(dt, k.ToString("000"), rr[i]["TD_CODE"].ToString(), CITY, SYS_ID, "REL_CEN_FUND");
                    j++;
                }
                sheet1tmp[(line + i), dataCol + j * 4].Value = sCEN_FUND.ToString();
                sheet1tmp[(line + i), dataCol + j * 4 + 1].Value = sLOC_FUND.ToString();
                sheet1tmp[(line + i), dataCol + j * 4 + 2].Value = sYR_FUND.ToString();
                sheet1tmp[(line + i), dataCol + j * 4 + 3].Value = sREL_CEN_FUND.ToString();
            }

            sheet1tmp["B" + (line + rr.Length)].Text = "合計";
            double tsCEN_FUND = 0, tsLOC_FUND = 0, tsYR_FUND = 0, tsREL_CEN_FUND = 0;
            int totalJ = 0;
            int totalCol = 6;
            for (int k = isSYR; k <= ieSYR; k++)
            {
                tsCEN_FUND += Convert.ToDouble(getFUND(dt, k.ToString("000"), "", CITY, SYS_ID, "CEN_FUND"));
                tsLOC_FUND += Convert.ToDouble(getFUND(dt, k.ToString("000"), "", CITY, SYS_ID, "LOC_FUND"));
                tsYR_FUND += Convert.ToDouble(getFUND(dt, k.ToString("000"), "", CITY, SYS_ID, "YR_FUND"));
                tsREL_CEN_FUND += Convert.ToDouble(getFUND(dt, k.ToString("000"), "", CITY, SYS_ID, "REL_CEN_FUND"));

                sheet1tmp[(line + rr.Length), totalCol + totalJ * 4].Value = getFUND(dt, k.ToString("000"), "", CITY, SYS_ID, "CEN_FUND");
                sheet1tmp[(line + rr.Length), totalCol + totalJ * 4 + 1].Value = getFUND(dt, k.ToString("000"), "", CITY, SYS_ID, "LOC_FUND");
                sheet1tmp[(line + rr.Length), totalCol + totalJ * 4 + 2].Value = getFUND(dt, k.ToString("000"), "", CITY, SYS_ID, "YR_FUND");
                sheet1tmp[(line + rr.Length), totalCol + totalJ * 4 + 3].Value = getFUND(dt, k.ToString("000"), "", CITY, SYS_ID, "REL_CEN_FUND");
                totalJ++;
            }
            sheet1tmp[(line + rr.Length), totalCol + totalJ * 4].Value = tsCEN_FUND.ToString();
            sheet1tmp[(line + rr.Length), totalCol + totalJ * 4 + 1].Value = tsLOC_FUND.ToString();
            sheet1tmp[(line + rr.Length), totalCol + totalJ * 4 + 2].Value = tsYR_FUND.ToString();
            sheet1tmp[(line + rr.Length), totalCol + totalJ * 4 + 3].Value = tsREL_CEN_FUND.ToString();
        }

        private string getFUND(DataTable dt, string SYR, string TD_CODE, string CITY, string SYS_ID, string FUND_TYPE)
        {
            double sFUND = 0;
            string sWhere = "isnull(TD_CODE,'')<>''";
            if (SYR != "") sWhere += " and PID_SYR='" + SYR + "'";
            if (TD_CODE != "") sWhere += " and TD_CODE='" + TD_CODE + "'";
            if (CITY != "") sWhere += " and SYS_CITY='" + CITY + "'";
            if (SYS_ID != "") sWhere += " and SYS_ID='" + SYS_ID + "'";

            DataRow[] rr = dt.Select(sWhere);
            for (int i = 0; i < rr.Length; i++)
            {
                sFUND += chkdouble(rr[i][FUND_TYPE].ToString());
            }
            return sFUND.ToString("0");
        }

        private DataTable GetT_SQL_DataTable(string sql, string ConnStr)
        {
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection(ConnStr))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            dt.Columns.Add(reader.GetName(i));
                        }
                        while (reader.Read())
                        {
                            DataRow dtrow = dt.NewRow();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                dtrow[reader.GetName(i)] = reader[reader.GetName(i)];
                            }
                            dt.Rows.Add(dtrow);
                        }
                    }
                }
            }
            return dt;
        }

        private double chkdouble(string val)
        {
            double ddt = 0;
            return double.TryParse(val, out ddt) ? ddt : 0;
        }

    }
}