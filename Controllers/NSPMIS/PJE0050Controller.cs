using Newtonsoft.Json.Linq;
using PagedList;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using static WCals.classes.john.BASE_DAO;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMEPMSModels;
using WCals.ViewModels.NewModels;
using System.IO;
using Aspose.Cells;

namespace WCals.Controllers.NSPMIS
{
    public class PJE0050Controller : Controller {
        private readonly Models.CMEPMSModels.CMEPMSEntities db_pms = new Models.CMEPMSModels.CMEPMSEntities();
        private readonly Models.CMETMSModels.CMETMSEntities db_tms = new Models.CMETMSModels.CMETMSEntities();
        private readonly Models.CMPSGSModels.CMPSGSEntities db_sgs = new Models.CMPSGSModels.CMPSGSEntities();
        int pageSize = 10;        

        public ActionResult Index(PJE0030Model model) {

            string userid = Helpers.WebSiteHelper.getUserId();
            int page = (model.prePage == 0) ? 1 : model.prePage;

            SqlCommand cmd = new SqlCommand("select B.PJ_CODE, B.PID, W.SYR, W.PJ_NAME as WPJ_NAME, B.PJ_NAME as BPJ_NAME, B.PJ_FUND," +
                "case (select count(1) from PJDFDCR where PID = B.PID)" +
                "when 0 then '原預算'else '第' + convert(nvarchar(3), " +
                "(select max(SEQ_NO) from PJDFDCR where PID = B.PID)) +'期修正'" +
                "end as SEQ_NAME from PJMFORM B left join PJMFORM W on B.AN_PID = W.PID and W.PJ_TYPE = 'W'" +
                "where B.PJ_TYPE = 'B'and B.ADV_UID='" + userid + "'order by SYR desc");

            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            cmd = new SqlCommand("select B.PJ_CODE, B.PID, W.SYR, W.PJ_NAME as WPJ_NAME, B.PJ_NAME as BPJ_NAME, B.PJ_FUND," +
                "case (select count(1) from PJDFDCR where PID = B.PID)" +
                "when 0 then '原預算'else '第' + convert(nvarchar(3), " +
                "(select max(SEQ_NO) from PJDFDCR where PID = B.PID)) +'期修正'" +
                "end as SEQ_NAME from (PJMFORM B left join PJMFORM W on B.AN_PID = W.PID and W.PJ_TYPE = 'W')left join PJDAUTH a ON B.PID=A.PID " +
                "where B.PJ_TYPE = 'B'and A.EXE_UID='" + userid + "'order by SYR desc");
            ps.Merge(BASE_DAO.SqlDao.CMEPMSQuery4List(cmd));
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.MainList = plist;
            return View("PJE0050Index", model);
        }

        //查詢
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SearchIndex")]
        public ActionResult SearchIndex(PJE0030Model model) {

            string userid = Helpers.WebSiteHelper.getUserId();
            int page = (model.prePage == 0) ? 1 : model.prePage;
            StringBuilder sql = new StringBuilder("select B.PJ_CODE, B.PID, W.SYR, W.PJ_NAME as WPJ_NAME, B.PJ_NAME as BPJ_NAME, B.PJ_FUND," +
                "case (select count(1) from PJDFDCR where PID = B.PID)" +
                "when 0 then '原預算'else '第' + convert(nvarchar(3), " +
                "(select max(SEQ_NO) from PJDFDCR where PID = B.PID)) +'期修正'" +
                "end as SEQ_NAME from PJMFORM B left join PJMFORM W on B.AN_PID = W.PID and W.PJ_TYPE = 'W'" +
                "where B.PJ_TYPE = 'B'and B.ADV_UID='" + userid + "'");
            if(model.stype.Equals("SYR")) {
                sql.Append("AND W." + model.stype + " LIKE '%" + model.keyword + "%' order by SYR desc");
            }
            if(model.stype.Equals("PJ_NAME")) {
                sql.Append("AND B." + model.stype + " LIKE '%" + model.keyword + "%' order by SYR desc");
            }
            if(model.stype.Equals("PJ_CODE")) {
                sql.Append("AND B." + model.stype + " LIKE '%" + model.keyword + "%' order by SYR desc");
            }

            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(sql.ToString());
            sql = new StringBuilder("select B.PJ_CODE, B.PID, W.SYR, W.PJ_NAME as WPJ_NAME, B.PJ_NAME as BPJ_NAME, B.PJ_FUND," +
                "case (select count(1) from PJDFDCR where PID = B.PID)" +
                "when 0 then '原預算'else '第' + convert(nvarchar(3), " +
                "(select max(SEQ_NO) from PJDFDCR where PID = B.PID)) +'期修正'" +
                "end as SEQ_NAME from (PJMFORM B left join PJMFORM W on B.AN_PID = W.PID and W.PJ_TYPE = 'W')left join PJDAUTH a ON B.PID=A.PID " +
                "where B.PJ_TYPE = 'B'and A.EXE_UID='" + userid + "'");
            if(model.stype.Equals("SYR")) {
                sql.Append("AND W." + model.stype + " LIKE '%" + model.keyword + "%' order by SYR desc");
            }
            if(model.stype.Equals("PJ_NAME")) {
                sql.Append("AND B." + model.stype + " LIKE '%" + model.keyword + "%' order by SYR desc");
            }
            if(model.stype.Equals("PJ_CODE")) {
                sql.Append("AND B." + model.stype + " LIKE '%" + model.keyword + "%' order by SYR desc");
            }
            ps.Merge(BASE_DAO.SqlDao.CMEPMSQuery4List(sql.ToString()));
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, 1, pageSize);
            model.MainList = plist;
            return View("PJE0050Index", model);
        }

        //列印
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Print")]
        public ActionResult Print(PJE0030Model model) {

            // 載入模板
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/PJE0050.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            Worksheet sheet = workbook.Worksheets[0];

            //做預存程序
            DataTable dt;
            using(Models.CMEPMSModels.CMEPMSEntities cmepms = new Models.CMEPMSModels.CMEPMSEntities())
            using(SqlConnection conn = new SqlConnection(cmepms.Database.Connection.ConnectionString))
            using(SqlCommand cmd = new SqlCommand("SP_RCD0010_UDF003_3", conn)) {
                cmd.CommandType = CommandType.StoredProcedure;
                // 輸入參數
                cmd.Parameters.AddWithValue("@PID", model.PID);
                cmd.Parameters.AddWithValue("@SEQ_NO", model.SEQ_NO);
                dt = new DataTable();
                conn.Open();
                using(SqlDataReader reader = cmd.ExecuteReader()) {
                    dt.Load(reader);  // 將結果讀入 DataTable
                }
            }
            int startRow = 2;
            foreach(DataRow d in dt.Rows) {
                sheet.Cells[startRow, 0].PutValue(d["TD_CODE"]);
                sheet.Cells[startRow, 1].PutValue(d["SYR"]);
                sheet.Cells[startRow, 2].PutValue(d["W_CODE"]);
                sheet.Cells[startRow, 3].PutValue(d["PJ_CODE"]);
                sheet.Cells[startRow, 4].PutValue(d["IPJ_CODE"]);
                sheet.Cells[startRow, 5].PutValue(d["CITY_CODE"]);
                sheet.Cells[startRow, 6].PutValue(d["TL_CODE"]);
                sheet.Cells[startRow, 7].PutValue(d["TOD_CODE"]);
                sheet.Cells[startRow, 8].PutValue(d["TD_NAME"]);
                sheet.Cells[startRow, 9].PutValue(d["YR_FUND"]);
                sheet.Cells[startRow, 10].PutValue(d["CEN_FUND"]);
                sheet.Cells[startRow, 11].PutValue(d["LOC_FUND"]);
                sheet.Cells[startRow, 12].PutValue(d["FIT_DESC"]);
                sheet.Cells[startRow, 13].PutValue(d["TIT_DESC"]);
                sheet.Cells[startRow, 14].PutValue(d["FIT_DESC"]);
                sheet.Cells[startRow, 15].PutValue(d["TIT_DESC"]);
                //sheet.Cells[startRow, 12].Formula = $"=B{startRow + 1}*C{startRow + 1}";

                startRow++;
            }

            sheet.AutoFitColumns();

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"PJE0050_{timestamp}.xlsx";
            using(var ms = new MemoryStream()) {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                ms.Position = 0;
                return File(ms.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName);
            }
        }
        
        //維護
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Maintain")]
        public ActionResult Maintain(PJE0030Model model) {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            int page = (model.page == 0) ? 1 : model.page;
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJDFDCR p WHERE p.PID=@PID ORDER BY SEQ_NO DESC");
            //cmd.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.List1 = plist;
            model.PJDFDCR = new PJDFDCR();
            model.PJMFORM = db_pms.PJMFORM.Where(w => w.PID == model.PID).FirstOrDefault();
            return View("PJE0050Maintain", model);
        }        

        [HttpPost]
        public JsonResult GetPJDFDCR(string PID, string SEQ_NO) {
            SqlCommand cmd = new SqlCommand("SELECT m.FLOCK_FG, p.* FROM PJDFDCR p, PJMFORM m WHERE m.PID=p.PID AND p.PID=@PID AND p.SEQ_NO=@SEQ_NO");
            cmd.Parameters.Add(new SqlParameter("PID", PID));
            cmd.Parameters.Add(new SqlParameter("SEQ_NO", SEQ_NO));
            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 管制
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Manage")]
        public ActionResult Manage(PJE0030Model model) {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            int page = (model.page == 0) ? 1 : model.page;


            int SEQ_MAX = util.GetSEQ_NO(model.PID);
            int SEQ_NO;
            if(model.SEQ_NO == 0) {
                //0為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model.SEQ_NO = SEQ_NO;
                //GetPJDFDCR_Max(model.PID, model);
                model = GetPJDFDCR_Max(model.PID, model, page, 1);
            } else {
                //已選取期數
                SEQ_NO = model.SEQ_NO;
                if(SEQ_NO == SEQ_MAX) {
                    //取最新
                    //GetPJDFDCR_Max(model.PID, model);
                    model = GetPJDFDCR_Max(model.PID, model, page, 1);
                } else {
                    //取某期
                    model = GetPJDFDCR_By(model.PID, model, page, 1);
                }
            }

            Session["fund"] =
                SqlDao.CMEPMSQuery4List("select sum(YR_FUND)as YR_FUND, " +
                "sum(CEN_FUND)as CEN_FUND, sum(LOC_FUND)as LOC_FUND from " +
                "TDV_TDDCFTS where PID =" + model.PID);
            GetSelectList(model.PID);
            return View("PJE0050Manage", model);
        }

        public ActionResult ManageInfo(int PID, int SEQ, string SYR) {

            int page = 1;
            PJE0030Model model = new PJE0030Model();
            model.PID = PID;
            model.SEQ_NO = SEQ;
            model.SYR = SYR;

            int SEQ_MAX = util.GetSEQ_NO(model.PID);
            int SEQ_NO;
            if(model.SEQ_NO == 0) {
                //0為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model.SEQ_NO = SEQ_NO;
                //GetPJDFDCR_Max(model.PID, model);
                model = GetPJDFDCR_Max(model.PID, model, page, 1);
            } else {
                //已選取期數
                SEQ_NO = model.SEQ_NO;
                if(SEQ_NO == SEQ_MAX) {
                    //取最新
                    //GetPJDFDCR_Max(model.PID, model);
                    model = GetPJDFDCR_Max(model.PID, model, page, 1);
                } else {
                    //取某期
                    model = GetPJDFDCR_By(model.PID, model, page, 1);
                }
            }
            return View("PJE0050ManageInfo", model);
        }


        public ActionResult ManageList1(
            string PJ_CATY_CODE, string TD_NAME,
            string CITY, string TOWN,
            string TD_CATY, string EXE_TYPE,
            string IPJ_CODE, string N_IPJ_CODE,
            string EMG_CATY, string ACC_ID,
            int PID, int page, int SEQ_NO
            ) {
            PJE0030Model model = new PJE0030Model();
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.ACC_ID = ACC_ID;
            model.PID = PID;
            model.SEQ_NO = SEQ_NO;
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.EMG_CATY = EMG_CATY;
            model.ACC_ID = ACC_ID;

            page = (page == 0) ? 1 : page;
            model.page = page;
            int SEQ_MAX = util.GetSEQ_NO(model.PID);
            //int SEQ_NO;
            if(SEQ_NO == 0) {
                //0為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model = GetPJDFDCR_Max(model.PID, model, model.page, 1);
            } else {
                //已選取期數                
                if(SEQ_NO == SEQ_MAX) {
                    //取最新
                    model = GetPJDFDCR_Max(model.PID, model, model.page, 1);
                } else {
                    //取某期
                    model = GetPJDFDCR_By(model.PID, model, model.page, 1);
                }
            }
            //model.SEQ_NO = SEQ_NO;
            GetSelectList(model.PID);
            return View("PJE0050ManageInfo", model);
        }

        public ActionResult Search(int PID, string PJ_CATY_CODE, string TD_NAME,
            string CITY, string TOWN, string TD_CATY, string EXE_TYPE,
            string IPJ_CODE, string N_IPJ_CODE, string EMG_CATY, string ACC_ID, int zero) {
            PJE0030Model model = new PJE0030Model();
            int max = util.GetSEQ_NO(model.PID);
            model.PID = PID;
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.EMG_CATY = EMG_CATY;
            model.ACC_ID = ACC_ID;
            model.page = 1;

            if(model.SEQ_NO == max) {
                model = GetPJDFDCR_Max(model.PID, model, 1, 1);//2:經費為限
            } else {
                model = GetPJDFDCR_By(model.PID, model, 1, 1);
            }
            GetSelectList(model.PID);
            return View("PJE0050ManageInfo", model);
        }


        //經費為0頁
        public ActionResult ManageList2(
            string PJ_CATY_CODE, string TD_NAME,
            string CITY, string TOWN,
            string TD_CATY, string EXE_TYPE,
            string IPJ_CODE, string N_IPJ_CODE,
            string EMG_CATY, string ACC_ID,
            int PID, int page, int SEQ_NO
            ) {
            PJE0030Model model = new PJE0030Model();
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.ACC_ID = ACC_ID;
            model.PID = PID;
            model.SEQ_NO = SEQ_NO;
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.EMG_CATY = EMG_CATY;
            model.ACC_ID = ACC_ID;


            page = (page == 0) ? 1 : page;
            model.page = page;
            int SEQ_MAX = util.GetSEQ_NO(model.PID);

            if(SEQ_NO == 0) {
                //0為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model = GetPJDFDCR_Max(model.PID, model, model.page, 2);
            } else {
                //已選取期數
                //SEQ_NO = model.SEQ_NO;
                if(SEQ_NO == SEQ_MAX) {
                    //取最新
                    model = GetPJDFDCR_Max(model.PID, model, model.page, 2);
                } else {
                    //取某期
                    model = GetPJDFDCR_By(model.PID, model, model.page, 2);
                }
            }
            //model.SEQ_NO = SEQ_NO;
            GetSelectList(model.PID);
            return View("PJE0050ManageInfo", model);
        }

        /// <summary>
        /// 保留款頁
        /// </summary>
        /// <param name="tid"></param>
        /// <returns></returns>
        public ActionResult EditCRDMNDS(decimal tid, decimal pid) {
            PJE0030Model model = new PJE0030Model();
            model.PID = pid;
            try {
                model.CRDMNDS = db_pms.CRDMNDS.Where(w => w.TID == tid).FirstOrDefault();
            } catch {
                model.CRDMNDS = new CRDMNDS();
            }
            return View("PJE0050EditCRDMNDS", model);
        }


        //奊 
        public ActionResult ManageList3(
            string PJ_CATY_CODE, string TD_NAME,
            string CITY, string TOWN,
            string TD_CATY, string EXE_TYPE,
            string IPJ_CODE, string N_IPJ_CODE,
            string EMG_CATY, string ACC_ID,
            int PID, int page, int SEQ_NO
            ) {
            PJE0030Model model = new PJE0030Model();

            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.ACC_ID = ACC_ID;
            model.PID = PID;
            model.SEQ_NO = SEQ_NO;
            model.PJ_CATY_CODE = PJ_CATY_CODE;
            model.TD_NAME = TD_NAME;
            model.CITY = CITY;
            model.TOWN = TOWN;
            model.TD_CATY = TD_CATY;
            model.EXE_TYPE = EXE_TYPE;
            model.IPJ_CODE = IPJ_CODE;
            model.N_IPJ_CODE = N_IPJ_CODE;
            model.EMG_CATY = EMG_CATY;
            model.ACC_ID = ACC_ID;


            page = (page == 0) ? 1 : page;
            model.page = page;
            int SEQ_MAX = util.GetSEQ_NO(model.PID);
            //int SEQ_NO;
            if(SEQ_NO == 0) {
                //0為初次進入
                SEQ_NO = util.GetSEQ_NO(model.PID);
                model = GetPJDFDCR_Max(model.PID, model, model.page, 3);
            } else {
                //已選取期數                
                if(SEQ_NO == SEQ_MAX) {
                    //取最新
                    model = GetPJDFDCR_Max(model.PID, model, model.page, 3);
                } else {
                    //取某期
                    model = GetPJDFDCR_By(model.PID, model, model.page, 3);
                }
            }
            model.PJMFORM = db_pms.PJMFORM.Where(w => w.PID == model.PID).FirstOrDefault();
            GetSelectList(model.PID);

            return View("PJE0050ManageInfo", model);
        }

        //各項下拉選單
        private void GetSelectList(object PID) {
            SqlCommand cmd = new SqlCommand("select SEQ_NO from PJDFDCR where PID = @PID ORDER BY SEQ_NO DESC");
            cmd.Parameters.Add(new SqlParameter("PID", PID));
            ViewBag.sels = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
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
            ViewBag.PJ_CATY_CODE = tools.ToSelectList(BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT CATY_CODE as id, CATY_NAME as name FROM PJDCATY WHERE PID=" + PID), "id", "name", null);
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

        }

        //取最新期CTDMNDS
        private PJE0030Model GetPJDFDCR_Max(object PID, PJE0030Model m, int page, int zero) {
            //均取最新期TD
            StringBuilder sb = new StringBuilder("SELECT * FROM TDV_TDDCFTS WHERE PID = @PID ");
            string sql;

            //分項
            if(!string.IsNullOrEmpty(m.PJ_CATY_CODE)) {
                sb.Append("AND PJ_CATY_CODE='" + m.PJ_CATY_CODE + "'");
            }
            //名稱
            if(!string.IsNullOrEmpty(m.TD_NAME)) {
                sb.Append("AND TD_NAME LIKE'%" + m.TD_NAME + "%'");
            }
            //縣市
            if(!string.IsNullOrEmpty(m.CITY)) {
                sb.Append("AND CITY='" + m.CITY + "'");
            }
            //鄉鎮
            if(!string.IsNullOrEmpty(m.TOWN)) {
                sb.Append("AND TOWN='" + m.TOWN + "'");
            }
            //工作類型
            if(!string.IsNullOrEmpty(m.TD_CATY)) {
                sb.Append("AND TD_CATY='" + m.TD_CATY + "'");
            }
            //執行方式
            if(!string.IsNullOrEmpty(m.EXE_TYPE)) {
                sb.Append("AND EXE_TYPE='" + m.EXE_TYPE + "'");
            }
            //經費層級
            if(!string.IsNullOrEmpty(m.IPJ_CODE)) {
                sb.Append("AND IPJ_CODE='" + m.IPJ_CODE + "'");
            }
            //辦理層級
            if(!string.IsNullOrEmpty(m.N_IPJ_CODE)) {
                sb.Append("AND N_IPJ_CODE='" + m.N_IPJ_CODE + "'");
            }
            //分類
            if(!string.IsNullOrEmpty(m.EMG_CATY)) {
                sb.Append("AND EMG_CATY='" + m.EMG_CATY + "'");
            }
            //會計科目
            if(!string.IsNullOrEmpty(m.ACC_ID)) {
                sb.Append("AND ACC_ID='" + m.ACC_ID + "'");
            }

            sql = sb.ToString();
            SqlCommand cmd;
            List<DataRow> list;
            DataTable ps;
            IPagedList<DataRow> plist;
            if(zero == 1) {
                sql += "AND(YR_FUND>0 OR YR_FUND IS NULL) ORDER BY TID DESC";
                cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));
                //ViewBag.List1 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List1 = plist;

                sb = new StringBuilder("select sum(YR_FUND)as YR_FUND, " +
                "sum(CEN_FUND)as CEN_FUND, sum(LOC_FUND)as LOC_FUND from " +
                "TDV_TDDCFTS where TID IN(");
                foreach(DataRow d in plist) {
                    sb.Append("'" + d["TID"] + "',");
                }
                sb.Length--;
                sb.Append(")");
                //ViewBag.fund1 = SqlDao.CMEPMSQuery4List(sb.ToString());
                if(plist.Count > 0) {
                    ViewBag.fund1 = SqlDao.CMEPMSQuery4List(sb.ToString());

                    DataTable fund1 = SqlDao.CMEPMSQuery4List(sb.ToString());

                    if(!string.IsNullOrEmpty(Convert.ToString(fund1.Rows[0][0]))) {
                        ViewBag.fund1 = fund1;
                    } else {
                        ViewBag.fund1 = null;
                    }
                }
            }
            if(zero == 2) {
                //經費為0
                sql = sb.ToString();
                sql += "AND YR_FUND=0 ORDER BY TID DESC";
                cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));
                //ViewBag.List2 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List2 = plist;
                sb.Length--;
                sb = new StringBuilder("select sum(YR_FUND)as YR_FUND, " +
                "sum(CEN_FUND)as CEN_FUND, sum(LOC_FUND)as LOC_FUND from " +
                "TDV_TDDCFTS where TID IN(");
                foreach(DataRow d in plist) {
                    sb.Append("'" + d["TID"] + "',");
                }
                sb.Length--;
                sb.Append(")");

                if(plist.Count > 0) {
                    ViewBag.fund2 = SqlDao.CMEPMSQuery4List(sb.ToString());

                    DataTable fund2 = SqlDao.CMEPMSQuery4List(sb.ToString());

                    if(!string.IsNullOrEmpty(Convert.ToString(fund2.Rows[0][0]))) {
                        ViewBag.fund2 = fund2;
                    } else {
                        ViewBag.fund2 = null;
                    }
                }

            }
            if(zero == 3) {
                //保留款
                cmd = new SqlCommand("select * FROM PJV_CRDMNDS where PID = @PID ORDER BY PID DESC");
                cmd.Parameters.Add(new SqlParameter("PID", PID));
                //ViewBag.List3 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List3 = plist;


                sb = new StringBuilder("select sum(YR_FUND)as YR_FUND, " +
                "sum(CEN_FUND)as CEN_FUND, sum(LOC_FUND)as LOC_FUND from " +
                "TDV_TDDCFTS where TID IN(");
                foreach(DataRow d in plist) {
                    sb.Append("'" + d["TID"] + "',");
                }
                sb.Length--;
                sb.Append(")");

                if(plist.Count > 0) {
                    //ViewBag.fund3 = SqlDao.CMEPMSQuery4List(sb.ToString());

                    DataTable fund3 = SqlDao.CMEPMSQuery4List(sb.ToString());

                    if(!string.IsNullOrEmpty(Convert.ToString(fund3.Rows[0][0]))) {
                        ViewBag.fund3 = fund3;
                    } else {
                        ViewBag.fund3 = null;
                    }
                }

            }
            return m;
        }

        //取某期CTHMNDS
        private PJE0030Model GetPJDFDCR_By(object PID, PJE0030Model m, int page, int zero) {

            StringBuilder sb = new StringBuilder("SELECT * FROM TDV_TDHCFTS WHERE PID = @PID ");
            string sql;

            //分項
            if(!string.IsNullOrEmpty(m.PJ_CATY_CODE)) {
                sb.Append("AND PJ_CATY_CODE='" + m.PJ_CATY_CODE + "'");
            }
            //名稱
            if(!string.IsNullOrEmpty(m.TD_NAME)) {
                sb.Append("AND TD_NAME LIKE'%" + m.TD_NAME + "%'");
            }
            //縣市
            if(!string.IsNullOrEmpty(m.CITY)) {
                sb.Append("AND CITY='" + m.CITY + "'");
            }
            //鄉鎮
            if(!string.IsNullOrEmpty(m.TOWN)) {
                sb.Append("AND TOWN='" + m.TOWN + "'");
            }
            //工作類型
            if(!string.IsNullOrEmpty(m.TD_CATY)) {
                sb.Append("AND TD_CATY='" + m.TD_CATY + "'");
            }
            //執行方式
            if(!string.IsNullOrEmpty(m.EXE_TYPE)) {
                sb.Append("AND EXE_TYPE='" + m.EXE_TYPE + "'");
            }
            //經費層級
            if(!string.IsNullOrEmpty(m.IPJ_CODE)) {
                sb.Append("AND IPJ_CODE='" + m.IPJ_CODE + "'");
            }
            //辦理層級
            if(!string.IsNullOrEmpty(m.N_IPJ_CODE)) {
                sb.Append("AND N_IPJ_CODE='" + m.N_IPJ_CODE + "'");
            }
            //分類
            if(!string.IsNullOrEmpty(m.EMG_CATY)) {
                sb.Append("AND EMG_CATY='" + m.EMG_CATY + "'");
            }
            //會計科目
            if(!string.IsNullOrEmpty(m.ACC_ID)) {
                sb.Append("AND ACC_ID='" + m.ACC_ID + "'");
            }
            //SEQ_NO
            sb.Append("AND SEQ_NO='" + m.SEQ_NO + "'");

            sql = sb.ToString();
            SqlCommand cmd;
            List<DataRow> list;
            DataTable ps;
            IPagedList<DataRow> plist;
            if(zero == 1) {
                sql += "AND(YR_FUND>0 OR YR_FUND IS NULL) ORDER BY TID DESC";
                cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));
                //ViewBag.List1 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List1 = plist;

            }
            if(zero == 2) {
                //經費為0
                sql = sb.ToString();
                sql += "AND YR_FUND=0 ORDER BY TID DESC";
                cmd = new SqlCommand(sql);
                cmd.Parameters.Add(new SqlParameter("PID", PID));
                //ViewBag.List2 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List2 = plist;
            }
            if(zero == 3) {
                //保留款
                cmd = new SqlCommand("select * from PJV_CRDMNDS where PID = @PID ORDER BY PID_SYR DESC, PID DESC");
                cmd.Parameters.Add(new SqlParameter("PID", PID));
                //ViewBag.List3 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
                list = ps.AsEnumerable().ToList();
                plist = new PagedList<DataRow>(list, page, pageSize);
                m.List3 = plist;
            }
            return m;
        }

        //管制 - 查詢
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Search")]
        public ActionResult Search(PJE0030Model model) {
            int max = util.GetSEQ_NO(model.PID);
            model.page = (model.page == 0) ? 1 : model.page;

            if(model.SEQ_NO == max) {
                model = GetPJDFDCR_Max(model.PID, model, 1, 1);
            } else {
                model = GetPJDFDCR_By(model.PID, model, 1, 1);
            }
            GetSelectList(model.PID);
            return View("PJE0050ManageInfo", model);
        }

        
        //管制 - CTDMNDS進入版本frame
        public ActionResult EditMnd(decimal PID, decimal TID, string Lock) {
            if(!string.IsNullOrEmpty(Lock)) {
                ViewBag.LOCK = Lock;
            }
            PJE0030Model model = new PJE0030Model();
            int max = util.GetSEQ_NO(PID);
            //最新版本
            CTDMNDS m = db_pms.CTDMNDS.Where(f => f.TID == TID).FirstOrDefault();
            TDV_TDDCFTS t = db_pms.TDV_TDDCFTS.Where(f => f.TID == TID).FirstOrDefault();
            model.TDV_TDHCFTS = new TDV_TDHCFTS();


            //選單            
            GetSelectList(PID);
            model.TID = TID;
            model.PID = PID;
            //經費
            ViewBag.CTDITFD = BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT*FROM CTDITFD c WHERE c.TID=" + TID + " ORDER BY c.FID DESC");
            ViewBag.FIT_CODES = BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT CATY_COD as id, CODE_DESC as name  FROM SMCCODE c WHERE c.SYS_ID='FD' AND CATY_ID='05'");
            ViewBag.CTDUPRC = BASE_DAO.SqlDao.CMEPMSQuery4List("SELECT*FROM CTDUPRC c WHERE c.TID=" + TID);
            if(t == null) {
                //return Content("<xml>資料對應有誤,請刪除並重新建立</xml>", "text/xml");
                return View("PJE0030ManageEdit", model);
            }

            try {
                model.TDV_TDHCFTS.CEN_SP_RATIO_S = t.CEN_SP_RATIO_S;
            } catch {
                model.TDV_TDHCFTS.CEN_SP_RATIO_S = 0;
            }

            try {
                model.TDV_TDHCFTS.CEN_SP_RATIO = t.CEN_SP_RATIO;
            } catch {
                model.TDV_TDHCFTS.CEN_SP_RATIO = 0;
            }

            model.TDV_TDHCFTS.TID = TID;
            model.TDV_TDHCFTS.PID = t.PID;
            model.TDV_TDHCFTS.SIP_TD_CODE = t.SIP_TD_CODE;
            model.TDV_TDHCFTS.TD_CODE = t.TD_CODE;
            model.TDV_TDHCFTS.FD_CODE = t.FD_CODE;
            model.TDV_TDHCFTS.TD_NAME = t.TD_NAME;
            model.TDV_TDHCFTS.TD_CATY = t.TD_CATY;
            model.TDV_TDHCFTS.IPJ_CODE_DESC = t.IPJ_CODE_DESC;
            model.TDV_TDHCFTS.YR_FUND = t.YR_FUND;
            model.TDV_TDHCFTS.ACC_DESC = t.ACC_DESC;
            model.TDV_TDHCFTS.CATY_CODE = t.CATY_CODE;
            model.TDV_TDHCFTS.CEN_SP_RATIO = t.CEN_SP_RATIO;
            model.TDV_TDHCFTS.CEN_SP_RATIO_S = t.CEN_SP_RATIO_S;
            //model.TDV_TDHCFTS.CEN_SP_RATIO_U = t.CEN_SP_RATIO_U;
            model.TDV_TDHCFTS.NOCTRL_FG = t.NOCTRL_FG;
            model.TDV_TDHCFTS.FCAL_FG = t.FCAL_FG;
            model.TDV_TDHCFTS.UCNT_FG = t.UCNT_FG;
            //model.TDV_TDHCFTS.WORKTYPE = t.WORKTYPE;
            model.TDV_TDHCFTS.IPJ_CODE = t.IPJ_CODE;
            model.TDV_TDHCFTS.EXE_TYPE = t.EXE_TYPE;
            model.TDV_TDHCFTS.EMG_CATY = t.EMG_CATY;
            model.TDV_TDHCFTS.ADV_DP = t.ADV_DP;
            if(!string.IsNullOrEmpty(model.TDV_TDHCFTS.ADV_UID))
                model.TDV_TDHCFTS.ADV_UID = t.ADV_UID;
            model.TDV_TDHCFTS.UPD_DT = t.UPD_DT;
            model.TDV_TDHCFTS.UPD_UID = t.UPD_UID;
            model.TDV_TDHCFTS.PD_FG = t.PD_FG;
            model.TDV_TDHCFTS.CITY = t.CITY;
            model.TDV_TDHCFTS.TOWN = t.TOWN;
            model.TDV_TDHCFTS.M_01 = t.M_01;
            model.TDV_TDHCFTS.M_02 = t.M_02;
            model.TDV_TDHCFTS.M_03 = t.M_03;
            model.TDV_TDHCFTS.M_04 = t.M_04;
            model.TDV_TDHCFTS.M_05 = t.M_05;
            model.TDV_TDHCFTS.M_06 = t.M_06;
            model.TDV_TDHCFTS.M_07 = t.M_07;
            model.TDV_TDHCFTS.M_08 = t.M_08;
            model.TDV_TDHCFTS.M_09 = t.M_09;
            model.TDV_TDHCFTS.M_10 = t.M_10;
            model.TDV_TDHCFTS.M_11 = t.M_11;
            model.TDV_TDHCFTS.M_12 = t.M_12;
            model.TDV_TDHCFTS.TD_GROUP = t.TD_GROUP;
            model.TDV_TDHCFTS.B_GROUP = t.B_GROUP;
            model.TDV_TDHCFTS.F_GROUP = t.F_GROUP;
            model.TDV_TDHCFTS.PJ_CATY_CODE = t.PJ_CATY_CODE;
            model.TDV_TDHCFTS.TD_CATY = t.TD_CATY;
            model.TDV_TDHCFTS.YR_WORK = t.YR_WORK;
            return View("PJE0050ManageEdit", model);
        }

        //預算統計
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BudgetCount")]
        public ActionResult BudgetCount(PJE0030Model model) {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            if(model.SEQ_NO == 0) {
                model.SEQ_NO = util.GetSEQ_NO(model.PID);
            }


            GetSelectList(model.PID);

            DataTable xline = SqlDao.CMEPMSQuery4List("select CATY_CODE, CATY_NAME from PJDCATY where PID=" + model.PID);
            DataTable yline = SqlDao.CMEPMSQuery4List("select  CITY, count(*)as cnt from TDV_TDDCFTS where PID=" + model.PID + " group by CITY order by CITY");
            yline.Columns.Remove("cnt");

            foreach(DataRow x in xline.Rows) {
                //yline.Columns.Add(x["CATY_NAME"]+"_NAME");//名
                yline.Columns.Add(x["CATY_CODE"] + "_QUE");//數量
                yline.Columns.Add(x["CATY_CODE"] + "_AMO");//金額
            }


            DataTable myTDDCFTS = SqlDao.CMEPMSQuery4List("select  CATY_NAME, CITY, COUNT(*)as QUE,  FORMAT(SUM(YR_FUND),'###,###.##')as AMO from TDV_TDDCFTS where PID='" +
                model.PID + "' GROUP BY CITY, CATY_NAME ORDER BY CATY_NAME, city");

            foreach(DataRow y in yline.Rows) {
                foreach(DataRow x in xline.Rows) {
                    try {
                        //y[x["CATY_NAME"] + "_NAME"] = x["CATY_NAME"];

                        foreach(DataRow m in myTDDCFTS.Rows) {
                            if(Convert.ToString(m["CITY"]).Equals(Convert.ToString(y["CITY"])) &&
                                Convert.ToString(m["CATY_NAME"]).Equals(Convert.ToString(x["CATY_NAME"]))

                                ) {

                                y[x["CATY_CODE"] + "_QUE"] = m["QUE"];
                                y[x["CATY_CODE"] + "_AMO"] = m["AMO"];
                            }
                        }

                    } catch(Exception e) {
                        System.Diagnostics.Debug.WriteLine(e.Message);
                    }
                }
            }
            ViewBag.xline = xline;
            ViewBag.yline = yline;








            return View("PJE0050BudgetCount", model);
        }

        //預算調整
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BudgetEdit")]
        public ActionResult BudgetEdit(PJE0030Model model) {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            //DataTable core = SqlDao.CMEPMSQuery4List("select CATY_CODE, CATY_NAME, CATY_FUND from PJDCATY where PID=" + model.PID);
            //DataTable core = SqlDao.CMEPMSQuery4List("SELECT (SELECT COUNT(*)FROM TDV_TDDCFTS WHERE PID="+model.PID+")as cnt,p.* FROM PJMFORM p WHERE  p.PID=" + model.PID);

            DataTable list;
            try {
                //list = SqlDao.CMEPMSQuery4List("exec SP_PJE0030_PJDFAD " + model.PID + " ,'','01,02,03,04,20,21'");
                DataTable catys = SqlDao.CMEPMSQuery4List("SELECT CATY_CODE FROM PJDCATY WHERE PID=" + model.PID);
                StringBuilder sb = new StringBuilder("exec SP_PJE0030_PJDFAD " + model.PID + " ,'','");
                foreach(DataRow d in catys.Rows) {
                    sb.Append(d["CATY_CODE"] + ",");
                }
                sb.Length--;
                sb.Append("'");
                //list = SqlDao.CMEPMSQuery4List("exec SP_PJE0030_PJDFAD " + model.PID + " ,'','07,08,09,20,21'");
                list = SqlDao.CMEPMSQuery4List(sb.ToString());
            } catch(Exception e) {

                ViewBag.msg = "預存程序SP_PJE0030_PJDFAD: " + e.Message;
                return Index(model);
            }


            DataTable core = list.Select("AN_CODE IS null").CopyToDataTable();


            //core.Columns.Add("citi", typeof(DataTable));
            core.Columns.Add("prj", typeof(DataTable));
            DataTable citi;
            DataTable prj;
            string ancode, code, keycode;
            //citi.Columns.Remove("cnt");
            foreach(DataRow dr in core.Rows) {

                //上層
                //prj = SqlDao.CMEPMSQuery4List("SELECT (SELECT COUNT(*)FROM TDV_TDDCFTS WHERE PID="+model.PID+" AND CATY_CODE=PJDCATY.CATY_CODE)as cnt, CATY_CODE, CATY_NAME, CATY_FUND from PJDCATY where PID=" + model.PID);
                try {
                    prj = list.Select("AN_CODE='" + dr["KEYCODE"] + "'").CopyToDataTable();
                    prj.Columns.Add("citi", typeof(DataTable));
                    dr["prj"] = prj;
                } catch {
                    prj = list.Copy();
                    prj.Columns.Add("citi", typeof(DataTable));
                    dr["prj"] = prj;
                }




                foreach(DataRow p in prj.Rows) {
                    //citi = SqlDao.CMEPMSQuery4List("select TID, COUNT(*)as cnt, SUM(TDV_TDDCFTS.YR_FUND)as fund, CITY from TDV_TDDCFTS where PID=" + model.PID + " and PJ_CATY_CODE=" + p["CATY_CODE"] +" group by CITY, TID");
                    try {
                        keycode = p["KEYCODE"].ToString();
                        citi = list.Select("AN_CODE='" + p["KEYCODE"] + "'AND REC_TP='D'").CopyToDataTable();
                        p["citi"] = citi;
                    } catch(Exception e) {
                        //citi = list.Select("AN_CODE='" + p["CODE"] + "'").CopyToDataTable();
                        p["citi"] = list.Clone();
                        //prj.Rows.Remove(p);
                        System.Diagnostics.Debug.WriteLine(e);
                    }

                }
            }

            ViewBag.core = core;
            return View("PJE0050BudgetEdit", model);
        }

    }

}