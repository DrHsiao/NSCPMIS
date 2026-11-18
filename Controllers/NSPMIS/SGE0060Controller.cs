using Aspose.Cells;
using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMPSGSModels;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS
{
    public class CpcClass
    {
        public class MyFunc
        {
            public double MyCDbl(object obj)
            {
                if (obj == null || obj == DBNull.Value || string.IsNullOrEmpty(obj.ToString()))
                {
                    return 0.0;
                }
                double.TryParse(obj.ToString(), out double result);
                return result;
            }
        }
    }

    //2.2.3
    public class SGE0060Controller : Controller
    {
        private readonly Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();
        private readonly Models.CMPSGSModels.CMPSGSEntities sgsdb = new Models.CMPSGSModels.CMPSGSEntities();
        private readonly Models.CMBASEModels.CMBASEEntities cmbase = new Models.CMBASEModels.CMBASEEntities();

        int pageSize = 10;

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CheckMyProposal")]
        public ActionResult CheckMyProposal(SGE0010Model model)
        {

            //CITY內含不分區權限
            /*
            if(util.IsPrincipal(WebSiteHelper.getUserId())) {
                BASE_DAO.SqlDao.CMPSGSExecuteSql("UPDATE APMFORM SET AP_STAT='00' WHERE PPID='1' AND AP_STAT='-1'");
                return ProposalManage(model);
            }
            */
            //TOWN不分區權限
            DataTable myCities = util.GetMyCities(WebSiteHelper.getUserId());
            //分區權限
            DataTable myTowns = util.GetMyTowns(WebSiteHelper.getUserId());
            StringBuilder sql = new StringBuilder("UPDATE APMFORM SET AP_STAT='00' WHERE AP_STAT='-1' AND PPID= " + model.PPID);
            if (myCities.Rows.Count > 0)
            {
                //加入不分區縣市
                sql.Append("AND (CITY IN(");
                foreach (DataRow d in myCities.Rows)
                {
                    sql.Append("'" + d["CITY"] + "',");
                }
                sql.Length--;
                sql.Append(")");
            }

            if (myTowns.Rows.Count > 0)
            {
                //加入分區鄉鎮
                sql.Append("OR TOWN IN(");
                foreach (DataRow d in myTowns.Rows)
                {
                    sql.Append("'" + d["TOWN"] + "',");
                }
                sql.Length--;
                sql.Append("))");
            }
            BASE_DAO.SqlDao.CMPSGSExecuteSql(sql.ToString());
            ViewBag.msg = "已收件";

            return ProposalManage(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Search")]
        public ActionResult Search(SGE0010Model model)
        {
            string userid = Helpers.WebSiteHelper.getUserId();
            int page1 = (model.page1 == 0) ? 1 : model.page1;
            model.page1 = page1;
            if (model.PJMFORM == null)
            {
                model.PJMFORM = new Models.CMPSGSModels.PJMFORM_CMPSGS();
            }

            SqlCommand cmd;
            if (!string.IsNullOrEmpty(model.keyword))
            {
                cmd = new SqlCommand("select * from PJMFORM where PJ_CODE in " +
            "(select PJ_CODE from PJCSGDP where AP_STAT='15' and DEPT_NO =@DP)AND PJ_NAME LIKE '%' + @KEY + '%'");
                cmd.Parameters.Add(new SqlParameter("KEY", model.keyword));
            }
            else
            {
                cmd = new SqlCommand("select * from PJMFORM where PJ_CODE in " +
            "(select PJ_CODE from PJCSGDP where AP_STAT='15' and DEPT_NO =@DP)");
            }
            cmd.Parameters.Add(new SqlParameter("DP", WebSiteHelper.getUserDept()));

            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page1, 10);
            model.MainList = plist;

            return View("SGE0060Index", model);
        }
        public ActionResult Index(SGE0010Model model)
        {

            //string userid = Helpers.WebSiteHelper.getUserId();
            int page1 = (model.page1 == 0) ? 1 : model.page1;
            model.page1 = page1;
            if (model.PJMFORM == null)
            {
                model.PJMFORM = new Models.CMPSGSModels.PJMFORM_CMPSGS();
            }

            SqlCommand cmd = new SqlCommand("select * from PJMFORM where PJ_CODE in " +
                "(select PJ_CODE from PJCSGDP where AP_STAT='15' and DEPT_NO =@DP)");
            // TODO：AND PPID >= 21 是目前正式環境開放給雨水科測試用，之後要還原成上面這段 // 20251210
            //SqlCommand cmd = new SqlCommand("select * from PJMFORM where PJ_CODE in " +
            //    "(select PJ_CODE from PJCSGDP where AP_STAT='15' and DEPT_NO =@DP) AND PPID >= 21");

            cmd.Parameters.Add(new SqlParameter("DP", WebSiteHelper.getUserDept()));
            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page1, 10);
            model.MainList = plist;


            return View("SGE0060Index", model);
        }


        public ActionResult EditAdvice(decimal aid, string locked)
        {

            //處理R
            APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == apmform.PPID).FirstOrDefault();
            //R的情況
            if (pjmform.STAT_NO.Equals("R"))
            {
                ViewBag.isR = true;
            }

            if (!string.IsNullOrEmpty(locked))
            {
                ViewBag.locked = "*";
            }

            SGE0010Model model = new SGE0010Model();
            //model.AID = aid;
            model.apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            model.PJMFORM = pjmform;
            SqlCommand cmd = new SqlCommand("SELECT*FROM APV_APMFORM WHERE AID=@AID");
            cmd.Parameters.Add(new SqlParameter("AID", aid));
            ViewBag.APV_APMFORM = SqlDao.CMESGSQuery4Object(cmd);
            try
            {
                ViewBag.S2_FN_DT = tools.GetDateByTWDate((DateTime)model.apmform.S2_FN_DT);
            }
            catch
            {
                ViewBag.S2_FN_DT = "";
            }
            try
            {
                ViewBag.S2_OF_DT = tools.GetDateByTWDate((DateTime)model.apmform.S2_OF_DT);
            }
            catch
            {
                ViewBag.S2_OF_DT = "";
            }


            cmd = new SqlCommand("SELECT*FROM APDFUND WHERE AID=@AID");
            cmd.Parameters.Add(new SqlParameter("AID", aid));
            ViewBag.apdfund = SqlDao.CMPSGSQuery4List(cmd);
            //辦理層級
            ViewBag.IPJ_CODES = tools.ToSelectList(util.GetCMETMSCODES("SM", "02"), "id", "name", null);
            //預算分類
            ViewBag.BUDGET_TP = tools.ToSelectList(util.GetCMPSGSCODES("TP", "01"), "id", "name", null);


            return View("EditAdvice", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SaveAdvice")]
        public ActionResult SaveAdvice(SGE0010Model model)
        {
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();


            a.S2_FN_NO = model.apmform.S2_FN_NO;
            a.S2_OF_NO = model.apmform.S2_OF_NO;
            a.S2_COMT = model.apmform.S2_COMT;
            a.UPD_DT = DateTime.Now;
            a.UPD_UID = WebSiteHelper.getUserId();

            try
            {
                a.S2_FN_DT = tools.GetDateByTWDate(model.FN_DT);
            }
            catch
            {

            }
            try
            {
                a.S2_OF_DT = tools.GetDateByTWDate(model.OF_DT);
            }
            catch
            {

            }


            //審查評比
            a.S1_ORDER = null;
            a.SC1_SCORE = null;
            a.SC2_SCORE = null;
            a.SC3_SCORE = null;
            a.SC4_SCORE = null;
            a.SC5_SCORE = null;
            a.SC6_SCORE = null;
            a.SC7_SCORE = null;
            a.PRV_CON_SEWERS = null;
            a.CHK_STAT = null;
            a.IPJ_CODE = null;
            a.BUDGET_TP = null;

            //notR
            a.SC8_SCORE = null;
            a.SC1_CMT = null;
            a.SC2_CMT = null;
            a.SC3_CMT = null;
            a.SC4_CMT = null;
            a.SC5_CMT = null;
            a.SC6_CMT = null;
            a.SC7_CMT = null;
            a.SC8_CMT = null;
            a.SCS_CMT = null;

            if (model.apmform.PRV_CON_SEWERS != null) a.PRV_CON_SEWERS = model.apmform.PRV_CON_SEWERS;
            if (model.apmform.S1_ORDER != null) a.S1_ORDER = model.apmform.S1_ORDER;
            if (model.apmform.SC1_SCORE != null) a.SC1_SCORE = model.apmform.SC1_SCORE;
            if (model.apmform.SC2_SCORE != null) a.SC2_SCORE = model.apmform.SC2_SCORE;
            if (model.apmform.SC3_SCORE != null) a.SC3_SCORE = model.apmform.SC3_SCORE;
            if (model.apmform.SC4_SCORE != null) a.SC4_SCORE = model.apmform.SC4_SCORE;
            if (model.apmform.SC5_SCORE != null) a.SC5_SCORE = model.apmform.SC5_SCORE;
            if (model.apmform.SC6_SCORE != null) a.SC6_SCORE = model.apmform.SC6_SCORE;
            if (model.apmform.SC7_SCORE != null) a.SC7_SCORE = model.apmform.SC7_SCORE;
            if (model.apmform.SC8_SCORE != null) a.SC8_SCORE = model.apmform.SC8_SCORE;

            if (model.apmform.CHK_STAT != null) a.CHK_STAT = model.apmform.CHK_STAT;
            if (model.apmform.IPJ_CODE != null) a.IPJ_CODE = model.apmform.IPJ_CODE;
            if (model.apmform.BUDGET_TP != null) a.BUDGET_TP = model.apmform.BUDGET_TP;

            //notR

            if (model.apmform.SC1_CMT != null) a.SC1_CMT = model.apmform.SC1_CMT;
            if (model.apmform.SC2_CMT != null) a.SC2_CMT = model.apmform.SC2_CMT;
            if (model.apmform.SC3_CMT != null) a.SC3_CMT = model.apmform.SC3_CMT;
            if (model.apmform.SC4_CMT != null) a.SC4_CMT = model.apmform.SC4_CMT;
            if (model.apmform.SC5_CMT != null) a.SC5_CMT = model.apmform.SC5_CMT;
            if (model.apmform.SC6_CMT != null) a.SC6_CMT = model.apmform.SC6_CMT;
            if (model.apmform.SC7_CMT != null) a.SC7_CMT = model.apmform.SC7_CMT;
            if (model.apmform.SC8_CMT != null) a.SC8_CMT = model.apmform.SC8_CMT;
            if (model.apmform.SCS_CMT != null) a.SCS_CMT = model.apmform.SCS_CMT;
            sgsdb.SaveChanges();
            string cmd;
            if (model.FID != null)
            {
                for (int i = 0; i < model.FID.Length; i++)
                {
                    string descValue = string.IsNullOrWhiteSpace(model.FUND_DESC[i]) ? "NULL" : $"'{model.FUND_DESC[i].Replace("'", "''")}'";
                    bool isValid = decimal.TryParse(model.DCS_FUND[i]?.Replace(",", "").Trim(), out decimal fundValue);
                    cmd = $"UPDATE APDFUND SET FUND_DESC={descValue}, DCS_FUND={(isValid ? fundValue.ToString("0") : "NULL")} WHERE FID={model.FID[i]}";
                    SqlDao.CMPSGSExecuteSql(cmd);
                }
            }
            ViewBag.msg = "已儲存";
            return EditAdvice(aid, null);
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ProposalSearch")]
        public ActionResult ProposalSearch(SGE0010Model model)
        {
            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
            int page1 = (model.page1 == 0) ? 1 : model.page1;
            int page2 = 1;// (model.page2 == 0) ? 1 : model.page2;
            DataTable yrcs = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC WHERE PPID=" + model.PPID + " ORDER BY YRC_ID DESC");
            ViewBag.yrcs = tools.ToSelectList(yrcs, "YRC_ID", "PSG_NAME", null);
            model.tab = 1;
            DataTable ps = util.GetAPMFORMByEzway(model.PPID, "00", model.YRC_ID);
            List<DataRow> list = ps.AsEnumerable().ToList();

            if (model.city != null)
            {
                list = ps.AsEnumerable().Where(r => string.IsNullOrEmpty(model.city) || r.Field<string>("CITY") == model.city).ToList();
            }

            // 【新增】工程名稱 (AP_NAME) 的模糊查詢過濾
            if (!string.IsNullOrEmpty(model.keyword2))
            {
                // 確保欄位不是 Null，並且包含使用者輸入的關鍵字
                list = list.Where(r => r.Field<string>("AP_NAME") != null &&
                                       r.Field<string>("AP_NAME").Contains(model.keyword2)).ToList();
            }

            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page2, pageSize);
            ViewBag.noTab = true;
            model.List1 = plist;
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == model.PPID).FirstOrDefault();
            model.PJMFORM = pjmform;
            ViewBag.PJ_NAME = pjmform?.PJ_NAME ?? "";
            return View("SGE0060ProposalManage", model);
            //return ProposalManage(model);
        }

        public ActionResult ProposalManage(int tab, string ppid, int page1, int page2, string yrcid, string city = null, string keyword2 = null)
        {
            string userid = WebSiteHelper.getUserId();
            string dept = WebSiteHelper.getUserDept();

            decimal id = Convert.ToDecimal(ppid);
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == id).FirstOrDefault();
            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
            PJDSYRC pjdsyrc = sgsdb.PJDSYRC.Where(w => w.PPID == id).FirstOrDefault();
            ViewBag.PJ_NAME = pjmform.PJ_NAME;
            ViewBag.PSG_NAME = pjdsyrc.PSG_NAME;


            SGE0010Model model = new SGE0010Model();
            model.PJMFORM = pjmform;
            model.AP_STAT = pjmform.STAT_NO;
            model.page1 = (page1 == 0) ? 1 : page1;
            model.page2 = (page2 == 0) ? 1 : page2; ;
            model.tab = tab;
            model.YRC_ID = yrcid;
            model.city = city;
            model.keyword2 = keyword2;
            model.PPID = tools.safe2Decimal(ppid);
            ViewBag.yrcs = tools.ToSelectList(SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC WHERE PPID=" + model.PPID + " ORDER BY YRC_ID DESC"), "YRC_ID", "PSG_NAME", null);

            DataTable ps = null;
            List<DataRow> list;
            IPagedList<DataRow> plist;
            //2.2.3 鎖到部門
            //Z1: 不使用 GetAPMFORMByDeptSP
            //Z2: GetAPMFORMByDeptSP(AP_STAT = 'Z2', NAP_DTAT = '15')
            //ZZ: GetAPMFORMByDeptSP(AP_STAT = 'ZZ', NAP_DTAT = '15')

            //待收件00
            if (tab == 1)
            {
                ps = util.GetAPMFORMByEzway(ppid, "00", yrcid);
            }
            //初審中10
            if (tab == 2)
            {
                ps = util.GetAPMFORMByEzway(ppid, "10", yrcid);
            }
            //複審15
            if (tab == 3)
            {
                ps = util.GetAPMFORMByEzway(model.PPID, "15", yrcid);
                //ps.Merge(util.GetAPMFORM_Attach(ppid, "15", userid, dept));//追加本單位
            }
            //核定20
            if (tab == 4)
            {
                ps = util.GetAPMFORMByEzway(model.PPID, "20", yrcid);
                //ps.Merge(util.GetAPMFORM_Attach(ppid, "20", userid, dept));//追加本單位
            }
            //成案ZZ
            if (tab == 5)
            {
                ps = util.GetAPMFORMByEzway(model.PPID, "90", yrcid);
            }
            //未通過初審
            if (tab == 6)
            {
                ps = util.GetAPMFORMByEzway(model.PPID, "Z1", yrcid);
                //ps.Merge(util.GetAPMFORM_Attach(ppid, "Z1", userid, dept));//追加本單位
            }
            //未通過複審
            if (tab == 7)
            {
                ps = util.GetAPMFORMByEzway(model.PPID, "Z2", yrcid);
                //ps.Merge(util.GetAPMFORM_Attach(ppid, "Z2", userid, dept));//追加本單位
            }

            // 【修正】移除重覆寫法
            list = ps?.AsEnumerable().ToList() ?? new List<DataRow>();
            if (!string.IsNullOrEmpty(city))
            {
                list = list.Where(r => r.Field<string>("CITY") == city).ToList();
            }
            if (string.IsNullOrEmpty(keyword2) == false)
            {
                list = list.Where(r => r.Field<string>("AP_NAME") != null &&
                                       r.Field<string>("AP_NAME").Contains(keyword2)).ToList();
            }
            plist = new PagedList<DataRow>(list, page2, pageSize);
            model.List1 = plist;

            return View("SGE0060ProposalManage", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ProposalManage")]
        public ActionResult ProposalManage(SGE0010Model model)
        {
            int page2 = (model.page == 0) ? 1 : model.page2;
            DataTable yrcs = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC WHERE PPID=" + model.PPID + " ORDER BY YRC_ID DESC");
            ViewBag.yrcs = tools.ToSelectList(yrcs, "YRC_ID", "PSG_NAME", null);
            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
            List<DataRow> list;
            IPagedList<DataRow> plist;
            string yrcid = Convert.ToString(yrcs.Rows[0]["YRC_ID"]);
            DataTable ps = util.GetAPMFORMByEzway(model.PPID, "00", yrcid);
            model.tab = 1;
            model.YRC_ID = yrcid;

            list = ps.AsEnumerable().ToList();
            plist = new PagedList<DataRow>(list, page2, pageSize);
            model.List1 = plist;


            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == model.PPID).FirstOrDefault();
            PJDSYRC pjdsyrc = sgsdb.PJDSYRC.Where(w => w.PPID == model.PPID).FirstOrDefault();
            ViewBag.PJ_NAME = pjmform.PJ_NAME;
            ViewBag.PSG_NAME = pjdsyrc.PSG_NAME;
            model.PJMFORM = pjmform;
            model.AP_STAT = pjmform.STAT_NO;
            return View("SGE0060ProposalManage", model);
        }

        //送件 (複審中->核定) 15 -> 20 ST
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendForward")]
        public ActionResult SendForward(SGE0010Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            // AP_STAT = 15(複審中) 才能送件
            if (a.AP_STAT != "15")
            {
                ViewBag.msg = "狀態已變更，請勿重複操作";
                return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
            }
            string o_status = a.AP_STAT;
            a.AP_STAT = code.AP_STAT_Forward(a.AP_STAT);
            sgsdb.SaveChanges();


            util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "ST", now, userid, WebSiteHelper.getUserDept(), "");
            ViewBag.msg = "已送件";
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        // 批次送件複審  15 -> 20(或下一個狀態) ST
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendForwardBatch")]
        public ActionResult SendForwardBatch(SGE0010Model model) // 參數型別維持與您提供的一致
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            string deptid = WebSiteHelper.getUserDept(); // 將取得部門 ID 移到迴圈外，節省效能

            // 防呆檢查：確認前端確實有傳入 AID 字串
            if (string.IsNullOrEmpty(model.AID))
            {
                ViewBag.msg = "未選擇任何要送件的資料";
                return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
            }

            // 將前端傳來逗號分隔的字串切成陣列 (例如 "1001,1002" -> ["1001", "1002"])
            string[] aidArray = model.AID.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            int successCount = 0; // 用來記錄成功處理的筆數

            foreach (string aidStr in aidArray)
            {
                decimal aid = tools.safe2Decimal(aidStr);
                APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();

                // 確保該筆資料存在，且 AP_STAT == "15" (複審中) 才能送件
                if (a != null && a.AP_STAT == "15")
                {
                    string o_status = a.AP_STAT;
                    a.AP_STAT = code.AP_STAT_Forward(a.AP_STAT);

                    // 寫入 Log (注意這裡的 deptid 已經在迴圈外宣告好了)
                    util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "ST", now, userid, deptid, "");

                    successCount++;
                }
            }

            // 檢查是否有成功處理的資料，統一執行一次資料庫儲存 (大幅提升效能)
            if (successCount > 0)
            {
                sgsdb.SaveChanges();
                ViewBag.msg = $"已成功批次送件 {successCount} 筆";
            }
            else
            {
                ViewBag.msg = "勾選的資料狀態皆已變更，無符合送件條件的項目，請重新查詢";
            }

            // 返回原本的列表頁面，並帶入原本的查詢參數
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        //退件 (複審中->初審中) 15 -> 10 BK
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendBackward")]
        public ActionResult SendBackward(SGE0010Model model, string PRC_NOTE)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);

            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            // AP_STAT = 15(複審中) 才能退件
            if (a.AP_STAT != "15")
            {
                ViewBag.msg = "狀態已變更，請勿重複操作";
                return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
            }
            string o_status = a.AP_STAT;
            a.AP_STAT = code.AP_STAT_Backward(a.AP_STAT);

            sgsdb.SaveChanges();

            util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "BK", now, userid, WebSiteHelper.getUserDept(), PRC_NOTE);
            ViewBag.msg = "已退件";
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        // 💡 新增：複審批次退件 (多筆退回初審) 15 -> 10 BK
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendBackwardBatch")]
        public ActionResult SendBackwardBatch(SGE0010Model model, string PRC_NOTE)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();

            // 💡 1. 取得前端傳來的多筆 AID 字串 (例如 "1001,1002,1003")
            // 使用 Request.Form 接值，避免 model.AID 如果是數字型態會產生綁定錯誤
            string aidStr = Request.Form["AID"];

            if (!string.IsNullOrEmpty(aidStr))
            {
                string[] aidList = aidStr.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                int successCount = 0;

                // 💡 2. 跑迴圈，一筆一筆處理退件邏輯
                foreach (var idStr in aidList)
                {
                    decimal aid = Convert.ToDecimal(idStr.Trim());
                    APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();

                    // 必須存在，且 AP_STAT = 15(複審中) 才能退件
                    if (a != null && a.AP_STAT == "15")
                    {
                        string o_status = a.AP_STAT;
                        a.AP_STAT = code.AP_STAT_Backward(a.AP_STAT);

                        // 寫入歷程紀錄 (帶入委員填寫的批次退件原因 PRC_NOTE)
                        util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "BK", now, userid, WebSiteHelper.getUserDept(), PRC_NOTE);

                        successCount++;
                    }
                }

                // 💡 3. 一次性將所有更新儲存進資料庫
                sgsdb.SaveChanges();

                ViewBag.msg = $"已成功批次退件 {successCount} 筆";
            }
            else
            {
                ViewBag.msg = "未接收到任何退件資料";
            }

            // 💡 4. 返回原本的列表畫面
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        //未過複審 (複審中->未過複審) 15 -> Z2 ST
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendDroppend")]
        public ActionResult SendDroppend(SGE0010Model model, string PRC_NOTE)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            // AP_STAT = 15(複審中) 才能送件
            if (a.AP_STAT != "15")
            {
                ViewBag.msg = "狀態已變更，請勿重複操作";
                return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
            }
            string o_status = a.AP_STAT;
            a.AP_STAT = "Z2";
            sgsdb.SaveChanges();
            ViewBag.msg = "未過複審";

            util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "ST", now, userid, WebSiteHelper.getUserDept(), PRC_NOTE);
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        //複審中 (未過複審->複審中) Z2 -> 15 BK
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendPrevious")]
        public ActionResult SendPrevious(SGE0010Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            // AP_STAT = Z2(未過複審) 才能退件
            if (a.AP_STAT != "Z2")
            {
                ViewBag.msg = "狀態已變更，請勿重複操作";
                return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
            }
            string o_status = a.AP_STAT;
            a.AP_STAT = "15";
            sgsdb.SaveChanges();
            ViewBag.msg = "已恢復為複審中";

            util.UPDATE_APMFORM_Log(a.AID, o_status, "15", "BK", now, userid, WebSiteHelper.getUserDept(), "恢復複審中");
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        //恢復初審中 (未過複審->初審中) Z2 -> 10 BK
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendBeforePrevious")]
        public ActionResult SendBeforePrevious(SGE0010Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            // AP_STAT = Z2(未過複審) 才能退件
            if (a.AP_STAT != "Z2")
            {
                ViewBag.msg = "狀態已變更，請勿重複操作";
                return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
            }
            string o_status = a.AP_STAT;
            a.AP_STAT = "10";
            sgsdb.SaveChanges();
            ViewBag.msg = "已恢復為初審中";

            util.UPDATE_APMFORM_Log(a.AID, o_status, "10", "BK", now, userid, WebSiteHelper.getUserDept(), "恢復初審中");
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        //列印
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Print")]
        public ActionResult Print(SGE0010Model model)
        {

            // 載入模板
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/PJE0030.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            Worksheet sheet = workbook.Worksheets[0];

            //做預存程序
            DataTable dt;
            using (Models.CMEPMSModels.CMEPMSEntities cmepms = new Models.CMEPMSModels.CMEPMSEntities())
            using (SqlConnection conn = new SqlConnection(cmepms.Database.Connection.ConnectionString))
            using (SqlCommand cmd = new SqlCommand("SP_RCD0010_UDF003_3", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                // 輸入參數
                cmd.Parameters.AddWithValue("@PID", model.PID);
                cmd.Parameters.AddWithValue("@SEQ_NO", model.SEQ_NO);

                dt = new DataTable();
                conn.Open();

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    dt.Load(reader);  // 將結果讀入 DataTable
                }


            }
            int startRow = 2;
            foreach (DataRow d in dt.Rows)
            {
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
            string fileName = $"PJE0030_{timestamp}.xlsx";
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                ms.Position = 0;
                return File(ms.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName);
            }
        }


        private string GetYrc(string yrc_id)
        {
            return SqlDao.CMPSGSQuery4Object($"SELECT * FROM PJDSYRC WHERE YRC_ID = '{yrc_id}'")?["PSG_NAME"]?.ToString() ?? "";
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "print3")]
        public ActionResult print3(SGE0010Model m)
        {
            string bodyType = "15";//鎖定複審中:15
            decimal PPID = m.PPID;
            string YRC_ID = m.YRC_ID;
            string sCity = m.city;

            // 初始化 SqlCommand 與 基礎 Where 條件
            SqlCommand cmd = new SqlCommand();
            string sApWhere = "PPID = @PPID And YRC_ID = @YRC_ID ";
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            // 處理直轄市及縣市合併問題 (參數化)
            if (!string.IsNullOrEmpty(sCity))
            {
                switch (sCity)
                {
                    case "新北市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "新北市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "台北縣"));
                        break;
                    case "台中市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "台中市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "台中縣"));
                        break;
                    case "台南市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "台南市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "台南縣"));
                        break;
                    case "高雄市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "高雄市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "高雄縣"));
                        break;
                    case "桃園市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "桃園市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "桃園縣"));
                        break;
                    default:
                        sApWhere += "And CITY = @City1 ";
                        cmd.Parameters.Add(new SqlParameter("City1", sCity));
                        break;
                }
            }
            // 【新增】工程名稱模糊查詢 (參數化)
            if (!string.IsNullOrEmpty(m.keyword2))
            {
                sApWhere += "And (AP_NAME Like @Keyword) ";
                cmd.Parameters.Add(new SqlParameter("Keyword", "%" + m.keyword2 + "%"));
            }

            string sApOrder = "";
            string sApFundField = "";
            switch (bodyType)
            {
                case "15": // 複審中
                    sApWhere += "And AP_STAT='15' "; sApOrder = "Order By CITY,S1_ORDER,AP_NAME";
                    //sApFund = ",IsNull(FUND,0)/1000 As FUND,IsNull(FUND,0)/1000 As DCS_FUND,IsNull(CEN_FUND,0)/1000 As CEN_DCS_FUND,IsNull(LOC_FUND,0)/1000 As LOC_DCS_FUND"; //全年度
                    sApFundField += ",SC1_SCORE As SC1,SC2_SCORE As SC2,SC3_SCORE As SC3,SC4_SCORE As SC4,SC5_SCORE As SC5,SC6_SCORE As SC6,SC7_SCORE As SC7";
                    sApFundField += ",M_FUND,M_CEN_FUND,M_LOC_FUND";
                    sApFundField += ",M_JUD_FUND As S_FUND,M_CEN_JUD_FUND As S_CEN_FUND,M_LOC_JUD_FUND As S_LOC_FUND,S1_ORDER As S_ORDER";
                    break;
                case "20": // 核定中
                    sApWhere += "And AP_STAT='20' "; sApOrder = "ORDER BY S1_ORDER, AP_NAME"; //S2_ORDER
                    //sApFund = ",IsNull(FUND,0)/1000 As FUND,IsNull(DCS_FUND,0)/1000 As DCS_FUND,IsNull(CEN_DCS_FUND,0)/1000 As CEN_DCS_FUND,IsNull(LOC_DCS_FUND,0)/1000 As LOC_DCS_FUND"; //全年度
                    sApFundField += ",SC1_SCORE As SC1,SC2_SCORE As SC2,SC3_SCORE As SC3,SC4_SCORE As SC4,SC5_SCORE As SC5,SC6_SCORE As SC6,SC7_SCORE As SC7";
                    //sApFundField += ",SC1_SCORE1 As SC1,SC2_SCORE1 As SC2,SC3_SCORE1 As SC3,SC4_SCORE1 As SC4,SC5_SCORE1 As SC5,SC6_SCORE1 As SC6,SC7_SCORE1 As SC7";
                    sApFundField += ",M_FUND,M_CEN_FUND,M_LOC_FUND";
                    sApFundField += ",M_DCS_FUND As S_FUND,M_CEN_DCS_FUND As S_CEN_FUND,M_LOC_DCS_FUND As S_LOC_FUND,S1_ORDER As S_ORDER"; //S2_ORDER
                    break;
                case "Z2": // 未過複審
                    sApWhere += "And AP_STAT='Z2' "; sApOrder = "Order By CITY,S1_ORDER,AP_NAME";
                    //sApFund = ",IsNull(FUND,0)/1000 As FUND,IsNull(FUND,0)/1000 As DCS_FUND,IsNull(CEN_FUND,0)/1000 As CEN_DCS_FUND,IsNull(LOC_FUND,0)/1000 As LOC_DCS_FUND"; //全年度
                    sApFundField += ",SC1_SCORE As SC1,SC2_SCORE As SC2,SC3_SCORE As SC3,SC4_SCORE As SC4,SC5_SCORE As SC5,SC6_SCORE As SC6,SC7_SCORE As SC7";
                    //sApFundField += ",SC1_SCORE1 As SC1,SC2_SCORE1 As SC2,SC3_SCORE1 As SC3,SC4_SCORE1 As SC4,SC5_SCORE1 As SC5,SC6_SCORE1 As SC6,SC7_SCORE1 As SC7";
                    sApFundField += ",M_FUND,M_CEN_FUND,M_LOC_FUND";
                    sApFundField += ",M_DCS_FUND As S_FUND,M_CEN_DCS_FUND As S_CEN_FUND,M_LOC_DCS_FUND As S_LOC_FUND,S1_ORDER As S_ORDER"; //S2_ORDER
                    break;
                case "90": // 已成案
                    sApWhere += "And AP_STAT='90' "; sApOrder = "Order By CITY,AP_NAME";
                    //sApFund = ",IsNull(FUND,0)/1000 As FUND,IsNull(DCS_FUND,0)/1000 As DCS_FUND,IsNull(CEN_DCS_FUND,0)/1000 As CEN_DCS_FUND,IsNull(LOC_DCS_FUND,0)/1000 As LOC_DCS_FUND"; //全年度
                    sApFundField += ",SC1_SCORE As SC1,SC2_SCORE As SC2,SC3_SCORE As SC3,SC4_SCORE As SC4,SC5_SCORE As SC5,SC6_SCORE As SC6,SC7_SCORE As SC7";
                    //sApFundField += ",SC1_SCORE1 As SC1,SC2_SCORE1 As SC2,SC3_SCORE1 As SC3,SC4_SCORE1 As SC4,SC5_SCORE1 As SC5,SC6_SCORE1 As SC6,SC7_SCORE1 As SC7";
                    sApFundField += ",M_FUND,M_CEN_FUND,M_LOC_FUND";
                    sApFundField += ",M_DCS_FUND As S_FUND,M_CEN_DCS_FUND As S_CEN_FUND,M_LOC_DCS_FUND As S_LOC_FUND,S1_ORDER As S_ORDER"; //S2_ORDER
                    break;
            }

            cmd.CommandText = "select CITY,SYR,AID,TD_CODE,AP_CODE,AP_NAME,AP_PJ_NAME,AP_ALL_NOTE,AP_YEAR_NOTE,AP_NOTE,M_T_FUND As T_FUND,M_R_FUND As R_FUND,CHK_STAT,BUDGET_TP,PRV_CON_SEWERS, BUDGET_TP_N" + sApFundField + ",EXT_FG from APV_APMFORM a where " + sApWhere + sApOrder;
            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            DataView dvAPM = dt.DefaultView;

            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == m.PPID).FirstOrDefault();
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0060/SGE0035_NEW.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            Worksheet sheet = workbook.Worksheets[0];

            // 標題
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
            var yrcInfo = GetYrcInfo(YRC_ID);
            string sExcelTitle = sCity + yrcInfo?.BGN_SYR + "年度" + pjmform?.PJ_NAME + "評分表"; //○○縣(市)○○年度污水下水道建設計畫評分表

            ExportSheet.SewageScore sewageScore = new ExportSheet.SewageScore();
            sheet = sewageScore.Sge0035MyPrint1_NEW(sheet, dvAPM, dt);

            sheet.Cells[1, 3].Value = sExcelTitle;
            //sheet.Cells[1, 3].Value = sCity + yrcName + "計畫評分表";

            //sheet.AutoFitColumns();
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                byte[] fileBytes = ms.ToArray();

                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"{sExcelTitle}_{timestamp}.xlsx");
            }
        }

        /// <summary>
        /// 取得中長程計畫-經費管制資料
        /// </summary>
        /// <param name="yrcId"> 年度編號 </param>
        /// <returns></returns>
        private PJDSYRC GetYrcInfo(string yrcId)
        {
            if (!decimal.TryParse(yrcId, out decimal yrcIdDecimal))
            {
                // 無法轉 decimal，直接回傳空白
                return null;
            }

            return sgsdb.PJDSYRC.SingleOrDefault(x => x.YRC_ID == yrcIdDecimal);
        }


        private string GetStat(string sChkStat)
        {
            sChkStat = sChkStat.ToLower();
            switch (sChkStat)
            {
                case "y": sChkStat = "是"; break;
                case "1": sChkStat = "是"; break;
                case "true": sChkStat = "是"; break;
                case "n": sChkStat = "否"; break;
                case "0": sChkStat = "否"; break;
                case "false": sChkStat = "否"; break;
            }
            return sChkStat;
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "print4")]
        public ActionResult print4(SGE0010Model m)
        {
            string bodyType = "1";
            decimal PPID = m.PPID;
            string YRC_ID = m.YRC_ID;
            string sCity = m.city;

            // 初始化 SqlCommand 與 基礎 Where 條件
            SqlCommand cmd = new SqlCommand();
            string sApWhere = "PPID = @PPID And YRC_ID = @YRC_ID ";
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            // 處理直轄市及縣市合併問題 (參數化)
            if (!string.IsNullOrEmpty(sCity))
            {
                switch (sCity)
                {
                    case "新北市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "新北市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "台北縣"));
                        break;
                    case "台中市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "台中市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "台中縣"));
                        break;
                    case "台南市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "台南市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "台南縣"));
                        break;
                    case "高雄市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "高雄市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "高雄縣"));
                        break;
                    case "桃園市":
                        sApWhere += "And CITY In (@City1, @City2) ";
                        cmd.Parameters.Add(new SqlParameter("City1", "桃園市"));
                        cmd.Parameters.Add(new SqlParameter("City2", "桃園縣"));
                        break;
                    default:
                        sApWhere += "And CITY = @City1 ";
                        cmd.Parameters.Add(new SqlParameter("City1", sCity));
                        break;
                }
            }
            // 【新增】工程名稱模糊查詢 (參數化)
            if (!string.IsNullOrEmpty(m.keyword2))
            {
                sApWhere += "And (AP_NAME Like @Keyword) ";
                cmd.Parameters.Add(new SqlParameter("Keyword", "%" + m.keyword2 + "%"));
            }

            string sApOrder = "";
            switch (bodyType)
            {
                case "1": sApWhere += "And AP_STAT='15' "; sApOrder = "Order By CITY,S1_ORDER,AP_NAME"; break;
                case "2": sApWhere += "And AP_STAT='20' "; sApOrder = "Order By S1_ORDER,AP_NAME"; break; // S2_ORDER
                case "3": sApWhere += "And AP_STAT='Z2' "; sApOrder = "Order By CITY,S1_ORDER,AP_NAME"; break;
                case "4": sApWhere += "And AP_STAT='90' "; sApOrder = "Order By CITY,AP_NAME"; break;
            }

            cmd.CommandText = "select *,CHK_STAT, S1_ORDER AS S_ORDER,EXT_FG from APV_APMFORM a where " + sApWhere + sApOrder;
            DataTable DT = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            DataView dvAPM = DT.DefaultView;

            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0060/SGE0035Base.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            Worksheet sheet = workbook.Worksheets[0];

            // 標題
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == m.PPID).FirstOrDefault();
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
            var yrcInfo = GetYrcInfo(YRC_ID);
            string sExcelTitle = m.city + yrcInfo?.BGN_SYR + "年度" + pjmform?.PJ_NAME + "明細表"; //○○縣(市)○○年度污水下水道建設計畫明細表

            ExportSheet.SewageDetail sewageDetail = new ExportSheet.SewageDetail();
            sheet = sewageDetail.Sge0035MyPrint2_NEW_Aspose(sheet, dvAPM, DT);

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                byte[] fileBytes = ms.ToArray();

                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     $"{sExcelTitle}_{timestamp}.xlsx");
            }
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "PrintSGE0010")]
        public ActionResult PrintSGE0010(SGE0010Model model)
        {
            // 載入模板
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0010.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            StringBuilder sql = new StringBuilder("SELECT F.* FROM APV_APMFORM F LEFT OUTER JOIN PJDSYRC S on F.PPID = S.PPID AND S.YRC_ID=F.YRC_ID " +
                "left join CMBASE.dbo.OGCDEPT D on F.ADV_DP = D.DEPT_NO left join CMBASE.dbo.OGCUSER U on F.ADV_UID= U.USR_ID WHERE " +
                "F.AP_STAT='15' AND F.PPID =@PPID AND S.OPEN_FG=1 ");
            if (!string.IsNullOrEmpty(model.YRC_ID))
            {
                sql.Append("AND F.YRC_ID=@YRC_ID ");
            }
            if (!string.IsNullOrEmpty(model.city))
            {
                sql.Append("AND F.CITY=@CITY ");
            }
            sql.Append("ORDER BY F.PPID, S.SYR Desc, PSG_NAME");

            SqlCommand cmd = new SqlCommand(sql.ToString());



            cmd.Parameters.Add(new SqlParameter("PPID", model.PPID));
            cmd.Parameters.Add(new SqlParameter("STAT", model.AP_STAT));
            if (!string.IsNullOrEmpty(model.YRC_ID))
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", model.YRC_ID));
            }
            if (!string.IsNullOrEmpty(model.city))
            {
                cmd.Parameters.Add(new SqlParameter("CITY", model.city));
            }


            DataTable ts = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            if (ts.Rows.Count == 0)
            {


                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                    ms.Position = 0;
                    return File(ms.ToArray(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "無資料.xlsx");
                }
            }

            int startRow = 5;
            int sn = 1;
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Replace("$SYR&", ts.Rows[0]["SYR"].ToString());
            sheet.Replace("$PJ_NAME&", model.PJ_NAME);
            foreach (DataRow d in ts.Rows)
            {
                //
                sheet.Cells[startRow, 0].PutValue(sn);
                sheet.Cells[startRow, 1].PutValue(d["AP_NAME"]);
                //2是個洞
                sheet.Cells[startRow, 3].PutValue(d["TD_LINE"]);
                sheet.Cells[startRow, 4].PutValue(d["FUND"]);
                sheet.Cells[startRow, 5].PutValue(d["R_FUND"]);
                sheet.Cells[startRow, 6].PutValue(d["DCS_FUND"]);
                sheet.Cells[startRow, 7].PutValue(d["CEN_DCS_FUND"]);
                sheet.Cells[startRow, 8].PutValue(d["LOC_DCS_FUND"]);

                try
                {
                    if (Convert.ToBoolean(d["SWW1_FG"]))
                    {
                        try { sheet.Cells[startRow, 9].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 10].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 9].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 10].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 9].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 10].PutValue(""); } catch { }
                }

                try
                {
                    if (Convert.ToBoolean(d["SWW2_FG"]))
                    {
                        try { sheet.Cells[startRow, 11].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 12].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 11].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 12].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 11].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 12].PutValue(""); } catch { }
                }

                try
                {
                    if (Convert.ToBoolean(d["SWW3_FG"]))
                    {
                        try { sheet.Cells[startRow, 13].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 14].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 13].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 14].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 13].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 14].PutValue(""); } catch { }
                }

                try
                {
                    if (Convert.ToBoolean(d["SWW4_FG"]))
                    {
                        try { sheet.Cells[startRow, 15].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 16].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 15].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 16].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 15].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 16].PutValue(""); } catch { }
                }
                try
                {
                    if (Convert.ToBoolean(d["SWW5_FG"]))
                    {
                        try { sheet.Cells[startRow, 17].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 18].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 17].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 18].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 17].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 18].PutValue(""); } catch { }
                }
                try
                {
                    if (Convert.ToBoolean(d["SWW6_FG"]))
                    {
                        try { sheet.Cells[startRow, 19].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 20].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 19].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 20].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 19].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 20].PutValue(""); } catch { }
                }

                try
                {
                    if (Convert.ToBoolean(d["SWW7_FG"]))
                    {
                        try { sheet.Cells[startRow, 21].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 22].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 21].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 22].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 21].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 22].PutValue(""); } catch { }
                }

                try
                {
                    if (Convert.ToBoolean(d["SWW8_FG"]))
                    {
                        try { sheet.Cells[startRow, 23].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 24].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 23].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 24].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 23].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 24].PutValue(""); } catch { }
                }

                try
                {
                    if (Convert.ToBoolean(d["SWW9_FG"]))
                    {
                        try { sheet.Cells[startRow, 25].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 26].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 25].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 26].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 25].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 26].PutValue(""); } catch { }
                }

                try
                {
                    if (Convert.ToBoolean(d["SWW10_FG"]))
                    {
                        try { sheet.Cells[startRow, 27].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 28].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 27].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 28].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 27].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 28].PutValue(""); } catch { }
                }

                try
                {
                    if (Convert.ToBoolean(d["SWW11_FG"]))
                    {
                        try { sheet.Cells[startRow, 29].PutValue("V"); } catch { }
                        try { sheet.Cells[startRow, 30].PutValue(""); } catch { }
                    }
                    else
                    {
                        try { sheet.Cells[startRow, 29].PutValue(""); } catch { }
                        try { sheet.Cells[startRow, 30].PutValue("V"); } catch { }
                    }
                }
                catch
                {
                    try { sheet.Cells[startRow, 29].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, 30].PutValue(""); } catch { }
                }
                try { sheet.Cells[startRow, 30].PutValue(""); } catch { }
                sheet.Cells[startRow, 31].PutValue(d["SWWOT_DESC"]);
                sheet.Cells[startRow, 32].PutValue(d["SWWOT_DESC"]);
                startRow++;
                sn++;
                sheet.Cells.InsertRows(startRow, 1, true);
            }

            sheet.AutoFitColumns();

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"PJE0030_{timestamp}.xlsx";
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                ms.Position = 0;
                return File(ms.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName);
            }
        }

        void SetFlagValue(Dictionary<string, object> d, Worksheet sheet, int startRow, string key, int trueCol, int falseCol)
        {
            try
            {
                if (Convert.ToBoolean(d[key]))
                {
                    try { sheet.Cells[startRow, trueCol].PutValue("V"); } catch { }
                    try { sheet.Cells[startRow, falseCol].PutValue(""); } catch { }
                }
                else
                {
                    try { sheet.Cells[startRow, trueCol].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, falseCol].PutValue("V"); } catch { }
                }
            }
            catch
            {
                try { sheet.Cells[startRow, trueCol].PutValue(""); } catch { }
                try { sheet.Cells[startRow, falseCol].PutValue(""); } catch { }
            }
        }

        #region 雨水-匯出複審中計畫評分表

        /// <summary>
        /// 匯出計畫評分表-污水/雨水
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        //[HttpPost]
        //[MultiButton(MatchFormKey = "action", MatchFormValue = "ExportScoreToExcel")]
        //public ActionResult ExportScoreToExcel(SGE0010Model m)
        //{
        //    var data = sgsdb.PJMFORM.SingleOrDefault(s => s.PPID == m.PPID)?.STAT_NO;
        //    if (data == "R") // 雨水中長程計畫
        //        return ExportRainScoreSheet(m);
        //    // 污水中長程計畫
        //    return ExportSewageScoreSheet(m);
        //}

        // 新版 Aspose 雨水計畫提案評分表(分3個工作表)

        /// <summary>
        /// 匯出雨水計畫提案評分表
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ExportRainScoreSheet")]
        public ActionResult ExportRainScoreSheet(SGE0010Model m)
        {
            ExportSheet.RainScore rainScore = new ExportSheet.RainScore();

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");

            using (var ms = new MemoryStream())
            {
                byte[] fileBytes = rainScore.GenerateScoreReportBytes(sgsdb, m.PPID, m.YRC_ID, m.city, m.tab, m.keyword2, out string sExcelTitle);

                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{sExcelTitle}評分表_{timestamp}.xlsx");
            }
        }

        /// (修正)已被 ExportSheet.RainScore.GenerateScoreReportBytes 取代
        /// <summary>
        /// 設定雨水計畫提案評分表工作簿內容
        /// </summary>
        /// <param name="PPID"></param>
        /// <param name="YRC_ID"></param>
        /// <param name="cityText"></param>
        /// <param name="tab"></param>
        /// <param name="apName"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        private byte[] GenerateScoreReportBytes(decimal PPID, string YRC_ID, string cityText, int tab, string apName, out string title)
        {
            ExportSheet.RainScore rainScore = new ExportSheet.RainScore();

            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0060/SGE0060Score_R.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // 標題資料取得
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == PPID).FirstOrDefault();
            string syrText = "";
            if (decimal.TryParse(YRC_ID, out var yrcIdDecimal))
            {
                var yrcInfo = sgsdb.PJDSYRC.SingleOrDefault(w => w.PPID == PPID && w.YRC_ID == yrcIdDecimal);
                syrText = (yrcInfo != null) ? $"{yrcInfo.BGN_SYR}年度" : "";
            }
            title = $"{cityText}{syrText}{pjmform.PJ_NAME}";

            // 擬辦工程(含用地)評分表
            DataTable dt = rainScore.GetRainScoreTable(1, PPID, YRC_ID, cityText, tab, apName);
            string header = $"{title} 擬辦工程(含用地)評分表"; // ○○縣(市)○○年度因應氣候變遷縣市管河川及排水整體改善計畫（115-118年）擬辦工程(含用地)評分表
            rainScore.SetRainScoreSheet(1, workbook, dt, header);

            // 擬辦規劃(檢討)評分表
            dt = rainScore.GetRainScoreTable(2, PPID, YRC_ID, cityText, tab, apName);
            header = $"{title} 擬辦規劃(檢討)評分表";
            rainScore.SetRainScoreSheet(2, workbook, dt, header);

            // 擬辦非工程措施評分表
            dt = rainScore.GetRainScoreTable(3, PPID, YRC_ID, cityText, tab, apName);
            header = $"{title} 擬辦非工程措施評分表";
            rainScore.SetRainScoreSheet(3, workbook, dt, header);

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// 取得案件類別查詢字串
        /// </summary>
        /// <param name="type"> 類型(1：工程，2：規劃，3：非工程) </param>
        /// <returns></returns>
        private string GetPJ_CATY_CODE(int type)
        {
            switch (type)
            {
                case 1: // 工程
                    return "'CA','CB','CC','CD','CE'";
                case 2: // 規劃
                    return "'PA','PB'";
                case 3: // 非工程
                    return "'OA','OB','OC'";
                default:
                    return "''"; // 不符合 1~3
            }
        }

        #endregion

        #region 雨水-匯出複審中計畫明細表

        ///// <summary>
        ///// 匯出計畫提案明細表-污水/雨水
        ///// </summary>
        ///// <param name="m"></param>
        ///// <returns></returns>
        //[HttpPost]
        //[MultiButton(MatchFormKey = "action", MatchFormValue = "ExportDetailsToExcel")]
        //public ActionResult ExportDetailsToExcel(SGE0010Model m)
        //{
        //    var data = sgsdb.PJMFORM.SingleOrDefault(s => s.PPID == m.PPID)?.STAT_NO;
        //    if (data == "R") // 雨水中長程計畫
        //        return ExportRainDetailsSheet(m);
        //    // 污水中長程計畫
        //    return ExportSewageDetailsSheet(m);
        //}

        /// <summary>
        /// 匯出雨水計畫提案明細表
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ExportRainDetailsSheet")]
        public ActionResult ExportRainDetailsSheet(SGE0010Model m)
        {
            ExportSheet.RainDetail rainDetail = new ExportSheet.RainDetail();

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");

            using (var ms = new MemoryStream())
            {
                byte[] fileBytes = rainDetail.GenerateReportBytes(sgsdb, m.PPID, m.YRC_ID, m.tab, m.city, m.keyword2, out string sExcelTitle);

                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{sExcelTitle}明細表_{timestamp}.xlsx");
            }
        }

        /// (修正)已被 ExportSheet.RainDetail.GenerateReportBytes 取代
        /// <summary>
        /// 設定雨水計畫提案明細表工作簿內容
        /// </summary>
        /// <param name="PPID"></param>
        /// <param name="YRC_ID"></param>
        /// <param name="tab"></param>
        /// <param name="cityText"></param>
        /// <param name="apName"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        private byte[] GenerateReportBytes(decimal PPID, string YRC_ID, int tab, string cityText, string apName, out string title)
        {
            ExportSheet.RainDetail rainDetail = new ExportSheet.RainDetail();

            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0060/SGE0060Details_R.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // 頁首頁尾資料取得
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == PPID).FirstOrDefault();
            string syrText = "";
            if (decimal.TryParse(YRC_ID, out var yrcIdDecimal))
            {
                var yrcInfo = sgsdb.PJDSYRC.SingleOrDefault(w => w.PPID == PPID && w.YRC_ID == yrcIdDecimal);
                syrText = (yrcInfo != null) ? $"{yrcInfo.BGN_SYR}年度" : "";
            }
            title = $"{cityText}{syrText}{pjmform.PJ_NAME}";
            string footer = "";

            // 擬辦工程(含用地)明細表
            DataTable dt = rainDetail.GetRainDetailTable(1, PPID, YRC_ID, tab, cityText, apName);
            string header = $"{title} 擬辦工程(含用地)明細表"; // ○○縣(市)○○年度因應氣候變遷縣市管河川及排水整體改善計畫（115-118年）擬辦工程(含用地)明細表
            rainDetail.SetRainDetailsSheet(1, workbook, dt, header, footer);

            // 擬辦規劃(檢討)明細表
            dt = rainDetail.GetRainDetailTable(2, PPID, YRC_ID, tab, cityText, apName);
            header = $"{title} 擬辦規劃(檢討)明細表";
            rainDetail.SetRainDetailsSheet(2, workbook, dt, header, footer);

            // 擬辦非工程措施明細表
            dt = rainDetail.GetRainDetailTable(3, PPID, YRC_ID, tab, cityText, apName);
            header = $"{title} 擬辦非工程措施明細表";
            rainDetail.SetRainDetailsSheet(3, workbook, dt, header, footer);

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                return ms.ToArray();
            }
        }

        #endregion

        #region 永續經營組-匯出評分表/明細表

        private string GetApStatByTab(int Tab)
        {
            switch (Tab)
            {
                case 1:
                    return "00"; /// 待收件

                case 2:
                    return "10"; /// 初審中

                case 3:
                    return "15"; /// 複審中

                case 4:
                    return "20"; /// 核定中

                case 5:
                    return "90"; /// 已成案

                case 6:
                    return "Z1"; /// 未過初審

                case 7:
                    return "Z2"; /// 未過複審

                default:
                    return "";
            }
        }

        /// <summary>
        /// 永續經營組 - 評分表[資料來源]
        /// </summary>
        /// <param name="ppid">計畫 PPID</param>
        /// <param name="apStat">案件狀態 AP_STAT</param>
        /// <param name="m">查詢條件 Model</param>
        /// <returns>評分表資料</returns>
        private DataTable GetEnvironmentScoreData(decimal PPID, string APStat, SGE0010Model m, string YRC_ID, string CATY_COD)
        {
            string sql = $@"SELECT
                        APM.PPID,
                        ISNULL(APM.SYR, '') AS SYR,
                        ISNULL(APM.AP_NAME, '') AS AP_NAME,
                        ISNULL(APM.AP_ALL_NOTE, '') AS ALL_NOTE,
                        ISNULL(APM.AP_YEAR_NOTE, '') AS YEAR_NOTE,
                        (ISNULL(APM.TCEN_FUND, 0) + ISNULL(APM.TLOC_FUND, 0)) / 1000 AS M_T_FUND, -- 總工程經費
                        (ISNULL(APM.RCEN_FUND, 0) + ISNULL(APM.RLOC_FUND, 0)) / 1000 AS M_R_FUND, -- 以前年度執行經費
                        ISNULL(FD.M_FUND, 0) / 1000 AS M_FUND,      -- 本年度申請經費
                        ISNULL(FD.CEN_FUND, 0) / 1000 AS CEN_FUND,  -- 中央款
                        ISNULL(FD.LOC_FUND, 0) / 1000 AS LOC_FUND,  -- 地方款
                        ISNULL(APM.SC1_SCORE, 0) AS SC1_SCORE,      -- (一)地方政府內部競爭序位分數(22分)
                        ISNULL(APM.SC2_SCORE, 0) AS SC2_SCORE,      -- (二)計畫總體規劃完善性(10分)
                        ISNULL(APM.SC3_SCORE, 0) AS SC3_SCORE,      -- (三)地方政府發展重點區域(6分)
                        ISNULL(APM.SC4_SCORE, 0) AS SC4_SCORE,      -- (四)與污水下水道建設計畫配合情形(12分)
                        ISNULL(APM.SC5_SCORE, 0) AS SC5_SCORE,      -- (五)水域生態及水質改善效益(20分)
                        ISNULL(APM.SC6_SCORE, 0) AS SC6_SCORE,      -- (六)水環境改善效益(12分)
                        ISNULL(APM.SC7_SCORE, 0) AS SC7_SCORE,      -- (七)民眾認同度(6分)
                        ISNULL(APM.SC8_SCORE, 0) AS SC8_SCORE,      -- (八)營運管理計畫完整性(6分)
                        ISNULL(APM.SC9_SCORE, 0) AS SC9_SCORE,      -- (九)地方配合款編列情況及過去3年相關計畫執行績效(6分)
                        ISNULL(APM.SWW1_FG, '') AS SWW1_FG,         -- 水環境是否完成
                        ISNULL(APM.SWWOT_DESC, '') AS SWWOT_DESC,   -- 延續/檢討修正
                        ISNULL(APM.SWW2_FG, '') AS SWW2_FG,         -- 用地取得
                        ISNULL(APM.S1_ORDER, '') AS S1_ORDER,       -- 辦理優先次序
                        APM.CHK_STAT,                               -- 本年度是否補助
                        ISNULL(APM.S1_COMT, '') AS S1_COMT,         -- 初評意見
                        ISNULL(APM.S2_COMT, '') AS S2_COMT,         -- 複審意見
                        ISNULL(APM.AP_NOTE, '') AS AP_NOTE,                 -- 備註
                        ISNULL(SMC_STAT.CODE_DESC, '無填寫') AS EXT_DESC,    -- 辦理狀態
                        APM.PJ_CATY_CODE,
                        ISNULL(SMC_CATY.CODE_DESC, '無填寫') AS CATY_DESC,   -- 工程分類
                        ISNULL(APM.AP_STAT, '') AS AP_STAT,                 -- 案件狀態
                        APM.BUDGET_TP AS BUDGET_TP
                    FROM APMFORM APM
                    LEFT JOIN SMCCODE SMC_STAT ON SMC_STAT.CATY_COD = CASE APM.EXT_FG WHEN '0' THEN '01' WHEN '1' THEN '02' END AND SMC_STAT.SYS_ID = 'TT' AND SMC_STAT.CATY_ID = '01' -- 辦理狀態
                    LEFT JOIN SMCCODE SMC_CATY ON SMC_CATY.CATY_COD = APM.PJ_CATY_CODE AND SMC_CATY.SYS_ID = 'TT' AND SMC_CATY.CATY_ID = '05' -- 工程分類
                    OUTER APPLY (
                        SELECT
                            SUM(ISNULL(FUND, 0)) AS M_FUND,
                            SUM(CASE WHEN FIT_CODE = '00' THEN ISNULL(FUND, 0) ELSE 0 END) AS CEN_FUND,
                            SUM(CASE WHEN FIT_CODE = '11' THEN ISNULL(FUND, 0) ELSE 0 END) AS LOC_FUND
                        FROM APDFUND FD WHERE FD.AID = APM.AID AND FD.SYR = APM.SYR) FD
                    WHERE APM.PPID IN ( @PPID ) AND APM.YRC_ID = @YRC_ID AND APM.AP_STAT = @AP_STAT AND APM.PJ_CATY_CODE = @CATY_COD ";

            SqlCommand cmd = new SqlCommand(sql);

            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            cmd.Parameters.Add(new SqlParameter("AP_STAT", APStat));
            cmd.Parameters.Add(new SqlParameter("CATY_COD", CATY_COD));

            if (!string.IsNullOrEmpty(m.keyword2))
            {
                sql += " AND APM.AP_NAME LIKE @KEYWORD ";
                cmd.Parameters.Add(new SqlParameter("KEYWORD", "%" + m.keyword2 + "%"));
            }

            if (!string.IsNullOrEmpty(m.city))
            {
                sql += " AND APM.CITY = @CITY ";
                cmd.Parameters.Add(new SqlParameter("CITY", m.city));
            }

            cmd.CommandText = sql;

            Debug.WriteLine($"PPID : {PPID}");
            Debug.WriteLine($"YRC_ID : {YRC_ID}");
            Debug.WriteLine($"APStat : {APStat}");
            Debug.WriteLine($"CATY_COD : {CATY_COD}");
            Debug.WriteLine($"AP_NAME : {m.keyword2}");
            Debug.WriteLine($"sql = {sql}");

            return BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
        }

        /// <summary>
        /// 永續經營組 - 評分表[Sheet1]:水域環境改善景觀營造評分表
        /// </summary>
        /// <param name="WorkBook">Excel Workbook</param>
        /// <param name="DT">資料來源</param>
        /// <param name="PJNAME">計畫名稱</param>
        /// <param name="m">查詢條件</param>
        private void WriteEnvironmentScoreSheet1(Workbook WorkBook, DataTable DT, string PJNAME, SGE0010Model m)
        {
            Worksheet ws = WorkBook.Worksheets[0];

            /// Step 1: 寫入標題
            ws.Cells["A2"].PutValue(PJNAME + " 評分表");

            /// Step 2: 設定模板列位置
            /// Excel 第 6 列 = index 5：藍色分類列模板
            /// Excel 第 7 列 = index 6：資料列模板
            /// Excel 第 10 列 = index 9：總計列模板
            int sourceCategoryTemplateRow = 5;
            int sourceDataTemplateRow = 6;
            int sourceTotalTemplateRow = 9;

            /// Step 3: 先把模板列複製到很下面，避免寫資料時把模板列覆蓋掉
            /// 注意：後面不使用 InsertRow / DeleteRows，所以 temp row 不會位移
            int tempCategoryTemplateRow = ws.Cells.MaxDataRow + 20;
            int tempDataTemplateRow = tempCategoryTemplateRow + 1;
            int tempTotalTemplateRow = tempCategoryTemplateRow + 2;

            ws.Cells.CopyRow(ws.Cells, sourceCategoryTemplateRow, tempCategoryTemplateRow);
            ws.Cells.CopyRow(ws.Cells, sourceDataTemplateRow, tempDataTemplateRow);
            ws.Cells.CopyRow(ws.Cells, sourceTotalTemplateRow, tempTotalTemplateRow);

            /// Step 4: Sheet1 資料起始列
            /// Excel 第 6 列 = index 5
            int row = 5;

            /// Step 5: 分類順序
            string[] categoryOrder = new string[] { "無分類", "業務費", "設備投資費", "獎補助費" };

            /// Step 6: 總計變數
            double totalTFund = 0;
            double totalRFund = 0;
            double totalMFund = 0;
            double totalCenFund = 0;
            double totalLocFund = 0;
            int Item = 1;

            /// 資料列統一使用整列範圍再次套用置中，避免複製模板列後個別儲存格對齊未生效
            Style dataRowCenterStyle = WorkBook.CreateStyle();
            dataRowCenterStyle.HorizontalAlignment = TextAlignmentType.Center;
            dataRowCenterStyle.VerticalAlignment = TextAlignmentType.Center;
            dataRowCenterStyle.IsTextWrapped = true;
            dataRowCenterStyle.IndentLevel = 0;

            StyleFlag dataRowCenterFlag = new StyleFlag();
            dataRowCenterFlag.Alignments = true;
            dataRowCenterFlag.WrapText = true;
            dataRowCenterFlag.Indent = true;

            /// Step 10: 依照分類順序寫入
            foreach (string categoryName in categoryOrder)
            {
                List<DataRow> dataRows = DT.AsEnumerable()
                    .Where(dr =>
                    {
                        string budgetTp = null;

                        /// BUDGET_TP 必須保留 NULL
                        /// NULL / 空白 / 未定義代碼 => 無分類
                        if (DT.Columns.Contains("BUDGET_TP") && dr["BUDGET_TP"] != null && dr["BUDGET_TP"] != DBNull.Value)
                        {
                            budgetTp = Convert.ToString(dr["BUDGET_TP"]).Trim();
                        }

                        string currentCategory = "無分類";

                        if (string.IsNullOrEmpty(budgetTp))
                        {
                            currentCategory = "無分類";
                        }
                        else if (budgetTp == "01" || budgetTp == "02" || budgetTp == "03")
                        {
                            currentCategory = "業務費";
                        }
                        else if (budgetTp == "04" || budgetTp == "09")
                        {
                            currentCategory = "設備投資費";
                        }
                        else if (budgetTp == "05" || budgetTp == "06" || budgetTp == "07" || budgetTp == "08" || budgetTp == "10")
                        {
                            currentCategory = "獎補助費";
                        }
                        else
                        {
                            currentCategory = "無分類";
                        }

                        return currentCategory == categoryName;
                    }).ToList();

                if (dataRows.Count == 0) { continue; }

                /// Step 10-1: 複製藍色分類列模板
                ws.Cells.CopyRow(ws.Cells, tempCategoryTemplateRow, row);

                /// Step 10-2: 解除目前列 A~AG 可能存在的合併
                /// A = 0, AG = 32，共 33 欄
                for (int i = ws.Cells.MergedCells.Count - 1; i >= 0; i--)
                {
                    CellArea area = (CellArea)ws.Cells.MergedCells[i];

                    if (area.StartRow <= row && area.EndRow >= row)
                    {
                        ws.Cells.UnMerge(
                            area.StartRow,
                            area.StartColumn,
                            area.EndRow - area.StartRow + 1,
                            area.EndColumn - area.StartColumn + 1
                        );
                    }
                }

                /// Step 10-3: 清空 A~AG
                for (int col = 0; col <= 32; col++)
                {
                    ws.Cells[row, col].PutValue("");
                }

                ws.Cells.Merge(row, 0, 1, 33); /// 合併 A ~ AG
                ws.Cells[row, 0].PutValue(categoryName); /// 寫入分類名稱

                /// Step 10-6: 強制設定分類列格式：藍色底色 + A~AG 全框線 + 文字靠左
                for (int col = 0; col <= 32; col++)
                {
                    Style style = ws.Cells[row, col].GetStyle();

                    style.Pattern = BackgroundType.Solid;
                    style.ForegroundColor = Color.FromArgb(183, 222, 232);

                    /// 分類列：水平靠左、垂直置中
                    style.HorizontalAlignment = TextAlignmentType.Left;
                    style.VerticalAlignment = TextAlignmentType.Center;

                    /// 縮排，避免分類文字貼齊左框線
                    style.IndentLevel = 1;

                    style.Font.IsBold = true;
                    style.IsTextWrapped = true;

                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    style.Borders[BorderType.TopBorder].Color = Color.Black;
                    style.Borders[BorderType.BottomBorder].Color = Color.Black;
                    style.Borders[BorderType.LeftBorder].Color = Color.Black;
                    style.Borders[BorderType.RightBorder].Color = Color.Black;

                    ws.Cells[row, col].SetStyle(style);
                }

                row++;

                /// Step 10-7: 寫入該分類資料
                foreach (DataRow DR in dataRows)
                {
                    ws.Cells.CopyRow(ws.Cells, tempDataTemplateRow, row);

                    /// 清空 A~AG
                    for (int col = 0; col <= 32; col++)
                    {
                        ws.Cells[row, col].PutValue("");
                    }

                    double Score1 = Convert.ToDouble(DR["SC1_SCORE"].ToString());
                    double Score2 = Convert.ToDouble(DR["SC2_SCORE"].ToString());
                    double Score3 = Convert.ToDouble(DR["SC3_SCORE"].ToString());
                    double Score4 = Convert.ToDouble(DR["SC4_SCORE"].ToString());
                    double Score5 = Convert.ToDouble(DR["SC5_SCORE"].ToString());
                    double Score6 = Convert.ToDouble(DR["SC6_SCORE"].ToString());
                    double Score7 = Convert.ToDouble(DR["SC7_SCORE"].ToString());
                    double Score8 = Convert.ToDouble(DR["SC8_SCORE"].ToString());
                    double Score9 = Convert.ToDouble(DR["SC9_SCORE"].ToString());

                    double TotalScore = Score1 + Score2 + Score3 + Score4 + Score5 + Score6 + Score7 + Score8 + Score9;

                    /// A ~ AG 欄位寫入
                    ws.Cells[row, 0].PutValue(Item++);                                          /// A: 項次
                    ws.Cells[row, 1].PutValue(DR["AP_NAME"].ToString());                        /// B: 工程名稱
                    ws.Cells[row, 2].PutValue(DR["ALL_NOTE"].ToString());                       /// C: 全期內容
                    ws.Cells[row, 3].PutValue(DR["YEAR_NOTE"].ToString());                      /// D: 本年度辦理內容
                    ws.Cells[row, 4].PutValue(Convert.ToDouble(DR["M_T_FUND"].ToString()));     /// E: 工程總經費
                    ws.Cells[row, 5].PutValue(Convert.ToDouble(DR["M_R_FUND"].ToString()));     /// F: 以前年度已執行經費
                    ws.Cells[row, 6].PutValue(Convert.ToDouble(DR["M_FUND"].ToString()));       /// G: 本年度申請經費
                    ws.Cells[row, 7].PutValue(Convert.ToDouble(DR["CEN_FUND"].ToString()));     /// H: 中央補助款
                    ws.Cells[row, 8].PutValue(Convert.ToDouble(DR["LOC_FUND"].ToString()));     /// I: 地方配合款
                    ws.Cells[row, 9].PutValue(Score1);                                          /// J: SC1
                    ws.Cells[row, 10].PutValue(Score2);                                         /// K: SC2
                    ws.Cells[row, 11].PutValue(Score3);                                         /// L: SC3
                    ws.Cells[row, 12].PutValue(Score4);                                         /// M: SC4
                    ws.Cells[row, 13].PutValue(Score5);                                         /// N: SC5
                    ws.Cells[row, 14].PutValue(Score6);                                         /// O: SC6
                    ws.Cells[row, 15].PutValue(Score7);                                         /// P: SC7
                    ws.Cells[row, 16].PutValue(Score8);                                         /// Q: SC8
                    ws.Cells[row, 17].PutValue(Score9);                                         /// R: SC9
                    ws.Cells[row, 18].PutValue(TotalScore);                                     /// S: 總計分
                    ws.Cells[row, 19].PutValue(DR["SWW1_FG"].ToString() == "0" ? "是" : "否");   /// T: 水環境空間發展藍圖是否完成
                    ws.Cells[row, 20].PutValue(DR["SWWOT_DESC"]);                               /// U
                    ws.Cells[row, 21].PutValue(DR["SWW2_FG"].ToString() == "0" ? "是" : "否");   /// V : 用地取得
                    ws.Cells[row, 22].PutValue(DR["S1_ORDER"]);                                 /// W
                    string chkStat = DR["CHK_STAT"] == DBNull.Value ? "" : DR["CHK_STAT"].ToString().Trim().ToUpper();
                    ws.Cells[row, 23].PutValue(chkStat == "Y" ? "V" : "");   /// X欄：是
                    ws.Cells[row, 24].PutValue(chkStat == "N" ? "V" : "");   /// Y欄：否
                    ws.Cells[row, 25].PutValue(Convert.ToDouble(DR["M_FUND"].ToString()));      /// Z: 總經費
                    ws.Cells[row, 26].PutValue(Convert.ToDouble(DR["CEN_FUND"].ToString()));    /// AA: 中央補助款
                    ws.Cells[row, 27].PutValue(Convert.ToDouble(DR["LOC_FUND"].ToString()));    /// AB: 地方配合款
                    ws.Cells[row, 28].PutValue(DR["S1_COMT"]);                                  /// AC: 初評意見
                    ws.Cells[row, 29].PutValue(DR["S2_COMT"]);                                  /// AD: 複評意見
                    ws.Cells[row, 30].PutValue(DR["AP_NOTE"]);                                  /// AE: 備註
                    ws.Cells[row, 31].PutValue(DR["EXT_DESC"]);                                 /// AF: 辦理狀態
                    ws.Cells[row, 32].PutValue(DR["CATY_DESC"]);                                /// AG: 工程分類

                    for (int col = 0; col <= 32; col++)
                    {
                        Style dataStyle = ws.Cells[row, col].GetStyle();

                        dataStyle.HorizontalAlignment = TextAlignmentType.Center;
                        dataStyle.VerticalAlignment = TextAlignmentType.Center;
                        dataStyle.IsTextWrapped = true;
                        dataStyle.IndentLevel = 0;

                        dataStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                        dataStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                        dataStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                        dataStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                        dataStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                        dataStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                        dataStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                        dataStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                        StyleFlag dataFlag = new StyleFlag();
                        dataFlag.HorizontalAlignment = true;
                        dataFlag.VerticalAlignment = true;
                        dataFlag.WrapText = true;
                        dataFlag.Indent = true;
                        dataFlag.Borders = true;

                        ws.Cells[row, col].SetStyle(dataStyle, dataFlag);
                    }

                    Range dataRowRange = ws.Cells.CreateRange(row, 0, 1, 33);
                    dataRowRange.ApplyStyle(dataRowCenterStyle, dataRowCenterFlag);

                    /// 固定資料列高度，讓垂直置中效果正常呈現
                    ws.Cells.SetRowHeight(row, 45);

                    totalTFund += Convert.ToDouble(DR["M_T_FUND"].ToString());
                    totalRFund += Convert.ToDouble(DR["M_R_FUND"].ToString());
                    totalMFund += Convert.ToDouble(DR["M_FUND"].ToString());
                    totalCenFund += Convert.ToDouble(DR["CEN_FUND"].ToString());
                    totalLocFund += Convert.ToDouble(DR["LOC_FUND"].ToString());
                    row++;
                }
            }

            /// 寫入總計列
            #region ====================================================================================

            ws.Cells.CopyRow(ws.Cells, tempTotalTemplateRow, row);

            for (int col = 0; col <= 32; col++)
            {
                ws.Cells[row, col].PutValue("");
            }

            ws.Cells[row, 1].PutValue("總計");
            ws.Cells[row, 4].PutValue(totalTFund);
            ws.Cells[row, 5].PutValue(totalRFund);
            ws.Cells[row, 6].PutValue(totalMFund);
            ws.Cells[row, 7].PutValue(totalCenFund);
            ws.Cells[row, 8].PutValue(totalLocFund);
            ws.Cells[row, 25].PutValue(totalMFund);
            ws.Cells[row, 26].PutValue(totalCenFund);
            ws.Cells[row, 27].PutValue(totalLocFund);

            /// Step 11-1: 總計列 A~AG：綠色底色 + 全框線 + 置中
            for (int col = 0; col <= 32; col++)
            {
                Style totalStyle = ws.Cells[row, col].GetStyle();

                totalStyle.Pattern = BackgroundType.Solid;
                totalStyle.ForegroundColor = Color.FromArgb(146, 208, 80);
                totalStyle.HorizontalAlignment = TextAlignmentType.Center;
                totalStyle.VerticalAlignment = TextAlignmentType.Center;
                totalStyle.Font.IsBold = true;
                totalStyle.IsTextWrapped = true;
                totalStyle.IndentLevel = 0;

                totalStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                totalStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                totalStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                totalStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                totalStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                totalStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                totalStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                totalStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                ws.Cells[row, col].SetStyle(totalStyle);
            }

            #endregion =================================================================================

            /// 設定欄寬
            ws.Cells.SetColumnWidth(0, 6);    // A 項次
            ws.Cells.SetColumnWidth(1, 35);   // B 工程名稱
            ws.Cells.SetColumnWidth(2, 35);   // C 全期內容
            ws.Cells.SetColumnWidth(3, 35);   // D 本年度辦理內容
            ws.Cells.SetColumnWidth(4, 20);   // E 工程總經費
            ws.Cells.SetColumnWidth(5, 20);   // F 以前年度已執行經費
            ws.Cells.SetColumnWidth(6, 20);   // G 本年度申請經費
            ws.Cells.SetColumnWidth(7, 20);   // H 中央補助款
            ws.Cells.SetColumnWidth(8, 20);   // I 地方配合款
            ws.Cells.SetColumnWidth(9, 12);   // J
            ws.Cells.SetColumnWidth(10, 12);  // K
            ws.Cells.SetColumnWidth(11, 12);  // L
            ws.Cells.SetColumnWidth(12, 12);  // M
            ws.Cells.SetColumnWidth(13, 12);  // N
            ws.Cells.SetColumnWidth(14, 12);  // O
            ws.Cells.SetColumnWidth(15, 12);  // P
            ws.Cells.SetColumnWidth(16, 12);  // Q
            ws.Cells.SetColumnWidth(17, 12);  // R
            ws.Cells.SetColumnWidth(18, 12);  // S
            ws.Cells.SetColumnWidth(19, 12);  // T
            ws.Cells.SetColumnWidth(20, 12);  // U
            ws.Cells.SetColumnWidth(21, 12);  // V
            ws.Cells.SetColumnWidth(22, 12);  // W
            ws.Cells.SetColumnWidth(23, 12);   // X 是
            ws.Cells.SetColumnWidth(24, 12);   // Y 否
            ws.Cells.SetColumnWidth(25, 20);  // Z
            ws.Cells.SetColumnWidth(26, 20);  // AA
            ws.Cells.SetColumnWidth(27, 20);  // AB
            ws.Cells.SetColumnWidth(28, 35);  // AC
            ws.Cells.SetColumnWidth(29, 35);  // AD
            ws.Cells.SetColumnWidth(30, 35);  // AE
            ws.Cells.SetColumnWidth(31, 35);  // AF
            ws.Cells.SetColumnWidth(32, 35);  // AG

            /// Step 13: 隱藏暫存模板列，不刪除
            /// 不刪除是為了避免列位移造成格式錯亂。
            ws.Cells.HideRows(tempCategoryTemplateRow, 3);
        }

        /// <summary>
        /// 永續經營組 - 評分表[Sheet2]:汙水廠水質改善評分表
        /// </summary>
        /// <param name="WorkBook"></param>
        /// <param name="DT"></param>
        /// <param name="PJNAME"></param>
        /// <param name="m"></param>
        private void WriteEnvironmentScoreSheet2(Workbook WorkBook, DataTable DT, string PJNAME, SGE0010Model m)
        {
            Worksheet ws = WorkBook.Worksheets[1];

            /// Step 1: 寫入標題
            ws.Cells["A2"].PutValue(PJNAME + " 評分表");

            /// Step 2: 設定模板列位置
            /// Excel 第 6 列 = index 5：藍色分類列模板
            /// Excel 第 7 列 = index 6：資料列模板
            /// Excel 第 10 列 = index 9：總計列模板
            int sourceCategoryTemplateRow = 5;
            int sourceDataTemplateRow = 6;
            int sourceTotalTemplateRow = 9;

            /// Step 3: 先把模板列複製到很下面，避免寫資料時把模板列覆蓋掉
            /// 注意：後面不使用 InsertRow / DeleteRows，所以 temp row 不會位移
            int tempCategoryTemplateRow = ws.Cells.MaxDataRow + 20;
            int tempDataTemplateRow = tempCategoryTemplateRow + 1;
            int tempTotalTemplateRow = tempCategoryTemplateRow + 2;

            ws.Cells.CopyRow(ws.Cells, sourceCategoryTemplateRow, tempCategoryTemplateRow);
            ws.Cells.CopyRow(ws.Cells, sourceDataTemplateRow, tempDataTemplateRow);
            ws.Cells.CopyRow(ws.Cells, sourceTotalTemplateRow, tempTotalTemplateRow);

            /// Step 4: Sheet1 資料起始列
            /// Excel 第 6 列 = index 5
            int row = 5;

            /// Step 5: 分類順序
            string[] categoryOrder = new string[] { "無分類", "業務費", "設備投資費", "獎補助費" };

            /// Step 6: 總計變數
            double totalTFund = 0;
            double totalRFund = 0;
            double totalMFund = 0;
            double totalCenFund = 0;
            double totalLocFund = 0;
            int Item = 1;

            /// 資料列統一使用整列範圍再次套用置中，避免複製模板列後個別儲存格對齊未生效
            Style dataRowCenterStyle = WorkBook.CreateStyle();
            dataRowCenterStyle.HorizontalAlignment = TextAlignmentType.Center;
            dataRowCenterStyle.VerticalAlignment = TextAlignmentType.Center;
            dataRowCenterStyle.IsTextWrapped = true;
            dataRowCenterStyle.IndentLevel = 0;

            StyleFlag dataRowCenterFlag = new StyleFlag();
            dataRowCenterFlag.Alignments = true;
            dataRowCenterFlag.WrapText = true;
            dataRowCenterFlag.Indent = true;

            /// Step 10: 依照分類順序寫入
            foreach (string categoryName in categoryOrder)
            {
                List<DataRow> dataRows = DT.AsEnumerable()
                    .Where(dr =>
                    {
                        string budgetTp = null;

                        /// BUDGET_TP 必須保留 NULL
                        /// NULL / 空白 / 未定義代碼 => 無分類
                        if (DT.Columns.Contains("BUDGET_TP") && dr["BUDGET_TP"] != null && dr["BUDGET_TP"] != DBNull.Value)
                        {
                            budgetTp = Convert.ToString(dr["BUDGET_TP"]).Trim();
                        }

                        string currentCategory = "無分類";

                        if (string.IsNullOrEmpty(budgetTp))
                        {
                            currentCategory = "無分類";
                        }
                        else if (budgetTp == "01" || budgetTp == "02" || budgetTp == "03")
                        {
                            currentCategory = "業務費";
                        }
                        else if (budgetTp == "04" || budgetTp == "09")
                        {
                            currentCategory = "設備投資費";
                        }
                        else if (budgetTp == "05" || budgetTp == "06" || budgetTp == "07" || budgetTp == "08" || budgetTp == "10")
                        {
                            currentCategory = "獎補助費";
                        }
                        else
                        {
                            currentCategory = "無分類";
                        }

                        return currentCategory == categoryName;
                    }).ToList();

                if (dataRows.Count == 0) { continue; }

                /// Step 10-1: 複製藍色分類列模板
                ws.Cells.CopyRow(ws.Cells, tempCategoryTemplateRow, row);

                /// Step 10-2: 解除目前列 A~AG 可能存在的合併
                /// A = 0, AG = 32，共 33 欄
                for (int i = ws.Cells.MergedCells.Count - 1; i >= 0; i--)
                {
                    CellArea area = (CellArea)ws.Cells.MergedCells[i];

                    if (area.StartRow <= row && area.EndRow >= row)
                    {
                        ws.Cells.UnMerge(
                            area.StartRow,
                            area.StartColumn,
                            area.EndRow - area.StartRow + 1,
                            area.EndColumn - area.StartColumn + 1
                        );
                    }
                }

                /// Step 10-3: 清空 A~AG
                for (int col = 0; col <= 32; col++)
                {
                    ws.Cells[row, col].PutValue("");
                }

                ws.Cells.Merge(row, 0, 1, 33); /// 合併 A ~ AG
                ws.Cells[row, 0].PutValue(categoryName); /// 寫入分類名稱

                /// Step 10-6: 強制設定分類列格式：藍色底色 + A~AG 全框線 + 文字靠左
                for (int col = 0; col <= 32; col++)
                {
                    Style style = ws.Cells[row, col].GetStyle();

                    style.Pattern = BackgroundType.Solid;
                    style.ForegroundColor = Color.FromArgb(183, 222, 232);

                    /// 分類列：水平靠左、垂直置中
                    style.HorizontalAlignment = TextAlignmentType.Left;
                    style.VerticalAlignment = TextAlignmentType.Center;

                    /// 縮排，避免分類文字貼齊左框線
                    style.IndentLevel = 1;

                    style.Font.IsBold = true;
                    style.IsTextWrapped = true;

                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    style.Borders[BorderType.TopBorder].Color = Color.Black;
                    style.Borders[BorderType.BottomBorder].Color = Color.Black;
                    style.Borders[BorderType.LeftBorder].Color = Color.Black;
                    style.Borders[BorderType.RightBorder].Color = Color.Black;

                    ws.Cells[row, col].SetStyle(style);
                }

                row++;

                /// Step 10-7: 寫入該分類資料
                foreach (DataRow DR in dataRows)
                {
                    ws.Cells.CopyRow(ws.Cells, tempDataTemplateRow, row);

                    /// 清空 A~AG
                    for (int col = 0; col <= 32; col++)
                    {
                        ws.Cells[row, col].PutValue("");
                    }

                    double Score1 = Convert.ToDouble(DR["SC1_SCORE"].ToString());
                    double Score2 = Convert.ToDouble(DR["SC2_SCORE"].ToString());
                    double Score3 = Convert.ToDouble(DR["SC3_SCORE"].ToString());
                    double Score4 = Convert.ToDouble(DR["SC4_SCORE"].ToString());
                    double Score5 = Convert.ToDouble(DR["SC5_SCORE"].ToString());
                    double Score6 = Convert.ToDouble(DR["SC6_SCORE"].ToString());
                    double Score7 = Convert.ToDouble(DR["SC7_SCORE"].ToString());
                    double Score8 = Convert.ToDouble(DR["SC8_SCORE"].ToString());
                    double Score9 = Convert.ToDouble(DR["SC9_SCORE"].ToString());

                    double TotalScore = Score1 + Score2 + Score3 + Score4 + Score5 + Score6 + Score7 + Score8 + Score9;

                    /// A ~ AG 欄位寫入
                    ws.Cells[row, 0].PutValue(Item++);                                          /// A: 項次
                    ws.Cells[row, 1].PutValue(DR["AP_NAME"].ToString());                        /// B: 工程名稱
                    ws.Cells[row, 2].PutValue(DR["ALL_NOTE"].ToString());                       /// C: 全期內容
                    ws.Cells[row, 3].PutValue(DR["YEAR_NOTE"].ToString());                      /// D: 本年度辦理內容
                    ws.Cells[row, 4].PutValue(Convert.ToDouble(DR["M_T_FUND"].ToString()));     /// E: 工程總經費
                    ws.Cells[row, 5].PutValue(Convert.ToDouble(DR["M_R_FUND"].ToString()));     /// F: 以前年度已執行經費
                    ws.Cells[row, 6].PutValue(Convert.ToDouble(DR["M_FUND"].ToString()));       /// G: 本年度申請經費
                    ws.Cells[row, 7].PutValue(Convert.ToDouble(DR["CEN_FUND"].ToString()));     /// H: 中央補助款
                    ws.Cells[row, 8].PutValue(Convert.ToDouble(DR["LOC_FUND"].ToString()));     /// I: 地方配合款
                    ws.Cells[row, 9].PutValue(Score1);                                          /// J: SC1
                    ws.Cells[row, 10].PutValue(Score2);                                         /// K: SC2
                    ws.Cells[row, 11].PutValue(Score3);                                         /// L: SC3
                    ws.Cells[row, 12].PutValue(Score4);                                         /// M: SC4
                    ws.Cells[row, 13].PutValue(Score5);                                         /// N: SC5
                    ws.Cells[row, 14].PutValue(Score6);                                         /// O: SC6
                    ws.Cells[row, 15].PutValue(Score7);                                         /// P: SC7
                    ws.Cells[row, 16].PutValue(Score8);                                         /// Q: SC8
                    ws.Cells[row, 17].PutValue(Score9);                                         /// R: SC9
                    ws.Cells[row, 18].PutValue(TotalScore);                                     /// S: 總計分
                    ws.Cells[row, 19].PutValue(DR["SWW1_FG"].ToString() == "0" ? "是" : "否");   /// T: 水環境空間發展藍圖是否完成
                    ws.Cells[row, 20].PutValue(DR["SWWOT_DESC"]);                               /// U
                    ws.Cells[row, 21].PutValue(DR["SWW2_FG"].ToString() == "0" ? "是" : "否");   /// V : 用地取得
                    ws.Cells[row, 22].PutValue(DR["S1_ORDER"]);                                 /// W
                    string chkStat = DR["CHK_STAT"] == DBNull.Value ? "" : DR["CHK_STAT"].ToString().Trim().ToUpper();
                    ws.Cells[row, 23].PutValue(chkStat == "Y" ? "V" : "");   /// X欄：是
                    ws.Cells[row, 24].PutValue(chkStat == "N" ? "V" : "");   /// Y欄：否
                    ws.Cells[row, 25].PutValue(Convert.ToDouble(DR["M_FUND"].ToString()));      /// Z: 總經費
                    ws.Cells[row, 26].PutValue(Convert.ToDouble(DR["CEN_FUND"].ToString()));    /// AA: 中央補助款
                    ws.Cells[row, 27].PutValue(Convert.ToDouble(DR["LOC_FUND"].ToString()));    /// AB: 地方配合款
                    ws.Cells[row, 28].PutValue(DR["S1_COMT"]);                                  /// AC: 初評意見
                    ws.Cells[row, 29].PutValue(DR["S2_COMT"]);                                  /// AD: 複評意見
                    ws.Cells[row, 30].PutValue(DR["AP_NOTE"]);                                  /// AE: 備註
                    ws.Cells[row, 31].PutValue(DR["EXT_DESC"]);                                 /// AF: 辦理狀態
                    ws.Cells[row, 32].PutValue(DR["CATY_DESC"]);                                /// AG: 工程分類

                    for (int col = 0; col <= 32; col++)
                    {
                        Style dataStyle = ws.Cells[row, col].GetStyle();

                        dataStyle.HorizontalAlignment = TextAlignmentType.Center;
                        dataStyle.VerticalAlignment = TextAlignmentType.Center;
                        dataStyle.IsTextWrapped = true;
                        dataStyle.IndentLevel = 0;

                        dataStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                        dataStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                        dataStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                        dataStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                        dataStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                        dataStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                        dataStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                        dataStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                        StyleFlag dataFlag = new StyleFlag();
                        dataFlag.HorizontalAlignment = true;
                        dataFlag.VerticalAlignment = true;
                        dataFlag.WrapText = true;
                        dataFlag.Indent = true;
                        dataFlag.Borders = true;

                        ws.Cells[row, col].SetStyle(dataStyle, dataFlag);
                    }

                    Range dataRowRange = ws.Cells.CreateRange(row, 0, 1, 33);
                    dataRowRange.ApplyStyle(dataRowCenterStyle, dataRowCenterFlag);

                    /// 固定資料列高度，讓垂直置中效果正常呈現
                    ws.Cells.SetRowHeight(row, 45);

                    totalTFund += Convert.ToDouble(DR["M_T_FUND"].ToString());
                    totalRFund += Convert.ToDouble(DR["M_R_FUND"].ToString());
                    totalMFund += Convert.ToDouble(DR["M_FUND"].ToString());
                    totalCenFund += Convert.ToDouble(DR["CEN_FUND"].ToString());
                    totalLocFund += Convert.ToDouble(DR["LOC_FUND"].ToString());
                    row++;
                }
            }

            /// 寫入總計列
            #region ====================================================================================

            ws.Cells.CopyRow(ws.Cells, tempTotalTemplateRow, row);

            for (int col = 0; col <= 32; col++)
            {
                ws.Cells[row, col].PutValue("");
            }

            ws.Cells[row, 1].PutValue("總計");
            ws.Cells[row, 4].PutValue(totalTFund);
            ws.Cells[row, 5].PutValue(totalRFund);
            ws.Cells[row, 6].PutValue(totalMFund);
            ws.Cells[row, 7].PutValue(totalCenFund);
            ws.Cells[row, 8].PutValue(totalLocFund);
            ws.Cells[row, 25].PutValue(totalMFund);
            ws.Cells[row, 26].PutValue(totalCenFund);
            ws.Cells[row, 27].PutValue(totalLocFund);

            /// Step 11-1: 總計列 A~AG：綠色底色 + 全框線 + 置中
            for (int col = 0; col <= 32; col++)
            {
                Style totalStyle = ws.Cells[row, col].GetStyle();

                totalStyle.Pattern = BackgroundType.Solid;
                totalStyle.ForegroundColor = Color.FromArgb(146, 208, 80);
                totalStyle.HorizontalAlignment = TextAlignmentType.Center;
                totalStyle.VerticalAlignment = TextAlignmentType.Center;
                totalStyle.Font.IsBold = true;
                totalStyle.IsTextWrapped = true;
                totalStyle.IndentLevel = 0;

                totalStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                totalStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                totalStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                totalStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                totalStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                totalStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                totalStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                totalStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                ws.Cells[row, col].SetStyle(totalStyle);
            }

            #endregion =================================================================================

            /// 設定欄寬
            ws.Cells.SetColumnWidth(0, 6);    // A 項次
            ws.Cells.SetColumnWidth(1, 35);   // B 工程名稱
            ws.Cells.SetColumnWidth(2, 35);   // C 全期內容
            ws.Cells.SetColumnWidth(3, 35);   // D 本年度辦理內容
            ws.Cells.SetColumnWidth(4, 20);   // E 工程總經費
            ws.Cells.SetColumnWidth(5, 20);   // F 以前年度已執行經費
            ws.Cells.SetColumnWidth(6, 20);   // G 本年度申請經費
            ws.Cells.SetColumnWidth(7, 20);   // H 中央補助款
            ws.Cells.SetColumnWidth(8, 20);   // I 地方配合款
            ws.Cells.SetColumnWidth(9, 12);   // J
            ws.Cells.SetColumnWidth(10, 12);  // K
            ws.Cells.SetColumnWidth(11, 12);  // L
            ws.Cells.SetColumnWidth(12, 12);  // M
            ws.Cells.SetColumnWidth(13, 12);  // N
            ws.Cells.SetColumnWidth(14, 12);  // O
            ws.Cells.SetColumnWidth(15, 12);  // P
            ws.Cells.SetColumnWidth(16, 12);  // Q
            ws.Cells.SetColumnWidth(17, 12);  // R
            ws.Cells.SetColumnWidth(18, 12);  // S
            ws.Cells.SetColumnWidth(19, 12);  // T
            ws.Cells.SetColumnWidth(20, 12);  // U
            ws.Cells.SetColumnWidth(21, 12);  // V
            ws.Cells.SetColumnWidth(22, 12);  // W
            ws.Cells.SetColumnWidth(23, 12);   // X 是
            ws.Cells.SetColumnWidth(24, 12);   // Y 否
            ws.Cells.SetColumnWidth(25, 20);  // Z
            ws.Cells.SetColumnWidth(26, 20);  // AA
            ws.Cells.SetColumnWidth(27, 20);  // AB
            ws.Cells.SetColumnWidth(28, 35);  // AC
            ws.Cells.SetColumnWidth(29, 35);  // AD
            ws.Cells.SetColumnWidth(30, 35);  // AE
            ws.Cells.SetColumnWidth(31, 35);  // AF
            ws.Cells.SetColumnWidth(32, 35);  // AG

            /// Step 13: 隱藏暫存模板列，不刪除
            /// 不刪除是為了避免列位移造成格式錯亂。
            ws.Cells.HideRows(tempCategoryTemplateRow, 3);
        }

        /// <summary>
        /// 永續經營組 - 評分表[Sheet3]:功能提升評分表
        /// </summary>
        /// <param name="WorkBook"></param>
        /// <param name="DT"></param>
        /// <param name="PJNAME"></param>
        /// <param name="m"></param>
        private void WriteEnvironmentScoreSheet3(Workbook WorkBook, DataTable DT, string PJNAME, SGE0010Model m)
        {
            Worksheet ws = WorkBook.Worksheets[2];

            /// Step 1: 寫入標題
            ws.Cells["A2"].PutValue(PJNAME + " 評分表");

            /// Step 2: 設定模板列位置
            /// Excel 第 6 列 = index 5：藍色分類列模板
            /// Excel 第 7 列 = index 6：資料列模板
            /// Excel 第 10 列 = index 9：總計列模板
            int sourceCategoryTemplateRow = 5;
            int sourceDataTemplateRow = 6;
            int sourceTotalTemplateRow = 9;

            /// Step 3: 先把模板列複製到很下面，避免寫資料時把模板列覆蓋掉
            /// 注意：後面不使用 InsertRow / DeleteRows，所以 temp row 不會位移
            int tempCategoryTemplateRow = ws.Cells.MaxDataRow + 20;
            int tempDataTemplateRow = tempCategoryTemplateRow + 1;
            int tempTotalTemplateRow = tempCategoryTemplateRow + 2;

            ws.Cells.CopyRow(ws.Cells, sourceCategoryTemplateRow, tempCategoryTemplateRow);
            ws.Cells.CopyRow(ws.Cells, sourceDataTemplateRow, tempDataTemplateRow);
            ws.Cells.CopyRow(ws.Cells, sourceTotalTemplateRow, tempTotalTemplateRow);

            /// Step 4: Sheet1 資料起始列
            /// Excel 第 6 列 = index 5
            int row = 5;

            /// Step 5: 分類順序
            string[] categoryOrder = new string[] { "無分類", "業務費", "設備投資費", "獎補助費" };

            /// Step 6: 總計變數
            double totalTFund = 0;
            double totalRFund = 0;
            double totalMFund = 0;
            double totalCenFund = 0;
            double totalLocFund = 0;
            int Item = 1;

            /// 資料列統一使用整列範圍再次套用置中，避免複製模板列後個別儲存格對齊未生效
            Style dataRowCenterStyle = WorkBook.CreateStyle();
            dataRowCenterStyle.HorizontalAlignment = TextAlignmentType.Center;
            dataRowCenterStyle.VerticalAlignment = TextAlignmentType.Center;
            dataRowCenterStyle.IsTextWrapped = true;
            dataRowCenterStyle.IndentLevel = 0;

            StyleFlag dataRowCenterFlag = new StyleFlag();
            dataRowCenterFlag.Alignments = true;
            dataRowCenterFlag.WrapText = true;
            dataRowCenterFlag.Indent = true;

            /// Step 10: 依照分類順序寫入
            foreach (string categoryName in categoryOrder)
            {
                List<DataRow> dataRows = DT.AsEnumerable()
                    .Where(dr =>
                    {
                        string budgetTp = null;

                        /// BUDGET_TP 必須保留 NULL
                        /// NULL / 空白 / 未定義代碼 => 無分類
                        if (DT.Columns.Contains("BUDGET_TP") && dr["BUDGET_TP"] != null && dr["BUDGET_TP"] != DBNull.Value)
                        {
                            budgetTp = Convert.ToString(dr["BUDGET_TP"]).Trim();
                        }

                        string currentCategory = "無分類";

                        if (string.IsNullOrEmpty(budgetTp))
                        {
                            currentCategory = "無分類";
                        }
                        else if (budgetTp == "01" || budgetTp == "02" || budgetTp == "03")
                        {
                            currentCategory = "業務費";
                        }
                        else if (budgetTp == "04" || budgetTp == "09")
                        {
                            currentCategory = "設備投資費";
                        }
                        else if (budgetTp == "05" || budgetTp == "06" || budgetTp == "07" || budgetTp == "08" || budgetTp == "10")
                        {
                            currentCategory = "獎補助費";
                        }
                        else
                        {
                            currentCategory = "無分類";
                        }

                        return currentCategory == categoryName;
                    }).ToList();

                if (dataRows.Count == 0) { continue; }

                /// Step 10-1: 複製藍色分類列模板
                ws.Cells.CopyRow(ws.Cells, tempCategoryTemplateRow, row);

                /// Step 10-2: 解除目前列 A~AG 可能存在的合併
                /// A = 0, AG = 32，共 33 欄
                for (int i = ws.Cells.MergedCells.Count - 1; i >= 0; i--)
                {
                    CellArea area = (CellArea)ws.Cells.MergedCells[i];

                    if (area.StartRow <= row && area.EndRow >= row)
                    {
                        ws.Cells.UnMerge(
                            area.StartRow,
                            area.StartColumn,
                            area.EndRow - area.StartRow + 1,
                            area.EndColumn - area.StartColumn + 1
                        );
                    }
                }

                /// Step 10-3: 清空 A~AG
                for (int col = 0; col <= 32; col++)
                {
                    ws.Cells[row, col].PutValue("");
                }

                ws.Cells.Merge(row, 0, 1, 33); /// 合併 A ~ AG
                ws.Cells[row, 0].PutValue(categoryName); /// 寫入分類名稱

                /// Step 10-6: 強制設定分類列格式：藍色底色 + A~AG 全框線 + 文字靠左
                for (int col = 0; col <= 32; col++)
                {
                    Style style = ws.Cells[row, col].GetStyle();

                    style.Pattern = BackgroundType.Solid;
                    style.ForegroundColor = Color.FromArgb(183, 222, 232);

                    /// 分類列：水平靠左、垂直置中
                    style.HorizontalAlignment = TextAlignmentType.Left;
                    style.VerticalAlignment = TextAlignmentType.Center;

                    /// 縮排，避免分類文字貼齊左框線
                    style.IndentLevel = 1;

                    style.Font.IsBold = true;
                    style.IsTextWrapped = true;

                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    style.Borders[BorderType.TopBorder].Color = Color.Black;
                    style.Borders[BorderType.BottomBorder].Color = Color.Black;
                    style.Borders[BorderType.LeftBorder].Color = Color.Black;
                    style.Borders[BorderType.RightBorder].Color = Color.Black;

                    ws.Cells[row, col].SetStyle(style);
                }

                row++;

                /// Step 10-7: 寫入該分類資料
                foreach (DataRow DR in dataRows)
                {
                    ws.Cells.CopyRow(ws.Cells, tempDataTemplateRow, row);

                    /// 清空 A~AG
                    for (int col = 0; col <= 32; col++)
                    {
                        ws.Cells[row, col].PutValue("");
                    }

                    double Score1 = Convert.ToDouble(DR["SC1_SCORE"].ToString());
                    double Score2 = Convert.ToDouble(DR["SC2_SCORE"].ToString());
                    double Score3 = Convert.ToDouble(DR["SC3_SCORE"].ToString());
                    double Score4 = Convert.ToDouble(DR["SC4_SCORE"].ToString());
                    double Score5 = Convert.ToDouble(DR["SC5_SCORE"].ToString());
                    double Score6 = Convert.ToDouble(DR["SC6_SCORE"].ToString());
                    double Score7 = Convert.ToDouble(DR["SC7_SCORE"].ToString());
                    double Score8 = Convert.ToDouble(DR["SC8_SCORE"].ToString());
                    double Score9 = Convert.ToDouble(DR["SC9_SCORE"].ToString());

                    double TotalScore = Score1 + Score2 + Score3 + Score4 + Score5 + Score6 + Score7 + Score8 + Score9;

                    /// A ~ AG 欄位寫入
                    ws.Cells[row, 0].PutValue(Item++);                                          /// A: 項次
                    ws.Cells[row, 1].PutValue(DR["AP_NAME"].ToString());                        /// B: 工程名稱
                    ws.Cells[row, 2].PutValue(DR["ALL_NOTE"].ToString());                       /// C: 全期內容
                    ws.Cells[row, 3].PutValue(DR["YEAR_NOTE"].ToString());                      /// D: 本年度辦理內容
                    ws.Cells[row, 4].PutValue(Convert.ToDouble(DR["M_T_FUND"].ToString()));     /// E: 工程總經費
                    ws.Cells[row, 5].PutValue(Convert.ToDouble(DR["M_R_FUND"].ToString()));     /// F: 以前年度已執行經費
                    ws.Cells[row, 6].PutValue(Convert.ToDouble(DR["M_FUND"].ToString()));       /// G: 本年度申請經費
                    ws.Cells[row, 7].PutValue(Convert.ToDouble(DR["CEN_FUND"].ToString()));     /// H: 中央補助款
                    ws.Cells[row, 8].PutValue(Convert.ToDouble(DR["LOC_FUND"].ToString()));     /// I: 地方配合款
                    ws.Cells[row, 9].PutValue(Score1);                                          /// J: SC1
                    ws.Cells[row, 10].PutValue(Score2);                                         /// K: SC2
                    ws.Cells[row, 11].PutValue(Score3);                                         /// L: SC3
                    ws.Cells[row, 12].PutValue(Score4);                                         /// M: SC4
                    ws.Cells[row, 13].PutValue(Score5);                                         /// N: SC5
                    ws.Cells[row, 14].PutValue(Score6);                                         /// O: SC6
                    ws.Cells[row, 15].PutValue(Score7);                                         /// P: SC7
                    ws.Cells[row, 16].PutValue(Score8);                                         /// Q: SC8
                    ws.Cells[row, 17].PutValue(Score9);                                         /// R: SC9
                    ws.Cells[row, 18].PutValue(TotalScore);                                     /// S: 總計分
                    ws.Cells[row, 19].PutValue(DR["SWW1_FG"].ToString() == "0" ? "是" : "否");   /// T: 水環境空間發展藍圖是否完成
                    ws.Cells[row, 20].PutValue(DR["SWWOT_DESC"]);                               /// U
                    ws.Cells[row, 21].PutValue(DR["SWW2_FG"].ToString() == "0" ? "是" : "否");   /// V : 用地取得
                    ws.Cells[row, 22].PutValue(DR["S1_ORDER"]);                                 /// W
                    string chkStat = DR["CHK_STAT"] == DBNull.Value ? "" : DR["CHK_STAT"].ToString().Trim().ToUpper();
                    ws.Cells[row, 23].PutValue(chkStat == "Y" ? "V" : "");   /// X欄：是
                    ws.Cells[row, 24].PutValue(chkStat == "N" ? "V" : "");   /// Y欄：否
                    ws.Cells[row, 25].PutValue(Convert.ToDouble(DR["M_FUND"].ToString()));      /// Z: 總經費
                    ws.Cells[row, 26].PutValue(Convert.ToDouble(DR["CEN_FUND"].ToString()));    /// AA: 中央補助款
                    ws.Cells[row, 27].PutValue(Convert.ToDouble(DR["LOC_FUND"].ToString()));    /// AB: 地方配合款
                    ws.Cells[row, 28].PutValue(DR["S1_COMT"]);                                  /// AC: 初評意見
                    ws.Cells[row, 29].PutValue(DR["S2_COMT"]);                                  /// AD: 複評意見
                    ws.Cells[row, 30].PutValue(DR["AP_NOTE"]);                                  /// AE: 備註
                    ws.Cells[row, 31].PutValue(DR["EXT_DESC"]);                                 /// AF: 辦理狀態
                    ws.Cells[row, 32].PutValue(DR["CATY_DESC"]);                                /// AG: 工程分類

                    for (int col = 0; col <= 32; col++)
                    {
                        Style dataStyle = ws.Cells[row, col].GetStyle();

                        dataStyle.HorizontalAlignment = TextAlignmentType.Center;
                        dataStyle.VerticalAlignment = TextAlignmentType.Center;
                        dataStyle.IsTextWrapped = true;
                        dataStyle.IndentLevel = 0;

                        dataStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                        dataStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                        dataStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                        dataStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                        dataStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                        dataStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                        dataStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                        dataStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                        StyleFlag dataFlag = new StyleFlag();
                        dataFlag.HorizontalAlignment = true;
                        dataFlag.VerticalAlignment = true;
                        dataFlag.WrapText = true;
                        dataFlag.Indent = true;
                        dataFlag.Borders = true;

                        ws.Cells[row, col].SetStyle(dataStyle, dataFlag);
                    }

                    Range dataRowRange = ws.Cells.CreateRange(row, 0, 1, 33);
                    dataRowRange.ApplyStyle(dataRowCenterStyle, dataRowCenterFlag);

                    /// 固定資料列高度，讓垂直置中效果正常呈現
                    ws.Cells.SetRowHeight(row, 45);

                    totalTFund += Convert.ToDouble(DR["M_T_FUND"].ToString());
                    totalRFund += Convert.ToDouble(DR["M_R_FUND"].ToString());
                    totalMFund += Convert.ToDouble(DR["M_FUND"].ToString());
                    totalCenFund += Convert.ToDouble(DR["CEN_FUND"].ToString());
                    totalLocFund += Convert.ToDouble(DR["LOC_FUND"].ToString());
                    row++;
                }
            }

            /// 寫入總計列
            #region ====================================================================================

            ws.Cells.CopyRow(ws.Cells, tempTotalTemplateRow, row);

            for (int col = 0; col <= 32; col++)
            {
                ws.Cells[row, col].PutValue("");
            }

            ws.Cells[row, 1].PutValue("總計");
            ws.Cells[row, 4].PutValue(totalTFund);
            ws.Cells[row, 5].PutValue(totalRFund);
            ws.Cells[row, 6].PutValue(totalMFund);
            ws.Cells[row, 7].PutValue(totalCenFund);
            ws.Cells[row, 8].PutValue(totalLocFund);
            ws.Cells[row, 25].PutValue(totalMFund);
            ws.Cells[row, 26].PutValue(totalCenFund);
            ws.Cells[row, 27].PutValue(totalLocFund);

            /// Step 11-1: 總計列 A~AG：綠色底色 + 全框線 + 置中
            for (int col = 0; col <= 32; col++)
            {
                Style totalStyle = ws.Cells[row, col].GetStyle();

                totalStyle.Pattern = BackgroundType.Solid;
                totalStyle.ForegroundColor = Color.FromArgb(146, 208, 80);
                totalStyle.HorizontalAlignment = TextAlignmentType.Center;
                totalStyle.VerticalAlignment = TextAlignmentType.Center;
                totalStyle.Font.IsBold = true;
                totalStyle.IsTextWrapped = true;
                totalStyle.IndentLevel = 0;

                totalStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                totalStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                totalStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                totalStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                totalStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                totalStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                totalStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                totalStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                ws.Cells[row, col].SetStyle(totalStyle);
            }

            #endregion =================================================================================

            /// 設定欄寬
            ws.Cells.SetColumnWidth(0, 6);    // A 項次
            ws.Cells.SetColumnWidth(1, 35);   // B 工程名稱
            ws.Cells.SetColumnWidth(2, 35);   // C 全期內容
            ws.Cells.SetColumnWidth(3, 35);   // D 本年度辦理內容
            ws.Cells.SetColumnWidth(4, 20);   // E 工程總經費
            ws.Cells.SetColumnWidth(5, 20);   // F 以前年度已執行經費
            ws.Cells.SetColumnWidth(6, 20);   // G 本年度申請經費
            ws.Cells.SetColumnWidth(7, 20);   // H 中央補助款
            ws.Cells.SetColumnWidth(8, 20);   // I 地方配合款
            ws.Cells.SetColumnWidth(9, 12);   // J
            ws.Cells.SetColumnWidth(10, 12);  // K
            ws.Cells.SetColumnWidth(11, 12);  // L
            ws.Cells.SetColumnWidth(12, 12);  // M
            ws.Cells.SetColumnWidth(13, 12);  // N
            ws.Cells.SetColumnWidth(14, 12);  // O
            ws.Cells.SetColumnWidth(15, 12);  // P
            ws.Cells.SetColumnWidth(16, 12);  // Q
            ws.Cells.SetColumnWidth(17, 12);  // R
            ws.Cells.SetColumnWidth(18, 12);  // S
            ws.Cells.SetColumnWidth(19, 12);  // T
            ws.Cells.SetColumnWidth(20, 12);  // U
            ws.Cells.SetColumnWidth(21, 12);  // V
            ws.Cells.SetColumnWidth(22, 12);  // W
            ws.Cells.SetColumnWidth(23, 12);   // X 是
            ws.Cells.SetColumnWidth(24, 12);   // Y 否
            ws.Cells.SetColumnWidth(25, 20);  // Z
            ws.Cells.SetColumnWidth(26, 20);  // AA
            ws.Cells.SetColumnWidth(27, 20);  // AB
            ws.Cells.SetColumnWidth(28, 35);  // AC
            ws.Cells.SetColumnWidth(29, 35);  // AD
            ws.Cells.SetColumnWidth(30, 35);  // AE
            ws.Cells.SetColumnWidth(31, 35);  // AF
            ws.Cells.SetColumnWidth(32, 35);  // AG

            /// Step 13: 隱藏暫存模板列，不刪除
            /// 不刪除是為了避免列位移造成格式錯亂。
            ws.Cells.HideRows(tempCategoryTemplateRow, 3);
        }

        /// <summary>
        /// 永續經營組 - 評分表[Excel]
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ExportEnvironmentScoreToExcel")]
        public ActionResult ExportEnvironmentScoreToExcel(SGE0010Model m)
        {
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0060/PJV0054.xlsx"));

            /// Step 1: 報表標題
            var PID = m.PPID;
            PJMFORM_CMPSGS PJMFORM = sgsdb.PJMFORM.FirstOrDefault(X => X.PPID == PID);
            string PJNAME = PJMFORM.PJ_NAME;

            /// Step 1-1: 查詢 YRC_ID
            string sqlYrc = @" SELECT TOP 1 YRC_ID FROM PJDSYRC WHERE PPID = @PPID ORDER BY YRC_ID DESC ";

            SqlCommand cmdYrc = new SqlCommand(sqlYrc);
            cmdYrc.Parameters.Add(new SqlParameter("PPID", PID));

            DataTable dtYrc = SqlDao.CMPSGSQuery4List(cmdYrc);

            string YRC_ID = string.Empty;

            if (dtYrc.Rows.Count > 0)
            {
                YRC_ID = dtYrc.Rows[0]["YRC_ID"].ToString();
            }
            else
            {
                return Content("查無 YRC_ID");
            }

            /// Step 1-2: 查 CATY_COD => EA / EB / EC 分類
            string CATY_COD = @" SELECT CATY_COD, CODE_DESC FROM SMCCODE WHERE SYS_ID = 'TT' AND CATY_ID = '05' ";

            SqlCommand CATYCOD_Command = new SqlCommand(CATY_COD);
            DataTable CATYCOD_Table = SqlDao.CMPSGSQuery4List(CATYCOD_Command);

            Dictionary<string, string> CATYCOD_MAP = CATYCOD_Table.AsEnumerable().ToDictionary(row => row["CATY_COD"].ToString(), row => row["CODE_DESC"].ToString());

            string EA = "EA";
            string EB = "EB";
            string EC = "EC";

            string EA_DESC = CATYCOD_MAP["EA"];
            string EB_DESC = CATYCOD_MAP["EB"];
            string EC_DESC = CATYCOD_MAP["EC"];

            /// Step 2: 根據目前頁籤取得 AP_STAT
            string APStat = GetApStatByTab(m.tab);

            /// Step 3: 查詢資料
            DataTable DT1 = GetEnvironmentScoreData(PID, APStat, m, YRC_ID, EA);
            DataTable DT2 = GetEnvironmentScoreData(PID, APStat, m, YRC_ID, EB);
            DataTable DT3 = GetEnvironmentScoreData(PID, APStat, m, YRC_ID, EC);

            /// Step 4: 將 DataTable 寫入 3 個 Sheet
            WriteEnvironmentScoreSheet1(workbook, DT1, PJNAME, m);
            WriteEnvironmentScoreSheet2(workbook, DT2, PJNAME, m);
            WriteEnvironmentScoreSheet3(workbook, DT3, PJNAME, m);

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);

                string fileName = $"{PJNAME}評分表{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        /// <summary>
        /// 永續經營組 - 明細(概述)表[資料來源]
        /// </summary>
        /// <param name="ppid">計畫 PPID</param>
        /// <param name="apStat">案件狀態 AP_STAT</param>
        /// <param name="m">查詢條件 Model</param>
        /// <returns>概述表資料</returns>
        private DataTable GetEnvironmentDetailsData(decimal PPID, string APStat, SGE0010Model m, string YRC_ID)
        {
            string sql = $@"SELECT
	                            APM.AID,
	                            APM.PPID,
	                            APM.SYR,
	                            ISNULL(APM.S1_ORDER,0) AS S1_ORDER,				-- 縣市自評優先排序
	                            APM.CITY,	-- 縣市名稱
	                            APM.EXE_MCH,-- 執行機關
	                            ISNULL(APM.AP_PJ_NAME,'無填寫') AS AP_PJ_NAME,	-- 整體計畫名稱
	                            APM.AP_NAME,									-- 分項案件名稱
	                            APM.CONT_NAME,									-- 中央主管機關
	                            ISNULL(LOCP.TOWN,'無填寫') AS TOWN,				-- 鄉鎮市區別
	                            APM.PJ_CATY_CODE,								-- 工作項目分類代碼
	                            SMC.CODE_DESC,									-- 工作項目分類名稱
	                            ISNULL(APM.AP_ALL_NOTE,'') AS AP_ALL_NOTE,		-- 工程內容
	                            ISNULL(FD.CEN_FUND,0) / 1000 AS CEN_FUND,										-- 該年度 ""中央款""
	                            ISNULL(FD.LOC_FUND,0) / 1000 AS LOC_FUND,										-- 該年度 ""地方款""
	                            (ISNULL(APM.TCEN_FUND, 0) + ISNULL(APM.TLOC_FUND, 0)) / 1000 AS M_T_FUND,　		-- 總經費需求(仟元)
	                            (ISNULL(APM.RCEN_FUND, 0) + ISNULL(APM.RLOC_FUND, 0)) / 1000 AS M_R_FUND,　		-- 過去年度已支用金額
	                            (ISNULL(APM.TCEN_FUND,0)) / 1000 AS TCEN_FUND,									-- 中央補助(該年度)
	                            (ISNULL(APM.FCEN_FUND,0)) / 1000 AS FCEN_FUND,									-- 中央補助(以後年度)
	                            (ISNULL(FD.CEN_FUND,0) + ISNULL(APM.FCEN_FUND,0)) / 1000 AS TOTAL_CEN_FUND,		-- 該年度 ""中央款"" + 中央補助(以後年度)
	                            (ISNULL(APM.TLOC_FUND,0)) / 1000 AS TLOC_FUND,									-- 地方分擔(該年度)
	                            (ISNULL(APM.FLOC_FUND,0)) / 1000 AS FLOC_FUND,									-- 地方分擔(以後年度)
	                            (ISNULL(FD.LOC_FUND,0) + ISNULL(APM.FLOC_FUND,0)) / 1000 AS TOTAL_LOC_FUND,		-- 該年度 ""地方款"" + 地方分擔(以後年度)
	                            ISNULL(APM.SWW1_FG, '') AS SWW1_FG,				-- 水環境是否完成
                                ISNULL(APM.SWWOT_DESC, '') AS SWWOT_DESC,		-- 延續/檢討修正
	                            ISNULL(APM.SWW2_FG, '') AS SWW2_FG,				-- 用地取得
	                            ISNULL(APM.IMPALLAREA, 0) AS IMPALLAREA,		-- 環境營造面積(HA)
	                            ISNULL(APM.PRTPOPS ,0) AS PRTPOPS ,				-- 受益人口
	                            APM.START_DT,									-- 開始時間
	                            APM.END_DT,										-- 完成時間
	                            CAST(ISNULL(TRY_CAST(NULLIF(LOCP.GIS_X, '') AS DECIMAL(18, 4)), 0) AS BIGINT) AS GIS_X,	-- 工程座標點位經度X座標
	                            CAST(ISNULL(TRY_CAST(NULLIF(LOCP.GIS_Y, '') AS DECIMAL(18, 4)), 0) AS BIGINT) AS GIS_Y,	-- 工程座標點位緯度Y座標
	                            ISNULL(APM.BENEFITS,'') AS BENEFITS,			-- 預期效益
	                            ISNULL(APM.AP_NOTE,'') AS AP_NOTE				-- 備註
                            FROM APMFORM APM
                            LEFT JOIN APDLOCP LOCP ON LOCP.AID = APM.AID
                            LEFT JOIN SMCCODE SMC ON SMC.CATY_COD = APM.PJ_CATY_CODE AND SMC.SYS_ID = 'TT' AND SMC.CATY_ID = '05'
                            OUTER APPLY (
                                SELECT
                                    SUM(ISNULL(FUND, 0)) AS M_FUND,
                                    SUM(CASE WHEN FIT_CODE = '00' THEN ISNULL(FUND, 0) ELSE 0 END) AS CEN_FUND,
                                    SUM(CASE WHEN FIT_CODE = '11' THEN ISNULL(FUND, 0) ELSE 0 END) AS LOC_FUND
                                FROM APDFUND FD WHERE FD.AID = APM.AID AND FD.SYR = APM.SYR) FD
                            WHERE APM.PPID IN ( @PPID ) AND APM.YRC_ID = @YRC_ID AND APM.AP_STAT = @AP_STAT ";

            SqlCommand cmd = new SqlCommand(sql);

            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            cmd.Parameters.Add(new SqlParameter("AP_STAT", APStat));

            if (!string.IsNullOrEmpty(m.keyword2))
            {
                sql += " AND APM.AP_NAME LIKE @KEYWORD ";
                cmd.Parameters.Add(new SqlParameter("KEYWORD", "%" + m.keyword2 + "%"));
            }

            if (!string.IsNullOrEmpty(m.city))
            {
                sql += " AND APM.CITY = @CITY ";
                cmd.Parameters.Add(new SqlParameter("CITY", m.city));
            }

            cmd.CommandText = sql;

            Debug.WriteLine($"PPID : {PPID}");
            Debug.WriteLine($"YRC_ID : {YRC_ID}");
            Debug.WriteLine($"APStat : {APStat}");
            Debug.WriteLine($"AP_NAME : {m.keyword2}");
            Debug.WriteLine($"sql = {sql}");

            return BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
        }

        /// <summary>
        /// 永續經營組 - 明細(概述)表[Sheet1]
        /// </summary>
        /// <param name="WorkBook">Excel Workbook</param>
        /// <param name="DT">資料來源</param>
        /// <param name="PJNAME">計畫名稱</param>
        /// <param name="m">查詢條件</param>
        private void WriteEnvironmentDetailsSheet(Workbook WorkBook, DataTable DT, string PJNAME, SGE0010Model m)
        {
            int StartRow = 4;   /// 資料開始列
            int TempRow = 4;    /// Excel 第 5 列
            int StartCol = 0;   /// A  欄
            int EndCol = 32;    /// AG 欄，如果你的樣板只到其他欄位，這裡再調整

            Worksheet ws = WorkBook.Worksheets[0];

            ws.Cells["A1"].PutValue(PJNAME + " 概述表");

            for (int i = 0; i < DT.Rows.Count; i++)
            {
                int CurrentRow = StartRow + i;

                ws.Cells.CopyRows(ws.Cells, TempRow, CurrentRow, 1);    /// 複製第 5 列樣板格式
                ws.Cells.ClearContents(CurrentRow, StartCol, CurrentRow, EndCol);

                DataRow Dr = DT.Rows[i];

                /// Step 2: 填入資料
                ws.Cells[CurrentRow, 0].PutValue(Convert.ToDouble(Dr["S1_ORDER"].ToString()));                          /// A:縣市自評優先順序
                ws.Cells[CurrentRow, 1].PutValue(Dr["CITY"].ToString());                                                /// B:縣市名稱
                ws.Cells[CurrentRow, 2].PutValue(Dr["EXE_MCH"].ToString());                                             /// C:執行機關
                ws.Cells[CurrentRow, 3].PutValue(Dr["AP_PJ_NAME"].ToString());                                          /// D:整體計畫名稱
                ws.Cells[CurrentRow, 4].PutValue(Dr["AP_NAME"].ToString());                                             /// E:分項案件名稱
                ws.Cells[CurrentRow, 5].PutValue(Dr["CONT_NAME"].ToString());                                           /// F:中央主管機關
                ws.Cells[CurrentRow, 6].PutValue(Dr["TOWN"].ToString());                                                /// G:鄉鎮市區別
                ws.Cells[CurrentRow, 7].PutValue(Dr["CODE_DESC"].ToString());                                           /// H:工作項目分類
                ws.Cells[CurrentRow, 8].PutValue(Dr["AP_ALL_NOTE"].ToString());                                        /// I:工程內容
                ws.Cells[CurrentRow, 9].PutValue(Convert.ToDouble(Dr["M_T_FUND"].ToString()));                          /// J:總經費需求(仟元)
                ws.Cells[CurrentRow, 10].PutValue(Convert.ToDouble(Dr["M_R_FUND"].ToString()));                         /// K:過去年度已支用金額
                ws.Cells[CurrentRow, 11].PutValue(Convert.ToDouble(Dr["CEN_FUND"].ToString()));                         /// L:115年度
                ws.Cells[CurrentRow, 12].PutValue(Convert.ToDouble(Dr["FCEN_FUND"].ToString()));                        /// M:以後年度
                ws.Cells[CurrentRow, 13].PutValue(Convert.ToDouble(Dr["TOTAL_CEN_FUND"].ToString()));                   /// N:合計
                ws.Cells[CurrentRow, 14].PutValue(Convert.ToDouble(Dr["LOC_FUND"].ToString()));                         /// O:115年度
                ws.Cells[CurrentRow, 15].PutValue(Convert.ToDouble(Dr["FLOC_FUND"].ToString()));                        /// P:以後年度
                ws.Cells[CurrentRow, 16].PutValue(Convert.ToDouble(Dr["TOTAL_LOC_FUND"].ToString()));                   /// Q:合計
                ws.Cells[CurrentRow, 17].PutValue(Dr["SWW1_FG"].ToString() == "0" ? "否" : "是");                       /// R:是否完成
                ws.Cells[CurrentRow, 18].PutValue(Dr["SWWOT_DESC"].ToString());                                         /// S:說明(延續/檢討修正)
                ws.Cells[CurrentRow, 19].PutValue(Dr["SWW2_FG"].ToString() == "0" ? "否" : "是");                       /// T:用地是否取得
                ws.Cells[CurrentRow, 20].PutValue(Convert.ToDouble(Dr["IMPALLAREA"].ToString()));                       /// U:環境營造面積(HA)
                ws.Cells[CurrentRow, 21].PutValue(Convert.ToInt32(Dr["PRTPOPS"].ToString()));                           /// V:受益人口
                ws.Cells[CurrentRow, 22].PutValue(Dr["START_DT"] == DBNull.Value ? "" : $"{Convert.ToDateTime(Dr["START_DT"]).Year - 1911}.{Convert.ToDateTime(Dr["START_DT"]).Month:00}.{Convert.ToDateTime(Dr["START_DT"]).Day:00}"); /// W:開始時間
                ws.Cells[CurrentRow, 23].PutValue(Dr["END_DT"] == DBNull.Value ? "" : $"{Convert.ToDateTime(Dr["END_DT"]).Year - 1911}.{Convert.ToDateTime(Dr["END_DT"]).Month:00}.{Convert.ToDateTime(Dr["END_DT"]).Day:00}");         /// X:完成時間
                ws.Cells[CurrentRow, 24].PutValue(Dr["GIS_X"].ToString());                             /// Y:工程座標點位經度X座標
                ws.Cells[CurrentRow, 25].PutValue(Dr["GIS_Y"].ToString());                             /// Z:工程座標點位緯度Y座標
                ws.Cells[CurrentRow, 26].PutValue(Dr["BENEFITS"] == DBNull.Value ? "" : Dr["BENEFITS"].ToString());     /// AA:預期效益
                ws.Cells[CurrentRow, 27].PutValue(Dr["AP_NOTE"] == DBNull.Value ? "" : Dr["AP_NOTE"].ToString());       /// AB:備註
            }

            int TotalRow = StartRow + DT.Rows.Count;

            ws.Cells.CopyRows(ws.Cells, TempRow, TotalRow, 1);
            ws.Cells.ClearContents(TotalRow, StartCol, TotalRow, EndCol);

            ws.Cells[TotalRow, 1].PutValue("合計");
            ws.Cells[TotalRow, 9].Formula = $"=SUM(J{StartRow + 1}:J{TotalRow})";   // J:總經費需求
            ws.Cells[TotalRow, 10].Formula = $"=SUM(K{StartRow + 1}:K{TotalRow})";  // K:過去年度已支用金額
            ws.Cells[TotalRow, 11].Formula = $"=SUM(L{StartRow + 1}:L{TotalRow})";  // L:中央補助115年度
            ws.Cells[TotalRow, 12].Formula = $"=SUM(M{StartRow + 1}:M{TotalRow})";  // M:中央補助以後年度
            ws.Cells[TotalRow, 13].Formula = $"=SUM(N{StartRow + 1}:N{TotalRow})";  // N:中央補助合計
            ws.Cells[TotalRow, 14].Formula = $"=SUM(O{StartRow + 1}:O{TotalRow})";  // O:地方分擔115年度
            ws.Cells[TotalRow, 15].Formula = $"=SUM(P{StartRow + 1}:P{TotalRow})";  // P:地方分擔以後年度
            ws.Cells[TotalRow, 16].Formula = $"=SUM(Q{StartRow + 1}:Q{TotalRow})";  // Q:地方分擔合計
            ws.Cells[TotalRow, 20].Formula = $"=SUM(U{StartRow + 1}:U{TotalRow})";  // U:環境營造面積
            ws.Cells[TotalRow, 21].Formula = $"=SUM(V{StartRow + 1}:V{TotalRow})";  // V:受益人口

            /// 手動設定欄寬
            ws.Cells.SetColumnWidth(0, 6);    // A 項次
            ws.Cells.SetColumnWidth(1, 10);   // B 工程名稱
            ws.Cells.SetColumnWidth(2, 20);   // C 全期內容
            ws.Cells.SetColumnWidth(3, 35);   // D 本年度辦理內容
            ws.Cells.SetColumnWidth(4, 60);   // E 工程總經費
            ws.Cells.SetColumnWidth(5, 20);   // F 以前年度已執行經費
            ws.Cells.SetColumnWidth(6, 20);   // G 本年度申請經費
            ws.Cells.SetColumnWidth(7, 20);   // H 中央補助款
            ws.Cells.SetColumnWidth(8, 60);   // I 地方配合款
            ws.Cells.SetColumnWidth(9, 18);   // J
            ws.Cells.SetColumnWidth(10, 18);  // K
            ws.Cells.SetColumnWidth(11, 18);  // L
            ws.Cells.SetColumnWidth(12, 18);  // M
            ws.Cells.SetColumnWidth(13, 18);  // N
            ws.Cells.SetColumnWidth(14, 18);  // O
            ws.Cells.SetColumnWidth(15, 18);  // P
            ws.Cells.SetColumnWidth(16, 18);  // Q
            ws.Cells.SetColumnWidth(17, 18);  // R
            ws.Cells.SetColumnWidth(18, 18);  // S
            ws.Cells.SetColumnWidth(19, 16);  // T
            ws.Cells.SetColumnWidth(20, 16);  // U
            ws.Cells.SetColumnWidth(21, 16);  // V
            ws.Cells.SetColumnWidth(22, 16);  // W
            ws.Cells.SetColumnWidth(23, 16);   // X 是
            ws.Cells.SetColumnWidth(24, 16);   // Y 否
            ws.Cells.SetColumnWidth(25, 16);  // Z
            ws.Cells.SetColumnWidth(26, 60);  // AA
            ws.Cells.SetColumnWidth(27, 60);  // AB

            /// 手動設定列高
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                int CurrentRow = StartRow + i;

                ws.Cells.SetRowHeight(CurrentRow, 45);
            }
            ws.Cells.SetRowHeight(TotalRow, 45);
        }

        /// <summary>
        /// 永續經營組 - 明細(概述)表[Excel]
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ExportEnvironmentDetailsToExcel")]
        public ActionResult ExportEnvironmentDetailsToExcel(SGE0010Model m)
        {
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0060/PJV0053.xlsx"));

            /// Step 1: 報表標題
            var PID = m.PPID;
            PJMFORM_CMPSGS PJMFORM = sgsdb.PJMFORM.FirstOrDefault(X => X.PPID == PID);
            string PJNAME = PJMFORM.PJ_NAME;

            /// Step 1-1: 查詢 YRC_ID
            string sqlYrc = @" SELECT TOP 1 YRC_ID FROM PJDSYRC WHERE PPID = @PPID ORDER BY YRC_ID DESC ";

            SqlCommand cmdYrc = new SqlCommand(sqlYrc);
            cmdYrc.Parameters.Add(new SqlParameter("PPID", PID));

            DataTable dtYrc = SqlDao.CMPSGSQuery4List(cmdYrc);

            string YRC_ID = string.Empty;

            if (dtYrc.Rows.Count > 0)
            {
                YRC_ID = dtYrc.Rows[0]["YRC_ID"].ToString();
            }
            else
            {
                return Content("查無 YRC_ID");
            }

            /// Step 2: 根據目前頁籤取得 AP_STAT
            string APStat = GetApStatByTab(m.tab);

            /// Step 3: 查詢資料
            DataTable DT = GetEnvironmentDetailsData(PID, APStat, m, YRC_ID);

            /// Step 4: 寫入資料到 Excel
            WriteEnvironmentDetailsSheet(workbook, DT, PJNAME, m);

            Debug.WriteLine("Over View Data");
            foreach (DataRow rows in DT.Rows)
            {
                foreach (DataColumn cols in DT.Columns)
                {
                    Debug.Write($"{cols.ColumnName}:{rows[cols]} \t");
                }
                Debug.WriteLine("");
            }

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);

                string fileName = $"{PJNAME}明細表{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        #endregion
    }

}