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
using System.Web;
using System.Web.Hosting;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMPSGSModels;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS
{
    public class PJE0011Controller : Controller
    {
        private readonly Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();
        private readonly Models.CMPSGSModels.CMPSGSEntities sgsdb = new Models.CMPSGSModels.CMPSGSEntities();
        private readonly Models.CMBASEModels.CMBASEEntities cmbase = new Models.CMBASEModels.CMBASEEntities();

        int pageSize = 10;
        //查詢
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Search")]
        public ActionResult Search(PJE0011Model model)
        {

            string userdp = Helpers.WebSiteHelper.getUserDept();
            int page1 = (model.page1 == 0) ? 1 : model.page1;
            model.page1 = page1;
            if (model.PJMFORM == null)
            {
                model.PJMFORM = new Models.CMPSGSModels.PJMFORM_CMPSGS();
            }

            /// 20260424 修改版本 => 修改顯示資料取得方式(原本帳號改成用 PJCSGDP 裡 AP_STAT = 20 的帳號可看)
            /// StringBuilder sql = new StringBuilder("SELECT * FROM PJMFORM p WHERE p.ADV_DP = @DP ");
            StringBuilder sql = new StringBuilder("SELECT * FROM PJMFORM WHERE PJ_CODE IN (SELECT PJ_CODE FROM PJCSGDP WHERE AP_STAT = '20' AND DEPT_NO = @DP )");

            if (!string.IsNullOrEmpty(model.keyword))
            {
                sql.Append("AND PJ_NAME LIKE'%" + model.keyword + "%'");

            }
            SqlCommand cmd = new SqlCommand(sql.ToString());
            cmd.Parameters.Add(new SqlParameter("DP", userdp));
            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page1, 5);
            model.MainList = plist;
            return View("PJE0011Index", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CheckMyProposal")]
        public ActionResult CheckMyProposal(PJE0011Model model)
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
            StringBuilder sql = new StringBuilder("UPDATE APMFORM SET AP_STAT='00' WHERE AP_STAT='-1' AND PPID=1 ");
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

        public ActionResult Index(PJE0011Model model)
        {
            string userdp = Helpers.WebSiteHelper.getUserDept();
            int page1 = (model.page1 == 0) ? 1 : model.page1;
            model.page1 = page1;
            if (model.PJMFORM == null)
            {
                model.PJMFORM = new Models.CMPSGSModels.PJMFORM_CMPSGS();
            }

            /// 20260424 修改版本 => 修改顯示資料取得方式(原本帳號改成用 PJCSGDP 裡 AP_STAT = 20 的帳號可看)
            /// StringBuilder sql = new StringBuilder("SELECT*FROM PJMFORM p WHERE p.ADV_DP=@DP ");
            StringBuilder sql = new StringBuilder("SELECT * FROM PJMFORM WHERE PJ_CODE IN (SELECT PJ_CODE FROM PJCSGDP WHERE AP_STAT = '20' AND DEPT_NO = @DP)");

            if (!string.IsNullOrEmpty(model.keyword))
            {
                sql.Append("AND PJ_NAME LIKE'%" + model.keyword + "%'");
            }
            SqlCommand cmd = new SqlCommand(sql.ToString());
            cmd.Parameters.Add(new SqlParameter("DP", userdp));
            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page1, 5);
            model.MainList = plist;
            return View("PJE0011Index", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ProposalSearch")]
        public ActionResult ProposalSearch(PJE0011Model model)
        {
            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
            int page1 = (model.page1 == 0) ? 1 : model.page1;
            int page2 = 1;// (model.page2 == 0) ? 1 : model.page2;
            DataTable yrcs = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC WHERE PPID=" + model.PPID + " ORDER BY YRC_ID DESC");
            ViewBag.yrcs = tools.ToSelectList(yrcs, "YRC_ID", "PSG_NAME", null);
            DataTable ps = util.GetAPMFORMByEzway(model.PPID, "00", model.YRC_ID);
            List<DataRow> list = ps.AsEnumerable().ToList();
            if (model.city != null)
            {
                list = ps.AsEnumerable().Where(r => string.IsNullOrEmpty(model.city) || r.Field<string>("CITY") == model.city).ToList();
            }
            // 2. 【新增】：工程名稱 (AP_NAME) 的模糊查詢過濾
            if (!string.IsNullOrEmpty(model.keyword2))
            {
                // 確保欄位不是 Null，並且包含使用者輸入的關鍵字
                list = list.Where(r => r.Field<string>("AP_NAME") != null &&
                                       r.Field<string>("AP_NAME").Contains(model.keyword2)).ToList();
            }
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page2, pageSize);
            model.tab = 1;
            model.List1 = plist;
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == model.PPID).FirstOrDefault();
            ViewBag.PJ_NAME = pjmform.PJ_NAME;
            ViewBag.STAT_NO = pjmform.STAT_NO;
            return View("PJE0011ProposalManage", model);
        }

        public ActionResult ProposalManage(int tab, string ppid, int page1, int page2, string yrcid, string city = null, string keyword2 = null)
        {
            //要求加顯示
            decimal id = Convert.ToDecimal(ppid);
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == id).FirstOrDefault();
            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
            PJDSYRC pjdsyrc = sgsdb.PJDSYRC.Where(w => w.PPID == id).FirstOrDefault();
            ViewBag.PJ_NAME = pjmform.PJ_NAME;
            ViewBag.PSG_NAME = pjdsyrc.PSG_NAME;
            ViewBag.STAT_NO = pjmform.STAT_NO;


            string userid = WebSiteHelper.getUserId();
            string dept = WebSiteHelper.getUserDept();
            PJE0011Model model = new PJE0011Model();
            model.YRC_ID = yrcid;
            model.page2 = (model.page2 == 0) ? 1 : model.page2; ;
            model.page1 = (model.page1 == 0) ? 1 : model.page1; ;
            model.tab = tab;
            model.city = city;
            model.keyword2 = keyword2;
            model.PPID = tools.safe2Decimal(ppid);
            DataTable yrcs = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC WHERE PPID=" + model.PPID + " ORDER BY YRC_ID DESC");
            ViewBag.yrcs = tools.ToSelectList(yrcs, "YRC_ID", "PSG_NAME", null);
            DataTable ps = null;
            List<DataRow> list;
            IPagedList<DataRow> plist;
            //00 待收件
            if (tab == 1)
            {
                ps = util.GetAPMFORMByEzway(ppid, "00", yrcid);
                list = ps.AsEnumerable().ToList();
                if (!string.IsNullOrEmpty(city))
                {
                    list = list.Where(r => r.Field<string>("CITY") == city).ToList();
                }
                plist = new PagedList<DataRow>(list, page2, pageSize);
                model.List1 = plist;
            }
            //2.1.1 鎖到部門
            //Z1: 不使用 GetAPMFORMByDeptSP
            //Z2: 不使用 GetAPMFORMByDeptSP
            //ZZ: GetAPMFORMByDeptSP(AP_STAT = 'ZZ', NAP_DTAT = '20')
            //初審中10
            if (tab == 2)
            {
                ps = util.GetAPMFORMByEzway(ppid, "10", yrcid);
            }
            //15 複審中 
            if (tab == 3)
            {
                ps = util.GetAPMFORMByEzway(model.PPID, "15", yrcid);
            }
            // 核定20
            if (tab == 4)
            {
                ps = util.GetAPMFORMByEzway(model.PPID, "20", yrcid);
            }
            //成案
            if (tab == 5)
            {
                ps = util.GetAPMFORMByEzway(model.PPID, "90", yrcid);
            }

            //未通過初審
            if (tab == 6)
            {
                ps = util.GetAPMFORMByEzway(model.PPID, "Z1", yrcid);
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

            return View("PJE0011ProposalManage", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ProposalManage")]
        public ActionResult ProposalManage(PJE0011Model model)
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
            ViewBag.STAT_NO = pjmform.STAT_NO;
            return View("PJE0011ProposalManage", model);
        }

        //PJDCATY
        public ActionResult EditPJDCATY(PJE0011Model model)
        {

            ViewBag.PJDCATY = (from pj in sgsdb.PJDCATY where pj.PPID == model.PPID select pj).ToList();
            return View("EditPJDCATY", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SavePJDCATY")]
        public ActionResult SavePJDCATY(PJE0011Model model)
        {
            PJDCATY_CMPSGS p;
            PJMFORM_CMPSGS m = sgsdb.PJMFORM.Where(x => x.PPID == model.PPID).FirstOrDefault();
            decimal cid;
            DateTime now = DateTime.Now;
            for (int i = 0; i < model.PJDCATY_CID.Length; i++)
            {
                if (string.IsNullOrEmpty(model.PJDCATY_CATY_NAME[i]))
                {
                    continue;
                }
                if (string.IsNullOrEmpty(model.PJDCATY_CID[i]))
                {
                    p = new PJDCATY_CMPSGS();
                }
                else
                {
                    cid = tools.safe2Decimal(model.PJDCATY_CID[i]);
                    p = sgsdb.PJDCATY.Where(x => x.CID == cid).FirstOrDefault();
                }



                p.PPID = model.PPID;
                p.CATY_CODE = model.PJDCATY_CATY_CODE[i];
                p.CATY_NAME = model.PJDCATY_CATY_NAME[i];
                if (model.PJDCATY_EVRF_FG[i].Equals("1"))
                {
                    p.EVRF_FG = true;
                }
                DateTime temp;
                p.OTDSG_DT = DateTime.TryParse(model.PJDCATY_OTDSG_DT[i], out temp) ? (DateTime?)temp : null;
                p.OTDSG_BUSS = model.PJDCATY_OTDSG_BUSS[i];
                p.DSGRL_DT = DateTime.TryParse(model.PJDCATY_DSGRL_DT[i], out temp) ? (DateTime?)temp : null;
                p.PCCJDG_DT = DateTime.TryParse(model.PJDCATY_PCCJDG_DT[i], out temp) ? (DateTime?)temp : null;
                p.PCCJDG_FUND = tools.safe2Decimal(model.PJDCATY_PCCJDG_FUND[i]);
                p.PJ_CATY_CODE = m.PJ_CATY;
                p.UPD_DT = now;
                p.UPD_UID = WebSiteHelper.getUserId();
                if (string.IsNullOrEmpty(model.PJDCATY_CID[i]))
                {
                    p.ADD_DT = now;
                    sgsdb.PJDCATY.Add(p);
                }
            }
            sgsdb.SaveChanges();
            ViewBag.msg = "已儲存";
            return EditPJDCATY(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "DelPJDCATY")]
        public ActionResult DelPJDCATY(PJE0011Model model)
        {

            BASE_DAO.SqlDao.CMPSGSExecuteSql("DELETE FROM PJDCATY WHERE CID='" + model.DeleteId1 + "'");

            ViewBag.msg = "已刪除";
            return EditPJDCATY(model);
        }

        //PJDCMCT
        public ActionResult EditPJDCMCT(PJE0011Model model)
        {

            ViewBag.PJDCMCT = (from pj in sgsdb.PJDCMCT where pj.PPID == model.PPID select pj).ToList();
            //model.PPID = ppid;            
            return View("EditPJDCMCT", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SavePJDCMCT")]
        public ActionResult SavePJDCMCT(PJE0011Model model)
        {

            PJDCMCT p;
            string cmscode;
            decimal ppid;
            for (int i = 0; i < model.PJDCMCT_CMS_CODE.Length; i++)
            {
                if (string.IsNullOrEmpty(model.PJDCMCT_CMS_CODE[i]) && string.IsNullOrEmpty(model.PJDCMCT_CMS_NAME[i]))
                {
                    continue;
                }
                try
                {
                    cmscode = model.PJDCMCT_CMS_CODE[i];
                    ppid = model.PPID;
                    p = sgsdb.PJDCMCT.Where(w => w.PPID == ppid && w.CMS_CODE == cmscode).FirstOrDefault();
                    p.CMS_CODE = model.PJDCMCT_CMS_CODE[i];
                    p.CMS_NAME = model.PJDCMCT_CMS_NAME[i];
                    p.CMS_NOTE = model.PJDCMCT_CMS_NOTE[i];
                    p.CMS_UNIT = model.PJDCMCT_CMS_UNIT[i];

                }
                catch
                {
                    p = new PJDCMCT();
                    p.PPID = model.PPID;
                    p.CMS_CODE = model.PJDCMCT_CMS_CODE[i];
                    p.CMS_NAME = model.PJDCMCT_CMS_NAME[i];
                    p.CMS_NOTE = model.PJDCMCT_CMS_NOTE[i];
                    p.CMS_UNIT = model.PJDCMCT_CMS_UNIT[i];
                    sgsdb.PJDCMCT.Add(p);

                }
            }
            sgsdb.SaveChanges();
            ViewBag.msg = "已儲存";
            return EditPJDCMCT(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "DelPJDCMCT")]
        public ActionResult DelPJDCMCT(PJE0011Model model)
        {

            BASE_DAO.SqlDao.CMPSGSExecuteSql("DELETE FROM PJDCMCT WHERE PPID='" + model.DeleteId1 + "'AND CMS_CODE='" +
                model.DeleteId2 + "'");

            ViewBag.msg = "已刪除";
            return EditPJDCMCT(model);
        }

        //PJDFUND
        public ActionResult EditPJDFUND(PJE0011Model model)
        {
            ViewBag.PJDFUND = (from pj in sgsdb.PJDFUND where pj.PPID == model.PPID select pj).ToList();
            //經費層級
            DataTable f = util.GetCMEPMSCODES("SM", "02");
            //ViewBag.IPJ_CODE = tools.ToSelectListNotItem(f, "id", "name", null);
            // new SelectList( dt.AsDataView( ) , "stor_id" , "stor_name" ); 
            ViewBag.IPJ_CODE = util.GetCMEPMSCODES("SM", "02");
            model.PPID = model.PPID;
            return View("EditPJDFUND", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "DelPJDFUND")]
        public ActionResult DelPJDFUND(PJE0011Model model)
        {
            string sql;
            for (int i = 0; i < model.DeleteIds.Length; i++)
            {
                if (!string.IsNullOrEmpty(model.DeleteIds[i]))
                {
                    sql = "DELETE FROM PJDFUND WHERE PPID=" + model.PPID + " AND SYR='" +
                        model.PJDFUND_SYR[i] + "'AND IPJ_CODE='" + model.PJDFUND_IPJ_CODE[i] + "'";
                    SqlDao.CMPSGSExecuteSql(sql);
                }
            }
            ViewBag.msg = "已刪除";
            return EditPJDFUND(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SavePJDFUND")]
        public ActionResult SavePJDFUND(PJE0011Model model)
        {
            DateTime now = DateTime.Now;

            decimal ppid = model.PPID;
            string syr, ipj;
            PJDFUND p;
            for (int i = 0; i < model.PJDFUND_FUND.Length; i++)
            {

                if (string.IsNullOrEmpty(model.PJDFUND_FUND[i]) ||
                   string.IsNullOrEmpty(model.PJDFUND_SYR[i]))
                {
                    continue;
                }
                ipj = model.PJDFUND_IPJ_CODE[i];
                syr = model.PJDFUND_SYR[i];
                //APDFILE file = sgsdb.APDFILE.Where(w => w.FID == fid && w.AID == aid).FirstOrDefault();
                //p = sgsdb.PJDFUND.Where(w =>  w.PPID == ppid && w.SYR == syr && w.IPJ_CODE == ipj).FirstOrDefault();
                SqlDao.CMPSGSExecuteSql("DELETE FROM PJDFUND WHERE PPID=" +
                    model.PPID + " AND IPJ_CODE='" + ipj + "'AND SYR='" + syr + "'");

                p = new PJDFUND();
                p.PPID = model.PPID;
                p.IPJ_CODE = model.PJDFUND_IPJ_CODE[i];
                p.SYR = model.PJDFUND_SYR[i];
                p.FUND = tools.safe2Decimal(model.PJDFUND_FUND[i]);
                p.FUND_DESC = model.PJDFUND_FUND_DESC[i];
                p.YBG_FUND = tools.safe2Decimal(model.PJDFUND_YBG_FUND[i]);
                p.UPD_DT = now;
                p.ADD_DT = now;
                p.UPD_UID = WebSiteHelper.getUserId();
                sgsdb.PJDFUND.Add(p);

                try
                {
                    sgsdb.PJDFUND.Add(p);
                    sgsdb.SaveChanges();
                }
                catch (Exception e)
                {
                    ViewBag.msg = "儲存失敗" + e.Message;
                    return EditPJDFUND(model);
                }


            }

            ViewBag.msg = "已儲存";
            return EditPJDFUND(model);
        }

        //EditPJDDOCS
        public ActionResult EditPJDDOCS(PJE0011Model model)
        {
            ViewBag.PJDDOCS = (from pj in sgsdb.PJDDOCS where pj.PPID == model.PPID select pj).ToList();
            return View("EditPJDDOCS", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SavePJDDOCS")]
        public ActionResult SavePJDDOCS(PJE0011Model model)
        {

            PJDDOCS doc;
            decimal fid;
            MemoryStream memoryStream;
            byte[] data;
            for (int i = 0; i < model.FID.Length; i++)
            {

                DateTime now = DateTime.Now;
                if (!string.IsNullOrEmpty(model.FID[i]))
                {
                    //舊有
                    fid = Convert.ToDecimal(model.FID[i]);
                    doc = sgsdb.PJDDOCS.Where(x => x.FID == fid).FirstOrDefault();
                    if (model.files[i] != null)
                    {
                        //新檔案
                        using (Stream inputStream = model.files[i].InputStream)
                        {
                            memoryStream = inputStream as MemoryStream;
                            if (memoryStream == null)
                            {
                                memoryStream = new MemoryStream();
                                inputStream.CopyTo(memoryStream);
                            }
                            data = memoryStream.ToArray();
                            doc.FILE_BODY = data;
                            doc.FILE_NAME = model.files[i].FileName;
                            doc.FILE_PATH = model.files[i].FileName;
                            doc.FILE_SIZE = data.Length;
                            doc.PPID = model.PPID;
                        }
                    }
                    doc.UPD_UID = WebSiteHelper.getUserId();
                    doc.ADD_DT = now;
                    doc.FILE_DESC = model.FILE_DESC[i];
                    sgsdb.SaveChanges();
                }
                else
                {
                    //新建
                    if (model.files[i] == null)
                    {
                        continue;
                    }
                    doc = new PJDDOCS();
                    if (model.files[i] != null)
                    {
                        //新檔案
                        using (Stream inputStream = model.files[i].InputStream)
                        {
                            memoryStream = inputStream as MemoryStream;
                            if (memoryStream == null)
                            {
                                memoryStream = new MemoryStream();
                                inputStream.CopyTo(memoryStream);
                            }
                            data = memoryStream.ToArray();
                            doc.FILE_BODY = data;
                            doc.FILE_NAME = model.files[i].FileName;
                            doc.FILE_PATH = model.files[i].FileName;
                            doc.FILE_SIZE = data.Length;
                        }
                        doc.PPID = model.PPID;
                        doc.UPD_UID = WebSiteHelper.getUserId();
                        doc.ADD_DT = now;
                        doc.FILE_DESC = model.FILE_DESC[i];
                        //doc.FID = null;
                        sgsdb.PJDDOCS.Add(doc);
                        sgsdb.SaveChanges();
                    }
                }
            }
            ViewBag.msg = "已儲存";
            return EditPJDDOCS(model);
        }

        public ActionResult DownloadAPDFILE(decimal fid)
        {
            PJDDOCS file = sgsdb.PJDDOCS.Where(w => w.FID == fid).FirstOrDefault();
            if (file == null)
            {
                return HttpNotFound();
            }
            var contentType = MimeMapping.GetMimeMapping(file.FILE_PATH);
            return File(file.FILE_BODY, contentType);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "DelPJDDOCS")]
        public ActionResult DelPJDDOCS(PJE0011Model model)
        {

            SqlCommand cmd = new SqlCommand("DELETE FROM PJDDOCS WHERE FID=@FID");
            cmd.Parameters.Add(new SqlParameter("FID", model.DeleteId1));

            BASE_DAO.SqlDao.CMESGSExecuteSql(cmd);
            ViewBag.msg = "已刪除";
            return EditPJDDOCS(model);
        }

        /// <summary>
        /// PJDSUPD
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult EditPJDSUPD(PJE0011Model model)
        {
            //右邊
            ViewBag.PJDSUPD = SqlDao.CMPSGSQuery4List("SELECT d.DEPT_NAME, d.DEPT_NO FROM PJDSUPD p, CMBASE.dbo.OGCDEPT d WHERE p.ADV_DP=d.DEPT_NO AND p.PPID='" + model.PPID + "'");
            //ViewBag.PJDSYRC = (from pj in sgsdb.PJDSYRC where pj.PPID == model.PPID select pj).ToList();
            ViewBag.deptV1 = SqlDao.CMBASEQuery4List("SELECT*FROM CMBASE.dbo.OGCDEPT WHERE AN_DEPT_NO IS NULL ORDER BY DEPT_NAME");
            ViewBag.depts = SqlDao.CMBASEQuery4List("SELECT*FROM CMBASE.dbo.OGCDEPT ORDER BY DEPT_NAME");

            return View("EditPJDSUPD", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SavePJDSUPD")]
        public ActionResult SavePJDSUPD(PJE0011Model model)
        {
            SqlCommand cmd = new SqlCommand("DELETE FROM PJDSUPD WHERE PPID=@PPID");
            cmd.Parameters.Add(new SqlParameter("PPID", model.PPID));
            BASE_DAO.SqlDao.CMESGSExecuteSql(cmd);

            PJDSUPD p;
            foreach (string d in model.PJDSUPD_DEPT_NO)
            {
                if (!d.Equals(""))
                {
                    p = new PJDSUPD();
                    p.PPID = model.PPID;
                    p.ADV_DP = d;
                    p.ADD_DT = DateTime.Now;
                    p.UPD_UID = WebSiteHelper.getUserId();
                    sgsdb.PJDSUPD.Add(p);
                }

            }
            sgsdb.SaveChanges();
            ViewBag.msg = "已儲存";
            return EditPJDSUPD(model);
        }

        /// <summary>
        /// PJDSYRC
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult EditPJDSYRC(PJE0011Model model)
        {
            ViewBag.PJDSUPD = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSUPD WHERE PPID='" + model.PPID + "'");
            ViewBag.PJDSYRC = (from pj in sgsdb.PJDSYRC where pj.PPID == model.PPID select pj).ToList();
            ViewBag.deptV1 = SqlDao.CMBASEQuery4List("SELECT*FROM OGCDEPT");
            ViewBag.PJDCMCT = (from pj in sgsdb.PJDCMCT where pj.PPID == model.PPID select pj).ToList();
            return View("EditPJDSYRC", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "OpenPJDSYRC")]
        public ActionResult OpenPJDSYRC(PJE0011Model model)
        {
            SqlCommand cmd = new SqlCommand("UPDATE PJDSYRC SET OPEN_FG=1 WHERE YRC_ID=@YRC_ID");
            cmd.Parameters.Add(new SqlParameter("YRC_ID", model.DeleteId1));

            BASE_DAO.SqlDao.CMESGSExecuteSql(cmd);
            ViewBag.msg = "已儲存";
            return EditPJDSYRC(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ClosePJDSYRC")]
        public ActionResult ClosePJDSYRC(PJE0011Model model)
        {
            SqlCommand cmd = new SqlCommand("UPDATE PJDSYRC SET OPEN_FG=0 WHERE YRC_ID=@YRC_ID");
            cmd.Parameters.Add(new SqlParameter("YRC_ID", model.DeleteId1));

            BASE_DAO.SqlDao.CMESGSExecuteSql(cmd);
            ViewBag.msg = "已儲存";
            return EditPJDSYRC(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "DelPJDSYRC")]
        public ActionResult DelPJDSYRC(PJE0011Model model)
        {
            if (!decimal.TryParse(model.DeleteId1, out decimal deleteId))
            {
                ViewBag.msg = "無法刪除資料";
                return EditPJDSYRC(model);
            }
            var isUsed = sgsdb.APMFORM.Any(a => a.YRC_ID == deleteId);
            if (isUsed)
            {
                ViewBag.msg = "無法刪除資料，已有提案使用此來源";
                return EditPJDSYRC(model);
            }

            SqlCommand cmd = new SqlCommand("DELETE FROM PJDSYRC WHERE YRC_ID=@YRC_ID");
            cmd.Parameters.Add(new SqlParameter("YRC_ID", model.DeleteId1));

            BASE_DAO.SqlDao.CMESGSExecuteSql(cmd);
            ViewBag.msg = "已刪除";
            return EditPJDSYRC(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SavePJDSYRC")]
        public ActionResult SavePJDSYRC(PJE0011Model model)
        {
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == model.PPID).FirstOrDefault();
            for (int i = 0; i < model.PJDSYRC_PSG_NAME.Length; i++)
            {
                if (!model.PJDSYRC_PSG_NAME[i].Equals("") ||
                    !model.PJDSYRC_BGN_SYR[i].Equals("") ||
                    !model.PJDSYRC_END_SYR[i].Equals(""))
                {

                    PJDSYRC p = new PJDSYRC();
                    p.PSG_NAME = model.PJDSYRC_PSG_NAME[i];
                    p.PPID = model.PPID;
                    p.ADV_DP = WebSiteHelper.getUserDept();
                    p.ADD_DT = DateTime.Now;
                    p.UPD_UID = WebSiteHelper.getUserId();
                    p.SYR = model.PJDSYRC_BGN_SYR[i];
                    p.BGN_SYR = model.PJDSYRC_BGN_SYR[i];
                    p.END_SYR = model.PJDSYRC_END_SYR[i];
                    p.OPEN_FG = true;
                    sgsdb.PJDSYRC.Add(p);
                }
                sgsdb.SaveChanges();
            }

            ViewBag.msg = "已儲存";
            return EditPJDSYRC(model);
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Edit")]
        public ActionResult Edit(PJE0011Model model)
        {
            int page2 = (model.page == 0) ? 1 : model.page2;


            /*
            DataTable pjcatys = new DataTable();
            pjcatys.Columns.Add("id");
            pjcatys.Columns.Add("name");
            DataRow dr = pjcatys.NewRow();
            dr["id"] = "R"; dr["name"] = "R";
            pjcatys.Rows.Add(dr);
            dr = pjcatys.NewRow();
            dr["id"] = "C"; dr["name"] = "C";
            pjcatys.Rows.Add(dr);
            dr = pjcatys.NewRow();
            dr["id"] = "W"; dr["name"] = "W";
            pjcatys.Rows.Add(dr);
            ViewBag.PJ_CATYS = tools.ToSelectList(pjcatys, "id", "name", "");

            ViewBag.depts = tools.ToSelectList(BASE_DAO.SqlDao.CMPSGSQuery4List("SELECT * FROM OGCDEPT o WHERE o.AN_DEPT_NO='"+WebSiteHelper.getUserDept()+"'"), "DEPT_NAME", "DEPT_NAME", "");
            ViewBag.users = tools.ToSelectList(BASE_DAO.SqlDao.CMPSGSQuery4List("SELECT u.USR_ID, u.USR_NAME FROM OGCUSER u, OGCDEPT d, OGCDPUR p WHERE d.DEPT_NO='" + WebSiteHelper.getUserDept() + "' AND p.USR_ID=u.USR_ID AND d.DEPT_NO=p.DEPT_NO"), "DEPT_NAME", "DEPT_NAME", "");
            model.PJMFORM = sgsdb.PJMFORM.Where(p => p.PPID==model.PPID).FirstOrDefault();
            ViewBag.PJDCATY = sgsdb.PJDCATY.Where(p=>p.PPID==model.PPID).ToList();
            ViewBag.PJDCMCT= sgsdb.PJDCMCT.Where(p => p.PPID == model.PPID).ToList();
            ViewBag.PJDFUND = sgsdb.PJDFUND.Where(p => p.PPID == model.PPID).ToList();
            ViewBag.PJDDOCS = sgsdb.PJDDOCS.Where(p => p.PPID == model.PPID).ToList();
            ViewBag.PJDSYRC = sgsdb.PJDSYRC.Where(p => p.PPID == model.PPID).ToList();
            ViewBag.PJDSUPD = BASE_DAO.SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSUPD WHERE PPID="+model.PPID);
            ViewBag.deptV1 = SqlDao.CMBASEQuery4List("SELECT*FROM OGCDEPT");
            */


            model.PJMFORM = sgsdb.PJMFORM.Where(p => p.PPID == model.PPID).FirstOrDefault();
            //ui
            ViewBag.depts = new SelectList(cmbase.OGCDEPT.AsNoTracking().ToList(), "DEPT_NO", "DEPT_NAME");
            ViewBag.users = new SelectList(cmbase.OGCUSER.AsNoTracking().ToList(), "USR_ID", "USR_NAME");
            List<SelectListItem> PJ_CATYS = new List<SelectListItem>();
            PJ_CATYS.Add(new SelectListItem() { Text = "R", Value = "R", Selected = false });
            PJ_CATYS.Add(new SelectListItem() { Text = "C", Value = "C", Selected = false });
            PJ_CATYS.Add(new SelectListItem() { Text = "W", Value = "W", Selected = false });
            ViewBag.PJ_CATYS = new SelectList(PJ_CATYS, "Value", "Text");

            //ViewBag.PJDCATY = (from pj in sgsdb.PJDCATY where pj.PPID == model.PPID select pj).ToList();
            //model.PJDCATY = (from pj in sgsdb.PJDCATY where pj.PPID == model.PPID select pj).ToArray();
            //ViewBag.PJDCMCT = (from pj in sgsdb.PJDCMCT where pj.PPID == model.PPID select pj).ToList();

            ViewBag.PJDDOCS = (from pj in sgsdb.PJDDOCS where pj.PPID == model.PPID select pj).ToList();
            ViewBag.PJDSUPD = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSUPD WHERE PPID='" + model.PPID + "'");
            ViewBag.PJDSYRC = (from pj in sgsdb.PJDSYRC where pj.PPID == model.PPID select pj).ToList();
            ViewBag.deptV1 = SqlDao.CMBASEQuery4List("SELECT*FROM OGCDEPT");

            try { model.LEV1_DT = tools.GetDateByTWDate(model.PJMFORM.LEV1_DT.Value); } catch { }
            try { model.LEV2_DT = tools.GetDateByTWDate(model.PJMFORM.LEV2_DT.Value); } catch { }
            try { model.LEV3_DT = tools.GetDateByTWDate(model.PJMFORM.LEV3_DT.Value); } catch { }

            return View("PJE0011Edit", model);
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SaveOuther")]
        public ActionResult SaveOuther(PJE0011Model m)
        {


            return Edit(m);
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Save")]
        public ActionResult Save(PJE0011Model m)
        {
            PJMFORM_CMPSGS p = sgsdb.PJMFORM.Where(x => x.PPID == m.PPID).FirstOrDefault();
            p.PJ_NAME = m.PJMFORM.PJ_NAME;
            p.PJ_CODE = m.PJMFORM.PJ_CODE;
            p.PJ_CATY = m.PJMFORM.PJ_CATY;
            p.ADV_DP = m.PJMFORM.ADV_DP;
            p.ADV_UID = m.PJMFORM.ADV_UID;
            p.PJ_CONT = m.PJMFORM.PJ_CONT;
            p.STT_DT = m.PJMFORM.STT_DT;
            p.END_DT = m.PJMFORM.END_DT;

            try { p.LEV1_DT = tools.GetDateByTWDate(m.LEV1_DT); } catch { }
            p.LEV1_FN = m.PJMFORM.LEV1_FN;
            try { p.LEV2_DT = tools.GetDateByTWDate(m.LEV2_DT); } catch { }
            p.LEV2_FN = m.PJMFORM.LEV2_FN;
            try { p.LEV3_DT = tools.GetDateByTWDate(m.LEV3_DT); } catch { }
            p.LEV3_FN = m.PJMFORM.LEV3_FN;
            sgsdb.SaveChanges();
            ViewBag.msg = "儲存完成";
            return Edit(m);
        }

        /// <summary>
        /// 送件
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendForward")]
        public ActionResult SendForward(PJE0011Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            string o_status = a.AP_STAT;
            a.AP_STAT = code.AP_STAT_Forward(a.AP_STAT);
            sgsdb.SaveChanges();
            ViewBag.msg = "已送件";

            util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "ST", now, userid, WebSiteHelper.getUserDept(), "");

            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC, model.city, model.keyword2);
        }

        /// <summary>
        /// 退件
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendBackward")]
        public ActionResult SendBackward(PJE0011Model model)
        {
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            a.AP_STAT = code.AP_STAT_Backward(a.AP_STAT);
            sgsdb.SaveChanges();
            ViewBag.msg = "已退件";

            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC, model.city, model.keyword2);
        }

        /// <summary>
        /// 成案 (核定中送件) 20 -> 90 ST
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendDeal")]
        public ActionResult SendDeal(PJE0011Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            // AP_STAT = 20(核定中) 才能送件
            if (a.AP_STAT != "20")
            {
                ViewBag.msg = "狀態已變更，請勿重複操作";
                return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
            }
            string o_status = a.AP_STAT;
            a.AP_STAT = "90";
            sgsdb.SaveChanges();
            ViewBag.msg = "已成案";
            util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "ST", now, userid, WebSiteHelper.getUserDept(), "");
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        /// <summary>
        /// 批次成案 (核定中批次送件) 20 -> 90 ST
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendDealBatch")]
        public ActionResult SendDealBatch(PJE0011Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            string deptid = WebSiteHelper.getUserDept(); // 將取得部門 ID 移到迴圈外，節省效能

            // 防呆檢查：確認前端確實有傳入 AID 字串
            if (string.IsNullOrEmpty(model.AID))
            {
                ViewBag.msg = "未選擇任何要成案的資料";
                return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
            }

            // 將前端傳來逗號分隔的字串切成陣列 (例如 "1001,1002" -> ["1001", "1002"])
            string[] aidArray = model.AID.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            int successCount = 0; // 用來記錄成功處理的筆數

            foreach (string aidStr in aidArray)
            {
                decimal aid = tools.safe2Decimal(aidStr);
                APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();

                // 確保該筆資料存在，且 AP_STAT == "20" (核定中) 才能成案
                if (a != null && a.AP_STAT == "20")
                {
                    string o_status = a.AP_STAT;
                    a.AP_STAT = "90"; // 改為成案狀態

                    // 寫入 Log (注意這裡的 deptid 已經在迴圈外宣告好了)
                    util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "ST", now, userid, deptid, "");

                    successCount++;
                }
            }

            // 檢查是否有成功處理的資料，統一執行一次資料庫儲存 (大幅提升效能)
            if (successCount > 0)
            {
                sgsdb.SaveChanges();
                ViewBag.msg = $"已成功批次成案 {successCount} 筆";
            }
            else
            {
                ViewBag.msg = "勾選的資料狀態皆已變更，無符合成案條件的項目，請重新查詢";
            }

            // 返回原本的列表頁面，並帶入原本的查詢參數
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        /// <summary>
        /// 回複審中 (核定中退回) 20 -> 15 BK
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendPrevious")]
        public ActionResult SendPrevious(SGE0010Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            // AP_STAT = 20 (核定中) 才能退件
            if (a.AP_STAT != "20")
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

        // 💡 新增：批次重新審核 (多筆退回複審中) 20 -> 15 BK
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendPreviousBatch")]
        public ActionResult SendPreviousBatch(SGE0010Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();

            // 💡 1. 取得前端傳來的多筆 AID 字串 (例如 "1001,1002,1003")
            string aidStr = Request.Form["AID"];

            if (!string.IsNullOrEmpty(aidStr))
            {
                string[] aidList = aidStr.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                int successCount = 0;

                // 💡 2. 跑迴圈，一筆一筆處理退回邏輯
                foreach (var idStr in aidList)
                {
                    decimal aid = Convert.ToDecimal(idStr.Trim());
                    APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();

                    // 必須存在，且 AP_STAT = 20(核定中) 才能恢復狀態
                    if (a != null && a.AP_STAT == "20")
                    {
                        string o_status = a.AP_STAT;
                        a.AP_STAT = "15";

                        // 儲存狀態變更
                        sgsdb.SaveChanges();

                        // 寫入歷程紀錄
                        util.UPDATE_APMFORM_Log(a.AID, o_status, "15", "BK", now, userid, WebSiteHelper.getUserDept(), "批次恢復複審中");

                        successCount++;
                    }
                }

                ViewBag.msg = $"已成功批次重新審核 {successCount} 筆，恢復為複審中";
            }
            else
            {
                ViewBag.msg = "未接收到任何重新審核資料";
            }

            // 💡 3. 返回原本的列表畫面
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        /// <summary>
        /// 回核定中 (已成案退回) 90 -> 20 BK
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendBeforePrevious")]
        public ActionResult SendBeforePrevious(SGE0010Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            // AP_STAT = 90(已成案) 才能退件
            if (a.AP_STAT != "90")
            {
                ViewBag.msg = "狀態已變更，請勿重複操作";
                return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
            }
            string o_status = a.AP_STAT;
            a.AP_STAT = "20";
            sgsdb.SaveChanges();
            ViewBag.msg = "已恢復為核定中";

            util.UPDATE_APMFORM_Log(a.AID, o_status, "20", "BK", now, userid, WebSiteHelper.getUserDept(), "恢復核定中");
            return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2, model.YRC_ID, model.city, model.keyword2);
        }

        #region 污水-匯出核定中/已成案計畫評分表

        /// <summary>
        /// 中長程計畫-匯出核定中計畫評分表
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        //[HttpPost]
        //[MultiButton(MatchFormKey = "action", MatchFormValue = "ExportSewageScoreSheet")]
        public ActionResult ExportSewageScoreSheet(SGE0010Model m)
        {
            //string bodyType = "20"; // 核定中
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
            switch (m.tab)
            //switch (bodyType)
            {
                //case "15":
                //    sApWhere += "And AP_STAT='15' "; sApOrder = "Order By CITY,S1_ORDER,AP_NAME";
                //    //sApFund = ",IsNull(FUND,0)/1000 As FUND,IsNull(FUND,0)/1000 As DCS_FUND,IsNull(CEN_FUND,0)/1000 As CEN_DCS_FUND,IsNull(LOC_FUND,0)/1000 As LOC_DCS_FUND"; //全年度
                //    sApFundField += ",SC1_SCORE1 As SC1,SC2_SCORE1 As SC2,SC3_SCORE1 As SC3,SC4_SCORE1 As SC4,SC5_SCORE1 As SC5,SC6_SCORE1 As SC6,SC7_SCORE1 As SC7";
                //    sApFundField += ",M_FUND,M_CEN_FUND,M_LOC_FUND";
                //    sApFundField += ",M_JUD_FUND As S_FUND,M_CEN_JUD_FUND As S_CEN_FUND,M_LOC_JUD_FUND As S_LOC_FUND,S1_ORDER As S_ORDER";
                //    break;
                //case "20":
                case 4: // 核定中
                    sApWhere += "And AP_STAT='20' "; sApOrder = "ORDER BY CITY, S1_ORDER, AP_NAME"; //原S2_ORDER, AP_NAME
                    //sApFund = ",IsNull(FUND,0)/1000 As FUND,IsNull(DCS_FUND,0)/1000 As DCS_FUND,IsNull(CEN_DCS_FUND,0)/1000 As CEN_DCS_FUND,IsNull(LOC_DCS_FUND,0)/1000 As LOC_DCS_FUND"; //全年度
                    sApFundField += ",SC1_SCORE As SC1,SC2_SCORE As SC2,SC3_SCORE As SC3,SC4_SCORE As SC4,SC5_SCORE As SC5,SC6_SCORE As SC6,SC7_SCORE As SC7";
                    //sApFundField += ",SC1_SCORE1 As SC1,SC2_SCORE1 As SC2,SC3_SCORE1 As SC3,SC4_SCORE1 As SC4,SC5_SCORE1 As SC5,SC6_SCORE1 As SC6,SC7_SCORE1 As SC7";
                    sApFundField += ",M_FUND,M_CEN_FUND,M_LOC_FUND";
                    sApFundField += ",M_DCS_FUND As S_FUND,M_CEN_DCS_FUND As S_CEN_FUND,M_LOC_DCS_FUND As S_LOC_FUND,S1_ORDER As S_ORDER"; //S2_ORDER
                    break;
                //case "Z2":
                //    sApWhere += "And AP_STAT='Z2' "; sApOrder = "Order By CITY,S1_ORDER,AP_NAME";
                //    //sApFund = ",IsNull(FUND,0)/1000 As FUND,IsNull(FUND,0)/1000 As DCS_FUND,IsNull(CEN_FUND,0)/1000 As CEN_DCS_FUND,IsNull(LOC_FUND,0)/1000 As LOC_DCS_FUND"; //全年度
                //    sApFundField += ",SC1_SCORE As SC1,SC2_SCORE As SC2,SC3_SCORE As SC3,SC4_SCORE As SC4,SC5_SCORE As SC5,SC6_SCORE As SC6,SC7_SCORE As SC7";
                //    //sApFundField += ",SC1_SCORE1 As SC1,SC2_SCORE1 As SC2,SC3_SCORE1 As SC3,SC4_SCORE1 As SC4,SC5_SCORE1 As SC5,SC6_SCORE1 As SC6,SC7_SCORE1 As SC7";
                //    sApFundField += ",M_FUND,M_CEN_FUND,M_LOC_FUND";
                //    sApFundField += ",M_DCS_FUND As S_FUND,M_CEN_DCS_FUND As S_CEN_FUND,M_LOC_DCS_FUND As S_LOC_FUND,S1_ORDER As S_ORDER"; //S2_ORDER
                //    break;
                //case "90":
                case 5: // 已成案
                    sApWhere += "And AP_STAT='90' "; sApOrder = "Order By CITY, S1_ORDER, AP_NAME"; //原CITY, AP_NAME
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

        #endregion

        #region 污水-匯出核定中/已成案計畫明細表

        //[HttpPost]
        //[MultiButton(MatchFormKey = "action", MatchFormValue = "ExportSewageDetailsSheet")]
        public ActionResult ExportSewageDetailsSheet(SGE0010Model m)
        {
            //string bodyType = "2"; // 核定中
            decimal PPID = m.PPID;
            string YRC_ID = m.YRC_ID;
            string sCity = m.city;

            // 初始化 SqlCommand 與 基礎 Where 條件
            SqlCommand cmd = new SqlCommand();
            string sApWhere = "PPID = @PPID And YRC_ID = @YRC_ID ";
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            // 處理城市篩選 (參數化)
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
            switch (m.tab)
            {
                case 4: sApWhere += "And AP_STAT='20' "; sApOrder = "ORDER BY CITY, S1_ORDER, AP_NAME"; break; //原S2_ORDER, AP_NAME
                case 5: sApWhere += "And AP_STAT='90' "; sApOrder = "Order By CITY, S1_ORDER, AP_NAME"; break; //原CITY, AP_NAME
                    //case 3: sApWhere += "And AP_STAT='15' "; sApOrder = "Order By CITY,S1_ORDER,AP_NAME"; break;
                    //case 7: sApWhere += "And AP_STAT='Z2' "; sApOrder = "Order By CITY,S1_ORDER,AP_NAME"; break;
            }

            cmd.CommandText = "select *,CHK_STAT, S1_ORDER AS S_ORDER from APV_APMFORM a where " + sApWhere + sApOrder;
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
            string sExcelTitle = sCity + yrcInfo?.BGN_SYR + "年度" + pjmform?.PJ_NAME + "明細表"; //○○縣(市)○○年度污水下水道建設計畫明細表

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

        #endregion

        #region 雨水-匯出核定中/已成案計畫評分表

        /// <summary>
        /// 匯出計畫評分表-污水/雨水
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ExportScoreToExcel")]
        public ActionResult ExportScoreToExcel(SGE0010Model m)
        {
            var data = sgsdb.PJMFORM.SingleOrDefault(s => s.PPID == m.PPID)?.STAT_NO;
            if (data == "R") // 雨水中長程計畫
                return ExportRainScoreSheet(m);
            // 污水中長程計畫
            return ExportSewageScoreSheet(m);
        }

        // 新版 Aspose 雨水計畫提案評分表(分3個工作表)
        /// <summary>
        /// 匯出雨水計畫提案評分表
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        //[HttpPost]
        //[MultiButton(MatchFormKey = "action", MatchFormValue = "ExportRainScoreSheet")]
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

        #region 雨水-匯出核定中/已成案計畫明細表

        /// <summary>
        /// 匯出計畫提案明細表-污水/雨水
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ExportDetailsToExcel")]
        public ActionResult ExportDetailsToExcel(SGE0010Model m)
        {
            var data = sgsdb.PJMFORM.SingleOrDefault(s => s.PPID == m.PPID)?.STAT_NO;
            if (data == "R") // 雨水中長程計畫
                return ExportRainDetailsSheet(m);
            // 污水中長程計畫
            return ExportSewageDetailsSheet(m);
        }

        /// <summary>
        /// 匯出雨水計畫提案明細表
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        //[HttpPost]
        //[MultiButton(MatchFormKey = "action", MatchFormValue = "ExportRainDetailsSheet")]
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
            Debug.WriteLine($"CITY : {m.city}");
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
        /// 永續經營組 - 明細(概述)表[Sheet1]:
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

    /// <summary>
    /// 先期作業匯出excel工作表類別
    /// </summary>
    public class ExportSheet
    {
        /// <summary>
        /// 共用輔助函式
        /// </summary>
        public class Helper
        {
            /// <summary>
            /// excel列印設定
            /// </summary>
            /// <param name="ps"> 列印頁面設定物件 </param>
            /// <param name="sheet"> 目前工作表 </param>
            public static void SetPageSetup(Aspose.Cells.PageSetup ps, Aspose.Cells.Worksheet sheet)
            {
                // 設定 A3 紙張
                ps.PaperSize = PaperSizeType.PaperA3;

                // 建議橫向，欄位比較好塞
                ps.Orientation = PageOrientationType.Landscape;

                // 所有欄放入 1 頁
                ps.FitToPagesWide = 1;

                // 所有列也放入 1 頁（若只管欄，可拿掉這行）
                //ps.FitToPagesTall = 1;

                // 列印時高度要縮放成幾頁 0 = 不限制高度
                ps.FitToPagesTall = 0;

                // 關閉固定縮放比例
                //ps.Zoom = 100;

                //// 設定列印範圍
                string printArea = $"A1:{CellsHelper.CellIndexToName(sheet.Cells.MaxDataRow + 1 - 1, sheet.Cells.MaxDataColumn + 1 - 1)}";
                ps.PrintArea = printArea;
                //ps.PrintArea = "";
            }

            /// <summary>
            /// 取得中長程計畫-經費管制資料
            /// </summary>
            /// <param name="sgsdb"> 資料庫 </param>
            /// <param name="yrcId"> 年度編號 </param>
            /// <returns></returns>
            public static PJDSYRC GetYrcInfo(CMPSGSEntities sgsdb, string yrcId)
            {
                if (!decimal.TryParse(yrcId, out decimal yrcIdDecimal))
                {
                    // 無法轉 decimal，直接回傳空白
                    return null;
                }

                return sgsdb.PJDSYRC.SingleOrDefault(x => x.YRC_ID == yrcIdDecimal);
            }

            /// <summary>
            /// 取得案件類別查詢字串
            /// </summary>
            /// <param name="type"> 類型(1：工程，2：規劃，3：非工程) </param>
            /// <returns></returns>
            public static string GetPJ_CATY_CODE(int type)
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

            /// <summary>
            /// 備註設定(直接顯示在資料最下面)
            /// </summary>
            /// <param name="sheet"> 目前工作表 </param>
            /// <param name="type"> 類型(1：工程，2：規劃，3：非工程) </param>
            /// <param name="startRow"> 備註置放位置的上一列 </param>
            /// <param name="hiddenCol"> 合併儲存格的最後一行位置(需加1) </param>
            public static void SetFooterRemark(Worksheet sheet, int type, int startRow, int hiddenCol)
            {
                string footer1 = "註1：工程分類係指：01.都市排水整體改善工程(新建)/02.抽水站及滯洪池工程(新建)/03.雨水下水道管線(修繕)/04.抽水站及滯洪池工程(更新)/05.用地";
                if (type == 2)
                    footer1 = "註1：規劃分類係指：01.規劃/02.規劃檢討";
                else if (type == 3)
                    footer1 = "註1：非工程分類係指：01.下水道普查數化/02.抽水站感測器建置及介接/03.委外研究";
                string footer2 = "註2：「初評分數」、「初評意見」欄位由本署填報；「初評分數」80分以上本署原則依所提報經費期程核定辦理，80分以下則由本署衡酌計畫整體執行效益調整。";
                string footer3 = "註3：各工程請依行政院公共工程委員會「公共工程生態檢核機制」、「公共工程節能減碳注意事項」規定辦理。";
                string footer4 = "填表人：";
                string footer5 = "縣(市)政府單位主管：";
                string footer6 = "電話：";
                int rmkRow = startRow + 1;
                sheet.Cells[rmkRow, 0].PutValue(footer1);
                sheet.Cells[rmkRow + 1, 0].PutValue(footer2);
                sheet.Cells[rmkRow + 2, 0].PutValue(footer3);
                sheet.Cells[rmkRow + 4, 0].PutValue(footer4);
                sheet.Cells[rmkRow + 4, 4].PutValue(footer5);
                sheet.Cells[rmkRow + 5, 0].PutValue(footer6);
                Aspose.Cells.Style style = sheet.Cells[rmkRow, 0].GetStyle();
                sheet.Cells.Merge(rmkRow, 0, 1, hiddenCol);
                style = sheet.Cells[rmkRow + 1, 0].GetStyle();
                sheet.Cells.Merge(rmkRow + 1, 0, 1, hiddenCol);
                style = sheet.Cells[rmkRow + 2, 0].GetStyle();
                sheet.Cells.Merge(rmkRow + 2, 0, 1, hiddenCol);
            }

            /// <summary>
            /// 設定儲存格總邊框
            /// </summary>
            /// <param name="sheet"> 目前工作表 </param>
            /// <param name="r1_idx"> 從第X列開始 </param>
            /// <param name="c1_idx"> 從第X行開始 </param>
            /// <param name="r2_idx"> 第X列結束 </param>
            /// <param name="c2_idx"> 第X行結束(需加1) </param>
            public static void SetTotalBorder(Worksheet sheet, int r1_idx, int c1_idx, int r2_idx, int c2_idx)
            {
                Range dataRange = sheet.Cells.CreateRange(r1_idx, c1_idx, r2_idx, c2_idx);

                // SetOutlineBorder 只會設定範圍的 "外" 邊框
                dataRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                dataRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                dataRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                dataRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                // 所有儲存格邊框設定
                Style allBordersStyle = sheet.Cells[r1_idx, c1_idx].GetStyle();
                allBordersStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                allBordersStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                allBordersStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                allBordersStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                allBordersStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                allBordersStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                allBordersStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                allBordersStyle.Borders[BorderType.RightBorder].Color = Color.Black;
                StyleFlag allBordersFlag = new StyleFlag() { Borders = true };
                dataRange.ApplyStyle(allBordersStyle, allBordersFlag);
            }
        }

        /// <summary>
        /// 雨水計畫評分表
        /// </summary>
        public class RainScore
        {
            //private readonly CMPSGSEntities cmpsgsDB = new CMPSGSEntities();

            public byte[] GenerateScoreReportBytes(CMPSGSEntities cmpsgsDB, decimal PPID, string YRC_ID, string cityText, int tab, string apName, out string title)
            {
                ExportSheet.RainScore rainScore = new ExportSheet.RainScore();

                Workbook workbook = new Workbook(HostingEnvironment.MapPath("~/Templates/SGE0060/SGE0060Score_R.xlsx"));
                WorkbookDesigner designer = new WorkbookDesigner();
                designer.Workbook = workbook;

                // 標題資料取得
                PJMFORM_CMPSGS pjmform = cmpsgsDB.PJMFORM.Where(w => w.PPID == PPID).FirstOrDefault();
                string syrText = "";
                if (decimal.TryParse(YRC_ID, out var yrcIdDecimal))
                {
                    var yrcInfo = cmpsgsDB.PJDSYRC.SingleOrDefault(w => w.PPID == PPID && w.YRC_ID == yrcIdDecimal);
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
            /// 取得雨水計畫提案評分表資料
            /// </summary>
            /// <param name="type"> 類型(1：工程，2：規劃，3：非工程) </param>
            /// <param name="PPID"> 中長程計畫編號 </param>
            /// <param name="YRC_ID"> 年度編號 </param>
            /// <param name="city"> 縣市 </param>
            /// <param name="tab"> 頁籤(2：初審中；3：複審中；4：核定中；5：已成案) </param>
            /// <param name="apName"> 工程名稱(模糊查詢) </param>
            /// <param name="isSGE0010"> 是否為提案初審頁面(表示初審中帳號只顯示自己的案件) </param>
            /// <returns></returns>
            public DataTable GetRainScoreTable(int type, decimal PPID, string YRC_ID, string city, int tab, string apName, bool isSGE0010 = false)
            {
                string paramValue = Helper.GetPJ_CATY_CODE(type);
                SqlCommand cmd = new SqlCommand();
                string sApOrder = ""; // 排序字串
                string sApFundField = ""; // 選取字串
                string sApWhere = ""; // 條件字串
                string AP_STAT = "15"; // 案件狀態
                switch (tab)
                {
                    case 1: //待收件(使用初審金額)
                        AP_STAT = "00";
                        sApFundField += ", M_JUD_FUND As S_FUND, M_CEN_JUD_FUND As S_CEN_FUND, M_LOC_JUD_FUND As S_LOC_FUND, a.S1_ORDER As S_ORDER";
                        sApOrder = "ORDER BY a.CITY, a.S1_ORDER, a.AP_NAME";
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                    case 2: // 初審中
                        AP_STAT = "10";
                        sApFundField += ", M_JUD_FUND As S_FUND, M_CEN_JUD_FUND As S_CEN_FUND, M_LOC_JUD_FUND As S_LOC_FUND, a.S1_ORDER As S_ORDER";
                        sApOrder = "ORDER BY a.CITY, a.S1_ORDER, a.AP_NAME ";
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                    case 3: // 複審中
                        AP_STAT = "15";
                        sApFundField += ", M_DCS_FUND As S_FUND, M_CEN_DCS_FUND As S_CEN_FUND, M_LOC_DCS_FUND As S_LOC_FUND, a.S1_ORDER As S_ORDER";
                        sApOrder = "ORDER BY a.CITY, a.S1_ORDER, a.AP_NAME";
                        break;
                    case 4: // 核定中
                        AP_STAT = "20";
                        sApFundField += ", M_DCS_FUND As S_FUND, M_CEN_DCS_FUND As S_CEN_FUND, M_LOC_DCS_FUND As S_LOC_FUND, a.S1_ORDER As S_ORDER"; //S2_ORDER
                        sApOrder = "ORDER BY a.CITY, a.S1_ORDER, a.AP_NAME"; //S2_ORDER
                        break;
                    case 5: // 已成案
                        AP_STAT = "90";
                        sApFundField += ", M_DCS_FUND As S_FUND, M_CEN_DCS_FUND As S_CEN_FUND, M_LOC_DCS_FUND As S_LOC_FUND, a.S1_ORDER As S_ORDER"; //S2_ORDER
                        sApOrder = "ORDER BY a.CITY, a.S1_ORDER, a.AP_NAME";
                        break;

                    case 6: //未過初審(使用初審金額)
                        AP_STAT = "Z1";
                        sApFundField += ", M_JUD_FUND As S_FUND, M_CEN_JUD_FUND As S_CEN_FUND, M_LOC_JUD_FUND As S_LOC_FUND, a.S1_ORDER As S_ORDER";
                        sApOrder = "ORDER BY a.CITY, a.S1_ORDER, a.AP_NAME";
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                    case 7: // 未過複審(使用複審金額)
                        AP_STAT = "Z2";
                        sApFundField += ", M_DCS_FUND As S_FUND, M_CEN_DCS_FUND As S_CEN_FUND, M_LOC_DCS_FUND As S_LOC_FUND, a.S1_ORDER As S_ORDER";
                        sApOrder = "ORDER BY a.CITY, a.S1_ORDER, a.AP_NAME";
                        break;
                }
                string sql = $"SELECT a.SYR, a.AID, a.AP_CODE, a.AP_NAME, a.AP_ALL_NOTE, AP_YEAR_NOTE, a.AP_NOTE, CHK_STAT, BUDGET_TP, BUDGET_TP_N," +
                            $" a.SWW2_FG, a.SWW4_FG, a.S1_COMT, S2_COMT " +
                            $", a.SC1_SCORE, a.SC2_SCORE, a.SC3_SCORE, a.SC4_SCORE, a.SC5_SCORE, a.SC6_SCORE, a.SC7_SCORE, a.SC8_SCORE, a.SC_SCORE, a.CITY " +
                            $", M_T_FUND, M_R_FUND, M_FUND, M_CEN_FUND, M_LOC_FUND,a.EXT_FG,aa.CODE_DESC,a.S1_TOT_FUND " +
                            $"{sApFundField} " +
                            $"FROM APV_APMFORM a left join APV_APDRPAL aa on a.AID=aa.AID " +
                            $"WHERE a.PJ_CATY_CODE IN({paramValue}) AND " +
                            $"a.AP_STAT = @AP_STAT AND a.PPID = @PPID AND a.YRC_ID LIKE @YRC_ID";
                cmd.Parameters.Add(new SqlParameter("AP_STAT", AP_STAT));
                cmd.Parameters.Add(new SqlParameter("PPID", PPID));
                cmd.Parameters.Add(new SqlParameter("YRC_ID", string.IsNullOrEmpty(YRC_ID) ? "%" : YRC_ID));
                if (!string.IsNullOrEmpty(city))
                {
                    var values = new List<string>() { city };
                    switch (city)
                    {
                        case "新北市": values.Add("台北縣"); break;
                        case "台中市": values.Add("台中縣"); break;
                        case "台南市": values.Add("台南縣"); break;
                        case "高雄市": values.Add("高雄縣"); break;
                        case "桃園市": values.Add("桃園縣"); break;
                    }
                    if (values.Any())
                    {
                        var paramNames = values
                            .Select((v, i) => $"@p{i}")
                            .ToArray();
                        sql += $" AND a.CITY IN ({string.Join(",", paramNames)})";
                        for (int i = 0; i < values.Count; i++)
                        {
                            cmd.Parameters.AddWithValue(paramNames[i], values[i]);
                        }
                    }
                }
                // 【新增】工程名稱 (AP_NAME) 的模糊查詢過濾
                if (!string.IsNullOrEmpty(apName))
                {
                    sql += $" AND a.AP_NAME LIKE @AP_NAME";
                    cmd.Parameters.Add(new SqlParameter("AP_NAME", $"%{apName}%"));
                }
                sql += $" {sApWhere} ";
                sql += $" {sApOrder}";
                cmd.CommandText = sql;
                DataTable ts = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);

                return ts;
            }

            /// <summary>
            /// 設定雨水計畫提案評分表工作表內容
            /// </summary>
            /// <param name="type"> 類型(1：工程，2：規劃，3：非工程) </param>
            /// <param name="workbook"></param>
            /// <param name="oDt"></param>
            /// <param name="header"></param>
            public void SetRainScoreSheet(int type, Workbook workbook, DataTable oDt, string header)
            {
                Worksheet sheet = workbook.Worksheets[type - 1];

                CpcClass.MyFunc CpcFun = new CpcClass.MyFunc();
                // Aspose 使用 0 基底索引，但原始邏輯從 1 基底的第 5 行開始
                // 我們將保持 iRow 為 1 基底，但在所有 Aspose 呼叫中都減 1
                int iRow = 5;
                int iRow_SUB = 0;
                Double iS_FUND = 0;
                Double iS_CEN_FUND = 0;
                Double iS_LOC_FUND = 0;
                Double tS_FUND = 0;
                Double tS_CEN_FUND = 0;
                Double tS_LOC_FUND = 0;

                Double tT_FUND = 0;
                Double tR_FUND = 0;
                Double tM_FUND = 0;
                Double tM_CEN_FUND = 0;
                Double tM_LOC_FUND = 0;
                string sumRow = "";
                int iCol = 1; // 同樣，保持 1 基底，但在呼叫時減 1
                string sCHK_STAT = "";
                int iRow_B = 0;
                DataRow[] oRow_B = null;
                DataRow DR_B = null;
                DataRow[] oRow_APM = null;
                DataRow rowViewAp = null;
                int iRow_APM = 0;
                DataTable oDt_BUDGET_TP = oDt.DefaultView.ToTable(true, new string[] { "BUDGET_TP", "BUDGET_TP_N" });
                oRow_B = oDt_BUDGET_TP.Select();
                string BUDGET_TP = "";
                string BUDGET_TP_N = "";
                iRow_B = oRow_B.Length;

                int colIdx_ST = 25; // 不同工作表"總經費"欄位位置(1基底)
                switch (type)
                {
                    case 1: colIdx_ST = 25; break;
                    case 2: colIdx_ST = 23 - 2; break;
                    case 3: colIdx_ST = 21 - 2; break;
                }

                sheet.Cells[1, 0].Value = header; // 標題
                sheet.Cells.Merge(1, 0, 1, colIdx_ST + 5); // 合併儲存格

                for (int i = 0; i < iRow_B; i++)
                {
                    iRow += 1;
                    iRow_SUB = iRow;
                    iCol = 1;
                    DR_B = oRow_B[i];
                    BUDGET_TP = DR_B["BUDGET_TP"].ToString();
                    if (string.IsNullOrEmpty(BUDGET_TP))
                    {
                        BUDGET_TP_N = "無分類";
                    }
                    else
                    {
                        BUDGET_TP_N = DR_B["BUDGET_TP_N"].ToString();
                    }

                    // 預算分類表頭處理
                    // Aspose 索引: [row, col] -> [row-1, col-1]
                    Range mergeRange = sheet.Cells.CreateRange(iRow - 1, 0, 1, colIdx_ST - 1);
                    mergeRange.Merge(); // 合併儲存格
                                        // 存取儲存格並設定樣式
                    Cell headerCell = sheet.Cells[iRow - 1, iCol - 1]; // iCol is 1, so col-index is 0
                    headerCell.PutValue(BUDGET_TP_N); //項次
                    Style headerStyle = headerCell.GetStyle();
                    headerStyle.IsTextWrapped = true;
                    headerStyle.HorizontalAlignment = TextAlignmentType.Left;
                    headerStyle.ForegroundColor = System.Drawing.Color.LightBlue;
                    headerStyle.Pattern = BackgroundType.Solid; // Aspose 需要 Pattern 才能顯示 FG Color
                    headerCell.SetStyle(headerStyle);

                    // 表頭最後五欄合併處理
                    mergeRange = sheet.Cells.CreateRange(iRow - 1, colIdx_ST + 2, 1, 5);
                    mergeRange.Merge(); // 合併儲存格
                    Cell cellLast = sheet.Cells[iRow - 1, colIdx_ST + 2]; // 26 - 1 = 25
                    cellLast.SetStyle(headerStyle);

                    // 取得該預算分類的資料列
                    oRow_APM = oDt.Select("isnull(BUDGET_TP,'')='" + BUDGET_TP + "'", "CITY, S_ORDER, AP_NAME");
                    iRow_APM = oRow_APM.Length;
                    iS_FUND = 0;
                    iS_CEN_FUND = 0;
                    iS_LOC_FUND = 0;
                    for (int j = 0; j < iRow_APM; j++)
                    {
                        iRow += 1;
                        iCol = 1;
                        rowViewAp = oRow_APM[j];

                        // --- 輔助函式 (Helper function) ---
                        // 為了簡潔，我們將使用一個本地函式來設定儲存格的值、換行和數字格式
                        // 這會使轉換後的程式碼更易讀
                        Action<int, object, string, bool> SetCell =
                            (colIdx, value, format, wrap) =>
                            {
                                Cell cell = sheet.Cells[iRow - 1, colIdx - 1];

                                // 嘗試將值設為 double (如果格式是數字)
                                if (!string.IsNullOrEmpty(format))
                                {
                                    double numValue = Convert.ToDouble(value);
                                    cell.PutValue(numValue);
                                }
                                else
                                {
                                    cell.PutValue(value);
                                }

                                Style style = cell.GetStyle();
                                if (wrap)
                                {
                                    style.IsTextWrapped = true;
                                }
                                if (!string.IsNullOrEmpty(format))
                                {
                                    //style.NumberFormat = format;
                                }
                                cell.SetStyle(style);
                            };
                        // ---------------------------------

                        // 基本資料
                        SetCell(iCol, (j + 1).ToString(), null, false); iCol++; //項次
                        SetCell(iCol, rowViewAp["AP_NAME"].ToString(), null, true); iCol++; //工程名稱
                        SetCell(iCol, rowViewAp["AP_ALL_NOTE"].ToString(), null, true); iCol++; //全期內容
                        SetCell(iCol, rowViewAp["AP_YEAR_NOTE"].ToString(), null, true); iCol++; //本年度辦理內容

                        // 經費需求(千元)
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["M_T_FUND"]), "###,##0", false); iCol++; //工程總經費
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["M_R_FUND"]), "###,##0", false); iCol++; //以前年度執行經費
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["M_FUND"]), "###,##0", false); iCol++; //本年度申請經費
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["M_CEN_FUND"]), "###,##0", false); iCol++; //中央補助款
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["M_LOC_FUND"]), "###,##0", false); iCol++; //地方配合款

                        try { tT_FUND += CpcFun.MyCDbl(rowViewAp["M_T_FUND"]); } catch { }
                        try { tR_FUND += CpcFun.MyCDbl(rowViewAp["M_R_FUND"]); } catch { }
                        try { tM_FUND += CpcFun.MyCDbl(rowViewAp["M_FUND"]); } catch { }
                        try { tM_CEN_FUND += CpcFun.MyCDbl(rowViewAp["M_CEN_FUND"]); } catch { }
                        try { tM_LOC_FUND += CpcFun.MyCDbl(rowViewAp["M_LOC_FUND"]); } catch { }

                        // 中央審核欄(內政部評分)
                        SetCell(iCol, rowViewAp["SC1_SCORE"] != DBNull.Value ? rowViewAp["SC1_SCORE"] : "-", null, false); iCol++; // 評分項目1
                        SetCell(iCol, rowViewAp["SC2_SCORE"] != DBNull.Value ? rowViewAp["SC2_SCORE"] : "-", null, false); iCol++; // 評分項目2
                        SetCell(iCol, rowViewAp["SC3_SCORE"] != DBNull.Value ? rowViewAp["SC3_SCORE"] : "-", null, false); iCol++; // 評分項目3
                        SetCell(iCol, rowViewAp["SC4_SCORE"] != DBNull.Value ? rowViewAp["SC4_SCORE"] : "-", null, false); iCol++; // 評分項目4
                        if (type == 1 || type == 2)
                        {
                            SetCell(iCol, rowViewAp["SC5_SCORE"] != DBNull.Value ? rowViewAp["SC5_SCORE"] : "-", null, false); iCol++; // 評分項目5
                            SetCell(iCol, rowViewAp["SC6_SCORE"] != DBNull.Value ? rowViewAp["SC6_SCORE"] : "-", null, false); iCol++; // 評分項目6
                        }
                        if (type == 1)
                        {
                            SetCell(iCol, rowViewAp["SC7_SCORE"] != DBNull.Value ? rowViewAp["SC7_SCORE"] : "-", null, false); iCol++; // 評分項目7
                            SetCell(iCol, rowViewAp["SC8_SCORE"] != DBNull.Value ? rowViewAp["SC8_SCORE"] : "-", null, false); iCol++; // 評分項目8
                        }
                        SetCell(iCol, rowViewAp["SC_SCORE"] != DBNull.Value ? rowViewAp["SC_SCORE"] : "-", null, false); iCol++; // 總評分

                        if (type == 1)
                        {
                            string SWW2_ST = rowViewAp["SWW2_FG"] == DBNull.Value ? "否" : (bool.Parse(rowViewAp["SWW2_FG"].ToString()) ? "是" : "否");
                            SetCell(iCol, SWW2_ST, null, false); iCol++; // 是否完成規劃或治理計畫
                            string SWW4_ST = rowViewAp["SWW4_FG"] == DBNull.Value ? "否" : (bool.Parse(rowViewAp["SWW4_FG"].ToString()) ? "是" : "否");
                            SetCell(iCol, SWW4_ST, null, false); iCol++; // 用地取得
                        }

                        // 中央評估結果
                        try { SetCell(iCol, rowViewAp["S_ORDER"], "##0", false); iCol++; } catch { iCol++; } //辦理優先次序
                        sCHK_STAT = rowViewAp["CHK_STAT"].ToString(); //本年度是否補助
                        switch (sCHK_STAT)
                        {
                            case "Y":
                                sheet.Cells[iRow - 1, iCol - 1].PutValue("V"); iCol++; //是
                                iCol++; //否
                                break;
                            case "N":
                                iCol++; //是
                                sheet.Cells[iRow - 1, iCol - 1].PutValue("V"); iCol++; //否
                                break;
                            default:
                                iCol++; //是
                                iCol++; //否
                                break;
                        }


                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["S1_TOT_FUND"]), "###,##0", false); iCol++; //總經費

                        // 核列經費(千元)
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["S_FUND"]), "###,##0", false); iCol++; //本年度核定經費
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["S_CEN_FUND"]), "###,##0", false); iCol++; //中央補助款
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["S_LOC_FUND"]), "###,##0", false); iCol++; //地方配合款

                        try { iS_FUND += CpcFun.MyCDbl(rowViewAp["S_FUND"]); } catch { }
                        try { iS_CEN_FUND += CpcFun.MyCDbl(rowViewAp["S_CEN_FUND"]); } catch { }
                        try { iS_LOC_FUND += CpcFun.MyCDbl(rowViewAp["S_LOC_FUND"]); } catch { }

                        SetCell(iCol, rowViewAp["S1_COMT"], null, false); iCol++; // 初評意見
                        SetCell(iCol, rowViewAp["S2_COMT"], null, false); iCol++; // 複評意見
                        SetCell(iCol, rowViewAp["AP_NOTE"], null, false); iCol++; // 備註
                        SetCell(iCol, Convert.ToBoolean(rowViewAp["EXT_FG"]) == true ? "延續性工程" : "新建工程", null, false); iCol++; // 辦理狀態
                        SetCell(iCol, rowViewAp["CODE_DESC"], null, false); iCol++; // 工程分類
                    }

                    tS_FUND += iS_FUND;
                    tS_CEN_FUND += iS_CEN_FUND;
                    tS_LOC_FUND += iS_LOC_FUND;

                    // 設定小計 (Subtotal)
                    // 輔助函式 (Helper function)
                    Action<int, double, string, Color> SetSubTotalCell =
                        (colIdx, value, format, bgColor) =>
                        {
                            Cell cell = sheet.Cells[iRow_SUB - 1, colIdx - 1];
                            cell.PutValue(value);
                            Style style = cell.GetStyle();
                            //style.NumberFormat = format;
                            style.ForegroundColor = bgColor;
                            style.Pattern = BackgroundType.Solid;
                            cell.SetStyle(style);
                        };

                    SetSubTotalCell(colIdx_ST, iS_FUND, "###,##0", Color.LightPink);
                    SetSubTotalCell(colIdx_ST + 1, iS_CEN_FUND, "###,##0", Color.LightPink);
                    SetSubTotalCell(colIdx_ST + 2, iS_LOC_FUND, "###,##0", Color.LightPink);
                }

                iRow += 1;

                sheet.Cells[iRow - 1, 0].PutValue("總計"); // col 2 -> index 1

                // 輔助函式 (Helper function)
                Action<int, double, string> SetGrandTotalCell =
                    (colIdx, value, format) =>
                    {
                        Cell cell = sheet.Cells[iRow - 1, colIdx - 1];
                        cell.PutValue(value);
                        Style style = cell.GetStyle();
                        //style.NumberFormat = format;
                        cell.SetStyle(style);
                    };

                // 總計值
                int colIdx_GT = 5;
                SetGrandTotalCell(colIdx_GT, tT_FUND, "###,##0"); //工程總經費
                SetGrandTotalCell(colIdx_GT + 1, tR_FUND, "###,##0"); //以前年度執行經費
                SetGrandTotalCell(colIdx_GT + 2, tM_FUND, "###,##0"); //本年度申請經費
                SetGrandTotalCell(colIdx_GT + 3, tM_CEN_FUND, "###,##0"); //中央補助款
                SetGrandTotalCell(colIdx_GT + 4, tM_LOC_FUND, "###,##0"); //地方配合款

                SetGrandTotalCell(colIdx_ST, tS_FUND, "###,##0"); //本年度核定經費
                SetGrandTotalCell(colIdx_ST + 1, tS_CEN_FUND, "###,##0"); //中央補助款
                SetGrandTotalCell(colIdx_ST + 2, tS_LOC_FUND, "###,##0"); //地方配合款

                // 總計列的樣式和邊框
                // 原始: sheet[iRow, 1, iRow, 27]
                Range gTotalRange = sheet.Cells.CreateRange(iRow - 1, 0, 1, colIdx_ST + 7); // (row, col, totalRows, totalCols)
                Style gTotalStyle = sheet.Cells[iRow - 1, 0].GetStyle(); // 取得一個樣式物件來修改
                gTotalStyle.ForegroundColor = System.Drawing.Color.LawnGreen;
                gTotalStyle.Pattern = BackgroundType.Solid;

                // 設定邊框
                gTotalStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                gTotalStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                gTotalStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                gTotalStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                gTotalStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                gTotalStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                gTotalStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                gTotalStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                // 建立一個 StyleFlag 來僅套用邊框和底色
                StyleFlag flag = new StyleFlag();
                flag.Borders = true;
                flag.CellShading = true;

                gTotalRange.ApplyStyle(gTotalStyle, flag);

                // 儲存格邊框設定
                Helper.SetTotalBorder(sheet, 5, 0, iRow - 5, sheet.Cells.MaxDataColumn + 1);

                // 設定列印範圍
                Aspose.Cells.PageSetup ps = sheet.PageSetup;
                string printArea = $"A1:{CellsHelper.CellIndexToName(sheet.Cells.MaxDataRow + 1 - 1, sheet.Cells.MaxDataColumn + 1 - 1)}";
                ps.PrintArea = printArea;
            }
        }

        /// <summary>
        /// 雨水計畫明細表
        /// </summary>
        public class RainDetail
        {
            //private readonly CMPSGSEntities cmpsgsDB = new CMPSGSEntities();

            // [k]
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
            public byte[] GenerateReportBytes(CMPSGSEntities cmpsgsDB, decimal PPID, string YRC_ID, int tab, string cityText, string apName, out string title)
            {
                ExportSheet.RainDetail rainDetail = new ExportSheet.RainDetail();

                Workbook workbook = new Workbook(HostingEnvironment.MapPath("~/Templates/SGE0060/SGE0060Details_R.xlsx"));
                WorkbookDesigner designer = new WorkbookDesigner();
                designer.Workbook = workbook;

                // 頁首頁尾資料取得
                PJMFORM_CMPSGS pjmform = cmpsgsDB.PJMFORM.Where(w => w.PPID == PPID).FirstOrDefault();
                string syrText = "";
                if (decimal.TryParse(YRC_ID, out var yrcIdDecimal))
                {
                    var yrcInfo = cmpsgsDB.PJDSYRC.SingleOrDefault(w => w.PPID == PPID && w.YRC_ID == yrcIdDecimal);
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

            /// <summary>
            /// 取得雨水計畫提案明細表資料
            /// </summary>
            /// <param name="type"> 類型(1：工程，2：規劃，3：非工程) </param>
            /// <param name="PPID"></param>
            /// <param name="YRC_ID"></param>
            /// <param name="tab"></param>
            /// <param name="city"></param>
            /// <param name="apName"></param>
            /// <param name="isSGE0010"> 是否為提案初審頁面(表示初審中帳號只顯示自己的案件) </param>
            /// <returns></returns>
            public DataTable GetRainDetailTable(int type, decimal PPID, string YRC_ID, int tab, string city, string apName, bool isSGE0010 = false)
            {
                string paramValue = Helper.GetPJ_CATY_CODE(type);
                SqlCommand cmd = new SqlCommand();
                string sql = $"SELECT * FROM APV_APDRPAL a WHERE PJ_CATY_CODE IN({paramValue}) AND " +
                    $"AP_STAT = @AP_STAT AND PPID = @PPID AND YRC_ID LIKE @YRC_ID";
                string sApOrder = ""; // 排序字串
                string sApWhere = ""; // 條件字串
                string AP_STAT = "15"; // 案件狀態
                switch (tab)
                {
                    case 1: //提案中
                    case 2: //待收件(使用初審金額)
                        AP_STAT = "00";
                        sApOrder = "ORDER BY CITY, S1_ORDER, AP_NAME";
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                    case 3: // 初審中
                        AP_STAT = "10";
                        sApOrder = "ORDER BY CITY, S1_ORDER, AP_NAME";
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                    case 4: // 複審中
                        AP_STAT = "15";
                        sApOrder = "ORDER BY CITY, S1_ORDER, AP_NAME";
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                    case 5: // 核定中
                        AP_STAT = "20";
                        sApOrder = "ORDER BY CITY, S1_ORDER, AP_NAME"; //S2_ORDER
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                    case 6: // 已成案
                        AP_STAT = "90";
                        sApOrder = "ORDER BY CITY, S1_ORDER, AP_NAME";
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                    case 7: //未過初審(使用初審金額)
                        AP_STAT = "Z1";
                        sApOrder = "ORDER BY CITY, S1_ORDER, AP_NAME";
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                    case 8: // 未過複審(使用複審金額)
                        AP_STAT = "Z2";
                        sApOrder = "ORDER BY CITY, S1_ORDER, AP_NAME";
                        if (isSGE0010) sApWhere = $" AND (SELECT TOP 1 PRC_DP FROM APDPROC WHERE AID = a.AID and NAP_STAT = '{AP_STAT}' AND PROC_ACT ='RV' ORDER BY ADD_DT DESC) = '{WebSiteHelper.getUserDept()}' ";
                        break;
                }
                cmd.Parameters.Add(new SqlParameter("AP_STAT", AP_STAT));
                cmd.Parameters.Add(new SqlParameter("PPID", PPID));
                cmd.Parameters.Add(new SqlParameter("YRC_ID", string.IsNullOrEmpty(YRC_ID) ? "%" : YRC_ID));
                if (!string.IsNullOrEmpty(city))
                {
                    var values = new List<string>() { city };
                    switch (city)
                    {
                        case "新北市": values.Add("台北縣"); break;
                        case "台中市": values.Add("台中縣"); break;
                        case "台南市": values.Add("台南縣"); break;
                        case "高雄市": values.Add("高雄縣"); break;
                        case "桃園市": values.Add("桃園縣"); break;
                    }
                    if (values.Any())
                    {
                        var paramNames = values
                            .Select((v, i) => $"@p{i}")
                            .ToArray();
                        sql += $" AND CITY IN ({string.Join(",", paramNames)})";
                        for (int i = 0; i < values.Count; i++)
                        {
                            cmd.Parameters.AddWithValue(paramNames[i], values[i]);
                        }
                    }
                }
                // 【新增】工程名稱 (AP_NAME) 的模糊查詢過濾
                if (!string.IsNullOrEmpty(apName))
                {
                    sql += $" AND AP_NAME LIKE @AP_NAME";
                    cmd.Parameters.Add(new SqlParameter("AP_NAME", $"%{apName}%"));
                }
                sql += $" {sApWhere} ";
                sql += $" {sApOrder}";
                cmd.CommandText = sql;
                DataTable ts = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);

                return ts;
            }

            /// <summary>
            /// 設定雨水計畫提案明細表工作表內容
            /// </summary>
            /// <param name="type"> 類型(1：工程，2：規劃，3：非工程) </param>
            /// <param name="workbook"></param>
            /// <param name="ts"></param>
            /// <param name="header"></param>
            /// <param name="footer"></param>
            public void SetRainDetailsSheet(int type, Workbook workbook, DataTable ts, string header, string footer)
            {
                Worksheet sheet = workbook.Worksheets[type - 1];
                workbook.Settings.ReCalculateOnOpen = true;

                int startRow = 3;
                int hiddenCol = 31;
                switch (type)
                {
                    case 1: // 工程
                        hiddenCol = 31;
                        if (ts.Rows.Count > 0)
                        {
                            sheet.Cells[2, 10].PutValue(ts.Rows[0]["SYR"] + "年");
                            sheet.Cells[2, 13].PutValue(ts.Rows[0]["SYR"] + "年");
                        }
                        foreach (DataRow d in ts.Rows)
                        {
                            sheet.Cells.CopyRow(sheet.Cells, startRow, startRow + 1); //複制格式到下行
                            sheet.Cells[startRow, 0].PutValue(d["S1_ORDER"]);
                            sheet.Cells[startRow, 1].PutValue(d["CODE_DESC"]);
                            sheet.Cells[startRow, 2].PutValue(d["CITY"]);
                            sheet.Cells[startRow, 3].PutValue(d["TOWN"]);
                            sheet.Cells[startRow, 4].PutValue(d["WSYS_NAME"]);
                            sheet.Cells[startRow, 5].PutValue(d["RSYS_NAME"]); // TODO：3.1.8雨水科實施計畫名稱修改後，需將RSYS_NAME改回AP_PJ_NAME // 20251222
                            sheet.Cells[startRow, 6].PutValue(d["AP_NAME"]);
                            sheet.Cells[startRow, 7].PutValue(d["AP_ALL_NOTE"]);
                            sheet.Cells[startRow, 17].PutValue(d["SWW2_ST"]);
                            sheet.Cells[startRow, 18].PutValue(d["SWW4_ST"]);
                            sheet.Cells[startRow, 19].PutValue(d["IMPFLDAREA"]);
                            sheet.Cells[startRow, 20].PutValue(d["PRTPOPS"]);
                            sheet.Cells[startRow, 21].PutValue(d["START_DT"]);
                            sheet.Cells[startRow, 22].PutValue(d["END_DT"]);
                            sheet.Cells[startRow, 23].PutValue(d["DEPT_NAME"]);
                            sheet.Cells[startRow, 24].PutValue(d["SC_SCORE1"] != DBNull.Value ? d["SC_SCORE1"] : "-");
                            sheet.Cells[startRow, 25].PutValue(d["SC_SCORE"] != DBNull.Value ? d["SC_SCORE"] : "-");
                            sheet.Cells[startRow, 26].PutValue(d["S1_COMT"]); // 初評意見
                            sheet.Cells[startRow, 27].PutValue(d["GIS_X"]);
                            sheet.Cells[startRow, 28].PutValue(d["GIS_Y"]);
                            sheet.Cells[startRow, 29].PutValue(d["AP_NOTE"]);
                            sheet.Cells[startRow, 30].PutValue((d["EXT_FG"] != DBNull.Value && Convert.ToBoolean(d["EXT_FG"])) ? "延續性工程" : "新建工程");
                            // 計算用，需隱藏
                            sheet.Cells[startRow, 31].PutValue(d["REXE_FUND"]);
                            sheet.Cells[startRow, 32].PutValue(d["CCEN_FUND"]);
                            sheet.Cells[startRow, 33].PutValue(d["FCEN_FUND"]);
                            sheet.Cells[startRow, 34].PutValue(d["CLOC_FUND"]);
                            sheet.Cells[startRow, 35].PutValue(d["FLOC_FUND"]);                            // 套用格式到目標行
                            startRow++;
                        }
                        break;
                    case 2: // 規劃
                        hiddenCol = 24;
                        if (ts.Rows.Count > 0)
                        {
                            sheet.Cells[2, 8].PutValue(ts.Rows[0]["SYR"] + "年");
                            sheet.Cells[2, 11].PutValue(ts.Rows[0]["SYR"] + "年");
                        }
                        foreach (DataRow d in ts.Rows)
                        {
                            sheet.Cells.CopyRow(sheet.Cells, startRow, startRow + 1);
                            sheet.Cells[startRow, 0].PutValue(d["S1_ORDER"]);
                            sheet.Cells[startRow, 1].PutValue(d["CODE_DESC"]);
                            sheet.Cells[startRow, 2].PutValue(d["CITY"]);
                            sheet.Cells[startRow, 3].PutValue(d["TOWN"]);
                            sheet.Cells[startRow, 4].PutValue(d["WSYS_NAME"]);
                            sheet.Cells[startRow, 5].PutValue(d["AP_NAME"]);
                            sheet.Cells[startRow, 15].PutValue(d["START_DT"]);
                            sheet.Cells[startRow, 16].PutValue(d["END_DT"]);
                            sheet.Cells[startRow, 17].PutValue(d["IMPFLDAREA"]);
                            sheet.Cells[startRow, 18].PutValue(d["DEPT_NAME"]);
                            sheet.Cells[startRow, 19].PutValue(d["SC_SCORE1"] != DBNull.Value ? d["SC_SCORE1"] : "-");
                            sheet.Cells[startRow, 20].PutValue(d["SC_SCORE"] != DBNull.Value ? d["SC_SCORE"] : "-");
                            sheet.Cells[startRow, 21].PutValue(d["S1_COMT"]); // 初評意見
                            sheet.Cells[startRow, 22].PutValue(d["AP_NOTE"]);
                            sheet.Cells[startRow, 23].PutValue((d["EXT_FG"] != DBNull.Value && Convert.ToBoolean(d["EXT_FG"])) ? "延續性工程" : "新建工程");
                            // 計算用，需隱藏
                            sheet.Cells[startRow, 24].PutValue(d["REXE_FUND"]);
                            sheet.Cells[startRow, 25].PutValue(d["CCEN_FUND"]);
                            sheet.Cells[startRow, 26].PutValue(d["FCEN_FUND"]);
                            sheet.Cells[startRow, 27].PutValue(d["CLOC_FUND"]);
                            sheet.Cells[startRow, 28].PutValue(d["FLOC_FUND"]);
                            startRow++;
                        }
                        break;
                    case 3: // 非工程
                        hiddenCol = 24;
                        if (ts.Rows.Count > 0)
                        {
                            sheet.Cells[2, 9].PutValue(ts.Rows[0]["SYR"] + "年");
                            sheet.Cells[2, 12].PutValue(ts.Rows[0]["SYR"] + "年");
                        }
                        foreach (DataRow d in ts.Rows)
                        {
                            sheet.Cells.CopyRow(sheet.Cells, startRow, startRow + 1);
                            sheet.Cells[startRow, 0].PutValue(d["S1_ORDER"]);
                            sheet.Cells[startRow, 1].PutValue(d["CITY"]);
                            sheet.Cells[startRow, 2].PutValue(d["TOWN"]);
                            sheet.Cells[startRow, 3].PutValue(d["WSYS_NAME"]);
                            sheet.Cells[startRow, 4].PutValue(d["AP_NAME"]);
                            sheet.Cells[startRow, 5].PutValue(d["AP_ALL_NOTE"]);
                            sheet.Cells[startRow, 6].PutValue(d["BENEFITS"]);
                            sheet.Cells[startRow, 16].PutValue(d["START_DT"]);
                            sheet.Cells[startRow, 17].PutValue(d["END_DT"]);
                            sheet.Cells[startRow, 18].PutValue(d["DEPT_NAME"]);
                            sheet.Cells[startRow, 19].PutValue(d["SC_SCORE1"] != DBNull.Value ? d["SC_SCORE1"] : "-");
                            sheet.Cells[startRow, 20].PutValue(d["SC_SCORE"] != DBNull.Value ? d["SC_SCORE"] : "-");
                            sheet.Cells[startRow, 21].PutValue(d["S1_COMT"]); // 初評意見
                            sheet.Cells[startRow, 22].PutValue(d["AP_NOTE"]);
                            sheet.Cells[startRow, 23].PutValue((d["EXT_FG"] != DBNull.Value && Convert.ToBoolean(d["EXT_FG"])) ? "延續性工程" : "新建工程");
                            // 計算用，需隱藏
                            sheet.Cells[startRow, 24].PutValue(d["REXE_FUND"]);
                            sheet.Cells[startRow, 25].PutValue(d["CCEN_FUND"]);
                            sheet.Cells[startRow, 26].PutValue(d["FCEN_FUND"]);
                            sheet.Cells[startRow, 27].PutValue(d["CLOC_FUND"]);
                            sheet.Cells[startRow, 28].PutValue(d["FLOC_FUND"]);
                            startRow++;
                        }
                        break;
                }

                sheet.Cells.DeleteRow(startRow);
                sheet.AutoFitColumns();
                sheet.Cells.HideColumns(hiddenCol, 5);

                // 備註設定(直接顯示在資料最下面)
                Helper.SetFooterRemark(sheet, type, startRow, hiddenCol);

                // 頁首頁尾設定、列印設定
                Aspose.Cells.PageSetup ps = sheet.PageSetup;
                ps.SetHeader(1, header);
                ps.SetFooter(0, footer);
                Helper.SetPageSetup(ps, sheet);
            }
        }

        /// <summary>
        /// 污水計畫評分表
        /// </summary>
        public class SewageScore
        {
            /// <summary>
            /// 將原始方法轉換為使用 Aspose.Cells.Worksheet
            /// </summary>
            /// <param name="sheet">Aspose.Cells.Worksheet 物件</param>
            /// <param name="dvAPM">DataView</param>
            /// <param name="oDt">DataTable</param>
            /// <returns>設定完成的 Aspose.Cells.Worksheet</returns>
            public Worksheet Sge0035MyPrint1_NEW(Worksheet sheet, DataView dvAPM, DataTable oDt)
            {
                CpcClass.MyFunc CpcFun = new CpcClass.MyFunc();
                int dvApLen = dvAPM.Count;

                // Aspose 使用 0 基底索引，但原始邏輯從 1 基底的第 5 行開始
                // 我們將保持 iRow 為 1 基底，但在所有 Aspose 呼叫中都減 1
                int iRow = 5;
                int iRow_SUB = 0;
                Double iS_FUND = 0;
                Double iS_CEN_FUND = 0;
                Double iS_LOC_FUND = 0;
                Double tS_FUND = 0;
                Double tS_CEN_FUND = 0;
                Double tS_LOC_FUND = 0;

                Double tSC1 = 0;
                Double tSC2 = 0;
                Double tSC3 = 0;
                Double tSC4 = 0;
                Double tSC5 = 0;
                Double tSC6 = 0;
                Double tSC7 = 0;
                Double t = 0;

                Double tT_FUND = 0;
                Double tR_FUND = 0;
                Double tM_FUND = 0;
                Double tM_CEN_FUND = 0;
                Double tM_LOC_FUND = 0;
                string sumRow = "";
                int iCol = 1; // 同樣，保持 1 基底，但在呼叫時減 1
                string sCHK_STAT = "";
                int iRow_B = 0;
                DataRow[] oRow_B = null;
                DataRow DR_B = null;
                DataRow[] oRow_APM = null;
                DataRow rowViewAp = null;
                int iRow_APM = 0;
                DataTable oDt_BUDGET_TP = oDt.DefaultView.ToTable(true, new string[] { "BUDGET_TP", "BUDGET_TP_N" });
                oRow_B = oDt_BUDGET_TP.Select();
                string BUDGET_TP = "";
                string BUDGET_TP_N = "";
                iRow_B = oRow_B.Length;

                for (int i = 0; i < iRow_B; i++)
                {
                    iRow += 1;
                    iRow_SUB = iRow;
                    iCol = 1;
                    DR_B = oRow_B[i];
                    BUDGET_TP = DR_B["BUDGET_TP"].ToString();
                    if (string.IsNullOrEmpty(BUDGET_TP))
                    {
                        BUDGET_TP_N = "無分類";
                    }
                    else
                    {
                        BUDGET_TP_N = DR_B["BUDGET_TP_N"].ToString();
                    }

                    // Aspose 索引: [row, col] -> [row-1, col-1]
                    // Aspose 範圍: CreateRange(firstRow, firstCol, totalRows, totalCols)
                    // 原始: sheet[iRow, 1, iRow, 23].Merge();
                    Range mergeRange = sheet.Cells.CreateRange(iRow - 1, 0, 1, 23);
                    mergeRange.Merge();

                    // 存取儲存格並設定樣式
                    Cell headerCell = sheet.Cells[iRow - 1, iCol - 1]; // iCol is 1, so col-index is 0
                    headerCell.PutValue(BUDGET_TP_N); //項次
                    Style headerStyle = headerCell.GetStyle();
                    headerStyle.IsTextWrapped = true;
                    headerStyle.HorizontalAlignment = TextAlignmentType.Left;
                    headerStyle.ForegroundColor = System.Drawing.Color.LightBlue;
                    headerStyle.Pattern = BackgroundType.Solid; // Aspose 需要 Pattern 才能顯示 FG Color
                    headerCell.SetStyle(headerStyle);

                    // sheet[iRow, 26].CellStyle.Color = System.Drawing.Color.LightBlue;
                    Cell cell26 = sheet.Cells[iRow - 1, 25]; // 26 - 1 = 25
                    Style style26 = cell26.GetStyle();
                    style26.ForegroundColor = System.Drawing.Color.LightBlue;
                    style26.Pattern = BackgroundType.Solid;
                    cell26.SetStyle(style26);

                    oRow_APM = oDt.Select("isnull(BUDGET_TP,'')='" + BUDGET_TP + "'", "CITY, S_ORDER, AP_NAME");
                    iRow_APM = oRow_APM.Length;
                    iS_FUND = 0;
                    iS_CEN_FUND = 0;
                    iS_LOC_FUND = 0;

                    for (int j = 0; j < iRow_APM; j++)
                    {
                        iRow += 1;
                        iCol = 1;
                        rowViewAp = oRow_APM[j];

                        // --- 輔助函式 (Helper function) ---
                        // 為了簡潔，我們將使用一個本地函式來設定儲存格的值、換行和數字格式
                        // 這會使轉換後的程式碼更易讀
                        Action<int, object, string, bool> SetCell =
                            (colIdx, value, format, wrap) =>
                            {
                                Cell cell = sheet.Cells[iRow - 1, colIdx - 1];

                                // 嘗試將值設為 double (如果格式是數字)
                                if (!string.IsNullOrEmpty(format))
                                {
                                    double numValue = Convert.ToDouble(value);
                                    cell.PutValue(numValue);
                                }
                                else
                                {
                                    cell.PutValue(value);
                                }

                                Style style = cell.GetStyle();
                                if (wrap)
                                {
                                    style.IsTextWrapped = true;
                                }
                                if (!string.IsNullOrEmpty(format))
                                {
                                    //style.NumberFormat = format;
                                }
                                cell.SetStyle(style);
                            };
                        // ---------------------------------




                        //基本資料
                        SetCell(iCol, (j + 1).ToString(), null, false); iCol++; //項次
                        SetCell(iCol, rowViewAp["AP_PJ_NAME"].ToString(), null, true); iCol++; //申請項目 Or 申請案件類別
                        SetCell(iCol, rowViewAp["TD_CODE"].ToString(), null, true); iCol++; //工程編號
                        SetCell(iCol, rowViewAp["AP_NAME"].ToString(), null, true); iCol++; //工程名稱
                        SetCell(iCol, rowViewAp["AP_ALL_NOTE"].ToString(), null, true); iCol++; //全期內容
                        SetCell(iCol, rowViewAp["AP_YEAR_NOTE"].ToString(), null, true); iCol++; //本年度辦理內容

                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["T_FUND"]), "###,##0", false); iCol++; //工程總經費
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["R_FUND"]), "###,##0", false); iCol++; //以前年度執行經費
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["M_FUND"]), "###,##0", false); iCol++; //本年度申請經費
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["M_CEN_FUND"]), "###,##0", false); iCol++; //中央補助款
                        SetCell(iCol, CpcFun.MyCDbl(rowViewAp["M_LOC_FUND"]), "###,##0", false); iCol++; //地方配合款
                        try
                        {
                            SetCell(iCol, rowViewAp["SC1"] != DBNull.Value ? rowViewAp["SC1"] : "", null, false); iCol++; // 評分項目1
                            SetCell(iCol, rowViewAp["SC2"] != DBNull.Value ? rowViewAp["SC2"] : "", null, false); iCol++; // 評分項目2
                            SetCell(iCol, rowViewAp["SC3"] != DBNull.Value ? rowViewAp["SC3"] : "", null, false); iCol++; // 評分項目3
                            SetCell(iCol, rowViewAp["SC4"] != DBNull.Value ? rowViewAp["SC4"] : "", null, false); iCol++; // 評分項目4
                            SetCell(iCol, rowViewAp["SC5"] != DBNull.Value ? rowViewAp["SC5"] : "", null, false); iCol++; // 評分項目5
                            SetCell(iCol, rowViewAp["SC6"] != DBNull.Value ? rowViewAp["SC6"] : "", null, false); iCol++; // 評分項目6
                            SetCell(iCol, rowViewAp["SC7"] != DBNull.Value ? rowViewAp["SC7"] : "", null, false); iCol++; // 評分項目7

                            //try { SetCell(iCol, rowViewAp["SC1"], "##0", false); } catch { SetCell(iCol, 0, "##0", false); }
                            //iCol++; //1
                            //try { SetCell(iCol, rowViewAp["SC2"], "##0", false); } catch { SetCell(iCol, 0, "##0", false); }
                            //iCol++; //2
                            //try { SetCell(iCol, rowViewAp["SC3"], "##0", false); } catch { SetCell(iCol, 0, "##0", false); }
                            //iCol++; //3
                            //try { SetCell(iCol, rowViewAp["SC4"], "##0", false); } catch { SetCell(iCol, 0, "##0", false); }
                            //iCol++; //4
                            //try { SetCell(iCol, rowViewAp["SC5"], "##0", false); } catch { SetCell(iCol, 0, "##0", false); }
                            //iCol++; //5
                            //try { SetCell(iCol, rowViewAp["SC6"], "##0", false); } catch { SetCell(iCol, 0, "##0", false); }
                            //iCol++; //6
                            //try { SetCell(iCol, rowViewAp["SC7"], "##0", false); } catch { SetCell(iCol, 0, "##0", false); }
                            //iCol++; //7

                            try { tT_FUND += CpcFun.MyCDbl(rowViewAp["T_FUND"]); } catch { }
                            //try { tT_FUND += CpcFun.MyCDbl((string)rowViewAp["T_FUND"]); } catch { }
                            try { tR_FUND += CpcFun.MyCDbl(rowViewAp["R_FUND"]); } catch { }
                            try { tM_FUND += CpcFun.MyCDbl(rowViewAp["M_FUND"]); } catch { }
                            try { tM_CEN_FUND += CpcFun.MyCDbl(rowViewAp["M_CEN_FUND"]); } catch { }
                            try { tM_LOC_FUND += CpcFun.MyCDbl(rowViewAp["M_LOC_FUND"]); } catch { }

                            try { tSC1 += CpcFun.MyCDbl(rowViewAp["SC1"]); } catch { }
                            try { tSC2 += CpcFun.MyCDbl(rowViewAp["SC2"]); } catch { }
                            try { tSC3 += CpcFun.MyCDbl(rowViewAp["SC3"]); } catch { }
                            try { tSC4 += CpcFun.MyCDbl(rowViewAp["SC4"]); } catch { }
                            try { tSC5 += CpcFun.MyCDbl(rowViewAp["SC5"]); } catch { }
                            try { tSC6 += CpcFun.MyCDbl(rowViewAp["SC6"]); } catch { }
                            try { tSC7 += CpcFun.MyCDbl(rowViewAp["SC7"]); } catch { }
                        }
                        catch
                        {
                            continue;
                        }


                        // 公式: Aspose 也接受 A1 標記法，公式保持不變 (iRow 是 1 基底的)
                        Cell formulaCell = sheet.Cells[iRow - 1, iCol - 1];
                        formulaCell.Formula = "=SUM(L" + iRow.ToString() + ":R" + iRow.ToString() + ")";
                        Style formulaStyle = formulaCell.GetStyle();
                        //formulaStyle.NumberFormat = "##0";
                        formulaCell.SetStyle(formulaStyle);
                        iCol++; //總評分

                        try { SetCell(iCol, rowViewAp["PRV_CON_SEWERS"], "###,##0", false); iCol++; } catch { iCol++; }//當年度可提供用戶接管數
                        try { SetCell(iCol, rowViewAp["S_ORDER"], "##0", false); iCol++; } catch { iCol++; } //辦理優先次序

                        sCHK_STAT = rowViewAp["CHK_STAT"].ToString(); //本年度是否補助
                        switch (sCHK_STAT)
                        {
                            case "Y":
                                sheet.Cells[iRow - 1, iCol - 1].PutValue("V"); iCol++; //是
                                iCol++; //否
                                break;
                            case "N":
                                iCol++; //是
                                sheet.Cells[iRow - 1, iCol - 1].PutValue("V"); iCol++; //否
                                break;
                            default:
                                iCol++; //是
                                iCol++; //否
                                break;
                        }

                        try { SetCell(iCol, CpcFun.MyCDbl(rowViewAp["S_FUND"]), "###,##0", false); iCol++; } catch { }//本年度核定經費
                        try { SetCell(iCol, CpcFun.MyCDbl(rowViewAp["S_CEN_FUND"]), "###,##0", false); iCol++; } catch { }//中央補助款
                        try { SetCell(iCol, CpcFun.MyCDbl(rowViewAp["S_LOC_FUND"]), "###,##0", false); iCol++; } catch { }//地方配合款

                        try { iS_FUND += CpcFun.MyCDbl(rowViewAp["S_FUND"]); } catch { }
                        try { iS_CEN_FUND += CpcFun.MyCDbl(rowViewAp["S_CEN_FUND"]); } catch { }
                        try { iS_LOC_FUND += CpcFun.MyCDbl(rowViewAp["S_LOC_FUND"]); } catch { }
                    }
                    tS_FUND += iS_FUND;
                    tS_CEN_FUND += iS_CEN_FUND;
                    tS_LOC_FUND += iS_LOC_FUND;

                    // 邊框設定:
                    // 原始碼: sheet[3, 1, iRows, iCols].CellStyle.Borders...
                    // 這看起來是在迴圈內反覆設定一個大範圍的邊框，這很奇怪且效率低下。
                    // 我們將其移到迴圈外部，在所有資料都填入後再執行。
                    // (請參見迴圈後的 "總邊框設定")

                    // 設定小計 (Subtotal)
                    // 輔助函式 (Helper function)
                    Action<int, double, string, Color> SetSubTotalCell =
                        (colIdx, value, format, bgColor) =>
                        {
                            Cell cell = sheet.Cells[iRow_SUB - 1, colIdx - 1];
                            cell.PutValue(value);
                            Style style = cell.GetStyle();
                            //style.NumberFormat = format;
                            style.ForegroundColor = bgColor;
                            style.Pattern = BackgroundType.Solid;
                            cell.SetStyle(style);
                        };

                    SetSubTotalCell(24, iS_FUND, "###,##0", Color.LightPink);
                    SetSubTotalCell(25, iS_CEN_FUND, "###,##0", Color.LightPink);
                    SetSubTotalCell(26, iS_LOC_FUND, "###,##0", Color.LightPink);
                }

                iRow += 1;

                sheet.Cells[iRow - 1, 1].PutValue("總計"); // col 2 -> index 1

                // 輔助函式 (Helper function)
                Action<int, double, string> SetGrandTotalCell =
                    (colIdx, value, format) =>
                    {
                        Cell cell = sheet.Cells[iRow - 1, colIdx - 1];
                        cell.PutValue(value);
                        Style style = cell.GetStyle();
                        //style.NumberFormat = format;
                        cell.SetStyle(style);
                    };

                // 總計值
                SetGrandTotalCell(7, tT_FUND, "###,##0"); //工程總經費
                SetGrandTotalCell(8, tR_FUND, "###,##0"); //以前年度執行經費
                SetGrandTotalCell(9, tM_FUND, "###,##0"); //本年度申請經費
                SetGrandTotalCell(10, tM_CEN_FUND, "###,##0"); //中央補助款
                SetGrandTotalCell(11, tM_LOC_FUND, "###,##0"); //地方配合款

                // 評分總計 (原始碼中被註解了)
                //SetGrandTotalCell(12, tSC1, "###,##0");
                // ... (tSC2 to tSC7) ...

                SetGrandTotalCell(24, tS_FUND, "###,##0"); //本年度核定經費
                SetGrandTotalCell(25, tS_CEN_FUND, "###,##0"); //中央補助款
                SetGrandTotalCell(26, tS_LOC_FUND, "###,##0"); //地方配合款


                // 總計列的樣式和邊框
                // 原始: sheet[iRow, 1, iRow, 27]
                Range gTotalRange = sheet.Cells.CreateRange(iRow - 1, 0, 1, 27); // (row, col, totalRows, totalCols)
                Style gTotalStyle = sheet.Cells[iRow - 1, 0].GetStyle(); // 取得一個樣式物件來修改
                gTotalStyle.ForegroundColor = System.Drawing.Color.LawnGreen;
                gTotalStyle.Pattern = BackgroundType.Solid;

                // 設定邊框
                gTotalStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                gTotalStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                gTotalStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                gTotalStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                gTotalStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                gTotalStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                gTotalStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                gTotalStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                // 建立一個 StyleFlag 來僅套用邊框和底色
                StyleFlag flag = new StyleFlag();
                flag.Borders = true;
                flag.CellShading = true;

                gTotalRange.ApplyStyle(gTotalStyle, flag);


                // *** 總邊框設定 ***
                // 原始碼中這段在迴圈內，這裡我們移到外面
                // 原始: sheet[3, 1, iRows, iCols]
                // iRows 和 iCols 在這裡是 "總行數" 和 "總列數"
                // 我們將範圍設定為從 (3,1) 到 (iRow-1, 27) (總計列的上一列)
                // r1=3, c1=1, r2=iRow-1, c2=27 (假設 27 是最大欄寬)
                int r1_idx = 3 - 1; // 2
                int c1_idx = 1 - 1; // 0
                int r2_idx = (iRow - 1) - 1; // 總計列的上一列
                int c2_idx = 27 - 1; // 26

                Range dataRange = sheet.Cells.CreateRange(r1_idx, c1_idx, (r2_idx - r1_idx + 1), (c2_idx - c1_idx + 1));

                // SetOutlineBorder 只會設定範圍的 "外" 邊框
                // 原始碼的 .Borders[ExcelBordersIndex.EdgeLeft] 看起來也是這個意思
                dataRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                dataRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                dataRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                dataRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                // 如果您需要 "所有" 儲存格都有邊框，請取消註解下面這段
                /*
                Style allBordersStyle = sheet.Cells[r1_idx, c1_idx].GetStyle();
                allBordersStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                allBordersStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                allBordersStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                allBordersStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                allBordersStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                allBordersStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                allBordersStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                allBordersStyle.Borders[BorderType.RightBorder].Color = Color.Black;
                StyleFlag allBordersFlag = new StyleFlag() { Borders = true };
                dataRange.ApplyStyle(allBordersStyle, allBordersFlag);
                */

                // 備註
                iRow += 2;
                sheet.Cells[iRow - 1, 4].PutValue("備註："); // col 5 -> index 4
                sheet.Cells[iRow - 1, 5].PutValue("一、中央審核係依據「中央對直轄市及縣(市)政府申請計畫型補助款補助之處理原則」規定辦理，其評分項目、標準如下："); // col 6 -> index 5
                sheet.Cells.Merge(iRow - 1, 5, 1, 20); // 合併儲存格
                iRow += 1;
                sheet.Cells[iRow - 1, 5].PutValue("  （一）應行配合事項之辦理情形（縣市政府勘查評估事項1-6項）、（二）基本設計完成進度（縣市政府勘查評估事項7項）、（三）以前年度計畫及預算執行情形、（四）可行性及預期效益、");
                sheet.Cells.Merge(iRow - 1, 5, 1, 20); // 合併儲存格
                iRow += 1;
                sheet.Cells[iRow - 1, 5].PutValue("  （五）與中央政策配合程度、（六）負擔經費之財源籌措及相關財務計畫規劃情形（縣市政府勘查評估事項9-10項）、（七）其他審查及評比之項目（人力、收費標準等）");
                sheet.Cells.Merge(iRow - 1, 5, 1, 20); // 合併儲存格
                iRow += 1;
                sheet.Cells[iRow - 1, 5].PutValue("二、各工程之總評分，由各單項評分後加總。");
                sheet.Cells.Merge(iRow - 1, 5, 1, 20); // 合併儲存格
                iRow += 1;
                sheet.Cells[iRow - 1, 5].PutValue("三、以前年度計畫項目及預算執行情形之審核，得包括工程進度及工程品質。");
                sheet.Cells.Merge(iRow - 1, 5, 1, 20); // 合併儲存格
                iRow += 1;
                sheet.Cells[iRow - 1, 5].PutValue("四、其他審查及評比之項目，包括下水道使用費徵收自治條例核定情形及用戶接管辦理情形等。");
                sheet.Cells.Merge(iRow - 1, 5, 1, 20); // 合併儲存格

                // 設定列印範圍
                Aspose.Cells.PageSetup ps = sheet.PageSetup;
                string printArea = $"A1:{CellsHelper.CellIndexToName(sheet.Cells.MaxDataRow + 1 - 1, sheet.Cells.MaxDataColumn + 1 - 1)}";
                ps.PrintArea = printArea;

                return sheet;
            }
        }

        /// <summary>
        /// 污水計畫明細表
        /// </summary>
        public class SewageDetail
        {
            /// <summary>
            /// 將 Sge0035MyPrint2_NEW (使用 Spire/IWorksheet) 轉換為使用 Aspose.Cells.Worksheet
            /// </summary>
            /// <param name="sheet">Aspose.Cells.Worksheet 物件</param>
            /// <param name="dvAPM">DataView</param>
            /// <param name="oDt">DataTable</param>
            /// <returns>設定完成的 Aspose.Cells.Worksheet</returns>
            public Worksheet Sge0035MyPrint2_NEW_Aspose(Worksheet sheet, DataView dvAPM, DataTable oDt)
            {
                CpcClass.MyFunc CpcFun = new CpcClass.MyFunc();
                int dvApLen = dvAPM.Count;

                // Aspose 使用 0 基底索引，但原始邏輯從 1 基底的第 3 行開始
                int iRow = 3;
                int iCol = 1;

                DataRow[] oRow_B = null;
                DataRow DR_B = null;
                DataRow[] oRow_APM = null;
                DataRow rowViewAp = null;
                int iRow_APM = 0;
                DataTable oDt_BUDGET_TP = oDt.DefaultView.ToTable(true, new string[] { "BUDGET_TP", "BUDGET_TP_N" });
                oRow_B = oDt_BUDGET_TP.Select();
                string BUDGET_TP = "";
                string BUDGET_TP_N = "";
                int iRow_B = oRow_B.Length;

                for (int i = 0; i < iRow_B; i++)
                {
                    iRow += 1; // 標題列 (第 4 行開始)
                    iCol = 1;
                    DR_B = oRow_B[i];
                    BUDGET_TP = DR_B["BUDGET_TP"].ToString();
                    if (string.IsNullOrEmpty(BUDGET_TP)) // 原始: BUDGET_TP == ""
                    {
                        BUDGET_TP_N = "無分類";
                    }
                    else
                    {
                        BUDGET_TP_N = DR_B["BUDGET_TP_N"].ToString();
                    }

                    // 原始: sheet[iRow, 1, iRow, 59].Merge();
                    // Aspose: CreateRange(firstRow, firstCol, totalRows, totalCols)
                    Range mergeRange = sheet.Cells.CreateRange(iRow - 1, 0, 1, 60);
                    mergeRange.Merge();

                    // 原始: sheet[iRow, iCol].Value = BUDGET_TP_N;
                    // 原始: sheet[iRow, iCol].CellStyle.WrapText = true;
                    Cell headerCell = sheet.Cells[iRow - 1, iCol - 1]; // iCol is 1, so col-index is 0
                    headerCell.PutValue(BUDGET_TP_N);
                    Style headerStyle = headerCell.GetStyle();
                    headerStyle.IsTextWrapped = true;
                    headerCell.SetStyle(headerStyle);

                    oRow_APM = oDt.Select("isnull(BUDGET_TP,'')='" + BUDGET_TP + "'", "CITY, S_ORDER, AP_NAME");
                    iRow_APM = oRow_APM.Length;

                    for (int j = 0; j < iRow_APM; j++)
                    {
                        iRow += 1; // 資料列
                        iCol = 1;
                        rowViewAp = oRow_APM[j];

                        // --- 輔助函式 (Helper function) ---
                        // 為了簡潔，我們將使用一個本地函式來設定儲存格的值、換行和數字格式
                        Action<int, object, string, bool> SetCell =
                            (colIdx, value, format, wrap) =>
                            {
                                Cell cell = sheet.Cells[iRow - 1, colIdx - 1];

                                // 嘗試將值設為 double (如果格式是數字)
                                if (!string.IsNullOrEmpty(format))
                                {
                                    double numValue = CpcFun.MyCDbl(value);
                                    cell.PutValue(numValue);
                                }
                                else
                                {
                                    // PutValue 可以安全地處理字串或 DBNull.ToString()
                                    cell.PutValue(value);
                                }

                                Style style = cell.GetStyle();
                                //if(wrap) {
                                //    style.IsTextWrapped = true;
                                //}
                                //if(format != null) {
                                //    style.NumberFormat = format;
                                //}
                                cell.SetStyle(style);
                            };
                        // ---------------------------------

                        //基本資料
                        SetCell(iCol, (j + 1).ToString(), null, false); iCol++; //項次
                        SetCell(iCol, rowViewAp["PJ_CATY_NAME"].ToString(), null, true); iCol++; //申請項目 Or 申請案件類別
                        SetCell(iCol, rowViewAp["TD_CODE"].ToString(), null, true); iCol++; //工程編號
                        SetCell(iCol, rowViewAp["AP_NAME"].ToString(), null, true); iCol++; //工程名稱
                        SetCell(iCol, rowViewAp["AP_PJ_NAME"].ToString(), null, false); iCol++; //實施計畫名稱 (原始碼沒有 WrapText)
                        SetCell(iCol, rowViewAp["CITY"].ToString(), null, false); iCol++; //縣市別
                        SetCell(iCol, rowViewAp["FILL_DP_NAME"].ToString(), null, true); iCol++; //受補助單位
                        SetCell(iCol, rowViewAp["EXE_MCH"].ToString(), null, true); iCol++; //預定執行單位
                        SetCell(iCol, rowViewAp["START_DT_CHT"].ToString(), null, false); iCol++; //預定開始日期
                        SetCell(iCol, rowViewAp["END_DT_CHT"].ToString(), null, false); iCol++; //預定完成日期
                        SetCell(iCol, rowViewAp["BENEFITS"].ToString(), null, true); iCol++; //預期效益
                        SetCell(iCol, rowViewAp["AP_ALL_NOTE"].ToString(), null, true); iCol++; //全期內容
                        SetCell(iCol, rowViewAp["AP_YEAR_NOTE"].ToString(), null, true); iCol++; //本年度辦理內容
                        SetCell(iCol, rowViewAp["AP_NOTE"].ToString(), null, true); iCol++; //備註說明

                        //計畫聯絡人
                        SetCell(iCol, rowViewAp["ADV_DP_NAME"].ToString(), null, true); iCol++; //填表單位
                        SetCell(iCol, rowViewAp["CONT_NAME"].ToString(), null, true); iCol++; //姓名
                        SetCell(iCol, rowViewAp["CONT_TITLE"].ToString(), null, true); iCol++; //職稱
                        SetCell(iCol, rowViewAp["CONT_TEL"].ToString(), null, true); iCol++; //電話
                        SetCell(iCol, rowViewAp["CONT_FAX"].ToString(), null, true); iCol++; //傳真
                        SetCell(iCol, rowViewAp["CONT_EMAIL"].ToString(), null, true); iCol++; //E-Mail

                        //預定
                        SetCell(iCol, rowViewAp["TF106_DT_CHT"].ToString(), null, false); iCol++; //預定徵收公告日
                        SetCell(iCol, rowViewAp["TF108_DT_CHT"].ToString(), null, false); iCol++; //預定用地取得完成日
                        SetCell(iCol, rowViewAp["TK103_DT_CHT"].ToString(), null, false); iCol++; //預定設計完成日
                        SetCell(iCol, rowViewAp["TK105_DT_CHT"].ToString(), null, false); iCol++; //預定發包日
                        SetCell(iCol, rowViewAp["TK106_DT_CHT"].ToString(), null, false); iCol++; //預定開工日
                        SetCell(iCol, rowViewAp["TK107_DT_CHT"].ToString(), null, false); iCol++; //預定完工日
                        SetCell(iCol, rowViewAp["SYS_CATY_NAME"].ToString(), null, false); iCol++; //系統別
                        SetCell(iCol, rowViewAp["IPJ_CODE_DESC"].ToString(), null, false); iCol++; //辦理層級

                        //初審作業
                        // 原始: .Value = ...ToString(); .NumberFormat = "##0"
                        SetCell(iCol, rowViewAp["SC1_SCORE"], "##0", false); iCol++; //1.應行配合事項之辦理情形(15分)
                        SetCell(iCol, rowViewAp["SC2_SCORE"], "##0", false); iCol++; //2.基本設計完成進度(15分)
                        SetCell(iCol, rowViewAp["SC3_SCORE"], "##0", false); iCol++; //3.以前年度計畫及預算執行情形(15分)
                        SetCell(iCol, rowViewAp["SC4_SCORE"], "##0", false); iCol++; //4.可行性及預期效益(15分)
                        SetCell(iCol, rowViewAp["SC5_SCORE"], "##0", false); iCol++; //5.與中央政策配合程度(15分)
                        SetCell(iCol, rowViewAp["SC6_SCORE"], "##0", false); iCol++; //6.負擔經費之財源籌措及相關財務計畫規劃情形(15分)
                        SetCell(iCol, rowViewAp["SC7_SCORE"], "##0", false); iCol++; //7.其他審查及評比之項目(10分)
                        SetCell(iCol, rowViewAp["S1_ORDER"].ToString(), null, false); iCol++; //縣市評估優先次序
                        SetCell(iCol, GetStat(rowViewAp["CHK_STAT"].ToString()), null, false); iCol++; //本年度是否補助
                        //縣市政府勘查評估事項
                        SetCell(iCol, GetStat(rowViewAp["SWW1_FG"].ToString()), null, false); iCol++; //1.符合系統規劃
                        SetCell(iCol, GetStat(rowViewAp["SWW2_FG"].ToString()), null, false); iCol++; //2.實施計畫是否通過
                        SetCell(iCol, GetStat(rowViewAp["SWW3_FG"].ToString()), null, false); iCol++; //3.用地變更程序
                        SetCell(iCol, GetStat(rowViewAp["SWW4_FG"].ToString()), null, false); iCol++; //4.用地取得
                        SetCell(iCol, GetStat(rowViewAp["SWW5_FG"].ToString()), null, false); iCol++; //5.下游管線廠站建設期程是否配合
                        SetCell(iCol, GetStat(rowViewAp["SWW6_FG"].ToString()), null, false); iCol++; //6.交維問題可否解決
                        SetCell(iCol, GetStat(rowViewAp["SWW7_FG"].ToString()), null, false); iCol++; //7.地下管線拆遷
                        SetCell(iCol, GetStat(rowViewAp["SWW8_FG"].ToString()), null, false); iCol++; //8.本工程已完成設計
                        SetCell(iCol, rowViewAp["SWW9_RATE"], "##0", false); iCol++; //9.本工程前一年度執行率(%)(新增工程免填)
                        SetCell(iCol, GetStat(rowViewAp["SWW10_FG"].ToString()), null, false); iCol++; //10.前一年度配合款編列情形
                        SetCell(iCol, GetStat(rowViewAp["SWW11_FG"].ToString()), null, false); iCol++; //11.是否可配合納入預算
                        SetCell(iCol, rowViewAp["SWWOT_DESC"].ToString(), null, true); iCol++; //其他

                        // 經費 (取代 SetIRange)
                        string fundFormat = "###,##0";
                        SetCell(iCol, rowViewAp["M_TCEN_FUND"], fundFormat, false); iCol++; //計畫總經費 中央補助款
                        SetCell(iCol, rowViewAp["M_TLOC_FUND"], fundFormat, false); iCol++; //計畫總經費 地方配合款
                        SetCell(iCol, rowViewAp["M_RCEN_FUND"], fundFormat, false); iCol++; //以前年度已編列 中央補助款
                        SetCell(iCol, rowViewAp["M_RLOC_FUND"], fundFormat, false); iCol++; //以前年度已編列 地方配合款
                        SetCell(iCol, rowViewAp["M_CEN_FUND"], fundFormat, false); iCol++; //本年度申請經費 中央補助款
                        SetCell(iCol, rowViewAp["M_LOC_FUND"], fundFormat, false); iCol++; //本年度申請經費 地方配合款
                        SetCell(iCol, rowViewAp["M_CEN_JUD_FUND"], fundFormat, false); iCol++; //本年度核定經費 初審 中央補助款
                        SetCell(iCol, rowViewAp["M_LOC_JUD_FUND"], fundFormat, false); iCol++; //本年度核定經費 初審 地方配合款
                        SetCell(iCol, rowViewAp["M_CEN_DCS_FUND"], fundFormat, false); iCol++; //本年度核定經G 複審 中央補助款
                        SetCell(iCol, rowViewAp["M_LOC_DCS_FUND"], fundFormat, false); iCol++; //本年度核定經G 複審 地方配合款

                        SetCell(iCol, Convert.ToBoolean(rowViewAp["EXT_FG"]) == true ? "延續性工程" : "新建工程", null, true); iCol++; //辦理狀態
                    }
                }

                // 原始:
                // int iRows = sheet.Rows.Length;
                // int iCols = sheet.Columns.Length;
                // sheet[1, 1, iRows, iCols].CellStyle.Borders...

                // Aspose 版本:
                // 我們從 (1,1) 開始設定邊框
                // 總行數是 iRow
                // 總欄數是 59 (根據上面的 merge)

                int r1_idx = 0; // 1 - 1
                int c1_idx = 0; // 1 - 1
                int r2_idx = iRow - 1; // 總行數
                int c2_idx = 60 - 1; // 總欄數

                // +1 是因為 CreateRange 需要 "總行數", "總欄數"
                Range dataRange = sheet.Cells.CreateRange(r1_idx, c1_idx, (r2_idx - r1_idx + 1), (c2_idx - c1_idx + 1));

                // SetOutlineBorder 對應於 ExcelBordersIndex.Edge...
                dataRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                dataRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                dataRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                dataRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                // 設定列印範圍
                Aspose.Cells.PageSetup ps = sheet.PageSetup;
                string printArea = $"A1:{CellsHelper.CellIndexToName(sheet.Cells.MaxDataRow + 1 - 1, sheet.Cells.MaxDataColumn + 1 - 1)}";
                ps.PrintArea = printArea;

                return sheet;
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
        }
    }
}