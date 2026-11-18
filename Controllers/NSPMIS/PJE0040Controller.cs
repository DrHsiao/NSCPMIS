using PagedList;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.ViewModels.NewModels;

namespace WCals.Controllers.NSPMIS
{
    public class PJE0040Controller : Controller
    {
        private readonly Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();
        int pageSize = 10;
        public ActionResult Index(PJE0040Model model)
        {
            //string userid = Helpers.WebSiteHelper.getUserId();
            int page = (model.page == 0) ? 1 : model.page;

            if (!string.IsNullOrEmpty(model.SEQ_NO_M))
            {




                return View("PJE0040Index", model);
            }
            /*
            string sql = "select P.AN_PID, '('+F.SYR+')'+F.PJ_NAME as AN_NAME from PJMFORM P left join " +
            "PJMFORM F on P.AN_PID = F.PID where P.PID in (select PID from TDV_TDDCFTS where " +
            "ADM_DP = '" + WebSiteHelper.getUserDept() + "'and F.PJ_TYPE='W')group by P.AN_PID, " +
            "F.SYR,F.PJ_NAME order by F.SYR desc,F.PJ_NAME";
            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List(sql);

            if(dt.Rows.Count == 0) {
                return View("PJE0040Index", model);
            }

            ViewBag.SYRS = tools.ToSelectList(dt, "AN_PID", "AN_NAME", null);

            sql = "select PID, '('+SYR+')'+PJ_NAME as PNAME\r\nfrom PJMFORM \r\nwhere AN_PID = " + dt.Rows[0]["AN_PID"] + "\r\nand PID in (select PID from TDV_TDDCFTS where ADM_DP ='" + WebSiteHelper.getUserId() + "')\r\ngroup by PID,SYR,PJ_NAME\r\norder by PID";
            dt = BASE_DAO.SqlDao.CMEPMSQuery4List(sql);

            ViewBag.SEQ_NO_MS = tools.ToSelectList(dt, "PID", "PNAME", null); ;
            string pid = Convert.ToString(dt.Rows[0]["PID"]);
            dt = BASE_DAO.SqlDao.CMEPMSQuery4List("select SEQ_NO, '第'+Convert(varchar,SEQ_NO)+'期'as SEQ_NAME from PJDFDCR where PID ='" + dt.Rows[0]["PID"] + "'ORDER BY SEQ_NO DESC ");
            DataRow dr = dt.NewRow();
            dr["SEQ_NO"] = 0;
            dr["SEQ_NAME"] = "原預算";
            dt.Rows.Add(dr);
            ViewBag.SEQ_NO_SS = tools.ToSelectList(dt, "SEQ_NO", "SEQ_NAME", null);
            */



            string sql = "select P.AN_PID, '('+AP.SYR + ')'+AP.PJ_NAME as PJ_NAME " +
            "from PJMFORM P left join PJMFORM AP on P.AN_PID = AP.PID and AP.PJ_TYPE = 'W'" +
            "where P.PID in (select PID from TDV_TDDCFLS where CITY in (select CITY from CMBASE.dbo.SECCITY " +
            "where DEPT_NO = '" + WebSiteHelper.getUserDept() + "'))and isnull(AP.PJ_NAME,'') != ''group by P.AN_PID, AP.SYR, AP.PJ_NAME " +
            "order by AP.SYR desc";

            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List(sql);

            if (dt.Rows.Count == 0)
            {
                return View("PJE0040Index", model);
            }

            ViewBag.SYRS = tools.ToSelectList(dt, "AN_PID", "PJ_NAME", "請選擇計畫");




            //均取最新期TD
            /*
            SqlCommand cmd = new SqlCommand("select * from PJV_CTDMNDS where PID = @PID ORDER BY TID DESC");
            cmd.Parameters.Add(new SqlParameter("PID", pid));
            ViewBag.List1 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            cmd = new SqlCommand("select * from PJV_CTDMNDS where PID = @PID ORDER BY TID DESC");
            cmd.Parameters.Add(new SqlParameter("PID", pid));
            ViewBag.List2 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            cmd = new SqlCommand("select * from PJV_CRDMNDS where PID = @PID ORDER BY SYR DESC, PID DESC");
            cmd.Parameters.Add(new SqlParameter("PID", pid));
            ViewBag.List3 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            */








            return View("PJE0040Index", model);
        }

        //取年度計畫
        [HttpPost]
        public JsonResult GetSEQ_NO_M(string PID)
        {
            //DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List("select PID, '('+SYR+')'+PJ_NAME as PNAME\r\nfrom PJMFORM \r\nwhere AN_PID = " + PID + "\r\nand PID in (select PID from TDV_TDDCFTS where ADM_DP ='" + WebSiteHelper.getUserId() + "')\r\ngroup by PID,SYR,PJ_NAME\r\norder by PID");
            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List("select PID, '(' + SYR + ')' +PJ_NAME as PNAME " +
                "from PJMFORM where PJ_TYPE='B'and AN_PID =" + PID + " AND PID in " +
                "(select PID from TDV_TDDCFLS where CITY in (select CITY from CMBASE.dbo.SECCITY " +
                "where DEPT_NO = '" + WebSiteHelper.getUserId() + "')) order by PID");

            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        //取分支計畫
        [HttpPost]
        public JsonResult GetSEQ_NO_S(string PID)
        {
            DataTable dt = BASE_DAO.SqlDao.CMEPMSQuery4List("select SEQ_NO, case SEQ_NO when 0 then '原預算'" +
                "else '第'+convert(nvarchar(2),SEQ_NO)+'次修正'end as f2, (select max(SEQ_NO) from PJDFDCR where PID =" + PID + ") as maxseq " +
                "from PJDFDCR where PID=" + PID + " order by SEQ_NO desc");
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        //查詢
        public ActionResult Search(string PID, int SEQ_NO)
        {

            PJE0040Model model = new PJE0040Model();

            model.PID = Convert.ToDecimal(PID);
            //int page = (model.page == 0) ? 1 : model.page;

            int SEQ_MAX = util.GetSEQ_NO(PID);

            string userDept = WebSiteHelper.getUserDept(); // 提出來只查一次，節省效能

            // 🟢 核心需求：判斷是否為最新期別，決定是否唯讀
            bool isReadOnly = (SEQ_NO != SEQ_MAX);
            ViewBag.IsReadOnly = isReadOnly;

            if (SEQ_NO == SEQ_MAX)
            {

                //均取最新期TD
                SqlCommand cmd = new SqlCommand("select * from TDV_TDDCFTS " +
                    "where PID = " + PID + " and CITY in " +
                    "(select CITY from CMBASE.dbo.SECCITY where DEPT_NO ='" + WebSiteHelper.getUserDept() + "')ORDER BY FD_CODE");
                cmd.Parameters.Add(new SqlParameter("PID", model.PID));
                ViewBag.List1 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

                cmd = new SqlCommand("select * from TDV_TDDCFTS where PID = " + PID +
                    " and CITY in(select CITY from CMBASE.dbo.SECCITY where DEPT_NO = '" + WebSiteHelper.getUserDept() + "')and YR_FUND=0 ORDER BY FD_CODE");
                cmd.Parameters.Add(new SqlParameter("PID", model.PID));
                ViewBag.List2 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

                cmd = new SqlCommand("select * from TDV_TDDRFLS where PID = " + PID + " and CITY in " +
                    "(select CITY from CMBASE.dbo.SECCITY where DEPT_NO = '" + WebSiteHelper.getUserDept() + "')ORDER BY FD_CODE");
                cmd.Parameters.Add(new SqlParameter("PID", model.PID));
                ViewBag.List3 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            }
            else
            {
                //取舊期數 
                SqlCommand cmd = new SqlCommand("select * from TDV_TDHCFTS " +
                "where PID = " + PID + " and SEQ_NO = " + SEQ_NO + " and CITY in " +
                "(select CITY from CMBASE.dbo.SECCITY where DEPT_NO = '" + WebSiteHelper.getUserDept() + "')ORDER BY FD_CODE");
                cmd.Parameters.Add(new SqlParameter("PID", model.PID));
                ViewBag.List1 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

                cmd = new SqlCommand("select * from TDV_TDDCFTS " +
                    "where PID = " + PID + " and SEQ_NO = " + SEQ_NO + " and CITY in (select CITY from " +
                    "CMBASE.dbo.SECCITY where DEPT_NO = '" + WebSiteHelper.getUserDept() + "')and YR_FUND=0 ORDER BY FD_CODE");
                cmd.Parameters.Add(new SqlParameter("PID", model.PID));
                ViewBag.List2 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

                cmd = new SqlCommand("select * from TDV_TDDRFLS where PID = " + PID + " and CITY in " +
                    "(select CITY from CMBASE.dbo.SECCITY where DEPT_NO = '" + WebSiteHelper.getUserDept() + "')ORDER BY FD_CODE");
                cmd.Parameters.Add(new SqlParameter("PID", model.PID));
                ViewBag.List3 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            }
            return View("PJE0040PAGE1", model);

        }

        public ActionResult PJE0040PAGE2(PJE0040Model model)
        {



            return View("PJE0040PAGE1", model);
        }

        public ActionResult PJE0040PAGE1(PJE0040Model model)
        {



            return View("PJE0040PAGE1", model);
        }

        //維護
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Maintain")]
        public ActionResult Maintain(PJE0030Model model)
        {
            int page = (model.page == 0) ? 1 : model.page;
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJDFDCR p WHERE p.PID=@PID");
            //cmd.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.MainList = plist;
            return View("PJE0040Maintain", model);
        }

        //管制
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Manage")]
        public ActionResult Manage(PJE0030Model model)
        {
            //int page = (model.page == 0) ? 1 : model.page;

            SqlCommand cmd = new SqlCommand("select 0 as SEQ_NO, '原預算' as SEQ_NAME union select SEQ_NO , '第'+ convert(nvarchar(3),SEQ_NO) +'期修正' from PJDFDCR where PID = @PID");
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            ViewBag.sels = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);


            cmd = new SqlCommand("select * from PJV_CTDMNDS where PID = @PID");
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            ViewBag.List1 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);


            cmd = new SqlCommand("select * from PJV_CTDMNDS where PID = @PID");
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            ViewBag.List2 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            cmd = new SqlCommand("select * from PJV_CRDMNDS where PID = @PID");
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            ViewBag.List3 = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);

            return View("PJE0040Manage");
        }

        //預算統計
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BudgetCount")]
        public ActionResult BudgetCount()
        {



            return View("PJE0040BudgetCount");
        }

        //預算調整
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BudgetEdit")]
        public ActionResult BudgetEdit()
        {



            return View("PJE0040BudgetEdit");
        }


    }

}