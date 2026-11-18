using Aspose.Cells;
using Aspose.Words;
using Newtonsoft.Json;
using PagedList;
using Syncfusion.CompoundFile.DocIO.Native;
using Syncfusion.CompoundFile.XlsIO.Native;
using Syncfusion.DocIO.DLS;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.XlsIO.Implementation.PivotAnalysis;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Security.Cryptography;
using System.ServiceModel.Configuration;
using System.Text;
using System.Web.Mvc;
using System.Web.UI;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEPMSModels;
using WCals.Models.CMETMSModels;
using WCals.Models.CMPSGSModels;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS {


    public class SGEAPController : Controller {
        //private readonly Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();
        int pageSize = 10;
        public ActionResult Index(SGEModel model) {
            string userid = Helpers.WebSiteHelper.getUserId();
            int page = (model.page == 0) ? 1 : model.page;
            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List("SELECT * FROM PJMFORM m, PJDSYRC p WHERE m.PPID=p.PPID AND m.ADV_UID='"+ userid + "'");
            List<DataRow> list = dt.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.List1 = plist;
            return View("SGEAPIndex");

        }
    }

    
    
    public class SGE0025Controller : Controller {
        private readonly Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();
        int pageSize = 10;

        public ActionResult Index(SGEModel model) {
            string userid = Helpers.WebSiteHelper.getUserId();
            int page = (model.page == 0) ? 1 : model.page;
            model.page = page;
            if(model.PJMFORM == null) {
                model.PJMFORM = new Models.CMPSGSModels.PJMFORM_CMPSGS();
            }
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJMFORM p WHERE p.ADV_UID=@UID");
            cmd.Parameters.Add(new SqlParameter("UID", userid));
            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.MainList = plist;
            return View("SGE0025Index", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Search")]
        public ActionResult Search(SGEModel model) {

            string userid = Helpers.WebSiteHelper.getUserId();
            int page = (model.page == 0) ? 1 : model.page;
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJMFORM p WHERE p.ADV_UID=@UID AND PJ_NAME LIKE @PJ_NAME OR PPID LIKE @PPID");

            cmd.Parameters.Add(new SqlParameter("UID", userid));
            if(string.IsNullOrEmpty(model.keyword)) {
                cmd.Parameters.Add(new SqlParameter("PJ_NAME", "%"));
                cmd.Parameters.Add(new SqlParameter("PPID", "%"));
            }
            else {
                cmd.Parameters.Add(new SqlParameter("PJ_NAME", "%" + model.keyword + "%"));
                cmd.Parameters.Add(new SqlParameter("PPID", "%" + model.keyword + "%"));
            }


            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);

            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.MainList = plist;
            return View("SGE0025Index", model);

        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Review")]
        public ActionResult Review(SGEModel model) {

            return View("SGE0025Index", model);

        }

    }

    public class SGE0035Controller : Controller {
        private readonly Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();
        int pageSize = 10;

        public ActionResult Index(SGEModel model) {
            string userid = Helpers.WebSiteHelper.getUserId();
            int page = (model.page == 0) ? 1 : model.page;
            model.page = page;
            if(model.PJMFORM == null) {
                model.PJMFORM = new Models.CMPSGSModels.PJMFORM_CMPSGS();
            }
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJMFORM p WHERE p.ADV_UID=@UID");
            cmd.Parameters.Add(new SqlParameter("UID", userid));
            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.MainList = plist;
            return View("SGE0035Index", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Search")]
        public ActionResult Search(SGEModel model) {

            string userid = Helpers.WebSiteHelper.getUserId();
            int page = (model.page == 0) ? 1 : model.page;
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJMFORM p WHERE p.ADV_UID=@UID AND PJ_NAME LIKE @PJ_NAME OR PPID LIKE @PPID");

            cmd.Parameters.Add(new SqlParameter("UID", userid));
            if(string.IsNullOrEmpty(model.keyword)) {
                cmd.Parameters.Add(new SqlParameter("PJ_NAME", "%"));
                cmd.Parameters.Add(new SqlParameter("PPID", "%"));
            }
            else {
                cmd.Parameters.Add(new SqlParameter("PJ_NAME", "%" + model.keyword + "%"));
                cmd.Parameters.Add(new SqlParameter("PPID", "%" + model.keyword + "%"));
            }


            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);

            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.MainList = plist;
            return View("SGE0035Index", model);

        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Review")]
        public ActionResult Review(SGEModel model) {
            return View("SGE0035Index", model);
        }
    }
}