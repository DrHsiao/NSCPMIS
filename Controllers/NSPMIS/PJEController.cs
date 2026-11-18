using Aspose.Cells;
using Aspose.Words;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PagedList;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Parser.Biff_Records;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMEPMSModels;
using WCals.Models.CMETMSModels;
using WCals.Models.CMPSGSModels;
using WCals.ViewModels.NewModels;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.PJE {
    public class ASYNCController : Controller {
        private readonly Models.CMPSGSModels.CMPSGSEntities sgsdb = new Models.CMPSGSModels.CMPSGSEntities();
        public ActionResult DownloadAPDFILE(decimal aid, decimal fid) {
            APDFILE file = sgsdb.APDFILE.Where(w => w.FID == fid && w.AID == aid).FirstOrDefault();

            if(file == null) {
                return HttpNotFound();
            }

            var contentType = MimeMapping.GetMimeMapping(file.FILE_PATH);
            return File(file.FILE_BODY, contentType);
        }

        public ActionResult GetAPDPROC(decimal aid) {
            string sql = "SELECT a.*, u.USR_NAME FROM APDPROC a, CMBASE.dbo.OGCUSER u WHERE a.UPD_UID=u.USR_ID AND a.AID=" + aid + " ORDER BY a.DID DESC";
            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List(sql);
            //dt.Columns["UPD_DT"].DataType = typeof(string);
            for(int i = 0; i < dt.Rows.Count; i++) {
                dt.Rows[i]["PROC_ACT"] = util.GetPROC_NAME(Convert.ToString(dt.Rows[i]["PROC_ACT"]));
                dt.Rows[i]["OAP_STAT"] = util.GetSTAT_NAME(Convert.ToString(dt.Rows[i]["OAP_STAT"]));
                dt.Rows[i]["NAP_STAT"] = util.GetSTAT_NAME(Convert.ToString(dt.Rows[i]["NAP_STAT"]));
                //dt.Rows[i]["UPD_DT"] = Convert.ToDateTime(dt.Rows[i]["UPD_DT"]).ToString("yyyy/MM/dd");
            }
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 依單位取使用者 
        /// </summary>
        /// <param name="DEPT_NO"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetUsers(string DEPT_NO) {
            string sql = "SELECT u.USR_ID, u.USR_NAME FROM OGCUSER u, OGCDEPT d, OGCDPUR p WHERE d.DEPT_NO='" + DEPT_NO + "' AND p.USR_ID=u.USR_ID AND d.DEPT_NO=p.DEPT_NO";
            DataTable dt = BASE_DAO.SqlDao.CMBASEQuery4List(sql);
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 依縣市取分區11111111111
        /// </summary>
        /// <param name="city"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetTowns(string city) {
            DataTable dt = util.GetTowns(city);
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetPJCMFORMbyCity(string city) {
            DataTable dt = SqlDao.CMETMSQuery4List("SELECT PJ_NAME FROM PJCFORM WHERE SER_CITY='" + city + "'");
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }


        //新的依縣市取分區
        public JsonResult GetTownsById(string city) {
            DataTable dt = util.GetTownsById(city);
            if(dt.Rows.Count == 0) {
                dt = util.GetTownsByName(city);
            }
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetTownsOldandNew(string city)
        {
            DataTable dt = util.GetTownsOldandNew(city);
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }


        public JsonResult GetTownsByName(string city) {
            DataTable dt = util.GetTownsByName(city);
            if(dt.Rows.Count == 0) {
                dt = util.GetTownsById(city);
            }
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetApmformBy(string city, string town, string no, string name) {
            DataTable dt = util.GetTownsByName(city);
            if(dt.Rows.Count == 0) {
                dt = util.GetTownsById(city);
            }
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 依PID取KATY名稱
        /// </summary>
        /// <param name="PID"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetCATYS(string PID) {
            return null;
        }

        /// <summary>
        /// 取計劃名
        /// </summary>
        /// <param name="city"></param>
        /// <returns></returns>
        public JsonResult GetPJ_NAMES(string city) {
            DataTable dt = SqlDao.CMETMSQuery4List("SELECT PJ_NAME as name FROM PJCFORM WHERE PJ_ID_TYPE='C' and SER_CITY='" + city + "' ");
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetPJDSYRC(string PPID) {
            DataTable dt = SqlDao.CMPSGSQuery4List("Select YRC_ID, PSG_NAME from PJDSYRC Where PPID = " + PPID + " Order by SYR desc");
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }


    }

    
    
    
    
    
    
}