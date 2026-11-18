using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WCals.Models;
using WCals.Service;
using WCals.Service.AllInterFace;

namespace WCals.Controllers
{
    

    public class AjaxDataController : Controller
    {
        private string connstringSt = "Data Source=10.1.1.112\\PMIS;Initial Catalog=WCals;User Id=***;Password=***";
        //service
        private PlanIUService planService = new PlanService();

        Models.WCalsEntities db = new Models.WCalsEntities();

        //取得供水範圍
        [Route("getSysnameRange")]
        [HttpGet]
        public string getSysnameRange()
        {
            string sysnameData = this.Request.QueryString["sysnameData"];
            var sysnameRange = (from p in db.DBO_WATERRANGE
                                where p.SYSID == sysnameData
                                select p.RANGE).ToList();
            var jj = "";

            if (sysnameRange != null)
            {
                foreach (var f in sysnameRange)
                {
                    jj = f;

                }
            }
            return jj;







        }
        //ajax取得鄉政名稱(撈取[SMV_SMCTOWN])
        [Route("getTownName")]
        [HttpGet]
        public string getTownName()
        {

            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet ds = new DataSet();


            APDFUNDObj jj = new APDFUNDObj();
            jj.saveData();

            //string connstringSt = "Data Source=10.1.1.45;Initial Catalog=WCals;User Id=***;Password=***";
            var queryString = "select * from PJMFORM where PJ_CODE = @pjcode";

            var conn = new SqlConnection(this.connstringSt);
            //var conntwo = new SqlConnection(ConfigurationManager.ConnectionStrings["Data Source=10.1.1.45;Initial Catalog=WCals;User Id=***;Password=***"].ConnectionString());
            conn.Open();
            SqlCommand sqlSt = new SqlCommand(queryString, conn);
            sqlSt.Parameters.Add(new SqlParameter("@pjcode", "1"));

            SqlDataReader reader = sqlSt.ExecuteReader();
            while (reader.Read())
            {
                var dd = reader["PJ_NAME"].ToString() + "," + reader["PJ_CITY"].ToString();
                Console.WriteLine(dd);



            }




            string townNameData = this.Request.QueryString["townNameData"];
            var townNameDataRange = (from p in db.SMV_SMCTOWN
                                     where p.CITY == townNameData
                                     select p.TOWN).ToList();
            string jsonStr = JsonConvert.SerializeObject(townNameDataRange);
            return jsonStr;
        }

        //批次動作處理ajax
        [Route("deltePualData")]
        [HttpGet]
        public string deltePualData(string test)
        {

            string param1 = this.Request.QueryString["sampleItems"];
            var ff = Convert.ToInt32(param1);
            var pjdbatc = from p in db.PJDBATC
                          where p.DID == ff
                          select p;
            foreach (var pj in pjdbatc)
            {
                db.PJDBATC.Remove(pj);
            }
            try
            {
                db.SaveChanges();

            }
            catch (Exception)
            {
                throw;
            }

            return "刪除成功";





        }

    }
}