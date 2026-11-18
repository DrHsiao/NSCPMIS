using Syncfusion.Pdf.Tables;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Web.Mvc;
using Newtonsoft.Json;
using System.Web.Script.Serialization;
using WCals.Models.CMEPMSModels;
using Syncfusion.XlsIO.Parser.Biff_Records;
using WCals.Helpers;
using static WCals.classes.john.BASE_DAO;
using System.Linq.Expressions;
using WCals.Models.CMPSGSModels;
using System.Net.NetworkInformation;
using Syncfusion.DocIO.DLS;
using WCals.Models.CMETMSModels;
using TDMFORM = WCals.Models.CMETMSModels.TDMFORM;

namespace WCals.classes.john {

    //頁面多按鈕處理
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = true)]
    public class MultiButtonAttribute : ActionNameSelectorAttribute {
        public string MatchFormKey { get; set; }
        public string MatchFormValue { get; set; }

        public override bool IsValidName(ControllerContext controllerContext, string actionName, MethodInfo methodInfo) {
            return controllerContext.HttpContext.Request[MatchFormKey] != null &&
                controllerContext.HttpContext.Request[MatchFormKey] == MatchFormValue;
        }
    }

    public class code {


        /// <summary>
        /// AP_STAT前進
        /// </summary>
        /// <param name="stat"></param>
        /// <returns></returns>
        public static string AP_STAT_Forward(string stat) {
            switch(stat) {
                case "-1"://提案中↓
                    return "00";
                case "00"://待收件↓
                    return "10";
                case "10"://初審↓
                    return "15";

                case "15"://複審↓
                    return "20";
                case "20"://成案↓
                    return "90";
                //Z1未過初審
                //Z2未過複審
                default:
                    return "00";
            }
        }

        /// <summary>
        /// AP_STAT後退
        /// </summary>
        /// <param name="stat"></param>
        /// <returns></returns>
        public static string AP_STAT_Backward(string stat) {
            switch(stat) {
                case "-1"://提案中↓
                    return "-1";
                case "00"://待收件↓
                    return "-1";
                case "10"://初審↓
                    return "00";

                case "15"://複審↓
                    return "10";
                case "20"://成案↓
                    return "15";
                case "90"://成案↓
                    return "20";
                default:
                    return "00";
            }
        }
    }


    public class util {

        public static DataTable GetTownsOldandNew(string city) {
            StringBuilder sql = new StringBuilder("SELECT TOWN,CPA_TOWN_ID as TOWN_ID FROM SMCPOSTN " +
                "WHERE TOWN<>'' and CITY" +
                " UNION " +
                " SELECT TOWN,CPA_TOWN_ID as TOWN_ID FROM SMCPOST WHERE TOWN<>'' and CITY "
                );
            SqlCommand cmd = new SqlCommand();

            if(city != null) {
                sql = sql.Replace("and CITY", " AND CITY=@city");
                cmd.Parameters.Add(new SqlParameter("@city", city));
            }
            sql.Append(" ORDER BY TOWN, CPA_TOWN_ID");

            cmd.CommandText = sql.ToString();
            DataTable dt = BASE_DAO.SqlDao.CMBASEQuery4List(cmd);

            return dt;
        }
        public static string getACC_ID(decimal PID, string CATY_CODE) {
            try {
                SqlCommand cmd = new SqlCommand("select ACC_ID from PJDCATY where PID=@PID and CATY_CODE=@CATY_CODE");
                cmd.Parameters.Add(new SqlParameter("@PID", PID));
                cmd.Parameters.Add(new SqlParameter("@CATY_CODE", CATY_CODE));

                DataRow result = SqlDao.CMEPMSQuery4Object(cmd);
                return result != null ? result["ACC_ID"].ToString() : "";
            } catch(Exception e) {
                System.Diagnostics.Debug.WriteLine(e);
                return "";
            }

        }

        /// <summary>
        /// 1.1.3新增管制點
        /// </summary>
        /// <param name="c"></param>
        /// <param name="tdmform"></param>
        public static void addTDDSDWB(CTDMNDS c, TDMFORM tdmform) {
            try {
                string sSql = "insert into CMETMS.dbo.TDDSDWB(TD_CODE, IT_CODE, AN_IT_CODE, IT_NAME, IT_TYPE, GP_ORDER, IT_GP_NAME, IT_WGT,  WRO_FG, DEL_FG, EDIT_FG) " +
                           "select '" + c.TD_CODE + "', IT_CODE, AN_IT_CODE, IT_NAME, IT_TYPE, GP_ORDER, IT_GP_NAME, IT_WGT,  WRO_FG, DEL_FG, EDIT_FG " +
                           "from CMETMS.dbo.SMWSDTP where TD_TYPE = '" + tdmform.CATY_CODE + "' ";
                BASE_DAO.SqlDao.CMETMSExecuteSql(sSql);
                sSql = "insert into CMETMS.dbo.TDMADMC(TD_CODE, ADV_DP, ADV_UID, OPEN_FG) Values ('" + c.TD_CODE + "','" + WebSiteHelper.getUserDept() + "','" + WebSiteHelper.getUserId() + "',1)";
                BASE_DAO.SqlDao.CMETMSExecuteSql(sSql);

                if(tdmform.ADV_DP.Trim().Length > 0 && WebSiteHelper.getUserId() != tdmform.ADV_UID && WebSiteHelper.getUserDept() != tdmform.ADV_DP) {
                    sSql = "insert into CMETMS.dbo.TDMADMC(TD_CODE, ADV_DP, ADM_LEV, ADV_UID, OPEN_FG) Values ('" + c.TD_CODE + "','" + tdmform.ADV_DP + "','1','" + tdmform.ADV_UID + "',0)";
                    BASE_DAO.SqlDao.CMETMSExecuteSql(sSql);
                }
                sSql = "insert into CMETMS.dbo.TDDBGRC(TD_CODE, REC_TP, SEQ_NO, ADD_DT, UPD_DT, UPD_UID) values ('" + c.TD_CODE + "', 'B', 0, getdate(), getdate(), '" + WebSiteHelper.getUserId() + "')";
                BASE_DAO.SqlDao.CMETMSExecuteSql(sSql);
                sSql = "insert into CMETMS.dbo.TDDBGRC(TD_CODE, REC_TP, SEQ_NO, ADD_DT, UPD_DT, UPD_UID) values ('" + c.TD_CODE + "', 'C', 0, getdate(), getdate(), '" + WebSiteHelper.getUserId() + "')";
                BASE_DAO.SqlDao.CMETMSExecuteSql(sSql);
                sSql = "insert into CMETMS.dbo.TDDBGRC(TD_CODE, REC_TP, SEQ_NO, ADD_DT, UPD_DT, UPD_UID) values ('" + c.TD_CODE + "', 'S', 0, getdate(), getdate(), '" + WebSiteHelper.getUserId() + "')";
                BASE_DAO.SqlDao.CMETMSExecuteSql(sSql);
                sSql = "insert into CMETMS.dbo.TDDBGRC(TD_CODE, REC_TP, SEQ_NO, ADD_DT, UPD_DT, UPD_UID) values ('" + c.TD_CODE + "', 'F', 0, getdate(), getdate(), '" + WebSiteHelper.getUserId() + "')";
                BASE_DAO.SqlDao.CMETMSExecuteSql(sSql);

            } catch(Exception e) {
                System.Diagnostics.Debug.Write(e.Message);
            }
        }

        /// <summary>
        /// 修改管制點
        /// </summary>
        /// <param name="c"></param>
        /// <param name="tdmform"></param>
        public static void updateTDDSDWB(CTDMNDS c, string NEW_CATY_CODE) {
            SqlDao.CMETMSExecuteSql("DELETE FROM TDDSDWB WHERE TD_CODE='" + c.TD_CODE + "'");
            try {
                string newSeq = "insert into CMETMS.dbo.TDDSDWB(TD_CODE, IT_CODE, AN_IT_CODE, IT_NAME, IT_TYPE, GP_ORDER, IT_GP_NAME, IT_WGT, WGT_FG, WRO_FG, DEL_FG, EDIT_FG, ADD_DT, UPD_DT, UPD_UID)" +
                    "select @TD_CODE, IT_CODE, AN_IT_CODE, IT_NAME, IT_TYPE, GP_ORDER, IT_GP_NAME, IT_WGT, WGT_FG, WRO_FG, DEL_FG, EDIT_FG, getdate(), getdate(), @userid " +
                    "from CMETMS.dbo.SMWSDTP where TD_TYPE=@CATY_CODE";

                SqlCommand cmd = new SqlCommand(newSeq);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", c.TD_CODE));
                cmd.Parameters.Add(new SqlParameter("userid", WebSiteHelper.getUserId()));
                cmd.Parameters.Add(new SqlParameter("CATY_CODE", NEW_CATY_CODE));
                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

            } catch(Exception e) {
                System.Diagnostics.Debug.WriteLine(e);
            }
        }

        /// <summary>
        /// 新增管制點
        /// </summary>
        /// <param name="TD_CODE"></param>
        /// <param name="CATY_CODE"></param>
        public static void addTDDSDWB(string TD_CODE, string CATY_CODE) {
            try {
                string newSeq = "insert into CMETMS.dbo.TDDSDWB(TD_CODE, IT_CODE, AN_IT_CODE, IT_NAME, IT_TYPE, " +
                    "GP_ORDER, IT_GP_NAME, IT_WGT, WGT_FG, WRO_FG, DEL_FG, EDIT_FG, ADD_DT, UPD_DT, UPD_UID)" +
                    "select @TD_CODE, IT_CODE, AN_IT_CODE, IT_NAME, IT_TYPE, GP_ORDER, IT_GP_NAME, IT_WGT, " +
                    "WGT_FG, WRO_FG, DEL_FG, EDIT_FG, getdate(), getdate(), @userid " +
                    "from CMETMS.dbo.SMWSDTP where TD_TYPE=@CATY_CODE";
                SqlCommand cmd = new SqlCommand(newSeq);
                cmd.Parameters.Add(new SqlParameter("TD_CODE", TD_CODE));
                cmd.Parameters.Add(new SqlParameter("userid", WebSiteHelper.getUserId()));
                cmd.Parameters.Add(new SqlParameter("CATY_CODE", CATY_CODE));
                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
            } catch(Exception e) {
                System.Diagnostics.Debug.WriteLine(e);
            }

        }

        public static string GetPROC_NAME(string s) {
            switch(s) {
                case "NW":
                    return "申請";
                case "ST":
                    return "送件";
                case "RV":
                    return "收件";
                case "BK":
                    return "退件";
                default:
                    return "";
            }
        }

        public static string GetSTAT_NAME(string s) {
            switch(s) {
                case "-1":
                    return "新增提案";
                case "00":
                    return "待收件";
                case "10":
                    return "初審中";
                case "15": return "複審中";
                case "20":
                    return "核定中";
                case "90":
                    return "已成案";
                case "Z1":
                    return "未過初審";
                case "Z2":
                    return "未過複審";

                default:
                    return "00";
            }
        }

        public static string GetAP_STAT(int tab) {
            switch(tab) {
                case 0:
                    return "-1";
                case 1:
                    return "00";
                case 2:
                    return "10";
                case 3:
                    return "15";
                case 4:
                    return "20";
                case 5:
                    return "90";
                case 6:
                    return "Z1";
                case 7:
                    return "Z2";

                default:
                    return "00";
            }

        }

        /// <summary>
        /// Z2、ZZ 用
        /// </summary>
        /// <param name="PPID"></param>
        /// <param name="AP_STAT"></param>
        /// <param name="YRC"></param>
        /// <returns></returns>
        public static DataTable GetAPMFORMByDeptSP(object PPID, string AP_STAT, string YRC) {
            StringBuilder sb = new StringBuilder("SELECT * FROM APMFORM a WHERE AP_STAT = '" + AP_STAT + "' " +
                "AND a.PPID=" + PPID + " AND (select top 1 PRC_DP FROM APDPROC WHERE AID=a.AID and NAP_STAT='" + AP_STAT + "'" +
                " ORDER BY ADD_DT DESC)= '" + WebSiteHelper.getUserDept() + "'");
            if(!string.IsNullOrEmpty(YRC)) {
                sb.Append("AND a.YRC_ID='" + YRC + "'");
            }
            try {
                return SqlDao.CMPSGSQuery4List(sb.ToString());
            } catch {
                return null;
            }
        }

        /// <summary>
        /// 以部門取APMFORM
        /// 2.2.2 從初審開始全部鎖部門
        /// 2.2.3 從複審開始全部鎖部門
        /// Z1 等於10的邏輯
        /// Z2 等於15的邏輯
        /// ZZ 等於20的邏輯
        /// </summary>
        /// <param name="PPID"></param>
        /// <param name="stat"></param>
        /// <param name="YRC"></param>
        /// <returns></returns>
        public static DataTable GetAPMFORMByDeptSP(object PPID, string AP_STAT, string NAP_DTAT, string PROC_ACT, string YRC) {
            StringBuilder sb = new StringBuilder("SELECT * FROM APMFORM a WHERE AP_STAT = '" + AP_STAT + "' " +
                "AND a.PPID=" + PPID + " AND (select top 1 PRC_DP FROM APDPROC WHERE AID=a.AID and NAP_STAT='" + NAP_DTAT + "'" +
                "AND PROC_ACT ='" + PROC_ACT + "' ORDER BY ADD_DT DESC)= '" + WebSiteHelper.getUserDept() + "'");
            if(!string.IsNullOrEmpty(YRC)) {
                sb.Append("AND a.YRC_ID='" + YRC + "'");
            }
            try {
                return SqlDao.CMPSGSQuery4List(sb.ToString());
            } catch {
                return null;
            }
        }

        /// <summary>
        /// 20 用
        /// </summary>
        /// <param name="PPID"></param>
        /// <param name="AP_STAT"></param>
        /// <param name="PROC_ACT"></param>
        /// <param name="YRC"></param>
        /// <returns></returns>
        public static DataTable GetAPMFORMByDeptSP(object PPID, string AP_STAT, string PROC_ACT, string YRC) {
            StringBuilder sb = new StringBuilder("SELECT * FROM APMFORM a WHERE AP_STAT = '" + AP_STAT + "' " +
                "AND a.PPID=" + PPID + " AND (select top 1 PRC_DP FROM APDPROC WHERE AID=a.AID and NAP_STAT='20'" +
                "AND PROC_ACT ='" + PROC_ACT + "' ORDER BY ADD_DT DESC)= '" + WebSiteHelper.getUserDept() + "'");
            if(!string.IsNullOrEmpty(YRC)) {
                sb.Append("AND a.YRC_ID='" + YRC + "'");
            }
            try {
                return SqlDao.CMPSGSQuery4List(sb.ToString());
            } catch {
                return null;
            }
        }

        public static DataTable GetAPMFORMByEzway(object PPID, string stat, string YRC) {
            StringBuilder sb = new StringBuilder("SELECT(SELECT SUM(FUND)FROM APDFUND WHERE AID=a.AID AND FIT_CODE='00')as cen1," +
                "(SELECT SUM(FUND)FROM APDFUND WHERE AID=a.AID AND FIT_CODE='11')as loc1," +
                "(SELECT SUM(JUD_FUND)FROM APDFUND WHERE AID=a.AID AND FIT_CODE='00')as cen2," +
                "(SELECT SUM(JUD_FUND)FROM APDFUND WHERE AID=a.AID AND FIT_CODE='11')as loc2," +
                "a.PJ_CATY_CODE, a.AP_STAT, a.AID, a.AP_CODE, a.AP_PJ_NAME,a.AP_NAME,a.SYR,a.CITY, " +
                "a.TOWN FROM APMFORM a WHERE a.PPID='" + PPID + "'AND a.AP_STAT='" + stat + "'");
            if(!string.IsNullOrEmpty(YRC)) {
                sb.Append("AND a.YRC_ID='" + YRC + "'");
            }
            try {
                return SqlDao.CMPSGSQuery4List(sb.ToString());
            } catch {
                return null;
            }
        }

        /// <summary>
        /// 以人員取APMFORM
        /// 2.2.1
        /// </summary>
        /// <param name="PPID"></param>
        /// <param name="stat"></param>
        /// <param name="YRC"></param>
        /// <returns></returns>
        public static DataTable GetAPMFORM(object PPID, string stat, string YRC) {
            StringBuilder sb = new StringBuilder("SELECT a.AP_STAT,a.AID, a.AP_CODE, a.AP_PJ_NAME,a.AP_NAME,a.SYR,a.CITY, " +
                "a.TOWN FROM APMFORM a WHERE a.PPID='" + PPID + "'AND a.AP_STAT='" + stat +
                //"'AND ADV_DP='"+WebSiteHelper.getUserDept()+"'AND ADV_UID='"+WebSiteHelper.getUserId()+"'");
                "'AND ADV_UID='" + WebSiteHelper.getUserId() + "'");
            if(!string.IsNullOrEmpty(YRC)) {
                sb.Append("AND a.YRC_ID='" + YRC + "'");
            }
            try {
                return SqlDao.CMPSGSQuery4List(sb.ToString());
            } catch {
                return null;
            }
        }

        /// <summary>
        /// 以部門取APMFORM
        /// 2.1.1, 
        /// 2.2.2, 
        /// 2.2.3
        /// </summary>
        /// <param name="PPID"></param>
        /// <param name="stat"></param>
        /// <param name="YRC"></param>
        /// <returns></returns>
        public static DataTable GetAPMFORMByDept(object PPID, string stat, string YRC) {
            StringBuilder sb = new StringBuilder("SELECT a.AP_STAT,a.AID, a.AP_CODE, a.AP_PJ_NAME,a.AP_NAME,a.SYR,a.CITY, " +
                "a.TOWN FROM APMFORM a WHERE a.PPID='" + PPID + "'AND a.AP_STAT='" + stat +
                "'AND ADV_DP='" + WebSiteHelper.getUserDept() + "'");
            if(!string.IsNullOrEmpty(YRC)) {
                sb.Append("AND a.YRC_ID='" + YRC + "'");
            }
            try {
                return SqlDao.CMPSGSQuery4List(sb.ToString());
            } catch {
                return null;
            }
        }

        public static DataTable GetAPMFORM_Attach(object PPID, string stat, string userid, string dept) {
            try {
                //string sql = "SELECT a.* FROM APDPROC p, APMFORM a WHERE a.PPID='" + PPID + "'AND a.AID=p.AID AND a.AP_STAT='" + stat + "' AND p.PRC_DP='" + dept + "' AND a.ADV_UID<>'" + userid + "' ORDER BY p.ADD_DT DESC";
                string sql = "SELECT a.AID, COUNT(*) AS RecordCount, a.AP_CODE, a.AP_PJ_NAME,a.AP_NAME,a.SYR,a.CITY, a.TOWN " +
                "FROM APMFORM a INNER JOIN APDPROC p ON a.AID = p.AID " +
                "WHERE a.PPID='" + PPID + "'AND a.AP_STAT = '" + stat + "'AND p.PRC_DP = '" + dept + "'AND a.ADV_UID <> '" + userid + "'" +
                "GROUP BY a.AID,a.AP_CODE,a.AP_PJ_NAME,a.AP_NAME,a.SYR,a.CITY, a.TOWN ORDER BY a.AID;";
                return SqlDao.CMPSGSQuery4List(sql);
            } catch {
                return null;
            }
        }

        /// <summary>
        /// 取城市列表
        /// </summary>
        /// <returns></returns>
        public static DataTable GetCities() {
            return BASE_DAO.SqlDao.CMBASEQuery4List("SELECT s.CITY, MAX(s.CITY_ID) AS CITY_ID FROM SMCPOST s GROUP BY s.CITY");
        }

        /// <summary>
        /// 依使用者部門取得授權縣市 (3.1.7 需求：對應 CMBASE.dbo.SECCITY)
        /// </summary>
        public static DataTable GetCitiesByDept(string dept) {
            SqlCommand cmd = new SqlCommand("SELECT CITY FROM SECCITY WHERE DEPT_NO=@dept GROUP BY CITY");
            cmd.Parameters.Add(new SqlParameter("@dept", dept ?? ""));
            return BASE_DAO.SqlDao.CMBASEQuery4List(cmd);
        }

        /// <summary>
        /// 取城市列表 - 新
        /// </summary>
        /// <returns></returns>
        public static DataTable GetCitiesNew() {
            return BASE_DAO.SqlDao.CMBASEQuery4List("SELECT s.CITY, MAX(s.CPA_CITY_ID) AS CITY_ID FROM SMCPOSTN s GROUP BY s.CITY");
        }


        public static string GetTownName(string city, string town) {
            try {
                SqlCommand cmd = new SqlCommand("SELECT TOWN FROM SMCPOSTN WHERE CPA_CITY_ID=@city AND CPA_TOWN_ID=@town");
                cmd.Parameters.Add(new SqlParameter("@city", city));
                cmd.Parameters.Add(new SqlParameter("@town", town));

                DataRow result = BASE_DAO.SqlDao.CMBASEQuery4Object(cmd);
                return result != null ? result["TOWN"].ToString() : "";
            } catch(Exception e) {
                System.Diagnostics.Debug.WriteLine(e);
                return "";
            }
        }

        public static string GetCityName(string city) {
            try {
                SqlCommand cmd = new SqlCommand("SELECT s.CITY FROM SMCPOSTN s WHERE s.CPA_CITY_ID=@city GROUP BY s.CITY");
                cmd.Parameters.Add(new SqlParameter("@city", city));

                DataRow result = BASE_DAO.SqlDao.CMBASEQuery4Object(cmd);
                return result != null ? result["CITY"].ToString() : "";
            } catch(Exception e) {
                System.Diagnostics.Debug.WriteLine(e);
                return "";
            }
        }


        /// <summary>
        /// 取鄉鎮列表
        /// </summary>
        /// <param name="city"></param>
        /// <returns></returns>
        public static DataTable GetTowns(string city) {
            if(string.IsNullOrEmpty(city)) {
                city = null;
            }
            DataTable dt;
            StringBuilder sql = new StringBuilder("SELECT * FROM SMCPOSTN WHERE TOWN<>''");
            SqlCommand cmd = new SqlCommand();

            if(city != null) {
                sql.Append(" AND CITY=@city");
                cmd.Parameters.Add(new SqlParameter("@city", city));
            }

            cmd.CommandText = sql.ToString();
            dt = BASE_DAO.SqlDao.CMBASEQuery4List(cmd);

            if(dt.Rows.Count == 0) {
                sql = new StringBuilder("SELECT * FROM SMCPOSTN WHERE TOWN<>''");
                cmd = new SqlCommand();

                if(city != null) {
                    sql.Append(" AND CITY_ID=@city");
                    cmd.Parameters.Add(new SqlParameter("@city", city));
                }

                cmd.CommandText = sql.ToString();
                dt = BASE_DAO.SqlDao.CMBASEQuery4List(cmd);
            }

            return dt;
        }

        public static DataTable GetTownNew(string city) {
            return GetTownsById(city);

        }


        /// <summary>
        /// 取鄉鎮列表 - 新
        /// </summary>
        /// <param name="city"></param>
        /// <returns></returns>
        public static DataTable GetTownsById(string city) {
            StringBuilder sql = new StringBuilder("SELECT TOWN,CPA_TOWN_ID as TOWN_ID FROM SMCPOSTN WHERE TOWN<>''");
            SqlCommand cmd = new SqlCommand();

            if(city != null) {
                sql.Append(" AND CITY_ID=@city");
                cmd.Parameters.Add(new SqlParameter("@city", city));
            }
            sql.Append(" ORDER BY CITY, CPA_TOWN_ID");

            cmd.CommandText = sql.ToString();
            DataTable dt = BASE_DAO.SqlDao.CMBASEQuery4List(cmd);

            if(dt.Rows.Count == 0) {
                sql = new StringBuilder("SELECT TOWN,CPA_TOWN_ID as TOWN_ID FROM SMCPOSTN WHERE TOWN<>''");
                cmd = new SqlCommand();

                if(city != null) {
                    sql.Append(" AND CITY=@city");
                    cmd.Parameters.Add(new SqlParameter("@city", city));
                }
                sql.Append(" ORDER BY CITY, CPA_TOWN_ID");

                cmd.CommandText = sql.ToString();
                dt = BASE_DAO.SqlDao.CMBASEQuery4List(cmd);
            }

            return dt;
        }

        public static DataTable GetTownsByName(string city) {
            StringBuilder sql = new StringBuilder("SELECT TOWN,CPA_TOWN_ID as TOWN_ID FROM SMCPOSTN WHERE TOWN<>''");
            SqlCommand cmd = new SqlCommand();

            if(city != null) {
                sql.Append(" AND CITY=@city");
                cmd.Parameters.Add(new SqlParameter("@city", city));
            }
            sql.Append(" ORDER BY CITY, CPA_TOWN_ID");

            cmd.CommandText = sql.ToString();
            return BASE_DAO.SqlDao.CMBASEQuery4List(cmd);
        }

        /// <summary>
        /// 取CMBASE的SMCCODE
        /// </summary>
        /// <param name="SYS_ID"></param>
        /// <param name="CATY_ID"></param>
        /// <returns></returns>
        public static DataTable GetCMBASECODES(string SYS_ID, string CATY_ID) {
            SqlCommand cmd = new SqlCommand("SELECT c.CATY_COD as id, c.CODE_DESC as name FROM SMCCODE c WHERE c.SYS_ID=@SYS_ID AND c.CATY_ID=@CATY_ID ORDER BY c.CATY_COD");
            cmd.Parameters.Add(new SqlParameter("@SYS_ID", SYS_ID));
            cmd.Parameters.Add(new SqlParameter("@CATY_ID", CATY_ID));
            return BASE_DAO.SqlDao.CMBASEQuery4List(cmd);
        }

        /// <summary>
        /// 取CMEPMS的SMCCODE
        /// </summary>
        /// <param name="SYS_ID"></param>
        /// <param name="CATY_ID"></param>
        /// <returns></returns>
        public static DataTable GetCMEPMSCODES(string SYS_ID, string CATY_ID) {
            SqlCommand cmd = new SqlCommand("SELECT c.CATY_COD as id, c.CODE_DESC as name FROM SMCCODE c WHERE c.SYS_ID=@SYS_ID AND c.CATY_ID=@CATY_ID ORDER BY c.CATY_COD");
            cmd.Parameters.Add(new SqlParameter("@SYS_ID", SYS_ID));
            cmd.Parameters.Add(new SqlParameter("@CATY_ID", CATY_ID));
            return BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
        }

        /// <summary>
        /// 取CMETMS的SMCCODE
        /// </summary>
        /// <param name="SYS_ID"></param>
        /// <param name="CATY_ID"></param>
        /// <returns></returns>
        public static DataTable GetCMETMSCODES(string SYS_ID, string CATY_ID) {
            SqlCommand cmd = new SqlCommand("SELECT c.CATY_COD as id, c.CODE_DESC as name FROM SMCCODE c WHERE c.SYS_ID=@SYS_ID AND c.CATY_ID=@CATY_ID ORDER BY c.CATY_COD");
            cmd.Parameters.Add(new SqlParameter("@SYS_ID", SYS_ID));
            cmd.Parameters.Add(new SqlParameter("@CATY_ID", CATY_ID));
            return BASE_DAO.SqlDao.CMETMSQuery4List(cmd);
        }

        public static string GetIPJCODE(string ipjcode) {
            try {
                SqlCommand cmd = new SqlCommand("SELECT c.CODE_DESC FROM SMCCODE c WHERE c.SYS_ID='FD' AND c.CATY_ID='05' AND CATY_COD=@ipjcode");
                cmd.Parameters.Add(new SqlParameter("@ipjcode", ipjcode));

                DataRow result = BASE_DAO.SqlDao.CMPSGSQuery4Object(cmd);
                return result != null ? result["CODE_DESC"].ToString() : ipjcode;
            } catch {
                return ipjcode;
            }
        }

        public static DataTable GetCMPSGSCODES(string SYS_ID, string CATY_ID) {
            SqlCommand cmd = new SqlCommand("SELECT c.CATY_COD as id, c.CODE_DESC as name FROM SMCCODE c WHERE c.SYS_ID=@SYS_ID AND c.CATY_ID=@CATY_ID ORDER BY c.CATY_COD");
            cmd.Parameters.Add(new SqlParameter("@SYS_ID", SYS_ID));
            cmd.Parameters.Add(new SqlParameter("@CATY_ID", CATY_ID));
            return BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
        }

        /// <summary>
        /// 取最新版本期數:-1為無期數
        /// </summary>
        /// <param name="PID"></param>
        /// <returns></returns>
        public static int GetSEQ_NO(object PID) {
            try {
                SqlCommand cmd = new SqlCommand("SELECT SEQ_NO FROM PJDFDCR WHERE PID=@PID ORDER BY SEQ_NO DESC");
                cmd.Parameters.Add(new SqlParameter("@PID", PID));

                DataRow result = BASE_DAO.SqlDao.CMEPMSQuery4Object(cmd);
                return result != null ? Convert.ToInt32(result["SEQ_NO"]) : -1;
            } catch {
                return -1;
            }
        }

        /// <summary>
        /// 取縣市權
        /// </summary>
        public static DataTable GetMyCities(string userid) {
            SqlCommand cmd = new SqlCommand("SELECT s.CITY FROM SECCITY s WHERE s.DEPT_NO=@userid AND (s.TOWN='不分區' OR s.TOWN='') GROUP BY s.CITY");
            cmd.Parameters.Add(new SqlParameter("@userid", userid));
            return BASE_DAO.SqlDao.CMBASEQuery4List(cmd);
        }

        /// <summary>
        /// 取鄉鎮權
        /// </summary>
        public static DataTable GetMyTowns(string userid) {
            SqlCommand cmd = new SqlCommand("SELECT s.TOWN FROM SECCITY s WHERE s.DEPT_NO=@userid AND s.TOWN<>'' GROUP BY s.TOWN");
            cmd.Parameters.Add(new SqlParameter("@userid", userid));
            return BASE_DAO.SqlDao.CMBASEQuery4List(cmd);
        }

        /// <summary>
        /// 是否全域授權
        /// </summary>
        public static Boolean IsPrincipal(string userid) {
            SqlCommand cmd = new SqlCommand("SELECT s.CITY FROM SECCITY s WHERE s.DEPT_NO=@userid AND s.CITY='不分區'");
            cmd.Parameters.Add(new SqlParameter("@userid", userid));
            DataTable myCities = BASE_DAO.SqlDao.CMBASEQuery4List(cmd);
            return myCities.Rows.Count > 0;
        }


        /// <summary>
        /// 處理動作PROC_ACT(NW:新建;ST:送件;RV:收件;BK:退件)
        /// </summary>
        /// <param name="aid"></param>
        /// <param name="o_status"></param>
        /// <param name="AP_STAT"></param>
        /// <param name="PROC_ACT"></param>
        public static void UPDATE_APMFORM_Log(decimal aid, string o_status, string AP_STAT, string PROC_ACT, DateTime now, string userid, string userdp, string note) {

            string dtime = now.ToString("yyyy-MM-dd HH:mm:ss");
            BASE_DAO.SqlDao.CMPSGSExecuteSql(
            "INSERT INTO APDPROC(AID, OAP_STAT, NAP_STAT, PROC_ACT, ADD_DT, UPD_DT, UPD_UID, PRC_DP, PRC_NOTE)" +
            "VALUES(" + aid + ", '" + o_status + "', '" + AP_STAT + "', '" + PROC_ACT + "', '" + dtime + "', '" + dtime + "', '" + userid + "', '" + userdp + "', '" + note + "');");
            /*
            APDPROC r = new APDPROC();
            r.AID = aid;
            r.OAP_STAT = o_status;
            r.NAP_STAT = AP_STAT;
            
            r.PROC_ACT = "ST";
            r.ADD_DT = now;
            r.UPD_DT = now;
            r.UPD_UID = userid;
            r.PRC_DP = WebSiteHelper.getUserDept();

            CMPSGSEntities cmpsgs = new Models.CMPSGSModels.CMPSGSEntities();
            cmpsgs.APDPROC.Add(r);
            cmpsgs.SaveChanges();
            */
        }
    }

    public class tools {

        public static DataTable ResetNull(DataTable dt) {
            foreach(DataRow row in dt.Rows) {
                for(int i = 0; i < dt.Columns.Count; i++) {
                    if(row[i] == null) {
                        row[i] = "";
                    }
                }
            }
            return dt;
        }


        /// <summary>
        /// 民國7碼轉西元年
        /// </summary>
        /// <param name="yyyMMdd"></param>
        /// <returns></returns>
        public static DateTime GetDateByTWDate(string yyyMMdd) {
            
            yyyMMdd = "0" + yyyMMdd;
            DateTime csharpDate = DateTime.ParseExact(yyyMMdd, "yyyyMMdd", null);

            csharpDate = DateTime.ParseExact(yyyMMdd, "yyyyMMdd", null);
            csharpDate = csharpDate.AddYears(1911);


            return csharpDate;
        }

        /// <summary>
        /// 西元年轉民國7碼
        /// </summary>
        /// <param name="yyyMMdd"></param>
        /// <returns></returns>
        public static string GetDateByTWDate(DateTime date) {

            date = date.AddYears(-1911); // 重新賦值
            return date.ToString("yyyMMdd");
        }

        public static decimal safe2Decimal(object o) {

            try {
                return Convert.ToDecimal(o);

            } catch {
                return 0;
            }


        }

        /// <summary>
        /// 取狀態對應中文
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string GetStat(string s) {
            string target = "";
            switch(s) {
                case "00": target = "提案"; break;
                case "10": target = "初審"; break;
                case "15": target = "複審"; break;
                case "20": target = "核定"; break;
                case "90": target = "已成案"; break;
                case "Z1": target = "未通過初審"; break;
                case "Z2": target = "未通過複審"; break;
                default: target = "提案"; break;
            }
            return target;
        }

        public static SelectList ToSelectList(DataTable table, string valueField, string textField, string first) {

            List<SelectListItem> list = new List<SelectListItem>();
            if(!string.IsNullOrEmpty(first)) {
                list.Add(new SelectListItem() {
                    Text = first,
                    Value = ""
                });
            }
            foreach(DataRow row in table.Rows) {
                list.Add(new SelectListItem() {
                    Text = row[textField].ToString(),
                    Value = row[valueField].ToString()
                });
            }

            return new SelectList(list, "Value", "Text");
        }

        public static SelectList ToSelectListNotItem(DataTable table, string valueField, string textField, string first) {

            List<SelectListItem> list = new List<SelectListItem>();

            foreach(DataRow row in table.Rows) {
                list.Add(new SelectListItem() {
                    Text = row[textField].ToString(),
                    Value = row[valueField].ToString()
                });
            }

            return new SelectList(list, valueField, textField);
        }

        public static decimal safe2Decimal(string s) {
            try {
                return Convert.ToDecimal(s);
            } catch {
                return 0;
            }
        }

        //處理注入攻擊
        private static string sqlFilterImp(object sql) {
            string str = sql.ToString();
            str = str.Replace("'", "");
            str = str.Replace("\"", "");
            str = str.Replace(";", "");
            str = str.Replace("\\", "");

            str = str.ToLower();
            //非法字符
            String[] keywords = { "master", "truncate", "insert", "select", "delete", "update", "declare", "alter", "drop" };
            foreach(string k in keywords) {
                if(str.IndexOf(k) != -1) {
                    return "";
                }
            }
            str = null;
            return sql.ToString();
        }

        /// <summary>
        /// 簡易型sql過濾器
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static String SqlFilter(object sql) {

            return sqlFilterImp(sql);
        }

        public static int SqlFilterInt(object sql) {

            try {
                return Convert.ToInt32(sqlFilterImp(sql));
            } catch {
                return 0;
            }

        }

        /// <summary>讀取 Json 檔，回傳 DataTable</summary>
        public static DataTable JsonFileReadToDataTable(string readFilePath) {
            DataTable dtData = null;

            if(File.Exists(readFilePath) == true) {
                bool readType = false;

                try {
                    using(StreamReader sr = new StreamReader(readFilePath, Encoding.Unicode)) {
                        dtData = JsonConvert.DeserializeObject<DataTable>(sr.ReadToEnd());

                        sr.Close();
                        sr.Dispose();

                        readType = true;
                    }
                } catch {
                    try {
                        using(StreamReader sr = new StreamReader(readFilePath, Encoding.UTF8)) {
                            dtData = JsonConvert.DeserializeObject<DataTable>(sr.ReadToEnd());

                            sr.Close();
                            sr.Dispose();

                            readType = true;
                        }
                    } catch {
                        readType = false;
                    }
                }
            }

            return dtData;
        }

        /// <summary>
        /// DataRow轉json序列
        /// </summary>
        /// <param name="datarow"></param>
        /// <returns></returns>
        public static string Convert_DataRowToJson(DataRow datarow) {
            var dict = new Dictionary<string, object>();
            foreach(DataColumn col in datarow.Table.Columns) {
                dict.Add(col.ColumnName, datarow[col]);
            }

            var jsSerializer = new JavaScriptSerializer();
            //TODO取出問題
            return jsSerializer.Serialize(dict);
        }

        /// <summary>
        /// DataTable轉json序列
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static string ConvertDataTabletoString(DataTable table) {
            string JSONString = string.Empty;
            JSONString = JsonConvert.SerializeObject(table);
            return JSONString;
        }

        /// <summary>
        /// 依登入者所屬縣市建立縣市下拉選單。
        /// 縣市政府帳號只呈現該縣市（五都另含改制前舊縣名）；中央帳號則呈現全部縣市。
        /// </summary>
        public static SelectList BuildCitySelectList(string loginCity) {
            if (!string.IsNullOrEmpty(loginCity)) {
                List<SelectListItem> singleCity = new List<SelectListItem>();

                if (loginCity == "新北市") {
                    singleCity.Add(new SelectListItem { Text = "新北市", Value = "新北市" });
                    singleCity.Add(new SelectListItem { Text = "台北縣", Value = "台北縣" });
                } else if (loginCity == "台中市") {
                    singleCity.Add(new SelectListItem { Text = "台中市", Value = "台中市" });
                    singleCity.Add(new SelectListItem { Text = "台中縣", Value = "台中縣" });
                } else if (loginCity == "台南市") {
                    singleCity.Add(new SelectListItem { Text = "台南市", Value = "台南市" });
                    singleCity.Add(new SelectListItem { Text = "台南縣", Value = "台南縣" });
                } else if (loginCity == "高雄市") {
                    singleCity.Add(new SelectListItem { Text = "高雄市", Value = "高雄市" });
                    singleCity.Add(new SelectListItem { Text = "高雄縣", Value = "高雄縣" });
                } else if (loginCity == "桃園市") {
                    singleCity.Add(new SelectListItem { Text = "桃園市", Value = "桃園市" });
                    singleCity.Add(new SelectListItem { Text = "桃園縣", Value = "桃園縣" });
                } else {
                    // 非五都之常規縣市，僅呈現登入縣市
                    singleCity.Add(new SelectListItem { Text = loginCity, Value = loginCity });
                }

                return new SelectList(singleCity, "Value", "Text", loginCity);
            } else {
                // 中央帳號，顯示所有縣市
                return tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);
            }
        }

    }

    //應急資料處理
    internal class BASE_DAO {

        public class SqlDao {
            private static readonly Models.CMBASEModels.CMBASEEntities cmbase = new Models.CMBASEModels.CMBASEEntities();
            private static readonly Models.CMEPMSModels.CMEPMSEntities cmepms = new Models.CMEPMSModels.CMEPMSEntities();
            private static readonly Models.CMPSGSModels.CMPSGSEntities cmpsgs = new Models.CMPSGSModels.CMPSGSEntities();
            private static readonly Models.CMETMSModels.CMETMSEntities cmetms = new Models.CMETMSModels.CMETMSEntities();
            private static readonly Models.CMEDOCModels.CMEDOCEntities cmedoc = new Models.CMEDOCModels.CMEDOCEntities();

            static string cmbaseConstr = cmbase.Database.Connection.ConnectionString;
            static string cmepmsConstr = cmepms.Database.Connection.ConnectionString;
            static string cmpsgsConstr = cmpsgs.Database.Connection.ConnectionString;//cmpsgs
            static string cmetmsConstr = cmetms.Database.Connection.ConnectionString;//cmetms
            static string cmedocConstr = cmedoc.Database.Connection.ConnectionString;//cmedoc

            //CMEDOC
            public static DataTable CMEDOCQuery4List(String sql) {
                SqlConnection mycon = new SqlConnection(cmedocConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                return dt;
            }
            public static DataTable CMEDOCQuery4List(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmedocConstr);
                mycon.Open();
                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                return dt;
            }

            public static void CMEDOCExecuteSql(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmedocConstr);
                mycom.Connection = mycon;
                mycon.Open();
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch {
                    mycon.Close();
                }
            }

            public static void CMEDOCExecuteSql(string sql) {
                SqlConnection mycon = new SqlConnection(cmedocConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch {
                    mycon.Close();
                }
            }

            public static DataRow CMEDOCQuery4Object(String sql) {
                SqlConnection mycon = new SqlConnection(cmedocConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }

            public static DataRow CMEDOCQuery4Object(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmedocConstr);
                mycon.Open();
                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }


            //CMBASE
            public static DataTable CMBASEQuery4List(String sql) {
                SqlConnection mycon = new SqlConnection(cmbaseConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);

                mydr.Close();
                mycon.Close();
                return dt;
            }
            public static DataTable CMBASEQuery4List(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmbaseConstr);
                mycon.Open();
                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                return dt;
            }
            public static DataRow CMBASEQuery4Object(String sql) {
                SqlConnection mycon = new SqlConnection(cmbaseConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }

            public static DataRow CMBASEQuery4Object(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmbaseConstr);
                mycon.Open();
                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }
            public static void CMBASEExecuteSql(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmbaseConstr);
                mycom.Connection = mycon;
                mycon.Open();
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch {
                    mycon.Close();
                }
            }

            public static void CMBASEExecuteSql(string sql) {
                SqlConnection mycon = new SqlConnection(cmbaseConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch {
                    mycon.Close();
                }
                //mydr.Close();        
            }

            //CMETMS
            public static DataTable CMETMSQuery4List(String sql) {
                //EShopEntities db = new EShopEntities();
                //string constr = ConfigurationManager.ConnectionStrings["EShopEntities"].ConnectionString;        
                SqlConnection mycon = new SqlConnection(cmetmsConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);

                mydr.Close();
                mycon.Close();
                return dt;
            }
            public static DataTable CMETMSQuery4List(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmetmsConstr);
                mycon.Open();
                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                return dt;
            }
            public static DataRow CMETMSQuery4Object(String sql) {
                //EShopEntities db = new EShopEntities();
                //string constr = ConfigurationManager.ConnectionStrings["EShopEntities"].ConnectionString;        
                SqlConnection mycon = new SqlConnection(cmetmsConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }

            public static DataRow CMETMSQuery4Object(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmetmsConstr);
                mycon.Open();

                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);

                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }
            public static void CMETMSExecuteSql(SqlCommand mycom) {

                SqlConnection mycon = new SqlConnection(cmetmsConstr);
                mycom.Connection = mycon;
                mycon.Open();
                //SqlCommand mycom = new SqlCommand(sql, mycon);
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch(Exception e) {
                    System.Diagnostics.Debug.WriteLine(e.StackTrace);
                    mycon.Close();
                    throw; // 重新拋出異常讓呼叫者知道執行失敗
                }
            }
            public static void CMETMSExecuteSql(string sql) {
                SqlConnection mycon = new SqlConnection(cmetmsConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch {
                    mycon.Close();
                }
                //mydr.Close();        
            }

            //CMEPMS
            public static DataTable CMEPMSQuery4List(String sql) {
                //EShopEntities db = new EShopEntities();
                //string constr = ConfigurationManager.ConnectionStrings["EShopEntities"].ConnectionString;        
                SqlConnection mycon = new SqlConnection(cmepmsConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);

                mydr.Close();
                mycon.Close();
                return dt;
            }
            public static DataTable CMEPMSQuery4List(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmepmsConstr);
                mycon.Open();
                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                return dt;
            }
            public static DataRow CMEPMSQuery4Object(String sql) {
                //EShopEntities db = new EShopEntities();
                //string constr = ConfigurationManager.ConnectionStrings["EShopEntities"].ConnectionString;        
                SqlConnection mycon = new SqlConnection(cmepmsConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }

            public static DataRow CMEPMSQuery4Object(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmepmsConstr);
                mycon.Open();

                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);

                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }
            public static void CMEPMSExecuteSql(SqlCommand mycom) {

                SqlConnection mycon = new SqlConnection(cmepmsConstr);
                mycom.Connection = mycon;
                mycon.Open();
                //SqlCommand mycom = new SqlCommand(sql, mycon);
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch(Exception e) {
                    System.Diagnostics.Debug.WriteLine(e.StackTrace);
                    mycon.Close();
                }
            }
            public static void CMEPMSExecuteSql(string sql) {
                SqlConnection mycon = new SqlConnection(cmepmsConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch {
                    mycon.Close();
                }
                //mydr.Close();        
            }

            //CMPSGS
            public static DataTable CMPSGSQuery4List(String sql) {
                //EShopEntities db = new EShopEntities();
                //string constr = ConfigurationManager.ConnectionStrings["EShopEntities"].ConnectionString;        
                SqlConnection mycon = new SqlConnection(cmpsgsConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);

                mydr.Close();
                mycon.Close();
                return dt;
            }

            public static DataTable CMPSGSQuery4List(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmpsgsConstr);
                mycon.Open();
                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                return dt;
            }
            public static DataRow CMPSGSQuery4Object(String sql) {
                //EShopEntities db = new EShopEntities();
                //string constr = ConfigurationManager.ConnectionStrings["EShopEntities"].ConnectionString;  
                SqlConnection mycon = new SqlConnection(cmpsgsConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }

            public static DataRow CMPSGSQuery4Object(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmpsgsConstr);
                mycon.Open();
                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);
                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }

            public static void CMPSGSExecuteSql(string sql) {
                SqlConnection mycon = new SqlConnection(cmpsgsConstr);
                mycon.Open();
                SqlCommand mycom = new SqlCommand(sql, mycon);
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch {
                    mycon.Close();
                }
                //mydr.Close();        
            }
            public static void CMESGSExecuteSql(SqlCommand mycom) {

                SqlConnection mycon = new SqlConnection(cmpsgsConstr);
                mycom.Connection = mycon;
                mycon.Open();
                //SqlCommand mycom = new SqlCommand(sql, mycon);
                try {
                    mycom.ExecuteNonQuery();
                    mycon.Close();
                } catch(Exception e) {
                    System.Diagnostics.Debug.WriteLine(e.StackTrace);
                    mycon.Close();
                }
            }
            public static DataRow CMESGSQuery4Object(SqlCommand mycom) {
                SqlConnection mycon = new SqlConnection(cmpsgsConstr);
                mycon.Open();

                mycom.Connection = mycon;
                SqlDataReader mydr = mycom.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(mydr);

                mydr.Close();
                mycon.Close();
                if(dt.Rows.Count > 0) {
                    return dt.Rows[0];
                } else {
                    return null;
                }
            }

        }


    }




}