using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Security.Principal;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using System.Web.UI.WebControls;
using WCals.Models.CMBASEModels;

namespace WCals.Helpers
{
    public class WebSiteHelper
    {
        /// <summary>
        /// department type MainDept總處 AreaDept區處 AllNone不屬於總處或區處.
        /// </summary>
        /// <returns></returns>
        public enum deptType { MainDept, AreaDept, AllNone };

        /// <summary>
        /// Gets the Cities.
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<Models.CMBASEModels.SMV_SMCCITY> GetCities()
        {
            //CMBASEEntities 裡沒有建立 SMV_SMCCITY
            using (Models.CMBASEModels.CMBASEEntities db = new CMBASEEntities())
            {
                var query = db.SMV_SMCCITY;
                return query.ToList();
            }
        }

        /// <summary>
        /// Gets the Towns.
        /// </summary>
        /// <param name="city">The city.</param>
        /// <returns></returns>
        public static IEnumerable<Models.CMBASEModels.SMV_SMCTOWN> GetTowns(string city)
        {
            //CMBASEEntities 裡沒有建立 SMV_SMCTOWN
            using (CMBASEEntities db = new CMBASEEntities())
            {
                var query = db.SMV_SMCTOWN.Where(x => x.CITY == city);
                return query.ToList();
            }
        }

        /// <summary>
        /// Gets the Towns.
        /// </summary>
        /// <param name="city">The city.</param>
        /// <returns></returns>
        /*public static IEnumerable<Models.SMCPOSN> GetVilg(String city, String town)
        {
            //CMBASEEntities 裡沒有建立 SMCPOSN
            using (WCalsEntities db = new WCalsEntities())
            {
                var query = (from f in db.SMCPOSN
                             where f.CITY == city && f.TOWN == town
                             select f);

                return query.ToList();
            }
        }*/

        /// <summary>
        /// Gets the Depts.
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<Models.CMBASEModels.OGCDEPT> GetDepts()
        {
            using (CMBASEEntities db = new CMBASEEntities())
            {
                var query = (from f in db.OGCDEPT
                             where f.DATA_TP == "D"
                             select f);
                return query.ToList();
            }
        }

        public static IEnumerable<Models.CMBASEModels.OGCDEPT> GetAreaDepts()
        {
            string[] Depts = new string[] {
                "0100001","0200002","0300003","0400004","0500005",
                "0600006","0700007","0800008","0900009","1000002",
                "1100003","1200004","5000000","5100001","5200002","6100003","6200004"};
            CMBASEEntities db = new CMBASEEntities();
            var OGCDEPTs = db.OGCDEPT.Where(f => Depts.Contains(f.DEPT_NO)).ToList();
            return OGCDEPTs;
        }

        public static string getUserId()
        {
            string backUserId = ""; //string.Empty;
            var identity = HttpContext.Current.User.Identity as FormsIdentity;
            if (identity != null)
            {
                backUserId = identity.Name;
            }

            //return HttpContext.Current.User.Identity.Name;
            return backUserId;
        }

        public static string getUserName()
        {
            string backUserId = getUserId();
            if (backUserId == "")
            {
                return string.Empty;
            }
            else
            {
                return getUserName(backUserId);
            }
        }

        /// <summary>
        /// Systems the name of the user.
        /// </summary>
        /// <param name="id">The identifier.</param>
        /// <returns></returns>
        public static string getUserName(string UserId)
        {
            string userName = string.Empty;

            if (string.IsNullOrWhiteSpace(UserId))
            {
                return userName;
            }

            if (UserId.Equals(String.Empty))
            {
                userName = "系統預設";
            }
            else
            {
                using (CMBASEEntities db = new CMBASEEntities())
                {
                    var user = db.OGCUSER.FirstOrDefault(x => x.USR_ID == UserId);
                    userName = (user == null) ? string.Empty : user.USR_NAME;
                }
            }
            return userName;
        }

        public static string getDeptName(string DEPT_NO)
        {
            CMBASEEntities db = new CMBASEEntities();
            var ogcdept = (from f in db.OGCDEPT
                           where f.DEPT_NO == DEPT_NO && f.DATA_TP == "D"
                           select f).FirstOrDefault();

            if (ogcdept != null)
            {
                return ogcdept.DEPT_NAME;
            }
            else
            {
                return "";
            }
        }
        /// <summary>
        /// 取得目前登入者的部門編號
        /// </summary>
        /// <returns></returns>

        public static string getUserDept(string userID = "")
        {
            if (userID == null || userID.Trim() == "")
            {
                userID = getUserId();
            }

            CMBASEEntities db = new CMBASEEntities();
            var ogcdpur = (from f in db.OGCDPUR
                           where f.USR_ID == userID && f.DATA_TP == "D"
                           select f).FirstOrDefault();

            if (ogcdpur != null)
            {
                return ogcdpur.DEPT_NO;
            }
            else
            {
                return "";
            }

        }

        public static void DataAuthProcess<T>(T t)
        {


        }

        public static Boolean isSystemAdministrator()
        {
            CMBASEEntities db = new CMBASEEntities();
            String userID = HttpContext.Current.User.Identity.Name;
            var ogcdpur = (from f in db.OGCDPUR
                           where f.USR_ID == userID && f.DEPT_NO == "SYS_Admin"
                           select f).FirstOrDefault();

            if (ogcdpur != null)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public static List<String> chiefOfDepartment(string userID)
        {
            var httpContext = HttpContext.Current;
            CMBASEEntities db = new CMBASEEntities();

            var ChiefDepts = (from f in db.OGCDEPT
                              where f.DEPT_CHIEF == userID
                              select f).ToList();

            if (ChiefDepts.Count() != 0)
            {
                List<String> depts = new List<String>();
                foreach (Models.CMBASEModels.OGCDEPT item in ChiefDepts)
                {
                    depts.Add(item.DEPT_NO);
                }
                return depts;
            }
            else
            {
                return null;
            }

        }

        ///getUserDepartmentType 取得目前登入User屬於總處或是區處
        public static deptType getUserDepartmentType()
        {
            CMBASEEntities db = new CMBASEEntities();
            string userID = HttpContext.Current.User.Identity.Name;
            var ogcdpur = (from f in db.OGCDPUR
                           join d in db.OGCDEPT on f.DEPT_NO equals d.DEPT_NO
                           where f.USR_ID == userID && f.DATA_TP == "D"
                           select new { f, d }).ToList();

            Boolean deptMain = false;
            Boolean deptArea = false;
            foreach (var item in ogcdpur)
            {
                if (item.f.DEPT_NO.Substring(0, 2) == "00")
                {
                    deptMain = true;
                }

                if (item.f.DEPT_NO.Substring(0, 2) != "00")
                {
                    deptArea = true;
                }

                //if (item.d.AN_DEPT_NO== "0000113")
                //{
                //    deptArea = true;
                //}

            }

            if (deptArea == true)
            {
                return deptType.AreaDept;
            }
            if (deptMain == true)
            {
                return deptType.MainDept;
            }
            else
            {
                return deptType.AllNone;
            }
        }

        public static string ValidateUser(string ComputerName, string UserName, string Password)
        {
            if (ComputerName.IndexOf('.') != -1)
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ComputerName, UserName, Password);
                //AD上如有建sync帳號就用CN=sync，如無，用CN=administrator也可以
                try
                {
                    string objectSid = (new SecurityIdentifier((byte[])entry.Properties["objectSid"].Value, 0).Value);
                    return objectSid;
                }
                catch
                {
                    return null;
                }
                finally
                {
                    entry.Dispose();
                }
            }
            else
            {
                DirectoryEntry entry = new DirectoryEntry("WinNT://" + ComputerName, UserName, Password);
                try
                {
                    string objectSid = (new SecurityIdentifier((byte[])entry.Properties["objectSid"].Value, 0).Value);
                    return objectSid;
                }
                catch
                {
                    return null;
                }
                finally
                {
                    entry.Dispose();
                }
            }
        }

        public static String replaceFileName(String file_name, String extraInfo)
        {
            String fExt = Path.GetExtension(file_name);

            return String.Format(@"{0}{1}{2}", Path.GetRandomFileName(), extraInfo, fExt);
        }

        public static String GetIPAddress()
        {
            System.Web.HttpContext context = System.Web.HttpContext.Current;
            String ipAddress = context.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];

            if (!String.IsNullOrEmpty(ipAddress))
            {
                String[] addresses = ipAddress.Split(',');
                if (addresses.Length != 0)
                {
                    return addresses[0];
                }
            }

            return context.Request.ServerVariables["REMOTE_ADDR"];
        }

        /*依照權限取得左邊功能選單
         * 
         */
        public static MvcHtmlString getMenuItemsOld()
        {
            var httpContext = HttpContext.Current;
            var identity = httpContext.User.Identity as FormsIdentity;
            var userID = identity.Name;

            String rtnHTML = "";
            CMBASEEntities db = new CMBASEEntities();

            var usergroup = (from f in db.OGCDPUR
                             where f.USR_ID == userID
                             select f.DEPT_NO).ToList();

            var smcauth = (from f in db.SMCAUTH
                           where usergroup.Contains(f.DEPT_NO)
                           select f.SYS_ID).ToList();

            var menus = (from f in db.SMCMENU3
                         where smcauth.Contains(f.SYS_ID)
                         orderby f.SYS_NAME, f.SYS_ORD
                         select f).ToList();

            //取所有的Menu資料以免再讀資料浪費IO時間
            var allmenu = (from f in db.SMCMENU3
                           orderby f.SYS_NAME, f.SYS_ORD
                           select f).ToList();
            DataTable allmenuDt = Utilities.CopyGenericToDataTable(allmenu);

            //複製Menu Table的結構
            DataTable finMenuDt = new DataTable();
            finMenuDt = allmenuDt.Clone();

            DataTable dt = Utilities.CopyGenericToDataTable(menus);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                String AN_SYS_ID = dt.Rows[i]["AN_SYS_ID"].ToString();
                DataRow[] r = finMenuDt.Select("SYS_ID='" + AN_SYS_ID + "'");
                if (r.Count() == 0)
                {
                    DataRow[] rr = allmenuDt.Select("SYS_ID='" + AN_SYS_ID + "'");
                    if (rr.Count() != 0)
                    {
                        DataRow nr = finMenuDt.NewRow();
                        nr["SYS_ID"] = rr[0]["SYS_ID"].ToString();
                        nr["DATA_TP"] = rr[0]["DATA_TP"].ToString();
                        nr["AN_SYS_ID"] = rr[0]["AN_SYS_ID"].ToString();
                        nr["SYS_NAME"] = rr[0]["SYS_NAME"].ToString();
                        nr["APP_PATH"] = rr[0]["APP_PATH"].ToString();
                        nr["SYS_ORD"] = rr[0]["SYS_ORD"].ToString();
                        nr["ADD_DT"] = rr[0]["ADD_DT"];
                        nr["UPD_DT"] = rr[0]["UPD_DT"];
                        nr["UPD_UID"] = rr[0]["UPD_UID"].ToString();
                        nr["URL"] = rr[0]["URL"].ToString();
                        finMenuDt.Rows.Add(nr);
                    }
                }
                DataRow nnr = finMenuDt.NewRow();
                nnr["SYS_ID"] = dt.Rows[i]["SYS_ID"].ToString();
                nnr["DATA_TP"] = dt.Rows[i]["DATA_TP"].ToString();
                nnr["AN_SYS_ID"] = dt.Rows[i]["AN_SYS_ID"].ToString();
                nnr["SYS_NAME"] = dt.Rows[i]["SYS_NAME"].ToString();
                nnr["APP_PATH"] = dt.Rows[i]["APP_PATH"].ToString();
                nnr["SYS_ORD"] = dt.Rows[i]["SYS_ORD"].ToString();
                nnr["ADD_DT"] = dt.Rows[i]["ADD_DT"];
                nnr["UPD_DT"] = dt.Rows[i]["UPD_DT"];
                nnr["UPD_UID"] = dt.Rows[i]["UPD_UID"].ToString();
                nnr["URL"] = dt.Rows[i]["URL"].ToString();
                finMenuDt.Rows.Add(nnr);
            }
            DataRow[] mr = finMenuDt.Select("AN_SYS_ID=''");
            for (int i = 0; i < mr.Count(); i++)
            {
                rtnHTML = rtnHTML + "<li>";
                String SYS_ID = mr[i]["SYS_ID"].ToString();
                DataRow[] dr = finMenuDt.Select("AN_SYS_ID='" + SYS_ID + "'");
                if (dr.Count() != 0) //有子選單
                {
                    rtnHTML = rtnHTML + "<a href='javascript:; ' data-toggle='collapse' title='" + mr[i]["SYS_NAME"].ToString() + "' data-target='#" + mr[i]["SYS_ID"].ToString() + "'>　" + mr[i]["SYS_NAME"].ToString() + "<i class='fa fa-fw fa-caret-down'></i></a><ul id='" + mr[i]["SYS_ID"].ToString() + "' class='collapse' >";
                    for (int j = 0; j < dr.Count(); j++)
                    {
                        rtnHTML = rtnHTML + "<li>";
                        rtnHTML = rtnHTML + "<a href=/SEPMIS2/'" + dr[j]["APP_PATH"].ToString() + "'>" + dr[j]["SYS_NAME"].ToString() + "</a>";
                        rtnHTML = rtnHTML + "</li>";
                    }
                    rtnHTML = rtnHTML + "</ul>";
                }
                else //無子選單
                {
                    rtnHTML = rtnHTML + "<a href=/SEPMIS2/'" + mr[i]["APP_PATH"].ToString() + "' title='" + mr[i]["SYS_NAME"].ToString() + "'>" + mr[i]["SYS_NAME"].ToString() + "<i class='fa fa-fw fa-caret-down'></i></a>";
                }
                rtnHTML = rtnHTML + "</li>";
            }
            return MvcHtmlString.Create(rtnHTML);
        }

        public static MvcHtmlString getMenuItems2()
        {
            var httpContext = HttpContext.Current;
            var identity = httpContext.User.Identity as FormsIdentity;
            if (identity == null)
            {
                return MvcHtmlString.Create("");
            }

            var userID = identity.Name;

            string rtnHTML = "";
            CMBASEEntities db = new CMBASEEntities();

            var usergroup = (from f in db.OGCDPUR
                             where f.USR_ID == userID
                             select f.DEPT_NO).ToList();

            var smcauth = (from f in db.SMCAUTH
                           where usergroup.Contains(f.DEPT_NO)
                           select f.SYS_ID).ToList();

            var menuA = (from f in db.SMCMENU3
                         where f.AN_SYS_ID != null && f.DATA_TP == "A" && smcauth.Contains(f.SYS_ID) //&& (f.SYS_ID.Contains("CMBASE") || f.SYS_ID.Contains("CMRSET") || f.SYS_ID.Contains("CMRSPT"))
                         orderby f.SYS_ORD
                         select f).ToList();
            var menuAAnList = (from f in menuA select f.AN_SYS_ID);


            var menuN = (from f in db.SMCMENU3
                         where f.AN_SYS_ID != null && f.DATA_TP == "N" && menuAAnList.Contains(f.SYS_ID) //&& (f.SYS_ID.Contains("CMBASE") || f.SYS_ID.Contains("CMRSET") || f.SYS_ID.Contains("CMRSPT"))
                         orderby f.SYS_ORD
                         select f).ToList();
            var menuNAnList = (from f in menuN select f.AN_SYS_ID);

            var menuR = (from f in db.SMCMENU3
                         where f.DATA_TP == "R" && menuNAnList.Contains(f.SYS_ID) //&& (f.SYS_ID.Contains("CMBASE") || f.SYS_ID.Contains("CMRSET") || f.SYS_ID.Contains("CMRSPT"))
                         orderby f.SYS_ORD
                         select f).ToList();

            string patternR = @"^(5|6|7|8|9)\.";
            //if (EnvHelper.IsTestMode)
            //    patternR = @"^(00|5|6|7|8|9|10)\.";
            //else
            //    patternR = @"^(1|3|4|5|6|7|8|9|10)\."; 
            //[v250225]先移除新系統尚未完成的功能頁面 
            // TODO：[260212]放到正式環境時需更改成正式環境使用的資料，如下
            // 目前正式環境只開放2，其他功能開放時再更新
            // 正式環境：string patternR = @"^(00|0|1|3|4|5|6|7|8|9|10)\.";
            // 測試環境：string patternR = @"^(00|5|6|7|8|9|10)\.";
            // [260331]正式環境與測試環境開放功能相同patternR = @"^(5|6|7|8|9)\.";

            string patternA = @"CMPSGS.SGE00(10|25|35|60)"; //[v250225]移除新系統1和2管理功能內不需要的功能頁面
            string sysIdR = "";
            string sysIdN = "";
            //string sysIdA = "";
            foreach (var itemR in menuR)
            {
                if (Regex.IsMatch(itemR.SYS_NAME, patternR)) continue;

                sysIdR = itemR.SYS_ID;
                var itemNList = menuN.Where(f => f.AN_SYS_ID == sysIdR).ToList();
                if (itemNList != null && itemNList.Count > 0)
                {
                    rtnHTML = rtnHTML + "<li class='dropdown'>";
                    rtnHTML = rtnHTML + "<a href='#' class='dropdown-toggle' data-toggle='dropdown' role='button' aria-haspopup='true' aria-expanded='false'>" + itemR.SYS_NAME + "</a>";
                    rtnHTML = rtnHTML + "<ul class='dropdown-menu'>";

                    foreach (var itemN in itemNList)
                    {
                        sysIdN = itemN.SYS_ID;
                        var itemAList = menuA.Where(f => f.AN_SYS_ID == sysIdN && !Regex.IsMatch(f.SYS_ID, patternA)).ToList();
                        if (itemAList != null && itemAList.Count > 0)
                        {
                            rtnHTML = rtnHTML + "<li class='dropdown-submenu'><a href='#' class='dropdown-toggle' data-toggle='dropdown' role='button' aria-haspopup='true' aria-expanded='false'>" + itemN.SYS_NAME + "</a>";
                            rtnHTML = rtnHTML + "<ul class='dropdown-menu'>";

                            foreach (var itemA in itemAList)
                            {
                                //rtnHTML = rtnHTML + "<div style='margin-left:55px'><br />";
                                //rtnHTML = rtnHTML + "<a style='color:#ffffff' href='" + itemA.FILE_PATH + "'>" + itemA.SYS_NAME + "</a>";
                                rtnHTML = rtnHTML + "<li><a href='/SEPMIS2/" + itemA.APP_PATH + "'>" + itemA.SYS_NAME + "</a></li>";
                            }
                            rtnHTML = rtnHTML + "</ul>";
                            //rtnHTML = rtnHTML + "</div>";
                            rtnHTML = rtnHTML + "</li>";
                        }
                    }

                    rtnHTML = rtnHTML + "</ul>";
                    rtnHTML = rtnHTML + "</li>";
                }
            }

            return MvcHtmlString.Create(rtnHTML);
        }

        /// <summary>
        /// BS5 版主選單產生器（供 _Layout2_v2 使用）。
        ///
        /// 行為邏輯完全比照 getMenuItems2()（git 版為準，含歷次依使用者需求做的調整）：
        ///   - patternR / patternA 的過濾條件與 getMenuItems2() 一致
        ///   - patternA 於 LINQ 內先過濾再判斷 Count，中層若無可見子項則整個不輸出
        ///     （若改成在 foreach 內 continue，會畫出空的子選單）
        ///   - menuR 不加任何額外條件（例如 SUGGEST 特例）
        /// 與 getMenuItems2() 的唯一差異是輸出的 HTML 改為 Bootstrap 5 標記：
        ///   nav-item / nav-link / dropdown-item / data-bs-toggle / submenu。
        /// 兩者的過濾條件日後必須同步修改。
        /// </summary>
        public static MvcHtmlString getMenuItems2_v2()
        {
            var httpContext = HttpContext.Current;
            var identity = httpContext.User.Identity as FormsIdentity;
            if (identity == null)
            {
                return MvcHtmlString.Create("");
            }

            var userID = identity.Name;

            string rtnHTML = "";
            CMBASEEntities db = new CMBASEEntities();

            var usergroup = (from f in db.OGCDPUR
                             where f.USR_ID == userID
                             select f.DEPT_NO).ToList();

            var smcauth = (from f in db.SMCAUTH
                           where usergroup.Contains(f.DEPT_NO)
                           select f.SYS_ID).ToList();

            var menuA = (from f in db.SMCMENU3
                         where f.AN_SYS_ID != null && f.DATA_TP == "A" && smcauth.Contains(f.SYS_ID)
                         orderby f.SYS_ORD
                         select f).ToList();
            var menuAAnList = (from f in menuA select f.AN_SYS_ID);

            var menuN = (from f in db.SMCMENU3
                         where f.AN_SYS_ID != null && f.DATA_TP == "N" && menuAAnList.Contains(f.SYS_ID)
                         orderby f.SYS_ORD
                         select f).ToList();
            var menuNAnList = (from f in menuN select f.AN_SYS_ID);

            var menuR = (from f in db.SMCMENU3
                         where f.DATA_TP == "R" && menuNAnList.Contains(f.SYS_ID)
                         orderby f.SYS_ORD
                         select f).ToList();

            // [260331]正式環境與測試環境開放功能相同（與 getMenuItems2 保持一致）
            string patternR = @"^(5|6|7|8|9)\.";
            // [v250225]先期新系統1和2管理功能內不需要的功能頁面
            string patternA = @"CMPSGS.SGE00(10|25|35|60)";

            string sysIdR = "";
            string sysIdN = "";
            foreach (var itemR in menuR)
            {
                if (Regex.IsMatch(itemR.SYS_NAME, patternR)) continue;

                sysIdR = itemR.SYS_ID;
                var itemNList = menuN.Where(f => f.AN_SYS_ID == sysIdR).ToList();
                if (itemNList != null && itemNList.Count > 0)
                {
                    rtnHTML = rtnHTML + "<li class='nav-item dropdown'>";
                    rtnHTML = rtnHTML + "<a href='#' class='nav-link dropdown-toggle' data-bs-toggle='dropdown' data-bs-auto-close='outside' role='button' aria-haspopup='true' aria-expanded='false'>" + itemR.SYS_NAME + "</a>";
                    rtnHTML = rtnHTML + "<ul class='dropdown-menu'>";

                    foreach (var itemN in itemNList)
                    {
                        sysIdN = itemN.SYS_ID;
                        // 先濾掉 patternA 再判斷有無子項，避免畫出空的子選單
                        var itemAList = menuA.Where(f => f.AN_SYS_ID == sysIdN && !Regex.IsMatch(f.SYS_ID, patternA)).ToList();
                        if (itemAList != null && itemAList.Count > 0)
                        {
                            rtnHTML = rtnHTML + "<li class='dropdown-submenu dropdown-item-container'><a href='#' class='dropdown-item dropdown-toggle d-flex justify-content-between align-items-center' role='button' aria-haspopup='true' aria-expanded='false'>" + itemN.SYS_NAME + "</a>";
                            rtnHTML = rtnHTML + "<ul class='dropdown-menu submenu'>";

                            foreach (var itemA in itemAList)
                            {
                                rtnHTML = rtnHTML + "<li><a class='dropdown-item' href='/SEPMIS2/" + itemA.APP_PATH + "'>" + itemA.SYS_NAME + "</a></li>";
                            }
                            rtnHTML = rtnHTML + "</ul>";
                            rtnHTML = rtnHTML + "</li>";
                        }
                    }

                    rtnHTML = rtnHTML + "</ul>";
                    rtnHTML = rtnHTML + "</li>";
                }
            }

            return MvcHtmlString.Create(rtnHTML);
        }


        public static MvcHtmlString getMenuItems5()
        {
            var httpContext = HttpContext.Current;
            var identity = httpContext.User.Identity as FormsIdentity;
            if (identity == null) return MvcHtmlString.Create("");

            var userID = identity.Name;
            string rtnHTML = "";

            using (CMBASEEntities db = new CMBASEEntities())
            {
                var usergroup = db.OGCDPUR
                    .Where(f => f.USR_ID == userID)
                    .Select(f => f.DEPT_NO)
                    .ToList();

                var smcauth = db.SMCAUTH
                    .Where(f => usergroup.Contains(f.DEPT_NO))
                    .Select(f => f.SYS_ID)
                    .ToList();

                var menuA = db.SMCMENU3
                    .Where(f => f.AN_SYS_ID != null && f.DATA_TP == "A" && smcauth.Contains(f.SYS_ID))
                    .OrderBy(f => f.SYS_ORD)
                    .ToList();

                var menuAAnList = menuA.Select(f => f.AN_SYS_ID);

                var menuN = db.SMCMENU3
                    .Where(f => f.AN_SYS_ID != null && f.DATA_TP == "N" && menuAAnList.Contains(f.SYS_ID))
                    .OrderBy(f => f.SYS_ORD)
                    .ToList();

                var menuNAnList = menuN.Select(f => f.AN_SYS_ID);

                var menuR = db.SMCMENU3
                    .Where(f => f.DATA_TP == "R" && menuNAnList.Contains(f.SYS_ID))
                    .OrderBy(f => f.SYS_ORD)
                    .ToList();

                string patternR = @"^(4|5|6|7|8|9|10)\.";
                string patternA = @"CMPSGS.SGE00(10|25|35)";
                /*
                <li class='dropdown-submenu'>
  


                */
                foreach (var itemR in menuR)
                {
                    if (Regex.IsMatch(itemR.SYS_NAME, patternR)) continue;

                    var itemNList = menuN.Where(f => f.AN_SYS_ID == itemR.SYS_ID).ToList();
                    if (itemNList.Count == 0) continue;

                    rtnHTML += "<li class='nav-item dropdown'>";
                    rtnHTML += $"<a class='nav-link dropdown-toggle' href='#' role='button' data-bs-toggle='dropdown' aria-expanded='false'>{itemR.SYS_NAME}</a>";
                    rtnHTML += "<ul class='dropdown-menu'>";

                    foreach (var itemN in itemNList)
                    {
                        var itemAList = menuA.Where(f => f.AN_SYS_ID == itemN.SYS_ID && !Regex.IsMatch(f.SYS_ID, patternA)).ToList();
                        if (itemAList.Count == 0) continue;

                        rtnHTML += "<li class='dropend'>";
                        rtnHTML += $"<a class='dropdown-item dropdown-toggle' href='#' data-bs-toggle='dropdown' aria-expanded='false'>{itemN.SYS_NAME}</a>";
                        rtnHTML += "<ul class='dropdown-menu'>";

                        foreach (var itemA in itemAList)
                        {
                            rtnHTML += $"<li><a class='dropdown-item' href='/SEPMIS2/{itemA.APP_PATH}'>{itemA.SYS_NAME}</a></li>";
                        }

                        rtnHTML += "</ul></li>"; // close itemN
                    }

                    rtnHTML += "</ul></li>"; // close itemR
                }
            }

            return MvcHtmlString.Create(rtnHTML);
        }


        public static MvcHtmlString getMenuItems()
        {
            var httpContext = HttpContext.Current;
            var identity = httpContext.User.Identity as FormsIdentity;
            if (identity == null)
            {
                return MvcHtmlString.Create("");
            }

            var userID = identity.Name;

            string rtnHTML = "";
            CMBASEEntities db = new CMBASEEntities();

            var usergroup = (from f in db.OGCDPUR
                             where f.USR_ID == userID
                             select f.DEPT_NO).ToList();

            var smcauth = (from f in db.SMCAUTH
                           where usergroup.Contains(f.DEPT_NO)
                           select f.SYS_ID).ToList();

            var menuA = (from f in db.SMCMENU3
                         where f.AN_SYS_ID != null && f.DATA_TP == "A" && smcauth.Contains(f.SYS_ID) //&& (f.SYS_ID.Contains("CMBASE") || f.SYS_ID.Contains("CMRSET") || f.SYS_ID.Contains("CMRSPT"))
                         orderby f.SYS_ORD
                         select f).ToList();
            var menuAAnList = (from f in menuA select f.AN_SYS_ID);


            var menuN = (from f in db.SMCMENU3
                         where f.AN_SYS_ID != null && f.DATA_TP == "N" && menuAAnList.Contains(f.SYS_ID) //&& (f.SYS_ID.Contains("CMBASE") || f.SYS_ID.Contains("CMRSET") || f.SYS_ID.Contains("CMRSPT"))
                         orderby f.SYS_ORD
                         select f).ToList();
            var menuNAnList = (from f in menuN select f.AN_SYS_ID);

            var menuR = (from f in db.SMCMENU3
                         where f.DATA_TP == "R" && menuNAnList.Contains(f.SYS_ID) //&& (f.SYS_ID.Contains("CMBASE") || f.SYS_ID.Contains("CMRSET") || f.SYS_ID.Contains("CMRSPT"))
                         orderby f.SYS_ORD
                         select f).ToList();



            string sysIdR = "";
            string sysIdN = "";
            string sysIdA = "";
            foreach (var itemR in menuR)
            {
                sysIdR = itemR.SYS_ID;
                var itemNList = menuN.Where(f => f.AN_SYS_ID == sysIdR).ToList();
                if (itemNList != null && itemNList.Count > 0)
                {
                    rtnHTML = rtnHTML + "<li class='menuli'>";
                    rtnHTML = rtnHTML + "<a href='javascript:; ' data-toggle='collapse' title='" + itemR.SYS_NAME + "' data-target='#" + itemR.SYS_ID + "'>　" + itemR.SYS_NAME + "<i class='fa fa-fw fa-caret-down'></i></a>";
                    rtnHTML = rtnHTML + "<ul id='" + itemR.SYS_ID + "' class='collapse' >";

                    foreach (var itemN in itemNList)
                    {
                        sysIdN = itemN.SYS_ID;
                        var itemAList = menuA.Where(f => f.AN_SYS_ID == sysIdN).ToList();
                        if (itemAList != null && itemAList.Count > 0)
                        {
                            rtnHTML = rtnHTML + "<li class='menuli'>";
                            rtnHTML = rtnHTML + "<a href='javascript:; ' data-toggle='collapse' title='" + itemN.SYS_NAME + "' onclick=\"menuOnclick('" + itemN.SYS_ID + "')\"  class='' aria-expanded='true'>　" + itemN.SYS_NAME + "<i class='fa fa-fw fa-caret-down'></i></a>";
                            rtnHTML = rtnHTML + "<div id='" + itemN.SYS_ID + "' style='display:none; visibility:hidden' >";

                            foreach (var itemA in itemAList)
                            {
                                rtnHTML = rtnHTML + "<div style='margin-left:55px'><br />";
                                //rtnHTML = rtnHTML + "<a style='color:#ffffff' href='" + itemA.FILE_PATH + "'>" + itemA.SYS_NAME + "</a>";
                                rtnHTML = rtnHTML + "<a style='color:#ffffff' href='/SEPMIS2/" + itemA.APP_PATH + "'>" + itemA.SYS_NAME + "</a>";
                                rtnHTML = rtnHTML + "</div>";
                            }

                            rtnHTML = rtnHTML + "</div>";
                            rtnHTML = rtnHTML + "</li>";
                        }
                    }

                    rtnHTML = rtnHTML + "</ul>";
                    rtnHTML = rtnHTML + "</li>";
                }
            }

            return MvcHtmlString.Create(rtnHTML);
        }

        public static List<SMCCODE> getCmbaseSysCodes(String sysId, String catyId)
        {
            CMBASEEntities db = new CMBASEEntities();
            var result = (from f in db.SMCCODE
                          where f.SYS_ID == sysId && f.CATY_ID == catyId
                          select f).ToList();

            return result;
        }

        public static void logFunctionInfo(String UserId, String Function, String Action)
        {
            //CMBASEEntities 裡沒有建立 SMCLOGS
            CMBASEEntities db = new CMBASEEntities();
            Models.CMBASEModels.SMCLOGS log = new SMCLOGS();
            log.USR_ID = UserId;
            log.LOGIN_DT = DateTime.Now;
            log.LOGIN_IP = GetIPAddress();
            log.ACTION_DESC = string.Format("使用者: {0}，使用功能: {1}，動作: {2}", UserId, Function, Action);
            log.ACTION_TYPE = "F";
            try
            {
                db.SMCLOGS.Add(log);
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static void logInfo(String UserId, String Action)
        {
            //CMBASEEntities 裡沒有建立 SMCLOGS
            CMBASEEntities db = new CMBASEEntities();
            Models.CMBASEModels.SMCLOGS log = new SMCLOGS();
            log.USR_ID = UserId;
            log.LOGIN_DT = DateTime.Now;
            log.LOGIN_IP = GetIPAddress();
            log.ACTION_DESC = Action;
            try
            {
                db.SMCLOGS.Add(log);
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static void logUserInfo(String UserId)
        {
            //CMBASEEntities 裡沒有建立 SMCLOGS
            CMBASEEntities db = new CMBASEEntities();
            Models.CMBASEModels.SMCLOGS log = new SMCLOGS();
            log.USR_ID = UserId;
            log.LOGIN_DT = DateTime.Now;
            log.LOGIN_IP = GetIPAddress();
            log.ACTION_DESC = "Login";
            log.ACTION_TYPE = "L";
            try
            {
                db.SMCLOGS.Add(log);
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        /*public static IQueryable<Models.SMCOGLS> getQueryForSMCOGLSData(ViewModels.SMCOGLSSearchViewModel SearchParameter, bool isCanceledData)
        {
            //CMBASEEntities 裡沒有建立 SMCOGLS
            WCalsEntities db = new WCalsEntities();
            var query = db.SMCOGLS.AsQueryable();

            if (null != SearchParameter)
            {
                if (!string.IsNullOrWhiteSpace(SearchParameter.MCH_CODE))
                {
                    SearchParameter.MCH_CODE = SearchParameter.MCH_CODE.Trim();
                    query = query.Where(
                        f => f.MCH_CODE.Contains(SearchParameter.MCH_CODE));
                }
                if (!string.IsNullOrWhiteSpace(SearchParameter.MCH_NAME))
                {
                    SearchParameter.MCH_NAME = SearchParameter.MCH_NAME.Trim();
                    query = query.Where(
                        f => f.MCH_NAME.Contains(SearchParameter.MCH_NAME));
                }
                if (!string.IsNullOrWhiteSpace(SearchParameter.ADDRESS))
                {
                    SearchParameter.ADDRESS = SearchParameter.ADDRESS.Trim();
                    query = query.Where(
                        f => f.ADDRESS.Contains(SearchParameter.ADDRESS));
                }
            }

            if (isCanceledData)
            {
                query = query.Where(
                    f => f.CAN_FG.Equals("是"));
            }
            else
            {
                query = query.Where(
                    f => f.CAN_FG == null || f.CAN_FG != "是");
            }

            return query = query.OrderBy(f => f.MCH_CODE);
        }*/

        /// <summary>
        /// 依照登入者部門，取得直到總處的父部門清單
        /// </summary>
        /// <returns></returns>
        public static List<string> getParentDeptInfo(String UserId)
        {
            CMBASEEntities db = new CMBASEEntities();
            var ogcdpur = (from f in db.OGCDPUR
                           where f.USR_ID == UserId && f.DATA_TP == "D"
                           select f).FirstOrDefault();

            List<string> parentDeptList = new List<string>();
            if (null != ogcdpur)
            {
                iterateParentDept(parentDeptList, ogcdpur.DEPT_NO);
                //以最上層(總處)為第一件做排序
                parentDeptList.Reverse();
            }
            else
            {
                parentDeptList.Add("0000113");
            }

            return parentDeptList;
        }


        /// <summary>
        /// 依照登入者部門，取得直到總處的父部門清單
        /// </summary>
        /// <returns></returns>
        public static DataTable getParentDept(String UserId)
        {
            CMBASEEntities db = new CMBASEEntities();
            var ogcdpur = (from f in db.OGCDPUR
                           where f.USR_ID == UserId && f.DATA_TP == "D"
                           select f).FirstOrDefault();

            DataTable parentDeptList = new DataTable();
            parentDeptList.Columns.Add("id");
            parentDeptList.Columns.Add("name");
            if (null != ogcdpur)
            {
                iterateParentDept(parentDeptList, ogcdpur.DEPT_NO);
                //以最上層(總處)為第一件做排序
                //parentDeptList.Reverse();
            }
            else
            {
                //parentDeptList.Add("0000113");
            }

            return parentDeptList;
        }

        private static void iterateParentDept(DataTable parentDeptList, string DEPT_NO)
        {
            CMBASEEntities db = new CMBASEEntities();

            var ogcdept = (from f in db.OGCDEPT
                           where f.DEPT_NO == DEPT_NO
                           select f).FirstOrDefault();

            if (null != ogcdept && !string.IsNullOrEmpty(ogcdept.DEPT_NO))
            {
                DataRow dr = parentDeptList.NewRow();
                dr["id"] = ogcdept.DEPT_NO;
                dr["name"] = ogcdept.DEPT_NAME;
                parentDeptList.Rows.Add(dr);
                if (!ogcdept.DEPT_NO.Equals("0000113") && !string.IsNullOrEmpty(ogcdept.AN_DEPT_NO))
                {
                    iterateParentDept(parentDeptList, ogcdept.AN_DEPT_NO);
                }
            }
        }



        private static void iterateParentDept(List<string> parentDeptList, string DEPT_NO)
        {
            CMBASEEntities db = new CMBASEEntities();

            var ogcdept = (from f in db.OGCDEPT
                           where f.DEPT_NO == DEPT_NO
                           select f).FirstOrDefault();

            if (null != ogcdept && !string.IsNullOrEmpty(ogcdept.DEPT_NO))
            {
                parentDeptList.Add(ogcdept.DEPT_NO);
                if (!ogcdept.DEPT_NO.Equals("0000113") && !string.IsNullOrEmpty(ogcdept.AN_DEPT_NO))
                {
                    iterateParentDept(parentDeptList, ogcdept.AN_DEPT_NO);
                }
            }
        }

        public static HashSet<string> getAuthDepts(String UserId)
        {

            // 取得下層部門與跨部門管理之部門
            HashSet<string> authorizedDepts = getChildDeptInfo(UserId);

            // 取得身為主管的部門
            List<string> depts = chiefOfDepartment(UserId);

            // 是部門主管，將部門放入權限，利用Set資料型態過濾重複部門
            if (depts != null)
            {
                foreach (string dept in depts)
                {
                    authorizedDepts.Add(dept);
                }
            }

            return authorizedDepts;
        }

        /// <summary>
        /// 依照登入者部門，取得所有子部門與跨部門管理之資料
        /// </summary>
        /// <returns></returns>
        public static HashSet<string> getChildDeptInfo(string UserId)
        {
            //CMBASEEntities 裡沒有建立 SMV_SMCADMC
            CMBASEEntities db = new CMBASEEntities();
            var ogcdpur = (from f in db.OGCDPUR
                           where f.USR_ID == UserId && f.DATA_TP == "D"
                           select f).FirstOrDefault();

            var childDeptSet = new HashSet<string>();
            /*if (null != ogcdpur)
            {
                iterateChildDept(ref childDeptSet, ogcdpur.DEPT_NO);

                // 跨部門管理
                var smv_smcadmc = (from f in db.SMV_SMCADMC
                                   where f.ADV_DP == ogcdpur.DEPT_NO
                                   select f).ToList();

                if (null != smv_smcadmc && smv_smcadmc.Count > 0)
                {
                    foreach (SMV_SMCADMC item in smv_smcadmc)
                    {
                        childDeptSet.Add(item.CADV_DP);
                    }
                }
            }*/


            return childDeptSet;
        }


        private static void iterateChildDept(ref HashSet<string> childDeptSet, string DEPT_NO)
        {
            CMBASEEntities db = new CMBASEEntities();

            var ogcdept = (from f in db.OGCDEPT
                           where f.AN_DEPT_NO == DEPT_NO && f.DATA_TP == "D"
                           select f).ToList();

            if (null != ogcdept && ogcdept.Count() != 0)
            {
                foreach (Models.CMBASEModels.OGCDEPT item in ogcdept)
                {
                    childDeptSet.Add(item.DEPT_NO);
                    iterateChildDept(ref childDeptSet, item.DEPT_NO);
                }
            }
        }

        public static byte[] getFileFromPath(String hrefFileName)
        {
            String UploadPath = System.Web.Configuration.WebConfigurationManager.AppSettings["uploadPath"];
            string fullPath = HttpContext.Current.Server.MapPath("~/" + UploadPath) + Path.GetFileName(hrefFileName);
            byte[] fileBytes = File.ReadAllBytes(fullPath);

            return fileBytes;
        }

        //提供工料自動編碼用的資料來源
        /// <summary>
        /// Gets the 章碼分類.
        /// </summary>
        /// <param name="">章碼分類</param>
        /// <returns></returns>
        /*public static IEnumerable<ViewModels.AutoNumViewModel> GetChapCaty(string CostKind)
        {
            //CMBASEEntities 裡沒有建立 AutoNumA
            using (CMBASEEntities db = new CMBASEEntities())
            {
                IEnumerable<ViewModels.AutoNumViewModel> query;

                if (CostKind != "EL")
                {
                    query = (from f in db.AutoNumA
                             where f.levelNo == "1"
                             select new ViewModels.AutoNumViewModel
                             {
                                 itemCode = f.itemCode.Substring(0, 2),
                                 icName = f.itemCode.Substring(0, 2) + "," + f.cName
                             });

                }
                else
                {
                    query = (from f in db.AutoNumA
                             where f.levelNo == "2" && f.parent != "L" && f.parent != "E"

                             select new ViewModels.AutoNumViewModel
                             {
                                 itemCode = f.itemCode.Substring(0, 5),
                                 icName = f.itemCode.Substring(0, 5) + "," + f.cName
                             });

                }
                return query.ToList();

            }
        }*/

        /*public static IEnumerable<ViewModels.AutoNumViewModel> GetChapCode(string ChapCaty)
        {
            //CMBASEEntities 裡沒有建立 AutoNumA
            using (WCalsEntities db = new WCalsEntities())
            {
                var query = (from f in db.AutoNumA
                             where f.levelNo == "2" && f.parent == ChapCaty
                             select new ViewModels.AutoNumViewModel
                             {
                                 itemCode = f.itemCode.Substring(0, 5),
                                 icName = f.itemCode.Substring(0, 5) + "," + f.cName.Trim() + f.IsShow
                             });
                return query.ToList();
            }
        }

        public static IEnumerable<ViewModels.AutoNumViewModel> GetELNNCode(string CostKind, string CodeSection, string ChapCode, string key)
        {
            //CMBASEEntities 裡沒有建立 AutoNumB_12
            WCalsEntities db = new WCalsEntities();
            var query = db.AutoNumB_12.AsNoTracking().Where(f => f.CodeSection == CodeSection && f.resType == CostKind && (f.Code != "" && f.Code != null)).AsQueryable();
            if (string.IsNullOrEmpty(ChapCode) == false)
            {
                query = query.Where(f => f.ChapCode == ChapCode);
            }

            if (string.IsNullOrEmpty(key) == false && key.IndexOf(',') != -1)
            {
                string[] keys = key.Split(',');
                if (string.IsNullOrEmpty(keys[1]) == false)
                {
                    int minRow = Convert.ToInt32(keys[1]);
                    query = query.Where(f => f.MinRow <= minRow);
                }
                if (string.IsNullOrEmpty(keys[2]) == false)
                {
                    int maxRow = Convert.ToInt32(keys[2]);
                    query = query.Where(f => f.MaxRow >= maxRow);
                }
            }

            var datas = query.OrderBy(f => f.Code).GroupBy(f => new { f.Code, f.MinRow, f.MaxRow, f.Content });
            var models = new List<ViewModels.AutoNumViewModel>();

            foreach (var entry in datas)
            {
                var model = new ViewModels.AutoNumViewModel();

                model.itemCode = entry.Key.Code.ToString() + "," + entry.Key.MinRow.ToString() + "," + entry.Key.MaxRow.ToString();
                if (string.IsNullOrEmpty(key) == false && (entry.Key.MinRow.ToString() != key.Split(',')[1] || entry.Key.MaxRow.ToString() != key.Split(',')[2]))
                {
                    model.itemCode = entry.Key.Code.ToString() + "," + key.Split(',')[1] + "," + key.Split(',')[2];
                }

                model.icName = entry.Key.Code.ToString() + "," + entry.Key.Content;
                models.Add(model);
            }

            if (datas.Count() == 0)
            {
                var model = new ViewModels.AutoNumViewModel();
                model.itemCode = "0," + (String.IsNullOrEmpty(key) ? "" : key.Split(',')[1] + "," + key.Split(',')[2]);

                //model.itemCode = "0," + key.Split(',')[1] + "," + key.Split(',')[2];
                model.icName = "0,";
                models.Add(model);
            }

            return models;
        }

        public static IEnumerable<ViewModels.AutoNumViewModel> GetPSNNCode(string lev, string ChapCode, string key)
        {
            //CMBASEEntities 裡沒有建立 AutoNumB
            WCalsEntities db = new WCalsEntities();

            var query = db.AutoNumB.AsNoTracking().Where(f => f.CodeSection == lev && f.IsCustom == "N" && f.Version == "0" && (f.Code != "" && f.Code != null) && f.ChapCode == ChapCode).AsQueryable();
            //if (lev!= "09" && lev!= "10")
            if (lev != "10")
            {
                if (string.IsNullOrEmpty(key) == false && key.IndexOf(',') != -1)
                {
                    string[] keys = key.Split(',');
                    if (string.IsNullOrEmpty(keys[1]) == false)
                    {
                        int minRow = Convert.ToInt32(keys[1]);
                        query = query.Where(f => f.MinRow <= minRow);
                    }
                    if (string.IsNullOrEmpty(keys[2]) == false)
                    {
                        int maxRow = Convert.ToInt32(keys[2]);
                        query = query.Where(f => f.MaxRow >= maxRow);
                    }
                }
            }
            //var datas = query = query.GroupBy(f=> f.Code, f=> f.Content);

            //var datas = query.OrderBy(f=> f.Code).GroupBy( f=> new { f.Code, f.MinRow, f.MaxRow, f.Content });
            var datas = query.OrderBy(f => f.Code).GroupBy(f => new { f.Code, f.MinRow, f.MaxRow, f.Content });

            var models = new List<ViewModels.AutoNumViewModel>();
            foreach (var entry in datas)
            {
                var model = new ViewModels.AutoNumViewModel();

                model.itemCode = entry.Key.Code.Substring(0, 1) + "," + entry.Key.MinRow.ToString() + "," + entry.Key.MaxRow.ToString();

                if (string.IsNullOrEmpty(key) == false && (entry.Key.MinRow.ToString() != key.Split(',')[1] || entry.Key.MaxRow.ToString() != key.Split(',')[2]))
                {
                    model.itemCode = entry.Key.Code.Substring(0, 1) + "," + key.Split(',')[1] + "," + key.Split(',')[2];
                }

                model.icName = entry.Key.Code.Substring(0, 1) + "," + entry.Key.Content;
                models.Add(model);
            }
            if (datas.Count() == 0)
            {
                var model = new ViewModels.AutoNumViewModel();
                //model.itemCode = "0,-1,-1";
                //model.itemCode = "0," + key.Split(',')[1] + "," + key.Split(',')[2];
                model.itemCode = "0," + (String.IsNullOrEmpty(key) ? "" : key.Split(',')[1] + "," + key.Split(',')[2]);
                model.icName = "0,";
                models.Add(model);
            }
            return models;
        }*/

        public static string sendMail(string sendTo, string subject, string emailBody)
        {
            try
            {
                int port = Convert.ToInt16(System.Web.Configuration.WebConfigurationManager.AppSettings["EmailPort"].ToString());
                String EmailHost = System.Web.Configuration.WebConfigurationManager.AppSettings["EmailHost"];
                bool enableSSL = false;
                if (System.Web.Configuration.WebConfigurationManager.AppSettings["EmailSSL"].ToString() == "Y")
                {
                    enableSSL = true;
                }
                System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                if (sendTo.IndexOf(';') == -1)
                {
                    mail.To.Add(sendTo);
                }
                else //多筆Email
                {
                    string[] eAddress = sendTo.Split(';');
                    for (int i = 0; i < eAddress.Length; i++)
                    {
                        mail.To.Add(eAddress[i]);
                    }
                }
                mail.From = new MailAddress(System.Web.Configuration.WebConfigurationManager.AppSettings["EmailFrom"], "系統管理員", System.Text.Encoding.UTF8);
                mail.Subject = subject;
                mail.SubjectEncoding = System.Text.Encoding.UTF8;

                mail.Body = emailBody;
                mail.BodyEncoding = System.Text.Encoding.UTF8;
                mail.IsBodyHtml = true;
                SmtpClient client = new SmtpClient();
                client.Port = port;
                client.Host = EmailHost;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential(System.Web.Configuration.WebConfigurationManager.AppSettings["EmailID"], System.Web.Configuration.WebConfigurationManager.AppSettings["EmailPWD"]);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = enableSSL;
                client.Send(mail);

                return "0";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        /// <summary>
        /// 取得目前所在的主計畫名稱(+分支計畫名稱)
        /// </summary>
        /// <param name="anPid"> 主計畫編號 </param>
        /// <param name="pid"> 分支計畫編號 </param>
        /// <returns></returns>
        public static string GetCurrentProjectName(decimal anPid, decimal pid = 0)
        {
            Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();

            var anPjName = string.Empty;
            var pjName = string.Empty;

            if (pid != 0)
            {
                //若有分支計畫編號(無主計畫編號)
                var p = db.PJMFORM.FirstOrDefault(f => f.PID == pid); //取得分支計畫資料
                pjName = p?.PJ_NAME ?? ""; //取得分支計畫名稱
                anPid = (anPid == 0) ? p?.AN_PID ?? 0 : anPid; //取得主計畫編號

                if (anPid != 0)
                {
                    var ap = db.PJMFORM.FirstOrDefault(f => f.PID == anPid); //取得主計畫資料
                    anPjName = string.Format("[{1}]{0}>", ap?.PJ_NAME, ap?.SYR); //取得主計畫名稱+主計畫年度
                }

                return string.Format("{0}{1}", anPjName, pjName);
            }
            else
            {
                //若有主計畫編號(無分支計畫編號)
                var ap = db.PJMFORM.FirstOrDefault(f => f.PID == anPid); //取得主計畫資料
                anPjName = string.Format("[{1}]{0}", ap?.PJ_NAME, ap?.SYR); //取得主計畫名稱+主計畫年度

                return anPjName;
            }
        }
    }
}