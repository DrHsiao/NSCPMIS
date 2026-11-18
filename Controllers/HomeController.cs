using PagedList;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using WCals.Models.CMBASEModels;
using WCals.ViewModels;

namespace WCals.Controllers
{
    public partial class Constants
    {
        public static string AUTH_DEPT = "authDept";
        public static string E_MAIL_DOMAIN_NAME = "@mail.water.gov.tw";
        public const string PAGE_NUMBER_KEY = "tempPN";
        public const string SEARCH_PARAMETER_KEY = "tempSP";
    }

    [Authorize]
    public class HomeController : Controller
    {
        private readonly Models.CMBASEModels.CMBASEEntities db = new Models.CMBASEModels.CMBASEEntities();

        // GET: Home
        public ActionResult Index()
        {
            var result = db.SMCNEWS
                .Where(x => (x.STARTDATE != null && DbFunctions.TruncateTime(x.STARTDATE) <= DbFunctions.TruncateTime(DateTime.Today)) &&
                (x.ENDDATE == null || DbFunctions.TruncateTime(x.ENDDATE) >= DbFunctions.TruncateTime(DateTime.Today)))
                .OrderByDescending(o => o.STARTDATE);
            var news = result.Where(x => x.TYPE == "01").ToList();
            var smcdocs = result.Where(x => x.TYPE == "02").ToList();

            HomeIndexViewModel viewModel = new HomeIndexViewModel();
            viewModel.ListSMCNEWS = news;
            viewModel.ListSMCDOCS = smcdocs;
            return View(viewModel);
        }

        public ActionResult News(String id)
        {
            try
            {
                int infoid = Convert.ToInt32(id);
                var news = (from f in db.SMCNEWS
                            where f.NEWSID == infoid
                            select f).FirstOrDefault();

                if (news == null)
                {
                    return RedirectToAction("ErrorPage", "Error", new { errorCode = 999 });
                }

                // 取得公告檔案
                var files = db.SMCNEWSDOC.Where(f => f.NEWSID == infoid).AsEnumerable();

                var model = new HomeIndexViewModel
                {
                    dataSMCNEWS = news,
                    ListSMCNEWSDOC = files
                };
                return View(model);
            }
            catch
            {
                return RedirectToAction("ErrorPage", "Error", new { errorCode = 999 });
            }
        }

        public ActionResult Docs(String id)
        {
            if (String.IsNullOrEmpty(id) == false)
            {
                int infoid = Convert.ToInt16(id);
                var docs = (from f in db.SMCNEWS
                            where f.TYPE == "02"
                            select f).FirstOrDefault();
                var model = new HomeIndexViewModel
                {
                    dataSMCDOCS = docs,
                };
                return View(model);
            }
            else
            {
                return View("/Views/Shared/Error.cshtml");
            }
        }

        /*public ActionResult DownloadFile(String id)
        {
            int id_num = int.Parse(id);
            var smcdocs = (from f in db.SMCDOCS
                           where f.ID == id_num
                           select f).FirstOrDefault();

            String UploadPath = System.Web.Configuration.WebConfigurationManager.AppSettings["uploadPath"];
            string fullPath = System.Web.HttpContext.Current.Server.MapPath("~/" + UploadPath) + System.IO.Path.GetFileName(smcdocs.FILENAME);
            byte[] fileBytes = System.IO.File.ReadAllBytes(fullPath);
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, smcdocs.DESP);
        }*/

        //===============================================================

        [AllowAnonymous]
        public ActionResult Logon()
        {
            //String onlinemode = ;
            prepareMainDeptViewInfo();
            ViewBag.Online = System.Web.Configuration.WebConfigurationManager.AppSettings["onlinemode"];
            return View();
        }

        protected string CryptographyPassword(string password, string salt)
        {
            string cryptographyPassword =
                FormsAuthentication.HashPasswordForStoringInConfigFile(password + salt, "sha1");

            return cryptographyPassword;
        }

        private void LoginProcess(Models.CMBASEModels.OGCUSER user, bool isRemeber)
        {
            var now = DateTime.Now;
            //string roles = string.Join(",", user.SystemRoles.Select(x => x.Name).ToArray());

            var ticket = new FormsAuthenticationTicket(
                version: 1,
                name: user.USR_ID.ToString(),
                issueDate: now,
                expiration: now.AddMinutes(30),
                isPersistent: isRemeber,
                userData: "",
                cookiePath: FormsAuthentication.FormsCookiePath);

            var encryptedTicket = FormsAuthentication.Encrypt(ticket);
            var cookie = new HttpCookie(FormsAuthentication.FormsCookieName, encryptedTicket);
            Response.Cookies.Add(cookie);

            // 資料權限放入Session
            //Session[Constants.AUTH_DEPT] = Helpers.WebSiteHelper.getAuthDepts(user.USR_ID.ToString());

            // 紀錄使用者登入Log
            Helpers.WebSiteHelper.logUserInfo(user.USR_ID.ToString());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [AllowAnonymous]
        public ActionResult Logon(LogonViewModel logonModel)
        {
            if (!ModelState.IsValid)
            {
                return View();
            }

            var systemuser = db.OGCUSER
                .FirstOrDefault(x => x.USR_ID == logonModel.Account);

            if (systemuser == null)
            {
                ModelState.AddModelError("", "請輸入正確的帳號或密碼!");
                return View();
            }
            //密碼加密的寫法
            var password = logonModel.Password;

            if (systemuser.USR_PSW.Equals(password))
            {
                this.LoginProcess(systemuser, logonModel.Remember);
                return RedirectToAction("Index", "Home");
            }
            else
            {
                ModelState.AddModelError("", "請輸入正確的帳號或密碼!");
                return View();
            }
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [AllowAnonymous]
        public ActionResult LogonAD(LogonViewModel logonModel)
        {
            prepareMainDeptViewInfo();

            if (!ModelState.IsValid)
            {
                return View("Logon");
            }





            var systemuser = db.OGCUSER
                        .FirstOrDefault(x => x.USR_ID == logonModel.Account);

            if (systemuser == null)
            {
                ModelState.AddModelError("", "請輸入正確的帳號或密碼!");
                return View("Logon");
            }
            //密碼加密的寫法
            var password = logonModel.Password;

            // 非上線模式不強制檢核密碼
            if (null == systemuser.USR_PSW || password.Equals(systemuser.USR_PSW))
            {
                LoginProcess(systemuser, logonModel.Remember);
                return RedirectToAction("Index", "Home");
            }
            else
            {
                ModelState.AddModelError("", "請輸入正確的帳號或密碼!");
                return View("Logon");
            }


















        }

        private void prepareLogonADViewInfo()
        {
            var ogcdept = (from f in db.OGCDEPT
                           where f.DATA_TP == "D"
                           orderby f.DEPT_NO
                           select f);
            ViewBag.OGCDEPT = new SelectList(ogcdept, "DEPT_NO", "DEPT_NAME");
        }

        private void prepareMainDeptViewInfo()
        {
            //var ogcscls = (from f in db.OGV_ADMDLST
            //                   //where f.DATA_TP == "D"
            //               orderby f.SC_NO
            //               select f);
            //ViewBag.OGCSCLS = new SelectList(ogcscls, "SC_NO", "DEPT_NAME");
        }


        [AllowAnonymous]
        public ActionResult LogonAD()
        {
            prepareLogonADViewInfo();
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [AllowAnonymous]
        public ActionResult ConnectAD(ADAccountViewModel model)
        {
            if (!ModelState.IsValid)
            {
                prepareLogonADViewInfo();
                return View("LogonAD", model);
            }

            // 有輸入舊帳號
            if (!string.IsNullOrEmpty(model.Account))
            {
                var OGCUSER = (from f in db.OGCUSER
                               where f.USR_ID == model.Account
                               select f).FirstOrDefault();

                var dpur = (from f in db.OGCDPUR
                            where f.USR_ID == model.Account && f.DATA_TP == "D"
                            select f).FirstOrDefault();

                if (OGCUSER != null)
                {
                    //OGCUSER.USR_STAT = "";
                    OGCUSER.UPD_DT = DateTime.Now;
                    //OGCUSER.AD_USR_ID = model.ADAccount;
                    OGCUSER.USR_NAME = model.UserName;
                    //OGCUSER.ADCONT = true;
                    OGCUSER.EMAIL = model.ADAccount + Constants.E_MAIL_DOMAIN_NAME;
                    using (var transaction = db.Database.BeginTransaction())
                    {
                        try
                        {
                            db.SaveChanges();
                            if (dpur != null)
                            {
                                if (dpur.DEPT_NO != model.DeptNo)
                                {
                                    DateTime addDtBk = (null == dpur.ADD_DT) ? DateTime.Now : dpur.ADD_DT.Value;
                                    db.OGCDPUR.Remove(dpur);
                                    //db.SaveChanges();

                                    Models.CMBASEModels.OGCDPUR ogcdpur = new Models.CMBASEModels.OGCDPUR();
                                    ogcdpur.DEPT_NO = model.DeptNo;
                                    ogcdpur.USR_ID = model.Account;
                                    ogcdpur.DATA_TP = "D";
                                    ogcdpur.ADD_DT = addDtBk;
                                    ogcdpur.UPD_DT = DateTime.Now;
                                    ogcdpur.UPD_UID = model.Account;
                                    db.OGCDPUR.Add(ogcdpur);
                                    db.SaveChanges();
                                }
                            }
                            else
                            {
                                Models.CMBASEModels.OGCDPUR ogcdpur = new Models.CMBASEModels.OGCDPUR();
                                ogcdpur.DEPT_NO = model.DeptNo;
                                ogcdpur.USR_ID = model.Account;
                                ogcdpur.DATA_TP = "D";
                                ogcdpur.ADD_DT = DateTime.Now;
                                ogcdpur.UPD_UID = model.Account;
                                db.OGCDPUR.Add(ogcdpur);
                                db.SaveChanges();
                            }
                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
            else
            {
                Models.CMBASEModels.OGCUSER ogcuser = new Models.CMBASEModels.OGCUSER();
                //ogcuser.USR_STAT = "";
                ogcuser.USR_NAME = model.UserName;
                ogcuser.ADD_DT = DateTime.Now;
                ogcuser.USR_ID = model.ADAccount;
                //ogcuser.AD_USR_ID = model.ADAccount;
                ogcuser.EMAIL = model.ADAccount + Constants.E_MAIL_DOMAIN_NAME;
                ogcuser.UPD_UID = model.ADAccount;

                using (var transaction = db.Database.BeginTransaction())
                {
                    db.OGCUSER.Add(ogcuser);
                    try
                    {
                        db.SaveChanges();
                        Models.CMBASEModels.OGCDPUR ogcdpur = new Models.CMBASEModels.OGCDPUR();
                        ogcdpur.DEPT_NO = model.DeptNo;
                        ogcdpur.USR_ID = model.ADAccount;
                        ogcdpur.DATA_TP = "D";
                        ogcdpur.ADD_DT = DateTime.Now;
                        ogcdpur.UPD_UID = model.ADAccount;
                        db.OGCDPUR.Add(ogcdpur);
                        db.SaveChanges();
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }

            return RedirectToAction("ConnectADOK", "Home");
        }

        [AllowAnonymous]
        public ActionResult ConnectADOK()
        {
            return View();
        }

        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();

            //清除所有的 session
            Session.RemoveAll();

            //建立一個同名的 Cookie 來覆蓋原本的 Cookie
            HttpCookie cookie1 = new HttpCookie(FormsAuthentication.FormsCookieName, "");
            cookie1.Expires = DateTime.Now.AddYears(-1);
            Response.Cookies.Add(cookie1);

            //建立 ASP.NET 的 Session Cookie 同樣是為了覆蓋
            HttpCookie cookie2 = new HttpCookie("ASP.NET_SessionId", "");
            cookie2.Expires = DateTime.Now.AddYears(-1);
            Response.Cookies.Add(cookie2);
            LogonViewModel m = new LogonViewModel();
            m.Account = "";
            m.Password = "";
            return RedirectToAction("Logon", "Home", m);
        }
    }

    //[Authorize] //之後要補上
    public class Sepmis_HomeController : Controller
    {
        CMBASEEntities db = new CMBASEEntities();
        ALL all = new ALL();
        int pageSize = 10;

        public ActionResult test()
        {
            return View();
        }

        public ActionResult Delete(string ID)
        {
            var user = db.OGCUSER.Where(m => m.USR_ID == ID).FirstOrDefault();
            db.OGCUSER.Remove(user);
            db.SaveChanges();

            var user2 = db.OGCORIGI.Where(m => m.USER_ID == ID).FirstOrDefault();
            db.OGCORIGI.Remove(user2);
            db.SaveChanges();

            return RedirectToAction("Employee", "Home");
        }

        public ActionResult DeptDelete(string ID, string DEPTID, string DATA)
        {
            var dept = db.OGCDPUR.Where(m => m.USR_ID == ID && m.DEPT_NO == DEPTID && m.DATA_TP == DATA).FirstOrDefault();
            db.OGCDPUR.Remove(dept);
            db.SaveChanges();
            return RedirectToAction("Edit", "Home", new { UserID = ID });
        }
        public ActionResult Edit(string UserID, string TestQuest, string DeptError)
        {
            ViewBag.Errorr = TestQuest;
            ViewBag.Error = DeptError;
            var User = db.OGCUSER
                .Where(m => m.USR_ID == UserID).FirstOrDefault();

            //string encryptedPassword = User.USR_PSW;
            //byte[] encryptionKey = Encoding.UTF8.GetBytes("cpc2k110cpc2k110cpc2k110");
            //string decryptedPassword = DecryptPassword(encryptedPassword, encryptionKey);
            //User.USR_PSW = decryptedPassword;
            List<Models.CMBASEModels.OGCDPUR> USER = db.OGCDPUR.Where(m => m.USR_ID == UserID).ToList();
            DeptGroupViewModel VM = new DeptGroupViewModel();
            VM.DEPTID = db.OGCDPUR.Where(m => m.USR_ID == (string)UserID).ToList();
            VM.DEPTNAME = db.OGCDEPT.ToList();
            var VMM = from s in VM.DEPTID
                      join st in VM.DEPTNAME on s.DEPT_NO equals st.DEPT_NO
                      select new DEPTGROUP_ViewModel { DEPTID = s, DEPTNAME = st };
            ViewBag.DEPTGROUP = VMM;
            return View(User);
        }
        [HttpPost]
        public ActionResult Edit(Models.CMBASEModels.OGCUSER User, string USR_PSW, string USR_PSW2)
        {
            string TestQuest = "";
            var temp = db.OGCUSER.Where(m => m.USR_ID == User.USR_ID).FirstOrDefault();
            string UserIDD = temp.USR_ID;
            if (USR_PSW2 != USR_PSW)
            {
                TestQuest = "確認密碼輸入錯誤";
                return RedirectToAction("Edit", "Home", new { UserID = UserIDD, TestQuest = TestQuest });
            }
            if (ModelState.IsValid)
            {
                //var original = db.OGCORIGI.Where(m => m.USER_ID == User.USR_ID).FirstOrDefault();
                //original.ORIGINAL_PSW = USR_PSW;
                //db.SaveChanges();


                //string originalPassword = USR_PSW;
                //byte[] encryptionKey = Encoding.UTF8.GetBytes("cpc2k110cpc2k110cpc2k110");   /////密碼至少需要16字元
                //string encryptedPassword = EncryptPassword(originalPassword, encryptionKey);

                temp.USR_NAME = User.USR_NAME;
                temp.EMAIL = User.EMAIL;
                temp.TEL_O = User.TEL_O;
                temp.TEL_M = User.TEL_M;
                temp.USR_STAT = User.USR_STAT;
                temp.NOTE = User.NOTE;
                temp.USR_PSW = USR_PSW/*encryptedPassword*/;
                temp.UPD_DT = DateTime.Now;
                temp.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                //之後要加部門群組的部分

                try
                {
                    db.SaveChanges();
                    TestQuest = "修改成功";
                }
                catch (Exception)
                {
                    TestQuest = "修改失敗";
                }
                return RedirectToAction("Edit", "Home", new { UserID = UserIDD, TestQuest = TestQuest });
            }
            return View(User);
        }
        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Create(Models.CMBASEModels.OGCUSER User, string USR_PSW, string USR_PSW2)
        {
            //string originalPassword = USR_PSW;
            //byte[] encryptionKey = Encoding.UTF8.GetBytes("cpc2k110cpc2k110cpc2k110");   /////密碼至少需要16字元
            //string encryptedPassword = EncryptPassword(originalPassword, encryptionKey);


            if (USR_PSW2 != USR_PSW)
            {
                ViewBag.IDERROR = "確認密碼輸入錯誤";
                return View();
            }
            try
            {
                User.USR_PSW = USR_PSW/*encryptedPassword*/;
                User.ADD_DT = DateTime.Now;
                User.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                db.OGCUSER.Add(User);
                db.SaveChanges();
                OGCORIGI Original = new OGCORIGI();
                Original.USER_ID = User.USR_ID;
                Original.ORIGINAL_PSW = USR_PSW;
                db.OGCORIGI.Add(Original);
                db.SaveChanges();
                return RedirectToAction("Employee", "Home");
            }
            catch
            {
                ViewBag.IDERROR = "此員工編號已註冊過";
            }
            return View();
        }

        public ActionResult Download()
        {
            // 資料夾路徑，這裡假設您要下載的資料夾在特定路徑下
            string Path = Server.MapPath("~/Download/111-114年度下水道工程建設計畫資訊整合應用網維運及擴改_教育訓練.pdf");
            // 提供下載
            byte[] fileBytes = System.IO.File.ReadAllBytes(Path);
            return File(fileBytes, "application/pdf", "111-114年度下水道工程建設計畫資訊整合應用網維運及擴改_教育訓練.pdf");
        }

        public ActionResult CreateDeptGroup(string USR_ID, string State, string Value)
        {
            Models.CMBASEModels.OGCDPUR User = new Models.CMBASEModels.OGCDPUR();
            User.USR_ID = USR_ID;
            string DeptError = "";
            if (State == "部門")
            {
                User.DATA_TP = "D";
                List<Models.CMBASEModels.OGCDPUR> U = db.OGCDPUR.Where(m => m.USR_ID == User.USR_ID).ToList();
                List<Models.CMBASEModels.OGCDPUR> M = U.Where(m => m.DATA_TP == "D").ToList();
                if (M.Count() >= 1)
                {
                    DeptError = "部門只能新增一個";
                    return RedirectToAction("Edit", "Home", new { UserID = USR_ID, DeptError = DeptError });
                }

            }
            if (State == "群組")
            {
                User.DATA_TP = "G";
            }
            User.DEPT_NO = db.OGCDEPT.Where(m => m.DEPT_NAME == Value).FirstOrDefault().DEPT_NO;
            try
            {
                User.ADD_DT = DateTime.Now;
                User.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                db.OGCDPUR.Add(User);
                db.SaveChanges();
                DeptError = "";
            }
            catch (Exception) { DeptError = "資料重複，無法新增"; }

            return RedirectToAction("Edit", "Home", new { UserID = USR_ID, DeptError = DeptError });
        }

        public ActionResult Cancel()
        {
            TempData["Station"] = null;
            return RedirectToAction("Employee", "Home");
        }
        public ActionResult Select()
        {
            return RedirectToAction("Employee", "Home");
        }
        public ActionResult Choose(string select)
        {
            var OGCDPUR = db.OGCDPUR.Where(m => m.USR_ID == select).ToList();
            var OGCUSER = db.OGCUSER.Where(m => m.USR_ID == select).ToList();
            var OGCDEPT = db.OGCDEPT.ToList();
            var Query = from a in OGCDPUR
                        join b in OGCUSER on a.USR_ID equals b.USR_ID
                        join c in OGCDEPT on a.DEPT_NO equals c.DEPT_NO
                        select new ChooseViewModel { OGCUSER = b, OGCDPUR = a, OGCDEPT = c };
            //ViewBag.DEPT = TempData["DEPT"];
            //ViewBag.Group = TempData["Group"];
            return View(Query);
        }
        [HttpGet]
        public ActionResult Employee(string year, string chooseSelect, string USR_ID_search, string USR_NAME_search, int page = 1)
        {
            //Select 選單 table

            string M = "外單位";
            string N = "副主任室";
            ViewBag.ChooseSelect = db.OGCDEPT.Where(m => m.DEPT_NAME != M && m.DEPT_NAME != N).ToList();
            int currentPage = page < 1 ? 1 : page;
            var employee = db.OGCUSER.OrderBy(m => m.USR_ID).ToList();
            if (USR_ID_search != null)
            {
                //employee = db.OGCUSER.OrderBy(m => m.USR_ID).Where(s => s.USR_ID.Contains(USR_ID_search)).ToList();
                employee = employee.OrderBy(m => m.USR_ID).Where(s => s.USR_ID.Contains(USR_ID_search)).ToList();

            }
            if (USR_NAME_search != null)
            {
                employee = employee.OrderBy(m => m.USR_ID).Where(s => s.USR_NAME.Contains(USR_NAME_search)).ToList();
            }
            var result = employee.ToPagedList(currentPage, pageSize);
            ViewBag.Group = null;
            if (chooseSelect != null)
            {
                TempData["Station"] = "1";
                var choose_btn = db.OGCUSER.Where(m => m.USR_ID == chooseSelect).FirstOrDefault();
                TempData["USERID"] = choose_btn.USR_ID;
                TempData["USERNAME"] = choose_btn.USR_NAME;
                TempData["TEL_O"] = choose_btn.TEL_O;
                TempData["TEL_M"] = choose_btn.TEL_M;
                TempData["MAIL"] = choose_btn.EMAIL;
                TempData["USERSTAT"] = choose_btn.USR_STAT;
                TempData["NOTE"] = choose_btn.NOTE;
                var dept_group = db.OGCDPUR.Where(m => m.USR_ID == chooseSelect).ToList();
                var DEPT = dept_group.Where(m => m.DATA_TP == "D").FirstOrDefault();
                try
                {
                    var DEPT_List = db.OGCDEPT.Where(m => m.DEPT_NO == DEPT.DEPT_NO).FirstOrDefault();
                    TempData["DEPT"] = DEPT_List.DEPT_NAME;
                    var GROUP = dept_group.Where(m => m.DATA_TP == "G").ToList();
                    List<Models.CMBASEModels.OGCDEPT> GROUPLIST = new List<Models.CMBASEModels.OGCDEPT>();
                    TempData["Group"] = "";
                    foreach (var item in GROUP)
                    {
                        var GROUP_element = db.OGCDEPT.Where(m => m.DEPT_NO == item.DEPT_NO).FirstOrDefault();
                        GROUPLIST.Add(GROUP_element);
                        TempData["Group"] = GROUPLIST;
                    }
                }
                catch (Exception e) { Console.WriteLine("{0} Exception caught.", e); }
                return RedirectToAction("Choose", "Home", new { select = chooseSelect });
            }
            return View(result);
        }

        [HttpPost]
        public JsonResult Orders(string State)
        {
            List<KeyValuePair<string, string>> items = new List<KeyValuePair<string, string>>();

            if (!string.IsNullOrWhiteSpace(State))
            {
                var orders = this.GetOrders(State);
                if (orders.Count() > 0)
                {
                    foreach (var order in orders)
                    {
                        items.Add(new KeyValuePair<string, string>(
                            order.DEPT_NAME.ToString(),
                            string.Format("{0}", order.DEPT_NAME)));
                    }
                }
            }
            return Json(items);
        }

        private IEnumerable<Models.CMBASEModels.OGCDEPT> GetOrders(string State)
        {
            using (CMBASEEntities db = new CMBASEEntities())
            {
                //var query = db.OGCDPUR.Where(x => x.USR_ID == customerID).OrderBy(x => x.DEPT_NO);
                string M = "外單位";
                string N = "副主任室";
                if (State == "部門")
                {
                    var query = db.OGCDEPT.Where(m => m.DEPT_NAME != M && m.DEPT_NAME != N);

                    return query.ToList();
                }
                else
                {
                    var query = db.OGCDEPT.Where(m => m.DEPT_NAME != M && m.DEPT_NAME != N && !m.DEPT_NO.Contains("A27"));
                    return query.ToList();
                }

            }
        }
        // GET: Home
        [Authorize]
        public ActionResult Index()
        {
            MenuViewModel Side = new MenuViewModel();
            Side.Menu1 = db.SMCMENU3.Where(m => m.DATA_TP == "R").OrderBy(m => m.SYS_ORD).ToList();
            Side.Menu2 = db.SMCMENU3.Where(m => m.DATA_TP == "N").OrderBy(m => m.SYS_ORD).ToList();
            Side.Menu3 = db.SMCMENU3.Where(m => m.DATA_TP == "A").OrderBy(m => m.SYS_ORD).ToList();
            return View(Side);

        }

        private IEnumerable<Models.CMBASEModels.OGCUSER> GetCustomers()
        {
            using (CMBASEEntities db = new CMBASEEntities())
            {
                var query = db.OGCUSER.OrderBy(x => x.USR_ID);
                return query.ToList();
            }
        }


        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();   // 登出
            GetValidateCode();
            return RedirectToAction("Login", GetValidateCode());
        }

        [AllowAnonymous]
        public ActionResult Login()
        {
            return View();
        }

        [AllowAnonymous]
        [HttpPost]
        public ActionResult Login(string txtUid, string txtPwd, string chptcha)
        {
            TempData["User"] = txtUid;
            string originalPassword = txtPwd;
            //byte[] encryptionKey = Encoding.UTF8.GetBytes("cpc2k110cpc2k110cpc2k110");   /////密碼至少需要16字元
            //string encryptedPassword = EncryptPassword(originalPassword, encryptionKey);
            var member = db.OGCUSER.Where(m => m.USR_ID == txtUid).FirstOrDefault();


            // 循序搜尋法
            if (member.FALL_TIME == 5)
            {
                ViewBag.Err = "帳號已被鎖定";
                return View();
            }

            member = db.OGCUSER
            .Where(m => m.USR_ID == txtUid && m.USR_PSW == /*encryptedPassword*/originalPassword)
            .FirstOrDefault();

            if (member == null)
            {
                member = db.OGCUSER.Where(m => m.USR_ID == txtUid).FirstOrDefault();
                if (member.FALL_TIME == null)
                {
                    member.FALL_TIME = 1;
                    db.SaveChanges();
                }
                else
                {
                    member.FALL_TIME += 1;
                    db.SaveChanges();
                }
                ViewBag.Err = "帳密錯誤，登入失敗" + member.FALL_TIME + "次" + "，還能嘗試" + (5 - member.FALL_TIME) + "次";
                return View();
            }
            //if (chptcha != TempData["code"].ToString())
            //{
            //    ViewBag.Err = "驗證碼錯誤，登入錯誤";
            //    return View();
            //}
            var Dept_ID = db.OGCDPUR.Where(m => m.USR_ID == member.USR_ID).FirstOrDefault();
            var Dept_Name = db.OGCDEPT.Where(m => m.DEPT_NO == Dept_ID.DEPT_NO).FirstOrDefault();
            Session["UsrName"] = member.USR_NAME;
            Session["UsrId"] = member.USR_ID;
            Session["Dept"] = Dept_Name.DEPT_NAME;
            Session["DeptId"] = Dept_Name.DEPT_NO;
            FormsAuthentication.RedirectFromLoginPage(txtUid, true);
            member.FALL_TIME = null;
            db.SaveChanges();
            //return RedirectToAction("Index", "Member"); 
            return RedirectToAction("SideBar", "Home");
        }
        [Authorize]
        public ActionResult SideBar()
        {
            MenuViewModel Side = new MenuViewModel();
            String user = (string)TempData["User"];
            TempData.Keep();
            Side.DEPT = db.OGCDPUR.Where(m => m.USR_ID == user).ToList();
            Side.FUNCTION = db.SMCAUTH.ToList();
            Side.Menu = db.SMCMENU3.ToList();

            //List<SMCMENU> MENU= new List<SMCMENU>();

            var vm = from s in Side.DEPT
                     join st in Side.FUNCTION on s.DEPT_NO equals st.DEPT_NO
                     select new DEPT_NO_ViewModel { DEPTVM = s, FUNCTIONVM = st };
            vm = vm.DistinctBy(m => m.FUNCTIONVM.SYS_ID).ToList();

            var vmm = from s in vm
                      join st in Side.Menu on s.FUNCTIONVM.SYS_ID equals st.SYS_ID
                      select new DEPT_NO_ViewModel { MENUVM = st };

            ALL all = new ALL();
            all.Menu1 = db.SMCMENU3.Where(m => m.DATA_TP == "R").OrderBy(m => m.SYS_ORD).ToList();
            all.Menu2 = db.SMCMENU3.Where(m => m.DATA_TP == "N").OrderBy(m => m.SYS_ORD).ToList();
            all.Menu3 = db.SMCMENU3.Where(m => m.DATA_TP == "A").OrderBy(m => m.SYS_ORD).ToList();
            all.Menu4 = vmm.Where(m => m.MENUVM.DATA_TP == "A").OrderBy(m => m.MENUVM.SYS_ORD).ToList();
            all.Customer = db.OGCUSER.ToList();
            return View(all);
        }
        public ActionResult AuthorizeUser()
        {
            MenuViewModel Side = new MenuViewModel();
            Side.Menu1 = db.SMCMENU3.Where(m => m.DATA_TP == "R").OrderBy(m => m.SYS_ORD).ToList();
            Side.Menu2 = db.SMCMENU3.Where(m => m.DATA_TP == "N").OrderBy(m => m.SYS_ORD).ToList();
            Side.Menu3 = db.SMCMENU3.Where(m => m.DATA_TP == "A").OrderBy(m => m.SYS_ORD).ToList();
            return View(Side);
        }


        public ActionResult GetValidateCode()
        {
            ValidateCode vCode = new ValidateCode();
            string code = vCode.CreateValidateCode(5);
            TempData["code"] = code;
            byte[] bytes = vCode.CreateValidateGraphic(code);
            return File(bytes, @"image/jpeg");
        }
        public class ValidateCode
        {
            /// <summary>
            /// 生成驗證碼
            /// </summary>
            /// <param name="length">指定驗證碼的長度</param>
            /// <returns></returns>
            public string CreateValidateCode(int length)
            {
                const string validateTemplate = "23456789abcdefghijkmnopqrstuvwxyzABCDEFGHIJKLMNPQRSTUVWXYZ";
                string validateStr = "";
                int n = validateTemplate.Length - 1;

                //設置隨機值生成器
                int seekSeek = unchecked((int)DateTime.Now.Ticks);
                Random seekRand = new Random(seekSeek);
                int beginSeek = 0;

                //生成隨機驗證碼
                for (int i = 0; i < length; i++)
                {
                    beginSeek = seekRand.Next(0, n);
                    validateStr += validateTemplate.Substring(beginSeek, 1);
                }
                return validateStr;
            }

            /// <summary>
            /// 創建驗證碼的圖片
            /// </summary>
            /// <param name="containsPage">要輸出到的page對象</param>
            /// <param name="validateNum">驗證碼</param>
            public byte[] CreateValidateGraphic(string validateCode)
            {
                Bitmap image = new Bitmap((int)Math.Ceiling(validateCode.Length * 13.0), 22);
                Graphics g = Graphics.FromImage(image);
                try
                {
                    //生成隨機生成器
                    Random random = new Random();
                    //清空圖片背景色
                    g.Clear(Color.White);
                    //畫圖片的干擾線
                    for (int i = 0; i < 25; i++)
                    {
                        int x1 = random.Next(image.Width);
                        int x2 = random.Next(image.Width);
                        int y1 = random.Next(image.Height);
                        int y2 = random.Next(image.Height);
                        g.DrawLine(new Pen(Color.Silver), x1, y1, x2, y2);
                    }
                    Font font = new Font("Arial", 12, (FontStyle.Bold | FontStyle.Italic));
                    LinearGradientBrush brush = new LinearGradientBrush(new Rectangle(0, 0, image.Width, image.Height),
                     Color.Blue, Color.DarkRed, 1.2f, true);
                    g.DrawString(validateCode, font, brush, 3, 2);
                    //畫圖片的前景干擾點
                    for (int i = 0; i < 100; i++)
                    {
                        int x = random.Next(image.Width);
                        int y = random.Next(image.Height);
                        image.SetPixel(x, y, Color.FromArgb(random.Next()));
                    }
                    //畫圖片的邊框線
                    g.DrawRectangle(new Pen(Color.Silver), 0, 0, image.Width - 1, image.Height - 1);
                    //保存圖片數據
                    MemoryStream stream = new MemoryStream();
                    image.Save(stream, ImageFormat.Jpeg);
                    //輸出圖片流
                    return stream.ToArray();
                }
                finally
                {
                    g.Dispose();
                    image.Dispose();
                }
            }
        }


        public string EncryptPassword(string password, byte[] key)
        {
            using (AesManaged aes = new AesManaged())
            {
                aes.Key = key;
                aes.Mode = CipherMode.ECB;
                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);

                byte[] bytes = Encoding.UTF8.GetBytes(password);
                byte[] encryptedbytes = encryptor.TransformFinalBlock(bytes, 0, bytes.Length);

                return Convert.ToBase64String(encryptedbytes);
            }
        }

        public string DecryptPassword(string encryptedPassword, byte[] key)
        {
            using (AesManaged aes = new AesManaged())
            {
                aes.KeySize = key.Length * 8;  // 根據金鑰的位元組數設置金鑰大小
                aes.Key = key;
                aes.Mode = CipherMode.ECB;
                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);

                byte[] encryptedbytes = Convert.FromBase64String(encryptedPassword);
                byte[] decryptedbytes = decryptor.TransformFinalBlock(encryptedbytes, 0, encryptedbytes.Length);

                return Encoding.UTF8.GetString(decryptedbytes);
            }
        }
    }
}