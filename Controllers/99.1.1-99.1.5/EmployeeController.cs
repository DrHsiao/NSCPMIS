using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using WCals.Models.CMBASEModels;

namespace WCals.Controllers
{
    public class EmployeeController : Controller
    {
        CMBASEEntities db = new CMBASEEntities();
        ALL all = new ALL();
        int pageSize = 10;

        public ActionResult Delete(string ID)
        {
            var user = db.OGCUSER.Where(m => m.USR_ID == ID).FirstOrDefault();
            db.OGCUSER.Remove(user);
            db.SaveChanges();

            var user2 = db.OGCORIGI.Where(m => m.USER_ID == ID).FirstOrDefault();
            db.OGCORIGI.Remove(user2);
            db.SaveChanges();

            return RedirectToAction("Employee", "Employee");
        }

        public ActionResult DeptDelete(string ID, string DEPTID, string DATA)
        {
            var dept = db.OGCDPUR.Where(m => m.USR_ID == ID && m.DEPT_NO == DEPTID && m.DATA_TP == DATA).FirstOrDefault();
            db.OGCDPUR.Remove(dept);
            db.SaveChanges();
            return RedirectToAction("Edit", "Employee", new { UserID = ID });
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
            List<OGCDPUR> USER = db.OGCDPUR.Where(m => m.USR_ID == UserID).ToList();
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
        public ActionResult Edit(OGCUSER User, string USR_PSW, string USR_PSW2)
        {
            string TestQuest = "";
            var temp = db.OGCUSER.Where(m => m.USR_ID == User.USR_ID).FirstOrDefault();
            string UserIDD = temp.USR_ID;
            if (USR_PSW2 != USR_PSW)
            {
                TestQuest = "確認密碼輸入錯誤";
                return RedirectToAction("Edit", "Employee", new { UserID = UserIDD, TestQuest = TestQuest });
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
                return RedirectToAction("Edit", "Employee", new { UserID = UserIDD, TestQuest = TestQuest });
            }
            return View(User);
        }

        public ActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Create(OGCUSER User, string USR_PSW, string USR_PSW2)
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
                return RedirectToAction("Employee", "Employee");
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
            OGCDPUR User = new OGCDPUR();
            User.USR_ID = USR_ID;
            string DeptError = "";
            if (State == "部門")
            {
                User.DATA_TP = "D";
                List<OGCDPUR> U = db.OGCDPUR.Where(m => m.USR_ID == User.USR_ID).ToList();
                List<OGCDPUR> M = U.Where(m => m.DATA_TP == "D").ToList();
                if (M.Count() >= 1)
                {
                    DeptError = "部門只能新增一個";
                    return RedirectToAction("Edit", "Employee", new { UserID = USR_ID, DeptError = DeptError });
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

            return RedirectToAction("Edit", "Employee", new { UserID = USR_ID, DeptError = DeptError });
        }

        public ActionResult Cancel()
        {
            TempData["Station"] = null;
            return RedirectToAction("Employee", "Employee");
        }

        public ActionResult Select()
        {
            return RedirectToAction("Employee", "Employee");
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
                    List<OGCDEPT> GROUPLIST = new List<OGCDEPT>();
                    TempData["Group"] = "";
                    foreach (var item in GROUP)
                    {
                        var GROUP_element = db.OGCDEPT.Where(m => m.DEPT_NO == item.DEPT_NO).FirstOrDefault();
                        GROUPLIST.Add(GROUP_element);
                        TempData["Group"] = GROUPLIST;
                    }
                }
                catch (Exception e) { Console.WriteLine("{0} Exception caught.", e); }
                return RedirectToAction("Choose", "Employee", new { select = chooseSelect });
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

        private IEnumerable<OGCDEPT> GetOrders(string State)
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
    }
}