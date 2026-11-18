using System;
using System.Linq;
using System.Web.Mvc;
using WCals.Models.CMBASEModels;

namespace WCals.Controllers
{
    public class GroupController : Controller
    {
        CMBASEEntities db = new CMBASEEntities();
        // GET: Group
        public ActionResult Index()
        {
            var group = db.OGCDEPT.OrderBy(m => m.DEPT_NO).Where(m => m.DATA_TP == "G").ToList();
            return View(group);
        }

        public ActionResult choose(string id)
        {
            TempData["DEPTNO"] = id;
            GroupChooseViewModel ViewModel = new GroupChooseViewModel();
            ViewModel.ChooseDEPT = db.OGCDEPT.Where(m => m.DEPT_NO == id).FirstOrDefault();   ////按下選取      擁有此編號的名稱
            ViewModel.ChooseDPUR = db.OGCDPUR.Where(m => m.DEPT_NO == id).ToList();    ////表單內顯示的畫面   有選取此群組的成員
            ViewModel.USER = db.OGCUSER.ToList();
            var VM = from s in ViewModel.ChooseDPUR
                     join st in ViewModel.USER on s.USR_ID equals st.USR_ID
                     select new GroupUserViewModel { DPUR = s, USER = st };
            ViewBag.GROUPUSER = VM;
            return View(ViewModel);
        }

        public ActionResult Edit(string DEPT_NO, string GroupError, string NameError)
        {
            ViewBag.GroupError = GroupError;
            ViewBag.NameError = NameError;
            var Group = db.OGCDEPT.Where(m => m.DEPT_NO == DEPT_NO).FirstOrDefault();
            ViewBag.DEPT_NO = Group.DEPT_NO;
            ViewBag.DEPT_NAME = Group.DEPT_NAME;
            GroupChooseViewModel ViewModel = new GroupChooseViewModel();
            ViewModel.ChooseDEPT = db.OGCDEPT.Where(m => m.DEPT_NO == DEPT_NO).FirstOrDefault();   ////按下選取      擁有此編號的名稱
            ViewModel.ChooseDPUR = db.OGCDPUR.Where(m => m.DEPT_NO == DEPT_NO).ToList();    ////表單內顯示的畫面   有選取此群組的成員
            ViewModel.USER = db.OGCUSER.ToList();
            var VM = from s in ViewModel.ChooseDPUR
                     join st in ViewModel.USER on s.USR_ID equals st.USR_ID
                     select new GroupUserViewModel { DPUR = s, USER = st };
            ViewBag.GROUPUSER = VM;
            var result = from a in db.OGCUSER
                         join b in db.OGCDPUR on a.USR_ID equals b.USR_ID into temp
                         from b in temp.DefaultIfEmpty()
                         where b.DEPT_NO.StartsWith("A27")
                         select new GroupUserViewModel { USER = a, DPUR = b };
            ViewBag.USER = result;

            return View(ViewModel);
        }
        [HttpPost]
        public ActionResult Edit(string DEPT_NO, string DEPT_NAME)
        {
            string NameError = "";
            var temp = db.OGCDEPT.Where(m => m.DEPT_NO == DEPT_NO).FirstOrDefault();
            temp.DEPT_NAME = DEPT_NAME;
            temp.UPD_DT = DateTime.Now;
            temp.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            try
            {
                db.SaveChanges();
                NameError = "修改成功";
            }
            catch (Exception) { NameError = "修改失敗"; }
            return RedirectToAction("Edit", "Group", new { DEPT_NO = DEPT_NO, NameError = NameError });
        }

        public ActionResult Create(string NameERRORR)
        {
            ViewBag.NameERRORR = NameERRORR;
            return View();
        }
        [HttpPost]
        public ActionResult Create(OGCDEPT New)
        {
            string NameERRORR = "";
            New.DATA_TP = "G";
            New.ADD_DT = DateTime.Now;
            New.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            try
            {
                db.OGCDEPT.Add(New);
                db.SaveChanges();
                NameERRORR = "";
            }
            catch (Exception)
            {
                NameERRORR = "資料重複，無法新增";
                return RedirectToAction("Create", "Group", new { NameERRORR = NameERRORR });
            }
            return RedirectToAction("Index", "Group");
        }
        [HttpPost]
        public ActionResult CreateUser(string DEPT_NO, string choose)
        {
            string GroupError = "";
            int dotIndex = choose.IndexOf('.');
            string result = choose.Substring(0, dotIndex);
            OGCDPUR GROUP_USER = new OGCDPUR();
            GROUP_USER.DEPT_NO = DEPT_NO;
            GROUP_USER.USR_ID = result;
            GROUP_USER.DATA_TP = "G";
            GROUP_USER.ADD_DT = DateTime.Now;
            GROUP_USER.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            try
            {
                db.OGCDPUR.Add(GROUP_USER);
                db.SaveChanges();
                GroupError = "";
            }
            catch (Exception) { GroupError = "資料重複，無法新增"; }
            return RedirectToAction("Edit", "Group", new { DEPT_NO = DEPT_NO, GroupError = GroupError });
        }

        public ActionResult Delete(string DEPT_NO)
        {
            var Group = db.OGCDEPT.Where(m => m.DEPT_NO == DEPT_NO).FirstOrDefault();
            db.OGCDEPT.Remove(Group);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        public ActionResult GroupDelete(string DEPT_ID, string USR_ID)
        {
            var Group = db.OGCDPUR.Where(m => m.DEPT_NO == DEPT_ID && m.USR_ID == USR_ID).FirstOrDefault();
            db.OGCDPUR.Remove(Group);
            db.SaveChanges();
            return RedirectToAction("Edit", "Group", new { DEPT_NO = DEPT_ID });
        }
    }
}