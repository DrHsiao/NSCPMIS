using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data.Entity.Migrations;
using System.Linq;
using System.Web.Mvc;
using WCals.Models.CMBASEModels;

namespace WCals.Controllers
{
    public class AuthController : Controller
    {
        CMBASEEntities db = new CMBASEEntities();
        // GET: Auth
        public ViewResult Index()
        {
            AuthViewModel avm = new AuthViewModel();
            avm.menu1 = db.SMCMENU3.Where(m => m.DATA_TP == "R").OrderBy(m => m.SYS_ORD).ToList();
            avm.menu2 = db.SMCMENU3.Where(m => m.DATA_TP == "N").OrderBy(m => m.SYS_ORD).ToList();
            avm.menu3 = db.SMCMENU3.Where(m => m.DATA_TP == "A").OrderBy(m => m.SYS_ORD).ToList();
            return View(avm);

        }
        [HttpPost]
        public ContentResult Index(string sysid)
        {
            var auth = db.SMCMENU3.Join(db.SMCAUTH, m => m.SYS_ID, n => n.SYS_ID, (du, d) =>
            new { d.DEPT_NO, du.SYS_ID, d.CUS_FG }).Where(ddu => ddu.SYS_ID == sysid).
            Join(db.OGCDEPT, m => m.DEPT_NO, n => n.DEPT_NO, (ddu, u) =>
            new { DEPT_NO = ddu.DEPT_NO, DEPT_NAME = u.DEPT_NAME, SYS_ID = ddu.SYS_ID, CUS_FG = ddu.CUS_FG }).ToList();


            string json = JsonConvert.SerializeObject(auth);

            return Content(json, "application/json");
        }
        public ActionResult Delete(string DEPT_NO, string SYS_ID)
        {
            var auth = db.SMCAUTH.Where(m => m.DEPT_NO == DEPT_NO && m.SYS_ID == SYS_ID).FirstOrDefault();
            db.SMCAUTH.Remove(auth);
            db.SaveChanges();

            return RedirectToAction("Index");
        }
        public ActionResult Create(string SYS_ID)
        {

            List<AddViewModel> avm1 = new List<AddViewModel>();
            var auth = db.SMCAUTH.Join(db.OGCDEPT, m => m.DEPT_NO, n => n.DEPT_NO, (d, u) =>
            new { d.SYS_ID, d.CUS_FG, u.DEPT_NO, u.DEPT_NAME, d.DATA_TP }).Where(m => m.SYS_ID == SYS_ID).ToList();

            var excludeDeptNos = auth.Select(a => a.DEPT_NO).Distinct().ToList();
            var auth1 = db.OGCDEPT.Except(db.OGCDEPT.Where(o => excludeDeptNos.Contains(o.DEPT_NO))).ToList();
            int i = 0;
            for (i = 0; i < auth1.Count; i++)
            {
                AddViewModel avm = new AddViewModel();
                avm.Index = i;
                avm.SelectedAdd = false;
                avm.DATA_TP = "D";
                avm.CUS_FG = false;
                avm.DEPT_NO = auth1[i].DEPT_NO;
                avm.DEPT_NAME = auth1[i].DEPT_NAME;
                avm.SYS_ID = SYS_ID;
                avm1.Add(avm);
            }
            for (int j = i; j < auth.Count + i; j++)
            {
                AddViewModel avm = new AddViewModel();
                avm.Index = j;
                avm.SelectedAdd = true;
                avm.DATA_TP = "D";
                avm.CUS_FG = auth[j - i].CUS_FG;
                avm.DEPT_NO = auth[j - i].DEPT_NO;
                avm.DEPT_NAME = auth[j - i].DEPT_NAME;
                avm.SYS_ID = SYS_ID;
                avm1.Add(avm);
            }
            return View(avm1);
        }
        [HttpPost]
        public ActionResult Create(List<AddViewModel> avm)
        {
            foreach (var item in avm)
            {
                var auth1 = db.SMCAUTH.Where(m => m.DEPT_NO == item.DEPT_NO && m.SYS_ID == item.SYS_ID).FirstOrDefault();

                if (item.SelectedAdd)
                {
                    // Add new item to database
                    //有資料且全部都一樣不動作
                    if (auth1 != null && auth1.CUS_FG == item.CUS_FG) { }
                    //無資料或修改過新增或更新
                    else
                    {
                        SMCAUTH auth = new SMCAUTH()
                        {
                            DEPT_NO = item.DEPT_NO,
                            SYS_ID = item.SYS_ID,
                            DATA_TP = "D",
                            CUS_FG = item.CUS_FG,
                        };
                        db.SMCAUTH.AddOrUpdate(auth);
                    }
                }
                else
                {
                    // Remove item from database
                    var auth = db.SMCAUTH.FirstOrDefault(a => a.SYS_ID == item.SYS_ID && a.DEPT_NO == item.DEPT_NO);
                    if (auth != null)
                    {
                        db.SMCAUTH.Remove(auth);
                    }
                }

            }

            db.SaveChanges();

            return RedirectToAction("Index");
        }
        //[HttpPost]
        //public ActionResult Create(AddViewModel avm)
        //{

        //    //foreach (var item in avm)
        //    //{
        //    if (avm.SelectedAdd)
        //    {
        //        // Add new item to database
        //        SMCAUTH auth = new SMCAUTH()
        //        {
        //            DEPT_NO = avm.DEPT_NO,
        //            SYS_ID = avm.SYS_ID,
        //            DATA_TP = "D",
        //            CUS_FG = avm.CUS_FG,
        //        };
        //        db.SMCAUTH.Add(auth);
        //    }
        //    else
        //    {
        //        // Remove item from database
        //        var auth = db.SMCAUTH.FirstOrDefault(a => a.SYS_ID == avm.SYS_ID && a.DEPT_NO == avm.DEPT_NO);
        //        if (auth != null)
        //        {
        //            db.SMCAUTH.Remove(auth);
        //        }
        //    }
        //    //}

        //    db.SaveChanges();

        //    return RedirectToAction("Index");
        //}
    }
}