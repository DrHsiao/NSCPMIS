using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web.Mvc;
using WCals.Models.CMBASEModels;

namespace WCals.Controllers
{
    public class DeptController : Controller
    {
        CMBASEEntities db = new CMBASEEntities();

        // GET: Dept
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult TreeVieww()
        {
            DeptViewModel dvm = new DeptViewModel();
            dvm.Dept1 = db.OGCDEPT.Where(m => m.DATA_TP == "D" && m.AN_DEPT_NO == null && m.DEPT_NO != "A040000").OrderBy(m => m.DEPT_NAME).ToList();
            dvm.Dept2 = db.OGCDEPT.Where(m => m.DATA_TP == "D").OrderBy(m => m.DEPT_NO).ToList();

            return View(dvm);
        }
        [HttpPost]
        public ContentResult TreeVieww(string deptno)
        {
            //dvm.Dept1 = db.OGCDEPT.Where(m => m.DATA_TP == "D" && m.AN_DEPT_NO == null && m.DEPT_NO != "A040000").OrderBy(m => m.DEPT_NAME).ToList();
            //dvm.Dept2 = db.OGCDEPT.Where(m => m.DATA_TP == "D").OrderBy(m => m.DEPT_NO).ToList();
            //dvm.DeptUser = db.OGCDPUR.Join(db.OGCDEPT, m => m.DEPT_NO, n => n.DEPT_NO, (du, d) =>
            //new { d.DEPT_NO, d.DEPT_NAME, d.MGR_FG, d.DEPT_CHIEF, du.USR_ID }).
            //    Where(ddu => ddu.DEPT_NO == deptno).
            //    Join(db.OGCUSER, m => m.USR_ID, n => n.USR_ID, (ddu, u) =>
            //    new DeptUserViewModel() { DEPTNO = ddu.DEPT_NO,DEPTNAME=ddu.DEPT_NAME, USRID = ddu.USR_ID, USRNAME = u.USR_NAME }).ToList();

            var dept = db.OGCDPUR.Join(db.OGCDEPT, m => m.DEPT_NO, n => n.DEPT_NO, (du, d) =>
            new { d.DEPT_NO, d.DEPT_NAME, d.MGR_FG, d.DEPT_CHIEF, du.USR_ID }).
                Where(ddu => ddu.DEPT_NO == deptno).
                Join(db.OGCUSER, m => m.USR_ID, n => n.USR_ID, (ddu, u) =>
                new { DEPT_NO = ddu.DEPT_NO, DEPT_NAME = ddu.DEPT_NAME, USR_ID = ddu.USR_ID, USR_NAME = u.USR_NAME, DEPT_CHIEF = ddu.DEPT_CHIEF ?? "", MGR_FG = ddu.MGR_FG ?? false }).ToList();


            string json = JsonConvert.SerializeObject(dept);
            if (json == "[]")
            {
                var dept1 = db.OGCDEPT.Where(m => m.DEPT_NO == deptno).Select(m =>
                new { DEPT_NO = m.DEPT_NO, DEPT_NAME = m.DEPT_NAME ?? "", DEPT_CHIEF = m.DEPT_CHIEF ?? "", MGR_FG = m.MGR_FG ?? false, USR_ID = "", USR_NAME = "" }).ToList();
                json = JsonConvert.SerializeObject(dept1);
            }



            // 將 List 物件轉換為 JSON 字串
            //string json = serializer.Serialize(dept);
            return Content(json, "application/json");
        }
        [HttpGet]
        public ActionResult Create()
        {
            DeptViewModel dvm = new DeptViewModel();
            dvm.id_name = db.OGCUSER.Select(x => new SelectListItem
            {
                Value = x.USR_ID,
                Text = x.USR_ID + "." + x.USR_NAME
            }).ToList();
            dvm.an_id_name = db.OGCDEPT.Where(x => x.DATA_TP == "D").Select(x => new SelectListItem
            {
                Value = x.DEPT_NO,
                Text = x.DEPT_NAME
            }).ToList();
            return View(dvm);
        }
        [HttpPost]
        public ActionResult Create(DeptViewModel dvm)
        {
            if (dvm.DEPT_CHIEF == "請選擇")
            {
                dvm.DEPT_CHIEF = null;
            }
            OGCDEPT dept = new OGCDEPT()
            {
                DEPT_NO = dvm.DEPT_NO,
                DEPT_NAME = dvm.DEPT_NAME,
                DEPT_CHIEF = dvm.DEPT_CHIEF,
                MGR_FG = dvm.MGR_FG,
                AN_DEPT_NO = dvm.AN_DEPT_NO,
                DATA_TP = "D"
            };
            db.OGCDEPT.Add(dept);
            //List傳回Usr_Id
            if (dvm.List != null)
            {
                for (int i = 0; i < dvm.List.Count(); i++)
                {
                    if (dvm.List[i].Id != null)
                    {
                        OGCDPUR dpur = new OGCDPUR()
                        {
                            DEPT_NO = dvm.DEPT_NO,
                            USR_ID = dvm.List[i].Id,
                            DATA_TP = "D",
                            ADD_DT = DateTime.Now,
                            UPD_DT = DateTime.Now,
                            UPD_UID = Helpers.WebSiteHelper.getUserId() //(string)Session["UsrId"]
                        };
                        db.OGCDPUR.Add(dpur);
                    }
                }
            }
            db.SaveChanges();
            return RedirectToAction("TreeVieww");
        }
        public ActionResult Delete(string id)
        {
            var dept = db.OGCDEPT.Where(m => m.DEPT_NO == id).FirstOrDefault();
            db.OGCDEPT.Remove(dept);
            var dpurList = db.OGCDPUR.Where(m => m.DEPT_NO == id).ToList();
            foreach (var dpur in dpurList)
            {
                db.OGCDPUR.Remove(dpur);
            }
            db.SaveChanges();

            return RedirectToAction("TreeVieww");
        }
        public ActionResult Edit(string id)
        {
            var dept = db.OGCDPUR.Join(db.OGCDEPT, m => m.DEPT_NO, n => n.DEPT_NO, (du, d) =>
            new { d.DEPT_NO, d.DEPT_NAME, d.MGR_FG, d.DEPT_CHIEF, du.USR_ID }).
            Where(ddu => ddu.DEPT_NO == id).
            Join(db.OGCUSER, m => m.USR_ID, n => n.USR_ID, (ddu, u) =>
            new { DEPT_NO = ddu.DEPT_NO, DEPT_NAME = ddu.DEPT_NAME, USR_ID = ddu.USR_ID, USR_NAME = u.USR_NAME, DEPT_CHIEF = ddu.DEPT_CHIEF ?? "", MGR_FG = ddu.MGR_FG ?? false }).ToList();
            List<DeptUserViewModel> dvmList = new List<DeptUserViewModel>();

            if (dept.Count() == 0)
            {
                var dept1 = db.OGCDEPT.Where(m => m.DEPT_NO == id).ToList();
                foreach (var item in dept1)
                {
                    DeptUserViewModel dvm = new DeptUserViewModel();
                    dvm.DEPTNO = item.DEPT_NO;
                    dvm.DEPTNAME = item.DEPT_NAME;
                    dvm.DEPT_CHIEF = item.DEPT_CHIEF;
                    if (item.DEPT_CHIEF == "" || item.DEPT_CHIEF == null)
                    {
                        dvm.DEPT_CHIEF = "請選擇";
                    }
                    dvm.MGR_FG = (bool)item.MGR_FG;
                    // 新增 dvm 內容到 dvm1 列表中
                    dvmList.Add(dvm);
                }
                DeptUserViewModel dvm1 = new DeptUserViewModel();
                dvm1.DEPTNO = id;
                dvm1.id_name = db.OGCUSER.Select(x => new SelectListItem
                {
                    Value = x.USR_ID,
                    Text = x.USR_ID + "." + x.USR_NAME
                }).ToList();
                dvmList.Add(dvm1);
            }
            else
            {
                foreach (var item in dept)
                {
                    DeptUserViewModel dvm = new DeptUserViewModel();
                    dvm.DEPTNO = item.DEPT_NO;
                    dvm.DEPTNAME = item.DEPT_NAME;
                    dvm.DEPT_CHIEF = item.DEPT_CHIEF;
                    if (item.DEPT_CHIEF == "")
                    {
                        dvm.DEPT_CHIEF = "請選擇";
                    }
                    dvm.MGR_FG = item.MGR_FG;
                    dvm.USRID = item.USR_ID;
                    dvm.USRNAME = item.USR_NAME;

                    // 新增 dvm 內容到 dvm1 列表中
                    dvmList.Add(dvm);
                }

                DeptUserViewModel dvm1 = new DeptUserViewModel();
                dvm1.DEPTNO = id;

                dvm1.id_name = db.OGCUSER.Select(x => new SelectListItem
                {
                    Value = x.USR_ID,
                    Text = x.USR_ID + "." + x.USR_NAME
                }).ToList();
                //dvm1去除和dvm有相同usrid的資料
                dvm1.id_name = dvm1.id_name.Where(x => x.Value == null || !dvmList.Any(d => d.USRID == x.Value)).ToList();
                dvmList.Add(dvm1);
            }

            return View(dvmList);

        }
        [HttpPost]
        public ActionResult Edit(List<DeptUserViewModel> duvm)
        {

            foreach (var item in duvm)
            {
                OGCDEPT dept = db.OGCDEPT.Find(item.DEPTNO, "D");
                if (dept != null)
                {
                    dept.DEPT_NAME = item.DEPTNAME;
                    dept.MGR_FG = item.MGR_FG;
                    if (item.DEPT_CHIEF == "請選擇")
                    {
                        dept.DEPT_CHIEF = null;
                    }
                    else
                    {
                        dept.DEPT_CHIEF = item.DEPT_CHIEF;
                    }
                    db.Entry(dept).State = EntityState.Modified;
                }
                db.SaveChanges();
            }

            return RedirectToAction("TreeVieww", "Dept", new { id = duvm[0].DEPTNO });
        }


        public ActionResult Edit_Delete(string id1, string id2)
        {
            var dpur = db.OGCDPUR.Where(m => m.DEPT_NO == id1 && m.USR_ID == id2).FirstOrDefault();
            db.OGCDPUR.Remove(dpur);
            var user = db.OGCUSER.Where(m => m.USR_ID == id2).FirstOrDefault();
            var dept = db.OGCDEPT.Where(m => m.DEPT_NO == id1).FirstOrDefault();

            if (user.USR_NAME == dept.DEPT_CHIEF)
            {
                db.OGCDEPT.Where(m => m.DEPT_NO == id1).FirstOrDefault().DEPT_CHIEF = null;
            }
            db.SaveChanges();

            return RedirectToAction("Edit", "Dept", new { id = id1 });
        }
        public ActionResult Edit_Create(string selectedvalue, string selectedid)
        {
            OGCDPUR dpur = new OGCDPUR()
            {
                USR_ID = selectedvalue,
                DEPT_NO = selectedid,
                DATA_TP = "D",
                UPD_DT = DateTime.Now
            };
            db.OGCDPUR.Add(dpur);
            db.SaveChanges();
            return RedirectToAction("Edit", "Dept", new { id = selectedid });
        }
    }
}