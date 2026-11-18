using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;

namespace WCals.Controllers
{
    [Authorize]
    public class WSMC0701Controller : Controller
    {
        private readonly WCals.Models.CMBASEModels.CMBASEEntities db = new WCals.Models.CMBASEModels.CMBASEEntities();
        //private readonly Models.WCalsEntities db = new Models.WCalsEntities();
        // GET: WSMC0701
        [LogAction]
        public ActionResult Index()
        {
            
            var smcmenu = db.SMCMENU3.Where(f => f.DATA_TP == "N").OrderBy(f => f.SYS_ORD).ToList();
            return View(smcmenu);
        }

        public ActionResult SubMenu(String id)
        {
            var smcmenu = db.SMCMENU3.Where(f => f.AN_SYS_ID == id).OrderBy(f => f.SYS_ORD).ToList();
            return View(smcmenu);
        }

        public ActionResult Edit(String id)
        {
            var smcmenu = db.SMCMENU3.Where(f => f.SYS_ID == id).OrderBy(f => f.SYS_ORD).FirstOrDefault();

            var smcauth_D = (from f in db.SMCAUTH
                           join g in db.OGCDEPT on f.DEPT_NO equals g.DEPT_NO
                           where f.SYS_ID == id && g.DATA_TP == "D"
                           select new { f.DEPT_NO, f.DATA_TP, f.SYS_ID, g.DEPT_NAME,  METHOD="original" }).ToList();

            var smcauth_G = (from f in db.SMCAUTH
                           join g in db.OGCDEPT on f.DEPT_NO equals g.DEPT_NO
                           where f.SYS_ID == id && g.DATA_TP == "G"
                           select new { f.DEPT_NO, f.DATA_TP, f.SYS_ID, g.DEPT_NAME, METHOD = "original" }).ToList();

            var ogcdept_G = (from f in db.OGCDEPT
                           where f.DATA_TP == "G"
                           select f).ToList();

            var ogcdept_D = (from f in db.OGCDEPT
                           where f.DATA_TP == "D"
                           select f).ToList();
            ViewBag.DEPT_NO_G = new SelectList(ogcdept_G, "DEPT_NO", "DEPT_NAME");
            ViewBag.DEPT_NO_D = new SelectList(ogcdept_D, "DEPT_NO", "DEPT_NAME");

            var model = new ViewModels.AUTHViewModel
            {
                DataSMCMENU = smcmenu,
                jsonSMCAUTH_D = JsonConvert.SerializeObject(smcauth_D),
                jsonSMCAUTH_G = JsonConvert.SerializeObject(smcauth_G)
            };
            return View(model);
        }

        [HttpPost]
        public ActionResult Edit(ViewModels.AUTHViewModel model)
        {
            String jsonSMCAUTH_D = model.jsonSMCAUTH_D;
            String jsonSMCAUTH_G = model.jsonSMCAUTH_G;

            proceedEdit(model.DataSMCMENU.SYS_ID, jsonSMCAUTH_D, "D");
            proceedEdit(model.DataSMCMENU.SYS_ID, jsonSMCAUTH_G, "G");

            return RedirectToAction("Edit", "WSMC0701", new { @id= model.DataSMCMENU.SYS_ID });
        }

        private void proceedEdit(string SYS_ID, string json, string type)
        {
            if (String.IsNullOrEmpty(json) == false)
            {
                dynamic x = JsonConvert.DeserializeObject(json);
                for (int i = 0; i < x.Count; i++)
                {
                    if (x[i].METHOD == "new")
                    {
                        //Models.SMCAUTH_WCLAS table = new Models.SMCAUTH_WCLAS();
                        var table = new WCals.Models.CMBASEModels.SMCAUTH();
                        //table.SYS_ID = model.DataSMCMENU.SYS_ID;
                        table.SYS_ID = SYS_ID;
                        table.DEPT_NO = x[i].DEPT_NO;
                        table.DATA_TP = type;
                        db.SMCAUTH.Add(table);
                        db.SaveChanges();
                    }
                    else if (x[i].METHOD == "edit")
                    {
                        if (x[i].DEPT_NO != null && x[i].SYS_ID != null)
                        {
                            String deptNo = x[i].DEPT_NO;
                            String sysId = x[i].SYS_ID;
                            var query = (from f in db.SMCAUTH
                                         where f.DEPT_NO == deptNo && f.SYS_ID == sysId
                                         select f).FirstOrDefault();
                            query.SYS_ID = x[i].SYS_ID;
                            query.DEPT_NO = x[i].DEPT_NO;
                            query.DATA_TP = type;
                            db.SaveChanges();
                        }
                    }
                    else if (x[i].METHOD == "delete")
                    {
                        if (x[i].DEPT_NO != null && x[i].SYS_ID != null)
                        {
                            String deptNo = x[i].DEPT_NO;
                            String sysId = x[i].SYS_ID;
                            var query = db.SMCAUTH.Where(f => f.DEPT_NO == deptNo).Where(f => f.SYS_ID == sysId).FirstOrDefault();
                            db.SMCAUTH.Remove(query);
                            db.SaveChanges();
                        }
                    }
                }
            }
        }

    }
}