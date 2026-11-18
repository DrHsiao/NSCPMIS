using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMEPMSModels;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS
{
    public class PJE0010Controller : Controller
    {

        /// private readonly Models.CMBASEModels.CMBASEEntities cmbase = new Models.CMBASEModels.CMBASEEntities();
        private readonly Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();

        int pageSize = 10;

        private IEnumerable<Models.CMEPMSModels.PJV_PJMFORM> GetAllPJMFORM()
        {
            string userdept = WebSiteHelper.getUserDept();
            string userid = Helpers.WebSiteHelper.getUserId();


            if (userdept.Equals("B150100"))
            {
                var pjmform = db.PJV_PJMFORM.AsNoTracking().Where(p => p.PJ_TYPE == "W").AsQueryable();
                pjmform = pjmform.OrderByDescending(p => p.SYR).
                ThenByDescending(p => p.END_SYR).
                ThenByDescending(p => p.PID);
                return pjmform.ToList();
            }
            else
            {
                var pjmform = db.PJV_PJMFORM.AsNoTracking().Where(p => p.ADV_UID == userid && p.PJ_TYPE == "W").AsQueryable();
                pjmform = pjmform.OrderByDescending(p => p.SYR).
                ThenByDescending(p => p.END_SYR).
                ThenByDescending(p => p.PID);
                return pjmform.ToList();
            }




            //return pjmform.AsQueryable().ToList();

        }

        /// <summary>
        /// 取所有專案
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private IEnumerable<Models.CMEPMSModels.PJV_PJMFORM> GetAnPJMFORM(PJEModel model)
        {
            string userid = Helpers.WebSiteHelper.getUserId();
            var pjmform = db.PJV_PJMFORM.AsNoTracking().Where(p => p.AN_PID == model.PID).AsQueryable();
            pjmform = pjmform.OrderByDescending(p => p.SYR).
                ThenByDescending(p => p.END_SYR).
                ThenByDescending(p => p.PID);
            //return pjmform.AsQueryable().ToList();
            return pjmform.ToList();
        }

        public ActionResult Index(PJEModel model)
        {

            var pjmf = GetAllPJMFORM();//欄位自動更新會壞
            int page = (model.page == 0) ? 1 : model.page;
            //if(pageNumber == 0) pageNumber = 1;
            //var ps = GetAllPJMFORM();

            //string sql = "SELECT * FROM PJV_PJMFORM ORDER BY @SYR, @END_SYR, @PID";
            //SqlCommand cmd = new SqlCommand("SELECT * FROM PJV_PJMFORM ORDER BY SYR, END_SYR, PID");
            //SqlCommand cmd = new SqlCommand("SELECT SYR, END_SYR, PID, * FROM PJV_PJMFORM ORDER BY @SYR, @END_SYR, @PID");
            //cmd.Parameters.Add(new SqlParameter("SYR", filter("SYR")));
            //cmd.Parameters.Add(new SqlParameter("END_SYR", "END_SYR"));
            //cmd.Parameters.Add(new SqlParameter("PID", "1"));
            //DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            //List<DataRow> list=ps.AsEnumerable().ToList();
            //IPagedList<DataRow> plist= new PagedList<DataRow>(list, page ?? 1, 10); 
            IPagedList<PJV_PJMFORM> plist = new PagedList<PJV_PJMFORM>(pjmf, page, pageSize);
            model.page = page;
            model.PJV_PJMFORM = plist;
            return View("PJE0010Index", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Add")]
        public ActionResult Add()
        {
            string year = Convert.ToString(DateTime.Now.AddYears(-1911).Year);
            PJEModel model = new PJEModel();
            model.PJMFORM = new PJMFORM_CMEPMS();
            model.PJMFORM.SYR = year;
            model.PJMFORM.END_SYR = year;

            return View("PJE0010Edit", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Edit")]
        public ActionResult Edit(PJEModel model)
        {

            model.PJMFORM = db.PJMFORM.First(f => f.PID == model.PID);
            model.PID = model.PJMFORM.PID;
            return View("PJE0010Edit", model);

        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Del")]
        public ActionResult Del(PJEModel model)
        {


            var ps = db.PJMFORM.Where(f => f.AN_PID == model.PID);
            if (ps.Count() > 0)
            {
                ViewBag.msg = "存在分支計畫不允許刪除";
                ViewBag.desc = "請確認資料";
                return Index(model);
            }




            PJMFORM_CMEPMS p = db.PJMFORM.Where(f => f.PID == model.PID).FirstOrDefault();

            /*if(p == null) {
                ViewBag.msg = "資料不存在";
                ViewBag.desc = "請確認資料";
                return Index(model);
            }*/
            db.PJMFORM.Remove(p);
            db.SaveChanges();
            ViewBag.msg = "已刪除";

            return Index(model);

        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Save")]
        public ActionResult Save(PJEModel model)
        {

            PJMFORM_CMEPMS p;
            if (model.PID == 0)
            {
                p = new PJMFORM_CMEPMS();
            }
            else
            {
                p = db.PJMFORM.Where(f => f.PID == model.PID).FirstOrDefault();
            }
            p.SYR = model.SYR;
            p.END_SYR = model.END_SYR;
            p.PJ_NAME = model.PJ_NAME;
            p.WK_CODE = model.WK_CODE;
            p.PJ_FUND = tools.safe2Decimal(model.PJ_FUND);
            p.ADV_DP = WCals.Helpers.WebSiteHelper.getUserDept();
            p.UPD_UID = WCals.Helpers.WebSiteHelper.getUserId();
            p.ADV_UID = WCals.Helpers.WebSiteHelper.getUserId();
            p.PJ_TYPE = "W";
            if (model.PID == 0)
            {
                p.PJ_TYPE = "W";
                db.PJMFORM.Add(p);
            }
            db.SaveChanges();
            model.PID = p.PID;
            ViewBag.msg = "已儲存";
            ViewBag.desc = "確認完成後，點選返回至列表";
            return Edit(model);
        }

        public ActionResult Branch1(decimal pid)
        {
            return Branch(pid, 0, 0, "", 1);
        }

        public ActionResult Branch(decimal pid, decimal EditPID = 0, int page = 0, string ANPID = "", int branchPage = 0)
        {
            PJEModel model = new PJEModel();
            model.PID = pid;
            model.EditPID = EditPID;
            model.ANPID = string.IsNullOrEmpty(ANPID) ? pid.ToString() : ANPID;
            model.page = (page == 0) ? 1 : page;
            model.branchPage = (branchPage == 0) ? 1 : branchPage;

            var pjmf = GetAnPJMFORM(model);
            if (pjmf.Count() == 0)
            {
                ViewBag.msg = "無分支計畫";
                ViewBag.desc = "無所屬分支計畫可列表";
            }

            IPagedList<PJV_PJMFORM> plist = new PagedList<PJV_PJMFORM>(pjmf, model.branchPage, pageSize);
            model.PJV_PJMFORM = plist;
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)
            return View("PJE0010Branch", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Branch")]
        public ActionResult Branch(PJEModel model)
        {
            var pjmf = GetAnPJMFORM(model);
            if (pjmf.Count() == 0)
            {
                ViewBag.msg = "無分支計畫";
                ViewBag.desc = "無所屬分支計畫可列表";
            }
            model.ANPID = string.IsNullOrEmpty(model.ANPID) ? model.PID.ToString() : model.ANPID;
            model.page = (model.page == 0) ? 1 : model.page;
            model.branchPage = (model.branchPage == 0) ? 1 : model.branchPage;
            IPagedList<PJV_PJMFORM> plist = new PagedList<PJV_PJMFORM>(pjmf, model.branchPage, pageSize);
            model.PJV_PJMFORM = plist;
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)
            return View("PJE0010Branch", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BranchEdit")]
        public ActionResult BranchEdit(PJEModel model)
        {
            decimal pid = tools.safe2Decimal(model.ANPID);
            if (model.EditPID != 0)
            {
                model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(pid, model.EditPID); //取得目前所在的主計畫名稱(+分支計畫名稱)
                                                                                                    //編輯                
                model.PJMFORM = db.PJMFORM.First(f => f.PID == model.EditPID);
                model.PJMFORM_P = db.PJMFORM.First(f => f.PID == pid);
            }
            else
            {
                model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)
                                                                                           //新增
                model.PJMFORM = new PJMFORM_CMEPMS();

                PJMFORM_CMEPMS p = db.PJMFORM.Where(w => w.PID == pid).FirstOrDefault();
                if (p != null)
                {


                    //PJMFORM_CMEPMS p = db.PJMFORM.Where(w => w.PID == pid).FirstOrDefault();
                    if (p != null)
                    {
                        model.PJMFORM = new PJMFORM_CMEPMS();
                        model.PJMFORM.SYR = p.SYR;
                        model.PJMFORM.END_SYR = p.END_SYR;
                        model.PJMFORM.WK_CODE = p.WK_CODE;


                        model.PJMFORM.WK_CODE = p.WK_CODE;
                    }
                }
            }


            //經費類別
            ViewBag.FUND_CATY = tools.ToSelectList(util.GetCMEPMSCODES("FD", "06"), "id", "name", null);



            return View("PJE0010BranchEdit", model);

        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BranchSave")]
        public ActionResult BranchSave(PJEModel model)
        {


            if (string.IsNullOrEmpty(model.PJMFORM.PJ_NAME))
            {
                ViewBag.msg = "錯誤";
                ViewBag.desc = "請檢查欄位";
                return BranchEdit(model);
            }
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            string userdp = WebSiteHelper.getUserDept();
            PJMFORM_CMEPMS p;
            if (model.EditPID == 0)
            {
                p = new PJMFORM_CMEPMS();
            }
            else
            {
                p = db.PJMFORM.Where(f => f.PID == model.EditPID).FirstOrDefault();
            }
            p.SYR = model.PJMFORM.SYR;
            p.END_SYR = model.PJMFORM.END_SYR;
            p.PJ_NAME = model.PJMFORM.PJ_NAME;


            p.PJ_CODE = model.PJMFORM.PJ_CODE;
            p.PJ_FUND = model.PJMFORM.PJ_FUND;

            p.W_CODE = model.PJMFORM.W_CODE;
            p.E_CODE = model.PJMFORM.E_CODE;

            p.E_CODE_N = model.PJMFORM.E_CODE_N;
            p.W_CODE_N = model.PJMFORM.W_CODE_N;

            p.Water_Type = model.PJMFORM.Water_Type;

            p.ADV_DP = WCals.Helpers.WebSiteHelper.getUserDept();
            p.UPD_UID = WCals.Helpers.WebSiteHelper.getUserId();
            p.ADV_UID = WCals.Helpers.WebSiteHelper.getUserId();
            p.AN_PID = Convert.ToDecimal(model.ANPID);
            p.PJ_TYPE = "B";

            p.CONT_DESC = model.PJMFORM.CONT_DESC;
            p.GP_ORDER = model.PJMFORM.GP_ORDER;
            p.FUND_CATY = model.PJMFORM.FUND_CATY;

            p.SIP_W_CODE = model.PJMFORM.SIP_W_CODE;

            p.WK_CODE = model.PJMFORM.WK_CODE;
            p.SIP_W_CODE = model.PJMFORM.SIP_W_CODE;
            p.SIP_W_CODE = model.PJMFORM.SIP_W_CODE;

            p.UPD_DT = now;

            if (model.EditPID == 0)
            {
                //新增
                p.ADD_DT = now;
                p.ADV_UID = userid;
                p.ADV_DP = userdp;
                db.PJMFORM.Add(p);
                db.SaveChanges();
                //第0期
                PJDFDCR f = new PJDFDCR();
                f.SEQ_NO = 0;
                f.PID = p.PID;
                f.ADD_DT = now;
                f.UPD_DT = now;
                f.UPD_UID = userid;
                db.PJDFDCR.Add(f);
                db.SaveChanges();
            }
            else
            {
                //非新增
                db.SaveChanges();
            }





            //model.PID = p.PID;
            model.EditPID = p.PID;
            ViewBag.msg = "已儲存";
            ViewBag.desc = "確認完成後，點選返回至列表";
            return BranchEdit(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BranchDel")]
        public ActionResult BranchDel(PJEModel model)
        {


            var ps = db.PJMFORM.Where(f => f.AN_PID == model.EditPID);
            if (ps.Count() > 0)
            {
                ViewBag.msg = "存在分支計畫";
                ViewBag.desc = "請確認資料";
                model.PID = Convert.ToDecimal(model.ANPID);
                return Branch(model);
            }
            /*
            PJMFORM_CMEPMS p = db.PJMFORM.Where(f => f.PID == model.PID).FirstOrDefault();
            if(p == null) {
                ViewBag.msg = "資料不存在";
                ViewBag.desc = "請確認資料";
                model.PID = Convert.ToDecimal(model.ANPID);
                return Branch(model);
            }
            db.PJMFORM.Remove(p);
            db.SaveChanges();
            model.PID = Convert.ToDecimal(model.ANPID);
            */

            SqlDao.CMEPMSExecuteSql("DELETE FROM PJMFORM WHERE PID=" + model.EditPID);


            ViewBag.msg = "已刪除";
            //ViewBag.desc = "請確認資料";
            return Branch(model);

        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "BranchAdd")]
        public ActionResult BranchAdd(PJEModel model)
        {
            decimal pid = tools.safe2Decimal(model.ANPID);
            PJMFORM_CMEPMS p = db.PJMFORM.Where(w => w.PID == pid).FirstOrDefault();
            if (p != null)
            {
                model.SYR = p.SYR;

                model.PJMFORM = new PJMFORM_CMEPMS();
                model.PJMFORM.WK_CODE = p.WK_CODE;
            }





            return View("PJE0010BranchEdit", model);
        }

        /// <summary>
        /// 資料授權
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Assign")]
        public ActionResult Assign(PJEModel model)
        {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.EditPID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            model.PJMFORM = new PJMFORM_CMEPMS();
            SqlCommand cmd = new SqlCommand("select A.DID, A.PID, D.DEPT_NAME, U.USR_NAME from PJDAUTH A left join CMBASE.dbo.OGCDEPT D on A.EXE_DP = D.DEPT_NO left join CMBASE.dbo.OGCUSER U on A.EXE_UID = U.USR_ID WHERE A.PID=@PID");

            cmd.Parameters.Add(new SqlParameter("PID", model.EditPID));

            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.list = ps;
            //ViewBag.depts = tools.ToSelectList(SqlDao.CMBASEQuery4List("SELECT DEPT_NO, DEPT_NAME FROM OGCDEPT WHERE AN_DEPT_NO IS NULL"), "DEPT_NO", "DEPT_NAME", null);
            string sql = "SELECT o.DEPT_NO,\r\n    MAX(o.DEPT_NAME) AS DEPT_NAME, \r\n COUNT(r.USR_ID) AS USER_COUNT \r\nFROM OGCDPUR r\r\nJOIN OGCDEPT o ON o.DEPT_NO = r.DEPT_NO\r\nJOIN OGCUSER u ON u.USR_ID = r.USR_ID\r\nGROUP BY o.DEPT_NO;";
            ViewBag.depts = tools.ToSelectList(SqlDao.CMBASEQuery4List(sql), "DEPT_NO", "DEPT_NAME", "請選擇單位");
            return View("PJE0010BranchAssign", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "AssignDelUser")]
        public ActionResult AssignDelUser(PJEModel model)
        {

            BASE_DAO.SqlDao.CMEPMSExecuteSql("DELETE FROM PJDAUTH WHERE DID='" + model.DID + "'");
            ViewBag.msg = "刪除完成";
            return Assign(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "AssignAddUser")]
        public ActionResult AssignAddUser(PJEModel model)
        {

            if (!string.IsNullOrEmpty(model.EXE_UID) && !string.IsNullOrEmpty(model.EXE_DP))
            {

                try
                {

                    PJDAUTH p = new PJDAUTH();
                    p.PID = model.EditPID;
                    p.EXE_DP = model.EXE_DP;
                    p.EXE_UID = model.EXE_UID;
                    p.ADD_DT = DateTime.Now;
                    p.UPD_UID = WebSiteHelper.getUserId();
                    db.PJDAUTH.Add(p);
                    db.SaveChanges();
                    model.EXE_DP = "";
                    ViewBag.msg = "新增完成";
                }
                catch
                {
                    ViewBag.msg = "新增失敗";
                }


            }
            else
            {
                ViewBag.msg = "請選擇單位和人員";
            }


            return Assign(model);
        }

        [HttpPost]
        public JsonResult GetUsers(string DEPT_NO)
        {
            DataTable dt = BASE_DAO.SqlDao.CMBASEQuery4List("SELECT u.USR_ID, u.USR_NAME FROM OGCUSER u, OGCDEPT d, OGCDPUR p WHERE d.AN_DEPT_NO='" + DEPT_NO + "' AND p.USR_ID=u.USR_ID AND d.DEPT_NO=p.DEPT_NO");
            return Json(tools.ConvertDataTabletoString(dt), JsonRequestBehavior.AllowGet);
        }
    }

}
