using PagedList;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using static WCals.classes.john.BASE_DAO;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMEPMSModels;
using WCals.ViewModels.NewModels;

namespace WCals.Controllers.NSPMIS
{
    public class PJE0020Controller : Controller
    {

        private readonly Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();
        int pageSize = 10;

        //分支
        private IEnumerable<Models.CMEPMSModels.PJV_PJMFORM> GetAllPJMFORM_B()
        {
            string userid = Helpers.WebSiteHelper.getUserId();
            //List<PJV_PJMFORM> Pjmform = db.PJV_PJMFORM.Where(x => x.ADV_UID == userid && x.PJ_TYPE == "B"
            List<PJV_PJMFORM> Pjmform = db.PJV_PJMFORM.Where(x => x.ADV_UID == userid && x.PJ_TYPE == "B")
            .OrderByDescending(x => x.SYR)
            .ThenByDescending(x => x.AN_PID)
            .ThenBy(x => x.PJ_CODE)
            //.ThenBy(x => x.PID)
            //.Select(x => x.p)  // 只選擇 PJV_PJMFORM 的 p
            //.Distinct()         // 去除重複
            .ToList();          // 轉換為 List<PJV_PJMFORM>

            List<PJV_PJMFORM> aDdPjmform = db.PJV_PJMFORM
            .Join(db.PJDAUTH,
                p => p.PID,
                a => a.PID,
                (p, a) => new { p, a })
            .Where(x => x.a.EXE_UID == userid && x.p.PJ_TYPE == "B")
            .OrderByDescending(x => x.p.SYR)
            .ThenByDescending(x => x.p.AN_PID)
            .ThenBy(x => x.p.PJ_CODE)
            //.ThenBy(x => x.p.PID)
            .Select(x => x.p)  // 只選擇 PJV_PJMFORM 的 p
                               //.Distinct()         // 去除重複
            .ToList();          // 轉換為 List<PJV_PJMFORM>
            Pjmform.AddRange(aDdPjmform);
            return Pjmform;
        }



        public ActionResult Index(PJE0020Model model)
        {
            int page = (model.page == 0) ? 1 : model.page;
            var pjmf = GetAllPJMFORM_B();//欄位自動更新會壞
            IPagedList<PJV_PJMFORM> plist = new PagedList<PJV_PJMFORM>(pjmf, page, pageSize);
            model.PJV_PJMFORM = plist;
            return View("PJE0020Index", model);
        }

        //分項
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SubItem")]
        public ActionResult SubItem(PJE0020Model model)
        {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            model.PJMFORM = new PJMFORM_CMEPMS();
            SqlCommand cmd = new SqlCommand("select C.*, S1.CODE_DESC, S2.CODE_DESC from PJDCATY C left join SMCCODE S1 on " +
                "C.ACC_ID =S1.CATY_COD and S1.SYS_ID='FD' and S1.CATY_ID='02' left join SMCCODE S2 on " +
                "C.CATY_TYPE =S2.CATY_COD and S2.SYS_ID='FD' and S2.CATY_ID='03'where PID=@PID ORDER BY CID DESC");

            cmd.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.list = ps;
            //ViewBag.depts = tools.ToSelectList(SqlDao.CMBASEQuery4List("SELECT DEPT_NO, DEPT_NAME FROM OGCDEPT WHERE AN_DEPT_NO IS NULL"), "DEPT_NO", "DEPT_NAME", true);
            //ViewBag.depts = SqlDao.CMBASEQuery4List("SELECT DEPT_NO, DEPT_NAME FROM OGCDEPT WHERE AN_DEPT_NO IS NULL");
            string sql = "SELECT o.DEPT_NO, MAX(o.DEPT_NAME) AS DEPT_NAME, COUNT(r.USR_ID)" +
                "AS USER_COUNT FROM OGCDPUR r JOIN OGCDEPT o ON o.DEPT_NO = r.DEPT_NO JOIN " +
                "OGCUSER u ON u.USR_ID = r.USR_ID GROUP BY o.DEPT_NO;";
            ViewBag.depts = SqlDao.CMBASEQuery4List(sql);
            return View("PJE0020SubItem", model);
        }

        //編輯分項
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SubItemEdit")]
        public ActionResult SubItemEdit(PJE0020Model model)
        {

            model.PJMFORM = new PJMFORM_CMEPMS();
            SqlCommand cmd = new SqlCommand("select CID, CATY_CODE, CATY_NAME, CATY_FUND,PJ_CATY_CODE, S1.CODE_DESC, S2.CODE_DESC from PJDCATY C left join SMCCODE S1 on  C.ACC_ID =S1.CATY_COD and S1.SYS_ID='FD' and S1.CATY_ID='02' left join SMCCODE S2 on  C.CATY_TYPE =S2.CATY_COD and S2.SYS_ID='FD' and S2.CATY_ID='03'where PID=@PID AND CID=@CID");

            cmd.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            cmd.Parameters.Add(new SqlParameter("CID", tools.SqlFilterInt(model.CID)));
            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.list = ps;

            return View("PJE0020SubItemEdit", model);
        }

        //儲存分項
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SubItemSave")]
        public ActionResult SubItemSave(PJE0020Model model)
        {

            decimal seq;
            PJDCATY_CMEPMS plast;

            if (model.PIDS == null)
            {
                ViewBag.msg = "資料不完整";
                return SubItem(model);
            }

            if (model.CATY_CODES.GroupBy(n => n).Any(g => g.Count() > 1))
            {
                ViewBag.msg = "資料重複";
                return SubItem(model);
            }
            DateTime now = DateTime.Now;
            PJDCATY_CMEPMS p;
            decimal cid, pid = model.PID;
            StringBuilder err = new StringBuilder("");
            for (int i = 0; i < model.PIDS.Length; i++)
            {

                //新增
                if (string.IsNullOrEmpty(model.CIDS[i]))
                {

                    plast = db.PJDCATY.Where(w => w.PID == model.PID).OrderByDescending(g => g.CID).FirstOrDefault();
                    if (plast == null)
                    {
                        seq = 1;
                    }
                    else
                    {
                        seq = (decimal)plast.CID + 1;
                    }
                    //新增一筆
                    p = new PJDCATY_CMEPMS();
                    p.PID = model.PID;
                    p.CID = seq;
                    p.CATY_CODE = model.CATY_CODES[i];
                    p.CATY_FUND = tools.safe2Decimal(model.CATY_FUNDS[i]);
                    p.CATY_TYPE = model.CATY_TYPES[i];
                    p.CATY_NAME = model.CATY_NAMES[i];
                    p.ACC_ID = model.ACC_IDS[i];
                    p.ADV_DP = WebSiteHelper.getUserDept();
                    p.ADV_UID = WebSiteHelper.getUserId();
                    p.ADD_DT = DateTime.Now;
                    try
                    {
                        db.PJDCATY.Add(p);
                        db.SaveChanges();
                    }
                    catch (Exception e)
                    {
                        err.Append("第" + (i + 1) + "筆資料異常: " + e);
                        continue;
                    }


                    continue;


                }



                //修改
                pid = Convert.ToDecimal(model.PIDS[i]);
                cid = Convert.ToDecimal(model.CIDS[i]);

                p = db.PJDCATY.Where(w => w.PID == pid && w.CID == cid).FirstOrDefault();


                p.CATY_CODE = model.CATY_CODES[i];
                p.CATY_FUND = tools.safe2Decimal(model.CATY_FUNDS[i]);
                p.ACC_ID = model.ACC_IDS[i];
                p.CATY_TYPE = model.CATY_TYPES[i];
                p.CATY_NAME = model.CATY_NAMES[i];

                p.ADV_DP = WebSiteHelper.getUserDept();
                p.ADV_UID = WebSiteHelper.getUserId();
                p.UPD_DT = now;



                try
                {
                    db.SaveChanges();
                }
                catch (Exception e)
                {
                    err.Append("第" + (i + 1) + "筆資料異常: " + e);
                    break;
                }

            }

            if (err.Length > 0)
            {
                ViewBag.msg = "錯誤";
                err.Length -= 2;
                ViewBag.desc = err.ToString();
            }
            else
            {
                ViewBag.msg = "已儲存";
            }


            return SubItem(model);
        }

        //新增分項
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SubItemAdd")]
        public ActionResult SubItemAdd(PJE0020Model model)
        {


            if (model.CATY_CODES[0].Equals("") || model.CATY_FUNDS[0].Equals(""))
            {
                ViewBag.msg = "請檢查欄位";
                return SubItem(model);
            }

            //最後一筆CID
            decimal seq;
            PJDCATY_CMEPMS plast = db.PJDCATY.Where(w => w.PID == model.PID).OrderByDescending(g => g.CID).FirstOrDefault();
            if (plast == null)
            {
                seq = 1;
            }
            else
            {
                seq = (decimal)plast.CID + 1;
            }
            //新增一筆
            PJDCATY_CMEPMS p = new PJDCATY_CMEPMS();
            p.PID = model.PID;
            p.CID = seq;
            p.CATY_CODE = model.CATY_CODES[0];
            p.CATY_FUND = Convert.ToDecimal(model.CATY_FUNDS[0]);
            p.CATY_TYPE = model.CATY_TYPES[0];
            p.CATY_NAME = model.CATY_NAMES[0];
            p.ACC_ID = model.ACC_IDS[0];
            p.ADV_DP = WebSiteHelper.getUserDept();
            p.ADV_UID = WebSiteHelper.getUserId();
            p.ADD_DT = DateTime.Now;
            db.PJDCATY.Add(p);
            db.SaveChanges();

            /*model.CATY_CODES[0]));
             *model.CATY_FUNDS[0])); 
             *model.CATY_TYPES[0])); 
             *model.ADV_DPS[0])); SqlCommand cmd = new SqlCommand("SELECT*FROM PJDCATY WHERE PID=@PID ORDER BY CID DESC");
            cmmodel.ADV_UIDS[0]));d.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            Damodel.CATY_NAMES[0]));taRow p = BASE_DAO.SqlDao.CMEPMSQuery4Object(cmd);
            demodel.ACC_IDS[0]));cimal seq;
            if(p == null) {
                seq = 1;
            } else {
                seq = tools.safe2Decimal(p["CID"].ToString());
                seq++;
            }*

            //新增一筆
            PJDCATY_CMEPMS plast
            cmd = new SqlCommand("INSERT INTO PJDCATY(PID, CID, CATY_CODE, CATY_FUND, CATY_TYPE, ADV_DP, ADV_UID, CATY_NAME, ACC_ID)VALUES" +
                "(@PID, @CID, @CATY_CODE, @CATY_FUND, @CATY_TYPE, @ADV_DP, @ADV_UID, @CATY_NAME, @ACC_ID)");
            cmd.Parameters.Add(new SqlParameter("PID", model.PID));
            //cmd.Parameters.Add(new SqlParameter("CID", model.CIDS[0]));
            cmd.Parameters.Add(new SqlParameter("CID", seq));
            cmd.Parameters.Add(new SqlParameter("CATY_CODE", model.CATY_CODES[0]));
            cmd.Parameters.Add(new SqlParameter("CATY_FUND", model.CATY_FUNDS[0]));
            cmd.Parameters.Add(new SqlParameter("CATY_TYPE", model.CATY_TYPES[0]));
            //cmd.Parameters.Add(new SqlParameter("ADV_DP",  model.ADV_DPS[0]));
            //cmd.Parameters.Add(new SqlParameter("ADV_UID", model.ADV_UIDS[0]));
            cmd.Parameters.Add(new SqlParameter("CATY_NAME", model.CATY_NAMES[0]));
            cmd.Parameters.Add(new SqlParameter("ACC_ID",    model.ACC_IDS[0]));

            BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
            /*
             * cmd.Parameters.Add(new SqlParameter("PID", model.PIDS[i]));



            cmd = new SqlCommand("INSERT INTO PJDCATY(PID, CID)VALUES(@PID, @CID, @CATY_CODE)");            
            cmd.Parameters.Add(new SqlParameter("PID", Utilities.SafeDecimal(p["PID"])));
            cmd.Parameters.Add(new SqlParameter("CID", Utilities.SafeDecimal(p["CID"])));
            cmd.Parameters.Add(new SqlParameter("CATY_CODE", p["CID"]));
            BASE_DAO.SqlDao.CMEPMSExecuteSql(cmd);
            model.CID = Utilities.SafeDecimal(p["CID"]);
            */
            ViewBag.msg = "已新增";
            return SubItem(model);
        }

        //刪除分項
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SubItemDel")]
        public ActionResult SubItemDel(PJE0020Model model, string dPID, string dCID)
        {
            SqlCommand cmd;

            try
            {
                cmd = new SqlCommand("DELETE FROM PJDCATY WHERE PID=@PID AND CID=@CID");
                cmd.Parameters.Add(new SqlParameter("PID", dPID));
                cmd.Parameters.Add(new SqlParameter("CID", dCID));
                SqlDao.CMEPMSExecuteSql(cmd);

            }
            catch (Exception e)
            {

            }

            ViewBag.msg = "已刪除";
            return SubItem(model);
        }

        //月分配
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Monthly")]
        public ActionResult Monthly(PJE0020Model model)
        {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            model.PJMFORM = db.PJMFORM.Where(w => w.PID == model.PID).FirstOrDefault();
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJDFDDS p WHERE p.PID=@PID");
            cmd.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.list = ps;

            return View("PJE0020Monthly", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "MonthlySave")]
        public ActionResult MonthlySave(PJE0020Model model, string[] delKey)
        {
            DateTime now = DateTime.Now;
            /*
            model.PJMFORM = new PJMFORM_CMEPMS();
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJDFDDS p WHERE p.PID=@PID");
            cmd.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);
            ViewBag.list = ps;
            */
            PJDFDDS p;
            string syr;
            for (int i = 1; i < model.PIDS.Length; i++)
            {


                if (string.IsNullOrEmpty(model.PIDS[i]))
                {

                    if (string.IsNullOrEmpty(model.SYRS[i]))
                    {
                        continue;
                    }
                    p = new PJDFDDS();
                    p.ADD_DT = now;
                    p.PID = model.PID;

                }
                else
                {
                    syr = model.SYRS[i];
                    p = db.PJDFDDS.Where(f => f.PID == model.PID && f.SYR == syr).FirstOrDefault();
                }



                p.SYR = model.SYRS[i];
                p.UPD_DT = now;
                p.UPD_UID = WebSiteHelper.getUserId();
                p.M_01 = tools.safe2Decimal(model.M_01[i]);
                p.M_02 = tools.safe2Decimal(model.M_02[i]);
                p.M_03 = tools.safe2Decimal(model.M_03[i]);
                p.M_04 = tools.safe2Decimal(model.M_04[i]);
                p.M_05 = tools.safe2Decimal(model.M_05[i]);
                p.M_06 = tools.safe2Decimal(model.M_06[i]);
                p.M_07 = tools.safe2Decimal(model.M_07[i]);
                p.M_08 = tools.safe2Decimal(model.M_08[i]);
                p.M_09 = tools.safe2Decimal(model.M_09[i]);
                p.M_10 = tools.safe2Decimal(model.M_10[i]);
                p.M_11 = tools.safe2Decimal(model.M_11[i]);
                p.M_12 = tools.safe2Decimal(model.M_12[i]);


                if (string.IsNullOrEmpty(model.PIDS[i]))
                {
                    db.PJDFDDS.Add(p);
                }
                try
                {
                    db.SaveChanges();

                }
                catch (Exception e)
                {

                    ViewBag.msg = "年度資料重複:";

                    return Monthly(model);
                }

            }

            ViewBag.msg = "已儲存";
            return Monthly(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "MonthlyAdd")]
        public ActionResult MonthlyAdd(PJE0020Model model)
        {
            if (string.IsNullOrEmpty(model.SYRS[0]))
            {
                ViewBag.msg = "年份為必要欄位";
                return Monthly(model);
            }
            PJDFDDS p = new PJDFDDS();
            p.PID = model.PID;
            p.SYR = model.SYRS[0];
            p.ADD_DT = DateTime.Now;
            p.UPD_UID = WebSiteHelper.getUserId();
            p.M_01 = tools.safe2Decimal(model.M_01[0]);
            p.M_02 = tools.safe2Decimal(model.M_02[0]);
            p.M_03 = tools.safe2Decimal(model.M_03[0]);
            p.M_04 = tools.safe2Decimal(model.M_04[0]);
            p.M_05 = tools.safe2Decimal(model.M_05[0]);
            p.M_06 = tools.safe2Decimal(model.M_06[0]);
            p.M_07 = tools.safe2Decimal(model.M_07[0]);
            p.M_08 = tools.safe2Decimal(model.M_08[0]);
            p.M_09 = tools.safe2Decimal(model.M_09[0]);
            p.M_10 = tools.safe2Decimal(model.M_10[0]);
            p.M_11 = tools.safe2Decimal(model.M_11[0]);
            p.M_12 = tools.safe2Decimal(model.M_12[0]);
            db.PJDFDDS.Add(p);
            db.SaveChanges();
            ViewBag.msg = "已新增";
            return Monthly(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "MonthlyDel")]
        public ActionResult MonthlyDel(PJE0020Model model, string[] delKey)
        {
            for (int i = 0; i < delKey.Length; i++)
            {
                if (delKey[i].Equals("*"))
                {
                    BASE_DAO.SqlDao.CMEPMSExecuteSql("DELETE FROM PJDFDDS WHERE PID='" + model.PID + "' AND SYR='" + model.SYRS[i] + "'");
                    break;
                }
            }
            ViewBag.msg = "已刪除";
            return Monthly(model);
        }

        //中央款比 
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Central")]
        public ActionResult Central(PJE0020Model model)
        {
            model.CurrentProjectName = WebSiteHelper.GetCurrentProjectName(0, model.PID); //取得目前所在的主計畫名稱(+分支計畫名稱)

            DataTable poop = SqlDao.CMBASEQuery4List("SELECT ''as SUP_RATIO, CITY FROM SMCPOSTN GROUP BY CITY");
            poop.Columns["SUP_RATIO"].ReadOnly = false;
            poop.Columns["SUP_RATIO"].MaxLength = 9999;
            //model.PJMFORM = new PJMFORM_CMEPMS();
            SqlCommand cmd = new SqlCommand("SELECT*FROM PJDCTSP p WHERE p.PID=@PID");
            //cmd.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            cmd.Parameters.Add(new SqlParameter("PID", tools.SqlFilterInt(model.PID)));
            DataTable ps = BASE_DAO.SqlDao.CMEPMSQuery4List(cmd);




            foreach (DataRow p in poop.Rows)
            {

                foreach (DataRow s in ps.Rows)
                {

                    if (Convert.ToString(p["CITY"]).Equals(s["CITY"]))
                    {
                        p["SUP_RATIO"] = s["SUP_RATIO"];
                        break;
                    }
                }



            }










            ViewBag.list = poop;

            return View("PJE0020Central", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CentralAdd")]
        public ActionResult CentralAdd(PJE0020Model model)
        {



            ViewBag.msg = "新增完成";
            return Central(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CentralSave")]
        public ActionResult CentralSave(PJE0020Model model, string[] SUP_RATIOS, string[] CITYS)
        {


            SqlDao.CMEPMSExecuteSql("DELETE FROM PJDCTSP WHERE PID='" + model.PID + "'");
            PJDCTSP j;
            DataRow p;
            for (int i = 0; i < CITYS.Length; i++)
            {
                //淦這個表不能查
                //p = db.PJDCTSP.Where(p => p.PID == model.PID && p.CITY == CITYS[i]).FirstOrDefault();
                p = SqlDao.CMEPMSQuery4Object("SELECT*FROM PJDCTSP WHERE PID='" + model.PID + "' AND CITY='" + CITYS[i] + "'");
                if (p == null)
                {
                    j = new PJDCTSP();
                    j.PID = model.PID;
                    j.CITY = CITYS[i];
                    j.SUP_RATIO = tools.safe2Decimal(SUP_RATIOS[i]);
                    db.PJDCTSP.Add(j);
                    db.SaveChanges();
                }
                else
                {
                    SqlDao.CMEPMSExecuteSql("UPDATE PJDCTSP SET SUP_RATIO='" + SUP_RATIOS[i] + "'  WHERE PID='" + model.PID + "' AND CITY='" + CITYS[i] + "'");
                }
            }



            ViewBag.msg = "已儲存";
            return Central(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "CentralDel")]
        public ActionResult CentralDel(PJE0020Model model)
        {



            ViewBag.msg = "已刪除";
            return Central(model);
        }




    }

}