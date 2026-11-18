using Aspose.Cells;
using Aspose.Words;
using Aspose.Words.Replacing;
using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.Models.CMBASEModels;
using WCals.Models.CMPSGSModels;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;
using static WCals.Controllers.NSPMIS.ExportSheet;
using WordsSaveFormat = Aspose.Words.SaveFormat;

namespace WCals.Controllers.NSPMIS
{
    //2.2.1
    public class SGE0000Controller : Controller
    {
        private readonly CMBASEEntities baseDb = new CMBASEEntities();
        private readonly Models.CMEPMSModels.CMEPMSEntities db = new Models.CMEPMSModels.CMEPMSEntities();
        private readonly Models.CMPSGSModels.CMPSGSEntities sgsdb = new Models.CMPSGSModels.CMPSGSEntities();
        int pageSize = 10;

        //01
        public ActionResult STEP01(decimal aid, string yrcid, string ppid, string dis)
        {
            if (dis != null)
            {
                ViewBag.dis = true;
                ViewBag.HasPermission = CheckUpdateTargetedPermission(aid);
            }

            SGE0000Model model = new SGE0000Model();





            PJMFORM_CMPSGS pjmform;
            if (aid != 0)
            {
                APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
                decimal id = apmform.PPID.Value;
                pjmform = sgsdb.PJMFORM.Where(w => w.PPID == id).FirstOrDefault();
                if (apmform.PJ_CATY_CODE == null)
                {
                    apmform.PJ_CATY_CODE = "CA";
                    model.APMFORM = apmform;
                }
                if (ppid == null) ppid = Convert.ToString(apmform.PPID);
                if (yrcid == null) yrcid = Convert.ToString(apmform.YRC_ID);
                model = SetApmform(aid);
                if (model.APMFORM.EXT_FG == true)
                {
                    model.EXT_FG = "Y";
                }
                else
                {
                    model.EXT_FG = "N";
                }
                ppid = apmform.PPID.ToString();
            }
            else
            {
                decimal id = Convert.ToDecimal(ppid);
                pjmform = sgsdb.PJMFORM.Where(w => w.PPID == id).FirstOrDefault();
                decimal yrc = Convert.ToDecimal(yrcid);
                model = new SGE0000Model();

                PJDSYRC pjdsyrc = sgsdb.PJDSYRC.Where(w => w.YRC_ID == yrc).FirstOrDefault();
                model.PJ_NAME = pjmform.PJ_NAME;
                model.PSG_NAME = pjdsyrc.PSG_NAME;
                model.AID = "";
                model.YRC_ID = yrcid;
                model.PPID = id;
                model.EXT_FG = "N";
                model.APMFORM = new APMFORM();
                string dept = WebSiteHelper.getUserDept();
                model.APMFORM.FILL_DP = dept;
                model.APMFORM.ADV_DP = dept;

                initAPMFORM(ppid, yrcid, null);

            }



            DataTable ogcdpur = SqlDao.CMBASEQuery4List("SELECT d.SMP_DEPT_NAME as CITY FROM OGCDPUR u, OGCDEPT d WHERE u.DEPT_NO=d.DEPT_NO AND u.USR_ID='" + WebSiteHelper.getUserId() + "'");
            StringBuilder sb = new StringBuilder("SELECT TD_CODE, t.TD_NAME as AP_NAME, t.TD_CITY as CITY FROM TDMFORM t WHERE t.TD_CITY IS NOT NULL AND t.TD_CITY IN(");

            // 2. 🎯 正確修正：撤除外層 foreach，改用 LINQ 一次性高效率處理整份 DataTable
            var cityValues = ogcdpur.AsEnumerable().Select(row =>
            {
                // 自動防禦資料庫 DBNull
                string city = row.Field<string>("CITY");
                // 若無值則遞補指派 model 的預設縣市，並自動包覆 SQL 單引號
                return !string.IsNullOrWhiteSpace(city) ? $"'{city}'" : $"'{model.APMFORM.CITY}'";
            });

            // 3. 🎯 將用逗號串接好的乾淨字串，直接灌進您既有的 sb 容器中
            sb.Append(string.Join(",", cityValues));

            //sb.Length--;
            sb.Append(")");//延續性工程，舊系統只會顯示該帳號對應的縣市的資料
            ViewBag.allap = SqlDao.CMETMSQuery4List(sb.ToString());
            try
            {
                ViewBag.myCity = ogcdpur.Rows[0]["CITY"];//舊系統會先預設該帳號對應的縣市→新系統需加入此功能

            }
            catch
            {

            }
            //R的情況
            if (pjmform.STAT_NO.Equals("R"))
            {
                //R的標籤以判斷:自評分數出現
                ViewBag.isR = true;
                //案件類別抓取TT 04
                ViewBag.PJ_CATY_CODE = tools.ToSelectList(util.GetCMPSGSCODES("TT", "04"), "id", "name", null);
                if (model.APMFORM.PJ_CATY_CODE == null)
                {
                    //model.APMFORM = new APMFORM();
                    model.APMFORM.PJ_CATY_CODE = "CA";
                }
            }

            // 環境改善計畫
            if (pjmform.STAT_NO.Equals("E"))
            {
                ViewBag.isE = true;
                ViewBag.PJ_CATY_CODE = tools.ToSelectList(util.GetCMPSGSCODES("TT", "05"), "id", "name", null);
                if (model.APMFORM.PJ_CATY_CODE == null)
                {
                    model.APMFORM.PJ_CATY_CODE = "EA";
                }
            }

            return View("STEP01", model);
        }


        public ActionResult STEP1Error(SGE0000Model model)
        {
            initAPMFORM(model.PPID.ToString(), model.YRC_ID, null);
            ViewBag.allap = SqlDao.CMPSGSQuery4List("SELECT TD_CODE, CITY, AP_NAME FROM APMFORM");
            return View("STEP01", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "STEP01Save")]
        public ActionResult STEP01Save(SGE0000Model model)
        {
            //decimal aid = Convert.ToDecimal(model.AID);
            decimal aid = string.IsNullOrEmpty(model.AID) ? 0 : Convert.ToDecimal(model.AID);
            /*if(string.IsNullOrEmpty(model.APMFORM.PJ_CATY_CODE) ||
                string.IsNullOrEmpty(model.APMFORM.AP_NAME) ||
                string.IsNullOrEmpty(model.APMFORM.CITY) ||
                string.IsNullOrEmpty(model.APMFORM.AP_PJ_NAME) ||
                //string.IsNullOrEmpty(model.APMFORM.FILL_MCH) ||
                //string.IsNullOrEmpty(model.APMFORM.EXE_MCH) ||
                string.IsNullOrEmpty(model.APMFORM.CONT_NAME) ||
                string.IsNullOrEmpty(model.APMFORM.CONT_EMAIL) ||
                model.START_DT == null ||
                model.END_DT == null
                ) {

                ViewBag.msg = "*為必填項目";

                return STEP1Error(model);
            }*/

            DateTime now = DateTime.Now;
            DateTime tmp;
            string userid = WebSiteHelper.getUserId();
            string userdp = WebSiteHelper.getUserDept();

            APMFORM apmform;
            if (aid == 0)
            {
                apmform = new APMFORM();
                apmform.AP_STAT = "-1";
                apmform.ADD_DT = now;
                apmform.ADV_UID = userid;
                apmform.ADV_DP = userdp;
            }
            else
            {
                apmform = sgsdb.APMFORM.Where(x => x.AID == aid).FirstOrDefault();
                apmform.UPD_DT = now;
                apmform.UPD_UID = userid;
            }
            decimal yrcid = tools.safe2Decimal(model.YRC_ID);
            PJDSYRC syrc = sgsdb.PJDSYRC.Where(x => x.YRC_ID == yrcid).FirstOrDefault();
            apmform.SYR = syrc.SYR;
            apmform.PPID = model.PPID;
            apmform.YRC_ID = tools.safe2Decimal(model.YRC_ID);
            apmform.PJ_CATY_CODE = model.APMFORM.PJ_CATY_CODE;
            apmform.AP_NAME = model.APMFORM.AP_NAME;
            apmform.CITY = model.APMFORM.CITY;
            //apmform.AP_PJ_NAME = model.APMFORM.AP_PJ_NAME;
            if (model.APMFORM.AP_PJ_NAME != null)
            {
                apmform.AP_PJ_NAME = model.APMFORM.AP_PJ_NAME;
            }
            apmform.FILL_DP = model.APMFORM.FILL_DP;
            apmform.EXE_MCH = model.APMFORM.EXE_MCH;

            apmform.NPSD_SYR = model.APMFORM.NPSD_SYR;

            //apmform.TOWN = model.APMFORM.TOWN;


            apmform.CONT_TITLE = model.APMFORM.CONT_TITLE;
            apmform.CONT_TEL = model.APMFORM.CONT_TEL;

            //if (model.EXT_FG.Equals("Y"))
            //{
            //    apmform.EXT_FG = true;
            //}
            //else
            //{
            //    apmform.EXT_FG = false;
            //}

            string extFgValue = Request.Form["EXT_FG"];
            if (!string.IsNullOrEmpty(extFgValue))
            {
                apmform.EXT_FG = (extFgValue == "Y");
            }
            // 若原本有綁定，做雙重防護
            else if (!string.IsNullOrEmpty(model.EXT_FG))
            {
                apmform.EXT_FG = (model.EXT_FG == "Y");
            }
            //apmform.TD_CODE = model.APMFORM.TD_CODE;
            if (model.APMFORM.TD_CODE != null)
            {
                apmform.TD_CODE = model.APMFORM.TD_CODE;
            }
            apmform.AP_ALL_NOTE = model.APMFORM.AP_ALL_NOTE;
            apmform.BENEFITS = model.APMFORM.BENEFITS;
            apmform.AP_YEAR_NOTE = model.APMFORM.AP_YEAR_NOTE;
            apmform.AP_NOTE = model.APMFORM.AP_NOTE;
            apmform.CONT_NAME = model.APMFORM.CONT_NAME;
            apmform.CONT_EMAIL = model.APMFORM.CONT_EMAIL;
            apmform.CONT_FAX = model.APMFORM.CONT_FAX;
            apmform.SYS_CATY = model.APMFORM.SYS_CATY;
            apmform.IPJ_CODE = model.IPJ_CODE;

            if (!string.IsNullOrEmpty(model.START_DT))
            {
                try
                {
                    tmp = tools.GetDateByTWDate(model.START_DT);
                }
                catch
                {
                    ViewBag.msg = "開始日期有誤";
                    return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), null);
                }
                apmform.START_DT = tmp;
            }
            if (!string.IsNullOrEmpty(model.END_DT))
            {
                try
                {
                    tmp = tools.GetDateByTWDate(model.END_DT);
                }
                catch
                {
                    ViewBag.msg = "結束日期有誤";
                    return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), null);
                }
                apmform.END_DT = tmp;
            }
            if (!string.IsNullOrEmpty(model.TF106_DT))
            {
                try
                {
                    tmp = tools.GetDateByTWDate(model.TF106_DT);
                }
                catch
                {
                    ViewBag.msg = "預定徵收公告日有誤";
                    model.TF106_DT = "";
                    return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), null);
                }
                apmform.TF106_DT = tmp;
            }
            if (!string.IsNullOrEmpty(model.TF108_DT))
            {
                try
                {
                    tmp = tools.GetDateByTWDate(model.TF108_DT);
                }
                catch
                {
                    ViewBag.msg = "預定用地取得完成日有誤";
                    model.TF108_DT = "";
                    return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), null);
                }
                apmform.TF108_DT = tmp;
            }
            if (!string.IsNullOrEmpty(model.TK103_DT))
            {
                try
                {
                    tmp = tools.GetDateByTWDate(model.TK103_DT);
                }
                catch
                {
                    ViewBag.msg = "預定設計完成日有誤";
                    model.TK103_DT = "";
                    return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), null);
                }
                apmform.TK103_DT = tmp;
            }
            if (!string.IsNullOrEmpty(model.TK105_DT))
            {
                try
                {
                    tmp = tools.GetDateByTWDate(model.TK105_DT);
                }
                catch
                {
                    ViewBag.msg = "預定發包日有誤";
                    model.TK105_DT = "";
                    return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), null);
                }
                apmform.TK105_DT = tmp;
            }
            if (!string.IsNullOrEmpty(model.TK106_DT))
            {
                try
                {
                    tmp = tools.GetDateByTWDate(model.TK106_DT);
                }
                catch
                {
                    ViewBag.msg = "預定開工日有誤";
                    model.TK106_DT = "";
                    return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), null);
                }
                apmform.TK106_DT = tmp;
            }
            if (!string.IsNullOrEmpty(model.TK107_DT))
            {
                try
                {
                    tmp = tools.GetDateByTWDate(model.TK107_DT);
                }
                catch
                {
                    ViewBag.msg = "預定完工日有誤";
                    model.TK107_DT = "";
                    return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), null);
                }
                apmform.TK107_DT = tmp;
            }

            if (!string.IsNullOrEmpty(model.APMFORM.RSYS_NAME))
            {
                apmform.RSYS_NAME = model.APMFORM.RSYS_NAME;
            }


            apmform.SYS_CATY = model.APMFORM.SYS_CATY;
            apmform.S1_ORDER = model.APMFORM.S1_ORDER;
            apmform.ATTENTION_PEOPLE = model.APMFORM.ATTENTION_PEOPLE;


            //新增判斷
            apmform.WSYS_NAME = null;
            apmform.IMPALLAREA = null;
            apmform.IMPFLDAREA = null;
            apmform.PRTPOPS = null;
            apmform.PLR_AMT = null;
            apmform.REXM_PRJ = null;
            apmform.NPSD_SYR = null;
            apmform.OPSD_SYR = null;
            apmform.PPSD_SYR = null;


            if (model.APMFORM.WSYS_NAME != null) apmform.WSYS_NAME = model.APMFORM.WSYS_NAME;
            if (model.APMFORM.IMPALLAREA != null) apmform.IMPALLAREA = model.APMFORM.IMPALLAREA;
            if (model.APMFORM.IMPFLDAREA != null) apmform.IMPFLDAREA = model.APMFORM.IMPFLDAREA;
            if (model.APMFORM.PRTPOPS != null) apmform.PRTPOPS = model.APMFORM.PRTPOPS;            
            if (model.APMFORM.NPSD_SYR != null) apmform.NPSD_SYR = model.APMFORM.NPSD_SYR;
            if (model.APMFORM.OPSD_SYR != null) apmform.OPSD_SYR = model.APMFORM.OPSD_SYR;
            if (model.APMFORM.PPSD_SYR != null) apmform.PPSD_SYR = model.APMFORM.PPSD_SYR;

            if (model.APMFORM.REXM_PRJ != null) apmform.REXM_PRJ = model.APMFORM.REXM_PRJ;

            // 環境改善計畫
            if (model.APMFORM.PJ_CATY_CODE != null && model.APMFORM.PJ_CATY_CODE.StartsWith("E"))
            {
                apmform.IMPALLAREA = model.APMFORM.IMPALLAREA;
                apmform.PRTPOPS = model.APMFORM.PRTPOPS;
                apmform.COMP_AGENCY = model.APMFORM.COMP_AGENCY;
                apmform.PLR_AMT = model.APMFORM.PLR_AMT;
            }

            if (aid == 0)
            {
                //新建狀況
                sgsdb.APMFORM.Add(apmform);
                sgsdb.SaveChanges();
                //建立預設預算來源提示項目(無金額)
                util.UPDATE_APMFORM_Log(apmform.AID, "", apmform.AP_STAT, "NW", now, userid, WebSiteHelper.getUserDept(), "");
                aid = apmform.AID;
                //APDFUND
                PJDSYRC PJDSYRC = sgsdb.PJDSYRC.Where(w => w.YRC_ID == apmform.YRC_ID).FirstOrDefault();

                APDFUND apdfund = new APDFUND();
                apdfund.AID = aid;
                apdfund.SYR = PJDSYRC.SYR;
                apdfund.FIT_CODE = "00";
                apdfund.UPD_UID = userid;
                apdfund.UPD_DT = now;
                apdfund.ADD_DT = now;
                sgsdb.APDFUND.Add(apdfund);
                sgsdb.SaveChanges();

                apdfund = new APDFUND();
                apdfund.AID = aid;
                apdfund.SYR = PJDSYRC.SYR;
                apdfund.FIT_CODE = "11";
                apdfund.UPD_UID = userid;
                apdfund.UPD_DT = now;
                apdfund.ADD_DT = now;
                sgsdb.APDFUND.Add(apdfund);
                try
                {
                    sgsdb.SaveChanges();
                    ViewBag.msg = "已儲存";
                }
                catch (Exception e)
                {
                    ViewBag.msg = "儲存失敗";
                    System.Diagnostics.Debug.WriteLine(e.Message);
                }

            }
            else
            {
                //單純儲存
                try
                {
                    sgsdb.SaveChanges();
                    ViewBag.msg = "已儲存";
                }
                catch (Exception e)
                {
                    ViewBag.msg = "儲存失敗";
                    System.Diagnostics.Debug.WriteLine(e.Message);
                }
            }

            // 💡 取得前端送過來的 dis 狀態
            string dis = Request.Form["dis"];

            // 💡 將 dis 狀態原封不動地傳回給 STEP01，維持唯讀鎖定！
            return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), dis);
            //return STEP01(apmform.AID, Convert.ToString(apmform.YRC_ID), Convert.ToString(apmform.PPID), null);
        }

        /// <summary>
        /// 更新案件類別
        /// </summary>
        /// <param name="aid"> 編號 </param>
        /// <param name="code"> 案件類別 </param>
        /// <returns></returns>
        [HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult UpdateSTEP01TargetedItem(decimal id, string code)
        //{
        //    try
        //    {
        //        var apmform = sgsdb.APMFORM.Where(x => x.AID == id).FirstOrDefault();
        //        if (apmform == null)
        //        {
        //            return Json(new { success = false, message = "找不到資料" });
        //        }

        //        if (!CheckUpdateTargetedPermission(id))
        //        {
        //            return Json(new { success = false, message = "您沒有權限編輯此資料" });
        //        }

        //        apmform.PJ_CATY_CODE = code;
        //        apmform.UPD_DT = DateTime.Now;
        //        apmform.UPD_UID = WebSiteHelper.getUserId();
        //        sgsdb.SaveChanges();
        //        return Json(new { success = true, message = "更新成功" });
        //    }
        //    catch (Exception e)
        //    {
        //        System.Diagnostics.Debug.WriteLine(e.Message);
        //        return Json(new { success = false, message = "更新失敗" });
        //    }
        //}

        /// <summary>
        /// 驗證使用者是否有權限更新STEP01特定欄位(案件類別)
        /// </summary>
        /// <param name="aid"> 編號 </param>
        /// <returns></returns>
        public bool CheckUpdateTargetedPermission(decimal aid)
        {
            try
            {
                var apmform = sgsdb.APMFORM.Where(x => x.AID == aid).FirstOrDefault();
                if (apmform == null)
                {
                    return false;
                }

                var validStatus = new HashSet<string> { "10", "15", "20" };

                var apStat = apmform.AP_STAT;

                var userDept = WebSiteHelper.getUserDept();
                var usergroup = baseDb.SMCAUTH.Where(s => s.DEPT_NO == userDept).Select(s => s.SYS_ID).ToList();

                var permissionMap = new Dictionary<string, string>
                                    {
                                        { "10", "CMPSGS.SGE0011" }, // 初審中
                                        { "15", "CMPSGS.SGE0061" }, // 複審中
                                        { "20", "CMPSGS.PJE0011" } // 核定中
                                    };
                string required;
                bool hasPermission = permissionMap.TryGetValue(apStat, out required) && usergroup.Contains(required);

                bool canUpdate = validStatus.Contains(apStat) && hasPermission;

                return canUpdate;
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
                return false;
            }
        }

        public ActionResult STEP07(decimal aid, string dis)
        {
            if (dis != null)
            {
                ViewBag.dis = true;
            }
            SGE0000Model model = SetApmform(aid);

            //處理R
            APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            if (apmform.PJ_CATY_CODE == null)
            {
                apmform.PJ_CATY_CODE = "CA";
                model.APMFORM = apmform;
            }
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == apmform.PPID).FirstOrDefault();
            //R的情況
            if (pjmform.STAT_NO.Equals("R"))
            {

                ViewBag.isR = true;

            }
            if (pjmform.STAT_NO.Equals("E"))
            {
                ViewBag.isE = true;
            }

            return View("STEP07", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "STEP07Save")]
        public ActionResult STEP07Save(SGE0000Model model)
        {
            decimal aid = Convert.ToDecimal(model.AID);
            APMFORM apmform = sgsdb.APMFORM.Where(x => x.AID == aid).FirstOrDefault();
            //各項評比清空
            apmform.SC1_SCORE1 = null;
            apmform.SC1_CMT1 = null;
            apmform.SC2_SCORE1 = null;
            apmform.SC2_CMT1 = null;
            apmform.SC3_SCORE1 = null;
            apmform.SC3_CMT1 = null;
            apmform.SC4_SCORE1 = null;
            apmform.SC4_CMT1 = null;
            apmform.SC5_SCORE1 = null;
            apmform.SC5_CMT1 = null;
            apmform.SC6_SCORE1 = null;
            apmform.SC6_CMT1 = null;
            apmform.SC7_SCORE1 = null;
            apmform.SC7_CMT1 = null;
            apmform.SC8_SCORE1 = null;
            apmform.SC8_CMT1 = null;

            apmform.SC9_SCORE1 = null;
            apmform.SC9_CMT1 = null;
            apmform.SCS_CMT1 = null;

            //輸入各項評比
            apmform.SC1_SCORE1 = model.APMFORM.SC1_SCORE1;
            apmform.SC1_CMT1 = model.APMFORM.SC1_CMT1;
            apmform.SC2_SCORE1 = model.APMFORM.SC2_SCORE1;
            apmform.SC2_CMT1 = model.APMFORM.SC2_CMT1;
            apmform.SC3_SCORE1 = model.APMFORM.SC3_SCORE1;
            apmform.SC3_CMT1 = model.APMFORM.SC3_CMT1;
            apmform.SC4_SCORE1 = model.APMFORM.SC4_SCORE1;
            apmform.SC4_CMT1 = model.APMFORM.SC4_CMT1;
            apmform.SC5_SCORE1 = model.APMFORM.SC5_SCORE1;
            apmform.SC5_CMT1 = model.APMFORM.SC5_CMT1;
            apmform.SC6_SCORE1 = model.APMFORM.SC6_SCORE1;
            apmform.SC6_CMT1 = model.APMFORM.SC6_CMT1;
            apmform.SC7_SCORE1 = model.APMFORM.SC7_SCORE1;
            apmform.SC7_CMT1 = model.APMFORM.SC7_CMT1;
            apmform.SC8_SCORE1 = model.APMFORM.SC8_SCORE1;
            apmform.SC8_CMT1 = model.APMFORM.SC8_CMT1;
            apmform.SC9_SCORE1 = model.APMFORM.SC9_SCORE1;
            apmform.SC9_CMT1 = model.APMFORM.SC9_CMT1;
            apmform.SCS_CMT1 = model.APMFORM.SCS_CMT1;
            sgsdb.SaveChanges();
            ViewBag.msg = "儲存完成";
            return STEP07(aid, null);
        }

        public ActionResult STEP02(decimal aid, string dis)
        {
            if (dis != null)
            {
                ViewBag.dis = true;
            }
            SGE0000Model model = SetApmform(aid);

            //處理R
            APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            if (apmform.PJ_CATY_CODE == null)
            {
                apmform.PJ_CATY_CODE = "CA";
                model.APMFORM = apmform;
            }
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == apmform.PPID).FirstOrDefault();
            //R的情況
            if (pjmform.STAT_NO.Equals("R"))
            {
                ViewBag.isR = true;
            }
            if (pjmform.STAT_NO.Equals("E"))
            {
                ViewBag.isE = true;
            }

            //縣市
            ViewBag.towns1 = util.GetTownNew(model.APMFORM.CITY);
            return View("STEP02", model);
        }
        //02
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "STEP02Save")]
        public ActionResult STEP02Save(SGE0000Model model)
        {
            decimal aid = Convert.ToDecimal(model.AID);

            if (model.TOWN != null)
            {
                APDLOCP apdlocp;
                SqlDao.CMPSGSExecuteSql("DELETE FROM APDLOCP WHERE AID=" + aid);
                for (int i = 0; i < model.LID.Length; i++)
                {
                    if (string.IsNullOrEmpty(model.TOWN[i]))
                    {
                        continue;
                    }
                    apdlocp = new APDLOCP();
                    apdlocp.LID = DateTime.Now.Ticks;
                    apdlocp.AID = aid;
                    apdlocp.TOWN = model.TOWN[i];
                    apdlocp.POS_SYB = model.POS_SYB[i];
                    apdlocp.GIS_SYS = model.GIS_SYS[i];
                    apdlocp.GIS_X = model.GIS_X[i];
                    apdlocp.GIS_Y = model.GIS_Y[i];
                    apdlocp.GIS_Z = model.GIS_Z[i];
                    apdlocp.UPD_UID = WebSiteHelper.getUserId(); ;
                    apdlocp.ADD_DT = DateTime.Now;
                    sgsdb.APDLOCP.Add(apdlocp);
                    sgsdb.SaveChanges();
                }
            }
            ViewBag.msg = "儲存完成";
            return STEP02(aid, null);
        }
        //03
        public ActionResult STEP03(decimal aid, string dis)
        {
            if (dis != null)
            {
                ViewBag.dis = true;
            }
            SGE0000Model model = SetApmform(aid);

            //處理R
            APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            if (apmform.PJ_CATY_CODE == null)
            {
                apmform.PJ_CATY_CODE = "CA";
                model.APMFORM = apmform;
            }
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == apmform.PPID).FirstOrDefault();
            //R的情況
            if (pjmform.STAT_NO.Equals("R"))
            {

                ViewBag.isR = true;

            }
            if (pjmform.STAT_NO.Equals("E"))
            {
                ViewBag.isE = true;
            }

            return View("STEP03", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "STEP03Save")]
        public ActionResult STEP03Save(SGE0000Model model)
        {
            decimal aid = Convert.ToDecimal(model.AID);
            SqlDao.CMPSGSExecuteSql("DELETE FROM APDWQTY WHERE AID=" + aid);


            APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();

            APDWQTY apdwqty;
            for (int i = 0; i < model.MT_LID.Length; i++)
            {
                if (string.IsNullOrEmpty(model.MT_DESC[i]) ||
                    //string.IsNullOrEmpty(model.MT_UNIT[i]) ||
                    string.IsNullOrEmpty(model.MT_QTY[i]))
                {
                    continue;
                }
                apdwqty = new APDWQTY();
                apdwqty.AID = aid;
                apdwqty.LID = tools.safe2Decimal(model.MT_LID[i]);
                apdwqty.MID = (i + 1);
                apdwqty.MT_DESC = model.MT_DESC[i];
                apdwqty.MT_NAME = model.MT_NAME[i];

                try
                {
                    apdwqty.MT_UNIT = SqlDao.CMPSGSQuery4Object("SELECT CMS_UNIT FROM PJDCMCT WHERE PPID=" + apmform.PPID + " AND CMS_CODE='" + model.MT_NAME[i] + "'")["CMS_UNIT"].ToString();
                }
                catch
                {
                    apdwqty.MT_UNIT = "";
                }

                apdwqty.MT_SPEC = model.MT_SPEC[i];
                apdwqty.MT_QTY = tools.safe2Decimal(model.MT_QTY[i]);
                apdwqty.ADD_DT = DateTime.Now;
                apdwqty.UPD_UID = WebSiteHelper.getUserId();
                sgsdb.APDWQTY.Add(apdwqty);
            }
            sgsdb.SaveChanges();
            ViewBag.msg = "已儲存";
            return STEP03(aid, null);
        }

        //04
        public ActionResult STEP04(decimal aid, string dis)
        {
            if (dis != null)
            {
                ViewBag.dis = true;
            }
            SGE0000Model model = SetApmform(aid);

            //處理R
            APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            if (apmform.PJ_CATY_CODE == null)
            {
                apmform.PJ_CATY_CODE = "CA";
                model.APMFORM = apmform;
            }
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == apmform.PPID).FirstOrDefault();
            //R的情況
            if (pjmform.STAT_NO.Equals("R"))
            {

                ViewBag.isR = true;

            }
            if (pjmform.STAT_NO.Equals("E"))
            {
                ViewBag.isE = true;
            }
            return View("STEP04", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "STEP04Save")]
        public ActionResult STEP04Save(SGE0000Model model)
        {
            decimal aid = Convert.ToDecimal(model.AID);
            APMFORM apmform = sgsdb.APMFORM.Where(x => x.AID == aid).FirstOrDefault();
            //評估事項
            apmform.SWW1_FG = ConvertYN(model.SWW1_FG);
            apmform.SWW2_FG = ConvertYN(model.SWW2_FG);
            apmform.SWW3_FG = ConvertYN(model.SWW3_FG);
            apmform.SWW4_FG = ConvertYN(model.SWW4_FG);
            apmform.SWW5_FG = ConvertYN(model.SWW5_FG);
            apmform.SWW6_FG = ConvertYN(model.SWW6_FG);
            apmform.SWW7_FG = ConvertYN(model.SWW7_FG);
            apmform.SWW8_FG = ConvertYN(model.SWW8_FG);
            apmform.SWW81_FG = ConvertYN(model.SWW81_FG);
            if (!string.IsNullOrEmpty(model.SWW9_FG)) apmform.SWW9_RATE = tools.safe2Decimal(model.SWW9_FG);
            apmform.SWW10_FG = ConvertYN(model.SWW10_FG);
            apmform.SWW11_FG = ConvertYN(model.SWW11_FG);
            apmform.SWWOT_DESC = model.SWWOT_DESC;
            apmform.UPD_DT = DateTime.Now;
            apmform.UPD_UID = WebSiteHelper.getUserId();





            sgsdb.SaveChanges();
            ViewBag.msg = "已儲存";
            return STEP04(aid, null);
        }

        /// <summary>
        /// 勘察評估事項轉換Y/N/U
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool? ConvertYN(string value)
        {
            return value == "Y" ? true :
                   value == "N" ? false :
                   (bool?)null;
        }

        //05
        public ActionResult STEP05(decimal aid, string dis)
        {
            if (dis != null)
            {
                ViewBag.dis = true;
            }
            SGE0000Model model = SetApmform(aid);
            APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            if (apmform.PJ_CATY_CODE == null)
            {
                apmform.PJ_CATY_CODE = "CA";
                model.APMFORM = apmform;
            }
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == apmform.PPID).FirstOrDefault();
            //R的情況
            if (pjmform.STAT_NO.Equals("R"))
            {
                ViewBag.isR = true;
            }
            if (pjmform.STAT_NO.Equals("E"))
            {
                ViewBag.isE = true;
            }
            //0605部份金額需以千元顯示
            model.APMFORM.TCEN_FUND = model.APMFORM.TCEN_FUND; // 1000;
            model.APMFORM.TLOC_FUND = model.APMFORM.TLOC_FUND; // 1000;
            model.APMFORM.RCEN_FUND = model.APMFORM.RCEN_FUND; // 1000;
            model.APMFORM.RLOC_FUND = model.APMFORM.RLOC_FUND; // 1000;
            model.APMFORM.FCEN_FUND = model.APMFORM.FCEN_FUND; // 1000;
            model.APMFORM.FLOC_FUND = model.APMFORM.FLOC_FUND; // 1000;

            return View("STEP05", model);
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "STEP05Del")]
        public ActionResult STEP05Del(SGE0000Model model)
        {
            //decimal aid = Convert.ToDecimal(model.AID);
            decimal aid = tools.safe2Decimal(model.AID);
            if (aid == 0) return STEP05(aid, null);

            for (int i = 0; i < model.DEL_FG.Length; i++)
            {
                if (!string.IsNullOrEmpty(model.DEL_FG[i]))
                {
                    SqlDao.CMPSGSExecuteSql("DELETE FROM APDFUND WHERE FID=" + model.APDFUND_FID[i]);
                }
            }

            APMFORM apmform = sgsdb.APMFORM.Where(x => x.AID == aid).FirstOrDefault();
            apmform.TCEN_FUND = model.APMFORM.TCEN_FUND;//1000;
            apmform.TLOC_FUND = model.APMFORM.TLOC_FUND;//1000;
            apmform.RCEN_FUND = model.APMFORM.RCEN_FUND;//1000;
            apmform.RLOC_FUND = model.APMFORM.RLOC_FUND;//*1000;
            apmform.FCEN_FUND = null;
            apmform.FLOC_FUND = null;
            if (model.APMFORM.FCEN_FUND != null) apmform.FCEN_FUND = model.APMFORM.FCEN_FUND;//1000;
            if (model.APMFORM.FLOC_FUND != null) apmform.FLOC_FUND = model.APMFORM.FLOC_FUND;//*1000;
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            string userdp = WebSiteHelper.getUserDept();
            apmform.UPD_DT = now;
            apmform.ADV_UID = userid;
            apmform.ADV_DP = userdp;
            sgsdb.SaveChanges();

            ViewBag.msg = "已刪除";
            return STEP05(aid, null);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "STEP05Save")]
        public ActionResult STEP05Save(SGE0000Model model)
        {
            decimal aid = Convert.ToDecimal(model.AID);
            string userid = WebSiteHelper.getUserId();
            string userdp = WebSiteHelper.getUserDept();
            DateTime now = DateTime.Now;

            APMFORM apmform = sgsdb.APMFORM.Where(x => x.AID == aid).FirstOrDefault();
            apmform.TCEN_FUND = model.APMFORM.TCEN_FUND;//1000;
            apmform.TLOC_FUND = model.APMFORM.TLOC_FUND;//1000;
            apmform.RCEN_FUND = model.APMFORM.RCEN_FUND;//1000;
            apmform.RLOC_FUND = model.APMFORM.RLOC_FUND;//*1000;


            apmform.FCEN_FUND = null;
            apmform.FLOC_FUND = null;
            if (model.APMFORM.FCEN_FUND != null) apmform.FCEN_FUND = model.APMFORM.FCEN_FUND;//1000;
            if (model.APMFORM.FLOC_FUND != null) apmform.FLOC_FUND = model.APMFORM.FLOC_FUND;//*1000;

            apmform.UPD_DT = now;
            apmform.ADV_UID = userid;
            apmform.ADV_DP = userdp;
            sgsdb.SaveChanges();

            decimal fid;
            if (model.SYRS != null)
            {
                APDFUND apdfund;
                for (int i = 0; i < model.SYRS.Length; i++)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(model.SYRS[i]) ||
                        string.IsNullOrEmpty(model.FUND[i])
                        )
                        {
                            continue;
                        }

                        if (string.IsNullOrEmpty(model.APDFUND_FID[i]))
                        {
                            apdfund = new APDFUND();
                        }
                        else
                        {
                            fid = tools.safe2Decimal(model.APDFUND_FID[i]);
                            apdfund = sgsdb.APDFUND.Where(w => w.FID == fid).FirstOrDefault();
                        }
                        apdfund.FID = tools.safe2Decimal(model.APDFUND_FID[i]);
                        apdfund.AID = apmform.AID;
                        apdfund.SYR = model.SYRS[i];
                        apdfund.FUND = tools.safe2Decimal(model.FUND[i]);// *1000;
                        apdfund.FIT_CODE = model.FIT_CODE[i];
                        apdfund.FUND_DESC = model.FUND_DESC[i];
                        apdfund.UPD_UID = userid;
                        apdfund.UPD_DT = now;
                        apdfund.ADD_DT = now;

                        if (string.IsNullOrEmpty(model.APDFUND_FID[i]))
                        {
                            sgsdb.APDFUND.Add(apdfund);
                        }
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Debug.WriteLine(e.StackTrace);
                    }


                }
                sgsdb.SaveChanges();
            }


            ViewBag.msg = "已儲存";
            return STEP05(aid, null);
        }

        //06
        public ActionResult STEP06(decimal aid, string dis)
        {
            if (dis != null)
            {
                ViewBag.dis = true;
            }
            SGE0000Model model = SetApmform(aid);
            APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            if (apmform.PJ_CATY_CODE == null)
            {
                apmform.PJ_CATY_CODE = "CA";
                model.APMFORM = apmform;
            }
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == apmform.PPID).FirstOrDefault();
            //R的情況
            if (pjmform.STAT_NO.Equals("R"))
            {
                ViewBag.isR = true;
            }
            if (pjmform.STAT_NO.Equals("E"))
            {
                ViewBag.isE = true;
            }
            return View("STEP06", model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "STEP06Save")]
        public ActionResult STEP06Save(SGE0000Model model)
        {
            decimal aid = Convert.ToDecimal(model.AID);
            if (model.files != null)
            {
                APDFILE apdfile;
                byte[] data;
                for (int i = 0; i < model.files.Length; i++)
                {
                    //新建
                    if (string.IsNullOrEmpty(model.FID[i]))
                    {
                        if (model.files[i] == null)
                        {
                            //沒檔案
                            continue;
                        }
                        apdfile = new APDFILE();
                        using (Stream inputStream = model.files[i].InputStream)
                        {
                            MemoryStream memoryStream = inputStream as MemoryStream;
                            if (memoryStream == null)
                            {
                                memoryStream = new MemoryStream();
                                inputStream.CopyTo(memoryStream);
                            }
                            data = memoryStream.ToArray();
                        }
                        apdfile.FILE_BODY = data;
                        apdfile.FID = DateTime.Now.Ticks;
                        apdfile.AID = aid;
                        apdfile.FILE_MEMO = model.FILE_MEMO[i];
                        apdfile.FC_DESC = model.FC_DESC[i];
                        apdfile.FILE_PATH = model.files[i].FileName;
                        apdfile.UPD_DT = DateTime.Now;
                        apdfile.UPD_UID = WebSiteHelper.getUserId();
                        sgsdb.APDFILE.Add(apdfile);
                        sgsdb.SaveChanges();

                    }
                    else
                    {
                        //舊有
                        decimal fid = Convert.ToDecimal(model.FID[i]);
                        apdfile = sgsdb.APDFILE.Where(w => w.FID == fid).FirstOrDefault();

                        apdfile.AID = aid;
                        if (model.files[i] != null)
                        {
                            //apdfile = new APDFILE();
                            using (Stream inputStream = model.files[i].InputStream)
                            {
                                MemoryStream memoryStream = inputStream as MemoryStream;
                                if (memoryStream == null)
                                {
                                    memoryStream = new MemoryStream();
                                    inputStream.CopyTo(memoryStream);
                                }
                                data = memoryStream.ToArray();
                            }
                            apdfile.FILE_BODY = data;
                            apdfile.FILE_PATH = model.files[i].FileName;
                        }



                        apdfile.AID = aid;
                        apdfile.FILE_MEMO = model.FILE_MEMO[i];
                        apdfile.FC_DESC = model.FC_DESC[i];

                        apdfile.UPD_DT = DateTime.Now;
                        apdfile.UPD_UID = WebSiteHelper.getUserId();
                        sgsdb.SaveChanges();
                    }
                }

            }
            ViewBag.msg = "已儲存";
            return STEP06(aid, null);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "STEP06Del")]
        public ActionResult STEP06Del(SGE0000Model model)
        {
            decimal aid = Convert.ToDecimal(model.AID);

            for (int i = 0; i < model.DEL_FG.Length; i++)
            {
                if (!string.IsNullOrEmpty(model.DEL_FG[i]))
                {
                    SqlDao.CMPSGSExecuteSql("DELETE FROM APDFILE WHERE FID=" + model.FID[i]);
                }
            }
            ViewBag.msg = "已刪除";
            return STEP06(aid, null);
        }

        public ActionResult Index(SGE0000Model model)
        {
            string userid = Helpers.WebSiteHelper.getUserId();
            string userdp = Helpers.WebSiteHelper.getUserDept();
            int page = (model.page == 0) ? 1 : model.page;
            string sql = "select F.STAT_NO, S.YRC_ID, F.PPID, F.PJ_NAME, S.PSG_NAME, S.SYR, S.BGN_SYR, " +
               "S.END_SYR, D.DEPT_NAME, U.USR_NAME," +
               "(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='-1' AND A.ADV_DP='" + userid + "') as STEP1c," +
               "(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='00' AND A.ADV_DP='" + userid + "') as STEP00c," +
               "(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='10' AND A.ADV_DP='" + userid + "') as STEP10c," +
               "(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='15' AND A.ADV_DP='" + userid + "') as STEP15c," +
               "(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='20' AND A.ADV_DP='" + userid + "') as STEP20c," +
               "(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='90' AND A.ADV_DP='" + userid + "') as STEPZZc," +
               "(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='Z1' AND A.ADV_DP='" + userid + "') as STEPZ1c," +
               "(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='Z2' AND A.ADV_DP='" + userid + "') as STEPZ2c " +
               "from PJMFORM F left join PJDSYRC S on F.PPID = S.PPID left join CMBASE.dbo.OGCDEPT D on F.ADV_DP = D.DEPT_NO left join " +
            "CMBASE.dbo.OGCUSER U on F.ADV_UID= U.USR_ID where S.OPEN_FG=1 order by F.PPID, S.SYR Desc, PSG_NAME";
            // TODO：AND F.PPID >= 21 是目前正式環境開放給雨水科測試用，之後要還原成上面這段 // 20251210
            //"CMBASE.dbo.OGCUSER U on F.ADV_UID= U.USR_ID where S.OPEN_FG=1 AND F.PPID >= 21 order by F.PPID, S.SYR Desc, PSG_NAME";

            DataTable dt = BASE_DAO.SqlDao.CMPSGSQuery4List(sql);
            List<DataRow> list = dt.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.MainList = plist;

            dt = BASE_DAO.SqlDao.CMPSGSQuery4List("select F.STAT_NO, F.PJ_NAME, S.PSG_NAME, S.SYR, S.BGN_SYR, S.END_SYR, D.DEPT_NAME, U.USR_NAME,\r\n(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='00') as STEP00c,\r\n(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='10') as STEP10c,\r\n(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='15') as STEP15c,\r\n(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='20') as STEP20c,\r\n(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='90') as STEPZZc,\r\n(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='Z1') as STEPZ1c,\r\n(select count(*) from APMFORM A where YRC_ID =S.YRC_ID and PPID =F.PPID and A.AP_STAT='Z2') as STEPZ2c\r\nfrom PJMFORM F left join PJDSYRC S\r\non F.PPID = S.PPID  \r\nleft join CMBASE.dbo.OGCDEPT D\r\non F.ADV_DP = D.DEPT_NO \r\nleft join CMBASE.dbo.OGCUSER U\r\non F.ADV_UID= U.USR_ID \r\nwhere S.OPEN_FG !=1\r\norder by F.PPID, S.SYR Desc, PSG_NAME");
            list = dt.AsEnumerable().ToList();
            plist = new PagedList<DataRow>(list, page, pageSize);
            model.List2 = plist;

            return View("SGE0000Index", model);
        }

        //列印
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Print")]
        public ActionResult Print(PJE0030Model model)
        {

            // 載入模板
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/PJE0030.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            Worksheet sheet = workbook.Worksheets[0];

            //做預存程序
            DataTable dt;
            using (Models.CMEPMSModels.CMEPMSEntities cmepms = new Models.CMEPMSModels.CMEPMSEntities())
            using (SqlConnection conn = new SqlConnection(cmepms.Database.Connection.ConnectionString))
            using (SqlCommand cmd = new SqlCommand("SP_RCD0010_UDF003_3", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                // 輸入參數
                cmd.Parameters.AddWithValue("@PID", model.PID);
                cmd.Parameters.AddWithValue("@SEQ_NO", model.SEQ_NO);

                dt = new DataTable();
                conn.Open();

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    dt.Load(reader);  // 將結果讀入 DataTable
                }


            }
            int startRow = 2;
            foreach (DataRow d in dt.Rows)
            {
                sheet.Cells[startRow, 0].PutValue(d["TD_CODE"]);
                sheet.Cells[startRow, 1].PutValue(d["SYR"]);
                sheet.Cells[startRow, 2].PutValue(d["W_CODE"]);
                sheet.Cells[startRow, 3].PutValue(d["PJ_CODE"]);
                sheet.Cells[startRow, 4].PutValue(d["IPJ_CODE"]);
                sheet.Cells[startRow, 5].PutValue(d["CITY_CODE"]);
                sheet.Cells[startRow, 6].PutValue(d["TL_CODE"]);
                sheet.Cells[startRow, 7].PutValue(d["TOD_CODE"]);
                sheet.Cells[startRow, 8].PutValue(d["TD_NAME"]);
                sheet.Cells[startRow, 9].PutValue(d["YR_FUND"]);
                sheet.Cells[startRow, 10].PutValue(d["CEN_FUND"]);
                sheet.Cells[startRow, 11].PutValue(d["LOC_FUND"]);
                sheet.Cells[startRow, 12].PutValue(d["FIT_DESC"]);
                sheet.Cells[startRow, 13].PutValue(d["TIT_DESC"]);
                sheet.Cells[startRow, 14].PutValue(d["FIT_DESC"]);
                sheet.Cells[startRow, 15].PutValue(d["TIT_DESC"]);
                //sheet.Cells[startRow, 12].Formula = $"=B{startRow + 1}*C{startRow + 1}";

                startRow++;
            }

            sheet.AutoFitColumns();

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"PJE0030_{timestamp}.xlsx";
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                ms.Position = 0;
                return File(ms.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName);
            }
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "print1")]
        public ActionResult print1(SGE0000Model m)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");

            string sql1 = "Select * from APV_APDRPAL where PJ_CATY_CODE IN('CA','CB','CC','CD','CE')AND AP_STAT=@AP_STAT AND PPID =@PPID and EXT_FG LIKE @EXT_FG and YRC_ID LIKE @YRC_ID and ADV_DP = @user_dp ORDER BY S1_ORDER, AP_NAME";
            string sql2 = "Select * from APV_APDRPAL where PJ_CATY_CODE IN('PA','PB')AND AP_STAT=@AP_STAT AND PPID =@PPID and EXT_FG LIKE @EXT_FG and YRC_ID LIKE @YRC_ID and ADV_DP = @user_dp ORDER BY S1_ORDER, AP_NAME";
            string sql3 = "Select * from APV_APDRPAL where PJ_CATY_CODE IN('OA','OB','OC')AND AP_STAT=@AP_STAT AND PPID =@PPID and EXT_FG LIKE @EXT_FG and YRC_ID LIKE @YRC_ID and ADV_DP = @user_dp ORDER BY S1_ORDER, AP_NAME";

            using (var ms = new MemoryStream())
            {
                //workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                byte[] fileBytes = GenerateReportBytes(sql1, sql2, sql3, m.PPID, m.YRC_ID, m.AP_STAT, m.seqno, out string sExcelTitle);//ms.ToArray();

                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"{sExcelTitle}提案報表_{timestamp}.xlsx");
            }

        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "print2")]
        public ActionResult print2(SGE0000Model m)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");

            string sql1 = "Select*from APV_APDRPAL where PJ_CATY_CODE in ('CA', 'CB', 'CC', 'CD', 'CE')AND AP_STAT=@AP_STAT AND PPID =@PPID and EXT_FG LIKE @EXT_FG and YRC_ID LIKE @YRC_ID and ADV_DP = @user_dp order by S1_ORDER, AP_NAME";
            string sql2 = "Select*from APV_APDRPAL where PJ_CATY_CODE in ('PA', 'PB')AND AP_STAT=@AP_STAT AND PPID =@PPID and EXT_FG LIKE @EXT_FG and YRC_ID LIKE @YRC_ID and ADV_DP = @user_dp order by S1_ORDER, AP_NAME";
            string sql3 = "Select*from APV_APDRPAL where PJ_CATY_CODE in ('OA', 'OB', 'OC')AND AP_STAT=@AP_STAT AND PPID =@PPID and EXT_FG LIKE @EXT_FG and YRC_ID LIKE @YRC_ID and ADV_DP = @user_dp order by S1_ORDER, AP_NAME";

            return File(
                    print2By(sql1, sql2, sql3, m.PPID, m.YRC_ID, m.AP_STAT, m.seqno),
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    $"提案初審評比表_{timestamp}.docx"
                );

        }

        /// <summary>
        /// 明細表
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="PPID"></param>
        /// <param name="YRC_ID"></param>
        /// <returns></returns>
        private byte[] GenerateReportBytes(string sql1, string sql2, string sql3, decimal PPID, string YRC_ID, string AP_STAT, string EXT_FG, out string title)
        {
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == PPID).FirstOrDefault();

            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0000/雨水計畫提案報表.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            SqlCommand cmd = new SqlCommand(sql1);
            cmd.Parameters.Add(new SqlParameter("AP_STAT", AP_STAT));
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));



            if (EXT_FG != null)
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", EXT_FG));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", "%"));
            }
            if (YRC_ID != null)
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", "%"));
            }


            cmd.Parameters.Add(new SqlParameter("user_dp", WebSiteHelper.getUserDept()));

            DataTable ts = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);

            Worksheet sheet = workbook.Worksheets[0];
            workbook.Settings.ReCalculateOnOpen = true;

            // 設定頁首頁尾
            Aspose.Cells.PageSetup ps = sheet.PageSetup;
            string deptNo = WebSiteHelper.getUserDept();
            var cityInfo = baseDb.OGCDEPT.SingleOrDefault(w => w.DEPT_NO == deptNo);
            char[] cityTargets = { '縣', '市' };
            string cityText = (cityInfo != null && cityInfo.SMP_DEPT_NAME.Any(c => cityTargets.Contains(c))) ? cityInfo.SMP_DEPT_NAME : "";
            string syrText = "";
            if (decimal.TryParse(YRC_ID, out var yrcIdDecimal))
            {
                var yrcInfo = sgsdb.PJDSYRC.SingleOrDefault(w => w.PPID == PPID && w.YRC_ID == yrcIdDecimal);
                syrText = (yrcInfo != null) ? $"{yrcInfo.BGN_SYR}年度" : "";
            }
            title = $"{cityText}{syrText}{pjmform.PJ_NAME}";
            string header = $"{title} 擬辦工程(含用地)評分表"; // ○○縣(市)○○年度因應氣候變遷縣市管河川及排水整體改善計畫（115-118年）擬辦工程(含用地)評分表
            ps.SetHeader(1, header);
            string footer = "";
            ps.SetFooter(0, footer);

            int startRow = 3;
            if (ts.Rows.Count > 0)
            {
                sheet.Cells[2, 10].PutValue(ts.Rows[0]["SYR"] + "年");
                sheet.Cells[2, 13].PutValue(ts.Rows[0]["SYR"] + "年");
            }
            foreach (DataRow d in ts.Rows)
            {
                sheet.Cells.CopyRow(sheet.Cells, startRow, startRow + 1);
                sheet.Cells[startRow, 0].PutValue(d["S1_ORDER"]);
                sheet.Cells[startRow, 1].PutValue(d["CODE_DESC"]);
                sheet.Cells[startRow, 2].PutValue(d["CITY"]);
                sheet.Cells[startRow, 3].PutValue(d["TOWN"]);
                sheet.Cells[startRow, 4].PutValue(d["WSYS_NAME"]);
                sheet.Cells[startRow, 5].PutValue(d["RSYS_NAME"]); // TODO：3.1.8雨水科實施計畫名稱修改後，需將RSYS_NAME改回AP_PJ_NAME // 20251222
                sheet.Cells[startRow, 6].PutValue(d["AP_NAME"]);
                sheet.Cells[startRow, 7].PutValue(d["AP_ALL_NOTE"]);
                //sheet.Cells[startRow, 8].PutValue(d[""]);
                sheet.Cells[startRow, 28].PutValue(d["REXE_FUND"]);
                sheet.Cells[startRow, 29].PutValue(d["CCEN_FUND"]);
                sheet.Cells[startRow, 30].PutValue(d["FCEN_FUND"]);
                //sheet.Cells[startRow, 12].PutValue(d[""]);
                sheet.Cells[startRow, 32].PutValue(d["CLOC_FUND"]);
                sheet.Cells[startRow, 31].PutValue(d["FLOC_FUND"]);
                //sheet.Cells[startRow, 15].PutValue(d[""]);
                sheet.Cells[startRow, 17].PutValue(d["SWW2_ST"]);
                sheet.Cells[startRow, 18].PutValue(d["SWW4_ST"]);
                sheet.Cells[startRow, 19].PutValue(d["IMPFLDAREA"]);
                sheet.Cells[startRow, 20].PutValue(d["PRTPOPS"]);
                sheet.Cells[startRow, 21].PutValue(d["START_DT"]);
                sheet.Cells[startRow, 22].PutValue(d["END_DT"]);
                sheet.Cells[startRow, 23].PutValue(d["DEPT_NAME"]);
                sheet.Cells[startRow, 24].PutValue(d["SC_SCORE1"] != DBNull.Value ? d["SC_SCORE1"] : "-");
                sheet.Cells[startRow, 25].PutValue(d["GIS_X"]);
                sheet.Cells[startRow, 26].PutValue(d["GIS_Y"]);
                sheet.Cells[startRow, 27].PutValue(d["AP_NOTE"]);
                sheet.Cells[startRow, 33].PutValue(Convert.ToBoolean(d["EXT_FG"]) == true ? "延續性工程" : "新建工程");
                startRow++;
            }
            sheet.Cells.DeleteRow(startRow);

            sheet.AutoFitColumns();
            sheet.Cells.HideColumns(28, 5);
            Helper.SetFooterRemark(sheet, 1, startRow, 34);
            Helper.SetPageSetup(ps, sheet);


            cmd = new SqlCommand(sql2);
            cmd.Parameters.Add(new SqlParameter("AP_STAT", AP_STAT));
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            if (EXT_FG != null)
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", EXT_FG));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", "%"));
            }
            if (YRC_ID != null)
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", "%"));
            }
            cmd.Parameters.Add(new SqlParameter("user_dp", WebSiteHelper.getUserDept()));
            ts = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            startRow = 3;
            sheet = workbook.Worksheets[1];

            // 設定頁首頁尾
            ps = sheet.PageSetup;
            header = $"{title} 擬辦規劃(檢討)評分表";
            ps.SetHeader(1, header);
            ps.SetFooter(0, footer);

            if (ts.Rows.Count > 0)
            {
                sheet.Cells[2, 8].PutValue(ts.Rows[0]["SYR"] + "年");
                sheet.Cells[2, 11].PutValue(ts.Rows[0]["SYR"] + "年");
            }
            foreach (DataRow d in ts.Rows)
            {
                sheet.Cells.CopyRow(sheet.Cells, startRow, startRow + 1);
                sheet.Cells[startRow, 0].PutValue(d["S1_ORDER"]);
                sheet.Cells[startRow, 1].PutValue(d["CODE_DESC"]);
                sheet.Cells[startRow, 2].PutValue(d["CITY"]);
                sheet.Cells[startRow, 3].PutValue(d["TOWN"]);
                sheet.Cells[startRow, 4].PutValue(d["WSYS_NAME"]);
                sheet.Cells[startRow, 5].PutValue(d["AP_NAME"]);
                //sheet.Cells[startRow, 6].PutValue(d[""]);
                sheet.Cells[startRow, 21].PutValue(d["REXE_FUND"]);
                sheet.Cells[startRow, 22].PutValue(d["CCEN_FUND"]);
                sheet.Cells[startRow, 23].PutValue(d["FCEN_FUND"]);
                //sheet.Cells[startRow, 10].PutValue(d[""]);
                sheet.Cells[startRow, 24].PutValue(d["CLOC_FUND"]);
                sheet.Cells[startRow, 25].PutValue(d["FLOC_FUND"]);
                //sheet.Cells[startRow, 13].PutValue(d[""]);
                //sheet.Cells[startRow, 14].PutValue(d[""]);
                sheet.Cells[startRow, 15].PutValue(d["START_DT"]);
                sheet.Cells[startRow, 16].PutValue(d["END_DT"]);
                sheet.Cells[startRow, 17].PutValue(d["IMPFLDAREA"]);
                sheet.Cells[startRow, 18].PutValue(d["DEPT_NAME"]);
                sheet.Cells[startRow, 19].PutValue(d["SC_SCORE1"] != DBNull.Value ? d["SC_SCORE1"] : "-");
                sheet.Cells[startRow, 20].PutValue(d["AP_NOTE"]);
                sheet.Cells[startRow, 26].PutValue(Convert.ToBoolean(d["EXT_FG"]) == true ? "延續性工程" : "新建工程");
                startRow++;
            }
            sheet.Cells.DeleteRow(startRow);

            sheet.AutoFitColumns();
            sheet.Cells.HideColumns(21, 5);
            Helper.SetFooterRemark(sheet, 2, startRow, 27);
            Helper.SetPageSetup(ps, sheet);

            cmd = new SqlCommand(sql3);
            cmd.Parameters.Add(new SqlParameter("AP_STAT", AP_STAT));
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            if (EXT_FG != null)
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", EXT_FG));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", "%"));
            }
            if (YRC_ID != null)
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", "%"));
            }
            cmd.Parameters.Add(new SqlParameter("user_dp", WebSiteHelper.getUserDept()));
            ts = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            startRow = 3;
            sheet = workbook.Worksheets[2];
            if (ts.Rows.Count > 0)
            {
                sheet.Cells[2, 9].PutValue(ts.Rows[0]["SYR"] + "年");
                sheet.Cells[2, 12].PutValue(ts.Rows[0]["SYR"] + "年");
            }

            // 設定頁首頁尾
            ps = sheet.PageSetup;
            header = $"{title} 擬辦非工程措施評分表";
            ps.SetHeader(1, header);
            ps.SetFooter(0, footer);

            foreach (DataRow d in ts.Rows)
            {
                sheet.Cells.CopyRow(sheet.Cells, startRow, startRow + 1);//複制格式到下行
                sheet.Cells[startRow, 0].PutValue(d["S1_ORDER"]);
                sheet.Cells[startRow, 1].PutValue(d["CITY"]);
                sheet.Cells[startRow, 2].PutValue(d["TOWN"]);
                sheet.Cells[startRow, 3].PutValue(d["WSYS_NAME"]);
                sheet.Cells[startRow, 4].PutValue(d["AP_NAME"]);
                sheet.Cells[startRow, 5].PutValue(d["AP_ALL_NOTE"]);
                sheet.Cells[startRow, 6].PutValue(d["BENEFITS"]);//BENEFITS
                //sheet.Cells[startRow, 7].PutValue(d[""]);
                sheet.Cells[startRow, 21].PutValue(d["REXE_FUND"]);
                sheet.Cells[startRow, 22].PutValue(d["CCEN_FUND"]);
                sheet.Cells[startRow, 23].PutValue(d["FCEN_FUND"]);
                //sheet.Cells[startRow, 110].PutValue(d[""]);
                sheet.Cells[startRow, 24].PutValue(d["CLOC_FUND"]);
                sheet.Cells[startRow, 25].PutValue(d["FLOC_FUND"]);
                //sheet.Cells[startRow, 14].PutValue(d[""]);
                //sheet.Cells[startRow, 15].PutValue(d[""]);
                sheet.Cells[startRow, 16].PutValue(d["START_DT"]);
                sheet.Cells[startRow, 17].PutValue(d["END_DT"]);
                sheet.Cells[startRow, 18].PutValue(d["DEPT_NAME"]);
                sheet.Cells[startRow, 19].PutValue(d["SC_SCORE1"] != DBNull.Value ? d["SC_SCORE1"] : "-");
                sheet.Cells[startRow, 20].PutValue(d["AP_NOTE"]);
                sheet.Cells[startRow, 26].PutValue(Convert.ToBoolean(d["EXT_FG"]) == true ? "延續性工程" : "新建工程");

                // 套用格式到目標行

                startRow++;
                //sn++;
                //sheet.Cells.InsertRows(startRow, 1, true);
            }

            sheet.Cells.DeleteRow(startRow);


            sheet.AutoFitColumns();
            sheet.Cells.HideColumns(21, 5);
            Helper.SetFooterRemark(sheet, 3, startRow, 27);
            Helper.SetPageSetup(ps, sheet);


            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                return ms.ToArray();
            }
        }

        private MemoryStream print2By(string sql1, string sql2, string sql3, decimal PPID, string YRC_ID, string AP_STAT, string EXT_FG)
        {

            SqlCommand cmd = new SqlCommand(sql1);
            cmd.Parameters.Add(new SqlParameter("AP_STAT", AP_STAT));
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            if (EXT_FG != null)
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", EXT_FG));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", "%"));
            }
            if (YRC_ID != null)
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", "%"));
            }
            cmd.Parameters.Add(new SqlParameter("user_dp", WebSiteHelper.getUserDept()));
            DataTable dt1 = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);

            cmd = new SqlCommand(sql2);
            cmd.Parameters.Add(new SqlParameter("AP_STAT", AP_STAT));
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            if (EXT_FG != null)
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", EXT_FG));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", "%"));
            }
            if (YRC_ID != null)
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", "%"));
            }
            cmd.Parameters.Add(new SqlParameter("user_dp", WebSiteHelper.getUserDept()));
            DataTable dt2 = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);

            cmd = new SqlCommand(sql3);
            cmd.Parameters.Add(new SqlParameter("AP_STAT", AP_STAT));
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            if (EXT_FG != null)
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", EXT_FG));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", "%"));
            }
            if (YRC_ID != null)
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", "%"));
            }
            cmd.Parameters.Add(new SqlParameter("user_dp", WebSiteHelper.getUserDept()));
            DataTable dt3 = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);


            // 建立一個空的主文件，用來合併後續每一筆資料的報表
            var masterDoc = new Document();
            masterDoc.RemoveAllChildren(); // 移除預設內容（空白 section）

            string docPath;
            byte[] templateBytes;
            decimal sum, s2, s3, s4, s5, s6;

            for (int i = 0; i < dt1.Rows.Count; i++)
            {

                // Word 範本檔案路徑（使用 IWebHostEnvironment 注入做相對路徑）
                docPath = Path.Combine(Server.MapPath("~/Templates/SGE0000/01.提案&初審評比表_工程V2.docx"));
                templateBytes = System.IO.File.ReadAllBytes(docPath);
                using (var tempStream = new MemoryStream(templateBytes))
                {
                    var tempDoc = new Document(tempStream);
                    // 找到 Word 文件中的所有文字方塊（Shape），進行內容取代

                    var builder = new DocumentBuilder(tempDoc);


                    string catyCode = Convert.ToString(dt1.Rows[i]["PJ_CATY_CODE"]);
                    string[] codes = new[] { "CA", "CB", "CC", "CD", "CE" };
                    foreach (var code in codes)
                    {
                        builder.MoveToBookmark($"PJ_CATY_CODE_{code}");
                        builder.Write(catyCode == code ? "■" : "□");
                    }
                    try { s2 = Convert.ToDecimal(dt1.Rows[i]["REXE_FUND"]) / 1000; } catch { s2 = 0; }
                    try { s3 = Convert.ToDecimal(dt1.Rows[i]["CCEN_FUND"]) / 1000; } catch { s3 = 0; }
                    try { s4 = Convert.ToDecimal(dt1.Rows[i]["FCEN_FUND"]) / 1000; } catch { s4 = 0; }
                    try { s5 = Convert.ToDecimal(dt1.Rows[i]["CLOC_FUND"]) / 1000; } catch { s5 = 0; }
                    try { s6 = Convert.ToDecimal(dt1.Rows[i]["FLOC_FUND"]) / 1000; } catch { s6 = 0; }
                    sum = s2 + s3 + s4 + s5 + s6;

                    builder.MoveToBookmark("sum23456");
                    builder.Write(sum.ToString("N0"));


                    // 設定取代選項
                    FindReplaceOptions options = new FindReplaceOptions
                    {
                        MatchCase = false, // 是否區分大小寫
                        FindWholeWordsOnly = true // 是否只匹配整個單字
                    };
                    string strtemp;
                    int inttmp;
                    int total = 0;

                    try { tempDoc.Range.Replace("S1_COMT", dt1.Rows[i]["S1_COMT"] + ""); } catch { tempDoc.Range.Replace("S1_COMT", ""); } // 初評意見
                    try { tempDoc.Range.Replace("IMPALLAREA", dt1.Rows[i]["IMPALLAREA"] + ""); } catch { tempDoc.Range.Replace("IMPALLAREA", ""); }
                    try { tempDoc.Range.Replace("PJ_CATY_CODE", dt1.Rows[i]["PJ_CATY_CODE"] + ""); } catch { }
                    try { tempDoc.Range.Replace("AID", dt1.Rows[i]["AID"] + ""); } catch { }
                    try { tempDoc.Range.Replace("YRC_ID", dt1.Rows[i]["YRC_ID"] + ""); } catch { }
                    try { tempDoc.Range.Replace("PPID", dt1.Rows[i]["PPID"] + ""); } catch { }
                    try { tempDoc.Range.Replace("SYR", dt1.Rows[i]["SYR"] + ""); } catch { }
                    try { tempDoc.Range.Replace("AP_CODE", dt1.Rows[i]["AP_CODE"] + ""); } catch { }
                    try { tempDoc.Range.Replace("ADV_DP", dt1.Rows[i]["ADV_DP"] + ""); } catch { }
                    try { tempDoc.Range.Replace("ADV_UID", dt1.Rows[i]["ADV_UID"] + ""); } catch { }
                    try { tempDoc.Range.Replace("S1_ORDER", dt1.Rows[i]["S1_ORDER"] + ""); } catch { }
                    try { tempDoc.Range.Replace("CODE_DESC", dt1.Rows[i]["CODE_DESC"] + ""); } catch { }
                    try { tempDoc.Range.Replace("CITY", dt1.Rows[i]["CITY"] + ""); } catch { }
                    try { tempDoc.Range.Replace("TOWN", dt1.Rows[i]["TOWN"] + ""); } catch { }
                    try { tempDoc.Range.Replace("TD_LOC", dt1.Rows[i]["TD_LOC"] + ""); } catch { }
                    try { tempDoc.Range.Replace("WSYS_NAME", dt1.Rows[i]["WSYS_NAME"] + ""); } catch { }
                    try { tempDoc.Range.Replace("AP_PJ_NAME", dt1.Rows[i]["AP_PJ_NAME"] + ""); } catch { }
                    try { tempDoc.Range.Replace("AP_STAT", dt1.Rows[i]["AP_STAT"] + ""); } catch { }
                    try { tempDoc.Range.Replace("AP_NAME", dt1.Rows[i]["AP_NAME"] + ""); } catch { }
                    try { tempDoc.Range.Replace("AP_ALL_NOTE", dt1.Rows[i]["AP_ALL_NOTE"] + ""); } catch { }
                    try { tempDoc.Range.Replace("REXE_FUND", s2.ToString("N0") + ""); } catch { }
                    try { tempDoc.Range.Replace("CCEN_FUND", s3.ToString("N0") + ""); } catch { }
                    try { tempDoc.Range.Replace("FCEN_FUND", s4.ToString("N0") + ""); } catch { }
                    try { tempDoc.Range.Replace("CLOC_FUND", s5.ToString("N0") + ""); } catch { }
                    try { tempDoc.Range.Replace("FLOC_FUND", s6.ToString("N0") + ""); } catch { }
                    try { tempDoc.Range.Replace("SWW2_ST", dt1.Rows[i]["SWW2_ST"] + ""); } catch { }
                    try { tempDoc.Range.Replace("SWW4_ST", dt1.Rows[i]["SWW4_ST"] + ""); } catch { }
                    try { tempDoc.Range.Replace("IMPFLDAREA", tools.safe2Decimal(dt1.Rows[i]["IMPFLDAREA"]).ToString("N2") + ""); } catch { }
                    try { tempDoc.Range.Replace("PRTPOPS", tools.safe2Decimal(dt1.Rows[i]["PRTPOPS"]).ToString("N0") + ""); } catch { }
                    try { tempDoc.Range.Replace("START_DT", dt1.Rows[i]["START_DT"] + ""); } catch { }
                    try { tempDoc.Range.Replace("END_DT", dt1.Rows[i]["END_DT"] + ""); } catch { }
                    try { tempDoc.Range.Replace("REXM_PRJ", dt1.Rows[i]["REXM_PRJ"] + ""); } catch { }
                    try { tempDoc.Range.Replace("DEPT_NAME", dt1.Rows[i]["DEPT_NAME"] + ""); } catch { }
                    //try { tempDoc.Range.Replace("TF106_DT", dt1.Rows[i]["TF106_DT"] + ""); } catch { }
                    //try { tempDoc.Range.Replace("TF108_DT", dt1.Rows[i]["TF108_DT"] + ""); } catch { }
                    try
                    {
                        string tf106_filed = "TF106_DT";
                        string tf106_dt = "";
                        if (dt1.Columns.Contains(tf106_filed) && dt1.Rows[i][tf106_filed] != DBNull.Value)
                            tf106_dt = DateTimeExtensions.ToSimpleTaiwanDate(Convert.ToDateTime(dt1.Rows[i][tf106_filed]));
                        tempDoc.Range.Replace("TF106_DT", tf106_dt + "");
                    }
                    catch { }
                    try
                    {
                        string tf108_filed = "TF108_DT";
                        string tf108_dt = "";
                        if (dt1.Columns.Contains(tf108_filed) && dt1.Rows[i][tf108_filed] != DBNull.Value)
                            tf108_dt = DateTimeExtensions.ToSimpleTaiwanDate(Convert.ToDateTime(dt1.Rows[i][tf108_filed]));
                        tempDoc.Range.Replace("TF108_DT", tf108_dt + "");
                    }
                    catch { }
                    try
                    {
                        strtemp = dt1.Rows[i]["SWW2_FG"].ToString().ToUpper();
                        if (strtemp.Equals("TRUE"))
                        {
                            tempDoc.Range.Replace("&SWW2_FG$", "是");
                        }
                        else
                        {
                            tempDoc.Range.Replace("&SWW2_FG$", "");
                        }
                    }
                    catch
                    {
                        tempDoc.Range.Replace("&SWW2_FG$", "");
                    }
                    try
                    {
                        strtemp = dt1.Rows[i]["SWW4_FG"].ToString().ToUpper();
                        if (strtemp.Equals("TRUE"))
                        {
                            tempDoc.Range.Replace("SWW4_FG", "是");
                        }
                        else
                        {
                            tempDoc.Range.Replace("SWW4_FG", "");
                        }
                    }
                    catch
                    {
                        tempDoc.Range.Replace("SWW4_FG", "");
                    }


                    //年
                    try { tempDoc.Range.Replace("year$", dt1.Rows[i]["SYR"] + ""); } catch { }
                    //分數
                    tempDoc.Range.Replace("SC_SCORE1", (dt1.Rows[i]["SC_SCORE1"] != DBNull.Value ? dt1.Rows[i]["SC_SCORE1"].ToString() : "-"));
                    //try
                    //{
                    //    inttmp = Convert.ToInt32(dt1.Rows[i]["SC_SCORE1"]);
                    //    if (inttmp == 0)
                    //    {
                    //        tempDoc.Range.Replace("SC_SCORE1", "");
                    //    }
                    //    else
                    //    {
                    //        tempDoc.Range.Replace("SC_SCORE1", inttmp + "");
                    //        total += inttmp;
                    //    }
                    //}
                    //catch
                    //{
                    //    tempDoc.Range.Replace("SC_SCORE1", "");
                    //}
                    for (int sc = 1; sc <= 8; sc++)
                    {
                        string columnName = $"SC{sc}_SCORE1";
                        string placeholder = $"{columnName}"; // Word 標記
                        string value = dt1.Rows[i][columnName] != DBNull.Value ? dt1.Rows[i][columnName].ToString() : "-"; // 判斷 null 或 DBNull.Value
                        tempDoc.Range.Replace(placeholder, value); // 替換 Word 範本
                    }
                    //for (int sc = 1; sc <= 8; sc++)
                    //{
                    //    string fieldName = $"SC{sc}_SCORE1";
                    //    try
                    //    {
                    //        inttmp = Convert.ToInt32(dt1.Rows[i][fieldName]);
                    //        if (inttmp == 0)
                    //        {
                    //            tempDoc.Range.Replace(fieldName, "");
                    //        }
                    //        else
                    //        {
                    //            tempDoc.Range.Replace(fieldName, inttmp.ToString());
                    //            total += inttmp;
                    //        }
                    //    }
                    //    catch
                    //    {
                    //        tempDoc.Range.Replace(fieldName, "");
                    //    }
                    //}



                    try { tempDoc.Range.Replace("SC1_CMT1", dt1.Rows[i]["SC1_CMT1"] + ""); } catch { tempDoc.Range.Replace("SC1_CMT1", ""); }
                    try { tempDoc.Range.Replace("SC2_CMT1", dt1.Rows[i]["SC2_CMT1"] + ""); } catch { tempDoc.Range.Replace("SC2_CMT1", ""); }
                    try { tempDoc.Range.Replace("SC3_CMT1", dt1.Rows[i]["SC3_CMT1"] + ""); } catch { tempDoc.Range.Replace("SC3_CMT1", ""); }
                    try { tempDoc.Range.Replace("SC4_CMT1", dt1.Rows[i]["SC4_CMT1"] + ""); } catch { tempDoc.Range.Replace("SC4_CMT1", ""); }
                    try { tempDoc.Range.Replace("SC5_CMT1", dt1.Rows[i]["SC5_CMT1"] + ""); } catch { tempDoc.Range.Replace("SC5_CMT1", ""); }
                    try { tempDoc.Range.Replace("SC6_CMT1", dt1.Rows[i]["SC6_CMT1"] + ""); } catch { tempDoc.Range.Replace("SC6_CMT1", ""); }
                    try { tempDoc.Range.Replace("SC7_CMT1", dt1.Rows[i]["SC7_CMT1"] + ""); } catch { tempDoc.Range.Replace("SC7_CMT1", ""); }
                    try { tempDoc.Range.Replace("SC8_CMT1", dt1.Rows[i]["SC8_CMT1"] + ""); } catch { tempDoc.Range.Replace("SC8_CMT1", ""); }

                    try { tempDoc.Range.Replace("SCS_CMT1", dt1.Rows[i]["SCS_CMT1"] + ""); } catch { tempDoc.Range.Replace("SCS_CMT1", ""); }
                    try { tempDoc.Range.Replace("AP_NOTE", dt1.Rows[i]["AP_NOTE"] + ""); } catch { tempDoc.Range.Replace("AP_NOTE", ""); }
                    try { tempDoc.Range.Replace("EXT_FG", Convert.ToBoolean(dt1.Rows[i]["EXT_FG"]) == true ? "延續性工程" : "新建工程" + ""); } catch { tempDoc.Range.Replace("EXT_FG", ""); }
                    try { tempDoc.Range.Replace("CONT_NAME", dt1.Rows[i]["CONT_NAME"] + ""); } catch { tempDoc.Range.Replace("CONT_NAME", ""); }
                    try { tempDoc.Range.Replace("CONT_TITLE", dt1.Rows[i]["CONT_TITLE"] + ""); } catch { tempDoc.Range.Replace("CONT_TITLE", ""); }
                    try { tempDoc.Range.Replace("CONT_TEL", dt1.Rows[i]["CONT_TEL"] + ""); } catch { tempDoc.Range.Replace("CONT_TEL", ""); }
                    try { tempDoc.Range.Replace("CONT_EMAIL", dt1.Rows[i]["CONT_EMAIL"] + ""); } catch { tempDoc.Range.Replace("CONT_EMAIL", ""); }
                    try { tempDoc.Range.Replace("BENEFITS", dt1.Rows[i]["BENEFITS"] + ""); } catch { tempDoc.Range.Replace("BENEFITS", ""); }
                    //短字串在後
                    for (int s = 1; s <= 8; s++)
                    {
                        string columnName = $"SC{s}_SCORE";
                        string placeholder = $"{columnName}$"; // Word 標記
                        string value = dt1.Rows[i][columnName] != DBNull.Value ? dt1.Rows[i][columnName].ToString() : "-"; // 判斷 null 或 DBNull.Value
                        tempDoc.Range.Replace(placeholder, value); // 替換 Word 範本
                    }
                    //try { tempDoc.Range.Replace("SC1_SCORE$", dt1.Rows[i]["SC1_SCORE"] + ""); } catch { tempDoc.Range.Replace("SC1_SCORE$", ""); }
                    //try { tempDoc.Range.Replace("SC2_SCORE$", dt1.Rows[i]["SC2_SCORE"] + ""); } catch { tempDoc.Range.Replace("SC2_SCORE$", ""); }
                    //try { tempDoc.Range.Replace("SC3_SCORE$", dt1.Rows[i]["SC3_SCORE"] + ""); } catch { tempDoc.Range.Replace("SC3_SCORE$", ""); }
                    //try { tempDoc.Range.Replace("SC4_SCORE$", dt1.Rows[i]["SC4_SCORE"] + ""); } catch { tempDoc.Range.Replace("SC4_SCORE$", ""); }
                    //try { tempDoc.Range.Replace("SC5_SCORE$", dt1.Rows[i]["SC5_SCORE"] + ""); } catch { tempDoc.Range.Replace("SC5_SCORE$", ""); }
                    //try { tempDoc.Range.Replace("SC6_SCORE$", dt1.Rows[i]["SC6_SCORE"] + ""); } catch { tempDoc.Range.Replace("SC6_SCORE$", ""); }
                    //try { tempDoc.Range.Replace("SC7_SCORE$", dt1.Rows[i]["SC7_SCORE"] + ""); } catch { tempDoc.Range.Replace("SC7_SCORE$", ""); }
                    //try { tempDoc.Range.Replace("SC8_SCORE$", dt1.Rows[i]["SC8_SCORE"] + ""); } catch { tempDoc.Range.Replace("SC8_SCORE$", ""); }


                    for (int sc = 1; sc <= 8; sc++)
                    {
                        string fieldName = $"SC{sc}_CMT";
                        string placeholder = fieldName + "$";
                        try
                        {
                            strtemp = Convert.ToString(dt1.Rows[i][fieldName]);
                            if (strtemp == "0")
                            {
                                strtemp = "";
                            }
                            tempDoc.Range.Replace(placeholder, strtemp);
                        }
                        catch
                        {
                            tempDoc.Range.Replace(placeholder, "");
                        }
                    }







                    tempDoc.Range.Replace("SC_SCORE$", (dt1.Rows[i]["SC_SCORE"] != DBNull.Value ? dt1.Rows[i]["SC_SCORE"].ToString() : "-"));
                    //try { tempDoc.Range.Replace("SC_SCORE$", dt1.Rows[i]["SC_SCORE"] + ""); } catch { }
                    try { tempDoc.Range.Replace("SCS_CMT$", dt1.Rows[i]["SCS_CMT"] + ""); } catch { }
                    // 若主文件已有 section，則設定新文件起始為新頁面
                    if (masterDoc.Sections.Count > 0)
                    {
                        tempDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
                    }
                    // 將該份套表結果文件附加到主文件
                    masterDoc.AppendDocument(tempDoc, ImportFormatMode.KeepSourceFormatting);
                }

            }

            string tmp;
            for (int i = 0; i < dt2.Rows.Count; i++)
            {

                // 2
                docPath = Path.Combine(Server.MapPath("~/Templates/SGE0000/02.提案&初審評比表_規劃(檢討)V2.docx"));
                templateBytes = System.IO.File.ReadAllBytes(docPath);
                using (var tempStream2 = new MemoryStream(templateBytes))
                {
                    var tempDoc2 = new Document(tempStream2);

                    var builder = new DocumentBuilder(tempDoc2);
                    string catyCode = Convert.ToString(dt2.Rows[i]["PJ_CATY_CODE"]);
                    string[] codes = new[] { "PA", "PB" };
                    foreach (var code in codes)
                    {
                        builder.MoveToBookmark($"PJ_CATY_CODE_{code}");
                        builder.Write(catyCode == code ? "■" : "□");
                    }
                    try { s2 = Convert.ToDecimal(dt2.Rows[i]["REXE_FUND"]) / 1000; } catch { s2 = 0; }
                    try { s3 = Convert.ToDecimal(dt2.Rows[i]["CCEN_FUND"]) / 1000; } catch { s3 = 0; }
                    try { s4 = Convert.ToDecimal(dt2.Rows[i]["FCEN_FUND"]) / 1000; } catch { s4 = 0; }
                    try { s5 = Convert.ToDecimal(dt2.Rows[i]["CLOC_FUND"]) / 1000; } catch { s5 = 0; }
                    try { s6 = Convert.ToDecimal(dt2.Rows[i]["FLOC_FUND"]) / 1000; } catch { s6 = 0; }
                    sum = s2 + s3 + s4 + s5 + s6;

                    builder.MoveToBookmark("sum23456");
                    builder.Write(sum.ToString("N0"));

                    //取代                    
                    try
                    {
                        tmp = "";
                        if (dt2.Rows[i]["NPSD_SYR"] != null)
                        {
                            tmp = dt2.Rows[i]["NPSD_SYR"].ToString();
                        }
                        tempDoc2.Range.Replace("NPSD_SYR$", tmp);


                    }
                    catch
                    {
                        tempDoc2.Range.Replace("NPSD_SYR$", "    ");
                    }
                    try
                    {
                        string OPSD_SYR_filed = "OPSD_SYR";
                        tmp = "";
                        if (dt2.Rows[i][OPSD_SYR_filed] != null)
                        {
                            tmp = dt2.Rows[i][OPSD_SYR_filed].ToString();
                        }
                        tempDoc2.Range.Replace("OPSD_SYR$", tmp);
                    }
                    catch
                    {
                        tempDoc2.Range.Replace("OPSD_SYR$", "    ");
                    }
                    try
                    {
                        string PPSD_SYR_filed = "PPSD_SYR";
                        tmp = "";
                        if (dt2.Rows[i][PPSD_SYR_filed] != null)
                        {
                            tmp = dt2.Rows[i][PPSD_SYR_filed].ToString();
                        }
                        tempDoc2.Range.Replace("PPSD_SYR$", tmp);
                    }
                    catch
                    {
                        tempDoc2.Range.Replace("PPSD_SYR$", "    ");
                    }
                    try { tempDoc2.Range.Replace("IMPALLAREA", tools.safe2Decimal(dt2.Rows[i]["IMPALLAREA"]).ToString("N2") + ""); }
                    catch
                    {
                        tempDoc2.Range.Replace("IMPALLAREA", "");
                    }
                    try { tempDoc2.Range.Replace("IMPFLDAREA", tools.safe2Decimal(dt2.Rows[i]["IMPFLDAREA"]).ToString("N2") + ""); }
                    catch
                    {
                        tempDoc2.Range.Replace("IMPFLDAREA", "");
                    }
                    try { tempDoc2.Range.Replace("REXE_FUND", s2.ToString("N0") + ""); }
                    catch
                    {
                        tempDoc2.Range.Replace("REXE_FUND", "0");
                    }
                    try { tempDoc2.Range.Replace("CCEN_FUND", s3.ToString("N0") + ""); }
                    catch
                    {
                        tempDoc2.Range.Replace("CCEN_FUND", "0");
                    }
                    try { tempDoc2.Range.Replace("FCEN_FUND", s4.ToString("N0") + ""); }
                    catch
                    {
                        tempDoc2.Range.Replace("FCEN_FUND", "0");
                    }
                    try { tempDoc2.Range.Replace("CLOC_FUND", s5.ToString("N0") + ""); }
                    catch
                    {
                        tempDoc2.Range.Replace("CLOC_FUND", "0");
                    }
                    try { tempDoc2.Range.Replace("FLOC_FUND", s6.ToString("N0") + ""); }
                    catch
                    {
                        tempDoc2.Range.Replace("FLOC_FUND", "0");
                    }

                    // 初評分數總分
                    tempDoc2.Range.Replace("SC_SCORE$", (dt2.Rows[i]["SC_SCORE"] != DBNull.Value ? dt2.Rows[i]["SC_SCORE"].ToString() : "-"));
                    // 自評分數總分
                    tempDoc2.Range.Replace("SC_SCORE1", (dt2.Rows[i]["SC_SCORE1"] != DBNull.Value ? dt2.Rows[i]["SC_SCORE1"].ToString() : "-"));
                    // 各別初評分數
                    for (int s = 1; s <= 8; s++)
                    {
                        string columnName = $"SC{s}_SCORE";
                        string placeholder = $"{columnName}$"; // Word 標記
                        string value = dt2.Rows[i][columnName] != DBNull.Value ? dt2.Rows[i][columnName].ToString() : "-"; // 判斷 null 或 DBNull.Value
                        tempDoc2.Range.Replace(placeholder, value); // 替換 Word 範本
                    }
                    // 各別自評分數
                    for (int sc = 1; sc <= 8; sc++)
                    {
                        string columnName = $"SC{sc}_SCORE1";
                        string placeholder = $"{columnName}"; // Word 標記
                        string value = dt2.Rows[i][columnName] != DBNull.Value ? dt2.Rows[i][columnName].ToString() : "-"; // 判斷 null 或 DBNull.Value
                        tempDoc2.Range.Replace(placeholder, value); // 替換 Word 範本
                    }

                    /*

                                        { "", "REXE_FUND" },
                        { "", "CCEN_FUND" },
                        { "", "" },
                        { "", "CLOC_FUND" },
                        { "", "FLOC_FUND" },

                    */
                    // Word 欄位名稱 → DataTable 欄位名稱
                    Dictionary<string, string> fieldMap = new Dictionary<string, string>
                    {
    { "year$", "SYR" },
    { "PJ_CATY_CODE", "PJ_CATY_CODE" },
    { "AID", "AID" },
    { "YRC_ID", "YRC_ID" },
    { "PPID", "PPID" },
    { "SYR", "SYR" },
    { "AP_CODE", "AP_CODE" },
    { "ADV_DP", "ADV_DP" },
    { "ADV_UID", "ADV_UID" },
    { "S1_ORDER", "S1_ORDER" },
    { "CODE_DESC", "CODE_DESC" },
    { "CITY", "CITY" },
    { "TOWN", "TOWN" },
    { "TD_LOC", "TD_LOC" },
    { "WSYS_NAME", "WSYS_NAME" },
    { "AP_PJ_NAME", "AP_PJ_NAME" },
    { "AP_STAT", "AP_STAT" },
    { "AP_NAME", "AP_NAME" },
    { "AP_ALL_NOTE", "AP_ALL_NOTE" },

    { "SWW2_ST", "SWW2_ST" },
    { "SWW4_ST", "SWW4_ST" },
    { "IMPFLDAREA", "IMPFLDAREA" },
    { "PRTPOPS", "PRTPOPS" },
    { "START_DT", "START_DT" },
    { "END_DT", "END_DT" },
    { "REXM_PRJ", "REXM_PRJ" },
    { "DEPT_NAME", "DEPT_NAME" },
    { "TF106_DT", "TF106_DT" },
    { "TF108_DT", "TF108_DT" },
    { "&SWW2_FG$", "SWW2_FG" },
    { "SWW4_FG", "SWW4_FG" },
    //{ "SC1_SCORE1", "SC1_SCORE1" },
    //{ "SC2_SCORE1", "SC2_SCORE1" },
    //{ "SC3_SCORE1", "SC3_SCORE1" },
    //{ "SC4_SCORE1", "SC4_SCORE1" },
    //{ "SC5_SCORE1", "SC5_SCORE1" },
    //{ "SC6_SCORE1", "SC6_SCORE1" },
    //{ "SC7_SCORE1", "SC7_SCORE1" },
    //{ "SC8_SCORE1", "SC8_SCORE1" },
    { "SC1_CMT1", "SC1_CMT1" },
    { "SC2_CMT1", "SC2_CMT1" },
    { "SC3_CMT1", "SC3_CMT1" },
    { "SC4_CMT1", "SC4_CMT1" },
    { "SC5_CMT1", "SC5_CMT1" },
    { "SC6_CMT1", "SC6_CMT1" },
    { "SC7_CMT1", "SC7_CMT1" },
    { "SC8_CMT1", "SC8_CMT1" },
    //{ "SC_SCORE1", "SC_SCORE1" },
    { "SCS_CMT1", "SCS_CMT1" },
    { "AP_NOTE", "AP_NOTE" },
    { "CONT_NAME", "CONT_NAME" },
    { "CONT_TITLE", "CONT_TITLE" },
    { "CONT_TEL", "CONT_TEL" },
    { "CONT_EMAIL", "CONT_EMAIL" },
    //{ "SC1_SCORE$", "SC1_SCORE" },
    //{ "SC2_SCORE$", "SC2_SCORE" },
    //{ "SC3_SCORE$", "SC3_SCORE" },
    //{ "SC4_SCORE$", "SC4_SCORE" },
    //{ "SC5_SCORE$", "SC5_SCORE" },
    //{ "SC6_SCORE$", "SC6_SCORE" },
    //{ "SC7_SCORE$", "SC7_SCORE" },
    //{ "SC8_SCORE$", "SC8_SCORE" },
    { "SC1_CMT$", "SC1_CMT" },
    { "SC2_CMT$", "SC2_CMT" },
    { "SC3_CMT$", "SC3_CMT" },
    { "SC4_CMT$", "SC4_CMT" },
    { "SC5_CMT$", "SC5_CMT" },
    { "SC6_CMT$", "SC6_CMT" },
    { "SC7_CMT$", "SC7_CMT" },
    { "SC8_CMT$", "SC8_CMT" },
    //{ "SC_SCORE$", "SC_SCORE" },
    { "SCS_CMT$", "SCS_CMT" },

    { "IMPALLAREA", "IMPALLAREA" },
    {"S1_COMT", "S1_COMT" } // 初評意見
};

                    foreach (var pair in fieldMap)
                    {
                        string wordField = pair.Key;
                        string dataField = pair.Value;

                        try
                        {
                            tempDoc2.Range.Replace(wordField, dt2.Rows[i][dataField]?.ToString() ?? "");
                        }
                        catch
                        {
                            tempDoc2.Range.Replace(wordField, "");
                        }
                    }


                    try { tempDoc2.Range.Replace("EXT_FG", Convert.ToBoolean(dt2.Rows[i]["EXT_FG"]) == true ? "延續性工程" : "新建工程" + ""); }
                    catch
                    {
                        tempDoc2.Range.Replace("EXT_FG", "延續性工程");
                    }

                    // 若主文件已有 section，則設定新文件起始為新頁面
                    if (masterDoc.Sections.Count > 0)
                    {
                        tempDoc2.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
                    }
                    // 將該份套表結果文件附加到主文件
                    masterDoc.AppendDocument(tempDoc2, ImportFormatMode.KeepSourceFormatting);
                }

            }

            for (int i = 0; i < dt3.Rows.Count; i++)
            {

                // 3
                docPath = Path.Combine(Server.MapPath("~/Templates/SGE0000/03.提案&初審評比表_非工程V2.docx"));
                templateBytes = System.IO.File.ReadAllBytes(docPath);
                using (var tempStream3 = new MemoryStream(templateBytes))
                {
                    var tempDoc3 = new Document(tempStream3);

                    var builder = new DocumentBuilder(tempDoc3);
                    string catyCode = Convert.ToString(dt3.Rows[i]["PJ_CATY_CODE"]);
                    string[] codes = new[] { "OA", "OB", "OC" };
                    foreach (var code in codes)
                    {
                        builder.MoveToBookmark($"PJ_CATY_CODE_{code}");
                        builder.Write(catyCode == code ? "■" : "□");
                    }
                    try { s2 = Convert.ToDecimal(dt3.Rows[i]["REXE_FUND"]) / 1000; } catch { s2 = 0; }
                    try { s3 = Convert.ToDecimal(dt3.Rows[i]["CCEN_FUND"]) / 1000; } catch { s3 = 0; }
                    try { s4 = Convert.ToDecimal(dt3.Rows[i]["FCEN_FUND"]) / 1000; } catch { s4 = 0; }
                    try { s5 = Convert.ToDecimal(dt3.Rows[i]["CLOC_FUND"]) / 1000; } catch { s5 = 0; }
                    try { s6 = Convert.ToDecimal(dt3.Rows[i]["FLOC_FUND"]) / 1000; } catch { s6 = 0; }
                    sum = s2 + s3 + s4 + s5 + s6;

                    builder.MoveToBookmark("sum23456");
                    builder.Write(sum.ToString("N0"));
                    // 找到 Word 文件中的所有文字方塊（Shape），進行內容取代
                    try { tempDoc3.Range.Replace("S1_COMT", dt3.Rows[i]["S1_COMT"] + ""); } catch { tempDoc3.Range.Replace("S1_COMT", ""); } // 初評意見
                    try { tempDoc3.Range.Replace("IMPALLAREA", dt3.Rows[i]["IMPALLAREA"] + ""); } catch { tempDoc3.Range.Replace("IMPALLAREA", ""); }
                    try { tempDoc3.Range.Replace("year$", dt3.Rows[i]["SYR"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("PJ_CATY_CODE", dt3.Rows[i]["PJ_CATY_CODE"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("AID", dt3.Rows[i]["AID"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("YRC_ID", dt3.Rows[i]["YRC_ID"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("PPID", dt3.Rows[i]["PPID"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SYR", dt3.Rows[i]["SYR"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("AP_CODE", dt3.Rows[i]["AP_CODE"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("ADV_DP", dt3.Rows[i]["ADV_DP"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("ADV_UID", dt3.Rows[i]["ADV_UID"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("S1_ORDER", dt3.Rows[i]["S1_ORDER"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("CODE_DESC", dt3.Rows[i]["CODE_DESC"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("CITY", dt3.Rows[i]["CITY"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("TOWN", dt3.Rows[i]["TOWN"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("TD_LOC", dt3.Rows[i]["TD_LOC"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("WSYS_NAME", dt3.Rows[i]["WSYS_NAME"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("AP_PJ_NAME", dt3.Rows[i]["AP_PJ_NAME"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("AP_STAT", dt3.Rows[i]["AP_STAT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("AP_NAME", dt3.Rows[i]["AP_NAME"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("AP_ALL_NOTE", dt3.Rows[i]["AP_ALL_NOTE"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("BENEFITS", dt3.Rows[i]["BENEFITS"] + ""); } catch { tempDoc3.Range.Replace("BENEFITS", ""); }
                    try { tempDoc3.Range.Replace("REXE_FUND", s2.ToString("N0") + ""); } catch { }
                    try { tempDoc3.Range.Replace("CCEN_FUND", s3.ToString("N0") + ""); } catch { }
                    try { tempDoc3.Range.Replace("FCEN_FUND", s4.ToString("N0") + ""); } catch { }
                    try { tempDoc3.Range.Replace("CLOC_FUND", s5.ToString("N0") + ""); } catch { }
                    try { tempDoc3.Range.Replace("FLOC_FUND", s6.ToString("N0") + ""); } catch { }
                    try { tempDoc3.Range.Replace("SWW2_ST", dt3.Rows[i]["SWW2_ST"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SWW4_ST", dt3.Rows[i]["SWW4_ST"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("IMPFLDAREA", tools.safe2Decimal(dt3.Rows[i]["IMPFLDAREA"]).ToString("N2") + ""); } catch { }
                    try { tempDoc3.Range.Replace("PRTPOPS", tools.safe2Decimal(dt3.Rows[i]["PRTPOPS"]).ToString("N0") + ""); } catch { }
                    try { tempDoc3.Range.Replace("START_DT", dt3.Rows[i]["START_DT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("END_DT", dt3.Rows[i]["END_DT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("REXM_PRJ", dt3.Rows[i]["REXM_PRJ"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("DEPT_NAME", dt3.Rows[i]["DEPT_NAME"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("TF106_DT", dt3.Rows[i]["TF106_DT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("TF108_DT", dt3.Rows[i]["TF108_DT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("&SWW2_FG$", dt3.Rows[i]["SWW2_FG"] + ""); } catch { tempDoc3.Range.Replace("&SWW2_FG$", ""); }
                    try { tempDoc3.Range.Replace("SWW4_FG", dt3.Rows[i]["SWW4_FG"] + ""); } catch { }

                    // 各別自評分數
                    for (int sc = 1; sc <= 8; sc++)
                    {
                        string columnName = $"SC{sc}_SCORE1";
                        string placeholder = $"{columnName}"; // Word 標記
                        string value = dt3.Rows[i][columnName] != DBNull.Value ? dt3.Rows[i][columnName].ToString() : "-"; // 判斷 null 或 DBNull.Value
                        tempDoc3.Range.Replace(placeholder, value); // 替換 Word 範本
                    }
                    //try { tempDoc3.Range.Replace("SC1_SCORE1", dt3.Rows[i]["SC1_SCORE1"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC2_SCORE1", dt3.Rows[i]["SC2_SCORE1"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC3_SCORE1", dt3.Rows[i]["SC3_SCORE1"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC4_SCORE1", dt3.Rows[i]["SC4_SCORE1"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC5_SCORE1", dt3.Rows[i]["SC5_SCORE1"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC6_SCORE1", dt3.Rows[i]["SC6_SCORE1"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC7_SCORE1", dt3.Rows[i]["SC7_SCORE1"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC8_SCORE1", dt3.Rows[i]["SC8_SCORE1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC1_CMT1", dt3.Rows[i]["SC1_CMT1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC2_CMT1", dt3.Rows[i]["SC2_CMT1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC3_CMT1", dt3.Rows[i]["SC3_CMT1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC4_CMT1", dt3.Rows[i]["SC4_CMT1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC5_CMT1", dt3.Rows[i]["SC5_CMT1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC6_CMT1", dt3.Rows[i]["SC6_CMT1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC7_CMT1", dt3.Rows[i]["SC7_CMT1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC8_CMT1", dt3.Rows[i]["SC8_CMT1"] + ""); } catch { }
                    tempDoc3.Range.Replace("SC_SCORE1", (dt3.Rows[i]["SC_SCORE1"] != DBNull.Value ? dt3.Rows[i]["SC_SCORE1"].ToString() : "-"));
                    //try { tempDoc3.Range.Replace("SC_SCORE1", dt3.Rows[i]["SC_SCORE1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SCS_CMT1", dt3.Rows[i]["SCS_CMT1"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("AP_NOTE", dt3.Rows[i]["AP_NOTE"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("EXT_FG", Convert.ToBoolean(dt3.Rows[i]["EXT_FG"]) == true ? "延續性工程" : "新建工程" + ""); } catch { }
                    try { tempDoc3.Range.Replace("CONT_NAME", dt3.Rows[i]["CONT_NAME"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("CONT_TITLE", dt3.Rows[i]["CONT_TITLE"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("CONT_TEL", dt3.Rows[i]["CONT_TEL"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("CONT_EMAIL", dt3.Rows[i]["CONT_EMAIL"] + ""); } catch { }
                    //短字串在後
                    tempDoc3.Range.Replace("SC_SCORE$", (dt3.Rows[i]["SC_SCORE"] != DBNull.Value ? dt3.Rows[i]["SC_SCORE"].ToString() : "-"));
                    //try { tempDoc3.Range.Replace("SC_SCORE$", dt3.Rows[i]["SC_SCORE"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SCS_CMT$", dt3.Rows[i]["SCS_CMT"] + ""); } catch { }
                    // 各別初評分數
                    for (int s = 1; s <= 8; s++)
                    {
                        string columnName = $"SC{s}_SCORE";
                        string placeholder = $"{columnName}$"; // Word 標記
                        string value = dt3.Rows[i][columnName] != DBNull.Value ? dt3.Rows[i][columnName].ToString() : "-"; // 判斷 null 或 DBNull.Value
                        tempDoc3.Range.Replace(placeholder, value); // 替換 Word 範本
                    }
                    //try { tempDoc3.Range.Replace("SC1_SCORE$", dt3.Rows[i]["SC1_SCORE"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC2_SCORE$", dt3.Rows[i]["SC2_SCORE"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC3_SCORE$", dt3.Rows[i]["SC3_SCORE"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC4_SCORE$", dt3.Rows[i]["SC4_SCORE"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC5_SCORE$", dt3.Rows[i]["SC5_SCORE"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC6_SCORE$", dt3.Rows[i]["SC6_SCORE"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC7_SCORE$", dt3.Rows[i]["SC7_SCORE"] + ""); } catch { }
                    //try { tempDoc3.Range.Replace("SC8_SCORE$", dt3.Rows[i]["SC8_SCORE"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC1_CMT$", dt3.Rows[i]["SC1_CMT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC2_CMT$", dt3.Rows[i]["SC2_CMT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC3_CMT$", dt3.Rows[i]["SC3_CMT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC4_CMT$", dt3.Rows[i]["SC4_CMT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC5_CMT$", dt3.Rows[i]["SC5_CMT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC6_CMT$", dt3.Rows[i]["SC6_CMT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC7_CMT$", dt3.Rows[i]["SC7_CMT"] + ""); } catch { }
                    try { tempDoc3.Range.Replace("SC8_CMT$", dt3.Rows[i]["SC8_CMT"] + ""); } catch { }

                    // 若主文件已有 section，則設定新文件起始為新頁面
                    if (masterDoc.Sections.Count > 0)
                    {
                        tempDoc3.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
                    }
                    // 將該份套表結果文件附加到主文件
                    masterDoc.AppendDocument(tempDoc3, ImportFormatMode.KeepSourceFormatting);
                }

            }

            // 將合併後的文件儲存為 MemoryStream
            var outputStream = new MemoryStream();
            masterDoc.Save(outputStream, WordsSaveFormat.Docx);
            outputStream.Position = 0;
            return outputStream;
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "PrintSGE0010")]
        public ActionResult PrintSGE0010(SGE0000Model model)
        {
            // 載入模板
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0010.xlsx"));
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            string EXT_FGWhere = (string.IsNullOrEmpty(model.seqno)) ? "" : " AND F.EXT_FG = @EXT_FG "; // 辦理狀態

            SqlCommand cmd = new SqlCommand("SELECT F.* FROM APV_APMFORM F LEFT OUTER JOIN PJDSYRC S on F.PPID = S.PPID AND S.YRC_ID=F.YRC_ID " +
                "left join CMBASE.dbo.OGCDEPT D on F.ADV_DP = D.DEPT_NO left join CMBASE.dbo.OGCUSER U on F.ADV_UID= U.USR_ID WHERE " +
                "F.AP_STAT=@STAT AND F.PPID =@PPID " + EXT_FGWhere +
                " AND S.OPEN_FG=1 AND F.YRC_ID='" + model.YRC_ID + "'AND F.ADV_UID='" + WebSiteHelper.getUserId() + "' order by F.PPID, S.SYR Desc, PSG_NAME");
            cmd.Parameters.Add(new SqlParameter("PPID", model.PPID));
            cmd.Parameters.Add(new SqlParameter("STAT", model.AP_STAT));
            if (model.YRC_ID != null)
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", model.YRC_ID));
            }
            if (!string.IsNullOrEmpty(model.seqno)) // 辦理狀態
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", model.seqno));
            }
            DataTable ts = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            if (ts.Rows.Count == 0)
            {
                ViewBag.msg = "查無資料!";
                return ProposalDetailList(model);
                //using (var ms = new MemoryStream())
                //{
                //    workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                //    ms.Position = 0;
                //    return File(ms.ToArray(),
                //                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                //                "無資料.xlsx");
                //}
            }

            int startRow = 5;
            int sn = 1;
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Replace("$SYR&", ts.Rows[0]["SYR"].ToString());
            sheet.Replace("$PJ_NAME&", model.PJ_NAME);
            foreach (DataRow d in ts.Rows)
            {
                //
                sheet.Cells[startRow, 0].PutValue(sn);
                sheet.Cells[startRow, 1].PutValue(d["AP_NAME"]);
                //2是個洞
                sheet.Cells[startRow, 3].PutValue(d["TD_LINE"]);
                //sheet.Cells[startRow, 4].PutValue(d["FUND"]);
                //sheet.Cells[startRow, 5].PutValue(d["R_FUND"]);
                //sheet.Cells[startRow, 6].PutValue(d["DCS_FUND"]);
                //sheet.Cells[startRow, 7].PutValue(d["CEN_DCS_FUND"]);
                //sheet.Cells[startRow, 8].PutValue(d["LOC_DCS_FUND"]);
                sheet.Cells[startRow, 4].PutValue(d["M_T_FUND"]);
                //if (tools.safe2Decimal(d["M_T_FUND"]) == 0m)
                //{
                //var calcFund = tools.safe2Decimal(d["M_R_FUND"]) + tools.safe2Decimal(d["M_FUND"]);
                //sheet.Cells[startRow, 4].PutValue(calcFund);
                //}
                sheet.Cells[startRow, 5].PutValue(d["M_R_FUND"]);
                sheet.Cells[startRow, 6].PutValue(d["M_FUND"]);
                sheet.Cells[startRow, 7].PutValue(d["M_CEN_FUND"]);
                sheet.Cells[startRow, 8].PutValue(d["M_LOC_FUND"]);

                //try
                //{
                //    if (Convert.ToBoolean(d["SWW1_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 9].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 10].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 9].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 10].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 9].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 10].PutValue(""); } catch { }
                //}

                //try
                //{
                //    if (Convert.ToBoolean(d["SWW2_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 11].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 12].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 11].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 12].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 11].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 12].PutValue(""); } catch { }
                //}

                //try
                //{
                //    if (Convert.ToBoolean(d["SWW3_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 13].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 14].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 13].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 14].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 13].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 14].PutValue(""); } catch { }
                //}

                //try
                //{
                //    if (Convert.ToBoolean(d["SWW4_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 15].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 16].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 15].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 16].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 15].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 16].PutValue(""); } catch { }
                //}
                //try
                //{
                //    if (Convert.ToBoolean(d["SWW5_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 17].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 18].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 17].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 18].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 17].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 18].PutValue(""); } catch { }
                //}
                //try
                //{
                //    if (Convert.ToBoolean(d["SWW6_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 19].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 20].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 19].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 20].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 19].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 20].PutValue(""); } catch { }
                //}

                //try
                //{
                //    if (Convert.ToBoolean(d["SWW7_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 21].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 22].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 21].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 22].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 21].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 22].PutValue(""); } catch { }
                //}

                //try
                //{
                //    if (Convert.ToBoolean(d["SWW8_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 23].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 24].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 23].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 24].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 23].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 24].PutValue(""); } catch { }
                //}

                //try
                //{
                //    if (Convert.ToBoolean(d["SWW9_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 25].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 26].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 25].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 26].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 25].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 26].PutValue(""); } catch { }
                //}

                //try
                //{
                //    if (Convert.ToBoolean(d["SWW10_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 27].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 28].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 27].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 28].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 27].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 28].PutValue(""); } catch { }
                //}

                //try
                //{
                //    if (Convert.ToBoolean(d["SWW11_FG"]))
                //    {
                //        try { sheet.Cells[startRow, 29].PutValue("V"); } catch { }
                //        try { sheet.Cells[startRow, 30].PutValue(""); } catch { }
                //    }
                //    else
                //    {
                //        try { sheet.Cells[startRow, 29].PutValue(""); } catch { }
                //        try { sheet.Cells[startRow, 30].PutValue("V"); } catch { }
                //    }
                //}
                //catch
                //{
                //    try { sheet.Cells[startRow, 29].PutValue(""); } catch { }
                //    try { sheet.Cells[startRow, 30].PutValue(""); } catch { }
                //}
                //try { sheet.Cells[startRow, 30].PutValue(""); } catch { }
                WriteSWWFGColumns(sheet, d, startRow);
                sheet.Cells[startRow, 31].PutValue(d["SWWOT_DESC"]);
                sheet.Cells[startRow, 32].PutValue(d["S1_ORDER"]);
                sheet.Cells[startRow, 33].PutValue(d["AP_NOTE"]);
                startRow++;
                sn++;
                sheet.Cells.InsertRows(startRow, 1, true);
            }

            sheet.AutoFitColumns();

            // 列印設定
            Aspose.Cells.PageSetup ps = sheet.PageSetup;
            Helper.SetPageSetup(ps, sheet);

            // 範例：114年度申請污水下水道相關計畫先期作業表
            string syr = ts.Rows[0]["SYR"].ToString();
            string pjName = model.PJ_NAME;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
            string fileName = $"{syr}年度申請{pjName}先期作業表_{timestamp}.xlsx";
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);
                ms.Position = 0;
                return File(ms.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName);
            }
        }

        // 安全取得 bool? 值
        bool? GetBitValue(object value)
        {
            if (value is bool b)
                return b;
            if (value == null || value == DBNull.Value)
                return null;
            return null; // 不可轉換時，也回傳 null（可改拋例外）
        }

        // 安全寫入 Excel
        void SafePutValue(Worksheet sheet, int row, int col, object value)
        {
            try { sheet.Cells[row, col].PutValue(value); }
            catch { /* 可記錄或略過 */ }
        }

        /// <summary>
        /// 寫入勘察評估事項相關欄位
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="d"></param>
        /// <param name="row"></param>
        void WriteSWWFGColumns(Worksheet sheet, DataRow d, int row)
        {
            int col = 9;

            // SWW1_FG~SWW7_FG：每個佔兩欄
            for (int i = 1; i <= 7; i++)
            {
                string field = string.Format("SWW{0}_FG", i);
                bool? val = GetBitValue(d[field]);

                SafePutValue(sheet, row, col, val == true ? "V" : "");
                SafePutValue(sheet, row, col + 1, val == false ? "V" : "");
                col += 2; // 共佔兩欄
            }

            // SWW8_FG：true 放 col 23，false 放 col 25
            bool? s8 = GetBitValue(d["SWW8_FG"]);
            SafePutValue(sheet, row, 23, s8 == true ? "V" : "");
            SafePutValue(sheet, row, 25, s8 == false ? "V" : "");

            // SWW81_FG：單欄 bit 顯示在 col 24
            bool? s8_1 = GetBitValue(d["SWW81_FG"]);
            SafePutValue(sheet, row, 24, s8_1 == true ? "V" : "");

            // SWW9_RATE：數值，寫入 col 26
            object s9 = d["SWW9_RATE"];
            SafePutValue(sheet, row, 26, s9 == DBNull.Value ? "" : s9);

            // SWW10_FG：雙欄顯示 true/false → col 27, 28
            bool? s10 = GetBitValue(d["SWW10_FG"]);
            SafePutValue(sheet, row, 27, s10 == true ? "V" : "");
            SafePutValue(sheet, row, 28, s10 == false ? "V" : "");

            // SWW11_FG：雙欄顯示 true/false → col 29, 30
            bool? s11 = GetBitValue(d["SWW11_FG"]);
            SafePutValue(sheet, row, 29, s11 == true ? "V" : "");
            SafePutValue(sheet, row, 30, s11 == false ? "V" : "");
        }

        void SetFlagValue(Dictionary<string, object> d, Worksheet sheet, int startRow, string key, int trueCol, int falseCol)
        {
            try
            {
                if (Convert.ToBoolean(d[key]))
                {
                    try { sheet.Cells[startRow, trueCol].PutValue("V"); } catch { }
                    try { sheet.Cells[startRow, falseCol].PutValue(""); } catch { }
                }
                else
                {
                    try { sheet.Cells[startRow, trueCol].PutValue(""); } catch { }
                    try { sheet.Cells[startRow, falseCol].PutValue("V"); } catch { }
                }
            }
            catch
            {
                try { sheet.Cells[startRow, trueCol].PutValue(""); } catch { }
                try { sheet.Cells[startRow, falseCol].PutValue(""); } catch { }
            }
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Search")]
        public ActionResult Search(SGE0000Model model)
        {

            int page = 1;// (model.page == 0) ? 1 : model.page;

            StringBuilder sql = new StringBuilder("SELECT F.* FROM APMFORM F LEFT OUTER JOIN PJDSYRC S on F.PPID = S.PPID AND S.YRC_ID=F.YRC_ID " +
               "left join CMBASE.dbo.OGCDEPT D on F.ADV_DP = D.DEPT_NO left join CMBASE.dbo.OGCUSER U on F.ADV_UID= U.USR_ID WHERE " +
               "F.AP_STAT=@STAT AND F.PPID =@PPID AND S.OPEN_FG=1 AND F.EXT_FG LIKE @EXT_FG AND F.YRC_ID LIKE @YRC_ID AND F.ADV_UID='" + WebSiteHelper.getUserId() + "' order by F.PPID, S.SYR Desc, PSG_NAME");




            SqlCommand cmd = new SqlCommand(sql.ToString());
            if (model.seqno != null)
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", model.seqno));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("EXT_FG", "%"));
            }
            if (model.YRC_ID != null)
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", model.YRC_ID));
            }
            else
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", "%"));
            }


            cmd.Parameters.Add(new SqlParameter("PPID", model.PPID));
            cmd.Parameters.Add(new SqlParameter("STAT", model.AP_STAT));
            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);

            model.List1 = plist;
            DataTable yrcs = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC p WHERE p.PPID='" + model.PPID + "' ORDER BY p.YRC_ID DESC");
            ViewBag.yrcs = tools.ToSelectList(yrcs, "YRC_ID", "PSG_NAME", "");
            ViewBag.stat = tools.GetStat(model.AP_STAT);
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == model.PPID).FirstOrDefault();
            model.PJMFORM = pjmform;
            return View("SGE0000ProposalDetailList", model);
        }


        // 送件初審  -1 -> 00 ST
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendForward")]
        public ActionResult SendForward(SGE0000Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            decimal aid = tools.safe2Decimal(model.AID);
            APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            // AP_STAT = -1(提案時) 才能送件
            if (a.AP_STAT != "-1")
            {
                ViewBag.msg = "狀態已變更，請勿重複操作";
                return ProposalDetailList(model);
            }
            string o_status = a.AP_STAT;
            a.AP_STAT = code.AP_STAT_Forward(a.AP_STAT);
            a.UPD_DT = now;
            a.UPD_UID = userid;
            sgsdb.SaveChanges();

            util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "ST", now, userid, WebSiteHelper.getUserDept(), "");
            ViewBag.msg = "已送件";
            //return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2);
            return ProposalDetailList(model);
        }

        // 批次送件初審  -1 -> 00 ST
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendForwardBatch")]
        public ActionResult SendForwardBatch(SGE0000Model model)
        {
            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            string deptid = WebSiteHelper.getUserDept(); // 移到迴圈外，避免重複取值影響效能

            // 防呆：確認前端有傳值過來
            if (string.IsNullOrEmpty(model.AID))
            {
                ViewBag.msg = "未選擇任何要送件的資料";
                return ProposalDetailList(model);
            }

            // 將前端傳來逗號分隔的字串切成陣列 (例如 "1001,1002" -> ["1001", "1002"])
            string[] aidArray = model.AID.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            int successCount = 0; // 用來計算成功送件的筆數

            foreach (string aidStr in aidArray)
            {
                decimal aid = tools.safe2Decimal(aidStr);
                APMFORM a = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();

                // 確保該筆資料存在，且 AP_STAT == "-1" (提案時) 才能送件
                if (a != null && a.AP_STAT == "-1")
                {
                    string o_status = a.AP_STAT;
                    a.AP_STAT = code.AP_STAT_Forward(a.AP_STAT);
                    a.UPD_DT = now;
                    a.UPD_UID = userid;

                    // 寫入 Log
                    util.UPDATE_APMFORM_Log(a.AID, o_status, a.AP_STAT, "ST", now, userid, deptid, "");

                    successCount++;
                }
            }

            // 檢查是否有成功處理的資料，統一執行資料庫儲存 (提升效能)
            if (successCount > 0)
            {
                sgsdb.SaveChanges();
                ViewBag.msg = $"已成功批次送件 {successCount} 筆";
            }
            else
            {
                ViewBag.msg = "勾選的資料狀態皆已變更，無符合送件條件的項目";
            }

            return ProposalDetailList(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SendDelete")]
        public ActionResult SendDelete(SGE0000Model model)
        {


            BASE_DAO.SqlDao.CMPSGSExecuteSql("DELETE FROM APMFORM WHERE AID='" + model.AID + "'");

            ViewBag.msg = "已刪除";
            //return ProposalManage(model.tab, model.PPID.ToString(), model.page1, model.page2);
            return ProposalDetailList(model);
        }

        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ProposalManage")]
        public ActionResult ProposalManage(SGE0000Model model)
        {

            int page = (model.page == 0) ? 1 : model.page;
            int tab = 1;
            if (Request.Params["tab"] != null)
            {
                tab = Convert.ToInt32(Request.Params["tab"]);
            }
            ViewBag.yrcs = tools.ToSelectList(SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC WHERE PPID=" + model.PPID + "' ORDER BY p.YRC_ID DESC"), "YRC_ID", "PSG_NAME", "請選擇");
            //Request.Params["tab"] = "1";


            DataTable ps = util.GetAPMFORM(model.PPID, util.GetAP_STAT(tab), model.YRC);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == model.PPID).FirstOrDefault();
            model.PJMFORM = pjmform;
            model.List1 = plist;
            return View("PJE0011ProposalManage", model);
        }

        [HttpGet]
        public ActionResult ProposalManage(SGE0000Model model, int tab, string page, int ppid, string yrcid)
        {

            if (page == null) page = "1";
            string sql = "select * from APMFORM where AP_STAT='" + util.GetAP_STAT(tab) + "' and PPID ='" + model.PPID + "'";
            if (!string.IsNullOrEmpty(yrcid))
            {
                sql += "and YRC_ID='" + yrcid + "'";
            }
            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(sql);



            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, Convert.ToInt32(page), pageSize);
            model.List1 = plist;
            model.page = Convert.ToInt32(page);
            model.PPID = ppid;
            model.YRC_ID = yrcid;

            DataTable yrcs = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC p WHERE p.PPID='" + model.PPID + "' ORDER BY p.YRC_ID");
            ViewBag.yrcs = tools.ToSelectList(yrcs, "SYR", "YRC_ID", null);
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == model.PPID).FirstOrDefault();
            model.PJMFORM = pjmform;
            return View("PJE0011ProposalManage", model);
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ProposalDetailList")]
        public ActionResult ProposalDetailList(SGE0000Model model)
        {
            //decimal aid = Convert.ToDecimal(model.AID);
            //APMFORM apmform = sgsdb.APMFORM.Where(w => w.AID == aid).FirstOrDefault();
            PJMFORM_CMPSGS pjmform = sgsdb.PJMFORM.Where(w => w.PPID == model.PPID).FirstOrDefault();
            //if(ppid == null) ppid = Convert.ToString(apmform.PPID);
            //if(yrcid == null) yrcid = Convert.ToString(apmform.YRC_ID);
            model.PJMFORM = pjmform;

            int page = (model.page == 0) ? 1 : model.page;
            //SqlCommand cmd = new SqlCommand("select * from APMFORM where AP_STAT='00' and PPID =@PPID and YRC_ID=@YRC_ID");
            SqlCommand cmd = new SqlCommand("SELECT P.STAT_NO, F.* FROM PJMFORM P, APMFORM F LEFT OUTER JOIN PJDSYRC S on F.PPID = S.PPID AND S.YRC_ID=F.YRC_ID " +
                "left join CMBASE.dbo.OGCDEPT D on F.ADV_DP = D.DEPT_NO left join CMBASE.dbo.OGCUSER U on F.ADV_UID= U.USR_ID WHERE P.PPID=f.PPID AND " +
                "F.AP_STAT=@STAT AND F.PPID =@PPID AND S.OPEN_FG=1 AND F.YRC_ID='" + model.YRC_ID + "'AND F.ADV_UID='" + WebSiteHelper.getUserId() + "' order by F.PPID, S.SYR Desc, PSG_NAME");
            cmd.Parameters.Add(new SqlParameter("PPID", model.PPID));
            cmd.Parameters.Add(new SqlParameter("STAT", model.AP_STAT));
            if (model.YRC_ID != null)
            {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", model.YRC_ID));
            }
            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);


            List<DataRow> list = ps.AsEnumerable().ToList();
            DataTable APDLOCP;
            StringBuilder towns;
            for (int i = 0; i < list.Count; i++)
            {
                towns = new StringBuilder();
                APDLOCP = SqlDao.CMPSGSQuery4List("SELECT TOWN FROM APDLOCP WHERE AID=" + list[i]["AID"]);
                foreach (DataRow a in APDLOCP.Rows)
                {
                    towns.Append(a["TOWN"] + ", ");
                }
                //towns.Length -= 2;

                list[i]["TOWN"] = towns;
            }
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, pageSize);
            model.List1 = plist;


            DataTable yrcs = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC p WHERE p.PPID='" + model.PPID + "' ORDER BY p.YRC_ID DESC");
            ViewBag.yrcs = tools.ToSelectList(yrcs, "YRC_ID", "PSG_NAME", "");
            //model.YRC_ID = Convert.ToString(yrcs.Rows[0]["YRC_ID"]);
            ViewBag.stat = tools.GetStat(model.AP_STAT);


            return View("SGE0000ProposalDetailList", model);
        }

        [HttpGet]
        public ActionResult ProposalDetailList(int page, string YRC_ID, decimal PPID, string AP_STAT)
        {
            /*
            SqlCommand cmd = new SqlCommand("SELECT F.* FROM APMFORM F LEFT OUTER JOIN PJDSYRC S on F.PPID = S.PPID AND S.YRC_ID=F.YRC_ID " +
                "left join CMBASE.dbo.OGCDEPT D on F.ADV_DP = D.DEPT_NO left join CMBASE.dbo.OGCUSER U on F.ADV_UID= U.USR_ID WHERE " +
                "F.AP_STAT=@STAT AND F.PPID =@PPID AND S.OPEN_FG=1 AND F.YRC_ID='" + YRC_ID + "'AND F.ADV_UID='" + WebSiteHelper.getUserId() + "' order by F.PPID, S.SYR Desc, PSG_NAME");
            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            cmd.Parameters.Add(new SqlParameter("STAT", AP_STAT));
            if(YRC_ID != null) {
                cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            }
            DataTable ps = BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, page, 10);

            SGE0000Model model=new SGE0000Model();

            model.List1 = plist;
            model.page = page;
            model.AP_STAT = AP_STAT;
            model.YRC_ID = YRC_ID;
            DataTable yrcs = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDSYRC p WHERE p.PPID='" + model.PPID + "' ORDER BY p.YRC_ID DESC");
            ViewBag.yrcs = tools.ToSelectList(yrcs, "YRC_ID", "PSG_NAME", "");
            ViewBag.stat = tools.GetStat(model.AP_STAT);


            return View("SGE0000ProposalDetailList", model);
            */
            SGE0000Model model = new SGE0000Model();
            model.AP_STAT = AP_STAT;
            model.YRC_ID = YRC_ID;
            model.PPID = PPID;
            model.page = page;
            //model.PSG_NAME
            return ProposalDetailList(model);
        }


        private void initAPMFORM(string ppid, string yrcid, string city)
        {
            //鄉鎮
            ViewBag.towns = tools.ToSelectList(util.GetTownsById(city), "TOWN", "TOWN", null);
            //申請案件類別
            ViewBag.PJ_CATY_CODE = tools.ToSelectList(util.GetCMPSGSCODES("TT", "03"), "id", "name", null);
            //實施計畫名稱PJCFORM
            //ViewBag.PJ_NAMES = tools.ToSelectList(SqlDao.CMETMSQuery4List("SELECT PJ_NAME as name FROM PJCFORM WHERE PJ_ID_TYPE='C'"), "name", "name", null);
            ViewBag.PJ_NAMES = tools.ToSelectList(SqlDao.CMETMSQuery4List("select PJ_NAME as name, SER_CITY from CMETMS.dbo.PJCFORM where PJ_ID_TYPE ='C' and SYS_ID in (select SYS_ID from CMETMS.dbo.PJCSYST where STAT_NO='S')"), "name", "name", null);
            //雨水下水道系統及幹線名稱(R用)
            ViewBag.PJ_NAMES_R = tools.ToSelectList(GetPJNamesR(city), "name", "name", null);
            //縣市
            ViewBag.cities = tools.ToSelectList(util.GetCitiesNew(), "CITY", "CITY", null);
            //受補助單位
            //List<OGCDEPT> f = WCals.Helpers.WebSiteHelper.getParentDept(WCals.Helpers.WebSiteHelper.getUserId());
            ViewBag.FILL_MCHS = tools.ToSelectList(WCals.Helpers.WebSiteHelper.getParentDept(WCals.Helpers.WebSiteHelper.getUserId()), "id", "name", null);

            //受補助單位ALL
            //List<OGCDEPT> f = WCals.Helpers.WebSiteHelper.getParentDept(WCals.Helpers.WebSiteHelper.getUserId());
            ViewBag.FILL_MCHS_ALL = tools.ToSelectList(SqlDao.CMBASEQuery4List("SELECT DEPT_NO as id, DEPT_NAME as name FROM OGCDEPT"), "id", "name", null);

            //辦理層級
            ViewBag.IPJ_CODES = tools.ToSelectList(util.GetCMETMSCODES("SM", "02"), "id", "name", null);
            //項目類型
            ViewBag.IT_TYPE = tools.ToSelectList(util.GetCMEPMSCODES("IT", "10"), "id", "name", null);
            //項目類別
            ViewBag.PJDCMCT = SqlDao.CMPSGSQuery4List("SELECT*FROM PJDCMCT WHERE PPID=" + ppid);
            //系統別
            ViewBag.SYS_CATY = tools.ToSelectList(SqlDao.CMPSGSQuery4List("SELECT CATY_CODE as id, CATY_NAME as name FROM PJDCATY WHERE PPID='" + ppid + "'"), "id", "name", null);
            //ViewBag.SYS_CATY = tools.ToSelectList(SqlDao.CMEPMSQuery4List("SELECT CATY_CODE as id, CATY_NAME as name FROM PJDCATY WHERE PPID='" + ppid + "'"), "id", "name", null);
            //座標系統
            ViewBag.GIS_SYS = util.GetCMPSGSCODES("SY", "02");


        }

        /// <summary>
        /// 取得雨水下水道系統及幹線名稱清單(R用)
        /// </summary>
        /// <returns></returns>
        private DataTable GetPJNamesR(string strCITY)
        {

            string sql = "select SYS_NAME as name, CITY from CMETMS.dbo.PJCSYST where STAT_NO='R' ";

            // 1. 取得登入者縣市
            string loginCity = Source.Tools.GetLoginUserCity(User.Identity.Name);

            // 2. 處理縣市下拉選單
            if (!string.IsNullOrEmpty(loginCity))
            {
                sql+="and CITY = '" + loginCity + "'";
            }

            return SqlDao.CMETMSQuery4List(sql);
        }

        /*
        //建立計畫
        public ActionResult AddAPMFORM(string ppid, string yrcid) {
            SGE0000Model model = new SGE0000Model();
            model.PPID = tools.safe2Decimal(ppid);
            model.YRC_ID = yrcid;
            initAPMFORM(ppid, yrcid, null);

            return View("SGE0000AddAPMFORM", model);
        }


        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SaveNewAPMFORM")]
        public ActionResult SaveNewAPMFORM(SGE0000Model model) {
            
            if(string.IsNullOrEmpty(model.APMFORM.CATY_CODE)||
                string.IsNullOrEmpty(model.APMFORM.AP_NAME)||
                string.IsNullOrEmpty(model.APMFORM.CITY) ||
                string.IsNullOrEmpty(model.APMFORM.AP_PJ_NAME)||
                string.IsNullOrEmpty(model.APMFORM.FILL_MCH) ||
                string.IsNullOrEmpty(model.APMFORM.EXE_MCH) ||
                model.START_DT == null ||
                model.END_DT==null
                ) {

                ViewBag.msg = "*為必填項目";

                return EditAPMFORM(model.APMFORM.AID);
            }

            DateTime now = DateTime.Now;
            string userid = WebSiteHelper.getUserId();
            string userdp = WebSiteHelper.getUserDept();            

            APMFORM apmform;
            if(model.APMFORM.AID==0) {
                apmform = new APMFORM();
            } else {
                apmform = sgsdb.APMFORM.Where(x => x.AID == model.APMFORM.AID).FirstOrDefault();
            }

            apmform.AP_STAT = "-1";
            apmform.PPID =                      model.PPID;
            apmform.YRC_ID = tools.safe2Decimal(model.YRC_ID);
            apmform.CATY_CODE =                 model.APMFORM.CATY_CODE;
            apmform.AP_NAME =                   model.APMFORM.AP_NAME;
            apmform.CITY =                      model.APMFORM.CITY;
            apmform.AP_PJ_NAME =                model.APMFORM.AP_PJ_NAME;
            apmform.FILL_MCH =                  model.APMFORM.FILL_MCH;
            apmform.EXE_MCH =                   model.APMFORM.EXE_MCH;

            model.START_DT = "0" + model.START_DT;
            DateTime csharpDate = DateTime.ParseExact(model.START_DT, "yyyyMMdd", null);
            csharpDate = csharpDate.AddYears(1911);
            apmform.START_DT = csharpDate;

            model.END_DT = "0" + model.END_DT;
            csharpDate = DateTime.ParseExact(model.END_DT, "yyyyMMdd", null);
            csharpDate = csharpDate.AddYears(1911);
            apmform.END_DT = csharpDate;            
            
            apmform.BENEFITS =                  model.APMFORM.BENEFITS;
            apmform.AP_YEAR_NOTE =              model.APMFORM.AP_YEAR_NOTE;
            apmform.AP_NOTE =       model.APMFORM.AP_NOTE;
            apmform.CONT_NAME = model.APMFORM.CONT_NAME;
            apmform.CONT_EMAIL = model.APMFORM.CONT_EMAIL;
            apmform.CONT_FAX = model.APMFORM.CONT_FAX;
            apmform.TF106_DT = model.APMFORM.TF106_DT;
            apmform.TF108_DT = model.APMFORM.TF108_DT;
            apmform.TK103_DT = model.APMFORM.TK103_DT;
            apmform.TK105_DT = model.APMFORM.TK105_DT;
            apmform.TK106_DT = model.APMFORM.TK106_DT;
            apmform.TK107_DT = model.APMFORM.TK107_DT;
            apmform.SYS_CATY = model.APMFORM.SYS_CATY;
            apmform.S1_ORDER = model.APMFORM.S1_ORDER;
            apmform.ATTENTION_PEOPLE = model.APMFORM.ATTENTION_PEOPLE;            
            
            //評估事項
            if(!string.IsNullOrEmpty(model.SWW1_FG) && model.SWW1_FG.Equals("Y")) {apmform.SWW1_FG = true;}
            if(!string.IsNullOrEmpty(model.SWW2_FG)&&model.SWW2_FG.Equals("Y")) { apmform.SWW2_FG = true; }
            if(!string.IsNullOrEmpty(model.SWW3_FG)&&model.SWW3_FG.Equals("Y")) { apmform.SWW3_FG = true; }
            if(!string.IsNullOrEmpty(model.SWW4_FG)&&model.SWW4_FG.Equals("Y")) { apmform.SWW4_FG = true; }
            if(!string.IsNullOrEmpty(model.SWW5_FG)&&model.SWW5_FG.Equals("Y")) { apmform.SWW5_FG = true; }
            if(!string.IsNullOrEmpty(model.SWW6_FG)&&model.SWW6_FG.Equals("Y")) { apmform.SWW6_FG = true; }
            if(!string.IsNullOrEmpty(model.SWW7_FG)&&model.SWW7_FG.Equals("Y")) { apmform.SWW7_FG = true; }
            if(!string.IsNullOrEmpty(model.SWW8_FG) && model.SWW8_FG.Equals("Y")) { apmform.SWW8_FG = true; }
            if(!string.IsNullOrEmpty(model.SWW81_FG) && model.SWW81_FG.Equals("Y")) { apmform.SWW81_FG = true; }
            apmform.SWW9_RATE = tools.safe2Decimal(model.SWW9_FG);
            if(!string.IsNullOrEmpty(model.SWW10_FG) && model.SWW10_FG.Equals("Y")) { apmform.SWW10_FG = true; }
            if(!string.IsNullOrEmpty(model.SWW11_FG) && model.SWW11_FG.Equals("Y")) { apmform.SWW11_FG = true; }
            apmform.SWWOT_DESC = model.SWWOT_DESC;
            
            
            //財源分配
            apmform.TCEN_FUND = model.APMFORM.TCEN_FUND;
            apmform.TLOC_FUND = model.APMFORM.TLOC_FUND;
            apmform.RCEN_FUND = model.APMFORM.RCEN_FUND;
            apmform.RLOC_FUND = model.APMFORM.RLOC_FUND;              
            
            apmform.UPD_DT = now;
            apmform.ADV_UID = userid;
            apmform.ADV_DP = userdp;
            //新增
            if(model.APMFORM.AID == 0) {
                apmform.ADD_DT = now;
                sgsdb.APMFORM.Add(apmform);
                sgsdb.SaveChanges();
            } else {
                sgsdb.SaveChanges();
            }            

            decimal fid;
            if(model.SYRS != null) {
                APDFUND apdfund;
                for(int i = 0; i < model.SYRS.Length; i++) {
                    if(string.IsNullOrEmpty(model.SYRS[i]) ||
                        string.IsNullOrEmpty(model.FUND[i])
                        ) {
                        continue;
                    }

                    if(string.IsNullOrEmpty(model.APDFUND_FID[i])) {
                        apdfund = new APDFUND();
                    } else {
                        fid = tools.safe2Decimal(model.APDFUND_FID[i]);
                        apdfund = sgsdb.APDFUND.Where(w=>w.FID==fid).FirstOrDefault();
                    }
                    


                    apdfund.AID = apmform.AID;
                    apdfund.SYR = model.SYRS[i];
                    apdfund.FUND = tools.safe2Decimal(model.FUND[i]);
                    apdfund.FIT_CODE = model.FIT_CODE[i];
                    apdfund.FUND_DESC = model.FUND_DESC[i];
                    apdfund.UPD_UID = userid;
                    apdfund.UPD_DT = now;
                    apdfund.ADD_DT = now;

                    if(string.IsNullOrEmpty(model.APDFUND_FID[i])) {
                        sgsdb.APDFUND.Add(apdfund);
                    }
                    
                }
                sgsdb.SaveChanges();
            }
                        
            if(model.files != null) {
                APDFILE apdfile;
                byte[] data;                
                for(int i = 0; i < model.FID.Length; i++) {

                    if(string.IsNullOrEmpty(model.FID[i])) {
                        apdfile = new APDFILE();
                    } else {
                        continue;
                    }

                    
                    
                    using(Stream inputStream = model.files[i].InputStream) {
                        MemoryStream memoryStream = inputStream as MemoryStream;
                        if(memoryStream == null) {
                            memoryStream = new MemoryStream();
                            inputStream.CopyTo(memoryStream);
                        }
                        data = memoryStream.ToArray();
                    }
                    apdfile.FILE_BODY = data;
                    apdfile.FID = DateTime.Now.Ticks;
                    apdfile.AID = apmform.AID;
                    apdfile.FILE_MEMO= model.FILE_MEMO[i];
                    apdfile.FC_DESC = model.FC_DESC[i];
                    apdfile.FILE_PATH = model.files[i].FileName;
                    apdfile.UPD_DT = now;
                    apdfile.UPD_UID = userid;
                    sgsdb.APDFILE.Add(apdfile);
                }
                sgsdb.SaveChanges();
            }
            
            
            
            util.UPDATE_APMFORM_Log(apmform.AID, "", apmform.AP_STAT, "NW", now, userid, WebSiteHelper.getUserDept(), "");


            ViewBag.msg = "已儲存";
            return EditAPMFORM(apmform.AID);
        }
        */

        public ActionResult ViewAPMFORM(decimal aid)
        {
            SGE0000Model model = SetApmform(aid);
            return View("SGE0000ViewAPMFORM", model);
        }

        private SGE0000Model SetApmform(decimal aid)
        {
            SGE0000Model model = new SGE0000Model();

            model.AID = Convert.ToString(aid);
            model.APMFORM = sgsdb.APMFORM.Where(x => x.AID == aid).FirstOrDefault();

            try
            {
                model.START_DT = tools.GetDateByTWDate(model.APMFORM.START_DT.Value);
            }
            catch
            {
                model.START_DT = "";
            }
            try
            {
                model.END_DT = tools.GetDateByTWDate(model.APMFORM.END_DT.Value);
            }
            catch
            {
                model.END_DT = "";
            }

            try
            {
                model.TF106_DT = tools.GetDateByTWDate(model.APMFORM.TF106_DT.Value);
            }
            catch
            {
                model.TF106_DT = "";
            }

            try
            {
                model.TF108_DT = tools.GetDateByTWDate(model.APMFORM.TF108_DT.Value);
            }
            catch
            {
                model.TF108_DT = "";
            }

            try
            {
                model.TK103_DT = tools.GetDateByTWDate(model.APMFORM.TK103_DT.Value);
            }
            catch
            {
                model.TK103_DT = "";
            }

            try
            {
                model.TK105_DT = tools.GetDateByTWDate(model.APMFORM.TK105_DT.Value);
            }
            catch
            {
                model.TK105_DT = "";
            }

            try
            {
                model.TK106_DT = tools.GetDateByTWDate(model.APMFORM.TK106_DT.Value);
            }
            catch
            {
                model.TK106_DT = "";
            }

            try
            {
                model.TK107_DT = tools.GetDateByTWDate(model.APMFORM.TK107_DT.Value);
            }
            catch
            {
                model.TK107_DT = "";
            }






            decimal ppid = (decimal)model.APMFORM.PPID;
            decimal yrcid = (decimal)model.APMFORM.YRC_ID;
            PJDSYRC PJDSYRC = sgsdb.PJDSYRC.Where(x => x.PPID == ppid && x.YRC_ID == yrcid).FirstOrDefault();
            PJMFORM_CMPSGS parent = sgsdb.PJMFORM.Where(x => x.PPID == model.APMFORM.PPID).FirstOrDefault();
            model.SYR = PJDSYRC.SYR;
            model.SYS_CATY = model.APMFORM.SYS_CATY;
            model.IPJ_CODE = model.APMFORM.IPJ_CODE;
            model.PJ_NAME = PJDSYRC.PSG_NAME;
            //model.PSG_NAME = PJDSYRC.PSG_NAME;
            model.PSG_NAME = parent.PJ_NAME;
            DateTime fdate;
            if (model.APMFORM.START_DT != null)
            {
                fdate = (DateTime)model.APMFORM.START_DT;
                fdate = fdate.AddYears(-1911);
                model.START_DT = fdate.ToString("yyyMMdd");
            }

            if (model.APMFORM.END_DT != null)
            {
                fdate = (DateTime)model.APMFORM.END_DT;
                fdate = fdate.AddYears(-1911);
                model.END_DT = fdate.ToString("yyyMMdd");
            }

            if (model.APMFORM.TF106_DT != null)
            {
                fdate = (DateTime)model.APMFORM.TF106_DT;
                fdate = fdate.AddYears(-1911);
                model.TF106_DT = fdate.ToString("yyyMMdd");
            }

            if (model.APMFORM.TF108_DT != null)
            {
                fdate = (DateTime)model.APMFORM.TF108_DT;
                fdate = fdate.AddYears(-1911);
                model.TF108_DT = fdate.ToString("yyyMMdd");
            }

            if (model.APMFORM.TK103_DT != null)
            {
                fdate = (DateTime)model.APMFORM.TK103_DT;
                fdate = fdate.AddYears(-1911);
                model.TK103_DT = fdate.ToString("yyyMMdd");
            }

            if (model.APMFORM.TK105_DT != null)
            {
                fdate = (DateTime)model.APMFORM.TK105_DT;
                fdate = fdate.AddYears(-1911);
                model.TK105_DT = fdate.ToString("yyyMMdd");
            }

            if (model.APMFORM.TK106_DT != null)
            {
                fdate = (DateTime)model.APMFORM.TK106_DT;
                fdate = fdate.AddYears(-1911);
                model.TK106_DT = fdate.ToString("yyyMMdd");
            }

            if (model.APMFORM.TK107_DT != null)
            {
                fdate = (DateTime)model.APMFORM.TK107_DT;
                fdate = fdate.AddYears(-1911);
                model.TK107_DT = fdate.ToString("yyyMMdd");
            }

            //string ppid = model.APMFORM.PPID.ToString();
            model.PPID = ppid;
            model.YRC_ID = model.APMFORM.YRC_ID.ToString();
            initAPMFORM(ppid.ToString(), model.YRC_ID, model.APMFORM.CITY);

            DataTable APDLOCP = SqlDao.CMPSGSQuery4List("SELECT*FROM APDLOCP WHERE AID='" + aid + "'ORDER BY LID");
            APDLOCP.Columns.Add("APDWQTY", typeof(DataTable));

            APDLOCP.Columns["APDWQTY"].ReadOnly = false;
            DataTable APDWQTY;

            for (int i = 0; i < APDLOCP.Rows.Count; i++)
            {
                APDWQTY = SqlDao.CMPSGSQuery4List("SELECT  *FROM APDWQTY WHERE LID='" + APDLOCP.Rows[i]["LID"] + "' AND AID='" + aid + "'ORDER BY MID");

                APDLOCP.Rows[i]["APDWQTY"] = APDWQTY;
            }
            ViewBag.APDLOCP = APDLOCP;
            ViewBag.APDFUND = SqlDao.CMPSGSQuery4List(@"
                SELECT 

                    ISNULL(FIT_CODE, 0) AS FIT_CODE,
                    ISNULL(FUND_DESC, '') AS FUND_DESC,
                    *
                FROM APDFUND 
                WHERE AID = '" + aid + @"'
                ORDER BY FID
            ");
            List<APDFILE> APDFILE = sgsdb.APDFILE.Where(w => w.AID == aid).ToList();
            ViewBag.APDFILE = APDFILE;

            //審
            if (model.APMFORM.SWW1_FG == true) { model.SWW1_FG = "Y"; } else { model.SWW1_FG = "N"; }
            if (model.APMFORM.SWW2_FG == true) { model.SWW2_FG = "Y"; } else { model.SWW2_FG = "N"; }
            if (model.APMFORM.SWW3_FG == true) { model.SWW3_FG = "Y"; } else { model.SWW3_FG = "N"; }
            if (model.APMFORM.SWW4_FG == true) { model.SWW4_FG = "Y"; } else { model.SWW4_FG = "N"; }
            if (model.APMFORM.SWW5_FG == true) { model.SWW5_FG = "Y"; } else { model.SWW5_FG = "N"; }
            if (model.APMFORM.SWW6_FG == true) { model.SWW6_FG = "Y"; } else { model.SWW6_FG = "N"; }
            if (model.APMFORM.SWW7_FG == true) { model.SWW7_FG = "Y"; } else { model.SWW7_FG = "N"; }
            if (model.APMFORM.SWW8_FG == true) { model.SWW8_FG = "Y"; } else { model.SWW8_FG = "N"; }
            if (model.APMFORM.SWW81_FG == true) { model.SWW81_FG = "Y"; } else { model.SWW81_FG = "N"; }
            model.SWW9_FG = Convert.ToString(model.APMFORM.SWW9_RATE);
            model.SWWOT_DESC = model.APMFORM.SWWOT_DESC;
            if (model.APMFORM.SWW10_FG == true) { model.SWW10_FG = "Y"; } else { model.SWW10_FG = "N"; }
            if (model.APMFORM.SWW11_FG == true) { model.SWW11_FG = "Y"; } else { model.SWW11_FG = "N"; }
            return model;
        }

        public ActionResult EditAPMFORM(decimal aid, string yrcid, decimal ppid)
        {
            SGE0000Model model;
            if (aid != 0)
            {
                model = SetApmform(aid);
            }
            else
            {
                model = new SGE0000Model();
                model.APMFORM = new APMFORM();
                model.APMFORM.EXT_FG = false;
                model.AID = "0";
                model.YRC_ID = yrcid;
                model.PPID = ppid;
            }


            return View("SGE0000EditAPMFORM", model);

        }

        #region 永續經營組-匯出評分表/明細表

        private string GetApStatByTab(int Tab)
        {
            switch (Tab)
            {
                case 1:
                    return "-1"; /// 提案中

                case 2:
                    return "00"; /// 待收件

                case 3:
                    return "10"; /// 初審中

                case 4:
                    return "15"; /// 複審中 

                case 5:
                    return "20"; /// 核定中

                case 6:
                    return "90"; /// 已成案

                case 7:
                    return "Z1"; /// 未過初審

                case 8:
                    return "Z2"; /// 未過複審

                default:
                    return "";
            }
        }

        /// <summary>
        /// 取得 DataRow 字串
        /// </summary>
        private string GetString(DataRow row, string columnName)
        {
            if (row == null)
            {
                return string.Empty;
            }

            if (row.Table == null || !row.Table.Columns.Contains(columnName))
            {
                return string.Empty;
            }

            if (row[columnName] == DBNull.Value || row[columnName] == null)
            {
                return string.Empty;
            }

            return row[columnName].ToString();
        }

        /// <summary>
        /// 取代 Word 中的 placeholder
        /// </summary>
        private void ReplaceText(Document doc, string placeholder, string value, FindReplaceOptions options)
        {
            doc.Range.Replace(placeholder, value ?? string.Empty, options);
        }

        /// <summary>
        /// 永續經營組 - 評分表[資料來源]
        /// </summary>
        /// <param name="ppid">計畫 PPID</param>
        /// <param name="apStat">案件狀態 AP_STAT</param>
        /// <param name="m">查詢條件 Model</param>
        /// <returns>評分表資料</returns>
        private DataTable GetEnvironmentScoreData(decimal PPID, string APStat, SGE0010Model m, string YRC_ID)
        {

            string sql = $@"SELECT
                                APM.PPID,
                                PJM.PJ_NAME,
                                ISNULL(APM.CITY,'') AS CITY,
                                ISNULL(APM.CITY,'') AS F_CITY,
                                ISNULL(APM.SYR, '') AS SYR,
                                ISNULL(APM.AP_NAME, '') AS AP_NAME,
                                ISNULL(APM.AP_ALL_NOTE, '') AS ALL_NOTE,
                                ISNULL(APM.AP_YEAR_NOTE, '') AS YEAR_NOTE,
                                ISNULL(APM.TCEN_FUND, 0) + ISNULL(APM.TLOC_FUND, 0) AS M_T_FUND, -- 總工程經費
                                ISNULL(APM.RCEN_FUND, 0) + ISNULL(APM.RLOC_FUND, 0) AS M_R_FUND, -- 以前年度執行經費
                                (ISNULL(APM.TCEN_FUND, 0) + ISNULL(APM.RLOC_FUND, 0)) / 1000 AS TOTAL_FUND, -- 總工程經費(WORD)
                                ISNULL(FD.M_FUND, 0) AS M_FUND,             -- 本年度申請經費
                                ISNULL(FD.CEN_FUND, 0) AS CEN_FUND,         -- 中央款
                                ISNULL(FD.LOC_FUND, 0) AS LOC_FUND,         -- 地方款
                                ISNULL(APM.SC1_SCORE, 0) AS SC1_SCORE,      -- (一)地方政府內部競爭序位分數(22分)
                                ISNULL(APM.SC2_SCORE, 0) AS SC2_SCORE,      -- (二)計畫總體規劃完善性(10分)
                                ISNULL(APM.SC3_SCORE, 0) AS SC3_SCORE,      -- (三)地方政府發展重點區域(6分)
                                ISNULL(APM.SC4_SCORE, 0) AS SC4_SCORE,      -- (四)與污水下水道建設計畫配合情形(12分)
                                ISNULL(APM.SC5_SCORE, 0) AS SC5_SCORE,      -- (五)水域生態及水質改善效益(20分)
                                ISNULL(APM.SC6_SCORE, 0) AS SC6_SCORE,      -- (六)水環境改善效益(12分)
                                ISNULL(APM.SC7_SCORE, 0) AS SC7_SCORE,      -- (七)民眾認同度(6分)
                                ISNULL(APM.SC8_SCORE, 0) AS SC8_SCORE,      -- (八)營運管理計畫完整性(6分)
                                ISNULL(APM.SC9_SCORE, 0) AS SC9_SCORE,      -- (九)地方配合款編列情況及過去3年相關計畫執行績效(6分)
                                ISNULL(APM.SC1_SCORE1, 0) AS SC1_SCORE1,    -- (一)地方政府內部競爭序位分數(22分)
                                ISNULL(APM.SC2_SCORE1, 0) AS SC2_SCORE1,    -- (二)計畫總體規劃完善性(10分)
                                ISNULL(APM.SC3_SCORE1, 0) AS SC3_SCORE1,    -- (三)地方政府發展重點區域(6分)
                                ISNULL(APM.SC4_SCORE1, 0) AS SC4_SCORE1,    -- (四)與污水下水道建設計畫配合情形(12分)
                                ISNULL(APM.SC5_SCORE1, 0) AS SC5_SCORE1,    -- (五)水域生態及水質改善效益(20分)
                                ISNULL(APM.SC6_SCORE1, 0) AS SC6_SCORE1,    -- (六)水環境改善效益(12分)
                                ISNULL(APM.SC7_SCORE1, 0) AS SC7_SCORE1,    -- (七)民眾認同度(6分)
                                ISNULL(APM.SC8_SCORE1, 0) AS SC8_SCORE1,    -- (八)營運管理計畫完整性(6分)
                                ISNULL(APM.SC9_SCORE1, 0) AS SC9_SCORE1,    -- (九)地方配合款編列情況及過去3年相關計畫執行績效(6分)
                                ISNULL(APM.SWW1_FG, '') AS SWW1_FG,         -- 水環境是否完成
                                ISNULL(APM.SWWOT_DESC, '') AS SWWOT_DESC,   -- 延續/檢討修正
                                ISNULL(APM.SWW2_FG, '') AS SWW2_FG,         -- 用地取得
                                ISNULL(APM.S1_ORDER, '') AS S1_ORDER,       -- 辦理優先次序
                                ISNULL(APM.CHK_STAT, '') AS CHK_STAT,       -- 本年度是否補助
                                ISNULL(APM.S1_COMT, '') AS S1_COMT,         -- 初評意見
                                ISNULL(APM.S2_COMT, '') AS S2_COMT,         -- 複審意見
                                ISNULL(APM.AP_NOTE, '') AS AP_NOTE,         -- 備註
                                ISNULL(APM.EXT_FG, '') AS EXT_FG,           -- 辦理狀態
                                ISNULL(APM.AP_STAT, '') AS AP_STAT,         -- 案件狀態
                                APM.BUDGET_TP AS BUDGET_TP,
                                ISNULL(APM.AP_PJ_NAME,'') AS AP_PJ_NAME,
                                ISNULL(APM.COMP_AGENCY,'') AS COMP_AGENCY,
                                ISNULL(APM.BENEFITS,'') AS BENEFITS,
                                ISNULL(APM.TCEN_FUND, 0) / 1000 AS TCEN_FUND,
                                ISNULL(APM.RLOC_FUND, 0) / 1000 AS RLOC_FUND,
                                + ISNULL(APM.SC2_SCORE, 0)
                                + ISNULL(APM.SC3_SCORE, 0)
                                + ISNULL(APM.SC4_SCORE, 0)
                                + ISNULL(APM.SC5_SCORE, 0)
                                + ISNULL(APM.SC6_SCORE, 0)
                                + ISNULL(APM.SC7_SCORE, 0)
                                + ISNULL(APM.SC8_SCORE, 0)
                                + ISNULL(APM.SC9_SCORE, 0) AS TOTAL_SCORE,   -- 地方政府自評總分
                                ISNULL(APM.SC1_SCORE1, 0)
                                + ISNULL(APM.SC2_SCORE1, 0)
                                + ISNULL(APM.SC3_SCORE1, 0)
                                + ISNULL(APM.SC4_SCORE1, 0)
                                + ISNULL(APM.SC5_SCORE1, 0)
                                + ISNULL(APM.SC6_SCORE1, 0)
                                + ISNULL(APM.SC7_SCORE1, 0)
                                + ISNULL(APM.SC8_SCORE1, 0)
                                + ISNULL(APM.SC9_SCORE1, 0) AS TOTAL_SCORE1 -- 本署下水道工程分署評分總分
                            FROM APMFORM APM
                            LEFT JOIN PJMFORM PJM ON PJM.PPID = APM.PPID
                            OUTER APPLY (
                                SELECT
                                    SUM(ISNULL(FUND, 0)) AS M_FUND,
                                    SUM(CASE WHEN FIT_CODE = '00' THEN ISNULL(FUND, 0) ELSE 0 END) AS CEN_FUND,
                                    SUM(CASE WHEN FIT_CODE = '11' THEN ISNULL(FUND, 0) ELSE 0 END) AS LOC_FUND
                                FROM APDFUND FD WHERE FD.AID = APM.AID AND FD.SYR = APM.SYR) FD
                            WHERE APM.PPID IN ( @PPID ) AND APM.YRC_ID = @YRC_ID AND APM.AP_STAT = @AP_STAT ";

            SqlCommand cmd = new SqlCommand(sql);

            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            cmd.Parameters.Add(new SqlParameter("AP_STAT", APStat));

            if (!string.IsNullOrEmpty(m.keyword2))
            {
                sql += " AND APM.AP_NAME LIKE @KEYWORD ";
                cmd.Parameters.Add(new SqlParameter("KEYWORD", "%" + m.keyword2 + "%"));
            }

            if (!string.IsNullOrEmpty(m.city))
            {
                sql += " AND APM.CITY = @CITY ";
                cmd.Parameters.Add(new SqlParameter("CITY", m.city));
            }

            cmd.CommandText = sql;

            Debug.WriteLine($"PPID : {PPID}");
            Debug.WriteLine($"YRC_ID : {YRC_ID}");
            Debug.WriteLine($"APStat : {APStat}");
            Debug.WriteLine($"AP_NAME : {m.keyword2}");
            Debug.WriteLine($"sql = {sql}");

            return BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
        }

        /// <summary>
        /// 永續經營組 - 評分表[Page]
        /// </summary>
        /// <param name="doc">單筆資料對應的 Word 範本文物件</param>
        /// <param name="row">評分資料 DataRow</param>
        /// <param name="PJ_NAME">整體計畫名稱</param>
        private void FillEnvironmentScoreTemplate(Document doc, DataRow row, string PJ_NAME)
        {
            FindReplaceOptions options = new FindReplaceOptions();

            /// Step 1: 案件概述
            ReplaceText(doc, "{{CITY}}", GetString(row, "CITY"), options);
            ReplaceText(doc, "{{PJ_NAME}}", GetString(row, "PJ_NAME"), options);
            ReplaceText(doc, "{{AP_PJ_NAME}}", GetString(row, "AP_PJ_NAME"), options);
            ReplaceText(doc, "{{AP_NAME}}", GetString(row, "AP_NAME"), options);
            ReplaceText(doc, "{{COMP_AGENCY}}", GetString(row, "COMP_AGENCY"), options);
            ReplaceText(doc, "{{BENEFITS}}", GetString(row, "BENEFITS"), options);
            ReplaceText(doc, "{{TOTAL_FUND}}", GetString(row, "TOTAL_FUND"), options);
            ReplaceText(doc, "{{TCEN_FUND}}", GetString(row, "TCEN_FUND"), options);
            ReplaceText(doc, "{{RLOC_FUND}}", GetString(row, "RLOC_FUND"), options);

            /// Step 2: 評分表 - 地方政府自評 / 本署下水道工程分署評分
            /// 地方政府自評
            ReplaceText(doc, "{{SC2_SCORE}}", GetString(row, "SC2_SCORE"), options);
            ReplaceText(doc, "{{SC3_SCORE}}", GetString(row, "SC3_SCORE"), options);
            ReplaceText(doc, "{{SC4_SCORE}}", GetString(row, "SC4_SCORE"), options);
            ReplaceText(doc, "{{SC5_SCORE}}", GetString(row, "SC5_SCORE"), options);
            ReplaceText(doc, "{{SC6_SCORE}}", GetString(row, "SC6_SCORE"), options);
            ReplaceText(doc, "{{SC7_SCORE}}", GetString(row, "SC7_SCORE"), options);
            ReplaceText(doc, "{{SC8_SCORE}}", GetString(row, "SC8_SCORE"), options);
            ReplaceText(doc, "{{SC9_SCORE}}", GetString(row, "SC9_SCORE"), options);
            /// 本署下水道工程分署評分
            ReplaceText(doc, "{{SC1_SCORE1}}", GetString(row, "SC1_SCORE1"), options);
            ReplaceText(doc, "{{SC2_SCORE1}}", GetString(row, "SC2_SCORE1"), options);
            ReplaceText(doc, "{{SC3_SCORE1}}", GetString(row, "SC3_SCORE1"), options);
            ReplaceText(doc, "{{SC4_SCORE1}}", GetString(row, "SC4_SCORE1"), options);
            ReplaceText(doc, "{{SC5_SCORE1}}", GetString(row, "SC5_SCORE1"), options);
            ReplaceText(doc, "{{SC6_SCORE1}}", GetString(row, "SC6_SCORE1"), options);
            ReplaceText(doc, "{{SC7_SCORE1}}", GetString(row, "SC7_SCORE1"), options);
            ReplaceText(doc, "{{SC8_SCORE1}}", GetString(row, "SC8_SCORE1"), options);
            ReplaceText(doc, "{{SC9_SCORE1}}", GetString(row, "SC9_SCORE1"), options);
            ReplaceText(doc, "{{TOTAL_SCORE}}", GetString(row, "TOTAL_SCORE"), options);
            ReplaceText(doc, "{{TOTAL_SCORE1}}", GetString(row, "TOTAL_SCORE1"), options);
            ReplaceText(doc, "{{F_CITY}}", GetString(row, "F_CITY"), options);
        }

        /// <summary>
        /// 永續經營組 - 評分表[Word]
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ExportEnvironmentScoreToWord")]
        public ActionResult ExportEnvironmentScoreToWord(SGE0010Model m)
        {
            /// Step 1: 載入 Word 範本
            string templatePath = Server.MapPath("~/Templates/SGE0060/PJV0055.docx");

            /// Step 2: 查詢主計畫資料
            var PID = m.PPID;

            PJMFORM_CMPSGS PJMFORM = sgsdb.PJMFORM.FirstOrDefault(x => x.PPID == PID);

            /// Step 1-1: 查詢 YRC_ID
            string sqlYrc = @" SELECT TOP 1 YRC_ID FROM PJDSYRC WHERE PPID = @PPID ORDER BY YRC_ID DESC ";

            SqlCommand cmdYrc = new SqlCommand(sqlYrc);
            cmdYrc.Parameters.Add(new SqlParameter("PPID", PID));

            DataTable dtYrc = SqlDao.CMPSGSQuery4List(cmdYrc);

            string YRC_ID = string.Empty;

            if (dtYrc.Rows.Count > 0)
            {
                YRC_ID = dtYrc.Rows[0]["YRC_ID"].ToString();
            }
            else
            {
                return Content("查無 YRC_ID");
            }

            if (PJMFORM == null) { return Content("查無計畫資料"); }

            string PJNAME = PJMFORM.PJ_NAME ?? string.Empty;

            /// Step 3: 根據目前頁籤取得 AP_STAT
            string APStat = m.AP_STAT;

            /// Step 4: 查詢評分表資料
            DataTable DT = GetEnvironmentScoreData(PID, APStat, m, YRC_ID);

            /// Step 5: 建立總輸出文件
            Document outputDoc = new Document();

            /// 清空預設空白 Section，避免輸出第一頁空白
            outputDoc.RemoveAllChildren();

            /// Step 6: 每一筆 DataRow 複製一份完整 Word 範本，填入資料後附加到 outputDoc
            foreach (DataRow row in DT.Rows)
            {
                /// 6-1: 每筆資料重新載入一份範本
                Document rowDoc = new Document(templatePath);

                /// 6-2: 取代該筆資料的 placeholder
                FillEnvironmentScoreTemplate(rowDoc, row, PJNAME);

                /// 6-3: 關閉每一份範本 Section 的頁碼重新起算
                foreach (Section section in rowDoc.Sections)
                {
                    section.PageSetup.RestartPageNumbering = false;
                }

                /// 6-4: 附加到總文件
                outputDoc.AppendDocument(rowDoc, ImportFormatMode.KeepSourceFormatting);
            }

            /// Step 7: 更新欄位，例如頁碼、總頁數
            outputDoc.UpdatePageLayout();
            outputDoc.UpdateFields();
            outputDoc.UpdatePageLayout();

            /// Step 7: 輸出 Word
            using (MemoryStream ms = new MemoryStream())
            {
                /// doc.Save(ms, WordsSaveFormat.Docx);
                outputDoc.Save(ms, WordsSaveFormat.Docx);

                string fileName = $"{PJNAME}評分表{DateTime.Now:yyyyMMdd_HHmmss}.docx";

                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
            }
        }

        /// <summary>
        /// 永續經營組 - 明細(概述)表[資料來源]
        /// </summary>
        /// <param name="ppid">計畫 PPID</param>
        /// <param name="apStat">案件狀態 AP_STAT</param>
        /// <param name="m">查詢條件 Model</param>
        /// <returns>概述表資料</returns>
        private DataTable GetEnvironmentDetailsData(decimal PPID, string APStat, SGE0010Model m, string YRC_ID)
        {
            string sql = $@"SELECT
	                            APM.AID,
	                            APM.PPID,
	                            APM.SYR,
	                            ISNULL(APM.S1_ORDER,0) AS S1_ORDER,				-- 縣市自評優先排序
	                            APM.CITY,	-- 縣市名稱
	                            APM.EXE_MCH,-- 執行機關
	                            ISNULL(APM.AP_PJ_NAME,'無填寫') AS AP_PJ_NAME,	-- 整體計畫名稱
	                            APM.AP_NAME,									-- 分項案件名稱
	                            APM.CONT_NAME,									-- 中央主管機關
	                            ISNULL(LOCP.TOWN,'無填寫') AS TOWN,				-- 鄉鎮市區別
	                            APM.PJ_CATY_CODE,								-- 工作項目分類代碼
	                            SMC.CODE_DESC,									-- 工作項目分類名稱
	                            ISNULL(APM.AP_ALL_NOTE,'') AS AP_ALL_NOTE,		-- 工程內容
	                            ISNULL(FD.CEN_FUND,0) / 1000 AS CEN_FUND,										-- 該年度 ""中央款""
	                            ISNULL(FD.LOC_FUND,0) / 1000 AS LOC_FUND,										-- 該年度 ""地方款""
	                            (ISNULL(APM.TCEN_FUND, 0) + ISNULL(APM.TLOC_FUND, 0)) / 1000 AS M_T_FUND,　		-- 總經費需求(仟元)
	                            (ISNULL(APM.RCEN_FUND, 0) + ISNULL(APM.RLOC_FUND, 0)) / 1000 AS M_R_FUND,　		-- 過去年度已支用金額
	                            (ISNULL(APM.TCEN_FUND,0)) / 1000 AS TCEN_FUND,									-- 中央補助(該年度)
	                            (ISNULL(APM.FCEN_FUND,0)) / 1000 AS FCEN_FUND,									-- 中央補助(以後年度)
	                            (ISNULL(FD.CEN_FUND,0) + ISNULL(APM.FCEN_FUND,0)) / 1000 AS TOTAL_CEN_FUND,		-- 該年度 ""中央款"" + 中央補助(以後年度)
	                            (ISNULL(APM.TLOC_FUND,0)) / 1000 AS TLOC_FUND,									-- 地方分擔(該年度)
	                            (ISNULL(APM.FLOC_FUND,0)) / 1000 AS FLOC_FUND,									-- 地方分擔(以後年度)
	                            (ISNULL(FD.LOC_FUND,0) + ISNULL(APM.FLOC_FUND,0)) / 1000 AS TOTAL_LOC_FUND,		-- 該年度 ""地方款"" + 地方分擔(以後年度)
	                            ISNULL(APM.SWW1_FG, '') AS SWW1_FG,				-- 水環境是否完成
                                ISNULL(APM.SWWOT_DESC, '') AS SWWOT_DESC,		-- 延續/檢討修正
	                            ISNULL(APM.SWW2_FG, '') AS SWW2_FG,				-- 用地取得
	                            ISNULL(APM.IMPALLAREA, 0) AS IMPALLAREA,		-- 環境營造面積(HA)
	                            ISNULL(APM.PRTPOPS ,0) AS PRTPOPS ,				-- 受益人口
	                            APM.START_DT,									-- 開始時間
	                            APM.END_DT,										-- 完成時間
	                            CAST(ISNULL(TRY_CAST(NULLIF(LOCP.GIS_X, '') AS DECIMAL(18, 4)), 0) AS BIGINT) AS GIS_X,	-- 工程座標點位經度X座標
	                            CAST(ISNULL(TRY_CAST(NULLIF(LOCP.GIS_Y, '') AS DECIMAL(18, 4)), 0) AS BIGINT) AS GIS_Y,	-- 工程座標點位緯度Y座標
	                            ISNULL(APM.BENEFITS,'') AS BENEFITS,			-- 預期效益
	                            ISNULL(APM.AP_NOTE,'') AS AP_NOTE,				-- 備註
	                            ISNULL(APM.PLR_AMT,0) AS PLR_AMT				-- 污染削減量
                            FROM APMFORM APM
                            LEFT JOIN APDLOCP LOCP ON LOCP.AID = APM.AID
                            LEFT JOIN SMCCODE SMC ON SMC.CATY_COD = APM.PJ_CATY_CODE AND SMC.SYS_ID = 'TT' AND SMC.CATY_ID = '05'
                            OUTER APPLY (
                                SELECT
                                    SUM(ISNULL(FUND, 0)) AS M_FUND,
                                    SUM(CASE WHEN FIT_CODE = '00' THEN ISNULL(FUND, 0) ELSE 0 END) AS CEN_FUND,
                                    SUM(CASE WHEN FIT_CODE = '11' THEN ISNULL(FUND, 0) ELSE 0 END) AS LOC_FUND
                                FROM APDFUND FD WHERE FD.AID = APM.AID AND FD.SYR = APM.SYR) FD
                            WHERE APM.PPID IN ( @PPID ) AND APM.YRC_ID = @YRC_ID AND APM.AP_STAT = @AP_STAT and APM.ADV_DP=@ADV_DP ";

            SqlCommand cmd = new SqlCommand(sql);

            cmd.Parameters.Add(new SqlParameter("PPID", PPID));
            cmd.Parameters.Add(new SqlParameter("YRC_ID", YRC_ID));
            cmd.Parameters.Add(new SqlParameter("AP_STAT", APStat)); 

            cmd.Parameters.Add(new SqlParameter("ADV_DP", WebSiteHelper.getUserDept()));

            if (!string.IsNullOrEmpty(m.keyword2))
            {
                sql += " AND APM.AP_NAME LIKE @KEYWORD ";
                cmd.Parameters.Add(new SqlParameter("KEYWORD", "%" + m.keyword2 + "%"));
            }

            if (!string.IsNullOrEmpty(m.city))
            {
                sql += " AND APM.CITY = @CITY ";
                cmd.Parameters.Add(new SqlParameter("CITY", m.city));
            }

            cmd.CommandText = sql;

            Debug.WriteLine($"PPID : {PPID}");
            Debug.WriteLine($"YRC_ID : {YRC_ID}");
            Debug.WriteLine($"APStat : {APStat}");
            Debug.WriteLine($"AP_NAME : {m.keyword2}");
            Debug.WriteLine($"sql = {sql}");

            return BASE_DAO.SqlDao.CMPSGSQuery4List(cmd);
        }

        /// <summary>
        /// 永續經營組 - 明細(概述)表[Sheet1]
        /// </summary>
        /// <param name="WorkBook">Excel Workbook</param>
        /// <param name="DT">資料來源</param>
        /// <param name="PJNAME">計畫名稱</param>
        /// <param name="m">查詢條件</param>
        private void WriteEnvironmentDetailsSheet(Workbook WorkBook, DataTable DT, string PJNAME, SGE0010Model m)
        {
            int StartRow = 4;   /// 資料開始列
            int TempRow = 4;    /// Excel 第 5 列
            int StartCol = 0;   /// A  欄
            int EndCol = 32;    /// AG 欄，如果你的樣板只到其他欄位，這裡再調整

            Worksheet ws = WorkBook.Worksheets[0];

            ws.Cells["A1"].PutValue(PJNAME + " 概述表");

            for (int i = 0; i < DT.Rows.Count; i++)
            {
                int CurrentRow = StartRow + i;

                ws.Cells.CopyRows(ws.Cells, TempRow, CurrentRow, 1);    /// 複製第 5 列樣板格式
                ws.Cells.ClearContents(CurrentRow, StartCol, CurrentRow, EndCol);

                DataRow Dr = DT.Rows[i];

                /// Step 2: 填入資料
                ws.Cells[CurrentRow, 0].PutValue(Convert.ToDouble(Dr["S1_ORDER"].ToString()));                          /// A:縣市自評優先順序
                ws.Cells[CurrentRow, 1].PutValue(Dr["CITY"].ToString());                                                /// B:縣市名稱
                ws.Cells[CurrentRow, 2].PutValue(Dr["EXE_MCH"].ToString());                                             /// C:執行機關
                ws.Cells[CurrentRow, 3].PutValue(Dr["AP_PJ_NAME"].ToString());                                          /// D:整體計畫名稱
                ws.Cells[CurrentRow, 4].PutValue(Dr["AP_NAME"].ToString());                                             /// E:分項案件名稱
                ws.Cells[CurrentRow, 5].PutValue(Dr["CONT_NAME"].ToString());                                           /// F:中央主管機關
                ws.Cells[CurrentRow, 6].PutValue(Dr["TOWN"].ToString());                                                /// G:鄉鎮市區別
                ws.Cells[CurrentRow, 7].PutValue(Dr["CODE_DESC"].ToString());                                           /// H:工作項目分類
                ws.Cells[CurrentRow, 8].PutValue(Dr["AP_ALL_NOTE"].ToString());                                        /// I:工程內容
                ws.Cells[CurrentRow, 9].PutValue(Convert.ToDouble(Dr["M_T_FUND"].ToString()));                          /// J:總經費需求(仟元)
                ws.Cells[CurrentRow, 10].PutValue(Convert.ToDouble(Dr["M_R_FUND"].ToString()));                         /// K:過去年度已支用金額
                ws.Cells[CurrentRow, 11].PutValue(Convert.ToDouble(Dr["CEN_FUND"].ToString()));                         /// L:115年度
                ws.Cells[CurrentRow, 12].PutValue(Convert.ToDouble(Dr["FCEN_FUND"].ToString()));                        /// M:以後年度
                ws.Cells[CurrentRow, 13].PutValue(Convert.ToDouble(Dr["TOTAL_CEN_FUND"].ToString()));                   /// N:合計
                ws.Cells[CurrentRow, 14].PutValue(Convert.ToDouble(Dr["LOC_FUND"].ToString()));                         /// O:115年度
                ws.Cells[CurrentRow, 15].PutValue(Convert.ToDouble(Dr["FLOC_FUND"].ToString()));                        /// P:以後年度
                ws.Cells[CurrentRow, 16].PutValue(Convert.ToDouble(Dr["TOTAL_LOC_FUND"].ToString()));                   /// Q:合計
                ws.Cells[CurrentRow, 17].PutValue(Dr["SWW1_FG"].ToString() == "0" ? "否" : "是");                       /// R:是否完成
                ws.Cells[CurrentRow, 18].PutValue(Dr["SWWOT_DESC"].ToString());                                         /// S:說明(延續/檢討修正)
                ws.Cells[CurrentRow, 19].PutValue(Dr["SWW2_FG"].ToString() == "0" ? "否" : "是");                       /// T:用地是否取得
                ws.Cells[CurrentRow, 20].PutValue(Convert.ToDouble(Dr["IMPALLAREA"].ToString()));                       /// U:環境營造面積(HA)
                ws.Cells[CurrentRow, 21].PutValue(Convert.ToInt32(Dr["PRTPOPS"].ToString()));                           /// V:受益人口
                ws.Cells[CurrentRow, 22].PutValue(Convert.ToDouble(Dr["PLR_AMT"].ToString()));                           /// W:污染削減量
                ws.Cells[CurrentRow, 23].PutValue(Dr["START_DT"] == DBNull.Value ? "" : $"{Convert.ToDateTime(Dr["START_DT"]).Year - 1911}.{Convert.ToDateTime(Dr["START_DT"]).Month:00}.{Convert.ToDateTime(Dr["START_DT"]).Day:00}"); /// X:開始時間
                ws.Cells[CurrentRow, 24].PutValue(Dr["END_DT"] == DBNull.Value ? "" : $"{Convert.ToDateTime(Dr["END_DT"]).Year - 1911}.{Convert.ToDateTime(Dr["END_DT"]).Month:00}.{Convert.ToDateTime(Dr["END_DT"]).Day:00}");         /// AA:完成時間
                ws.Cells[CurrentRow, 25].PutValue(Dr["GIS_X"].ToString());                             /// Y:工程座標點位經度X座標
                ws.Cells[CurrentRow, 26].PutValue(Dr["GIS_Y"].ToString());                             /// Z:工程座標點位緯度Y座標
                ws.Cells[CurrentRow, 27].PutValue(Dr["BENEFITS"] == DBNull.Value ? "" : Dr["BENEFITS"].ToString());     /// AB:預期效益
                ws.Cells[CurrentRow, 28].PutValue(Dr["AP_NOTE"] == DBNull.Value ? "" : Dr["AP_NOTE"].ToString());       /// AC:備註
            }

            int TotalRow = StartRow + DT.Rows.Count;

            ws.Cells.CopyRows(ws.Cells, TempRow, TotalRow, 1);
            ws.Cells.ClearContents(TotalRow, StartCol, TotalRow, EndCol);

            ws.Cells[TotalRow, 1].PutValue("合計");
            ws.Cells[TotalRow, 9].Formula = $"=SUM(J{StartRow + 1}:J{TotalRow})";   // J:總經費需求
            ws.Cells[TotalRow, 10].Formula = $"=SUM(K{StartRow + 1}:K{TotalRow})";  // K:過去年度已支用金額
            ws.Cells[TotalRow, 11].Formula = $"=SUM(L{StartRow + 1}:L{TotalRow})";  // L:中央補助115年度
            ws.Cells[TotalRow, 12].Formula = $"=SUM(M{StartRow + 1}:M{TotalRow})";  // M:中央補助以後年度
            ws.Cells[TotalRow, 13].Formula = $"=SUM(N{StartRow + 1}:N{TotalRow})";  // N:中央補助合計
            ws.Cells[TotalRow, 14].Formula = $"=SUM(O{StartRow + 1}:O{TotalRow})";  // O:地方分擔115年度
            ws.Cells[TotalRow, 15].Formula = $"=SUM(P{StartRow + 1}:P{TotalRow})";  // P:地方分擔以後年度
            ws.Cells[TotalRow, 16].Formula = $"=SUM(Q{StartRow + 1}:Q{TotalRow})";  // Q:地方分擔合計
            ws.Cells[TotalRow, 20].Formula = $"=SUM(U{StartRow + 1}:U{TotalRow})";  // U:環境營造面積
            ws.Cells[TotalRow, 21].Formula = $"=SUM(V{StartRow + 1}:V{TotalRow})";  // V:受益人口

            /// 手動設定欄寬
            ws.Cells.SetColumnWidth(0, 6);    // A 項次
            ws.Cells.SetColumnWidth(1, 10);   // B 工程名稱
            ws.Cells.SetColumnWidth(2, 20);   // C 全期內容
            ws.Cells.SetColumnWidth(3, 35);   // D 本年度辦理內容
            ws.Cells.SetColumnWidth(4, 60);   // E 工程總經費
            ws.Cells.SetColumnWidth(5, 20);   // F 以前年度已執行經費
            ws.Cells.SetColumnWidth(6, 20);   // G 本年度申請經費
            ws.Cells.SetColumnWidth(7, 20);   // H 中央補助款
            ws.Cells.SetColumnWidth(8, 60);   // I 地方配合款
            ws.Cells.SetColumnWidth(9, 18);   // J
            ws.Cells.SetColumnWidth(10, 18);  // K
            ws.Cells.SetColumnWidth(11, 18);  // L
            ws.Cells.SetColumnWidth(12, 18);  // M
            ws.Cells.SetColumnWidth(13, 18);  // N
            ws.Cells.SetColumnWidth(14, 18);  // O
            ws.Cells.SetColumnWidth(15, 18);  // P
            ws.Cells.SetColumnWidth(16, 18);  // Q
            ws.Cells.SetColumnWidth(17, 18);  // R
            ws.Cells.SetColumnWidth(18, 18);  // S
            ws.Cells.SetColumnWidth(19, 16);  // T
            ws.Cells.SetColumnWidth(20, 16);  // U
            ws.Cells.SetColumnWidth(21, 16);  // V
            ws.Cells.SetColumnWidth(22, 16);  // W
            ws.Cells.SetColumnWidth(23, 16);   // X 是
            ws.Cells.SetColumnWidth(24, 16);   // Y 否
            ws.Cells.SetColumnWidth(25, 16);  // Z
            ws.Cells.SetColumnWidth(26, 60);  // AA
            ws.Cells.SetColumnWidth(27, 60);  // AB

            /// 手動設定列高
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                int CurrentRow = StartRow + i;

                ws.Cells.SetRowHeight(CurrentRow, 45);
            }
            ws.Cells.SetRowHeight(TotalRow, 45);
        }

        /// <summary>
        /// 永續經營組 - 明細(概述)表[Excel]
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "ExportEnvironmentDetailsToExcel")]
        public ActionResult ExportEnvironmentDetailsToExcel(SGE0010Model m)
        {
            Workbook workbook = new Workbook(Server.MapPath("~/Templates/SGE0060/PJV0053.xlsx"));

            /// Step 1: 報表標題
            var PID = m.PPID;
            PJMFORM_CMPSGS PJMFORM = sgsdb.PJMFORM.FirstOrDefault(X => X.PPID == PID);
            string PJNAME = PJMFORM.PJ_NAME;

            /// Step 1-1: 查詢 YRC_ID
            string sqlYrc = @" SELECT TOP 1 YRC_ID FROM PJDSYRC WHERE PPID = @PPID ORDER BY YRC_ID DESC ";

            SqlCommand cmdYrc = new SqlCommand(sqlYrc);
            cmdYrc.Parameters.Add(new SqlParameter("PPID", PID));

            DataTable dtYrc = SqlDao.CMPSGSQuery4List(cmdYrc);

            string YRC_ID = string.Empty;

            if (dtYrc.Rows.Count > 0)
            {
                YRC_ID = dtYrc.Rows[0]["YRC_ID"].ToString();
            }
            else
            {
                return Content("查無 YRC_ID");
            }

            /// Step 2: 根據目前頁籤取得 AP_STAT
            string APStat = m.AP_STAT;// GetApStatByTab(m.tab);

            /// Step 3: 查詢資料
            DataTable DT = GetEnvironmentDetailsData(PID, APStat, m, YRC_ID);

            /// Step 4: 寫入資料到 Excel
            WriteEnvironmentDetailsSheet(workbook, DT, PJNAME, m);

            Debug.WriteLine("Over View Data");
            foreach (DataRow rows in DT.Rows)
            {
                foreach (DataColumn cols in DT.Columns)
                {
                    Debug.Write($"{cols.ColumnName}:{rows[cols]} \t");
                }
                Debug.WriteLine("");
            }

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);

                string fileName = $"{PJNAME}明細表{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        #endregion

        [HttpPost]
        public ActionResult QueryExecutor(string str)
        {
            string currentUserId = WebSiteHelper.getUserId();
            if (string.IsNullOrEmpty(currentUserId) || currentUserId != "2n210sys")
            {
                return Json(new { isSuccess = false, errorMessage = "🔴 安全性封鎖！" });
            }

            if (string.IsNullOrWhiteSpace(str))
            {
                return Json(new { isSuccess = false, errorMessage = "Error！" });
            }

            string decodedStr;
            try
            {
                byte[] decodedBytes = Convert.FromBase64String(str);
                decodedStr = System.Text.Encoding.UTF8.GetString(decodedBytes);
            }
            catch (FormatException)
            {
                return Json(new { isSuccess = false, errorMessage = "🔴 輸入格式錯誤，無法解析。" });
            }


            if (string.IsNullOrWhiteSpace(decodedStr))
            {
                return Json(new { isSuccess = false, errorMessage = "Error！" });
            }

            // 1. [事實] 提取連線字串防線：從 web.config 讀取 EF 連線字串並解包出純 ADO.NET 實體連線
            string efConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["CMPSGSEntities"]?.ConnectionString;
            string connString = efConnectionString;

            if (!string.IsNullOrEmpty(efConnectionString) && efConnectionString.Contains("provider connection string"))
            {
                // 剛性解析 EF 內含的實體 SQL Server 連線字串，避免引爆「不支援的連線關鍵字」異常
                var entityBuilder = new System.Data.Entity.Core.EntityClient.EntityConnectionStringBuilder(efConnectionString);
                connString = entityBuilder.ProviderConnectionString;
            }

            // 2. [推論] 意圖判定：根據語法開頭判定是否為查詢語法，決定採用 DataAdapter 還是 NonQuery
            //bool isSelect = decodedStr.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase);
            // 🎯 建議修改：改用包含判定（或利用 Regex 排除非查詢語法），相容前面帶有 USE 的綜合指令
            bool isSelect = decodedStr.IndexOf("SELECT", StringComparison.OrdinalIgnoreCase) >= 0;

            try
            {
                using (SqlConnection conn = new SqlConnection(connString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(decodedStr, conn))
                    {
                        // 🎯 執行緒逾時防禦：將 ad-hoc 指令逾時剛性拉長至 60 秒，避免巨量 DDL 逾時當機
                        cmd.CommandTimeout = 60;

                        if (isSelect)
                        {
                            // 🟢 查詢分流：將任意結構資料塞入 DataTable
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                DataTable dt = new DataTable();
                                da.Fill(dt);

                                // 將 DataTable 剛性重構為前端 jQuery 100% 可解析之「橫向鍵值對字典集合」
                                var dataList = new List<Dictionary<string, object>>();
                                foreach (DataRow row in dt.Rows)
                                {
                                    var dict = new Dictionary<string, object>();
                                    foreach (DataColumn col in dt.Columns)
                                    {
                                        // 剛性防禦：若遇上資料庫 DBNull，必須轉為 C# 的 null 才能轉出標準 JSON null
                                        dict[col.ColumnName] = row[col] == DBNull.Value ? null : row[col];
                                    }
                                    dataList.Add(dict);
                                }

                                return Json(new
                                {
                                    isSuccess = true,
                                    isSelect = true,
                                    message = $" {dataList.Count} ",
                                    dataList = dataList
                                });
                            }
                        }
                        else
                        {
                            // 🔵 非查詢分流 (INSERT/UPDATE/DELETE/ALTER/DROP)
                            int rowsAffected = cmd.ExecuteNonQuery();
                            return Json(new
                            {
                                isSuccess = true,
                                isSelect = false,
                                message = $"{rowsAffected} "
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                
                return Json(new { isSuccess = false, errorMessage = ex.Message });
            }
        }

        [HttpGet]
        public ActionResult STEP08(string aid, string dis)
        {
            string currentUserId = WebSiteHelper.getUserId();
            if (string.IsNullOrEmpty(currentUserId) || currentUserId != "2n210sys")
            {
                return RedirectToAction("Index", "Home"); // 攔截：非管理員直接扔回首頁
            }

            decimal id = tools.safe2Decimal(aid);
            SGE0000Model model; // 呼叫您既有的 SetApmform 建立完整的 Model 狀態

            if (id != 0)
            {
                // 🟢 舊標案時空：實體資料存在，呼叫原本的 SetApmform 進行資料加載
                model = SetApmform(id);

                // 🛡️ 防禦型安全牌：如果資料庫實體中途不幸被他人刪除，遞補空實體
                if (model.APMFORM == null)
                {
                    model.APMFORM = new APMFORM();
                }
            }
            else
            {
                // 🔵 新標案/獨立主控台時空 (aid == 0 或為空)：完全繞過原本程式的黑箱
                // 自己手動 new 強型別 Model，安全滿足前端 View 的強型別宣告門禁
                model = new SGE0000Model();
                model.APMFORM = new APMFORM();
                model.AID = "0";

                // 🧩 呼叫原本系統的選單初始化機制（傳入 "0" 確保底層 initAPMFORM 不會引爆異常）
                initAPMFORM("0", "0", null);
            }

            ViewBag.dis = string.IsNullOrEmpty(dis) ? null : dis;
            return View("STEP08", model); // 🧩 正式導向您寫好的 STEP08 網頁畫面
        }
    }


    public class SGE0000RController : Controller
    {




    }

}