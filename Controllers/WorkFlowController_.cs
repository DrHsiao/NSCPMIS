using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http.Cors;
using System.Web.Http;
using System.Data.Entity.Validation;

namespace WCals.Controllers
{
    public class FunctionResult
    {
        public int RtnCode { get; set; }
        public string RtnDesc { get; set; }
        public object Result { get; set; }
        public FunctionResult(int rtnCode, string rtnDesc, object result)
        {
            RtnCode = RtnCode;
            RtnDesc = rtnDesc;
            Result = result;
        }
        public FunctionResult()
        {

        }
    }

    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class WorkFlowController : ApiController
    {
        private readonly WCals.Models.WCalsEntities db = new Models.WCalsEntities();

        public IHttpActionResult Get()
        {
            try
            {
                ViewModels.WorkFlowViewModel model = new ViewModels.WorkFlowViewModel();
                model.Actions = db.SMCCODE.Where(f => f.SYS_ID == "FC" && f.CATY_ID == "01").ToList();
                model.Steps = db.SMCCODE.Where(f => f.SYS_ID == "FC" && f.CATY_ID == "02").ToList();
                var deptusers = (from f in db.OGCDPUR
                                 join u in db.OGCUSER on f.USR_ID equals u.USR_ID
                                 join d in db.OGCDEPT on f.DEPT_NO equals d.DEPT_NO
                                 select new ViewModels.DeptUsersViewModel { DEPT_NO = f.DEPT_NO, DEPT_NAME = d.DEPT_NAME, USR_ID = u.USR_ID, USR_NAME = u.USR_NAME }).ToList();
                model.Depts = db.OGCDEPT.ToList();
                model.DeptUsers = deptusers;
                model.Users = db.OGCUSER.ToList();
                return Ok(new { result = model, rtncode = 0, rtndesc = "資料讀取成功" });
            }
            catch (Exception e)
            {
                return Ok(new { result = "", rtncode = -1, rtndesc = e.Message });
            }
        }

        public IHttpActionResult Get(string deptno, string stepno)
        {
            //流程資料下載
            try
            {
                ViewModels.WorkFlowViewModel model = new ViewModels.WorkFlowViewModel();
                if (string.IsNullOrEmpty(deptno) == false && string.IsNullOrEmpty(stepno) == false)
                {
                    model.FCWPOOL = (from f in db.TDV_FCWPOOL
                                     where f.DEPT_NO == deptno && f.STEP_NO == stepno
                                     select f).ToList();
                }
                return Ok(new { result = model, rtncode = 0, rtndesc = "資料讀取成功" });
            }
            catch (Exception e)
            {
                return Ok(new { result = "", rtncode = -1, rtndesc = e.Message });
            }
        }
        private FunctionResult UpdateAPMFORM(string TD_CODE, string STEP_NO, string USR_ID, string DEPT_NO, string SC_NO)
        {
            var TDMFORM = db.TDMFORM.Where(f => f.TD_CODE == TD_CODE).FirstOrDefault();
            if (TDMFORM != null)
            {
                TDMFORM.STEP_NO = STEP_NO;
                TDMFORM.ADV_DP = DEPT_NO;
                TDMFORM.ADV_UID = USR_ID;
                TDMFORM.ADV_SC = SC_NO;
                TDMFORM.UPD_DT = DateTime.Now;
                TDMFORM.UPD_UID = USR_ID;

                switch (STEP_NO)
                {
                    case "02": //設計
                        TDMFORM.DSG_UID = USR_ID;
                        TDMFORM.DSG_DP = DEPT_NO;
                        TDMFORM.DSG_SC = SC_NO;
                        break;
                    case "03": //發包
                        TDMFORM.BID_UID = USR_ID;
                        TDMFORM.BID_DP = DEPT_NO;
                        TDMFORM.BID_SC = SC_NO;
                        break;
                    case "04": //施工
                        TDMFORM.CNS_UID = USR_ID;
                        TDMFORM.CNS_DP = DEPT_NO;
                        TDMFORM.CNS_SC = SC_NO;
                        break;
                    case "05": //決算
                        TDMFORM.AIT_UID = USR_ID;
                        TDMFORM.AIT_DP = DEPT_NO;
                        TDMFORM.AIT_SC = SC_NO;
                        break;
                }
                db.SaveChanges();
                return new FunctionResult(0, "處理完成", "");
            }
            else
            {
                return new FunctionResult(-1, "工程資料主檔不存在, 請確認後再試", "");
            }
        }
        private void AddFCDJOUR(string TD_CODE, string STEP_NO, string Action, string ActionDesc, string USR_ID, string DEPT_NO, string SC_NO)
        {
            Models.FCDJOUR FCDJOUR = new Models.FCDJOUR();
            FCDJOUR.TD_CODE = TD_CODE;
            FCDJOUR.ACTION_COD = Action;
            FCDJOUR.ACTION_DESC = ActionDesc;
            FCDJOUR.STEP_NO = STEP_NO;
            FCDJOUR.ACTION_DT = DateTime.Now;
            FCDJOUR.UPD_UID = USR_ID;
            FCDJOUR.DEPT_NO = DEPT_NO;
            FCDJOUR.SC_NO = SC_NO;
            db.FCDJOUR.Add(FCDJOUR);
            db.SaveChanges();
        }
        private void AddFCWPOOL(ViewModels.WorkFlowViewModel model, DateTime? ActionDt)
        {
            var FCWPOOL = new Models.FCWPOOL();
            FCWPOOL.TD_CODE = model.TD_CODE;
            FCWPOOL.ACTION_NO = model.ACTION;
            if (ActionDt == null)
            {
                FCWPOOL.ACTION_DT = DateTime.Now;
            }
            else
            {
                FCWPOOL.ACTION_DT = ActionDt;
            }
            FCWPOOL.STEP_NO = model.STEP_NO;
            FCWPOOL.DEPT_NO = model.DEPT_NO;
            FCWPOOL.SC_NO = model.SC_NO;
            FCWPOOL.UPD_UID = model.USR_ID;
            db.FCWPOOL.Add(FCWPOOL);
            db.SaveChanges();
        }
        public IHttpActionResult Get(string tdcode, string action, string userid, string deptno, string scno)
        {
            //收件動作
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    FunctionResult result = null;

                    ViewModels.WorkFlowViewModel model = new ViewModels.WorkFlowViewModel();
                    var FCWPOOL = db.TDV_FCWPOOL.Where(f => f.TD_CODE == tdcode).FirstOrDefault();

                    model.TDV_TDMFORM = db.TDV_TDMFORM.Where(f => f.TD_CODE == tdcode).ToList();
                    model.TDMFORM = db.TDMFORM.Where(f => f.TD_CODE == tdcode).ToList();
                    model.TDDYPSC = db.TDDYPSC.Where(f => f.TD_CODE == tdcode).ToList();
                    var APMFORM = db.TDMFORM.Where(f => f.TD_CODE == tdcode).FirstOrDefault();
                    if (action == "00") //收件
                    {
                        var fcwpool = db.FCWPOOL.Where(f => f.TD_CODE == tdcode).FirstOrDefault();
                        string stepNo = fcwpool.STEP_NO;
                        if (fcwpool != null)
                        {
                            db.FCWPOOL.Remove(fcwpool);
                        }
                        var FCWFILE = db.FCWFILE.Where(f => f.TD_CODE == tdcode).FirstOrDefault();
                        model.fileArray = FCWFILE.FILECONT;
                        AddFCDJOUR(tdcode, stepNo, action, "收件", userid, deptno, scno);
                        if (stepNo== "20") //收件議定(完成)特別處理將APMFORM Step改回施工中
                        {
                            //收件
                            result = UpdateAPMFORM(tdcode, "04", userid, deptno, scno);
                        }
                        else
                        {
                            result = UpdateAPMFORM(tdcode, stepNo, userid, deptno, scno);
                        }
                        if (result.RtnCode!= 0)
                        {
                            return Ok(new { result = model, rtncode = -1, rtndesc = result.RtnDesc });
                        }
                    }
                    else if (action == "02") //自辦
                    {
                        if (APMFORM.ADV_UID != userid)
                        {
                            return Ok(new { result = "", rtncode = -1, rtndesc = "案件擁有者並非案件處理人員，請確認後再進行自辦。" });
                        }

                        var fcwpool = db.FCWPOOL.Where(f => f.TD_CODE == tdcode).FirstOrDefault();
                        string stepNo = fcwpool.STEP_NO;
                        if (fcwpool != null)
                        {
                            db.FCWPOOL.Remove(fcwpool);
                        }
                        var FCWFILE = db.FCWFILE.Where(f => f.TD_CODE == tdcode).FirstOrDefault();
                        model.fileArray = FCWFILE.FILECONT;
                        AddFCDJOUR(tdcode, stepNo, action, "自辦", userid, deptno, scno);
                        result = UpdateAPMFORM(tdcode, stepNo, userid, deptno, scno);
                        if (result.RtnCode != 0)
                        {
                            return Ok(new { result = model, rtncode = -1, rtndesc = result.RtnDesc });
                        }
                    }
                    else if (action == "03") //退件
                    {
                        if (APMFORM.ADV_UID != userid)
                        {
                            return Ok(new { result = "", rtncode = -1, rtndesc = "案件擁有者並非案件處理人員，請確認後再進行退件。" });
                        }
                        var fcdjours = db.FCDJOUR.Where(f => f.TD_CODE == tdcode && (f.ACTION_COD == "01" || f.ACTION_COD == "04")).OrderByDescending(f => f.seqNo).ToList();
                        var fcdjour = fcdjours.FirstOrDefault();
                        model.TD_CODE = tdcode;
                        model.DEPT_NO = fcdjour.DEPT_NO;
                        model.STEP_NO = fcdjour.STEP_NO;
                        model.ACTION = fcdjour.ACTION_COD;
                        model.SC_NO = fcdjour.SC_NO;
                        AddFCWPOOL(model, Convert.ToDateTime(fcdjour.ACTION_DT));
                        AddFCDJOUR(tdcode, fcdjour.STEP_NO, action, "退件", userid, fcdjour.DEPT_NO, fcdjour.SC_NO);
                        result= UpdateAPMFORM(tdcode, fcdjour.STEP_NO, userid, fcdjour.DEPT_NO, fcdjour.SC_NO);
                        if (result.RtnCode != 0)
                        {
                            return Ok(new { result = model, rtncode = -1, rtndesc = result.RtnDesc });
                        }
                    }
                    else if (action == "04") //轉送
                    {
                        if (APMFORM.ADV_UID != userid)
                        {
                            return Ok(new { result = "", rtncode = -1, rtndesc = "案件擁有者並非案件處理人員，請確認後再進行轉送。" });
                        }
                        var fcdjours = db.FCDJOUR.Where(f => f.TD_CODE == tdcode).OrderByDescending(f => f.seqNo).ToList();
                        var fcdjour = fcdjours.FirstOrDefault();
                        model.TD_CODE = tdcode;
                        model.DEPT_NO = deptno;
                        model.STEP_NO = fcdjour.STEP_NO;
                        model.ACTION = fcdjour.ACTION_COD;
                        model.SC_NO = scno;
                        AddFCWPOOL(model, null);
                        AddFCDJOUR(tdcode, fcdjour.STEP_NO, action, "轉送", userid, fcdjour.DEPT_NO, fcdjour.SC_NO);
                        result= UpdateAPMFORM(tdcode, fcdjour.STEP_NO, userid, fcdjour.DEPT_NO, fcdjour.SC_NO);
                        if (result.RtnCode != 0)
                        {
                            return Ok(new { result = model, rtncode = -1, rtndesc = result.RtnDesc });
                        }
                    }
                    else if (action == "05") //誤收件
                    {
                        if (APMFORM.ADV_UID != userid)
                        {
                            return Ok(new { result = "", rtncode = -1, rtndesc = "案件擁有者並非案件處理人員，請確認後再進行誤收件。" });
                        }
                        var fcdjours = db.FCDJOUR.Where(f => f.TD_CODE == tdcode && (f.ACTION_COD == "01" || f.ACTION_COD == "04")).OrderByDescending(f => f.seqNo).ToList();
                        var fcdjour = fcdjours.FirstOrDefault();
                        model.TD_CODE = tdcode;
                        model.DEPT_NO = fcdjour.DEPT_NO;
                        model.STEP_NO = fcdjour.STEP_NO;
                        model.ACTION = fcdjour.ACTION_COD;
                        model.SC_NO = fcdjour.SC_NO;
                        AddFCWPOOL(model, Convert.ToDateTime(fcdjour.ACTION_DT));
                        AddFCDJOUR(tdcode, fcdjour.STEP_NO, action, "誤收件", userid, fcdjour.DEPT_NO, fcdjour.SC_NO);
                        result= UpdateAPMFORM(tdcode, fcdjour.STEP_NO, userid, fcdjour.DEPT_NO, fcdjour.SC_NO);
                        if (result.RtnCode != 0)
                        {
                            return Ok(new { result = model, rtncode = -1, rtndesc = result.RtnDesc });
                        }
                    }
                    transaction.Commit();
                    return Ok(new { result = model, rtncode = 0, rtndesc = "資料讀取成功" });
                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    return Ok(new { result = "", rtncode = -1, rtndesc = e.Message });
                }
            }
        }

        public IHttpActionResult Put(ViewModels.WorkFlowViewModel model)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    /*
                     * (各類計畫)會計科目經費額度檢查TDDYPAC 上傳時對應工程所屬計畫的PJDYPAC 
                       年度SYR與科目ACC_ID是否已經超過目前該計畫年度科目下各工程使用的總金額
                    */
                    var tmTDMFORM = db.TDMFORM.AsNoTracking().Where(f => f.TD_CODE == model.TD_CODE).FirstOrDefault();
                    var tmPJMFORM = db.PJMFORM.AsNoTracking().Where(f => f.PJ_CODE == tmTDMFORM.PJ_CODE).FirstOrDefault();
                    //上傳使用者合理性檢查
                    if (tmTDMFORM.ADV_UID!= model.USR_ID)
                    {
                        return Ok(new { result = "", rtncode = -1, rtndesc = "案件擁有者並非案件處理人員，請確認後再進行處理。" });
                    }
                    //local Data
                    if (model.TDDYPAC != null && model.TDDYPAC.Count()!= 0)
                    {
                        var SYRs = model.TDDYPAC.OrderBy(f=> f.SYR).Select(f=> f.SYR).Distinct();
                        if (tmPJMFORM.PJ_MODE == "F") //自由提案型計畫
                        {
                            foreach (var Entry in SYRs)
                            {
                                string SYR = Entry;

                                var TDDYPACs = model.TDDYPAC.Where(f => f.SYR == SYR).ToList();
                                foreach (var TDDYPAC in TDDYPACs)
                                {
                                    var PJFUND = db.PJDYPAC.AsNoTracking().Where(f => f.SYR == SYR && f.ACC_FUND != null && f.ACC_ID == TDDYPAC.ACC_ID).Select(f => f.ACC_FUND).Sum();
                                    var TD_USED_FUND = db.TDDYPAC.AsNoTracking().Where(f => f.SYR == SYR && f.ACC_FUND != null && f.ACC_ID == TDDYPAC.ACC_ID).Select(f => f.ACC_FUND).Sum();
                                    var TOTAL_USED_FUND = TDDYPAC.ACC_FUND + TD_USED_FUND;
                                    if (TOTAL_USED_FUND > PJFUND)
                                    {
                                        transaction.Rollback();
                                        return Ok(new { result = "", rtncode = -1, rtndesc = "工程編號: " + model.TD_CODE + " 之工程案件不存在，請查明後再試..." });
                                    }
                                }
                                /*
                                (自由提案型計畫) PJ_MODE='F')單位分配經費額度檢查TDDYPAC 上傳
                                時對應所屬單位的分配額是否已經超過目前該計畫年度單位分配金額PJDFDIS上限             
                                */
                                foreach (var TDDYPAC in TDDYPACs)
                                {
                                    //ADVDP的判斷要調整一下, 改從APMFORM取得
                                    var tmAPM = db.APMFORM.Where(f => f.AP_CODE ==  tmTDMFORM.AP_CODE).FirstOrDefault();
                                    string relADV_DP = tmAPM.ADV_DP.ToString();

                                    
                                    var DP_YR_DIS_FUND = db.PJDFDIS.Where(f => f.PJ_CODE == tmPJMFORM.PJ_CODE && f.SYR == SYR && f.SC_NO.Substring(0, 2) == relADV_DP.Substring(0, 2)).Select(f => f.DIST_FEE).Sum();
                                    //var DP_YR_DIS_FUND = db.PJDFDIS.Where(f => f.SYR == SYR && f.SC_NO.Substring(0, 2) == tmTDMFORM.ADV_DP.Substring(0, 2)).Select(f => f.DIST_FEE).Sum();
                                    var tmTDMFORMs = db.TDMFORM.AsNoTracking().Where(f => f.PJ_CODE == tmPJMFORM.PJ_CODE && f.ADV_DP.Substring(0, 2) == tmTDMFORM.ADV_DP.Substring(0, 2)).Select(f => f.TD_CODE).ToList();
                                    var TD_USED_DIS_FUND = db.TDDYPAC.AsNoTracking().Where(f => tmTDMFORMs.Contains(f.TD_CODE) && f.SYR == SYR && f.ACC_FUND != null && f.ACC_ID == TDDYPAC.ACC_ID).Select(f => f.ACC_FUND).Sum();
                                    var TOTAL_USED_FUND = TDDYPAC.ACC_FUND + TD_USED_DIS_FUND;
                                    if (TOTAL_USED_FUND > DP_YR_DIS_FUND)
                                    {
                                        transaction.Rollback();
                                        return Ok(new { result = "", rtncode = -1, rtndesc = "工程編號: " + model.TD_CODE + " 之經費超過年度部門分配金額上限, 不允許上傳!" });
                                    }

                                }
                            }
                        }
                        else
                        {
                            foreach (var Entry in SYRs)
                            {
                                string SYR = Entry;
                                var TDDYPACs = model.TDDYPAC.Where(f => f.SYR == SYR).ToList();
                                foreach (var TDDYPAC in TDDYPACs)
                                {
                                    var PJFUND = db.PJDYPAC.AsNoTracking().Where(f => f.SYR == SYR && f.PJ_CODE== tmPJMFORM.PJ_CODE && f.ACC_FUND != null && f.ACC_ID == TDDYPAC.ACC_ID).Select(f => f.ACC_FUND).Sum();
                                    var tmTDMFORMs = db.TDMFORM.AsNoTracking().Where(f => f.PJ_CODE == tmPJMFORM.PJ_CODE).Select(f=> f.TD_CODE).ToList();
                                    var TD_USED_FUND = db.TDDYPAC.AsNoTracking().Where(f => tmTDMFORMs.Contains(f.TD_CODE) && f.SYR == SYR && f.ACC_FUND != null && f.ACC_ID == TDDYPAC.ACC_ID).Select(f => f.ACC_FUND).Sum();
                                    var TOTAL_USED_FUND = TDDYPAC.ACC_FUND + TD_USED_FUND;
                                    if (TOTAL_USED_FUND > PJFUND)
                                    {
                                        transaction.Rollback();
                                        return Ok(new { result = "", rtncode = -1, rtndesc = "工程編號: " + model.TD_CODE + " 之經費超過年度科目金額上限, 不允許上傳!" });
                                    }
                                }
                            }
                        }
                    }

                    if (model.ACTION == "01" || model.ACTION == "02") //送件 或自辦
                    {
                        if (model.ACTION== "01")
                        {
                            //Put FCWPOOL
                            var FCWPOOL = db.TDV_FCWPOOL.Where(f => f.TD_CODE == model.TD_CODE).FirstOrDefault();
                            if (FCWPOOL != null) //有資料先刪除
                            {
                                var fcwpool = db.FCWPOOL.Where(f => f.TD_CODE == model.TD_CODE).FirstOrDefault();
                                if (fcwpool != null)
                                {
                                    db.FCWPOOL.Remove(fcwpool);
                                }
                                fcwpool = db.FCWPOOL.Where(f => f.AP_CODE == model.TD_CODE).FirstOrDefault();
                                if (fcwpool != null)
                                {
                                    db.FCWPOOL.Remove(fcwpool);
                                }
                            }
                            AddFCWPOOL(model, null);
                            //Put FCDJOUR
                            AddFCDJOUR(model.TD_CODE, model.ACTION, model.STEP_NO, "送件", model.USR_ID, model.DEPT_NO, model.SC_NO);
                        }
                        else if (model.ACTION== "02")
                        {
                            //Put FCDJOUR
                            AddFCDJOUR(model.TD_CODE, model.ACTION, model.STEP_NO, "自辦", model.USR_ID, model.DEPT_NO, model.SC_NO);
                            
                            if (model.STEP_NO== "09") //自辦議定
                            {
                                //找到施工單位
                                var wkTDMFORM= db.TDMFORM.AsNoTracking().Where(f => f.TD_CODE == model.TD_CODE).FirstOrDefault();
                                if (wkTDMFORM== null)
                                {
                                    transaction.Rollback();
                                    return Ok(new { result = "", rtncode = -1, rtndesc = "工程編號: " +model.TD_CODE+ " 之工程案件不存在，請查明後再試..." });
                                }
                                if (string.IsNullOrEmpty(wkTDMFORM.CNS_DP)== true)
                                {
                                    transaction.Rollback();
                                    return Ok(new { result = "", rtncode = -1, rtndesc = "工程編號: " + model.TD_CODE + " 無法取得施工部門資訊，請查明後再試..." });
                                }

                                //Put FCDJOUR
                                AddFCDJOUR(model.TD_CODE, model.ACTION, model.STEP_NO, "自辦", model.USR_ID, model.DEPT_NO, model.SC_NO);
                                //put FCWPOOL
                                var FCWPOOL = db.TDV_FCWPOOL.Where(f => f.TD_CODE == model.TD_CODE).FirstOrDefault();
                                if (FCWPOOL != null) //有資料先刪除
                                {
                                    var fcwpool = db.FCWPOOL.Where(f => f.TD_CODE == model.TD_CODE).FirstOrDefault();
                                    if (fcwpool != null)
                                    {
                                        db.FCWPOOL.Remove(fcwpool);
                                    }
                                    fcwpool = db.FCWPOOL.Where(f => f.AP_CODE == model.TD_CODE).FirstOrDefault();
                                    if (fcwpool != null)
                                    {
                                        db.FCWPOOL.Remove(fcwpool);
                                    }
                                }
                                //先記住原來的部門
                                string tmDEPT_NO = model.DEPT_NO;
                                string tmSTEP_NO = model.STEP_NO;
                                model.DEPT_NO = wkTDMFORM.CNS_DP;
                                model.STEP_NO = "20"; //議定完成
                                AddFCWPOOL(model, null);
                                //回寫原來的部門以免後續更新資料寫到施工部門
                                model.DEPT_NO = tmDEPT_NO;
                                model.STEP_NO = tmSTEP_NO;
                            }
                        }
                        //Put Files
                        var FCWFILE = db.FCWFILE.Where(f => f.TD_CODE == model.TD_CODE).FirstOrDefault();
                        if (FCWFILE != null)
                        {
                            FCWFILE.TD_CODE = model.TD_CODE;
                            FCWFILE.FILE_NAME = model.TD_CODE + ".zip";
                            FCWFILE.FILECONT = model.fileArray;
                            FCWFILE.UPD_DT = DateTime.Now;
                            FCWFILE.UPD_UID = model.USR_ID;
                        }
                        else
                        {
                            FCWFILE = new Models.FCWFILE();
                            FCWFILE.TD_CODE = model.TD_CODE;
                            FCWFILE.FILE_NAME = model.TD_CODE + ".zip";
                            FCWFILE.FILECONT = model.fileArray;
                            FCWFILE.ADD_DT = DateTime.Now;
                            FCWFILE.UPD_UID = model.USR_ID;
                            db.FCWFILE.Add(FCWFILE);
                        }
                        db.SaveChanges();

                        //處理資料基本資料更新
                        if (model.TDMFORM.Count() != 0)
                        {
                            var TDMFORM= db.Set<Models.TDMFORM>().AsNoTracking().Where(f => f.TD_CODE == model.TD_CODE).FirstOrDefault();
                            tmTDMFORM = model.TDMFORM.FirstOrDefault();
                            if (model.ACTION== "02") //自辦
                            {
                                switch (model.STEP_NO)
                                {
                                    case "02": //設計
                                        tmTDMFORM.DSG_UID = model.USR_ID;
                                        tmTDMFORM.DSG_DP = model.DEPT_NO;
                                        tmTDMFORM.DSG_SC = model.SC_NO;
                                        break;
                                    case "03": //發包
                                        tmTDMFORM.BID_UID = model.USR_ID;
                                        tmTDMFORM.BID_DP = model.DEPT_NO;
                                        tmTDMFORM.BID_SC = model.SC_NO;
                                        break;
                                    case "04": //施工
                                        tmTDMFORM.CNS_UID = model.USR_ID;
                                        tmTDMFORM.CNS_DP = model.DEPT_NO;
                                        tmTDMFORM.CNS_SC = model.SC_NO;
                                        break;
                                    case "05": //決算
                                        tmTDMFORM.AIT_UID = model.USR_ID;
                                        tmTDMFORM.AIT_DP = model.DEPT_NO;
                                        tmTDMFORM.AIT_SC = model.SC_NO;
                                        break;
                                }
                                if (model.STEP_NO!= "UU") //自辦更新時不更新流程資料
                                {
                                    tmTDMFORM.STEP_NO = model.STEP_NO;
                                    tmTDMFORM.ADV_UID = model.USR_ID;
                                    tmTDMFORM.ADV_DP = model.DEPT_NO;
                                }
                                else
                                {
                                    tmTDMFORM.ADV_UID = model.USR_ID;
                                    tmTDMFORM.ADV_DP = model.DEPT_NO;
                                }
                            }
                            if (TDMFORM != null)
                            {
                                tmTDMFORM.UPD_DT = DateTime.Now;
                                tmTDMFORM.UPD_UID = model.USR_ID;
                                db.Entry(tmTDMFORM).State= System.Data.Entity.EntityState.Modified;
                            }
                            else
                            {
                                tmTDMFORM.ADD_DT = DateTime.Now;
                                db.Entry(tmTDMFORM).State= System.Data.Entity.EntityState.Added;
                            }

                            //TDDYPSC
                            var tmTDDYPSC = db.TDDYPSC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDYPSC.RemoveRange(tmTDDYPSC);
                            if (model.TDDYPSC != null)
                            {
                                foreach (var entity in model.TDDYPSC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }
                            //TDDYPAC
                            var tmTDDYPAC = db.TDDYPAC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDYPAC.RemoveRange(tmTDDYPAC);
                            if (model.TDDYPAC != null)
                            {
                                foreach (var entity in model.TDDYPAC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDMSCD
                            var tmTDDMSCD = db.TDDMSCD.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDMSCD.RemoveRange(tmTDDMSCD);
                            if (model.TDDMSCD != null)
                            {
                                foreach (var entity in model.TDDMSCD)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDACPT
                            var tmTDDACPT = db.TDDACPT.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDACPT.RemoveRange(tmTDDACPT);
                            if (model.TDDACPT != null)
                            {
                                foreach (var entity in model.TDDACPT)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDARBK
                            var tmTDDARBK = db.TDDARBK.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDARBK.RemoveRange(tmTDDARBK);
                            if (model.TDDARBK != null)
                            {
                                foreach (var entity in model.TDDARBK)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDARRC
                            var tmTDDARRC = db.TDDARRC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDARRC.RemoveRange(tmTDDARRC);
                            if (model.TDDARRC != null)
                            {
                                foreach (var entity in model.TDDARRC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDARRE
                            var tmTDDARRE = db.TDDARRE.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDARRE.RemoveRange(tmTDDARRE);
                            if (model.TDDARRE != null)
                            {
                                foreach (var entity in model.TDDARRE)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDBDCM
                            var tmTDDBDCM = db.TDDBDCM.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDBDCM.RemoveRange(tmTDDBDCM);
                            if (model.TDDBDCM != null)
                            {
                                foreach (var entity in model.TDDBDCM)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDBDRC
                            var tmTDDBDRC = db.TDDBDRC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDBDRC.RemoveRange(tmTDDBDRC);
                            if (model.TDDBDRC != null)
                            {
                                foreach (var entity in model.TDDBDRC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDBGET
                            var tmTDDBGET = db.TDDBGET.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDBGET.RemoveRange(tmTDDBGET);
                            if (model.TDDBGET != null)
                            {
                                foreach (var entity in model.TDDBGET)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDBGLS
                            var tmTDDBGLS = db.TDDBGLS.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDBGLS.RemoveRange(tmTDDBGLS);
                            if (model.TDDBGLS != null)
                            {
                                foreach (var entity in model.TDDBGLS)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

//
                            //TDDBGRC
                            var tmTDDBGRC = db.TDDBGRC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDBGRC.RemoveRange(tmTDDBGRC);
                            if (model.TDDBGRC != null)
                            {
                                foreach (var entity in model.TDDBGRC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDCHKR
                            var tmTDDCHKR = db.TDDCHKR.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDCHKR.RemoveRange(tmTDDCHKR);
                            if (model.TDDCHKR != null)
                            {
                                foreach (var entity in model.TDDCHKR)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }


                            //TDDCPGP
                            var tmTDDCPGP = db.TDDCPGP.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDCPGP.RemoveRange(tmTDDCPGP);
                            if (model.TDDCPGP != null)
                            {
                                foreach (var entity in model.TDDCPGP)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDCQTY
                            var tmTDDCQTY = db.TDDCQTY.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDCQTY.RemoveRange(tmTDDCQTY);
                            if (model.TDDCQTY != null)
                            {
                                foreach (var entity in model.TDDCQTY)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDCMSL
                            var tmTDDCMSL = db.TDDCMSL.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDCMSL.RemoveRange(tmTDDCMSL);
                            if (model.TDDCMSL != null)
                            {
                                foreach (var entity in model.TDDCMSL)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDDRCR
                            var tmTDDDRCR = db.TDDDRCR.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDDRCR.RemoveRange(tmTDDDRCR);
                            if (model.TDDDRCR != null)
                            {
                                foreach (var entity in model.TDDDRCR)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }


                            //TDDECCR
                            var tmTDDECCR = db.TDDECCR.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDECCR.RemoveRange(tmTDDECCR);
                            if (model.TDDECCR != null)
                            {
                                foreach (var entity in model.TDDECCR)
                                {
                                    entity.TD_CODE = model.TD_CODE;
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDECRF
                            var tmTDDECRF = db.TDDECRF.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDECRF.RemoveRange(tmTDDECRF);
                            if (model.TDDECRF != null)
                            {
                                foreach (var entity in model.TDDECRF)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDECWT
                            var tmTDDECWT = db.TDDECWT.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDECWT.RemoveRange(tmTDDECWT);
                            if (model.TDDECWT != null)
                            {
                                foreach (var entity in model.TDDECWT)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDFDIS
                            var tmTDDFDIS = db.TDDFDIS.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDFDIS.RemoveRange(tmTDDFDIS);
                            if (model.TDDFDIS != null)
                            {
                                foreach (var entity in model.TDDFDIS)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDFDSC
                            var tmTDDFDSC = db.TDDFDSC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDFDSC.RemoveRange(tmTDDFDSC);
                            if (model.TDDFDSC != null)
                            {
                                foreach (var entity in model.TDDFDSC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDFUND
                            var tmTDDFUND = db.TDDFUND.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDFUND.RemoveRange(tmTDDFUND);
                            if (model.TDDFUND != null)
                            {
                                foreach (var entity in model.TDDFUND)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDGCRC
                            var tmTDDGCRC = db.TDDGCRC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDGCRC.RemoveRange(tmTDDGCRC);
                            if (model.TDDGCRC != null)
                            {
                                foreach (var entity in model.TDDGCRC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDMATL
                            var tmTDDMATL = db.TDDMATL.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDMATL.RemoveRange(tmTDDMATL);
                            if (model.TDDMATL != null)
                            {
                                foreach (var entity in model.TDDMATL)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDREVF
                            var tmTDDREVF = db.TDDREVF.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDREVF.RemoveRange(tmTDDREVF);
                            if (model.TDDREVF != null)
                            {
                                foreach (var entity in model.TDDREVF)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }


                            //TDDSTWR
                            var tmTDDSTWR = db.TDDSTWR.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDSTWR.RemoveRange(tmTDDSTWR);
                            if (model.TDDSTWR != null)
                            {
                                foreach (var entity in model.TDDSTWR)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDSUPM
                            var tmTDDSUPM = db.TDDSUPM.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDSUPM.RemoveRange(tmTDDSUPM);
                            if (model.TDDSUPM != null)
                            {
                                foreach (var entity in model.TDDSUPM)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }


                            //TDDUCLF
                            var tmTDDUCLF = db.TDDUCLF.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDUCLF.RemoveRange(tmTDDUCLF);
                            if (model.TDDUCLF != null)
                            {
                                foreach (var entity in model.TDDUCLF)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDWBSC
                            var tmTDDWBSC = db.TDDWBSC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDWBSC.RemoveRange(tmTDDWBSC);
                            if (model.TDDWBSC != null)
                            {
                                foreach (var entity in model.TDDWBSC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDWKEN
                            var tmTDDWKEN = db.TDDWKEN.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDWKEN.RemoveRange(tmTDDWKEN);
                            if (model.TDDWKEN != null)
                            {
                                foreach (var entity in model.TDDWKEN)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDWKLS
                            var tmTDDWKLS = db.TDDWKLS.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDWKLS.RemoveRange(tmTDDWKLS);
                            if (model.TDDWKLS != null)
                            {
                                foreach (var entity in model.TDDWKLS)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDDYPFD
                            var tmTDDYPFD = db.TDDYPFD.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDYPFD.RemoveRange(tmTDDYPFD);
                            if (model.TDDYPFD != null)
                            {
                                foreach (var entity in model.TDDYPFD)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDHBGLS
                            var tmTDHBGLS = db.TDHBGLS.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDHBGLS.RemoveRange(tmTDHBGLS);
                            if (model.TDHBGLS != null)
                            {
                                foreach (var entity in model.TDHBGLS)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDHBGET
                            var tmTDHBGET = db.TDHBGET.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDHBGET.RemoveRange(tmTDHBGET);
                            if (model.TDHBGET != null)
                            {
                                foreach (var entity in model.TDHBGET)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDHYPAC
                            var tmTDHYPAC = db.TDHYPAC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDHYPAC.RemoveRange(tmTDHYPAC);
                            if (model.TDHYPAC != null)
                            {
                                foreach (var entity in model.TDHYPAC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDHYPFD
                            var tmTDHYPFD = db.TDHYPFD.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDHYPFD.RemoveRange(tmTDHYPFD);
                            if (model.TDHYPFD != null)
                            {
                                foreach (var entity in model.TDHYPFD)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }

                            //TDHYPSC
                            var tmTDHYPSC = db.TDHYPSC.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDHYPSC.RemoveRange(tmTDHYPSC);
                            if (model.TDHYPSC != null)
                            {
                                foreach (var entity in model.TDHYPSC)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }


                            //TDDNPAN 由NPCals 呼叫web service上傳
                            //var tmTDDNPAN = db.TDDNPAN.Where(f => f.TD_CODE == model.TD_CODE);
                            //db.TDDNPAN.RemoveRange(tmTDDNPAN);
                            //if (model.TDDNPAN != null)
                            //{
                            //    foreach (var entity in model.TDDNPAN)
                            //    {
                            //        db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                            //    }
                            //}


                            //TDDRMPR  由NPCals 呼叫web service上傳
                            //var tmTDDRMPR = db.TDDRMPR.Where(f => f.TD_CODE == model.TD_CODE);
                            //db.TDDRMPR.RemoveRange(tmTDDRMPR);
                            //if (model.TDDRMPR != null)
                            //{
                            //    foreach (var entity in model.TDDRMPR)
                            //    {
                            //        db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                            //    }
                            //}

                            //TDDPRMT 暫不更新
                            /*
                            var tmTDDPRMT = db.TDDPRMT.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDDPRMT.RemoveRange(tmTDDPRMT);
                            if (model.TDDPRMT != null)
                            {
                                foreach (var entity in model.TDDPRMT)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }
                            */


                            //TDMIOMT 暫不更新
                            /*
                            var tmTDMIOMT = db.TDMIOMT.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDMIOMT.RemoveRange(tmTDMIOMT);
                            if (model.TDMIOMT != null)
                            {
                                foreach (var entity in model.TDMIOMT)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }
                            */
                            //TDMPRMT 暫不更新
                            /*
                            var tmTDMPRMT = db.TDMPRMT.Where(f => f.TD_CODE == model.TD_CODE);
                            db.TDMPRMT.RemoveRange(tmTDMPRMT);
                            if (model.TDMPRMT != null)
                            {
                                foreach (var entity in model.TDMPRMT)
                                {
                                    db.Entry(entity).State = System.Data.Entity.EntityState.Added;
                                }
                            }
                            */
                            db.SaveChanges();
                        }

                        transaction.Commit();
                    }
                    return Ok(new { result = model, rtncode = 0, rtndesc = "資料處理完成" });
                }
                catch (Exception e)
                {

                    //System.Diagnostics.Debug.Write(e.StackTrace);
                    transaction.Rollback();
                    if (e is DbEntityValidationException)
                    {
                        var entityErrorMessage = ((DbEntityValidationException)e).EntityValidationErrors
                            .Select(x => x.Entry)
                            .Select(x => x.Entity).FirstOrDefault();


                        var errorMessages = ((DbEntityValidationException)e).EntityValidationErrors
                                .SelectMany(x => x.ValidationErrors)
                                .Select(x => x.ErrorMessage).FirstOrDefault();

                        // Join the list to a single string.
                        var fullErrorMessage = string.Join("; ", errorMessages);
                        return Ok(new { result = "", rtncode = -1, rtndesc = entityErrorMessage + " : " + fullErrorMessage });
                    }
                    else
                    {
                        return Ok(new { result = "", rtncode = -1, rtndesc = e.Message });
                    }
                    // Retrieve the error messages as a list of strings.
                }
            }
        }
    }
}
