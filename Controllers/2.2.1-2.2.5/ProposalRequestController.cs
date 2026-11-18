using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Web;
using System.Web.Mvc;
using WCals.Models;
using WCals.Models.CMBASEModels;
using WCals.Models.CMETMSModels;
using WCals.Models.CMPSGSModels;

namespace WCals.Controllers._2._2._1_2._2._5
{
    public class ProposalRequestController : Controller
    {
        CMBASEEntities db = new CMBASEEntities();
        CMETMSEntities db3 = new CMETMSEntities();
        CMPSGSEntities db4 = new CMPSGSEntities();

        // GET: ProposalRequest
        public ActionResult Index(int page = 1)
        {
            int incurrentPage = page < 1 ? 1 : page;

            var PPID_List = db4.PJDSUPD.Where(m => m.ADV_DP == User.Identity.Name).ToList();
            var PJDSYRC = db4.PJDSYRC.Where(m=>m.OPEN_FG == true).ToList();
            var PJMFORM = (from a in db4.PJMFORM.ToList()
                           join b in PPID_List on a.PPID equals b.PPID
                           join c in PJDSYRC on a.PPID equals c.PPID
                           select new ProposalRequestViewModel
                           {
                               PPID = a.PPID,
                               YRC_ID = c.YRC_ID,
                               SYR = c.BGN_SYR + "~" + c.END_SYR,
                               PJ_NAME = a.PJ_NAME
                           }).ToList();
            
            ViewBag.PJMFORM = PJMFORM;

            var SMCCODE1 = db3.SMCCODE.Where(m => m.SYS_ID == "SM" && m.CATY_ID == "02").ToList(); ///辦理層級
            ViewBag.SMCCODE1 = SMCCODE1;

            var SMCCODE2 = db4.SMCCODE.Where(m => m.SYS_ID == "TT" && m.CATY_ID == "03").ToList(); ///申請案件類別
            ViewBag.SMCCODE2 = SMCCODE2;

            var cities = new[] { "不分區", "新北市", "基隆市", "台北市", "桃園市", "新竹縣", "新竹市", "苗栗縣", "台中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市", "台南市", "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣", "澎湖縣", "金門縣", "連江縣" };
            var PJNames = (from a in db3.PJCFORM
                           where cities.Contains(a.SER_CITY) && a.PJ_ID_TYPE == "C"
                           select a.PJ_NAME).ToList();

            ViewBag.PJName = PJNames;

            var List = new[] {
                            "不分區", "新北市", "基隆市", "台北市", "桃園市", "新竹縣", "新竹市",
                            "苗栗縣", "台中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市",
                            "台南市", "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣", "澎湖縣",
                            "金門縣", "連江縣", ""
                                             };
            var TDMFORM = (from a in db3.TDMFORM
                                  where List.Contains(a.TD_CITY)
                                  select new
                                  {
                                      a.TD_CODE,
                                      a.TD_NAME,
                                      a.TD_CITY,
                                      a.TD_TOWN
                                  }).ToList();

            var WindowsTD_NAME = (from a in TDMFORM
                                  select new Models.CMETMSModels.TDMFORM
                                  {
                                      TD_CODE = a.TD_CODE,
                                      TD_NAME = a.TD_NAME,
                                      TD_CITY = a.TD_CITY,
                                      TD_TOWN = a.TD_TOWN
                                  }).OrderBy(m=>m.TD_CODE).ToList();

            ViewBag.Count = WindowsTD_NAME.Count/10;

            WindowsTD_NAME = WindowsTD_NAME.Take(10).ToList();

            ViewBag.WindowsTDName = WindowsTD_NAME;

            var City = db.SMCPOSTN.DistinctBy(m=>m.CITY).ToList();
            ViewBag.City = City;
            ViewBag.User = db.OGCUSER.Where(m=>m.USR_ID == User.Identity.Name).FirstOrDefault().USR_NAME;
            ViewBag.Page = page;

            return View();
        }

        public ActionResult NextPage(int Page , string Search_thing)
        {
            int First_Page = 10 * Page;
            Page = Page + 1;

            var List = new[] {
                            "不分區", "新北市", "基隆市", "台北市", "桃園市", "新竹縣", "新竹市",
                            "苗栗縣", "台中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市",
                            "台南市", "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣", "澎湖縣",
                            "金門縣", "連江縣", ""
                                             };
            var TDMFORM = (from a in db3.TDMFORM
                           where List.Contains(a.TD_CITY)
                           select new
                           {
                               a.TD_CODE,
                               a.TD_NAME,
                               a.TD_CITY,
                               a.TD_TOWN
                           }).ToList();

            if(Search_thing != null) {

                TDMFORM = TDMFORM.Where(m => m.TD_NAME.Contains(Search_thing)).ToList();

            }

            var WindowsTD_NAME = (from a in TDMFORM
                                  select new Models.CMETMSModels.TDMFORM
                                  {
                                      TD_CODE = a.TD_CODE,
                                      TD_NAME = a.TD_NAME,
                                      TD_CITY = a.TD_CITY,
                                      TD_TOWN = a.TD_TOWN
                                  }).Skip(First_Page).Take(10).ToList();

            return Json(WindowsTD_NAME);
        }

        public ActionResult LastPage(int Page, string Search_thing)
        {
            Page = Page - 1;
            int First_Page = 10 * (Page - 1);

            var List = new[] {
                            "不分區", "新北市", "基隆市", "台北市", "桃園市", "新竹縣", "新竹市",
                            "苗栗縣", "台中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市",
                            "台南市", "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣", "澎湖縣",
                            "金門縣", "連江縣", ""
                                             };
            var TDMFORM = (from a in db3.TDMFORM
                           where List.Contains(a.TD_CITY)
                           select new
                           {
                               a.TD_CODE,
                               a.TD_NAME,
                               a.TD_CITY,
                               a.TD_TOWN
                           }).ToList();

            if (Search_thing != null)
            {

                TDMFORM = TDMFORM.Where(m => m.TD_NAME.Contains(Search_thing)).ToList();

            }

            var WindowsTD_NAME = (from a in TDMFORM
                                  select new Models.CMETMSModels.TDMFORM
                                  {
                                      TD_CODE = a.TD_CODE,
                                      TD_NAME = a.TD_NAME,
                                      TD_CITY = a.TD_CITY,
                                      TD_TOWN = a.TD_TOWN
                                  }).Skip(First_Page).Take(10).ToList();

            return Json(WindowsTD_NAME);
        }

        public ActionResult Search(string Search_thing)
        {
            var List = new[] {
                            "不分區", "新北市", "基隆市", "台北市", "桃園市", "新竹縣", "新竹市",
                            "苗栗縣", "台中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市",
                            "台南市", "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣", "澎湖縣",
                            "金門縣", "連江縣", ""
                                             };
            var TDMFORM = (from a in db3.TDMFORM
                           where List.Contains(a.TD_CITY)
                           select new
                           {
                               a.TD_CODE,
                               a.TD_NAME,
                               a.TD_CITY,
                               a.TD_TOWN
                           }).ToList();

            TDMFORM = TDMFORM.Where(m=>m.TD_NAME.Contains(Search_thing)).ToList();

            var WindowsTD_NAME = (from a in TDMFORM
                                  select new Models.CMETMSModels.TDMFORM
                                  {
                                      TD_CODE = a.TD_CODE,
                                      TD_NAME = a.TD_NAME,
                                      TD_CITY = a.TD_CITY,
                                      TD_TOWN = a.TD_TOWN
                                  }).ToList();


            return Json(WindowsTD_NAME);
        }

        public ActionResult GetTown(string City)
        {
            var result = (from mp in db.MACPOST
                         where mp.CITY == City
                         group mp by new { TownName = mp.TOWN == "" ? "不分區" : mp.TOWN, Town = mp.TOWN } into g
                         orderby g.Key.Town
                         select new
                         {
                             TOWNName = g.Key.TownName,
                             TOWN = g.Key.Town
                         }).ToList();
            return Json(result);
        }
        [HttpPost]
        public ActionResult Save( PJDSYRC Source1, APMFORM Source2, APDFILE Source3, IEnumerable<HttpPostedFileBase> FILE_PATH, TownList TownBlocks, int[] count, ItemList Items, FundList Funds)
        { 
            APMFORM model = new APMFORM();
            model = Source2;
            var time = DateTime.Now;
            model.ADD_DT = time;
            model.YRC_ID = Source1.YRC_ID;
            var PJDSYRC = db4.PJDSYRC.Where(m=>m.YRC_ID == Source1.YRC_ID).FirstOrDefault();
            model.PPID = PJDSYRC.PPID;
            model.SYR = PJDSYRC.SYR;
            model.ADV_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            model.ADV_DP = Helpers.WebSiteHelper.getUserDept(); //(string)Session["DeptId"];
            db4.APMFORM.Add(model);
            db4.SaveChanges();
            var AID = model.AID;

            // 將 IEnumerable<HttpPostedFileBase> 轉換為 List<HttpPostedFileBase>
            var fileList = FILE_PATH.ToList();

            try { 
                for (int i = 0; i < fileList.Count; i++)
                {
                    APDFILE model2 = new APDFILE();
                    model2.FID = i + 1;
                    model2.FILE_PATH = fileList[i]?.FileName;
                    byte[] fileBytes;
                    using (var memoryStream = new MemoryStream())
                    {
                        fileList[i].InputStream.CopyTo(memoryStream);
                        fileBytes = memoryStream.ToArray();
                    }
                    model2.FILE_BODY = fileBytes;
                    model2.AID = AID;
                    model2.ADD_DT = time;
                    db4.APDFILE.Add(model2);
                    db4.SaveChanges();
                }
            }catch (Exception ) {  }

            for (int i = 0; i < TownBlocks.TOWN.Length; i++)
            {
                APDLOCP model3 = new APDLOCP();
                model3.LID = i + 1;
                model3.TOWN = TownBlocks.TOWN[i];
                model3.GIS_X = TownBlocks.GIS_X[i];
                model3.GIS_Y = TownBlocks.GIS_Y[i];
                model3.GIS_Z = TownBlocks.GIS_Z[i];
                model3.POS_SYB = TownBlocks.POS_SYB[i];
                model3.AID = AID;
                model3.ADD_DT = time;
                db4.APDLOCP.Add(model3);
                db4.SaveChanges();
            }

            for(int j = 0, M = 0, N = 0; j < TownBlocks.TOWN.Length; j++)
            {
                for(int k = 0; k < Items.MT_NAME.Length; k++)
                {
                    if (count[j] ==  k) 
                    {
                        M = count[j];
                        break;
                    }
                    if (M != 0)
                    {
                        N = M;
                    }
                   
                    APDWQTY model4 = new APDWQTY();
                    model4.LID = j + 1;
                    model4.MID = k + 1;
                    model4.AID = AID;
                    model4.ADD_DT = time;
                    model4.MT_NAME = Items.MT_NAME[N];
                    model4.MT_UNIT = Items.MT_UNIT[N];
                    model4.MT_QTY = Items.MT_QTY[N];
                    model4.MT_DESC = Items.MT_DESC[N];
                    model4.MT_SPEC = Items.MT_SPEC[N];
                    db4.APDWQTY.Add(model4) ;
                    db4.SaveChanges();
                    N++;
                }
            }

            for(int i = 0; i < Funds.SYR.Length; i++)
            {
                APDFUND model5 = new APDFUND();
                model5.AID = AID;
                model5.ADD_DT = time;
                model5.SYR = Funds.SYR[i];
                model5.FIT_CODE = Funds.FIT_CODE[i];
                model5.FUND_DESC = Funds.FUND_DESC[i];
                model5.FUND = Funds.FUND[i];
                db4.APDFUND.Add(model5); 
                db4.SaveChanges();
            }

            return RedirectToAction("Index");
        }
    }
}