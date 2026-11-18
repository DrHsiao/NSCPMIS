using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;
using Syncfusion.XlsIO;
using WCals.Helpers;
using WCals.Models;
using WCals.ViewModels;

namespace WCals.Controllers
{
    [Authorize]
    public class WTDM0405Controller : Controller
    {
        private readonly WCals.Models.WCalsEntities db = new Models.WCalsEntities();

        private int pageSize = 15;
        // key constants for storing authorities
        public const string AUTHORITY_TYPE_STRING_KEY = "authority";
        public const string SEARCH_PARAMETERS = "SP";
        public const string PAGE_QUERY_STRING_KEY = "page";

        // 收到request時檢核權限
        protected override void OnActionExecuting(System.Web.Mvc.ActionExecutingContext filterContext)
        {
            if (null == filterContext.Controller.TempData[AUTHORITY_TYPE_STRING_KEY])
            {
                filterContext.Controller.TempData.Add(AUTHORITY_TYPE_STRING_KEY, isAuthorizedToEdit());
            }
            else
            {
                filterContext.Controller.TempData.Keep(AUTHORITY_TYPE_STRING_KEY);
            }

            base.OnActionExecuting(filterContext);
        }

        private IQueryable<Models.MTDPINF> queryData(MTDPINFSearchViewModel SearchParameter)
        {
            //var query = db.MTDPINF.AsQueryable();
            string USER_DP = Helpers.WebSiteHelper.getUserDept().Substring(1, 2);
            var query = db.MTDPINF.Where(c => c.ADV_DP == USER_DP).AsQueryable();

            if (null != SearchParameter)
            {
                if (!string.IsNullOrWhiteSpace(SearchParameter.SYR))
                {
                    SearchParameter.SYR = SearchParameter.SYR.Trim();
                    query = query.Where(
                        f => f.SYR.Contains(SearchParameter.SYR));
                }
                if (!string.IsNullOrWhiteSpace(SearchParameter.SMN))
                {
                    query = query.Where(
                        f => f.SMN == SearchParameter.SMN);
                }
            }

            return query = query.OrderBy(f => f.SYR).ThenBy(f=>f.SMN);
        }

        // GET: WTDM0405
        [LogAction]
        public ActionResult Index(int? page)
        {
            var pageNumeber = page ?? 1;

            var searchParameters = (MTDPINFSearchViewModel)TempData.Peek(SEARCH_PARAMETERS);
            var query = queryData(searchParameters);

            var result = new MTDPINFListViewModel()
            {
                SearchParameter = searchParameters ?? new MTDPINFSearchViewModel(),
                MTDPINF = query.ToPagedList(pageNumeber, pageSize)
            };

            ViewBag.MONTHES = Helpers.TenderHelper.getDefaultMonthes();
            TempData[PAGE_QUERY_STRING_KEY] = page;

            return View(result);
        }


        [HttpPost]
        public ActionResult Index(MTDPINFListViewModel Model)
        {
            TempData[SEARCH_PARAMETERS] = Model.SearchParameter;

            var query = queryData(Model.SearchParameter);

            var result = new MTDPINFListViewModel
            {
                SearchParameter = Model.SearchParameter,
                MTDPINF = query.ToPagedList(1, pageSize),
            };

            ViewBag.itemAreas = Helpers.TenderHelper.getItemAreasOptions();
            ViewBag.MONTHES = Helpers.TenderHelper.getDefaultMonthes();

            return View(result);
        }

        private bool isAuthorizedToEdit()
        {
            if (WebSiteHelper.isSystemAdministrator())
            {
                return true;
            }
            WebSiteHelper.deptType currentDetpType = WebSiteHelper.getUserDepartmentType();
            if (currentDetpType == WebSiteHelper.deptType.MainDept || currentDetpType == WebSiteHelper.deptType.AreaDept)
            {
                return true;
            }

            return false;
        }


        public ActionResult New()
        {
            var model = new MTDPINFEditViewModel()
            {
                MTDPINF = new Models.MTDPINF()
            };

            ViewBag.MONTHES = Helpers.TenderHelper.getDefaultMonthes();

            return View(model);
        }

        [HttpPost]
        public ActionResult New(MTDPINFEditViewModel model)
        {
            if (ModelState.IsValid)
            {
                var mtdpinf = new Models.MTDPINF();

                // 檢核欄位回壓
                mtdpinf.SYR = model.VALIDATE_SYR;
                mtdpinf.SMN = model.VALIDATE_SMN;
                mtdpinf.ADV_DP = Helpers.WebSiteHelper.getUserDept().Substring(1, 2);
                mtdpinf.DM01_UP = model.VALIDATE_DM01_UP;
                mtdpinf.DM02_UP = model.VALIDATE_DM02_UP;
                mtdpinf.DM03_UP = model.VALIDATE_DM03_UP;
                mtdpinf.DM04_UP = model.VALIDATE_DM04_UP;
                mtdpinf.DM05_UP = model.VALIDATE_DM05_UP;
                mtdpinf.DM06_UP = model.VALIDATE_DM06_UP;
                mtdpinf.DM_NOTE = model.MTDPINF.DM_NOTE;

                mtdpinf.RM01_UP = model.VALIDATE_RM01_UP;
                mtdpinf.RM02_UP = model.VALIDATE_RM02_UP;
                mtdpinf.RM_NOTE = model.MTDPINF.RM_NOTE;

                mtdpinf.MM01_UP = model.VALIDATE_MM01_UP;
                mtdpinf.MM02_UP = model.VALIDATE_MM02_UP;
                mtdpinf.MM03_UP = model.VALIDATE_MM03_UP;
                mtdpinf.MM04_UP = model.VALIDATE_MM04_UP;
                mtdpinf.MM05_UP = model.VALIDATE_MM05_UP;
                mtdpinf.MM06_UP = model.VALIDATE_MM06_UP;
                mtdpinf.MM07_UP = model.VALIDATE_MM07_UP;
                mtdpinf.MM_NOTE = model.MTDPINF.MM_NOTE;

                mtdpinf.ADD_DT = DateTime.Now;
                mtdpinf.UPD_UID = User.Identity.Name;
                db.MTDPINF.Add(mtdpinf);
                db.SaveChanges();

                return RedirectToAction("Index", "WTDM0405");
            }

            ViewBag.MONTHES = Helpers.TenderHelper.getDefaultMonthes();

            return View(model);
        }

        public ActionResult Edit(String id)
        {
            int did = Convert.ToInt16(id);

            var mtdpinf = (from f in db.MTDPINF
                           where f.DID == did
                           select f).FirstOrDefault();

            var model = new MTDPINFEditViewModel()
            {
                MTDPINF = mtdpinf,
                VALIDATE_SYR = mtdpinf.SYR,
                VALIDATE_SMN = mtdpinf.SMN,
                VALIDATE_DM01_UP = mtdpinf.DM01_UP,
                VALIDATE_DM02_UP = mtdpinf.DM02_UP,
                VALIDATE_DM03_UP = mtdpinf.DM03_UP,
                VALIDATE_DM04_UP = mtdpinf.DM04_UP,
                VALIDATE_DM05_UP = mtdpinf.DM05_UP,
                VALIDATE_DM06_UP = mtdpinf.DM06_UP,
                VALIDATE_RM01_UP = mtdpinf.RM01_UP,
                VALIDATE_RM02_UP = mtdpinf.RM02_UP,
                VALIDATE_MM01_UP = mtdpinf.MM01_UP,
                VALIDATE_MM02_UP = mtdpinf.MM02_UP,
                VALIDATE_MM03_UP = mtdpinf.MM03_UP,
                VALIDATE_MM04_UP = mtdpinf.MM04_UP,
                VALIDATE_MM05_UP = mtdpinf.MM05_UP,
                VALIDATE_MM06_UP = mtdpinf.MM06_UP,
                VALIDATE_MM07_UP = mtdpinf.MM07_UP
            };

            ViewBag.MONTHES = Helpers.TenderHelper.getDefaultMonthes();

            return View(model);
        }

        [HttpPost]
        public ActionResult Edit(MTDPINFEditViewModel model)
        {
            if (ModelState.IsValid)
            {
                // 檢核欄位回壓
                var mtdpinf = (from f in db.MTDPINF
                               where f.DID == model.MTDPINF.DID
                               select f).FirstOrDefault();

                mtdpinf.DM01_UP = model.VALIDATE_DM01_UP;
                mtdpinf.DM02_UP = model.VALIDATE_DM02_UP;
                mtdpinf.DM03_UP = model.VALIDATE_DM03_UP;
                mtdpinf.DM04_UP = model.VALIDATE_DM04_UP;
                mtdpinf.DM05_UP = model.VALIDATE_DM05_UP;
                mtdpinf.DM06_UP = model.VALIDATE_DM06_UP;
                mtdpinf.DM_NOTE = model.MTDPINF.DM_NOTE;

                mtdpinf.RM01_UP = model.VALIDATE_RM01_UP;
                mtdpinf.RM02_UP = model.VALIDATE_RM02_UP;
                mtdpinf.RM_NOTE = model.MTDPINF.RM_NOTE;

                mtdpinf.MM01_UP = model.VALIDATE_MM01_UP;
                mtdpinf.MM02_UP = model.VALIDATE_MM02_UP;
                mtdpinf.MM03_UP = model.VALIDATE_MM03_UP;
                mtdpinf.MM04_UP = model.VALIDATE_MM04_UP;
                mtdpinf.MM05_UP = model.VALIDATE_MM05_UP;
                mtdpinf.MM06_UP = model.VALIDATE_MM06_UP;
                mtdpinf.MM07_UP = model.VALIDATE_MM07_UP;
                mtdpinf.MM_NOTE = model.MTDPINF.MM_NOTE;

                mtdpinf.UPD_DT = DateTime.Now;
                mtdpinf.UPD_UID = User.Identity.Name;

                db.SaveChanges();

                return RedirectToAction("Index", "WTDM0405");
            }

            ViewBag.MONTHES = Helpers.TenderHelper.getDefaultMonthes();

            return View(model);
        }

        public ActionResult Delete(String id)
        {
            int did = Convert.ToInt16(id);
            var mtdpinf = (from f in db.MTDPINF
                           where f.DID == did
                           select f).FirstOrDefault();
            db.MTDPINF.Remove(mtdpinf);
            db.SaveChanges();

            //reset查詢
            TempData.Remove(SEARCH_PARAMETERS);

            return RedirectToAction("Index", "WTDM0405");

        }

        public ActionResult Reset()
        {
            TempData.Remove(SEARCH_PARAMETERS);

            return RedirectToAction("Index", "WTDM0405");
        }

        //[HttpPost]
        //public ActionResult ExportData(SMCCPIXListViewModel Model)
        //{
        //    var query = queryData(Model.SearchParameter);

        //    using (ExcelEngine excelEngine = new Syncfusion.XlsIO.ExcelEngine())
        //    {
        //        string title = "營建物價資料明細表";
        //        List<SMCCPIX> items = query.ToList();
        //        IApplication application = excelEngine.Excel;
        //        application.DefaultVersion = ExcelVersion.Excel2007;
        //        String fileName = Server.MapPath("/Templates/營建物價資料明細表.xlsx");
        //        IWorkbook sourceWorkbook = application.Workbooks.Open(fileName);

        //        string[] dataTableColumns = new string[] { "itemCode", "Description", "SYR", "SMN", "ITEMAREA", "Price", "Unit", "PriceCat", "StandardDiff", "sampleCount" };

        //        DataTable reports = new DataTable();

        //        foreach (string column in dataTableColumns)
        //        {
        //            reports.Columns.Add(column);
        //        }

        //        foreach (SMCCPIX item in items)
        //        {
        //            reports.Rows.Add(item.itemCode, item.Description, item.SYR, item.SMN, item.ITEMAREA, item.Price, item.Unit, item.PriceCat, item.StandardDiff, item.sampleCount);
        //        }
        //        sourceWorkbook.Worksheets[0].ImportDataTable(reports, false, 3, 1, true);

        //        if (items.Count > 0)
        //        {
        //            setDefaultStyleForExportExcel(sourceWorkbook, 0, items.Count, "J");
        //        }

        //        sourceWorkbook.SaveAs(HttpUtility.UrlEncode(title) + ".xlsx", HttpContext.ApplicationInstance.Response, ExcelDownloadType.Open);
        //    }

        //    return View();
        //}

        //private void setDefaultStyleForExportExcel(IWorkbook workbook, int sheetsIndex, int itemCnt, string columnIdx)
        //{
        //    string finalRowAsString = (itemCnt + 2).ToString();
        //    string dataRange = "A2:" + columnIdx + finalRowAsString;
        //    workbook.Worksheets[sheetsIndex].Range[dataRange].WrapText = true;
        //    workbook.Worksheets[sheetsIndex].Range[dataRange].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
        //    workbook.Worksheets[sheetsIndex].Range[dataRange].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
        //    workbook.Worksheets[sheetsIndex].Range[dataRange].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
        //    workbook.Worksheets[sheetsIndex].Range[dataRange].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
        //}

    }
}