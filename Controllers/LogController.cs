using PagedList;
using System;
using System.Data.Entity;
using System.Linq;
using System.Web.Mvc;
using WCals.Helpers;
using WCals.Models.CMBASEModels;
using WCals.ViewModels;

namespace WCals.Controllers
{
    public class LogController : Controller
    {
        private readonly CMBASEEntities db = new CMBASEEntities();
        private readonly int pageSize = 10;

        #region 系統操作紀錄-查詢頁面動作

        /// <summary>
        /// 系統操作紀錄-資料查詢
        /// </summary>
        /// <param name="SearchParameter"> 查詢條件 </param>
        /// <remarks> 取得操作類型為M[點擊功能頁面]的資料 </remarks>
        /// <returns></returns>
        private IQueryable<LogListItemViewModel> GetQueryData(LogListSearchViewModel SearchParameter)
        {
            // 取得使用者帳號清單→供模糊查詢用
            ViewBag.User = db.OGCUSER.AsNoTracking()
                .Select(s => new UsrComponentSearchViewModel
                {
                    USR_ID = s.USR_ID,
                    USR_NAME = s.USR_NAME
                });

            // 取得系統操作紀錄資料
            var loginLogs = db.SMCLOGS.AsNoTracking().Where(l => l.ACTION_TYPE == "L");

            var query = from m in db.SMCLOGS.AsNoTracking()
                        where m.ACTION_TYPE == "M"
                        join u in db.OGCUSER on m.USR_ID equals u.USR_ID
                        // LEFT JOIN 最近一次登入 log
                        from lastLogin in loginLogs
                            .Where(l => l.USR_ID == m.USR_ID && l.LOGIN_DT < m.LOGIN_DT)
                            .OrderByDescending(l => l.LOGIN_DT)
                            .Take(1)
                            .DefaultIfEmpty()
                        select new LogListItemViewModel
                        {
                            USR_ID = m.USR_ID,
                            USR_NAME = u.USR_NAME,
                            ACTION_DESC = m.ACTION_DESC,
                            ACTION_DT = m.LOGIN_DT,
                            LASTEST_LOGIN_DT = (DateTime?)lastLogin.LOGIN_DT
                        };

            #region 查詢條件

            if (null != SearchParameter)
            {
                // 使用者帳號
                if (!string.IsNullOrWhiteSpace(SearchParameter.USR_ID))
                {
                    query = query.Where(f => f.USR_ID.Contains(SearchParameter.USR_ID.Trim()));
                }
                // 使用者名稱
                if (!string.IsNullOrEmpty(SearchParameter.USR_NAME))
                {
                    query = query.Where(f => f.USR_NAME.Contains(SearchParameter.USR_NAME.Trim()));
                }
                // 起始日期
                DateTime parsedDATE_START;
                if (DateTime.TryParse(SearchParameter.ACTION_DT_START, out parsedDATE_START))
                {
                    query = query.Where(m => DbFunctions.TruncateTime(m.ACTION_DT) >= parsedDATE_START.Date);
                }
                // 迄止日期
                DateTime parsedDATE_END;
                if (DateTime.TryParse(SearchParameter.ACTION_DT_END, out parsedDATE_END))
                {
                    query = query.Where(m => DbFunctions.TruncateTime(m.ACTION_DT) <= parsedDATE_END.Date);
                }
            }

            #endregion

            return query.OrderByDescending(f => f.ACTION_DT);
        }

        /// <summary>
        /// 系統操作紀錄查詢-初始/切換頁面
        /// </summary>
        /// <param name="page"> 分頁 </param>
        /// <param name="edit"> 切換頁面時為1 </param>
        /// <returns></returns>
        //[LogAction(SkipIfParamsHaveValue = new[] { "edit" })] // 加入此行會新增操作記錄到資料庫
        public ActionResult LogList(int? page, int? edit)
        {
            var pageNumber = page ?? 1;
            TempData[Constants.LOG_PAGE_NUMBER_KEY] = pageNumber;

            if (edit == null) TempData.Remove(Constants.LOG_SEARCH_PARAMETER_KEY);
            var searchParameters = TempData.Peek(Constants.LOG_SEARCH_PARAMETER_KEY) as LogListSearchViewModel;
            if (searchParameters == null)
            {
                searchParameters = new LogListSearchViewModel()
                {
                    ACTION_DT_START = DateTime.Now.ToString("yyyy-MM-dd"),
                    ACTION_DT_END = DateTime.Now.ToString("yyyy-MM-dd")
                };
                TempData[Constants.LOG_SEARCH_PARAMETER_KEY] = searchParameters;
            }

            var query = GetQueryData(searchParameters);
            var result = new LogListViewModel()
            {
                LogSearch = searchParameters,
                LogList = query.ToPagedList(pageNumber, pageSize)
            };

            return View(result);
        }

        /// <summary>
        /// 系統操作紀錄查詢-查詢按鈕
        /// </summary>
        /// <param name="Model"> 查詢條件 </param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult LogList(LogListViewModel Model)
        {
            TempData[Constants.LOG_PAGE_NUMBER_KEY] = 1;
            TempData[Constants.LOG_SEARCH_PARAMETER_KEY] = Model.LogSearch;

            var query = GetQueryData(Model.LogSearch);

            var result = new LogListViewModel()
            {
                LogSearch = Model.LogSearch,
                LogList = query.ToPagedList(1, pageSize)
            };

            return View(result);
        }

        #endregion

        #region 系統操作記錄-新增動作

        /// <summary> 儲存系統操作紀錄(使用者點擊功能頁面) </summary>
        /// <param name="path_name"> 連結名稱 </param>
        /// <remarks> 當使用者點擊選單時儲存點擊哪個功能項目的紀錄 </remarks>
        /// <returns></returns>
        [HttpPost]
        public ActionResult MenuClick(string path_name)
        {
            string usrId = WebSiteHelper.getUserId();
            if (string.IsNullOrWhiteSpace(path_name) || string.IsNullOrWhiteSpace(usrId))
            {
                return Json(new { success = false, message = "para empty" });
            }

            var menuText = db.SMCMENU3.FirstOrDefault(m => m.SYS_NAME == path_name.Trim())?.SYS_NAME;
            if (menuText == null)
            {
                return Json(new { success = false, message = "para error" });
            }

            try
            {
                var logEntry = new SMCLOGS
                {
                    USR_ID = usrId,
                    ACTION_TYPE = "M",
                    ACTION_DESC = menuText,
                    LOGIN_IP = WebSiteHelper.GetIPAddress(),
                    LOGIN_DT = DateTime.Now,
                    ADD_DT = DateTime.Now
                };
                db.SMCLOGS.Add(logEntry);
                db.SaveChanges();

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "save error", error = ex.Message });
            }
        }

        #endregion
    }
}