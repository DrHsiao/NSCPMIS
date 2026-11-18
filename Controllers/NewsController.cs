using Microsoft.Security.Application;
using PagedList;
using System;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using WCals.Exceptions;
using WCals.Helpers;
using WCals.Models.CMBASEModels;
using WCals.ViewModels;

namespace WCals.Controllers
{
    public partial class Constants
    {
        public const string LOG_PAGE_NUMBER_KEY = "tempLOGPN";
        public const string LOG_SEARCH_PARAMETER_KEY = "tempLOGSP";
        public const string NEWS_PAGE_NUMBER_KEY = "tempNEWSPN";
        public const string NEWS_SEARCH_PARAMETER_KEY = "tempNEWSSP";
        public const string MQA_PAGE_NUMBER_KEY = "tempMQAPN";
        public const string MQA_SEARCH_PARAMETER_KEY = "tempMQASP";
    }

    public class NewsController : Controller
    {
        private readonly CMBASEEntities db = new CMBASEEntities();
        private readonly int pageSize = 10;

        #region 最新消息維護-查詢頁面操作

        /// <summary>
        /// 取得最新消息訊息類別清單
        /// </summary>
        /// <param name="includePlaceholder"> 是否需加入「請選擇」選項 </param>
        /// <returns></returns>
        private SelectList GetNewsTypesList(bool includePlaceholder)
        {
            var items = db.SMCNEWSCODE
                .Where(w => w.SYS_ID == "SY" && w.CATY_ID == "01")
                .Select(x => new SelectListItem
                {
                    Value = x.CATY_COD,
                    Text = x.CODE_DESC
                }).ToList();

            if (includePlaceholder)
            {
                items.Insert(0, new SelectListItem { Text = "請選擇", Value = "" });
            }

            return new SelectList(items, "Value", "Text");
        }

        /// <summary>
        /// 最新消息維護查詢-資料查詢
        /// </summary>
        /// <param name="SearchParameter"> 查詢條件 </param>
        /// <returns></returns>
        private IQueryable<NewsListItemViewModel> GetQueryData(NewsSearchViewModel SearchParameter)
        {
            ViewBag.NewsTypeList = GetNewsTypesList(true);

            // 取得最新消息資料
            var query = db.SMCNEWS.AsNoTracking()
                .Join(db.SMCNEWSCODE,
                news => news.TYPE,
                code => code.CATY_COD,
                (news, code) => new NewsListItemViewModel
                {
                    IsAnnounced = ((news.STARTDATE != null && DbFunctions.TruncateTime(news.STARTDATE) <= DbFunctions.TruncateTime(DateTime.Today)) &&
                (news.ENDDATE == null || DbFunctions.TruncateTime(news.ENDDATE) >= DbFunctions.TruncateTime(DateTime.Today)))
                ? "T" : "F",
                    NEWSID = news.NEWSID,
                    HEADLINE = news.HEADLINE != null && news.HEADLINE.Length >= 15 ? news.HEADLINE.Substring(0, 15) + "..." : news.HEADLINE,
                    SUMMARY = news.SUMMARY != null && news.SUMMARY.Length >= 20 ? news.SUMMARY.Substring(0, 20) + "..." : news.SUMMARY,
                    STARTDATE = news.STARTDATE,
                    ENDDATE = news.ENDDATE,
                    TYPE_NAME = code.CODE_DESC,
                    TYPE = news.TYPE
                }).AsQueryable();

            #region 查詢條件

            if (null != SearchParameter)
            {
                // 訊息類別
                if (!string.IsNullOrWhiteSpace(SearchParameter.TYPE))
                {
                    query = query.Where(f => f.TYPE == SearchParameter.TYPE.Trim());
                }
                // 標題
                if (!string.IsNullOrWhiteSpace(SearchParameter.HEADLINE))
                {
                    query = query.Where(f => f.HEADLINE.Contains(SearchParameter.HEADLINE.Trim()));
                }
                // 內容
                if (!string.IsNullOrEmpty(SearchParameter.SUMMARY))
                {
                    query = query.Where(f => f.SUMMARY.Contains(SearchParameter.SUMMARY.Trim()));
                }
                // (西元年)
                // 起始日期
                DateTime parsedDATE_START;
                if (DateTime.TryParse(SearchParameter.STARTDATE, out parsedDATE_START))
                {
                    query = query.Where(m => DbFunctions.TruncateTime(m.STARTDATE) >= parsedDATE_START.Date
                    ); //|| DbFunctions.TruncateTime(m.ENDDATE) >= parsedDATE_START.Date); // [可變動]若要顯示在查詢區間公告中的資料，則需加入此行條件；若只是查詢公告起始日期的區間，則可移除此行條件
                }
                // 迄止日期
                DateTime parsedDATE_END;
                if (DateTime.TryParse(SearchParameter.ENDDATE, out parsedDATE_END))
                {
                    query = query.Where(m => DbFunctions.TruncateTime(m.STARTDATE) <= parsedDATE_END.Date);
                }
            }

            #endregion

            return query.OrderByDescending(f => f.NEWSID);
        }

        /// <summary>
        /// 最新消息維護查詢-初始/切換頁面
        /// </summary>
        /// <param name="page"> 分頁 </param>
        /// <param name="edit"> 切換頁面時為1 </param>
        /// <returns></returns>
        //[LogAction(SkipIfParamsHaveValue = new[] { "edit" })] // 加入此行會新增操作記錄到資料庫(但當參數值為 "edit" 時會跳過記錄)
        public ActionResult NewsList(int? page, int? edit)
        {
            var pageNumber = page ?? 1;
            TempData[Constants.NEWS_PAGE_NUMBER_KEY] = pageNumber;

            if (edit == null) TempData.Remove(Constants.NEWS_SEARCH_PARAMETER_KEY);
            var searchParameters = TempData.Peek(Constants.NEWS_SEARCH_PARAMETER_KEY) as NewsSearchViewModel;
            if (searchParameters == null)
            {
                searchParameters = new NewsSearchViewModel()
                {
                    STARTDATE = DateTime.Now.AddMonths(-3).ToString("yyyy-MM-dd"),
                };
                TempData[Constants.NEWS_SEARCH_PARAMETER_KEY] = searchParameters;
            }

            var query = GetQueryData(searchParameters);
            var result = new NewsListViewModel()
            {
                NewsSearch = searchParameters,
                NewsList = query.ToPagedList(pageNumber, pageSize)
            };

            return View(result);
        }

        /// <summary>
        /// 最新消息維護查詢-查詢按鈕
        /// </summary>
        /// <param name="Model"> 查詢條件 </param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult NewsList(NewsListViewModel Model)
        {
            TempData[Constants.NEWS_PAGE_NUMBER_KEY] = 1;
            TempData[Constants.NEWS_SEARCH_PARAMETER_KEY] = Model.NewsSearch;

            var query = GetQueryData(Model.NewsSearch);

            var result = new NewsListViewModel()
            {
                NewsSearch = Model.NewsSearch,
                NewsList = query.ToPagedList(1, pageSize)
            };

            return View(result);
        }

        #endregion

        #region 最新消息維護-表單操作

        /// <summary>
        /// 顯示最新消息維護-新增/編輯畫面
        /// </summary>
        /// <param name="id"> 最新消息編號(如為0則新增) </param>
        /// <returns></returns>
        public ActionResult NewsEdit(int id = 0)
        {
            TempData.Keep(Constants.NEWS_SEARCH_PARAMETER_KEY);
            ViewBag.NewsTypeList = GetNewsTypesList(false);

            var model = new NewsDataViewModel();

            if (id > 0) // 編輯模式
            {
                model = db.SMCNEWS
                        .Where(w => w.NEWSID == id)
                        .Select(s => new NewsDataViewModel
                        {
                            NEWSID = s.NEWSID,
                            HEADLINE = s.HEADLINE,
                            BODY = s.BODY,
                            TYPE = s.TYPE,
                            STARTDATE = DbFunctions.TruncateTime(s.STARTDATE),
                            ENDDATE = DbFunctions.TruncateTime(s.ENDDATE),
                        }).FirstOrDefault();
                if (model == null)
                {
                    return RedirectToAction("ErrorPage", "Error", new { errorCode = 404 });
                }

                ViewBag.IsAnnounced = (model.STARTDATE != null && model.STARTDATE <= DateTime.Today);

                ViewBag.NewsFiles = db.SMCNEWSDOC.Where(w => w.NEWSID == model.NEWSID).ToList();
            }

            return View(model);
        }

        /// <summary>
        /// 儲存最新消息資料
        /// </summary>
        /// <param name="model"> 資料來源 </param>
        /// <param name="email_confirm"> bot-Honeypot欄位 </param>
        /// <param name="form_loaded"> bot-填表時間欄位 </param>
        /// <returns></returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult NewsEdit(NewsDataViewModel model, string email_confirm, long form_loaded)
        {
            if (this.IsBot(email_confirm, form_loaded))
            {
                return new HttpStatusCodeResult(200); // 偵測到自動機器人行為，直接返回200 OK(不處理請求)
            }

            ViewBag.NewsTypeList = GetNewsTypesList(false);

            if (!ModelState.IsValid)
            {
                return ReturnIfSaveInvalid(model, "請檢查輸入資料是否正確。");
            }

            try
            {
                SaveNews(model);
                ProcessNewsFiles(model);

                TempData["msg"] = "儲存成功";
                return RedirectToAction("NewsEdit", new { id = model.NEWSID }); // 使用RedirectToAction防止重複提交表單
            }
            catch (BusinessException ex)
            {
                return ReturnIfSaveInvalid(model, ex.Message);
            }
            catch (Exception ex)
            {
                return ReturnIfSaveInvalid(model, "儲存失敗，請聯絡系統管理員。" + ex.ToString());
            }
        }

        /// <summary>
        /// 返回儲存表單時的錯誤訊息
        /// </summary>
        /// <param name="model"></param>
        /// <param name="error"> 錯誤訊息 </param>
        /// <returns></returns>
        private ActionResult ReturnIfSaveInvalid(NewsDataViewModel model, string error)
        {
            if (model.NEWSID > 0) ViewBag.NewsFiles = db.SMCNEWSDOC.Where(w => w.NEWSID == model.NEWSID).ToList();
            ModelState.AddModelError("", error);
            return View(model);
        }

        /// <summary>
        /// 儲存最新消息資料
        /// </summary>
        /// <param name="model"></param>
        /// <exception cref="Exception"></exception>
        private void SaveNews(NewsDataViewModel model)
        {
            // 處理HTML內容
            string cleanedHtml = Sanitizer.GetSafeHtmlFragment(model.BODY);
            string noHtmlTags = Regex.Replace(cleanedHtml, "<.*?>", string.Empty);
            string cleanedTxt = WebUtility.HtmlDecode(noHtmlTags);

            SMCNEWS news;
            if (model.NEWSID > 0) // 編輯模式
            {
                news = db.SMCNEWS.FirstOrDefault(w => w.NEWSID == model.NEWSID);
                if (news != null)
                {
                    news.HEADLINE = model.HEADLINE;
                    news.SUMMARY = cleanedTxt;
                    news.BODY = cleanedHtml;
                    news.TYPE = model.TYPE;
                    news.STARTDATE = model.STARTDATE;
                    news.ENDDATE = model.ENDDATE;
                    news.UPD_DT = DateTime.Now;
                    news.UPD_UID = WebSiteHelper.getUserId();
                }
                else
                {
                    throw new BusinessException("找不到要編輯的最新消息資料。");
                }
            }
            else // 新增模式
            {
                news = new SMCNEWS
                {
                    HEADLINE = model.HEADLINE,
                    SUMMARY = cleanedTxt,
                    BODY = cleanedHtml,
                    TYPE = model.TYPE,
                    STARTDATE = model.STARTDATE,
                    ENDDATE = model.ENDDATE,
                    STATUS = 0,
                    ADD_DT = DateTime.Now,
                    UPD_UID = WebSiteHelper.getUserId()
                };
                db.SMCNEWS.Add(news);
            }

            try
            {
                db.SaveChanges();
                model.NEWSID = news.NEWSID; // 取得新增的NEWSID
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// 處理最新消息檔案上傳與儲存
        /// </summary>
        /// <param name="model"></param>
        /// <exception cref="Exception"></exception>
        private void ProcessNewsFiles(NewsDataViewModel model)
        {
            if (model.UploadedFiles != null)
            {
                string[] allowedExtensions = { ".pdf", ".xlsx", ".xls", ".docx", ".doc", ".zip", ".jpg", ".jpeg", ".png" };
                int maxSizeBase = 20;
                int maxSize = maxSizeBase * 1024 * 1024; // 20MB

                bool isAdd = false;
                SMCNEWSDOC newsFile;
                byte[] data;
                for (int i = 0; i < model.UploadedFiles.Length; i++)
                {
                    HttpPostedFileBase file = model.UploadedFiles[i];
                    if (file == null || file.ContentLength <= 0)
                    {
                        continue;
                    }

                    string ext = Path.GetExtension(file.FileName).ToLower();
                    if (!allowedExtensions.Contains(ext))
                    {
                        throw new BusinessException("只允許上傳 PDF、EXCEL、WORD、ZIP、JPG、PNG 檔案。");
                    }

                    if (file.ContentLength > maxSize)
                    {
                        throw new BusinessException($"檔案大小不能超過{maxSizeBase}MB。");
                    }

                    if (string.IsNullOrEmpty(model.FID[i]))
                    {
                        // 新增檔案
                        isAdd = true;
                        newsFile = new SMCNEWSDOC();
                        newsFile.NEWSID = model.NEWSID;
                        newsFile.FILID = DateTime.Now.Ticks;
                        newsFile.ADD_DT = DateTime.Now;
                    }
                    else
                    {
                        // 修改檔案
                        decimal fid = Convert.ToDecimal(model.FID[i]);
                        newsFile = db.SMCNEWSDOC
                            .Where(w => w.NEWSID == model.NEWSID && w.FILID == fid)
                            .FirstOrDefault();
                        if (newsFile == null)
                        {
                            continue;
                        }
                    }

                    using (Stream inputStream = file.InputStream)
                    {
                        MemoryStream memoryStream = inputStream as MemoryStream;
                        if (memoryStream == null)
                        {
                            memoryStream = new MemoryStream();
                            inputStream.CopyTo(memoryStream);
                        }
                        data = memoryStream.ToArray();
                    }

                    newsFile.FILE_PATH = file.FileName;
                    newsFile.FILE_BODY = data;
                    newsFile.FILE_SIZE = data.Length;
                    newsFile.UPD_DT = DateTime.Now;
                    newsFile.UPD_UID = WebSiteHelper.getUserId();
                    if (isAdd) db.SMCNEWSDOC.Add(newsFile);
                }

                try
                {
                    db.SaveChanges();
                }
                catch
                {
                    throw;
                }
            }
        }

        /// <summary>
        /// 刪除最新消息資料
        /// </summary>
        /// <param name="id"> 最新消息編號(NEWSID) </param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult NewsDelete(int id)
        {
            try
            {
                var data = db.SMCNEWS.Find(id);
                if (data == null)
                {
                    return Json(new { success = false, message = "資料不存在" });
                }

                // 一併刪除相關的檔案
                var files = db.SMCNEWSDOC.Where(b => b.NEWSID == id).ToList();
                db.SMCNEWSDOC.RemoveRange(files);

                db.SMCNEWS.Remove(data);
                db.SaveChanges();

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// 刪除最新消息檔案資料
        /// </summary>
        /// <param name="id"> 最新消息編號(NEWSID) </param>
        /// <param name="fid"> 檔案編號(FID) </param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult NewsFileDelete(decimal id, decimal fid)
        {
            try
            {
                var data = db.SMCNEWSDOC.Where(b => b.NEWSID == id && b.FILID == fid).FirstOrDefault();
                if (data == null)
                {
                    return Json(new { success = false, message = "資料不存在" });
                }

                db.SMCNEWSDOC.Remove(data);
                db.SaveChanges();

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// 下載最新消息檔案
        /// </summary>
        /// <param name="id"> 最新消息編號(NEWSID) </param>
        /// <param name="fid"> 檔案編號(FID) </param>
        /// <returns></returns>
        public ActionResult DownloadNewsFile(decimal id, decimal fid)
        {
            var file = db.SMCNEWSDOC.Where(b => b.NEWSID == id && b.FILID == fid).FirstOrDefault();
            if (file == null)
            {
                return HttpNotFound();
            }

            var contentType = MimeMapping.GetMimeMapping(file.FILE_PATH);
            return File(file.FILE_BODY, contentType, file.FILE_PATH);
        }

        #endregion
    }
}