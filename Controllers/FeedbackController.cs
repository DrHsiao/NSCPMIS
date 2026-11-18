using PagedList;
using System;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WCals.Exceptions;
using WCals.Helpers;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEMQAModels;
using WCals.ViewModels;

namespace WCals.Controllers
{
    public class FeedbackController : Controller
    {
        private readonly CMBASEEntities dbBASE = new CMBASEEntities();
        private readonly CMEMQAEntities dbMQA = new CMEMQAEntities();
        const int pageSize = 10;
        const string MQA_FILE_PATH = "MQAFILE";

        #region 使用者意見紀錄-查詢頁面操作

        /// <summary>
        /// 取得使用者意見內容分類清單
        /// </summary>
        /// <param name="includePlaceholder"> 是否加入「請選擇」選項 </param>
        /// <returns></returns>
        private SelectList GetFeedbackTypesList(bool includePlaceholder)
        {
            var items = dbMQA.SMCCODE
                .Where(w => w.SYS_ID == "DT" && w.CATY_ID == "02")
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
        /// 使用者意見紀錄查詢-資料庫查詢
        /// </summary>
        /// <param name="SearchParameter"> 查詢條件 </param>
        /// <returns></returns>
        private IQueryable<FeedbackListItemViewModel> GetQueryData(FeedbackSearchViewModel SearchParameter)
        {
            ViewBag.FeedbackTypeList = GetFeedbackTypesList(true);

            // 取得使用者意見資料
            var query = dbMQA.MQAMAIN.AsNoTracking()
                .Join(dbMQA.SMCCODE,
                feedback => feedback.APL_TYPE,
                code => code.CATY_COD,
                 (feedback, code) => new { feedback, code })
                .Where(w => w.code.SYS_ID == "DT" && w.code.CATY_ID == "02")
                .Select(result => new FeedbackListItemViewModel
                {
                    MAIN_ID = result.feedback.MAIN_ID,
                    APL_TYPE = result.feedback.APL_TYPE,
                    APL_TYPE_NAME = result.code.CODE_DESC,
                    QUE_HEAD = result.feedback.QUE_HEAD != null && result.feedback.QUE_HEAD.Length >= 20 ? result.feedback.QUE_HEAD.Substring(0, 20) + "..." : result.feedback.QUE_HEAD,
                    QUE_DATA = result.feedback.QUE_DATA,
                    OPEN_PEP = result.feedback.OPEN_PEP ?? 0,
                    RETURN_PEP = result.feedback.RETURN_PEP ?? 0,
                    QUE_USER = result.feedback.QUE_USER,
                    ADD_DT = result.feedback.ADD_DT
                }).AsQueryable();

            #region 查詢條件

            if (null != SearchParameter)
            {
                // 內容分類
                if (!string.IsNullOrWhiteSpace(SearchParameter.APL_TYPE))
                {
                    query = query.Where(f => f.APL_TYPE == SearchParameter.APL_TYPE.Trim());
                }
                // 問題主旨
                if (!string.IsNullOrWhiteSpace(SearchParameter.QUE_HEAD))
                {
                    query = query.Where(f => f.QUE_HEAD.Contains(SearchParameter.QUE_HEAD.Trim()));
                }
                // 問題內容
                if (!string.IsNullOrEmpty(SearchParameter.QUE_DATA))
                {
                    query = query.Where(f => f.QUE_DATA.Contains(SearchParameter.QUE_DATA.Trim()));
                }
                // (西元年)
                // 起始日期
                DateTime parsedDATE_START;
                if (DateTime.TryParse(SearchParameter.STARTDATE, out parsedDATE_START))
                {
                    query = query.Where(m => DbFunctions.TruncateTime(m.ADD_DT) >= parsedDATE_START.Date);
                }
                // 迄止日期
                DateTime parsedDATE_END;
                if (DateTime.TryParse(SearchParameter.ENDDATE, out parsedDATE_END))
                {
                    query = query.Where(m => DbFunctions.TruncateTime(m.ADD_DT) <= parsedDATE_END.Date);
                }
            }

            #endregion

            return query.OrderByDescending(f => f.MAIN_ID);
        }

        /// <summary>
        /// 使用者意見紀錄查詢-初始/切換頁面
        /// </summary>
        /// <param name="page"> 分頁 </param>
        /// <param name="edit"> 切換頁面時為1 </param>
        /// <returns></returns>
        //[LogAction(SkipIfParamsHaveValue = new[] { "edit" })] // 加入此行會新增操作記錄到資料庫
        public ActionResult FeedbackList(int? page, int? edit)
        {
            var pageNumber = page ?? 1;
            TempData[Constants.MQA_PAGE_NUMBER_KEY] = pageNumber;

            if (edit == null) TempData.Remove(Constants.MQA_SEARCH_PARAMETER_KEY);
            var searchParameters = TempData.Peek(Constants.MQA_SEARCH_PARAMETER_KEY) as FeedbackSearchViewModel;
            if (searchParameters == null)
            {
                searchParameters = new FeedbackSearchViewModel()
                {
                    STARTDATE = DateTime.Now.AddMonths(-3).ToString("yyyy-MM-dd"),
                };
                TempData[Constants.MQA_SEARCH_PARAMETER_KEY] = searchParameters;
            }

            var query = GetQueryData(searchParameters);
            var result = new FeedbackListViewModel()
            {
                FeedbackSearch = searchParameters,
                FeedbackList = query.ToPagedList(pageNumber, pageSize)
            };

            return View(result);
        }

        /// <summary>
        /// 使用者意見紀錄查詢-查詢按鈕操作
        /// </summary>
        /// <param name="Model"> 查詢條件 </param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult FeedbackList(FeedbackListViewModel Model)
        {
            TempData[Constants.MQA_PAGE_NUMBER_KEY] = 1;
            TempData[Constants.MQA_SEARCH_PARAMETER_KEY] = Model.FeedbackSearch;

            var query = GetQueryData(Model.FeedbackSearch);

            var result = new FeedbackListViewModel()
            {
                FeedbackSearch = Model.FeedbackSearch,
                FeedbackList = query.ToPagedList(1, pageSize)
            };

            return View(result);
        }

        #endregion

        #region 使用者意見紀錄-檢視/回覆頁面操作

        /// <summary>
        /// 顯示使用者意見紀錄-詳細資料畫面(檢視意見內容、相關檔案、回覆等)
        /// </summary>
        /// <param name="id"> 意見編號 </param>
        /// <returns></returns>
        public ActionResult FeedbackDetails(string id)
        {
            var data = dbMQA.MQAMAIN
                .Join(dbMQA.SMCCODE,
                feedback => feedback.APL_TYPE,
                code => code.CATY_COD,
                 (feedback, code) => new { feedback, code })
                .Where(w => w.code.SYS_ID == "DT" && w.code.CATY_ID == "02")
                .FirstOrDefault(x => x.feedback.MAIN_ID == id);
            if (data == null)
            {
                return RedirectToAction("ErrorPage", "Error", new { errorCode = 404 });
            }

            // 檢查是否已有 cookie(避免重複更新點擊數)
            string cookieKey = $"viewed_feedback_{id}";
            if (Request.Cookies[cookieKey] == null)
            {
                // 更新點擊數
                data.feedback.OPEN_PEP++;
                dbMQA.SaveChanges();

                // 設定 cookie
                var cookie = new HttpCookie(cookieKey, "1");
                cookie.Expires = DateTime.Now.AddDays(1); // 1天內不再記錄
                Response.Cookies.Add(cookie);
            }

            var fileList = dbMQA.MQAFILE.Where(x => x.MAIN_ID == id).ToList();
            var replyList = dbMQA.MQAFORM.Where(x => x.MAIN_ID == id).OrderByDescending(x => x.ADD_DT).ToList();

            var model = new FeedbackDetailsViewModel();
            model.FeedbackDetails = new FeedbackListItemViewModel()
            {
                MAIN_ID = data.feedback.MAIN_ID,
                APL_TYPE = data.feedback.APL_TYPE,
                APL_TYPE_NAME = data.code.CODE_DESC,
                QUE_HEAD = data.feedback.QUE_HEAD,
                QUE_DATA = data.feedback.QUE_DATA,
                OPEN_PEP = data.feedback.OPEN_PEP ?? 0,
                RETURN_PEP = data.feedback.RETURN_PEP ?? 0,
                QUE_USER = data.feedback.QUE_USER,
                ADD_DT = data.feedback.ADD_DT
            };
            model.FeedbackFileList = fileList;
            model.FeedbackReplyList = replyList;
            model.UserCanReply = GetUserCanReply();
            return View(model);
        }

        /// <summary>
        /// 檢視使用者意見相關檔案
        /// </summary>
        /// <param name="id"> 檔案編號 </param>
        /// <returns></returns>
        public ActionResult ViewFile(string id)
        {
            var file = dbMQA.MQAFILE.FirstOrDefault(s => s.FID == id);
            if (file == null)
            {
                return RedirectToAction("ErrorPage", "Error", new { errorCode = 999, message = "找不到檔案" });
            }

            string physicalPath = Server.MapPath(file.FILE_PATH);
            if (!System.IO.File.Exists(physicalPath))
            {
                return RedirectToAction("ErrorPage", "Error", new { errorCode = 999, message = "找不到檔案" });
            }

            string contentType = MimeMapping.GetMimeMapping(physicalPath);
            return File(physicalPath, contentType);
        }

        /// <summary>
        /// 新增使用者意見回覆資料
        /// </summary>
        /// <param name="id"> 意見編號(MAIN_ID) </param>
        /// <param name="content"> 回覆內容 </param>
        /// <returns></returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult CreateReply(string id, string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                TempData["msg"] = "回覆內容不能空白";
                return RedirectToAction("FeedbackDetails", "Feedback", new { id });
            }

            var data = dbMQA.MQAMAIN.FirstOrDefault(x => x.MAIN_ID == id);
            if (data == null)
                return RedirectToAction("ErrorPage", "Error", new { errorCode = 999, message = "找不到相關的使用者意見紀錄" });

            string usrID = WebSiteHelper.getUserId();
            if (!GetUserCanReply(usrID))
                return RedirectToAction("ErrorPage", "Error", new { errorCode = 999, message = "您沒有權限回覆使用者意見" });

            string usrName = WebSiteHelper.getUserName();
            var reply = new MQAFORM
            {
                FRM_ID = GetDateTimeNum(),
                MAIN_ID = id,
                FRM_DATA = content,
                FRM_USER = usrName,
                ADV_UID = usrID,
                ADD_DT = DateTime.Now
            };

            try
            {
                dbMQA.MQAFORM.Add(reply);
                data.RETURN_PEP += 1; // 更新回覆數
                dbMQA.SaveChanges();

                TempData["msg"] = "儲存成功";
                return RedirectToAction("FeedbackDetails", "Feedback", new { id });
            }
            catch (Exception ex)
            {
                return RedirectToAction("ErrorPage", "Error", new { errorCode = 999, message = "新增失敗，請聯絡系統管理員。" + ex.Message });
            }
        }

        /// <summary>
        /// 更新使用者意見回覆資料
        /// </summary>
        /// <param name="id"> 回覆編號(FRM_ID) </param>
        /// <param name ="content"> 回覆內容 </param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult EditReply(string id, string content)
        {
            if (string.IsNullOrWhiteSpace(content))
                return new HttpStatusCodeResult(403);

            string usrID = WebSiteHelper.getUserId();
            if (!GetUserCanReply(usrID))
                return new HttpStatusCodeResult(403);

            var reply = dbMQA.MQAFORM.Find(id);
            if (reply == null || reply.ADV_UID != usrID)
                return new HttpStatusCodeResult(403);

            try
            {
                reply.FRM_DATA = content;
                reply.ADD_DT = DateTime.Now; // 更新回覆時間(會覆蓋原本建立時間) // => TODO：需確認是否符合業務需求
                dbMQA.SaveChanges();
                return new HttpStatusCodeResult(200);
            }
            catch (Exception ex)
            {
                return new HttpStatusCodeResult(500, "修改失敗，請聯絡系統管理員。" + ex.Message);
            }
        }

        /// <summary>
        /// 刪除使用者意見回覆資料
        /// </summary>
        /// <param name="id"> 回覆編號(FRM_ID) </param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult DeleteReply(string id)
        {
            string usrID = WebSiteHelper.getUserId();
            if (!GetUserCanReply(usrID))
                return new HttpStatusCodeResult(403);

            var reply = dbMQA.MQAFORM.Find(id);
            if (reply == null || reply.ADV_UID != usrID)
                return new HttpStatusCodeResult(403);

            var data = dbMQA.MQAMAIN.FirstOrDefault(x => x.MAIN_ID == reply.MAIN_ID);
            if (data == null)
                return HttpNotFound("找不到相關的使用者意見紀錄");

            try
            {
                dbMQA.MQAFORM.Remove(reply);
                data.RETURN_PEP -= 1; // 更新回覆數
                dbMQA.SaveChanges();
                return new HttpStatusCodeResult(200);
            }
            catch (Exception ex)
            {
                return new HttpStatusCodeResult(500, "刪除失敗，請聯絡系統管理員。" + ex.Message);
            }
        }

        /// <summary>
        /// 取得使用者是否有權限回覆意見
        /// </summary>
        /// <param name="usrID"> 使用者帳號 </param>
        /// <remarks> 目前只開放群組為SA的帳號使用 </remarks>
        /// <returns></returns>
        public bool GetUserCanReply(string usrID = "")
        {
            usrID = string.IsNullOrEmpty(usrID) ? Helpers.WebSiteHelper.getUserId() : usrID;
            return dbBASE.OGCDPUR.Any(w => w.DATA_TP == "G" && w.DEPT_NO == "SA" && w.USR_ID == usrID);
        }

        #endregion

        #region 使用者意見紀錄-新增表單操作

        /// <summary>
        /// 顯示使用者意見紀錄-新增表單畫面
        /// </summary>
        /// <returns></returns>
        public ActionResult FeedbackCreate()
        {
            ViewBag.FeedbackTypeList = GetFeedbackTypesList(false);
            var model = new FeedbackDataViewModel();
            return View(model);
        }

        /// <summary>
        /// 新增使用者意見資料
        /// </summary>
        /// <param name="model"> 資料來源 </param>
        /// <param name="email_confirm"> bot-Honeypot欄位 </param>
        /// <param name="form_loaded"> bot-填表時間欄位 </param>
        /// <returns></returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult FeedbackCreate(FeedbackDataViewModel model, string email_confirm, long form_loaded)
        {
            if (this.IsBot(email_confirm, form_loaded))
            {
                return new HttpStatusCodeResult(200); // 偵測到自動機器人行為，直接返回200 OK(不處理請求)
            }

            if (!ModelState.IsValid)
            {
                // 因瀏覽器安全限制，檔案上傳欄位無法在 View(model) 時自動保留內容，因此若有上傳檔案需提示要重新上傳
                if (model.FileUploads.Any(f => f.File != null))
                    ModelState.AddModelError("", "因瀏覽器安全限制，請重新上傳檔案。");
                return ReturnIfCreateInvalid(model, "請檢查輸入資料是否正確。");
            }

            try
            {
                string mainID = GetDateTimeNum();

                ProcessFeedbackData(model, mainID);
                ProcessFeedbackFiles(model, mainID);

                dbMQA.SaveChanges();

                TempData["msg"] = "新增成功";
                return RedirectToAction("FeedbackList");
            }
            catch (BusinessException ex)
            {
                return ReturnIfCreateInvalid(model, ex.Message);
            }
            catch (Exception ex)
            {
                return ReturnIfCreateInvalid(model, "儲存失敗，請聯絡系統管理員。" + ex.ToString());
            }
        }

        /// <summary>
        /// 返回新增表單時的錯誤訊息
        /// </summary>
        /// <param name="model"></param>
        /// <param name="error"> 錯誤訊息 </param>
        /// <returns></returns>
        private ActionResult ReturnIfCreateInvalid(object model, string error)
        {
            ViewBag.FeedbackTypeList = GetFeedbackTypesList(false);
            ModelState.AddModelError("", error);
            return View(model); // 返回原表單
        }

        /// <summary>
        /// 處理使用者意見資料
        /// </summary>
        /// <param name="model"></param>
        /// <exception cref="Exception"></exception>"
        private void ProcessFeedbackData(FeedbackDataViewModel model, string mainID)
        {
            // 取得使用者名稱和部門名稱
            string usrId = Helpers.WebSiteHelper.getUserId();
            var result = (from t1 in dbBASE.OGCUSER
                          join t2 in dbBASE.OGCDPUR on t1.USR_ID equals t2.USR_ID
                          join t3 in dbBASE.OGCDEPT on t2.DEPT_NO equals t3.DEPT_NO
                          where t1.USR_ID == usrId && t2.DATA_TP == "D" && t3.DATA_TP == "D"
                          select new { usrName = t1.USR_NAME, deptName = t3.DEPT_NAME }).FirstOrDefault();

            var message = new MQAMAIN
            {
                MAIN_ID = mainID,
                MAIN_TYPE = "TD",
                QUE_HEAD = model.QUE_HEAD,
                QUE_DATA = model.QUE_DATA,
                QUE_DEPT = result.deptName,
                QUE_USER = result.usrName,
                USER_MAIL = model.USER_MAIL,
                USER_TEL = model.USER_TEL,
                APL_TYPE = model.APL_TYPE,
                OPEN_PEP = 0,
                RETURN_PEP = 0,
                SHOW_FG = true,
                ADV_UID = usrId,
                ADD_DT = DateTime.Now
            };
            dbMQA.MQAMAIN.Add(message);
        }

        /// <summary>
        /// 處理使用者意見相關檔案
        /// </summary>
        /// <param name="model"></param>
        /// <param name="mainID"></param>
        /// <exception cref="Exception"></exception>
        private void ProcessFeedbackFiles(FeedbackDataViewModel model, string mainID)
        {
            string[] allowedExtensions = { ".jpg", ".jpeg", ".png" };
            int maxSizeBase = 20;
            int maxSize = maxSizeBase * 1024 * 1024; // 20MB

            foreach (var item in model.FileUploads)
            {
                if (item.File == null || item.File.ContentLength <= 0)
                {
                    continue;
                }

                string originalFileName = Path.GetFileName(item.File.FileName);
                string ext = Path.GetExtension(originalFileName).ToLower();
                //ext = ""; // 這行是為了測試不符合副檔名的情況
                if (!allowedExtensions.Contains(ext))
                {
                    throw new BusinessException("只允許上傳JPG、PNG 檔案。");
                }
                //maxSize = 0; // 這行是為了測試檔案大小超過限制的情況
                if (item.File.ContentLength > maxSize)
                {
                    throw new BusinessException($"檔案大小不能超過{maxSizeBase}MB。");
                }

                string fid = GetDateTimeNum();
                string newFileName = fid + ext;

                // 資料夾路徑：MQAFILE/yyyy-MM-dd/
                string date = DateTime.Now.ToString("yyyy-MM-dd");
                string folderPath = Path.Combine(MQA_FILE_PATH, date);
                string physicalFolderPath = Server.MapPath($"~/{folderPath}");
                if (!Directory.Exists(physicalFolderPath))
                {
                    Directory.CreateDirectory(physicalFolderPath); // 若資料夾不存在就建立
                }
                // 儲存檔案
                string path = Path.Combine(physicalFolderPath, newFileName);
                item.File.SaveAs(path);

                var file = new MQAFILE
                {
                    FID = fid,
                    MAIN_ID = mainID,
                    FILE_NAME = newFileName,
                    O_FILE_NAME = originalFileName,
                    FILE_DESC = item.FILE_DESC ?? "",
                    FILE_PATH = $"~/{MQA_FILE_PATH}/{date}/{newFileName}" // 資料庫存檔路徑：~/MQAFILE/yyyy-MM-dd/檔名
                };
                dbMQA.MQAFILE.Add(file);
            }
        }

        /// <summary>
        /// 取得時間數字當作ID
        /// </summary>
        /// <returns></returns>
        public string GetDateTimeNum()
        {
            DateTime now = DateTime.Now;
            return now.Year.ToString("0000") + now.Month.ToString("00") + now.Day.ToString("00") + now.Hour.ToString("00") + now.Minute.ToString("00") + now.Second.ToString("00") + now.Millisecond.ToString("000");
        }

        #endregion
    }
}