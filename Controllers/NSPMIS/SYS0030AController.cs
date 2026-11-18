using PagedList;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using System.Web.UI;
using WCals.classes.john;
using WCals.Helpers;
using WCals.ViewModels.NewModels;
using static System.Runtime.CompilerServices.RuntimeHelpers;

namespace WCals.Controllers.NSPMIS {
    public class SYS0030AController : Controller {

        // GET: SYS0030A
        public ActionResult Index(SYS0030Model model) {

            ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);
            ViewBag.towns = tools.ToSelectList(util.GetTowns(null), "TOWN", "TOWN", null);
            // 加入污水下水道系統下拉選單資料 
            SqlCommand cmdSystems = new SqlCommand("SELECT SYS_ID, SYS_NAME FROM PJCSYST WHERE STAT_NO='R' ORDER BY SYS_NAME");
            DataTable dtSystems = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSystems);
            ViewBag.systems = tools.ToSelectList(dtSystems, "SYS_ID", "SYS_NAME", null);

            return View("SYS0030AIndex", model);
        }

        int pageSize = 5;


        public ActionResult Search(string sys, string city, string doc_name, string doc_type, int page) {
            SYS0030Model model = new SYS0030Model {
                sys = sys,
                CITY = city,
                doc_name = doc_name,
                doc_type = doc_type,
                page = page
            };

            return Search(model);
        }

        // POST: SYS0030A首頁查詢
        [HttpPost]
        public ActionResult Search(SYS0030Model model) {

            if(model.page == 0) {
                model.page = 1;
            }

            StringBuilder sb = new StringBuilder("SELECT s.SYS_NAME, p.* FROM PJCFORM p, PJCSYST s WHERE p.SYS_ID in (select SYS_ID from PJCSYST where STAT_NO='R') AND p.SYS_ID=s.SYS_ID ");

            SqlCommand cmd = new SqlCommand();
            if(!string.IsNullOrEmpty(model.sys)) {
                cmd.Parameters.Add(new SqlParameter("SYS_NAME", model.sys + "%"));
                sb.Append("AND s.SYS_NAME LIKE @SYS_NAME ");
            }

            if(!string.IsNullOrEmpty(model.CITY)) {
                cmd.Parameters.Add(new SqlParameter("SER_CITY", model.CITY));
                sb.Append("AND p.SER_CITY LIKE @SER_CITY ");
            }

            if(!string.IsNullOrEmpty(model.doc_type)) {
                cmd.Parameters.Add(new SqlParameter("PJ_ID_TYPE", model.doc_type));
                sb.Append("AND p.PJ_ID_TYPE LIKE @PJ_ID_TYPE ");
            }

            cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, model.page, pageSize);
            model.List1 = plist;
            return Index(model);
        }

        /// <summary>
        /// 系統別管理
        /// </summary>
        /// <returns></returns>
        public ActionResult SYSManager(SYS0030Model model) {

            ViewBag.cities = util.GetCities();
            if(model.page == 0) {
                model.page = 1;
            }
            
            return View("SYSManager", model);
        }


        public ActionResult SYSManagerSearch(string CITY, string SYS_ID, string SYS_NAME, string CONST_TP, string IS_SUP, int page = 1) {

            SYS0030Model model = new SYS0030Model {
                CITY = CITY,
                SYS_ID = SYS_ID,
                SYS_NAME = SYS_NAME,
                CONST_TP = CONST_TP,
                IS_SUP = IS_SUP,
                page = page
            };

            return SYSManagerSearch(model);
        }

        /// <summary>
        /// 系統別管理
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SYSManagerSearch")]
        public ActionResult SYSManagerSearch(SYS0030Model model) {

            ViewBag.cities = util.GetCities();
            if(model.page == 0) {
                model.page = 1;
            }

            StringBuilder sb = new StringBuilder("SELECT*FROM PJCSYST WHERE STAT_NO='S' ");

            SqlCommand cmd = new SqlCommand();
            if(!string.IsNullOrEmpty(model.CITY)) {
                cmd.Parameters.Add(new SqlParameter("CITY", model.CITY));
                sb.Append("AND CITY=@CITY ");
            }
            if(!string.IsNullOrEmpty(model.SYS_ID)) {
                cmd.Parameters.Add(new SqlParameter("SYS_ID", model.SYS_ID));
                sb.Append("AND SYS_ID=@SYS_ID ");
            }
            if(!string.IsNullOrEmpty(model.SYS_NAME)) {
                cmd.Parameters.Add(new SqlParameter("SYS_NAME", "%" + model.SYS_NAME + "%"));
                sb.Append("AND SYS_NAME LIKE @SYS_NAME ");
            }
            if(!string.IsNullOrEmpty(model.CONST_TP)) {
                cmd.Parameters.Add(new SqlParameter("CONST_TP", "%" + model.CONST_TP + "%"));
                sb.Append("AND CONST_TP LIKE @CONST_TP ");
            }
            if(!string.IsNullOrEmpty(model.IS_SUP)) {
                cmd.Parameters.Add(new SqlParameter("IS_SUP", model.IS_SUP));
                sb.Append("AND IS_SUP = @IS_SUP ");
            }
            cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, model.page, pageSize);
            model.List1 = plist;
            //return Index(model);

            ModelState.Clear(); // ★ 關鍵：讓 DropDownList 不吃舊值
            return SYSManager(model);
        }


        /// <summary>
        /// 系統別管理
        /// </summary>
        /// <returns></returns>
        public ActionResult SYSManagerSave(SYS0030Model model) {

            ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);
            if(model.page == 0) {
                model.page = 1;
            }

            StringBuilder sb = new StringBuilder("SELECT*FROM PJCSYST WHERE STAT_NO='R' ");

            SqlCommand cmd = new SqlCommand();

            cmd.CommandText = sb.ToString();
            DataTable ps = BASE_DAO.SqlDao.CMETMSQuery4List(cmd);

            List<DataRow> list = ps.AsEnumerable().ToList();
            IPagedList<DataRow> plist = new PagedList<DataRow>(list, model.page, pageSize);
            model.List1 = plist;

            return SYSManager(model);
        }

        [HttpPost]
        public ActionResult SYSManagerAdd(SYS0030Model model) {
            try {
                // 驗證必填欄位
                if(string.IsNullOrEmpty(model.CITY)) {
                    ViewBag.Message = "請選擇縣市別";
                    return SYSManager(model);
                }
                if(string.IsNullOrEmpty(model.SYS_ID)) {
                    ViewBag.Message = "請輸入下水道系統編號";
                    return SYSManager(model);
                }
                if(string.IsNullOrEmpty(model.SYS_NAME)) {
                    ViewBag.Message = "請輸入下水道系統名稱";
                    return SYSManager(model);
                }

                // 檢查 SYS_ID 是否已存在
                SqlCommand checkCmd = new SqlCommand("SELECT COUNT(*) FROM PJCSYST WHERE SYS_ID = @SYS_ID");
                checkCmd.Parameters.Add(new SqlParameter("SYS_ID", model.SYS_ID));
                DataTable checkDt = BASE_DAO.SqlDao.CMETMSQuery4List(checkCmd);
                if(checkDt.Rows.Count > 0 && Convert.ToInt32(checkDt.Rows[0][0]) > 0) {
                    ViewBag.Message = "此系統編號已存在，請使用其他編號";
                    return SYSManager(model);
                }

                // 新增資料到 PJCSYST
                DateTime now = DateTime.Now;
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO PJCSYST (SYS_ID, CITY, SYS_NAME, CONST_TP, IS_SUP, ADD_DT, UPD_DT, UPD_UID, STAT_NO) " +
                    "VALUES (@SYS_ID, @CITY, @SYS_NAME, @CONST_TP, @IS_SUP, @ADD_DT, @UPD_DT, @UPD_UID, @STAT_NO)"
                );
                cmd.Parameters.Add(new SqlParameter("SYS_ID", model.SYS_ID));
                cmd.Parameters.Add(new SqlParameter("CITY", model.CITY));
                cmd.Parameters.Add(new SqlParameter("SYS_NAME", model.SYS_NAME ?? ""));
                cmd.Parameters.Add(new SqlParameter("CONST_TP", model.CONST_TP ?? ""));
                cmd.Parameters.Add(new SqlParameter("IS_SUP", model.IS_SUP ?? "N"));
                cmd.Parameters.Add(new SqlParameter("ADD_DT", now));
                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId() ?? ""));
                cmd.Parameters.Add(new SqlParameter("STAT_NO", "S"));

                BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                ViewBag.Message = "新增成功";

                // 清空表單並重新查詢（顯示所有資料）
                model.CITY = "";
                model.SYS_ID = "";
                model.SYS_NAME = "";
                model.CONST_TP = "";
                model.IS_SUP = "";

                return SYSManagerSearch(model);
            } catch(Exception ex) {
                ViewBag.Message = "新增失敗：" + ex.Message;
                return SYSManager(model);
            }
        }

        public ActionResult SYSInfo(string pjId, string mode = null) {
            try {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();
                WCals.Models.CMETMSModels.PJCFORM pjcform = null;

                if(!string.IsNullOrEmpty(pjId)) {
                    // 編輯模式或複製模式：載入既有記錄
                    System.Diagnostics.Debug.WriteLine("SYSInfo - 查詢 pjId: " + pjId);
                    pjcform = db.PJCFORM.FirstOrDefault(x => x.PJ_ID == pjId);

                    if(pjcform != null) {
                        System.Diagnostics.Debug.WriteLine("SYSInfo - 找到記錄: " + pjcform.PJ_NAME);

                        // 複製模式：清空 PJ_ID，讓儲存時作為新增處理
                        if(mode == "copy") {
                            pjcform.PJ_ID = null;
                        }
                    } else {
                        System.Diagnostics.Debug.WriteLine("SYSInfo - 未找到記錄");
                    }
                }

                if(pjcform == null) {
                    // 新增模式：建立空物件
                    System.Diagnostics.Debug.WriteLine("SYSInfo - 創建新物件");
                    pjcform = new WCals.Models.CMETMSModels.PJCFORM();
                }

                // 傳遞 mode 給 View
                ViewBag.Mode = mode;

                // 加入縣市下拉選單資料
                ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);

                // 加入污水下水道系統下拉選單資料 (使用 SQL 查詢)
                SqlCommand cmdSystems = new SqlCommand("SELECT SYS_ID, SYS_NAME FROM PJCSYST WHERE STAT_NO='R' ORDER BY SYS_NAME");
                DataTable dtSystems = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSystems);
                ViewBag.systems = tools.ToSelectList(dtSystems, "SYS_ID", "SYS_NAME", null);

                return View("SYSInfoA", pjcform);
            } catch(Exception ex) {
                System.Diagnostics.Debug.WriteLine("SYSInfo - 異常: " + ex.Message);
                // 即使發生異常也要返回新物件，避免 null reference
                ViewBag.cities = tools.ToSelectList(util.GetCities(), "CITY", "CITY", null);
                ViewBag.systems = new SelectList(new List<object>(), "SYS_ID", "SYS_NAME");
                return View("SYSInfoA", new WCals.Models.CMETMSModels.PJCFORM());
            }
        }

        public ActionResult SYSStat() {
            return View("SYSStatA");
        }

        /// <summary>
        /// 新增/修改規劃報告 (SYS0030A 只處理頁面上實際顯示的欄位)
        /// </summary>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "SaveSYSInfo")]
        public ActionResult SaveSYSInfo(WCals.Models.CMETMSModels.PJCFORM model) {
            try {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();
                string pjId = model.PJ_ID;

                if(!string.IsNullOrEmpty(model.PJ_ID)) {
                    // 編輯模式：更新既有記錄 (只更新頁面上有的欄位)
                    var existingPjcform = db.PJCFORM.FirstOrDefault(x => x.PJ_ID == model.PJ_ID);
                    if(existingPjcform != null) {
                        existingPjcform.SYS_ID = model.SYS_ID;
                        existingPjcform.PJ_ID_TYPE = model.PJ_ID_TYPE;
                        existingPjcform.SER_FN_1 = model.SER_FN_1;
                        existingPjcform.SER_FN_2 = model.SER_FN_2;
                        existingPjcform.SER_FN_3 = model.SER_FN_3;
                        existingPjcform.SER_FN_4 = model.SER_FN_4;
                        existingPjcform.SER_FN_5 = model.SER_FN_5;
                        existingPjcform.SER_FN_6 = model.SER_FN_6;
                        existingPjcform.SER_FN_7 = model.SER_FN_7;
                        existingPjcform.SER_CITY = model.SER_CITY;
                        existingPjcform.PJ_NAME = model.PJ_NAME;
                        existingPjcform.UPD_DT = DateTime.Now;

                        db.SaveChanges();
                        ViewBag.Message = "修改成功";
                    } else {
                        ViewBag.Message = "找不到該記錄";
                        return SYSInfo(pjId);
                    }
                } else {
                    // 新增模式：建立新記錄
                    // 驗證造冊編碼
                    if (string.IsNullOrEmpty(model.SER_FN_1)) {
                        ViewBag.Message = "請填寫造冊編碼";
                        return SYSInfo(null);
                    }

                    pjId = (model.SER_FN_1 ?? "") + (model.SER_FN_2 ?? "") + (model.SER_FN_3 ?? "") +
                           (model.SER_FN_4 ?? "") + (model.SER_FN_5 ?? "") + (model.SER_FN_6 ?? "") + (model.SER_FN_7 ?? "");

                    // 檢查 PJ_ID 是否已存在
                    var existingRecord = db.PJCFORM.FirstOrDefault(x => x.PJ_ID == pjId);
                    if (existingRecord != null) {
                        ViewBag.Message = "此造冊編碼已存在，請使用其他編碼";
                        return SYSInfo(null);
                    }

                    var newPjcform = new WCals.Models.CMETMSModels.PJCFORM();
                    newPjcform.PJ_ID = pjId;
                    newPjcform.SYS_ID = model.SYS_ID;
                    newPjcform.PJ_ID_TYPE = model.PJ_ID_TYPE;
                    newPjcform.SER_FN_1 = model.SER_FN_1;
                    newPjcform.SER_FN_2 = model.SER_FN_2;
                    newPjcform.SER_FN_3 = model.SER_FN_3;
                    newPjcform.SER_FN_4 = model.SER_FN_4;
                    newPjcform.SER_FN_5 = model.SER_FN_5;
                    newPjcform.SER_FN_6 = model.SER_FN_6;
                    newPjcform.SER_FN_7 = model.SER_FN_7;
                    newPjcform.SER_CITY = model.SER_CITY;
                    newPjcform.PJ_NAME = model.PJ_NAME;
                    newPjcform.UPD_DT = DateTime.Now;

                    db.PJCFORM.Add(newPjcform);
                    db.SaveChanges();
                    ViewBag.Message = "新增成功";
                }

                // 處理檔案上傳 - 先檢查是否有實際檔案
                var files = Request.Files;
                bool hasFiles = false;
                for (int i = 0; i < files.Count; i++) {
                    if (files[i] != null && files[i].ContentLength > 0) {
                        hasFiles = true;
                        break;
                    }
                }

                if (hasFiles) {
                    int uploadCount = 0;
                    DateTime now = DateTime.Now;
                    int rocYear = now.Year - 1911;
                    string month = now.Month.ToString();
                    string day = now.Day.ToString();

                    string basePath = GetBaseFilePath();
                    string relativePath = $@"{rocYear}\{month}\{day}\{pjId}";
                    string fullPath = System.IO.Path.Combine(basePath, relativePath);

                    if(!System.IO.Directory.Exists(fullPath)) {
                        System.IO.Directory.CreateDirectory(fullPath);
                    }

                    for(int i = 0; i < files.Count; i++) {
                        HttpPostedFileBase file = files[i];

                        if(file != null && file.ContentLength > 0) {
                            string fileName = System.IO.Path.GetFileName(file.FileName);
                            string filePath = System.IO.Path.Combine(fullPath, fileName);

                            file.SaveAs(filePath);

                            decimal fileSizeKB = (decimal)file.ContentLength / 1024;
                            decimal fileId = decimal.Parse(DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                            string filePathValue = System.IO.Path.Combine(relativePath, fileName);

                            SqlCommand cmd = new SqlCommand(
                             "INSERT INTO PJCFILE (FILE_ID, PJ_ID, FILE_LINK, FILE_PATH, FILE_NAME, FILE_SIZE, ADD_DT, UPD_DT, UPD_UID) " +
                            "VALUES (@FILE_ID, @PJ_ID, @FILE_LINK, @FILE_PATH, @FILE_NAME, @FILE_SIZE, @ADD_DT, @UPD_DT, @UPD_UID)"
                            );
                            cmd.Parameters.Add(new SqlParameter("FILE_ID", fileId));
                            cmd.Parameters.Add(new SqlParameter("PJ_ID", pjId));
                            cmd.Parameters.Add(new SqlParameter("FILE_LINK", filePathValue));
                            cmd.Parameters.Add(new SqlParameter("FILE_PATH", filePathValue));
                            cmd.Parameters.Add(new SqlParameter("FILE_NAME", fileName));
                            cmd.Parameters.Add(new SqlParameter("FILE_SIZE", fileSizeKB));
                            cmd.Parameters.Add(new SqlParameter("ADD_DT", now));
                            cmd.Parameters.Add(new SqlParameter("UPD_DT", now));
                            cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId()));

                            BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);
                            uploadCount++;
                        }
                    }

                    if(uploadCount > 0) {
                        ViewBag.Message += $", 成功上傳 {uploadCount} 個檔案";
                    }
                }

                // 重新載入頁面
                return SYSInfo(pjId);
            } catch(Exception ex) {
                // 顯示詳細錯誤訊息，包含 InnerException
                string errorMsg = ex.Message;
                if (ex.InnerException != null) {
                    errorMsg += " | Inner: " + ex.InnerException.Message;
                    if (ex.InnerException.InnerException != null) {
                        errorMsg += " | Inner2: " + ex.InnerException.InnerException.Message;
                    }
                }
                ViewBag.Message = "操作失敗：" + errorMsg;
                return SYSInfo(model.PJ_ID);
            }
        }

        /// <summary>
        /// 刪除規劃報告
        /// </summary>
        [HttpPost]
        [MultiButton(MatchFormKey = "action", MatchFormValue = "Del")]
        public ActionResult DeleteSYSInfo(SYS0030Model m) {
            try {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();

                var pjcform = db.PJCFORM.FirstOrDefault(x => x.PJ_ID == m.PJ_ID);
                if(pjcform != null) {
                    db.PJCFORM.Remove(pjcform);
                    db.SaveChanges();
                    //return Json(new { success = true, message = "刪除成功" });
                    ViewBag.msg = "刪除成功";
                    return Search(m);
                } else {
                    //return Json(new { success = false, message = "找不到該記錄" });
                    ViewBag.msg = "找不到該記錄";
                    return Search(m);
                }
            } catch(Exception ex) {
                //return Json(new { success = false, message = "刪除失敗：" + ex.Message });
                ViewBag.msg = "刪除失敗";
                return Search(m);
            }
        }

        /// <summary>
        /// 取得檔案列表
        /// </summary>
        [HttpGet]
        public ActionResult GetFileList(string pjId) {
            try {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();
                var dbFiles = db.PJCFILE
                     .Where(x => x.PJ_ID == pjId)
                     .OrderByDescending(x => x.ADD_DT)
                     .ToList();

                var files = dbFiles.Select(x => new {
                    FILE_ID = x.FILE_ID.ToString(),
                    x.FILE_NAME,
                    x.FILE_SIZE,
                    x.ADD_DT
                }).ToList();

                return Json(new { success = true, files = files }, JsonRequestBehavior.AllowGet);
            } catch(Exception ex) {
                return Json(new { success = false, message = "取得檔案列表失敗：" + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 多檔案上傳
        /// </summary>
        [HttpPost]
        public ActionResult UploadFiles(string pjId) {
            try {
                if(string.IsNullOrEmpty(pjId)) {
                    return Json(new { success = false, message = "PJ_ID 不可為空" });
                }

                var files = Request.Files;
                if(files.Count == 0) {
                    return Json(new { success = false, message = "請選擇要上傳的檔案" });
                }

                var uploadedFiles = new List<object>();

                // 建立路徑: C:\backup\SEPMIS\CMETMS\Files\PJCFILE\民國年\月\日\PJ_ID\
                DateTime now = DateTime.Now;
                int rocYear = now.Year - 1911; // 轉換為民國年
                string month = now.Month.ToString();
                string day = now.Day.ToString();

                string basePath = GetBaseFilePath();
                string relativePath = $@"{rocYear}\{month}\{day}\{pjId}";
                string fullPath = System.IO.Path.Combine(basePath, relativePath);

                // 確保目錄存在
                if(!System.IO.Directory.Exists(fullPath)) {
                    System.IO.Directory.CreateDirectory(fullPath);
                }

                for(int i = 0; i < files.Count; i++) {
                    HttpPostedFileBase file = files[i];

                    if(file != null && file.ContentLength > 0) {
                        // 取得檔案名稱
                        string fileName = System.IO.Path.GetFileName(file.FileName);
                        string filePath = System.IO.Path.Combine(fullPath, fileName);

                        // 儲存檔案
                        file.SaveAs(filePath);

                        // 計算檔案大小 (KB)
                        decimal fileSizeKB = (decimal)file.ContentLength / 1024;

                        // 產生 FILE_ID - 使用時間戳記
                        decimal fileId = decimal.Parse(DateTime.Now.ToString("yyyyMMddHHmmssfff"));

                        // FILE_LINK 和 FILE_PATH 存相同的值
                        string filePathValue = System.IO.Path.Combine(relativePath, fileName);

                        // 使用 ADO.NET 插入資料
                        SqlCommand cmd = new SqlCommand(
                             "INSERT INTO PJCFILE (FILE_ID, PJ_ID, FILE_LINK, FILE_PATH, FILE_NAME, FILE_SIZE, ADD_DT, UPD_DT, UPD_UID) " +
                         "VALUES (@FILE_ID, @PJ_ID, @FILE_LINK, @FILE_PATH, @FILE_NAME, @FILE_SIZE, @ADD_DT, @UPD_DT, @UPD_UID)"
                  );
                        cmd.Parameters.Add(new SqlParameter("FILE_ID", fileId));
                        cmd.Parameters.Add(new SqlParameter("PJ_ID", pjId));
                        cmd.Parameters.Add(new SqlParameter("FILE_LINK", filePathValue));
                        cmd.Parameters.Add(new SqlParameter("FILE_PATH", filePathValue));
                        cmd.Parameters.Add(new SqlParameter("FILE_NAME", fileName));
                        cmd.Parameters.Add(new SqlParameter("FILE_SIZE", fileSizeKB));
                        cmd.Parameters.Add(new SqlParameter("ADD_DT", now));
                        cmd.Parameters.Add(new SqlParameter("UPD_DT", now));
                        cmd.Parameters.Add(new SqlParameter("UPD_UID", WebSiteHelper.getUserId()));

                        BASE_DAO.SqlDao.CMETMSExecuteSql(cmd);

                        uploadedFiles.Add(new {
                            fileId = fileId.ToString(),
                            fileName = fileName,
                            fileSize = fileSizeKB,
                            addDate = now
                        });
                    }
                }

                return Json(new {
                    success = true,
                    message = $"成功上傳 {uploadedFiles.Count} 個檔案",
                    files = uploadedFiles
                });
            } catch(Exception ex) {
                return Json(new { success = false, message = "上傳失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// 刪除檔案
        /// </summary>
        [HttpPost]
        public ActionResult DeleteFile(string fileId) {
            try {
                //decimal id = decimal.Parse(fileId);

                // Use direct SQL to find file path to avoid EF caching/mapping issues
                string sql = "SELECT FILE_PATH FROM PJCFILE WHERE FILE_ID = @FILE_ID";
                SqlCommand cmdSelect = new SqlCommand(sql);
                cmdSelect.Parameters.Add(new SqlParameter("FILE_ID", fileId));

                DataTable dt = BASE_DAO.SqlDao.CMETMSQuery4List(cmdSelect);

                if (dt.Rows.Count > 0) {
                    string filePath = dt.Rows[0]["FILE_PATH"].ToString();
                    string basePath = GetBaseFilePath();
                    string fullPath = System.IO.Path.Combine(basePath, filePath);

                    if (System.IO.File.Exists(fullPath)) {
                        System.IO.File.Delete(fullPath);
                    }

                    // Use direct SQL to delete record
                    SqlCommand cmdDelete = new SqlCommand("DELETE FROM PJCFILE WHERE FILE_ID = @FILE_ID");
                    cmdDelete.Parameters.Add(new SqlParameter("FILE_ID", fileId));
                    BASE_DAO.SqlDao.CMETMSExecuteSql(cmdDelete);

                    return Json(new { success = true, message = "刪除成功" });
                } else {
                    // If record not found, treat as success per user request
                    return Json(new { success = true, message = "找不到該檔案記錄(已視為刪除)" });
                }
            } catch (Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// 下載檔案
        /// </summary>
        [HttpGet]
        public ActionResult DownloadFile(string fileId) {
            try {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();
                decimal id = decimal.Parse(fileId);
                var file = db.PJCFILE.FirstOrDefault(x => x.FILE_ID == id);

                if(file != null) {
                    string basePath = GetBaseFilePath();
                    string fullPath = System.IO.Path.Combine(basePath, file.FILE_PATH);

                    if(System.IO.File.Exists(fullPath)) {
                        var contentType = MimeMapping.GetMimeMapping(file.FILE_NAME);
                        return File(fullPath, contentType, file.FILE_NAME);
                    } else {
                        return HttpNotFound("檔案不存在");
                    }
                } else {
                    return HttpNotFound("找不到該檔案記錄");
                }
            } catch(Exception ex) {
                return Content("下載失敗：" + ex.Message);
            }
        }

        private string GetBaseFilePath() {
            string uploadPath = WebConfigurationManager.AppSettings["uploadPath"];
            if (string.IsNullOrEmpty(uploadPath)) {
                uploadPath = "File/Doc/";
            }
            if (!uploadPath.StartsWith("~") && !System.IO.Path.IsPathRooted(uploadPath)) {
                uploadPath = "~/" + uploadPath;
            }
            return Server.MapPath(uploadPath);
        }

        /// <summary>
        /// 刪除規劃報告/實施計畫
        /// </summary>
        /// <param name="id"> 編號 </param>
        /// <param name="check"> 是否檢查關聯 </param>
        /// <returns></returns>
        [HttpPost]
        //[ValidateAntiForgeryToken]
        // 相容現行前端：AJAX 送出的參數名為 pjId，未提供時回退使用 id
        public ActionResult DeleteSYSInfo(string id, bool check = false, string pjId = null)
        {
            id = string.IsNullOrEmpty(id) ? pjId : id;

            try
            {
                var db = new WCals.Models.CMETMSModels.CMETMSEntities();

                // (刪除)取得是否有使用該實施計畫的工程登記冊資料
                if (check)
                {
                    var usedTDMs = db.TDMFORM.Where(a => a.PJ_ID == id).Select(a => a.TD_CODE).ToList();

                    if (usedTDMs.Any())
                    {
                        bool isUnused = !usedTDMs.Any();

                        return Json(new
                        {
                            success = false,
                            isUnused,
                            usedIn = new
                            {
                                tdms = usedTDMs
                            }
                        });
                    }
                }

                var pjcform = db.PJCFORM.SingleOrDefault(x => x.PJ_ID == id);

                if (pjcform != null)
                {
                    db.PJCFORM.Remove(pjcform);
                    db.SaveChanges();
                    return Json(new { success = true, message = "刪除成功" });
                }
                else
                {
                    return Json(new { success = false, message = "找不到該記錄" });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

    }
}
