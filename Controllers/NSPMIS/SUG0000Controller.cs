using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using WCals.classes.john;
using WCals.Helpers;
using WCals.ViewModels.NewModels;
using static WCals.classes.john.BASE_DAO;

namespace WCals.Controllers.NSPMIS {
    /// <summary>
    /// SUG0000 建議案管理 Controller (無單位分派)
    /// 頁籤1: 列管中 (SUG_CLOSE=0)
    /// 頁籤2: 已結案 (SUG_CLOSE=1)
    /// </summary>
    public class SUG0000Controller : Controller {

        private readonly int pageSize = 20;

        #region 主頁面與頁籤

        // GET: SUG0000
        public ActionResult Index(SUG0010Model model) {
            if (model == null) {
                model = new SUG0010Model();
            }

            // 預設頁籤為1
            if (model.tab == 0) {
                model.tab = 1;
            }

            return View("SUG0000Index", model);
        }

        /// <summary>
        /// Tab 頁面 (供 iframe 載入)
        /// </summary>
        public ActionResult Tab(int tab = 1, string INSTRUCT_NAME = null, string SUG_NAME = null,
            string SUG_CONTENT = null, string START_DATE = null, string END_DATE = null,
            int page1 = 1, int page2 = 1) {

            SUG0010Model model = new SUG0010Model {
                tab = tab,
                INSTRUCT_NAME = INSTRUCT_NAME,
                SUG_NAME = SUG_NAME,
                SUG_CONTENT = SUG_CONTENT,
                START_DATE = START_DATE,
                END_DATE = END_DATE,
                page1 = page1,
                page2 = page2
            };

            // 根據當前頁籤載入對應資料
            switch (tab) {
                case 1:
                    LoadTab1Data(model);
                    break;
                case 2:
                    LoadTab2Data(model);
                    break;
            }

            return View("SUG0000Tab", model);
        }

        /// <summary>
        /// 頁籤1: 列管中 (SUG_CLOSE=0)
        /// </summary>
        private void LoadTab1Data(SUG0010Model model) {
            if (model.page1 == 0) model.page1 = 1;

            StringBuilder sb = new StringBuilder();
            sb.Append(@"SELECT M.SUG_ID,
                M.INSTRUCT_DATE,
                M.INSTRUCT_NAME,
                M.SUG_DATE,
                M.SUG_JOB, M.SUG_NAME, M.SUG_CONTENT,
                (SELECT COUNT(*) FROM SUGGEST.dbo.SUGFILE_H F WHERE F.SUG_ID = CONVERT(VARCHAR, M.SUG_ID)) AS FILE_COUNT
                FROM SUGGEST.dbo.SUGMASTER M
                WHERE M.SUG_CLOSE = 0 ");

            SqlCommand cmd = new SqlCommand();

            // 查詢條件
            if (!string.IsNullOrEmpty(model.INSTRUCT_NAME)) {
                sb.Append("AND M.INSTRUCT_NAME LIKE @INSTRUCT_NAME ");
                cmd.Parameters.Add(new SqlParameter("INSTRUCT_NAME", "%" + model.INSTRUCT_NAME + "%"));
            }
            if (!string.IsNullOrEmpty(model.SUG_NAME)) {
                sb.Append("AND M.SUG_NAME LIKE @SUG_NAME ");
                cmd.Parameters.Add(new SqlParameter("SUG_NAME", "%" + model.SUG_NAME + "%"));
            }
            if (!string.IsNullOrEmpty(model.SUG_CONTENT)) {
                sb.Append("AND M.SUG_CONTENT LIKE @SUG_CONTENT ");
                cmd.Parameters.Add(new SqlParameter("SUG_CONTENT", "%" + model.SUG_CONTENT + "%"));
            }

            sb.Append("ORDER BY M.SUG_ID DESC");
            cmd.CommandText = sb.ToString();

            DataTable dt = SqlDao.CMBASEQuery4List(cmd);
            List<DataRow> list = dt.AsEnumerable().ToList();
            model.List_Tab1 = new PagedList<DataRow>(list, model.page1, pageSize);
        }

        /// <summary>
        /// 頁籤2: 已結案 (SUG_CLOSE=1)
        /// </summary>
        private void LoadTab2Data(SUG0010Model model) {
            if (model.page2 == 0) model.page2 = 1;

            StringBuilder sb = new StringBuilder();
            sb.Append(@"SELECT SUG_ID,
                INSTRUCT_DATE,
                INSTRUCT_NAME,
                SUG_DATE,
                SUG_JOB, SUG_NAME, SUG_CONTENT, FINAL_CONTENT,
                AGREE, KEYWORDS
                FROM SUGGEST.dbo.SUGMASTER
                WHERE SUG_CLOSE = 1 ");

            SqlCommand cmd = new SqlCommand();

            // 查詢條件
            if (!string.IsNullOrEmpty(model.INSTRUCT_NAME)) {
                sb.Append("AND INSTRUCT_NAME LIKE @INSTRUCT_NAME ");
                cmd.Parameters.Add(new SqlParameter("INSTRUCT_NAME", "%" + model.INSTRUCT_NAME + "%"));
            }
            if (!string.IsNullOrEmpty(model.SUG_NAME)) {
                sb.Append("AND SUG_NAME LIKE @SUG_NAME ");
                cmd.Parameters.Add(new SqlParameter("SUG_NAME", "%" + model.SUG_NAME + "%"));
            }
            if (!string.IsNullOrEmpty(model.SUG_CONTENT)) {
                sb.Append("AND (SUG_CONTENT LIKE @SUG_CONTENT OR FINAL_CONTENT LIKE @SUG_CONTENT) ");
                cmd.Parameters.Add(new SqlParameter("SUG_CONTENT", "%" + model.SUG_CONTENT + "%"));
            }
            if (!string.IsNullOrEmpty(model.START_DATE)) {
                sb.Append("AND INSTRUCT_DATE >= @START_DATE ");
                cmd.Parameters.Add(new SqlParameter("START_DATE", model.START_DATE));
            }
            if (!string.IsNullOrEmpty(model.END_DATE)) {
                sb.Append("AND INSTRUCT_DATE <= @END_DATE ");
                cmd.Parameters.Add(new SqlParameter("END_DATE", model.END_DATE + " 23:59:59"));
            }

            sb.Append("ORDER BY SUG_ID DESC");
            cmd.CommandText = sb.ToString();

            DataTable dt = SqlDao.CMBASEQuery4List(cmd);
            List<DataRow> list = dt.AsEnumerable().ToList();
            model.List_Tab2 = new PagedList<DataRow>(list, model.page2, pageSize);
        }

        #endregion

        #region 新增/編輯建議案

        /// <summary>
        /// 新增/編輯頁面
        /// </summary>
        public ActionResult Edit(decimal? SUG_ID = null) {
            SUG0010Model model = new SUG0010Model();

            if (SUG_ID.HasValue && SUG_ID.Value > 0) {
                // 編輯模式 - 載入資料
                SqlCommand cmd = new SqlCommand(@"SELECT * FROM SUGGEST.dbo.SUGMASTER WHERE SUG_ID = @SUG_ID");
                cmd.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID.Value));
                DataTable dt = SqlDao.CMBASEQuery4List(cmd);

                if (dt.Rows.Count > 0) {
                    DataRow row = dt.Rows[0];
                    model.SUG_ID = Convert.ToDecimal(row["SUG_ID"]);
                    model.INSTRUCT_DATE = row["INSTRUCT_DATE"] != DBNull.Value ? (DateTime?)row["INSTRUCT_DATE"] : null;
                    model.INSTRUCT_NAME = row["INSTRUCT_NAME"] != DBNull.Value ? row["INSTRUCT_NAME"].ToString() : "";
                    model.SUG_DATE = row["SUG_DATE"] != DBNull.Value ? (DateTime?)row["SUG_DATE"] : null;
                    model.SUG_JOB = row["SUG_JOB"] != DBNull.Value ? row["SUG_JOB"].ToString() : "";
                    model.SUG_NAME = row["SUG_NAME"] != DBNull.Value ? row["SUG_NAME"].ToString() : "";
                    model.SUG_CONTENT = row["SUG_CONTENT"] != DBNull.Value ? row["SUG_CONTENT"].ToString() : "";
                    model.KEYWORDS = row["KEYWORDS"] != DBNull.Value ? row["KEYWORDS"].ToString() : "";
                }

                // 載入附件
                SqlCommand cmdFiles = new SqlCommand(@"SELECT * FROM SUGGEST.dbo.SUGFILE_H WHERE SUG_ID = @SUG_ID");
                cmdFiles.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID.Value.ToString()));
                model.FileList = SqlDao.CMBASEQuery4List(cmdFiles);
            }

            return View("SUG0000Edit", model);
        }

        /// <summary>
        /// 儲存建議案
        /// </summary>
        [HttpPost]
        public JsonResult Save(SUG0010Model model) {
            try {
                // TODO: 權限控制
                string userId = WebSiteHelper.getUserId() ?? "";
                string userDept = WebSiteHelper.getUserDept() ?? "";
                DateTime now = DateTime.Now;

                if (model.SUG_ID > 0) {
                    // 更新
                    SqlCommand cmd = new SqlCommand(@"UPDATE SUGGEST.dbo.SUGMASTER SET
                        INSTRUCT_DATE = @INSTRUCT_DATE,
                        INSTRUCT_NAME = @INSTRUCT_NAME,
                        SUG_DATE = @SUG_DATE,
                        SUG_JOB = @SUG_JOB,
                        SUG_NAME = @SUG_NAME,
                        SUG_CONTENT = @SUG_CONTENT,
                        KEYWORDS = @KEYWORDS,
                        UPD_DT = @UPD_DT,
                        UPD_UID = @UPD_UID
                        WHERE SUG_ID = @SUG_ID");
                    cmd.Parameters.Add(new SqlParameter("SUG_ID", model.SUG_ID));
                    cmd.Parameters.Add(new SqlParameter("INSTRUCT_DATE", model.INSTRUCT_DATE.HasValue ? (object)model.INSTRUCT_DATE.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("INSTRUCT_NAME", string.IsNullOrEmpty(model.INSTRUCT_NAME) ? (object)DBNull.Value : model.INSTRUCT_NAME));
                    cmd.Parameters.Add(new SqlParameter("SUG_DATE", model.SUG_DATE.HasValue ? (object)model.SUG_DATE.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("SUG_JOB", string.IsNullOrEmpty(model.SUG_JOB) ? (object)DBNull.Value : model.SUG_JOB));
                    cmd.Parameters.Add(new SqlParameter("SUG_NAME", string.IsNullOrEmpty(model.SUG_NAME) ? (object)DBNull.Value : model.SUG_NAME));
                    cmd.Parameters.Add(new SqlParameter("SUG_CONTENT", string.IsNullOrEmpty(model.SUG_CONTENT) ? (object)DBNull.Value : model.SUG_CONTENT));
                    cmd.Parameters.Add(new SqlParameter("KEYWORDS", string.IsNullOrEmpty(model.KEYWORDS) ? (object)DBNull.Value : model.KEYWORDS));
                    cmd.Parameters.Add(new SqlParameter("UPD_DT", now));
                    cmd.Parameters.Add(new SqlParameter("UPD_UID", userId));
                    SqlDao.CMBASEExecuteSql(cmd);

                    return Json(new { success = true, message = "更新成功", SUG_ID = model.SUG_ID });
                } else {
                    // 新增
                    SqlCommand cmd = new SqlCommand(@"INSERT INTO SUGGEST.dbo.SUGMASTER
                        (INSTRUCT_DATE, INSTRUCT_NAME, SUG_DATE, SUG_JOB, SUG_NAME, SUG_CONTENT,
                         KEYWORDS, SUG_CLOSE, CREATE_DEPT, ADD_DT, UPD_UID)
                        VALUES (@INSTRUCT_DATE, @INSTRUCT_NAME, @SUG_DATE, @SUG_JOB, @SUG_NAME, @SUG_CONTENT,
                         @KEYWORDS, 0, @CREATE_DEPT, @ADD_DT, @UPD_UID);
                        SELECT SCOPE_IDENTITY();");
                    cmd.Parameters.Add(new SqlParameter("INSTRUCT_DATE", model.INSTRUCT_DATE.HasValue ? (object)model.INSTRUCT_DATE.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("INSTRUCT_NAME", string.IsNullOrEmpty(model.INSTRUCT_NAME) ? (object)DBNull.Value : model.INSTRUCT_NAME));
                    cmd.Parameters.Add(new SqlParameter("SUG_DATE", model.SUG_DATE.HasValue ? (object)model.SUG_DATE.Value : DBNull.Value));
                    cmd.Parameters.Add(new SqlParameter("SUG_JOB", string.IsNullOrEmpty(model.SUG_JOB) ? (object)DBNull.Value : model.SUG_JOB));
                    cmd.Parameters.Add(new SqlParameter("SUG_NAME", string.IsNullOrEmpty(model.SUG_NAME) ? (object)DBNull.Value : model.SUG_NAME));
                    cmd.Parameters.Add(new SqlParameter("SUG_CONTENT", string.IsNullOrEmpty(model.SUG_CONTENT) ? (object)DBNull.Value : model.SUG_CONTENT));
                    cmd.Parameters.Add(new SqlParameter("KEYWORDS", string.IsNullOrEmpty(model.KEYWORDS) ? (object)DBNull.Value : model.KEYWORDS));
                    cmd.Parameters.Add(new SqlParameter("CREATE_DEPT", userDept));
                    cmd.Parameters.Add(new SqlParameter("ADD_DT", now));
                    cmd.Parameters.Add(new SqlParameter("UPD_UID", userId));

                    DataTable dt = SqlDao.CMBASEQuery4List(cmd);
                    decimal newId = dt.Rows.Count > 0 ? Convert.ToDecimal(dt.Rows[0][0]) : 0;

                    return Json(new { success = true, message = "新增成功", SUG_ID = newId });
                }
            } catch (Exception ex) {
                return Json(new { success = false, message = "儲存失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// 刪除建議案
        /// </summary>
        [HttpPost]
        public JsonResult Delete(decimal SUG_ID) {
            try {
                // TODO: 權限控制

                // 刪除附件
                SqlCommand cmdFiles = new SqlCommand("DELETE FROM SUGGEST.dbo.SUGFILE_H WHERE SUG_ID = @SUG_ID");
                cmdFiles.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID.ToString()));
                SqlDao.CMBASEExecuteSql(cmdFiles);

                // 刪除主檔
                SqlCommand cmd = new SqlCommand("DELETE FROM SUGGEST.dbo.SUGMASTER WHERE SUG_ID = @SUG_ID");
                cmd.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID));
                SqlDao.CMBASEExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });
            } catch (Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        #endregion

        #region 結案功能

        /// <summary>
        /// 結案頁面
        /// </summary>
        public ActionResult Close(decimal SUG_ID) {
            SUG0010Model model = new SUG0010Model {
                SUG_ID = SUG_ID
            };

            // 載入建議案資料
            SqlCommand cmd = new SqlCommand(@"SELECT * FROM SUGGEST.dbo.SUGMASTER WHERE SUG_ID = @SUG_ID");
            cmd.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID));
            DataTable dt = SqlDao.CMBASEQuery4List(cmd);
            if (dt.Rows.Count > 0) {
                DataRow row = dt.Rows[0];
                model.SUG_CONTENT = row["SUG_CONTENT"] != DBNull.Value ? row["SUG_CONTENT"].ToString() : "";
                model.FINAL_CONTENT = row["FINAL_CONTENT"] != DBNull.Value ? row["FINAL_CONTENT"].ToString() : "";
                model.AGREE = row["AGREE"] != DBNull.Value ? (bool?)row["AGREE"] : null;
                model.KEYWORDS = row["KEYWORDS"] != DBNull.Value ? row["KEYWORDS"].ToString() : "";
            }

            return View("SUG0000Close", model);
        }

        /// <summary>
        /// 執行結案
        /// </summary>
        [HttpPost]
        public JsonResult DoClose(decimal SUG_ID, string FINAL_CONTENT, bool? AGREE, string KEYWORDS) {
            try {
                // TODO: 權限控制
                string userId = WebSiteHelper.getUserId() ?? "";
                DateTime now = DateTime.Now;

                SqlCommand cmd = new SqlCommand(@"UPDATE SUGGEST.dbo.SUGMASTER SET
                    FINAL_CONTENT = @FINAL_CONTENT,
                    AGREE = @AGREE,
                    KEYWORDS = @KEYWORDS,
                    SUG_CLOSE = 1,
                    UPD_DT = @UPD_DT,
                    UPD_UID = @UPD_UID
                    WHERE SUG_ID = @SUG_ID");
                cmd.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID));
                cmd.Parameters.Add(new SqlParameter("FINAL_CONTENT", string.IsNullOrEmpty(FINAL_CONTENT) ? (object)DBNull.Value : FINAL_CONTENT));
                cmd.Parameters.Add(new SqlParameter("AGREE", AGREE.HasValue ? (object)AGREE.Value : DBNull.Value));
                cmd.Parameters.Add(new SqlParameter("KEYWORDS", string.IsNullOrEmpty(KEYWORDS) ? (object)DBNull.Value : KEYWORDS));
                cmd.Parameters.Add(new SqlParameter("UPD_DT", now));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId));
                SqlDao.CMBASEExecuteSql(cmd);

                return Json(new { success = true, message = "結案成功" });
            } catch (Exception ex) {
                return Json(new { success = false, message = "結案失敗：" + ex.Message });
            }
        }

        #endregion

        #region 附件管理

        /// <summary>
        /// 上傳附件
        /// </summary>
        [HttpPost]
        public JsonResult UploadFile(decimal SUG_ID, HttpPostedFileBase file) {
            try {
                if (file == null || file.ContentLength == 0) {
                    return Json(new { success = false, message = "請選擇檔案" });
                }

                // TODO: 權限控制
                string userId = WebSiteHelper.getUserId() ?? "";
                DateTime now = DateTime.Now;

                // 取得新的 F_ID
                SqlCommand cmdMaxId = new SqlCommand(@"SELECT ISNULL(MAX(F_ID), 0) + 1
                    FROM SUGGEST.dbo.SUGFILE_H WHERE SUG_ID = @SUG_ID");
                cmdMaxId.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID.ToString()));
                DataTable dtMaxId = SqlDao.CMBASEQuery4List(cmdMaxId);
                decimal newFId = dtMaxId.Rows.Count > 0 ? Convert.ToDecimal(dtMaxId.Rows[0][0]) : 1;

                // 讀取檔案內容
                byte[] fileBytes;
                using (BinaryReader br = new BinaryReader(file.InputStream)) {
                    fileBytes = br.ReadBytes(file.ContentLength);
                }

                // 新增檔案記錄
                SqlCommand cmd = new SqlCommand(@"INSERT INTO SUGGEST.dbo.SUGFILE_H
                    (SUG_ID, F_ID, F_PATH, F_FILE, ADD_DT, UPD_UID)
                    VALUES (@SUG_ID, @F_ID, @F_PATH, @F_FILE, @ADD_DT, @UPD_UID)");
                cmd.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID.ToString()));
                cmd.Parameters.Add(new SqlParameter("F_ID", newFId));
                cmd.Parameters.Add(new SqlParameter("F_PATH", file.FileName));
                cmd.Parameters.Add(new SqlParameter("F_FILE", fileBytes));
                cmd.Parameters.Add(new SqlParameter("ADD_DT", now));
                cmd.Parameters.Add(new SqlParameter("UPD_UID", userId));
                SqlDao.CMBASEExecuteSql(cmd);

                return Json(new { success = true, message = "上傳成功", F_ID = newFId });
            } catch (Exception ex) {
                return Json(new { success = false, message = "上傳失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// 下載附件
        /// </summary>
        public ActionResult DownloadFile(string SUG_ID, decimal F_ID) {
            try {
                SqlCommand cmd = new SqlCommand(@"SELECT F_PATH, F_FILE FROM SUGGEST.dbo.SUGFILE_H
                    WHERE SUG_ID = @SUG_ID AND F_ID = @F_ID");
                cmd.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID));
                cmd.Parameters.Add(new SqlParameter("F_ID", F_ID));
                DataTable dt = SqlDao.CMBASEQuery4List(cmd);

                if (dt.Rows.Count > 0) {
                    string fileName = dt.Rows[0]["F_PATH"].ToString();
                    byte[] fileBytes = (byte[])dt.Rows[0]["F_FILE"];
                    return File(fileBytes, "application/octet-stream", fileName);
                }

                return Content("檔案不存在");
            } catch (Exception ex) {
                return Content("下載失敗：" + ex.Message);
            }
        }

        /// <summary>
        /// 刪除附件
        /// </summary>
        [HttpPost]
        public JsonResult DeleteFile(string SUG_ID, decimal F_ID) {
            try {
                // TODO: 權限控制

                SqlCommand cmd = new SqlCommand(@"DELETE FROM SUGGEST.dbo.SUGFILE_H
                    WHERE SUG_ID = @SUG_ID AND F_ID = @F_ID");
                cmd.Parameters.Add(new SqlParameter("SUG_ID", SUG_ID));
                cmd.Parameters.Add(new SqlParameter("F_ID", F_ID));
                SqlDao.CMBASEExecuteSql(cmd);

                return Json(new { success = true, message = "刪除成功" });
            } catch (Exception ex) {
                return Json(new { success = false, message = "刪除失敗：" + ex.Message });
            }
        }

        #endregion
    }
}
