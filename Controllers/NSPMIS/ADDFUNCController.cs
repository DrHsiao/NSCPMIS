using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity.Core.EntityClient;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Web;
using System.Web.Mvc;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEPMSModels;
using WCals.Models.CMETMSModels;

namespace WCals.Controllers.NSPMIS
{
    public class ADDFUNCController : Controller
    {
        CMBASEEntities db = new CMBASEEntities();
        CMEPMSEntities db2 = new CMEPMSEntities();
        CMETMSEntities db3 = new CMETMSEntities();

        // TODO： 放到正式環境時需更改連線字串為正式環境
        //string connectionString = @"Data Source=192.168.7.14;initial catalog=CMEPMS;user id=***;password=***;";
        //string connectionString = @"Data Source=10.1.1.112\SEPMIS;initial catalog=CMEPMS_from_prod_260401;user id=***;password=***;";

        private string connectionString
        {
            get
            {
                string efConnectionString = ConfigurationManager.ConnectionStrings["CMEPMSEntities"].ConnectionString;
                var entityBuilder = new EntityConnectionStringBuilder(efConnectionString);
                return entityBuilder.ProviderConnectionString;
            }
        }
        //string connectionString = @"Data Source=192.168.1.26\SEPMIS;initial catalog=CMEPMS;user id=***;password=***;";
        public ActionResult Rainfunc()
        {
            return View();
        }

        public ActionResult Sewfunc()
        {
            return View();
        }

        /// <summary>
        /// 驗證上傳之 Excel 是否符合 [實支數匯入] 格式規範
        /// 目的：
        /// 1. 防止錯誤格式進入預覽或資料庫
        /// 2. 提前回饋錯誤給使用者
        /// </summary>
        /// <param name="File">使用者上傳的 Excel 檔案</param>
        /// <returns>
        /// success = true  : 格式正確
        /// success = false : 格式錯誤 + 錯誤原因
        /// </returns>
        [HttpPost]
        public JsonResult ValidateExcelFormat(HttpPostedFileBase File)
        {
            Aspose.Cells.License License = new Aspose.Cells.License();
            License.SetLicense("Aspose.Total.lic");

            try
            {
                if (File == null || File.ContentLength == 0)
                {
                    return Json(new { success = false, message = "未選擇任何檔案" });
                }

                var workbook = new Aspose.Cells.Workbook(File.InputStream);

                if (workbook.Worksheets.Count != 2)
                {
                    return Json(new
                    {
                        success = false,
                        message = "Excel 僅能包含 2 個工作表 [國土署全、水工分署全]"
                    });
                }

                bool hasMain = false;   /// 國土署全
                bool hasBranch = false; /// 水工分署全

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    string sheetName = sheet.Name.Trim();

                    if (sheetName.Contains("國土署全"))
                        hasMain = true;

                    if (sheetName.Contains("水工分署全"))
                        hasBranch = true;

                    string a3 = sheet.Cells["A3"]?.StringValue?.Trim();

                    if (string.IsNullOrEmpty(a3) || !a3.Contains("截止年月:"))
                    {
                        return Json(new
                        {
                            success = false,
                            message = $"工作表【{sheetName}】A3 未包含 [截止年月:] 標示, 請依範例檔案修改."
                        });
                    }
                }

                if (!hasMain || !hasBranch)
                {
                    return Json(new
                    {
                        success = false,
                        message = "工作表名稱必須包含 : 一個 [國土署全】和一個 [水工分署全], 請依範例檔案修改."
                    });
                }

                return Json(new
                {
                    success = true,
                    message = "Excel 格式正確"
                });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    success = false,
                    message = "Excel 解析失敗：" + ex.Message
                });
            }
        }

        /// <summary>
        /// 解析 Excel 內容為標準 Dictionary List - 供「預覽」與「資料庫寫入」共用
        /// </summary>
        private List<Dictionary<string, string>> ParseExcelRows(Workbook workbook)
        {
            var result = new List<Dictionary<string, string>>();

            foreach (var worksheet in workbook.Worksheets)
            {
                var sheetName = worksheet.Name;
                var cells = worksheet.Cells;

                int startRow = 4;
                int endRow = cells.MaxDataRow;

                string SYR = "";
                string SMN = "";

                string text = cells["A3"]?.StringValue?.Trim();
                if (!string.IsNullOrEmpty(text) && text.Contains("截止年月"))
                {
                    string[] parts = text.Split(new[] { ':', '：' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2)
                    {
                        string dateStr = parts[1].Trim();
                        if (dateStr.Length >= 5)
                        {
                            SYR = dateStr.Substring(0, 3);
                            SMN = dateStr.Substring(3, 2);
                        }
                    }
                }

                for (int row = startRow; row <= endRow; row++)
                {
                    string RAM = cells[row, 1]?.StringValue ?? "";
                    string FDCODE = RAM.Replace(" ", "").Split('.')[0].Trim();
                    if (string.IsNullOrEmpty(FDCODE)) continue;

                    string TDCODE = cells[row, 2]?.StringValue?.Trim() ?? "";
                    string NTDCODE = TDCODE.Replace("-", "");

                    var rowData = new Dictionary<string, string>
                    {
                        ["年份"] = SYR,
                        ["月份"] = SMN,
                        ["隸屬部門"] = sheetName,
                        ["經費編號"] = FDCODE,
                        ["工程編號"] = TDCODE,
                        ["項目名稱"] = cells[row, 3]?.StringValue?.Trim() ?? "",
                        ["預算數"] = cells[row, 4]?.StringValue?.Trim() ?? "",
                        ["(實支數)本月發生"] = cells[row, 5]?.StringValue?.Trim() ?? "",
                        ["(實支數)累至本月"] = cells[row, 6]?.StringValue?.Trim() ?? "",
                        ["(暫付數)本月發生"] = cells[row, 7]?.StringValue?.Trim() ?? "",
                        ["(暫付數)累至本月"] = cells[row, 8]?.StringValue?.Trim() ?? "",
                        ["餘額"] = cells[row, 11]?.StringValue?.Trim() ?? "",
                        ["上傳時間"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    };

                    result.Add(rowData);
                }
            }

            return result;
        }

        /// <summary>
        /// Excel 資料預覽（不寫入資料庫）
        /// </summary>
        /// <param name="File"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult PreviewExcel(HttpPostedFileBase File)
        {
            try
            {
                var workbook = new Workbook(File.InputStream);
                var result = ParseExcelRows(workbook);

                return Json(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "解析失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// 將 Excel 資料寫入 TDDCTET(實支數) - 使用 MERGE 避免重複資料
        /// </summary>
        /// <param name="File"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult InsertTDDCTET(HttpPostedFileBase File)
        {
            if (File == null || File.ContentLength == 0)
            {
                return Json(new { success = false, message = "未選擇任何檔案" });
            }

            try
            {
                var workbook = new Workbook(File.InputStream);
                var rows = ParseExcelRows(workbook);

                int successCount = 0;
                int failCount = 0;

                List<string> failList = new List<string>();

                //if (EnvHelper.IsTestMode)
                //    connectionString = @"Data Source=10.1.1.112\SEPMIS;initial catalog=CMEPMS;user id=***;password=***;";
                //else
                //    connectionString = @"Data Source=192.168.7.14;initial catalog=CMEPMS;user id=***;password=***;";

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    foreach (var row in rows)
                    {
                        string sql = @"
                                        MERGE INTO TDDCTET AS T
                                        USING (SELECT 
                                            @SYR AS SYR,
                                            @SMN AS SMN,
                                            @FD_CODE AS FD_CODE,
                                            @TD_CODE AS TD_CODE
                                            ) AS S
                                        ON  T.SYR = S.SYR
                                        AND T.SMN = S.SMN
                                        AND T.FD_CODE = S.FD_CODE
                                        AND T.TD_CODE = S.TD_CODE
                                        WHEN MATCHED THEN
                                            UPDATE SET
                                                T.BAM = @BAM,
                                                T.AE_TM = @AE_TM,
                                                T.AE_CTM = @AE_CTM,
                                                T.PP_TM = @PP_TM,
                                                T.PP_CTM = @PP_CTM,
                                                T.BBE = @BBE,
                                                T.TD_NAME = @TD_NAME,
                                                T.ImportTime = GETDATE()
                                        WHEN NOT MATCHED THEN
                                            INSERT (SYR, SMN, DEPT, FD_CODE, TD_CODE, TD_NAME, BAM, AE_TM, AE_CTM, PP_TM, PP_CTM, BBE, ImportTime)
                                        VALUES 
                                            (@SYR, @SMN, @DEPT, @FD_CODE, @TD_CODE, @TD_NAME, @BAM, @AE_TM, @AE_CTM, @PP_TM, @PP_CTM, @BBE, GETDATE());";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@SYR", row["年份"]);
                            cmd.Parameters.AddWithValue("@SMN", row["月份"]);
                            cmd.Parameters.AddWithValue("@DEPT", row["隸屬部門"]);
                            cmd.Parameters.AddWithValue("@FD_CODE", row["經費編號"]);
                            cmd.Parameters.AddWithValue("@TD_CODE", row["工程編號"]);
                            cmd.Parameters.AddWithValue("@TD_NAME", row["項目名稱"]);

                            cmd.Parameters.AddWithValue("@BAM", ToDecimalSafe(row["預算數"]));
                            cmd.Parameters.AddWithValue("@AE_TM", ToDecimalSafe(row["(實支數)本月發生"]));
                            cmd.Parameters.AddWithValue("@AE_CTM", ToDecimalSafe(row["(實支數)累至本月"]));
                            cmd.Parameters.AddWithValue("@PP_TM", ToDecimalSafe(row["(暫付數)本月發生"]));
                            cmd.Parameters.AddWithValue("@PP_CTM", ToDecimalSafe(row["(暫付數)累至本月"]));
                            cmd.Parameters.AddWithValue("@BBE", ToDecimalSafe(row["餘額"]));

                            try
                            {
                                cmd.ExecuteNonQuery();
                                successCount++;
                            }
                            catch (Exception exRow)
                            {
                                failCount++;

                                // === Debug 印出完整資料 ===
                                Debug.WriteLine("========== 匯入失敗資料 ==========");
                                foreach (var kv in row)
                                {
                                    Debug.WriteLine($"{kv.Key}: {kv.Value}");
                                }
                                Debug.WriteLine($"錯誤訊息: {exRow.Message}");
                                Debug.WriteLine("====================================");

                                failList.Add(
                                    $"經費 {row["經費編號"]} / 工程 {row["工程編號"]} 匯入失敗：{exRow.Message}"
                                );
                            }
                        }
                    }
                }

                return Json(new
                {
                    success = true,
                    message = $"匯入完成：成功 {successCount} 筆，失敗 {failCount} 筆",
                    errors = failList
                });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "匯入失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// 將 Excel 資料寫入 TDDSTET(實支數) - 使用 MERGE 避免重複資料
        /// </summary>
        /// <param name="File"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult InsertTDDSTET(HttpPostedFileBase File)
        {
            if (File == null || File.ContentLength == 0)
            {
                return Json(new { success = false, message = "未選擇任何檔案" });
            }

            try
            {
                var workbook = new Workbook(File.InputStream);
                var rows = ParseExcelRows(workbook);

                int successCount = 0;
                int failCount = 0;

                List<string> failList = new List<string>();

                //if (EnvHelper.IsTestMode)
                //    connectionString = @"Data Source=10.1.1.112\SEPMIS;initial catalog=CMEPMS;user id=***;password=***;";
                //else
                //    connectionString = @"Data Source=192.168.7.14;initial catalog=CMEPMS;user id=***;password=***;";

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    foreach (var row in rows)
                    {
                        string sql = @"
                                        MERGE INTO TDDSTET AS T
                                        USING (SELECT 
                                            @SYR AS SYR,
                                            @SMN AS SMN,
                                            @FD_CODE AS FD_CODE,
                                            @TD_CODE AS TD_CODE
                                            ) AS S
                                        ON  T.SYR = S.SYR
                                        AND T.SMN = S.SMN
                                        AND T.FD_CODE = S.FD_CODE
                                        AND T.TD_CODE = S.TD_CODE
                                        WHEN MATCHED THEN
                                            UPDATE SET
                                                T.BAM = @BAM,
                                                T.AE_TM = @AE_TM,
                                                T.AE_CTM = @AE_CTM,
                                                T.PP_TM = @PP_TM,
                                                T.PP_CTM = @PP_CTM,
                                                T.BBE = @BBE,
                                                T.TD_NAME = @TD_NAME,
                                                T.ImportTime = GETDATE()
                                        WHEN NOT MATCHED THEN
                                            INSERT (SYR, SMN, DEPT, FD_CODE, TD_CODE, TD_NAME, BAM, AE_TM, AE_CTM, PP_TM, PP_CTM, BBE, ImportTime)
                                        VALUES 
                                            (@SYR, @SMN, @DEPT, @FD_CODE, @TD_CODE, @TD_NAME, @BAM, @AE_TM, @AE_CTM, @PP_TM, @PP_CTM, @BBE, GETDATE());";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@SYR", row["年份"]);
                            cmd.Parameters.AddWithValue("@SMN", row["月份"]);
                            cmd.Parameters.AddWithValue("@DEPT", row["隸屬部門"]);
                            cmd.Parameters.AddWithValue("@FD_CODE", row["經費編號"]);
                            cmd.Parameters.AddWithValue("@TD_CODE", row["工程編號"]);
                            cmd.Parameters.AddWithValue("@TD_NAME", row["項目名稱"]);

                            cmd.Parameters.AddWithValue("@BAM", ToDecimalSafe(row["預算數"]));
                            cmd.Parameters.AddWithValue("@AE_TM", ToDecimalSafe(row["(實支數)本月發生"]));
                            cmd.Parameters.AddWithValue("@AE_CTM", ToDecimalSafe(row["(實支數)累至本月"]));
                            cmd.Parameters.AddWithValue("@PP_TM", ToDecimalSafe(row["(暫付數)本月發生"]));
                            cmd.Parameters.AddWithValue("@PP_CTM", ToDecimalSafe(row["(暫付數)累至本月"]));
                            cmd.Parameters.AddWithValue("@BBE", ToDecimalSafe(row["餘額"]));

                            try
                            {
                                cmd.ExecuteNonQuery();
                                successCount++;
                            }
                            catch (Exception exRow)
                            {
                                failCount++;

                                // === Debug 印出完整資料 ===
                                Debug.WriteLine("========== 匯入失敗資料 ==========");
                                foreach (var kv in row)
                                {
                                    Debug.WriteLine($"{kv.Key}: {kv.Value}");
                                }
                                Debug.WriteLine($"錯誤訊息: {exRow.Message}");
                                Debug.WriteLine("====================================");

                                failList.Add(
                                    $"經費 {row["經費編號"]} / 工程 {row["工程編號"]} 匯入失敗：{exRow.Message}"
                                );
                            }
                        }
                    }
                }

                return Json(new
                {
                    success = true,
                    message = $"匯入完成：成功 {successCount} 筆，失敗 {failCount} 筆",
                    errors = failList
                });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "匯入失敗：" + ex.Message });
            }
        }

        /// <summary>
        /// 字串轉 decimal（容錯）
        /// - 空值 → 0
        /// - 移除千分位
        /// - 非數字 → 0
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private decimal ToDecimalSafe(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return 0;

            input = input.Replace(",", "").Trim();

            decimal value;
            return decimal.TryParse(input, out value) ? value : 0;
        }

    }
}