using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity.Core.EntityClient;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using WCals.Models;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEPMSModels;
using WCals.Models.CMETMSModels;
using WCals.Source;

namespace WCals.Controllers
{
    public class RainReportController : System.Web.Mvc.Controller
    {
        CMBASEEntities db = new CMBASEEntities();
        CMEPMSEntities db2 = new CMEPMSEntities();
        CMETMSEntities db3 = new CMETMSEntities();

        /// 連線字串一律取自 EF (CMEPMSEntities) 的實際連線 —— 亦即 connections.config 中生效的那組，
        /// 與全站其他功能同源。本專案 PJE0030 / PJE0050 / SGE0000 / SGE0060 均採此寫法。
        ///
        /// 【勿改回硬編碼】原本此處寫死資料庫位址，導致部署到不同環境時
        /// RainReport 連到與 EF 不同的資料庫，各報表匯出因而發生伺服器錯誤。
        private string connectionString
        {
            get
            {
                string efConnectionString = ConfigurationManager.ConnectionStrings["CMEPMSEntities"].ConnectionString;
                var entityBuilder = new EntityConnectionStringBuilder(efConnectionString);
                return entityBuilder.ProviderConnectionString;
            }
        }        /// <summary>
                 /// I-1.雨水例行性管考報表-計畫名稱列表(List)
                 /// </summary>
                 /// <returns></returns>
        public ActionResult Index()
        {
            ViewBag.State = "Index";

            var pjNameList = new[]
            {
                "都市總合治水建設計畫(公務預算)",
                "因應氣候變遷縣市管河川及排水整體改善計畫(公務預算)"
            };

            ViewBag.Plan = db2.PJMFORM
                .Where(p => (p.ADV_UID == "A270400" || p.ADV_UID == "B150400")
                         && pjNameList.Contains(p.PJ_NAME))
                .OrderByDescending(p => p.SYR)
                .Select(p => new Report13ViewModel
                {
                    PID = p.PID,
                    AN_PID = 0,
                    SYR = p.SYR,
                    PJ_NAME = "(" + p.SYR + ((p.END_SYR == null || p.END_SYR == "") ? "" : "~" + p.END_SYR) + ")" + p.PJ_NAME
                }).ToList();

            return View();
        }

        /// <summary>
        /// I-1.雨水例行性管考報表-報表生成
        /// </summary>
        /// <param name="PID"> 案件編號 </param>
        /// <param name="Form"> 報表選擇 </param>
        /// <param name="SYR"> 年份 </param>
        /// <param name="SMN"> 月份</param>
        /// <param name="Button"> 報表輸出 </param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Index(decimal PID, string Form, string SYR, string SMN, string Button)
        {
            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            decimal AN_PID = PID;
            string DEPT = (string)Session["DeptId"];
            string PJ_NAME = db2.PJMFORM.Where(m => m.PID == PID).FirstOrDefault()?.PJ_NAME;

            switch (Form)
            {
                case "RPT0001":
                    PrintRPT0001(AN_PID, SYR, SMN, connectionString, DEPT, PJ_NAME, Button);
                    break;
                case "RPT0002":
                    PrintRPT0002_NEW(AN_PID, SYR, SMN, connectionString, DEPT, PJ_NAME, Button);
                    break;
                case "RPT0003":
                    PrintRPT0003_3(AN_PID, SYR, SMN, connectionString, DEPT, PJ_NAME, Button);
                    break;
            }

            ViewBag.Plan = (from p in db2.PJMFORM
                            where p.AN_PID == null && p.ADV_DP == "A270400"
                            orderby p.SYR descending
                            select new Report13ViewModel
                            {
                                PID = p.PID,
                                AN_PID = 0,
                                SYR = p.SYR,
                                PJ_NAME = "(" + p.SYR + (string.IsNullOrEmpty(p.END_SYR) ? "" : "~" + p.END_SYR) + ")" + p.PJ_NAME
                            }).ToList();

            return View(ViewBag.Plan);
        }

        /// <summary>
        /// I-1.雨水例行性管考報表[前瞻計畫執行情形表]-舊版(已無用)
        /// </summary>
        /// <param name="AN_PID"></param>
        /// <param name="SYR"> 年份 </param>
        /// <param name="SMN"> 月份 </param>
        /// <param name="connectionString"> SQL連接 </param>
        /// <param name="DEPT"> 部門 </param>
        /// <param name="PJ_NAME"> 案件名稱 </param>
        /// <param name="Button"> 報表生成 </param>
        private void PrintRPT0003(decimal AN_PID, string SYR, string SMN, string connectionString, string DEPT, string PJ_NAME, string Button)
        {
            string sFileName = "";
            string sFullFile = Server.MapPath("~/Templates/PJV0052.xlsx");
            FileInfo oFileInfo = new FileInfo(sFullFile);
            oFileInfo.IsReadOnly = false;

            Workbook workbook = new Workbook(sFullFile);
            Worksheet sheet1 = workbook.Worksheets["執行情形表"];

            int i;

            string sSYR = (DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (DateTime.Now.Month).ToString("00");
            string sPSYR = (DateTime.Now.Year - 1911 - 1).ToString("000");
            string sDAY = "31";
            DataTable dt = RainReport_Source.GetRPT0003Data(AN_PID, SYR, SMN, connectionString, DEPT);

            if (SYR != "")
            {
                sSYR = SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }
            if (SMN != "")
            {
                sSMN = SMN;
            }
            sDAY = Convert.ToDateTime((Convert.ToInt32(sSYR) + 1911).ToString() + "/" + sSMN + "/01").AddMonths(1).AddDays(-1).Day.ToString("00");

            for (i = 0; i < dt.Rows.Count; i++)
            {
                Range sourceRange = sheet1.Cells.CreateRange(4, 0, 1, 29);
                Range destinationRange = sheet1.Cells.CreateRange((i + 4), 0, 1, 29);
                destinationRange.Copy(sourceRange);
                sheet1.Cells.SetRowHeight((i + 4), sheet1.Cells.GetRowHeight(4));
            }

            sheet1.Cells["A1"].PutValue(PJ_NAME + "執行情形表");
            sheet1.Cells["A2"].PutValue("中華民國" + sSYR + "年01月01日起至" + sSYR + "年" + sSMN + "月" + sDAY + "日止");
            sheet1.Cells["J3"].PutValue(sPSYR + "年");
            sheet1.Cells["Q3"].PutValue(sSYR + "年度分配預算");
            sheet1.Cells["R3"].PutValue("預估至" + sSYR + "年底實現數(元)");
            sheet1.Cells["S3"].PutValue(sSYR + "年度撥款數");
            sheet1.Cells["T3"].PutValue(sPSYR + "~" + sSYR + "年度累計撥款數");
            sheet1.Cells["U3"].PutValue(sSYR + "年度已撥款現現數(元)");
            sheet1.Cells["V3"].PutValue(sSYR + ".01~" + sSYR + "." + sSMN + "累計實現數(元)");

            for (i = 0; i < dt.Rows.Count; i++)
            {
                sheet1.Cells["A" + (i + 5)].PutValue((i + 1).ToString());
                sheet1.Cells["B" + (i + 5)].PutValue(dt.Rows[i]["SIP_TD_CODE"].ToString());
                sheet1.Cells["C" + (i + 5)].PutValue(dt.Rows[i]["SIP_FD_CODE"].ToString());
                sheet1.Cells["E" + (i + 5)].PutValue(dt.Rows[i]["TD_CITY"].ToString());
                sheet1.Cells["G" + (i + 5)].PutValue(dt.Rows[i]["TD_NAME"].ToString());
                sheet1.Cells["H" + (i + 5)].PutValue(dt.Rows[i]["CEN_SP_RATIO_S"].ToString()); // 補助比例
                sheet1.Cells["I" + (i + 5)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["ALL_CEN_FUND"].ToString())); // 核定補助經費==>歷年中央款合計
                sheet1.Cells["J" + (i + 5)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PRE_CEN_FUND"].ToString())); // 分配預算==>前一年度中央款
                sheet1.Cells["K" + (i + 5)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PRE_PAY_FUND"].ToString())); // 撥款數==>實支數+ 已撥保留款
                sheet1.Cells["L" + (i + 5)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PRE_COST_FEE"].ToString())); // 實支數
                sheet1.Cells["M" + (i + 5)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PRE_PPAY_FEE"].ToString())); // 已撥保留款
                sheet1.Cells["N" + (i + 5)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PRE_REV_FUND"].ToString())); // 保留款中的權責數
                sheet1.Cells["Q" + (i + 5)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["CEN_FUND"].ToString())); // 年度分配預算==>本年度中央款
                sheet1.Cells["S" + (i + 5)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["COST_FEE"].ToString())); // 本年度實支數
                sheet1.Cells["V" + (i + 5)].PutValue(Initdate(dt.Rows[i]["TK105_DT"].ToString())); // 發包後請款期程==>實際發包日期(若無則為預訂)
                sheet1.Cells["X" + (i + 5)].PutValue(Initdate(dt.Rows[i]["TK107_PRE_DT"].ToString())); // 預定完工日期
                sheet1.Cells["Y" + (i + 5)].PutValue(Initdate(dt.Rows[i]["TK107_REL_DT"].ToString())); // 實際完工日期
                sheet1.Cells["Z" + (i + 5)].PutValue(Initdate(dt.Rows[i]["TK110_REL_DT"].ToString())); // 結案==>決算日期(若無則結算日期)
                sheet1.Cells["AB" + (i + 5)].PutValue(dt.Rows[i]["PRE_SCHD"].ToString());
                sheet1.Cells["AC" + (i + 5)].PutValue(dt.Rows[i]["REL_SCHD"].ToString());
                sheet1.Cells["AD" + (i + 5)].PutValue("總工程經費：" + dt.Rows[i]["BUDGET_FEE_B0"].ToString() + "\r\n" + "契約金額：" + dt.Rows[i]["BUDGET_FEE_C0"].ToString() + "\r\n" + "辦理情形：" + dt.Rows[i]["EXECT_DESC"].ToString());
            }

            sFileName = HttpUtility.UrlEncode("I-1.前瞻計畫執行情形表", System.Text.Encoding.UTF8);
            var savepath = Server.MapPath("~/Temp/" + sFileName + ".xlsx");

            workbook.Save(savepath, SaveFormat.Xlsx);

            if (Button == "1")
            {
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=" + sFileName + ".xlsx");
                Response.TransmitFile(savepath);
                Response.Flush();
                Response.End();
            }
            if (Button == "2")
            {
                using (var stream = new MemoryStream())
                {
                    var workbookdownload = new Workbook(savepath);

                    // Get the first worksheet
                    var worksheet = workbookdownload.Worksheets[0];

                    // Set the row height to automatically fit
                    worksheet.AutoFitRows();

                    var saveOptions = new OdsSaveOptions();

                    workbookdownload.Save(stream, saveOptions);
                    stream.Seek(0, SeekOrigin.Begin);
                    var odsBytes = stream.ToArray();

                    // Save the ODS file to the specified path on the server
                    var odsOutputFile = Server.MapPath("~/Temp/" + sFileName + ".ods");
                    System.IO.File.WriteAllBytes(odsOutputFile, odsBytes);

                    // ODS 檔下載
                    Response.Clear();
                    Response.ContentType = "application/vnd.oasis.opendocument.spreadsheet";
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + sFileName + ".ods");
                    Response.BinaryWrite(odsBytes);
                    Response.End();
                }
            }
        }

        /// <summary>
        /// I-1.雨水例行性管考報表[執行情形表]
        /// </summary>
        /// <param name="AN_PID"></param>
        /// <param name="SYR"> 年份 </param>
        /// <param name="SMN"> 月份 </param>
        /// <param name="connectionString"> SQL連接 </param>
        /// <param name="DEPT"> 部門 </param>
        /// <param name="PJ_NAME"> 案件名稱 </param>
        /// <param name="Button"> 報表生成 </param>
        private void PrintRPT0002(decimal AN_PID, string SYR, string SMN, string connectionString, string DEPT, string PJ_NAME, string Button)
        {
            string sFileName = "";
            string sFullFile = Server.MapPath("~/Templates/PJV0051.xlsx");
            FileInfo oFileInfo = new FileInfo(sFullFile);
            oFileInfo.IsReadOnly = false;
            Workbook workbook = new Workbook(sFullFile);
            Worksheet sheet1 = workbook.Worksheets["執行情形表"];

            int copyLine = 0;
            int runLine = 0;
            string runLineData = "";

            string sSYR = (DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (DateTime.Now.Month).ToString("00");
            string sPSYR = (DateTime.Now.Year - 1911 - 1).ToString("000");
            string sDAY = "31";

            DataView dvData = RainReport_Source.GetRPT0002Data(AN_PID, SYR, SMN, connectionString, DEPT);

            DataTable dtUnique = dvData.ToTable().DefaultView.ToTable();
            Dictionary<string, DataRow> dict = new Dictionary<string, DataRow>();

            foreach (DataRow row in dtUnique.Rows)
            {
                string key = row["TD_CODE"].ToString();

                if (!dict.ContainsKey(key))
                {
                    dict.Add(key, row);
                }
            }

            DataTable dtResult = new DataTable();
            dtResult = dict.Values.CopyToDataTable();
            DataView dvResult = dtResult.DefaultView;
            DataRowView rowV = null;

            double dTD_FUND = 0;
            double dCEN_FUND_1 = 0;
            double dCEN_FUND_2 = 0;
            double dCEN_FUND = 0;
            double dLOC_FUND = 0;
            double dOCR_FUND = 0;
            double dP_PLN_REL_COST = 0;
            double dYR_FUND = 0;
            double dL_PLN_COST = 0;
            double dL_REL_COST = 0;
            double dL_UNPAY_COST = 0;
            double dSAVE_FEE = 0;
            double dPRE_CON_LEN = 0;
            double dREL_CON_LEN = 0;
            double dPRE_CON_FLOOD = 0;
            double dREL_CON_FLOOD = 0;
            double dCON_LEN_PER = 0;
            double dCON_FLOOD_PER = 0;
            double dPRE_CON_LEN_T = 0;
            double dREL_CON_LEN_T = 0;
            double dPRE_CON_FLOOD_T = 0;
            double dREL_CON_FLOOD_T = 0;
            double dCON_LEN_PER_T = 0;
            double dCON_FLOOD_PER_T = 0;

            int lenDvData = dvResult.Count;
            int dvCnt;

            if (SYR != "")
            {
                sSYR = SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }

            if (SMN != "")
            {
                sSMN = SMN;
            }

            sDAY = Convert.ToDateTime((Convert.ToInt32(sSYR) + 1911).ToString() + "/" + sSMN + "/01").AddMonths(1).AddDays(-1).Day.ToString("00");
            copyLine = 6;

            for (dvCnt = 0; dvCnt <= lenDvData; dvCnt++)
            {
                runLine = copyLine + dvCnt;
                Aspose.Cells.Range sourceRange = sheet1.Cells.CreateRange(6, 1, 1, 40);
                Aspose.Cells.Range targetRange = sheet1.Cells.CreateRange(runLine, 1, 1, 40);
                sourceRange.Copy(targetRange);
                sheet1.Cells.SetRowHeight(runLine, sourceRange.RowHeight);
            }

            sheet1.Cells["A1"].PutValue(PJ_NAME + "執行情形表");
            sheet1.Cells["A2"].PutValue("中華民國" + sSYR + "年01月01日起至" + sSYR + "年" + sSMN + "月" + sDAY + "日止");
            sheet1.Cells["G4"].PutValue(sPSYR + "年\r\n中央款");
            sheet1.Cells["H4"].PutValue(sSYR + "年\r\n中央款");

            copyLine = 8;
            var TDV_TDDCFTS = db2.TDV_TDDCFTS.ToList();
            var TDDSTCF = db2.TDDSTCF.ToList();
            var TOTAL_FUND2 = 0;

            for (dvCnt = 0; dvCnt < lenDvData; dvCnt++)
            {
                rowV = dvResult[dvCnt];
                runLine = copyLine + dvCnt;
                runLineData = runLine.ToString();
                dCON_LEN_PER = 0;
                dCON_FLOOD_PER = 0;
                dPRE_CON_LEN = Convert.ToDouble(Tools.StringToDouble(rowV["PRE_CON_LEN"].ToString()));
                dREL_CON_LEN = Convert.ToDouble(Tools.StringToDouble(rowV["REL_CON_LEN"].ToString()));
                dPRE_CON_FLOOD = Convert.ToDouble(Tools.StringToDouble(rowV["PRE_CON_FLOOD"].ToString()));
                dREL_CON_FLOOD = Convert.ToDouble(Tools.StringToDouble(rowV["REL_CON_FLOOD"].ToString()));
                dPRE_CON_LEN_T += dPRE_CON_LEN;
                dREL_CON_LEN_T += dREL_CON_LEN;
                dPRE_CON_FLOOD_T += dPRE_CON_FLOOD;
                dREL_CON_FLOOD_T += dREL_CON_FLOOD;

                if (dPRE_CON_LEN != 0) { dCON_LEN_PER = (dREL_CON_LEN / dPRE_CON_LEN) * 100; }
                if (dPRE_CON_FLOOD != 0) { dCON_FLOOD_PER = (dREL_CON_FLOOD / dPRE_CON_FLOOD) * 100; }

                var TD_NAME_LINQ = rowV["TD_NAME"].ToString();
                var LAST_SYR = (int.Parse(SYR) - 1).ToString();
                var Money1 = TDV_TDDCFTS.Where(m => m.TD_NAME == TD_NAME_LINQ && SYR == LAST_SYR).FirstOrDefault()?.CEN_FUND ?? 0; /// 前年度核定數
                var Money2 = ((int)Source.Tools.StringToDouble(rowV["P_PLN_REL_COST"].ToString()));
                var Money3 = Money1 - Money2; /// 前一年度核定數－實支數費用
                var TD_CODE_LINQ = rowV["SIP_TD_CODE"].ToString();

                var Fund = TDDSTCF.Where(m => m.TD_CODE == TD_CODE_LINQ && m.SYR == SYR)
                    .GroupBy(b => new { b.SYR, b.SMN })
                    .Select(g => new
                    {
                        SYR = g.Key.SYR,
                        SMN = g.Key.SMN,
                        COST_FEE = g.Sum(b => b.COST_FEE),
                        PPAY_FEE = g.Sum(b => b.PPAY_FEE)
                    })
                    .OrderByDescending(b => b.SYR).ThenByDescending(b => b.SMN).FirstOrDefault();

                /// 當年度實支數加暫付
                var TOTAL_FUND = (Fund?.COST_FEE + Fund?.PPAY_FEE) ?? 0;
                long TOTALFUND = (long)TOTAL_FUND + TOTAL_FUND2;

                sheet1.Cells["O7"].PutValue(TOTALFUND);
                sheet1.Cells["A" + runLineData].PutValue((dvCnt + 1).ToString());
                sheet1.Cells["B" + runLineData].PutValue(rowV["CITY"].ToString());
                sheet1.Cells["C" + runLineData].PutValue(rowV["SIP_TD_CODE"].ToString());
                sheet1.Cells["D" + runLineData].PutValue(rowV["TD_NAME"].ToString());
                SetCellCenterAlign(sheet1.Cells, "D" + runLineData);
                sheet1.Cells["E" + runLineData].PutValue(Convert.ToDouble(rowV["TD_FUND"]));
                sheet1.Cells["F" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["CEN_SP_RATIO"].ToString()) * 100);
                sheet1.Cells["G" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["CEN_FUND_1"].ToString())); /// 選擇的計畫年度前一年總工程費
                SetCellCenterAlign(sheet1.Cells, "G" + runLineData);
                sheet1.Cells["H" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["CEN_FUND_2"].ToString())); /// 選擇的計畫當年度總工程費
                SetCellCenterAlign(sheet1.Cells, "H" + runLineData);
                sheet1.Cells["I" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["CEN_FUND"].ToString()));
                SetCellCenterAlign(sheet1.Cells, "I" + runLineData);
                sheet1.Cells["J" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["LOC_FUND"].ToString()));
                sheet1.Cells["K" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["OCR_FUND"].ToString()));
                sheet1.Cells["L" + runLineData].PutValue(Money3);
                sheet1.Cells["M" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["CEN_FUND_2"].ToString())); /// 2020412 修正項目 => 可支用預算(D)=當年度中央款(原本抓YR_FUND欄位值)
                SetCellCenterAlign(sheet1.Cells, "M" + runLineData);
                sheet1.Cells["N" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["L_PLN_COST"].ToString()));
                sheet1.Cells["O" + runLineData].PutValue(TOTAL_FUND);
                sheet1.Cells["P" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["L_UNPAY_COST"].ToString()));
                sheet1.Cells["Q" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["SAVE_FEE"].ToString()));
                sheet1.Cells["T" + runLineData].PutValue(Initdate(rowV["PRE_TK_103"].ToString()));
                sheet1.Cells["U" + runLineData].PutValue(Initdate(rowV["REL_TK_103"].ToString()));
                sheet1.Cells["V" + runLineData].PutValue(Initdate(rowV["PRE_TK_105"].ToString()));
                sheet1.Cells["W" + runLineData].PutValue(Initdate(rowV["REL_TK_105"].ToString()));
                /// 20250421 修改項目 => 標案工程進度預定(%), 改為百分比
                sheet1.Cells["X" + runLineData].PutValue(double.Parse(rowV["PRE_SCHD"].ToString()));
                Style styleX = sheet1.Cells["X" + runLineData].GetStyle(); styleX.Number = 9; sheet1.Cells["X" + runLineData].SetStyle(styleX);
                /// 20250421 修改項目 => 標案工程進度實際(%), 改為百分比
                sheet1.Cells["Y" + runLineData].PutValue(double.Parse(rowV["REL_SCHD"].ToString()));
                Style styleY = sheet1.Cells["Y" + runLineData].GetStyle(); styleY.Number = 9; sheet1.Cells["Y" + runLineData].SetStyle(styleY);
                /// 20250421 修改項目 => 標案工程進度差異(%), 改為百分比
                sheet1.Cells["Z" + runLineData].Formula = "=Y" + runLineData + "-X" + runLineData;
                Style styleZ = sheet1.Cells["Z" + runLineData].GetStyle(); styleZ.Number = 9; sheet1.Cells["Z" + runLineData].SetStyle(styleZ);
                sheet1.Cells["AA" + runLineData].PutValue(dPRE_CON_LEN.ToString());
                sheet1.Cells["AB" + runLineData].PutValue(dREL_CON_LEN.ToString());
                sheet1.Cells["AC" + runLineData].PutValue(dCON_LEN_PER.ToString("0.00") + "%");
                sheet1.Cells["AD" + runLineData].PutValue(dPRE_CON_FLOOD.ToString());
                sheet1.Cells["AE" + runLineData].PutValue(dREL_CON_FLOOD.ToString());
                sheet1.Cells["AF" + runLineData].PutValue(dCON_FLOOD_PER.ToString("0.00") + "%");
                /// 20250421 修改項目 => 標案完成日期預定, 重新取表抓值:EXTRA_PRE_SDT(PRE_SDT), 取自 CMETMS.dbo.TDDSDWB
                /// 20250421 修改項目 => 標案完成日期實際, 重新取表抓值:EXTRA_PRE_EDT(PRE_EDT), 取自 CMETMS.dbo.TDDSDWB
                sheet1.Cells["AG" + runLineData].PutValue(rowV["EXTRA_PRE_SDT"].ToString());
                sheet1.Cells["AH" + runLineData].PutValue(rowV["EXTRA_PRE_EDT"].ToString());
                sheet1.Cells["AI" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["BUDGET_FEE_B"].ToString()));
                sheet1.Cells["AJ" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["BUDGET_FEE_B_0101"].ToString()));
                sheet1.Cells["AK" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["BUDGET_FEE_B_N0101"].ToString()));
                sheet1.Cells["AL" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["BUDGET_FEE_C"].ToString()));
                sheet1.Cells["AM" + runLineData].PutValue(Source.Tools.StringToDouble(rowV["BUDGET_FEE_C_0101"].ToString()));

                /// -----------------------------------------------------------------------------------------------------------
                /// 2025/08/21 修改項目 => 取得 EXECT_DESC, 限制最多 200 字
                /// 2025/08/21 修改項目 => 每 20 個字換行
                string remark = rowV["EXECT_DESC"].ToString();

                if (remark.Length > 200)
                {
                    remark = remark.Substring(0, 200);
                }

                string formattedRemark = string.Concat(
                    remark.Select((c, i) => (i + 1) % 20 == 0 ? c + "\n" : c.ToString())
                );

                sheet1.Cells["AN" + runLineData].PutValue("　\n" + formattedRemark + "　");
                SetCellCenterAlign(sheet1.Cells, "AN" + runLineData);
                /// -----------------------------------------------------------------------------------------------------------

                dCEN_FUND = dCEN_FUND + Convert.ToDouble(Tools.StringToDouble(rowV["CEN_FUND"].ToString()));
                dTD_FUND = dTD_FUND + Convert.ToDouble(Tools.StringToDouble(rowV["TD_FUND"].ToString()));
                dCEN_FUND_1 = dCEN_FUND_1 + Convert.ToDouble(Tools.StringToDouble(rowV["CEN_FUND_1"].ToString()));
                dCEN_FUND_2 = dCEN_FUND_2 + Convert.ToDouble(Tools.StringToDouble(rowV["CEN_FUND_2"].ToString()));
                dLOC_FUND = dLOC_FUND + Convert.ToDouble(Tools.StringToDouble(rowV["LOC_FUND"].ToString()));
                dOCR_FUND = dOCR_FUND + Convert.ToDouble(Tools.StringToDouble(rowV["OCR_FUND"].ToString()));
                dP_PLN_REL_COST = dP_PLN_REL_COST + Convert.ToDouble(Tools.StringToDouble(rowV["P_PLN_REL_COST"].ToString()));
                dYR_FUND = dYR_FUND + Convert.ToDouble(Tools.StringToDouble(rowV["YR_FUND"].ToString()));
                dL_PLN_COST = dL_PLN_COST + Convert.ToDouble(Tools.StringToDouble(rowV["L_PLN_COST"].ToString()));
                dL_REL_COST = dL_REL_COST + Convert.ToDouble(Tools.StringToDouble(rowV["L_REL_COST"].ToString()));
                dL_UNPAY_COST = dL_UNPAY_COST + Convert.ToDouble(Tools.StringToDouble(rowV["L_UNPAY_COST"].ToString()));
                dSAVE_FEE = dSAVE_FEE + Convert.ToDouble(Tools.StringToDouble(rowV["SAVE_FEE"].ToString()));
            }

            sheet1.Cells.DeleteColumn(8);
            sheet1.AutoFitColumns();
            sheet1.AutoFitRows();

            dCON_LEN_PER_T = 0;
            dCON_FLOOD_PER_T = 0;
            if (dPRE_CON_LEN_T != 0) { dCON_LEN_PER_T = (dREL_CON_LEN_T / dPRE_CON_LEN_T) * 100; }
            if (dPRE_CON_FLOOD_T != 0) { dCON_FLOOD_PER_T = (dREL_CON_FLOOD_T / dPRE_CON_FLOOD_T) * 100; }

            sheet1.Cells.DeleteRow(6);                  /// 20241127 修改項目 => 刪除第6行                                            
            int lastRow = sheet1.Cells.MaxDataRow + 1;  /// 20241127 修改項目 => 重新計算總計
            sheet1.Cells.InsertRow(lastRow);
            sheet1.Cells[lastRow, 3].PutValue("總計");
            sheet1.Cells[lastRow, 4].PutValue(Source.Tools.StringToDouble(dTD_FUND.ToString()));
            sheet1.Cells[lastRow, 6].PutValue(Source.Tools.StringToDouble(dCEN_FUND_1.ToString()));
            sheet1.Cells[lastRow, 7].PutValue(Source.Tools.StringToDouble(dCEN_FUND_2.ToString()));
            sheet1.Cells[lastRow, 8].PutValue(Source.Tools.StringToDouble(dCEN_FUND.ToString()));
            sheet1.Cells[lastRow, 9].PutValue(Source.Tools.StringToDouble(dLOC_FUND.ToString()));
            sheet1.Cells[lastRow, 10].PutValue(Source.Tools.StringToDouble(dOCR_FUND.ToString()));
            sheet1.Cells[lastRow, 11].PutValue(Source.Tools.StringToDouble(dP_PLN_REL_COST.ToString()));
            sheet1.Cells[lastRow, 12].PutValue(Source.Tools.StringToDouble(dCEN_FUND_2.ToString()));     /// 2020413 修改 => 可支用預算(D) = 當年度中央款(原本抓dYR_FUND值)
            sheet1.Cells[lastRow, 13].PutValue(Source.Tools.StringToDouble(dL_PLN_COST.ToString()));
            sheet1.Cells[lastRow, 15].PutValue(Source.Tools.StringToDouble(dL_UNPAY_COST.ToString()));
            sheet1.Cells[lastRow, 26].PutValue(Source.Tools.StringToDouble(dPRE_CON_LEN_T.ToString()));
            sheet1.Cells[lastRow, 27].PutValue(Source.Tools.StringToDouble(dREL_CON_LEN_T.ToString()));
            sheet1.Cells[lastRow, 28].PutValue(dCON_LEN_PER_T.ToString("0.00") + "%");
            sheet1.Cells[lastRow, 29].PutValue(Source.Tools.StringToDouble(dPRE_CON_FLOOD_T.ToString()));
            sheet1.Cells[lastRow, 30].PutValue(Source.Tools.StringToDouble(dREL_CON_FLOOD_T.ToString()));
            sheet1.Cells[lastRow, 31].PutValue(dCON_FLOOD_PER_T.ToString("00.0") + "%");

            /// ======================== 報表樣式 ========================

            int maxRow = sheet1.Cells.MaxDataRow;
            int maxColumn = sheet1.Cells.MaxDataColumn;
            Range range = sheet1.Cells.CreateRange(0, 0, maxRow - 5, maxColumn + 1);

            Style style = sheet1.Workbook.CreateStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;

            StyleFlag styleFlag = new StyleFlag { Borders = true };
            range.ApplyStyle(style, styleFlag);

            sheet1.Cells.SetColumnWidth(0, 6);    /// A 項次
            sheet1.Cells.SetColumnWidth(1, 10);   /// B 計畫說明
            sheet1.Cells.SetColumnWidth(2, 25);   /// C
            sheet1.Cells.SetColumnWidth(3, 30);   /// D
            sheet1.Cells.SetColumnWidth(4, 10);   /// E
            sheet1.Cells.SetColumnWidth(5, 10);   /// F
            sheet1.Cells.SetColumnWidth(6, 10);   /// G 工作名稱
            sheet1.Cells.SetColumnWidth(7, 15);   /// H
            sheet1.Cells.SetColumnWidth(8, 18);   /// I
            sheet1.Cells.SetColumnWidth(9, 18);   /// J
            sheet1.Cells.SetColumnWidth(10, 18);  /// K
            sheet1.Cells.SetColumnWidth(11, 18);  /// L
            sheet1.Cells.SetColumnWidth(12, 15);  /// M
            sheet1.Cells.SetColumnWidth(13, 15);  /// N
            sheet1.Cells.SetColumnWidth(14, 15);  /// O
            sheet1.Cells.SetColumnWidth(15, 15);  /// P
            sheet1.Cells.SetColumnWidth(16, 15);  /// Q
            sheet1.Cells.SetColumnWidth(17, 15);  /// R
            sheet1.Cells.SetColumnWidth(18, 15);  /// S
            sheet1.Cells.SetColumnWidth(19, 15);  /// T
            sheet1.Cells.SetColumnWidth(20, 15);  /// U
            sheet1.Cells.SetColumnWidth(21, 15);  /// V
            sheet1.Cells.SetColumnWidth(22, 15);  /// W
            sheet1.Cells.SetColumnWidth(25, 15);  /// Z
            sheet1.Cells.SetColumnWidth(26, 15);  /// AA
            sheet1.Cells.SetColumnWidth(27, 10);  /// AB
            sheet1.Cells.SetColumnWidth(28, 18);  /// AC
            sheet1.Cells.SetColumnWidth(29, 18);  /// AD
            sheet1.Cells.SetColumnWidth(30, 18);  /// AE
            sheet1.Cells.SetColumnWidth(31, 18);  /// AF
            sheet1.Cells.SetColumnWidth(32, 18);  /// AG
            sheet1.Cells.SetColumnWidth(33, 18);  /// AH
            sheet1.Cells.SetColumnWidth(34, 18);  /// AI
            sheet1.Cells.SetColumnWidth(35, 18);  /// AJ
            sheet1.Cells.SetColumnWidth(36, 18);  /// AK
            sheet1.Cells.SetColumnWidth(37, 18);  /// AL
            sheet1.Cells.SetColumnWidth(38, 18);  /// AM

            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;

            sheet1.AutoFitRows();
            sheet1.PageSetup.Zoom = 80;

            /// =========================================================

            sFileName = HttpUtility.UrlEncode("I-1.執行情形表", System.Text.Encoding.UTF8);
            var savepath = Server.MapPath("~/Temp/" + sFileName + ".xlsx");
            workbook.Save(savepath, SaveFormat.Xlsx);

            if (Button == "1")
            {
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=" + sFileName + ".xlsx");
                Response.TransmitFile(savepath);
                Response.Flush();
                Response.End();
            }

            if (Button == "2")
            {
                using (var stream = new MemoryStream())
                {
                    var workbookdownload = new Workbook(savepath);
                    var worksheet = workbookdownload.Worksheets[0];
                    worksheet.AutoFitRows();
                    var saveOptions = new OdsSaveOptions();
                    workbookdownload.Save(stream, saveOptions);
                    stream.Seek(0, SeekOrigin.Begin);
                    var odsBytes = stream.ToArray();
                    var odsOutputFile = Server.MapPath("~/Temp/" + sFileName + ".ods");
                    System.IO.File.WriteAllBytes(odsOutputFile, odsBytes);
                    Response.Clear();
                    Response.ContentType = "application/vnd.oasis.opendocument.spreadsheet";
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + sFileName + ".ods");
                    Response.BinaryWrite(odsBytes);
                    Response.End();
                }
            }
        }

        public IEnumerable<SiteCityViewModel> User_Right(int CityState)
        {
            ViewBag.year = db2.PJMFORM.OrderByDescending(m => m.SYR).DistinctBy(m => m.SYR).ToList();
            ViewBag.year2 = db2.PJMFORM.OrderBy(m => m.SYR).DistinctBy(m => m.SYR).ToList();

            if (CityState == 11)
            {
                ViewBag.year = db2.PJMFORM.Where(m => m.SYR.CompareTo("106") >= 0).OrderByDescending(m => m.SYR).DistinctBy(m => m.SYR).ToList();
                ViewBag.year2 = db2.PJMFORM.Where(m => m.SYR.CompareTo("106") >= 0).OrderBy(m => m.SYR).DistinctBy(m => m.SYR).ToList();
            }

            if (CityState == 7)
            {
                List<PJMFORM_CMEPMS> yearList = Enumerable.Range(97, 5)
                 .Select(year => new PJMFORM_CMEPMS
                 {
                     SYR = year.ToString("D3")  // 將數字格式化為三位數，例如 "097"
                 }).ToList();

                ViewBag.year = yearList;
            }
            string UserID = (string)Session["UsrId"];
            var SECCITY = db.SECCITY.Where(m => m.USR_ID == UserID).ToList();
            var SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            var query = from a in SECCITY
                        join b in SMCPOSTN on a.CITY equals b.CITY
                        select new SiteCityViewModel { sECCITY = a, sMCPOSTN = b };
            if (CityState == 2 | CityState == 5) { }
            else
            {
                query = query.DistinctBy(m => m.sMCPOSTN.AREA);
            }

            if (query == null || !query.Any())
            {
                var OGCDPUR = db.OGCDPUR.Where(m => m.USR_ID == UserID && m.DATA_TP == "D").FirstOrDefault().DEPT_NO;

                var Dept = db.SECCITY.Where(m => m.DEPT_NO == OGCDPUR).ToList();
                query = from a in Dept
                        join b in SMCPOSTN on a.CITY equals b.CITY
                        select new SiteCityViewModel { sECCITY = a, sMCPOSTN = b };
                if (CityState == 2 | CityState == 5)
                {
                    query = query.DistinctBy(m => m.sMCPOSTN.CITY);
                }
                else
                {
                    query = query.DistinctBy(m => m.sMCPOSTN.AREA);
                }
            }
            return query;
        }

        public ActionResult State2_2()
        {
            ViewBag.State = "2_2";
            var query = User_Right(2);
            return View(query);
        }

        public ActionResult State6_2()
        {
            ViewBag.State = "6_2";
            var query = User_Right(6); /////之後要開啟 現在測試一下而已
            return View(/*query*/);
        }

        public ActionResult State8_4()
        {
            ViewBag.State = "8_4";
            var query = User_Right(9);
            state9ViewModel CityAndColumn = new state9ViewModel();

            var Plan_Name = db2.PJV_PJDPCAT.OrderByDescending(m => m.SYR).Where(m => m.PJ_TYPE == "B").ToList();

            CityAndColumn.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            CityAndColumn.PJV_PJDPCAT = Plan_Name;
            CityAndColumn.OGCTITLE = db.OGCTITLE.ToList();

            return View(CityAndColumn);
        }

        public ActionResult State9_2()
        {

            ViewBag.State = "9_2";
            var query = User_Right(9);
            state9ViewModel CityAndColumn = new state9ViewModel();
            var PJV_PJDPCAT = db2.PJV_PJDPCAT.OrderByDescending(m => m.SYR).Where(m => m.PJ_TYPE == "B" /*&& (m.ADV_DP == "A270100" || m.ADV_DP == "A270200")*/).ToList();
            var PJMFORM = db2.PJMFORM.ToList();
            var Plan_Name = (from a in PJV_PJDPCAT
                             join b in PJMFORM on a.PID equals b.PID
                             select new Water_Type_Use
                             {
                                 SYR = a.SYR,
                                 PID = a.PID,
                                 Water_Type = b.Water_Type,
                                 AN_PJ_NAME = a.AN_PJ_NAME,
                                 PJ_NAME = a.PJ_NAME
                             }).ToList();
            CityAndColumn.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            CityAndColumn.Water_Type = Plan_Name;
            CityAndColumn.OGCTITLE = db.OGCTITLE.ToList();

            return View(CityAndColumn);
        }

        public ActionResult State10_2()
        {
            ViewBag.State = "10_2";
            var query = User_Right(10);

            //////扣除前瞻和流域的分支計畫，帶回傳時處理
            var PJMFORM = db2.PJMFORM.Where(m => m.AN_PID != 118 || m.AN_PID != 122 || m.AN_PID != 134 || m.AN_PID != 176 || m.AN_PID != 198 || m.AN_PID != 215 || m.AN_PID != 246 || m.AN_PID != 61 || m.AN_PID != 96 || m.AN_PID != 139).ToList();

            ///////抓取所有的分支計畫
            var pjNames = (from m in PJMFORM
                           where m.AN_PID != null && m.Water_Type == "雨水使用"
                           join n in PJMFORM on m.AN_PID equals n.PID
                           select new
                           {
                               PJ = m.PJ_NAME,
                               PJ_NAME = "(" + db2.PJMFORM.Where(z => z.PID == m.AN_PID).FirstOrDefault().PJ_NAME + ")" + m.PJ_NAME,
                               ProjectNameID = m.AN_PID,
                               branchProjectNameID = m.PID
                           }).DistinctBy(m => m.PJ_NAME).ToList();

            // Create the SelectListItem objects using the list previously generated
            var PJ_BAG = pjNames.Select(m => new State10ViewModel
            {
                PJ = m.PJ,
                ProjectNameID = m.branchProjectNameID.ToString(),
                PJ_NAME = m.PJ_NAME.ToString(),
                MASTER_PID = m.ProjectNameID.ToString()
            }).OrderBy(m => m.PJ).ToList();
            ViewBag.PJ_BAG = PJ_BAG;

            ViewBag.CITY = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            ViewBag.TOWN = db.SMCPOSTN.ToList();

            return View(query);
        }

        public ActionResult State12_2()
        {
            ViewBag.State = "12_2";
            var query = User_Right(12);
            return View(query);
        }

        /// <summary>
        /// B-1 修正預算作業表
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="period"></param>
        /// <param name="Subject"></param>
        /// <param name="SYR"></param>
        /// <param name="ODS"></param>
        /// <param name="Choose"></param>
        /// <returns></returns>
        public ActionResult ReportPrint6(string PID, string period, string[] Subject, string SYR, string ODS, string Choose)
        {
            Trace.WriteLine($"ReportPrint6");
            Trace.WriteLine($"PID = {PID}");
            Trace.WriteLine($"Period = {period}");
            Trace.WriteLine($"SYR = {SYR}");
            Trace.WriteLine($"ODS = {ODS}");
            Trace.WriteLine($"Choose = {Choose}");

            string ADV_UID = User.Identity.Name;
            string sACC_ID = "";
            string aDV_DP = "";
            string Dept = db.OGCDPUR
                .Where(x => x.USR_ID == ADV_UID && x.DATA_TP == "D")
                .Select(x => x.DEPT_NO)
                .FirstOrDefault();

            if (Subject != null && Subject.Length > 0)
            {
                sACC_ID = "'" + string.Join("','", Subject) + "'";
            }

            if (Dept == "A270000")
            {
                var result1 = db.OGCDEPT.Where(m => m.DEPT_NO.Contains("A27")).ToList();

                foreach (var item in result1)
                {
                    aDV_DP += item.DEPT_NO + ",";
                }

                aDV_DP = aDV_DP.TrimEnd(',');
            }
            else if (db.OGCDEPT.Any(m => m.AN_DEPT_NO == Dept))
            {
                var result1 = db.OGCDEPT.Where(m => m.AN_DEPT_NO == Dept);

                foreach (var item in result1)
                {
                    aDV_DP += item.DEPT_NO + ",";
                }

                aDV_DP = aDV_DP.TrimEnd(',');
            }
            else
            {
                aDV_DP = Dept;
            }

            if (Dept == "B160000")
            {
                var result1 = db.OGCDEPT.Where(m => m.DEPT_NO.Contains("A2705")).ToList();

                aDV_DP += !String.IsNullOrEmpty(aDV_DP) ? "," : "";

                foreach (var item in result1)
                {
                    aDV_DP += item.DEPT_NO + ",";
                }

                aDV_DP = aDV_DP.TrimEnd(',');
            }

            string SEQ_NO = period;

            var PJNAME = string.Empty;

            DataSet DsMain1 = new DataSet();

            bool notSetName = true;

            /// 20241118 修改項目 => 接受多項 PID
            var PIDList = PID.Split(',').Select(pid => pid.Trim()).ToList();

            foreach (var singlePID in PIDList)
            {
                /// 若為前瞻基礎建設計畫，excel顯示名稱修正
                if (notSetName)
                {
                    PJNAME = db2.PJMFORM.Where(m => m.PID.ToString() == singlePID).FirstOrDefault().PJ_NAME;

                    var mergeName = new[] { "01辦理水與安全業務費", "02辦理水與安全獎補助費", "03辦理水與安全設備投資費" };

                    if (mergeName.Any(name => PJNAME.Contains(name)))
                    {
                        PJNAME = "前瞻基礎建設計畫-縣市管河川及區域排水整理改善計畫";
                        notSetName = false;
                    }
                }

                DataSet tempDataSet = GetSP_RCD0010_UDF003_New(singlePID, SEQ_NO, sACC_ID, ADV_UID, aDV_DP, Dept, Choose);

                foreach (DataTable table in tempDataSet.Tables)
                {
                    /// 複製 DataTable
                    DataTable clonedTable = table.Copy();

                    if (!DsMain1.Tables.Contains(table.TableName))
                    {
                        /// 如果 DsMain1 不存在該表格, 直接添加
                        DsMain1.Tables.Add(clonedTable);
                    }
                    else
                    {
                        /// 如果 DsMain1 已有該表格, 合併資料
                        DsMain1.Tables[table.TableName].Merge(clonedTable);
                    }
                }
            }

            DataTable DtS00 = DsMain1.Tables["DataTable"];
            DataTable DtS01 = DsMain1.Tables["DataTable"];
            DataTable DtS02 = DsMain1.Tables["DataTable1"];
            DataTable DtS03 = DsMain1.Tables["DataTable2"];
            DataTable DtS04 = DsMain1.Tables["DataTable3"];
            DataTable DtS05 = DsMain1.Tables["DataTable4"];
            DataTable DtS06 = DsMain1.Tables["DataTable5"];

            double ACC02_FUND_FINAL;

            if (DtS00.Rows.Count > 1)
            {
                ACC02_FUND_FINAL = ReportController.ObjectToDouble(DtS00.Rows[0]["ACC02_FUND"].ToString()) + ReportController.ObjectToDouble(DtS00.Rows[1]["ACC02_FUND"].ToString());
            }
            else
            {
                ACC02_FUND_FINAL = 0;
            }

            DataTable DtS000 = null;
            DataTable DtS11 = null;
            DataTable DtS12 = null;
            DataTable DtS13 = null;
            DataTable DtS14 = null;
            DataTable DtS15 = null;
            DataTable DtS16 = null;

            double ACC02_FUND_SECOND = 0;

            #region

            /// 將 PID 解析為列表
            var PIDLists = PID.Split(',').Select(pid => pid.Trim()).ToList();

            if (SEQ_NO != "0")
            {
                DataSet DsMain2 = new DataSet();
                SEQ_NO = Convert.ToString(Convert.ToInt32(SEQ_NO) - 1);

                foreach (var singlePID in PIDLists)
                {
                    DataSet tempDataSet = GetSP_RCD0010_UDF003_New(singlePID, SEQ_NO, sACC_ID, ADV_UID, aDV_DP, Dept, Choose);

                    foreach (DataTable table in tempDataSet.Tables)
                    {
                        DataTable clonedTable = table.Copy(); // 複製 DataTable

                        if (!DsMain2.Tables.Contains(table.TableName))
                            DsMain2.Tables.Add(clonedTable); // 如果 DsMain2 不存在該表格，直接添加
                        else
                            DsMain2.Tables[table.TableName].Merge(clonedTable); // 如果 DsMain2 已有該表格，合併資料
                    }
                }

                DtS000 = DsMain2.Tables["DataTable"];
                DtS11 = DsMain2.Tables["DataTable"];
                DtS12 = DsMain2.Tables["DataTable1"];
                DtS13 = DsMain2.Tables["DataTable2"];
                DtS14 = DsMain2.Tables["DataTable3"];
                DtS15 = DsMain2.Tables["DataTable4"];
                DtS16 = DsMain2.Tables["DataTable5"];
            }

            #endregion

            var result = db.SMCPOSTN
                           .GroupBy(p => p.CITY)
                           .Select(g => new
                           {
                               CPA_CITY_ID = g.Max(p => p.CPA_CITY_ID),
                               CITY = g.Key
                           })
                           .OrderBy(p => p.CPA_CITY_ID).ToList();

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("CPA_CITY_ID", typeof(int));
            dataTable.Columns.Add("CITY", typeof(string));

            foreach (var item in result)
            {
                dataTable.Rows.Add(item.CPA_CITY_ID, item.CITY);
            }

            var DtCITY = getReport02_CITY();

            /// Aspose 文件處理

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);

            #region -------------------- 總表(公辦) --------------------
            DataTable DtH01 = GetDtH01(DtS01, DtS11, "1", DtCITY);
            DataTable DtH01H = ReportController.GetHeadDt("1", PJNAME, period, SYR);
            #endregion

            #region -------------------- 總表(BOT) --------------------
            DataTable DtH04 = GetDtH01(DtS04, DtS14, "4", DtCITY);
            DataTable DtH04H = ReportController.GetHeadDt("4", PJNAME, period, SYR);
            #endregion

            #region -------------------- 總表 --------------------
            DataTable DtH00H = ReportController.GetHeadDt("0", PJNAME, period, SYR);
            DataTable DtH00 = GetDtH00(DtH01, DtH04, DtH00H);
            #endregion

            #region -------------------- 彙整明細、彙整明細簡易版 --------------------
            DataTable DtH06H = ReportController.GetHeadDt("6", PJNAME, period, SYR);
            List<B1DtBack> Lb1Dtback = new List<B1DtBack>();
            List<B1DtFooter> Lb1DtFooter = new List<B1DtFooter>();
            GetDtH02(DtS06, DtS16, "6", DtCITY, period, Lb1Dtback, Lb1DtFooter);
            #endregion

            var filepath = Server.MapPath(@"~/Templates/PJV0010_02.xlsx");
            var file = new FileInfo(filepath);
            var workbook = new Aspose.Cells.Workbook(filepath);

            /// 總表(公辦)
            workbook.Worksheets.Add("總表(公辦)");
            Worksheet ws2 = workbook.Worksheets["總表(公辦)"];
            Worksheet wsa = workbook.Worksheets[0];

            /// 總表(BOT)
            workbook.Worksheets.Add("總表(BOT)");
            Worksheet ws3 = workbook.Worksheets["總表(BOT)"];
            Worksheet wsb = workbook.Worksheets[2];

            /// 總表
            workbook.Worksheets.Add("總表");
            Worksheet ws1 = workbook.Worksheets["總表"];
            Worksheet wsc = workbook.Worksheets[4];
            ws1.MoveTo(5);

            /// 彙整明細
            workbook.Worksheets.Add("彙整明細");
            Worksheet ws4 = workbook.Worksheets["彙整明細"];
            Worksheet wsd = workbook.Worksheets[1];

            /// 彙整明細簡易版
            workbook.Worksheets.Add("彙整明細簡易版");
            Worksheet ws5 = workbook.Worksheets["彙整明細簡易版"];
            Worksheet wse = workbook.Worksheets[3];

            #region -------------------- NEW 總表 --------------------

            // 模板
            int templateStartRow = 3;   // A4:Q5 -> row 3~4
            int blockHeight = 2;        // 公辦 + BOT
            int spacerRows = 1;

            /// summary (你提供的範圍 $A$7:$Q$10)
            /// int summaryStartRow = 6;    // row 6~9
            /// int summaryHeight = 4;      // 4 rows


            /// string PID, string period, string[] Subject, string SYR, string ODS, string Choose
            /**
             * B-1.修正預算管考報表
             * PID = 298,299,300
             * Period = 2
             * SYR = 114
             * ODS = 
             * Chose = ALL
             */

            Trace.WriteLine("B-1.修正預算管考報表");
            Trace.WriteLine($"PID = {PID}");
            Trace.WriteLine($"Period = {period}");
            Trace.WriteLine($"SYR = {SYR}");
            Trace.WriteLine($"ODS = {ODS}");
            Trace.WriteLine($"Chose = {Choose}");

            Worksheet ws6 = workbook.Worksheets["Sheet6"];
            ws6.Name = "總表(New)";

            /// 暫存區(總計區塊)
            int summaryTempRow = ws6.Cells.MaxDataRow + 50;

            DataTable dt = RainReport_Source.B1RERPORTDATA(PID, SYR, period, connectionString);

            /// 測試程式碼
            Debug.WriteLine("DT1");
            foreach (DataRow rows in dt.Rows)
            {
                foreach (DataColumn cols in dt.Columns)
                {
                    Debug.Write($"{cols.ColumnName}:{rows[cols]} \t");
                }
                Debug.WriteLine("");
            }

            var cityGROUP = dt.AsEnumerable().GroupBy(r => r["CITY"].ToString()).OrderBy(g => g.Key);

            // 依 CITY 分組，空白 CITY 視為 不分區
            var cityGroups = dt.AsEnumerable().GroupBy(r => r["CITY"].ToString()).OrderBy(g => g.Key).ToList();

            ws6.Cells["A1"].PutValue("B-1.修正預算管考報表");

            // 2. 刪掉模板下面所有列
            ws6.Cells.DeleteRows(templateStartRow + blockHeight, 200);

            int currentStartRow = templateStartRow;

            /// Pperiod(期數換算) = string -> int -> string
            int TransPeriod = Convert.ToInt32(period);
            TransPeriod = TransPeriod - 1;
            string Pperiod = TransPeriod.ToString();

            /// 原預算
            if (period == "0")
            {
                ws6.Cells["C2"].PutValue($"第{period}期");
                ws6.Cells["H2"].PutValue($"第{Math.Max(Convert.ToInt32(period) - 1, 0)}期");

                for (int i = 0; i < cityGroups.Count; i++)
                {
                    var cityGroup = cityGroups[i];

                    ws6.Cells.CopyRows(ws6.Cells, templateStartRow, currentStartRow, blockHeight);

                    int govRow = currentStartRow;
                    int botRow = currentStartRow + 1;

                    string city = cityGroup.Key;

                    ws6.Cells[govRow, 0].PutValue(city);
                    ws6.Cells[botRow, 0].PutValue("");
                    ws6.Cells[govRow, 1].PutValue("公辦");
                    ws6.Cells[botRow, 1].PutValue("BOT");

                    /// 公辦 => 業務費 / 獎補助費 / 設備及投資
                    double gov200 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov400 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov300 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// BOT  => 業務費 / 獎補助費 / 設備及投資
                    double bot200 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot400 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot300 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// C - E：本期
                    ws6.Cells[govRow, 2].PutValue(gov300); // C
                    ws6.Cells[govRow, 3].PutValue(gov400); // D
                    ws6.Cells[govRow, 4].PutValue(gov200); // E
                    ws6.Cells[botRow, 2].PutValue(bot300); // C
                    ws6.Cells[botRow, 3].PutValue(bot400); // D
                    ws6.Cells[botRow, 4].PutValue(bot200); // E

                    /// H - J：前期
                    ws6.Cells[govRow, 7].PutValue(0); // H
                    ws6.Cells[govRow, 8].PutValue(0); // I
                    ws6.Cells[govRow, 9].PutValue(0); // J
                    ws6.Cells[botRow, 7].PutValue(0); // H
                    ws6.Cells[botRow, 8].PutValue(0); // I
                    ws6.Cells[botRow, 9].PutValue(0); // J

                    currentStartRow += (blockHeight + spacerRows);
                }

                /// 複製格式, 清除資料(僅保留格式)
                ws6.Cells.CopyRows(ws6.Cells, templateStartRow, currentStartRow, blockHeight);

                for (int r = currentStartRow; r < currentStartRow + blockHeight; r++)
                {
                    for (int c = 0; c <= 16; c++)
                    {
                        Cell cell = ws6.Cells[r, c];

                        if (string.IsNullOrEmpty(cell.Formula))
                        {
                            cell.Value = null;
                        }
                    }
                }

                /// 標題
                ws6.Cells[currentStartRow, 0].PutValue("總計");
                ws6.Cells[currentStartRow + 1, 0].PutValue("");
                ws6.Cells[currentStartRow, 1].PutValue("公辦");
                ws6.Cells[currentStartRow + 1, 1].PutValue("BOT");

                /// 公辦(本期)
                double sumGov300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumGov400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumGov200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                /// BOT(本期)
                double sumBot300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumBot400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumBot200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                ws6.Cells[currentStartRow, 2].PutValue(sumGov300); // C
                ws6.Cells[currentStartRow, 3].PutValue(sumGov400); // D
                ws6.Cells[currentStartRow, 4].PutValue(sumGov200); // E
                ws6.Cells[currentStartRow + 1, 2].PutValue(sumBot300); // C
                ws6.Cells[currentStartRow + 1, 3].PutValue(sumBot400); // D
                ws6.Cells[currentStartRow + 1, 4].PutValue(sumBot200); // E

                ws6.Cells[currentStartRow, 7].PutValue(0); // H
                ws6.Cells[currentStartRow, 8].PutValue(0); // I
                ws6.Cells[currentStartRow, 9].PutValue(0); // J
                ws6.Cells[currentStartRow + 1, 7].PutValue(0); // H
                ws6.Cells[currentStartRow + 1, 8].PutValue(0); // I
                ws6.Cells[currentStartRow + 1, 9].PutValue(0); // J
            }
            /// 含預算 > 0 期數
            else
            {
                ws6.Cells["C2"].PutValue($"第{Math.Max(Convert.ToInt32(period) - 1, 0)}期");
                ws6.Cells["H2"].PutValue($"第{period}期");

                for (int i = 0; i < cityGroups.Count; i++)
                {
                    var cityGroup = cityGroups[i];

                    ws6.Cells.CopyRows(ws6.Cells, templateStartRow, currentStartRow, blockHeight);

                    int govRow = currentStartRow;
                    int botRow = currentStartRow + 1;

                    string city = cityGroup.Key;

                    ws6.Cells[govRow, 0].PutValue(city);
                    ws6.Cells[botRow, 0].PutValue("");

                    ws6.Cells[govRow, 1].PutValue("公辦");
                    ws6.Cells[botRow, 1].PutValue("BOT");

                    /// 公辦 => 業務費 / 獎補助費 / 設備及投資
                    double gov200P = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov400P = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov300P = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// BOT  => 業務費 / 獎補助費 / 設備及投資
                    double bot200P = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot400P = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot300P = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// C - E : 前一期
                    ws6.Cells[govRow, 2].PutValue(gov300P); // C
                    ws6.Cells[govRow, 3].PutValue(gov400P); // D
                    ws6.Cells[govRow, 4].PutValue(gov200P); // E
                    ws6.Cells[botRow, 2].PutValue(bot300P); // C
                    ws6.Cells[botRow, 3].PutValue(bot400P); // D
                    ws6.Cells[botRow, 4].PutValue(bot200P); // E

                    /// 公辦 => 業務費 / 獎補助費 / 設備及投資
                    double gov200 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov400 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov300 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// BOT  => 業務費 / 獎補助費 / 設備及投資
                    double bot200 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot400 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot300 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// H - J : 本期
                    ws6.Cells[govRow, 7].PutValue(gov300); // H
                    ws6.Cells[govRow, 8].PutValue(gov400); // I
                    ws6.Cells[govRow, 9].PutValue(gov200); // J
                    ws6.Cells[botRow, 7].PutValue(bot300); // H
                    ws6.Cells[botRow, 8].PutValue(bot400); // I
                    ws6.Cells[botRow, 9].PutValue(bot200); // J

                    currentStartRow += (blockHeight + spacerRows);
                }

                /// 複製格式, 清除資料(僅保留格式)
                ws6.Cells.CopyRows(ws6.Cells, templateStartRow, currentStartRow, blockHeight);

                for (int r = currentStartRow; r < currentStartRow + blockHeight; r++)
                {
                    for (int c = 0; c <= 16; c++)
                    {
                        Cell cell = ws6.Cells[r, c];

                        if (string.IsNullOrEmpty(cell.Formula))
                        {
                            cell.Value = null;
                        }
                    }
                }

                /// 標題
                ws6.Cells[currentStartRow, 0].PutValue("總計");
                ws6.Cells[currentStartRow + 1, 0].PutValue("");
                ws6.Cells[currentStartRow, 1].PutValue("公辦");
                ws6.Cells[currentStartRow + 1, 1].PutValue("BOT");

                /// 公辦(前期)
                double sumGov300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumGov400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumGov200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                /// BOT(前期)
                double sumBot300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumBot400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumBot200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == Pperiod).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                /// 公辦(本期)
                double PsumGov300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double PsumGov400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double PsumGov200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                /// BOT(本期)
                double PsumBot300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double PsumBot400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double PsumBot200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == period).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                ws6.Cells[currentStartRow, 2].PutValue(sumGov300); // C
                ws6.Cells[currentStartRow, 3].PutValue(sumGov400); // D
                ws6.Cells[currentStartRow, 4].PutValue(sumGov200); // E
                ws6.Cells[currentStartRow + 1, 2].PutValue(sumBot300); // C
                ws6.Cells[currentStartRow + 1, 3].PutValue(sumBot400); // D
                ws6.Cells[currentStartRow + 1, 4].PutValue(sumBot200); // E

                ws6.Cells[currentStartRow, 7].PutValue(PsumGov300); // H
                ws6.Cells[currentStartRow, 8].PutValue(PsumGov400); // I
                ws6.Cells[currentStartRow, 9].PutValue(PsumGov200); // J
                ws6.Cells[currentStartRow + 1, 7].PutValue(PsumBot300); // H
                ws6.Cells[currentStartRow + 1, 8].PutValue(PsumBot400); // I
                ws6.Cells[currentStartRow + 1, 9].PutValue(PsumBot200); // J
            }


            #endregion

            #region -------------------- 總表(公辦) --------------------

            int row = 4;

            ReportController.CopyRange(wsa, ws2, 0, 2, 0, 21, 0, 2, 0, 21);

            ws2.Cells[0, 1].Value = DtH01H.Rows[0]["Title"].ToString();
            ws2.Cells[1, 3].Value = DtH01H.Rows[0]["TITLE1"].ToString();
            ws2.Cells[1, 7].Value = DtH01H.Rows[0]["TITLE2"].ToString();

            for (int i = 0; i < DtH01.Rows.Count; i++)
            {
                ReportController.CopyRange(wsa, ws2, 4, 7, 0, 21, row, row + 3, 0, 21);
                row = row + 5;
            }

            ReportController.CopyRange(wsa, ws2, 9, 12, 0, 21, row, row + 3, 0, 21);

            row = 4;

            for (int i = 0; i < DtH01.Rows.Count; i++)
            {
                double CntC11 = ReportController.ObjectToDouble(DtH01.Rows[i]["CntC11"].ToString());
                double CntP11 = ReportController.ObjectToDouble(DtH01.Rows[i]["CntP11"].ToString());
                double CntP01 = ReportController.ObjectToDouble(DtH01.Rows[i]["CntP01"].ToString());
                double CntC01 = ReportController.ObjectToDouble(DtH01.Rows[i]["CntC01"].ToString());
                double FundC11 = ReportController.ObjectToDouble(DtH01.Rows[i]["FundC11"].ToString());
                double FundP11 = ReportController.ObjectToDouble(DtH01.Rows[i]["FundP11"].ToString());
                double FundP01 = ReportController.ObjectToDouble(DtH01.Rows[i]["FundP01"].ToString());
                double FundC01 = ReportController.ObjectToDouble(DtH01.Rows[i]["FundC01"].ToString());
                double CntC22 = ReportController.ObjectToDouble(DtH01.Rows[i]["CntC22"].ToString());
                double CntP22 = ReportController.ObjectToDouble(DtH01.Rows[i]["CntP22"].ToString());
                double CntP02 = ReportController.ObjectToDouble(DtH01.Rows[i]["CntP02"].ToString());
                double CntC02 = ReportController.ObjectToDouble(DtH01.Rows[i]["CntC02"].ToString());
                double FundC22 = ReportController.ObjectToDouble(DtH01.Rows[i]["FundC22"].ToString());
                double FundP22 = ReportController.ObjectToDouble(DtH01.Rows[i]["FundP22"].ToString());
                double FundP02 = ReportController.ObjectToDouble(DtH01.Rows[i]["FundP02"].ToString());
                double FundC02 = ReportController.ObjectToDouble(DtH01.Rows[i]["FundC02"].ToString());
                double M_01 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_01"].ToString());
                double M_02 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_02"].ToString());
                double M_03 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_03"].ToString());
                double M_04 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_04"].ToString());
                double M_05 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_05"].ToString());
                double M_06 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_06"].ToString());
                double M_07 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_07"].ToString());
                double M_08 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_08"].ToString());
                double M_09 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_09"].ToString());
                double M_10 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_10"].ToString());
                double M_11 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_11"].ToString());
                double M_12 = ReportController.ObjectToDouble(DtH01.Rows[i]["M_12"].ToString());
                ws2.Cells[row, 0].Value = DtH01.Rows[i]["CITY"].ToString();
                ws2.Cells[row, 3].Value = CntP11; //修改規劃類和工程類資料相反問題 //-k
                ws2.Cells[row, 4].Value = CntC11; //修改規劃類和工程類資料相反問題 //-k
                ws2.Cells[row, 5].Value = CntP11 + CntC11;
                ws2.Cells[row, 6].Value = FundP11 + FundC11 + FundP22 + FundC22;
                ws2.Cells[row, 7].Value = CntP01;
                ws2.Cells[row, 8].Value = CntC01;
                ws2.Cells[row, 9].Value = CntP01 + CntC01;
                ws2.Cells[row, 10].Value = FundP01 + FundC01 + FundP02 + FundC02;
                ws2.Cells[row, 11].Value = FundP01 + FundC01 + FundP02 + FundC02 - (FundP11 + FundC11 + FundP22 + FundC22); //修改中央款增減數算法 //-k
                ws2.Cells[row, 12].Value = FundP01 + FundC01;
                ws2.Cells[row, 13].Value = FundP01 + FundC01 - (FundP11 + FundC11);
                ws2.Cells[row, 14].Value = FundP02 + FundC02;
                ws2.Cells[row, 15].Value = FundP02 + FundC02 - (FundP22 + FundC22);
                ws2.Cells[row, 16].Value = M_01;
                ws2.Cells[row, 17].Value = M_02;
                ws2.Cells[row, 18].Value = M_03;
                ws2.Cells[row, 19].Value = M_04;
                ws2.Cells[row, 20].Value = M_05;
                ws2.Cells[row, 21].Value = M_06;
                row++;
                ws2.Cells[row, 3].Value = FundP11;
                ws2.Cells[row, 4].Value = FundC11;
                ws2.Cells[row, 5].Value = FundP11 + FundC11;
                ws2.Cells[row, 7].Value = FundP01;
                ws2.Cells[row, 8].Value = FundC01;
                ws2.Cells[row, 9].Value = FundP01 + FundC01;
                row++;
                ws2.Cells[row, 3].Value = CntP22;
                ws2.Cells[row, 4].Value = CntC22;
                ws2.Cells[row, 5].Value = CntP22 + CntC22;
                ws2.Cells[row, 7].Value = CntP02;
                ws2.Cells[row, 8].Value = CntC02;
                ws2.Cells[row, 9].Value = CntP02 + CntC02;
                ws2.Cells[row, 16].Value = M_07;
                ws2.Cells[row, 17].Value = M_08;
                ws2.Cells[row, 18].Value = M_09;
                ws2.Cells[row, 19].Value = M_10;
                ws2.Cells[row, 20].Value = M_11;
                ws2.Cells[row, 21].Value = M_12;
                row++;
                ws2.Cells[row, 3].Value = FundP22;
                ws2.Cells[row, 4].Value = FundC22;
                ws2.Cells[row, 5].Value = FundP22 + FundC22;
                ws2.Cells[row, 7].Value = FundP02;
                ws2.Cells[row, 8].Value = FundC02;
                ws2.Cells[row, 9].Value = FundP02 + FundC02;
                row = row + 2;
            }

            double SumCntC11 = 0;
            double SumCntP11 = 0;
            double SumCntP01 = 0;
            double SumCntC01 = 0;
            double SumFundC11 = 0;
            double SumFundP11 = 0;
            double SumFundP01 = 0;
            double SumFundC01 = 0;
            double SumCntC22 = 0;
            double SumCntP22 = 0;
            double SumCntP02 = 0;
            double SumCntC02 = 0;
            double SumFundC22 = 0;
            double SumFundP22 = 0;
            double SumFundP02 = 0;
            double SumFundC02 = 0;
            double SumM_01 = 0;
            double SumM_02 = 0;
            double SumM_03 = 0;
            double SumM_04 = 0;
            double SumM_05 = 0;
            double SumM_06 = 0;
            double SumM_07 = 0;
            double SumM_08 = 0;
            double SumM_09 = 0;
            double SumM_10 = 0;
            double SumM_11 = 0;
            double SumM_12 = 0;

            for (int i = 0; i < DtH01.Rows.Count; i++)
            {
                SumCntC11 += ReportController.ObjectToDouble(DtH01.Rows[i]["CntC11"].ToString());
                SumCntP11 += ReportController.ObjectToDouble(DtH01.Rows[i]["CntP11"].ToString());
                SumCntP01 += ReportController.ObjectToDouble(DtH01.Rows[i]["CntP01"].ToString());
                SumCntC01 += ReportController.ObjectToDouble(DtH01.Rows[i]["CntC01"].ToString());
                SumFundC11 += ReportController.ObjectToDouble(DtH01.Rows[i]["FundC11"].ToString());
                SumFundP11 += ReportController.ObjectToDouble(DtH01.Rows[i]["FundP11"].ToString());
                SumFundP01 += ReportController.ObjectToDouble(DtH01.Rows[i]["FundP01"].ToString());
                SumFundC01 += ReportController.ObjectToDouble(DtH01.Rows[i]["FundC01"].ToString());
                SumCntC22 += ReportController.ObjectToDouble(DtH01.Rows[i]["CntC22"].ToString());
                SumCntP22 += ReportController.ObjectToDouble(DtH01.Rows[i]["CntP22"].ToString());
                SumCntP02 += ReportController.ObjectToDouble(DtH01.Rows[i]["CntP02"].ToString());
                SumCntC02 += ReportController.ObjectToDouble(DtH01.Rows[i]["CntC02"].ToString());
                SumFundC22 += ReportController.ObjectToDouble(DtH01.Rows[i]["FundC22"].ToString());
                SumFundP22 += ReportController.ObjectToDouble(DtH01.Rows[i]["FundP22"].ToString());
                SumFundP02 += ReportController.ObjectToDouble(DtH01.Rows[i]["FundP02"].ToString());
                SumFundC02 += ReportController.ObjectToDouble(DtH01.Rows[i]["FundC02"].ToString());
                SumM_01 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_01"].ToString());
                SumM_02 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_02"].ToString());
                SumM_03 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_03"].ToString());
                SumM_04 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_04"].ToString());
                SumM_05 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_05"].ToString());
                SumM_06 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_06"].ToString());
                SumM_07 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_07"].ToString());
                SumM_08 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_08"].ToString());
                SumM_09 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_09"].ToString());
                SumM_10 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_10"].ToString());
                SumM_11 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_11"].ToString());
                SumM_12 += ReportController.ObjectToDouble(DtH01.Rows[i]["M_12"].ToString());
            }

            ws2.Cells[row, 3].Value = SumCntP11; //修改規劃類和工程類資料相反問題 //-k
            ws2.Cells[row, 4].Value = SumCntC11; //修改規劃類和工程類資料相反問題 //-k
            ws2.Cells[row, 5].Value = SumCntP11 + SumCntC11;
            ws2.Cells[row, 6].Value = SumFundP11 + SumFundC11 + SumFundP22 + SumFundC22;
            ws2.Cells[row, 7].Value = SumCntP01;
            ws2.Cells[row, 8].Value = SumCntC01;
            ws2.Cells[row, 9].Value = SumCntP01 + SumCntC01;
            ws2.Cells[row, 10].Value = SumFundP01 + SumFundC01 + SumFundP02 + SumFundC02;
            ws2.Cells[row, 11].Value = SumFundP01 + SumFundC01 + SumFundP02 + SumFundC02 - (SumFundP11 + SumFundC11 + SumFundP22 + SumFundC22); //修改中央款增減數算法 //-k
            ws2.Cells[row, 12].Value = SumFundP01 + SumFundC01;
            ws2.Cells[row, 13].Value = SumFundP01 + SumFundC01 - (SumFundP11 + SumFundC11);
            ws2.Cells[row, 14].Value = SumFundP02 + SumFundC02;
            ws2.Cells[row, 15].Value = SumFundP02 + SumFundC02 - (SumFundP22 + SumFundC22);
            ws2.Cells[row, 16].Value = SumM_01;
            ws2.Cells[row, 17].Value = SumM_02;
            ws2.Cells[row, 18].Value = SumM_03;
            ws2.Cells[row, 19].Value = SumM_04;
            ws2.Cells[row, 20].Value = SumM_05;
            ws2.Cells[row, 21].Value = SumM_06;
            row++;

            ws2.Cells[row, 3].Value = SumFundP11;
            ws2.Cells[row, 4].Value = SumFundC11;
            ws2.Cells[row, 5].Value = SumFundP11 + SumFundC11;
            ws2.Cells[row, 7].Value = SumFundP01;
            ws2.Cells[row, 8].Value = SumFundC01;
            ws2.Cells[row, 9].Value = SumFundP01 + SumFundC01;
            row++;

            ws2.Cells[row, 3].Value = SumCntP22;
            ws2.Cells[row, 4].Value = SumCntC22;
            ws2.Cells[row, 5].Value = SumCntP22 + SumCntC22;
            ws2.Cells[row, 7].Value = SumCntP02;
            ws2.Cells[row, 8].Value = SumCntC02;
            ws2.Cells[row, 9].Value = SumCntP02 + SumCntC02;
            ws2.Cells[row, 16].Value = SumM_07;
            ws2.Cells[row, 17].Value = SumM_08;
            ws2.Cells[row, 18].Value = SumM_09;
            ws2.Cells[row, 19].Value = SumM_10;
            ws2.Cells[row, 20].Value = SumM_11;
            ws2.Cells[row, 21].Value = SumM_12;
            row++;

            ws2.Cells[row, 3].Value = SumFundP22;
            ws2.Cells[row, 4].Value = SumFundC22;
            ws2.Cells[row, 5].Value = SumFundP22 + SumFundC22;
            ws2.Cells[row, 7].Value = SumFundP02;
            ws2.Cells[row, 8].Value = SumFundC02;
            ws2.Cells[row, 9].Value = SumFundP02 + SumFundC02;

            #region Excel 格式設定

            /// 設定內容置中
            Style style = workbook.CreateStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;

            StyleFlag flag = new StyleFlag();
            flag.HorizontalAlignment = true;

            ws2.Cells.Columns[1].ApplyStyle(style, flag); // B
            ws2.Cells.Columns[2].ApplyStyle(style, flag); // C

            ws2.Cells.SetColumnWidth(0, 10);  // A 項次
            ws2.Cells.SetColumnWidth(1, 10);  // B 縣市別
            ws2.Cells.SetColumnWidth(2, 16);  // C 
            ws2.Cells.SetColumnWidth(3, 16);  // D
            ws2.Cells.SetColumnWidth(4, 16);  // E
            ws2.Cells.SetColumnWidth(5, 16);  // F
            ws2.Cells.SetColumnWidth(6, 16);  // G 
            ws2.Cells.SetColumnWidth(7, 16);  // H
            ws2.Cells.SetColumnWidth(8, 16);  // I
            ws2.Cells.SetColumnWidth(9, 16);  // J
            ws2.Cells.SetColumnWidth(10, 16); // K
            ws2.Cells.SetColumnWidth(11, 16); // L
            ws2.Cells.SetColumnWidth(12, 16); // M
            ws2.Cells.SetColumnWidth(13, 16); // N
            ws2.Cells.SetColumnWidth(14, 16); // O
            ws2.Cells.SetColumnWidth(15, 16); // P
            ws2.Cells.SetColumnWidth(16, 16); // Q
            ws2.PageSetup.Zoom = 70;

            #endregion

            /// 20240823 功能修改 => 調整列印窄邊界/列印範圍1頁寬

            var pageSetup2 = ws2.PageSetup;

            /// 設定列印窄邊界
            pageSetup2.HeaderMargin = 0.76d;
            pageSetup2.FooterMargin = 0.76d;
            pageSetup2.LeftMargin = 0.64d;
            pageSetup2.RightMargin = 0.64d;
            pageSetup2.TopMargin = 1.91d;
            pageSetup2.BottomMargin = 1.91d;

            /// 列印範圍調整成1頁寬
            pageSetup2.FitToPagesWide = 1;
            pageSetup2.FitToPagesTall = 0;

            #endregion

            #region -------------------- 總表(BOT) --------------------

            row = 4;

            ReportController.CopyRange(wsb, ws3, 0, 2, 0, 25, 0, 2, 0, 25);

            ws3.Cells[0, 1].Value = DtH04H.Rows[0]["Title"].ToString();
            ws3.Cells[1, 3].Value = DtH04H.Rows[0]["TITLE1"].ToString();
            ws3.Cells[1, 9].Value = DtH04H.Rows[0]["TITLE2"].ToString();

            for (int i = 0; i < DtH04.Rows.Count; i++)
            {
                ReportController.CopyRange(wsb, ws3, 4, 7, 0, 25, row, row + 3, 0, 25);
                row = row + 5;
            }

            ReportController.CopyRange(wsb, ws3, 9, 12, 0, 25, row, row + 3, 0, 25);
            row = 4;

            for (int i = 0; i < DtH04.Rows.Count; i++)
            {
                double CntC11 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntC11"].ToString());
                double CntP11 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntP11"].ToString());
                double CntM11 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntM11"].ToString());
                double CntL11 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntL11"].ToString());
                double CntM01 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntM01"].ToString());
                double CntL01 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntL01"].ToString());
                double FundM11 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundM11"].ToString());
                double FundL11 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundL11"].ToString());
                double FundM01 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundM01"].ToString());
                double FundL01 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundL01"].ToString());
                double CntM22 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntM22"].ToString());
                double CntL22 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntL22"].ToString());
                double CntM02 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntM02"].ToString());
                double CntL02 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntL02"].ToString());
                double FundM22 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundM22"].ToString());
                double FundL22 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundL22"].ToString());
                double FundM02 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundM02"].ToString());
                double FundL02 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundL02"].ToString());
                double CntP01 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntP01"].ToString());
                double CntC01 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntC01"].ToString());
                double FundC11 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundC11"].ToString());
                double FundP11 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundP11"].ToString());
                double FundP01 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundP01"].ToString());
                double FundC01 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundC01"].ToString());
                double CntC22 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntC22"].ToString());
                double CntP22 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntP22"].ToString());
                double CntP02 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntP02"].ToString());
                double CntC02 = ReportController.ObjectToDouble(DtH04.Rows[i]["CntC02"].ToString());
                double FundC22 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundC22"].ToString());
                double FundP22 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundP22"].ToString());
                double FundP02 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundP02"].ToString());
                double FundC02 = ReportController.ObjectToDouble(DtH04.Rows[i]["FundC02"].ToString());
                double M_01 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_01"].ToString());
                double M_02 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_02"].ToString());
                double M_03 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_03"].ToString());
                double M_04 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_04"].ToString());
                double M_05 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_05"].ToString());
                double M_06 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_06"].ToString());
                double M_07 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_07"].ToString());
                double M_08 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_08"].ToString());
                double M_09 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_09"].ToString());
                double M_10 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_10"].ToString());
                double M_11 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_11"].ToString());
                double M_12 = ReportController.ObjectToDouble(DtH04.Rows[i]["M_12"].ToString());

                ws3.Cells[row, 0].Value = DtH04.Rows[i]["CITY"].ToString();
                ws3.Cells[row, 3].Value = CntP11;
                ws3.Cells[row, 4].Value = CntC11;
                ws3.Cells[row, 5].Value = CntM11;
                ws3.Cells[row, 6].Value = CntL11;
                ws3.Cells[row, 7].Value = CntP11 + CntC11 + CntM11 + CntL11;
                ws3.Cells[row, 8].Value = FundP11 + FundC11 + FundM11 + FundL11 + FundP22 + FundC22 + FundM22 + FundL22;
                ws3.Cells[row, 9].Value = CntP01;
                ws3.Cells[row, 10].Value = CntC01;
                ws3.Cells[row, 11].Value = CntM01;
                ws3.Cells[row, 12].Value = CntL01;
                ws3.Cells[row, 13].Value = CntP01 + CntC01 + CntM01 + CntL01;
                ws3.Cells[row, 14].Value = FundP01 + FundC01 + FundM01 + FundL01 + FundP02 + FundC02 + FundM02 + FundL02;
                ws3.Cells[row, 15].Value = FundP01 + FundC01 + FundM01 + FundL01 + FundP02 + FundC02 + FundM02 + FundL02 - (FundP11 + FundC11 + FundM11 + FundL11 + FundP22 + FundC22 + FundM22 + FundL22);
                ws3.Cells[row, 16].Value = FundP01 + FundC01 + FundM01 + FundL01;
                ws3.Cells[row, 17].Value = FundP01 + FundC01 + FundM01 + FundL01 - (FundP11 + FundC11 + FundM11 + FundL11);
                ws3.Cells[row, 18].Value = FundP02 + FundC02 + FundM02 + FundL02;
                ws3.Cells[row, 19].Value = FundP02 + FundC02 + FundM02 + FundL02 - (FundP22 + FundC22 + FundM22 + FundL22);
                ws3.Cells[row, 20].Value = M_01;
                ws3.Cells[row, 21].Value = M_02;
                ws3.Cells[row, 22].Value = M_03;
                ws3.Cells[row, 23].Value = M_04;
                ws3.Cells[row, 24].Value = M_05;
                ws3.Cells[row, 25].Value = M_06;
                row++;

                ws3.Cells[row, 3].Value = FundP11;
                ws3.Cells[row, 4].Value = FundC11;
                ws3.Cells[row, 5].Value = FundM11;
                ws3.Cells[row, 6].Value = FundL11;
                ws3.Cells[row, 7].Value = FundP11 + FundC11 + FundM11 + FundL11;
                ws3.Cells[row, 9].Value = FundP01;
                ws3.Cells[row, 10].Value = FundC01;
                ws3.Cells[row, 11].Value = FundM01;
                ws3.Cells[row, 12].Value = FundL01;
                ws3.Cells[row, 13].Value = FundP01 + FundC01 + FundM01 + FundL01;
                row++;

                ws3.Cells[row, 3].Value = CntP22;
                ws3.Cells[row, 4].Value = CntC22;
                ws3.Cells[row, 5].Value = CntM22;
                ws3.Cells[row, 6].Value = CntL22;
                ws3.Cells[row, 7].Value = CntP22 + CntC22 + CntM22 + CntL22;
                ws3.Cells[row, 9].Value = CntP02;
                ws3.Cells[row, 10].Value = CntC02;
                ws3.Cells[row, 11].Value = CntM02;
                ws3.Cells[row, 12].Value = CntL02;
                ws3.Cells[row, 13].Value = CntP02 + CntC02 + CntM02 + CntL02;
                ws3.Cells[row, 20].Value = M_07;
                ws3.Cells[row, 21].Value = M_08;
                ws3.Cells[row, 22].Value = M_09;
                ws3.Cells[row, 23].Value = M_10;
                ws3.Cells[row, 24].Value = M_11;
                ws3.Cells[row, 25].Value = M_12;
                row++;

                ws3.Cells[row, 3].Value = FundP22;
                ws3.Cells[row, 4].Value = FundC22;
                ws3.Cells[row, 5].Value = FundM22;
                ws3.Cells[row, 6].Value = FundL22;
                ws3.Cells[row, 7].Value = FundP22 + FundC22 + FundM22 + FundL22;
                ws3.Cells[row, 9].Value = FundP02;
                ws3.Cells[row, 10].Value = FundC02;
                ws3.Cells[row, 11].Value = FundM02;
                ws3.Cells[row, 12].Value = FundL02;
                ws3.Cells[row, 13].Value = FundP02 + FundC02 + FundM02 + FundL02;
                row = row + 2;
            }

            double Sum3CntC11 = 0;
            double Sum3CntP11 = 0;
            double Sum3CntM11 = 0;
            double Sum3CntL11 = 0;
            double Sum3CntM01 = 0;
            double Sum3CntL01 = 0;
            double Sum3FundM11 = 0;
            double Sum3FundL11 = 0;
            double Sum3FundM01 = 0;
            double Sum3FundL01 = 0;
            double Sum3CntM22 = 0;
            double Sum3CntL22 = 0;
            double Sum3CntM02 = 0;
            double Sum3CntL02 = 0;
            double Sum3FundM22 = 0;
            double Sum3FundL22 = 0;
            double Sum3FundM02 = 0;
            double Sum3FundL02 = 0;
            double Sum3CntP01 = 0;
            double Sum3CntC01 = 0;
            double Sum3FundC11 = 0;
            double Sum3FundP11 = 0;
            double Sum3FundP01 = 0;
            double Sum3FundC01 = 0;
            double Sum3CntC22 = 0;
            double Sum3CntP22 = 0;
            double Sum3CntP02 = 0;
            double Sum3CntC02 = 0;
            double Sum3FundC22 = 0;
            double Sum3FundP22 = 0;
            double Sum3FundP02 = 0;
            double Sum3FundC02 = 0;
            double Sum3M_01 = 0;
            double Sum3M_02 = 0;
            double Sum3M_03 = 0;
            double Sum3M_04 = 0;
            double Sum3M_05 = 0;
            double Sum3M_06 = 0;
            double Sum3M_07 = 0;
            double Sum3M_08 = 0;
            double Sum3M_09 = 0;
            double Sum3M_10 = 0;
            double Sum3M_11 = 0;
            double Sum3M_12 = 0;

            for (int i = 0; i < DtH04.Rows.Count; i++)
            {
                Sum3CntC11 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntC11"].ToString());
                Sum3CntP11 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntP11"].ToString());
                Sum3CntM11 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntM11"].ToString());
                Sum3CntL11 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntL11"].ToString());
                Sum3CntM01 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntM01"].ToString());
                Sum3CntL01 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntL01"].ToString());
                Sum3FundM11 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundM11"].ToString());
                Sum3FundL11 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundL11"].ToString());
                Sum3FundM01 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundM01"].ToString());
                Sum3FundL01 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundL01"].ToString());
                Sum3CntM22 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntM22"].ToString());
                Sum3CntL22 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntL22"].ToString());
                Sum3CntM02 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntM02"].ToString());
                Sum3CntL02 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntL02"].ToString());
                Sum3FundM22 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundM22"].ToString());
                Sum3FundL22 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundL22"].ToString());
                Sum3FundM02 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundM02"].ToString());
                Sum3FundL02 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundL02"].ToString());
                Sum3CntP01 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntP01"].ToString());
                Sum3CntC01 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntC01"].ToString());
                Sum3FundC11 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundC11"].ToString());
                Sum3FundP11 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundP11"].ToString());
                Sum3FundP01 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundP01"].ToString());
                Sum3FundC01 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundC01"].ToString());
                Sum3CntC22 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntC22"].ToString());
                Sum3CntP22 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntP22"].ToString());
                Sum3CntP02 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntP02"].ToString());
                Sum3CntC02 += ReportController.ObjectToDouble(DtH04.Rows[i]["CntC02"].ToString());
                Sum3FundC22 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundC22"].ToString());
                Sum3FundP22 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundP22"].ToString());
                Sum3FundP02 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundP02"].ToString());
                Sum3FundC02 += ReportController.ObjectToDouble(DtH04.Rows[i]["FundC02"].ToString());
                Sum3M_01 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_01"].ToString());
                Sum3M_02 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_02"].ToString());
                Sum3M_03 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_03"].ToString());
                Sum3M_04 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_04"].ToString());
                Sum3M_05 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_05"].ToString());
                Sum3M_06 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_06"].ToString());
                Sum3M_07 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_07"].ToString());
                Sum3M_08 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_08"].ToString());
                Sum3M_09 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_09"].ToString());
                Sum3M_10 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_10"].ToString());
                Sum3M_11 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_11"].ToString());
                Sum3M_12 += ReportController.ObjectToDouble(DtH04.Rows[i]["M_12"].ToString());
            }

            ws3.Cells[row, 3].Value = Sum3CntP11;
            ws3.Cells[row, 4].Value = Sum3CntC11;
            ws3.Cells[row, 5].Value = Sum3CntM11;
            ws3.Cells[row, 6].Value = Sum3CntL11;
            ws3.Cells[row, 7].Value = Sum3CntP11 + Sum3CntC11 + Sum3CntM11 + Sum3CntL11;
            ws3.Cells[row, 8].Value = Sum3FundP11 + Sum3FundC11 + Sum3FundM11 + Sum3FundL11 + Sum3FundP22 + Sum3FundC22 + Sum3FundM22 + Sum3FundL22;
            ws3.Cells[row, 9].Value = Sum3CntP01;
            ws3.Cells[row, 10].Value = Sum3CntC01;
            ws3.Cells[row, 11].Value = Sum3CntM01;
            ws3.Cells[row, 12].Value = Sum3CntL01;
            ws3.Cells[row, 13].Value = Sum3CntP01 + Sum3CntC01 + Sum3CntM01 + Sum3CntL01;
            ws3.Cells[row, 14].Value = Sum3FundP01 + Sum3FundC01 + Sum3FundM01 + Sum3FundL01 + Sum3FundP02 + Sum3FundC02 + Sum3FundM02 + Sum3FundL02;
            ws3.Cells[row, 15].Value = Sum3FundP01 + Sum3FundC01 + Sum3FundM01 + Sum3FundL01 + Sum3FundP02 + Sum3FundC02 + Sum3FundM02 + Sum3FundL02 - (Sum3FundP11 + Sum3FundC11 + Sum3FundM11 + Sum3FundL11 + Sum3FundP22 + Sum3FundC22 + Sum3FundM22 + Sum3FundL22);
            ws3.Cells[row, 16].Value = Sum3FundP01 + Sum3FundC01 + Sum3FundM01 + Sum3FundL01;
            ws3.Cells[row, 17].Value = Sum3FundP01 + Sum3FundC01 + Sum3FundM01 + Sum3FundL01 - (Sum3FundP11 + Sum3FundC11 + Sum3FundM11 + Sum3FundL11);
            ws3.Cells[row, 18].Value = Sum3FundP02 + Sum3FundC02 + Sum3FundM02 + Sum3FundL02;
            ws3.Cells[row, 19].Value = Sum3FundP02 + Sum3FundC02 + Sum3FundM02 + Sum3FundL02 - (Sum3FundP22 + Sum3FundC22 + Sum3FundM22 + Sum3FundL22);
            ws3.Cells[row, 20].Value = Sum3M_01;
            ws3.Cells[row, 21].Value = Sum3M_02;
            ws3.Cells[row, 22].Value = Sum3M_03;
            ws3.Cells[row, 23].Value = Sum3M_04;
            ws3.Cells[row, 24].Value = Sum3M_05;
            ws3.Cells[row, 25].Value = Sum3M_06;
            row++;

            ws3.Cells[row, 3].Value = Sum3FundP11;
            ws3.Cells[row, 4].Value = Sum3FundC11;
            ws3.Cells[row, 5].Value = Sum3FundM11;
            ws3.Cells[row, 6].Value = Sum3FundL11;
            ws3.Cells[row, 7].Value = Sum3FundP11 + Sum3FundC11 + Sum3FundM11 + Sum3FundL11;
            ws3.Cells[row, 9].Value = Sum3FundP01;
            ws3.Cells[row, 10].Value = Sum3FundC01;
            ws3.Cells[row, 11].Value = Sum3FundM01;
            ws3.Cells[row, 12].Value = Sum3FundL01;
            ws3.Cells[row, 13].Value = Sum3FundP01 + Sum3FundC01 + Sum3FundM01 + Sum3FundL01;
            row++;

            ws3.Cells[row, 3].Value = Sum3CntP22;
            ws3.Cells[row, 4].Value = Sum3CntC22;
            ws3.Cells[row, 5].Value = Sum3CntM22;
            ws3.Cells[row, 6].Value = Sum3CntL22;
            ws3.Cells[row, 7].Value = Sum3CntP22 + Sum3CntC22 + Sum3CntM22 + Sum3CntL22;
            ws3.Cells[row, 9].Value = Sum3CntP02;
            ws3.Cells[row, 10].Value = Sum3CntC02;
            ws3.Cells[row, 11].Value = Sum3CntM02;
            ws3.Cells[row, 12].Value = Sum3CntL02;
            ws3.Cells[row, 13].Value = Sum3CntP02 + Sum3CntC02 + Sum3CntM02 + Sum3CntL02;
            ws3.Cells[row, 20].Value = Sum3M_07;
            ws3.Cells[row, 21].Value = Sum3M_08;
            ws3.Cells[row, 22].Value = Sum3M_09;
            ws3.Cells[row, 23].Value = Sum3M_10;
            ws3.Cells[row, 24].Value = Sum3M_11;
            ws3.Cells[row, 25].Value = Sum3M_12;
            row++;

            ws3.Cells[row, 3].Value = Sum3FundP22;
            ws3.Cells[row, 4].Value = Sum3FundC22;
            ws3.Cells[row, 5].Value = Sum3FundM22;
            ws3.Cells[row, 6].Value = Sum3FundL22;
            ws3.Cells[row, 7].Value = Sum3FundP22 + Sum3FundC22 + Sum3FundM22 + Sum3FundL22;
            ws3.Cells[row, 9].Value = Sum3FundP02;
            ws3.Cells[row, 10].Value = Sum3FundC02;
            ws3.Cells[row, 11].Value = Sum3FundM02;
            ws3.Cells[row, 12].Value = Sum3FundL02;
            ws3.Cells[row, 13].Value = Sum3FundP02 + Sum3FundC02 + Sum3FundM02 + Sum3FundL02;

            ws3.AutoFitColumns();

            /// 20240823 功能修改 => 調整列印窄邊界/列印範圍1頁寬

            var pageSetup3 = ws3.PageSetup;

            /// 設定列印窄邊界
            pageSetup3.HeaderMargin = 1.91d;
            pageSetup3.FooterMargin = 1.91d;
            pageSetup3.LeftMargin = 0.64d;
            pageSetup3.RightMargin = 0.64d;
            pageSetup3.TopMargin = 0.76d;
            pageSetup3.BottomMargin = 0.76d;

            /// 列印範圍調整成1頁寬
            pageSetup3.FitToPagesWide = 1;
            pageSetup3.FitToPagesTall = 0;

            #endregion

            #region -------------------- 總表 --------------------

            ReportController.CopyRange(wsc, ws1, 0, 2, 0, 16, 0, 2, 0, 16);
            ws1.Cells[0, 0].Value = DtH00H.Rows[0]["Title"].ToString();
            ws1.Cells[1, 2].Value = DtH00H.Rows[0]["TITLE1"].ToString();
            ws1.Cells[1, 7].Value = DtH00H.Rows[0]["TITLE2"].ToString(); /// 7 ~ 8

            row = 3;

            for (int i = 0; i < DtH00.Rows.Count; i++)
            {
                ReportController.CopyRange(wsc, ws1, 3, 4, 0, 16, row, row + 1, 0, 16);
                row = row + 2;
            }

            ReportController.CopyRange(wsc, ws1, 6, 9, 0, 16, row, row + 3, 0, 16);

            ws1.Cells[0, 0].Value = DtH00H.Rows[0]["Title"].ToString();
            ws1.Cells[1, 2].Value = DtH00H.Rows[0]["TITLE1"].ToString();
            ws1.Cells[1, 7].Value = DtH00H.Rows[0]["TITLE2"].ToString(); /// 7 ~ 8
            row = 3;

            double FEE_G_1_1 = 0;
            double FEE_G_1_2 = 0;
            double FEE_G_2_1 = 0;
            double FEE_G_2_2 = 0;
            double FEE_B_1_1 = 0;
            double FEE_B_1_2 = 0;
            double FEE_B_2_1 = 0;
            double FEE_B_2_2 = 0;

            /// 20241017 修改項目 => 業務費用
            double FEE_G_1_3 = 0;
            double FEE_G_2_3 = 0;
            double FEE_B_1_3 = 0;
            double FEE_B_2_3 = 0;


            double ACC02_FUND_FINAL_SUM = ACC02_FUND_FINAL;
            double ACC02_FUND_SECOND_SUM = ACC02_FUND_SECOND;

            for (int i = 0; i < DtH00.Rows.Count; i++)
            {
                if (i != 0)
                {
                    ACC02_FUND_FINAL = 0;
                    ACC02_FUND_SECOND = 0;
                }

                FEE_G_1_1 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_1_1"].ToString());
                FEE_G_1_2 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_1_2"].ToString());
                FEE_G_2_1 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_2_1"].ToString());
                FEE_G_2_2 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_2_2"].ToString());
                FEE_B_1_1 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_1_1"].ToString());
                FEE_B_1_2 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_1_2"].ToString());
                FEE_B_2_1 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_2_1"].ToString());
                FEE_B_2_2 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_2_2"].ToString());

                /// 20241017 修改項目 => 業務費用
                FEE_G_1_3 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_1_3"].ToString());
                FEE_G_2_3 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_2_3"].ToString());
                FEE_B_1_3 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_1_3"].ToString());
                FEE_B_2_3 = ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_2_3"].ToString());
                ACC02_FUND_SECOND = FEE_G_1_3;
                ACC02_FUND_FINAL = FEE_G_2_3;


                ws1.Cells[row, 0].Value = DtH00.Rows[i]["CITY"].ToString();
                ws1.Cells[row, 2].Value = FEE_G_1_1;
                ws1.Cells[row, 3].Value = FEE_G_1_2;
                ws1.Cells[row, 4].Value = ACC02_FUND_SECOND;
                ws1.Cells[row, 5].Value = FEE_G_1_1 + FEE_G_1_2 + ACC02_FUND_SECOND;
                ws1.Cells[row, 6].Value = FEE_G_1_1 + FEE_G_1_2 + FEE_B_1_1 + FEE_B_1_2 + ACC02_FUND_SECOND;
                ws1.Cells[row, 7].Value = FEE_G_2_1;
                ws1.Cells[row, 8].Value = FEE_G_2_2;
                ws1.Cells[row, 9].Value = ACC02_FUND_FINAL;
                ws1.Cells[row, 10].Value = FEE_G_2_1 + FEE_G_2_2 + ACC02_FUND_FINAL;
                ws1.Cells[row, 11].Value = FEE_G_2_1 + FEE_G_2_2 + FEE_B_2_1 + FEE_B_2_2 + ACC02_FUND_FINAL;
                ws1.Cells[row, 12].Value = FEE_G_2_1 - FEE_G_1_1;
                ws1.Cells[row, 13].Value = FEE_G_2_2 - FEE_G_1_2;
                ws1.Cells[row, 14].Value = ACC02_FUND_FINAL - ACC02_FUND_SECOND;
                ws1.Cells[row, 15].Value = FEE_G_2_1 + FEE_G_2_2 + ACC02_FUND_FINAL - (FEE_G_1_1 + FEE_G_1_2 + ACC02_FUND_SECOND);
                ws1.Cells[row, 16].Value = FEE_G_2_1 + FEE_G_2_2 + ACC02_FUND_FINAL - (FEE_G_1_1 + FEE_G_1_2 + ACC02_FUND_SECOND) + FEE_B_2_1 + FEE_B_2_2 - (FEE_B_1_1 + FEE_B_1_2);
                row++;

                ws1.Cells[row, 2].Value = FEE_B_1_1;
                ws1.Cells[row, 3].Value = FEE_B_1_2;
                ws1.Cells[row, 4].Value = 0;  /// 預設 0
                ws1.Cells[row, 5].Value = FEE_B_1_1 + FEE_B_1_2;
                ws1.Cells[row, 7].Value = FEE_B_2_1;
                ws1.Cells[row, 8].Value = FEE_B_2_2;
                ws1.Cells[row, 9].Value = 0;  /// 預設 0
                ws1.Cells[row, 10].Value = FEE_B_2_1 + FEE_B_2_2;
                ws1.Cells[row, 12].Value = FEE_B_2_1 - FEE_B_1_1;
                ws1.Cells[row, 13].Value = FEE_B_2_2 - FEE_B_1_2;
                ws1.Cells[row, 14].Value = 0; /// 預設 0
                ws1.Cells[row, 15].Value = FEE_B_2_1 + FEE_B_2_2 - (FEE_B_1_1 + FEE_B_1_2);
                row++;
            }

            double SumFEE_G_1_1 = 0;
            double SumFEE_G_1_2 = 0;
            double SumFEE_G_2_1 = 0;
            double SumFEE_G_2_2 = 0;
            double SumFEE_B_1_1 = 0;
            double SumFEE_B_1_2 = 0;
            double SumFEE_B_2_1 = 0;
            double SumFEE_B_2_2 = 0;

            /// 20241017 修改項目 => 業務費用
            double SumFEE_G_1_3 = 0;
            double SumFEE_G_2_3 = 0;
            double SumFEE_B_1_3 = 0;
            double SumFEE_B_2_3 = 0;

            for (int i = 0; i < DtH00.Rows.Count; i++)
            {
                SumFEE_G_1_1 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_1_1"].ToString());
                SumFEE_G_1_2 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_1_2"].ToString());
                SumFEE_G_2_1 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_2_1"].ToString());
                SumFEE_G_2_2 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_2_2"].ToString());
                SumFEE_B_1_1 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_1_1"].ToString());
                SumFEE_B_1_2 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_1_2"].ToString());
                SumFEE_B_2_1 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_2_1"].ToString());
                SumFEE_B_2_2 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_2_2"].ToString());

                /// 20241017 修改項目 => 業務費用
                SumFEE_G_1_3 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_1_3"].ToString());
                SumFEE_G_2_3 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_G_2_3"].ToString());
                SumFEE_B_1_3 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_1_3"].ToString());
                SumFEE_B_2_3 += ReportController.ObjectToDouble(DtH00.Rows[i]["FEE_B_2_3"].ToString());

                ACC02_FUND_SECOND_SUM = SumFEE_G_1_3;
                ACC02_FUND_FINAL_SUM = SumFEE_G_2_3;
            }

            ws1.Cells[row, 2].Value = SumFEE_G_1_1;
            ws1.Cells[row, 3].Value = SumFEE_G_1_2;
            ws1.Cells[row, 4].Value = ACC02_FUND_SECOND_SUM;
            ws1.Cells[row, 5].Value = SumFEE_G_1_1 + SumFEE_G_1_2 + ACC02_FUND_SECOND_SUM;
            ws1.Cells[row, 6].Value = SumFEE_G_1_1 + SumFEE_G_1_2 + SumFEE_B_1_1 + SumFEE_B_1_2 + ACC02_FUND_SECOND_SUM;
            ws1.Cells[row, 7].Value = SumFEE_G_2_1;
            ws1.Cells[row, 8].Value = SumFEE_G_2_2;
            ws1.Cells[row, 9].Value = ACC02_FUND_FINAL_SUM;
            ws1.Cells[row, 10].Value = SumFEE_G_2_1 + SumFEE_G_2_2 + ACC02_FUND_FINAL_SUM;
            ws1.Cells[row, 11].Value = SumFEE_G_2_1 + SumFEE_G_2_2 + SumFEE_B_2_1 + SumFEE_B_2_2 + ACC02_FUND_FINAL_SUM;
            ws1.Cells[row, 12].Value = SumFEE_G_2_1 - SumFEE_G_1_1;
            ws1.Cells[row, 13].Value = SumFEE_G_2_2 - SumFEE_G_1_2;
            ws1.Cells[row, 14].Value = ACC02_FUND_FINAL_SUM - ACC02_FUND_SECOND_SUM;
            ws1.Cells[row, 15].Value = SumFEE_G_2_1 + SumFEE_G_2_2 + ACC02_FUND_FINAL_SUM - (SumFEE_G_1_1 + SumFEE_G_1_2 + ACC02_FUND_SECOND_SUM);
            ws1.Cells[row, 16].Value = SumFEE_G_2_1 + SumFEE_G_2_2 + ACC02_FUND_FINAL_SUM - (SumFEE_G_1_1 + SumFEE_G_1_2 + ACC02_FUND_SECOND_SUM) + SumFEE_B_2_1 + SumFEE_B_2_2 - (SumFEE_B_1_1 + SumFEE_B_1_2);
            row++;

            ws1.Cells[row, 2].Value = SumFEE_B_1_1;
            ws1.Cells[row, 3].Value = SumFEE_B_1_2;
            ws1.Cells[row, 4].Value = 0;
            ws1.Cells[row, 5].Value = SumFEE_B_1_1 + SumFEE_B_1_2;
            ws1.Cells[row, 7].Value = SumFEE_B_2_1;
            ws1.Cells[row, 8].Value = SumFEE_B_2_2;
            ws1.Cells[row, 9].Value = 0;
            ws1.Cells[row, 10].Value = SumFEE_B_2_1 + SumFEE_B_2_2;
            ws1.Cells[row, 12].Value = SumFEE_B_2_1 - SumFEE_B_1_1;
            ws1.Cells[row, 13].Value = SumFEE_B_2_2 - SumFEE_B_1_2;
            ws1.Cells[row, 14].Value = 0;
            ws1.Cells[row, 15].Value = SumFEE_B_2_1 + SumFEE_B_2_2 - (SumFEE_B_1_1 + SumFEE_B_1_2);
            row++;

            double FEE_1_1 = ReportController.ObjectToDouble(DtH00H.Rows[0]["FEE_1_1"].ToString());
            double FEE_1_2 = ReportController.ObjectToDouble(DtH00H.Rows[0]["FEE_1_2"].ToString());
            double FEE_2_1 = ReportController.ObjectToDouble(DtH00H.Rows[0]["FEE_2_1"].ToString());
            double FEE_2_2 = ReportController.ObjectToDouble(DtH00H.Rows[0]["FEE_2_2"].ToString());
            double FEE_G_DIFF = ReportController.ObjectToDouble(DtH00H.Rows[0]["FEE_G_DIFF"].ToString());
            double FEE_B_DIFF = ReportController.ObjectToDouble(DtH00H.Rows[0]["FEE_B_DIFF"].ToString());

            ws1.Cells[row, 2].Value = FEE_1_1;
            ws1.Cells[row, 3].Value = FEE_1_2;
            ws1.Cells[row, 4].Value = ACC02_FUND_SECOND_SUM; /// 修改前期顯示業務費總計取得資料
            ws1.Cells[row, 5].Value = FEE_1_1 + FEE_1_2 + ACC02_FUND_SECOND_SUM;
            ws1.Cells[row, 6].Value = FEE_1_1 + FEE_1_2 + ACC02_FUND_SECOND_SUM;
            ws1.Cells[row, 7].Value = FEE_2_1;
            ws1.Cells[row, 8].Value = FEE_2_2;
            ws1.Cells[row, 9].Value = ACC02_FUND_FINAL_SUM;
            ws1.Cells[row, 10].Value = FEE_2_1 + FEE_2_2 + ACC02_FUND_FINAL_SUM;
            ws1.Cells[row, 11].Value = FEE_2_1 + FEE_2_2 + ACC02_FUND_FINAL_SUM;
            ws1.Cells[row, 12].Value = FEE_2_1 - FEE_1_1;
            ws1.Cells[row, 13].Value = FEE_2_2 - FEE_1_2;
            ws1.Cells[row, 14].Value = ACC02_FUND_FINAL_SUM - ACC02_FUND_SECOND_SUM; ;
            ws1.Cells[row, 15].Value = FEE_G_DIFF + ACC02_FUND_FINAL_SUM - ACC02_FUND_SECOND_SUM;
            ws1.Cells[row, 16].Value = FEE_G_DIFF + FEE_B_DIFF + ACC02_FUND_FINAL_SUM - ACC02_FUND_SECOND_SUM;
            row++;

            ws1.Cells[row, 15].Value = FEE_B_DIFF;

            /// 20240823 修改功能 => 總表(頁籤)格式修改為"#,##0"

            ws1.Cells.SetColumnWidth(0, 10);  // A 項次
            ws1.Cells.SetColumnWidth(1, 10);  // B 縣市別
            ws1.Cells.SetColumnWidth(2, 16);  // C 
            ws1.Cells.SetColumnWidth(3, 16);  // D
            ws1.Cells.SetColumnWidth(4, 16);  // E
            ws1.Cells.SetColumnWidth(5, 16);  // F
            ws1.Cells.SetColumnWidth(6, 16);  // G 
            ws1.Cells.SetColumnWidth(7, 16);  // H
            ws1.Cells.SetColumnWidth(8, 16);  // I
            ws1.Cells.SetColumnWidth(9, 16);  // J
            ws1.Cells.SetColumnWidth(10, 16); // K
            ws1.Cells.SetColumnWidth(11, 16); // L
            ws1.Cells.SetColumnWidth(12, 16); // M
            ws1.Cells.SetColumnWidth(13, 16); // N
            ws1.Cells.SetColumnWidth(14, 16); // O
            ws1.Cells.SetColumnWidth(15, 16); // P
            ws1.Cells.SetColumnWidth(16, 16); // Q
            ws1.PageSetup.Zoom = 70;

            /// 設定 C4 ~ Q4 所有欄位
            int StartRow = 3;
            int StartCol = CellsHelper.ColumnNameToIndex("C");
            int EndCol = CellsHelper.ColumnNameToIndex("Q");

            /// 遍歷所有 C4 ~ Q4 以下欄位
            for (int Excelrow = StartRow; Excelrow <= ws1.Cells.MaxDataRow; Excelrow++)
            {
                for (int Excelcol = StartCol; Excelcol <= EndCol; Excelcol++)
                {
                    /// 取得格式後進行修改項目
                    /// 數值格式("#,##0")
                    /// 所有欄位靠右對齊
                    Aspose.Cells.Style style1 = ws1.Cells[Excelrow, Excelcol].GetStyle();
                    style1.Custom = "#,##0";
                    style1.HorizontalAlignment = TextAlignmentType.Center;
                    ws1.Cells[Excelrow, Excelcol].SetStyle(style1);
                }
            }

            /// 20240822 修改功能 => 列印範圍設置調整成1頁寬
            var pageSetup1 = ws1.PageSetup;
            pageSetup1.FitToPagesWide = 1;
            pageSetup1.FitToPagesTall = 0;


            #endregion

            #region -------------------- 彙整明細 --------------------

            row = 0;

            var distinctCityCls = Lb1Dtback.Select(m => m.CityCls).Distinct().ToList();

            for (int i = 0; i < Lb1Dtback.Select(m => m.CityCls).Distinct().Count(); i++)
            {
                ReportController.CopyRange(wsd, ws4, 0, 25, 0, 21, row, row + 25, 0, 21);
                ws4.Cells[row, 0].Value = DtH06H.Rows[0]["Title"].ToString();
                ws4.Cells[row + 1, 2].Value = Lb1DtFooter[i].CITY.ToString() + "調整預算表";
                ws4.Cells[row + 2, 7].Value = Lb1DtFooter[i].TITLE1.ToString();
                ws4.Cells[row + 2, 10].Value = Lb1DtFooter[i].TITLE2.ToString();
                row = row + 4;

                /// 2024/10/16 修改 => 全縣市 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 22
                var list04 = Lb1Dtback.Where(m => m.CityCls == distinctCityCls[i]).OrderBy(m => m.Order).ToList();

                for (int j = 0; j < Convert.ToInt32(list04.Select(m => m.Order).Count()); j++)
                {
                    var x = (j + 1) % 11;
                    var y = 11 - x;

                    ws4.Cells[row, 0].Value = list04[j].SERIAL;
                    ws4.Cells[row, 1].Value = list04[j].SIP_TD_CODE;
                    ws4.Cells[row, 2].Value = list04[j].FD_CODE;
                    ws4.Cells[row, 3].Value = list04[j].TD_NAME;
                    ws4.Cells[row, 4].Value = list04[j].CATY_DESC;
                    ws4.Cells[row, 5].Value = list04[j].ACC_DESC;
                    ws4.Cells[row, 6].Value = list04[j].CEN_SP_RATIO;
                    ws4.Cells[row, 7].Value = ReportController.ObjectToDouble(list04[j].CEN_FUND2); // 修改格式
                    ws4.Cells[row, 8].Value = ReportController.ObjectToDouble(list04[j].LOC_FUND2); // 修改格式
                    ws4.Cells[row, 9].Value = ReportController.CalPlus(list04[j].CEN_FUND2, list04[j].LOC_FUND2);
                    ws4.Cells[row, 10].Value = ReportController.ObjectToDouble(list04[j].CEN_FUND1); // 修改格式
                    ws4.Cells[row, 11].Value = ReportController.ObjectToDouble(list04[j].LOC_FUND1); // 修改格式
                    ws4.Cells[row, 12].Value = ReportController.CalPlus(list04[j].CEN_FUND1, list04[j].LOC_FUND1);
                    ws4.Cells[row, 13].Value = ReportController.CalMinus(list04[j].CEN_FUND1, list04[j].CEN_FUND2);
                    ws4.Cells[row, 14].Value = ReportController.ObjectToDouble(list04[j].COST_FEE); // 修改格式
                    ws4.Cells[row, 15].Value = list04[j].FUND_DESC;
                    ws4.Cells[row, 16].Value = list04[j].M_01;
                    ws4.Cells[row, 17].Value = list04[j].M_02;
                    ws4.Cells[row, 18].Value = list04[j].M_03;
                    ws4.Cells[row, 19].Value = list04[j].M_04;
                    ws4.Cells[row, 20].Value = list04[j].M_05;
                    ws4.Cells[row, 21].Value = list04[j].M_06;

                    ws4.Cells[row + 1, 16].Value = list04[j].M_07;
                    ws4.Cells[row + 1, 17].Value = list04[j].M_08;
                    ws4.Cells[row + 1, 18].Value = list04[j].M_09;
                    ws4.Cells[row + 1, 19].Value = list04[j].M_10;
                    ws4.Cells[row + 1, 20].Value = list04[j].M_11;
                    ws4.Cells[row + 1, 21].Value = list04[j].M_12;
                    row = row + 2;

                    if (x == 0 && j != Convert.ToInt32(list04.Select(m => m.Order).Count()) - 1)
                    {
                        row = row + 2;
                        ReportController.CopyRange(wsd, ws4, 0, 25, 0, 21, row, row + 25, 0, 21);
                        ws4.Cells[row, 0].Value = DtH06H.Rows[0]["Title"].ToString();
                        ws4.Cells[row + 1, 2].Value = Lb1DtFooter[i].CITY.ToString() + "調整預算表";
                        ws4.Cells[row + 2, 7].Value = Lb1DtFooter[i].TITLE1.ToString();
                        ws4.Cells[row + 2, 10].Value = Lb1DtFooter[i].TITLE2.ToString();
                        row = row + 4;
                    }
                    else if (x != 0 && j == Convert.ToInt32(list04.Select(m => m.Order).Count()) - 1)
                    {
                        row = row + (y + 1) * 2;
                    }
                }
            }

            ws4.AutoFitColumns();

            /// 20240823 功能修改 => 調整列印窄邊界/列印範圍1頁寬

            var pageSetup4 = ws4.PageSetup;

            /// 設定列印窄邊界
            pageSetup4.HeaderMargin = 1.91d;
            pageSetup4.FooterMargin = 1.91d;
            pageSetup4.LeftMargin = 0.64d;
            pageSetup4.RightMargin = 0.64d;
            pageSetup4.TopMargin = 0.76d;
            pageSetup4.BottomMargin = 0.76d;

            /// 列印範圍調整成1頁寬
            pageSetup4.FitToPagesWide = 1;
            pageSetup4.FitToPagesTall = 0;

            #endregion

            #region -------------------- 彙整明細簡易版 --------------------

            row = 0;

            for (int i = 0; i < Lb1Dtback.Select(m => m.CityCls).Distinct().Count(); i++)
            {
                ReportController.CopyRange(wse, ws5, 0, 25, 0, 14, row, row + 25, 0, 14);

                ws5.Cells[row, 0].Value = DtH06H.Rows[0]["Title"].ToString();
                ws5.Cells[row + 1, 0].Value = Lb1DtFooter[i].CITY.ToString() + "調整預算表";
                ws5.Cells[row + 2, 7].Value = Lb1DtFooter[i].TITLE1.ToString();
                ws5.Cells[row + 2, 10].Value = Lb1DtFooter[i].TITLE2.ToString();
                row = row + 4;

                var list04 = Lb1Dtback.Where(m => m.CityCls == distinctCityCls[i]).OrderBy(m => m.Order).ToList();

                for (int j = 0; j < Convert.ToInt32(list04.Select(m => m.Order).Count()); j++)
                {
                    var x = (j + 1) % 11;
                    var y = 11 - x;

                    ws5.Cells[row, 0].Value = list04[j].SERIAL;
                    ws5.Cells[row, 1].Value = list04[j].SIP_TD_CODE;
                    ws5.Cells[row, 2].Value = list04[j].FD_CODE;
                    ws5.Cells[row, 3].Value = list04[j].TD_NAME;
                    ws5.Cells[row, 4].Value = list04[j].CATY_DESC;
                    ws5.Cells[row, 5].Value = list04[j].ACC_DESC;
                    ws5.Cells[row, 6].Value = list04[j].CEN_SP_RATIO;
                    ws5.Cells[row, 7].Value = ReportController.ObjectToDouble(list04[j].CEN_FUND2); // 修改格式
                    ws5.Cells[row, 8].Value = ReportController.ObjectToDouble(list04[j].LOC_FUND2); // 修改格式
                    ws5.Cells[row, 9].Value = ReportController.CalPlus(list04[j].CEN_FUND2, list04[j].LOC_FUND2);
                    ws5.Cells[row, 10].Value = ReportController.ObjectToDouble(list04[j].CEN_FUND1); // 修改格式
                    ws5.Cells[row, 11].Value = ReportController.ObjectToDouble(list04[j].LOC_FUND1); // 修改格式
                    ws5.Cells[row, 12].Value = ReportController.CalPlus(list04[j].CEN_FUND1, list04[j].LOC_FUND1);
                    ws5.Cells[row, 13].Value = ReportController.CalMinus(list04[j].CEN_FUND1, list04[j].CEN_FUND2);
                    ws5.Cells[row, 14].Value = ReportController.ObjectToDouble(list04[j].COST_FEE);  // 修改格式
                    row = row + 2;

                    /// 20240823 修改功能 => 總表(頁籤)格式修改為 "#,##0"

                    /// 設定 C4 ~ Q4 欄位
                    int ws5_StartRow = 4;
                    int ws5_StartCol = CellsHelper.ColumnNameToIndex("G");
                    int ws5_EndCol = CellsHelper.ColumnNameToIndex("O");

                    /// 遍歷 C4 ~ Q4 欄位
                    for (int ws5_row = ws5_StartRow; ws5_row <= ws5.Cells.MaxDataRow; ws5_row++)
                    {
                        for (int ws5_col = ws5_StartCol; ws5_col <= ws5_EndCol; ws5_col++)
                        {
                            object ws5_value = ws5.Cells[ws5_row, ws5_col].Value;

                            if (ws5_value != null && decimal.TryParse(ws5_value.ToString(), out decimal numericValue))
                            {
                                /// 將單元格的值轉換為數值型
                                ws5.Cells[ws5_row, ws5_col].Value = numericValue;
                            }

                            Aspose.Cells.Style style5 = ws5.Cells[ws5_row, ws5_col].GetStyle();
                            style5.Custom = "#,##0";
                            style5.HorizontalAlignment = TextAlignmentType.Right;
                            ws5.Cells[ws5_row, ws5_col].SetStyle(style5, true);
                        }
                    }

                    if (x == 0 && j != Convert.ToInt32(list04.Select(m => m.Order).Count()) - 1)
                    {
                        row = row + 2;
                        ReportController.CopyRange(wse, ws5, 0, 25, 0, 14, row, row + 25, 0, 14);
                        ws5.Cells[row, 0].Value = DtH06H.Rows[0]["Title"].ToString();
                        ws5.Cells[row + 1, 0].Value = Lb1DtFooter[i].CITY.ToString() + "調整預算表";
                        ws5.Cells[row + 2, 7].Value = Lb1DtFooter[i].TITLE1.ToString();
                        ws5.Cells[row + 2, 10].Value = Lb1DtFooter[i].TITLE2.ToString();
                        row = row + 4;
                    }
                    else if (x != 0 && j == Convert.ToInt32(list04.Select(m => m.Order).Count()) - 1)
                    {
                        row = row + (y + 1) * 2;
                    }
                }
            }

            #endregion

            ws5.AutoFitColumns();

            /// 20240823 功能修改 => 調整列印窄邊界/列印範圍1頁寬

            /// 取得 PageSetup 物件
            var pageSetup5 = ws5.PageSetup;

            /// 設定列印窄邊界
            pageSetup5.HeaderMargin = 1.91d;
            pageSetup5.FooterMargin = 1.91d;
            pageSetup5.LeftMargin = 0.64d;
            pageSetup5.RightMargin = 0.64d;
            pageSetup5.TopMargin = 0.76d;
            pageSetup5.BottomMargin = 0.76d;

            /// 將所有欄放入單一頁面
            pageSetup5.FitToPagesWide = 1;
            pageSetup5.FitToPagesTall = 0;

            for (int i = 0; i < 6; i++)
            {
                workbook.Worksheets.RemoveAt(0);
            }

            string filename = "B-1.修正預算管考報表" + DateTime.Now.ToString("yyyyMMddHHmmss");
            string directoryPath = Server.MapPath(@"~\Temp\");
            string newfilepath = System.IO.Path.Combine(directoryPath, filename + ".xlsx");
            workbook.Save(newfilepath);

            if (ODS == "1")
            {
                var OdsResult = OdsPrint(newfilepath, filename);
                return OdsResult;
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Seek(0, SeekOrigin.Begin);
                    var bytes = stream.ToArray();
                    stream.Close();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename + ".xlsx");
                }
            }
        }

        /// <summary>
        /// A-2 細項執行情形表(舊版)
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="SYR"></param>
        /// <param name="YearWrite"></param>
        /// <param name="SMN"></param>
        /// <param name="CityList"></param>
        /// <param name="funding"></param>
        /// <param name="ODS"></param>
        /// <param name="Site"></param>
        /// <param name="group"></param>
        /// <returns></returns>
        public ActionResult ReportPrint2(string PID, string SYR, string YearWrite, string SMN, string[] CityList, string funding, string ODS, string[] Site, string group)
        {
            var PIDs = PID.Split(',').Select(id => decimal.Parse(id)).ToList();
            var PJ = db2.PJMFORM.Where(m => PIDs.Contains(m.PID)).ToList();

            var path = Server.MapPath("~/Templates/PJV0010_03.xlsx");
            var file = new FileInfo(path);
            string City = "";                // 用來存儲城市名稱
            DataTable dt2 = new DataTable(); // 用來存儲查詢結果

            if (CityList == null || CityList.Length == 0)
            {
                string alertMessage = "請至少選擇一個縣市";
                return Content($"<script>alert('{alertMessage}'); window.location.href = document.referrer;</script>");
            }

            foreach (var item in CityList)
            {
                City = item;
                dt2.Merge(Source.Report_Source.getBillionProject08(PID, SYR, SMN, City, funding, connectionString));
            }

            if (group != null)
            {
                var result = db3.TDMFORM.Where(m => m.TD_GROUP == true).ToList();
                DataTable filteredData = dt2.Clone();
                var tdCodes = result.Select(r => r.TD_CODE).ToList();

                foreach (DataRow row in dt2.Rows)
                {
                    string tdCode = row["TD_CODE"].ToString();

                    if (tdCodes.Contains(tdCode))
                    {
                        filteredData.ImportRow(row);
                    }
                }

                dt2.Rows.Clear();
                foreach (DataRow row in filteredData.Rows)
                {
                    dt2.ImportRow(row);
                }
            }

            var TDV_TDDCFTS = db2.TDV_TDDCFTS;
            var TDMFORM = db3.TDMFORM;

            /// ================= Aspose.Cells ================= ///

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            var workbook = new Aspose.Cells.Workbook(path);
            var ws = workbook.Worksheets["Sheet1"];

            ws.Cells["A1"].PutValue("表2-5、內政部辦理細項執行情形");                // A1 => 標題
            ws.Cells["U2"].PutValue("資料截至：" + YearWrite + "年" + SMN + "月"); // U2 => 資料截至年月

            var a = ws.Cells.CreateRange("A5:AQ5");
            Aspose.Cells.Style style1 = ws.Cells["A6"].GetStyle();

            for (int i = 6; i < dt2.Rows.Count + 5; i++)
            {
                var b = ws.Cells.CreateRange("A" + i.ToString() + ":AQ" + i.ToString());
                b.Copy(a);
            }

            /// 加總欄位
            double B = 0;  /// 分支計畫核定經費總計(仟元)
            double I = 0;  /// 歷年累積撥發金額(仟元)
            double J = 0;  /// 當年累積撥發金額(仟元)
            double K = 0;  /// 當年度核定經費(仟元)(A)
            double L = 0;  /// 依進度可撥經費(仟元)
            double AD = 0; /// 總工程經費(仟元)

            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                int rowNumber = 5 + i;

                var TD_CODE = dt2.Rows[i]["TD_CODE"].ToString();
                var CEN_SP_RATIO = TDV_TDDCFTS.Where(m => m.TD_CODE == TD_CODE).FirstOrDefault()?.CEN_SP_RATIO ?? 0;
                var CEN_SP_RATIO_TWODECIMALPLACE = Math.Round((double)CEN_SP_RATIO, 2);
                var TD_SOURCE = TDMFORM.Where(m => m.TD_CODE == TD_CODE).FirstOrDefault();
                var TD_FINISH = TD_SOURCE.TD_FINISH;
                var TD_START_FUND = TD_SOURCE.START_FUND;
                var TD_HOUSEHOLD = TD_SOURCE.HOUSEHOLD;
                var TD_TOTAL_HOUSEHOLD = TD_SOURCE.TOTAL_HOUSEHOLD;

                ws.Cells["A" + rowNumber].PutValue(dt2.Rows[i]["CATY_NAME"].ToString()); // 計畫說明
                ws.Cells["B" + rowNumber].PutValue(Source.Tools.thousandchange3(dt2.Rows[i]["CATY_FUND_"].ToString())); // 分支計畫核定經費總計(仟元)
                Tools.SetCellstyle(ws, "B" + rowNumber);
                ws.Cells["C" + rowNumber].PutValue(i + 1); // 項次
                ws.Cells["D" + rowNumber].PutValue(dt2.Rows[i]["TD_CITY"].ToString());  // 縣市
                ws.Cells["E" + rowNumber].PutValue(dt2.Rows[i]["IPJ_DESC"].ToString()); // 辦理層級
                ws.Cells["F" + rowNumber].PutValue(CEN_SP_RATIO_TWODECIMALPLACE);       // 補助比例
                ws.Cells["G" + rowNumber].PutValue(dt2.Rows[i]["TD_NAME"].ToString());  // 工作名稱(工程標案名稱、勞務採購、其它(如用地補償…))

                if (TD_FINISH is false)
                {
                    ws.Cells["H" + rowNumber].PutValue("未結案");   // 結案情形
                }
                else if (TD_FINISH is true)
                {
                    ws.Cells["H" + rowNumber].PutValue("已結案");   // 結案情形
                }
                else
                {
                    ws.Cells["H" + rowNumber].PutValue(TD_FINISH); // 結案情形
                }

                ws.Cells["I" + rowNumber].PutValue(Source.Tools.thousandchange3(dt2.Rows[i]["ALL_PPAY_FEE"].ToString())); // 歷年累積撥發金額(仟元)
                Tools.SetCellstyle(ws, "I" + rowNumber);
                ws.Cells["J" + rowNumber].PutValue(Source.Tools.thousandchange3(dt2.Rows[i]["PPAY_FEE"].ToString()));     // 當年累積撥發金額(仟元)
                Tools.SetCellstyle(ws, "J" + rowNumber);
                ws.Cells["K" + rowNumber].PutValue(Source.Tools.thousandchange3(dt2.Rows[i]["YR_FUND"].ToString()));      // 當年度核定經費(仟元)(A)
                Tools.SetCellstyle(ws, "K" + rowNumber);
                ws.Cells["O" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK103"].ToString()));         // 預定細設完成日期
                ws.Cells["P" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK103"].ToString()));         // 實際細設完成日期
                ws.Cells["Q" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK105"].ToString()));         // 預定發包日期
                ws.Cells["R" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK105"].ToString()));         // 實際發包日期
                ws.Cells["M" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK102"].ToString()));         // 基設核訂日期

                string S = dt2.Rows[i]["PRE_SCHD"].ToString(); // 預定工程進度(%)
                double S_DOUBLE;

                if (!string.IsNullOrEmpty(S))
                {
                    S_DOUBLE = double.Parse(S);
                }
                else
                {
                    S_DOUBLE = 0;
                }

                ws.Cells["S" + rowNumber].Value = S;             // 預定工程進度(%)
                ws.Cells["N" + rowNumber].Value = TD_START_FUND; // 基設核定經費

                string T = dt2.Rows[i]["REL_SCHD"].ToString();
                double T_DOUBLE;

                if (!string.IsNullOrEmpty(T))
                {
                    T_DOUBLE = double.Parse(T);
                }
                else
                {
                    T_DOUBLE = 0;
                }

                ws.Cells["T" + rowNumber].Value = T; // 實際工程進度(%)

                if (string.IsNullOrEmpty(S) || string.IsNullOrEmpty(T))
                {
                    ws.Cells["U" + rowNumber].PutValue(""); // 進度差異(%)
                }
                else
                {
                    double sValue = double.Parse(S);
                    double tValue = double.Parse(T);
                    ws.Cells["U" + rowNumber].PutValue(tValue - sValue); // 進度差異(%)
                }

                ws.Cells["V" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK107"].ToString()));  // 預定完工日期
                ws.Cells["W" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK107"].ToString()));  // 實際完工日期
                ws.Cells["Z" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK108"].ToString()));  // 預定正驗日期
                ws.Cells["AA" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK108"].ToString())); // 實際正驗日期
                ws.Cells["AB" + rowNumber].PutValue(dt2.Rows[i]["EXECT_DESC"].ToString());                         // 辦理情形或落後原因說明
                ws.Cells["AC" + rowNumber].PutValue(dt2.Rows[i]["DELAY_RESN"].ToString());                         // 落後改善對策

                string an = dt2.Rows[i]["REL_TK105"].ToString();
                string ao = dt2.Rows[i]["BID_FEE"].ToString();
                string ap = dt2.Rows[i]["BG_FEE"].ToString();
                var TD_TOTAL_COST = Source.Tools.IfFormula(an, ap, ao);

                ws.Cells["AD" + rowNumber].PutValue(TD_TOTAL_COST); // 總工程經費(仟元)
                Tools.SetCellstyle(ws, "AD" + rowNumber);
                ws.Cells["L" + rowNumber].PutValue(TD_TOTAL_COST * 0.95 * T_DOUBLE / 100 * CEN_SP_RATIO_TWODECIMALPLACE / 100); // 依進度可撥經費(仟元)
                Tools.SetCellstyle(ws, "L" + rowNumber);
                ws.Cells["AE" + rowNumber].PutValue(dt2.Rows[i]["CTDELAY_RESN"].ToString());  // 備註
                ws.Cells["X" + rowNumber].PutValue(TD_HOUSEHOLD);                             // 計畫接管戶數
                ws.Cells["Y" + rowNumber].PutValue(TD_TOTAL_HOUSEHOLD);                       // 實際累計接管戶數
                ws.Cells["AF" + rowNumber].PutValue(dt2.Rows[i]["DSG_BUSS_NAME"].ToString()); // 設計單位
                ws.Cells["AG" + rowNumber].PutValue(dt2.Rows[i]["BUSS_NAME"].ToString());     // 廠商名稱

                // 20240124 修改項目 => 在最後一行加上合計
                B += Source.Tools.thousandchange3(dt2.Rows[i]["CATY_FUND_"].ToString());
                I += Source.Tools.thousandchange3(dt2.Rows[i]["ALL_PPAY_FEE"].ToString());
                J += Source.Tools.thousandchange3(dt2.Rows[i]["PPAY_FEE"].ToString());
                K += Source.Tools.thousandchange3(dt2.Rows[i]["YR_FUND"].ToString());
                L += (TD_TOTAL_COST * 0.95 * T_DOUBLE / 100 * CEN_SP_RATIO_TWODECIMALPLACE / 100);
                AD += TD_TOTAL_COST;

                if (dt2.Rows.Count > 0 && i == dt2.Rows.Count - 1)
                {
                    rowNumber += 1;
                    ws.Cells["A" + rowNumber].PutValue("合計");
                    ws.Cells["B" + rowNumber].PutValue(B);
                    ws.Cells["I" + rowNumber].PutValue(I);
                    ws.Cells["J" + rowNumber].PutValue(J);
                    ws.Cells["K" + rowNumber].PutValue(K);
                    ws.Cells["L" + rowNumber].PutValue(L);
                    ws.Cells["AD" + rowNumber].PutValue(AD);
                }
            }

            Aspose.Cells.Style style = workbook.DefaultStyle;
            style.Font.Name = "Microsoft JhengHei Light";
            workbook.DefaultStyle = style;
            ws.AutoFitColumns();
            ws.AutoFitRows();

            string filename = "A-2.細項執行情形表" + DateTime.Now.ToString("yyyyMMddHHmmss");
            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");
            workbook.Save(savepath, SaveFormat.Xlsx);

            if (ODS == "1")
            {
                var odsSavePath = Server.MapPath("~/Temp/" + filename + ".ods");
                workbook.Save(odsSavePath, SaveFormat.ODS);

                using (var stream = new FileStream(odsSavePath, FileMode.Open))
                {
                    var bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, (int)stream.Length);
                    return File(bytes, "application/vnd.oasis.opendocument.spreadsheet", filename + ".ods");
                }
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Seek(0, SeekOrigin.Begin);
                    var bytes = stream.ToArray();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename + ".xlsx");
                }
            }
        }

        public ActionResult ReportPrint4(string StartYear, string EndYear, string[] plan, string[] CITY, string Row, string Field, string[] COLUMN1, string PDF)
        {
            Debug.WriteLine("ReportPrint4");
            Debug.WriteLine($"StartYear = {StartYear}");
            Debug.WriteLine($"EndYear = {EndYear}");
            Debug.Write($"plan");
            foreach (var item in plan)
            {
                Debug.Write($"{item} ");
            }
            Debug.Write($"CITY");
            foreach (var item2 in CITY)
            {
                Debug.Write($"{item2} ");
            }
            Debug.WriteLine($"Row = {Row}");
            Debug.WriteLine($"Field = {Field}");
            Debug.Write($"COLUMN1");
            foreach (var item3 in COLUMN1)
            {
                Debug.Write($"{item3} ");
            }
            Debug.WriteLine($"PDF = {PDF}");


            var path = Server.MapPath("~/Templates/表1.xlsx");

            var file = new FileInfo(path);

            List<string> modifiedCities = new List<string>(CITY);

            if (Array.Exists(CITY, element => element == "新北市"))
            {
                modifiedCities.Add("台北縣");
            }

            if (Array.Exists(CITY, element => element == "桃園市"))
            {
                modifiedCities.Add("桃園縣");
            }

            if (Array.Exists(CITY, element => element == "台中市"))
            {
                modifiedCities.Add("台中縣");
            }

            if (Array.Exists(CITY, element => element == "台南市"))
            {
                modifiedCities.Add("台南縣");
            }

            if (Array.Exists(CITY, element => element == "高雄市"))
            {
                modifiedCities.Add("高雄縣");
            }

            CITY = modifiedCities.ToArray();

            string[] COLUMN = new string[COLUMN1.Length + 1];
            COLUMN[0] = "年分";
            Array.Copy(COLUMN1, 0, COLUMN, 1, COLUMN1.Length);
            var dataTable = new DataTable();

            /// 20241119 修改項目 => 計算總和相關參數設定
            /// =================================================
            var dict = new Dictionary<string, decimal>();
            dict.Add("中央款總和", 0);
            dict.Add("累積撥款數總和", 0);
            HashSet<string> procItem = new HashSet<string>();
            HashSet<string> procFD_CODE = new HashSet<string>();
            HashSet<string> procTD_CODE = new HashSet<string>();
            HashSet<string> procTD_CODE2 = new HashSet<string>();
            /// =================================================

            /// 取各年度區間各年度資料
            string[] allYears = GenerateYearsArray(Source.Tools.StringToInt(StartYear), Source.Tools.StringToInt(EndYear));
            foreach (var columnName in COLUMN)
            {
                if (columnName == "當年度核定數")
                {
                    foreach (var year in allYears)
                    {
                        dataTable.Columns.Add(year + "年核定數");
                        dict.Add(year + "年核定數總和", 0); //計算總和相關參數設定 //-k 2411-3
                    }
                }
                else
                {
                    dataTable.Columns.Add(columnName);
                }
            }

            //var CTDMNDS = db2.CTDMNDS.ToList();
            var OGCDEPT = db.OGCDEPT.ToList();
            var CTDITFD = db2.CTDITFD.Where(m => m.FIT_CODE == "11").ToList();
            var TDMFORM = db3.TDMFORM.ToList();

            var TDDBGLS = db3.TDDBGLS.ToList().Where(m => m.SEQ_NO == 0 && m.REC_TP == "B").GroupBy(m => m.TD_CODE).Select(p => new
            {
                TD_CODE = p.Key,
                Total = p.Sum(m => m.BUDGET_FEE)
            }).ToList();
            var TDDSDWB_CONTRATE = db3.TDDSDWB.Where(m => m.IT_NAME == "發包日期").ToList();
            var TDDSDWB_START = db3.TDDSDWB.Where(m => m.IT_NAME == "工程開工").ToList();
            var TDDSDWB_END = db3.TDDSDWB.Where(m => m.IT_NAME == "工程完工").ToList();

            string Year = "";
            string Master_Plan_Name = "";
            string Plan_Name = "";

            /// 抓取每個年度最後月份的經費
            var query = (from t in db2.TDDSTCF.ToList()
                         group t by new { t.FD_CODE, t.SYR } into g
                         select new
                         {
                             FD_CODE = g.Key.FD_CODE,
                             SYR = g.Key.SYR,
                             SMN = g.OrderByDescending(t => t.SMN).First().SMN,
                             COST_FEE = g.OrderByDescending(t => t.SMN).First().COST_FEE,
                             PPAY_FEE = g.OrderByDescending(t => t.SMN).First().PPAY_FEE,
                             TD_CODE1 = g.OrderByDescending(t => t.SMN).First().TD_CODE1
                         }).OrderBy(m => m.FD_CODE).ThenBy(m => m.SYR).ToList();

            var query_insert_column = query.Select((item, index) => new TDDSTCF_INSERT_CLOMNS
            {
                FD_CODE = item.FD_CODE,
                SYR = item.SYR,
                SMN = item.SMN,
                COST_FEE = item.COST_FEE,
                PPAY_FEE = item.PPAY_FEE,
                TD_CODE1 = item.TD_CODE1,
                MATH = (index > 0 && item.FD_CODE == query[index - 1].FD_CODE)
                        ? item.COST_FEE - query[index - 1].COST_FEE : item.COST_FEE
            }).ToList();

            var query2 = from t in db3.TDDMSCD.ToList()
                         group t by new { t.TD_CODE, t.SYR } into g
                         select new
                         {
                             TD_CODE = g.Key.TD_CODE,
                             SYR = g.Key.SYR,
                             SMN = g.Max(t => t.SMN),
                             PRE_SCHD = g.Max(t => t.PRE_SCHD),
                             REL_SCHD = g.Max(t => t.REL_SCHD),
                             EXECT_DESC = g.Max(t => t.EXECT_DESC),
                             DELAY_RESN = g.Max(t => t.DELAY_RESN),
                             CTDELAY_RESN = g.Max(t => t.CTDELAY_RESN),
                         };

            var TDV_TDDCFTS = db2.TDV_TDDCFTS.Where(m => m.YR_FUND != 0 || (m.FCAL_FG ?? 0 == 1)).ToList();

            var TDV_TDDCFTS2 = (TDV_TDDCFTS.Select(m => new
            {
                m.PID_SYR, /// 20241119 修改項目
                m.FD_CODE, /// 20241119 修改項目
                SYR_FD_CODE = m.PID_SYR + m.FD_CODE, /// 20241119 修改項目
                m.PID,
                m.TD_CODE,
                m.TD_NAME,
                m.CITY,
                m.TOWN,
                m.IPJ_CODE,
                m.IPJ_DESC,
                m.ADV_DP,
                m.TID,
                m.CEN_SP_RATIO,
                m.CEN_FUND,
                SYR_TD_CODE = m.PID_SYR + m.TD_CODE
            }
            )).ToList();

            var TDV_TDDCFTS_SUM = db2.TDV_TDDCFTS.GroupBy(m => m.TD_CODE).Select(p => new
            {
                TD_CODE = p.Key,
                Total = p.Sum(m => m.CEN_FUND)
            }).ToList();

            /// 20241119 修改項目
            /// 取得工程編號各個年度累積的撥款資料
            var financialDataQuery = query
                .GroupBy(f => new { f.TD_CODE1, f.SYR })
                .Select(g => new
                {
                    g.Key.TD_CODE1,
                    g.Key.SYR,
                    SUM_FEE = g.Sum(f => (f.COST_FEE ?? 0) + (f.PPAY_FEE ?? 0)),
                    SYR_TD_CODE = g.Key.SYR + g.Key.TD_CODE1
                });

            for (int w = 0; w < plan.Count(); w++)
            {
                string[] sArray;
                sArray = plan[w].Split('+');

                /// 20240131 修改項目 => 全部資料抓最後一年
                Year = EndYear;
                Master_Plan_Name = sArray[0];
                Plan_Name = sArray[1];
                var TDDSTCF = query_insert_column.Where(m => m.SYR == Year).ToList(); /// 這個是某年度的累積撥款數
                while (TDDSTCF.Count == 0 && StringToDouble(Year) >= StringToDouble(StartYear))
                {
                    TDDSTCF = query_insert_column.Where(m => m.SYR == Year).ToList();
                    double Year_double = StringToDouble(Year);
                    if (TDDSTCF.Count == 0)
                    {
                        Year_double--;
                        Year = Year_double.ToString();
                    }
                }
                var PJMFORM = db2.PJMFORM.Where(m => m.SYR.CompareTo(Year) <= 0 && m.SYR.CompareTo(StartYear) >= 0 && m.PJ_NAME == Plan_Name).ToList();

                /// 計畫下的每個工程
                int yearNumber = int.Parse(Year);

                var TDDSTCF_TOTAL = query.Where(m => int.Parse(m.SYR) <= yearNumber).GroupBy(m => m.TD_CODE1).Select(p => new
                {
                    TDCODE1 = p.Key,
                    Total = p.Sum(m => m.COST_FEE) + p.Sum(m => m.PPAY_FEE)
                }).ToList();

                var TDDMSCD = query2.Where(m => m.SYR == Year).ToList();

                var SOURCE_MERGE = (from a in PJMFORM
                                    join b in TDV_TDDCFTS2 on a.PID equals b.PID into Joinb
                                    from b in Joinb.DefaultIfEmpty()
                                    join c in TDV_TDDCFTS_SUM on b?.TD_CODE equals c?.TD_CODE into Joinc
                                    from c in Joinc.DefaultIfEmpty()
                                    join d in TDMFORM on b?.TD_CODE equals d.TD_CODE into Joind
                                    from d in Joind.DefaultIfEmpty()
                                    join e in OGCDEPT on b?.ADV_DP equals e.DEPT_NO into Joine
                                    from e in Joine.DefaultIfEmpty()
                                    join f in TDDBGLS on b?.TD_CODE equals f.TD_CODE into Joinf
                                    from f in Joinf.DefaultIfEmpty()
                                    join g in financialDataQuery on b?.SYR_TD_CODE equals g.SYR_TD_CODE into Joing
                                    from g in Joing.DefaultIfEmpty()
                                    join i in TDDSDWB_CONTRATE on b?.TD_CODE equals i.TD_CODE into contrateJoin
                                    from i in contrateJoin.DefaultIfEmpty()
                                    join m in TDDSDWB_START on b?.TD_CODE equals m.TD_CODE into startJoin
                                    from m in startJoin.DefaultIfEmpty()
                                    join n in TDDSDWB_END on b?.TD_CODE equals n.TD_CODE into endJoin
                                    from n in endJoin.DefaultIfEmpty()
                                    join l in TDDMSCD on b?.TD_CODE equals l.TD_CODE into query2Join
                                    from l in query2Join.DefaultIfEmpty()
                                    where b != null
                                    select new CTDMNDSmergeTDV_TDDCFTS
                                    {
                                        FD_CODE = b?.FD_CODE, /// 20241119 修改項目
                                        PJ_NAME = Plan_Name,
                                        PID = b?.PID != null ? (decimal)b.PID : 0,
                                        TD_NAME = b.TD_NAME,
                                        TD_CITY = b.CITY,
                                        TD_TOWN = b.TOWN,
                                        TD_CODE = b?.TD_CODE,
                                        SYR = b?.SYR_TD_CODE.Substring(0, 3),
                                        IPJ_CODE = b.IPJ_CODE,
                                        IPJ_DESC = b.IPJ_DESC, /// 辦理層級
                                        ADV_DP = b.ADV_DP,
                                        TID = b?.TID != null ? (decimal)b.TID : 0,
                                        FUND = c?.Total != null ? (decimal)c.Total : 0, /// 中央款
                                        PERCENT_FUND = b.CEN_SP_RATIO != null ? (decimal)b.CEN_SP_RATIO : 0, /// 工程費比例
                                        TOTAL_ALL_BUDGET_FEE = b?.CEN_FUND != null ? (decimal)b.CEN_FUND : 0,
                                        TD_LINE = d?.TD_LINE,
                                        TD_NOTE = d?.TD_NOTE,
                                        TD_MEMO = d?.TD_MEMO,
                                        SMP_DEPT_NAME = e?.SMP_DEPT_NAME,
                                        TOTAL_FUND = f?.Total != null ? (decimal)f?.Total : 0, /// 總工程費(元)
                                        COST_FEE = g?.SUM_FEE ?? 0,
                                        PRE_CONTRACT_DATE = i?.PRE_EDT,
                                        REL_CONTRACT_DATE = i?.REL_EDT,
                                        PRE_START_DATE = m?.PRE_EDT,
                                        REL_START_DATE = m?.REL_EDT,
                                        PRE_END_DATE = n?.PRE_EDT,
                                        REL_END_DATE = n?.REL_EDT,
                                        PRE_SCHD = l?.PRE_SCHD != null ? (decimal)l.PRE_SCHD : 0,
                                        REL_SCHD = l?.REL_SCHD != null ? (decimal)l.REL_SCHD : 0,
                                        EXECT_DESC = l?.EXECT_DESC,
                                        DELAY_RESN = l?.DELAY_RESN,
                                        CTDELAY_RESN = l?.CTDELAY_RESN,
                                        PROGRESS_DIFFERENCE = (l?.PRE_SCHD != null ? (decimal)l.PRE_SCHD : 0) - (l?.REL_SCHD != null ? (decimal)l.REL_SCHD : 0),
                                    }).GroupBy(m => m.TD_CODE)
                                    .Select(group => group.OrderByDescending(m => m.SYR).First())
                                    .ToList();

                Trace.WriteLine("SOURCE_MERGE 工程數 = " + SOURCE_MERGE.Select(x => x.TD_CODE).Distinct().Count());


                string MAX_IPJ_CODE = SOURCE_MERGE.Max(m => m.IPJ_CODE);
                SOURCE_MERGE = SOURCE_MERGE.Where(m => CITY.Any(value => m.TD_CITY.Contains(value))).ToList();
                SOURCE_MERGE = SOURCE_MERGE.OrderBy(m => m.TD_CITY).ToList();

                foreach (var item in SOURCE_MERGE)
                {
                    var row = dataTable.NewRow();

                    /// 20241119 修改項目
                    var TD_CODE = item.GetType().GetProperty("TD_CODE")?.GetValue(item) ?? 0;
                    var SYR = item.GetType().GetProperty("SYR")?.GetValue(item) ?? "";

                    /// 同工程編號同年度已有資料則顯示一筆就好
                    string sKey = TD_CODE + "-" + SYR;
                    if (!procItem.Contains(sKey))
                    {
                        procItem.Add(sKey);
                    }
                    /// 已有資料則不再加入 
                    else
                    {
                        continue;
                    }
                    /// 20241119 修改項目

                    var value = item.GetType().GetProperty("SYR")?.GetValue(item);
                    row["年分"] = value;
                    foreach (var columnName in COLUMN)
                    {
                        if (columnName == "年分")
                        {
                            continue;
                        }
                        else if (columnName.Contains("本署補助經費(中央款)(元)"))
                        {
                            /// 20241119 修改項目
                            /// 計算本署補助經費(中央款)(元)總和
                            var Fund = item.GetType().GetProperty("FUND")?.GetValue(item) ?? 0;
                            row["本署補助經費(中央款)(元)"] = Fund;
                            string idKey = TD_CODE.ToString();
                            if (!procTD_CODE.Contains(idKey)) /// 如果沒處理過，則加總並標記為已處理
                            {
                                dict["中央款總和"] += (decimal)Fund;
                                procTD_CODE.Add(idKey);
                            }
                        }
                        else if (columnName.Contains("年度核定數"))
                        {
                            /// 抓取年度區間內各年度的核定數(只會取得該年度之前的)
                            /// 每筆工程編號每年只會有一筆資料
                            foreach (var year in allYears)
                            {
                                /// if (year.CompareTo(SYR.ToString()) > 0) continue;

                                var pid = item.PID;

                                decimal Fund = TDV_TDDCFTS2
                                    .Where(m => m.TD_CODE == TD_CODE.ToString() && m.PID_SYR == year.ToString() && m.PID == pid)
                                    .Select(s => s.CEN_FUND).DefaultIfEmpty(0).Sum();

                                row[year + "年核定數"] = Fund;

                                /// 計算xxx年度核定數總和
                                string idYearKey = TD_CODE + "-" + year;

                                if (!procFD_CODE.Contains(idYearKey)) /// 如果沒處理過，則加總並標記為已處理
                                {
                                    dict[year + "年核定數總和"] += Fund;
                                    procFD_CODE.Add(idYearKey);
                                }
                            }
                        }
                        else if (columnName == "累積撥款數") //-k2411-3
                        {

                            //計算每年累積撥款數
                            //每筆工程編號每年只會有一筆資料
                            if (!string.IsNullOrEmpty(SYR.ToString()))
                            {
                                var preSum = financialDataQuery.Where(s => s.TD_CODE1 == TD_CODE.ToString())
                                .Where(s => s.SYR.CompareTo(SYR) <= 0).Sum(s => s.SUM_FEE);
                                row["累積撥款數"] = preSum;

                                //計算累積撥款數總和
                                if (SYR.ToString() == financialDataQuery.Where(s => s.TD_CODE1 == TD_CODE.ToString()).OrderByDescending(s => s.SYR).FirstOrDefault()?.SYR)
                                {
                                    string idYearKey = TD_CODE.ToString();
                                    if (!procTD_CODE2.Contains(idYearKey)) //如果沒處理過，則加總並標記為已處理
                                    {
                                        dict["累積撥款數總和"] += preSum;
                                        procTD_CODE2.Add(idYearKey);
                                    }
                                }
                            }
                        }
                        else
                        {
                            var COLUMN_IDNAME = db.OGCTITLE.Where(m => m.TITLE == columnName).FirstOrDefault().NAME;
                            value = item.GetType().GetProperty(COLUMN_IDNAME)?.GetValue(item);
                            row[columnName] = value;
                        }
                    }
                    dataTable.Rows.Add(row);
                }
            }

            var rowcount = dataTable.Rows.Count;

            ///////////////Aspose.Cells套件使用////////////////////
            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            var workbook = new Aspose.Cells.Workbook(path);
            var ws1 = workbook.Worksheets["工程細項執行表(縣市在列)"];
            var Column_length = Report_Source.NumberToColumnLetter(COLUMN.Length + allYears.Length);

            var A = ws1.Cells.CreateRange("A1:A1");
            var B = ws1.Cells.CreateRange("A1:" + Column_length + (rowcount + 2));
            var R = ws1.Cells.CreateRange("A1", Column_length + (rowcount + 2 + 1));

            Aspose.Cells.Style style1 = workbook.CreateStyle();


            Aspose.Cells.Style borderstyle = ws1.Workbook.CreateStyle();

            // Set outer borders
            Aspose.Cells.Border outerBorder = borderstyle.Borders[BorderType.TopBorder];
            outerBorder.LineStyle = CellBorderType.Thin;

            outerBorder = borderstyle.Borders[BorderType.BottomBorder];
            outerBorder.LineStyle = CellBorderType.Thin;

            outerBorder = borderstyle.Borders[BorderType.LeftBorder];
            outerBorder.LineStyle = CellBorderType.Thin;

            outerBorder = borderstyle.Borders[BorderType.RightBorder];
            outerBorder.LineStyle = CellBorderType.Thin;


            // Set inner borders
            Aspose.Cells.Border innerBorder = borderstyle.Borders[BorderType.Vertical];
            innerBorder.LineStyle = CellBorderType.Thin;

            innerBorder = borderstyle.Borders[BorderType.Horizontal];
            innerBorder.LineStyle = CellBorderType.Thin;
            borderstyle.IsTextWrapped = true;  ////自動換行
            borderstyle.HorizontalAlignment = TextAlignmentType.Center;
            R.SetStyle(borderstyle);

            B.Copy(A);
            int k = 0;
            for (int i = 2; i < COLUMN.Length + 2; i++)
            {
                var title = COLUMN[i - 2];
                if (title == "當年度核定數")
                {
                    foreach (var year in allYears)
                    {
                        var ABC = Report_Source.NumberToColumnLetter(i + k);
                        var a = ws1.Cells[ABC + "2"];
                        a.PutValue(year + "年核定數");
                        a.GetStyle().Borders.SetStyle(CellBorderType.Thin);
                        if (k < allYears.Length - 1)
                        {
                            k++;
                        }

                    }
                }
                else
                {
                    var ABC = Report_Source.NumberToColumnLetter(i + k);
                    var a = ws1.Cells[ABC + "2"];
                    a.PutValue(title);
                    a.GetStyle().Borders.SetStyle(CellBorderType.Thin);
                }
            }

            int j = 3;

            for (var n = 0; n < rowcount; n++)
            {
                var a0 = ws1.Cells["A" + j];
                a0.PutValue(dataTable.Rows[j - 3]["工程地點(址)"]);
                k = 0;
                for (int i = 2; i < COLUMN.Length + 2; i++)
                {
                    var ABC = Report_Source.NumberToColumnLetter(i + k);
                    var a1 = ws1.Cells[ABC + j];
                    Aspose.Cells.Style a1_Style = a1.GetStyle();
                    a1_Style.Number = 5;
                    var title = COLUMN[i - 2];

                    if (title == "序號")
                    {
                        a1.PutValue(j - 2);
                    }
                    else if (title == "補助對象")
                    {
                        a1.PutValue(dataTable.Rows[j - 3]["工程地點(址)"] + "政府");
                    }
                    else if (title == "落後原因")
                    {
                        a1.PutValue(dataTable.Rows[j - 3]["落後原因"]);
                    }
                    else if (title == "總工程費(元)")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["總工程費(元)"]);
                        a1.PutValue(value/*.ToString("#,##0")*/);
                        a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    else if (title == "本署補助經費(中央款)(元)")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["本署補助經費(中央款)(元)"]);
                        a1.PutValue(value/*.ToString("#,##0")*/);
                        a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    else if (title == "工程費比例")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["工程費比例"]);
                        a1.PutValue(value/*.ToString("#,##0.00")*/);
                        a1_Style.Custom = "0.00_);[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    //
                    else if (title == "當年度核定數")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                        foreach (var year in allYears)
                        {
                            decimal value = 0;
                            ABC = Report_Source.NumberToColumnLetter(i + k);
                            a1 = ws1.Cells[ABC + j];
                            a1_Style = a1.GetStyle();
                            a1_Style.Number = 5;
                            try
                            {
                                var test = dataTable.Rows[j - 3][year + "年核定數"];
                                value = Source.Tools.ToDecimal(dataTable.Rows[j - 3][year + "年核定數"]);
                            }
                            catch
                            {
                                var test = dataTable.Rows[j - 3][year + "年核定數"];
                                value = 0;
                            }
                            a1.PutValue(value/*.ToString("#,##0")*/);
                            a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                            a1.SetStyle(a1_Style);
                            if (k < allYears.Length - 1)
                            {
                                k++;
                            }
                        }

                    }
                    else if (title == "當年度撥款數")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["當年度撥款數"]);
                        a1.PutValue(value/*.ToString("#,##0")*/);
                        a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    else if (title == "累積撥款數")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["累積撥款數"]);
                        a1.PutValue(value/*.ToString("#,##0")*/);
                        a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    else
                    {
                        a1.PutValue(dataTable.Rows[j - 3][title].ToString());
                    }
                }
                j++;
            }

            ////////從這裡開始增加一個加總欄位
            int bottom_count = rowcount + 1 + 2; ///////最底部的層數
            k = 0;
            for (int t = 0; t < COLUMN.Length; t++)
            {
                var ABC = Report_Source.NumberToColumnLetter(t + 2 + k);
                ws1.Cells["A" + bottom_count.ToString()].PutValue("合計");
                var b = ws1.Cells[ABC + bottom_count];
                Aspose.Cells.Style b_Style = b.GetStyle();
                decimal sum = 0;
                if (COLUMN[t] == "總工程費(元)")
                {

                    foreach (DataRow row in dataTable.Rows)
                    {
                        if (row["總工程費(元)"] != DBNull.Value) // 確認值不為 DBNull
                        {
                            sum += Source.Tools.ToDecimal(row["總工程費(元)"]);
                        }
                    }

                    b.Value = sum;
                    b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                    b.SetStyle(b_Style);
                }
                else if (COLUMN[t] == "本署補助經費(中央款)(元)")
                {
                    //顯示已處理過的總和
                    b.Value = dict["中央款總和"];

                    b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                    b.SetStyle(b_Style);
                }
                else if (COLUMN[t] == "當年度核定數")
                {
                    foreach (var year in allYears)
                    {
                        ABC = Report_Source.NumberToColumnLetter(t + 2 + k);
                        b = ws1.Cells[ABC + bottom_count];
                        b_Style = b.GetStyle();
                        sum = 0;

                        //顯示已處理過的總和
                        b.Value = dict[year + "年核定數總和"];

                        b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        b.SetStyle(b_Style);
                        if (k < allYears.Length - 1)
                        {
                            k++;
                        }
                    }
                }
                else if (COLUMN[t] == "當年度撥款數")
                {
                    foreach (DataRow row in dataTable.Rows)
                    {
                        if (row["當年度撥款數"] != DBNull.Value) // 確認值不為 DBNull
                        {
                            sum += Source.Tools.ToDecimal(row["當年度撥款數"]);
                        }
                    }

                    b.Value = sum;
                    b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                    b.SetStyle(b_Style);
                }
                else if (COLUMN[t] == "累積撥款數")
                {
                    b.Value = dict["累積撥款數總和"];

                    b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                    b.SetStyle(b_Style);
                }
                else
                {
                }
            }

            var mergedRange = ws1.Cells.CreateRange("A1", Column_length + "1");
            mergedRange.Merge();
            mergedRange[0, 0].PutValue("工程細項執行表");

            Aspose.Cells.Style style = workbook.DefaultStyle;
            style.Font.Name = "Microsoft JhengHei Light";
            workbook.DefaultStyle = style;
            ws1.AutoFitColumns();
            ws1.AutoFitRows();

            for (int i = 0; i < COLUMN.Length; i++)
            {
                var Excel_column = Report_Source.NumberToColumnLetter(i + 1);
                var data_column = dataTable.Columns[i].ColumnName;
                if (data_column == "年分")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 4);
                }
                else if (data_column == "序號")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 4);
                }
                else if (data_column == "補助計畫名稱")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 27.42);
                }
                else if (data_column == "法規依據")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 7.5);
                }
                else if (data_column == "計畫期程")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 7.5);
                }
                else if (data_column == "補助對象")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 9.33);
                }
                else if (data_column == "工程編號")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 19.08);
                }
                else if (data_column == "工程名稱")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "工程地點(址)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 10.42);
                }
                else if (data_column == "行政區")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 15.33);
                }
                else if (data_column == "工程主辦機關")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 11.08);
                }
                else if (data_column == "填寫本署(自辦、代辦或補助)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 22.83);
                }
                else if (data_column == "總工程費(元)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 12.42);
                }
                else if (data_column == "分標工程費(元)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 12.17);
                }
                else if (data_column == "本署補助經費(中央款)(元)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 20.42);
                }
                else if (data_column == "工程費比例")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 5.83);
                }
                else if (data_column == "當年度核定數")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 11.5);
                }
                else if (data_column == "當年度撥款數")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 11.5);
                }
                else if (data_column == "累積撥款數")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 24.83);
                }
                else if (data_column == "工程內容") //-k 2411-3
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "辦理情形")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "預定發包日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "實際發包日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "預定開工日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "實際開工日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "預定開工典禮日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 14.67);
                }
                else if (data_column == "實際開工典禮日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 14.67);
                }
                else if (data_column == "預定竣工日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "實際竣工日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 7.5);
                }
                else if (data_column == "預定竣工典禮日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 14.67);
                }
                else if (data_column == "實際竣工典禮日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 14.67);
                }
                else if (data_column == "預定進度")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 5.33);
                }
                else if (data_column == "實際進度")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.17);
                }
                else if (data_column == "進度差異(預定-實際)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 16.33);
                }
                else if (data_column == "落後原因或辦理情形說明") //-k 2411-3
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "落後解決對策") //-k 2411-3
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "備註")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }

            }

            ws1.Cells.DeleteColumn(1);

            string filename = "A-4.跨年度細項執行情形表" + DateTime.Now.ToString("yyyyMMddHHmmss");

            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");
            workbook.Save(savepath, SaveFormat.Xlsx); // 保存为 Excel 文件

            if (PDF == "1")
            {
                var odsSavePath = Server.MapPath("~/Temp/" + filename + ".ods");
                workbook.Save(odsSavePath, SaveFormat.ODS);

                // 返回 .ods 檔案至客戶端
                using (var stream = new FileStream(odsSavePath, FileMode.Open))
                {
                    var bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, (int)stream.Length);
                    return File(bytes, "application/vnd.oasis.opendocument.spreadsheet", filename + ".ods");
                }
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx); // 将 Excel 文件保存到内存流中
                    stream.Seek(0, SeekOrigin.Begin);
                    var bytes = stream.ToArray();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename + ".xlsx");
                }
            }
            ////////////////////////////////////////////////////////////////////

        }

        private string Initdate(string sdt)
        {
            string rtn = "";
            if (sdt.Length == 7)
            {
                rtn = sdt.Substring(0, 3) + "." + sdt.Substring(3, 2) + "." + sdt.Substring(5, 2);
            }
            else
            {
                rtn = sdt;
            }
            return rtn;
        }

        private ActionResult OdsPrint(string filepath, string filename)
        {
            Workbook workbook = new Workbook(filepath);
            for (int j = 0; j < workbook.Worksheets.Count; j++)
            {
                workbook.Worksheets[j].AutoFitRows();
                workbook.Worksheets[j].AutoFitColumns();
            }

            string odsOutputFile = Server.MapPath("~/Temp/" + filename + ".ods");
            workbook.Save(odsOutputFile, SaveFormat.ODS);

            byte[] bytes = System.IO.File.ReadAllBytes(odsOutputFile);
            return File(bytes, "application/vnd.oasis.opendocument.spreadsheet", filename + ".ods");
        }

        public DataTable getReport02_CITY()
        {

            DataTable dt = new DataTable();

            string sql = "SELECT MAX(CPA_CITY_ID) AS CPA_CITY_ID, CITY FROM CMBASE.dbo.SMCPOSTN GROUP BY CITY ORDER BY CPA_CITY_ID";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    conn.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        dt.Load(reader);
                    }
                }
            }
            return dt;
        }

        private DataTable GetDtH01(DataTable DtS1, DataTable DtS2, string sType, DataTable DtCITY)
        {
            DataTable DtBack = new DataTable();

            DtBack.Columns.Add("CITY");
            DtBack.Columns.Add("CntP01");
            DtBack.Columns.Add("FundP01");
            DtBack.Columns.Add("CntC01");
            DtBack.Columns.Add("FundC01");
            DtBack.Columns.Add("CntM01");
            DtBack.Columns.Add("FundM01");
            DtBack.Columns.Add("CntL01");
            DtBack.Columns.Add("FundL01");

            DtBack.Columns.Add("CntP02");
            DtBack.Columns.Add("FundP02");
            DtBack.Columns.Add("CntC02");
            DtBack.Columns.Add("FundC02");
            DtBack.Columns.Add("CntM02");
            DtBack.Columns.Add("FundM02");
            DtBack.Columns.Add("CntL02");
            DtBack.Columns.Add("FundL02");

            DtBack.Columns.Add("CntP11");
            DtBack.Columns.Add("FundP11");
            DtBack.Columns.Add("CntC11");
            DtBack.Columns.Add("FundC11");
            DtBack.Columns.Add("CntM11");
            DtBack.Columns.Add("FundM11");
            DtBack.Columns.Add("CntL11");
            DtBack.Columns.Add("FundL11");

            DtBack.Columns.Add("CntP22");
            DtBack.Columns.Add("FundP22");
            DtBack.Columns.Add("CntC22");
            DtBack.Columns.Add("FundC22");
            DtBack.Columns.Add("CntM22");
            DtBack.Columns.Add("FundM22");
            DtBack.Columns.Add("CntL22");
            DtBack.Columns.Add("FundL22");

            DtBack.Columns.Add("ACC02_FUND");    /// 新加入的
            DtBack.Columns.Add("ACC03_FUND");
            DtBack.Columns.Add("ACC04_FUND");
            DtBack.Columns.Add("ACC02_FUNDSUM"); /// 新加入的
            DtBack.Columns.Add("ACC03_FUNDSUM");
            DtBack.Columns.Add("ACC04_FUNDSUM");

            DtBack.Columns.Add("M_01");
            DtBack.Columns.Add("M_02");
            DtBack.Columns.Add("M_03");
            DtBack.Columns.Add("M_04");
            DtBack.Columns.Add("M_05");
            DtBack.Columns.Add("M_06");
            DtBack.Columns.Add("M_07");
            DtBack.Columns.Add("M_08");
            DtBack.Columns.Add("M_09");
            DtBack.Columns.Add("M_10");
            DtBack.Columns.Add("M_11");
            DtBack.Columns.Add("M_12");

            /// 加入前一期各費用顯示欄位
            DtBack.Columns.Add("ACC02_FUND_PREV");
            DtBack.Columns.Add("ACC03_FUND_PREV");
            DtBack.Columns.Add("ACC04_FUND_PREV");

            string sCITY = "";
            DataRow NewRw = null;
            switch (sType)
            {
                case "1":
                case "4":
                    for (int x = 0; x < DtCITY.Rows.Count; x++)
                    {
                        NewRw = null;

                        DataRow[] DtS1RowS = DtS1.Select("CITY='" + DtCITY.Rows[x]["CITY"].ToString() + "'");
                        for (int i = 0; i < DtS1RowS.Length; i++)
                        {
                            DataRow Drow1 = DtS1RowS[i];
                            if (sCITY != Drow1["CITY"].ToString())
                            {
                                sCITY = Drow1["CITY"].ToString();

                                NewRw = DtBack.NewRow();
                                NewRw["CITY"] = sCITY;

                                NewRw["M_01"] = "0";
                                NewRw["M_02"] = "0";
                                NewRw["M_03"] = "0";
                                NewRw["M_04"] = "0";
                                NewRw["M_05"] = "0";
                                NewRw["M_06"] = "0";
                                NewRw["M_07"] = "0";
                                NewRw["M_08"] = "0";
                                NewRw["M_09"] = "0";
                                NewRw["M_10"] = "0";
                                NewRw["M_11"] = "0";
                                NewRw["M_12"] = "0";
                            }

                            string sTD_CATY = "";
                            switch (sType)
                            {
                                case "1":
                                    sTD_CATY = Drow1["GOV_TD_CATY"].ToString();
                                    break;
                                case "4":
                                    sTD_CATY = Drow1["BOT_TD_CATY"].ToString();
                                    break;
                            }

                            string sIPJ_CODE = Drow1["IPJ_CODE"].ToString();
                            if (sTD_CATY == "J") { sTD_CATY = "P"; }

                            SetBackValue(Drow1, NewRw, sTD_CATY, sIPJ_CODE);

                            NewRw["ACC02_FUND"] = GetNumSum(NewRw, Drow1, "ACC02_FUND"); /////新加的
                            NewRw["ACC03_FUND"] = GetNumSum(NewRw, Drow1, "ACC03_FUND");
                            NewRw["ACC04_FUND"] = GetNumSum(NewRw, Drow1, "ACC04_FUND");

                            NewRw["ACC02_FUNDSUM"] = GetNumSum(NewRw, Drow1, "ACC02_FUNDSUM", "ACC02_FUND"); /////新加的
                            NewRw["ACC03_FUNDSUM"] = GetNumSum(NewRw, Drow1, "ACC03_FUNDSUM", "ACC03_FUND");
                            NewRw["ACC04_FUNDSUM"] = GetNumSum(NewRw, Drow1, "ACC04_FUNDSUM", "ACC04_FUND");

                            NewRw["M_01"] = GetNumSum(NewRw, Drow1, "M_01");
                            NewRw["M_02"] = GetNumSum(NewRw, Drow1, "M_02");
                            NewRw["M_03"] = GetNumSum(NewRw, Drow1, "M_03");
                            NewRw["M_04"] = GetNumSum(NewRw, Drow1, "M_04");
                            NewRw["M_05"] = GetNumSum(NewRw, Drow1, "M_05");
                            NewRw["M_06"] = GetNumSum(NewRw, Drow1, "M_06");
                            NewRw["M_07"] = GetNumSum(NewRw, Drow1, "M_07");
                            NewRw["M_08"] = GetNumSum(NewRw, Drow1, "M_08");
                            NewRw["M_09"] = GetNumSum(NewRw, Drow1, "M_09");
                            NewRw["M_10"] = GetNumSum(NewRw, Drow1, "M_10");
                            NewRw["M_11"] = GetNumSum(NewRw, Drow1, "M_11");
                            NewRw["M_12"] = GetNumSum(NewRw, Drow1, "M_12");

                            if (DtS2 != null)
                            {
                                string sWhere = "";
                                switch (sType)
                                {
                                    case "1":
                                        sWhere = string.Format("CITY='{0}' And GOV_TD_CATY='{1}' And IPJ_CODE='{2}'", Drow1["CITY"].ToString(), Drow1["GOV_TD_CATY"].ToString(), sIPJ_CODE);
                                        break;
                                    case "4":
                                        sWhere = string.Format("CITY='{0}' And BOT_TD_CATY='{1}' And IPJ_CODE='{2}'", Drow1["CITY"].ToString(), Drow1["BOT_TD_CATY"].ToString(), sIPJ_CODE);
                                        break;
                                }

                                DataRow[] xDrow = DtS2.Select(sWhere);
                                if (xDrow.Length > 0)
                                {
                                    DataRow Drow2 = xDrow[0];
                                    if (sIPJ_CODE == "01") { sIPJ_CODE = "11"; }
                                    if (sIPJ_CODE == "02") { sIPJ_CODE = "22"; }

                                    SetBackValue(Drow2, NewRw, sTD_CATY, sIPJ_CODE);

                                    NewRw["ACC02_FUND_PREV"] = GetNumSum(NewRw, Drow2, "ACC02_FUND_PREV", "ACC02_FUND");
                                    NewRw["ACC03_FUND_PREV"] = GetNumSum(NewRw, Drow2, "ACC03_FUND_PREV", "ACC03_FUND");
                                    NewRw["ACC04_FUND_PREV"] = GetNumSum(NewRw, Drow2, "ACC04_FUND_PREV", "ACC04_FUND");
                                    NewRw["ACC02_FUNDSUM"] = ReportController.GetNumSub(Drow2, NewRw, "ACC02_FUND", "ACC02_FUNDSUM"); /// 新加的
                                    NewRw["ACC03_FUNDSUM"] = ReportController.GetNumSub(Drow2, NewRw, "ACC03_FUND", "ACC03_FUNDSUM");
                                    NewRw["ACC04_FUNDSUM"] = ReportController.GetNumSub(Drow2, NewRw, "ACC04_FUND", "ACC04_FUNDSUM");

                                    DtS2.Rows.Remove(Drow2);
                                }
                            }
                        }

                        if (DtS2 != null)
                        {
                            DtS2.AcceptChanges();

                            DataRow[] DtS2RowS = DtS2.Select("CITY='" + DtCITY.Rows[x]["CITY"].ToString() + "'");

                            for (int i = 0; i < DtS2RowS.Length; i++)
                            {
                                DataRow Drow2 = DtS2RowS[i];
                                if (sCITY != Drow2["CITY"].ToString())
                                {
                                    sCITY = Drow2["CITY"].ToString();
                                    NewRw = DtBack.NewRow();
                                    NewRw["CITY"] = sCITY;
                                    NewRw["M_01"] = "0";
                                    NewRw["M_02"] = "0";
                                    NewRw["M_03"] = "0";
                                    NewRw["M_04"] = "0";
                                    NewRw["M_05"] = "0";
                                    NewRw["M_06"] = "0";
                                    NewRw["M_07"] = "0";
                                    NewRw["M_08"] = "0";
                                    NewRw["M_09"] = "0";
                                    NewRw["M_10"] = "0";
                                    NewRw["M_11"] = "0";
                                    NewRw["M_12"] = "0";
                                }

                                string sTD_CATY = Drow2["GOV_TD_CATY"].ToString();
                                string sIPJ_CODE = Drow2["IPJ_CODE"].ToString();
                                if (sTD_CATY == "J") { sTD_CATY = "P"; }
                                if (sIPJ_CODE == "01") { sIPJ_CODE = "11"; }
                                if (sIPJ_CODE == "02") { sIPJ_CODE = "22"; }

                                NewRw["M_01"] = GetNumSum(NewRw, NewRw, "M_01");
                                NewRw["M_02"] = GetNumSum(NewRw, NewRw, "M_02");
                                NewRw["M_03"] = GetNumSum(NewRw, NewRw, "M_03");
                                NewRw["M_04"] = GetNumSum(NewRw, NewRw, "M_04");
                                NewRw["M_05"] = GetNumSum(NewRw, NewRw, "M_05");
                                NewRw["M_06"] = GetNumSum(NewRw, NewRw, "M_06");
                                NewRw["M_07"] = GetNumSum(NewRw, NewRw, "M_07");
                                NewRw["M_08"] = GetNumSum(NewRw, NewRw, "M_08");
                                NewRw["M_09"] = GetNumSum(NewRw, NewRw, "M_09");
                                NewRw["M_10"] = GetNumSum(NewRw, NewRw, "M_10");
                                NewRw["M_11"] = GetNumSum(NewRw, NewRw, "M_11");
                                NewRw["M_12"] = GetNumSum(NewRw, NewRw, "M_12");

                                SetBackValue(Drow2, NewRw, sTD_CATY, sIPJ_CODE);
                            }
                        }

                        if (NewRw != null) { DtBack.Rows.Add(NewRw); }
                    }
                    break;
            }
            return DtBack;
        }

        private void SetBackValue(DataRow Drow, DataRow NewRw, string sTD_CATY, string sIPJ_CODE)
        {
            // 修正(前瞻基礎建設計畫)多個計畫的資料可合併
            NewRw["Cnt" + sTD_CATY + sIPJ_CODE] = GetNumSum(NewRw, Drow, "Cnt" + sTD_CATY + sIPJ_CODE, "Cnt");
            NewRw["Fund" + sTD_CATY + sIPJ_CODE] = GetNumSum(NewRw, Drow, "Fund" + sTD_CATY + sIPJ_CODE, "Fund");
        }

        private string GetNumSum(DataRow NewRw, DataRow Drow, string FidName)
        {
            return GetNumSum(NewRw, Drow, FidName, FidName);
        }

        private string GetNumSum(DataRow NewRw, DataRow Drow, string FidName1, string FidName2)
        {
            if (NewRw[FidName1] != DBNull.Value)
            {
                if (Drow[FidName2] != DBNull.Value)
                {
                    return Convert.ToString(Convert.ToDouble(NewRw[FidName1].ToString()) + Convert.ToDouble(Drow[FidName2].ToString()));
                }
                else
                {
                    return NewRw[FidName1].ToString();
                }
            }
            else
            {
                if (Drow[FidName2] != DBNull.Value)
                {
                    return Drow[FidName2].ToString();
                }
                else
                {
                    return "0";
                }
            }
        }

        public double StringToDouble(string strValue)
        {
            if (!string.IsNullOrEmpty(strValue))
            {
                string numericString = Regex.Replace(strValue, "[^0-9.]", "");
                double doubleValue;
                if (Double.TryParse(numericString, out doubleValue))
                {
                    return doubleValue;
                }
                else
                {
                    return 0.0;
                }
            }
            else
            {
                return 0.0;
            }
        }

        private DataTable GetDtH00(DataTable DtS1, DataTable DtS2, DataTable DtS3)
        {
            DataTable DtBack = new DataTable();

            DtBack.Columns.Add("CITY");
            DtBack.Columns.Add("FEE_G_1_1");
            DtBack.Columns.Add("FEE_G_1_2");
            DtBack.Columns.Add("FEE_G_2_1");
            DtBack.Columns.Add("FEE_G_2_2");
            DtBack.Columns.Add("FEE_B_1_1");
            DtBack.Columns.Add("FEE_B_1_2");
            DtBack.Columns.Add("FEE_B_2_1");
            DtBack.Columns.Add("FEE_B_2_2");

            //加入業務費用顯示欄位
            DtBack.Columns.Add("FEE_G_1_3");
            DtBack.Columns.Add("FEE_G_2_3");
            DtBack.Columns.Add("FEE_B_1_3");
            DtBack.Columns.Add("FEE_B_2_3");

            DataRow NewRw = null;

            for (int i = 0; i < DtS1.Rows.Count; i++)
            {
                NewRw = DtBack.NewRow();
                NewRw["CITY"] = DtS1.Rows[i]["CITY"].ToString();

                //修改顯示取得資料
                NewRw["FEE_G_1_1"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["ACC03_FUND_PREV"].ToString()));
                NewRw["FEE_G_1_2"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["ACC04_FUND_PREV"].ToString()));
                NewRw["FEE_G_1_3"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["ACC02_FUND_PREV"].ToString()));
                NewRw["FEE_G_2_1"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["ACC03_FUND"].ToString()));
                NewRw["FEE_G_2_2"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["ACC04_FUND"].ToString()));
                NewRw["FEE_G_2_3"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["ACC02_FUND"].ToString()));

                NewRw["FEE_B_1_1"] = 0;
                NewRw["FEE_B_1_2"] = 0;
                NewRw["FEE_B_2_1"] = 0;
                NewRw["FEE_B_2_2"] = 0;
                DtBack.Rows.Add(NewRw);
            }

            int index = -1;

            for (int i = 0; i < DtS2.Rows.Count; i++)
            {
                index = -1;

                for (int j = 0; j < DtBack.Rows.Count; j++)
                {
                    if (DtS2.Rows[i]["CITY"].ToString() == DtBack.Rows[j]["CITY"].ToString())
                    {
                        index = j;
                    }
                }

                if (index == -1)
                {
                    NewRw = DtBack.NewRow();

                    NewRw["CITY"] = DtS2.Rows[i]["CITY"].ToString();
                    NewRw["FEE_G_1_1"] = 0;
                    NewRw["FEE_G_1_2"] = 0;
                    NewRw["FEE_G_2_1"] = 0;
                    NewRw["FEE_G_2_2"] = 0;

                    // 修改顯示取得資料
                    NewRw["FEE_B_1_1"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC03_FUND_PREV"].ToString()));
                    NewRw["FEE_B_1_2"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC04_FUND_PREV"].ToString()));
                    NewRw["FEE_B_1_3"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC02_FUND_PREV"].ToString()));
                    NewRw["FEE_B_2_1"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC03_FUND"].ToString()));
                    NewRw["FEE_B_2_2"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC04_FUND"].ToString()));
                    NewRw["FEE_B_2_3"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC02_FUND"].ToString()));


                    DtBack.Rows.Add(NewRw);
                }
                else
                {
                    //修改顯示取得資料
                    DtBack.Rows[index]["FEE_B_1_1"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC03_FUND_PREV"].ToString()));
                    DtBack.Rows[index]["FEE_B_1_2"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC04_FUND_PREV"].ToString()));
                    DtBack.Rows[index]["FEE_B_1_3"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC02_FUND_PREV"].ToString()));
                    DtBack.Rows[index]["FEE_B_2_1"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC03_FUND"].ToString()));
                    DtBack.Rows[index]["FEE_B_2_2"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC04_FUND"].ToString()));
                    DtBack.Rows[index]["FEE_B_2_3"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["ACC02_FUND"].ToString()));
                }
            }

            if (DtS3.Rows.Count > 0)
            {
                Double fee11 = 0;
                Double fee12 = 0;
                Double fee21 = 0;
                Double fee22 = 0;
                Double feeGDiff = 0;
                Double feeBDiff = 0;

                for (int i = 0; i < DtBack.Rows.Count; i++)
                {
                    fee11 = fee11 + Convert.ToDouble(DtBack.Rows[i]["FEE_G_1_1"].ToString()) + Convert.ToDouble(DtBack.Rows[i]["FEE_B_1_1"].ToString());
                    fee12 = fee12 + Convert.ToDouble(DtBack.Rows[i]["FEE_G_1_2"].ToString()) + Convert.ToDouble(DtBack.Rows[i]["FEE_B_1_2"].ToString());
                    fee21 = fee21 + Convert.ToDouble(DtBack.Rows[i]["FEE_G_2_1"].ToString()) + Convert.ToDouble(DtBack.Rows[i]["FEE_B_2_1"].ToString());
                    fee22 = fee22 + Convert.ToDouble(DtBack.Rows[i]["FEE_G_2_2"].ToString()) + Convert.ToDouble(DtBack.Rows[i]["FEE_B_2_2"].ToString());
                    feeGDiff = feeGDiff + Convert.ToDouble(DtBack.Rows[i]["FEE_G_2_1"].ToString()) + Convert.ToDouble(DtBack.Rows[i]["FEE_G_2_2"].ToString()) - Convert.ToDouble(DtBack.Rows[i]["FEE_G_1_1"].ToString()) - Convert.ToDouble(DtBack.Rows[i]["FEE_G_1_2"].ToString());
                    feeBDiff = feeBDiff + Convert.ToDouble(DtBack.Rows[i]["FEE_B_2_1"].ToString()) + Convert.ToDouble(DtBack.Rows[i]["FEE_B_2_2"].ToString()) - Convert.ToDouble(DtBack.Rows[i]["FEE_B_1_1"].ToString()) - Convert.ToDouble(DtBack.Rows[i]["FEE_B_1_2"].ToString());
                }

                DtS3.Columns.Add("FEE_1_1");
                DtS3.Columns.Add("FEE_1_2");
                DtS3.Columns.Add("FEE_2_1");
                DtS3.Columns.Add("FEE_2_2");
                DtS3.Columns.Add("FEE_G_DIFF");
                DtS3.Columns.Add("FEE_B_DIFF");
                DtS3.Rows[0]["FEE_1_1"] = fee11;
                DtS3.Rows[0]["FEE_1_2"] = fee12;
                DtS3.Rows[0]["FEE_2_1"] = fee21;
                DtS3.Rows[0]["FEE_2_2"] = fee22;
                DtS3.Rows[0]["FEE_G_DIFF"] = feeGDiff;
                DtS3.Rows[0]["FEE_B_DIFF"] = feeBDiff;
            }

            return DtBack;
        }

        private void GetDtH02(DataTable DtS1, DataTable DtS2, string sType, DataTable DtCITY, string SEQ_NO, List<B1DtBack> Lb1Dtback, List<B1DtFooter> Lb1DtFooter)
        {
            //SEQ_NO = Convert.ToString(Convert.ToInt32(SEQ_NO) + 1);

            B1DtBack b1DtBack = new B1DtBack();
            B1DtBack b1DtBack2 = new B1DtBack();
            B1DtFooter b1DtFooter = new B1DtFooter();

            double SUM_CEN_FUND2 = 0;
            double SUM_LOC_FUND2 = 0;
            switch (sType)
            {
                case "2":
                case "3":
                case "5":
                case "6":
                    int CityCls = 0;
                    for (int x = 0; x < DtCITY.Rows.Count; x++)
                    {
                        int iShowTitle = 0;
                        int iShowSerial = 0;

                        int Order = 0;

                        b1DtBack2 = new B1DtBack();

                        b1DtFooter = new B1DtFooter();
                        ReportController.GetDtTitleName(SEQ_NO, b1DtFooter);
                        b1DtFooter.CITY = DtCITY.Rows[x]["CITY"].ToString();
                        for (int i = 1; i < 6; i++)
                        {
                            bool bNewNumYesNo = false;
                            b1DtBack = new B1DtBack();
                            b1DtBack.TD_NAME = "小計";
                            b1DtBack.CityCls = CityCls;
                            string sWhere = string.Format("CITY='{0}' And SUM_CATY='{1}'", DtCITY.Rows[x]["CITY"].ToString(), "C" + i.ToString());
                            DataRow[] DataRowS = DtS1.Select(sWhere);
                            if (DataRowS.Length > 0)
                            {
                                for (int z = 0; z < DataRowS.Length; z++)
                                {
                                    DataRow Drow1 = DataRowS[z];
                                    bool print = true;
                                    if (DtS2 != null)
                                    {
                                        double cenFund1 = 0;
                                        if (Drow1["CEN_FUND"].ToString() != null && Drow1["CEN_FUND"].ToString() != "")
                                        {
                                            cenFund1 = Convert.ToDouble(Drow1["CEN_FUND"].ToString());
                                        }
                                        double locFund1 = 0;
                                        if (Drow1["LOC_FUND"].ToString() != null && Drow1["LOC_FUND"].ToString() != "")
                                        {
                                            locFund1 = Convert.ToDouble(Drow1["LOC_FUND"].ToString());
                                        }
                                        double cenFund2 = 0;
                                        double locFund2 = 0;
                                        sWhere = string.Format("TD_CODE='{0}'", Drow1["TD_CODE"].ToString());
                                        DataRow[] xDrow = DtS2.Select(sWhere);
                                        if (xDrow.Length > 0)
                                        {
                                            DataRow Drow2 = xDrow[0];
                                            if (Drow2["CEN_FUND"].ToString() != null && Drow2["CEN_FUND"].ToString() != "")
                                            {
                                                cenFund2 = Convert.ToDouble(Drow2["CEN_FUND"].ToString());
                                            }
                                            if (Drow2["LOC_FUND"].ToString() != null && Drow2["LOC_FUND"].ToString() != "")
                                            {
                                                locFund2 = Convert.ToDouble(Drow2["LOC_FUND"].ToString());
                                            }
                                        }
                                        if (cenFund1 + locFund1 == 0 && cenFund2 + locFund2 == 0)
                                        {
                                            print = false;
                                            if (xDrow.Length > 0)
                                            {
                                                DtS2.Rows.Remove(xDrow[0]);
                                            }
                                        }
                                    }

                                    if (print == true)
                                    {
                                        iShowTitle = iShowTitle + 1;
                                        iShowSerial = iShowSerial + 1;
                                        b1DtBack2 = new B1DtBack();
                                        b1DtBack2.CityCls = CityCls;
                                        b1DtBack2.Order = iShowTitle;
                                        b1DtBack2.SERIAL = iShowSerial.ToString();
                                        b1DtBack2.SIP_TD_CODE = Drow1["SIP_TD_CODE"].ToString();
                                        b1DtBack2.TD_NAME = Drow1["TD_NAME"].ToString();
                                        b1DtBack2.FD_CODE = Drow1["FD_CODE"].ToString();
                                        b1DtBack2.CATY_DESC = Drow1["CATY_DESC"].ToString();
                                        b1DtBack2.ACC_DESC = Drow1["IPJ_DESC"].ToString();
                                        b1DtBack2.CEN_SP_RATIO = Drow1["CEN_SP_RATIO"].ToString();
                                        SetFUND(b1DtBack2, b1DtBack, b1DtFooter, Drow1, "1");
                                        b1DtBack2.COST_FEE = Drow1["COST_FEE"].ToString();
                                        b1DtBack2.FUND_DESC = Drow1["EXECT_DESC"].ToString();
                                        b1DtFooter.COST_FEESUM = GetNumSumStrb1footer(b1DtFooter, Drow1, "COST_FEESUM", "COST_FEE");
                                        b1DtBack.COST_FEE = GetNumSumStrNew(b1DtBack, Drow1, "COST_FEE", "COST_FEE");
                                        SetMonth(b1DtBack2, Drow1);
                                        SetMonthSumB(b1DtBack, Drow1, "");
                                        SetMonthSumF(b1DtFooter, Drow1, "SUM");
                                        if (DtS2 != null)
                                        {
                                            sWhere = string.Format("TD_CODE='{0}'", Drow1["TD_CODE"].ToString());
                                            DataRow[] xDrow = DtS2.Select(sWhere);
                                            if (xDrow.Length > 0)
                                            {
                                                DataRow Drow2 = xDrow[0];
                                                SetFUND(b1DtBack2, b1DtBack, b1DtFooter, Drow2, "2");
                                                SUM_CEN_FUND2 = Convert.ToDouble(b1DtFooter.CEN_FUNDSUM2.ToString());
                                                SUM_LOC_FUND2 = Convert.ToDouble(b1DtFooter.LOC_FUNDSUM2.ToString());
                                                DtS2.Rows.Remove(Drow2);
                                            }
                                        }
                                        b1DtBack2.Order = Order;
                                        Order++;
                                        Lb1Dtback.Add(b1DtBack2);
                                        bNewNumYesNo = true;
                                    }
                                }

                            }
                            if (DtS2 != null)
                            {
                                DtS2.AcceptChanges();
                                sWhere = string.Format("CITY='{0}' And SUM_CATY='{1}'", DtCITY.Rows[x]["CITY"].ToString(), "C" + i.ToString());
                                DataRowS = DtS2.Select(sWhere);
                                if (DataRowS.Length > 0)
                                {
                                    for (int z = 0; z < DataRowS.Length; z++)
                                    {
                                        DataRow Drow2 = DataRowS[z];
                                        iShowTitle = iShowTitle + 1;
                                        iShowSerial = iShowSerial + 1;
                                        b1DtBack2 = new B1DtBack();
                                        b1DtBack2.CityCls = CityCls;
                                        b1DtBack.Order = Order;
                                        Order++;
                                        b1DtBack2.SERIAL = iShowSerial.ToString();
                                        b1DtBack2.SIP_TD_CODE = Drow2["SIP_TD_CODE"].ToString();
                                        b1DtBack2.TD_NAME = Drow2["TD_NAME"].ToString();
                                        b1DtBack2.FD_CODE = Drow2["FD_CODE"].ToString();
                                        b1DtBack2.CATY_DESC = Drow2["CATY_DESC"].ToString();
                                        b1DtBack2.ACC_DESC = Drow2["IPJ_DESC"].ToString();
                                        b1DtBack2.CEN_SP_RATIO = Drow2["CEN_SP_RATIO"].ToString();
                                        SetMonth(b1DtBack2, null);
                                        SetFUND(b1DtBack2, b1DtBack, b1DtFooter, Drow2, "2");
                                        SUM_CEN_FUND2 = SUM_CEN_FUND2 + Convert.ToDouble(b1DtFooter.CEN_FUNDSUM2.ToString());
                                        SUM_LOC_FUND2 = SUM_LOC_FUND2 + Convert.ToDouble(b1DtFooter.LOC_FUNDSUM2.ToString());
                                        SetMonthSumB(b1DtBack, null, "");
                                        SetMonthSumF(b1DtFooter, null, "SUM");
                                        Lb1Dtback.Add(b1DtBack2);
                                        bNewNumYesNo = true;
                                        DtS2.Rows.Remove(Drow2);
                                    }
                                }
                            }
                            if (bNewNumYesNo == true)
                            {
                                iShowTitle = iShowTitle + 1;
                                b1DtBack.Order = Order;
                                Order++;
                                Lb1Dtback.Add(b1DtBack);
                            }
                        }

                        b1DtBack = new B1DtBack();
                        b1DtBack.CityCls = CityCls;
                        b1DtBack.Order = Order;
                        Order++;
                        b1DtBack.TD_NAME = "總計";
                        b1DtBack.CEN_FUND1 = DtS1.Compute("SUM(CEN_FUND)", "CITY='" + DtCITY.Rows[x]["CITY"].ToString() + "'").ToString();
                        b1DtBack.LOC_FUND1 = DtS1.Compute("SUM(LOC_FUND)", "CITY='" + DtCITY.Rows[x]["CITY"].ToString() + "'").ToString();
                        b1DtBack.COST_FEE = DtS1.Compute("SUM(COST_FEE)", "CITY='" + DtCITY.Rows[x]["CITY"].ToString() + "'").ToString();
                        b1DtBack.CEN_FUND2 = SUM_CEN_FUND2.ToString();
                        b1DtBack.LOC_FUND2 = SUM_LOC_FUND2.ToString();

                        if (ReportController.NullList(b1DtBack2))
                        {
                            Lb1Dtback.Add(b1DtBack);
                            b1DtFooter.CEN_FUNDSUM = GetNumSub(b1DtFooter, b1DtFooter, "CEN_FUNDSUM2", "CEN_FUNDSUM1");
                            b1DtFooter.CityCls = CityCls;
                            Lb1DtFooter.Add(b1DtFooter);
                        }
                        CityCls++;
                    }
                    var a = Lb1Dtback;
                    var b = Lb1DtFooter;
                    break;
            }
        }

        private void SetMonthSumF(B1DtFooter b1DtFooter, DataRow Drow, string sExStr)
        {
            if (Drow != null)
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtFooter.GetType().GetProperty("M_" + i.ToString("00") + sExStr).SetValue(b1DtFooter, ReportController.ObjectToDouble(b1DtFooter.GetType().GetProperty("M_" + i.ToString("00") + sExStr)?.GetValue(b1DtFooter)) + StringToDouble(Drow["M_" + i.ToString("00")].ToString()));

                }
            }
            else
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtFooter.GetType().GetProperty("M_" + i.ToString("00") + sExStr).SetValue(b1DtFooter, ReportController.ObjectToDouble(b1DtFooter.GetType().GetProperty("M_" + i.ToString("00") + sExStr)?.GetValue(b1DtFooter)) + 0);
                }
            }
        }

        private void SetMonthSumB(B1DtBack b1DtBack, DataRow Drow, string sExStr)
        {
            if (Drow != null)
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtBack.GetType().GetProperty("M_" + i.ToString("00") + sExStr).SetValue(b1DtBack, ReportController.ObjectToDouble(b1DtBack.GetType().GetProperty("M_" + i.ToString("00") + sExStr)?.GetValue(b1DtBack)) + StringToDouble(Drow["M_" + i.ToString("00")].ToString()));

                }
            }
            else
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtBack.GetType().GetProperty("M_" + i.ToString("00") + sExStr).SetValue(b1DtBack, ReportController.ObjectToDouble(b1DtBack.GetType().GetProperty("M_" + i.ToString("00") + sExStr)?.GetValue(b1DtBack)) + 0);
                }
            }
        }

        private void SetMonth(B1DtBack b1DtBack2, DataRow Drow)
        {
            if (Drow != null)
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtBack2.GetType().GetProperty("M_" + i.ToString("00")).SetValue(b1DtBack2, StringToDouble(Drow["M_" + i.ToString("00")].ToString()));
                }
            }
            else
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtBack2.GetType().GetProperty("M_" + i.ToString("00")).SetValue(b1DtBack2, 0);
                }
            }
        }

        private void SetFUND(B1DtBack b2, B1DtBack b, B1DtFooter f, DataRow Drow, string sNumStr)
        {
            b2.GetType().GetProperty("CEN_FUND" + sNumStr).SetValue(b2, Drow["CEN_FUND"].ToString());
            b2.GetType().GetProperty("LOC_FUND" + sNumStr).SetValue(b2, Drow["LOC_FUND"].ToString());
            f.GetType().GetProperty("CEN_FUNDSUM" + sNumStr).SetValue(f, ReportController.ConvertDoubleValue(GetNumSumStrb1footer(f, Drow, "CEN_FUNDSUM" + sNumStr, "CEN_FUND")));
            f.GetType().GetProperty("LOC_FUNDSUM" + sNumStr).SetValue(f, ReportController.ConvertDoubleValue(GetNumSumStrb1footer(f, Drow, "LOC_FUNDSUM" + sNumStr, "LOC_FUND")));
            f.GetType().GetProperty("FUNDSUM" + sNumStr).SetValue(f, ReportController.ConvertDoubleValue(GetNumSumStrb1footer(f, Drow, "FUNDSUM" + sNumStr, "CEN_FUND")));
            f.GetType().GetProperty("FUNDSUM" + sNumStr).SetValue(f, ReportController.ConvertDoubleValue(GetNumSumStrb1footer(f, Drow, "FUNDSUM" + sNumStr, "LOC_FUND")));
            b.GetType().GetProperty("CEN_FUND" + sNumStr).SetValue(b, GetNumSumStrNew(b, Drow, "CEN_FUND" + sNumStr, "CEN_FUND"));
            b.GetType().GetProperty("LOC_FUND" + sNumStr).SetValue(b, GetNumSumStrNew(b, Drow, "LOC_FUND" + sNumStr, "LOC_FUND"));
        }

        private string GetNumSumStrb1footer(B1DtFooter f, DataRow Drow, string FidName1, string FidName2)
        {
            if (f.GetType().GetProperty(FidName1)?.GetValue(f) != null)
            {
                if (Drow[FidName2] != DBNull.Value)
                {
                    return Convert.ToString(Convert.ToDouble(f.GetType().GetProperty(FidName1)?.GetValue(f).ToString()) + Convert.ToDouble(Drow[FidName2].ToString()));
                }
                else
                {
                    return f.GetType().GetProperty(FidName1)?.GetValue(f).ToString();
                }
            }
            else
            {
                if (Drow[FidName2] != DBNull.Value)
                {
                    return Drow[FidName2].ToString();
                }
                else
                {
                    return "0";
                }
            }
        }

        private string GetNumSumStrNew(B1DtBack b, DataRow Drow, string FidName1, string FidName2)
        {
            if (b.GetType().GetProperty(FidName1)?.GetValue(b) != null)
            {
                if (Drow[FidName2] != DBNull.Value)
                {
                    return Convert.ToString(Convert.ToDouble(b.GetType().GetProperty(FidName1)?.GetValue(b).ToString()) + Convert.ToDouble(Drow[FidName2].ToString()));
                }
                else
                {
                    return b.GetType().GetProperty(FidName1)?.GetValue(b).ToString();
                }
            }
            else
            {
                if (Drow[FidName2] != DBNull.Value)
                {
                    return Drow[FidName2].ToString();
                }
                else
                {
                    return "0";
                }
            }
        }

        private string GetNumSub(B1DtFooter b, B1DtFooter Drow2, string FidName1, string FidName2)
        {
            if (Drow2.GetType().GetProperty(FidName2)?.GetValue(Drow2) != null)
            {
                if (b.GetType().GetProperty(FidName1)?.GetValue(b) != null)
                {
                    return Convert.ToString(Convert.ToDouble(Drow2.GetType().GetProperty(FidName2)?.GetValue(Drow2).ToString()) - Convert.ToDouble(b.GetType().GetProperty(FidName1)?.GetValue(b).ToString()));
                }
                else
                {
                    return Drow2.GetType().GetProperty(FidName2)?.GetValue(Drow2).ToString();
                }
            }
            else
            {
                if (b.GetType().GetProperty(FidName1)?.GetValue(b) != null)
                {
                    return "0";
                }
                else
                {
                    return "";
                }
            }
        }

        public Worksheet CopyWorksheet(Workbook workbook, Worksheet worksheet, string sourceSheetName, string newSheetName)
        {
            Worksheet sheetToCopy = workbook.Worksheets[sourceSheetName];
            int newSheetIndex = workbook.Worksheets.Add();
            worksheet = workbook.Worksheets[newSheetIndex];
            worksheet.Name = newSheetName;
            worksheet.Copy(sheetToCopy);
            return worksheet;
        }

        static string[] GenerateYearsArray(int startYear, int endYear)
        {
            int numberOfYears = endYear - startYear + 1;

            string[] yearsArray = new string[numberOfYears];

            for (int i = 0; i < numberOfYears; i++)
            {
                yearsArray[i] = (startYear + i).ToString();
            }

            return yearsArray;
        }

        void SetCellCenterAlign(Cells cells, string cellName)
        {
            Cell cell = cells[cellName];

            Style style = cell.GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;

            StyleFlag flag = new StyleFlag
            {
                HorizontalAlignment = true,
                VerticalAlignment = true
            };

            /// Fix: Replace ApplyStyle with SetStyle as per the Cell class definition  
            cell.SetStyle(style, flag);
        }

        private void SetCellLeftAlign(Cells cells, string cellName)
        {
            Style style = cells[cellName].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Left;   /// 靠左
            style.VerticalAlignment = TextAlignmentType.Top;      /// 如果想靠上
            cells[cellName].SetStyle(style);
        }

        public ActionResult State4_2()
        {
            ViewBag.State = "4_2";
            var query = User_Right(9);
            state9ViewModel CityAndColumn = new state9ViewModel();

            var Plan_Name = db2.PJV_PJDPCAT.OrderByDescending(m => m.SYR).Where(m => m.PJ_TYPE == "B").ToList();

            CityAndColumn.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            CityAndColumn.PJV_PJDPCAT = Plan_Name;
            CityAndColumn.OGCTITLE = db.OGCTITLE.ToList();

            return View(CityAndColumn);
        }

        /// <summary>
        /// A-2.細詳執行情形表: 日期格式處理
        /// 格式 => 114.01.12
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        string FormatROCDate(object val)
        {
            string s = val?.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Replace(".", ""); // 移除小數點
            if (s.Length == 7 || s.Length == 8)
                return $"{s.Substring(0, 3)}.{s.Substring(3, 2)}.{s.Substring(5, 2)}";
            return s; // 不符合格式就原樣輸出
        }

        /// <summary>
        /// B-1 修正預算作業表 : 報表資料來源
        /// </summary>
        /// <param name="pID"></param>
        /// <param name="SEQ_NO"></param>
        /// <param name="sACC_ID"></param>
        /// <param name="ADV_UID"></param>
        /// <param name="aDV_DP"></param>
        /// <param name="Dept"></param>
        /// <param name="Choose"></param>
        /// <returns></returns>
        public DataSet GetSP_RCD0010_UDF003_New(string pID, string SEQ_NO, string sACC_ID, string ADV_UID, string aDV_DP, string Dept, string Choose)
        {
            aDV_DP = aDV_DP.Replace(",", "','");
            aDV_DP = "'" + aDV_DP + "'";

            if (!string.IsNullOrEmpty(pID) && !string.IsNullOrEmpty(SEQ_NO) && !string.IsNullOrEmpty(aDV_DP))
            {
                using (var db = new CMEPMSEntities())
                {
                    var RetVal = new SqlParameter("RetVal", SqlDbType.Int);
                    RetVal.Direction = ParameterDirection.ReturnValue;

                    var pidParam = new SqlParameter("@PID", SqlDbType.Decimal) { Precision = 18, Scale = 0, Value = Convert.ToDecimal(pID) };
                    var seqParam = new SqlParameter("@SEQ_NO", SqlDbType.Decimal) { Size = 18, Value = Convert.ToDecimal(SEQ_NO) };
                    var advParam = new SqlParameter("@ADV_DP_LIST", SqlDbType.VarChar) { Size = 2000, Value = aDV_DP };
                    var accParam = new SqlParameter("@ACC_ID", SqlDbType.VarChar) { Size = 2000, Value = sACC_ID };
                    var uidParam = new SqlParameter("@ADV_UID", SqlDbType.VarChar) { Size = 2000, Value = ADV_UID };
                    var deptParam = new SqlParameter("@Dept", SqlDbType.VarChar) { Size = 2000, Value = Dept };
                    var chooseParam = new SqlParameter("@Choose", SqlDbType.VarChar) { Size = 2000, Value = Choose }; // 新加入的分署篩選

                    var command = db.Database.Connection.CreateCommand();
                    command.CommandText = "EXEC CMEPMS.dbo.SP_RCD0010_UDF003_New @PID,@SEQ_NO,@ADV_DP_LIST,@ACC_ID,@ADV_UID,@Dept,@Choose";
                    command.Parameters.Add(pidParam);
                    command.Parameters.Add(seqParam);
                    command.Parameters.Add(advParam);
                    command.Parameters.Add(accParam);
                    command.Parameters.Add(uidParam);
                    command.Parameters.Add(deptParam);
                    command.Parameters.Add(chooseParam);
                    command.Parameters.Add(RetVal);
                    command.CommandTimeout = 600;

                    /// 執行存儲過程並將結果填入 DataSet
                    var adapter = new SqlDataAdapter((SqlCommand)command);
                    var dataSet = new DataSet();
                    adapter.Fill(dataSet, "DataTable");

                    if (Convert.ToInt16(RetVal.Value) == 0)
                    {
                        return dataSet;
                    }

                    return null;
                }
            }

            DataSet newDataSet = new DataSet();
            DataTable newTable = new DataTable("DataTable");
            newTable.Columns.Add("COLUMN1", typeof(string));
            newTable.Rows.Add("未填入");
            newDataSet.Tables.Add(newTable);

            return newDataSet;
        }

        #region =============================================== 新版報表 ===============================================

        /// <summary>
        /// I-1.雨水例行性管考報表[經費調整表]
        /// </summary>
        /// <param name="AN_PID"></param>
        /// <param name="SYR"> 年份 </param> 
        /// <param name="SMN"> 月份 </param>
        /// <param name="connectionString"> 連線字串 </param>
        /// <param name="DEPT"> 部門編號 </param>
        /// <param name="PJ_NAME"> 工程名稱 </param>
        /// <param name="Button"> 輸出格式 </param>
        /// <see cref="GetRPT0001Data"/> 資料來源
        private void PrintRPT0001(decimal AN_PID, string SYR, string SMN, string connectionString, string DEPT, string PJ_NAME, string Button)
        {
            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            string SEQ_NO = "";
            string PSEQ_NO = "";
            string sFullFile = Server.MapPath("~/Templates/PJV0050.xlsx");
            var workbook = new Aspose.Cells.Workbook(sFullFile);
            Aspose.Cells.Worksheet sheet1 = workbook.Worksheets["Sheet1"];     // (已無用)
            Aspose.Cells.Worksheet sheet2 = workbook.Worksheets["Sheet2"];     // (已無用)
            Aspose.Cells.Worksheet sheet3 = workbook.Worksheets["Sheet3"];     // (已無用)
            Aspose.Cells.Worksheet sheet4 = workbook.Worksheets["Sheet4"];     // (已無用)
            Aspose.Cells.Worksheet sheet5 = workbook.Worksheets["Sheet5"];     // (已無用)
            Aspose.Cells.Worksheet sheet6 = workbook.Worksheets["Sheet6"];     // (已無用)
            Aspose.Cells.Worksheet sheet1_1 = workbook.Worksheets["Sheet1_1"];
            Aspose.Cells.Worksheet sheet1_2 = workbook.Worksheets["Sheet1_2"];
            Aspose.Cells.Worksheet sheet8 = workbook.Worksheets["Sheet8"];     // 總表
            Aspose.Cells.Worksheet sheet9 = workbook.Worksheets["Sheet9"];     // 經費調整表
            Aspose.Cells.Worksheet sheet10 = workbook.Worksheets["Sheet10"];   // 經費調整表
            Aspose.Cells.Worksheet sheet11 = workbook.Worksheets["Sheet11"];   // 署辦
            Aspose.Cells.Worksheet sheet12 = workbook.Worksheets["Sheet12"];   // 地辦
            Aspose.Cells.Worksheet sheet13 = workbook.Worksheets["Sheet13"];   // 業務費
            DataTable dt = new DataTable();    /// 單一資料表
            DataTable ditdt = new DataTable(); /// 單一資料表
            DataSet ds = new DataSet();        /// 包含多張資料表
            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");
            SEQ_NO = "0";
            PSEQ_NO = "";
            ds = RainReport_Source.GetRPT0001Data(AN_PID, SYR, SMN, connectionString, DEPT);
            ditdt = ds.Tables["ditdt"];

            #region -------------- NEW總表 --------------

            DataSet ds2 = RainReport_Source.I1REPORTDATA(AN_PID, SYR, SMN, connectionString, DEPT);
            DataTable dt1 = ds2.Tables["dt1"];
            DataTable dt2 = ds2.Tables["dt2"];
            sheet8.Name = "總表";
            RainReport_Source.RoutineRainwaterInspectionReport(workbook, sheet8, dt1, dt2, ditdt, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);

            #endregion

            #region -------------- NEW經費調整表/經費調整表(明細) --------------

            sheet9.Name = "經費調整表";
            sheet10.Name = "經費調整表_明細";
            DataTable dt3 = ds2.Tables["dt3"];
            DataTable dt4 = ds2.Tables["dt4"];
            RainReport_Source.FundingAdjustmentTable(workbook, sheet9, dt3, dt4, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);
            RainReport_Source.FundingAdjustmentTableDetails(workbook, sheet10, dt3, dt4, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);

            #endregion

            #region -------------- NEW署辦 --------------

            sheet11.Name = "署辦";
            RainReport_Source.DepartmentOfficeTable(workbook, sheet11, dt1, dt2, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);

            #endregion

            #region -------------- NEW地辦 --------------

            sheet12.Name = "地辦";
            RainReport_Source.LocalOfficeTable(workbook, sheet12, dt1, dt2, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);

            #endregion

            #region -------------- NEW業務費 --------------

            sheet13.Name = "業務費";
            RainReport_Source.BusinessExpensesTable(workbook, sheet13, dt1, dt2, dt3, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);

            #endregion

            #region -------------- (舊版)署辦 --------------
            dt = ds.Tables["dtp1"];
            Report_Source.CreateSheet(workbook, sheet3, dt, ditdt, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);
            #endregion

            #region -------------- (舊版)地辦 --------------
            dt = ds.Tables["dtp2"];
            Report_Source.CreateSheet(workbook, sheet4, dt, ditdt, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);
            #endregion

            #region -------------- (舊版)業務費 --------------
            dt = ds.Tables["dtp3"];
            Report_Source.CreateSheet(workbook, sheet5, dt, ditdt, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);
            #endregion

            #region -------------- (舊版)經費調整表-明細 --------------
            dt = ds.Tables["dtp4"];
            Report_Source.CreateSheet2(workbook, sheet2, dt, ditdt, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);
            #endregion

            #region -------------- (舊版)經費調整表(已無用) --------------
            dt = ds.Tables["dtp4"];
            Report_Source.CreateSheet3(workbook, sheet6, dt, ditdt, SYR, PSEQ_NO, SEQ_NO, PJ_NAME);
            #endregion

            dt = ds.Tables["dtp"];

            /// sSYR(當年度) 
            /// sPSYR(前年度)
            if (SYR != "")
            {
                sSYR = SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }

            workbook.Worksheets.RemoveAt("Sheet1_1");
            workbook.Worksheets.RemoveAt("Sheet1_2");
            workbook.Worksheets.RemoveAt("Sheet1");
            workbook.Worksheets.RemoveAt("Sheet2");
            workbook.Worksheets.RemoveAt("Sheet3");
            workbook.Worksheets.RemoveAt("Sheet4");
            workbook.Worksheets.RemoveAt("Sheet5");
            workbook.Worksheets.RemoveAt("Sheet6");
            sheet8.MoveTo(0);
            sheet9.MoveTo(1);
            sheet10.MoveTo(2);
            sheet11.MoveTo(3);
            sheet12.MoveTo(4);
            sheet13.MoveTo(5);

            string filename = "I-1.經費調整表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            string savepath = Server.MapPath("~/Temp/" + filename);
            workbook.Save(savepath, SaveFormat.Xlsx);

            if (Button == "1")
            {
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=" + filename + ".xlsx");
                Response.TransmitFile(savepath);
                Response.Flush();
                Response.End();
            }

            if (Button == "2")
            {
                using (var stream = new MemoryStream())
                {
                    var workbookdownload = new Workbook(savepath);
                    var worksheet = workbookdownload.Worksheets[0];
                    worksheet.AutoFitRows();
                    var saveOptions = new OdsSaveOptions();
                    workbookdownload.Save(stream, saveOptions);
                    stream.Seek(0, SeekOrigin.Begin);
                    var odsBytes = stream.ToArray();
                    var odsOutputFile = Server.MapPath("~/Temp/" + filename + ".ods");
                    System.IO.File.WriteAllBytes(odsOutputFile, odsBytes);

                    Response.Clear();
                    Response.ContentType = "application/vnd.oasis.opendocument.spreadsheet";
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + filename + ".ods");
                    Response.BinaryWrite(odsBytes);
                    Response.End();
                }
            }
        }

        /// <summary>
        /// I-1.雨水例行性管考報表[執行情形表]
        /// </summary>
        /// <param name="AN_PID"></param>
        /// <param name="SYR"></param>
        /// <param name="SMN"></param>
        /// <param name="connectionString"></param>
        /// <param name="DEPT"></param>
        /// <param name="PJ_NAME"></param>
        /// <param name="Button"></param>
        private void PrintRPT0002_NEW(decimal AN_PID, string SYR, string SMN, string connectionString, string DEPT, string PJ_NAME, string Button)
        {
            string templatePath = Server.MapPath("~/Templates/PJV0051.xlsx");
            string fileName = HttpUtility.UrlEncode("I-1.執行情形表", System.Text.Encoding.UTF8);
            Workbook workbook = new Workbook(templatePath);
            Worksheet sheet = workbook.Worksheets["執行情形表"];

            int StartCow = 6;
            int Item = 0;
            var Money1 = 0;
            var Money2 = 0;
            var Money3 = 0;
            var TOTAL_FUND2 = 0;

            double dCON_LEN_PER = 0;
            double dCON_FLOOD_PER = 0;
            double dPRE_CON_LEN = 0;
            double dREL_CON_LEN = 0;
            double dPRE_CON_FLOOD = 0;
            double dREL_CON_FLOOD = 0;
            double dPRE_CON_LEN_T = 0;
            double dREL_CON_LEN_T = 0;
            double dPRE_CON_FLOOD_T = 0;
            double dREL_CON_FLOOD_T = 0;

            /// 報表標題[填入案件名稱及期間]
            string sSYR = string.IsNullOrEmpty(SYR) ? (DateTime.Now.Year - 1911).ToString("000") : SYR;
            string sSMN = string.IsNullOrEmpty(SMN) ? DateTime.Now.Month.ToString("00") : SMN;
            string sDAY = Convert.ToDateTime((Convert.ToInt32(sSYR) + 1911) + "/" + sSMN + "/01").AddMonths(1).AddDays(-1).Day.ToString("00");

            sheet.Cells["A1"].PutValue($"{SYR}年度 {PJ_NAME} 執行情形表");
            sheet.Cells["A2"].PutValue($"中華民國{sSYR}年01月01日起至{sSYR}年{sSMN}月{sDAY}日止");

            DataView dvData = RainReport_Source.GetRPT0002Data(AN_PID, SYR, SMN, connectionString, DEPT);

            if (dvData == null)
            {
                throw new Exception("dvData == null，getRPT0002Data 沒有回傳 DataView");
            }

            if (dvData.Table == null)
            {
                throw new Exception("dvData.Table == null");
            }

            if (dvData.Table.Rows.Count == 0)
            {
                throw new Exception("dvData.Table.Rows.Count == 0，沒有資料列");
            }

            foreach (DataRow datarow in dvData.Table.Rows)
            {
                /// (L)欄位 => 前一年度核定費用 - 實支數費用
                var TD_NAME_LINQ = datarow["TD_NAME"].ToString();
                var LAST_SYR = (int.Parse(SYR) - 1).ToString();
                Money1 = (int)(db2.TDV_TDDCFTS.Where(m => m.TD_NAME == TD_NAME_LINQ && SYR == LAST_SYR).FirstOrDefault()?.CEN_FUND ?? 0); /// 前一年度核定費用
                Money2 = ((int)Source.Tools.StringToDouble(datarow["P_PLN_REL_COST"].ToString()));
                Money3 = Money1 - Money2; /// 前一年度核定費用 - 實支數費用

                /// (O)欄位 => 年度累計實際支用(加上暫付數)
                var Fund = db2.TDDSTCF.Where(m => m.TD_CODE == TD_NAME_LINQ && m.SYR == SYR)
                    .GroupBy(b => new { b.SYR, b.SMN })
                    .Select(g => new
                    {
                        SYR = g.Key.SYR,
                        SMN = g.Key.SMN,
                        COST_FEE = g.Sum(b => b.COST_FEE),
                        PPAY_FEE = g.Sum(b => b.PPAY_FEE)
                    })
                    .OrderByDescending(b => b.SYR).ThenByDescending(b => b.SMN).FirstOrDefault();

                var TOTAL_FUND = (Fund?.COST_FEE + Fund?.PPAY_FEE) ?? 0;
                long TOTALFUND = (long)TOTAL_FUND + TOTAL_FUND2;

                dPRE_CON_LEN = Convert.ToDouble(Tools.StringToDouble(datarow["PRE_CON_LEN"].ToString()));
                dREL_CON_LEN = Convert.ToDouble(Tools.StringToDouble(datarow["REL_CON_LEN"].ToString()));
                dPRE_CON_FLOOD = Convert.ToDouble(Tools.StringToDouble(datarow["PRE_CON_FLOOD"].ToString()));
                dREL_CON_FLOOD = Convert.ToDouble(Tools.StringToDouble(datarow["REL_CON_FLOOD"].ToString()));
                dPRE_CON_LEN_T += dPRE_CON_LEN;
                dREL_CON_LEN_T += dREL_CON_LEN;
                dPRE_CON_FLOOD_T += dPRE_CON_FLOOD;
                dREL_CON_FLOOD_T += dREL_CON_FLOOD;

                if (dPRE_CON_LEN != 0) { dCON_LEN_PER = (dREL_CON_LEN / dPRE_CON_LEN) * 100; }
                if (dPRE_CON_FLOOD != 0) { dCON_FLOOD_PER = (dREL_CON_FLOOD / dPRE_CON_FLOOD) * 100; }

                sheet.Cells.CopyRow(sheet.Cells, 6, StartCow);
                sheet.Cells[StartCow, 0].PutValue(++Item);                                                                                                              /// (A)標案項次
                sheet.Cells[StartCow, 1].PutValue(datarow["CITY"].ToString());                                                                                          /// (B)縣市別
                sheet.Cells[StartCow, 2].PutValue(datarow["SIP_TD_CODE"].ToString());                                                                                   /// (C)工程編號
                sheet.Cells[StartCow, 3].PutValue(datarow["TD_NAME"].ToString());                                                                                       /// (D)標案名稱
                sheet.Cells[StartCow, 4].PutValue(Convert.ToDouble(datarow["TD_FUND"].ToString()));                                                                     /// (E)核定總工程經費
                sheet.Cells[StartCow, 5].PutValue(Convert.ToDouble(datarow["CEN_SP_RATIO"]));                                                                           /// (F)補助比例(%)
                sheet.Cells[StartCow, 6].PutValue(Convert.ToDouble(datarow["CEN_FUND_1"].ToString()));                                                                  /// (G)106 中央款
                sheet.Cells[StartCow, 7].PutValue(Convert.ToDouble(datarow["CEN_FUND_2"].ToString()));                                                                  /// (H)107 中央款
                sheet.Cells[StartCow, 8].PutValue(Convert.ToDouble(datarow["CEN_FUND"].ToString()));                                                                    /// (I)中央款
                sheet.Cells[StartCow, 9].PutValue(Convert.ToDouble(datarow["LOC_FUND"].ToString()));                                                                    /// (J)地方款
                sheet.Cells[StartCow, 10].PutValue(Convert.ToDouble(datarow["OCR_FUND"].ToString()));                                                                   /// (K)發生權責數
                sheet.Cells[StartCow, 11].PutValue(Convert.ToDouble(Money3));                                                                                           /// (L)前一年度核定費用 - 實支數費用
                sheet.Cells[StartCow, 12].PutValue(Convert.ToDouble(datarow["CEN_FUND"].ToString()));                                                                   /// (M)可支用預算(D) --
                sheet.Cells[StartCow, 13].PutValue(Convert.ToDouble(datarow["L_PLN_COST"].ToString()));                                                                 /// (N)標案分配數(E) --
                sheet.Cells[StartCow, 14].PutValue(Convert.ToDouble(TOTAL_FUND));                                                                                       /// (O)年度累計實際支用(加上暫付數)(F)
                sheet.Cells[StartCow, 15].PutValue(datarow["L_UNPAY_COST"] == DBNull.Value ? 0d : Convert.ToDouble(datarow["L_UNPAY_COST"].ToString()));                /// (P)標案應付未付數(G)
                sheet.Cells[StartCow, 16].PutValue(datarow["SAVE_FEE"] == DBNull.Value ? 0d : Convert.ToDouble(datarow["SAVE_FEE"].ToString()));                        /// (Q)標案節餘數(H)
                // sheet.Cells[StartCow, 17].PutValue(datarow[""].ToString());                                                                                          /// (R)標案執行率(F + G + H) / E [報表自動計算]
                // sheet.Cells[StartCow, 18].PutValue(datarow[""].ToString());                                                                                          /// (S)標案達成率(F + G + H) / D [報表自動計算]
                sheet.Cells[StartCow, 19].PutValue(datarow["PRE_TK_103"].ToString());                                                                                   /// (T)標案設計日期預定
                sheet.Cells[StartCow, 20].PutValue(datarow["REL_TK_103"].ToString());                                                                                   /// (U)標案設計日期實際
                sheet.Cells[StartCow, 21].PutValue(datarow["PRE_TK_105"].ToString());                                                                                   /// (V)標案發包日期預定
                sheet.Cells[StartCow, 22].PutValue(datarow["REL_TK_105"].ToString());                                                                                   /// (W)標案發包日期實際
                sheet.Cells[StartCow, 23].PutValue(Convert.ToDouble(datarow["PRE_SCHD"].ToString()));                                                                   /// (X)標案工程進度預定%(I)
                sheet.Cells[StartCow, 24].PutValue(Convert.ToDouble(datarow["REL_SCHD"].ToString()));                                                                   /// (Y)標案工程進度實際%(J)
                // sheet.Cells[StartCow, 25].PutValue(datarow[""].ToString());                                                                                          /// (Z)標案工程進度差異%(J - I) [報表自動計算]
                sheet.Cells[StartCow, 26].PutValue(Convert.ToDouble(datarow["PRE_CON_LEN"].ToString()));                                                                /// (AA)箱涵設計長(m)
                sheet.Cells[StartCow, 27].PutValue(Convert.ToDouble(datarow["REL_CON_LEN"].ToString()));                                                                /// (AB)累積箱涵施作長度(m)
                sheet.Cells[StartCow, 28].PutValue(Convert.ToDouble(dCON_LEN_PER));                                                                                     /// (AC)實際施作進度%
                sheet.Cells[StartCow, 29].PutValue(Convert.ToDouble(datarow["PRE_CON_FLOOD"].ToString()));                                                              /// (AD)設計滯洪量(m³)
                sheet.Cells[StartCow, 30].PutValue(Convert.ToDouble(datarow["REL_CON_FLOOD"].ToString()));                                                              /// (AE)累積設計施作滯洪量(m³)
                sheet.Cells[StartCow, 31].PutValue(Convert.ToDouble(dCON_FLOOD_PER));                                                                                   /// (AF)實際施作進度%
                sheet.Cells[StartCow, 32].PutValue(datarow["EXTRA_PRE_SDT"].ToString());                                                                                /// (AG)標案完成日期預定
                sheet.Cells[StartCow, 33].PutValue(datarow["EXTRA_PRE_EDT"].ToString());                                                                                /// (AH)標案完成日期實際
                sheet.Cells[StartCow, 34].PutValue(datarow["BUDGET_FEE_B"] == DBNull.Value ? 0d : Convert.ToDouble(datarow["BUDGET_FEE_B"]));                           /// (AI)總經費
                sheet.Cells[StartCow, 35].PutValue(datarow["BUDGET_FEE_B_0101"] == DBNull.Value ? 0d : Convert.ToDouble(datarow["BUDGET_FEE_B_0101"]));                 /// (AJ)施工費
                sheet.Cells[StartCow, 36].PutValue(datarow["BUDGET_FEE_B_N0101"] == DBNull.Value ? 0d : Convert.ToDouble(datarow["BUDGET_FEE_B_N0101"].ToString()));    /// (AK)間接費用
                sheet.Cells[StartCow, 37].PutValue(datarow["BUDGET_FEE_C"] == DBNull.Value ? 0d : Convert.ToDouble(datarow["BUDGET_FEE_C"].ToString()));                /// (AL)總經費
                sheet.Cells[StartCow, 38].PutValue(datarow["BUDGET_FEE_C_0101"] == DBNull.Value ? 0d : Convert.ToDouble(datarow["BUDGET_FEE_C_0101"].ToString()));      /// (AM)施工費
                sheet.Cells[StartCow, 39].PutValue(datarow["EXECT_DESC"].ToString());                                                                                   /// (AN)備註
                StartCow++;
            }

            /// 合計
            if (dvData.Table.Rows.Count > 0)
            {
                int dataStartRow = 6;          // 第一筆資料列
                int dataEndRow = StartCow - 1; // 最後一筆資料列
                int totalRow = StartCow;       // 合計列位置

                sheet.Cells.CopyRow(sheet.Cells, dataStartRow, totalRow);
                sheet.Cells.ClearContents(totalRow, 0, totalRow, sheet.Cells.MaxColumn);
                sheet.Cells[totalRow, 2].PutValue("合計");
                sheet.Cells[totalRow, 4].Formula = $"=SUM(E{dataStartRow + 1}:E{dataEndRow + 1})";
                sheet.Cells[totalRow, 6].Formula = $"=SUM(G{dataStartRow + 1}:G{dataEndRow + 1})";
                sheet.Cells[totalRow, 7].Formula = $"=SUM(H{dataStartRow + 1}:H{dataEndRow + 1})";
                sheet.Cells[totalRow, 8].Formula = $"=SUM(I{dataStartRow + 1}:I{dataEndRow + 1})";
                sheet.Cells[totalRow, 9].Formula = $"=SUM(J{dataStartRow + 1}:J{dataEndRow + 1})";
                sheet.Cells[totalRow, 11].Formula = $"=SUM(L{dataStartRow + 1}:L{dataEndRow + 1})";
                sheet.Cells[totalRow, 12].Formula = $"=SUM(M{dataStartRow + 1}:M{dataEndRow + 1})";
                sheet.Cells[totalRow, 13].Formula = $"=SUM(N{dataStartRow + 1}:N{dataEndRow + 1})";
                sheet.Cells[totalRow, 14].Formula = $"=SUM(O{dataStartRow + 1}:O{dataEndRow + 1})";
                sheet.Cells[totalRow, 15].Formula = $"=SUM(P{dataStartRow + 1}:P{dataEndRow + 1})";
                sheet.Cells[totalRow, 16].Formula = $"=SUM(Q{dataStartRow + 1}:Q{dataEndRow + 1})";
                sheet.Cells[totalRow, 34].Formula = $"=SUM(AI{dataStartRow + 1}:AI{dataEndRow + 1})";
                sheet.Cells[totalRow, 35].Formula = $"=SUM(AJ{dataStartRow + 1}:AJ{dataEndRow + 1})";
                sheet.Cells[totalRow, 36].Formula = $"=SUM(AK{dataStartRow + 1}:AK{dataEndRow + 1})";
                sheet.Cells[totalRow, 37].Formula = $"=SUM(AL{dataStartRow + 1}:AL{dataEndRow + 1})";
                sheet.Cells[totalRow, 38].Formula = $"=SUM(AM{dataStartRow + 1}:AM{dataEndRow + 1})";
            }

            #region 報表樣式設定
            sheet.Cells.SetColumnWidth(0, 6);   // A 項次
            sheet.Cells.SetColumnWidth(1, 6);   // B 縣市別
            sheet.Cells.SetColumnWidth(2, 26);  // C 
            sheet.Cells.SetColumnWidth(3, 26);  // D
            sheet.Cells.SetColumnWidth(4, 15);  // E
            sheet.Cells.SetColumnWidth(5, 8);   // F
            sheet.Cells.SetColumnWidth(6, 15);  // G 
            sheet.Cells.SetColumnWidth(7, 15);  // H
            sheet.Cells.SetColumnWidth(8, 15);  // I
            sheet.Cells.SetColumnWidth(9, 15);  // J
            sheet.Cells.SetColumnWidth(10, 8);  // K
            sheet.Cells.SetColumnWidth(11, 15); // L
            sheet.Cells.SetColumnWidth(12, 15); // M
            sheet.Cells.SetColumnWidth(13, 15); // N
            sheet.Cells.SetColumnWidth(14, 15); // O
            sheet.Cells.SetColumnWidth(15, 15); // P
            sheet.Cells.SetColumnWidth(16, 15); // Q
            sheet.Cells.SetColumnWidth(17, 15); // R
            sheet.Cells.SetColumnWidth(18, 15); // S
            sheet.Cells.SetColumnWidth(19, 10); // T
            sheet.Cells.SetColumnWidth(20, 10); // U
            sheet.Cells.SetColumnWidth(21, 10); // V
            sheet.Cells.SetColumnWidth(22, 10); // W
            sheet.Cells.SetColumnWidth(23, 10); // X
            sheet.Cells.SetColumnWidth(24, 10); // Y
            sheet.Cells.SetColumnWidth(25, 10); // Z
            sheet.Cells.SetColumnWidth(26, 10); // AA
            sheet.Cells.SetColumnWidth(27, 10); // AB
            sheet.Cells.SetColumnWidth(28, 10); // AC
            sheet.Cells.SetColumnWidth(29, 10); // AD
            sheet.Cells.SetColumnWidth(30, 10); // AE
            sheet.Cells.SetColumnWidth(31, 10); // AF
            sheet.Cells.SetColumnWidth(32, 10); // AG
            sheet.Cells.SetColumnWidth(33, 10); // AH
            sheet.Cells.SetColumnWidth(34, 15); // AI
            sheet.Cells.SetColumnWidth(35, 15); // AJ
            sheet.Cells.SetColumnWidth(36, 15); // AK
            sheet.Cells.SetColumnWidth(37, 15); // AL
            sheet.Cells.SetColumnWidth(38, 15); // AM
            sheet.Cells.SetColumnWidth(39, 38); // AN
            sheet.PageSetup.Zoom = 70;
            #endregion

            /// 20260529 版本修改 => 刪除 (G) | (H) 兩欄, 起始位置 6, 後數 2 欄(包含 6)
            sheet.Cells.DeleteColumns(6, 2, true);

            #region 報表產出
            string savePath = Server.MapPath("~/Temp/" + fileName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            workbook.Save(savePath, SaveFormat.Xlsx);

            if (Button == "1")
            {
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=" + fileName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                Response.TransmitFile(savePath);
                Response.Flush();
                Response.End();
            }
            else if (Button == "2")
            {
                using (var stream = new MemoryStream())
                {
                    Workbook odsBook = new Workbook(savePath);
                    odsBook.Save(stream, new OdsSaveOptions());
                    stream.Seek(0, SeekOrigin.Begin);

                    byte[] odsBytes = stream.ToArray();
                    Response.Clear();
                    Response.ContentType = "application/vnd.oasis.opendocument.spreadsheet";
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + fileName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".ods");
                    Response.BinaryWrite(odsBytes);
                    Response.End();
                }
            }
            #endregion
        }

        /// <summary>
        /// I-1.雨水例行性管考報表[前瞻計畫執行情形表]
        /// </summary>
        /// <param name="AN_PID"></param>
        /// <param name="SYR"></param>
        /// <param name="SMN"></param>
        /// <param name="connectionString"></param>
        /// <param name="DEPT"></param>
        /// <param name="PJ_NAME"></param>
        /// <param name="Button"></param>
        private void PrintRPT0003_3(decimal AN_PID, string SYR, string SMN, string connectionString, string DEPT, string PJ_NAME, string Button)
        {
            int ITEM = 1;
            int templateRow = 4; /// Excel 第5列
            int rowIndex = 4;
            string sSYR = (DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (DateTime.Now.Month).ToString("00");
            string sPSYR = (DateTime.Now.Year - 1911 - 1).ToString("000");
            string sDAY = "31";

            #region Excel 報表樣板獲取

            string sFileName = "";
            string sFullFile = Server.MapPath("~/Templates/PJV0052.xlsx");
            FileInfo oFileInfo = new FileInfo(sFullFile);

            Workbook workbook = new Workbook(sFullFile);
            Worksheet sheet1 = workbook.Worksheets["執行情形表"];

            #endregion

            #region Excel 資料流

            if (SYR != "")
            {
                sSYR = SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }

            if (SMN != "")
            {
                sSMN = SMN;
            }

            sDAY = Convert.ToDateTime((Convert.ToInt32(sSYR) + 1911).ToString() + "/" + sSMN + "/01").AddMonths(1).AddDays(-1).Day.ToString("00");

            sheet1.Cells["A1"].PutValue($"{SYR}年度 {PJ_NAME} 前瞻執行情形表");
            sheet1.Cells["A2"].PutValue("中華民國" + sSYR + "年01月01日起至" + sSYR + "年" + sSMN + "月" + sDAY + "日止");
            sheet1.Cells["J3"].PutValue(sPSYR + "年");
            sheet1.Cells["Q3"].PutValue(sSYR + "年度分配預算");
            sheet1.Cells["R3"].PutValue("預估至" + sSYR + "年底實現數(元)");
            sheet1.Cells["S3"].PutValue(sSYR + "年度撥款數");
            sheet1.Cells["T3"].PutValue(sPSYR + "~" + sSYR + "年度累計撥款數");
            sheet1.Cells["U3"].PutValue(sSYR + "年度已撥款現現數(元)");
            sheet1.Cells["V3"].PutValue(sSYR + ".01~" + sSYR + "." + sSMN + "累計實現數(元)");

            /// 取得資料
            DataTable dt = RainReport_Source.GetRPT0003NEWData(AN_PID, SYR, SMN, connectionString, DEPT);

            foreach (DataRow row in dt.Rows)
            {
                sheet1.Cells.CopyRow(sheet1.Cells, templateRow, rowIndex);

                sheet1.Cells[rowIndex, 0].PutValue(ITEM++);
                sheet1.Cells[rowIndex, 1].PutValue(row["SIP_TD_CODE"].ToString(), false);
                sheet1.Cells[rowIndex, 2].PutValue(row["FD_CODE"].ToString(), false);
                sheet1.Cells[rowIndex, 3].PutValue(0);
                sheet1.Cells[rowIndex, 4].PutValue(row["CITY"].ToString());
                sheet1.Cells[rowIndex, 5].PutValue(PJ_NAME);
                sheet1.Cells[rowIndex, 6].PutValue(row["TD_NAME"].ToString());
                sheet1.Cells[rowIndex, 7].PutValue(Convert.ToDouble(row["CEN_SP_RATIO"].ToString()));
                sheet1.Cells[rowIndex, 8].PutValue(Convert.ToDouble(row["CEN_FUND"].ToString()));
                sheet1.Cells[rowIndex, 9].PutValue(row["PCEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(row["PCEN_FUND"].ToString()));
                sheet1.Cells[rowIndex, 10].PutValue(row["LAST_PPAY_FEE"] == DBNull.Value ? 0 : Convert.ToDouble(row["LAST_PPAY_FEE"].ToString()));
                sheet1.Cells[rowIndex, 11].PutValue(row["AE_CTM"] == DBNull.Value ? 0 : Convert.ToDouble(row["AE_CTM"].ToString()));
                sheet1.Cells[rowIndex, 12].PutValue(0);
                sheet1.Cells[rowIndex, 13].PutValue(0);
                sheet1.Cells[rowIndex, 14].PutValue(0);
                sheet1.Cells[rowIndex, 15].PutValue("");
                /// sheet1.Cells[rowIndex, 16].PutValue(row["NCEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(row["NCEN_FUND"].ToString()));
                sheet1.Cells[rowIndex, 16].PutValue(Convert.ToDouble(row["CEN_FUND"].ToString()));
                sheet1.Cells[rowIndex, 17].PutValue(row["NAE_CTM"] == DBNull.Value ? 0 : Convert.ToDouble(row["NAE_CTM"].ToString()));
                sheet1.Cells[rowIndex, 18].PutValue(row["NOW_PPAY_FEE"] == DBNull.Value ? 0 : Convert.ToDouble(row["NOW_PPAY_FEE"].ToString()));
                sheet1.Cells[rowIndex, 20].PutValue(row["NOW_PPAY_FEE"] == DBNull.Value ? 0 : Convert.ToDouble(row["NOW_PPAY_FEE"].ToString()));
                sheet1.Cells[rowIndex, 21].PutValue("");
                sheet1.Cells[rowIndex, 22].PutValue(row["TK105_DT"] == DBNull.Value ? "" : FormatROCDate(row["TK105_DT"]));             /// 發包後請款工程期程
                sheet1.Cells[rowIndex, 23].PutValue(row["TK141_DT"] == DBNull.Value ? "" : FormatROCDate(row["TK141_DT"]));             /// 進度達 40％ 期程
                sheet1.Cells[rowIndex, 24].PutValue(row["TK107_PRE_DT"] == DBNull.Value ? "" : FormatROCDate(row["TK107_PRE_DT"]));     /// 完工日期 (預定)
                sheet1.Cells[rowIndex, 25].PutValue(row["TK107_REL_DT"] == DBNull.Value ? "" : FormatROCDate(row["TK107_REL_DT"]));     /// 完工日期 (實際)
                sheet1.Cells[rowIndex, 26].PutValue(row["TK110_REL_DT"] == DBNull.Value ? "" : FormatROCDate(row["TK110_REL_DT"]));     /// 結案
                sheet1.Cells[rowIndex, 27].PutValue(row["PRE_SCHD"] == DBNull.Value ? "" : row["PRE_SCHD"]);                            /// 工程進度(%)月底(預定)
                sheet1.Cells[rowIndex, 28].PutValue(row["REL_SCHD"] == DBNull.Value ? "" : row["REL_SCHD"]);                            /// 工程進度(%)月底(實際)
                sheet1.Cells[rowIndex, 29].PutValue("總工程經費：" + Convert.ToDouble(row["BUDGET_FEE_B0"].ToString()) + "\r\n" + "契約金額：" + Convert.ToDouble(row["BUDGET_FEE_C0"].ToString()) + "\r\n" + "辦理情形：" + row["EXECT_DESC"].ToString());    /// 總工程經費 / 契約金額 / 辦理情形

                rowIndex++;
            }

            /// 總計欄位
            sheet1.Cells.CopyRow(sheet1.Cells, templateRow, rowIndex);
            sheet1.Cells.ClearContents(rowIndex, 0, rowIndex, 29);

            sheet1.Cells[rowIndex, 5].PutValue("總計");

            /// CEN_FUND 欄位
            string sumFormula = $"=SUM(I{templateRow + 1}:I{rowIndex})";
            sheet1.Cells[rowIndex, 8].Formula = sumFormula;

            string sumFormulaJ = $"=SUM(J{templateRow + 1}:J{rowIndex})";
            sheet1.Cells[rowIndex, 9].Formula = sumFormulaJ;

            string sumFormulaK = $"=SUM(K{templateRow + 1}:K{rowIndex})";
            sheet1.Cells[rowIndex, 10].Formula = sumFormulaK;

            string sumFormulaL = $"=SUM(L{templateRow + 1}:L{rowIndex})";
            sheet1.Cells[rowIndex, 11].Formula = sumFormulaL;

            string sumFormulaM = $"=SUM(M{templateRow + 1}:M{rowIndex})";
            sheet1.Cells[rowIndex, 12].Formula = sumFormulaM;

            string sumFormulaN = $"=SUM(N{templateRow + 1}:N{rowIndex})";
            sheet1.Cells[rowIndex, 13].Formula = sumFormulaN;

            string sumFormulaO = $"=SUM(O{templateRow + 1}:O{rowIndex})";
            sheet1.Cells[rowIndex, 14].Formula = sumFormulaO;

            string sumFormulaQ = $"=SUM(Q{templateRow + 1}:Q{rowIndex})";
            sheet1.Cells[rowIndex, 16].Formula = sumFormulaQ;

            string sumFormulaR = $"=SUM(R{templateRow + 1}:R{rowIndex})";
            sheet1.Cells[rowIndex, 17].Formula = sumFormulaR;

            string sumFormulaS = $"=SUM(S{templateRow + 1}:S{rowIndex})";
            sheet1.Cells[rowIndex, 18].Formula = sumFormulaS;

            string sumFormulaT = $"=SUM(T{templateRow + 1}:T{rowIndex})";
            sheet1.Cells[rowIndex, 19].Formula = sumFormulaT;

            string sumFormulaU = $"=SUM(U{templateRow + 1}:U{rowIndex})";
            sheet1.Cells[rowIndex, 20].Formula = sumFormulaU;

            #endregion

            #region Excel 格式設定

            /// 設定內容置中
            Style style = workbook.CreateStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;

            StyleFlag flag = new StyleFlag();
            flag.HorizontalAlignment = true;

            sheet1.Cells.Columns[1].ApplyStyle(style, flag); /// B
            sheet1.Cells.Columns[2].ApplyStyle(style, flag); /// C

            sheet1.Cells.SetColumnWidth(0, 6);      /// A 項次
            sheet1.Cells.SetColumnWidth(1, 30);     /// B 縣市別
            sheet1.Cells.SetColumnWidth(2, 30);     /// C 
            sheet1.Cells.SetColumnWidth(3, 5);      /// D
            sheet1.Cells.SetColumnWidth(4, 10);     /// E
            sheet1.Cells.SetColumnWidth(5, 30);     /// F
            sheet1.Cells.SetColumnWidth(6, 30);     /// G 
            sheet1.Cells.SetColumnWidth(7, 5);      /// H
            sheet1.Cells.SetColumnWidth(8, 15);     /// I
            sheet1.Cells.SetColumnWidth(9, 16);     /// J
            sheet1.Cells.SetColumnWidth(10, 16);    /// K
            sheet1.Cells.SetColumnWidth(11, 16);    /// L
            sheet1.Cells.SetColumnWidth(12, 16);    /// M
            sheet1.Cells.SetColumnWidth(13, 16);    /// N
            sheet1.Cells.SetColumnWidth(14, 16);    /// O
            sheet1.Cells.SetColumnWidth(15, 16);    /// P
            sheet1.Cells.SetColumnWidth(16, 16);    /// Q
            sheet1.Cells.SetColumnWidth(17, 16);    /// R
            sheet1.Cells.SetColumnWidth(18, 16);    /// S
            sheet1.Cells.SetColumnWidth(19, 16);    /// T
            sheet1.Cells.SetColumnWidth(20, 16);    /// U
            sheet1.Cells.SetColumnWidth(21, 16);    /// V
            sheet1.Cells.SetColumnWidth(22, 16);    /// W
            sheet1.Cells.SetColumnWidth(23, 16);    /// X
            sheet1.Cells.SetColumnWidth(24, 16);    /// Y
            sheet1.Cells.SetColumnWidth(25, 16);    /// Z
            sheet1.Cells.SetColumnWidth(26, 16);    /// AA
            sheet1.Cells.SetColumnWidth(27, 16);    /// AB
            sheet1.Cells.SetColumnWidth(28, 16);    /// AC
            sheet1.Cells.SetColumnWidth(29, 50);    /// AD
            sheet1.PageSetup.Zoom = 70;

            #endregion

            #region Excel/ODS 報表產出

            /// 儲存位置 Temp 資料夾
            sFileName = HttpUtility.UrlEncode("I-1.前瞻計畫執行情形表", System.Text.Encoding.UTF8);
            string savepath = Server.MapPath("~/Temp/" + sFileName + ".xlsx");
            workbook.Save(savepath, SaveFormat.Xlsx);

            if (Button == "1")
            {
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=" + sFileName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                Response.TransmitFile(savepath);
                Response.Flush();
                Response.End();
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    Workbook workbookdownload = new Workbook(savepath);

                    var worksheet = workbookdownload.Worksheets[0];
                    worksheet.AutoFitRows();

                    var saveOptions = new OdsSaveOptions();
                    workbookdownload.Save(stream, saveOptions);

                    stream.Seek(0, SeekOrigin.Begin);
                    var odsBytes = stream.ToArray();

                    Response.Clear();
                    Response.ContentType = "application/vnd.oasis.opendocument.spreadsheet";
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + sFileName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".ods");
                    Response.BinaryWrite(odsBytes);
                    Response.End();
                }
            }

            #endregion
        }

        /// <summary>
        /// A-2.細項執行情形表
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="SYR"></param>
        /// <param name="YearWrite"></param>
        /// <param name="SMN"></param>
        /// <param name="CityList"></param>
        /// <param name="funding"></param>
        /// <param name="ODS"></param>
        /// <param name="Site"></param>
        /// <param name="group"></param>
        /// <returns></returns>
        public ActionResult ReportPrint2_22(string PID, string SYR, string YearWrite, string SMN, string[] CityList, string funding, string ODS, string[] Site, string group)
        {
            var TDMFORM = db3.TDMFORM;
            var TDV_TDDCFTS = db2.TDV_TDDCFTS;
            DataTable SourceTable = new DataTable();

            /// 錯誤警告
            #region =================================================================================================
            if (CityList == null || CityList.Length == 0)
            {
                string alertMessage = "請至少選擇一個縣市";
                return Content($"<script>alert('{alertMessage}'); window.location.href = document.referrer;</script>");
            }
            else if (SMN == null || SMN.Length == 0)
            {
                string alertMessage2 = "請選擇資料月份";
                return Content($"<script>alert('{alertMessage2}'); window.location.href = document.referrer;</script>");
            }
            #endregion ==============================================================================================

            /// A-2.細項執行情形表來源
            SourceTable.Merge(Source.RainReport_Source.getBillionProject08_NEW(PID, SYR, SMN, CityList, funding, connectionString));

            /// 取得 Aspose 授權, 開啟 EXCEL範本
            #region =================================================================================================
            var path = Server.MapPath("~/Templates/PJV0010_03(Rain).xlsx");
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense(Server.MapPath("~/Aspose/Aspose.Total.lic"));
            var workbook = new Aspose.Cells.Workbook(path);
            var ws = workbook.Worksheets["Sheet1"];
            BulidReport_A2(workbook, SourceTable, PID, SYR, SMN);
            #endregion ==============================================================================================

            /// 報表輸出(PDF/ODS)
            #region =================================================================================================

            string filename = "A-2.細項執行情形表" + DateTime.Now.ToString("yyyyMMddHHmmss");

            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");

            workbook.Save(savepath, SaveFormat.Xlsx);

            if (ODS == "1")
            {
                var odsSavePath = Server.MapPath("~/Temp/" + filename + ".ods");
                workbook.Save(odsSavePath, SaveFormat.ODS);

                using (var stream = new FileStream(odsSavePath, FileMode.Open))
                {
                    var bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, (int)stream.Length);
                    return File(bytes, "application/vnd.oasis.opendocument.spreadsheet", filename + ".ods");
                }
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Seek(0, SeekOrigin.Begin);
                    var bytes = stream.ToArray();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename + ".xlsx");
                }
            }

            #endregion ==============================================================================================

        }

        private void BulidReport_A2(Workbook workbook, DataTable source, string PID, string SYR, string SMN)
        {
            /// Debug 測試程式碼
            #region ==================================================================
            Debug.WriteLine("BulidReport_A2");

            foreach (DataRow rows in source.Rows)
            {
                foreach (DataColumn cols in source.Columns)
                {
                    Debug.Write($"{cols.ColumnName}:{rows[cols]} \t");
                }
                Debug.WriteLine("");
            }
            #endregion ===============================================================

            int templateRow = 4;              /// 樣板(第 4 列)
            int currentRow = templateRow + 1; /// 填值(第 5 列)
            int Number = 1;
            Worksheet ws = workbook.Worksheets["Sheet1"];

            string ReportTitle = GetReportTitle(PID, connectionString);
            ws.Cells["A1"].PutValue($"{SYR}年度 {ReportTitle} 內政部辦理細項執行情形");
            ws.Cells["AB2"].PutValue($"資料截至:{SYR}年{SMN}月");

            for (int i = 0; i < source.Rows.Count; i++)
            {
                DataRow dr = source.Rows[i];
                ws.Cells.CopyRow(ws.Cells, templateRow - 1, currentRow - 1);                                                          /// 複製樣板行格式
                ws.Cells.ClearContents(currentRow - 1, 0, currentRow - 1, ws.Cells.MaxColumn);                                        /// 清空內容，保留格式
                ws.Cells["A" + currentRow].PutValue(Number);                                                                          /// 項次
                ws.Cells["B" + currentRow].PutValue(dr["CATY_NAME"]);                                                                 /// 計畫說明 CATY_NAME                        
                ws.Cells["C" + currentRow].PutValue(double.TryParse(dr["CATY_FUND"]?.ToString(), out double C) ? (C / 1000) : 0);     /// 分支計畫核定經費總計(仟元) CATY_FUND
                ws.Cells["D" + currentRow].PutValue(dr["TD_CITY"]);                                                                   /// 縣市 TD_CITY
                ws.Cells["E" + currentRow].PutValue(dr["IPJ_DESC"]);                                                                  /// 辦理層級 IPJ_DESC
                ws.Cells["F" + currentRow].PutValue(dr["CEN_RATIO"]);                                                                 /// 補助比例 CEN_RATIO
                ws.Cells["G" + currentRow].PutValue(dr["TD_NAME"]);                                                                   /// 工作名稱 TD_NAME
                ws.Cells["H" + currentRow].PutValue(dr["TD_FINISH_DESC"]);                                                            /// 結案情形 TD_FINISH_DESC
                ws.Cells["I" + currentRow].PutValue(double.TryParse(dr["LAST_PPAY_FEE"]?.ToString(), out double I) ? (I / 1000) : 0); /// 歷年累積撥款金額(仟元)
                ws.Cells["J" + currentRow].PutValue(double.TryParse(dr["ALL_PPAY_FEE"]?.ToString(), out double J) ? (J / 1000) : 0);  /// 當年累積撥款金額(仟元)        
                ws.Cells["K" + currentRow].PutValue(double.TryParse(dr["CEN_FUND"]?.ToString(), out double K) ? (K / 1000) : 0);      /// 當年度核定經費(仟元)                       
                ws.Cells["L" + currentRow].PutValue(double.TryParse(dr["L_FUND"]?.ToString(), out double L) ? (L / 1000) : 0);        /// 依進度可撥經費(仟元)
                ws.Cells["M" + currentRow].PutValue(FormatROCDate(dr["REL_TK102"]));                                                  /// 基設核定日期
                ws.Cells["N" + currentRow].PutValue(double.TryParse(dr["START_FUND"]?.ToString(), out double N) ? (N / 1000) : 0);    /// 基設核定經費
                ws.Cells["O" + currentRow].PutValue(FormatROCDate(dr["PRE_TK103"]));                                                  /// 預定細設完成日期
                ws.Cells["P" + currentRow].PutValue(FormatROCDate(dr["REL_TK103"]));                                                  /// 實際細設完成日期
                ws.Cells["Q" + currentRow].PutValue(FormatROCDate(dr["PRE_TK105"]));                                                  /// 預定發包日期
                ws.Cells["R" + currentRow].PutValue(FormatROCDate(dr["REL_TK105"]));                                                  /// 實際發包日期
                ws.Cells["S" + currentRow].PutValue(dr["PRE_SCHD"]);                                                                  /// 預定工程進度
                ws.Cells["T" + currentRow].PutValue(dr["REL_SCHD"]);                                                                  /// 實際工程進度
                ws.Cells["U" + currentRow].PutValue(dr["SCHD_DIFF"]);                                                                 /// 進度差異
                ws.Cells["V" + currentRow].PutValue(FormatROCDate(dr["PRE_TK107"]));                                                  /// 預定完工日期
                ws.Cells["W" + currentRow].PutValue(FormatROCDate(dr["REL_TK107"]));                                                  /// 實際完工日期
                // ws.Cells["X" + currentRow].PutValue(dr[""]);                                                                       /// 預計雨水下水道建設長度(m)
                // ws.Cells["Y" + currentRow].PutValue(dr[""]);                                                                       /// 實際雨水下水道建設長度(m)
                ws.Cells["Z" + currentRow].PutValue(FormatROCDate(dr["PRE_TK108"]));                                                  /// 預定正驗日期
                ws.Cells["AA" + currentRow].PutValue(FormatROCDate(dr["REL_TK108"]));                                                 /// 實際正驗日期
                ws.Cells["AB" + currentRow].PutValue(dr["EXECT_DESC"]);                                                               /// 辦理情形或落後原因
                ws.Cells["AC" + currentRow].PutValue(dr["DELAY_RESN"]);                                                               /// 落後改善對策
                ws.Cells["AD" + currentRow].PutValue(double.TryParse(dr["B_FUND"]?.ToString(), out double AD) ? (AD / 1000) : 0);     /// 總工程經費(仟元)
                ws.Cells["AE" + currentRow].PutValue(dr["CTDELAY_RESN"]);                                                             /// 備註
                ws.Cells["AF" + currentRow].PutValue(dr["DSG_BUSS_NAME"]);                                                            /// 設計單位
                ws.Cells["AG" + currentRow].PutValue(dr["BUSS_NAME"]);                                                                /// 廠商名稱

                Number++;
                currentRow++;
            }

            #region A-2.細項執行情形表[合計欄位]

            /// 新增"合計欄位"
            int LastRow = currentRow;

            /// 合計欄位前置作業: 複製樣板, 清空內容
            ws.Cells.CopyRow(ws.Cells, templateRow - 1, LastRow - 1);
            ws.Cells.ClearContents(LastRow - 1, 0, LastRow - 1, ws.Cells.MaxColumn);

            ws.Cells["B" + LastRow].PutValue("合計");
            ws.Cells["C" + LastRow].Formula = $"=SUM(C{templateRow + 1}:C{LastRow - 1})";
            ws.Cells["I" + LastRow].Formula = $"=SUM(I{templateRow + 1}:I{LastRow - 1})";
            ws.Cells["J" + LastRow].Formula = $"=SUM(J{templateRow + 1}:J{LastRow - 1})";
            ws.Cells["K" + LastRow].Formula = $"=SUM(K{templateRow + 1}:K{LastRow - 1})";
            ws.Cells["L" + LastRow].Formula = $"=SUM(L{templateRow + 1}:L{LastRow - 1})";
            ws.Cells["N" + LastRow].Formula = $"=SUM(N{templateRow + 1}:N{LastRow - 1})";
            ws.Cells["AD" + LastRow].Formula = $"=SUM(AD{templateRow + 1}:AD{LastRow - 1})";

            ws.Cells.SetColumnWidth(0, 6);    // A 項次
            ws.Cells.SetColumnWidth(1, 30);   // B 計畫說明
            ws.Cells.SetColumnWidth(2, 18);   // C
            ws.Cells.SetColumnWidth(3, 10);   // D
            ws.Cells.SetColumnWidth(4, 12);   // E
            ws.Cells.SetColumnWidth(5, 8);    // F
            ws.Cells.SetColumnWidth(6, 28);   // G 工作名稱
            ws.Cells.SetColumnWidth(7, 15);   // H
            ws.Cells.SetColumnWidth(8, 18);   // I
            ws.Cells.SetColumnWidth(9, 18);   // J
            ws.Cells.SetColumnWidth(10, 18);  // K
            ws.Cells.SetColumnWidth(11, 18);  // L
            ws.Cells.SetColumnWidth(12, 15);  // M
            ws.Cells.SetColumnWidth(13, 15);  // N
            ws.Cells.SetColumnWidth(14, 15);  // O
            ws.Cells.SetColumnWidth(15, 15);  // P
            ws.Cells.SetColumnWidth(16, 15);  // Q
            ws.Cells.SetColumnWidth(17, 15);  // R
            ws.Cells.SetColumnWidth(18, 15);  // S
            ws.Cells.SetColumnWidth(19, 15);  // T
            ws.Cells.SetColumnWidth(20, 15);  // U
            ws.Cells.SetColumnWidth(21, 15);  // V
            ws.Cells.SetColumnWidth(22, 15);  // W
            ws.Cells.SetColumnWidth(25, 15);  // Z
            ws.Cells.SetColumnWidth(26, 15);  // AA
            ws.Cells.SetColumnWidth(27, 45);  // AB
            ws.Cells.SetColumnWidth(28, 18);  // AC
            ws.Cells.SetColumnWidth(29, 18);  // AD
            ws.Cells.SetColumnWidth(30, 18);  // AE
            ws.Cells.SetColumnWidth(31, 18);  // AF
            ws.Cells.SetColumnWidth(32, 18);  // AG

            #endregion

            /// 文字自動換行設定
            /// 需要換行的欄位列表 : B, G, AB, AC, AE, AF, AG
            int[] wrapColumns = { 1, 6, 27, 28, 30, 31, 32 };

            StyleFlag flag = new StyleFlag();
            flag.WrapText = true;

            foreach (int col in wrapColumns)
            {
                for (int row = templateRow; row < currentRow; row++)
                {
                    var cell = ws.Cells[row, col];
                    Style style = cell.GetStyle();
                    style.IsTextWrapped = true;
                    cell.SetStyle(style, flag);
                }
            }

            ws.AutoFitRows();
            ws.PageSetup.Zoom = 80;
        }

        /// <summary>
        /// A-4.跨年度細項執行情形表
        /// </summary>
        /// <param name="StartYear"></param>
        /// <param name="EndYear"></param>
        /// <param name="plan"></param>
        /// <param name="CITY"></param>
        /// <param name="Row"></param>
        /// <param name="Field"></param>
        /// <param name="COLUMN1"></param>
        /// <param name="FormatFile"></param>
        /// <returns></returns>
        public ActionResult ReportPrint4_2(string StartYear, string EndYear, string[] plan, string[] CITY, string Row, string Field, string[] COLUMN1, string FormatFile)
        {
            Trace.WriteLine($"StartYear = {StartYear}");
            Trace.WriteLine($"EndYear = {EndYear}");
            Trace.WriteLine($"Row = {Row}");
            Trace.WriteLine($"Field = {Field}");
            Trace.WriteLine($"FormatFile = {FormatFile}");

            Trace.WriteLine($"Plan = ");
            foreach (var item in plan)
            {
                Trace.Write($" {item}");
            }

            Trace.WriteLine($"CITY = ");
            foreach (var item in CITY)
            {
                Trace.Write($" {item}");
            }

            Trace.WriteLine($"COLUMN1 = ");
            foreach (var item in COLUMN1)
            {
                Trace.Write($" {item}");
            }

            var path = Server.MapPath("~/Templates/表1.xlsx");

            var file = new FileInfo(path);

            List<string> modifiedCities = new List<string>(CITY);

            if (Array.Exists(CITY, element => element == "新北市"))
            {
                modifiedCities.Add("台北縣");
            }

            if (Array.Exists(CITY, element => element == "桃園市"))
            {
                modifiedCities.Add("桃園縣");
            }

            if (Array.Exists(CITY, element => element == "台中市"))
            {
                modifiedCities.Add("台中縣");
            }

            if (Array.Exists(CITY, element => element == "台南市"))
            {
                modifiedCities.Add("台南縣");
            }

            if (Array.Exists(CITY, element => element == "高雄市"))
            {
                modifiedCities.Add("高雄縣");
            }

            CITY = modifiedCities.ToArray();

            string[] COLUMN = new string[COLUMN1.Length + 1];
            COLUMN[0] = "年分";
            Array.Copy(COLUMN1, 0, COLUMN, 1, COLUMN1.Length);
            var dataTable = new DataTable();

            /// 20241119 修改項目 => 計算總和相關參數設定
            /// =================================================
            var dict = new Dictionary<string, decimal>();
            dict.Add("中央款總和", 0);
            dict.Add("累積撥款數總和", 0);
            HashSet<string> procItem = new HashSet<string>();
            HashSet<string> procFD_CODE = new HashSet<string>();
            HashSet<string> procTD_CODE = new HashSet<string>();
            HashSet<string> procTD_CODE2 = new HashSet<string>();
            /// =================================================

            /// 取各年度區間各年度資料
            string[] allYears = GenerateYearsArray(Source.Tools.StringToInt(StartYear), Source.Tools.StringToInt(EndYear));
            foreach (var columnName in COLUMN)
            {
                if (columnName == "當年度核定數")
                {
                    foreach (var year in allYears)
                    {
                        dataTable.Columns.Add(year + "年核定數");

                        dict.Add(year + "年核定數總和", 0); /// 計算總和相關參數設定
                    }
                }
                else
                {
                    dataTable.Columns.Add(columnName);
                }
            }

            var OGCDEPT = db.OGCDEPT.ToList();
            var CTDITFD = db2.CTDITFD.Where(m => m.FIT_CODE == "11").ToList();
            var TDMFORM = db3.TDMFORM.ToList();

            /// CMETMS.dbo.TDDBGLS : 總工程費來源
            /// 20250630 修改項目
            /// REC_TP = C : 契約資料
            /// REC_TP = B : 預算資料
            /// OUTSRC_FG = 1 : 已發包
            /// 先抓取契約資料, 若無, 則抓取預算資料
            var list = db3.TDDBGLS
                    .Where(m => m.REC_TP == "B" || m.REC_TP == "C") /// 抓所有 C（契約）和 B（預算）
                    .ToList(); /// 先查回記憶體再分段處理，避免 EF 無法轉譯

            var TDDBGLS = list
                .GroupBy(m => m.TD_CODE)
                .Select(g =>
                {
                    /// 優先抓 OUTSRC_FG = 1 (代表已發包)
                    var contractList = g.Where(x => x.OUTSRC_FG == true);
                    var contractTotal = contractList.Sum(x => x.BUDGET_FEE ?? 0);

                    /// 如果 contract 為 0，就改抓 REC_TP = B (預算資料)
                    var budgetList = g.Where(x => x.REC_TP == "B");
                    var budgetTotal = budgetList.Sum(x => x.BUDGET_FEE ?? 0);

                    return new
                    {
                        TD_CODE = g.Key,
                        ChosenSource = contractTotal > 0 ? "OUTSRC_FG = 1" : "REC_TP = B",
                        Total = contractTotal > 0 ? contractTotal : budgetTotal
                    };
                }).ToList();


            var TDDSDWB_CONTRATE = db3.TDDSDWB.Where(m => m.IT_NAME == "發包日期").ToList();
            var TDDSDWB_START = db3.TDDSDWB.Where(m => m.IT_CODE == "TK107").ToList();
            var TDDSDWB_END = db3.TDDSDWB.Where(m => m.IT_CODE == "TK107").ToList();

            string Year = "";
            string Master_Plan_Name = "";
            string Plan_Name = "";

            /// 抓取每個年度最後月份的經費
            var query = (from t in db2.TDDSTCF.ToList()
                         group t by new { t.FD_CODE, t.SYR } into g
                         select new
                         {
                             FD_CODE = g.Key.FD_CODE,
                             SYR = g.Key.SYR,
                             SMN = g.OrderByDescending(t => t.SMN).First().SMN,
                             COST_FEE = g.OrderByDescending(t => t.SMN).First().COST_FEE,
                             PPAY_FEE = g.OrderByDescending(t => t.SMN).First().PPAY_FEE,
                             TD_CODE1 = g.OrderByDescending(t => t.SMN).First().TD_CODE1
                         }).OrderBy(m => m.FD_CODE).ThenBy(m => m.SYR).ToList();

            var query_insert_column = query.Select((item, index) => new TDDSTCF_INSERT_CLOMNS
            {
                FD_CODE = item.FD_CODE,
                SYR = item.SYR,
                SMN = item.SMN,
                COST_FEE = item.COST_FEE,
                PPAY_FEE = item.PPAY_FEE,
                TD_CODE1 = item.TD_CODE1,
                MATH = (index > 0 && item.FD_CODE == query[index - 1].FD_CODE)
                        ? item.COST_FEE - query[index - 1].COST_FEE : item.COST_FEE
            }).ToList();


            var query2 = from t in db3.TDDMSCD.ToList()
                         group t by new { t.TD_CODE, t.SYR } into g
                         select new
                         {
                             TD_CODE = g.Key.TD_CODE,
                             SYR = g.Key.SYR,
                             SMN = g.Max(t => t.SMN),
                             PRE_SCHD = g.Max(t => t.PRE_SCHD),
                             REL_SCHD = g.Max(t => t.REL_SCHD),
                             EXECT_DESC = g.Max(t => t.EXECT_DESC),
                             DELAY_RESN = g.Max(t => t.DELAY_RESN),
                             CTDELAY_RESN = g.Max(t => t.CTDELAY_RESN),
                         };

            /// 20241223 ==> 移除已被刪除的分項內容對應資料
            var TDV_TDDCFTS = db2.TDV_TDDCFTS.Where(m => m.YR_FUND != 0 || (m.FCAL_FG ?? 0 == 1)).Where(s => s.CATY_TYPE != null).ToList();

            var TDV_TDDCFTS2 = (TDV_TDDCFTS.Select(m => new
            {
                m.PID_SYR, /// 20241119 修改項目
                m.FD_CODE, /// 20241119 修改項目
                SYR_FD_CODE = m.PID_SYR + m.FD_CODE, /// 20241119 修改項目
                m.PID,
                m.TD_CODE,
                m.TD_NAME,
                m.CITY,
                m.TOWN,
                m.IPJ_CODE,
                m.IPJ_DESC,
                m.ADV_DP,
                m.TID,
                m.CEN_SP_RATIO,
                m.CEN_FUND,
                SYR_TD_CODE = m.PID_SYR + m.TD_CODE
            }
            )).ToList();

            /// 20241119 修改項目
            /// 取得工程編號各個年度累積的撥款資料
            var financialDataQuery = query
                .GroupBy(f => new { f.TD_CODE1, f.SYR })
                .Select(g => new
                {
                    g.Key.TD_CODE1,
                    g.Key.SYR,
                    SUM_FEE = g.Sum(f => (f.COST_FEE ?? 0) + (f.PPAY_FEE ?? 0)),
                    SYR_TD_CODE = g.Key.SYR + g.Key.TD_CODE1
                });

            for (int w = 0; w < plan.Count(); w++)
            {
                string[] sArray;

                sArray = plan[w].Split('+');

                /// 20240131 修改項目 => 全部資料抓最後一年
                Year = EndYear;

                Master_Plan_Name = sArray[0];

                Plan_Name = sArray[1];

                var PJMFORM = db2.PJMFORM.Where(m => m.SYR.CompareTo(EndYear) <= 0 && m.SYR.CompareTo(StartYear) >= 0 && m.PJ_NAME == Plan_Name).ToList();

                #region ======================== 本署補助經費(中央款)(元) ========================

                /// 20260417 修改 => 本署補助經費(中央款)
                /// 111 ~ 114 (進行累加) / 114 ~ 114 (算單年度)
                /// 計算指定計畫(PID)對應年份金額
                var pidList = PJMFORM.Select(x => (decimal?)x.PID).ToList();

                var TDV_TDDCFTS_SUM = TDV_TDDCFTS
                    .Where(m => m.PID != null
                             && pidList.Contains(m.PID)
                             && !string.IsNullOrEmpty(m.PID_SYR)
                             && m.PID_SYR.CompareTo(StartYear) >= 0
                             && m.PID_SYR.CompareTo(EndYear) <= 0)
                    .GroupBy(m => m.TD_CODE)
                    .Select(p => new
                    {
                        TD_CODE = p.Key,
                        Total = p.Sum(m => m.CEN_FUND)
                    })
                    .ToList();

                #endregion ====================================================================

                var TDDSTCF = query_insert_column.Where(m => m.SYR == Year).ToList(); /// 這個是某年度的累積撥款數

                while (TDDSTCF.Count == 0 && StringToDouble(Year) > StringToDouble(StartYear))
                {
                    TDDSTCF = query_insert_column.Where(m => m.SYR == Year).ToList();
                    double Year_double = StringToDouble(Year);
                    if (TDDSTCF.Count == 0)
                    {
                        Year_double--;
                        Year = Year_double.ToString();
                    }
                }

                /// 計畫下的每個工程
                int yearNumber = int.Parse(Year);

                var TDDSTCF_TOTAL = query.Where(m => int.Parse(m.SYR) <= yearNumber).GroupBy(m => m.TD_CODE1).Select(p => new
                {
                    TDCODE1 = p.Key,
                    Total = p.Sum(m => m.COST_FEE) + p.Sum(m => m.PPAY_FEE)
                }).ToList();

                var TDDMSCD = query2.Where(m => m.SYR == Year).ToList();

                var SOURCE_MERGE = (from a in PJMFORM
                                    join b in TDV_TDDCFTS2 on a.PID equals b.PID into Joinb
                                    from b in Joinb.DefaultIfEmpty()
                                    join c in TDV_TDDCFTS_SUM on b?.TD_CODE equals c?.TD_CODE into Joinc
                                    from c in Joinc.DefaultIfEmpty()
                                    join d in TDMFORM on b?.TD_CODE equals d.TD_CODE into Joind
                                    from d in Joind.DefaultIfEmpty()
                                    join e in OGCDEPT on b?.ADV_DP equals e.DEPT_NO into Joine
                                    from e in Joine.DefaultIfEmpty()
                                    join f in TDDBGLS on b?.TD_CODE equals f.TD_CODE into Joinf
                                    from f in Joinf.DefaultIfEmpty()
                                    join g in financialDataQuery on b?.SYR_TD_CODE equals g.SYR_TD_CODE into Joing
                                    from g in Joing.DefaultIfEmpty()
                                    join i in TDDSDWB_CONTRATE on b?.TD_CODE equals i.TD_CODE into contrateJoin
                                    from i in contrateJoin.DefaultIfEmpty()
                                    join m in TDDSDWB_START on b?.TD_CODE equals m.TD_CODE into startJoin
                                    from m in startJoin.Where(x => x.IT_CODE == "TK107").DefaultIfEmpty()
                                    join n in TDDSDWB_END on b?.TD_CODE equals n.TD_CODE into endJoin
                                    from n in endJoin.Where(x => x.IT_CODE == "TK107").DefaultIfEmpty()
                                    join l in TDDMSCD on b?.TD_CODE equals l.TD_CODE into query2Join
                                    from l in query2Join.DefaultIfEmpty()
                                    where b != null
                                    select new CTDMNDSmergeTDV_TDDCFTS
                                    {
                                        FD_CODE = b?.FD_CODE, /// 20241119 修改項目
                                        PJ_NAME = Plan_Name,
                                        PID = b?.PID != null ? (decimal)b.PID : 0,
                                        TD_NAME = b.TD_NAME,
                                        TD_CITY = b.CITY,
                                        TD_TOWN = b.TOWN,
                                        TD_CODE = b?.TD_CODE,
                                        SYR = b?.SYR_TD_CODE.Substring(0, 3),
                                        IPJ_CODE = b.IPJ_CODE,
                                        IPJ_DESC = b.IPJ_DESC, /// 辦理層級
                                        ADV_DP = b.ADV_DP,
                                        TID = b?.TID != null ? (decimal)b.TID : 0,
                                        FUND = c?.Total != null ? (decimal)c.Total : 0, /// 中央款
                                        PERCENT_FUND = b.CEN_SP_RATIO != null ? (decimal)b.CEN_SP_RATIO : 0, /// 工程費比例
                                        TOTAL_ALL_BUDGET_FEE = b?.CEN_FUND != null ? (decimal)b.CEN_FUND : 0,
                                        TD_LINE = d?.TD_LINE,
                                        TD_NOTE = d?.TD_NOTE,
                                        TD_MEMO = d?.TD_MEMO,
                                        SMP_DEPT_NAME = e?.SMP_DEPT_NAME,
                                        TOTAL_FUND = f?.Total != null ? (decimal)f?.Total : 0, /// 總工程費(元)
                                        COST_FEE = g?.SUM_FEE ?? 0,
                                        PRE_CONTRACT_DATE = i?.PRE_EDT,
                                        REL_CONTRACT_DATE = i?.REL_EDT,
                                        PRE_START_DATE = m?.PRE_EDT,
                                        REL_START_DATE = m?.REL_EDT,
                                        PRE_END_DATE = n?.PRE_EDT,
                                        REL_END_DATE = n?.REL_EDT,
                                        PRE_SCHD = l?.PRE_SCHD != null ? (decimal)l.PRE_SCHD : 0,
                                        REL_SCHD = l?.REL_SCHD != null ? (decimal)l.REL_SCHD : 0,
                                        EXECT_DESC = l?.EXECT_DESC,
                                        DELAY_RESN = l?.DELAY_RESN,
                                        CTDELAY_RESN = l?.CTDELAY_RESN,
                                        PROGRESS_DIFFERENCE = (l?.PRE_SCHD != null ? (decimal)l.PRE_SCHD : 0) - (l?.REL_SCHD != null ? (decimal)l.REL_SCHD : 0),
                                    }).GroupBy(m => m.TD_CODE)
                                    .Select(group => group.OrderByDescending(m => m.SYR).First())
                                    .ToList();


                string MAX_IPJ_CODE = SOURCE_MERGE.Max(m => m.IPJ_CODE);
                SOURCE_MERGE = SOURCE_MERGE.Where(m => CITY.Any(value => m.TD_CITY.Contains(value))).ToList();
                SOURCE_MERGE = SOURCE_MERGE.OrderBy(m => m.TD_CITY).ToList();


                foreach (var item in SOURCE_MERGE)
                {
                    var row = dataTable.NewRow();
                    /// 20241119 修改項目
                    var TD_CODE = item.GetType().GetProperty("TD_CODE")?.GetValue(item) ?? 0;
                    var SYR = item.GetType().GetProperty("SYR")?.GetValue(item) ?? "";

                    /// 同工程編號同年度已有資料則顯示一筆就好
                    string sKey = TD_CODE + "-" + SYR;
                    if (!procItem.Contains(sKey))
                    {
                        procItem.Add(sKey);
                    }
                    else
                    {
                        continue;
                    }

                    /// 20241119 修改項目
                    var value = item.GetType().GetProperty("SYR")?.GetValue(item);
                    row["年分"] = value;
                    foreach (var columnName in COLUMN)
                    {
                        if (columnName == "年分")
                        {
                            continue;
                        }
                        else if (columnName.Contains("本署補助經費(中央款)(元)"))
                        {
                            /// 20241119 修改項目
                            /// 計算本署補助經費(中央款)(元)總和
                            var Fund = item.GetType().GetProperty("FUND")?.GetValue(item) ?? 0;

                            row["本署補助經費(中央款)(元)"] = Fund;

                            string idKey = TD_CODE.ToString();

                            if (!procTD_CODE.Contains(idKey)) /// 如果沒處理過，則加總並標記為已處理
                            {
                                dict["中央款總和"] += (decimal)Fund;
                                procTD_CODE.Add(idKey);
                            }
                        }
                        else if (columnName.Contains("年度核定數"))
                        {

                            /// 抓取年度區間內各年度的核定數(只會取得該年度之前的)
                            /// 每筆工程編號每年只會有一筆資料
                            foreach (var year in allYears)
                            {

                                //decimal Fund = TDV_TDDCFTS2.Where(m => m.TD_CODE == TD_CODE.ToString() && m.PID_SYR == year.ToString())
                                //    .Select(s => s.CEN_FUND).DefaultIfEmpty(0).Sum();

                                decimal Fund = TDV_TDDCFTS2
                                        .Where(m => m.TD_CODE == TD_CODE.ToString()
                                                 && m.PID_SYR == year.ToString()
                                                 && m.PID != null
                                                 && pidList.Contains(m.PID))
                                        .Select(s => s.CEN_FUND)
                                        .DefaultIfEmpty(0)
                                        .Sum();

                                row[year + "年核定數"] = Fund;

                                /// 計算xxx年度核定數總和
                                string idYearKey = TD_CODE + "-" + year;

                                if (!procFD_CODE.Contains(idYearKey))
                                {
                                    dict[year + "年核定數總和"] += Fund;
                                    procFD_CODE.Add(idYearKey);
                                }
                            }
                        }
                        else if (columnName == "累積撥款數") /// 20241103
                        {
                            /// 計算每年累積撥款數
                            /// 每筆工程編號每年只會有一筆資料
                            if (!string.IsNullOrEmpty(SYR.ToString()))
                            {
                                var preSum = financialDataQuery.Where(s => s.TD_CODE1 == TD_CODE.ToString())
                                .Where(s => s.SYR.CompareTo(SYR) <= 0).Sum(s => s.SUM_FEE);

                                row["累積撥款數"] = preSum;

                                /// 計算累積撥款數總和
                                if (SYR.ToString() == financialDataQuery.Where(s => s.TD_CODE1 == TD_CODE.ToString()).OrderByDescending(s => s.SYR).FirstOrDefault()?.SYR)
                                {
                                    string idYearKey = TD_CODE.ToString();

                                    if (!procTD_CODE2.Contains(idYearKey))
                                    {
                                        dict["累積撥款數總和"] += preSum;
                                        procTD_CODE2.Add(idYearKey);
                                    }
                                }
                            }
                        }
                        else
                        {
                            var COLUMN_IDNAME = db.OGCTITLE.Where(m => m.TITLE == columnName).FirstOrDefault().NAME;

                            value = item.GetType().GetProperty(COLUMN_IDNAME)?.GetValue(item);

                            row[columnName] = value;
                        }
                    }
                    dataTable.Rows.Add(row);
                }
            }

            var rowcount = dataTable.Rows.Count;

            #region ================================== Aspose.Cells套件使用 ==================================

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);

            var workbook = new Aspose.Cells.Workbook(path);
            var ws1 = workbook.Worksheets["工程細項執行表(縣市在列)"];
            var Column_length = Report_Source.NumberToColumnLetter(COLUMN.Length + allYears.Length);

            var A = ws1.Cells.CreateRange("A1:A1");
            var B = ws1.Cells.CreateRange("A1:" + Column_length + (rowcount + 2));
            var R = ws1.Cells.CreateRange("A1", Column_length + (rowcount + 2 + 1));

            Aspose.Cells.Style style1 = workbook.CreateStyle();
            Aspose.Cells.Style borderstyle = ws1.Workbook.CreateStyle();

            Aspose.Cells.Border outerBorder = borderstyle.Borders[BorderType.TopBorder];
            outerBorder.LineStyle = CellBorderType.Thin;

            outerBorder = borderstyle.Borders[BorderType.BottomBorder];
            outerBorder.LineStyle = CellBorderType.Thin;

            outerBorder = borderstyle.Borders[BorderType.LeftBorder];
            outerBorder.LineStyle = CellBorderType.Thin;

            outerBorder = borderstyle.Borders[BorderType.RightBorder];
            outerBorder.LineStyle = CellBorderType.Thin;

            Aspose.Cells.Border innerBorder = borderstyle.Borders[BorderType.Vertical];
            innerBorder.LineStyle = CellBorderType.Thin;

            innerBorder = borderstyle.Borders[BorderType.Horizontal];
            innerBorder.LineStyle = CellBorderType.Thin;
            borderstyle.IsTextWrapped = true;
            borderstyle.HorizontalAlignment = TextAlignmentType.Center;

            R.SetStyle(borderstyle);

            B.Copy(A);

            int k = 0;

            for (int i = 2; i < COLUMN.Length + 2; i++)
            {
                var title = COLUMN[i - 2];

                if (title == "當年度核定數")
                {
                    foreach (var year in allYears)
                    {
                        var ABC = Report_Source.NumberToColumnLetter(i + k);

                        var a = ws1.Cells[ABC + "2"];

                        a.PutValue(year + "年核定數");
                        a.GetStyle().Borders.SetStyle(CellBorderType.Thin);

                        if (k < allYears.Length - 1)
                        {
                            k++;
                        }
                    }
                }
                else
                {
                    var ABC = Report_Source.NumberToColumnLetter(i + k);

                    var a = ws1.Cells[ABC + "2"];

                    a.PutValue(title);
                    a.GetStyle().Borders.SetStyle(CellBorderType.Thin);
                }
            }

            int j = 3;

            for (var n = 0; n < rowcount; n++)
            {
                var a0 = ws1.Cells["A" + j];

                a0.PutValue(dataTable.Rows[j - 3]["工程地點(址)"]);

                k = 0;

                for (int i = 2; i < COLUMN.Length + 2; i++)
                {
                    var ABC = Report_Source.NumberToColumnLetter(i + k);
                    var a1 = ws1.Cells[ABC + j];
                    Aspose.Cells.Style a1_Style = a1.GetStyle();
                    a1_Style.Number = 5;
                    var title = COLUMN[i - 2];

                    if (title == "序號")
                    {
                        a1.PutValue(j - 2);
                    }
                    else if (title == "補助對象")
                    {
                        a1.PutValue(dataTable.Rows[j - 3]["工程地點(址)"] + "政府");
                    }
                    else if (title == "落後原因")
                    {
                        a1.PutValue(dataTable.Rows[j - 3]["落後原因"]);
                    }
                    else if (title == "總工程費(元)")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");

                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["總工程費(元)"]);

                        a1.PutValue(value);
                        a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    else if (title == "本署補助經費(中央款)(元)")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");

                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["本署補助經費(中央款)(元)"]);

                        a1.PutValue(value);
                        a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    else if (title == "工程費比例")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");

                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["工程費比例"]);

                        a1.PutValue(value);
                        a1_Style.Custom = "0.00_);[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    else if (title == "當年度核定數")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");

                        foreach (var year in allYears)
                        {
                            decimal value = 0;

                            ABC = Report_Source.NumberToColumnLetter(i + k);

                            a1 = ws1.Cells[ABC + j];
                            a1_Style = a1.GetStyle();
                            a1_Style.Number = 5;

                            try
                            {
                                var test = dataTable.Rows[j - 3][year + "年核定數"];
                                value = Source.Tools.ToDecimal(dataTable.Rows[j - 3][year + "年核定數"]);
                            }
                            catch
                            {
                                var test = dataTable.Rows[j - 3][year + "年核定數"];
                                value = 0;
                            }

                            a1.PutValue(value);
                            a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                            a1.SetStyle(a1_Style);

                            if (k < allYears.Length - 1)
                            {
                                k++;
                            }
                        }
                    }
                    else if (title == "當年度撥款數")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["當年度撥款數"]);
                        a1.PutValue(value);
                        a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    else if (title == "累積撥款數")
                    {
                        CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                        decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["累積撥款數"]);
                        a1.PutValue(value);
                        a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);
                    }
                    else
                    {
                        a1.PutValue(dataTable.Rows[j - 3][title].ToString());
                    }
                }

                j++;
            }

            int bottom_count = rowcount + 1 + 2;

            k = 0;

            for (int t = 0; t < COLUMN.Length; t++)
            {
                var ABC = Report_Source.NumberToColumnLetter(t + 2 + k);

                ws1.Cells["A" + bottom_count.ToString()].PutValue("合計");

                var b = ws1.Cells[ABC + bottom_count];

                Aspose.Cells.Style b_Style = b.GetStyle();

                decimal sum = 0;

                if (COLUMN[t] == "總工程費(元)")
                {
                    foreach (DataRow row in dataTable.Rows)
                    {
                        if (row["總工程費(元)"] != DBNull.Value)
                        {
                            sum += Source.Tools.ToDecimal(row["總工程費(元)"]);
                        }
                    }

                    b.Value = sum;
                    b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                    b.SetStyle(b_Style);
                }
                else if (COLUMN[t] == "本署補助經費(中央款)(元)")
                {
                    /// 顯示已處理過的總和
                    b.Value = dict["中央款總和"];

                    b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                    b.SetStyle(b_Style);
                }
                else if (COLUMN[t] == "當年度核定數")
                {
                    foreach (var year in allYears)
                    {
                        ABC = Report_Source.NumberToColumnLetter(t + 2 + k);
                        b = ws1.Cells[ABC + bottom_count];
                        b_Style = b.GetStyle();
                        sum = 0;

                        /// 顯示已處理過的總和
                        b.Value = dict[year + "年核定數總和"];
                        b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        b.SetStyle(b_Style);

                        if (k < allYears.Length - 1)
                        {
                            k++;
                        }
                    }
                }
                else if (COLUMN[t] == "當年度撥款數")
                {
                    foreach (DataRow row in dataTable.Rows)
                    {
                        if (row["當年度撥款數"] != DBNull.Value)
                        {
                            sum += Source.Tools.ToDecimal(row["當年度撥款數"]);
                        }
                    }

                    b.Value = sum;
                    b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                    b.SetStyle(b_Style);
                }
                else if (COLUMN[t] == "累積撥款數")
                {
                    /// 每筆工程編號只需取得最新一筆累積撥款數
                    b.Value = dict["累積撥款數總和"];
                    b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                    b.SetStyle(b_Style);
                }
                else
                {

                }
            }

            var mergedRange = ws1.Cells.CreateRange("A1", Column_length + "1");
            mergedRange.Merge();
            mergedRange[0, 0].PutValue("工程細項執行表");

            Aspose.Cells.Style style = workbook.DefaultStyle;
            style.Font.Name = "Microsoft JhengHei Light";

            workbook.DefaultStyle = style;
            ws1.AutoFitColumns();
            ws1.AutoFitRows();

            for (int i = 0; i < COLUMN.Length; i++)
            {
                var Excel_column = Report_Source.NumberToColumnLetter(i + 1);
                var data_column = dataTable.Columns[i].ColumnName;

                if (data_column == "年分")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 4);
                }
                else if (data_column == "序號")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 4);
                }
                else if (data_column == "補助計畫名稱")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 27.42);
                }
                else if (data_column == "法規依據")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 7.5);
                }
                else if (data_column == "計畫期程")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 7.5);
                }
                else if (data_column == "補助對象")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 9.33);
                }
                else if (data_column == "工程編號")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 19.08);
                }
                else if (data_column == "工程名稱")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "工程地點(址)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 10.42);
                }
                else if (data_column == "行政區")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 15.33);
                }
                else if (data_column == "工程主辦機關")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 11.08);
                }
                else if (data_column == "填寫本署(自辦、代辦或補助)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 22.83);
                }
                else if (data_column == "總工程費(元)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 12.42);
                }
                else if (data_column == "分標工程費(元)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 12.17);
                }
                else if (data_column == "本署補助經費(中央款)(元)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 20.42);
                }
                else if (data_column == "工程費比例")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 5.83);
                }
                else if (data_column == "當年度核定數")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 11.5);
                }
                else if (data_column == "當年度撥款數")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 11.5);
                }
                else if (data_column == "累積撥款數")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 24.83);
                }
                else if (data_column == "工程內容") /// 20241103
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "辦理情形")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "預定發包日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "實際發包日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "預定開工日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "實際開工日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "預定開工典禮日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 14.67);
                }
                else if (data_column == "實際開工典禮日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 14.67);
                }
                else if (data_column == "預定竣工日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.67);
                }
                else if (data_column == "實際竣工日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 7.5);
                }
                else if (data_column == "預定竣工典禮日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 14.67);
                }
                else if (data_column == "實際竣工典禮日期")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 14.67);
                }
                else if (data_column == "預定進度")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 5.33);
                }
                else if (data_column == "實際進度")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 6.17);
                }
                else if (data_column == "進度差異(預定-實際)")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 16.33);
                }
                else if (data_column == "落後原因或辦理情形說明")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "落後解決對策")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "備註")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
            }

            ws1.Cells.DeleteColumn(1);

            /// 檔案輸出 => EXCEL | ODS
            string filename = "A-4.跨年度細項執行情形表" + DateTime.Now.ToString("yyyyMMddHHmmss");
            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");
            workbook.Save(savepath, SaveFormat.Xlsx);

            if (FormatFile == "ods")
            {
                var odsSavePath = Server.MapPath("~/Temp/" + filename + ".ods");

                workbook.Save(odsSavePath, SaveFormat.ODS);

                using (var stream = new FileStream(odsSavePath, FileMode.Open))
                {
                    var bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, (int)stream.Length);

                    return File(bytes, "application/vnd.oasis.opendocument.spreadsheet", filename + ".ods");
                }
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Seek(0, SeekOrigin.Begin);
                    var bytes = stream.ToArray();

                    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename + ".xlsx");
                }
            }

            #endregion 
        }

        /// <summary>
        /// B-1.修正預算作業表
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="period"></param>
        /// <param name="Subject"></param>
        /// <param name="SYR"></param>
        /// <param name="ODS"></param>
        /// <param name="Choose"></param>
        /// <returns></returns>
        public ActionResult ReportPrint6_NEW(string PID, string period, string[] Subject, string SYR, string ODS, string Choose)
        {
            #region Excel 報表樣板獲取

            string sFullFile = Server.MapPath("~/Templates/PJV0010_02.xlsx");
            Workbook workbook = new Workbook(sFullFile);

            Worksheet Sheet1 = workbook.Worksheets["Sheet1"];
            Worksheet Sheet2 = workbook.Worksheets["Sheet2"];
            Worksheet Sheet3 = workbook.Worksheets["Sheet3"];
            Worksheet Sheet4 = workbook.Worksheets["Sheet4"];
            Worksheet Sheet5 = workbook.Worksheets["Sheet5"];

            #endregion

            #region EXCEL 報表資料來源

            DataTable source = RainReport_Source.B1RERPORTDATA(PID, SYR, period, connectionString);

            B1_SummaryTable(Sheet1, source, SYR, period, PID);
            B1_PublicTable(Sheet2, source, SYR, period, PID);
            B1_BOTTable(Sheet3, source, SYR, period, PID);
            B1_SummaryDetails(Sheet4, source, SYR, period, PID);
            B1_SimplifiedDetails(Sheet5, source, SYR, period, PID);

            #endregion

            #region Excel/ODS 報表產出

            /// 儲存位置 Temp 資料夾
            string sFileName = "B-1.修正預算管考報表" + DateTime.Now.ToString("yyyyMMddHHmmss");

            using (var stream = new MemoryStream())
            {
                if (ODS == "1")
                {
                    var worksheet = workbook.Worksheets[0];
                    worksheet.AutoFitRows();

                    var saveOptions = new OdsSaveOptions();
                    workbook.Save(stream, saveOptions);
                    stream.Seek(0, SeekOrigin.Begin);
                    return File(stream.ToArray(), "application/vnd.oasis.opendocument.spreadsheet", sFileName + ".ods");
                }
                else
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Seek(0, SeekOrigin.Begin);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sFileName + ".xlsx");
                }
            }

            #endregion
        }

        /// <summary>
        /// B-1.修正預算作業表[總表]_Sheet1
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="source"></param>
        /// <param name="SYR"></param>
        /// <param name="SMN"></param>
        private void B1_SummaryTable(Worksheet worksheet1, DataTable source, string SYR, string SEQNO, string PID)
        {
            int templateStartRow = 3;   /// A4:Q5 -> row 3~4
            int blockHeight = 2;        /// 公辦 + BOT
            int spacerRows = 1;
            int summaryTempRow = worksheet1.Cells.MaxDataRow + 50; /// 暫存區(總計區塊)

            DataTable dt = source;
            worksheet1.Name = "總表";

            /// 依 CITY 分組，空白 CITY 視為 不分區
            var cityGroups = dt.AsEnumerable().GroupBy(r => r["CITY"].ToString()).OrderBy(g => g.Key).ToList();

            string ReportTitle = GetReportTitle(PID, connectionString);
            worksheet1.Cells["A1"].PutValue($"{SYR}年度 {ReportTitle} 總表");
            worksheet1.Cells.DeleteRows(templateStartRow + blockHeight, 200); /// 刪掉模板下面所有列

            int currentStartRow = templateStartRow;

            /// Pperiod(期數換算) = string -> int -> string
            int TransPeriod = Convert.ToInt32(SEQNO);
            TransPeriod = TransPeriod - 1;
            string PSEQNO = TransPeriod.ToString();

            /// 原預算
            if (SEQNO == "0")
            {
                worksheet1.Cells["C2"].PutValue($"第{SEQNO}期");
                worksheet1.Cells["H2"].PutValue($"第{Math.Max(Convert.ToInt32(SEQNO) - 1, 0)}期");

                for (int i = 0; i < cityGroups.Count; i++)
                {
                    var cityGroup = cityGroups[i];

                    worksheet1.Cells.CopyRows(worksheet1.Cells, templateStartRow, currentStartRow, blockHeight);

                    int govRow = currentStartRow;
                    int botRow = currentStartRow + 1;

                    string city = cityGroup.Key;

                    worksheet1.Cells[govRow, 0].PutValue(city);
                    worksheet1.Cells[botRow, 0].PutValue("");
                    worksheet1.Cells[govRow, 1].PutValue("公辦");
                    worksheet1.Cells[botRow, 1].PutValue("BOT");

                    /// 公辦 => 業務費 / 獎補助費 / 設備及投資
                    double gov200 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov400 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov300 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// BOT  => 業務費 / 獎補助費 / 設備及投資
                    double bot200 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot400 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot300 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// C - E：本期
                    worksheet1.Cells[govRow, 2].PutValue(gov300);   /// C
                    worksheet1.Cells[govRow, 3].PutValue(gov400);   /// D
                    worksheet1.Cells[govRow, 4].PutValue(gov200);   /// E
                    worksheet1.Cells[botRow, 2].PutValue(bot300);   /// C
                    worksheet1.Cells[botRow, 3].PutValue(bot400);   /// D
                    worksheet1.Cells[botRow, 4].PutValue(bot200);   /// E

                    /// H - J：前期
                    worksheet1.Cells[govRow, 7].PutValue(0);        /// H
                    worksheet1.Cells[govRow, 8].PutValue(0);        /// I
                    worksheet1.Cells[govRow, 9].PutValue(0);        /// J
                    worksheet1.Cells[botRow, 7].PutValue(0);        /// H
                    worksheet1.Cells[botRow, 8].PutValue(0);        /// I
                    worksheet1.Cells[botRow, 9].PutValue(0);        /// J

                    currentStartRow += (blockHeight + spacerRows);
                }

                /// 複製格式, 清除資料(僅保留格式)
                worksheet1.Cells.CopyRows(worksheet1.Cells, templateStartRow, currentStartRow, blockHeight);

                for (int r = currentStartRow; r < currentStartRow + blockHeight; r++)
                {
                    for (int c = 0; c <= 16; c++)
                    {
                        Cell cell = worksheet1.Cells[r, c];

                        if (string.IsNullOrEmpty(cell.Formula))
                        {
                            cell.Value = null;
                        }
                    }
                }

                /// 標題
                worksheet1.Cells[currentStartRow, 0].PutValue("總計");
                worksheet1.Cells[currentStartRow + 1, 0].PutValue("");
                worksheet1.Cells[currentStartRow, 1].PutValue("公辦");
                worksheet1.Cells[currentStartRow + 1, 1].PutValue("BOT");

                /// 公辦(本期)
                double sumGov300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumGov400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumGov200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                /// BOT(本期)
                double sumBot300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumBot400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumBot200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == "0").Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                worksheet1.Cells[currentStartRow, 2].PutValue(sumGov300);       /// C
                worksheet1.Cells[currentStartRow, 3].PutValue(sumGov400);       /// D
                worksheet1.Cells[currentStartRow, 4].PutValue(sumGov200);       /// E
                worksheet1.Cells[currentStartRow + 1, 2].PutValue(sumBot300);   /// C
                worksheet1.Cells[currentStartRow + 1, 3].PutValue(sumBot400);   /// D
                worksheet1.Cells[currentStartRow + 1, 4].PutValue(sumBot200);   /// E

                worksheet1.Cells[currentStartRow, 7].PutValue(0);       /// H
                worksheet1.Cells[currentStartRow, 8].PutValue(0);       /// I
                worksheet1.Cells[currentStartRow, 9].PutValue(0);       /// J
                worksheet1.Cells[currentStartRow + 1, 7].PutValue(0);   /// H
                worksheet1.Cells[currentStartRow + 1, 8].PutValue(0);   /// I
                worksheet1.Cells[currentStartRow + 1, 9].PutValue(0);   /// J
            }
            /// 含預算 > 0 期數
            else
            {
                worksheet1.Cells["C2"].PutValue($"第{Math.Max(Convert.ToInt32(SEQNO) - 1, 0)}期");
                worksheet1.Cells["H2"].PutValue($"第{SEQNO}期");

                for (int i = 0; i < cityGroups.Count; i++)
                {
                    var cityGroup = cityGroups[i];

                    worksheet1.Cells.CopyRows(worksheet1.Cells, templateStartRow, currentStartRow, blockHeight);

                    int govRow = currentStartRow;
                    int botRow = currentStartRow + 1;

                    string city = cityGroup.Key;

                    worksheet1.Cells[govRow, 0].PutValue(city);
                    worksheet1.Cells[botRow, 0].PutValue("");

                    worksheet1.Cells[govRow, 1].PutValue("公辦");
                    worksheet1.Cells[botRow, 1].PutValue("BOT");

                    /// 公辦 => 業務費 / 獎補助費 / 設備及投資
                    double gov200P = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov400P = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov300P = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// BOT  => 業務費 / 獎補助費 / 設備及投資
                    double bot200P = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot400P = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot300P = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// C - E : 前一期
                    worksheet1.Cells[govRow, 2].PutValue(gov300P); // C
                    worksheet1.Cells[govRow, 3].PutValue(gov400P); // D
                    worksheet1.Cells[govRow, 4].PutValue(gov200P); // E
                    worksheet1.Cells[botRow, 2].PutValue(bot300P); // C
                    worksheet1.Cells[botRow, 3].PutValue(bot400P); // D
                    worksheet1.Cells[botRow, 4].PutValue(bot200P); // E

                    /// 公辦 => 業務費 / 獎補助費 / 設備及投資
                    double gov200 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov400 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double gov300 = cityGroup.Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// BOT  => 業務費 / 獎補助費 / 設備及投資
                    double bot200 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot400 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                    double bot300 = cityGroup.Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                    /// H - J : 本期
                    worksheet1.Cells[govRow, 7].PutValue(gov300); // H
                    worksheet1.Cells[govRow, 8].PutValue(gov400); // I
                    worksheet1.Cells[govRow, 9].PutValue(gov200); // J
                    worksheet1.Cells[botRow, 7].PutValue(bot300); // H
                    worksheet1.Cells[botRow, 8].PutValue(bot400); // I
                    worksheet1.Cells[botRow, 9].PutValue(bot200); // J

                    currentStartRow += (blockHeight + spacerRows);
                }

                /// 複製格式, 清除資料(僅保留格式)
                worksheet1.Cells.CopyRows(worksheet1.Cells, templateStartRow, currentStartRow, blockHeight);

                for (int r = currentStartRow; r < currentStartRow + blockHeight; r++)
                {
                    for (int c = 0; c <= 16; c++)
                    {
                        Cell cell = worksheet1.Cells[r, c];

                        if (string.IsNullOrEmpty(cell.Formula))
                        {
                            cell.Value = null;
                        }
                    }
                }

                /// 標題
                worksheet1.Cells[currentStartRow, 0].PutValue("總計");
                worksheet1.Cells[currentStartRow + 1, 0].PutValue("");
                worksheet1.Cells[currentStartRow, 1].PutValue("公辦");
                worksheet1.Cells[currentStartRow + 1, 1].PutValue("BOT");

                /// 公辦(前期)
                double sumGov300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumGov400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumGov200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                /// BOT(前期)
                double sumBot300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumBot400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double sumBot200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == PSEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                /// 公辦(本期)
                double PsumGov300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double PsumGov400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double PsumGov200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "0" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                /// BOT(本期)
                double PsumBot300 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "設備及投資" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double PsumBot400 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "獎補助費" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));
                double PsumBot200 = dt.AsEnumerable().Where(r => r["BOT_FG"].ToString() == "1" && r["ACC_DESC"].ToString() == "業務費" && r["SEQ_NO"].ToString() == SEQNO).Sum(r => r["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(r["CEN_FUND"]));

                worksheet1.Cells[currentStartRow, 2].PutValue(sumGov300); // C
                worksheet1.Cells[currentStartRow, 3].PutValue(sumGov400); // D
                worksheet1.Cells[currentStartRow, 4].PutValue(sumGov200); // E
                worksheet1.Cells[currentStartRow + 1, 2].PutValue(sumBot300); // C
                worksheet1.Cells[currentStartRow + 1, 3].PutValue(sumBot400); // D
                worksheet1.Cells[currentStartRow + 1, 4].PutValue(sumBot200); // E

                worksheet1.Cells[currentStartRow, 7].PutValue(PsumGov300); // H
                worksheet1.Cells[currentStartRow, 8].PutValue(PsumGov400); // I
                worksheet1.Cells[currentStartRow, 9].PutValue(PsumGov200); // J
                worksheet1.Cells[currentStartRow + 1, 7].PutValue(PsumBot300); // H
                worksheet1.Cells[currentStartRow + 1, 8].PutValue(PsumBot400); // I
                worksheet1.Cells[currentStartRow + 1, 9].PutValue(PsumBot200); // J
            }
        }

        /// <summary>
        /// B-1.修正預算作業表[總表(公辦)]_Sheet2
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="source"></param>
        /// <param name="SYR"></param>
        /// <param name="SMN"></param>
        private void B1_PublicTable(Worksheet worksheet2, DataTable source, string SYR, string SEQNO, string PID)
        {
            worksheet2.Name = "總表(公辦)";
            DataTable Table = source;

            int templateStartRow = 4;   /// A5 : Index 4 (0 - based)
            int blockHeight = 4;        /// A5 ~ A8 => 共 4 列 
            int currentRow = templateStartRow;

            var cityGroups = Table.AsEnumerable().GroupBy(r => r["CITY"].ToString()).OrderBy(g => g.Key).ToList();

            /// Pperiod(前期數換算) : string => int => string
            int TransPeriod = Convert.ToInt32(SEQNO);
            TransPeriod = TransPeriod - 1;
            string PSEQNO = TransPeriod.ToString();

            /// 原預算
            if (SEQNO == "0")
            {
                string ReportTitle = GetReportTitle(PID, connectionString);
                worksheet2.Cells["B1"].PutValue($"{SYR}年度 {ReportTitle} 總表(公辦)");
                worksheet2.Cells["D2"].PutValue($"第{SEQNO}期");
                worksheet2.Cells["H2"].PutValue($"第{Math.Max(Convert.ToInt32(SEQNO) - 1, 0)}期");

                for (int i = 0; i < cityGroups.Count; i++)
                {
                    var cityGroup = cityGroups[i];
                    string city = cityGroup.Key;

                    worksheet2.Cells.CopyRows(worksheet2.Cells, templateStartRow, currentRow, blockHeight);

                    int row1 = currentRow;      /// A5
                    int row2 = currentRow + 1;  /// A6
                    int row3 = currentRow + 2;  /// A7
                    int row4 = currentRow + 3;  /// A8

                    /// 計算 [規劃類] "署辦" 件數 / "署辦" 金額
                    /// 計算 [工程類] "署辦" 件數 / "署辦" 金額
                    /// 計算 [規劃類] "地方辦" 件數 / "地方辦" 金額
                    /// 計算 [工程類] "地方辦" 件數 / "地方辦" 金額
                    /// 計算 1月 ~ 12 月 各月份總合
                    int AdminPlanCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() != "C").Count();
                    double AdminPlanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() != "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));
                    int AdminEngerCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() == "C").Count();
                    double AdminEngerFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() == "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));

                    int LocalPlanCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() != "C").Count();
                    double LocalPlanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() != "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));
                    int LocalEngerCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() == "C").Count();
                    double LocalEngerFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() == "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));

                    double JanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_01"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_01"]));  /// 1月
                    double FebFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_02"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_02"]));  /// 2月
                    double MarFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_03"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_03"]));  /// 3月
                    double AprFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_04"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_04"]));  /// 4月
                    double MayFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_05"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_05"]));  /// 5月
                    double JunFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_06"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_06"]));  /// 6月
                    double JulFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_07"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_07"]));  /// 7月
                    double AugFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_08"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_08"]));  /// 8月
                    double SepFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_09"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_09"]));  /// 9月
                    double OctFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_10"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_10"]));  /// 10月
                    double NovFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_11"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_11"]));  /// 11月
                    double DecFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_12"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_12"]));  /// 12月

                    #region ======================================== 填資料位置 ========================================

                    worksheet2.Cells[row1, 0].PutValue(city);

                    worksheet2.Cells[row1, 3].PutValue(AdminPlanCount);
                    worksheet2.Cells[row2, 3].PutValue(AdminPlanFund);
                    worksheet2.Cells[row3, 3].PutValue(LocalPlanCount);
                    worksheet2.Cells[row4, 3].PutValue(LocalPlanFund);
                    worksheet2.Cells[row1, 4].PutValue(AdminEngerCount);
                    worksheet2.Cells[row2, 4].PutValue(AdminEngerFund);
                    worksheet2.Cells[row3, 4].PutValue(LocalEngerCount);
                    worksheet2.Cells[row4, 4].PutValue(LocalEngerFund);

                    worksheet2.Cells[row1, 7].PutValue(0);
                    worksheet2.Cells[row2, 7].PutValue(0);
                    worksheet2.Cells[row3, 7].PutValue(0);
                    worksheet2.Cells[row4, 7].PutValue(0);
                    worksheet2.Cells[row1, 8].PutValue(0);
                    worksheet2.Cells[row2, 8].PutValue(0);
                    worksheet2.Cells[row3, 8].PutValue(0);
                    worksheet2.Cells[row4, 8].PutValue(0);

                    worksheet2.Cells[row1, 16].PutValue(JanFund);
                    worksheet2.Cells[row1, 17].PutValue(FebFund);
                    worksheet2.Cells[row1, 18].PutValue(MarFund);
                    worksheet2.Cells[row1, 19].PutValue(AprFund);
                    worksheet2.Cells[row1, 20].PutValue(MayFund);
                    worksheet2.Cells[row1, 21].PutValue(JunFund);
                    worksheet2.Cells[row3, 16].PutValue(JulFund);
                    worksheet2.Cells[row3, 17].PutValue(AugFund);
                    worksheet2.Cells[row3, 18].PutValue(SepFund);
                    worksheet2.Cells[row3, 19].PutValue(OctFund);
                    worksheet2.Cells[row3, 20].PutValue(NovFund);
                    worksheet2.Cells[row3, 21].PutValue(DecFund);


                    #endregion ========================================================================================


                    currentRow += blockHeight + 1;
                }
            }
            /// 含預算 > 0 期數
            else
            {
                string ReportTitle = GetReportTitle(PID, connectionString);
                worksheet2.Cells["B1"].PutValue($"{SYR}年度 {ReportTitle} 總表(公辦)");
                worksheet2.Cells["D2"].PutValue($"第{Math.Max(Convert.ToInt32(SEQNO) - 1, 0)}期");
                worksheet2.Cells["H2"].PutValue($"第{SEQNO}期");

                for (int i = 0; i < cityGroups.Count; i++)
                {
                    var cityGroup = cityGroups[i];
                    string city = cityGroup.Key;

                    worksheet2.Cells.CopyRows(worksheet2.Cells, templateStartRow, currentRow, blockHeight);

                    int row1 = currentRow;
                    int row2 = currentRow + 1;
                    int row3 = currentRow + 2;  /// A7
                    int row4 = currentRow + 3;  /// A8

                    /// 計算 [規劃類] "署辦" 件數 / "署辦" 金額
                    /// 計算 [工程類] "署辦" 件數 / "署辦" 金額
                    /// 計算 [規劃類] "地方辦" 件數 / "地方辦" 金額
                    /// 計算 [工程類] "地方辦" 件數 / "地方辦" 金額
                    /// 計算 1月 ~ 12 月 各月份總合
                    int PAdminPlanCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == PSEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() != "C").Count();
                    double PAdminPlanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == PSEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() != "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));
                    int PAdminEngerCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == PSEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() == "C").Count();
                    double PAdminEngerFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == PSEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() == "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));

                    int PLocalPlanCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == PSEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() != "C").Count();
                    double PLocalPlanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == PSEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() != "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));
                    int PLocalEngerCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == PSEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() == "C").Count();
                    double PLocalEngerFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == PSEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() == "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));

                    int AdminPlanCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() != "C").Count();
                    double AdminPlanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() != "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));
                    int AdminEngerCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() == "C").Count();
                    double AdminEngerFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && (s["IPJ_DESC"].ToString() == "署辦" || s["IPJ_DESC"].ToString() == "署委辦") && s["CATY_CODE"].ToString() == "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));

                    int LocalPlanCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() != "C").Count();
                    double LocalPlanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() != "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));
                    int LocalEngerCount = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() == "C").Count();
                    double LocalEngerFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO && s["IPJ_DESC"].ToString() == "地方辦" && s["CATY_CODE"].ToString() == "C").Sum(s => s["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(s["CEN_FUND"]));

                    double JanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_01"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_01"]));  /// 1月
                    double FebFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_02"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_02"]));  /// 2月
                    double MarFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_03"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_03"]));  /// 3月
                    double AprFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_04"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_04"]));  /// 4月
                    double MayFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_05"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_05"]));  /// 5月
                    double JunFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_06"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_06"]));  /// 6月
                    double JulFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_07"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_07"]));  /// 7月
                    double AugFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_08"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_08"]));  /// 8月
                    double SepFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_09"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_09"]));  /// 9月
                    double OctFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_10"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_10"]));  /// 10月
                    double NovFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_11"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_11"]));  /// 11月
                    double DecFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_12"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_12"]));  /// 12月

                    #region ======================================== 填資料位置 ========================================

                    worksheet2.Cells[row1, 0].PutValue(city);

                    worksheet2.Cells[row1, 3].PutValue(PAdminPlanCount);
                    worksheet2.Cells[row2, 3].PutValue(PAdminPlanFund);
                    worksheet2.Cells[row3, 3].PutValue(PLocalPlanCount);
                    worksheet2.Cells[row4, 3].PutValue(PLocalPlanFund);
                    worksheet2.Cells[row1, 4].PutValue(PAdminEngerCount);
                    worksheet2.Cells[row2, 4].PutValue(PAdminEngerFund);
                    worksheet2.Cells[row3, 4].PutValue(PLocalEngerCount);
                    worksheet2.Cells[row4, 4].PutValue(PLocalEngerFund);

                    worksheet2.Cells[row1, 7].PutValue(AdminPlanCount);
                    worksheet2.Cells[row2, 7].PutValue(AdminPlanFund);
                    worksheet2.Cells[row3, 7].PutValue(LocalPlanCount);
                    worksheet2.Cells[row4, 7].PutValue(LocalPlanFund);
                    worksheet2.Cells[row1, 8].PutValue(AdminEngerCount);
                    worksheet2.Cells[row2, 8].PutValue(AdminEngerFund);
                    worksheet2.Cells[row3, 8].PutValue(LocalEngerCount);
                    worksheet2.Cells[row4, 8].PutValue(LocalEngerFund);

                    worksheet2.Cells[row1, 16].PutValue(JanFund);
                    worksheet2.Cells[row1, 17].PutValue(FebFund);
                    worksheet2.Cells[row1, 18].PutValue(MarFund);
                    worksheet2.Cells[row1, 19].PutValue(AprFund);
                    worksheet2.Cells[row1, 20].PutValue(MayFund);
                    worksheet2.Cells[row1, 21].PutValue(JunFund);
                    worksheet2.Cells[row3, 16].PutValue(JulFund);
                    worksheet2.Cells[row3, 17].PutValue(AugFund);
                    worksheet2.Cells[row3, 18].PutValue(SepFund);
                    worksheet2.Cells[row3, 19].PutValue(OctFund);
                    worksheet2.Cells[row3, 20].PutValue(NovFund);
                    worksheet2.Cells[row3, 21].PutValue(DecFund);


                    #endregion ========================================================================================

                    currentRow += blockHeight + 1;
                }
            }
        }

        /// <summary>
        /// B-1.修正預算作業表[總表(BOT)]_Sheet3
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="source"></param>
        /// <param name="SYR"></param>
        /// <param name="SMN"></param>
        private void B1_BOTTable(Worksheet worksheet3, DataTable source, string SYR, string SEQNO, string PID)
        {
            worksheet3.Name = "總表(BOT)";
            DataTable Table = source;

            int templateStartRow = 4;   /// A5 = index 4（0-based）
            int blockHeight = 4;        /// A5 ~ A8 => 總共 4 列
            int currentRow = templateStartRow;

            var cityGroups = Table.AsEnumerable().GroupBy(r => r["CITY"].ToString()).OrderBy(g => g.Key).ToList();

            /// Pperiod(期數換算) = string -> int -> string
            int TransPeriod = Convert.ToInt32(SEQNO);
            TransPeriod = TransPeriod - 1;
            string PSEQNO = TransPeriod.ToString();

            /// 原預算
            if (SEQNO == "0")
            {
                string ReportTitle = GetReportTitle(PID, connectionString);
                worksheet3.Cells["B1"].PutValue($"{SYR}年度 {ReportTitle} 總表(BOT)");
                worksheet3.Cells["D2"].PutValue($"第{SEQNO}期");
                worksheet3.Cells["H2"].PutValue($"第{Math.Max(Convert.ToInt32(SEQNO) - 1, 0)}期");

                for (int i = 0; i < cityGroups.Count; i++)
                {
                    var cityGroup = cityGroups[i];
                    string city = cityGroup.Key;

                    worksheet3.Cells.CopyRows(worksheet3.Cells, templateStartRow, currentRow, blockHeight);

                    int row1 = currentRow;      /// A5
                    int row2 = currentRow + 1;  /// A6
                    int row3 = currentRow + 2;  /// A7
                    int row4 = currentRow + 3;  /// A8

                    /// 計算 [用地取得] "署辦" 件數 / "署辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int OLolCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "01").Count();
                    double OLolFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int LLolCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "02").Count();
                    double LLolFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 [先期規劃] "署辦" 件數 / "署辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int OPlnCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "01").Count();
                    double OPlnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int LPlnCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "02").Count();
                    double LPlnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 [專案管理] "地方辦" 件數 / "地方辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int OProCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "01").Count();
                    double OProFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int LProCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "02").Count();
                    double LProFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 [政府應辦事項] "地方辦" 件數 / "地方辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int OGovCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "01").Count();
                    double OGovnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int LGovCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "02").Count();
                    double LGovnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 1月 ~ 12 月 各月份總合
                    double JanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_01"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_01"]));  /// 1月
                    double FebFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_02"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_02"]));  /// 2月
                    double MarFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_03"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_03"]));  /// 3月
                    double AprFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_04"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_04"]));  /// 4月
                    double MayFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_05"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_05"]));  /// 5月
                    double JunFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_06"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_06"]));  /// 6月
                    double JulFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_07"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_07"]));  /// 7月
                    double AugFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_08"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_08"]));  /// 8月
                    double SepFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_09"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_09"]));  /// 9月
                    double OctFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_10"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_10"]));  /// 10月
                    double NovFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_11"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_11"]));  /// 11月
                    double DecFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_12"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_12"]));  /// 12月

                    #region ======================================== 填資料位置 ========================================

                    worksheet3.Cells[row1, 0].PutValue(city);

                    worksheet3.Cells[row1, 3].PutValue(OLolCount);
                    worksheet3.Cells[row2, 3].PutValue(OLolFund);
                    worksheet3.Cells[row3, 3].PutValue(LLolCount);
                    worksheet3.Cells[row4, 3].PutValue(LLolFund);
                    worksheet3.Cells[row1, 4].PutValue(OPlnCount);
                    worksheet3.Cells[row2, 4].PutValue(OPlnFund);
                    worksheet3.Cells[row3, 4].PutValue(LPlnCount);
                    worksheet3.Cells[row4, 4].PutValue(LPlnFund);
                    worksheet3.Cells[row1, 5].PutValue(OProCount);
                    worksheet3.Cells[row2, 5].PutValue(OProFund);
                    worksheet3.Cells[row3, 5].PutValue(LProCount);
                    worksheet3.Cells[row4, 5].PutValue(LProFund);
                    worksheet3.Cells[row1, 6].PutValue(OGovCount);
                    worksheet3.Cells[row2, 6].PutValue(OGovnFund);
                    worksheet3.Cells[row3, 6].PutValue(LGovCount);
                    worksheet3.Cells[row4, 6].PutValue(LGovnFund);

                    worksheet3.Cells[row1, 9].PutValue(0);
                    worksheet3.Cells[row2, 9].PutValue(0);
                    worksheet3.Cells[row3, 9].PutValue(0);
                    worksheet3.Cells[row4, 9].PutValue(0);
                    worksheet3.Cells[row1, 10].PutValue(0);
                    worksheet3.Cells[row2, 10].PutValue(0);
                    worksheet3.Cells[row3, 10].PutValue(0);
                    worksheet3.Cells[row4, 10].PutValue(0);
                    worksheet3.Cells[row1, 11].PutValue(0);
                    worksheet3.Cells[row2, 11].PutValue(0);
                    worksheet3.Cells[row3, 11].PutValue(0);
                    worksheet3.Cells[row4, 11].PutValue(0);
                    worksheet3.Cells[row1, 12].PutValue(0);
                    worksheet3.Cells[row2, 12].PutValue(0);
                    worksheet3.Cells[row3, 12].PutValue(0);
                    worksheet3.Cells[row4, 12].PutValue(0);

                    worksheet3.Cells[row1, 20].PutValue(JanFund);
                    worksheet3.Cells[row1, 21].PutValue(FebFund);
                    worksheet3.Cells[row1, 22].PutValue(MarFund);
                    worksheet3.Cells[row1, 23].PutValue(AprFund);
                    worksheet3.Cells[row1, 24].PutValue(MayFund);
                    worksheet3.Cells[row1, 25].PutValue(JunFund);
                    worksheet3.Cells[row3, 20].PutValue(JulFund);
                    worksheet3.Cells[row3, 21].PutValue(AugFund);
                    worksheet3.Cells[row3, 22].PutValue(SepFund);
                    worksheet3.Cells[row3, 23].PutValue(OctFund);
                    worksheet3.Cells[row3, 24].PutValue(NovFund);
                    worksheet3.Cells[row3, 25].PutValue(DecFund);


                    #endregion ========================================================================================


                    currentRow += blockHeight + 1;
                }
            }
            /// 含預算 > 0 期數
            else
            {
                string ReportTitle = GetReportTitle(PID, connectionString);
                worksheet3.Cells["B1"].PutValue($"{SYR}年度 {ReportTitle} 總表(BOT)");
                worksheet3.Cells["D2"].PutValue($"第{Math.Max(Convert.ToInt32(SEQNO) - 1, 0)}期");
                worksheet3.Cells["J2"].PutValue($"第{SEQNO}期");

                for (int i = 0; i < cityGroups.Count; i++)
                {
                    var cityGroup = cityGroups[i];
                    string city = cityGroup.Key;

                    worksheet3.Cells.CopyRows(worksheet3.Cells, templateStartRow, currentRow, blockHeight);

                    int row1 = currentRow;
                    int row2 = currentRow + 1;
                    int row3 = currentRow + 2;  /// A7
                    int row4 = currentRow + 3;  /// A8

                    #region ======================================== 舊一期位置 ========================================
                    /// 計算 [用地取得] "署辦" 件數 / "署辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int OLolCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "01").Count();
                    double OLolFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int LLolCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "02").Count();
                    double LLolFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 [先期規劃] "署辦" 件數 / "署辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int OPlnCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "01").Count();
                    double OPlnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int LPlnCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "02").Count();
                    double LPlnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 [專案管理] "地方辦" 件數 / "地方辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int OProCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "01").Count();
                    double OProFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int LProCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "02").Count();
                    double LProFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 [政府應辦事項] "地方辦" 件數 / "地方辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int OGovCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "01").Count();
                    double OGovnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int LGovCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "02").Count();
                    double LGovnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == PSEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));
                    #endregion ========================================================================================

                    #region ======================================== 新一期位置 ========================================
                    /// 計算 [用地取得] "署辦" 件數 / "署辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int SOLolCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "01").Count();
                    double SOLolFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int SLLolCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "02").Count();
                    double SLLolFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "L" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 [先期規劃] "署辦" 件數 / "署辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int SOPlnCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "01").Count();
                    double SOPlnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int SLPlnCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "02").Count();
                    double SLPlnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "J" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 [專案管理] "地方辦" 件數 / "地方辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int SOProCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "01").Count();
                    double SOProFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int SLProCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "02").Count();
                    double SLProFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() == "M" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    /// 計算 [政府應辦事項] "地方辦" 件數 / "地方辦" 金額 / "地方辦" 件數 / "地方辦" 金額
                    int SOGovCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "01").Count();
                    double SOGovnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "01").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));

                    int SLGovCount = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "02").Count();
                    double SLGovnFund = cityGroup.Where(T => T["SEQ_NO"].ToString() == SEQNO && T["BOT_FG"].ToString() == "1" && T["CATY_CODE"].ToString() != "L" && T["CATY_CODE"].ToString() != "J" && T["CATY_CODE"].ToString() != "M" && T["IPJ_CODE"].ToString() == "02").Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));
                    #endregion ========================================================================================

                    /// 計算 1月 ~ 12 月 各月份總合
                    double JanFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_01"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_01"]));  /// 1月
                    double FebFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_02"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_02"]));  /// 2月
                    double MarFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_03"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_03"]));  /// 3月
                    double AprFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_04"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_04"]));  /// 4月
                    double MayFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_05"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_05"]));  /// 5月
                    double JunFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_06"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_06"]));  /// 6月
                    double JulFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_07"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_07"]));  /// 7月
                    double AugFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_08"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_08"]));  /// 8月
                    double SepFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_09"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_09"]));  /// 9月
                    double OctFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_10"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_10"]));  /// 10月
                    double NovFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_11"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_11"]));  /// 11月
                    double DecFund = cityGroup.Where(s => s["SEQ_NO"].ToString() == SEQNO).Sum(s => s["M_12"] == DBNull.Value ? 0 : Convert.ToDouble(s["M_12"]));  /// 12月

                    #region ======================================== 填資料位置 ========================================

                    worksheet3.Cells[row1, 0].PutValue(city);

                    worksheet3.Cells[row1, 3].PutValue(OLolCount);
                    worksheet3.Cells[row2, 3].PutValue(OLolFund);
                    worksheet3.Cells[row3, 3].PutValue(LLolCount);
                    worksheet3.Cells[row4, 3].PutValue(LLolFund);
                    worksheet3.Cells[row1, 4].PutValue(OPlnCount);
                    worksheet3.Cells[row2, 4].PutValue(OPlnFund);
                    worksheet3.Cells[row3, 4].PutValue(LPlnCount);
                    worksheet3.Cells[row4, 4].PutValue(LPlnFund);
                    worksheet3.Cells[row1, 5].PutValue(OProCount);
                    worksheet3.Cells[row2, 5].PutValue(OProFund);
                    worksheet3.Cells[row3, 5].PutValue(LProCount);
                    worksheet3.Cells[row4, 5].PutValue(LProFund);
                    worksheet3.Cells[row1, 6].PutValue(OGovCount);
                    worksheet3.Cells[row2, 6].PutValue(OGovnFund);
                    worksheet3.Cells[row3, 6].PutValue(LGovCount);
                    worksheet3.Cells[row4, 6].PutValue(LGovnFund);

                    worksheet3.Cells[row1, 9].PutValue(SOLolCount);
                    worksheet3.Cells[row2, 9].PutValue(SOLolFund);
                    worksheet3.Cells[row3, 9].PutValue(SLLolCount);
                    worksheet3.Cells[row4, 9].PutValue(SLLolFund);
                    worksheet3.Cells[row1, 10].PutValue(SOPlnCount);
                    worksheet3.Cells[row2, 10].PutValue(SOPlnFund);
                    worksheet3.Cells[row3, 10].PutValue(SLPlnCount);
                    worksheet3.Cells[row4, 10].PutValue(SLPlnFund);
                    worksheet3.Cells[row1, 11].PutValue(SOProCount);
                    worksheet3.Cells[row2, 11].PutValue(SOProFund);
                    worksheet3.Cells[row3, 11].PutValue(SLProCount);
                    worksheet3.Cells[row4, 11].PutValue(SLProFund);
                    worksheet3.Cells[row1, 12].PutValue(SOGovCount);
                    worksheet3.Cells[row2, 12].PutValue(SOGovnFund);
                    worksheet3.Cells[row3, 12].PutValue(SLGovCount);
                    worksheet3.Cells[row4, 12].PutValue(SLGovnFund);

                    worksheet3.Cells[row1, 20].PutValue(JanFund);
                    worksheet3.Cells[row1, 21].PutValue(FebFund);
                    worksheet3.Cells[row1, 22].PutValue(MarFund);
                    worksheet3.Cells[row1, 23].PutValue(AprFund);
                    worksheet3.Cells[row1, 24].PutValue(MayFund);
                    worksheet3.Cells[row1, 25].PutValue(JunFund);
                    worksheet3.Cells[row3, 20].PutValue(JulFund);
                    worksheet3.Cells[row3, 21].PutValue(AugFund);
                    worksheet3.Cells[row3, 22].PutValue(SepFund);
                    worksheet3.Cells[row3, 23].PutValue(OctFund);
                    worksheet3.Cells[row3, 24].PutValue(NovFund);
                    worksheet3.Cells[row3, 25].PutValue(DecFund);


                    #endregion ========================================================================================

                    currentRow += blockHeight + 1;
                }
            }
        }

        /// <summary>
        /// B-1.修正預算作業表[彙整明細]_Sheet4
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="source"></param>
        /// <param name="SYR"></param>
        /// <param name="SMN"></param>
        private void B1_SummaryDetails(Worksheet worksheet4, DataTable source, string SYR, string SEQNO, string PID)
        {
            worksheet4.Name = "彙整明細";
            DataTable table = source;

            int count = 1;
            int templateStartRow = 1;   // A2 => index = 1
            int blockHeight = 25;       // 每個完整 block 25 列
            int itemPerBlock = 11;      // 每個 block 可放 11 個項目(每項目佔2列)
            int currentRow = templateStartRow;

            /// 分類前置作業 : CITY 分組
            var cityGroups = table.AsEnumerable().GroupBy(r => r["CITY"].ToString()).OrderBy(g => g.Key).ToList();

            /// 取前一期期數
            string PSEQNO = Math.Max(Convert.ToInt32(SEQNO) - 1, 0).ToString();

            /// 原預算
            if (SEQNO == "0")
            {
                foreach (var cityGroup in cityGroups)
                {
                    string city = cityGroup.Key;
                    List<DataRow> cityRows = cityGroup.ToList();

                    /// 先拆成 規劃類 / 工程類
                    List<DataRow> planningRows = cityRows.Where(r => r["CATY_CODE"].ToString() != "C").ToList();
                    List<DataRow> engineeringRows = cityRows.Where(r => r["CATY_CODE"].ToString() == "C").ToList();

                    /// 建立「實際輸出項目清單」
                    /// Type:
                    /// D = 明細
                    /// P = 規劃類小計
                    /// E = 工程類小計
                    /// T = 總計
                    var renderItems = new List<object>();

                    double N_CenFund = planningRows.Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));
                    double N_LocFund = planningRows.Sum(T => T["LOC_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["LOC_FUND"]));

                    /// 規劃類明細
                    foreach (DataRow row in planningRows)
                    {
                        renderItems.Add(new
                        {
                            Type = "D",
                            Data = row,
                            Label = "",
                            ItemCount = 0,
                            CenFund = N_CenFund,
                            LocFund = N_LocFund,
                            SumFund = N_CenFund + N_LocFund
                        });
                    }


                    /// 規劃類小計
                    if (planningRows.Count > 0)
                    {
                        renderItems.Add(new
                        {
                            Type = "P",
                            Data = (DataRow)null,
                            Label = "規劃類小計",
                            ItemCount = planningRows.Count,
                            CenFund = N_CenFund,
                            LocFund = N_LocFund,
                            SumFund = N_CenFund + N_LocFund
                        });
                    }

                    double C_CenFund = engineeringRows.Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));
                    double C_LocFund = engineeringRows.Sum(T => T["LOC_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["LOC_FUND"]));

                    /// 工程類明細
                    foreach (DataRow row in engineeringRows)
                    {
                        renderItems.Add(new
                        {
                            Type = "D",
                            Data = row,
                            Label = "",
                            ItemCount = 0,
                            CenFund = C_CenFund,
                            LocFund = C_LocFund,
                            SumFund = C_CenFund + C_LocFund
                        });
                    }

                    /// 工程類小計
                    if (engineeringRows.Count > 0)
                    {
                        renderItems.Add(new
                        {
                            Type = "E",
                            Data = (DataRow)null,
                            Label = "工程類小計",
                            ItemCount = engineeringRows.Count,
                            CenFund = C_CenFund,
                            LocFund = C_LocFund,
                            SumFund = C_CenFund + C_LocFund
                        });
                    }

                    /// 總計
                    renderItems.Add(new
                    {
                        Type = "T",
                        Data = (DataRow)null,
                        Label = "總計",
                        ItemCount = cityRows.Count,
                        CenFund = N_CenFund + C_CenFund,
                        LocFund = N_LocFund + C_LocFund,
                        SumFund = N_CenFund + C_CenFund + N_LocFund + C_LocFund

                    });

                    /// 依照輸出項目數量計算 block 數
                    int blockCount = (int)Math.Ceiling((double)renderItems.Count / itemPerBlock);

                    for (int blockIndex = 0; blockIndex < blockCount; blockIndex++)
                    {
                        /// 第一個 block 直接用原模板 ; 後續 block 複製
                        if (currentRow != templateStartRow)
                        {
                            worksheet4.Cells.CopyRows(worksheet4.Cells, templateStartRow, currentRow, blockHeight);
                        }

                        /// Block Layout
                        int titleRow = currentRow;           /// block 第1列
                        int periodRow = currentRow + 1;      /// block 第2列
                        int headerRow = currentRow + 2;      /// block 第3列
                        int dataStartRow = currentRow + 3;   /// block 第4列
                        int dataEndRow = currentRow + 24;    /// block 第25列

                        /// 清除資料區內容，保留格式
                        for (int r = dataStartRow; r <= dataEndRow; r++)
                        {
                            for (int c = 0; c <= 21; c++)   /// A ~ P，20260408 版本修正 15 => 21
                            {
                                worksheet4.Cells[r, c].PutValue(string.Empty);
                            }
                        }

                        /// 標題列
                        string ReportTitle = GetReportTitle(PID, connectionString);
                        worksheet4.Cells["A1"].PutValue($"{SYR}年度 {ReportTitle} 彙整明細");
                        worksheet4.Cells[titleRow, 0].PutValue($"{city}調整預算表");
                        worksheet4.Cells[periodRow, 10].PutValue($"暫無期別");
                        worksheet4.Cells[periodRow, 7].PutValue(SEQNO == "0" ? "原預算" : $"第{PSEQNO}期");

                        /// 取本 block 的資料
                        var blockItems = renderItems.Skip(blockIndex * itemPerBlock).Take(itemPerBlock).ToList();

                        for (int i = 0; i < blockItems.Count; i++)
                        {
                            int targetRow = dataStartRow + (i * 2);

                            dynamic item = blockItems[i];

                            /// 明細列
                            if (item.Type == "D")
                            {
                                DataRow data = item.Data;

                                worksheet4.Cells[targetRow, 0].PutValue(count++);                       // A 序號
                                worksheet4.Cells[targetRow, 1].PutValue(data["SIP_TD_CODE"]);           // B 工程編號
                                worksheet4.Cells[targetRow, 2].PutValue(data["FD_CODE"].ToString());    // C 經費編號
                                worksheet4.Cells[targetRow, 3].PutValue(data["TD_NAME"].ToString());    // D 工程名稱
                                worksheet4.Cells[targetRow, 4].PutValue(data["CATY_CODE"].ToString() == "C" ? "工程類" : "規劃類");
                                worksheet4.Cells[targetRow, 5].PutValue(data["IPJ_DESC"].ToString());
                                worksheet4.Cells[targetRow, 6].PutValue(data["CEN_SP_RATIO"] == DBNull.Value ? 0 : Convert.ToDouble(data["CEN_SP_RATIO"]));
                                worksheet4.Cells[targetRow, 7].PutValue(Convert.ToDouble(data["CEN_FUND"]));
                                worksheet4.Cells[targetRow, 8].PutValue(Convert.ToDouble(data["LOC_FUND"]));
                                worksheet4.Cells[targetRow, 9].PutValue(Convert.ToDouble(data["CEN_FUND"]) + Convert.ToDouble(data["LOC_FUND"]));
                                worksheet4.Cells[targetRow, 10].PutValue(0);
                                worksheet4.Cells[targetRow, 11].PutValue(0);
                                worksheet4.Cells[targetRow, 12].PutValue(0);
                                worksheet4.Cells[targetRow, 13].PutValue(0);
                                worksheet4.Cells[targetRow, 14].PutValue(0);
                                worksheet4.Cells[targetRow, 15].PutValue(data["EXECT_DESC"].ToString());
                                worksheet4.Cells[targetRow, 16].PutValue(data["M_01"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_01"]));
                                worksheet4.Cells[targetRow, 17].PutValue(data["M_02"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_02"]));
                                worksheet4.Cells[targetRow, 18].PutValue(data["M_03"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_03"]));
                                worksheet4.Cells[targetRow, 19].PutValue(data["M_04"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_04"]));
                                worksheet4.Cells[targetRow, 20].PutValue(data["M_05"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_05"]));
                                worksheet4.Cells[targetRow, 21].PutValue(data["M_06"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_06"]));
                                worksheet4.Cells[targetRow + 1, 16].PutValue(data["M_07"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_07"]));
                                worksheet4.Cells[targetRow + 1, 17].PutValue(data["M_08"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_08"]));
                                worksheet4.Cells[targetRow + 1, 18].PutValue(data["M_09"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_09"]));
                                worksheet4.Cells[targetRow + 1, 19].PutValue(data["M_10"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_10"]));
                                worksheet4.Cells[targetRow + 1, 20].PutValue(data["M_11"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_11"]));
                                worksheet4.Cells[targetRow + 1, 21].PutValue(data["M_12"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_12"]));
                            }
                            else
                            {
                                /// 小計 / 總計列
                                worksheet4.Cells[targetRow, 3].PutValue(item.Label);
                                worksheet4.Cells[targetRow, 7].PutValue(item.CenFund);
                                worksheet4.Cells[targetRow, 8].PutValue(item.LocFund);
                                worksheet4.Cells[targetRow, 9].PutValue(item.SumFund);
                                worksheet4.Cells[targetRow, 10].PutValue(0);
                                worksheet4.Cells[targetRow, 11].PutValue(0);
                                worksheet4.Cells[targetRow, 12].PutValue(0);
                            }
                        }

                        currentRow += blockHeight + 1;
                    }
                }
            }
            /// 調整預算後
            else
            {
                foreach (var cityGroup in cityGroups)
                {
                    string city = cityGroup.Key;
                    List<DataRow> cityRows = cityGroup.ToList();

                    /// 先拆成 規劃類 / 工程類
                    List<DataRow> planningRows = cityRows.Where(r => r["CATY_CODE"].ToString() != "C").ToList();
                    List<DataRow> engineeringRows = cityRows.Where(r => r["CATY_CODE"].ToString() == "C").ToList();

                    /// 建立「實際輸出項目清單」
                    /// Type:
                    /// D = 明細
                    /// P = 規劃類小計
                    /// E = 工程類小計
                    /// T = 總計
                    var renderItems = new List<object>();

                    /// 當前期數
                    double SEQ_N_CenFund = planningRows.Where(T => T["SEQ_NO"].ToString() == SEQNO).Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));
                    double SEQ_N_LocFund = planningRows.Where(T => T["SEQ_NO"].ToString() == SEQNO).Sum(T => T["LOC_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["LOC_FUND"]));

                    /// 前一期數
                    double PSEQ_N_CenFund = planningRows.Where(T => T["SEQ_NO"].ToString() == PSEQNO).Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));
                    double PSEQ_N_LocFund = planningRows.Where(T => T["SEQ_NO"].ToString() == PSEQNO).Sum(T => T["LOC_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["LOC_FUND"]));


                    /// 規劃類明細
                    foreach (DataRow row in planningRows.Where(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "D",
                            Data = row,
                            Label = "",
                            ItemCount = 0,
                            CenFund = SEQ_N_CenFund,
                            LocFund = SEQ_N_LocFund,
                            SumFund = SEQ_N_CenFund + SEQ_N_LocFund,
                            PCenFund = PSEQ_N_CenFund,
                            PLocFund = PSEQ_N_LocFund,
                            PSumFund = PSEQ_N_CenFund + PSEQ_N_LocFund
                        });
                    }

                    /// 規劃類小計
                    if (planningRows.Count > 0)
                    {
                        renderItems.Add(new
                        {
                            Type = "P",
                            Data = (DataRow)null,
                            Label = "規劃類小計",
                            ItemCount = planningRows.Count,
                            CenFund = SEQ_N_CenFund,
                            LocFund = SEQ_N_LocFund,
                            SumFund = SEQ_N_CenFund + SEQ_N_LocFund,
                            PCenFund = PSEQ_N_CenFund,
                            PLocFund = PSEQ_N_LocFund,
                            PSumFund = PSEQ_N_CenFund + PSEQ_N_LocFund
                        });
                    }

                    /// 當前期數
                    double SEQ_C_CenFund = engineeringRows.Where(T => T["SEQ_NO"].ToString() == SEQNO).Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));
                    double SEQ_C_LocFund = engineeringRows.Where(T => T["SEQ_NO"].ToString() == SEQNO).Sum(T => T["LOC_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["LOC_FUND"]));

                    /// 前一期數
                    double PSEQ_C_CenFund = engineeringRows.Where(T => T["SEQ_NO"].ToString() == PSEQNO).Sum(T => T["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["CEN_FUND"]));
                    double PSEQ_C_LocFund = engineeringRows.Where(T => T["SEQ_NO"].ToString() == PSEQNO).Sum(T => T["LOC_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(T["LOC_FUND"]));

                    /// 工程類明細
                    foreach (DataRow row in engineeringRows.Where(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "D",
                            Data = row,
                            Label = "",
                            ItemCount = 0,
                            CenFund = SEQ_C_CenFund,
                            LocFund = SEQ_C_LocFund,
                            SumFund = SEQ_C_CenFund + SEQ_C_LocFund,
                            PCenFund = PSEQ_C_CenFund,
                            PLocFund = PSEQ_C_LocFund,
                            PSumFund = PSEQ_C_CenFund + PSEQ_C_LocFund
                        });
                    }

                    /// 工程類小計
                    if (engineeringRows.Count > 0)
                    {
                        renderItems.Add(new
                        {
                            Type = "E",
                            Data = (DataRow)null,
                            Label = "工程類小計",
                            ItemCount = engineeringRows.Count,
                            CenFund = SEQ_C_CenFund,
                            LocFund = SEQ_C_LocFund,
                            SumFund = SEQ_C_CenFund + SEQ_C_LocFund,
                            PCenFund = PSEQ_C_CenFund,
                            PLocFund = PSEQ_C_LocFund,
                            PSumFund = PSEQ_C_CenFund + PSEQ_C_LocFund
                        });
                    }

                    /// 總計
                    renderItems.Add(new
                    {
                        Type = "T",
                        Data = (DataRow)null,
                        Label = "總計",
                        ItemCount = cityRows.Count,
                        CenFund = SEQ_N_CenFund + SEQ_C_CenFund,
                        LocFund = SEQ_N_LocFund + SEQ_C_LocFund,
                        SumFund = SEQ_N_CenFund + SEQ_C_CenFund + SEQ_N_LocFund + SEQ_C_LocFund,
                        PCenFund = PSEQ_N_CenFund + PSEQ_C_CenFund,
                        PLocFund = PSEQ_N_LocFund + PSEQ_C_LocFund,
                        PSumFund = PSEQ_N_CenFund + PSEQ_C_CenFund + PSEQ_N_LocFund + PSEQ_C_LocFund

                    });

                    /// 依照輸出項目數量計算 block 數
                    int blockCount = (int)Math.Ceiling((double)renderItems.Count / itemPerBlock);

                    for (int blockIndex = 0; blockIndex < blockCount; blockIndex++)
                    {
                        /// 第一個 block 直接用原模板 ; 後續 block 複製
                        if (currentRow != templateStartRow)
                        {
                            worksheet4.Cells.CopyRows(worksheet4.Cells, templateStartRow, currentRow, blockHeight);
                        }

                        /// Block Layout
                        int titleRow = currentRow;           /// block 第1列
                        int periodRow = currentRow + 1;      /// block 第2列
                        int headerRow = currentRow + 2;      /// block 第3列
                        int dataStartRow = currentRow + 3;   /// block 第4列
                        int dataEndRow = currentRow + 24;    /// block 第25列

                        /// 清除資料區內容，保留格式
                        for (int r = dataStartRow; r <= dataEndRow; r++)
                        {
                            for (int c = 0; c <= 21; c++)   /// A ~ P，20260408 版本修正 15 => 21
                            {
                                worksheet4.Cells[r, c].PutValue(string.Empty);
                            }
                        }

                        /// 標題列
                        string ReportTitle = GetReportTitle(PID, connectionString);
                        worksheet4.Cells["A1"].PutValue($"{SYR}年度 {ReportTitle} 彙整明細");
                        worksheet4.Cells[titleRow, 0].PutValue($"{city}調整預算表");
                        worksheet4.Cells[periodRow, 7].PutValue(SEQNO == "0" ? "暫無期別" : $"第{PSEQNO}期");
                        worksheet4.Cells[periodRow, 10].PutValue($"第{SEQNO}期");

                        /// 取本 block 的資料
                        var blockItems = renderItems.Skip(blockIndex * itemPerBlock).Take(itemPerBlock).ToList();

                        for (int i = 0; i < blockItems.Count; i++)
                        {
                            int targetRow = dataStartRow + (i * 2);

                            dynamic item = blockItems[i];

                            /// 明細列
                            if (item.Type == "D")
                            {
                                DataRow data = item.Data;
                                DataRow prevData = cityRows.FirstOrDefault(r => r["SEQ_NO"].ToString() == PSEQNO && r["FD_CODE"].ToString() == data["FD_CODE"].ToString());

                                /// 前一期數
                                double PSEQ_CenFund = prevData == null || prevData["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(prevData["CEN_FUND"]);
                                double PSEQ_LocFund = prevData == null || prevData["LOC_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(prevData["LOC_FUND"]);
                                double PSEQ_SumFund = PSEQ_CenFund + PSEQ_LocFund;

                                /// 當前期數
                                double SEQ_CenFund = data["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(data["CEN_FUND"]);
                                double SEQ_LocFund = data["LOC_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(data["LOC_FUND"]);
                                double SEQ_SumFund = SEQ_CenFund + SEQ_LocFund;

                                worksheet4.Cells[targetRow, 0].PutValue(count++);                       // A 序號
                                worksheet4.Cells[targetRow, 1].PutValue(data["SIP_TD_CODE"]);           // B 工程編號
                                worksheet4.Cells[targetRow, 2].PutValue(data["FD_CODE"].ToString());    // C 經費編號
                                worksheet4.Cells[targetRow, 3].PutValue(data["TD_NAME"].ToString());    // D 工程名稱
                                worksheet4.Cells[targetRow, 4].PutValue(data["CATY_CODE"].ToString() == "C" ? "工程類" : "規劃類");
                                worksheet4.Cells[targetRow, 5].PutValue(data["IPJ_DESC"].ToString());
                                worksheet4.Cells[targetRow, 6].PutValue(data["CEN_SP_RATIO"] == DBNull.Value ? 0 : Convert.ToDouble(data["CEN_SP_RATIO"]));
                                worksheet4.Cells[targetRow, 7].PutValue(PSEQ_CenFund);
                                worksheet4.Cells[targetRow, 8].PutValue(PSEQ_LocFund);
                                worksheet4.Cells[targetRow, 9].PutValue(PSEQ_SumFund);
                                worksheet4.Cells[targetRow, 10].PutValue(SEQ_CenFund);
                                worksheet4.Cells[targetRow, 11].PutValue(SEQ_LocFund);
                                worksheet4.Cells[targetRow, 12].PutValue(SEQ_SumFund);
                                worksheet4.Cells[targetRow, 13].PutValue(SEQ_SumFund - PSEQ_SumFund);
                                worksheet4.Cells[targetRow, 14].PutValue(0);
                                worksheet4.Cells[targetRow, 15].PutValue(data["EXECT_DESC"].ToString());

                                worksheet4.Cells[targetRow, 16].PutValue(data["M_01"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_01"]));
                                worksheet4.Cells[targetRow, 17].PutValue(data["M_02"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_02"]));
                                worksheet4.Cells[targetRow, 18].PutValue(data["M_03"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_03"]));
                                worksheet4.Cells[targetRow, 19].PutValue(data["M_04"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_04"]));
                                worksheet4.Cells[targetRow, 20].PutValue(data["M_05"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_05"]));
                                worksheet4.Cells[targetRow, 21].PutValue(data["M_06"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_06"]));

                                worksheet4.Cells[targetRow + 1, 16].PutValue(data["M_07"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_07"]));
                                worksheet4.Cells[targetRow + 1, 17].PutValue(data["M_08"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_08"]));
                                worksheet4.Cells[targetRow + 1, 18].PutValue(data["M_09"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_09"]));
                                worksheet4.Cells[targetRow + 1, 19].PutValue(data["M_10"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_10"]));
                                worksheet4.Cells[targetRow + 1, 20].PutValue(data["M_11"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_11"]));
                                worksheet4.Cells[targetRow + 1, 21].PutValue(data["M_12"] == DBNull.Value ? 0 : Convert.ToDouble(data["M_12"]));
                            }
                            else
                            {
                                /// 小計 / 總計列
                                worksheet4.Cells[targetRow, 3].PutValue(item.Label);
                                worksheet4.Cells[targetRow, 7].PutValue(item.PCenFund);
                                worksheet4.Cells[targetRow, 8].PutValue(item.PLocFund);
                                worksheet4.Cells[targetRow, 9].PutValue(item.PSumFund);
                                worksheet4.Cells[targetRow, 10].PutValue(item.CenFund);
                                worksheet4.Cells[targetRow, 11].PutValue(item.LocFund);
                                worksheet4.Cells[targetRow, 12].PutValue(item.SumFund);
                            }
                        }

                        currentRow += blockHeight + 1;
                    }
                }
            }
        }

        /// <summary>
        /// B-1.修正預算作業表[彙整明細(簡易版)]_Sheet5
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="source"></param>
        /// <param name="SYR"></param>
        /// <param name="SMN"></param>
        private void B1_SimplifiedDetails(Worksheet worksheet5, DataTable source, string SYR, string SEQNO, string PID)
        {
            worksheet5.Name = "彙整明細(簡易版)";
            DataTable Table = source;

            string ReportTitle = GetReportTitle(PID, connectionString);
            worksheet5.Cells["A1"].PutValue($"{SYR}年度 {ReportTitle} 彙整明細(簡易版)");

            int count = 1;
            int templateStartRow = 1;   // A2 => index = 1
            int blockHeight = 14;       // 每個完整 block 14 列
            int itemPerBlock = 11;      // 每個 block 可放 11 個項目(每項目佔2列)
            int currentRow = templateStartRow;

            /// 分類前置作業 : CITY 分組
            var cityGroups = Table.AsEnumerable().GroupBy(r => r["CITY"].ToString()).OrderBy(g => g.Key).ToList();

            /// 前一期
            string PSEQNO = Math.Max(Convert.ToInt32(SEQNO) - 1, 0).ToString();

            /// 轉 double 共用
            double ToDouble(object value)
            {
                return value == DBNull.Value || value == null || string.IsNullOrWhiteSpace(value.ToString())
                    ? 0
                    : Convert.ToDouble(value);
            }

            /// 原預算
            if (SEQNO == "0")
            {

                foreach (var cityGroup in cityGroups)
                {
                    string city = cityGroup.Key;
                    List<DataRow> cityRows = cityGroup.ToList();

                    /// 拆成 規劃類 / 工程類
                    List<DataRow> planningRows = cityRows.Where(r => r["CATY_CODE"].ToString() != "C").ToList();
                    List<DataRow> engineeringRows = cityRows.Where(r => r["CATY_CODE"].ToString() == "C").ToList();

                    var renderItems = new List<object>();

                    /// 規劃類 - 當前期數
                    double SEQ_N_CenFund = planningRows.Where(r => r["SEQ_NO"].ToString() == SEQNO).Sum(r => ToDouble(r["CEN_FUND"]));
                    double SEQ_N_LocFund = planningRows.Where(r => r["SEQ_NO"].ToString() == SEQNO).Sum(r => ToDouble(r["LOC_FUND"]));

                    /// 規劃類 - 前一期數
                    double PSEQ_N_CenFund = planningRows.Where(r => r["SEQ_NO"].ToString() == PSEQNO).Sum(r => ToDouble(r["CEN_FUND"]));
                    double PSEQ_N_LocFund = planningRows.Where(r => r["SEQ_NO"].ToString() == PSEQNO).Sum(r => ToDouble(r["LOC_FUND"]));

                    /// 規劃類明細
                    foreach (DataRow row in planningRows.Where(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "D",
                            Data = row,
                            Label = "",
                            ItemCount = 0,
                            CenFund = SEQ_N_CenFund,
                            LocFund = SEQ_N_LocFund,
                            SumFund = SEQ_N_CenFund + SEQ_N_LocFund,
                            PCenFund = PSEQ_N_CenFund,
                            PLocFund = PSEQ_N_LocFund,
                            PSumFund = PSEQ_N_CenFund + PSEQ_N_LocFund
                        });
                    }

                    /// 規劃類小計
                    if (planningRows.Any(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "P",
                            Data = (DataRow)null,
                            Label = "規劃類小計",
                            ItemCount = planningRows.Count(r => r["SEQ_NO"].ToString() == SEQNO),
                            CenFund = SEQ_N_CenFund,
                            LocFund = SEQ_N_LocFund,
                            SumFund = SEQ_N_CenFund + SEQ_N_LocFund,
                            PCenFund = PSEQ_N_CenFund,
                            PLocFund = PSEQ_N_LocFund,
                            PSumFund = PSEQ_N_CenFund + PSEQ_N_LocFund
                        });
                    }

                    /// 工程類 - 當前期數
                    double SEQ_C_CenFund = engineeringRows.Where(r => r["SEQ_NO"].ToString() == SEQNO).Sum(r => ToDouble(r["CEN_FUND"]));
                    double SEQ_C_LocFund = engineeringRows.Where(r => r["SEQ_NO"].ToString() == SEQNO).Sum(r => ToDouble(r["LOC_FUND"]));

                    /// 工程類 - 前一期數
                    double PSEQ_C_CenFund = engineeringRows.Where(r => r["SEQ_NO"].ToString() == PSEQNO).Sum(r => ToDouble(r["CEN_FUND"]));
                    double PSEQ_C_LocFund = engineeringRows.Where(r => r["SEQ_NO"].ToString() == PSEQNO).Sum(r => ToDouble(r["LOC_FUND"]));

                    /// 工程類明細
                    foreach (DataRow row in engineeringRows.Where(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "D",
                            Data = row,
                            Label = "",
                            ItemCount = 0,
                            CenFund = SEQ_C_CenFund,
                            LocFund = SEQ_C_LocFund,
                            SumFund = SEQ_C_CenFund + SEQ_C_LocFund,
                            PCenFund = PSEQ_C_CenFund,
                            PLocFund = PSEQ_C_LocFund,
                            PSumFund = PSEQ_C_CenFund + PSEQ_C_LocFund
                        });
                    }

                    /// 工程類小計
                    if (engineeringRows.Any(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "E",
                            Data = (DataRow)null,
                            Label = "工程類小計",
                            ItemCount = engineeringRows.Count(r => r["SEQ_NO"].ToString() == SEQNO),
                            CenFund = SEQ_C_CenFund,
                            LocFund = SEQ_C_LocFund,
                            SumFund = SEQ_C_CenFund + SEQ_C_LocFund,
                            PCenFund = PSEQ_C_CenFund,
                            PLocFund = PSEQ_C_LocFund,
                            PSumFund = PSEQ_C_CenFund + PSEQ_C_LocFund
                        });
                    }

                    /// 總計
                    renderItems.Add(new
                    {
                        Type = "T",
                        Data = (DataRow)null,
                        Label = "總計",
                        ItemCount = cityRows.Count(r => r["SEQ_NO"].ToString() == SEQNO),
                        CenFund = SEQ_N_CenFund + SEQ_C_CenFund,
                        LocFund = SEQ_N_LocFund + SEQ_C_LocFund,
                        SumFund = SEQ_N_CenFund + SEQ_C_CenFund + SEQ_N_LocFund + SEQ_C_LocFund,
                        PCenFund = PSEQ_N_CenFund + PSEQ_C_CenFund,
                        PLocFund = PSEQ_N_LocFund + PSEQ_C_LocFund,
                        PSumFund = PSEQ_N_CenFund + PSEQ_C_CenFund + PSEQ_N_LocFund + PSEQ_C_LocFund
                    });

                    /// block 數量
                    int blockCount = (int)Math.Ceiling((double)renderItems.Count / itemPerBlock);

                    for (int blockIndex = 0; blockIndex < blockCount; blockIndex++)
                    {
                        /// 第一個 block 直接用原模板；後續 block 複製
                        if (currentRow != templateStartRow)
                        {
                            worksheet5.Cells.CopyRows(worksheet5.Cells, templateStartRow, currentRow, blockHeight);
                        }

                        /// Block Layout
                        int titleRow = currentRow;
                        int periodRow = currentRow + 1;
                        int headerRow = currentRow + 2;
                        int dataStartRow = currentRow + 3;
                        int dataEndRow = currentRow + 24;

                        /// 清除資料區內容，保留格式
                        for (int r = dataStartRow; r <= dataEndRow; r++)
                        {
                            for (int c = 0; c <= 14; c++)   // A ~ O
                            {
                                worksheet5.Cells[r, c].PutValue(string.Empty);
                            }
                        }

                        /// 標題列
                        worksheet5.Cells["A1"].PutValue($"{SYR}年度 {ReportTitle} 彙整明細(簡易版)");
                        worksheet5.Cells[titleRow, 0].PutValue($"{city}調整預算表");
                        worksheet5.Cells[periodRow, 7].PutValue(SEQNO == "0" ? "原預算" : $"第{PSEQNO}期");
                        worksheet5.Cells[periodRow, 10].PutValue($"暫無期別");

                        /// 取本 block 的資料
                        var blockItems = renderItems.Skip(blockIndex * itemPerBlock).Take(itemPerBlock).ToList();

                        for (int i = 0; i < blockItems.Count; i++)
                        {
                            int targetRow = dataStartRow + i;
                            dynamic item = blockItems[i];

                            /// 明細列
                            if (item.Type == "D")
                            {
                                DataRow data = item.Data;
                                DataRow prevData = cityRows.FirstOrDefault(r =>
                                    r["SEQ_NO"].ToString() == PSEQNO &&
                                    r["FD_CODE"].ToString() == data["FD_CODE"].ToString());

                                /// 前一期
                                double PSEQ_CenFund = prevData == null ? 0 : ToDouble(prevData["CEN_FUND"]);
                                double PSEQ_LocFund = prevData == null ? 0 : ToDouble(prevData["LOC_FUND"]);
                                double PSEQ_SumFund = PSEQ_CenFund + PSEQ_LocFund;

                                /// 本期
                                double SEQ_CenFund = ToDouble(data["CEN_FUND"]);
                                double SEQ_LocFund = ToDouble(data["LOC_FUND"]);
                                double SEQ_SumFund = SEQ_CenFund + SEQ_LocFund;

                                worksheet5.Cells[targetRow, 0].PutValue(count++);                         // A 序號
                                worksheet5.Cells[targetRow, 1].PutValue(data["SIP_TD_CODE"].ToString());  // B 工程編號
                                worksheet5.Cells[targetRow, 2].PutValue(data["FD_CODE"].ToString());      // C 經費編號
                                worksheet5.Cells[targetRow, 3].PutValue(data["TD_NAME"].ToString());      // D 工程名稱
                                worksheet5.Cells[targetRow, 4].PutValue(data["CATY_CODE"].ToString() == "C" ? "工程類" : "規劃類");
                                worksheet5.Cells[targetRow, 5].PutValue(data["IPJ_DESC"].ToString());      // F 辦理層級
                                worksheet5.Cells[targetRow, 6].PutValue(ToDouble(data["CEN_SP_RATIO"]));  // G 中央款比例
                                worksheet5.Cells[targetRow, 7].PutValue(SEQ_CenFund);                    // H 前一期中央補助款
                                worksheet5.Cells[targetRow, 8].PutValue(SEQ_LocFund);                    // I 前一期地方配合款
                                worksheet5.Cells[targetRow, 9].PutValue(SEQ_SumFund);                    // J 前一期小計
                                worksheet5.Cells[targetRow, 10].PutValue(0);                                // K 本期中央補助款
                                worksheet5.Cells[targetRow, 11].PutValue(0);                    // L 本期地方配合款
                                worksheet5.Cells[targetRow, 12].PutValue(0);                    // M 本期小計
                                worksheet5.Cells[targetRow, 13].PutValue(SEQ_SumFund - PSEQ_SumFund);     // N 中央款增減數
                                worksheet5.Cells[targetRow, 14].PutValue(0);                              // O 截至目前為止實支金額
                            }
                            else
                            {
                                /// 小計 / 總計列
                                worksheet5.Cells[targetRow, 3].PutValue(item.Label);
                                worksheet5.Cells[targetRow, 7].PutValue(item.CenFund);
                                worksheet5.Cells[targetRow, 8].PutValue(item.LocFund);
                                worksheet5.Cells[targetRow, 9].PutValue(item.SumFund);
                                worksheet5.Cells[targetRow, 10].PutValue(0);
                                worksheet5.Cells[targetRow, 11].PutValue(0);
                                worksheet5.Cells[targetRow, 12].PutValue(0);
                                worksheet5.Cells[targetRow, 13].PutValue(item.SumFund - item.PSumFund);
                                worksheet5.Cells[targetRow, 14].PutValue(0);
                            }
                        }

                        currentRow += blockHeight + 1;
                    }
                }

            }
            /// 調整預算後
            else
            {

                foreach (var cityGroup in cityGroups)
                {
                    string city = cityGroup.Key;
                    List<DataRow> cityRows = cityGroup.ToList();

                    /// 拆成 規劃類 / 工程類
                    List<DataRow> planningRows = cityRows.Where(r => r["CATY_CODE"].ToString() != "C").ToList();
                    List<DataRow> engineeringRows = cityRows.Where(r => r["CATY_CODE"].ToString() == "C").ToList();

                    var renderItems = new List<object>();

                    /// 規劃類 - 當前期數
                    double SEQ_N_CenFund = planningRows.Where(r => r["SEQ_NO"].ToString() == SEQNO).Sum(r => ToDouble(r["CEN_FUND"]));
                    double SEQ_N_LocFund = planningRows.Where(r => r["SEQ_NO"].ToString() == SEQNO).Sum(r => ToDouble(r["LOC_FUND"]));

                    /// 規劃類 - 前一期數
                    double PSEQ_N_CenFund = planningRows.Where(r => r["SEQ_NO"].ToString() == PSEQNO).Sum(r => ToDouble(r["CEN_FUND"]));
                    double PSEQ_N_LocFund = planningRows.Where(r => r["SEQ_NO"].ToString() == PSEQNO).Sum(r => ToDouble(r["LOC_FUND"]));

                    /// 規劃類明細
                    foreach (DataRow row in planningRows.Where(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "D",
                            Data = row,
                            Label = "",
                            ItemCount = 0,
                            CenFund = SEQ_N_CenFund,
                            LocFund = SEQ_N_LocFund,
                            SumFund = SEQ_N_CenFund + SEQ_N_LocFund,
                            PCenFund = PSEQ_N_CenFund,
                            PLocFund = PSEQ_N_LocFund,
                            PSumFund = PSEQ_N_CenFund + PSEQ_N_LocFund
                        });
                    }

                    /// 規劃類小計
                    if (planningRows.Any(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "P",
                            Data = (DataRow)null,
                            Label = "規劃類小計",
                            ItemCount = planningRows.Count(r => r["SEQ_NO"].ToString() == SEQNO),
                            CenFund = SEQ_N_CenFund,
                            LocFund = SEQ_N_LocFund,
                            SumFund = SEQ_N_CenFund + SEQ_N_LocFund,
                            PCenFund = PSEQ_N_CenFund,
                            PLocFund = PSEQ_N_LocFund,
                            PSumFund = PSEQ_N_CenFund + PSEQ_N_LocFund
                        });
                    }

                    /// 工程類 - 當前期數
                    double SEQ_C_CenFund = engineeringRows.Where(r => r["SEQ_NO"].ToString() == SEQNO).Sum(r => ToDouble(r["CEN_FUND"]));
                    double SEQ_C_LocFund = engineeringRows.Where(r => r["SEQ_NO"].ToString() == SEQNO).Sum(r => ToDouble(r["LOC_FUND"]));

                    /// 工程類 - 前一期數
                    double PSEQ_C_CenFund = engineeringRows.Where(r => r["SEQ_NO"].ToString() == PSEQNO).Sum(r => ToDouble(r["CEN_FUND"]));
                    double PSEQ_C_LocFund = engineeringRows.Where(r => r["SEQ_NO"].ToString() == PSEQNO).Sum(r => ToDouble(r["LOC_FUND"]));

                    /// 工程類明細
                    foreach (DataRow row in engineeringRows.Where(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "D",
                            Data = row,
                            Label = "",
                            ItemCount = 0,
                            CenFund = SEQ_C_CenFund,
                            LocFund = SEQ_C_LocFund,
                            SumFund = SEQ_C_CenFund + SEQ_C_LocFund,
                            PCenFund = PSEQ_C_CenFund,
                            PLocFund = PSEQ_C_LocFund,
                            PSumFund = PSEQ_C_CenFund + PSEQ_C_LocFund
                        });
                    }

                    /// 工程類小計
                    if (engineeringRows.Any(r => r["SEQ_NO"].ToString() == SEQNO))
                    {
                        renderItems.Add(new
                        {
                            Type = "E",
                            Data = (DataRow)null,
                            Label = "工程類小計",
                            ItemCount = engineeringRows.Count(r => r["SEQ_NO"].ToString() == SEQNO),
                            CenFund = SEQ_C_CenFund,
                            LocFund = SEQ_C_LocFund,
                            SumFund = SEQ_C_CenFund + SEQ_C_LocFund,
                            PCenFund = PSEQ_C_CenFund,
                            PLocFund = PSEQ_C_LocFund,
                            PSumFund = PSEQ_C_CenFund + PSEQ_C_LocFund
                        });
                    }

                    /// 總計
                    renderItems.Add(new
                    {
                        Type = "T",
                        Data = (DataRow)null,
                        Label = "總計",
                        ItemCount = cityRows.Count(r => r["SEQ_NO"].ToString() == SEQNO),
                        CenFund = SEQ_N_CenFund + SEQ_C_CenFund,
                        LocFund = SEQ_N_LocFund + SEQ_C_LocFund,
                        SumFund = SEQ_N_CenFund + SEQ_C_CenFund + SEQ_N_LocFund + SEQ_C_LocFund,
                        PCenFund = PSEQ_N_CenFund + PSEQ_C_CenFund,
                        PLocFund = PSEQ_N_LocFund + PSEQ_C_LocFund,
                        PSumFund = PSEQ_N_CenFund + PSEQ_C_CenFund + PSEQ_N_LocFund + PSEQ_C_LocFund
                    });

                    /// block 數量
                    int blockCount = (int)Math.Ceiling((double)renderItems.Count / itemPerBlock);

                    for (int blockIndex = 0; blockIndex < blockCount; blockIndex++)
                    {
                        /// 第一個 block 直接用原模板；後續 block 複製
                        if (currentRow != templateStartRow)
                        {
                            worksheet5.Cells.CopyRows(worksheet5.Cells, templateStartRow, currentRow, blockHeight);
                        }

                        /// Block Layout
                        int titleRow = currentRow;
                        int periodRow = currentRow + 1;
                        int headerRow = currentRow + 2;
                        int dataStartRow = currentRow + 3;
                        int dataEndRow = currentRow + 24;

                        /// 清除資料區內容，保留格式
                        for (int r = dataStartRow; r <= dataEndRow; r++)
                        {
                            for (int c = 0; c <= 14; c++)   // A ~ O
                            {
                                worksheet5.Cells[r, c].PutValue(string.Empty);
                            }
                        }

                        /// 標題列
                        worksheet5.Cells["A1"].PutValue($"{SYR}年度 {ReportTitle} 彙整明細(簡易版)");
                        worksheet5.Cells[titleRow, 0].PutValue($"{city}調整預算表");
                        worksheet5.Cells[periodRow, 7].PutValue(SEQNO == "0" ? "暫無期別" : $"第{PSEQNO}期");
                        worksheet5.Cells[periodRow, 10].PutValue($"第{SEQNO}期");

                        /// 取本 block 的資料
                        var blockItems = renderItems.Skip(blockIndex * itemPerBlock).Take(itemPerBlock).ToList();

                        for (int i = 0; i < blockItems.Count; i++)
                        {
                            int targetRow = dataStartRow + i;
                            dynamic item = blockItems[i];

                            /// 明細列
                            if (item.Type == "D")
                            {
                                DataRow data = item.Data;
                                DataRow prevData = cityRows.FirstOrDefault(r =>
                                    r["SEQ_NO"].ToString() == PSEQNO &&
                                    r["FD_CODE"].ToString() == data["FD_CODE"].ToString());

                                /// 前一期
                                double PSEQ_CenFund = prevData == null ? 0 : ToDouble(prevData["CEN_FUND"]);
                                double PSEQ_LocFund = prevData == null ? 0 : ToDouble(prevData["LOC_FUND"]);
                                double PSEQ_SumFund = PSEQ_CenFund + PSEQ_LocFund;

                                /// 本期
                                double SEQ_CenFund = ToDouble(data["CEN_FUND"]);
                                double SEQ_LocFund = ToDouble(data["LOC_FUND"]);
                                double SEQ_SumFund = SEQ_CenFund + SEQ_LocFund;

                                worksheet5.Cells[targetRow, 0].PutValue(count++);                         // A 序號
                                worksheet5.Cells[targetRow, 1].PutValue(data["SIP_TD_CODE"].ToString());  // B 工程編號
                                worksheet5.Cells[targetRow, 2].PutValue(data["FD_CODE"].ToString());      // C 經費編號
                                worksheet5.Cells[targetRow, 3].PutValue(data["TD_NAME"].ToString());      // D 工程名稱
                                worksheet5.Cells[targetRow, 4].PutValue(data["CATY_CODE"].ToString() == "C" ? "工程類" : "規劃類");
                                worksheet5.Cells[targetRow, 5].PutValue(data["IPJ_DESC"].ToString());      // F 辦理層級
                                worksheet5.Cells[targetRow, 6].PutValue(ToDouble(data["CEN_SP_RATIO"]));  // G 中央款比例
                                worksheet5.Cells[targetRow, 7].PutValue(PSEQ_CenFund);                    // H 前一期中央補助款
                                worksheet5.Cells[targetRow, 8].PutValue(PSEQ_LocFund);                    // I 前一期地方配合款
                                worksheet5.Cells[targetRow, 9].PutValue(PSEQ_SumFund);                    // J 前一期小計
                                worksheet5.Cells[targetRow, 10].PutValue(SEQ_CenFund);                    // K 本期中央補助款
                                worksheet5.Cells[targetRow, 11].PutValue(SEQ_LocFund);                    // L 本期地方配合款
                                worksheet5.Cells[targetRow, 12].PutValue(SEQ_SumFund);                    // M 本期小計
                                worksheet5.Cells[targetRow, 13].PutValue(SEQ_SumFund - PSEQ_SumFund);     // N 中央款增減數
                                worksheet5.Cells[targetRow, 14].PutValue(0);                              // O 截至目前為止實支金額
                            }
                            else
                            {
                                /// 小計 / 總計列
                                worksheet5.Cells[targetRow, 3].PutValue(item.Label);
                                worksheet5.Cells[targetRow, 7].PutValue(item.PCenFund);
                                worksheet5.Cells[targetRow, 8].PutValue(item.PLocFund);
                                worksheet5.Cells[targetRow, 9].PutValue(item.PSumFund);
                                worksheet5.Cells[targetRow, 10].PutValue(item.CenFund);
                                worksheet5.Cells[targetRow, 11].PutValue(item.LocFund);
                                worksheet5.Cells[targetRow, 12].PutValue(item.SumFund);
                                worksheet5.Cells[targetRow, 13].PutValue(item.SumFund - item.PSumFund);
                                worksheet5.Cells[targetRow, 14].PutValue(0);
                            }
                        }

                        currentRow += blockHeight + 1;
                    }
                }

            }
        }

        /// <summary>
        /// E-1.功能細項執行表(自訂欄位)
        /// </summary>
        /// <param name="plan"></param>
        /// <param name="CITY"></param>
        /// <param name="Row"></param>
        /// <param name="Field"></param>
        /// <param name="COLUMN1"></param>
        /// <param name="PDF"></param>
        /// <returns></returns>
        public ActionResult ReportPrint9(string[] plan, string[] CITY, string Row, string Field, string[] COLUMN1, string PDF)
        {
            var path = Server.MapPath("~/Templates/表1.xlsx");
            var file = new FileInfo(path);

            List<string> modifiedCities = new List<string>(CITY);

            if (Array.Exists(CITY, element => element == "新北市"))
            {
                modifiedCities.Add("台北縣");
            }

            if (Array.Exists(CITY, element => element == "桃園市"))
            {
                modifiedCities.Add("桃園縣");
            }

            if (Array.Exists(CITY, element => element == "台中市"))
            {
                modifiedCities.Add("台中縣");
            }

            if (Array.Exists(CITY, element => element == "台南市"))
            {
                modifiedCities.Add("台南縣");
            }

            if (Array.Exists(CITY, element => element == "高雄市"))
            {
                modifiedCities.Add("高雄縣");
            }

            // 將修改後的城市列表轉換回陣列並重新指派給 CITY 變數
            CITY = modifiedCities.ToArray();

            string[] COLUMN = new string[COLUMN1.Length + 1];
            COLUMN[0] = "年分";
            Array.Copy(COLUMN1, 0, COLUMN, 1, COLUMN1.Length);

            var dataTable = new DataTable();


            foreach (var columnName in COLUMN)
            {
                dataTable.Columns.Add(columnName);
            }


            var TDMFORM = db3.TDMFORM.ToList();
            var OGCDEPT = db.OGCDEPT.ToList();
            var CTDITFD = db2.CTDITFD.Where(m => m.FIT_CODE == "11").ToList();

            // 20241024 修改項目 => 總工程費(元)
            // 查詢各項(REC_TP = B)子工程項目的最大預算(SEQ_NO)
            #region 程式碼(舊版)
            //var TDDBGLS = db3.TDDBGLS.ToList().Where(m => m.SEQ_NO == maxSeqNo && m.REC_TP == "B").GroupBy(m => m.TD_CODE).Select(p => new
            //{
            //    TD_CODE = p.Key,
            //    Total = p.Sum(m => m.BUDGET_FEE)
            //}).ToList();
            #endregion
            var TDDBGLS = db3.TDDBGLS
                .GroupBy(m => m.TD_CODE)
                .Select(g => new
                {
                    TD_CODE = g.Key,
                    MaxSeqNo = g.Max(m => m.SEQ_NO)
                })
                .ToList()
                .SelectMany(g => db3.TDDBGLS
                    .Where(m => m.TD_CODE == g.TD_CODE && m.SEQ_NO == g.MaxSeqNo && m.REC_TP == "B")
                    .GroupBy(m => m.TD_CODE)
                    .Select(p => new
                    {
                        TD_CODE = p.Key,
                        Total = p.Sum(m => m.BUDGET_FEE)
                    })
                )
                .ToList();


            // 20241024 修改項目 => 修改取得日期條件
            //var TDDSDWB_CONTRATE = db3.TDDSDWB.Where(m => m.IT_NAME == "發包日期").ToList();
            //var TDDSDWB_START = db3.TDDSDWB.Where(m => m.IT_NAME == "工程開工").ToList();
            //var TDDSDWB_END = db3.TDDSDWB.Where(m => m.IT_NAME == "工程完工").ToList();
            var TDDSDWB_CONTRATE = db3.TDDSDWB.Where(m => m.IT_CODE == "TK105").ToList(); //發包日期
            var TDDSDWB_START = db3.TDDSDWB.Where(m => m.IT_CODE == "TK106").ToList(); //工程開工日期
            var TDDSDWB_END = db3.TDDSDWB.Where(m => m.IT_CODE == "TK107").ToList(); //工程完工日期

            // 20241024 修改項目 => (細部)設計完成日期
            var TDDSDWB_TK103 = db3.TDDSDWB.Where(m => m.IT_CODE == "TK103").ToList();

            // 20241024 修改項目 => 已發包的預算款項
            var oriBUDGET = (from t in db3.TDDBGLS
                             join s in db3.SMCCODE on t.AC_ITEM_CODE equals s.CATY_COD
                             where t.REC_TP == "B"
                             && t.OUTSRC_FG == true
                             && t.SEQ_NO == (db3.TDDBGLS
                                             .Where(x => x.TD_CODE == t.TD_CODE && x.REC_TP == "B")
                                             .Max(x => x.SEQ_NO))
                             && s.SYS_ID == "SM" && s.CATY_ID == "03"
                             select new
                             {
                                 t.TD_CODE,
                                 s.CODE_DESC,
                                 t.SEQ_NO,
                                 t.IT_ORDER
                             } into temp
                             group temp by temp.TD_CODE into g
                             select new
                             {
                                 TD_CODE = g.Key,
                                 CODE_DESCs = g.OrderByDescending(x => x.SEQ_NO).OrderBy(x => x.IT_ORDER).Select(x => x.CODE_DESC).ToList()
                             }).ToList();

            // 20241024 修改項目 => 已發包的預算款項
            var BUDGET = oriBUDGET
                        .Select(x => new
                        {
                            x.TD_CODE,
                            ALL_OUTSRC_FG = string.Join(",\n", x.CODE_DESCs)
                        }).ToList();

            string Year = "";
            string Master_Plan_Name = "";
            string Plan_Name = "";

            /// 把重複年度月份的資料和再一起
            var query_source = (from item in db2.TDDSTCF
                                group item by new { item.SYR, item.SMN, item.TD_CODE1 } into grouped
                                select new
                                {
                                    SYR = grouped.Key.SYR,
                                    SMN = grouped.Key.SMN,
                                    TD_CODE = grouped.Key.TD_CODE1,
                                    Total_Cost = grouped.Sum(x => x.COST_FEE),
                                    Total_PPAY = grouped.Sum(x => x.PPAY_FEE),
                                }).ToList();

            /// query 是抓每個年度最後月份的經費
            var query = (from t in query_source.ToList()
                         group t by new { t.TD_CODE, t.SYR } into g
                         select new
                         {
                             SYR = g.Key.SYR,
                             SMN = g.OrderByDescending(t => t.SMN).First().SMN,
                             COST_FEE = g.OrderByDescending(t => t.SMN).First().Total_Cost,
                             PPAY_FEE = g.OrderByDescending(t => t.SMN).First().Total_PPAY,
                             TD_CODE1 = g.OrderByDescending(t => t.SMN).First().TD_CODE
                         }).OrderBy(m => m.SYR).ToList();

            /// NewList 是把 COST_FEE 依照年分加總前面年份的數值做的表
            var newList = query.Select(item => new
            {
                SYR = item.SYR,
                SMN = item.SMN,
                COST_FEE = item.COST_FEE, /// 年度實支數
                PPAY_FEE = item.PPAY_FEE, /// 年度暫付數
                TD_CODE1 = item.TD_CODE1,
                SUM_COST_FEE = query
                .Where(t => t.TD_CODE1 == item.TD_CODE1 && t.SYR.CompareTo(item.SYR) <= 0)
                .Sum(t => t.COST_FEE),
                SUM_PPAY_FEE = query
                .Where(t => t.TD_CODE1 == item.TD_CODE1 && t.SYR.CompareTo(item.SYR) <= 0)
                .Sum(t => t.PPAY_FEE),
                SYR_TD_CODE = item.SYR + item.TD_CODE1
            }).OrderBy(m => m.SYR).ToList();


            var query2 = from t in db3.TDDMSCD.ToList()
                         group t by new { t.TD_CODE, t.SYR } into g
                         select new
                         {
                             TD_CODE = g.Key.TD_CODE,
                             SYR = g.Key.SYR,
                             SMN = g.Max(t => t.SMN),
                             PRE_SCHD = g.Max(t => t.PRE_SCHD),
                             REL_SCHD = g.Max(t => t.REL_SCHD),
                             EXECT_DESC = g.Max(t => t.EXECT_DESC),
                             DELAY_RESN = g.Max(t => t.DELAY_RESN),
                             CTDELAY_RESN = g.Max(t => t.CTDELAY_RESN),
                         };

            var CTDITFD_SUM = db2.CTDITFD.GroupBy(m => m.TID).Select(p => new
            {
                TID = p.Key,
                Total = p.Sum(m => m.FUND)
            }).ToList();

            /// 新增一個新欄位 SYR + TD_CODE 給 TDV_TDDCFTS
            var TDV_TDDCFTS = db2.TDV_TDDCFTS.Where(m => m.YR_FUND != 0 || (m.FCAL_FG ?? 0 == 1)).Where(s => s.CATY_TYPE != null).ToList();

            var TDV_TDDCFTS2 = (TDV_TDDCFTS.Select(m => new
            {
                m.PID,
                m.TD_CODE,
                m.TD_NAME,
                m.CITY,
                m.TOWN,
                m.IPJ_CODE,
                m.IPJ_DESC,
                m.ADV_DP,
                m.TID,
                m.CEN_SP_RATIO,
                m.CEN_FUND,
                SYR_TD_CODE = m.PID_SYR + m.TD_CODE
            }
            )).ToList();

            //var TDV_TDDCFTS_SUM = db2.TDV_TDDCFTS.GroupBy(m => m.TD_CODE).Select(p => new
            //{
            //    TD_CODE = p.Key,
            //    Total = p.Sum(m => m.CEN_FUND)
            //}).ToList();

            /// Start Merge   
            for (int w = 0; w < plan.Count(); w++)
            {
                string[] sArray;
                sArray = plan[w].Split('+');
                Year = sArray[0];
                Master_Plan_Name = sArray[1];
                Plan_Name = sArray[2];

                var PJMFORM = db2.PJMFORM.Where(m => m.SYR == Year && m.PJ_NAME == Plan_Name).FirstOrDefault();

                var targetPid = PJMFORM?.PID;

                var TDV_TDDCFTS_SUM = TDV_TDDCFTS
                    .Where(m => m.PID == targetPid
                             && !string.IsNullOrEmpty(m.PID_SYR)
                             && m.PID_SYR == Year)
                    .GroupBy(m => m.TD_CODE)
                    .Select(p => new
                    {
                        TD_CODE = p.Key,
                        Total = p.Sum(m => m.CEN_FUND)
                    })
                    .ToList();

                var TDDSTCF = newList.Where(m => m.SYR == Year).ToList(); /// 某年度的累積撥款數

                int yearNumber = int.Parse(Year);

                var TDDSTCF_TOTAL = query.Where(m => int.Parse(m.SYR) <= yearNumber).GroupBy(m => m.TD_CODE1).Select(p => new
                {
                    TDCODE1 = p.Key,
                    Total = p.Sum(m => m.COST_FEE) + p.Sum(m => m.PPAY_FEE)
                }).ToList();

                var TDDMSCD = query2.Where(m => m.SYR == Year).ToList();

                var SOURCE_MERGE = (from a in db2.PJMFORM.Where(m => m.SYR == Year && m.PJ_NAME == Plan_Name).ToList()
                                    join b in TDV_TDDCFTS2 on a.PID equals b.PID into Joinb
                                    from b in Joinb.DefaultIfEmpty()
                                    join c in TDV_TDDCFTS_SUM on b.TD_CODE equals c.TD_CODE into Joinc
                                    from c in Joinc.DefaultIfEmpty()
                                    join d in TDMFORM on b.TD_CODE equals d.TD_CODE into Joind
                                    from d in Joind.DefaultIfEmpty()
                                    join e in OGCDEPT on b.ADV_DP equals e.DEPT_NO into Joine
                                    from e in Joine.DefaultIfEmpty()
                                    join f in TDDBGLS on b.TD_CODE equals f.TD_CODE into Joinf
                                    from f in Joinf.DefaultIfEmpty()
                                    join g in newList on b.SYR_TD_CODE equals g.SYR_TD_CODE into Joing
                                    from g in Joing.DefaultIfEmpty()
                                    join i in TDDSDWB_CONTRATE on b.TD_CODE equals i.TD_CODE into contrateJoin
                                    from i in contrateJoin.DefaultIfEmpty()
                                    join j in TDDSDWB_START on b.TD_CODE equals j.TD_CODE into startJoin
                                    from j in startJoin.DefaultIfEmpty()
                                    join k in TDDSDWB_END on b.TD_CODE equals k.TD_CODE into endJoin
                                    from k in endJoin.DefaultIfEmpty()
                                    join l in TDDMSCD on b.TD_CODE equals l.TD_CODE into query2Join
                                    from l in query2Join.DefaultIfEmpty()
                                        // 20241024 修改項目 => 取得預定和實際(細部)設計完成日期
                                    join tk103 in TDDSDWB_TK103 on b.TD_CODE equals tk103.TD_CODE into tk103Join
                                    from tk103 in tk103Join.DefaultIfEmpty()
                                        // 20241024 修改項目 => 取得已發包的預算款項名稱
                                    join bud in BUDGET on b.TD_CODE equals bud.TD_CODE into budJoin
                                    from bud in budJoin.DefaultIfEmpty()

                                    select new CTDMNDSmergeTDV_TDDCFTS
                                    {
                                        PJ_NAME = Plan_Name,
                                        PID = b?.PID != null ? (decimal)b.PID : 0,
                                        TD_NAME = b.TD_NAME,
                                        TD_CITY = b.CITY,
                                        TD_TOWN = b.TOWN,
                                        TD_CODE = b?.TD_CODE,
                                        SYR = b?.SYR_TD_CODE.Substring(0, 3),
                                        IPJ_CODE = b.IPJ_CODE,
                                        IPJ_DESC = b.IPJ_DESC, //// 辦理層級
                                        ADV_DP = b.ADV_DP,
                                        TID = b?.TID != null ? (decimal)b.TID : 0,
                                        FUND = c?.Total != null ? (decimal)c.Total : 0, //// 中央款
                                        PERCENT_FUND = b.CEN_SP_RATIO != null ? (decimal)b.CEN_SP_RATIO : 0, /// 工程費比例
                                        TOTAL_ALL_BUDGET_FEE = b?.CEN_FUND != null ? (decimal)b.CEN_FUND : 0,
                                        TD_LINE = d?.TD_LINE,
                                        TD_NOTE = d?.TD_NOTE,
                                        TD_MEMO = d?.TD_MEMO,
                                        SMP_DEPT_NAME = e?.SMP_DEPT_NAME,
                                        TOTAL_FUND = f?.Total != null ? (decimal)f?.Total : 0, /// 總工程費(元)
                                        COST_FEE = (g?.COST_FEE != null ? (decimal)g.COST_FEE : 0) + (g?.PPAY_FEE != null ? (decimal)g.PPAY_FEE : 0),/// 當年度撥款數
                                        COST_FEE_TOTAL = (g?.SUM_COST_FEE != null ? (decimal)g.SUM_COST_FEE : 0) + (g?.SUM_PPAY_FEE != null ? (decimal)g.SUM_PPAY_FEE : 0), /// 累積撥款數
                                        PRE_CONTRACT_DATE = i?.PRE_EDT,
                                        REL_CONTRACT_DATE = i?.REL_EDT,
                                        PRE_START_DATE = j?.PRE_EDT,
                                        REL_START_DATE = j?.REL_EDT,
                                        PRE_END_DATE = k?.PRE_EDT,
                                        REL_END_DATE = k?.REL_EDT,
                                        PRE_SCHD = l?.PRE_SCHD != null ? (decimal)l.PRE_SCHD : 0,
                                        REL_SCHD = l?.REL_SCHD != null ? (decimal)l.REL_SCHD : 0,
                                        EXECT_DESC = l?.EXECT_DESC,
                                        DELAY_RESN = l?.DELAY_RESN,
                                        CTDELAY_RESN = l?.CTDELAY_RESN,
                                        PROGRESS_DIFFERENCE = (l?.PRE_SCHD != null ? (decimal)l.PRE_SCHD : 0) - (l?.REL_SCHD != null ? (decimal)l.REL_SCHD : 0),
                                        // 20241024 修改項目 => 取得報表新增的需求欄位資料
                                        PRE_TK103_DATE = tk103?.PRE_EDT,
                                        REL_TK103_DATE = tk103?.REL_EDT,
                                        ALL_OUTSRC_FG = bud?.ALL_OUTSRC_FG,
                                        DSG_BUSS_NAME = d.DSG_BUSS_NAME,
                                        SVY_BUSS_NAME = d.SVY_BUSS_NAME,
                                        BUSS_NAME = d.BUSS_NAME,
                                        FINISH_FG = d.TD_FINISH == true ? "是" : "",
                                        START_FUND = d.START_FUND,
                                        HOUSEHOLD = d.HOUSEHOLD,
                                        TOTAL_HOUSEHOLD = d.TOTAL_HOUSEHOLD,
                                        WTR_PL_NEW_LEN = d.WTR_PL_NEW_LEN,
                                        WTR_PL_REP_LEN = d.WTR_PL_REP_LEN,
                                        WTR_FLOOD = d.WTR_FLOOD,
                                        WTR_TOTAL_PUMP = d.WTR_TOTAL_PUMP,
                                        WTR_GU_STA = d.WTR_GU_STA,
                                    }).GroupBy(m => m.TD_CODE)
                                    .Select(group => group.OrderByDescending(m => m.SYR).First()).ToList();

                string MAX_IPJ_CODE = SOURCE_MERGE.Max(m => m.IPJ_CODE);
                SOURCE_MERGE = SOURCE_MERGE.Where(m => CITY.Any(value => m.TD_CITY.Contains(value))).ToList();
                SOURCE_MERGE = SOURCE_MERGE.OrderBy(m => m.TD_CITY).ToList();

                foreach (var item in SOURCE_MERGE)
                {
                    var row = dataTable.NewRow();
                    var value = item.GetType().GetProperty("SYR")?.GetValue(item);
                    row["年分"] = value;
                    foreach (var columnName in COLUMN)
                    {
                        if (columnName == "年分")
                        {
                            continue;
                        }
                        var COLUMN_IDNAME = db.OGCTITLE.Where(m => m.TITLE == columnName).FirstOrDefault().NAME;
                        value = item.GetType().GetProperty(COLUMN_IDNAME)?.GetValue(item);
                        row[columnName] = value;
                    }
                    dataTable.Rows.Add(row);
                }
            }

            var rowcount = dataTable.Rows.Count;

            /// Aspose.Cells
            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            var workbook = new Aspose.Cells.Workbook(path);
            var ws1 = workbook.Worksheets["工程細項執行表(縣市在列)"];
            var Column_length = Report_Source.NumberToColumnLetter(COLUMN.Length + 1);

            var A = ws1.Cells.CreateRange("A1:A1");
            var B = ws1.Cells.CreateRange("A1:" + Column_length + (rowcount + 2));
            var R = ws1.Cells.CreateRange("A1", Column_length + (rowcount + 2 + 1));

            Aspose.Cells.Style style1 = workbook.CreateStyle();

            if (Field == "on")
            {
                ws1.Name = "工程細項執行表(縣市在欄)";
                Column_length = Report_Source.NumberToColumnLetter(rowcount + 1);
                R = ws1.Cells.CreateRange("A1", Column_length + (COLUMN.Length + 2));

                Aspose.Cells.Style borderstyle = ws1.Workbook.CreateStyle();
                // Set outer borders
                Aspose.Cells.Border outerBorder = borderstyle.Borders[Aspose.Cells.BorderType.TopBorder];
                outerBorder.LineStyle = Aspose.Cells.CellBorderType.Thin;
                outerBorder = borderstyle.Borders[Aspose.Cells.BorderType.BottomBorder];
                outerBorder.LineStyle = Aspose.Cells.CellBorderType.Thin;
                outerBorder = borderstyle.Borders[Aspose.Cells.BorderType.LeftBorder];
                outerBorder.LineStyle = Aspose.Cells.CellBorderType.Thin;
                outerBorder = borderstyle.Borders[Aspose.Cells.BorderType.RightBorder];
                outerBorder.LineStyle = Aspose.Cells.CellBorderType.Thin;

                // Set inner borders
                Aspose.Cells.Border innerBorder = borderstyle.Borders[Aspose.Cells.BorderType.Vertical];
                innerBorder.LineStyle = Aspose.Cells.CellBorderType.Thin;
                innerBorder = borderstyle.Borders[Aspose.Cells.BorderType.Horizontal];
                innerBorder.LineStyle = Aspose.Cells.CellBorderType.Thin;

                borderstyle.IsTextWrapped = true; // 自動換行
                borderstyle.HorizontalAlignment = Aspose.Cells.TextAlignmentType.Center;
                R.SetStyle(borderstyle);


                int j = 2;
                for (var k = 0; k < rowcount; k++)
                {
                    var ABC = Report_Source.NumberToColumnLetter(j);
                    var a1 = ws1.Cells[ABC + "2"];
                    a1.PutValue(dataTable.Rows[j - 2]["工程地點(址)"]);
                    a1.GetStyle().Borders.SetStyle(CellBorderType.Thin);

                    for (int i = 3; i < COLUMN.Length + 3; i++)
                    {
                        var b = ws1.Cells[ABC + i];
                        Aspose.Cells.Style b_Style = a1.GetStyle();
                        var title = COLUMN[i - 3];
                        if (title == "年分")
                        {
                            ws1.Cells[ABC + "3"].PutValue(dataTable.Rows[j - 2]["年分"]);
                        }
                        if (title == "序號")
                        {
                            ws1.Cells[ABC + "4"].PutValue(j - 1);
                        }
                        else if (title == "補助對象")
                        {
                            b.PutValue(dataTable.Rows[j - 2]["工程地點(址)"] + "政府");
                        }
                        else if (title == "落後原因")
                        {
                            b.PutValue(dataTable.Rows[j - 2]["落後原因"]);
                        }
                        else if (title == "總工程費(元)")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 2]["總工程費(元)"]);
                            b.PutValue(value.ToString("#,##0"));
                            b_Style.Custom = "0.00_);[Red]\\(0.00\\)";
                            b.SetStyle(b_Style);
                        }
                        else if (title == "本署補助經費(中央款)(元)")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 2]["本署補助經費(中央款)(元)"]);
                            b.PutValue(value.ToString("#,##0"));
                            b_Style.Custom = "0.00_);[Red]\\(0.00\\)";
                            b.SetStyle(b_Style);
                        }
                        else if (title == "工程費比例")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 2]["工程費比例"]);
                            b.PutValue(value.ToString("#,##0.00"));
                            b_Style.Custom = "0.00_);[Red]\\(0.00\\)";
                            b.SetStyle(b_Style);
                        }
                        else if (title == "當年度核定數")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 2]["當年度核定數"]);
                            b.PutValue(value.ToString("#,##0"));
                            b_Style.Custom = "0.00_);[Red]\\(0.00\\)";
                            b.SetStyle(b_Style);
                        }
                        else if (title == "當年度撥款數")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 2]["當年度撥款數"]);
                            b.PutValue(value.ToString("#,##0"));
                            b_Style.Custom = "0.00_);[Red]\\(0.00\\)";
                            b.SetStyle(b_Style);
                        }
                        else if (title == "累積撥款數")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 2]["累積撥款數"]);
                            b.PutValue(value.ToString("#,##0"));
                            b_Style.Custom = "0.00_);[Red]\\(0.00\\)";
                            b.SetStyle(b_Style);
                        }
                        else
                        {
                            b.PutValue(dataTable.Rows[j - 2][title].ToString());
                        }

                        var a = ws1.Cells["A" + i];
                        a.PutValue(title);
                        a.GetStyle().Borders.SetStyle(CellBorderType.Thin);
                    }
                    j++;
                }
            }
            else
            {
                Aspose.Cells.Style borderstyle = ws1.Workbook.CreateStyle();
                Aspose.Cells.Border outerBorder = borderstyle.Borders[BorderType.TopBorder];
                outerBorder.LineStyle = CellBorderType.Thin;

                outerBorder = borderstyle.Borders[BorderType.BottomBorder];
                outerBorder.LineStyle = CellBorderType.Thin;

                outerBorder = borderstyle.Borders[BorderType.LeftBorder];
                outerBorder.LineStyle = CellBorderType.Thin;

                outerBorder = borderstyle.Borders[BorderType.RightBorder];
                outerBorder.LineStyle = CellBorderType.Thin;

                Aspose.Cells.Border innerBorder = borderstyle.Borders[BorderType.Vertical];
                innerBorder.LineStyle = CellBorderType.Thin;

                innerBorder = borderstyle.Borders[BorderType.Horizontal];
                innerBorder.LineStyle = CellBorderType.Thin;
                borderstyle.IsTextWrapped = true;  /// 自動換行
                borderstyle.HorizontalAlignment = TextAlignmentType.Center;
                R.SetStyle(borderstyle);

                B.Copy(A);
                for (int i = 2; i < COLUMN.Length + 2; i++)
                {
                    var ABC = Report_Source.NumberToColumnLetter(i);
                    var a = ws1.Cells[ABC + "2"];
                    var title = COLUMN[i - 2];
                    a.PutValue(title);
                    a.GetStyle().Borders.SetStyle(CellBorderType.Thin);
                }

                int j = 3;
                for (var n = 0; n < rowcount; n++)
                {
                    var a0 = ws1.Cells["A" + j];
                    a0.PutValue(dataTable.Rows[j - 3]["工程地點(址)"]);

                    for (int i = 2; i < COLUMN.Length + 2; i++)
                    {

                        var ABC = Report_Source.NumberToColumnLetter(i);
                        var a1 = ws1.Cells[ABC + j];
                        Aspose.Cells.Style a1_Style = a1.GetStyle();
                        a1_Style.Number = 5;
                        var title = COLUMN[i - 2];

                        if (title == "序號")
                        {
                            a1.PutValue(j - 2);
                        }
                        else if (title == "補助對象")
                        {
                            a1.PutValue(dataTable.Rows[j - 3]["工程地點(址)"] + "政府");
                        }
                        else if (title == "落後原因")
                        {
                            a1.PutValue(dataTable.Rows[j - 3]["落後原因"]);
                        }
                        else if (title == "總工程費(元)")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["總工程費(元)"]);
                            a1.PutValue(value);
                            a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                            a1.SetStyle(a1_Style);
                        }
                        else if (title == "本署補助經費(中央款)(元)")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["本署補助經費(中央款)(元)"]);
                            a1.PutValue(value);
                            a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                            a1.SetStyle(a1_Style);
                        }
                        else if (title == "工程費比例")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["工程費比例"]);
                            a1.PutValue(value);
                            a1_Style.Custom = "0.00_);[Red]\\(0.00\\)";
                            a1.SetStyle(a1_Style);
                        }
                        else if (title == "當年度核定數")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["當年度核定數"]);
                            a1.PutValue(value);
                            a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                            a1.SetStyle(a1_Style);
                        }
                        else if (title == "當年度撥款數")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["當年度撥款數"]);
                            a1.PutValue(value);
                            a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                            a1.SetStyle(a1_Style);
                        }
                        else if (title == "累積撥款數")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["累積撥款數"]);
                            a1.PutValue(value);
                            a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                            a1.SetStyle(a1_Style);
                        }
                        else
                        {
                            a1.PutValue(dataTable.Rows[j - 3][title].ToString());
                        }
                    }
                    j++;
                }

                /// 從這裡開始增加一個加總欄位
                int bottom_count = rowcount + 1 + 2;
                for (int t = 0; t < COLUMN.Length; t++)
                {
                    var ABC = Report_Source.NumberToColumnLetter(t + 2);
                    ws1.Cells["A" + bottom_count.ToString()].PutValue("合計");
                    var b = ws1.Cells[ABC + bottom_count];
                    Aspose.Cells.Style b_Style = b.GetStyle();
                    decimal sum = 0;
                    if (COLUMN[t] == "總工程費(元)")
                    {

                        foreach (DataRow row in dataTable.Rows)
                        {
                            if (row["總工程費(元)"] != DBNull.Value) /// 確認值不為 DBNull
                            {
                                sum += Source.Tools.ToDecimal(row["總工程費(元)"]);
                            }
                        }

                        b.Value = sum;
                        b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        b.SetStyle(b_Style);
                    }
                    else if (COLUMN[t] == "本署補助經費(中央款)(元)")
                    {
                        foreach (DataRow row in dataTable.Rows)
                        {
                            if (row["本署補助經費(中央款)(元)"] != DBNull.Value) /// 確認值不為 DBNull
                            {
                                sum += Source.Tools.ToDecimal(row["本署補助經費(中央款)(元)"]);
                            }
                        }

                        b.Value = sum;
                        b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        b.SetStyle(b_Style);
                    }
                    else if (COLUMN[t] == "當年度核定數")
                    {
                        foreach (DataRow row in dataTable.Rows)
                        {
                            if (row["當年度核定數"] != DBNull.Value) /// 確認值不為 DBNull
                            {
                                sum += Source.Tools.ToDecimal(row["當年度核定數"]);
                            }
                        }

                        b.Value = sum;
                        b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        b.SetStyle(b_Style);
                    }
                    else if (COLUMN[t] == "當年度撥款數")
                    {
                        foreach (DataRow row in dataTable.Rows)
                        {
                            if (row["當年度撥款數"] != DBNull.Value) // 確認值不為 DBNull
                            {
                                sum += Source.Tools.ToDecimal(row["當年度撥款數"]);
                            }
                        }

                        b.Value = sum;
                        b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        b.SetStyle(b_Style);
                    }
                    else if (COLUMN[t] == "累積撥款數")
                    {
                        foreach (DataRow row in dataTable.Rows)
                        {
                            if (row["累積撥款數"] != DBNull.Value) /// 確認值不為 DBNull
                            {
                                sum += Source.Tools.ToDecimal(row["累積撥款數"]);
                            }
                        }

                        b.Value = sum;
                        b_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        b.SetStyle(b_Style);
                    }
                    else
                    {

                    }
                }

            }

            var mergedRange = ws1.Cells.CreateRange("A1", Column_length + "1");
            mergedRange.Merge();
            mergedRange[0, 0].PutValue("工程細項執行表");

            Aspose.Cells.Style style = workbook.DefaultStyle;
            style.Font.Name = "Microsoft JhengHei Light";
            workbook.DefaultStyle = style;
            ws1.AutoFitColumns();
            ws1.AutoFitRows();

            int maxRow = ws1.Cells.Rows.Count;
            int maxColumn = ws1.Cells.Columns.Count;

            for (int col = 0; col <= maxColumn; col++)
            {
                double columnWidth = ws1.Cells.GetColumnWidth(col);
                if (columnWidth > 35)
                {
                    ws1.Cells.SetColumnWidth(col, 35);
                }
            }

            for (int i = 0; i < COLUMN.Length; i++)
            {
                var Excel_column = Report_Source.NumberToColumnLetter(i + 1);
                var data_column = dataTable.Columns[i].ColumnName;
                if (data_column == "落後原因")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 35);
                }
            }

            string filename = "E-1.工程細項執行表(縣市在列)" + DateTime.Now.ToString("yyyyMMddHHmmss");

            if (Field == "on")
            {
                filename = "E-1.工程細項執行表(縣市在欄)" + DateTime.Now.ToString("yyyyMMddHHmmss");
            }

            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");
            workbook.Save(savepath, SaveFormat.Xlsx);

            if (PDF == "1")
            {
                var odsSavePath = Server.MapPath("~/Temp/" + filename + ".ods");
                workbook.Save(odsSavePath, SaveFormat.ODS);

                using (var stream = new FileStream(odsSavePath, FileMode.Open))
                {
                    var bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, (int)stream.Length);
                    return File(bytes, "application/vnd.oasis.opendocument.spreadsheet", filename + ".ods");
                }
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Seek(0, SeekOrigin.Begin);
                    var bytes = stream.ToArray();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename + ".xlsx");
                }
            }
        }

        /// <summary>
        /// F-1 相關計畫年度輔助經費表
        /// </summary>
        /// <param name="CITY"></param>
        /// <param name="CITY_TOWN"></param>
        /// <param name="year1"></param>
        /// <param name="year2"></param>
        /// <param name="Row"></param>
        /// <param name="Field"></param>
        /// <param name="master_plan"></param>
        /// <param name="plan"></param>
        /// <param name="ods"></param>
        /// <see cref="PJMFORM"/> 查詢經費相關數據的主資料表
        /// <see cref="TDDSTCF"/> 根據城市和行政區條件
        /// <see cref="TDV_TDDCFTS"/> 涉及細項分類(FD_CODE、CATY_DESC、ACC_DESC)，查詢不同經費類別或細項
        /// <returns></returns>
        public ActionResult ReportPrint10(string[] CITY, string[] CITY_TOWN, string year1, string year2, string Row, string Field, List<decimal> master_plan, List<decimal> plan, string ods)
        {
            int A = int.Parse(year1);
            int B = int.Parse(year2) + 1;
            List<string> Year_Range = new List<string>();
            for (int num = A; num < B; num++)
            {
                string year = num.ToString("D3");
                Year_Range.Add(year);
            }
            var Year_Range_Length = Year_Range.Count();

            if (master_plan is null)
            {
                master_plan = new List<decimal>();
            }

            if (plan is null)
            {
                plan = new List<decimal>();
            }

            //未勾選縣市／鄉鎮時表單不會送出對應欄位，若不防呆後續 CITY_TOWN.Count() 會拋 ArgumentNullException
            if (CITY is null)
            {
                CITY = new string[0];
            }

            if (CITY_TOWN is null)
            {
                CITY_TOWN = new string[0];
            }

            //取得主計畫所對該的分支計畫資料
            foreach (var PID in master_plan)
            {
                var branch_LIST = db2.PJMFORM.Where(m => m.AN_PID == PID).DistinctBy(m => m.PJ_NAME).ToList();
                foreach (var item in branch_LIST)
                {
                    plan.Add(item.PID);
                    Debug.Write($"PID:{item.PID}");
                }
            }

            string Plan_Name = "";
            var Source_Merge = new List<Report10ViewModel>();
            var PJ_CITY = "";
            var PJ_TOWN = "";

            var dataTable = new DataTable();
            var row2 = dataTable.NewRow();
            decimal total_sum = 0;
            decimal total_sum2 = 0;

            /// 實支數
            if (Row == "on")
            {
                dataTable.Columns.Add("計畫名稱");
                dataTable.Columns.Add("縣市");
                dataTable.Columns.Add("行政區");
                foreach (var columnName in Year_Range)
                {
                    string YEAR_COLUMN = columnName.ToString();
                    dataTable.Columns.Add(YEAR_COLUMN + "核定數");
                    dataTable.Columns.Add(YEAR_COLUMN + "實支數");
                }
                dataTable.Columns.Add("小計核定數");
                dataTable.Columns.Add("小計實支數");
                dataTable.Columns.Add("備註");

                foreach (var PID in plan)
                {
                    //取得計畫名稱=主計畫名稱+分支計畫名稱
                    var pInfo = db2.PJMFORM
                        .Join(db2.PJMFORM, p => p.AN_PID, pc => pc.PID, (p, pc) => new { P = p, PC = pc })
                        .Where(m => m.P.PID == PID)
                        .Select(s => new
                        {
                            AN_PJ_Name = s.PC.PJ_NAME,
                            PJ_Name = s.P.PJ_NAME,
                            Plan_Name = s.PC.PJ_NAME + "-" + s.P.PJ_NAME
                        })
                        .FirstOrDefault();
                    if (pInfo != null)
                    {
                        Plan_Name = pInfo.Plan_Name;
                    }

                    //取得該分支計畫名稱所有符合條件的資料
                    Source_Merge = (from a in db2.TDV_TDDCFTS
                                    join p in db2.PJMFORM on a.PID equals p.PID into PJF
                                    from p in PJF.DefaultIfEmpty()
                                    join mp in db2.PJMFORM on p.AN_PID equals mp.PID into MPJF
                                    from mp in MPJF.DefaultIfEmpty()
                                    join latest_b in (
                                                     from bb in db2.TDDSTCF
                                                     group bb by bb.FD_CODE into g
                                                     select new
                                                     {
                                                         FD = g.Key,
                                                         P_FEE = g.Where(s => s.ID == g.Max(m => m.ID)).Select(s => s.PPAY_FEE).FirstOrDefault(),
                                                         C_FEE = g.Where(s => s.ID == g.Max(m => m.ID)).Select(s => s.COST_FEE).FirstOrDefault()
                                                     }
                                                 ) on a.FD_CODE equals latest_b.FD into bGroup
                                    from b in bGroup.DefaultIfEmpty()
                                    where a.PID_SYR.CompareTo(year1) >= 0 && a.PID_SYR.CompareTo(year2) <= 0
                                    where mp.PJ_NAME.Contains(pInfo.AN_PJ_Name)
                                    where p.PJ_NAME.Contains(pInfo.PJ_Name)
                                    select new Report10ViewModel
                                    {
                                        PID = a.PID.Value,
                                        SYR = a.PID_SYR,
                                        TD_CODE = a.TD_CODE,
                                        TD_NAME = a.TD_NAME,
                                        CITY = a.CITY,
                                        TOWN = a.TOWN,
                                        FUND = (int)a.CEN_FUND,
                                        COST_FEE = (b.C_FEE.HasValue ? (int)b.C_FEE.Value : 0) + (b.P_FEE.HasValue ? (int)b.P_FEE.Value : 0)
                                    }).ToList();

                    for (int i = 0; i < CITY_TOWN.Count(); i++)
                    {
                        string[] sArray;
                        sArray = CITY_TOWN[i].Split('+');
                        PJ_CITY = sArray[0];
                        PJ_TOWN = sArray[1];

                        /// TOWN_TOTAL2 => 實支數 / TOWN_TOTAL => 核定數
                        var TOWN_SOURCE = Source_Merge.Where(m => m.CITY == PJ_CITY && m.TOWN == PJ_TOWN);
                        var TOWN_TOTAL = TOWN_SOURCE.GroupBy(m => m.SYR).Select(g => new Report10ViewModel { SYR = g.Key, FUND = g.Sum(m => m.FUND) });
                        var TOWN_TOTAL2 = TOWN_SOURCE.GroupBy(m => m.SYR).Select(g => new Report10ViewModel { SYR = g.Key, COST_FEE = g.Sum(m => m.COST_FEE) });

                        var row = dataTable.NewRow();
                        row["計畫名稱"] = Plan_Name;
                        row["縣市"] = PJ_CITY;
                        row["行政區"] = PJ_TOWN;

                        decimal Total = 0;
                        decimal Total2 = 0;

                        /// 依據年度範圍填入"核定數"/"實支數"
                        foreach (var year in Year_Range)
                        {
                            decimal Year_Fund = 0;
                            decimal Year_Fund2 = 0;

                            /// 核定數總計 / 實支數總計
                            var yearData = TOWN_TOTAL.FirstOrDefault(m => m.SYR == year.ToString());
                            var yearData2 = TOWN_TOTAL2.FirstOrDefault(m => m.SYR == year.ToString());

                            if (yearData != null)
                                Year_Fund = yearData.FUND;

                            if (yearData2 != null)
                                Year_Fund2 = yearData2.COST_FEE;

                            row[year + "核定數"] = Year_Fund;
                            row[year + "實支數"] = Year_Fund2;
                            Total += Year_Fund;
                            Total2 += Year_Fund2;
                        }
                        row["小計核定數"] = Total;
                        row["小計實支數"] = Total2;
                        dataTable.Rows.Add(row);
                    }
                }

                /// 加總後欄位 (合計"核定數"與合計"實支數")
                row2 = dataTable.NewRow();
                total_sum = 0;
                total_sum2 = 0;

                foreach (var year in Year_Range)
                {
                    decimal sum = 0;
                    decimal sum2 = 0;

                    foreach (DataRow rowFund in dataTable.Rows)
                    {
                        sum += Convert.ToDecimal(rowFund[year + "核定數"]);
                        sum2 += Convert.ToDecimal(rowFund[year + "實支數"]);
                    }
                    row2[year + "核定數"] = sum;
                    row2[year + "實支數"] = sum2;
                    total_sum += sum;
                    total_sum2 += sum2;
                }
                row2["小計核定數"] = total_sum;
                row2["小計實支數"] = total_sum2;
                dataTable.Rows.Add(row2);
            }
            /// 無實支數
            else
            {
                dataTable.Columns.Add("計畫名稱");
                dataTable.Columns.Add("縣市");
                dataTable.Columns.Add("行政區");
                foreach (var columnName in Year_Range)
                {
                    string YEAR_COLUMN = columnName.ToString();
                    dataTable.Columns.Add(YEAR_COLUMN);
                }
                dataTable.Columns.Add("小計");
                dataTable.Columns.Add("備註");

                foreach (var PID in plan)
                {
                    //取得計畫名稱=主計畫名稱+分支計畫名稱
                    var pInfo = db2.PJMFORM
                        .Join(db2.PJMFORM, p => p.AN_PID, pc => pc.PID, (p, pc) => new { P = p, PC = pc })
                        .Where(m => m.P.PID == PID)
                        .Select(s => new
                        {
                            AN_PJ_Name = s.PC.PJ_NAME,
                            PJ_Name = s.P.PJ_NAME,
                            Plan_Name = s.PC.PJ_NAME + "-" + s.P.PJ_NAME
                        })
                        .FirstOrDefault();
                    if (pInfo != null)
                    {
                        Plan_Name = pInfo.Plan_Name;
                    }

                    //取得該分支計畫名稱所有符合條件的資料
                    Source_Merge = (from a in db2.TDV_TDDCFTS
                                    join p in db2.PJMFORM on a.PID equals p.PID into PJF
                                    from p in PJF.DefaultIfEmpty()
                                    join mp in db2.PJMFORM on p.AN_PID equals mp.PID into MPJF
                                    from mp in MPJF.DefaultIfEmpty()
                                    where a.PID_SYR.CompareTo(year1) >= 0 && a.PID_SYR.CompareTo(year2) <= 0
                                    where mp.PJ_NAME.Contains(pInfo.AN_PJ_Name)
                                    where p.PJ_NAME.Contains(pInfo.PJ_Name)
                                    select new Report10ViewModel
                                    {
                                        PID = a.PID.Value,
                                        SYR = a.PID_SYR,
                                        TD_CODE = a.TD_CODE,
                                        TD_NAME = a.TD_NAME,
                                        CITY = a.CITY,
                                        TOWN = a.TOWN,
                                        FUND = (int)a.CEN_FUND
                                    }).ToList();

                    for (int i = 0; i < CITY_TOWN.Count(); i++)
                    {
                        string[] sArray;
                        sArray = CITY_TOWN[i].Split('+');
                        PJ_CITY = sArray[0];
                        PJ_TOWN = sArray[1];
                        var TOWN_SOURCE = Source_Merge.Where(m => m.CITY == PJ_CITY && m.TOWN == PJ_TOWN);
                        var TOWN_TOTAL = TOWN_SOURCE.GroupBy(m => m.SYR).Select(g => new Report10ViewModel { SYR = g.Key, FUND = g.Sum(m => m.FUND) });   /// 這裡有 CITY / TOWN / FUND 上面有 PJ_NAME

                        var row = dataTable.NewRow();
                        row["計畫名稱"] = Plan_Name;
                        row["縣市"] = PJ_CITY;
                        row["行政區"] = PJ_TOWN;

                        decimal Total = 0;
                        foreach (var year in Year_Range)
                        {
                            decimal Year_Fund = 0;
                            var yearData = TOWN_TOTAL.FirstOrDefault(m => m.SYR == year.ToString());
                            if (yearData != null)
                            {
                                Year_Fund = yearData.FUND;
                            }
                            row[year] = Year_Fund;
                            Total += Year_Fund;
                        }
                        row["小計"] = Total;
                        dataTable.Rows.Add(row);
                    }
                }

                /// 加總後欄位
                row2 = dataTable.NewRow();
                total_sum = 0;

                foreach (var year in Year_Range)
                {
                    decimal sum = 0;

                    foreach (DataRow rowFund in dataTable.Rows)
                    {
                        sum += Convert.ToDecimal(rowFund[year]);
                    }
                    row2[year] = sum;
                    total_sum += sum;
                }
                row2["小計"] = total_sum;
                dataTable.Rows.Add(row2);
            }

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);

            var path = Server.MapPath("~/Templates/表2.xlsx");
            var file = new FileInfo(path);
            var workbook = new Aspose.Cells.Workbook(path);
            Worksheet ws1 = workbook.Worksheets[0];
            ws1 = CopyWorksheet(workbook, ws1, "總表1", "相關計畫年度補助經費表");

            Aspose.Cells.Style numStyle = workbook.CreateStyle();
            numStyle.Number = 1;
            numStyle.Custom = "#,##0";
            numStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            numStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            numStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            numStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            numStyle.HorizontalAlignment = TextAlignmentType.Right;

            var Column_length = Report_Source.NumberToColumnLetter(Year_Range_Length + 5);
            var SOURCE = Source_Merge;

            if (Row == "on")
            {
                ws1.Name = "相關計畫年度執行表";
                Column_length = Report_Source.NumberToColumnLetter(Year_Range_Length * 2 + 6);

                /// 實支數 / 核定數
                for (int i = 4, j = 0; i < Year_Range_Length * 2 + 4; i = i + 2, j++)
                {
                    var ABC = Report_Source.NumberToColumnLetter(i);
                    var DEF = Report_Source.NumberToColumnLetter(i + 1);
                    var a = ws1.Cells.CreateRange(ABC + "4:" + DEF + "4");
                    a.Merge();

                    var title = Year_Range[i - 4 - j];
                    a.Value = title + "年度";
                    var w = ws1.Cells[ABC + "5"];
                    w.Value = "核定數";
                    var w2 = ws1.Cells[DEF + "5"];
                    w2.Value = "實支數";
                }

                var title2 = Report_Source.NumberToColumnLetter(Year_Range_Length * 2 + 4);
                var title2_2 = Report_Source.NumberToColumnLetter(Year_Range_Length * 2 + 5);

                var b = ws1.Cells.CreateRange("D3:" + title2 + "3");
                b.Merge();
                b.Value = "補助經費";

                var d = ws1.Cells.CreateRange(title2 + "4:" + title2_2 + "4");
                d.Merge();
                d.Value = "小計";

                var d1 = ws1.Cells[title2 + "5"];
                d1.Value = "核定數";

                var d2 = ws1.Cells[title2_2 + "5"];
                d2.Value = "實支數";

                var Note = Report_Source.NumberToColumnLetter(Year_Range_Length * 2 + 6);
                var e = ws1.Cells.CreateRange(Note + "3" + ":" + Note + "4");
                e.Merge();
                e.Value = "備註";

                var f = ws1.Cells[Note + "2"];
                f.Value = "單位：千元";

                var g = ws1.Cells.CreateRange("A3" + ":" + "A5");
                g.Merge();
                g.Value = "計畫名稱";

                var h = ws1.Cells.CreateRange("B3" + ":" + "B5");
                h.Value = "縣市";
                h.Merge();

                var k = ws1.Cells.CreateRange("C3" + ":" + "C5");
                k.Value = "行政區";
                k.Merge();

                int z = 6;
                int y = 0;
                int x = 0;

                for (int w = 0; w < plan.Count; w++)
                {
                    foreach (var item in CITY_TOWN)
                    {
                        x = 0;
                        var PJ_BLOCK = ws1.Cells["A" + z];
                        PJ_BLOCK.Value = dataTable.Rows[y]["計畫名稱"];
                        var CITY_BLOCK = ws1.Cells["B" + z];
                        CITY_BLOCK.Value = dataTable.Rows[y]["縣市"];
                        var TOWN_BLOCK = ws1.Cells["C" + z];
                        TOWN_BLOCK.Value = dataTable.Rows[y]["行政區"];

                        foreach (var item2 in Year_Range) /// 篩年度核定數實支數
                        {
                            var year_column = Report_Source.NumberToColumnLetter(x + 4);
                            var year_column2 = Report_Source.NumberToColumnLetter((x + 1) + 4);

                            var YEAR_BLOCK = ws1.Cells[year_column + z];
                            var YEAR_BLOCK2 = ws1.Cells[year_column2 + z];

                            /// YEAR_BLOCK.Value = int.Parse(dataTable.Rows[y][item2 + "核定數"].ToString()).ToString("#,##0");
                            YEAR_BLOCK.Value = Convert.ToDouble(dataTable.Rows[y][item2 + "核定數"].ToString()) / 1000;

                            /// YEAR_BLOCK2.Value = int.Parse(dataTable.Rows[y][item2 + "實支數"].ToString()).ToString("#,##0");
                            YEAR_BLOCK2.Value = Convert.ToDouble(dataTable.Rows[y][item2 + "實支數"].ToString()) / 1000;

                            YEAR_BLOCK.SetStyle(numStyle);
                            YEAR_BLOCK2.SetStyle(numStyle);

                            x = x + 2;
                        }

                        var ABC = Report_Source.NumberToColumnLetter(Year_Range.Count() * 2 + 4);
                        var DEF = Report_Source.NumberToColumnLetter(Year_Range.Count() * 2 + 5);

                        var TOTAL_BLOCK = ws1.Cells[ABC + z];
                        var TOTAL_BLOCK2 = ws1.Cells[DEF + z];

                        //TOTAL_BLOCK.Value = int.Parse(dataTable.Rows[y]["小計核定數"].ToString()).ToString("#,##0");
                        //TOTAL_BLOCK2.Value = int.Parse(dataTable.Rows[y]["小計實支數"].ToString()).ToString("#,##0");

                        TOTAL_BLOCK.Value = Convert.ToDouble(dataTable.Rows[y]["小計核定數"].ToString()) / 1000;
                        TOTAL_BLOCK2.Value = Convert.ToDouble(dataTable.Rows[y]["小計實支數"].ToString()) / 1000;

                        TOTAL_BLOCK.SetStyle(numStyle);
                        TOTAL_BLOCK2.SetStyle(numStyle);

                        z++;
                        y++;
                    }
                }

                var l = ws1.Cells.CreateRange("A" + (CITY_TOWN.Count() * plan.Count + 6) + ":C" + (CITY_TOWN.Count() * plan.Count + 6));
                l.Merge();
                l.Value = "合計";
                x = 0;

                foreach (var item in Year_Range)
                {
                    var year_column = Report_Source.NumberToColumnLetter(x + 4);
                    var year_column2 = Report_Source.NumberToColumnLetter((x + 1) + 4);
                    var sum = ws1.Cells[year_column + (CITY_TOWN.Count() * plan.Count + 6)];
                    var sum2 = ws1.Cells[year_column2 + (CITY_TOWN.Count() * plan.Count + 6)];
                    //sum.Value = int.Parse(dataTable.Rows[CITY_TOWN.Count() * plan.Count][item + "核定數"].ToString()).ToString("#,##0");
                    //sum2.Value = int.Parse(dataTable.Rows[CITY_TOWN.Count() * plan.Count][item + "實支數"].ToString()).ToString("#,##0");
                    sum.Value = Convert.ToDouble(dataTable.Rows[CITY_TOWN.Count() * plan.Count][item + "核定數"].ToString()) / 1000;
                    sum2.Value = Convert.ToDouble(dataTable.Rows[CITY_TOWN.Count() * plan.Count][item + "實支數"].ToString()) / 1000;
                    sum.SetStyle(numStyle);
                    sum2.SetStyle(numStyle);
                    x = x + 2;
                }
                var SUM_COLUMN = Report_Source.NumberToColumnLetter(Year_Range.Count() * 2 + 4);
                var SUM_COLUMN2 = Report_Source.NumberToColumnLetter(Year_Range.Count() * 2 + 5);
                var SUM_TOTAL_BLOCK = ws1.Cells[SUM_COLUMN + (CITY_TOWN.Count() * plan.Count + 6)];
                var SUM_TOTAL_BLOCK2 = ws1.Cells[SUM_COLUMN2 + (CITY_TOWN.Count() * plan.Count + 6)];
                //SUM_TOTAL_BLOCK.Value = int.Parse(dataTable.Rows[CITY_TOWN.Count() * plan.Count]["小計核定數"].ToString()).ToString("#,##0");
                //SUM_TOTAL_BLOCK2.Value = int.Parse(dataTable.Rows[CITY_TOWN.Count() * plan.Count]["小計實支數"].ToString()).ToString("#,##0");
                SUM_TOTAL_BLOCK.Value = Convert.ToDouble(dataTable.Rows[CITY_TOWN.Count() * plan.Count]["小計核定數"].ToString()) / 1000;
                SUM_TOTAL_BLOCK2.Value = Convert.ToDouble(dataTable.Rows[CITY_TOWN.Count() * plan.Count]["小計實支數"].ToString()) / 1000;
                SUM_TOTAL_BLOCK.SetStyle(numStyle);
                SUM_TOTAL_BLOCK2.SetStyle(numStyle);
            }
            else
            {
                for (int i = 4; i < Year_Range_Length + 4; i++)
                {
                    var ABC = Report_Source.NumberToColumnLetter(i);
                    var a = ws1.Cells[ABC + "4"];
                    var title = Year_Range[i - 4];
                    a.Value = title + "年度";
                }
                var title2 = Report_Source.NumberToColumnLetter(Year_Range_Length + 4);
                var b = ws1.Cells.CreateRange("D3:" + title2 + "3");
                b.Merge();
                b.Value = "補助經費";
                var d = ws1.Cells[title2 + "4"];
                d.Value = "小計";
                var Note = Report_Source.NumberToColumnLetter(Year_Range_Length + 5);
                var e = ws1.Cells.CreateRange(Note + "3" + ":" + Note + "4");
                e.Merge();
                e.Value = "備註";
                var f = ws1.Cells[Note + "2"];
                f.Value = "單位：千元";
                var g = ws1.Cells.CreateRange("A3" + ":" + "A4");
                g.Merge();
                g.Value = "計畫名稱";
                var h = ws1.Cells.CreateRange("B3" + ":" + "B4");
                h.Value = "縣市";
                h.Merge();
                var k = ws1.Cells.CreateRange("C3" + ":" + "C4");
                k.Value = "行政區";
                k.Merge();

                int z = 5;
                int y = 0;
                int x = 0;

                for (int w = 0; w < plan.Count; w++)
                {
                    foreach (var item in CITY_TOWN)
                    {
                        x = 0;
                        var PJ_BLOCK = ws1.Cells["A" + z];
                        PJ_BLOCK.Value = dataTable.Rows[y]["計畫名稱"];
                        var CITY_BLOCK = ws1.Cells["B" + z];
                        CITY_BLOCK.Value = dataTable.Rows[y]["縣市"];
                        var TOWN_BLOCK = ws1.Cells["C" + z];
                        TOWN_BLOCK.Value = dataTable.Rows[y]["行政區"];
                        foreach (var item2 in Year_Range)
                        {
                            var year_column = Report_Source.NumberToColumnLetter(x + 4);
                            var YEAR_BLOCK = ws1.Cells[year_column + z];
                            YEAR_BLOCK.Value = Convert.ToDouble(dataTable.Rows[y][item2].ToString()) / 1000;
                            YEAR_BLOCK.SetStyle(numStyle);
                            x++;
                        }
                        var ABC = Report_Source.NumberToColumnLetter(Year_Range.Count() + 4);
                        var TOTAL_BLOCK = ws1.Cells[ABC + z];
                        TOTAL_BLOCK.Value = Convert.ToDouble(dataTable.Rows[y]["小計"].ToString()) / 1000;
                        TOTAL_BLOCK.SetStyle(numStyle);
                        z++;
                        y++;
                    }
                }

                var l = ws1.Cells.CreateRange("A" + (CITY_TOWN.Count() * plan.Count + 5) + ":C" + (CITY_TOWN.Count() * plan.Count + 5));
                l.Merge();
                l.Value = "合計";
                x = 0;

                foreach (var item in Year_Range)
                {
                    var year_column = Report_Source.NumberToColumnLetter(x + 4);
                    var sum = ws1.Cells[year_column + (CITY_TOWN.Count() * plan.Count + 5)];
                    sum.Value = Convert.ToDouble(dataTable.Rows[CITY_TOWN.Count() * plan.Count][item].ToString()) / 1000;
                    sum.SetStyle(numStyle);
                    x++;
                }
                var SUM_COLUMN = Report_Source.NumberToColumnLetter(Year_Range.Count() + 4);
                var SUM_TOTAL_BLOCK = ws1.Cells[SUM_COLUMN + (CITY_TOWN.Count() * plan.Count + 5)];
                SUM_TOTAL_BLOCK.Value = Convert.ToDouble(dataTable.Rows[CITY_TOWN.Count() * plan.Count]["小計"].ToString()) / 1000;
                SUM_TOTAL_BLOCK.SetStyle(numStyle);
            }

            var Title_length = Report_Source.NumberToColumnLetter(Year_Range_Length + 5);

            if (Row == "on")
            {
                Title_length = Report_Source.NumberToColumnLetter(Year_Range_Length * 2 + 6);
            }
            var mergedCells = ws1.Cells.CreateRange("A1:" + Title_length + "1");
            mergedCells.Merge();
            mergedCells.Value = "相關計畫年度補助經費表";

            if (Row == "on")
            {
                mergedCells.Value = "相關計畫年度執行表";
            }
            Aspose.Cells.Style style = workbook.CreateStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            mergedCells.SetStyle(style);

            /// 20241021 => 修正加入格線位置
            Aspose.Cells.Range c = ws1.Cells.CreateRange("A1:" + Column_length + (CITY_TOWN.Count() * plan.Count + 5).ToString());

            if (Row == "on")
            {
                c = ws1.Cells.CreateRange("A1:" + Column_length + (CITY_TOWN.Count() * plan.Count + 6).ToString());
            }

            Aspose.Cells.Style borderstyle = ws1.Workbook.CreateStyle();
            Aspose.Cells.Border outerBorder = borderstyle.Borders[BorderType.TopBorder];
            outerBorder.LineStyle = CellBorderType.Thin;
            outerBorder = borderstyle.Borders[BorderType.BottomBorder];
            outerBorder.LineStyle = CellBorderType.Thin;
            outerBorder = borderstyle.Borders[BorderType.LeftBorder];
            outerBorder.LineStyle = CellBorderType.Thin;
            outerBorder = borderstyle.Borders[BorderType.RightBorder];
            outerBorder.LineStyle = CellBorderType.Thin;
            Aspose.Cells.Border innerBorder = borderstyle.Borders[BorderType.Vertical];
            innerBorder.LineStyle = CellBorderType.Thin;
            innerBorder = borderstyle.Borders[BorderType.Horizontal];
            innerBorder.LineStyle = CellBorderType.Thin;
            borderstyle.HorizontalAlignment = TextAlignmentType.Center;
            borderstyle.Custom = "#,##0.00_);[Red]\\(0.00\\)";
            c.SetStyle(borderstyle);

            string filename = "F-1.相關計畫年度補助經費表" + DateTime.Now.ToString("yyyyMMddHHmmss");

            if (Row == "on")
            {
                filename = "F-1.相關計畫年度執行表" + DateTime.Now.ToString("yyyyMMddHHmmss");
            }

            int endRow = dataTable.Rows.Count;
            int endCol = ws1.Cells.MaxColumn;
            var range = ws1.Cells.CreateRange(5, 3, endRow, endCol - 3);
            range.SetStyle(numStyle);
            workbook.Worksheets.RemoveAt(0);

            /// 設定欄寬
            /// A欄 計畫名稱
            ws1.Cells.SetColumnWidth(0, 55);

            /// B欄 縣市
            ws1.Cells.SetColumnWidth(1, 12);

            /// C欄 行政區
            ws1.Cells.SetColumnWidth(2, 12);

            /// 從 D 欄開始是年度欄
            int startCol = 3;   /// D = index 3
            int totalCols = ws1.Cells.MaxColumn;

            for (int i = startCol; i <= totalCols; i++)
            {
                ws1.Cells.SetColumnWidth(i, 15);
            }

            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");
            var afile = new FileInfo(savepath);
            string directoryPath = Server.MapPath(@"~\Temp\");
            string newfilepath = System.IO.Path.Combine(directoryPath, filename + ".xlsx");
            workbook.Save(newfilepath);

            if (ods == "1")
            {
                var OdsResult = OdsPrint(newfilepath, filename);
                return OdsResult;
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Seek(0, SeekOrigin.Begin);
                    var bytes = stream.ToArray();
                    stream.Close();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename + ".xlsx");
                }
            }
        }

        /// <summary>
        /// H-1.在建工程抽查 / 查核 / 督導統計表
        /// </summary>
        /// <param name="SYR"></param>
        /// <param name="PID"></param>
        /// <param name="DATE_START"></param>
        /// <param name="DATE_END"></param>
        /// <returns></returns>
        public ActionResult ReportPrint12(string SYR, string PID, DateTime DATE_START, DateTime DATE_END)
        {
            string sFullFile = Server.MapPath("~/Templates/PJV0010_09.xlsx");

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);

            int i;
            int j;
            int sheet1line;
            int sheet2line;
            DataRow[] rr;
            DataTable TDMFORM = new DataTable();
            DataTable TDDSPRS = new DataTable();
            string DATE_START1 = ReportController.ConvertToRepublicOfChinaDate(DATE_START);
            string DATE_END1 = ReportController.ConvertToRepublicOfChinaDate(DATE_END);

            Report_Source.getProject09(ref TDMFORM, ref TDDSPRS, PID, DATE_START1, DATE_END1, connectionString);

            var workbook = new Aspose.Cells.Workbook(sFullFile);
            Aspose.Cells.Worksheet sheet1 = workbook.Worksheets["Sheet1"];
            Aspose.Cells.Worksheet sheet2 = workbook.Worksheets["Sheet2"];
            Aspose.Cells.Worksheet sheet3 = workbook.Worksheets["總表"];
            Aspose.Cells.Worksheet sheet4 = workbook.Worksheets["詳細表"];

            sheet1line = 4;
            sheet2line = 1;

            Aspose.Cells.Range sourceRange = sheet3.Cells.CreateRange(0, 0, 3, 6);
            Aspose.Cells.Range destinationRange = sheet1.Cells.CreateRange(0, 0, 3, 6);
            destinationRange.Copy(sourceRange);

            sheet1.Cells.SetRowHeight(0, sheet3.Cells.GetRowHeight(0));
            sheet1.Cells.SetRowHeight(1, sheet3.Cells.GetRowHeight(1));
            sheet1.Cells.SetRowHeight(2, sheet3.Cells.GetRowHeight(2));

            for (i = 0; i < TDMFORM.Rows.Count; i++)
            {
                sheet1.Cells.CreateRange("A" + sheet1line).Copy(sheet3.Cells.CreateRange("A9:F9"));
                sheet1.Cells.SetRowHeight(sheet1line - 1, sheet3.Cells.GetRowHeight(3));

                sheet1.Cells["A" + sheet1line.ToString()].PutValue((i + 1).ToString());
                sheet1.Cells["B" + sheet1line.ToString()].PutValue(DATE_START1 + "~" + DATE_END1);
                sheet1.Cells["C" + sheet1line.ToString()].PutValue((TDMFORM.Rows[i]["TD_NAME"]).ToString());
                sheet1.Cells["D" + sheet1line.ToString()].PutValue((TDMFORM.Rows[i]["SP_03"]).ToString());
                sheet1.Cells["E" + sheet1line.ToString()].PutValue((TDMFORM.Rows[i]["SP_02"]).ToString());
                sheet1.Cells["F" + sheet1line.ToString()].PutValue((TDMFORM.Rows[i]["SP_01"]).ToString());
                sheet1line++;

                rr = TDDSPRS.Select("TD_CODE='" + TDMFORM.Rows[i]["TD_CODE"].ToString() + "'");

                if (rr.Length > 0)
                {
                    sheet2.Cells.CreateRange("A" + sheet2line).Copy(sheet4.Cells.CreateRange("A1:J3"));
                    sheet2.Cells["A" + sheet2line.ToString()].PutValue(TDMFORM.Rows[i]["TD_NAME"].ToString() + "抽查/查核/督導詳細表");
                    sheet2line = sheet2line + 3;
                    for (j = 0; j < rr.Length; j++)
                    {
                        sheet2.Cells.CreateRange("A" + sheet2line).Copy(sheet4.Cells.CreateRange("A4:J4"));
                        sheet2.Cells["A" + sheet2line.ToString()].PutValue((j + 1).ToString());
                        sheet2.Cells["B" + sheet2line.ToString()].PutValue(TDMFORM.Rows[i]["IPJ_DESC"].ToString());
                        sheet2.Cells["C" + sheet2line.ToString()].PutValue(rr[j]["CATEGORY_N"].ToString());// 類別
                        sheet2.Cells["D" + sheet2line.ToString()].PutValue(rr[j]["EXA_UNIT"].ToString());//單位
                        sheet2.Cells["E" + sheet2line.ToString()].PutValue(rr[j]["SP_ALL_SCORE"].ToString());//施工品質成績
                        sheet2.Cells["F" + sheet2line.ToString()].PutValue(rr[j]["SP_SCORE"].ToString());//職業安全成績
                        sheet2.Cells["G" + sheet2line.ToString()].PutValue(rr[j]["HIATUS_DESC"].ToString());//是否辦理施工品質缺失扣點
                        sheet2.Cells["H" + sheet2line.ToString()].PutValue(rr[j]["HIATUS2_DESC"].ToString());//是否辦理職業安全缺失扣點
                        sheet2.Cells["I" + sheet2line.ToString()].PutValue(rr[j]["IMP_DT"].ToString());//缺失改善期限
                        sheet2.Cells["J" + sheet2line.ToString()].PutValue(rr[j]["FIN_DT"].ToString());//已改善完成日期
                        sheet2line++;
                    }
                    sheet4.Cells.CreateRange(9, 0, 1, 10).Copy(sheet2.Cells.CreateRange(sheet2line, 0, 1, 10));
                    sheet2.Cells.SetRowHeight(sheet2line, sheet4.Cells.GetRowHeight(9));
                    sheet2line++;
                }
            }

            workbook.Worksheets.RemoveAt(sheet3.Index);
            workbook.Worksheets.RemoveAt(sheet4.Index);
            sheet1.Name = "總表";
            sheet2.Name = "詳細表";

            string sFileName = HttpUtility.UrlEncode("H-1.在建工程抽查/查核/督導統計表", System.Text.Encoding.UTF8);
            var savepath = Server.MapPath("~/Temp/" + sFileName + ".xlsx");

            workbook.Save(savepath, Aspose.Cells.SaveFormat.Xlsx);
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + sFileName + ".xlsx");
            Response.TransmitFile(savepath);
            Response.Flush();
            Response.End();

            return View();
        }

        /// <summary>
        /// G-1.雨水工程控管表
        /// </summary>
        /// <param name="plan"></param>
        /// <param name="CITY"></param>
        /// <param name="year1"></param>
        /// <param name="year2"></param>
        /// <param name="month"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult ReportPrint11(string PLAN, string CITY, string YEAR1, string YEAR2, string month, string TYPE)
        {
            int ICount = 1;
            int startRow = 3;
            int dataStartRow = startRow;

            #region 報表取樣

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);

            var path = Server.MapPath("~/Templates/表3.xlsx");

            var file = new FileInfo(path);
            var workbook = new Aspose.Cells.Workbook(path);
            var sheet = workbook.Worksheets["雨水前瞻控管表"];

            #endregion

            DataTable SourceTable = Report_Source.GetBillionData11(PLAN, CITY, YEAR1, YEAR2, month, TYPE, connectionString);

            #region G-1.雨水工程控管表[報表部分]

            /// Step 1. 年度範圍
            int y1 = int.Parse(YEAR1);
            int y2 = int.Parse(YEAR2);

            List<int> years = new List<int>();

            for (int y = y1; y <= y2; y++)
            {
                years.Add(y);
            }

            int yearColumnCount = years.Count * 2;

            /// Step 2. 插入年度欄位
            sheet.Cells.InsertColumns(8, yearColumnCount);

            /// Step 3. 年度標題
            int headerCol = 8;
            foreach (int y in years)
            {
                sheet.Cells[1, headerCol].PutValue($"{y}核定數");
                sheet.Cells[1, headerCol + 1].PutValue($"{y}撥款數");
                headerCol += 2;
            }

            int shift = yearColumnCount;

            Func<DataRow, string, double> N = (r, c) => r[c] == DBNull.Value ? 0D : Convert.ToDouble(r[c]);

            Func<DataRow, string, string> S = (r, c) => r[c] == DBNull.Value ? "" : r[c].ToString();

            Func<DataRow, int> RowYear = r =>
            {
                int yy;
                return int.TryParse(S(r, "PID_SYR"), out yy) ? yy : 0;
            };

            /// TD_CODE 工程編號分類
            var projectGroups = SourceTable.Rows.Cast<DataRow>().GroupBy(r => !string.IsNullOrWhiteSpace(S(r, "TD_CODE")) ? S(r, "TD_CODE") : S(r, "SIP_TD_CODE"));

            foreach (var g in projectGroups)
            {
                var rows = g.ToList();
                var first = rows.First();
                var latest = rows.OrderByDescending(RowYear).First();

                sheet.Cells.CopyRow(sheet.Cells, 3, startRow);

                sheet.Cells[startRow, 0].PutValue(ICount++);
                sheet.Cells[startRow, 1].PutValue(S(first, "SIP_TD_CODE"));
                sheet.Cells[startRow, 2].PutValue(S(first, "TD_CITY"));
                sheet.Cells[startRow, 3].PutValue(S(first, "TD_NAME"));
                sheet.Cells[startRow, 4].PutValue(N(first, "CEN_RATIO"));
                sheet.Cells[startRow, 5].PutValue(rows.Sum(r => N(r, "CEN_FUND")));
                sheet.Cells[startRow, 6].PutValue(rows.Max(r => N(r, "REV_FUND")));
                sheet.Cells[startRow, 7].PutValue(rows.Max(r => N(r, "OCR_FUND")));

                int col = 8;

                foreach (int y in years)
                {
                    var yearRows = rows.Where(r => RowYear(r) == y).ToList();

                    sheet.Cells[startRow, col].PutValue(yearRows.Sum(r => N(r, "CEN_FUND")));
                    sheet.Cells[startRow, col + 1].PutValue(yearRows.Sum(r => N(r, "ALL_PPAY_FEE")));

                    col += 2;
                }

                sheet.Cells[startRow, 8 + shift].PutValue("");
                sheet.Cells[startRow, 9 + shift].PutValue(latest["PRE_TK105"]);
                sheet.Cells[startRow, 10 + shift].PutValue(latest["REL_TK105"]);
                sheet.Cells[startRow, 11 + shift].PutValue(latest["PRE_TK107"]);
                sheet.Cells[startRow, 12 + shift].PutValue(latest["REL_TK107"]);
                sheet.Cells[startRow, 13 + shift].PutValue(latest["PRE_TK108"]);
                sheet.Cells[startRow, 14 + shift].PutValue(latest["REL_TK108"]);
                sheet.Cells[startRow, 15 + shift].PutValue(N(latest, "PRE_SCHD"));
                sheet.Cells[startRow, 16 + shift].PutValue(N(latest, "REL_SCHD"));
                sheet.Cells[startRow, 17 + shift].PutValue(latest["EXECT_DESC"]);

                startRow++;
            }

            Style numberStyle = workbook.CreateStyle();
            numberStyle.Custom = "#,##0";

            StyleFlag numberFlag = new StyleFlag();
            numberFlag.NumberFormat = true;

            for (int i = 0; i < yearColumnCount; i++)
            {
                sheet.Cells.ApplyColumnStyle(8 + i, numberStyle, numberFlag);
            }

            /// Step 5. 合計列 (Excel用)
            int totalRow = startRow;

            sheet.Cells.CopyRow(sheet.Cells, 3, totalRow);

            for (int col = 0; col <= 17 + shift; col++)
            {
                sheet.Cells[totalRow, col].PutValue(null);
            }

            sheet.Cells[totalRow, 3].PutValue("合計");
            sheet.Cells[totalRow, 5].Formula = $"SUM(F{dataStartRow + 1}:F{totalRow})";

            for (int i = 0; i < yearColumnCount; i++)
            {
                int col = 8 + i;

                string colName = CellsHelper.ColumnIndexToName(col);

                sheet.Cells[totalRow, col].Formula = $"SUM({colName}{dataStartRow + 1}:{colName}{totalRow})";
            }

            #endregion

            #region 報表格式調整

            /// A ~ H 固定欄位
            sheet.Cells.SetColumnWidth(0, 6);
            sheet.Cells.SetColumnWidth(1, 32);
            sheet.Cells.SetColumnWidth(2, 10);
            sheet.Cells.SetColumnWidth(3, 40);
            sheet.Cells.SetColumnWidth(4, 6);
            sheet.Cells.SetColumnWidth(5, 20);
            sheet.Cells.SetColumnWidth(6, 20);
            sheet.Cells.SetColumnWidth(7, 20);

            /// 年度欄位
            for (int i = 0; i < yearColumnCount; i++)
            {
                sheet.Cells.SetColumnWidth(8 + i, 20);
            }

            /// 後半部分固定欄位
            sheet.Cells.SetColumnWidth(8 + shift, 20);
            sheet.Cells.SetColumnWidth(9 + shift, 16);
            sheet.Cells.SetColumnWidth(10 + shift, 16);
            sheet.Cells.SetColumnWidth(11 + shift, 16);
            sheet.Cells.SetColumnWidth(12 + shift, 16);
            sheet.Cells.SetColumnWidth(13 + shift, 16);
            sheet.Cells.SetColumnWidth(14 + shift, 16);
            sheet.Cells.SetColumnWidth(15 + shift, 16);
            sheet.Cells.SetColumnWidth(16 + shift, 16);
            sheet.Cells.SetColumnWidth(17 + shift, 50);

            sheet.PageSetup.Zoom = 70;

            /// 自動換行
            Style style = workbook.CreateStyle();
            style.IsTextWrapped = true;

            StyleFlag flag = new StyleFlag();
            flag.WrapText = true;

            sheet.Cells.ApplyColumnStyle(3, style, flag);
            sheet.Cells.ApplyColumnStyle(17 + shift, style, flag);

            #endregion

            #region 檔案輸出 => EXCEL | ODS

            string Filename = "";


            Filename += DateTime.Now.ToString("yyyyMMddHHmmss");

            var savePath = Server.MapPath(@"~/Temp/");
            string newFilePath = System.IO.Path.Combine(savePath, Filename + ".xlsx");

            if (TYPE == "ODS")
            {
                Filename += ".ods";
                workbook.Save(newFilePath, SaveFormat.ODS);

                return File(System.IO.File.ReadAllBytes(newFilePath), "application/vnd.oasis.opendocument.spreadsheet", Filename);
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Seek(0, SeekOrigin.Begin);
                    var bytes = stream.ToArray();

                    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Filename + ".xlsx");
                }
            }

            #endregion
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private string GetReportTitle(string PID, string connectionString)
        {
            /// Step 1. 查詢 PJ_NAME
            string sql = $@" SELECT DISTINCT PJ_NAME FROM CMEPMS.dbo.PJMFORM WHERE PID IN ({PID}) ";

            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
            {
                conn.Open();
                da.Fill(dt);
            }

            /// Step 2. 轉成 List
            List<string> planNames = dt.AsEnumerable().Select(r => r["PJ_NAME"] == DBNull.Value ? "" : r["PJ_NAME"].ToString().Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

            if (planNames.Count == 0)
            {
                return "報表";
            }

            /// Step 3. 特殊報表判斷（最實務寫法）
            if (planNames.Any(s => s.Contains("01辦理水與安全業務費") || s.Contains("02辦理水與安全獎補助費") || s.Contains("03辦理水與安全設備投資費")))
            {
                return "前瞻基礎建設計畫-縣市管河川及區域排水整理改善計畫";
            }

            /// Step 4. 一般報表名稱
            if (planNames.Count == 1)
            {
                return $"{planNames[0]}";
            }

            return string.Join("、", planNames);
        }

        #endregion ====================================================================================================
    }
}