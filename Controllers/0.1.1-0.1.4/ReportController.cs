using Aspose.Cells;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Data.Entity.Core.EntityClient;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using WCals.Models;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEPMSModels;
using WCals.Models.CMETMSModels;
using WCals.Source;

namespace WCals.Controllers
{
    public class ReportController : Controller
    {
        CMBASEEntities db = new CMBASEEntities();
        CMEPMSEntities db2 = new CMEPMSEntities();
        CMETMSEntities db3 = new CMETMSEntities();

        // TODO： 放到正式環境時需更改連線字串為正式環境
        //private string connectionString = @"Data Source=192.168.7.14;initial catalog=CMEPMS;user id=***;password=***;" + "MultipleActiveResultSets=True;" + "initial catalog=CMETMS;";
        //private string connectionString2 = @"Data Source=192.168.7.14;initial catalog=CMBASE;user id=***;password=***;" + "MultipleActiveResultSets=True;" + "initial catalog=CMETMS;";
        //private string connectionString = @"data source=10.1.1.112\sepmis;initial catalog=cmepms_from_prod_260401;user id=***;password=***;" + "multipleactiveresultsets=true;" + "initial catalog=cmetms;";
        //private string connectionString2 = @"data source=10.1.1.112\sepmis;initial catalog = cmbase_from_prod_260401; user id=***;password=***;" + "multipleactiveresultsets=true;" + "initial catalog=cmetms;";
        /// 連線字串一律取自 EF (CMEPMSEntities) 的實際連線 —— 亦即 connections.config 中生效的那組，
        /// 與全站其他功能同源。本專案 PJE0030 / PJE0050 / SGE0000 / SGE0060 / RainReport 均採此寫法。
        ///
        /// 【勿改回硬編碼】原本此處寫死資料庫位址，導致部署到不同環境時
        /// 報表連到與 EF 不同的資料庫，各報表匯出因而發生伺服器錯誤。
        //private string connectionString { get { return db2.Database.Connection.ConnectionString; } }
        //string connectionString2 = @"Data Source=192.168.1.26;initial catalog=CMBASE;user id=***;password=***;";

        private string connectionString
        {
            get
            {
                string efConnectionString = ConfigurationManager.ConnectionStrings["CMEPMSEntities"].ConnectionString;
                var entityBuilder = new EntityConnectionStringBuilder(efConnectionString);
                return entityBuilder.ProviderConnectionString;
            }
        }

        private string connectionString2
        {
            get
            {
                string efConnectionString = ConfigurationManager.ConnectionStrings["CMBASEEntities"].ConnectionString;
                var entityBuilder = new EntityConnectionStringBuilder(efConnectionString);
                return entityBuilder.ProviderConnectionString;
            }
        }

        public static string ConvertToRepublicOfChinaDate(DateTime date)
        {
            int republicOfChinaYear = date.Year - 1911;

            string month = date.Month.ToString("00");
            string day = date.Day.ToString("00");
            string republicOfChinaDate = $"{republicOfChinaYear}{month}{day}";

            return republicOfChinaDate;
        }

        public ActionResult ReportPrint10(string[] CITY, string[] CITY_TOWN, string year1, string year2, string Row, string Field, List<decimal> master_plan, List<decimal> plan)
        {
            int A = int.Parse(year1);
            int B = int.Parse(year2);
            List<string> Year_Range = new List<string>();
            for (int num = A; num <= B; num++)
            {
                string year = num.ToString();
                if (num.ToString() == "97")
                {
                    year = "097";
                }
                else if (num.ToString() == "98")
                {
                    year = "098";
                }
                else if (num.ToString() == "99")
                {
                    year = "099";
                }
                Year_Range.Add(year);
            }
            var Year_Range_Length = Year_Range.Count();

            var TDV_TDDCFTS = db2.TDV_TDDCFTS.ToList();

            var TDDSTCF = (from a in db2.TDDSTCF
                           where !db2.TDDSTCF.Any(b =>
                                 b.SYR == a.SYR &&
                                 b.FD_CODE == a.FD_CODE &&
                                 string.Compare(b.SMN, a.SMN) > 0)
                           select new
                           {
                               SYR = a.SYR,
                               SMN = a.SMN,
                               COST_FEE = a.COST_FEE,
                               FD_CODE = a.FD_CODE
                           }).ToList();


            if (master_plan is null)
            {
                master_plan = new List<decimal>();
            }
            if (plan is null)
            {
                plan = new List<decimal>();
            }

            /// 把所有分支計畫塞進 plan[]
            foreach (var PID in master_plan)
            {
                var branch_LIST = db2.PJMFORM.Where(m => m.AN_PID == PID).DistinctBy(m => m.PJ_NAME).ToList();
                foreach (var item in branch_LIST)
                {
                    plan.Add(item.PID);
                }
            }

            string Plan_Name = "";

            foreach (var PID in plan)
            {
                Plan_Name = db2.PJMFORM.Where(m => m.PID == PID).FirstOrDefault().PJ_NAME;
            }

            var PJMFORM = db2.PJMFORM.Where(m => m.SYR.CompareTo(year1) >= 0 && m.SYR.CompareTo(year2) <= 0).ToList();

            var Source_Merge = from a in PJMFORM
                               join b in TDV_TDDCFTS on a.PID equals b.PID
                               join d in TDDSTCF on b.FD_CODE equals d.FD_CODE
                               select new Report10ViewModel
                               {
                                   PJ_NAME = a.PJ_NAME,
                                   SYR = b.PID_SYR,
                                   TD_CODE = b.TD_CODE,
                                   TD_NAME = b.TD_NAME,
                                   CITY = b.CITY,
                                   TOWN = b.TOWN,
                                   FUND = (int)b.CEN_FUND / 1000,
                                   COST_FEE = (int)Math.Round(((decimal)d.COST_FEE) / 1000)
                               };

            var PJ_CITY = "";
            var PJ_TOWN = "";

            var dataTable = new DataTable();
            var row2 = dataTable.NewRow();
            decimal total_sum = 0;
            decimal total_sum2 = 0;

            /// 如果有勾選實支數
            if (Row == "on")
            {

                /// 欄位建立
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
                    Plan_Name = db2.PJMFORM.Where(m => m.PID == PID).FirstOrDefault().PJ_NAME;


                    PJMFORM = db2.PJMFORM.Where(m => m.PJ_NAME.Contains(Plan_Name) && m.SYR.CompareTo(year1) >= 0 && m.SYR.CompareTo(year2) <= 0).ToList();

                    Source_Merge = from a in PJMFORM
                                   join b in TDV_TDDCFTS on a.PID equals b.PID
                                   join d in TDDSTCF on b.FD_CODE equals d.FD_CODE into dJoin
                                   from d in dJoin.DefaultIfEmpty()
                                   select new Report10ViewModel
                                   {
                                       PJ_NAME = a.PJ_NAME,
                                       SYR = b == null ? null : b.PID_SYR,
                                       TD_CODE = b == null ? null : b.TD_CODE,
                                       TD_NAME = b == null ? null : b.TD_NAME,
                                       CITY = b == null ? null : b.CITY,
                                       TOWN = b == null ? null : b.TOWN,
                                       FUND = b == null ? 0 : (int)b.CEN_FUND / 1000,
                                       COST_FEE = d == null ? 0 : (int)Math.Round(((decimal)d.COST_FEE) / 1000)
                                   };


                    for (int i = 0; i < CITY_TOWN.Count(); i++)
                    {
                        string[] sArray;
                        sArray = CITY_TOWN[i].Split('+');
                        PJ_CITY = sArray[0];
                        PJ_TOWN = sArray[1];
                        var TOWN_SOURCE = Source_Merge.Where(m => m.CITY == PJ_CITY && m.TOWN == PJ_TOWN);
                        //////實支數
                        var TOWN_TOTAL = TOWN_SOURCE.GroupBy(m => m.SYR).Select(g => new Report10ViewModel { SYR = g.Key, FUND = g.Sum(m => m.FUND) });   ///////這裡有 CITY  TOWN  FUND  上面有PJ_NAME
                                                                                                                                                          //////核定數
                        var TOWN_TOTAL2 = TOWN_SOURCE.GroupBy(m => m.SYR).Select(g => new Report10ViewModel { SYR = g.Key, COST_FEE = g.Sum(m => m.COST_FEE) });
                        var row = dataTable.NewRow();
                        row["計畫名稱"] = Plan_Name;
                        row["縣市"] = PJ_CITY;
                        row["行政區"] = PJ_TOWN;
                        decimal Total = 0;
                        decimal Total2 = 0;
                        foreach (var year in Year_Range)
                        {
                            decimal Year_Fund = 0;
                            decimal Year_Fund2 = 0;
                            var yearData = TOWN_TOTAL.FirstOrDefault(m => m.SYR == year.ToString());
                            var yearData2 = TOWN_TOTAL2.FirstOrDefault(m => m.SYR == year.ToString());

                            if (yearData != null)
                            {
                                Year_Fund = yearData.FUND;
                            }
                            if (yearData2 != null)
                            {
                                Year_Fund2 = yearData2.COST_FEE;
                            }

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

                /////////加總後欄位
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
            ///////////無實支數
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
                    Plan_Name = db2.PJMFORM.Where(m => m.PID == PID).FirstOrDefault().PJ_NAME;


                    PJMFORM = db2.PJMFORM.Where(m => m.PJ_NAME.Contains(Plan_Name) && m.SYR.CompareTo(year1) >= 0 && m.SYR.CompareTo(year2) <= 0).ToList();

                    Source_Merge = from a in PJMFORM
                                   join b in TDV_TDDCFTS on a.PID equals b.PID
                                   join d in TDDSTCF on b.FD_CODE equals d.FD_CODE
                                   select new Report10ViewModel
                                   {
                                       PJ_NAME = a.PJ_NAME,
                                       SYR = b.PID_SYR,
                                       TD_CODE = b.TD_CODE,
                                       TD_NAME = b.TD_NAME,
                                       CITY = b.CITY,
                                       TOWN = b.TOWN,
                                       FUND = (int)b.CEN_FUND / 1000,
                                       COST_FEE = (int)Math.Round(((decimal)d.COST_FEE) / 1000)
                                   };

                    for (int i = 0; i < CITY_TOWN.Count(); i++)
                    {
                        string[] sArray;
                        sArray = CITY_TOWN[i].Split('+');
                        PJ_CITY = sArray[0];
                        PJ_TOWN = sArray[1];
                        var TOWN_SOURCE = Source_Merge.Where(m => m.CITY == PJ_CITY && m.TOWN == PJ_TOWN);
                        var TOWN_TOTAL = TOWN_SOURCE.GroupBy(m => m.SYR).Select(g => new Report10ViewModel { SYR = g.Key, FUND = g.Sum(m => m.FUND) });   ///////這裡有 CITY  TOWN  FUND  上面有PJ_NAME
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

                    ///////////加總後欄位
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
            }

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);



            var path = Server.MapPath("~/Templates/表2.xlsx");
            var file = new FileInfo(path);
            var workbook = new Aspose.Cells.Workbook(path);

            Worksheet ws1 = workbook.Worksheets[0];
            ws1 = CopyWorksheet(workbook, ws1, "總表1", "相關計畫年度補助經費表");

            var Column_length = Report_Source.NumberToColumnLetter(Year_Range_Length + 5);
            var SOURCE = Source_Merge;

            ////////將欄位的表刻好，並把datatable的資料填進去
            if (Row == "on")
            {
                ws1.Name = "相關計畫年度執行表";
                Column_length = Report_Source.NumberToColumnLetter(Year_Range_Length * 2 + 6);
                for (int i = 4, j = 0; i < Year_Range_Length * 2 + 4; i = i + 2, j++)   // 填入欄位名稱
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

                        foreach (var item2 in Year_Range) //////篩年度核定數實支數
                        {
                            var year_column = Report_Source.NumberToColumnLetter(x + 4);
                            var year_column2 = Report_Source.NumberToColumnLetter((x + 1) + 4);
                            var YEAR_BLOCK = ws1.Cells[year_column + z];
                            var YEAR_BLOCK2 = ws1.Cells[year_column2 + z];
                            YEAR_BLOCK.GetStyle().Custom = "#,##0";
                            YEAR_BLOCK.Value = int.Parse(dataTable.Rows[y][item2 + "核定數"].ToString());
                            YEAR_BLOCK2.GetStyle().Custom = "#,##0";
                            YEAR_BLOCK2.Value = int.Parse(dataTable.Rows[y][item2 + "實支數"].ToString());
                            x = x + 2;
                        }
                        var ABC = Report_Source.NumberToColumnLetter(Year_Range.Count() * 2 + 4);
                        var DEF = Report_Source.NumberToColumnLetter(Year_Range.Count() * 2 + 5);

                        var TOTAL_BLOCK = ws1.Cells[ABC + z];
                        var TOTAL_BLOCK2 = ws1.Cells[DEF + z];

                        TOTAL_BLOCK.GetStyle().Custom = "#,##0";
                        TOTAL_BLOCK.Value = int.Parse(dataTable.Rows[y]["小計核定數"].ToString());
                        TOTAL_BLOCK2.GetStyle().Custom = "#,##0";
                        TOTAL_BLOCK2.Value = int.Parse(dataTable.Rows[y]["小計實支數"].ToString());

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

                    sum.GetStyle().Custom = "#,##0";
                    sum.Value = int.Parse(dataTable.Rows[CITY_TOWN.Count() * plan.Count][item + "核定數"].ToString());
                    sum2.GetStyle().Custom = "#,##0";
                    sum2.Value = int.Parse(dataTable.Rows[CITY_TOWN.Count() * plan.Count][item + "實支數"].ToString());

                    x = x + 2;
                }
                var SUM_COLUMN = Report_Source.NumberToColumnLetter(Year_Range.Count() * 2 + 4);
                var SUM_COLUMN2 = Report_Source.NumberToColumnLetter(Year_Range.Count() * 2 + 5);

                var SUM_TOTAL_BLOCK = ws1.Cells[SUM_COLUMN + (CITY_TOWN.Count() * plan.Count + 6)];
                var SUM_TOTAL_BLOCK2 = ws1.Cells[SUM_COLUMN2 + (CITY_TOWN.Count() * plan.Count + 6)];

                SUM_TOTAL_BLOCK.GetStyle().Custom = "#,##0";
                SUM_TOTAL_BLOCK.Value = int.Parse(dataTable.Rows[CITY_TOWN.Count() * plan.Count]["小計核定數"].ToString());
                SUM_TOTAL_BLOCK2.GetStyle().Custom = "#,##0";
                SUM_TOTAL_BLOCK2.Value = int.Parse(dataTable.Rows[CITY_TOWN.Count() * plan.Count]["小計實支數"].ToString());

            }
            else
            {
                // No need to convert SOURCE_MERGE to List, directly use it as IEnumerable

                for (int i = 4; i < Year_Range_Length + 4; i++)   // 填入欄位名稱
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
                var l = ws1.Cells.CreateRange("A" + (CITY_TOWN.Count() + 5) + ":C" + (CITY_TOWN.Count() + 5));
                l.Merge();
                l.Value = "合計";

                int z = 5;
                int y = 0;
                int x = 0;
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
                        YEAR_BLOCK.Value = dataTable.Rows[y][item2];
                        x++;
                    }
                    var ABC = Report_Source.NumberToColumnLetter(Year_Range.Count() + 4);
                    var TOTAL_BLOCK = ws1.Cells[ABC + z];
                    TOTAL_BLOCK.Value = dataTable.Rows[y]["小計"];
                    z++;
                    y++;
                }
                x = 0;
                foreach (var item in Year_Range)
                {
                    var year_column = Report_Source.NumberToColumnLetter(x + 4);
                    var sum = ws1.Cells[year_column + (CITY_TOWN.Count() + 5)];
                    sum.Value = dataTable.Rows[CITY_TOWN.Count()][item];
                    x++;
                }
                var SUM_COLUMN = Report_Source.NumberToColumnLetter(Year_Range.Count() + 4);
                var SUM_TOTAL_BLOCK = ws1.Cells[SUM_COLUMN + (CITY_TOWN.Count() + 5)];
                SUM_TOTAL_BLOCK.Value = dataTable.Rows[CITY_TOWN.Count()]["小計"];

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



            Aspose.Cells.Range c = ws1.Cells.CreateRange("A1:" + Column_length + (CITY_TOWN.Count() * plan.Count + 6).ToString());

            if (Row == "on")
            {
                c = ws1.Cells.CreateRange("A1:" + Column_length + (CITY_TOWN.Count() * plan.Count + 6).ToString());
            }

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

            borderstyle.HorizontalAlignment = TextAlignmentType.Center;
            borderstyle.Custom = "#,##0.00_);[Red]\\(0.00\\)";
            c.SetStyle(borderstyle);

            string filename = "F-1.相關計畫年度補助經費表" + DateTime.Now.ToString("yyyyMMddHHmmss");


            if (Row == "on")
            {
                filename = "F-1.相關計畫年度執行表" + DateTime.Now.ToString("yyyyMMddHHmmss");
            }

            workbook.Worksheets.RemoveAt(0);

            ws1.AutoFitColumns();

            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");
            var afile = new FileInfo(savepath); // 檔案路徑 
                                                // Specify the directory where you want to save the file
            string directoryPath = Server.MapPath(@"~\Temp\");

            // Combine the directory and file name to create the complete file path
            string newfilepath = System.IO.Path.Combine(directoryPath, filename + ".xlsx");

            // Save the workbook with the specified file path
            workbook.Save(newfilepath);


            using (var stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsx);

                // Set the position of the stream to the beginning
                stream.Seek(0, SeekOrigin.Begin);

                var bytes = stream.ToArray();

                // Close the stream
                stream.Close();

                // Return the Excel file
                return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename + ".xlsx");
            }
        }

        public ActionResult ReportPrint12(string SYR, string PID, DateTime DATE_START, DateTime DATE_END, string ODS)
        {
            Debug.WriteLine($"SYR:{SYR}");
            Debug.WriteLine($"PID:{PID}");
            Debug.WriteLine($"DATE_START:{DATE_START}");
            Debug.WriteLine($"DATE_END:{DATE_END}");
            Debug.WriteLine($"ODS:{ODS}");

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
            //DATE_START = DateTime.ParseExact(DATE_START, "yyyy-MM-dd", null).ToString("MMddyy");
            //DATE_END = DateTime.ParseExact(DATE_END, "yyyy-MM-dd", null).ToString("MMddyy");
            string DATE_START1 = ConvertToRepublicOfChinaDate(DATE_START);
            string DATE_END1 = ConvertToRepublicOfChinaDate(DATE_END);

            //string PID = MyTreeView1.SelectedValue;
            Report_Source.getProject09(ref TDMFORM, ref TDDSPRS, PID, DATE_START1, DATE_END1, connectionString);

            var workbook = new Aspose.Cells.Workbook(sFullFile);
            Aspose.Cells.Worksheet sheet1 = workbook.Worksheets["Sheet1"];
            Aspose.Cells.Worksheet sheet2 = workbook.Worksheets["Sheet2"];
            Aspose.Cells.Worksheet sheet3 = workbook.Worksheets["總表"];
            Aspose.Cells.Worksheet sheet4 = workbook.Worksheets["詳細表"];

            sheet1line = 4;
            sheet2line = 1;

            Aspose.Cells.Range sourceRange = sheet3.Cells.CreateRange(0, 0, 3, 6);

            // Specify the destination range in sheet1
            Aspose.Cells.Range destinationRange = sheet1.Cells.CreateRange(0, 0, 3, 6);

            // Copy values and styles from source range to destination range in sheet1
            destinationRange.Copy(sourceRange);

            // Set row heights in sheet1 based on row heights in sheet3
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
                    // Set row heights in sheet2 based on row heights in sheet4


                    // Set the text in cell A[sheet2line] in sheet2
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
                    //sheet4[10, 1, 10, 10].CopyTo(sheet2[sheet2line, 1]);
                    //sheet2[sheet2line, 1, sheet2line, 10].RowHeight = sheet4[10, 1, 10, 10].RowHeight;
                    sheet4.Cells.CreateRange(9, 0, 1, 10).Copy(sheet2.Cells.CreateRange(sheet2line, 0, 1, 10));
                    sheet2.Cells.SetRowHeight(sheet2line, sheet4.Cells.GetRowHeight(9));
                    sheet2line++;
                }
            }

            workbook.Worksheets.RemoveAt(sheet3.Index);
            workbook.Worksheets.RemoveAt(sheet4.Index);
            sheet1.Name = "總表";
            sheet2.Name = "詳細表";

            string filename = "H-1.在建工程抽查-查核-督導統計表" + DateTime.Now.ToString("yyyyMMddHHmmss");
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
        /// G-1.相關計畫年度補助經費表(含執行表)
        /// </summary>
        /// <param name="plan"></param>
        /// <param name="CITY"></param>
        /// <param name="year1"></param>
        /// <param name="year2"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult ReportPrint11(string plan, string CITY, string year1, string year2, string month, string type)
        {
            Trace.WriteLine($"ReportPrint11");
            Trace.WriteLine($"plan = {plan}");
            Trace.WriteLine($"CITY = {CITY}");
            Trace.WriteLine($"year1 = {year1}");
            Trace.WriteLine($"year2 = {year2}");
            Trace.WriteLine($"month = {month}");
            Trace.WriteLine($"type = {type}");


            int A = int.Parse(year1);
            int B = int.Parse(year2);

            List<string> Year_Range = new List<string>();

            for (int num = A; num <= B; num++)
            {
                string year = num.ToString();
                Year_Range.Add(year);
            }

            var Year_Range_Length = Year_Range.Count();

            var dataTable = new DataTable();
            dataTable.Columns.Add("項次");
            dataTable.Columns.Add("工程編號");
            dataTable.Columns.Add("縣市別");
            dataTable.Columns.Add("工程名稱");
            dataTable.Columns.Add("補助比例");
            dataTable.Columns.Add("總工程經費(中央款)");

            if (Year_Range_Length > 1)
            {
                for (int i = 0; i < Year_Range_Length - 1; i++)
                {
                    dataTable.Columns.Add(Year_Range[i]);
                }
                if (year1 != year2)
                {
                    dataTable.Columns.Add(year1 + "~" + year2 + "年度累計核定數");
                }
            }
            dataTable.Columns.Add("已撥保留數");
            dataTable.Columns.Add("已發生權責保留數");

            /// 第X期保留撥款數, 目前先保留之後會再加上
            /// dataTable.Columns.Add("保留核銷期程");
            dataTable.Columns.Add(year2 + "核定數");
            dataTable.Columns.Add(year2 + "撥款數");

            if (year1 != year2)
            {
                dataTable.Columns.Add(year1 + "~" + year2 + "年度累計撥款數");
            }

            dataTable.Columns.Add("最新年度可支用數");
            dataTable.Columns.Add("發包日期(預定)");
            dataTable.Columns.Add("發包日期(實際)");
            dataTable.Columns.Add("開工日期(預定)");
            dataTable.Columns.Add("開工日期(實際)");
            dataTable.Columns.Add("完工日期(預定)");
            dataTable.Columns.Add("完工日期(實際)");
            dataTable.Columns.Add("工程進度(%)月底(預定)");
            dataTable.Columns.Add("工程進度(%)月底(實際)");
            dataTable.Columns.Add("辦理情形");

            var path = Server.MapPath("~/Templates/表3.xlsx");
            var file = new FileInfo(path);
            string PJ_NAME = "";
            string PJ_CODE = "";
            string Plan_CITY = "";
            string[] Plan = plan.Split(',');

            int count = 1;

            var CRDMNDS = db2.CRDMNDS.Select(m => new
            {
                TD_CODE = m.TD_CODE,
                OCR_FUND = m.OCR_FUND,
                REV_FUND = m.REV_FUND
            }).ToList();

            var TDDSDWB = db3.TDDSDWB.Select(m => new
            {
                IT_CODE = m.IT_CODE, /// 修正以 IT_CODE 取得資料非 IT_NAME
                PRE_EDT = m.PRE_EDT,
                REL_EDT = m.REL_EDT,
                TD_CODE = m.TD_CODE,
                IT_NAME = m.IT_NAME
            }).ToList();

            var TDV_TDDCFTS = db2.TDV_TDDCFTS.ToList();

            var TDDMSCD = db3.TDDMSCD.ToList();

            var TDDMSCD_EXECT_DESC = db3.TDDMSCD.Where(m => m.EXECT_DESC != null && m.EXECT_DESC != "").ToList();

            /// 20260303 修改 => 李憨兒要求加上月份
            var TDDSTCF = db2.TDDSTCF.ToList();

            /// 將 FD_CODE 最後年分月份都抓出來
            var TDDSTCF_FINAL = TDDSTCF.OrderByDescending(x => x.SYR)
            .ThenByDescending(x => x.SMN)
            .GroupBy(x => x.FD_CODE)
            .Select(g => g.First())
            .ToList();

            /// 將整個工程的總撥款數抓出來
            var TDDSTCF_FINAL_TOTAL = TDDSTCF_FINAL.GroupBy(m => m.TD_CODE).Select(g => new
            {
                TD_CODE = g.Key,
                Total_Fund = g.Sum(m => m.COST_FEE) + g.Sum(m => m.PPAY_FEE)
            }).ToList();

            /// 將最後年分工程的總撥款數抓出來
            var TDDSTCF_LASTYEAR_FINAL_TOTAL = TDDSTCF_FINAL.Where(m => m.SYR == year2).GroupBy(m => m.TD_CODE).Select(g => new
            {
                TD_CODE = g.Key,
                Total_Fund = g.Sum(m => m.COST_FEE) + g.Sum(m => m.PPAY_FEE)
            }).ToList();

            var PJMFORM = new List<PJMFORM_CMEPMS>();

            for (int l = 0; l < Plan.Count(); l++)
            {
                PJ_CODE = Plan[l];

                switch (Plan[l])
                {
                    case "01":
                        PJ_NAME = "01辦理水安全業務費";
                        break;
                    case "02":
                        PJ_NAME = "02辦理水安全獎補助費";
                        break;
                    case "03":
                        PJ_NAME = "03辦理水安全設備投資費";
                        break;
                    case "04":
                        PJ_NAME = "都市總合治水建設計畫(公務預算)";
                        break;
                    case "07":
                        PJ_NAME = "因應氣候變遷縣市管河川及排水整體改善計畫";
                        break;
                }


                PJMFORM = db2.PJMFORM
                            .Where(m => m.PJ_NAME == PJ_NAME
                                     && m.SYR.CompareTo(year1) >= 0
                                     && m.SYR.CompareTo(year2) <= 0)
                            .ToList();

                string[] city = CITY.Split(',');

                for (int i = 0; i < city.Count(); i++)
                {
                    Plan_CITY = city[i];

                    var TDMFORM = db3.TDMFORM.Where(m => m.TD_CITY == Plan_CITY);

                    // 包含年份的工程都抓出來(所有會有很多一樣的工程名稱)但年份不同
                    var Source_Merge = (
                                         from a in PJMFORM
                                         join c in TDV_TDDCFTS on a.PID equals c.PID
                                         join d in TDMFORM.ToList() on c.TD_CODE equals d.TD_CODE
                                         select new Report11ViewModel
                                         {
                                             SYR = c.PID_SYR,
                                             /// TD_CODE = c.SIP_TD_CODE,
                                             TD_CODE = c.TD_CODE,
                                             TD_NAME = c.TD_NAME,
                                             CITY = c.CITY,
                                             TOWN = c.TOWN,
                                             TD_MEMO = d.TD_MEMO,
                                             TID = (decimal)c.TID,
                                             TD_FUND = d.TD_FUND ?? 0,
                                             FUND = c.CEN_FUND
                                         }
                                       ).ToList();

                    // 將工程名稱抓出來(不區分年分)

                    var groupedSource = (from item in Source_Merge
                                         group item by new { item.TD_CODE, item.TD_NAME, item.CITY, item.TOWN, item.TD_MEMO, item.TD_FUND } into grouped
                                         select new Report11ViewModel
                                         {
                                             TD_CODE = grouped.Key.TD_CODE,
                                             TD_NAME = grouped.Key.TD_NAME,
                                             CITY = grouped.Key.CITY,
                                             TOWN = grouped.Key.TOWN,
                                             TD_FUND = grouped.Key.TD_FUND,
                                             FUND = grouped.Sum(item => item.FUND),
                                             TD_MEMO = grouped.Key.TD_MEMO
                                         }).ToList();

                    var SameSYRTDCODEMerge = (from item in Source_Merge
                                              group item by new { item.TD_CODE, item.SYR } into grouped
                                              select new Report11ViewModel
                                              {
                                                  TD_CODE = grouped.Key.TD_CODE,
                                                  SYR = grouped.Key.SYR,
                                                  FUND = grouped.Sum(item => item.FUND)
                                              }).ToList();

                    // 分支計畫下全部工程
                    foreach (var item in groupedSource)
                    {
                        try
                        {
                            var row = dataTable.NewRow();
                            row["項次"] = count;
                            count++;
                            row["工程編號"] = item.TD_CODE;
                            row["縣市別"] = item.CITY;
                            row["工程名稱"] = item.TD_NAME;

                            if (year1 != year2)
                            {
                                row[year1 + "~" + year2 + "年度累計核定數"] = item.FUND;
                            }

                            double? CEN_SP_RATIPO_Nullable = null; // 宣告為 nullable 的 double
                            var model = TDV_TDDCFTS.Where(m => m.TD_CODE == item.TD_CODE.Replace("-", "")).OrderByDescending(m => m.PID_SYR).FirstOrDefault();

                            if (model != null)
                            {
                                CEN_SP_RATIPO_Nullable = (double?)model.CEN_SP_RATIO; // 將 model.CEN_SP_RATIO 賦值給 nullable double
                            }

                            double CEN_SP_RATIPO = CEN_SP_RATIPO_Nullable ?? 0; // 若 CEN_SP_RATIPO_Nullable 為 null,則使用 0

                            var TD_CODE2 = item.TD_CODE.Replace("-", "");
                            row["補助比例"] = CEN_SP_RATIPO.ToString("00") + "%";
                            row["總工程經費(中央款)"] = item?.TD_FUND.ToString();
                            row["已撥保留數"] = CRDMNDS.FirstOrDefault(m => m.TD_CODE == TD_CODE2)?.REV_FUND ?? 0;
                            row["已發生權責保留數"] = CRDMNDS.LastOrDefault(m => m.TD_CODE == TD_CODE2)?.OCR_FUND ?? 0;
                            row["發包日期(預定)"] = TDDSDWB.FirstOrDefault(m => m.TD_CODE == TD_CODE2 && m.IT_CODE == "TK105")?.PRE_EDT; /// 修正以 IT_CODE 取得資料非IT_NAME
                            row["發包日期(實際)"] = TDDSDWB.FirstOrDefault(m => m.TD_CODE == TD_CODE2 && m.IT_CODE == "TK105")?.REL_EDT; /// 修正以 IT_CODE 取得資料非IT_NAME
                            row["開工日期(預定)"] = TDDSDWB.FirstOrDefault(m => m.TD_CODE == TD_CODE2 && m.IT_CODE == "TK106")?.PRE_EDT; /// 修正以 IT_CODE 取得資料非IT_NAME
                            row["開工日期(實際)"] = TDDSDWB.FirstOrDefault(m => m.TD_CODE == TD_CODE2 && m.IT_CODE == "TK106")?.REL_EDT; /// 修正以 IT_CODE 取得資料非IT_NAME
                            row["完工日期(預定)"] = TDDSDWB.FirstOrDefault(m => m.TD_CODE == TD_CODE2 && m.IT_CODE == "TK107")?.PRE_EDT; /// 修正以 IT_CODE 取得資料非IT_NAME
                            row["完工日期(實際)"] = TDDSDWB.FirstOrDefault(m => m.TD_CODE == TD_CODE2 && m.IT_CODE == "TK107")?.REL_EDT; /// 修正以 IT_CODE 取得資料非IT_NAME
                            row["工程進度(%)月底(預定)"] = TDDMSCD.Where(m => m.TD_CODE == TD_CODE2 && m.SYR == year2).OrderByDescending(m => m.SYR).FirstOrDefault()?.PRE_SCHD;
                            row["工程進度(%)月底(實際)"] = TDDMSCD.Where(m => m.TD_CODE == TD_CODE2 && m.SYR == year2).OrderByDescending(m => m.SYR).FirstOrDefault()?.REL_SCHD;
                            row["辦理情形"] = TDDMSCD_EXECT_DESC.Where(m => m.TD_CODE == TD_CODE2).OrderByDescending(d => d.ID).FirstOrDefault()?.EXECT_DESC;

                            foreach (var item2 in Year_Range)
                            {
                                /// 修正最後一個年度核定數欄位沒有顯示資料
                                var column = (item2 == year2) ? item2 + "核定數" : item2;
                                row[column] = SameSYRTDCODEMerge.Where(m => m.TD_CODE == item.TD_CODE && m.SYR == item2).FirstOrDefault()?.FUND ?? 0;
                            }

                            row[year2 + "撥款數"] = TDDSTCF_LASTYEAR_FINAL_TOTAL.Where(m => m.TD_CODE == item.TD_CODE).FirstOrDefault()?.Total_Fund ?? 0;

                            /// 修正年度區間選擇同一年度時EXCEL無法顯示資料
                            if (year1 != year2)
                            {
                                row[year1 + "~" + year2 + "年度累計撥款數"] = TDDSTCF_FINAL_TOTAL.Where(m => m.TD_CODE == item.TD_CODE).FirstOrDefault()?.Total_Fund ?? 0;
                            }

                            dataTable.Rows.Add(row);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"發生錯誤: {ex.Message}");
                        }
                    }
                }

            }

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);

            var workbook = new Aspose.Cells.Workbook(path);

            var ws1 = workbook.Worksheets["雨水前瞻控管表"];
            var columns_length = dataTable.Columns.Count;
            var Excel_column = Report_Source.NumberToColumnLetter(columns_length);
            var First_Title = ws1.Cells.CreateRange("A1:" + Excel_column + "1");
            First_Title.Merge();

            //ws1.Name = PJ_NAME;
            //First_Title.Value = year1 + "~" + year2 + PJ_NAME;
            switch (plan)
            {
                case "04":
                    ws1.Name = "都市總合治水控管表";
                    First_Title.Value = year1 + "~" + year2 + "都市總合治水計畫控管表";
                    break;

                case "01":
                case "02":
                case "03":
                    First_Title.Value = year1 + "~" + year2 + "前瞻基礎建設計畫-縣市管河川及區域排水整體改善計畫";
                    break;

                case "07":
                    ws1.Name = "因應氣候變遷縣市管河川及排水整體改善計畫";
                    First_Title.Value = year1 + "~" + year2 + "因應氣候變遷縣市管河川及排水整體改善計畫";
                    break;
            }

            var All_block = ws1.Cells.CreateRange("A1:" + Excel_column + (dataTable.Rows.Count + 4));

            int commentIndex = ws1.Comments.Add("E2");

            var comment = workbook.Worksheets[0].Comments[commentIndex];

            comment.Note = "抓取最新年份工程的補助比例";
            ws1.AutoFitColumns();

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

            Aspose.Cells.Border innerBorder = borderstyle.Borders[BorderType.Vertical];
            innerBorder.LineStyle = CellBorderType.Thin;

            innerBorder = borderstyle.Borders[BorderType.Horizontal];
            innerBorder.LineStyle = CellBorderType.Thin;
            borderstyle.IsTextWrapped = true;  /// 自動換行
            borderstyle.HorizontalAlignment = TextAlignmentType.Center;
            All_block.SetStyle(borderstyle);

            for (int i = 0; i < columns_length; i++)
            {
                Excel_column = Report_Source.NumberToColumnLetter(i + 1);
                ws1.Cells.CreateRange(Excel_column + "2:" + Excel_column + "3").Merge();
                ws1.Cells[Excel_column + "2"].Value = dataTable.Columns[i];
            }

            double[] Sum_Array = new double[columns_length];
            int j = 4;
            foreach (DataRow row in dataTable.Rows)
            {
                for (int i = 0; i < columns_length; i++)
                {
                    Excel_column = Report_Source.NumberToColumnLetter(i + 1);
                    var a1 = ws1.Cells[Excel_column + j];
                    Aspose.Cells.Style a1_Style = a1.GetStyle();
                    a1_Style.Number = 5;
                    var data_column = dataTable.Columns[i].ColumnName;
                    string value = row[data_column].ToString();

                    bool containsNumber = Regex.IsMatch(data_column, @"\d");
                    bool containsNumber2 = Regex.IsMatch((string)value, @"\d");

                    decimal math_value;

                    if (year1 == year2) // 如果起訖年相同取消歷年核定數和累計核定數
                    {
                        Year_Range_Length = 0;
                    }
                    if (i == 5) // 總工程經費(中央款)
                    {
                        Sum_Array[i] += Source.Tools.StringToDouble(value);
                    }
                    else if (i == 8 + Year_Range_Length) // 核定數
                    {
                        Sum_Array[i] += Source.Tools.StringToDouble(value);
                    }
                    else if (i == 9 + Year_Range_Length) // 撥款數
                    {
                        Sum_Array[i] += Source.Tools.StringToDouble(value);
                    }

                    // 年度累計撥款數
                    // 修正合計欄位沒有顯示資料問題
                    if (i == 10 + Year_Range_Length && year1 != year2)
                    {
                        Sum_Array[i] += Source.Tools.StringToDouble(value);
                    }

                    // 各年度核定數及累計核定數
                    for (int m = 0; m < Year_Range_Length; m++)
                    {
                        if (i == 6 + m)
                        {
                            Sum_Array[i] += Source.Tools.StringToDouble(value);
                        }
                    }


                    if (containsNumber)
                    {
                        if (containsNumber2)
                        {
                            math_value = Convert.ToDecimal(value);
                            a1.PutValue(math_value);
                        }
                    }
                    else
                    {
                        if (data_column == "總工程經費(中央款)")
                        {
                            math_value = Convert.ToDecimal(value);
                            a1.PutValue(math_value);
                        }
                        else
                        {
                            a1.PutValue(value);
                        }
                    }
                    a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                    a1.SetStyle(a1_Style);

                }

                j++;

                if (j == dataTable.Rows.Count + 4)
                {
                    for (int i = 0; i < columns_length; i++)
                    {
                        Excel_column = Report_Source.NumberToColumnLetter(i + 1);
                        var a1 = ws1.Cells[Excel_column + j];
                        Aspose.Cells.Style a1_Style = a1.GetStyle();
                        a1_Style.Number = 5;

                        double value = Sum_Array[i];

                        if (i == 0)
                        {
                            a1.PutValue("合計");
                        }

                        if (value != 0)
                        {
                            a1.PutValue(value);
                        }

                        a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                        a1.SetStyle(a1_Style);

                    }
                }
            }

            for (int i = 0; i < columns_length; i++)
            {
                Excel_column = Report_Source.NumberToColumnLetter(i + 1);
                var data_column = dataTable.Columns[i].ColumnName;
                ws1.Cells.SetColumnWidth(i + 1, 16);

                if (data_column == "辦理情形")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 150);
                }
            }

            string filename = "";
            string mainPlan = Plan.Last();
            Debug.WriteLine($"mainPlan = {mainPlan}");
            switch (mainPlan)
            {
                case "04":
                    filename = "都市總合治水計畫控管表";
                    break;

                case "01":
                case "02":
                case "03":
                    filename = "前瞻基礎建設計畫-縣市管河川及區域排水整體改善計畫";
                    break;

                case "07":
                    filename = "因應氣候變遷縣市管河川及排水整體改善計畫";
                    break;
            }
            filename += DateTime.Now.ToString("yyyyMMddHHmmss");

            var savepath = Server.MapPath(@"~/Temp/");
            string newFilepath = System.IO.Path.Combine(savepath, filename + ".xlsx");

            if (type == "ODS")
            {
                filename += ".ods"; // 確保副檔名一致
                workbook.Save(newFilepath, SaveFormat.ODS); // 直接儲存為 ODS 格式
                return File(System.IO.File.ReadAllBytes(newFilepath),
                            "application/vnd.oasis.opendocument.spreadsheet",
                            filename);
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
        /// E-1.工程細項執行表(自訂欄位)
        /// </summary>
        /// <param name="plan"></param>
        /// <param name="CITY"></param>
        /// <param name="Row"></param>
        /// <param name="Field"></param>
        /// <param name="COLUMN1"></param>
        /// <param name="PDF"> 0/1 </param>
        /// <param name="abnormal"> 異常狀況:進度未登錄清單/施工落後表/已完工未驗收/逾期未完工表 </param>
        /// <returns></returns>
        public ActionResult ReportPrint9(string[] plan, string[] CITY, string Row, string Field, string[] COLUMN1, string PDF, string abnormal)
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

            // 總工程費用
            var TDDBGLS = db3.TDDBGLS.ToList()
                .Where(m => m.SEQ_NO == 0 && m.REC_TP == "B")
                .GroupBy(m => m.TD_CODE)
                .Select(p => new
                {
                    TD_CODE = p.Key,
                    Total = p.Sum(m => m.BUDGET_FEE)
                }).ToList();

            // 20241022修改項目 => 取得日期條件
            var TDDSDWB_CONTRATE = db3.TDDSDWB.Where(m => m.IT_CODE == "TK105").ToList(); // 發包日期
            var TDDSDWB_START = db3.TDDSDWB.Where(m => m.IT_CODE == "TK106").ToList();    // 工程開工日期
            var TDDSDWB_END = db3.TDDSDWB.Where(m => m.IT_CODE == "TK107").ToList();      // 工程完工日期                  
            var TDDSDWB_TK103 = db3.TDDSDWB.Where(m => m.IT_CODE == "TK103").ToList();    // (細部)設計完成日期
            var oriBUDGET = (from t in db3.TDDBGLS                                        // 已發包的預算款項  
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
                                 t.IT_ORDER
                             } into temp
                             group temp by temp.TD_CODE into g
                             select new
                             {
                                 TD_CODE = g.Key,
                                 CODE_DESCs = g.OrderBy(x => x.IT_ORDER).Select(x => x.CODE_DESC).ToList()
                             }).ToList();

            var BUDGET = oriBUDGET
                        .Select(x => new
                        {
                            x.TD_CODE,
                            ALL_OUTSRC_FG = string.Join(",\n", x.CODE_DESCs)
                        }).ToList();

            // 20241102 修改項目 => 實際驗收完成日期
            var TDDSDWB_TK108 = db3.TDDSDWB.Where(m => m.IT_CODE == "TK108").ToList();

            string Year = "";
            string Master_Plan_Name = "";
            string Plan_Name = "";

            // 把重複年度月份的資料和再一起
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

            // query 是抓每個年度最後月份的經費
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

            // newList 是把COST_FEE 依照年分加總前面年份的數值做的表
            var newList = query.Select(item => new
            {
                SYR = item.SYR,
                SMN = item.SMN,
                COST_FEE = item.COST_FEE, //////年度實支數
                PPAY_FEE = item.PPAY_FEE, //////年度暫付數
                TD_CODE1 = item.TD_CODE1,
                SUM_COST_FEE = query
                .Where(t => t.TD_CODE1 == item.TD_CODE1 && t.SYR.CompareTo(item.SYR) <= 0)
                .Sum(t => t.COST_FEE),
                SUM_PPAY_FEE = query
                .Where(t => t.TD_CODE1 == item.TD_CODE1 && t.SYR.CompareTo(item.SYR) <= 0)
                .Sum(t => t.PPAY_FEE),
                SUM_FUND = query
                .Where(t => t.TD_CODE1 == item.TD_CODE1 && t.SYR.CompareTo(item.SYR) <= 0)
                .Sum(t => t.COST_FEE) + query
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

            var TDV_TDDCFTS = db2.TDV_TDDCFTS.Where(m => m.YR_FUND != 0 || (m.FCAL_FG ?? 0 == 1)).ToList();

            // 新增一個新欄位 SYR + TD_CODE 給 TDV_TDDCFTS                                                                                                 
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

            var TDV_TDDCFTS_SUM = db2.TDV_TDDCFTS.GroupBy(m => m.TD_CODE).Select(p => new
            {
                TD_CODE = p.Key,
                Total = p.Sum(m => m.CEN_FUND)
            }).ToList();

            for (int w = 0; w < plan.Count(); w++)
            {
                string[] sArray;
                sArray = plan[w].Split('+');
                Year = sArray[0];
                Master_Plan_Name = sArray[1];
                Plan_Name = sArray[2];


                var PJMFORM = db2.PJMFORM.Where(m => m.SYR == Year && m.PJ_NAME == Plan_Name).FirstOrDefault();

                var TDDSTCF = newList.Where(m => m.SYR == Year).ToList(); // 這個是某年度的累積撥款數

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
                                        // 241022 => 取得預定和實際(細部)設計完成日期
                                    join tk103 in TDDSDWB_TK103 on b.TD_CODE equals tk103.TD_CODE into tk103Join
                                    from tk103 in tk103Join.DefaultIfEmpty()
                                        // 241022 => 取得已發包的預算款項名稱
                                    join bud in BUDGET on b.TD_CODE equals bud.TD_CODE into budJoin
                                    from bud in budJoin.DefaultIfEmpty()
                                        // 取得實際驗收完成日期
                                    join tk108 in TDDSDWB_TK108 on b.TD_CODE equals tk108.TD_CODE into tk108Join
                                    from tk108 in tk108Join.DefaultIfEmpty()

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
                                        IPJ_DESC = b.IPJ_DESC, // 辦理層級
                                        ADV_DP = b.ADV_DP,
                                        TID = b?.TID != null ? (decimal)b.TID : 0,
                                        FUND = c?.Total != null ? (decimal)c.Total : 0, // 中央款
                                        PERCENT_FUND = b.CEN_SP_RATIO != null ? (decimal)b.CEN_SP_RATIO : 0, // 工程費比例
                                        TOTAL_ALL_BUDGET_FEE = b?.CEN_FUND != null ? (decimal)b.CEN_FUND : 0,
                                        TD_LINE = d?.TD_LINE,
                                        TD_NOTE = d?.TD_NOTE,
                                        TD_MEMO = d?.TD_MEMO,
                                        SMP_DEPT_NAME = e?.SMP_DEPT_NAME,
                                        TOTAL_FUND = f?.Total != null ? (decimal)f?.Total : 0, // 總工程費(元)
                                        COST_FEE = (g?.COST_FEE != null ? (decimal)g.COST_FEE : 0) + (g?.PPAY_FEE != null ? (decimal)g.PPAY_FEE : 0),// 當年度撥款數
                                        COST_FEE_TOTAL = (g?.SUM_COST_FEE != null ? (decimal)g.SUM_COST_FEE : 0) + (g?.SUM_PPAY_FEE != null ? (decimal)g.SUM_PPAY_FEE : 0), // 累積撥款數
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
                                        // 20241022 修改項目 => 取得報表新增的需求欄位資料
                                        PRE_TK103_DATE = tk103?.PRE_EDT,
                                        REL_TK103_DATE = tk103?.REL_EDT,
                                        ALL_OUTSRC_FG = bud?.ALL_OUTSRC_FG,
                                        DSG_BUSS_NAME = d?.DSG_BUSS_NAME,
                                        SVY_BUSS_NAME = d?.SVY_BUSS_NAME,
                                        BUSS_NAME = d?.BUSS_NAME,
                                        FINISH_FG = d?.TD_FINISH == true ? "是" : "",
                                        START_FUND = d?.START_FUND,
                                        HOUSEHOLD = d?.HOUSEHOLD,
                                        TOTAL_HOUSEHOLD = d?.TOTAL_HOUSEHOLD,
                                        REL_TK108_DATE = tk108?.REL_EDT,
                                    });

                // 異常狀況篩選條件
                switch (abnormal)
                {
                    case "1":/// 進度未登錄清單：辦理情形空白
                        SOURCE_MERGE = SOURCE_MERGE.Where(s => string.IsNullOrEmpty(s.TD_MEMO)).ToList();
                        break;
                    case "2":/// 施工落後表：進度差異(預定-實際)是負的
                        SOURCE_MERGE = SOURCE_MERGE.Where(s => s.PROGRESS_DIFFERENCE < 0).ToList();
                        break;
                    case "3":/// 已完工未驗收：有實際完工日期但沒填實際驗收完成日期
                        SOURCE_MERGE = SOURCE_MERGE
                            .Where(s => !string.IsNullOrEmpty(s.REL_END_DATE)
                            && string.IsNullOrEmpty(s.REL_TK108_DATE)).ToList();
                        break;
                    case "4":/// 逾期未完工表：預定完工日期已過且沒填實際完工日期
                             /// 將當日轉成民國年顯示方式
                        DateTime now = DateTime.Now;
                        TaiwanCalendar taiwanCalendar = new TaiwanCalendar();
                        int rocYear = taiwanCalendar.GetYear(now);
                        string month = now.ToString("MM");
                        string day = now.ToString("dd");
                        string formattedDate = string.Format("{0:D3}{1}{2}", rocYear, month, day);

                        SOURCE_MERGE = SOURCE_MERGE
                            .Where(s => !string.IsNullOrEmpty(s.PRE_END_DATE)
                            && (s.PRE_END_DATE.CompareTo(formattedDate) < 0)
                            && string.IsNullOrEmpty(s.REL_END_DATE)).ToList();
                        break;
                }
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

            // ---------------------------- Aspose.Cells套件使用 ----------------------------
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
                        else if (title == "當年度核定數")
                        {
                            CultureInfo cultureInfo = CultureInfo.GetCultureInfo("zh-TW");
                            decimal value = Source.Tools.ToDecimal(dataTable.Rows[j - 3]["當年度核定數"]);
                            a1.PutValue(value/*.ToString("#,##0")*/);
                            a1_Style.Custom = "#,##0;[Red]\\(0.00\\)";
                            a1.SetStyle(a1_Style);
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
                        foreach (DataRow row in dataTable.Rows)
                        {
                            if (row["本署補助經費(中央款)(元)"] != DBNull.Value) // 確認值不為 DBNull
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
                            if (row["當年度核定數"] != DBNull.Value) // 確認值不為 DBNull
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
                            if (row["累積撥款數"] != DBNull.Value) // 確認值不為 DBNull
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
                        //////else就留空
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
                else if (data_column == "工程規模概述")
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
                else if (data_column == "落後原因")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "改善對策")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
                }
                else if (data_column == "備註")
                {
                    ws1.Cells.SetColumnWidth(i + 1, 40.17);
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

        static string GetYear(string input)
        {
            // Using regular expression to match the numeric part
            string pattern = @"\d+";
            Match match = Regex.Match(input, pattern);

            // If a match is found, return the numeric part
            if (match.Success)
            {
                return match.Value;
            }

            // If no match is found, return an empty string
            return string.Empty;
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
        /// <param name="PDF"></param>
        /// <returns></returns>
        public ActionResult ReportPrint4(string StartYear, string EndYear, string[] plan, string[] CITY, string Row, string Field, string[] COLUMN1, string PDF)
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
                    .Where(m => m.REC_TP == "B" || m.REC_TP == "C") // 抓所有 C（契約）和 B（預算）
                    .ToList(); // 先查回記憶體再分段處理，避免 EF 無法轉譯

            var TDDBGLS = list
                .GroupBy(m => m.TD_CODE)
                .Select(g =>
                {
                    // 優先抓 OUTSRC_FG = 1（代表已發包）
                    var contractList = g.Where(x => x.OUTSRC_FG == true);
                    var contractTotal = contractList.Sum(x => x.BUDGET_FEE ?? 0);

                    // 如果 contract 為 0，就改抓 REC_TP = B（預算資料）
                    var budgetList = g.Where(x => x.REC_TP == "B");
                    var budgetTotal = budgetList.Sum(x => x.BUDGET_FEE ?? 0);

                    return new
                    {
                        TD_CODE = g.Key,
                        ChosenSource = contractTotal > 0 ? "OUTSRC_FG=1" : "REC_TP=B",
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


            // v241223 移除已被刪除的分項內容對應資料
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
            // v241223 修改中央款取得資料表
            var TDV_TDDCFTS_SUM = TDV_TDDCFTS.GroupBy(m => m.TD_CODE).Select(p => new
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
                //var TD_CODE_LIST = db2.CTDMNDS.Where(m => m.PID == PJMFORM.PID).ToList();
                //var TDDSTCF = db2.TDDSTCF.Where(m => m.SYR == Year).OrderByDescending(m => m.SMN).Take(1).ToList();

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
                                    join g in financialDataQuery on b?.SYR_TD_CODE equals g.SYR_TD_CODE into Joing //-k 2411-3
                                    from g in Joing.DefaultIfEmpty()
                                    join i in TDDSDWB_CONTRATE on b?.TD_CODE equals i.TD_CODE into contrateJoin
                                    from i in contrateJoin.DefaultIfEmpty()
                                        //join m in TDDSDWB_START on b?.TD_CODE equals m.TD_CODE into startJoin
                                        //from m in startJoin.DefaultIfEmpty()
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
                                        COST_FEE = g?.SUM_FEE ?? 0,//-k2411-3 //(g?.COST_FEE != null ? (decimal)g.COST_FEE : 0) + (g?.PPAY_FEE != null ? (decimal)g.PPAY_FEE : 0),////當年度撥款數
                                                                   //-k2411-3 //COST_FEE_TOTAL = (g?.SUM_COST_FEE != null ? (decimal)g.SUM_COST_FEE : 0) + (g?.SUM_PPAY_FEE != null ? (decimal)g.SUM_PPAY_FEE : 0), ////累積撥款數
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

                    // 20241119 修改項目
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

                            // 抓取年度區間內各年度的核定數(只會取得該年度之前的)
                            // 每筆工程編號每年只會有一筆資料
                            foreach (var year in allYears)
                            {
                                if (year.CompareTo(SYR.ToString()) > 0)
                                    continue;

                                decimal Fund = TDV_TDDCFTS2.Where(m => m.TD_CODE == TD_CODE.ToString() && m.PID_SYR == year.ToString())
                                    .Select(s => s.CEN_FUND).DefaultIfEmpty(0).Sum();

                                row[year + "年核定數"] = Fund;

                                // 計算xxx年度核定數總和
                                string idYearKey = TD_CODE + "-" + year;

                                if (!procFD_CODE.Contains(idYearKey))
                                {
                                    dict[year + "年核定數總和"] += Fund;
                                    procFD_CODE.Add(idYearKey);
                                }
                            }
                        }
                        else if (columnName == "累積撥款數") //-k2411-3
                        {
                            // 計算每年累積撥款數
                            // 每筆工程編號每年只會有一筆資料
                            if (!string.IsNullOrEmpty(SYR.ToString()))
                            {
                                var preSum = financialDataQuery.Where(s => s.TD_CODE1 == TD_CODE.ToString())
                                .Where(s => s.SYR.CompareTo(SYR) <= 0).Sum(s => s.SUM_FEE);

                                row["累積撥款數"] = preSum;

                                //計算累積撥款數總和
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
            borderstyle.IsTextWrapped = true;  // 自動換行
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
                    // 顯示已處理過的總和
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

                        // 顯示已處理過的總和
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
                    // 每筆工程編號只需取得最新一筆累積撥款數
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

        public ActionResult ReportPrint5(string PID, string YearWrite, string SMN, string[] CityList, string funding, string ODS)
        {
            /// <summary> 只針對單一 PID(日後須改為接受多個PID) </summary>
            //decimal PIDD = decimal.Parse(PID);
            //var PJ = db2.PJMFORM.Where(m => m.PID == PIDD).FirstOrDefault().PJ_NAME;

            var path = Server.MapPath("~/Templates/PJV0010_08.xlsx");
            var file = new FileInfo(path);

            string City = "";
            DataTable dt = new DataTable();

            foreach (var item in CityList)
            {
                City = item;
                dt.Merge(Source.Report_Source.getBillionProject08(PID, YearWrite, SMN, City, funding, connectionString));
            }


            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);


            var workbook = new Aspose.Cells.Workbook(path);
            var ws1 = workbook.Worksheets["Sheet1"];
            var ws2 = workbook.Worksheets["Sheet2"];
            var ws3 = workbook.Worksheets["Sheet3"];

            #region ======================== Sheet[細項執行情形] ========================

            // 文件標題
            //ws1.Cells["A1"].PutValue("表A-5、內政部辦理" + PJ + "細項執行情形");
            ws1.Cells["A1"].PutValue("表A-5_內政部辦理細項執行情形");
            ws1.Cells["AI2"].PutValue("資料截至：" + YearWrite + "年" + SMN + "月");


            var a = ws1.Cells.CreateRange("A5:BM5");

            for (int i = 6; i < dt.Rows.Count + 5; i++)
            {
                var b = ws1.Cells.CreateRange("A" + i.ToString() + ":BM" + i.ToString());
                b.Copy(a);
            }

            //20240124經費相關欄位最後增加統計列
            double Total_B = 0; //分支計畫核定經費總計(仟元)
            double Total_H = 0; //歷年累積撥發金額(仟元)
            double Total_I = 0; //當年度核定經費(仟元)(A)
            double Total_J = 0; //撥款金額(仟元)(B)
            double Total_L = 0; //累積分配數(仟元)©
            double Total_M = 0; //實際支用數(仟元)(D)
            double Total_N = 0; //應付未付數(仟元)(E)
            double Total_O = 0; //節餘數(仟元)(F)
            double Total_P = 0; //D+E+F
            double Total_AK = 0;//總經費
            double Total_AL = 0;//施工費
            double Total_AM = 0;//間接費用
            double Total_AN = 0;//總經費
            double Total_AO = 0;//施工費

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string TD_CODE = dt.Rows[i]["TD_CODE"].ToString();
                int rowNumber = 5 + i;

                ws1.Cells["A" + rowNumber].PutValue(dt.Rows[i]["CATY_NAME"].ToString());
                ws1.Cells["B" + rowNumber].PutValue(Source.Tools.thousandchange1(dt.Rows[i]["L_UNPAY_COST"].ToString()));
                ws1.Cells["C" + rowNumber].PutValue(dt.Rows[i]["SEQ"].ToString());
                ws1.Cells["D" + rowNumber].PutValue(dt.Rows[i]["TD_CITY"].ToString());
                ws1.Cells["E" + rowNumber].PutValue(dt.Rows[i]["TD_NAME"].ToString());
                // ws1.Cells["F" + rowNumber].PutValue(dt.Rows[i]["IPJ_DESC"].ToString()); // 如果需要使用此行，取消註解並更改為 PutValue()
                ws1.Cells["G" + rowNumber].PutValue(dt.Rows[i]["CATY_CODE_DESC"].ToString());
                ws1.Cells["H" + rowNumber].PutValue(Source.Tools.thousandchange1(dt.Rows[i]["ALL_PPAY_FEE"].ToString()));
                double I = Source.Tools.thousandchange1(dt.Rows[i]["YR_FUND"].ToString());
                ws1.Cells["I" + rowNumber].PutValue(I);
                double J = Source.Tools.thousandchange1(dt.Rows[i]["PPAY_FEE"].ToString());
                ws1.Cells["J" + rowNumber].PutValue(J);

                double KValue;
                if (I == 0 || J == 0)
                {
                    KValue = 0;
                }
                else
                {
                    KValue = J / I;
                }
                ws1.Cells["K" + rowNumber].PutValue(KValue);

                double L = Source.Tools.thousandchange1(dt.Rows[i]["L_PLN_COST"].ToString());
                ws1.Cells["L" + rowNumber].PutValue(L);

                double M = Source.Tools.thousandchange1(dt.Rows[i]["L_REL_COST"].ToString());
                ws1.Cells["M" + rowNumber].PutValue(M);

                double N = Source.Tools.thousandchange1(dt.Rows[i]["L_UNPAY_COST"].ToString());
                ws1.Cells["N" + rowNumber].PutValue(N);

                double O = Source.Tools.thousandchange1(dt.Rows[i]["SAVE_FEE"].ToString());
                ws1.Cells["O" + rowNumber].PutValue(O);

                double P;
                if (M == 0 && N == 0 && O == 0)
                {
                    ws1.Cells["P" + (5 + i)].PutValue(null);
                    P = 0;
                }
                else
                {
                    double M1 = Source.Tools.IsNumberOrNot(M);
                    double N1 = Source.Tools.IsNumberOrNot(N);
                    double O1 = Source.Tools.IsNumberOrNot(O);
                    P = M1 + N1 + O1;
                    ws1.Cells["P" + (5 + i)].PutValue(P);
                }

                if (L == 0 || P == 0)
                {
                    ws1.Cells["Q" + (5 + i)].PutValue(null);
                }
                else
                {
                    double M2 = Source.Tools.IsNumberOrNot(M);
                    double N2 = Source.Tools.IsNumberOrNot(N);
                    double O2 = Source.Tools.IsNumberOrNot(O);
                    ws1.Cells["Q" + (5 + i)].PutValue((M2 + N2 + O2) / L);
                }

                if (I == 0 || P == 0)
                {
                    ws1.Cells["R" + (5 + i)].PutValue(null);
                }
                else
                {
                    double M2 = Source.Tools.IsNumberOrNot(M);
                    double N2 = Source.Tools.IsNumberOrNot(N);
                    double O2 = Source.Tools.IsNumberOrNot(O);
                    ws1.Cells["R" + (5 + i)].PutValue((M2 + N2 + O2) / I);
                }

                string S = dt.Rows[i]["PRE_TK102"].ToString();
                if (string.IsNullOrEmpty(S))
                {
                    ws1.Cells["S" + (5 + i)].PutValue("-");
                }
                else
                {
                    ws1.Cells["S" + (5 + i)].PutValue(Source.Tools.FormatDate(S));
                }

                string T = dt.Rows[i]["REL_TK102"].ToString();
                if (string.IsNullOrEmpty(T))
                {
                    ws1.Cells["T" + (5 + i)].PutValue("-");
                }
                else
                {
                    ws1.Cells["T" + (5 + i)].PutValue(Source.Tools.FormatDate(T));
                }

                string U = dt.Rows[i]["PRE_TK103"].ToString();
                if (string.IsNullOrEmpty(U))
                {
                    ws1.Cells["U" + (5 + i)].PutValue("-");
                }
                else
                {
                    ws1.Cells["U" + (5 + i)].PutValue(Source.Tools.FormatDate(U));
                }

                string V = dt.Rows[i]["REL_TK103"].ToString();
                string Z = dt.Rows[i]["REL_TK105"].ToString();

                if (string.IsNullOrEmpty(V))
                {
                    if (string.IsNullOrEmpty(Z))
                    {
                        ws1.Cells["V" + (5 + i)].PutValue(null);
                    }
                    else
                    {
                        ws1.Cells["V" + (5 + i)].PutValue("-");
                    }
                }
                else
                {
                    ws1.Cells["V" + (5 + i)].PutValue(Source.Tools.FormatDate(V));
                }

                string W = dt.Rows[i]["PRE_TK1031"].ToString();
                if (string.IsNullOrEmpty(W))
                {
                    ws1.Cells["W" + (5 + i)].PutValue("-");
                }
                else
                {
                    ws1.Cells["W" + (5 + i)].PutValue(Source.Tools.FormatDate(W));
                }

                string X = dt.Rows[i]["REL_TK1031"].ToString(); //X
                ws1.Cells["X" + (5 + i)].PutValue(string.IsNullOrEmpty(X) ? "-" : Source.Tools.FormatDate(X));

                string Y = dt.Rows[i]["PRE_TK105"].ToString(); //Y
                ws1.Cells["Y" + (5 + i)].PutValue(string.IsNullOrEmpty(Y) ? "-" : Source.Tools.FormatDate(Y));
                ws1.Cells["Z" + (5 + i)].PutValue(string.IsNullOrEmpty(Z) ? "-" : Source.Tools.FormatDate(Z));

                ws1.Cells["AA" + (5 + i)].PutValue(dt.Rows[i]["TK105_TYPE"].ToString());
                ws1.Cells["AC" + (5 + i)].PutValue(dt.Rows[i]["BUSS_NAME"].ToString());
                ws1.Cells["AB" + (5 + i)].PutValue(dt.Rows[i]["DSG_BUSS_NAME"].ToString());

                string AG = dt.Rows[i]["PRE_TK107"].ToString(); //AG
                ws1.Cells["AG" + (5 + i)].PutValue(AG);

                string AH = dt.Rows[i]["REL_TK107"].ToString(); //AH
                ws1.Cells["AH" + (5 + i)].PutValue(AH);

                ws1.Cells["AD" + (5 + i)].PutValue(string.IsNullOrEmpty(AG) ? "" : Source.Tools.FormatDate(AG));
                ws1.Cells["AE" + (5 + i)].PutValue(string.IsNullOrEmpty(AH) ? "" : Source.Tools.FormatDate(AH));

                string PRE_SCHD = GetTDDMSCD_SCHD(YearWrite, SMN, TD_CODE, "PRE_SCHD");
                string REL_SCHD = GetTDDMSCD_SCHD(YearWrite, SMN, TD_CODE, "REL_SCHD");
                ws1.Cells["AF" + (5 + i)].PutValue(PRE_SCHD);
                ws1.Cells["AG" + (5 + i)].PutValue(REL_SCHD);

                ws1.Cells["AI" + (5 + i)].PutValue(dt.Rows[i]["EXECT_DESC"].ToString());
                ws1.Cells["AJ" + (5 + i)].PutValue(dt.Rows[i]["DELAY_RESN"].ToString());

                ws1.Cells["AK" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["B_FUND"].ToString()));
                ws1.Cells["AL" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["B_0101_FUND"].ToString()));
                ws1.Cells["AM" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["B_0114_FUND"].ToString()));

                ws1.Cells["AN" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["C_FUND"].ToString()));
                ws1.Cells["AO" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["C_0101_FUND"].ToString()));

                ws1.Cells["AP" + (5 + i)].PutValue(dt.Rows[i]["CTDELAY_RESN"].ToString());
                //-----1070615二課調整---end
                ws1.Cells["AP" + (5 + i) + ""].Value = dt.Rows[i]["CTDELAY_RESN"].ToString();

                Total_B += Source.Tools.thousandchange1(dt.Rows[i]["L_UNPAY_COST"].ToString());
                Total_H += Source.Tools.thousandchange1(dt.Rows[i]["ALL_PPAY_FEE"].ToString());
                Total_I += I;
                Total_J += J;
                Total_L += L;
                Total_M += M;
                Total_N += N;
                Total_O += O;
                Total_P += 0;
                Total_AK += Source.Tools.StringToDouble(dt.Rows[i]["B_FUND"].ToString());
                Total_AL += Source.Tools.StringToDouble(dt.Rows[i]["B_0101_FUND"].ToString());
                Total_AM += Source.Tools.StringToDouble(dt.Rows[i]["B_0114_FUND"].ToString());
                Total_AN += Source.Tools.StringToDouble(dt.Rows[i]["C_FUND"].ToString());
                Total_AO += Source.Tools.StringToDouble(dt.Rows[i]["C_0101_FUND"].ToString());
                if (dt.Rows.Count > 0 && i == dt.Rows.Count - 1)
                {
                    i += 1;
                    ws1.Cells["A" + (5 + i)].PutValue("合計");
                    ws1.Cells["B" + (5 + i)].PutValue(Total_B);
                    ws1.Cells["H" + (5 + i)].PutValue(Total_H);
                    ws1.Cells["I" + (5 + i)].PutValue(Total_I);
                    ws1.Cells["J" + (5 + i)].PutValue(Total_J);
                    ws1.Cells["L" + (5 + i)].PutValue(Total_L);
                    ws1.Cells["M" + (5 + i)].PutValue(Total_M);
                    ws1.Cells["N" + (5 + i)].PutValue(Total_N);
                    ws1.Cells["O" + (5 + i)].PutValue(Total_O);
                    ws1.Cells["P" + (5 + i)].PutValue(Total_M + Total_N + Total_O);
                    ws1.Cells["AK" + (5 + i)].PutValue(Total_AK);
                    ws1.Cells["AL" + (5 + i)].PutValue(Total_AL);
                    ws1.Cells["AM" + (5 + i)].PutValue(Total_AM);
                    ws1.Cells["AN" + (5 + i)].PutValue(Total_AN);
                    ws1.Cells["AO" + (5 + i)].PutValue(Total_AO);
                }
            }

            ws1.Name = "細項執行情形";

            #endregion

            #region ======================== Sheet[細項執行情形(用戶接管)] ========================

            // 文件標題
            //ws2.Cells["A1"].PutValue("表2-5、內政部辦理" + PJ + "細項執行情形");
            ws2.Cells["A1"].PutValue("表2-5_內政部辦理細項執行情形");
            ws2.Cells["AI2"].PutValue("資料截至：" + YearWrite + "年" + SMN + "月");

            // 複製範圍到特定行
            Aspose.Cells.Range a2 = ws2.Cells.CreateRange("A6:BR6");
            for (int i = 7; i < dt.Rows.Count + 6; i++)
            {
                Aspose.Cells.Range b2 = ws2.Cells.CreateRange("A" + i.ToString() + ":BR" + i.ToString());
                b2.Copy(a2);
            }

            Total_B = 0;
            Total_H = 0;
            Total_I = 0;
            Total_J = 0;
            Total_L = 0;
            Total_M = 0;
            Total_N = 0;
            Total_O = 0;
            Total_P = 0;
            Total_AK = 0;
            Total_AL = 0;
            Total_AM = 0;
            Total_AN = 0;
            Total_AO = 0;
            double Total_AP = 0;
            double Total_AQ = 0;
            double Total_AR = 0;
            double Total_AS = 0;
            double Total_AT = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string TD_CODE = dt.Rows[i]["TD_CODE"].ToString();
                ws2.Cells["A" + (6 + i)].PutValue(dt.Rows[i]["CATY_NAME"].ToString());
                ws2.Cells["B" + (6 + i)].PutValue(Source.Tools.thousandchange1(dt.Rows[i]["L_UNPAY_COST"].ToString())); // N
                ws2.Cells["D" + (6 + i)].PutValue(dt.Rows[i]["TD_CITY"].ToString());
                ws2.Cells["C" + (6 + i)].PutValue(dt.Rows[i]["SEQ"].ToString());
                ws2.Cells["E" + (6 + i)].PutValue(dt.Rows[i]["TD_NAME"].ToString());
                // ws1.Cells["F" + (5 + i)].PutValue(dt.Rows[i]["IPJ_DESC"].ToString()); 
                ws2.Cells["G" + (6 + i)].PutValue(dt.Rows[i]["CATY_CODE_DESC"].ToString());
                ws2.Cells["H" + (6 + i)].PutValue(Source.Tools.thousandchange1(dt.Rows[i]["ALL_PPAY_FEE"].ToString())); // H
                double I = Source.Tools.thousandchange1(dt.Rows[i]["YR_FUND"].ToString());
                ws2.Cells["I" + (6 + i)].PutValue(I); // I
                double J = Source.Tools.thousandchange1(dt.Rows[i]["PPAY_FEE"].ToString());
                ws2.Cells["J" + (6 + i)].PutValue(J); // J

                if (I == 0 || J == 0)
                {
                    ws2.Cells["K" + (6 + i)].PutValue(null);
                }
                else
                {
                    ws2.Cells["K" + (6 + i)].PutValue(J / I);
                }

                double L = Source.Tools.thousandchange1(dt.Rows[i]["L_PLN_COST"].ToString());
                ws2.Cells["L" + (6 + i)].PutValue(L); // L
                double M = Source.Tools.thousandchange1(dt.Rows[i]["L_REL_COST"].ToString());
                ws2.Cells["M" + (6 + i)].PutValue(M); // M
                double N = Source.Tools.thousandchange1(dt.Rows[i]["L_UNPAY_COST"].ToString());
                ws2.Cells["N" + (6 + i)].PutValue(N); // N
                double O = Source.Tools.thousandchange1(dt.Rows[i]["SAVE_FEE"].ToString());
                ws2.Cells["O" + (6 + i)].PutValue(O); // O

                double P;
                if (M == 0 && N == 0 && O == 0)
                {
                    ws2.Cells["P" + (6 + i)].PutValue(null);
                    P = 0;
                }
                else
                {
                    double M1 = Source.Tools.IsNumberOrNot(M);
                    double N1 = Source.Tools.IsNumberOrNot(N);
                    double O1 = Source.Tools.IsNumberOrNot(O);
                    P = M1 + N1 + O1;
                    ws2.Cells["P" + (6 + i)].PutValue(P);
                }

                if (L == 0 || P == 0)
                {
                    ws2.Cells["Q" + (6 + i)].PutValue(null);
                }
                else
                {
                    double M2 = Source.Tools.IsNumberOrNot(M);
                    double N2 = Source.Tools.IsNumberOrNot(N);
                    double O2 = Source.Tools.IsNumberOrNot(O);
                    ws2.Cells["Q" + (6 + i)].PutValue((M2 + N2 + O2) / L);
                }

                if (I == 0 || P == 0)
                {
                    ws2.Cells["R" + (6 + i)].PutValue(null);
                }
                else
                {
                    double M2 = Source.Tools.IsNumberOrNot(M);
                    double N2 = Source.Tools.IsNumberOrNot(N);
                    double O2 = Source.Tools.IsNumberOrNot(O);
                    ws2.Cells["R" + (6 + i)].PutValue((M2 + N2 + O2) / I);
                }
                string S = dt.Rows[i]["PRE_TK102"].ToString();//S
                ws2.Cells["S" + (6 + i)].PutValue(string.IsNullOrEmpty(S) ? "-" : Source.Tools.FormatDate(S));

                string T = dt.Rows[i]["REL_TK102"].ToString();//T
                ws2.Cells["T" + (6 + i)].PutValue(string.IsNullOrEmpty(T) ? "-" : Source.Tools.FormatDate(T));

                string U = dt.Rows[i]["PRE_TK103"].ToString();//U
                ws2.Cells["U" + (6 + i)].PutValue(string.IsNullOrEmpty(U) ? "-" : Source.Tools.FormatDate(U));

                string V = dt.Rows[i]["REL_TK103"].ToString();//V
                string Z = dt.Rows[i]["REL_TK105"].ToString();//Z

                if (string.IsNullOrEmpty(V))
                {
                    ws2.Cells["V" + (6 + i)].PutValue(string.IsNullOrEmpty(Z) ? (object)null : "-");
                }
                else
                {
                    ws2.Cells["V" + (6 + i)].PutValue(Source.Tools.FormatDate(V));
                }

                string W = dt.Rows[i]["PRE_TK1031"].ToString();//W
                ws2.Cells["W" + (6 + i)].PutValue(string.IsNullOrEmpty(W) ? "-" : Source.Tools.FormatDate(W));

                string X = dt.Rows[i]["REL_TK1031"].ToString();//X
                ws2.Cells["X" + (6 + i)].PutValue(string.IsNullOrEmpty(X) ? "-" : Source.Tools.FormatDate(X));

                string Y = dt.Rows[i]["PRE_TK105"].ToString();//Y
                ws2.Cells["Y" + (6 + i)].PutValue(string.IsNullOrEmpty(Y) ? "-" : Source.Tools.FormatDate(Y));

                if (string.IsNullOrEmpty(Z))
                {
                    ws2.Cells["Z" + (6 + i)].PutValue("-");
                }
                else
                {
                    ws2.Cells["Z" + (6 + i)].PutValue(Source.Tools.FormatDate(Z));
                }

                ws2.Cells["AA" + (6 + i)].PutValue(dt.Rows[i]["TK105_TYPE"].ToString());
                ws2.Cells["AB" + (6 + i)].PutValue(dt.Rows[i]["DSG_BUSS_NAME"].ToString());
                ws2.Cells["AC" + (6 + i)].PutValue(dt.Rows[i]["BUSS_NAME"].ToString());

                string AG = dt.Rows[i]["PRE_TK107"].ToString();//AG
                ws2.Cells["AG" + (6 + i)].PutValue(AG);

                string AH = dt.Rows[i]["REL_TK107"].ToString();//AH
                ws2.Cells["AH" + (6 + i)].PutValue(AH);

                if (string.IsNullOrEmpty(AG))
                {
                    ws2.Cells["AD" + (6 + i)].PutValue("");
                }
                else
                {
                    ws2.Cells["AD" + (6 + i)].PutValue(Source.Tools.FormatDate(AG));
                }

                if (string.IsNullOrEmpty(AH))
                {
                    ws2.Cells["AE" + (6 + i)].PutValue("");
                }
                else
                {
                    ws2.Cells["AE" + (6 + i)].PutValue(Source.Tools.FormatDate(AH));
                }

                // ======================== 2018/06/15二課調整 ========================

                string PRE_SCHD = GetTDDMSCD_SCHD(YearWrite, SMN, TD_CODE, "PRE_SCHD");
                string REL_SCHD = GetTDDMSCD_SCHD(YearWrite, SMN, TD_CODE, "REL_SCHD");
                ws2.Cells["AF" + (6 + i)].PutValue(PRE_SCHD);
                ws2.Cells["AG" + (6 + i)].PutValue(REL_SCHD);

                ws2.Cells["AI" + (6 + i)].PutValue(dt.Rows[i]["EXECT_DESC"].ToString());
                ws2.Cells["AJ" + (6 + i)].PutValue(dt.Rows[i]["DELAY_RESN"].ToString());

                ws2.Cells["AM" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["B_0114_FUND"].ToString()));
                ws2.Cells["AL" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["B_0101_FUND"].ToString()));
                ws2.Cells["AK" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["B_FUND"].ToString()));
                ws2.Cells["AN" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["B_FUND_07"].ToString()));
                ws2.Cells["AO" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["B_0701_FUND"].ToString()));
                ws2.Cells["AP" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["B_0702_FUND"].ToString()));

                ws2.Cells["AQ" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["C_FUND"].ToString()));
                ws2.Cells["AR" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["C_0101_FUND"].ToString()));
                ws2.Cells["AS" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["C_FUND_07"].ToString()));
                ws2.Cells["AT" + (6 + i)].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["C_0701_FUND"].ToString()));

                // ======================== 2018/06/15二課調整 ========================

                ws2.Cells["AU" + (6 + i)].PutValue(dt.Rows[i]["CTDELAY_RESN"].ToString());
                Total_B += Source.Tools.thousandchange1(dt.Rows[i]["L_UNPAY_COST"].ToString());
                Total_H += Source.Tools.thousandchange1(dt.Rows[i]["ALL_PPAY_FEE"].ToString());
                Total_I += I;
                Total_J += J;
                Total_L += L;
                Total_M += M;
                Total_N += N;
                Total_O += O;
                Total_P += 0;
                Total_AK += Source.Tools.StringToDouble(dt.Rows[i]["B_FUND"].ToString());
                Total_AL += Source.Tools.StringToDouble(dt.Rows[i]["B_0101_FUND"].ToString());
                Total_AM += Source.Tools.StringToDouble(dt.Rows[i]["B_0114_FUND"].ToString());
                Total_AN += Source.Tools.StringToDouble(dt.Rows[i]["C_FUND"].ToString());
                Total_AO += Source.Tools.StringToDouble(dt.Rows[i]["C_0101_FUND"].ToString());
                Total_AP += Source.Tools.StringToDouble(dt.Rows[i]["B_0702_FUND"].ToString());
                Total_AQ += Source.Tools.StringToDouble(dt.Rows[i]["C_FUND"].ToString());
                Total_AR += Source.Tools.StringToDouble(dt.Rows[i]["C_0101_FUND"].ToString());
                Total_AS += Source.Tools.StringToDouble(dt.Rows[i]["C_FUND_07"].ToString());
                Total_AT += Source.Tools.StringToDouble(dt.Rows[i]["C_0701_FUND"].ToString());

                if (dt.Rows.Count > 0 && i == dt.Rows.Count - 1)
                {
                    i += 1;
                    ws2.Cells["A" + (6 + i)].PutValue("合計");
                    ws2.Cells["B" + (6 + i)].PutValue(Total_B);
                    ws2.Cells["H" + (6 + i)].PutValue(Total_H);
                    ws2.Cells["I" + (6 + i)].PutValue(Total_I);
                    ws2.Cells["J" + (6 + i)].PutValue(Total_J);
                    ws2.Cells["L" + (6 + i)].PutValue(Total_L);
                    ws2.Cells["M" + (6 + i)].PutValue(Total_M);
                    ws2.Cells["N" + (6 + i)].PutValue(Total_N);
                    ws2.Cells["O" + (6 + i)].PutValue(Total_O);
                    ws2.Cells["P" + (6 + i)].PutValue(Total_M + Total_N + Total_O);
                    ws2.Cells["AK" + (6 + i)].PutValue(Total_AK);
                    ws2.Cells["AL" + (6 + i)].PutValue(Total_AL);
                    ws2.Cells["AM" + (6 + i)].PutValue(Total_AM);
                    ws2.Cells["AN" + (6 + i)].PutValue(Total_AN);
                    ws2.Cells["AO" + (6 + i)].PutValue(Total_AO);
                    ws2.Cells["AP" + (6 + i)].PutValue(Total_AP);
                    ws2.Cells["AQ" + (6 + i)].PutValue(Total_AQ);
                    ws2.Cells["AR" + (6 + i)].PutValue(Total_AR);
                    ws2.Cells["AS" + (6 + i)].PutValue(Total_AS);
                    ws2.Cells["AT" + (6 + i)].PutValue(Total_AT);
                }
            }

            ws2.Name = "細項執行情形(用戶接管)";

            #endregion

            #region ======================== Sheet[細項執行情形(依縣市)] ========================

            // 文件標題
            // ws3.Cells["A1"].PutValue("表A-5、內政部辦理" + PJ + "細項執行情形");
            ws3.Cells["A1"].PutValue("表A-5_內政部辦理細項執行情形");
            ws3.Cells["AI2"].PutValue("資料截至：" + YearWrite + "年" + SMN + "月");

            DataTable SMCPOST = SourceFunction.DataTableGet("select distinct CPA_CITY_ID,CITY from CMBASE.dbo.SMCPOSTN order by CPA_CITY_ID", connectionString);
            DataTable dt1 = dt.Clone();
            DataRow[] rr;

            for (int i = 0; i < SMCPOST.Rows.Count; i++)
            {
                switch (SMCPOST.Rows[i]["CITY"].ToString())
                {
                    case "新北市":
                        rr = dt.Select("TD_CITY in ('台北縣','新北市')");
                        break;
                    case "桃園市":
                        rr = dt.Select("TD_CITY in ('桃園縣','桃園市')");
                        break;
                    case "台中市":
                        rr = dt.Select("TD_CITY in ('台中縣','台中市')");
                        break;
                    case "台南市":
                        rr = dt.Select("TD_CITY in ('台南縣','台南市')");
                        break;
                    case "高雄市":
                        rr = dt.Select("TD_CITY in ('高雄縣','高雄市')");
                        break;
                    default:
                        rr = dt.Select("TD_CITY='" + SMCPOST.Rows[i]["CITY"].ToString() + "'");
                        break;
                }

                if (rr.Length > 0)
                {
                    DataRow dr = dt1.NewRow();
                    dr["TD_CODE"] = "TD_CODE";
                    dr["TD_CITY"] = SMCPOST.Rows[i]["CITY"].ToString();
                    dt1.Rows.Add(dr);

                    foreach (DataRow row in rr)
                    {
                        dt1.ImportRow(row);
                    }
                }
            }

            Total_B = 0;
            Total_H = 0;
            Total_I = 0;
            Total_J = 0;
            Total_L = 0;
            Total_M = 0;
            Total_N = 0;
            Total_O = 0;
            Total_P = 0;
            Total_AK = 0;
            Total_AL = 0;
            Total_AM = 0;
            Total_AN = 0;
            Total_AO = 0;

            var a3 = ws3.Cells.CreateRange("A5:BM5");

            for (int i = 6; i < dt1.Rows.Count + 5; i++)
            {
                var b3 = ws3.Cells.CreateRange("A" + i.ToString() + ":BM" + i.ToString() + "");
                b3.Copy(a3);
            }

            int line = 5;
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                switch (dt1.Rows[i]["TD_CODE"].ToString())
                {
                    case "TD_CODE":
                        Aspose.Cells.Range range = ws3.Cells.CreateRange(line + i - 1, 1, 1, 41);
                        range.Merge();
                        ws3.Cells["A" + (line + i)].PutValue(dt1.Rows[i]["TD_CITY"].ToString());
                        break;
                    default:
                        string TD_CODE = dt1.Rows[i]["TD_CODE"].ToString();
                        ws3.Cells["A" + (5 + i)].PutValue(dt1.Rows[i]["CATY_NAME"].ToString());
                        ws3.Cells["B" + (5 + i)].PutValue(Source.Tools.thousandchange1(dt1.Rows[i]["L_UNPAY_COST"].ToString())); //N
                        ws3.Cells["C" + (5 + i)].PutValue(dt1.Rows[i]["SEQ"].ToString());
                        ws3.Cells["D" + (5 + i)].PutValue(dt1.Rows[i]["TD_CITY"].ToString());
                        ws3.Cells["E" + (5 + i)].PutValue(dt1.Rows[i]["TD_NAME"].ToString());
                        //ws3.Cells["F" + (5 + i)].PutValue(dt1.Rows[i]["IPJ_DESC"].ToString());
                        ws3.Cells["G" + (5 + i)].PutValue(dt1.Rows[i]["CATY_CODE_DESC"].ToString());
                        ws3.Cells["H" + (5 + i)].PutValue(Source.Tools.thousandchange1(dt1.Rows[i]["ALL_PPAY_FEE"].ToString())); //H
                        double I = Source.Tools.thousandchange1(dt1.Rows[i]["YR_FUND"].ToString());
                        ws3.Cells["I" + (5 + i)].PutValue(I); //I
                        double J = Source.Tools.thousandchange1(dt1.Rows[i]["PPAY_FEE"].ToString());
                        ws3.Cells["J" + (5 + i)].PutValue(J); //J

                        if (I == 0 || J == 0)
                        {
                            ws3.Cells["K" + (5 + i)].PutValue((object)null);
                        }
                        else if (I == 0)
                        {
                            ws3.Cells["K" + (5 + i)].PutValue((object)null);
                        }
                        else
                        {
                            ws3.Cells["K" + (5 + i)].PutValue(J / I);
                        }

                        double L = Source.Tools.thousandchange1(dt1.Rows[i]["L_PLN_COST"].ToString());
                        ws3.Cells["L" + (5 + i)].PutValue(L); //L
                        double M = Source.Tools.thousandchange1(dt1.Rows[i]["L_REL_COST"].ToString());
                        ws3.Cells["M" + (5 + i)].PutValue(M); //M
                        double N = Source.Tools.thousandchange1(dt1.Rows[i]["L_UNPAY_COST"].ToString());
                        ws3.Cells["N" + (5 + i)].PutValue(N); //N
                        double O = Source.Tools.thousandchange1(dt1.Rows[i]["SAVE_FEE"].ToString());
                        ws3.Cells["O" + (5 + i)].PutValue(O); //O

                        double P;
                        if (M == 0 && N == 0 && O == 0)
                        {
                            ws3.Cells["P" + (5 + i)].PutValue((object)null);
                            P = 0;
                        }
                        else
                        {
                            double M1 = Source.Tools.IsNumberOrNot(M);
                            double N1 = Source.Tools.IsNumberOrNot(N);
                            double O1 = Source.Tools.IsNumberOrNot(O);
                            P = M1 + N1 + O1;
                            ws3.Cells["P" + (5 + i)].PutValue(P);
                        }

                        if (L == 0 || P == 0)
                        {
                            ws3.Cells["Q" + (5 + i)].PutValue((object)null);
                        }
                        else
                        {
                            double M2 = Source.Tools.IsNumberOrNot(M);
                            double N2 = Source.Tools.IsNumberOrNot(N);
                            double O2 = Source.Tools.IsNumberOrNot(O);
                            ws3.Cells["Q" + (5 + i)].PutValue((M2 + N2 + O2) / L);
                        }

                        if (I == 0 || P == 0)
                        {
                            ws3.Cells["R" + (5 + i)].PutValue((object)null);
                        }
                        else
                        {
                            double M2 = Source.Tools.IsNumberOrNot(M);
                            double N2 = Source.Tools.IsNumberOrNot(N);
                            double O2 = Source.Tools.IsNumberOrNot(O);
                            ws3.Cells["R" + (5 + i)].PutValue((M2 + N2 + O2) / I);
                        }

                        string S = dt1.Rows[i]["PRE_TK102"].ToString(); // S
                        ws3.Cells["S" + (5 + i)].PutValue(string.IsNullOrEmpty(S) ? "-" : Source.Tools.FormatDate(S));

                        string T = dt1.Rows[i]["REL_TK102"].ToString(); // T
                        ws3.Cells["T" + (5 + i)].PutValue(string.IsNullOrEmpty(T) ? "-" : Source.Tools.FormatDate(T));
                        string U = dt1.Rows[i]["PRE_TK103"].ToString(); // U
                        ws3.Cells["U" + (5 + i)].PutValue(string.IsNullOrEmpty(U) ? "-" : Source.Tools.FormatDate(U));

                        string V = dt1.Rows[i]["REL_TK103"].ToString(); // V
                        string Z = dt1.Rows[i]["REL_TK105"].ToString(); // Z

                        if (string.IsNullOrEmpty(V))
                        {
                            ws3.Cells["V" + (5 + i)].PutValue(string.IsNullOrEmpty(Z) ? (object)null : "-");
                        }
                        else
                        {
                            ws3.Cells["V" + (5 + i)].PutValue(Source.Tools.FormatDate(V));
                        }

                        string W = dt1.Rows[i]["PRE_TK1031"].ToString(); // W
                        ws3.Cells["W" + (5 + i)].PutValue(string.IsNullOrEmpty(W) ? "-" : Source.Tools.FormatDate(W));

                        string X = dt1.Rows[i]["REL_TK1031"].ToString(); // X
                        ws3.Cells["X" + (5 + i)].PutValue(string.IsNullOrEmpty(X) ? "-" : Source.Tools.FormatDate(X));

                        string Y = dt1.Rows[i]["PRE_TK105"].ToString(); // Y
                        ws3.Cells["Y" + (5 + i)].PutValue(string.IsNullOrEmpty(Y) ? "-" : Source.Tools.FormatDate(Y));

                        string AA = dt1.Rows[i]["TK105_TYPE"].ToString(); // AA
                        ws3.Cells["AA" + (5 + i)].PutValue(AA);

                        string AC = dt1.Rows[i]["BUSS_NAME"].ToString(); // AC
                        ws3.Cells["AC" + (5 + i)].PutValue(AC);

                        string AB = dt1.Rows[i]["DSG_BUSS_NAME"].ToString(); // AB
                        ws3.Cells["AB" + (5 + i)].PutValue(AB);

                        string AG = dt1.Rows[i]["PRE_TK107"].ToString(); // AG
                        ws3.Cells["AG" + (5 + i)].PutValue(AG);

                        string AH = dt1.Rows[i]["REL_TK107"].ToString(); // AH
                        ws3.Cells["AH" + (5 + i)].PutValue(AH);

                        ws3.Cells["AD" + (5 + i)].PutValue(string.IsNullOrEmpty(AG) ? "" : Source.Tools.FormatDate(AG));
                        ws3.Cells["AE" + (5 + i)].PutValue(string.IsNullOrEmpty(AH) ? "" : Source.Tools.FormatDate(AH));

                        string PRE_SCHD = GetTDDMSCD_SCHD(YearWrite, SMN, TD_CODE, "PRE_SCHD");
                        string REL_SCHD = GetTDDMSCD_SCHD(YearWrite, SMN, TD_CODE, "REL_SCHD");
                        ws3.Cells["AF" + (5 + i)].PutValue(PRE_SCHD);
                        ws3.Cells["AG" + (5 + i)].PutValue(REL_SCHD);

                        ws3.Cells["AI" + (5 + i)].PutValue(dt1.Rows[i]["EXECT_DESC"].ToString());
                        ws3.Cells["AJ" + (5 + i)].PutValue(dt1.Rows[i]["DELAY_RESN"].ToString());

                        ws3.Cells["AK" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt1.Rows[i]["B_FUND"].ToString()));
                        ws3.Cells["AL" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt1.Rows[i]["B_0101_FUND"].ToString()));
                        ws3.Cells["AM" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt1.Rows[i]["B_0114_FUND"].ToString()));

                        ws3.Cells["AN" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt1.Rows[i]["C_FUND"].ToString()));
                        ws3.Cells["AO" + (5 + i)].PutValue(Source.Tools.StringToDouble(dt1.Rows[i]["C_0101_FUND"].ToString()));

                        ws3.Cells["AP" + (5 + i)].PutValue(dt1.Rows[i]["CTDELAY_RESN"].ToString());

                        Total_B += Source.Tools.thousandchange1(dt1.Rows[i]["L_UNPAY_COST"].ToString());
                        Total_H += Source.Tools.thousandchange1(dt1.Rows[i]["ALL_PPAY_FEE"].ToString());
                        Total_I += I;
                        Total_J += J;
                        Total_L += L;
                        Total_M += M;
                        Total_N += N;
                        Total_O += O;
                        Total_P += 0;
                        Total_AK += Source.Tools.StringToDouble(dt1.Rows[i]["B_FUND"].ToString());
                        Total_AL += Source.Tools.StringToDouble(dt1.Rows[i]["B_0101_FUND"].ToString());
                        Total_AM += Source.Tools.StringToDouble(dt1.Rows[i]["B_0114_FUND"].ToString());
                        Total_AN += Source.Tools.StringToDouble(dt1.Rows[i]["C_FUND"].ToString());
                        Total_AO += Source.Tools.StringToDouble(dt1.Rows[i]["C_0101_FUND"].ToString());

                        if (dt1.Rows.Count > 0 && i == dt1.Rows.Count - 1)
                        {
                            i += 1;
                            ws3.Cells["A" + (5 + i)].PutValue("合計");
                            ws3.Cells["B" + (5 + i)].PutValue(Total_B);
                            ws3.Cells["H" + (5 + i)].PutValue(Total_H);
                            ws3.Cells["I" + (5 + i)].PutValue(Total_I);
                            ws3.Cells["J" + (5 + i)].PutValue(Total_J);
                            ws3.Cells["L" + (5 + i)].PutValue(Total_L);
                            ws3.Cells["M" + (5 + i)].PutValue(Total_M);
                            ws3.Cells["N" + (5 + i)].PutValue(Total_N);
                            ws3.Cells["O" + (5 + i)].PutValue(Total_O);
                            ws3.Cells["P" + (5 + i)].PutValue(Total_M + Total_N + Total_O);
                            ws3.Cells["AK" + (5 + i)].PutValue(Total_AK);
                            ws3.Cells["AL" + (5 + i)].PutValue(Total_AL);
                            ws3.Cells["AM" + (5 + i)].PutValue(Total_AM);
                            ws3.Cells["AN" + (5 + i)].PutValue(Total_AN);
                            ws3.Cells["AO" + (5 + i)].PutValue(Total_AO);
                        }
                        break;

                }
            }

            ws3.Name = "細項執行情形(依縣市)";

            #endregion

            string filename = "細項執行情形表(前瞻計畫)" + DateTime.Now.ToString("yyyyMMddHHmmss");
            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");

            workbook.Save(savepath, SaveFormat.Xlsx); // 保存为 Excel 文件

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

        /// <summary>
        /// A-2.細項執行情形表
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="YearWrite"></param>
        /// <param name="SMN"></param>
        /// <param name="CityList"></param>
        /// <param name="funding"></param>
        /// <param name="ODS"></param>
        /// <param name="Site"></param>
        /// <param name="group"></param>
        /// <returns></returns>
        public ActionResult ReportPrint2(string PID, string YearWrite, string SMN, string[] CityList, string funding, string ODS, string[] Site, string group)
        {
            var path = Server.MapPath("~/Templates/PJV0010_03.xlsx");
            var file = new FileInfo(path);
            string City = "";
            DataTable dt2 = new DataTable();

            if (CityList == null || CityList.Length == 0)
            {
                string alertMessage = "請至少選擇一個縣市";
                return Content($"<script>alert('{alertMessage}'); window.location.href = document.referrer;</script>");
            }

            foreach (var item in CityList)
            {
                City = item;
                dt2.Merge(Source.Report_Source.getBillionProject08(PID, YearWrite, SMN, City, funding, connectionString));
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

            // ========================================================================================================

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            var workbook = new Aspose.Cells.Workbook(path);
            var ws = workbook.Worksheets["Sheet1"];

            // 標題檔
            ws.Cells["A1"].PutValue("表2-5_內政部辦理細項執行情形");
            ws.Cells["AB2"].PutValue("資料截至：" + YearWrite + "年" + SMN + "月");
            var a = ws.Cells.CreateRange("A5:AQ5");

            Aspose.Cells.Style style1 = ws.Cells["A6"].GetStyle();
            for (int i = 6; i < dt2.Rows.Count + 5; i++)
            {
                var b = ws.Cells.CreateRange("A" + i.ToString() + ":AQ" + i.ToString());
                b.Copy(a);
            }

            //加總欄位
            double B = 0; // 分支計畫核定經費總計(仟元)
            double I = 0; // 歷年累積撥發金額(仟元)
            double J = 0; // 當年累積撥發金額(仟元)
            double K = 0; // 當年度核定經費(仟元)(A)
            double L = 0; // 依進度可撥經費(仟元)
            double AD = 0;// 總工程經費(仟元)

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

                ws.Cells["A" + rowNumber].PutValue(dt2.Rows[i]["CATY_NAME"].ToString()); /// 計畫說明
                ws.Cells["B" + rowNumber].PutValue(Source.Tools.thousandchange3(dt2.Rows[i]["CATY_FUND_"].ToString())); /// 分支計畫核定經費總計(仟元)
                Tools.SetCellstyle(ws, "B" + rowNumber);

                ws.Cells["C" + rowNumber].PutValue(dt2.Rows[i]["SEQ"].ToString());      /// 項次
                ws.Cells["D" + rowNumber].PutValue(dt2.Rows[i]["TD_CITY"].ToString());  /// 縣市
                ws.Cells["E" + rowNumber].PutValue(dt2.Rows[i]["IPJ_DESC"].ToString()); /// 辦理層級
                ws.Cells["F" + rowNumber].PutValue(CEN_SP_RATIO_TWODECIMALPLACE);       /// 補助比例
                ws.Cells["G" + rowNumber].PutValue(dt2.Rows[i]["TD_NAME"].ToString());  /// 工作名稱(工程標案名稱、勞務採購、其它(如用地補償…))

                if (TD_FINISH is false)
                {
                    ws.Cells["H" + rowNumber].PutValue("未結案"); // 結案情形
                }
                else if (TD_FINISH is true)
                {
                    ws.Cells["H" + rowNumber].PutValue("已結案"); // 結案情形
                }
                else
                {
                    ws.Cells["H" + rowNumber].PutValue(TD_FINISH); // 結案情形
                }

                ws.Cells["I" + rowNumber].PutValue(Source.Tools.thousandchange3(dt2.Rows[i]["ALL_PPAY_FEE"].ToString())); // 歷年累積撥發金額(仟元)
                Tools.SetCellstyle(ws, "I" + rowNumber);
                ws.Cells["J" + rowNumber].PutValue(dt2.Rows[i]["PPAY_FEE"].ToString()); // 當年累積撥發金額(仟元)
                Tools.SetCellstyle(ws, "J" + rowNumber);
                ws.Cells["K" + rowNumber].PutValue(Source.Tools.thousandchange3(dt2.Rows[i]["YR_FUND"].ToString())); // 當年度核定經費(仟元)(A)
                Tools.SetCellstyle(ws, "K" + rowNumber);
                ws.Cells["O" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK103"].ToString())); // 預定細設完成日期
                ws.Cells["P" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK103"].ToString())); // 實際細設完成日期
                ws.Cells["Q" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK105"].ToString())); // 預定發包日期
                ws.Cells["R" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK105"].ToString())); // 實際發包日期
                ws.Cells["M" + rowNumber].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK102"].ToString())); // 基設核訂日期

                string S = dt2.Rows[i]["PRE_SCHD"].ToString();
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

                ws.Cells["T" + rowNumber].Value = T;             // 實際工程進度(%)

                if (string.IsNullOrEmpty(S) || string.IsNullOrEmpty(T))
                {
                    ws.Cells["U" + rowNumber].PutValue("");      // 進度差異(%)
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

                // 20240124在最後一行加上合計
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
            ws.PageSetup.PaperSize = PaperSizeType.PaperA4;             // 紙張大小(A4)
            ws.PageSetup.Orientation = PageOrientationType.Landscape;   // 紙張方向(橫向)
            ws.PageSetup.TopMargin = 1.91 / 2.54;    // 上邊界
            ws.PageSetup.BottomMargin = 1.91 / 2.54; // 下邊界
            ws.PageSetup.LeftMargin = 1.91 / 2.54;   // 左邊界
            ws.PageSetup.RightMargin = 1.91 / 2.54;  // 右邊
            ws.PageSetup.FitToPagesWide = 1; // 單頁寬度
            ws.PageSetup.FitToPagesTall = 0; // 高度不(允許多頁)
            ws.PageSetup.Zoom = 100;         // 縮放比例
            ws.AutoFitColumns();
            ws.AutoFitRows();

            // ========================================================================================================

            string filename = "A-2.細項執行情形表" + DateTime.Now.ToString("yyyyMMddHHmmss");
            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");
            workbook.Save(savepath, SaveFormat.Xlsx); // 保存为 Excel 文件

            if (ODS == "1")
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
        }

        public ActionResult ReportPrint7(string PID, string SYR, string SMN, string ODS)
        {
            decimal PIDD = decimal.Parse(PID);
            var PJ = db2.PJMFORM.Where(m => m.PID == PIDD).FirstOrDefault().PJ_NAME;

            var path = Server.MapPath("~/Templates/CD_16.xlsx");
            var file = new FileInfo(path);

            string ADV_UID = User.Identity.Name;
            string Dept = db.OGCDPUR
                .Where(x => x.USR_ID == ADV_UID && x.DATA_TP == "D")
                .Select(x => x.DEPT_NO)
                .FirstOrDefault();

            SqlConnection con = new SqlConnection();
            con.ConnectionString = connectionString;

            DataTable dt1 = Source.Report_Source.getBillionProject06Cities(ADV_UID, Dept, PID, SYR, SMN, connectionString);
            DataTable dt2 = Source.Report_Source.getBillionProject06Multi(ADV_UID, Dept, PID, SYR, SMN, connectionString);
            DataTable dt3 = Source.Report_Source.getBillionProject06Base(PID, SYR, SMN, connectionString);
            DataTable dt4 = Source.Report_Source.getBillionProject06Single(PID, SYR, SMN, connectionString);
            dt2.Columns.Add("SEQ");       ///SEQ 是下表計算列的資料

            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                dt2.Rows[i]["SEQ"] = i + 1;
            }

            string Day = "";
            switch (SMN)
            {
                case "01":
                case "03":
                case "05":
                case "07":
                case "08":
                case "10":
                case "12":
                    Day = "31";
                    break;
                case "02":
                    Day = "28";
                    break;
                default:
                    Day = "30";
                    break;
            }
            ///////////////////Aspose.Cells寫法////////////////////
            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            var workbook = new Aspose.Cells.Workbook(path);
            var ws1 = workbook.Worksheets["總表1"];
            var ws2 = workbook.Worksheets["總表2"];
            var ws3 = workbook.Worksheets["分縣市表"];
            var ws4 = workbook.Worksheets["分項表"];
            ws1.Cells["A1"].PutValue("資料截至：" + SYR + "年" + SMN + "月" + Day + "日");
            ws1.Cells["C5"].PutValue(PJ);

            // 设置ws2的单元格值
            ws2.Cells["A1"].PutValue("本署預算執行情形表(資料截至" + SYR + "年" + SMN + "月底)");
            ws2.Cells["B4"].PutValue(PJ);

            // 设置ws3的单元格值
            ws3.Cells["B2"].PutValue(PJ + "各縣市經費執行進度");
            string B4 = dt3.Rows[0]["NO_FUND"].ToString();
            ws3.Cells["B4"].PutValue(ObjectToDouble(B4));

            // 设置ws4的单元格值
            ws4.Cells["A2"].PutValue("中華民國" + SYR + "年01月01日起至" + SYR + "年" + SMN + "月" + Day + "日止");
            double AM8 = dt4.Rows[0]["NO_FUND"] as double? ?? 0;
            ws4.Cells["K8"].PutValue(AM8);
            ws4.Cells["I8"].PutValue(AM8);

            var a = ws4.Cells.CreateRange("A9:AK9");
            for (int i = 10; i < dt2.Rows.Count + 9; i++)
            {
                var b = ws4.Cells.CreateRange("A" + i.ToString() + ":AK" + i.ToString() + "");
                b.Copy(a, new PasteOptions() { PasteType = PasteType.All });
            }

            double B_All = 0;
            double C_All = 0;
            double D_All = 0;
            double E_All = 0;

            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                ws3.Cells["A" + (5 + i)].PutValue(dt1.Rows[i]["TD_CITY"]);
                double B = ObjectToDouble(dt1.Rows[i]["YR_FUND"].ToString());
                ws3.Cells["B" + (5 + i)].PutValue(B);
                double C = ObjectToDouble(dt1.Rows[i]["SMN_FUND"].ToString());
                ws3.Cells["C" + (5 + i)].PutValue(C);
                double I = ObjectToDouble(dt1.Rows[i]["FEE"].ToString());
                double D;
                double E;
                if (C == 0 || I == 0)
                {
                    D = 0;
                    ws3.Cells["D" + (5 + i)].PutValue(D);
                }
                else
                {
                    D = I / C * 100;
                    ws3.Cells["D" + (5 + i)].PutValue(D);
                }
                if (B == 0 || I == 0)
                {
                    E = 0;
                    ws3.Cells["E" + (5 + i)].PutValue(E);
                }
                else
                {
                    E = I / B * 100;
                    ws3.Cells["E" + (5 + i)].PutValue(E);
                }

                B_All += B;
                C_All += C;
                D_All += D;
                E_All += E;
            }

            ws3.Cells["B35"].PutValue(B_All + ObjectToDouble(B4));
            ws3.Cells["D35"].PutValue(D_All);
            ws3.Cells["C35"].PutValue(C_All);
            ws3.Cells["E35"].PutValue(E_All);

            double I_All = 0;
            double J_All = 0;
            double K_All = 0;
            double L_All = 0;
            double M_All = 0;
            double N_All = 0;
            double O_All = 0;
            double P_All = 0;
            double Q_All = 0;

            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                ws4.Cells["D" + (9 + i)].PutValue(dt2.Rows[i]["SEQ"].ToString());
                ws4.Cells["E" + (9 + i)].PutValue(dt2.Rows[i]["TD_NAME"].ToString());
                ws4.Cells["F" + (9 + i)].PutValue(dt2.Rows[i]["TD_CITY"].ToString());
                ws4.Cells["G" + (9 + i)].PutValue(dt2.Rows[i]["CODE_DESC"].ToString());
                ws4.Cells["H" + (9 + i)].PutValue(dt2.Rows[i]["TIT_CODE"].ToString());

                double J = ObjectToDouble(Tools.thousandchange3(dt2.Rows[i]["REV_FUND"].ToString()));
                double K = ObjectToDouble(Tools.thousandchange3(dt2.Rows[i]["FUND"].ToString()));
                ws4.Cells["J" + (9 + i)].PutValue(J);
                Tools.SetCellstyle(ws4, "J" + (9 + i));
                ws4.Cells["K" + (9 + i)].PutValue(K);
                Tools.SetCellstyle(ws4, "K" + (9 + i));
                double I = J + K;
                ws4.Cells["I" + (9 + i)].PutValue(I);
                Tools.SetCellstyle(ws4, "I" + (9 + i));

                double M = ObjectToDouble(Tools.thousandchange3(dt2.Rows[i]["REV_SMN_FUND"].ToString()));
                double N = ObjectToDouble(Tools.thousandchange3(dt2.Rows[i]["SMN_FUND"].ToString()));
                ws4.Cells["M" + (9 + i)].PutValue(M);
                Tools.SetCellstyle(ws4, "M" + (9 + i));

                ws4.Cells["N" + (9 + i)].PutValue(N);
                Tools.SetCellstyle(ws4, "N" + (9 + i));

                double L = M + N;
                ws4.Cells["L" + (9 + i)].PutValue(L);
                Tools.SetCellstyle(ws4, "L" + (9 + i));


                double AP = ObjectToDouble(Tools.thousandchange3(dt2.Rows[i]["REV_COST_FEE"].ToString()));
                double AR = ObjectToDouble(Tools.thousandchange3(dt2.Rows[i]["REV_PPAY_FEE"].ToString()));
                double P = AP + AR;
                ws4.Cells["P" + (9 + i)].PutValue(P);
                Tools.SetCellstyle(ws4, "P" + (9 + i));

                double AQ = ObjectToDouble(Tools.thousandchange3(dt2.Rows[i]["COST_FEE"].ToString()));
                double AS = ObjectToDouble(Tools.thousandchange3(dt2.Rows[i]["PPAY_FEE"].ToString()));
                double Q = AQ + AS;
                ws4.Cells["Q" + (9 + i)].PutValue(Q);
                Tools.SetCellstyle(ws4, "Q" + (9 + i));

                double O = AP + AR + AQ + AS;
                ws4.Cells["O" + (9 + i)].PutValue(O);
                Tools.SetCellstyle(ws4, "O" + (9 + i));

                if (O == 0 || L == 0)
                {
                    ws4.Cells["R" + (9 + i)].PutValue(0);
                }
                else
                {
                    ws4.Cells["R" + (9 + i)].PutValue(O / L);
                }

                if (O == 0 || I == 0)
                {
                    ws4.Cells["S" + (9 + i)].PutValue(0);
                }
                else
                {
                    ws4.Cells["S" + (9 + i)].PutValue(O / I);
                }

                ws4.Cells["T" + (9 + i)].PutValue(I);
                Tools.SetCellstyle(ws4, "T" + (9 + i));

                double U = L;
                ws4.Cells["U" + (9 + i)].PutValue(L);
                Tools.SetCellstyle(ws4, "U" + (9 + i));

                double V = AP + AQ;
                ws4.Cells["V" + (9 + i)].PutValue(V);
                Tools.SetCellstyle(ws4, "V" + (9 + i));

                double W = AR + AS;
                ws4.Cells["W" + (9 + i)].PutValue(W);
                Tools.SetCellstyle(ws4, "W" + (9 + i));

                double X = 0;
                ws4.Cells["X" + (9 + i)].PutValue(0);
                Tools.SetCellstyle(ws4, "X" + (9 + i));


                if (L == 0)
                {
                    ws4.Cells["Y" + (9 + i)].PutValue(0);
                }
                else
                {
                    ws4.Cells["Y" + (9 + i)].PutValue((V + W + X) / U);
                }

                if (I == 0)
                {
                    ws4.Cells["Z" + (9 + i)].PutValue(0);
                }
                else
                {
                    ws4.Cells["Z" + (9 + i)].PutValue((V + W + X) / I);
                }

                ws4.Cells["AA" + (9 + i)].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK102"].ToString()));
                ws4.Cells["AB" + (9 + i)].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK102"].ToString()));
                ws4.Cells["AC" + (9 + i)].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK105"].ToString()));
                ws4.Cells["AD" + (9 + i)].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK105"].ToString()));

                double AE = ObjectToDouble(dt2.Rows[i]["PRE_SCHD"].ToString()) / 100;
                ws4.Cells["AE" + (9 + i)].PutValue(AE);
                //Tools.SetCellstyle(ws4, "AE" + (9 + i));


                double AF = ObjectToDouble(dt2.Rows[i]["REL_SCHD"].ToString()) / 100;
                ws4.Cells["AF" + (9 + i)].PutValue(AF);
                //Tools.SetCellstyle(ws4, "AF" + (9 + i));

                if (AE == 0 || AF == 0)
                {
                    ws4.Cells["AG" + (9 + i)].PutValue(AF);
                    //Tools.SetCellstyle(ws4, "AG" + (9 + i));

                }
                else if (AF == 0)
                {
                    ws4.Cells["AG" + (9 + i)].PutValue(0 - AE);
                    //Tools.SetCellstyle(ws4, "AG" + (9 + i));


                }
                else
                {
                    ws4.Cells["AG" + (9 + i)].PutValue(AF - AE);
                    //Tools.SetCellstyle(ws4, "AG" + (9 + i));

                }

                ws4.Cells["AH" + (9 + i)].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK107"].ToString()));
                ws4.Cells["AI" + (9 + i)].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK107"].ToString()));
                ws4.Cells["AJ" + (9 + i)].PutValue(dt2.Rows[i]["EXECT_DESC"]);
                ws4.Cells["AK" + (9 + i)].PutValue(dt2.Rows[i]["MEMO"]);

                I_All += I;
                J_All += J;
                K_All += K;
                L_All += L;
                M_All += M;
                N_All += N;
                O_All += O;
                P_All += P;
                Q_All += Q;
            }

            ws4.Cells["I7"].PutValue(I_All);
            ws4.Cells["J7"].PutValue(J_All);
            ws4.Cells["K7"].PutValue(K_All);
            ws4.Cells["L7"].PutValue(L_All);
            ws4.Cells["M7"].PutValue(M_All);
            ws4.Cells["N7"].PutValue(N_All);
            ws4.Cells["O7"].PutValue(O_All);
            ws4.Cells["P7"].PutValue(P_All);
            ws4.Cells["Q7"].PutValue(Q_All);

            if (O_All == 0 || L_All == 0)
            {
                ws4.Cells["R7"].PutValue(0);
            }
            else
            {
                ws4.Cells["R7"].PutValue(O_All / L_All);
            }

            if (O_All == 0 || I_All == 0)
            {
                ws4.Cells["S7"].PutValue(0);
            }
            else
            {
                ws4.Cells["S7"].PutValue(O_All / I_All);
            }

            ws4.Cells["T7"].PutValue(I_All);
            ws4.Cells["U7"].PutValue(L_All);
            ws4.Cells["V7"].PutValue(O_All);
            ws4.Cells["W7"].PutValue(P_All);
            ws4.Cells["X7"].PutValue(Q_All);

            string filename = "99年振興經濟表" + DateTime.Now.ToString("yyyyMMddHHmmss");
            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");

            // 此处填充 workbook 的数据和内容

            workbook.Save(savepath, SaveFormat.Xlsx); // 保存为 Excel 文件

            if (ODS == "1")
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


            //////////////////////////////////////////////////////
        }

        public ActionResult ReportPrint8(string PID, string SYR, string SMN, string ODS)
        {
            //string UserID = (string)Session["UsrId"];
            string UserID = User.Identity.Name;
            //string Dept = (string)Session["DeptId"];
            string Dept = db.OGCDPUR
                .Where(x => x.USR_ID == UserID && x.DATA_TP == "D")
                .Select(x => x.DEPT_NO)
                .FirstOrDefault();
            SqlConnection con = new SqlConnection();
            con.ConnectionString = connectionString;

            string sql = Source.Report_Source.GetBillionProject07(PID, SYR, SMN, connectionString, UserID, Dept);
            SqlDataAdapter adp = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            adp.Fill(ds);
            DataTable dt2 = ds.Tables[0];

            dt2.Columns.Add("SEQ");       ///SEQ 是下表計算列的資料

            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                dt2.Rows[i]["SEQ"] = i + 1;
            }

            var path = Server.MapPath("~/Templates/PJV0010_04.xlsx");
            var file = new FileInfo(path);

            //////////////////Aspose.Cells 寫法//////////////////
            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            var workbook = new Aspose.Cells.Workbook(path);
            var ws = workbook.Worksheets["sheet1"];
            // 修改工作表名稱
            ws.Name = "執行情形";

            // 取得儲存格範圍
            Aspose.Cells.Range a = ws.Cells.CreateRange("A8", "AF8");
            var c = ws.Cells["P8"].Formula;
            for (int i = 9; i < dt2.Rows.Count + 8; i++)
            {
                var b = ws.Cells.CreateRange("A" + i.ToString(), "AF" + i.ToString());
                b.Copy(a);
            }
            var SDT = "";

            switch (SMN)
            {
                case "02":
                    SDT = "28";
                    break;
                case "04":
                case "06":
                case "09":
                case "11":
                    SDT = "30";
                    break;
                default:
                    SDT = "31";
                    break;
            }
            ws.Cells["A2"].PutValue("中華民國" + SYR + "年" + SMN + "月1日起至" + SYR + "年" + SMN + "月" + SDT + "日止");
            ws.Cells["G6"].PutValue("" + SYR + "年度\r\n預算");
            ws.Cells["J6"].PutValue("" + SYR + "年度\r\n預算分配數");
            ws.Cells["M6"].PutValue("" + SYR + "年度\r\n預算實支數\r\n(含暫付數)");
            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                ws.Cells[$"C{8 + i}"].PutValue(dt2.Rows[i]["SEQ"]);
                ws.Cells[$"D{8 + i}"].PutValue(dt2.Rows[i]["TD_NAME"].ToString());

                double AG = Convert.ToDouble(dt2.Rows[i]["REV_FUND"]);
                AG = Source.Tools.thousandchange2(AG);
                ws.Cells[$"F{8 + i}"].PutValue(AG);

                double AH = Convert.ToDouble(dt2.Rows[i]["FUND"]);
                AH = Source.Tools.thousandchange2(AH);
                ws.Cells[$"G{8 + i}"].PutValue(AH);

                ws.Cells[$"E{8 + i}"].PutValue(AG + AH);

                double AI = Convert.ToDouble(dt2.Rows[i]["REV_SMN_FUND"]);
                AI = Source.Tools.thousandchange2(AI);
                ws.Cells[$"I{8 + i}"].PutValue(AI);

                ws.Cells[$"F{8 + i}"].PutValue(AG);

                double AJ = Convert.ToDouble(dt2.Rows[i]["SMN_FUND"]);
                AJ = Source.Tools.thousandchange2(AJ);
                ws.Cells[$"J{8 + i}"].PutValue(AJ);

                ws.Cells[$"H{8 + i}"].PutValue(AI + AJ);

                double AK = Convert.ToDouble(dt2.Rows[i]["REV_COST_FEE"]);
                AK = Source.Tools.thousandchange2(AK);
                double AL = Convert.ToDouble(dt2.Rows[i]["REV_PPAY_FEE"]);
                AL = Source.Tools.thousandchange2(AL);

                ws.Cells[$"L{8 + i}"].PutValue(AK + AL);

                double AM = Convert.ToDouble(dt2.Rows[i]["COST_FEE"]);
                AM = Source.Tools.thousandchange2(AM);
                double AN = dt2.Rows[i]["PPAY_FEE"] as double? ?? 0;
                AN = Source.Tools.thousandchange2(AN);

                ws.Cells[$"M{8 + i}"].PutValue(AM + AN);
                ws.Cells[$"K{8 + i}"].PutValue(AK + AL + AM + AN);

                double N = ((AI + AJ) == 0) ? 0 : (AK + AL + AM + AN) / (AI + AJ);
                ws.Cells[$"N{8 + i}"].PutValue(N);

                double O = ((AG + AH) == 0) ? 0 : (AK + AL + AM + AN) / (AG + AH);
                ws.Cells[$"O{8 + i}"].PutValue(O);

                double P = AG + AH;
                ws.Cells[$"P{8 + i}"].PutValue(P);

                double Q = AI + AJ;
                ws.Cells[$"Q{8 + i}"].PutValue(Q);

                double R = Source.Tools.thousandchange2(AK) + Source.Tools.thousandchange2(AM);
                ws.Cells[$"R{8 + i}"].PutValue(R);

                double S = Source.Tools.thousandchange2(AL) + Source.Tools.thousandchange2(AN);
                ws.Cells[$"S{8 + i}"].PutValue(S);

                double T = 0;
                ws.Cells[$"T{8 + i}"].PutValue(T);

                double U = (Q == 0) ? 0 : (R + S + T) / Q;
                ws.Cells[$"U{8 + i}"].PutValue(U);

                double V = (P == 0) ? 0 : (R + S + T) / P;
                ws.Cells[$"V{8 + i}"].PutValue(V);

                ws.Cells[$"W{8 + i}"].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK102"].ToString()));
                ws.Cells[$"X{8 + i}"].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK102"].ToString()));
                ws.Cells[$"Y{8 + i}"].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK105"].ToString()));
                ws.Cells[$"Z{8 + i}"].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK105"].ToString()));

                double AA = Source.Tools.IsNumberOrNot(dt2.Rows[i]["PRE_SCHD"] as double? ?? 0);
                double AB = Source.Tools.IsNumberOrNot(dt2.Rows[i]["REL_SCHD"] as double? ?? 0);

                ws.Cells[$"AA{8 + i}"].PutValue(AA);
                ws.Cells[$"AB{8 + i}"].PutValue(AB);

                double AC = AB - AA;
                ws.Cells[$"AC{8 + i}"].PutValue(AC);

                ws.Cells[$"AD{8 + i}"].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["PRE_TK107"].ToString()));
                ws.Cells[$"AE{8 + i}"].PutValue(Source.Tools.FormatDate(dt2.Rows[i]["REL_TK107"].ToString()));

                string AX = dt2.Rows[i]["DELAY_RESN"].ToString();
                string AW = dt2.Rows[i]["EXECT_DESC"].ToString();

                string AF = string.IsNullOrEmpty(AX) ? AW : AX;
                ws.Cells[$"AF{8 + i}"].PutValue(AF);
            }

            string filename = "D-1.公共建設推動會報計畫執行情形表" + DateTime.Now.ToString("yyyyMMddHHmmss");
            var savepath = Server.MapPath("~/Temp/" + filename + ".xlsx");


            workbook.Save(savepath, SaveFormat.Xlsx); // 保存为 Excel 文件

            if (ODS == "1")
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
            /////////////////////////////////////////////////////
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
            // string UserID = (string)Session["UsrId"];
            string UserID = User.Identity.Name;
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

        [HttpPost]
        public JsonResult ChoosePeriod(string selectedPID)
        {
            #region 舊版程式碼 => 處理單一 PID (參數改用 int)

            //var Period = db2.PJDFDCR
            //    .OrderByDescending(m => m.SEQ_NO)
            //    .Where(m => m.PID == selectedPID)
            //    .ToList();

            #endregion

            #region 新版程式碼 => 處理多個 PID (參數改用 string)

            var PIDs = selectedPID.Split(',')
                .Select(id => decimal.Parse(id))
                .ToList();

            var query = db2.PJDFDCR.Where(m => PIDs.Contains(m.PID)).ToList();

            // 依照 SEQ_NO 分組,返回每一筆 PID 資料
            var Period = query
                .GroupBy(m => m.SEQ_NO)
                .Select(g => new
                {
                    SEQ_NO = g.Key,
                    Records = g.Select(x => new
                    {
                        x.PID,
                        x.AP_DT,
                        x.AP_FN,
                        x.APROV_DT,
                        x.APROV_FN,
                        x.FDCR_RESN,
                        x.AP_OUT_FUND,
                        x.ADD_DT,
                        x.UPD_DT,
                        x.UPD_UID
                    }).ToList()
                })
                .OrderByDescending(x => x.SEQ_NO)
                .ToList();

            #endregion

            string json = JsonConvert.SerializeObject(Period);
            return Json(json);
        }

        public JsonResult ChooseSubject(string selectedPID, string SYR)
        {
            #region 舊版程式碼 => 處理單一 PID (參數改用 int)

            //var Result = (from a in db2.PJDCATY
            //              join b in db2.SMCYACC on a.ACC_ID equals b.ACC_CODE
            //              where a.PID == selectedPID && b.SYR == SYR
            //              orderby a.ACC_ID
            //              select new { a.ACC_ID, b.ACC_DESC }).Distinct();

            #endregion

            #region 新版程式碼 => 處理多項 PID (參數改用 string)

            var PIDs = selectedPID.Split(',').Select(decimal.Parse).ToList();

            var Result = (from a in db2.PJDCATY
                          join b in db2.SMCYACC on a.ACC_ID equals b.ACC_CODE
                          where PIDs.Contains(a.PID) && b.SYR == SYR
                          orderby a.ACC_ID
                          select new { a.ACC_ID, b.ACC_DESC }).Distinct();

            #endregion

            string json = JsonConvert.SerializeObject(Result);
            return Json(json);
        }

        private bool IsProjectMerge(string projectName)
        {
            var mergeName = new[]
            {
                "01辦理水與安全業務費",
                "02辦理水與安全獎補助費",
                "03辦理水與安全設備投資費"
            };

            return mergeName.Any(name => projectName.Contains(name));
        }

        /// <summary>
        /// A-2.細項執行情形表, (擴充):取得合併專案
        /// </summary>
        private SelectListItem GetMergeProject(List<PJMFORM_CMEPMS> projects)
        {
            var Items = projects
                .Where(m => IsProjectMerge(m.PJ_NAME)) // 判斷是否符合合併條件
                .GroupBy(m => 1)                       // 將符合條件的專案進行分組
                .Select(g => new SelectListItem
                {
                    Value = string.Join(",", g.Select(m => m.PID.ToString())), // 將所有的 PID 合併為逗號分隔的字符串
                    Text = "前瞻基礎建設計畫-縣市管河川及區域排水整理改善計畫"        // 統一名稱
                })
                .FirstOrDefault();                                             // 取第一個結果

            return Items;
        }

        /// <summary>
        /// 取得計劃年度
        /// </summary>
        /// <param name="selectYear"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult Orders(string selectYear, string type)
        {
            #region 舊版程式碼

            //ReportViewModel reportViewModel = new ReportViewModel();

            //if (type != null)
            //{
            //    reportViewModel.PJ_NAME = db2.PJMFORM.Where(m => m.SYR == selectYear && m.AN_PID != null && m.Water_Type == type).Select(m => new SelectListItem
            //    {
            //        Value = m.PID.ToString(),
            //        Text = db2.PJMFORM.Where(n => n.PID == m.AN_PID).FirstOrDefault().PJ_NAME.ToString() + "-" + m.PJ_NAME
            //    }).ToList();
            //}
            //else
            //{
            //    reportViewModel.PJ_NAME = db2.PJMFORM.Where(m => m.SYR == selectYear && m.AN_PID != null).Select(m => new SelectListItem
            //    {
            //        Value = m.PID.ToString(),
            //        Text = db2.PJMFORM.Where(n => n.PID == m.AN_PID).FirstOrDefault().PJ_NAME.ToString() + "-" + m.PJ_NAME
            //    }).ToList();
            //}

            //string json = JsonConvert.SerializeObject(reportViewModel.PJ_NAME);

            //return Json(json);

            #endregion

            #region 新版程式碼

            ReportViewModel reportViewModel = new ReportViewModel();

            reportViewModel.PJ_NAME = new List<SelectListItem>();

            var projects = db2.PJMFORM // 查詢符合條件的專案
                .Where(m => m.SYR == selectYear && m.AN_PID != null && (type == null || m.Water_Type == type))
                .ToList();

            var MergeName = GetMergeProject(projects); // 取得合併項目

            if (MergeName != null)                     // 合併項目加入至 PJ_NAME
            {
                reportViewModel.PJ_NAME = reportViewModel.PJ_NAME.Concat(new List<SelectListItem> { MergeName }).ToList();
            }

            var nonMergedProjects = projects            // 處理不需合併項目
                .Where(m => !IsProjectMerge(m.PJ_NAME)) // 過濾已合併項目
                .Select(m => new SelectListItem
                {
                    Value = m.PID.ToString(),
                    Text = db2.PJMFORM
                        .Where(n => n.PID == m.AN_PID)
                        .FirstOrDefault()?.PJ_NAME?.ToString() + "-" + m.PJ_NAME
                }).ToList();

            reportViewModel.PJ_NAME = reportViewModel.PJ_NAME.Concat(nonMergedProjects).ToList();

            string JsonResult = JsonConvert.SerializeObject(reportViewModel.PJ_NAME);
            return Json(JsonResult);

            #endregion
        }

        /// <summary>
        /// 取得計畫名稱
        /// </summary>
        /// <param name="selectedPID"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult ChoosePJYear(string selectedPID)
        {
            #region 舊版程式碼

            //decimal PID = (decimal)selectedPID;

            //var a = db2.PJMFORM
            //    .Where(m => m.PID == PID)
            //    .FirstOrDefault();

            //var PJNAME = a.PJ_NAME;

            //var PJYEAR = db2.PJMFORM
            //    .Where(m => m.PJ_NAME == PJNAME)
            //    .OrderBy(m => m.SYR)
            //    .DistinctBy(m => m.SYR)
            //    .ToList();

            //string json = JsonConvert.SerializeObject(PJYEAR);

            //return Json(json);

            #endregion

            #region 新版程式碼

            var PIDs = selectedPID.Split(',').Select(id => decimal.Parse(id)).ToList();

            var projects = db2.PJMFORM.Where(m => PIDs.Contains(m.PID)).ToList();

            var PJNAME = projects.FirstOrDefault().PJ_NAME;

            var PJYEAR = db2.PJMFORM
                .Where(m => m.PJ_NAME == PJNAME)
                .OrderBy(m => m.SYR)
                .Distinct()
                .ToList();

            string JsonResult = JsonConvert.SerializeObject(PJYEAR);
            return Json(JsonResult);

            #endregion
        }

        public ActionResult ChooseYear(decimal PID)
        {
            var result = db2.PJV_PJMFORM.Where(m => m.AN_PID == PID).DistinctBy(m => m.SYR).ToList();

            if (result.Count == 0)
            {
                result = db2.PJV_PJMFORM.Where(m => m.PID == PID).DistinctBy(m => m.SYR).ToList();
            }

            string json = JsonConvert.SerializeObject(result);

            return Json(json);
        }

        public ActionResult ChooseState(string value)
        {
            string url = "";
            switch (value)
            {
                case "1":
                    url = "State1";

                    break;
                case "2":
                    url = "State2";
                    break;
                case "3":
                    url = "State3";
                    break;
                case "4":
                    url = "State4";
                    break;
                case "5":
                    url = "State5";
                    break;
                case "6":
                    url = "State6";
                    break;
                case "7":
                    url = "State7";
                    break;
                case "8":
                    url = "State8";
                    break;
                case "9":
                    url = "State9";
                    break;
                case "10":
                    url = "State10";
                    break;
                case "11":
                    url = "State11";
                    break;
                case "12":
                    url = "State12";
                    break;
                case "Index":
                    url = "Index";
                    break;
                case "2_2":
                    url = "State2_2";
                    break;
                case "4_2":
                    url = "State4_2";
                    break;
                case "6_2":
                    url = "State6_2";
                    break;
                case "9_2":
                    url = "State9_2";
                    break;
                case "10_2":
                    url = "State10_2";
                    break;
                case "12_2":
                    url = "State12_2";
                    break;
            }
            return Json(url);
        }

        public ActionResult State1(string message)
        {
            ViewBag.State = "1";
            ViewBag.Message = message;
            var query = User_Right(1);
            return View(query);
        }

        public ActionResult State2()
        {
            ViewBag.State = "2";
            var query = User_Right(2);
            return View(query);
        }

        public ActionResult State3()
        {
            ViewBag.State = "3";

            var query = User_Right(3);
            return View(query);
        }

        public ActionResult State4()
        {
            ViewBag.State = "4";
            var query = User_Right(9);
            state9ViewModel CityAndColumn = new state9ViewModel();

            var Plan_Name = db2.PJV_PJDPCAT.OrderByDescending(m => m.SYR).Where(m => m.PJ_TYPE == "B").ToList();
            CityAndColumn.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            CityAndColumn.PJV_PJDPCAT = Plan_Name;
            CityAndColumn.OGCTITLE = db.OGCTITLE.ToList();

            return View(CityAndColumn);
        }

        public ActionResult State5()
        {
            ViewBag.State = "5";

            var query = User_Right(5);
            return View(query);
        }

        public ActionResult State6()
        {
            ViewBag.State = "6";
            var query = User_Right(6);
            return View();
        }

        public ActionResult State7()
        {
            ViewBag.State = "7";

            var query = User_Right(7);
            return View(query);
        }

        public ActionResult State8()
        {
            ViewBag.State = "8";

            var query = User_Right(8);
            return View(query);
        }

        public ActionResult State9()
        {

            ViewBag.State = "9";
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

        public ActionResult State10()
        {
            ViewBag.State = "10";
            var query = User_Right(10);

            //////扣除前瞻和流域的分支計畫，帶回傳時處理
            var PJMFORM = db2.PJMFORM.Where(m => m.AN_PID != 118 || m.AN_PID != 122 || m.AN_PID != 134 || m.AN_PID != 176 || m.AN_PID != 198 || m.AN_PID != 215 || m.AN_PID != 246 || m.AN_PID != 61 || m.AN_PID != 96 || m.AN_PID != 139).ToList();

            ///////抓取所有的分支計畫
            var pjNames = (from m in PJMFORM
                           where m.AN_PID != null
                           join n in PJMFORM on m.AN_PID equals n.PID
                           select new
                           {
                               PJ = m.PJ_NAME,
                               PJ_NAME = "(" + db2.PJMFORM.Where(z => z.PID == m.AN_PID).FirstOrDefault().PJ_NAME + ")" + m.PJ_NAME,
                               ProjectNameID = m.AN_PID,
                               branchProjectNameID = m.PID,
                               Water_Type = m.Water_Type
                           }).DistinctBy(m => m.PJ_NAME).ToList();

            // Create the SelectListItem objects using the list previously generated
            var PJ_BAG = pjNames.Select(m => new State10ViewModel
            {
                PJ = m.PJ,
                ProjectNameID = m.branchProjectNameID.ToString(),
                PJ_NAME = m.PJ_NAME.ToString(),
                MASTER_PID = m.ProjectNameID.ToString(),
                Water_Type = m.Water_Type
            }).OrderBy(m => m.PJ).ToList();
            ViewBag.PJ_BAG = PJ_BAG;

            ViewBag.CITY = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            ViewBag.TOWN = db.SMCPOSTN.ToList();

            return View(query);
        }

        public ActionResult State11()
        {
            ViewBag.State = "11";

            var query = User_Right(11);

            ReportViewModel reportViewModel = new ReportViewModel();

            var pjNames = (from m in db2.PJMFORM
                           where m.AN_PID != null
                           join n in db2.PJMFORM on m.AN_PID equals n.PID
                           select new
                           {
                               PJ_NAME = m.PJ_NAME,
                               ProjectName = m.PJ_NAME     /////之後有時間要把這裡有空格的也去除
                           }).DistinctBy(m => m.ProjectName).ToList();

            // Create the SelectListItem objects using the list previously generated
            reportViewModel.PJ_NAME = pjNames.Select(m => new SelectListItem
            {
                Value = m.PJ_NAME.ToString(),
                Text = m.ProjectName
            }).OrderBy(m => m.Text).ToList();
            ViewBag.PJ_NAME = reportViewModel.PJ_NAME;

            ViewBag.CITY = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();

            return View(query);
        }

        public ActionResult State12()
        {
            ViewBag.State = "12";
            var query = User_Right(12);
            return View(query);
        }

        public ActionResult State13()
        {

            ViewBag.State = "13";

            var query = User_Right(13);
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
            return View(query);
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

        public static void CopyRange(Worksheet sourceWorksheet, Worksheet destinationWorksheet, int sourceStartRow, int sourceEndRow, int sourceStartColumn, int sourceEndColumn,
                             int destinationStartRow, int destinationEndRow, int destinationStartColumn, int destinationEndColumn)
        {

            destinationWorksheet.Cells.CreateRange(destinationStartRow, destinationStartColumn, destinationEndRow - destinationStartRow + 1, destinationEndColumn - destinationStartColumn + 1)
                 .Copy(sourceWorksheet.Cells.CreateRange(sourceStartRow, sourceStartColumn, sourceEndRow - sourceStartRow + 1, sourceEndColumn - sourceStartColumn + 1));
        }
        [HttpPost]
        public async Task<ActionResult> ReportPrint1(string PID, string SYR, string SMN, string Summary, string[] Site, string ODS)
        {
            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            decimal PIDD = decimal.Parse(PID);
            var PJNAME = db2.PJMFORM.Where(m => m.PID == PIDD).FirstOrDefault().PJ_NAME;

            // string aDV_UID = (string)Session["UsrId"];
            string aDV_UID = User.Identity.Name;
            // string Dept = (string)Session["DeptId"];
            string Dept = db.OGCDPUR
                .Where(x => x.USR_ID == aDV_UID && x.DATA_TP == "D")
                .Select(x => x.DEPT_NO)
                .FirstOrDefault();

            string aDV_DP = "";

            if (Dept == "A270000")
            {
                var result = db.OGCDEPT.Where(m => m.DEPT_NO.Contains("A27")).ToList();
                foreach (var item in result)
                {
                    aDV_DP += item.DEPT_NO + ",";
                }
                aDV_DP = aDV_DP.TrimEnd(',');
            }
            else if (db.OGCDEPT.Any(m => m.AN_DEPT_NO == Dept))
            {
                var result = db.OGCDEPT.Where(m => m.AN_DEPT_NO == Dept);
                foreach (var item in result)
                {
                    aDV_DP += item.DEPT_NO + ",";
                }
                aDV_DP = aDV_DP.TrimEnd(',');
            }
            else
            {
                aDV_DP = Dept;
            }
            DataTable Main = await GetSP_RCD0010_UDF002Async(PID, SYR, SMN, Site, aDV_DP, aDV_UID, Dept);
            if (Main != null && Main.Rows.Count > 1)
            {
                string PJ_DIST_FUND = Main.Rows[Main.Rows.Count - 1][0].ToString();
                Main.Rows.Remove(Main.Rows[Main.Rows.Count - 1]);
                MainDataViewModel gmdvm = GetMainDataViewModel(Main, "", "");
                var Detail = new DataTable();
                string sErrMsg = await GetSP_RCD0010_UDF002SAsync(Detail, PID, SYR, SMN, "", aDV_DP, aDV_UID, Dept);
                int i = 0;
                int row = 0;
                string filepath = Server.MapPath(@"~\Templates\cd_01.xlsx");
                var file = new FileInfo(filepath);
                var workbook = new Aspose.Cells.Workbook(filepath);

                for (int k = 0; k < Site.Length; k++)
                {
                    switch (Site[k])
                    {
                        case "all":

                            Worksheet wsa = workbook.Worksheets[0];
                            wsa = CopyWorksheet(workbook, wsa, "總表1", "總表");
                            //總表
                            wsa.Cells[1, 0].Value = DateTime.Now.ToString("yyyyMMdd");
                            double TotalSumYrFund = 0;
                            double PJ_DIST_FUND1 = StringToDouble(gmdvm.PJ_DIST_FUND);
                            string[] Site1 = new string[Site.Length - 1];
                            int index = 0;
                            for (i = 0; i < Site.Length; i++)
                            {
                                if (Site[i] != "all")
                                {
                                    Site1[index] = Site[i];
                                    index++;
                                }
                            }

                            //2023/05/08 => 先這樣設定
                            string[] Site2 = { "N", "C", "S" };
                            for (int j = 1; j < 3; j++)
                            {
                                for (i = 0; i < Site2.Length; i++)
                                {
                                    object SumYrFund = gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_SUMYR_FUND").GetValue(gmdvm, null);
                                    if (SumYrFund == null)
                                    {
                                        SumYrFund = "0";
                                    }
                                    TotalSumYrFund += Convert.ToDouble(SumYrFund);
                                }
                            }

                            //string []北中南
                            row = 13;
                            for (int j = 1; j < 3; j++)
                            {
                                for (i = 0; i < Site2.Length; i++)
                                {
                                    double SUMYR_FUND = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_SUMYR_FUND").GetValue(gmdvm, null));
                                    double SUMCOST_FEE = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_SUMCOST_FEE").GetValue(gmdvm, null));
                                    double SUMPPAY_FEE = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_SUMPPAY_FEE").GetValue(gmdvm, null));
                                    double LD_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_LD_CNT").GetValue(gmdvm, null));
                                    double LD_UG_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_LD_UG_CNT").GetValue(gmdvm, null));
                                    double LD_HG_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_LD_HG_CNT").GetValue(gmdvm, null));
                                    double DG_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_DG_CNT").GetValue(gmdvm, null));
                                    double DG_UD_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_DG_UD_CNT").GetValue(gmdvm, null));
                                    double DG_HD_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_DG_HD_CNT").GetValue(gmdvm, null));
                                    double CN_UD_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_CN_UD_CNT").GetValue(gmdvm, null));
                                    double CN_UB_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_CN_UB_CNT").GetValue(gmdvm, null));
                                    double CN_CN_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_CN_CN_CNT").GetValue(gmdvm, null));
                                    double CN_HF_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_CN_HF_CNT").GetValue(gmdvm, null));
                                    double M3_PF_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "01_M3_PF_CNT").GetValue(gmdvm, null));
                                    double M4_PF_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "01_M4_PF_CNT").GetValue(gmdvm, null));
                                    double M4_PU_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "01_M4_PU_CNT").GetValue(gmdvm, null));
                                    double COST_CN_FEE = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "02_COST_CN_FEE").GetValue(gmdvm, null));
                                    double COST_LD_FEE = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "02_COST_LD_FEE").GetValue(gmdvm, null));
                                    int row1 = row;
                                    if (Site2[i] == "C")
                                    {
                                        row1 = row + 1;
                                    }
                                    else if (Site2[i] == "S")
                                    {
                                        row1 = row + 2;
                                    }
                                    wsa.Cells[0, 0].Value = "內政部營建署（下水道工程處）　一億元以上工程計畫執行情形 (" + PJNAME + ")";

                                    wsa.Cells[row1, 1].Value = SUMYR_FUND;
                                    wsa.Cells[row1, 2].Value = Math.Round(PJ_DIST_FUND1 / TotalSumYrFund * SUMYR_FUND);
                                    wsa.Cells[row1, 3].Value = SUMCOST_FEE;
                                    wsa.Cells[row1, 4].Value = SUMPPAY_FEE;
                                    wsa.Cells[row1, 5].Value = CalDiv(SUMCOST_FEE + SUMPPAY_FEE, Math.Round(PJ_DIST_FUND1 / TotalSumYrFund * SUMYR_FUND));
                                    wsa.Cells[row1, 6].Value = LD_CNT + DG_CNT + CN_UD_CNT + CN_UB_CNT + CN_CN_CNT + CN_HF_CNT;
                                    wsa.Cells[row1, 7].Value = LD_CNT;
                                    wsa.Cells[row1, 8].Value = LD_UG_CNT;
                                    wsa.Cells[row1, 9].Value = LD_HG_CNT;
                                    wsa.Cells[row1, 10].Value = DG_CNT;
                                    wsa.Cells[row1, 11].Value = DG_UD_CNT;
                                    wsa.Cells[row1, 12].Value = DG_HD_CNT;
                                    wsa.Cells[row1, 13].Value = CN_UD_CNT + CN_UB_CNT + CN_CN_CNT + CN_HF_CNT;
                                    wsa.Cells[row1, 14].Value = CN_UD_CNT;
                                    wsa.Cells[row1, 15].Value = CN_UB_CNT;
                                    wsa.Cells[row1, 16].Value = CN_CN_CNT;
                                    wsa.Cells[row1, 17].Value = CN_HF_CNT;
                                    if (j.ToString() == "1")
                                    {
                                        wsa.Cells[row1, 18].Value = M3_PF_CNT;
                                        wsa.Cells[row1, 19].Value = M4_PF_CNT;
                                        wsa.Cells[row1, 20].Value = M4_PU_CNT;
                                    }
                                    else
                                    {
                                        wsa.Cells[row1, 18].Value = COST_CN_FEE;
                                        wsa.Cells[row1, 19].Value = COST_LD_FEE;
                                    }
                                }
                                row = row + 3;
                                for (i = 1; i < 21; i++)
                                {
                                    wsa.Cells[row, i].Value = ObjectToDouble(wsa.Cells[row - 1, i].Value) + ObjectToDouble(wsa.Cells[row - 2, i].Value) + ObjectToDouble(wsa.Cells[row - 3, i].Value);
                                }
                                row = row + 3;
                            }
                            row = row - 2; ;
                            for (i = 1; i < 20; i++)
                            {
                                wsa.Cells[row, i].Value = ObjectToDouble(wsa.Cells[row - 1, i].Value) + ObjectToDouble(wsa.Cells[row - 7, i].Value);
                            }
                            wsa.Cells[row, 18].Value = null;
                            wsa.Cells[row, 19].Value = null;

                            //總計
                            row = 5;
                            for (i = 1; i < 3; i++)
                            {
                                double TD_Fund = (double)gmdvm.GetType().GetProperty("TD_FUND_0" + i.ToString()).GetValue(gmdvm, null);
                                double FYR_FUND = (double)gmdvm.GetType().GetProperty("FYR_FUND_0" + i.ToString()).GetValue(gmdvm, null);
                                double REV_FUND = (double)gmdvm.GetType().GetProperty("REV_FUND_0" + i.ToString()).GetValue(gmdvm, null);
                                double COST_FEE = (double)gmdvm.GetType().GetProperty("COST_FEE_0" + i.ToString()).GetValue(gmdvm, null);
                                double REV_COST_FEE = (double)gmdvm.GetType().GetProperty("REV_COST_FEE_0" + i.ToString()).GetValue(gmdvm, null);
                                double SUMPPAY_FEE = (double)gmdvm.GetType().GetProperty("SUMPPAY_FEE_0" + i.ToString()).GetValue(gmdvm, null);
                                double FIN_SAVE = (double)gmdvm.GetType().GetProperty("FIN_SAVE_0" + i.ToString()).GetValue(gmdvm, null);

                                wsa.Cells[row, 1].Value = TD_Fund;
                                wsa.Cells[row, 2].Value = FYR_FUND;
                                wsa.Cells[row, 3].Value = 0;
                                wsa.Cells[row, 4].Value = REV_FUND;
                                wsa.Cells[row, 5].Value = FYR_FUND + 0 + REV_FUND;
                                if (i == 0)
                                {
                                    wsa.Cells[row, 6].Value = gmdvm.PJ_DIST_FUND;
                                }
                                else if (i == 1)
                                {
                                    wsa.Cells[row, 6].Value = ObjectToDouble(wsa.Cells[16, 2].Value);
                                }
                                else
                                {
                                    wsa.Cells[row, 6].Value = ObjectToDouble(wsa.Cells[22, 2].Value);
                                }
                                wsa.Cells[row, 7].Value = COST_FEE;
                                wsa.Cells[row, 8].Value = 0;
                                wsa.Cells[row, 9].Value = REV_COST_FEE;
                                wsa.Cells[row, 10].Value = COST_FEE + 0 + REV_COST_FEE;
                                wsa.Cells[row, 11].Value = SUMPPAY_FEE;
                                wsa.Cells[row, 12].Value = FIN_SAVE;
                                wsa.Cells[row, 13].Value = CalDiv(ObjectToDouble(wsa.Cells[row, 10].Value), ObjectToDouble(wsa.Cells[row, 6].Value));
                                wsa.Cells[row, 14].Value = CalDiv(ObjectToDouble(wsa.Cells[row, 10].Value) + SUMPPAY_FEE + FIN_SAVE, ObjectToDouble(wsa.Cells[row, 6].Value));
                                wsa.Cells[row, 15].Value = CalDiv((COST_FEE + 0 + REV_COST_FEE + SUMPPAY_FEE + FIN_SAVE), FYR_FUND + 0 + REV_FUND);
                                row++;
                            }
                            row = 4;
                            wsa.Cells[row, 0].Value = PJNAME;
                            for (i = 1; i < 13; i++)
                            {
                                wsa.Cells[row, i].Value = ObjectToDouble(wsa.Cells[row + 1, i].Value) + ObjectToDouble(wsa.Cells[row + 2, i].Value);
                            }
                            wsa.Cells[row, 13].Value = 0;
                            wsa.Cells[row, 14].Value = 0;
                            wsa.Cells[row, 15].Value = (ObjectToDouble(wsa.Cells[row, 10].Value) + ObjectToDouble(wsa.Cells[row, 11].Value) + ObjectToDouble(wsa.Cells[row, 12].Value)) / ObjectToDouble(wsa.Cells[row, 5].Value);

                            break;
                        case "N":
                            //用地
                            var Detail_a_n = Detail.Select("ADM_DIST='北區' and CATY_CODE='L'");
                            //自辦
                            var Detail_b_n = Detail.Select("ADM_DIST='北區' and IPJ_CODE='01' and CATY_CODE<>'L'");
                            //地方自辦
                            var Detail_c_n = Detail.Select("ADM_DIST='北區' and IPJ_CODE='02' and CATY_CODE<>'L'");
                            //ExcelWorksheet wsbn = excel.Workbook.Worksheets.Copy("分處", "北區分處");
                            Worksheet wsbn = workbook.Worksheets[0];
                            wsbn = CopyWorksheet(workbook, wsbn, "分處", "北區分處");
                            //ExcelWorksheet wscn = excel.Workbook.Worksheets.Copy("自辦", "北區自辦");
                            Worksheet wscn = workbook.Worksheets[0];
                            wscn = CopyWorksheet(workbook, wscn, "自辦", "北區自辦");
                            //ExcelWorksheet wsdn = excel.Workbook.Worksheets.Copy("用地取得", "北區用地取得");
                            Worksheet wsdn = workbook.Worksheets[0];
                            wsdn = CopyWorksheet(workbook, wsdn, "用地取得", "北區用地取得");
                            //ExcelWorksheet wsen = excel.Workbook.Worksheets.Copy("地方自辦", "北區地方自辦");
                            Worksheet wsen = workbook.Worksheets[0];
                            wsen = CopyWorksheet(workbook, wsen, "地方自辦", "北區地方自辦");
                            //北區分處
                            SheetSub(wsbn, gmdvm, PJNAME);
                            //北區用地取得

                            row = 3;
                            for (i = 0; i < Detail_a_n.Count(); i++)
                            {
                                //wsdn.Cells[4, 1, 5, 13].Copy(wsdn.Cells[row, 1, row + 1, 13]);
                                CopyRange(wsdn, wsdn, 3, 4, 0, 12, row, row + 1, 0, 12);
                                wsdn.Cells[0, 4].Value = "北區用地取得";
                                wsdn.Cells[row, 0].Value = i + 1;
                                wsdn.Cells[row, 1].Value = Detail_a_n[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_a_n[i]["TD_CODE"].ToString();
                                wsdn.Cells[row, 2].Value = Detail_a_n[i]["YR_FUND"];
                                wsdn.Cells[row, 3].Value = Detail_a_n[i]["COST_FEE"];
                                wsdn.Cells[row, 4].Value = CalDiv(ObjectToDouble(Detail_a_n[i]["COST_FEE"]), ObjectToDouble(Detail_a_n[i]["YR_FUND"]));
                                wsdn.Cells[row, 6].Value = Detail_a_n[i]["PRE_TF101"];
                                wsdn.Cells[row, 8].Value = Detail_a_n[i]["PRE_TF105"];
                                wsdn.Cells[row, 9].Value = Detail_a_n[i]["PRE_TF106"];
                                wsdn.Cells[row, 10].Value = Detail_a_n[i]["PRE_TF107"];
                                wsdn.Cells[row, 11].Value = Detail_a_n[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_a_n[i]["RESLV_MTHD"].ToString();
                                wsdn.Cells[row, 12].Value = Detail_a_n[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_a_n[i]["CTDELAY_RESN"].ToString();
                                row++;
                                wsdn.Cells[row, 6].Value = Detail_a_n[i]["REL_TF101"];
                                wsdn.Cells[row, 8].Value = Detail_a_n[i]["REL_TF105"];
                                wsdn.Cells[row, 9].Value = Detail_a_n[i]["REL_TF106"];
                                wsdn.Cells[row, 10].Value = Detail_a_n[i]["REL_TF107"];
                                row++;
                            }
                            //北區自辦
                            row = 3;
                            for (i = 0; i < Detail_b_n.Count(); i++)
                            {
                                //wscn.Cells[4, 1, 5, 17].Copy(wscn.Cells[row, 1, row + 1, 17]);
                                CopyRange(wscn, wscn, 3, 4, 0, 16, row, row + 1, 0, 16);
                                wscn.Cells[0, 6].Value = "北區自辦工程";
                                wscn.Cells[row, 0].Value = i + 1;
                                wscn.Cells[row, 1].Value = Detail_b_n[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_b_n[i]["TD_CODE"].ToString();
                                wscn.Cells[row, 3].Value = Detail_b_n[i]["PRE_SCHD"];
                                wscn.Cells[row, 4].Value = Detail_b_n[i]["REL_SCHD"];
                                wscn.Cells[row, 5].Value = ObjectToDouble(Detail_b_n[i]["REL_SCHD"]) - ObjectToDouble(Detail_b_n[i]["PRE_SCHD"]);
                                wscn.Cells[row, 6].Value = Detail_b_n[i]["PRE_TK101"];
                                wscn.Cells[row, 7].Value = Detail_b_n[i]["PRE_TK103"];
                                wscn.Cells[row, 8].Value = Detail_b_n[i]["PRE_TK105"];
                                wscn.Cells[row, 9].Value = Detail_b_n[i]["PRE_TK107"];
                                wscn.Cells[row, 10].Value = Detail_b_n[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_b_n[i]["RESLV_MTHD"].ToString();
                                wscn.Cells[row, 11].Value = Detail_b_n[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_b_n[i]["CTDELAY_RESN"].ToString();
                                wscn.Cells[row, 12].Value = Detail_b_n[i]["SAVED_FEE"];
                                wscn.Cells[row, 13].Value = Detail_b_n[i]["YR_FUND"];
                                wscn.Cells[row, 14].Value = Detail_b_n[i]["COST_FEE"];
                                wscn.Cells[row, 15].Value = Detail_b_n[i]["UPAY_FEE"];
                                row++;
                                wscn.Cells[row, 6].Value = Detail_b_n[i]["REL_TK101"];
                                wscn.Cells[row, 7].Value = Detail_b_n[i]["REL_TK103"];
                                wscn.Cells[row, 8].Value = Detail_b_n[i]["REL_TK105"];
                                wscn.Cells[row, 9].Value = Detail_b_n[i]["REL_TK107"];
                                row++;
                            }
                            //北區地方自辦
                            row = 3;
                            wsen.Cells[0, 5].Value = "北區地方自辦工程";
                            for (i = 0; i < Detail_c_n.Count(); i++)
                            {

                                //wsen.Cells[4, 1, 5, 17].Copy(wsen.Cells[row, 1, row + 1, 17]);
                                CopyRange(wsen, wsen, 3, 4, 0, 16, row, row + 1, 0, 16);
                                wsen.Cells[row, 0].Value = i + 1;
                                wsen.Cells[row, 1].Value = Detail_c_n[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_c_n[i]["TD_CODE"].ToString();
                                wsen.Cells[row, 3].Value = Detail_c_n[i]["PRE_SCHD"];
                                wsen.Cells[row, 4].Value = Detail_c_n[i]["REL_SCHD"];
                                wsen.Cells[row, 5].Value = ObjectToDouble(Detail_c_n[i]["REL_SCHD"]) - ObjectToDouble(Detail_c_n[i]["PRE_SCHD"]);
                                wsen.Cells[row, 6].Value = Detail_c_n[i]["PRE_TK101"];
                                wsen.Cells[row, 7].Value = Detail_c_n[i]["PRE_TK103"];
                                wsen.Cells[row, 8].Value = Detail_c_n[i]["PRE_TK105"];
                                wsen.Cells[row, 9].Value = Detail_c_n[i]["PRE_TK107"];
                                wsen.Cells[row, 10].Value = Detail_c_n[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_c_n[i]["RESLV_MTHD"].ToString();
                                wsen.Cells[row, 11].Value = Detail_c_n[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_c_n[i]["CTDELAY_RESN"].ToString();
                                wsen.Cells[row, 13].Value = Detail_c_n[i]["SAVED_FEE"];
                                wsen.Cells[row, 14].Value = Detail_c_n[i]["YR_FUND"];
                                wsen.Cells[row, 15].Value = Detail_c_n[i]["COST_FEE"];
                                row++;
                                wsen.Cells[row, 6].Value = Detail_c_n[i]["REL_TK101"];
                                wsen.Cells[row, 7].Value = Detail_c_n[i]["REL_TK103"];
                                wsen.Cells[row, 8].Value = Detail_c_n[i]["REL_TK105"];
                                wsen.Cells[row, 9].Value = Detail_c_n[i]["REL_TK107"];
                                row++;
                            }
                            break;
                        case "C":
                            //用地
                            var Detail_a_c = Detail.Select("ADM_DIST='中區' and CATY_CODE='L'");
                            //自辦
                            var Detail_b_c = Detail.Select("ADM_DIST='中區' and IPJ_CODE='01' and CATY_CODE<>'L'");
                            //地方自辦
                            var Detail_c_c = Detail.Select("ADM_DIST='中區' and IPJ_CODE='02' and CATY_CODE<>'L'");
                            //ExcelWorksheet wsbc = excel.Workbook.Worksheets.Copy("分處", "中區分處");
                            //ExcelWorksheet wscc = excel.Workbook.Worksheets.Copy("自辦", "中區自辦");
                            //ExcelWorksheet wsdc = excel.Workbook.Worksheets.Copy("用地取得", "中區用地取得");
                            //ExcelWorksheet wsec = excel.Workbook.Worksheets.Copy("地方自辦", "中區地方自辦");
                            Worksheet wsbc = workbook.Worksheets[0];
                            wsbc = CopyWorksheet(workbook, wsbc, "分處", "中區分處");
                            Worksheet wscc = workbook.Worksheets[0];
                            wscc = CopyWorksheet(workbook, wscc, "自辦", "中區自辦");
                            Worksheet wsdc = workbook.Worksheets[0];
                            wsdc = CopyWorksheet(workbook, wsdc, "用地取得", "中區用地取得");
                            Worksheet wsec = workbook.Worksheets[0];
                            wsec = CopyWorksheet(workbook, wsec, "地方自辦", "中區地方自辦");

                            //中區分處
                            SheetSub(wsbc, gmdvm, PJNAME);
                            //中區用地取得
                            row = 3;
                            wsdc.Cells[0, 4].Value = "中區用地取得";
                            for (i = 0; i < Detail_a_c.Count(); i++)
                            {
                                //wsdc.Cells[4, 1, 5, 13].Copy(wsdc.Cells[row, 1, row + 1, 13]);
                                CopyRange(wsdc, wsdc, 3, 4, 0, 12, row, row + 1, 0, 12);
                                //wsec.Cells["A4:Q4"].Copy(wsec.Cells["A6:Q6"]);

                                wsdc.Cells[row, 0].Value = i + 1;
                                wsdc.Cells[row, 1].Value = Detail_a_c[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_a_c[i]["TD_CODE"].ToString();
                                wsdc.Cells[row, 2].Value = Detail_a_c[i]["YR_FUND"];
                                wsdc.Cells[row, 3].Value = Detail_a_c[i]["COST_FEE"];
                                wsdc.Cells[row, 4].Value = CalDiv(ObjectToDouble(Detail_a_c[i]["COST_FEE"]), ObjectToDouble(Detail_a_c[i]["YR_FUND"]));
                                wsdc.Cells[row, 6].Value = Detail_a_c[i]["PRE_TF101"];
                                wsdc.Cells[row, 8].Value = Detail_a_c[i]["PRE_TF105"];
                                wsdc.Cells[row, 9].Value = Detail_a_c[i]["PRE_TF106"];
                                wsdc.Cells[row, 10].Value = Detail_a_c[i]["PRE_TF107"];
                                wsdc.Cells[row, 11].Value = Detail_a_c[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_a_c[i]["RESLV_MTHD"].ToString();
                                wsdc.Cells[row, 12].Value = Detail_a_c[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_a_c[i]["CTDELAY_RESN"].ToString();
                                row++;
                                wsdc.Cells[row, 6].Value = Detail_a_c[i]["REL_TF101"];
                                wsdc.Cells[row, 8].Value = Detail_a_c[i]["REL_TF105"];
                                wsdc.Cells[row, 9].Value = Detail_a_c[i]["REL_TF106"];
                                wsdc.Cells[row, 10].Value = Detail_a_c[i]["REL_TF107"];
                                row++;
                            }
                            //中區自辦
                            row = 3;
                            wscc.Cells[0, 6].Value = "中區自辦工程";
                            for (i = 0; i < Detail_b_c.Count(); i++)
                            {
                                //wscc.Cells[4, 1, 5, 17].Copy(wscc.Cells[row, 1, row + 1, 17]);
                                CopyRange(wscc, wscc, 3, 4, 0, 16, row, row + 1, 0, 16);

                                wscc.Cells[row, 0].Value = i + 1;
                                wscc.Cells[row, 1].Value = Detail_b_c[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_b_c[i]["TD_CODE"].ToString();
                                wscc.Cells[row, 3].Value = Detail_b_c[i]["PRE_SCHD"];
                                wscc.Cells[row, 4].Value = Detail_b_c[i]["REL_SCHD"];
                                wscc.Cells[row, 5].Value = ObjectToDouble(Detail_b_c[i]["REL_SCHD"]) - ObjectToDouble(Detail_b_c[i]["PRE_SCHD"]);
                                wscc.Cells[row, 6].Value = Detail_b_c[i]["PRE_TK101"];
                                wscc.Cells[row, 7].Value = Detail_b_c[i]["PRE_TK103"];
                                wscc.Cells[row, 8].Value = Detail_b_c[i]["PRE_TK105"];
                                wscc.Cells[row, 9].Value = Detail_b_c[i]["PRE_TK107"];
                                wscc.Cells[row, 10].Value = Detail_b_c[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_b_c[i]["RESLV_MTHD"].ToString();
                                wscc.Cells[row, 11].Value = Detail_b_c[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_b_c[i]["CTDELAY_RESN"].ToString();
                                wscc.Cells[row, 12].Value = Detail_b_c[i]["SAVED_FEE"];
                                wscc.Cells[row, 13].Value = Detail_b_c[i]["YR_FUND"];
                                wscc.Cells[row, 14].Value = Detail_b_c[i]["COST_FEE"];
                                wscc.Cells[row, 15].Value = Detail_b_c[i]["UPAY_FEE"];

                                row++;
                                wscc.Cells[row, 6].Value = Detail_b_c[i]["REL_TK101"];
                                wscc.Cells[row, 7].Value = Detail_b_c[i]["REL_TK103"];
                                wscc.Cells[row, 8].Value = Detail_b_c[i]["REL_TK105"];
                                wscc.Cells[row, 9].Value = Detail_b_c[i]["REL_TK107"];
                                row++;
                            }
                            //中區地方自辦
                            row = 3;
                            wsec.Cells[0, 5].Value = "中區地方自辦工程";
                            var q = Detail_c_c.Count();
                            for (i = 0; i < Detail_c_c.Count(); i++)
                            {
                                //wsec.Cells[4, 1, 5, 17].Copy(wsec.Cells[row, 1, row + 1, 17]);
                                CopyRange(wsec, wsec, 3, 4, 0, 16, row, row + 1, 0, 16);
                                wsec.Cells[row, 0].Value = i + 1;
                                wsec.Cells[row, 1].Value = Detail_c_c[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_c_c[i]["TD_CODE"].ToString();
                                wsec.Cells[row, 3].Value = Detail_c_c[i]["PRE_SCHD"];
                                wsec.Cells[row, 4].Value = Detail_c_c[i]["REL_SCHD"];
                                wsec.Cells[row, 5].Value = ObjectToDouble(Detail_c_c[i]["REL_SCHD"]) - ObjectToDouble(Detail_c_c[i]["PRE_SCHD"]);
                                wsec.Cells[row, 6].Value = Detail_c_c[i]["PRE_TK101"];
                                wsec.Cells[row, 7].Value = Detail_c_c[i]["PRE_TK103"];
                                wsec.Cells[row, 8].Value = Detail_c_c[i]["PRE_TK105"];
                                wsec.Cells[row, 9].Value = Detail_c_c[i]["PRE_TK107"];
                                wsec.Cells[row, 10].Value = Detail_c_c[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_c_c[i]["RESLV_MTHD"].ToString();
                                wsec.Cells[row, 11].Value = Detail_c_c[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_c_c[i]["CTDELAY_RESN"].ToString();
                                wsec.Cells[row, 13].Value = Detail_c_c[i]["SAVED_FEE"];
                                wsec.Cells[row, 14].Value = Detail_c_c[i]["YR_FUND"];
                                wsec.Cells[row, 15].Value = Detail_c_c[i]["COST_FEE"];
                                row++;
                                wsec.Cells[row, 6].Value = Detail_c_c[i]["REL_TK101"];
                                wsec.Cells[row, 7].Value = Detail_c_c[i]["REL_TK103"];
                                wsec.Cells[row, 8].Value = Detail_c_c[i]["REL_TK105"];
                                wsec.Cells[row, 9].Value = Detail_c_c[i]["REL_TK107"];
                                row++;
                            }

                            break;
                        case "S":
                            //用地
                            var Detail_a_s = Detail.Select("ADM_DIST='南區' and CATY_CODE='L'");
                            //自辦
                            var Detail_b_s = Detail.Select("ADM_DIST='南區' and IPJ_CODE='01' and CATY_CODE<>'L'");
                            //地方自辦
                            var Detail_c_s = Detail.Select("ADM_DIST='南區' and IPJ_CODE='02' and CATY_CODE<>'L'");
                            //ExcelWorksheet wsbs = excel.Workbook.Worksheets.Copy("分處", "南區分處");
                            //ExcelWorksheet wscs = excel.Workbook.Worksheets.Copy("自辦", "南區自辦");
                            //ExcelWorksheet wsds = excel.Workbook.Worksheets.Copy("用地取得", "南區用地取得");
                            //ExcelWorksheet wses = excel.Workbook.Worksheets.Copy("地方自辦", "南區地方自辦");
                            Worksheet wsbs = workbook.Worksheets[0];
                            wsbs = CopyWorksheet(workbook, wsbs, "分處", "南區分處");
                            Worksheet wscs = workbook.Worksheets[0];
                            wscs = CopyWorksheet(workbook, wscs, "自辦", "南區自辦");
                            Worksheet wsds = workbook.Worksheets[0];
                            wsds = CopyWorksheet(workbook, wsds, "用地取得", "南區用地取得");
                            Worksheet wses = workbook.Worksheets[0];
                            wses = CopyWorksheet(workbook, wses, "地方自辦", "南區地方自辦");

                            //南區分處
                            SheetSub(wsbs, gmdvm, PJNAME);
                            //南區用地取得
                            row = 3;
                            wsds.Cells[0, 4].Value = "南區用地取得";
                            for (i = 0; i < Detail_a_s.Count(); i++)
                            {
                                //wsds.Cells[4, 1, 5, 13].Copy(wsds.Cells[row, 1, row + 1, 13]);
                                CopyRange(wsds, wsds, 3, 4, 0, 12, row, row + 1, 0, 12);
                                wsds.Cells[row, 0].Value = i + 1;
                                wsds.Cells[row, 1].Value = Detail_a_s[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_a_s[i]["TD_CODE"].ToString();
                                wsds.Cells[row, 2].Value = Detail_a_s[i]["YR_FUND"];
                                wsds.Cells[row, 3].Value = Detail_a_s[i]["COST_FEE"];
                                wsds.Cells[row, 4].Value = CalDiv(ObjectToDouble(Detail_a_s[i]["COST_FEE"]), ObjectToDouble(Detail_a_s[i]["YR_FUND"]));
                                wsds.Cells[row, 6].Value = Detail_a_s[i]["PRE_TF101"];
                                wsds.Cells[row, 8].Value = Detail_a_s[i]["PRE_TF105"];
                                wsds.Cells[row, 9].Value = Detail_a_s[i]["PRE_TF106"];
                                wsds.Cells[row, 10].Value = Detail_a_s[i]["PRE_TF107"];
                                wsds.Cells[row, 11].Value = Detail_a_s[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_a_s[i]["RESLV_MTHD"].ToString();
                                wsds.Cells[row, 12].Value = Detail_a_s[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_a_s[i]["CTDELAY_RESN"].ToString();
                                row++;
                                wsds.Cells[row, 6].Value = Detail_a_s[i]["REL_TF101"];
                                wsds.Cells[row, 8].Value = Detail_a_s[i]["REL_TF105"];
                                wsds.Cells[row, 9].Value = Detail_a_s[i]["REL_TF106"];
                                wsds.Cells[row, 10].Value = Detail_a_s[i]["REL_TF107"];
                                row++;
                            }
                            //南區自辦

                            wscs.Cells[0, 6].Value = "南區自辦工程";
                            row = 3;
                            for (i = 0; i < Detail_b_s.Count(); i++)
                            {
                                //wscs.Cells[4, 1, 5, 17].Copy(wscs.Cells[row, 1, row + 1, 17]);
                                CopyRange(wscs, wscs, 3, 4, 0, 16, row, row + 1, 0, 16);
                                wscs.Cells[row, 0].Value = i + 1;
                                wscs.Cells[row, 1].Value = Detail_b_s[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_b_s[i]["TD_CODE"].ToString();
                                wscs.Cells[row, 3].Value = Detail_b_s[i]["PRE_SCHD"];
                                wscs.Cells[row, 4].Value = Detail_b_s[i]["REL_SCHD"];
                                wscs.Cells[row, 5].Value = ObjectToDouble(Detail_b_s[i]["REL_SCHD"]) - ObjectToDouble(Detail_b_s[i]["PRE_SCHD"]);
                                wscs.Cells[row, 6].Value = Detail_b_s[i]["PRE_TK101"];
                                wscs.Cells[row, 7].Value = Detail_b_s[i]["PRE_TK103"];
                                wscs.Cells[row, 8].Value = Detail_b_s[i]["PRE_TK105"];
                                wscs.Cells[row, 9].Value = Detail_b_s[i]["PRE_TK107"];
                                wscs.Cells[row, 10].Value = Detail_b_s[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_b_s[i]["RESLV_MTHD"].ToString();
                                wscs.Cells[row, 11].Value = Detail_b_s[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_b_s[i]["CTDELAY_RESN"].ToString();
                                wscs.Cells[row, 12].Value = Detail_b_s[i]["SAVED_FEE"];
                                wscs.Cells[row, 13].Value = Detail_b_s[i]["YR_FUND"];
                                wscs.Cells[row, 14].Value = Detail_b_s[i]["COST_FEE"];
                                wscs.Cells[row, 15].Value = Detail_b_s[i]["UPAY_FEE"];
                                row++;
                                wscs.Cells[row, 6].Value = Detail_b_s[i]["REL_TK101"];
                                wscs.Cells[row, 7].Value = Detail_b_s[i]["REL_TK103"];
                                wscs.Cells[row, 8].Value = Detail_b_s[i]["REL_TK105"];
                                wscs.Cells[row, 9].Value = Detail_b_s[i]["REL_TK107"];
                                row++;
                            }
                            //南區地方自辦
                            wses.Cells[0, 5].Value = "南區地方自辦工程";
                            row = 3;
                            for (i = 0; i < Detail_c_s.Count(); i++)
                            {

                                //wses.Cells[4, 1, 5, 17].Copy(wses.Cells[row, 1, row + 1, 17]);
                                CopyRange(wses, wses, 3, 4, 0, 16, row, row + 1, 0, 16);

                                wses.Cells[row, 0].Value = i + 1;
                                wses.Cells[row, 1].Value = Detail_c_s[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_c_s[i]["TD_CODE"].ToString();
                                wses.Cells[row, 3].Value = Detail_c_s[i]["PRE_SCHD"];
                                wses.Cells[row, 4].Value = Detail_c_s[i]["REL_SCHD"];
                                wses.Cells[row, 5].Value = ObjectToDouble(Detail_c_s[i]["REL_SCHD"]) - ObjectToDouble(Detail_c_s[i]["PRE_SCHD"]);
                                wses.Cells[row, 6].Value = Detail_c_s[i]["PRE_TK101"];
                                wses.Cells[row, 7].Value = Detail_c_s[i]["PRE_TK103"];
                                wses.Cells[row, 8].Value = Detail_c_s[i]["PRE_TK105"];
                                wses.Cells[row, 9].Value = Detail_c_s[i]["PRE_TK107"];
                                wses.Cells[row, 10].Value = Detail_c_s[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_c_s[i]["RESLV_MTHD"].ToString();
                                wses.Cells[row, 11].Value = Detail_c_s[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_c_s[i]["CTDELAY_RESN"].ToString();
                                wses.Cells[row, 13].Value = Detail_c_s[i]["SAVED_FEE"];
                                wses.Cells[row, 14].Value = Detail_c_s[i]["YR_FUND"];
                                wses.Cells[row, 15].Value = Detail_c_s[i]["COST_FEE"];
                                row++;
                                wses.Cells[row, 6].Value = Detail_c_s[i]["REL_TK101"];
                                wses.Cells[row, 7].Value = Detail_c_s[i]["REL_TK103"];
                                wses.Cells[row, 8].Value = Detail_c_s[i]["REL_TK105"];
                                wses.Cells[row, 9].Value = Detail_c_s[i]["REL_TK107"];
                                row++;
                            }
                            break;

                        default:
                            break;
                    }
                }

                for (i = 0; i < 5; i++)
                {
                    workbook.Worksheets.RemoveAt(0);
                }

                string filename = "A-1.一億元以上工程計畫執行情形表" + DateTime.Now.ToString("yyyyMMddHHmmss");
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
            else
            {
                var message = "查無資料";
                return RedirectToAction("State1", new { message = message });
            }
        }

        [HttpPost]
        public async Task<ActionResult> ReportPrint3(string PID, string SYR, string SMN, string[] PJYear, string[] Site, string ODS)
        {
            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            decimal PIDD = decimal.Parse(PID);
            var PJNAME = db2.PJMFORM.Where(m => m.PID == PIDD).FirstOrDefault().PJ_NAME;

            //string aDV_UID = (string)Session["UsrId"];
            string aDV_UID = User.Identity.Name;
            //string Dept = (string)Session["DeptId"];
            string Dept = db.OGCDPUR
                .Where(x => x.USR_ID == aDV_UID && x.DATA_TP == "D")
                .Select(x => x.DEPT_NO)
                .FirstOrDefault();

            string aDV_DP = "";
            if (Dept == "A270000")
            {

                var result = db.OGCDEPT.Where(m => m.DEPT_NO.Contains("A27")).ToList();
                foreach (var item in result)
                {
                    aDV_DP += item.DEPT_NO + ",";
                }
                aDV_DP = aDV_DP.TrimEnd(',');
            }
            else if (db.OGCDEPT.Any(m => m.AN_DEPT_NO == Dept))
            {
                var result = db.OGCDEPT.Where(m => m.AN_DEPT_NO == Dept);
                foreach (var item in result)
                {
                    aDV_DP += item.DEPT_NO + ",";
                }
                aDV_DP = aDV_DP.TrimEnd(',');
            }
            else
            {
                aDV_DP = Dept;
            }

            DataTable Main = await GetSP_RCD0010_UDF002Async(PID, SYR, SMN, Site, aDV_DP, aDV_UID, Dept);

            if (Main != null && Main.Rows.Count > 1)
            {
                string PJ_DIST_FUND = Main.Rows[Main.Rows.Count - 1][0].ToString();
                Main.Rows.Remove(Main.Rows[Main.Rows.Count - 1]);
                MainDataViewModel gmdvm = GetMainDataViewModel(Main, "", "");
                var Detail = new DataTable();
                string sErrMsg = await GetSP_RCD0010_UDF002SAsync(Detail, PID, SYR, SMN, "", aDV_DP, aDV_UID, Dept);


                string filepath = Server.MapPath(@"~\Templates\cd_01.xlsx");
                var file = new FileInfo(filepath); // 檔案路徑

                ReplaceMainDataViewModel(gmdvm, PJYear, PID, SYR, SMN, aDV_UID);
                int i = 0;
                int row = 0;
                for (int k = 0; k < Site.Length; k++)
                {
                    //switch (Site[k])
                    //{
                    //    case "all":
                    //        var workbook = new Aspose.Cells.Workbook(filepath);
                    //        Worksheet sheetToCopy = workbook.Worksheets["總表1"];
                    //        int newSheetIndex = workbook.Worksheets.Add();
                    //        Worksheet newSheet = workbook.Worksheets[newSheetIndex];
                    //        newSheet.Name = "總表";
                    //        newSheet.Copy(sheetToCopy);

                    //        // 將"SheetC"設置為第一張工作表
                    //        workbook.Worksheets.ActiveSheetIndex = newSheetIndex;

                    //}
                }
                ////////////////Aspose.Cells 套件產製////////////////////
                var workbook = new Aspose.Cells.Workbook(filepath);
                /////////////////////////////////////////////////////////
                //using (var excel = new ExcelPackage(file))
                //{
                for (int k = 0; k < Site.Length; k++)
                {
                    switch (Site[k])
                    {
                        case "all":
                            //ExcelWorksheet wsa = excel.Workbook.Worksheets.Copy("總表1", "總表");
                            Worksheet wsa = workbook.Worksheets[0];
                            wsa = CopyWorksheet(workbook, wsa, "總表1", "總表");

                            //總表
                            wsa.Cells[1, 1].Value = "內政部國土管理署（下水道工程處）　一億元以上工程計畫執行情形 (" + PJNAME + ")";
                            wsa.Cells[2, 1].Value = DateTime.Now.ToString("yyyyMMdd");

                            string[] Site1 = new string[Site.Length - 1];
                            int index = 0;
                            for (i = 0; i < Site.Length; i++)
                            {
                                if (Site[i] != "all")
                                {
                                    Site1[index] = Site[i];
                                    index++;
                                }
                            }

                            //20230508先這樣設定
                            string[] Site2 = { "N", "C", "S" };
                            double TotalSumYrFund = 0;
                            double PJ_DIST_FUND1 = StringToDouble(gmdvm.PJ_DIST_FUND);
                            for (int j = 1; j < 3; j++)
                            {
                                for (i = 0; i < Site2.Length; i++)
                                {
                                    object SumYrFund = gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_SUMYR_FUND").GetValue(gmdvm, null);
                                    if (SumYrFund == null)
                                    {
                                        SumYrFund = "0";
                                    }
                                    TotalSumYrFund += Convert.ToDouble(SumYrFund);
                                }
                            }
                            //string []北中南

                            row = 13;
                            for (int j = 1; j < 3; j++)
                            {
                                for (i = 0; i < Site2.Length; i++)
                                {
                                    double SUMYR_FUND = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_SUMYR_FUND").GetValue(gmdvm, null));
                                    double SUMCOST_FEE = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_SUMCOST_FEE").GetValue(gmdvm, null));
                                    double SUMPPAY_FEE = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_SUMPPAY_FEE").GetValue(gmdvm, null));
                                    double LD_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_LD_CNT").GetValue(gmdvm, null));
                                    double LD_UG_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_LD_UG_CNT").GetValue(gmdvm, null));
                                    double LD_HG_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_LD_HG_CNT").GetValue(gmdvm, null));
                                    double DG_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_DG_CNT").GetValue(gmdvm, null));
                                    double DG_UD_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_DG_UD_CNT").GetValue(gmdvm, null));
                                    double DG_HD_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_DG_HD_CNT").GetValue(gmdvm, null));
                                    double CN_UD_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_CN_UD_CNT").GetValue(gmdvm, null));
                                    double CN_UB_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_CN_UB_CNT").GetValue(gmdvm, null));
                                    double CN_CN_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_CN_CN_CNT").GetValue(gmdvm, null));
                                    double CN_HF_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + j.ToString("D2") + "_CN_HF_CNT").GetValue(gmdvm, null));
                                    double M3_PF_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "01_M3_PF_CNT").GetValue(gmdvm, null));
                                    double M4_PF_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "01_M4_PF_CNT").GetValue(gmdvm, null));
                                    double M4_PU_CNT = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "01_M4_PU_CNT").GetValue(gmdvm, null));
                                    double COST_CN_FEE = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "02_COST_CN_FEE").GetValue(gmdvm, null));
                                    double COST_LD_FEE = ObjectToDouble(gmdvm.GetType().GetProperty(Site2[i] + "02_COST_LD_FEE").GetValue(gmdvm, null));
                                    int row1 = row;
                                    if (Site2[i] == "C")
                                    {
                                        row1 = row + 1;
                                    }
                                    else if (Site2[i] == "S")
                                    {
                                        row1 = row + 2;
                                    }
                                    wsa.Cells[row1, 1].Value = SUMYR_FUND;
                                    wsa.Cells[row1, 2].Value = Math.Round(PJ_DIST_FUND1 / TotalSumYrFund * SUMYR_FUND);
                                    wsa.Cells[row1, 3].Value = SUMCOST_FEE;
                                    wsa.Cells[row1, 4].Value = SUMPPAY_FEE;
                                    wsa.Cells[row1, 5].Value = CalDiv(SUMCOST_FEE + SUMPPAY_FEE, Math.Round(PJ_DIST_FUND1 / TotalSumYrFund * SUMYR_FUND));
                                    wsa.Cells[row1, 6].Value = LD_CNT + DG_CNT + CN_UD_CNT + CN_UB_CNT + CN_CN_CNT + CN_HF_CNT;
                                    wsa.Cells[row1, 7].Value = LD_CNT;
                                    wsa.Cells[row1, 8].Value = LD_UG_CNT;
                                    wsa.Cells[row1, 9].Value = LD_HG_CNT;
                                    wsa.Cells[row1, 10].Value = DG_CNT;
                                    wsa.Cells[row1, 11].Value = DG_UD_CNT;
                                    wsa.Cells[row1, 12].Value = DG_HD_CNT;
                                    wsa.Cells[row1, 13].Value = CN_UD_CNT + CN_UB_CNT + CN_CN_CNT + CN_HF_CNT;
                                    wsa.Cells[row1, 14].Value = CN_UD_CNT;
                                    wsa.Cells[row1, 15].Value = CN_UB_CNT;
                                    wsa.Cells[row1, 16].Value = CN_CN_CNT;
                                    wsa.Cells[row1, 17].Value = CN_HF_CNT;
                                    if (j.ToString() == "1")
                                    {
                                        wsa.Cells[row1, 18].Value = M3_PF_CNT;
                                        wsa.Cells[row1, 19].Value = M4_PF_CNT;
                                        wsa.Cells[row1, 20].Value = M4_PU_CNT;
                                    }
                                    else
                                    {
                                        wsa.Cells[row1, 18].Value = COST_CN_FEE;
                                        wsa.Cells[row1, 19].Value = COST_LD_FEE;
                                    }
                                }
                                row = row + 3;
                                for (i = 1; i < 21; i++)
                                {
                                    wsa.Cells[row, i].Value = ObjectToDouble(wsa.Cells[row - 1, i].Value) + ObjectToDouble(wsa.Cells[row - 2, i].Value) + ObjectToDouble(wsa.Cells[row - 3, i].Value);
                                }
                                row = row + 3;
                            }
                            row = row - 2; ;
                            for (i = 1; i < 20; i++)
                            {
                                wsa.Cells[row, i].Value = ObjectToDouble(wsa.Cells[row - 1, i].Value) + ObjectToDouble(wsa.Cells[row - 7, i].Value);
                            }
                            wsa.Cells[row, 18].Value = null;
                            wsa.Cells[row, 19].Value = null;
                            //總計
                            row = 5;
                            for (i = 1; i < 3; i++)
                            {
                                double TD_Fund = (double)gmdvm.GetType().GetProperty("TD_FUND_0" + i.ToString()).GetValue(gmdvm, null);
                                double FYR_FUND = (double)gmdvm.GetType().GetProperty("FYR_FUND_0" + i.ToString()).GetValue(gmdvm, null);
                                double REV_FUND = (double)gmdvm.GetType().GetProperty("REV_FUND_0" + i.ToString()).GetValue(gmdvm, null);
                                double COST_FEE = (double)gmdvm.GetType().GetProperty("COST_FEE_0" + i.ToString()).GetValue(gmdvm, null);
                                double REV_COST_FEE = (double)gmdvm.GetType().GetProperty("REV_COST_FEE_0" + i.ToString()).GetValue(gmdvm, null);
                                double SUMPPAY_FEE = (double)gmdvm.GetType().GetProperty("SUMPPAY_FEE_0" + i.ToString()).GetValue(gmdvm, null);
                                double FIN_SAVE = (double)gmdvm.GetType().GetProperty("FIN_SAVE_0" + i.ToString()).GetValue(gmdvm, null);

                                wsa.Cells[row, 1].Value = TD_Fund;
                                wsa.Cells[row, 2].Value = FYR_FUND;
                                wsa.Cells[row, 3].Value = 0;
                                wsa.Cells[row, 4].Value = REV_FUND;
                                wsa.Cells[row, 5].Value = FYR_FUND + 0 + REV_FUND;
                                if (i == 0)
                                {
                                    wsa.Cells[row, 6].Value = gmdvm.PJ_DIST_FUND;
                                }
                                else if (i == 1)
                                {
                                    wsa.Cells[row, 6].Value = ObjectToDouble(wsa.Cells[16, 2].Value);
                                }
                                else
                                {
                                    wsa.Cells[row, 6].Value = ObjectToDouble(wsa.Cells[22, 2].Value);
                                }
                                wsa.Cells[row, 7].Value = COST_FEE;
                                wsa.Cells[row, 8].Value = 0;
                                wsa.Cells[row, 9].Value = REV_COST_FEE;
                                wsa.Cells[row, 10].Value = COST_FEE + 0 + REV_COST_FEE;
                                wsa.Cells[row, 11].Value = SUMPPAY_FEE;
                                wsa.Cells[row, 12].Value = FIN_SAVE;
                                wsa.Cells[row, 13].Value = CalDiv(ObjectToDouble(wsa.Cells[row, 10].Value), ObjectToDouble(wsa.Cells[row, 6].Value));
                                wsa.Cells[row, 14].Value = CalDiv(ObjectToDouble(wsa.Cells[row, 10].Value) + SUMPPAY_FEE + FIN_SAVE, ObjectToDouble(wsa.Cells[row, 6].Value));
                                wsa.Cells[row, 15].Value = CalDiv((COST_FEE + 0 + REV_COST_FEE + SUMPPAY_FEE + FIN_SAVE), FYR_FUND + 0 + REV_FUND);
                                row++;
                            }
                            row = 4;

                            for (i = 1; i < 13; i++)
                            {
                                wsa.Cells[row, i].Value = ObjectToDouble(wsa.Cells[row + 1, i].Value) + ObjectToDouble(wsa.Cells[row + 2, i].Value);
                            }
                            wsa.Cells[row, 13].Value = 0;
                            wsa.Cells[row, 14].Value = 0;
                            wsa.Cells[row, 15].Value = (ObjectToDouble(wsa.Cells[row, 10].Value) + ObjectToDouble(wsa.Cells[row, 11].Value) + ObjectToDouble(wsa.Cells[row, 12].Value)) / ObjectToDouble(wsa.Cells[row, 5].Value);

                            break;
                        case "N":
                            //用地
                            var Detail_a_n = Detail.Select("ADM_DIST='北區' and CATY_CODE='L'");
                            //自辦
                            var Detail_b_n = Detail.Select("ADM_DIST='北區' and IPJ_CODE='01' and CATY_CODE<>'L'");
                            //地方自辦
                            var Detail_c_n = Detail.Select("ADM_DIST='北區' and IPJ_CODE='02' and CATY_CODE<>'L'");
                            //ExcelWorksheet wsbn = excel.Workbook.Worksheets.Copy("分處", "北區分處");
                            Worksheet wsbn = workbook.Worksheets[0];
                            wsbn = CopyWorksheet(workbook, wsbn, "分處", "北區分處");
                            //ExcelWorksheet wscn = excel.Workbook.Worksheets.Copy("自辦", "北區自辦");
                            Worksheet wscn = workbook.Worksheets[0];
                            wscn = CopyWorksheet(workbook, wscn, "自辦", "北區自辦");
                            //ExcelWorksheet wsdn = excel.Workbook.Worksheets.Copy("用地取得", "北區用地取得");
                            Worksheet wsdn = workbook.Worksheets[0];
                            wsdn = CopyWorksheet(workbook, wsdn, "用地取得", "北區用地取得");
                            //ExcelWorksheet wsen = excel.Workbook.Worksheets.Copy("地方自辦", "北區地方自辦");
                            Worksheet wsen = workbook.Worksheets[0];
                            wsen = CopyWorksheet(workbook, wsen, "地方自辦", "北區地方自辦");
                            //北區分處
                            SheetSub(wsbn, gmdvm, PJNAME);
                            //北區用地取得
                            wsdn.Cells[0, 4].Value = "北區用地取得";
                            row = 3;
                            for (i = 0; i < Detail_a_n.Count(); i++)
                            {
                                //wsdn.Cells[4, 1, 5, 13].Copy(wsdn.Cells[row, 1, row + 1, 13]);
                                CopyRange(wsdn, wsdn, 3, 4, 0, 12, row, row + 1, 0, 12);

                                wsdn.Cells[row, 0].Value = i + 1;
                                wsdn.Cells[row, 1].Value = Detail_a_n[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_a_n[i]["TD_CODE"].ToString();
                                wsdn.Cells[row, 2].Value = Detail_a_n[i]["YR_FUND"];
                                wsdn.Cells[row, 3].Value = Detail_a_n[i]["COST_FEE"];
                                wsdn.Cells[row, 4].Value = CalDiv(ObjectToDouble(Detail_a_n[i]["COST_FEE"]), ObjectToDouble(Detail_a_n[i]["YR_FUND"]));
                                wsdn.Cells[row, 6].Value = Detail_a_n[i]["PRE_TF101"];
                                wsdn.Cells[row, 8].Value = Detail_a_n[i]["PRE_TF105"];
                                wsdn.Cells[row, 9].Value = Detail_a_n[i]["PRE_TF106"];
                                wsdn.Cells[row, 10].Value = Detail_a_n[i]["PRE_TF107"];
                                wsdn.Cells[row, 11].Value = Detail_a_n[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_a_n[i]["RESLV_MTHD"].ToString();
                                wsdn.Cells[row, 12].Value = Detail_a_n[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_a_n[i]["CTDELAY_RESN"].ToString();
                                row++;
                                wsdn.Cells[row, 6].Value = Detail_a_n[i]["REL_TF101"];
                                wsdn.Cells[row, 8].Value = Detail_a_n[i]["REL_TF105"];
                                wsdn.Cells[row, 9].Value = Detail_a_n[i]["REL_TF106"];
                                wsdn.Cells[row, 10].Value = Detail_a_n[i]["REL_TF107"];
                                row++;
                            }
                            //北區自辦
                            wscn.Cells[0, 6].Value = "北區自辦工程";
                            row = 3;
                            for (i = 0; i < Detail_b_n.Count(); i++)
                            {
                                //wscn.Cells[4, 1, 5, 17].Copy(wscn.Cells[row, 1, row + 1, 17]);
                                CopyRange(wscn, wscn, 3, 4, 0, 16, row, row + 1, 0, 16);


                                wscn.Cells[row, 0].Value = i + 1;
                                wscn.Cells[row, 1].Value = Detail_b_n[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_b_n[i]["TD_CODE"].ToString();
                                wscn.Cells[row, 3].Value = Detail_b_n[i]["PRE_SCHD"];
                                wscn.Cells[row, 4].Value = Detail_b_n[i]["REL_SCHD"];
                                wscn.Cells[row, 5].Value = ObjectToDouble(Detail_b_n[i]["REL_SCHD"]) - ObjectToDouble(Detail_b_n[i]["PRE_SCHD"]);
                                wscn.Cells[row, 6].Value = Detail_b_n[i]["PRE_TK101"];
                                wscn.Cells[row, 7].Value = Detail_b_n[i]["PRE_TK103"];
                                wscn.Cells[row, 8].Value = Detail_b_n[i]["PRE_TK105"];
                                wscn.Cells[row, 9].Value = Detail_b_n[i]["PRE_TK107"];
                                wscn.Cells[row, 10].Value = Detail_b_n[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_b_n[i]["RESLV_MTHD"].ToString();
                                wscn.Cells[row, 11].Value = Detail_b_n[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_b_n[i]["CTDELAY_RESN"].ToString();
                                wscn.Cells[row, 12].Value = Detail_b_n[i]["SAVED_FEE"];
                                wscn.Cells[row, 13].Value = Detail_b_n[i]["YR_FUND"];
                                wscn.Cells[row, 14].Value = Detail_b_n[i]["COST_FEE"];
                                wscn.Cells[row, 15].Value = Detail_b_n[i]["UPAY_FEE"];
                                row++;
                                wscn.Cells[row, 6].Value = Detail_b_n[i]["REL_TK101"];
                                wscn.Cells[row, 7].Value = Detail_b_n[i]["REL_TK103"];
                                wscn.Cells[row, 8].Value = Detail_b_n[i]["REL_TK105"];
                                wscn.Cells[row, 9].Value = Detail_b_n[i]["REL_TK107"];
                                row++;
                            }
                            //北區地方自辦
                            wsen.Cells[0, 5].Value = "北區地方自辦工程";
                            row = 3;
                            for (i = 0; i < Detail_c_n.Count(); i++)
                            {

                                //wsen.Cells[4, 1, 5, 17].Copy(wsen.Cells[row, 1, row + 1, 17]);
                                CopyRange(wsen, wsen, 3, 4, 0, 16, row, row + 1, 0, 16);


                                wsen.Cells[row, 0].Value = i + 1;
                                wsen.Cells[row, 1].Value = Detail_c_n[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_c_n[i]["TD_CODE"].ToString();
                                wsen.Cells[row, 3].Value = Detail_c_n[i]["PRE_SCHD"];
                                wsen.Cells[row, 4].Value = Detail_c_n[i]["REL_SCHD"];
                                wsen.Cells[row, 5].Value = ObjectToDouble(Detail_c_n[i]["REL_SCHD"]) - ObjectToDouble(Detail_c_n[i]["PRE_SCHD"]);
                                wsen.Cells[row, 6].Value = Detail_c_n[i]["PRE_TK101"];
                                wsen.Cells[row, 7].Value = Detail_c_n[i]["PRE_TK103"];
                                wsen.Cells[row, 8].Value = Detail_c_n[i]["PRE_TK105"];
                                wsen.Cells[row, 9].Value = Detail_c_n[i]["PRE_TK107"];
                                wsen.Cells[row, 10].Value = Detail_c_n[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_c_n[i]["RESLV_MTHD"].ToString();
                                wsen.Cells[row, 11].Value = Detail_c_n[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_c_n[i]["CTDELAY_RESN"].ToString();
                                wsen.Cells[row, 13].Value = Detail_c_n[i]["SAVED_FEE"];
                                wsen.Cells[row, 14].Value = Detail_c_n[i]["YR_FUND"];
                                wsen.Cells[row, 15].Value = Detail_c_n[i]["COST_FEE"];
                                row++;
                                wsen.Cells[row, 6].Value = Detail_c_n[i]["REL_TK101"];
                                wsen.Cells[row, 7].Value = Detail_c_n[i]["REL_TK103"];
                                wsen.Cells[row, 8].Value = Detail_c_n[i]["REL_TK105"];
                                wsen.Cells[row, 9].Value = Detail_c_n[i]["REL_TK107"];
                                row++;
                            }
                            break;
                        case "C":
                            //用地
                            var Detail_a_c = Detail.Select("ADM_DIST='中區' and CATY_CODE='L'");
                            //自辦
                            var Detail_b_c = Detail.Select("ADM_DIST='中區' and IPJ_CODE='01' and CATY_CODE<>'L'");
                            //地方自辦
                            var Detail_c_c = Detail.Select("ADM_DIST='中區' and IPJ_CODE='02' and CATY_CODE<>'L'");
                            //ExcelWorksheet wsbc = excel.Workbook.Worksheets.Copy("分處", "中區分處");
                            //ExcelWorksheet wscc = excel.Workbook.Worksheets.Copy("自辦", "中區自辦");
                            //ExcelWorksheet wsdc = excel.Workbook.Worksheets.Copy("用地取得", "中區用地取得");
                            //ExcelWorksheet wsec = excel.Workbook.Worksheets.Copy("地方自辦", "中區地方自辦");
                            Worksheet wsbc = workbook.Worksheets[0];
                            wsbc = CopyWorksheet(workbook, wsbc, "分處", "中區分處");
                            Worksheet wscc = workbook.Worksheets[0];
                            wscc = CopyWorksheet(workbook, wscc, "自辦", "中區自辦");
                            Worksheet wsdc = workbook.Worksheets[0];
                            wsdc = CopyWorksheet(workbook, wsdc, "用地取得", "中區用地取得");
                            Worksheet wsec = workbook.Worksheets[0];
                            wsec = CopyWorksheet(workbook, wsec, "地方自辦", "中區地方自辦");
                            //中區分處
                            SheetSub(wsbc, gmdvm, PJNAME);
                            //中區用地取得
                            wsdc.Cells[0, 4].Value = "中區用地取得";
                            row = 3;

                            for (i = 0; i < Detail_a_c.Count(); i++)
                            {
                                //wsdc.Cells[4, 1, 5, 13].Copy(wsdc.Cells[row, 1, row + 1, 13]);
                                CopyRange(wsdc, wsdc, 3, 4, 0, 12, row, row + 1, 0, 12);
                                //wsec.Cells["A4:Q4"].Copy(wsec.Cells["A6:Q6"]);

                                wsdc.Cells[row, 0].Value = i + 1;
                                wsdc.Cells[row, 1].Value = Detail_a_c[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_a_c[i]["TD_CODE"].ToString();
                                wsdc.Cells[row, 2].Value = Detail_a_c[i]["YR_FUND"];
                                wsdc.Cells[row, 3].Value = Detail_a_c[i]["COST_FEE"];
                                wsdc.Cells[row, 4].Value = CalDiv(ObjectToDouble(Detail_a_c[i]["COST_FEE"]), ObjectToDouble(Detail_a_c[i]["YR_FUND"]));
                                wsdc.Cells[row, 6].Value = Detail_a_c[i]["PRE_TF101"];
                                wsdc.Cells[row, 8].Value = Detail_a_c[i]["PRE_TF105"];
                                wsdc.Cells[row, 9].Value = Detail_a_c[i]["PRE_TF106"];
                                wsdc.Cells[row, 10].Value = Detail_a_c[i]["PRE_TF107"];
                                wsdc.Cells[row, 11].Value = Detail_a_c[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_a_c[i]["RESLV_MTHD"].ToString();
                                wsdc.Cells[row, 12].Value = Detail_a_c[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_a_c[i]["CTDELAY_RESN"].ToString();
                                row++;
                                wsdc.Cells[row, 6].Value = Detail_a_c[i]["REL_TF101"];
                                wsdc.Cells[row, 8].Value = Detail_a_c[i]["REL_TF105"];
                                wsdc.Cells[row, 9].Value = Detail_a_c[i]["REL_TF106"];
                                wsdc.Cells[row, 10].Value = Detail_a_c[i]["REL_TF107"];
                                row++;
                            }
                            //中區自辦
                            wscc.Cells[0, 6].Value = "中區自辦工程";
                            row = 3;
                            for (i = 0; i < Detail_b_c.Count(); i++)
                            {
                                //wscc.Cells[4, 1, 5, 17].Copy(wscc.Cells[row, 1, row + 1, 17]);
                                CopyRange(wscc, wscc, 3, 4, 0, 16, row, row + 1, 0, 16);


                                wscc.Cells[row, 0].Value = i + 1;
                                wscc.Cells[row, 1].Value = Detail_b_c[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_b_c[i]["TD_CODE"].ToString();
                                wscc.Cells[row, 3].Value = Detail_b_c[i]["PRE_SCHD"];
                                wscc.Cells[row, 4].Value = Detail_b_c[i]["REL_SCHD"];
                                wscc.Cells[row, 5].Value = ObjectToDouble(Detail_b_c[i]["REL_SCHD"]) - ObjectToDouble(Detail_b_c[i]["PRE_SCHD"]);
                                wscc.Cells[row, 6].Value = Detail_b_c[i]["PRE_TK101"];
                                wscc.Cells[row, 7].Value = Detail_b_c[i]["PRE_TK103"];
                                wscc.Cells[row, 8].Value = Detail_b_c[i]["PRE_TK105"];
                                wscc.Cells[row, 9].Value = Detail_b_c[i]["PRE_TK107"];
                                wscc.Cells[row, 10].Value = Detail_b_c[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_b_c[i]["RESLV_MTHD"].ToString();
                                wscc.Cells[row, 11].Value = Detail_b_c[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_b_c[i]["CTDELAY_RESN"].ToString();
                                wscc.Cells[row, 12].Value = Detail_b_c[i]["SAVED_FEE"];
                                wscc.Cells[row, 13].Value = Detail_b_c[i]["YR_FUND"];
                                wscc.Cells[row, 14].Value = Detail_b_c[i]["COST_FEE"];
                                wscc.Cells[row, 15].Value = Detail_b_c[i]["UPAY_FEE"];

                                row++;
                                wscc.Cells[row, 6].Value = Detail_b_c[i]["REL_TK101"];
                                wscc.Cells[row, 7].Value = Detail_b_c[i]["REL_TK103"];
                                wscc.Cells[row, 8].Value = Detail_b_c[i]["REL_TK105"];
                                wscc.Cells[row, 9].Value = Detail_b_c[i]["REL_TK107"];
                                row++;
                            }
                            //中區地方自辦
                            wsec.Cells[0, 5].Value = "中區地方自辦工程";
                            row = 3;
                            var q = Detail_c_c.Count();
                            for (i = 0; i < Detail_c_c.Count(); i++)
                            {
                                //wsec.Cells[4, 1, 5, 17].Copy(wsec.Cells[row, 1, row + 1, 17]);
                                CopyRange(wsec, wsec, 3, 4, 0, 16, row, row + 1, 0, 16);

                                wsec.Cells[row, 0].Value = i + 1;
                                wsec.Cells[row, 1].Value = Detail_c_c[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_c_c[i]["TD_CODE"].ToString();
                                wsec.Cells[row, 3].Value = Detail_c_c[i]["PRE_SCHD"];
                                wsec.Cells[row, 4].Value = Detail_c_c[i]["REL_SCHD"];
                                wsec.Cells[row, 5].Value = ObjectToDouble(Detail_c_c[i]["REL_SCHD"]) - ObjectToDouble(Detail_c_c[i]["PRE_SCHD"]);
                                wsec.Cells[row, 6].Value = Detail_c_c[i]["PRE_TK101"];
                                wsec.Cells[row, 7].Value = Detail_c_c[i]["PRE_TK103"];
                                wsec.Cells[row, 8].Value = Detail_c_c[i]["PRE_TK105"];
                                wsec.Cells[row, 9].Value = Detail_c_c[i]["PRE_TK107"];
                                wsec.Cells[row, 10].Value = Detail_c_c[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_c_c[i]["RESLV_MTHD"].ToString();
                                wsec.Cells[row, 11].Value = Detail_c_c[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_c_c[i]["CTDELAY_RESN"].ToString();
                                wsec.Cells[row, 13].Value = Detail_c_c[i]["SAVED_FEE"];
                                wsec.Cells[row, 14].Value = Detail_c_c[i]["YR_FUND"];
                                wsec.Cells[row, 15].Value = Detail_c_c[i]["COST_FEE"];
                                row++;
                                wsec.Cells[row, 6].Value = Detail_c_c[i]["REL_TK101"];
                                wsec.Cells[row, 7].Value = Detail_c_c[i]["REL_TK103"];
                                wsec.Cells[row, 8].Value = Detail_c_c[i]["REL_TK105"];
                                wsec.Cells[row, 9].Value = Detail_c_c[i]["REL_TK107"];
                                row++;
                            }

                            break;
                        case "S":
                            //用地
                            var Detail_a_s = Detail.Select("ADM_DIST='南區' and CATY_CODE='L'");
                            //自辦
                            var Detail_b_s = Detail.Select("ADM_DIST='南區' and IPJ_CODE='01' and CATY_CODE<>'L'");
                            //地方自辦
                            var Detail_c_s = Detail.Select("ADM_DIST='南區' and IPJ_CODE='02' and CATY_CODE<>'L'");
                            //ExcelWorksheet wsbs = excel.Workbook.Worksheets.Copy("分處", "南區分處");
                            //ExcelWorksheet wscs = excel.Workbook.Worksheets.Copy("自辦", "南區自辦");
                            //ExcelWorksheet wsds = excel.Workbook.Worksheets.Copy("用地取得", "南區用地取得");
                            //ExcelWorksheet wses = excel.Workbook.Worksheets.Copy("地方自辦", "南區地方自辦");
                            Worksheet wsbs = workbook.Worksheets[0];
                            wsbs = CopyWorksheet(workbook, wsbs, "分處", "南區分處");
                            Worksheet wscs = workbook.Worksheets[0];
                            wscs = CopyWorksheet(workbook, wscs, "自辦", "南區自辦");
                            Worksheet wsds = workbook.Worksheets[0];
                            wsds = CopyWorksheet(workbook, wsds, "用地取得", "南區用地取得");
                            Worksheet wses = workbook.Worksheets[0];
                            wses = CopyWorksheet(workbook, wses, "地方自辦", "南區地方自辦");
                            //南區分處
                            SheetSub(wsbs, gmdvm, PJNAME);
                            //南區用地取得
                            wsds.Cells[0, 4].Value = "南區用地取得";
                            row = 3;
                            for (i = 0; i < Detail_a_s.Count(); i++)
                            {
                                //wsds.Cells[4, 1, 5, 13].Copy(wsds.Cells[row, 1, row + 1, 13]);
                                CopyRange(wsds, wsds, 3, 4, 0, 12, row, row + 1, 0, 12);

                                wsds.Cells[row, 0].Value = i + 1;
                                wsds.Cells[row, 1].Value = Detail_a_s[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_a_s[i]["TD_CODE"].ToString();
                                wsds.Cells[row, 2].Value = Detail_a_s[i]["YR_FUND"];
                                wsds.Cells[row, 3].Value = Detail_a_s[i]["COST_FEE"];
                                wsds.Cells[row, 4].Value = CalDiv(ObjectToDouble(Detail_a_s[i]["COST_FEE"]), ObjectToDouble(Detail_a_s[i]["YR_FUND"]));
                                wsds.Cells[row, 6].Value = Detail_a_s[i]["PRE_TF101"];
                                wsds.Cells[row, 8].Value = Detail_a_s[i]["PRE_TF105"];
                                wsds.Cells[row, 9].Value = Detail_a_s[i]["PRE_TF106"];
                                wsds.Cells[row, 10].Value = Detail_a_s[i]["PRE_TF107"];
                                wsds.Cells[row, 11].Value = Detail_a_s[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_a_s[i]["RESLV_MTHD"].ToString();
                                wsds.Cells[row, 12].Value = Detail_a_s[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_a_s[i]["CTDELAY_RESN"].ToString();
                                row++;
                                wsds.Cells[row, 6].Value = Detail_a_s[i]["REL_TF101"];
                                wsds.Cells[row, 8].Value = Detail_a_s[i]["REL_TF105"];
                                wsds.Cells[row, 9].Value = Detail_a_s[i]["REL_TF106"];
                                wsds.Cells[row, 10].Value = Detail_a_s[i]["REL_TF107"];
                                row++;
                            }
                            //南區自辦
                            wscs.Cells[0, 6].Value = "南區自辦工程";
                            row = 3;
                            for (i = 0; i < Detail_b_s.Count(); i++)
                            {
                                //wscs.Cells[4, 1, 5, 17].Copy(wscs.Cells[row, 1, row + 1, 17]);
                                CopyRange(wscs, wscs, 3, 4, 0, 16, row, row + 1, 0, 16);


                                wscs.Cells[row, 0].Value = i + 1;
                                wscs.Cells[row, 1].Value = Detail_b_s[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_b_s[i]["TD_CODE"].ToString();
                                wscs.Cells[row, 3].Value = Detail_b_s[i]["PRE_SCHD"];
                                wscs.Cells[row, 4].Value = Detail_b_s[i]["REL_SCHD"];
                                wscs.Cells[row, 5].Value = ObjectToDouble(Detail_b_s[i]["REL_SCHD"]) - ObjectToDouble(Detail_b_s[i]["PRE_SCHD"]);
                                wscs.Cells[row, 6].Value = Detail_b_s[i]["PRE_TK101"];
                                wscs.Cells[row, 7].Value = Detail_b_s[i]["PRE_TK103"];
                                wscs.Cells[row, 8].Value = Detail_b_s[i]["PRE_TK105"];
                                wscs.Cells[row, 9].Value = Detail_b_s[i]["PRE_TK107"];
                                wscs.Cells[row, 10].Value = Detail_b_s[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_b_s[i]["RESLV_MTHD"].ToString();
                                wscs.Cells[row, 11].Value = Detail_b_s[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_b_s[i]["CTDELAY_RESN"].ToString();
                                wscs.Cells[row, 12].Value = Detail_b_s[i]["SAVED_FEE"];
                                wscs.Cells[row, 13].Value = Detail_b_s[i]["YR_FUND"];
                                wscs.Cells[row, 14].Value = Detail_b_s[i]["COST_FEE"];
                                wscs.Cells[row, 15].Value = Detail_b_s[i]["UPAY_FEE"];
                                row++;
                                wscs.Cells[row, 6].Value = Detail_b_s[i]["REL_TK101"];
                                wscs.Cells[row, 7].Value = Detail_b_s[i]["REL_TK103"];
                                wscs.Cells[row, 8].Value = Detail_b_s[i]["REL_TK105"];
                                wscs.Cells[row, 9].Value = Detail_b_s[i]["REL_TK107"];
                                row++;
                            }
                            //南區地方自辦
                            wses.Cells[0, 5].Value = "南區地方自辦工程";
                            row = 3;
                            for (i = 0; i < Detail_c_s.Count(); i++)
                            {

                                //wses.Cells[4, 1, 5, 17].Copy(wses.Cells[row, 1, row + 1, 17]);
                                CopyRange(wses, wses, 3, 4, 0, 16, row, row + 1, 0, 16);


                                wses.Cells[row, 0].Value = i + 1;
                                wses.Cells[row, 1].Value = Detail_c_s[i]["TD_NAME"].ToString() + Environment.NewLine + Detail_c_s[i]["TD_CODE"].ToString();
                                wses.Cells[row, 3].Value = Detail_c_s[i]["PRE_SCHD"];
                                wses.Cells[row, 4].Value = Detail_c_s[i]["REL_SCHD"];
                                wses.Cells[row, 5].Value = ObjectToDouble(Detail_c_s[i]["REL_SCHD"]) - ObjectToDouble(Detail_c_s[i]["PRE_SCHD"]);
                                wses.Cells[row, 6].Value = Detail_c_s[i]["PRE_TK101"];
                                wses.Cells[row, 7].Value = Detail_c_s[i]["PRE_TK103"];
                                wses.Cells[row, 8].Value = Detail_c_s[i]["PRE_TK105"];
                                wses.Cells[row, 9].Value = Detail_c_s[i]["PRE_TK107"];
                                wses.Cells[row, 10].Value = Detail_c_s[i]["EXECT_DESC"].ToString() + Environment.NewLine + Detail_c_s[i]["RESLV_MTHD"].ToString();
                                wses.Cells[row, 11].Value = Detail_c_s[i]["DELAY_RESN"].ToString() + Environment.NewLine + Detail_c_s[i]["CTDELAY_RESN"].ToString();
                                wses.Cells[row, 13].Value = Detail_c_s[i]["SAVED_FEE"];
                                wses.Cells[row, 14].Value = Detail_c_s[i]["YR_FUND"];
                                wses.Cells[row, 15].Value = Detail_c_s[i]["COST_FEE"];
                                row++;
                                wses.Cells[row, 6].Value = Detail_c_s[i]["REL_TK101"];
                                wses.Cells[row, 7].Value = Detail_c_s[i]["REL_TK103"];
                                wses.Cells[row, 8].Value = Detail_c_s[i]["REL_TK105"];
                                wses.Cells[row, 9].Value = Detail_c_s[i]["REL_TK107"];
                                row++;
                            }

                            break;
                        default:
                            break;
                    }
                }

                for (i = 0; i < 5; i++)
                {
                    workbook.Worksheets.RemoveAt(0);
                }

                string filename = "A-3.跨年度一億元以上工程計畫執行情形表" + DateTime.Now.ToString("yyyyMMddHHmmss");
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
            else
            {
                string message = "查無資料";
                return RedirectToAction("State3", new { message = message });
            }
        }

        /// <summary>
        /// B-1.修正預算管考報表
        /// </summary>
        /// <param name="PID"> 工程編號 </param>
        /// <param name="period"> 期數 </param>
        /// <param name="Subject">  </param>
        /// <param name="SYR"> 年度 </param>
        /// <param name="ODS"> .ods </param>
        /// <param name="Choose">  </param>
        /// <returns></returns>
        public ActionResult ReportPrint6(string PID, string period, string[] Subject, string SYR, string ODS, string Choose)
        {
            var PJNAME = "";
            //string ADV_UID = (string)Session["UsrId"];
            string ADV_UID = User.Identity.Name;
            string sACC_ID = "";
            string aDV_DP = "";
            //string Dept = (string)Session["DeptId"];
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

            /// B160000(國土署下水道永續營運組)篩選限制調整為 A270500 / A270500(汙設隊帳號) 相同
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

            if (PID.Length == 3)
            {
                decimal PIDD = decimal.Parse(PID);
                PJNAME = db2.PJMFORM.Where(m => m.PID == PIDD).FirstOrDefault().PJ_NAME;
            }
            else
            {
                PJNAME = "前瞻基礎建設計畫-縣市管河川及區域排水整理改善計畫";
            }

            DataSet DsMain1 = GetSP_RCD0010_UDF003_New(PID, SEQ_NO, sACC_ID, ADV_UID, aDV_DP, Dept, Choose);
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
                ACC02_FUND_FINAL = ObjectToDouble(DtS00.Rows[0]["ACC02_FUND"].ToString()) + ObjectToDouble(DtS00.Rows[1]["ACC02_FUND"].ToString());
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

            if (SEQ_NO != "0")
            {
                SEQ_NO = Convert.ToString(Convert.ToInt32(SEQ_NO) - 1);
                DataSet DsMain2 = GetSP_RCD0010_UDF003_New(PID, SEQ_NO, sACC_ID, ADV_UID, aDV_DP, Dept, Choose);
                DtS000 = DsMain2.Tables["DataTable"];
                DtS11 = DsMain2.Tables["DataTable"];
                DtS12 = DsMain2.Tables["DataTable1"];
                DtS13 = DsMain2.Tables["DataTable2"];
                DtS14 = DsMain2.Tables["DataTable3"];
                DtS15 = DsMain2.Tables["DataTable4"];
                DtS16 = DsMain2.Tables["DataTable5"];

                if (DtS00.Rows.Count > 1)
                {
                    ACC02_FUND_FINAL = ObjectToDouble(DtS00.Rows[0]["ACC02_FUND"].ToString()) + ObjectToDouble(DtS00.Rows[1]["ACC02_FUND"].ToString());
                }
                else
                {
                    ACC02_FUND_FINAL = 0;
                }

            }

            var result = db.SMCPOSTN.
                GroupBy(p => p.CITY).
                Select(g => new
                {
                    CPA_CITY_ID = g.Max(p => p.CPA_CITY_ID),
                    CITY = g.Key
                }).
                OrderBy(p => p.CPA_CITY_ID).ToList();

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("CPA_CITY_ID", typeof(int));
            dataTable.Columns.Add("CITY", typeof(string));

            foreach (var item in result)
            {
                dataTable.Rows.Add(item.CPA_CITY_ID, item.CITY);
            }

            var DtCITY = getReport02_CITY();

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);

            #region ================== Aspose.Cells 標題 ==================

            /// ================== 總表(公辦) ==================
            DataTable DtH01 = GetDtH01(DtS01, DtS11, "1", DtCITY);
            DataTable DtH01H = GetHeadDt("1", PJNAME, period, SYR);

            /// ================== 總表(BOT) ==================
            DataTable DtH04 = GetDtH01(DtS04, DtS14, "4", DtCITY);
            DataTable DtH04H = GetHeadDt("4", PJNAME, period, SYR);

            /// ================== 總表 ==================
            DataTable DtH00H = GetHeadDt("0", PJNAME, period, SYR);
            DataTable DtH00 = GetDtH00(DtH01, DtH04, DtH00H);

            /// ================== 彙整明細 ==================
            /// ================== 彙整明細(簡易版) ==================
            DataTable DtH06H = GetHeadDt("6", PJNAME, SEQ_NO, SYR);
            List<B1DtBack> Lb1Dtback = new List<B1DtBack>();
            List<B1DtFooter> Lb1DtFooter = new List<B1DtFooter>();
            GetDtH02(DtS06, DtS16, "6", DtCITY, SEQ_NO, Lb1Dtback, Lb1DtFooter);

            #endregion

            var filepath = Server.MapPath(@"~/Templates/PJV0010_02.xlsx");
            var file = new FileInfo(filepath);
            var workbook = new Aspose.Cells.Workbook(filepath);

            #region ================== Aspose.Cells Sheet ==================

            /// ================== Aspose.Cells 總表(公辦) 標題 ==================
            workbook.Worksheets.Add("總表(公辦)");
            Worksheet ws2 = workbook.Worksheets["總表(公辦)"];
            Worksheet wsa = workbook.Worksheets[0];

            /// ================== Aspose.Cells 總表(BOT) 標題 ==================
            workbook.Worksheets.Add("總表(BOT)");
            Worksheet ws3 = workbook.Worksheets["總表(BOT)"];
            Worksheet wsb = workbook.Worksheets[2];

            /// ================== Aspose.Cells 總表 標題 ==================
            workbook.Worksheets.Add("總表");
            Worksheet ws1 = workbook.Worksheets["總表"];
            Worksheet wsc = workbook.Worksheets[4];
            ws1.MoveTo(5);

            /// ================== Aspose.Cells 彙整明細 標題 ==================
            workbook.Worksheets.Add("彙整明細");
            Worksheet ws4 = workbook.Worksheets["彙整明細"];
            Worksheet wsd = workbook.Worksheets[1];

            /// ================== Aspose.Cells 彙整明細簡易版 標題 ==================
            workbook.Worksheets.Add("彙整明細簡易版");
            Worksheet ws5 = workbook.Worksheets["彙整明細簡易版"];
            Worksheet wse = workbook.Worksheets[3];

            #endregion

            #region ================== Aspose.Cells 總表(公辦) ==================

            int row = 4;
            CopyRange(wsa, ws2, 0, 2, 0, 21, 0, 2, 0, 21);

            ws2.Cells[0, 1].Value = DtH01H.Rows[0]["Title"].ToString();
            ws2.Cells[1, 3].Value = DtH01H.Rows[0]["TITLE1"].ToString();
            ws2.Cells[1, 7].Value = DtH01H.Rows[0]["TITLE2"].ToString();

            for (int i = 0; i < DtH01.Rows.Count; i++)
            {
                CopyRange(wsa, ws2, 4, 7, 0, 21, row, row + 3, 0, 21);
                row = row + 5;
            }

            CopyRange(wsa, ws2, 9, 12, 0, 21, row, row + 3, 0, 21);
            row = 4;
            for (int i = 0; i < DtH01.Rows.Count; i++)
            {
                double CntC11 = ObjectToDouble(DtH01.Rows[i]["CntC11"].ToString());
                double CntP11 = ObjectToDouble(DtH01.Rows[i]["CntP11"].ToString());
                double CntP01 = ObjectToDouble(DtH01.Rows[i]["CntP01"].ToString());
                double CntC01 = ObjectToDouble(DtH01.Rows[i]["CntC01"].ToString());

                /// 金額
                double FundC11 = ObjectToDouble(DtH01.Rows[i]["FundC11"].ToString());
                double FundP11 = ObjectToDouble(DtH01.Rows[i]["FundP11"].ToString());
                double FundP01 = ObjectToDouble(DtH01.Rows[i]["FundP01"].ToString());
                double FundC01 = ObjectToDouble(DtH01.Rows[i]["FundC01"].ToString());

                double CntC22 = ObjectToDouble(DtH01.Rows[i]["CntC22"].ToString());
                double CntP22 = ObjectToDouble(DtH01.Rows[i]["CntP22"].ToString());
                double CntP02 = ObjectToDouble(DtH01.Rows[i]["CntP02"].ToString());
                double CntC02 = ObjectToDouble(DtH01.Rows[i]["CntC02"].ToString());

                /// 金額
                double FundC22 = ObjectToDouble(DtH01.Rows[i]["FundC22"].ToString());
                double FundP22 = ObjectToDouble(DtH01.Rows[i]["FundP22"].ToString());
                double FundP02 = ObjectToDouble(DtH01.Rows[i]["FundP02"].ToString());
                double FundC02 = ObjectToDouble(DtH01.Rows[i]["FundC02"].ToString());

                /// 各月份處理
                double M_01 = ObjectToDouble(DtH01.Rows[i]["M_01"].ToString());
                double M_02 = ObjectToDouble(DtH01.Rows[i]["M_02"].ToString());
                double M_03 = ObjectToDouble(DtH01.Rows[i]["M_03"].ToString());
                double M_04 = ObjectToDouble(DtH01.Rows[i]["M_04"].ToString());
                double M_05 = ObjectToDouble(DtH01.Rows[i]["M_05"].ToString());
                double M_06 = ObjectToDouble(DtH01.Rows[i]["M_06"].ToString());
                double M_07 = ObjectToDouble(DtH01.Rows[i]["M_07"].ToString());
                double M_08 = ObjectToDouble(DtH01.Rows[i]["M_08"].ToString());
                double M_09 = ObjectToDouble(DtH01.Rows[i]["M_09"].ToString());
                double M_10 = ObjectToDouble(DtH01.Rows[i]["M_10"].ToString());
                double M_11 = ObjectToDouble(DtH01.Rows[i]["M_11"].ToString());
                double M_12 = ObjectToDouble(DtH01.Rows[i]["M_12"].ToString());

                ws2.Cells[row, 0].Value = DtH01.Rows[i]["CITY"].ToString();
                ws2.Cells[row, 4].Value = CntP11;
                ws2.Cells[row, 3].Value = CntC11;
                ws2.Cells[row, 5].Value = CntP11 + CntC11;
                ws2.Cells[row, 6].Value = FundP11 + FundC11 + FundP22 + FundC22;
                ws2.Cells[row, 7].Value = CntP01;
                ws2.Cells[row, 8].Value = CntC01;
                ws2.Cells[row, 9].Value = CntP01 + CntC01;
                ws2.Cells[row, 10].Value = FundP01 + FundC01 + FundP02 + FundC02;
                ws2.Cells[row, 11].Value = FundP11 + FundC11 + FundP22 + FundC22 - (FundP01 + FundC01 + FundP02 + FundC02);
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
                SumCntC11 += ObjectToDouble(DtH01.Rows[i]["CntC11"].ToString());
                SumCntP11 += ObjectToDouble(DtH01.Rows[i]["CntP11"].ToString());
                SumCntP01 += ObjectToDouble(DtH01.Rows[i]["CntP01"].ToString());
                SumCntC01 += ObjectToDouble(DtH01.Rows[i]["CntC01"].ToString());
                SumFundC11 += ObjectToDouble(DtH01.Rows[i]["FundC11"].ToString());
                SumFundP11 += ObjectToDouble(DtH01.Rows[i]["FundP11"].ToString());
                SumFundP01 += ObjectToDouble(DtH01.Rows[i]["FundP01"].ToString());
                SumFundC01 += ObjectToDouble(DtH01.Rows[i]["FundC01"].ToString());
                SumCntC22 += ObjectToDouble(DtH01.Rows[i]["CntC22"].ToString());
                SumCntP22 += ObjectToDouble(DtH01.Rows[i]["CntP22"].ToString());
                SumCntP02 += ObjectToDouble(DtH01.Rows[i]["CntP02"].ToString());
                SumCntC02 += ObjectToDouble(DtH01.Rows[i]["CntC02"].ToString());
                SumFundC22 += ObjectToDouble(DtH01.Rows[i]["FundC22"].ToString());
                SumFundP22 += ObjectToDouble(DtH01.Rows[i]["FundP22"].ToString());
                SumFundP02 += ObjectToDouble(DtH01.Rows[i]["FundP02"].ToString());
                SumFundC02 += ObjectToDouble(DtH01.Rows[i]["FundC02"].ToString());
                SumM_01 += ObjectToDouble(DtH01.Rows[i]["M_01"].ToString());
                SumM_02 += ObjectToDouble(DtH01.Rows[i]["M_02"].ToString());
                SumM_03 += ObjectToDouble(DtH01.Rows[i]["M_03"].ToString());
                SumM_04 += ObjectToDouble(DtH01.Rows[i]["M_04"].ToString());
                SumM_05 += ObjectToDouble(DtH01.Rows[i]["M_05"].ToString());
                SumM_06 += ObjectToDouble(DtH01.Rows[i]["M_06"].ToString());
                SumM_07 += ObjectToDouble(DtH01.Rows[i]["M_07"].ToString());
                SumM_08 += ObjectToDouble(DtH01.Rows[i]["M_08"].ToString());
                SumM_09 += ObjectToDouble(DtH01.Rows[i]["M_09"].ToString());
                SumM_10 += ObjectToDouble(DtH01.Rows[i]["M_10"].ToString());
                SumM_11 += ObjectToDouble(DtH01.Rows[i]["M_11"].ToString());
                SumM_12 += ObjectToDouble(DtH01.Rows[i]["M_12"].ToString());
            }

            ws2.Cells[row, 4].Value = SumCntP11;
            ws2.Cells[row, 3].Value = SumCntC11;
            ws2.Cells[row, 5].Value = SumCntP11 + SumCntC11;
            ws2.Cells[row, 6].Value = SumFundP11 + SumFundC11 + SumFundP22 + SumFundC22;
            ws2.Cells[row, 7].Value = SumCntP01;
            ws2.Cells[row, 8].Value = SumCntC01;
            ws2.Cells[row, 9].Value = SumCntP01 + SumCntC01;
            ws2.Cells[row, 10].Value = SumFundP01 + SumFundC01 + SumFundP02 + SumFundC02;
            ws2.Cells[row, 11].Value = SumFundP11 + SumFundC11 + SumFundP22 + SumFundC22 - (SumFundP01 + SumFundC01 + SumFundP02 + SumFundC02);
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

            ws2.AutoFitColumns();

            #endregion

            #region ================== Aspose.Cells 總表(BOT) ==================

            row = 4;
            CopyRange(wsb, ws3, 0, 2, 0, 25, 0, 2, 0, 25);
            ws3.Cells[0, 1].Value = DtH04H.Rows[0]["Title"].ToString();
            ws3.Cells[1, 3].Value = DtH04H.Rows[0]["TITLE1"].ToString();
            ws3.Cells[1, 9].Value = DtH04H.Rows[0]["TITLE2"].ToString();

            for (int i = 0; i < DtH04.Rows.Count; i++)
            {
                CopyRange(wsb, ws3, 4, 7, 0, 25, row, row + 3, 0, 25);
                row = row + 5;
            }

            CopyRange(wsb, ws3, 9, 12, 0, 25, row, row + 3, 0, 25);
            row = 4;

            for (int i = 0; i < DtH04.Rows.Count; i++)
            {
                double CntC11 = ObjectToDouble(DtH04.Rows[i]["CntC11"].ToString());
                double CntP11 = ObjectToDouble(DtH04.Rows[i]["CntP11"].ToString());
                double CntM11 = ObjectToDouble(DtH04.Rows[i]["CntM11"].ToString());
                double CntL11 = ObjectToDouble(DtH04.Rows[i]["CntL11"].ToString());
                double CntM01 = ObjectToDouble(DtH04.Rows[i]["CntM01"].ToString());
                double CntL01 = ObjectToDouble(DtH04.Rows[i]["CntL01"].ToString());
                double FundM11 = ObjectToDouble(DtH04.Rows[i]["FundM11"].ToString());
                double FundL11 = ObjectToDouble(DtH04.Rows[i]["FundL11"].ToString());
                double FundM01 = ObjectToDouble(DtH04.Rows[i]["FundM01"].ToString());
                double FundL01 = ObjectToDouble(DtH04.Rows[i]["FundL01"].ToString());
                double CntM22 = ObjectToDouble(DtH04.Rows[i]["CntM22"].ToString());
                double CntL22 = ObjectToDouble(DtH04.Rows[i]["CntL22"].ToString());
                double CntM02 = ObjectToDouble(DtH04.Rows[i]["CntM02"].ToString());
                double CntL02 = ObjectToDouble(DtH04.Rows[i]["CntL02"].ToString());
                double FundM22 = ObjectToDouble(DtH04.Rows[i]["FundM22"].ToString());
                double FundL22 = ObjectToDouble(DtH04.Rows[i]["FundL22"].ToString());
                double FundM02 = ObjectToDouble(DtH04.Rows[i]["FundM02"].ToString());
                double FundL02 = ObjectToDouble(DtH04.Rows[i]["FundL02"].ToString());
                double CntP01 = ObjectToDouble(DtH04.Rows[i]["CntP01"].ToString());
                double CntC01 = ObjectToDouble(DtH04.Rows[i]["CntC01"].ToString());
                double FundC11 = ObjectToDouble(DtH04.Rows[i]["FundC11"].ToString());
                double FundP11 = ObjectToDouble(DtH04.Rows[i]["FundP11"].ToString());
                double FundP01 = ObjectToDouble(DtH04.Rows[i]["FundP01"].ToString());
                double FundC01 = ObjectToDouble(DtH04.Rows[i]["FundC01"].ToString());
                double CntC22 = ObjectToDouble(DtH04.Rows[i]["CntC22"].ToString());
                double CntP22 = ObjectToDouble(DtH04.Rows[i]["CntP22"].ToString());
                double CntP02 = ObjectToDouble(DtH04.Rows[i]["CntP02"].ToString());
                double CntC02 = ObjectToDouble(DtH04.Rows[i]["CntC02"].ToString());
                double FundC22 = ObjectToDouble(DtH04.Rows[i]["FundC22"].ToString());
                double FundP22 = ObjectToDouble(DtH04.Rows[i]["FundP22"].ToString());
                double FundP02 = ObjectToDouble(DtH04.Rows[i]["FundP02"].ToString());
                double FundC02 = ObjectToDouble(DtH04.Rows[i]["FundC02"].ToString());
                double M_01 = ObjectToDouble(DtH04.Rows[i]["M_01"].ToString());
                double M_02 = ObjectToDouble(DtH04.Rows[i]["M_02"].ToString());
                double M_03 = ObjectToDouble(DtH04.Rows[i]["M_03"].ToString());
                double M_04 = ObjectToDouble(DtH04.Rows[i]["M_04"].ToString());
                double M_05 = ObjectToDouble(DtH04.Rows[i]["M_05"].ToString());
                double M_06 = ObjectToDouble(DtH04.Rows[i]["M_06"].ToString());
                double M_07 = ObjectToDouble(DtH04.Rows[i]["M_07"].ToString());
                double M_08 = ObjectToDouble(DtH04.Rows[i]["M_08"].ToString());
                double M_09 = ObjectToDouble(DtH04.Rows[i]["M_09"].ToString());
                double M_10 = ObjectToDouble(DtH04.Rows[i]["M_10"].ToString());
                double M_11 = ObjectToDouble(DtH04.Rows[i]["M_11"].ToString());
                double M_12 = ObjectToDouble(DtH04.Rows[i]["M_12"].ToString());

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
                Sum3CntC11 += ObjectToDouble(DtH04.Rows[i]["CntC11"].ToString());
                Sum3CntP11 += ObjectToDouble(DtH04.Rows[i]["CntP11"].ToString());
                Sum3CntM11 += ObjectToDouble(DtH04.Rows[i]["CntM11"].ToString());
                Sum3CntL11 += ObjectToDouble(DtH04.Rows[i]["CntL11"].ToString());
                Sum3CntM01 += ObjectToDouble(DtH04.Rows[i]["CntM01"].ToString());
                Sum3CntL01 += ObjectToDouble(DtH04.Rows[i]["CntL01"].ToString());
                Sum3FundM11 += ObjectToDouble(DtH04.Rows[i]["FundM11"].ToString());
                Sum3FundL11 += ObjectToDouble(DtH04.Rows[i]["FundL11"].ToString());
                Sum3FundM01 += ObjectToDouble(DtH04.Rows[i]["FundM01"].ToString());
                Sum3FundL01 += ObjectToDouble(DtH04.Rows[i]["FundL01"].ToString());
                Sum3CntM22 += ObjectToDouble(DtH04.Rows[i]["CntM22"].ToString());
                Sum3CntL22 += ObjectToDouble(DtH04.Rows[i]["CntL22"].ToString());
                Sum3CntM02 += ObjectToDouble(DtH04.Rows[i]["CntM02"].ToString());
                Sum3CntL02 += ObjectToDouble(DtH04.Rows[i]["CntL02"].ToString());
                Sum3FundM22 += ObjectToDouble(DtH04.Rows[i]["FundM22"].ToString());
                Sum3FundL22 += ObjectToDouble(DtH04.Rows[i]["FundL22"].ToString());
                Sum3FundM02 += ObjectToDouble(DtH04.Rows[i]["FundM02"].ToString());
                Sum3FundL02 += ObjectToDouble(DtH04.Rows[i]["FundL02"].ToString());
                Sum3CntP01 += ObjectToDouble(DtH04.Rows[i]["CntP01"].ToString());
                Sum3CntC01 += ObjectToDouble(DtH04.Rows[i]["CntC01"].ToString());
                Sum3FundC11 += ObjectToDouble(DtH04.Rows[i]["FundC11"].ToString());
                Sum3FundP11 += ObjectToDouble(DtH04.Rows[i]["FundP11"].ToString());
                Sum3FundP01 += ObjectToDouble(DtH04.Rows[i]["FundP01"].ToString());
                Sum3FundC01 += ObjectToDouble(DtH04.Rows[i]["FundC01"].ToString());
                Sum3CntC22 += ObjectToDouble(DtH04.Rows[i]["CntC22"].ToString());
                Sum3CntP22 += ObjectToDouble(DtH04.Rows[i]["CntP22"].ToString());
                Sum3CntP02 += ObjectToDouble(DtH04.Rows[i]["CntP02"].ToString());
                Sum3CntC02 += ObjectToDouble(DtH04.Rows[i]["CntC02"].ToString());
                Sum3FundC22 += ObjectToDouble(DtH04.Rows[i]["FundC22"].ToString());
                Sum3FundP22 += ObjectToDouble(DtH04.Rows[i]["FundP22"].ToString());
                Sum3FundP02 += ObjectToDouble(DtH04.Rows[i]["FundP02"].ToString());
                Sum3FundC02 += ObjectToDouble(DtH04.Rows[i]["FundC02"].ToString());
                Sum3M_01 += ObjectToDouble(DtH04.Rows[i]["M_01"].ToString());
                Sum3M_02 += ObjectToDouble(DtH04.Rows[i]["M_02"].ToString());
                Sum3M_03 += ObjectToDouble(DtH04.Rows[i]["M_03"].ToString());
                Sum3M_04 += ObjectToDouble(DtH04.Rows[i]["M_04"].ToString());
                Sum3M_05 += ObjectToDouble(DtH04.Rows[i]["M_05"].ToString());
                Sum3M_06 += ObjectToDouble(DtH04.Rows[i]["M_06"].ToString());
                Sum3M_07 += ObjectToDouble(DtH04.Rows[i]["M_07"].ToString());
                Sum3M_08 += ObjectToDouble(DtH04.Rows[i]["M_08"].ToString());
                Sum3M_09 += ObjectToDouble(DtH04.Rows[i]["M_09"].ToString());
                Sum3M_10 += ObjectToDouble(DtH04.Rows[i]["M_10"].ToString());
                Sum3M_11 += ObjectToDouble(DtH04.Rows[i]["M_11"].ToString());
                Sum3M_12 += ObjectToDouble(DtH04.Rows[i]["M_12"].ToString());
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

            #endregion

            #region ================== Aspose.Cells 總表 ==================

            CopyRange(wsc, ws1, 0, 2, 0, 16, 0, 2, 0, 16);
            ws1.Cells[0, 0].Value = DtH00H.Rows[0]["Title"].ToString();
            ws1.Cells[1, 2].Value = DtH00H.Rows[0]["TITLE1"].ToString();
            ws1.Cells[1, 7].Value = DtH00H.Rows[0]["TITLE2"].ToString(); // 7-8
            row = 3;

            // Excel 表格
            for (int i = 0; i < DtH00.Rows.Count; i++)
            {
                CopyRange(wsc, ws1, 3, 4, 0, 16, row, row + 1, 0, 16);
                row = row + 2;
            }

            CopyRange(wsc, ws1, 6, 9, 0, 16, row, row + 3, 0, 16);

            // Excel 資料處理
            ws1.Cells[0, 0].Value = DtH00H.Rows[0]["Title"].ToString();
            ws1.Cells[1, 2].Value = DtH00H.Rows[0]["TITLE1"].ToString();
            ws1.Cells[1, 7].Value = DtH00H.Rows[0]["TITLE2"].ToString(); ///7-8
            row = 3;

            double FEE_G_1_1 = 0;
            double FEE_G_1_2 = 0;
            double FEE_G_2_1 = 0;
            double FEE_G_2_2 = 0;
            double FEE_B_1_1 = 0;
            double FEE_B_1_2 = 0;
            double FEE_B_2_1 = 0;
            double FEE_B_2_2 = 0;
            double ACC02_FUND_FINAL_SUM = ACC02_FUND_FINAL;
            double ACC02_FUND_SECOND_SUM = ACC02_FUND_SECOND;

            for (int i = 0; i < DtH00.Rows.Count; i++)
            {
                if (i != 0)
                {
                    ACC02_FUND_FINAL = 0;
                    ACC02_FUND_SECOND = 0;
                }

                FEE_G_1_1 = ObjectToDouble(DtH00.Rows[i]["FEE_G_1_1"].ToString());
                FEE_G_1_2 = ObjectToDouble(DtH00.Rows[i]["FEE_G_1_2"].ToString());
                FEE_G_2_1 = ObjectToDouble(DtH00.Rows[i]["FEE_G_2_1"].ToString());
                FEE_G_2_2 = ObjectToDouble(DtH00.Rows[i]["FEE_G_2_2"].ToString());
                FEE_B_1_1 = ObjectToDouble(DtH00.Rows[i]["FEE_B_1_1"].ToString());
                FEE_B_1_2 = ObjectToDouble(DtH00.Rows[i]["FEE_B_1_2"].ToString());
                FEE_B_2_1 = ObjectToDouble(DtH00.Rows[i]["FEE_B_2_1"].ToString());
                FEE_B_2_2 = ObjectToDouble(DtH00.Rows[i]["FEE_B_2_2"].ToString());
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
                ws1.Cells[row, 4].Value = 0; // 預設 0

                ws1.Cells[row, 5].Value = FEE_B_1_1 + FEE_B_1_2;
                ws1.Cells[row, 7].Value = FEE_B_2_1;
                ws1.Cells[row, 8].Value = FEE_B_2_2;
                ws1.Cells[row, 9].Value = 0; // 預設 0

                ws1.Cells[row, 10].Value = FEE_B_2_1 + FEE_B_2_2;
                ws1.Cells[row, 12].Value = FEE_B_2_1 - FEE_B_1_1;
                ws1.Cells[row, 13].Value = FEE_B_2_2 - FEE_B_1_2;
                ws1.Cells[row, 14].Value = 0;// 預設 0

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

            for (int i = 0; i < DtH00.Rows.Count; i++)
            {
                SumFEE_G_1_1 += ObjectToDouble(DtH00.Rows[i]["FEE_G_1_1"].ToString());
                SumFEE_G_1_2 += ObjectToDouble(DtH00.Rows[i]["FEE_G_1_2"].ToString());
                SumFEE_G_2_1 += ObjectToDouble(DtH00.Rows[i]["FEE_G_2_1"].ToString());
                SumFEE_G_2_2 += ObjectToDouble(DtH00.Rows[i]["FEE_G_2_2"].ToString());
                SumFEE_B_1_1 += ObjectToDouble(DtH00.Rows[i]["FEE_B_1_1"].ToString());
                SumFEE_B_1_2 += ObjectToDouble(DtH00.Rows[i]["FEE_B_1_2"].ToString());
                SumFEE_B_2_1 += ObjectToDouble(DtH00.Rows[i]["FEE_B_2_1"].ToString());
                SumFEE_B_2_2 += ObjectToDouble(DtH00.Rows[i]["FEE_B_2_2"].ToString());
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

            double FEE_1_1 = ObjectToDouble(DtH00H.Rows[0]["FEE_1_1"].ToString());
            double FEE_1_2 = ObjectToDouble(DtH00H.Rows[0]["FEE_1_2"].ToString());
            double FEE_2_1 = ObjectToDouble(DtH00H.Rows[0]["FEE_2_1"].ToString());
            double FEE_2_2 = ObjectToDouble(DtH00H.Rows[0]["FEE_2_2"].ToString());
            double FEE_G_DIFF = ObjectToDouble(DtH00H.Rows[0]["FEE_G_DIFF"].ToString());
            double FEE_B_DIFF = ObjectToDouble(DtH00H.Rows[0]["FEE_B_DIFF"].ToString());

            ws1.Cells[row, 2].Value = FEE_1_1;
            ws1.Cells[row, 3].Value = FEE_1_2;
            ws1.Cells[row, 4].Value = ACC02_FUND_FINAL_SUM;
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
            ws1.AutoFitColumns();

            #endregion

            #region ================== Aspose.Cells 彙整明細 ==================

            ws4.PageSetup.PaperSize = Aspose.Cells.PaperSizeType.PaperA4;            // 紙張大小為 A4 
            ws4.PageSetup.Orientation = Aspose.Cells.PageOrientationType.Landscape;  // 紙張為橫向
            ws4.PageSetup.FitToPagesWide = 1;  // 寬度縮放為 1 頁
            ws4.PageSetup.FitToPagesTall = 0;  // 高度不限頁數
            ws4.PageSetup.TopMargin = 1.91;    // [設定窄邊界]上邊界 (英吋)
            ws4.PageSetup.BottomMargin = 1.91; // [設定窄邊界]下邊界 (英吋)
            ws4.PageSetup.LeftMargin = 0.64;   // [設定窄邊界]左邊界 (英吋)
            ws4.PageSetup.RightMargin = 0.64;  // [設定窄邊界]右邊界 (英吋)
            ws4.PageSetup.HeaderMargin = 0.76; // [設定窄邊界]頁首邊界 (英吋)
            ws4.PageSetup.FooterMargin = 0.76; // [設定窄邊界]頁尾邊界 (英吋)

            var distinctCityClss = Lb1Dtback.Select(m => m.CityCls).Distinct().ToList();

            row = 0;

            for (int i = 0; i < Lb1Dtback.Select(m => m.CityCls).Distinct().Count(); i++)
            {
                ws4.HorizontalPageBreaks.Add(28, 0);

                if (i > 0)
                {
                    ws4.HorizontalPageBreaks.Add(row, 0);
                }

                CopyRange(wsd, ws4, 0, 25, 0, 21, row, row + 25, 0, 21);
                ws4.Cells[row, 0].Value = DtH06H.Rows[0]["Title"].ToString();
                ws4.Cells[row + 1, 2].Value = Lb1DtFooter[i].CITY.ToString() + "調整預算表";
                ws4.Cells[row + 2, 7].Value = Lb1DtFooter[i].TITLE1.ToString();
                ws4.Cells[row + 2, 10].Value = Lb1DtFooter[i].TITLE2.ToString();
                row = row + 4;

                var list04 = Lb1Dtback.Where(m => m.CityCls == distinctCityClss[i]).OrderBy(m => m.Order).ToList();

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

                    double ratio = 0;
                    double.TryParse(list04[j].CEN_SP_RATIO?.ToString(), out ratio);
                    ws4.Cells[row, 6].PutValue(ratio / 100);  // 存成數值 (0.9823)
                    var style = ws4.Cells[row, 6].GetStyle();
                    style.Custom = "0%";                       // 顯示為百分比 (98%)
                    ws4.Cells[row, 6].SetStyle(style);
                    ws4.Cells[row, 7].PutValue(double.TryParse(list04[j].CEN_FUND2?.ToString(), out double v7) ? v7 : 0);
                    ws4.Cells[row, 8].PutValue(double.TryParse(list04[j].LOC_FUND2?.ToString(), out double v8) ? v8 : 0);
                    ws4.Cells[row, 9].PutValue(CalPlus(list04[j].CEN_FUND2, list04[j].LOC_FUND2));
                    ws4.Cells[row, 10].PutValue(double.TryParse(list04[j].CEN_FUND1?.ToString(), out double v10) ? v10 : 0);
                    ws4.Cells[row, 11].PutValue(double.TryParse(list04[j].LOC_FUND1?.ToString(), out double v11) ? v11 : 0);
                    ws4.Cells[row, 12].PutValue(CalPlus(list04[j].CEN_FUND1, list04[j].LOC_FUND1));
                    ws4.Cells[row, 13].PutValue(CalMinus(list04[j].CEN_FUND1, list04[j].CEN_FUND2));
                    ws4.Cells[row, 14].PutValue(double.TryParse(list04[j].COST_FEE?.ToString(), out double v) ? v : 0);
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
                        CopyRange(wsd, ws4, 0, 25, 0, 21, row, row + 25, 0, 21);
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
                    else if (x == 0 && j == Convert.ToInt32(list04.Select(m => m.Order).Count()) - 1)
                    {
                        row += 2;
                    }
                }

                ws4.AutoFitColumns();
            }

            #endregion

            #region ================== Aspose.Cells 彙整明細(簡易版) ==================

            row = 0;

            ws5.PageSetup.PaperSize = Aspose.Cells.PaperSizeType.PaperA4;            // 紙張大小為 A4 
            ws5.PageSetup.Orientation = Aspose.Cells.PageOrientationType.Landscape;  // 紙張為橫向
            ws5.PageSetup.FitToPagesWide = 1;  // 寬度縮放為 1 頁
            ws5.PageSetup.FitToPagesTall = 0;  // 高度不限頁數
            ws5.PageSetup.TopMargin = 1.91;    // [設定窄邊界]上邊界 (英吋)
            ws5.PageSetup.BottomMargin = 1.91; // [設定窄邊界]下邊界 (英吋)
            ws5.PageSetup.LeftMargin = 0.64;   // [設定窄邊界]左邊界 (英吋)
            ws5.PageSetup.RightMargin = 0.64;  // [設定窄邊界]右邊界 (英吋)
            ws5.PageSetup.HeaderMargin = 0.76; // [設定窄邊界]頁首邊界 (英吋)
            ws5.PageSetup.FooterMargin = 0.76; // [設定窄邊界]頁尾邊界 (英吋)

            var distinctCityCls = Lb1Dtback.Select(m => m.CityCls).Distinct().ToList();

            for (int i = 0; i < Lb1Dtback.Select(m => m.CityCls).Distinct().Count(); i++)
            {
                ws5.HorizontalPageBreaks.Add(28, 0);

                if (i > 0)
                {
                    ws5.HorizontalPageBreaks.Add(row, 0);
                }

                // 報表標題
                CopyRange(wse, ws5, 0, 25, 0, 14, row, row + 25, 0, 14);
                ws5.Cells[row, 0].Value = DtH06H.Rows[0]["Title"].ToString();
                ws5.Cells[row + 1, 0].Value = Lb1DtFooter[i].CITY.ToString() + "調整預算表";
                ws5.Cells[row + 2, 7].Value = Lb1DtFooter[i].TITLE1.ToString();
                ws5.Cells[row + 2, 10].Value = Lb1DtFooter[i].TITLE2.ToString();
                row = row + 4;

                // 縣市排列
                // 2024/12/25 修改 => 全縣市 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 22
                var list04 = Lb1Dtback.Where(m => m.CityCls == distinctCityCls[i]).OrderBy(m => m.CityCls).ToList();

                // 報表資料處理
                for (int j = 0; j < Convert.ToInt32(list04.Select(m => m.CityCls == i).Count()); j++)
                {
                    var x = (j + 1) % 11;
                    var y = 11 - x;

                    ws5.Cells[row, 0].Value = list04[j].SERIAL;
                    ws5.Cells[row, 1].Value = list04[j].SIP_TD_CODE;
                    ws5.Cells[row, 2].Value = list04[j].FD_CODE;
                    ws5.Cells[row, 3].Value = list04[j].TD_NAME;
                    ws5.Cells[row, 3].GetMergedRange().RowHeight = 35;
                    ws5.Cells[row, 4].Value = list04[j].CATY_DESC;
                    ws5.Cells[row, 5].Value = list04[j].ACC_DESC;
                    //ws5.Cells[row, 6].Value = list04[j].CEN_SP_RATIO;
                    ws5.Cells[row, 6].Value = Convert.ToDecimal(list04[j].CEN_SP_RATIO);
                    ws5.Cells[row, 7].Value = Convert.ToDecimal(list04[j].CEN_FUND2);
                    ws5.Cells[row, 8].Value = Convert.ToDecimal(list04[j].LOC_FUND2);
                    ws5.Cells[row, 9].Value = CalPlus(list04[j].CEN_FUND2, list04[j].LOC_FUND2);
                    ws5.Cells[row, 10].Value = Convert.ToDecimal(list04[j].CEN_FUND1);
                    ws5.Cells[row, 11].Value = Convert.ToDecimal(list04[j].LOC_FUND1);
                    ws5.Cells[row, 12].Value = CalPlus(list04[j].CEN_FUND1, list04[j].LOC_FUND1);
                    ws5.Cells[row, 13].Value = CalMinus(list04[j].CEN_FUND1, list04[j].CEN_FUND2);
                    //ws5.Cells[row, 14].Value = list04[j].COST_FEE;
                    ws5.Cells[row, 14].PutValue(double.TryParse(list04[j].COST_FEE?.ToString(), out double v) ? v : 0);


                    row = row + 2;

                    if (x == 0 && j != Convert.ToInt32(list04.Select(m => m.Order).Count()) - 1)
                    {
                        row = row + 2;
                        CopyRange(wse, ws5, 0, 25, 0, 14, row, row + 25, 0, 14);
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
                    else if (x == 0 && j == Convert.ToInt32(list04.Select(m => m.Order).Count()) - 1)
                    {
                        row += 2;
                    }
                }

                // 各欄位寬度設定
                ws5.Cells.SetColumnWidth(0, 4);
                ws5.Cells.SetColumnWidth(1, 19);
                ws5.Cells.SetColumnWidth(2, 15);
                ws5.Cells.SetColumnWidth(3, 53);
                ws5.Cells.SetColumnWidth(4, 6);
                ws5.Cells.SetColumnWidth(5, 6);
                ws5.Cells.SetColumnWidth(6, 10);
                ws5.Cells.SetColumnWidth(7, 10);
                ws5.Cells.SetColumnWidth(8, 10);
                ws5.Cells.SetColumnWidth(9, 10);
                ws5.Cells.SetColumnWidth(10, 10);
                ws5.Cells.SetColumnWidth(11, 10);
                ws5.Cells.SetColumnWidth(12, 10);
                ws5.Cells.SetColumnWidth(13, 10);
                ws5.Cells.SetColumnWidth(14, 10);
            }

            #endregion

            for (int i = 0; i < 5; i++)
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

        public ActionResult ReportPrint6_2(string PID, string period, string[] Subject, string SYR, string ODS, string Choose)
        {
            Debug.WriteLine($"ReportPrint6_2");
            Debug.WriteLine($"PID:{PID}");
            Debug.WriteLine($"period:{period}");
            Debug.WriteLine($"SYR:{SYR}");

            Aspose.Cells.License licenseE = new Aspose.Cells.License();
            var path2 = Server.MapPath("~/Aspose/Aspose.Total.lic");
            licenseE.SetLicense(path2);
            var filePath = Server.MapPath(@"~/Templates/PJV0010_02.xlsx");
            var workBook = new Aspose.Cells.Workbook(filePath);

            Aspose.Cells.Worksheet sheet1 = workBook.Worksheets["Sheet1"];
            Aspose.Cells.Worksheet sheet2 = workBook.Worksheets["Sheet2"];
            Aspose.Cells.Worksheet sheet3 = workBook.Worksheets["Sheet3"];
            Aspose.Cells.Worksheet sheet4 = workBook.Worksheets["Sheet4"];
            Aspose.Cells.Worksheet sheet5 = workBook.Worksheets["Sheet5"];

            // string ADV_UID = (string)Session["UsrId"];
            string ADV_UID = User.Identity.Name;
            string sACC_ID = "";
            string aDV_DP = "";
            // string Dept = (string)Session["DeptId"];
            string Dept = db.OGCDPUR
                .Where(x => x.USR_ID == ADV_UID && x.DATA_TP == "D")
                .Select(x => x.DEPT_NO)
                .FirstOrDefault();
            string SEQ_NO = period; // 期數

            DataSet ds = new DataSet();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();

            #region 部門代碼
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

            if (Dept == "B160000")  //B160000(國土署下水道永續營運組)篩選限制調整為與A270500 / A270500(汙設隊帳號)相同

            {
                var result1 = db.OGCDEPT.Where(m => m.DEPT_NO.Contains("A2705")).ToList();

                aDV_DP += !String.IsNullOrEmpty(aDV_DP) ? "," : "";

                foreach (var item in result1)
                {
                    aDV_DP += item.DEPT_NO + ",";
                }
                aDV_DP = aDV_DP.TrimEnd(',');
            }
            #endregion

            //ds = Report_Source.SP_RCD0010_UDF003(PID, SEQ_NO, SYR, sACC_ID, ADV_UID, aDV_DP, Dept, Choose, connectionString3);
            //dt1 = ds.Tables["dt1"];
            //dt2 = ds.Tables["dt2"];

            //#region =========== 總表 ===========
            //sheet1.Name = "總表";
            //Report_Source.emeTableSummary(workBook, sheet1, dt1, dt2, PID, SYR, SEQ_NO, sACC_ID, ADV_UID, aDV_DP, Dept, Choose);
            //#endregion

            //#region =========== 總表(公辦) ===========
            //sheet2.Name = "總表(公辦)";
            //Report_Source.emeTableSummary_PublicOffices(workBook, sheet2, dt1, dt2, PID, SEQ_NO, SYR, sACC_ID, ADV_UID, aDV_DP, Dept, Choose);
            //#endregion

            //#region =========== 總表(BOT) ===========
            //sheet3.Name = "總表(BOT)";
            //Report_Source.emeTableSummary_BOT(workBook, sheet3, dt1, dt2, PID, SEQ_NO, SYR, sACC_ID, ADV_UID, aDV_DP, Dept, Choose);
            //#endregion

            //#region =========== 彙整明細 ===========
            //sheet4.Name = "彙整明細";
            //Report_Source.emeTableSummaryDetails(workBook, sheet4, dt1, dt2, PID, SEQ_NO, SYR, sACC_ID, ADV_UID, aDV_DP, Dept, Choose);
            //#endregion

            //#region =========== 彙整明細簡易版 ===========
            //sheet5.Name = "彙整明細簡易版";
            //Report_Source.emeTableSummaryDetails_Simplify(workBook, sheet5, dt1, dt2, PID, SEQ_NO, SYR, sACC_ID, ADV_UID, aDV_DP, Dept, Choose);
            //#endregion


            string fileName = "B-1.修正預算管考報表" + DateTime.Now.ToString("yyyyMMddHHmmss");
            string directoryPath = Server.MapPath(@"~\Temp\");
            string newfilePath = System.IO.Path.Combine(directoryPath, fileName + ".xlsx");
            workBook.Save(newfilePath);

            if (ODS == "1")
            {
                var odsResult = OdsPrint(filePath, fileName);
                return odsResult;
            }
            else
            {
                using (var stream = new MemoryStream())
                {
                    workBook.Save(stream, SaveFormat.Xlsx);
                    stream.Seek(0, SeekOrigin.Begin);
                    var bytes = stream.ToArray();
                    stream.Close();
                    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName + ".xlsx");
                }
            }
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

        private void GetDtH02(DataTable DtS1, DataTable DtS2, string sType, DataTable DtCITY, string SEQ_NO, List<B1DtBack> Lb1Dtback, List<B1DtFooter> Lb1DtFooter)
        {
            SEQ_NO = Convert.ToString(Convert.ToInt32(SEQ_NO) + 1);

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
                        GetDtTitleName(SEQ_NO, b1DtFooter);
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
                                    //bNewNumYesNo = true;
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
                                        //if (((iShowTitle) % 11) == 0)
                                        //{
                                        //    SetSpaceFootPage(sType, SEQ_NO, Lb1Dtback, x, iShowTitle, Drow2);
                                        //    Lb1Dtback.Remove(b1DtBack2);
                                        //}

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

                        if (NullList(b1DtBack2))
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

        public static bool NullList(B1DtBack b)
        {
            if (b.SERIAL == null && b.SIP_TD_CODE == null && b.TD_NAME == null && b.FD_CODE == null && b.CATY_DESC == null && b.ACC_DESC == null && b.CEN_SP_RATIO == null && b.CEN_FUND1 == null && b.LOC_FUND1 == null && b.COST_FEE == null && b.FUND_DESC == null && b.CEN_FUND2 == null && b.LOC_FUND2 == null)
            {
                return false;
            }
            return true;
        }

        private void SetSpaceFootPage(string sType, string SEQ_NO, List<B1DtBack> Lb1Dtback, int x, int iShowTitle, DataRow Drow1)
        {
            List<B1DtFooter> ldf = new List<B1DtFooter>();
            B1DtFooter df = new B1DtFooter();
            df.CITY = Drow1["CITY"].ToString();
            GetDtH02TitleName(SEQ_NO, df);
            ldf.Add(df);
        }

        private void GetDtH02TitleName(string SEQ_NO, B1DtFooter df)
        {
            if (SEQ_NO == "0")
            {
                df.TITLE2 = "原預算 工作計畫項目";
            }
            else
            {
                df.TITLE2 = "第" + SEQ_NO + "次修正工作計畫項目";
                string sSEQ_NO = Convert.ToString(Convert.ToInt32(SEQ_NO) - 1);
                if (sSEQ_NO == "0")
                {
                    df.TITLE1 = "原預算 工作計畫項目";
                }
                else
                {
                    df.TITLE1 = "第" + sSEQ_NO + "次修正工作計畫項目";
                }
            }
        }

        private void SetMonthSumF(B1DtFooter b1DtFooter, DataRow Drow, string sExStr)
        {
            if (Drow != null)
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtFooter.GetType().GetProperty("M_" + i.ToString("00") + sExStr).SetValue(b1DtFooter, ObjectToDouble(b1DtFooter.GetType().GetProperty("M_" + i.ToString("00") + sExStr)?.GetValue(b1DtFooter)) + StringToDouble(Drow["M_" + i.ToString("00")].ToString()));

                }
            }
            else
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtFooter.GetType().GetProperty("M_" + i.ToString("00") + sExStr).SetValue(b1DtFooter, ObjectToDouble(b1DtFooter.GetType().GetProperty("M_" + i.ToString("00") + sExStr)?.GetValue(b1DtFooter)) + 0);
                }
            }
        }

        private void SetMonthSumB(B1DtBack b1DtBack, DataRow Drow, string sExStr)
        {
            if (Drow != null)
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtBack.GetType().GetProperty("M_" + i.ToString("00") + sExStr).SetValue(b1DtBack, ObjectToDouble(b1DtBack.GetType().GetProperty("M_" + i.ToString("00") + sExStr)?.GetValue(b1DtBack)) + StringToDouble(Drow["M_" + i.ToString("00")].ToString()));

                }
            }
            else
            {
                for (int i = 1; i < 13; i++)
                {
                    b1DtBack.GetType().GetProperty("M_" + i.ToString("00") + sExStr).SetValue(b1DtBack, ObjectToDouble(b1DtBack.GetType().GetProperty("M_" + i.ToString("00") + sExStr)?.GetValue(b1DtBack)) + 0);
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
            f.GetType().GetProperty("CEN_FUNDSUM" + sNumStr).SetValue(f, ConvertDoubleValue(GetNumSumStrb1footer(f, Drow, "CEN_FUNDSUM" + sNumStr, "CEN_FUND")));
            f.GetType().GetProperty("LOC_FUNDSUM" + sNumStr).SetValue(f, ConvertDoubleValue(GetNumSumStrb1footer(f, Drow, "LOC_FUNDSUM" + sNumStr, "LOC_FUND")));
            f.GetType().GetProperty("FUNDSUM" + sNumStr).SetValue(f, ConvertDoubleValue(GetNumSumStrb1footer(f, Drow, "FUNDSUM" + sNumStr, "CEN_FUND")));
            f.GetType().GetProperty("FUNDSUM" + sNumStr).SetValue(f, ConvertDoubleValue(GetNumSumStrb1footer(f, Drow, "FUNDSUM" + sNumStr, "LOC_FUND")));
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

            DataRow NewRw = null;
            for (int i = 0; i < DtS1.Rows.Count; i++)
            {
                NewRw = DtBack.NewRow();
                NewRw["CITY"] = DtS1.Rows[i]["CITY"].ToString();
                NewRw["FEE_G_1_1"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["FundP11"].ToString())) + Convert.ToDouble(StringToDouble(DtS1.Rows[i]["FundC11"].ToString()));
                NewRw["FEE_G_1_2"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["FundP22"].ToString())) + Convert.ToDouble(StringToDouble(DtS1.Rows[i]["FundC22"].ToString()));
                NewRw["FEE_G_2_1"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["FundP01"].ToString())) + Convert.ToDouble(StringToDouble(DtS1.Rows[i]["FundC01"].ToString()));
                NewRw["FEE_G_2_2"] = Convert.ToDouble(StringToDouble(DtS1.Rows[i]["FundP02"].ToString())) + Convert.ToDouble(StringToDouble(DtS1.Rows[i]["FundC02"].ToString()));
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
                    NewRw["FEE_B_1_1"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundP11"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundC11"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundM11"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundL11"].ToString()));
                    NewRw["FEE_B_1_2"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundP22"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundC22"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundM22"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundL22"].ToString()));
                    NewRw["FEE_B_2_1"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundP01"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundC01"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundM01"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundL01"].ToString()));
                    NewRw["FEE_B_2_2"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundP02"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundC02"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundM02"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundL02"].ToString()));
                    DtBack.Rows.Add(NewRw);
                }
                else
                {
                    DtBack.Rows[index]["FEE_B_1_1"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundP11"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundC11"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundM11"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundL11"].ToString()));
                    DtBack.Rows[index]["FEE_B_1_2"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundP22"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundC22"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundM22"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundL22"].ToString()));
                    DtBack.Rows[index]["FEE_B_2_1"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundP01"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundC01"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundM01"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundL01"].ToString()));
                    DtBack.Rows[index]["FEE_B_2_2"] = Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundP02"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundC02"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundM02"].ToString())) + Convert.ToDouble(StringToDouble(DtS2.Rows[i]["FundL02"].ToString()));
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

        public static double ObjectToDouble(object objValue)
        {
            double doubleValue;

            if (objValue != null)
            {
                if (Double.TryParse(objValue.ToString(), out doubleValue))
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

        public static DataTable GetHeadDt(string sType, string PJNAME, string SEQ_NO, string SYR)
        {
            DataTable DtBack = new DataTable();
            DtBack.Columns.Add("Year");
            DtBack.Columns.Add("PJName");
            DtBack.Columns.Add("Title");

            switch (sType)
            {
                case "0":
                case "1":
                case "4":
                    DtBack.Columns.Add("TITLE1");
                    DtBack.Columns.Add("TITLE2");
                    break;
            }

            DataRow NewRw = DtBack.NewRow();

            NewRw["Year"] = SYR;
            NewRw["PJName"] = PJNAME;

            switch (sType)
            {
                case "0":
                    NewRw["Title"] = SYR + "年度 " + PJNAME + " 預算分配表";
                    break;
                case "1":
                    NewRw["Title"] = SYR + "年度 " + PJNAME + " 修正預算總表(公辦)";
                    break;
                case "2":
                    NewRw["Title"] = SYR + "年度 " + PJNAME + " (公辦部分-署辦)";
                    break;
                case "3":
                    NewRw["Title"] = SYR + "年度 " + PJNAME + " (公辦部分-地方辦)";
                    break;
                case "4":
                    NewRw["Title"] = SYR + "年度 " + PJNAME + " 修正預算總表(BOT)";
                    break;
                case "5":
                    NewRw["Title"] = SYR + "年度 " + PJNAME + " (BOT部分)";
                    break;
                case "6":
                    NewRw["Title"] = SYR + "年度 " + PJNAME + " 彙整明細";
                    break;
            }

            switch (sType)
            {
                case "0":
                case "1":
                case "4":

                    if (SEQ_NO == "0")
                    {
                        NewRw["TITLE2"] = "原預算";
                    }
                    else
                    {
                        NewRw["TITLE2"] = "第" + SEQ_NO + "次修正";

                        SEQ_NO = Convert.ToString(Convert.ToInt32(SEQ_NO) - 1);

                        if (SEQ_NO == "0")
                        {
                            NewRw["TITLE1"] = "原預算";
                        }
                        else
                        {
                            NewRw["TITLE1"] = "第" + SEQ_NO + "次修正";
                        }
                    }

                    break;
            }

            DtBack.Rows.Add(NewRw);

            return DtBack;
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

            DtBack.Columns.Add("ACC02_FUND");///////新加入的
            DtBack.Columns.Add("ACC03_FUND");
            DtBack.Columns.Add("ACC04_FUND");
            DtBack.Columns.Add("ACC02_FUNDSUM"); //////新加入的
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

                            NewRw["ACC02_FUND"] = GetNumSum(NewRw, Drow1, "ACC02_FUND"); /// 新增項目
                            NewRw["ACC03_FUND"] = GetNumSum(NewRw, Drow1, "ACC03_FUND");
                            NewRw["ACC04_FUND"] = GetNumSum(NewRw, Drow1, "ACC04_FUND");

                            NewRw["ACC02_FUNDSUM"] = GetNumSum(NewRw, Drow1, "ACC02_FUNDSUM", "ACC02_FUND"); /// 新增項目
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

                                    NewRw["ACC02_FUNDSUM"] = GetNumSub(Drow2, NewRw, "ACC02_FUND", "ACC02_FUNDSUM"); /// 新增項目
                                    NewRw["ACC03_FUNDSUM"] = GetNumSub(Drow2, NewRw, "ACC03_FUND", "ACC03_FUNDSUM");
                                    NewRw["ACC04_FUNDSUM"] = GetNumSub(Drow2, NewRw, "ACC04_FUND", "ACC04_FUNDSUM");

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

        public static string GetNumSub(DataRow NewRw1, DataRow Drow2, string FidName1, string FidName2)
        {
            if (Drow2[FidName2] != DBNull.Value)
            {
                if (NewRw1[FidName1] != DBNull.Value)
                {
                    return Convert.ToString(Convert.ToDouble(Drow2[FidName2].ToString()) - Convert.ToDouble(NewRw1[FidName1].ToString()));
                }
                else
                {
                    return Drow2[FidName2].ToString();
                }
            }
            else
            {
                if (NewRw1[FidName1] != DBNull.Value)
                {
                    return "0";
                }
                else
                {
                    return "";
                }
            }
        }

        private void SetBackValue(DataRow Drow, DataRow NewRw, string sTD_CATY, string sIPJ_CODE)
        {
            Debug.WriteLine($"ReportController/SetBackValue");
            Debug.WriteLine($"sTD_CATY:{sTD_CATY}");
            Debug.WriteLine($"sIPJ_CODE:{sIPJ_CODE}");

            NewRw["CNT" + sTD_CATY + sIPJ_CODE] = Drow["CNT"].ToString();
            NewRw["FUND" + sTD_CATY + sIPJ_CODE] = Drow["FUND"].ToString();
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

        public static async Task<DataTable> GetSP_RCD0010_UDF002Async(string PID, string SYR, string SMN, string[] Site, string aDV_DP, string aDV_UID, string Dept)
        {
            string aDM_DIST = "";

            if (!string.IsNullOrEmpty(PID) && !string.IsNullOrEmpty(SYR) && !string.IsNullOrEmpty(SMN))
            {
                using (var db3 = new CMEPMSEntities())
                {
                    var pidParam = new SqlParameter("@PID", SqlDbType.Decimal) { Precision = 18, Scale = 0, Value = Convert.ToDecimal(PID) };
                    var syrParam = new SqlParameter("@SYR", SqlDbType.VarChar) { Size = 10, Value = SYR };
                    var smnParam = new SqlParameter("@SMN", SqlDbType.VarChar) { Size = 10, Value = SMN };
                    var admDistParam = new SqlParameter("@ADM_DIST", SqlDbType.VarChar) { Size = 20, Value = aDM_DIST };
                    var advDpParam = new SqlParameter("@ADV_DP", SqlDbType.VarChar) { Size = 2000, Value = aDV_DP };
                    var advUidParam = new SqlParameter("@ADV_UID", SqlDbType.VarChar) { Size = 20, Value = aDV_UID };
                    var DeptParam = new SqlParameter("@Dept", SqlDbType.VarChar) { Size = 20, Value = Dept };
                    var pdstFundParam = new SqlParameter("@PDST_FUND", SqlDbType.Decimal) { Precision = 19, Scale = 0, Direction = ParameterDirection.Output };

                    var connection = (SqlConnection)db3.Database.Connection;
                    var command = new SqlCommand("EXEC CMEPMS.dbo.SP_RCD0010_UDF002S @PID, @SYR, @SMN, @ADM_DIST, @ADV_DP, @ADV_UID, @Dept, @PDST_FUND OUTPUT", connection);
                    command.Parameters.Add(pidParam);
                    command.Parameters.Add(syrParam);
                    command.Parameters.Add(smnParam);
                    command.Parameters.Add(admDistParam);
                    command.Parameters.Add(advDpParam);
                    command.Parameters.Add(advUidParam);
                    command.Parameters.Add(DeptParam);
                    command.Parameters.Add(pdstFundParam);

                    await connection.OpenAsync();
                    var reader = await command.ExecuteReaderAsync();
                    var dataTable = new DataTable();
                    dataTable.Load(reader); // 使用同步的 Load 方法

                    var pdstFundValue = pdstFundParam.Value != DBNull.Value ? Convert.ToDecimal(pdstFundParam.Value) : 0;

                    if (Convert.ToInt32(pdstFundValue) == 0)
                    {
                        DataRow dr = dataTable.NewRow();
                        try
                        {
                            dr[0] = Convert.ToDecimal(pdstFundValue).ToString();
                        }
                        catch
                        {
                            dr[0] = 0;
                        }
                        dr[1] = "";
                        for (int i = 2; i < dataTable.Columns.Count; i++)
                        {
                            dr[i] = 0;
                        }
                        dataTable.Rows.Add(dr);
                        return dataTable;
                    }
                }
            }

            return null;
        }

        public async Task<string> GetSP_RCD0010_UDF002SAsync(DataTable Detail, string pID, string sYR, string sMN, string aDM_DIST, string aDV_DP, string aDV_UID, string Dept)
        {
            if (!string.IsNullOrEmpty(pID) && !string.IsNullOrEmpty(sYR) && !string.IsNullOrEmpty(sMN))
            {
                try
                {
                    using (var db3 = new CMEPMSEntities())
                    {
                        var RetVal = new SqlParameter("RetVal", SqlDbType.Int);
                        RetVal.Direction = ParameterDirection.ReturnValue;
                        var pidParam = new SqlParameter("@PID", SqlDbType.Decimal) { Precision = 18, Scale = 0, Value = Convert.ToDecimal(pID) };
                        var syrParam = new SqlParameter("@SYR", SqlDbType.VarChar) { Size = 10, Value = sYR };
                        var smnParam = new SqlParameter("@SMN", SqlDbType.VarChar) { Size = 10, Value = sMN };
                        var admDistParam = new SqlParameter("@ADM_DIST", SqlDbType.VarChar) { Size = 20, Value = aDM_DIST };
                        var advDpParam = new SqlParameter("@ADV_DP", SqlDbType.VarChar) { Size = 2000, Value = aDV_DP };
                        var advUidParam = new SqlParameter("@ADV_UID", SqlDbType.VarChar) { Size = 20, Value = aDV_UID };
                        var DeptParam = new SqlParameter("@Dept", SqlDbType.VarChar) { Size = 20, Value = Dept };
                        var connection = (SqlConnection)db3.Database.Connection;
                        var cmd = new SqlCommand("EXEC CMEPMS.dbo.SP_RCD0010_UDF002 @PID, @SYR, @SMN, @ADM_DIST, @ADV_DP, @ADV_UID,@Dept", connection);
                        cmd.Parameters.Add(pidParam);
                        cmd.Parameters.Add(RetVal);
                        cmd.Parameters.Add(syrParam);
                        cmd.Parameters.Add(smnParam);
                        cmd.Parameters.Add(admDistParam);
                        cmd.Parameters.Add(advDpParam);
                        cmd.Parameters.Add(advUidParam);
                        cmd.Parameters.Add(DeptParam);
                        cmd.CommandTimeout = 6000;

                        await connection.OpenAsync();
                        var dr = await cmd.ExecuteReaderAsync();
                        var oTable = new DataTable();
                        oTable.Load(dr); // 使用同步的 Load 方法

                        if (Convert.ToInt16(RetVal.Value) == 0)
                        {
                            Detail.Merge(oTable);
                            return "";
                        }

                        return RetVal.Value.ToString();
                    }
                }
                catch (Exception e)
                {
                    return e.Message;
                }
            }

            return "IsNullOrEmpty";
        }

        public DataSet GetSP_RCD0010_UDF003_New(string pID, string SEQ_NO, string sACC_ID, string ADV_UID, string aDV_DP, string Dept, string Choose)
        {
            aDV_DP = aDV_DP.Replace(",", "','");
            aDV_DP = "'" + aDV_DP + "'";

            Debug.WriteLine($"PID:{pID}");
            Debug.WriteLine($"SEQ_NO:{SEQ_NO}");
            Debug.WriteLine($"SACC_ID:{sACC_ID}");
            Debug.WriteLine($"ADV_UID:{ADV_UID}");
            Debug.WriteLine($"ADV_DP:{aDV_DP}");
            Debug.WriteLine($"Dept:{Dept}");
            Debug.WriteLine($"Choose:{Choose}");

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
                    var chooseParam = new SqlParameter("@Choose", SqlDbType.VarChar) { Size = 2000, Value = Choose }; // 加入分署篩選

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

            return null;
        }

        private MainDataViewModel GetMainDataViewModel(DataTable Main, string PID, string ADMs)
        {
            if (Main == null || Main.Rows.Count == 0)
            {
                return null;
            }

            string projectName = "";
            string printDate = string.Format("{0:yyyyMMdd}", DateTime.Now);

            Dictionary<string, Dictionary<string, double>> areaData = new Dictionary<string, Dictionary<string, double>>();

            DataRow[] rows = null;
            if (ADMs == null || ADMs == "")
            {
                rows = Main.Select("TD_GLRY='01' and ADM_DIST <> ''");
            }
            else
            {
                rows = Main.Select(string.Format("TD_GLRY='01' and ADM_DIST='{0}'", ADMs));
            }
            var results = Main.AsEnumerable()
                .GroupBy(row => row.Field<string>("TD_GLRY"))
                .Select(g => new
                {
                    TD_GLRY = g.Key,
                    TD_FUND = Convert.ToDouble(g.Sum(row => row.Field<decimal?>("TD_FUND") ?? 0)),
                    FYR_FUND = Convert.ToDouble(g.Sum(row => row.Field<decimal?>("YR_FUND") ?? 0)),
                    REV_FUND = Convert.ToDouble(g.Sum(row => row.Field<decimal?>("REV_FUND") ?? 0)),
                    COST_FEE = Convert.ToDouble(g.Sum(row => row.Field<decimal?>("COST_FEE") ?? 0)),
                    SUMPPAY_FEE = Convert.ToDouble(g.Sum(row => row.Field<int?>("SUMPPAY_FEE") ?? 0)),
                    FIN_SAVE = Convert.ToDouble(g.Sum(row => row.Field<decimal?>("FIN_SAVE") ?? 0)),
                    REV_COST_FEE = Convert.ToDouble(g.Sum(row => row.Field<decimal?>("REV_COST_FEE") ?? 0)),
                }).ToList();

            var a = results.Where(result => result.TD_GLRY == "01").FirstOrDefault();
            var b = results.Where(result => result.TD_GLRY == "02").FirstOrDefault();
            var mainListC01 = Main.AsEnumerable().Where(r => r.Field<string>("TD_GLRY") == "01" && r.Field<string>("ADM_DIST") == "中區").ToList();
            var mainListN01 = Main.AsEnumerable().Where(r => r.Field<string>("TD_GLRY") == "01" && r.Field<string>("ADM_DIST") == "北區").ToList();
            var mainListS01 = Main.AsEnumerable().Where(r => r.Field<string>("TD_GLRY") == "01" && r.Field<string>("ADM_DIST") == "南區").ToList();
            var mainListC02 = Main.AsEnumerable().Where(r => r.Field<string>("TD_GLRY") == "02" && r.Field<string>("ADM_DIST") == "中區").ToList();
            var mainListN02 = Main.AsEnumerable().Where(r => r.Field<string>("TD_GLRY") == "02" && r.Field<string>("ADM_DIST") == "北區").ToList();
            var mainListS02 = Main.AsEnumerable().Where(r => r.Field<string>("TD_GLRY") == "02" && r.Field<string>("ADM_DIST") == "南區").ToList();

            MainDataViewModel mainDataViewModel = new MainDataViewModel();

            mainDataViewModel.PRINT_DATE = printDate;
            mainDataViewModel.PJ_NAME = projectName;
            //01
            mainDataViewModel.TD_FUND_01 = a.TD_FUND;
            mainDataViewModel.FYR_FUND_01 = a.FYR_FUND;
            mainDataViewModel.REV_FUND_01 = a.REV_FUND;
            mainDataViewModel.COST_FEE_01 = a.COST_FEE;
            mainDataViewModel.SUMPPAY_FEE_01 = a.SUMPPAY_FEE;
            mainDataViewModel.FIN_SAVE_01 = a.FIN_SAVE;
            mainDataViewModel.REV_COST_FEE_01 = a.REV_COST_FEE;
            //02
            if (b is null)
            {
                mainDataViewModel.TD_FUND_02 = 0;
                mainDataViewModel.FYR_FUND_02 = 0;
                mainDataViewModel.REV_FUND_02 = 0;
                mainDataViewModel.COST_FEE_02 = 0;
                mainDataViewModel.SUMPPAY_FEE_02 = 0;
                mainDataViewModel.FIN_SAVE_02 = 0;
                mainDataViewModel.REV_COST_FEE_02 = 0;
            }
            else
            {
                mainDataViewModel.TD_FUND_02 = b.TD_FUND;
                mainDataViewModel.FYR_FUND_02 = b.FYR_FUND;
                mainDataViewModel.REV_FUND_02 = b.REV_FUND;
                mainDataViewModel.COST_FEE_02 = b.COST_FEE;
                mainDataViewModel.SUMPPAY_FEE_02 = b.SUMPPAY_FEE;
                mainDataViewModel.FIN_SAVE_02 = b.FIN_SAVE;
                mainDataViewModel.REV_COST_FEE_02 = b.REV_COST_FEE;
            }
            //中區01
            if (mainListC01.Count > 0)
            {
                mainDataViewModel.C01_TD_GLRY = mainListC01[0]["TD_GLRY"].ToString();
                mainDataViewModel.C01_ADM_DIST = mainListC01[0]["ADM_DIST"].ToString();
                mainDataViewModel.C01_FYR_FUND = ConvertDoubleValue(mainListC01[0]["FYR_FUND"].ToString());
                mainDataViewModel.C01_TD_FUND = ConvertDoubleValue(mainListC01[0]["TD_FUND"].ToString());
                mainDataViewModel.C01_YR_FUND = ConvertDoubleValue(mainListC01[0]["YR_FUND"].ToString());
                mainDataViewModel.C01_REV_FUND = ConvertDoubleValue(mainListC01[0]["REV_FUND"].ToString());
                mainDataViewModel.C01_MN_DIST_FUND = ConvertDoubleValue(mainListC01[0]["MN_DIST_FUND"].ToString());
                mainDataViewModel.C01_REV_MN_DIST_FUND = ConvertDoubleValue(mainListC01[0]["REV_MN_DIST_FUND"].ToString());
                mainDataViewModel.C01_COST_FEE = ConvertDoubleValue(mainListC01[0]["COST_FEE"].ToString());
                mainDataViewModel.C01_REV_COST_FEE = ConvertDoubleValue(mainListC01[0]["REV_COST_FEE"].ToString());
                mainDataViewModel.C01_PPAY_FEE = ConvertDoubleValue(mainListC01[0]["PPAY_FEE"].ToString());
                mainDataViewModel.C01_REV_PPAY_FEE = ConvertDoubleValue(mainListC01[0]["REV_PPAY_FEE"].ToString());
                mainDataViewModel.C01_FIN_SAVE = ConvertDoubleValue(mainListC01[0]["FIN_SAVE"].ToString());
                mainDataViewModel.C01_SUMYR_FUND = ConvertDoubleValue(mainListC01[0]["SUMYR_FUND"].ToString());
                mainDataViewModel.C01_SUMMN_DIST_FUND = ConvertDoubleValue(mainListC01[0]["SUMMN_DIST_FUND"].ToString());
                mainDataViewModel.C01_SUMCOST_FEE = ConvertDoubleValue(mainListC01[0]["SUMCOST_FEE"].ToString());
                mainDataViewModel.C01_SUMPPAY_FEE = ConvertDoubleValue(mainListC01[0]["SUMPPAY_FEE"].ToString());
                mainDataViewModel.C01_COST_CN_FEE = ConvertDoubleValue(mainListC01[0]["COST_CN_FEE"].ToString());
                mainDataViewModel.C01_COST_LD_FEE = ConvertDoubleValue(mainListC01[0]["COST_LD_FEE"].ToString());
                mainDataViewModel.C01_LD_CNT = ConvertDoubleValue(mainListC01[0]["LD_CNT"].ToString());
                mainDataViewModel.C01_LD_UG_CNT = ConvertDoubleValue(mainListC01[0]["LD_UG_CNT"].ToString());
                mainDataViewModel.C01_LD_HG_CNT = ConvertDoubleValue(mainListC01[0]["LD_HG_CNT"].ToString());
                mainDataViewModel.C01_DG_CNT = ConvertDoubleValue(mainListC01[0]["DG_CNT"].ToString());
                mainDataViewModel.C01_DG_UD_CNT = ConvertDoubleValue(mainListC01[0]["DG_UD_CNT"].ToString());
                mainDataViewModel.C01_DG_HD_CNT = ConvertDoubleValue(mainListC01[0]["DG_HD_CNT"].ToString());
                mainDataViewModel.C01_CN_CNT = ConvertDoubleValue(mainListC01[0]["CN_CNT"].ToString());
                mainDataViewModel.C01_CN_UD_CNT = ConvertDoubleValue(mainListC01[0]["CN_UD_CNT"].ToString());
                mainDataViewModel.C01_CN_UB_CNT = ConvertDoubleValue(mainListC01[0]["CN_UB_CNT"].ToString());
                mainDataViewModel.C01_CN_CN_CNT = ConvertDoubleValue(mainListC01[0]["CN_CN_CNT"].ToString());
                mainDataViewModel.C01_CN_HF_CNT = ConvertDoubleValue(mainListC01[0]["CN_HF_CNT"].ToString());
                mainDataViewModel.C01_M3_PF_CNT = ConvertDoubleValue(mainListC01[0]["M3_PF_CNT"].ToString());
                mainDataViewModel.C01_M4_PF_CNT = ConvertDoubleValue(mainListC01[0]["M4_PF_CNT"].ToString());
                mainDataViewModel.C01_M4_PU_CNT = ConvertDoubleValue(mainListC01[0]["M4_PU_CNT"].ToString());

            }
            //北區01
            if (mainListN01.Count > 0)
            {
                mainDataViewModel.N01_TD_GLRY = mainListN01[0]["TD_GLRY"].ToString();
                mainDataViewModel.N01_ADM_DIST = mainListN01[0]["ADM_DIST"].ToString();
                mainDataViewModel.N01_FYR_FUND = ConvertDoubleValue(mainListN01[0]["FYR_FUND"].ToString());
                mainDataViewModel.N01_TD_FUND = ConvertDoubleValue(mainListN01[0]["TD_FUND"].ToString());
                mainDataViewModel.N01_YR_FUND = ConvertDoubleValue(mainListN01[0]["YR_FUND"].ToString());
                mainDataViewModel.N01_REV_FUND = ConvertDoubleValue(mainListN01[0]["REV_FUND"].ToString());
                mainDataViewModel.N01_MN_DIST_FUND = ConvertDoubleValue(mainListN01[0]["MN_DIST_FUND"].ToString());
                mainDataViewModel.N01_REV_MN_DIST_FUND = ConvertDoubleValue(mainListN01[0]["REV_MN_DIST_FUND"].ToString());
                mainDataViewModel.N01_COST_FEE = ConvertDoubleValue(mainListN01[0]["COST_FEE"].ToString());
                mainDataViewModel.N01_REV_COST_FEE = ConvertDoubleValue(mainListN01[0]["REV_COST_FEE"].ToString());
                mainDataViewModel.N01_PPAY_FEE = ConvertDoubleValue(mainListN01[0]["PPAY_FEE"].ToString());
                mainDataViewModel.N01_REV_PPAY_FEE = ConvertDoubleValue(mainListN01[0]["REV_PPAY_FEE"].ToString());
                mainDataViewModel.N01_FIN_SAVE = ConvertDoubleValue(mainListN01[0]["FIN_SAVE"].ToString());
                mainDataViewModel.N01_SUMYR_FUND = ConvertDoubleValue(mainListN01[0]["SUMYR_FUND"].ToString());
                mainDataViewModel.N01_SUMMN_DIST_FUND = ConvertDoubleValue(mainListN01[0]["SUMMN_DIST_FUND"].ToString());
                mainDataViewModel.N01_SUMCOST_FEE = ConvertDoubleValue(mainListN01[0]["SUMCOST_FEE"].ToString());
                mainDataViewModel.N01_SUMPPAY_FEE = ConvertDoubleValue(mainListN01[0]["SUMPPAY_FEE"].ToString());
                mainDataViewModel.N01_COST_CN_FEE = ConvertDoubleValue(mainListN01[0]["COST_CN_FEE"].ToString());
                mainDataViewModel.N01_COST_LD_FEE = ConvertDoubleValue(mainListN01[0]["COST_LD_FEE"].ToString());
                mainDataViewModel.N01_LD_CNT = ConvertDoubleValue(mainListN01[0]["LD_CNT"].ToString());
                mainDataViewModel.N01_LD_UG_CNT = ConvertDoubleValue(mainListN01[0]["LD_UG_CNT"].ToString());
                mainDataViewModel.N01_LD_HG_CNT = ConvertDoubleValue(mainListN01[0]["LD_HG_CNT"].ToString());
                mainDataViewModel.N01_DG_CNT = ConvertDoubleValue(mainListN01[0]["DG_CNT"].ToString());
                mainDataViewModel.N01_DG_UD_CNT = ConvertDoubleValue(mainListN01[0]["DG_UD_CNT"].ToString());
                mainDataViewModel.N01_DG_HD_CNT = ConvertDoubleValue(mainListN01[0]["DG_HD_CNT"].ToString());
                mainDataViewModel.N01_CN_CNT = ConvertDoubleValue(mainListN01[0]["CN_CNT"].ToString());
                mainDataViewModel.N01_CN_UD_CNT = ConvertDoubleValue(mainListN01[0]["CN_UD_CNT"].ToString());
                mainDataViewModel.N01_CN_UB_CNT = ConvertDoubleValue(mainListN01[0]["CN_UB_CNT"].ToString());
                mainDataViewModel.N01_CN_CN_CNT = ConvertDoubleValue(mainListN01[0]["CN_CN_CNT"].ToString());
                mainDataViewModel.N01_CN_HF_CNT = ConvertDoubleValue(mainListN01[0]["CN_HF_CNT"].ToString());
                mainDataViewModel.N01_M3_PF_CNT = ConvertDoubleValue(mainListN01[0]["M3_PF_CNT"].ToString());
                mainDataViewModel.N01_M4_PF_CNT = ConvertDoubleValue(mainListN01[0]["M4_PF_CNT"].ToString());
                mainDataViewModel.N01_M4_PU_CNT = ConvertDoubleValue(mainListN01[0]["M4_PU_CNT"].ToString());

            }
            //南區01
            if (mainListS01.Count > 0)
            {
                mainDataViewModel.S01_TD_GLRY = mainListS01[0]["TD_GLRY"].ToString();
                mainDataViewModel.S01_ADM_DIST = mainListS01[0]["ADM_DIST"].ToString();
                mainDataViewModel.S01_FYR_FUND = ConvertDoubleValue(mainListS01[0]["FYR_FUND"].ToString());
                mainDataViewModel.S01_TD_FUND = ConvertDoubleValue(mainListS01[0]["TD_FUND"].ToString());
                mainDataViewModel.S01_YR_FUND = ConvertDoubleValue(mainListS01[0]["YR_FUND"].ToString());
                mainDataViewModel.S01_REV_FUND = ConvertDoubleValue(mainListS01[0]["REV_FUND"].ToString());
                mainDataViewModel.S01_MN_DIST_FUND = ConvertDoubleValue(mainListS01[0]["MN_DIST_FUND"].ToString());
                mainDataViewModel.S01_REV_MN_DIST_FUND = ConvertDoubleValue(mainListS01[0]["REV_MN_DIST_FUND"].ToString());
                mainDataViewModel.S01_COST_FEE = ConvertDoubleValue(mainListS01[0]["COST_FEE"].ToString());
                mainDataViewModel.S01_REV_COST_FEE = ConvertDoubleValue(mainListS01[0]["REV_COST_FEE"].ToString());
                mainDataViewModel.S01_PPAY_FEE = ConvertDoubleValue(mainListS01[0]["PPAY_FEE"].ToString());
                mainDataViewModel.S01_REV_PPAY_FEE = ConvertDoubleValue(mainListS01[0]["REV_PPAY_FEE"].ToString());
                mainDataViewModel.S01_FIN_SAVE = ConvertDoubleValue(mainListS01[0]["FIN_SAVE"].ToString());
                mainDataViewModel.S01_SUMYR_FUND = ConvertDoubleValue(mainListS01[0]["SUMYR_FUND"].ToString());
                mainDataViewModel.S01_SUMMN_DIST_FUND = ConvertDoubleValue(mainListS01[0]["SUMMN_DIST_FUND"].ToString());
                mainDataViewModel.S01_SUMCOST_FEE = ConvertDoubleValue(mainListS01[0]["SUMCOST_FEE"].ToString());
                mainDataViewModel.S01_SUMPPAY_FEE = ConvertDoubleValue(mainListS01[0]["SUMPPAY_FEE"].ToString());
                mainDataViewModel.S01_COST_CN_FEE = ConvertDoubleValue(mainListS01[0]["COST_CN_FEE"].ToString());
                mainDataViewModel.S01_COST_LD_FEE = ConvertDoubleValue(mainListS01[0]["COST_LD_FEE"].ToString());
                mainDataViewModel.S01_LD_CNT = ConvertDoubleValue(mainListS01[0]["LD_CNT"].ToString());
                mainDataViewModel.S01_LD_UG_CNT = ConvertDoubleValue(mainListS01[0]["LD_UG_CNT"].ToString());
                mainDataViewModel.S01_LD_HG_CNT = ConvertDoubleValue(mainListS01[0]["LD_HG_CNT"].ToString());
                mainDataViewModel.S01_DG_CNT = ConvertDoubleValue(mainListS01[0]["DG_CNT"].ToString());
                mainDataViewModel.S01_DG_UD_CNT = ConvertDoubleValue(mainListS01[0]["DG_UD_CNT"].ToString());
                mainDataViewModel.S01_DG_HD_CNT = ConvertDoubleValue(mainListS01[0]["DG_HD_CNT"].ToString());
                mainDataViewModel.S01_CN_CNT = ConvertDoubleValue(mainListS01[0]["CN_CNT"].ToString());
                mainDataViewModel.S01_CN_UD_CNT = ConvertDoubleValue(mainListS01[0]["CN_UD_CNT"].ToString());
                mainDataViewModel.S01_CN_UB_CNT = ConvertDoubleValue(mainListS01[0]["CN_UB_CNT"].ToString());
                mainDataViewModel.S01_CN_CN_CNT = ConvertDoubleValue(mainListS01[0]["CN_CN_CNT"].ToString());
                mainDataViewModel.S01_CN_HF_CNT = ConvertDoubleValue(mainListS01[0]["CN_HF_CNT"].ToString());
                mainDataViewModel.S01_M3_PF_CNT = ConvertDoubleValue(mainListS01[0]["M3_PF_CNT"].ToString());
                mainDataViewModel.S01_M4_PF_CNT = ConvertDoubleValue(mainListS01[0]["M4_PF_CNT"].ToString());
                mainDataViewModel.S01_M4_PU_CNT = ConvertDoubleValue(mainListS01[0]["M4_PU_CNT"].ToString());

            }
            //中區02
            if (mainListC02.Count > 0)
            {
                mainDataViewModel.C02_TD_GLRY = mainListC02[0]["TD_GLRY"].ToString();
                mainDataViewModel.C02_ADM_DIST = mainListC02[0]["ADM_DIST"].ToString();
                mainDataViewModel.C02_FYR_FUND = ConvertDoubleValue(mainListC02[0]["FYR_FUND"].ToString());
                mainDataViewModel.C02_TD_FUND = ConvertDoubleValue(mainListC02[0]["TD_FUND"].ToString());
                mainDataViewModel.C02_YR_FUND = ConvertDoubleValue(mainListC02[0]["YR_FUND"].ToString());
                mainDataViewModel.C02_REV_FUND = ConvertDoubleValue(mainListC02[0]["REV_FUND"].ToString());
                mainDataViewModel.C02_MN_DIST_FUND = ConvertDoubleValue(mainListC02[0]["MN_DIST_FUND"].ToString());
                mainDataViewModel.C02_REV_MN_DIST_FUND = ConvertDoubleValue(mainListC02[0]["REV_MN_DIST_FUND"].ToString());
                mainDataViewModel.C02_COST_FEE = ConvertDoubleValue(mainListC02[0]["COST_FEE"].ToString());
                mainDataViewModel.C02_REV_COST_FEE = ConvertDoubleValue(mainListC02[0]["REV_COST_FEE"].ToString());
                mainDataViewModel.C02_PPAY_FEE = ConvertDoubleValue(mainListC02[0]["PPAY_FEE"].ToString());
                mainDataViewModel.C02_REV_PPAY_FEE = ConvertDoubleValue(mainListC02[0]["REV_PPAY_FEE"].ToString());
                mainDataViewModel.C02_FIN_SAVE = ConvertDoubleValue(mainListC02[0]["FIN_SAVE"].ToString());
                mainDataViewModel.C02_SUMYR_FUND = ConvertDoubleValue(mainListC02[0]["SUMYR_FUND"].ToString());
                mainDataViewModel.C02_SUMMN_DIST_FUND = ConvertDoubleValue(mainListC02[0]["SUMMN_DIST_FUND"].ToString());
                mainDataViewModel.C02_SUMCOST_FEE = ConvertDoubleValue(mainListC02[0]["SUMCOST_FEE"].ToString());
                mainDataViewModel.C02_SUMPPAY_FEE = ConvertDoubleValue(mainListC02[0]["SUMPPAY_FEE"].ToString());
                mainDataViewModel.C02_COST_CN_FEE = ConvertDoubleValue(mainListC02[0]["COST_CN_FEE"].ToString());
                mainDataViewModel.C02_COST_LD_FEE = ConvertDoubleValue(mainListC02[0]["COST_LD_FEE"].ToString());
                mainDataViewModel.C02_LD_CNT = ConvertDoubleValue(mainListC02[0]["LD_CNT"].ToString());
                mainDataViewModel.C02_LD_UG_CNT = ConvertDoubleValue(mainListC02[0]["LD_UG_CNT"].ToString());
                mainDataViewModel.C02_LD_HG_CNT = ConvertDoubleValue(mainListC02[0]["LD_HG_CNT"].ToString());
                mainDataViewModel.C02_DG_CNT = ConvertDoubleValue(mainListC02[0]["DG_CNT"].ToString());
                mainDataViewModel.C02_DG_UD_CNT = ConvertDoubleValue(mainListC02[0]["DG_UD_CNT"].ToString());
                mainDataViewModel.C02_DG_HD_CNT = ConvertDoubleValue(mainListC02[0]["DG_HD_CNT"].ToString());
                mainDataViewModel.C02_CN_CNT = ConvertDoubleValue(mainListC02[0]["CN_CNT"].ToString());
                mainDataViewModel.C02_CN_UD_CNT = ConvertDoubleValue(mainListC02[0]["CN_UD_CNT"].ToString());
                mainDataViewModel.C02_CN_UB_CNT = ConvertDoubleValue(mainListC02[0]["CN_UB_CNT"].ToString());
                mainDataViewModel.C02_CN_CN_CNT = ConvertDoubleValue(mainListC02[0]["CN_CN_CNT"].ToString());
                mainDataViewModel.C02_CN_HF_CNT = ConvertDoubleValue(mainListC02[0]["CN_HF_CNT"].ToString());
                mainDataViewModel.C02_M3_PF_CNT = ConvertDoubleValue(mainListC02[0]["M3_PF_CNT"].ToString());
                mainDataViewModel.C02_M4_PF_CNT = ConvertDoubleValue(mainListC02[0]["M4_PF_CNT"].ToString());
                mainDataViewModel.C02_M4_PU_CNT = ConvertDoubleValue(mainListC02[0]["M4_PU_CNT"].ToString());

            }
            //北區02
            if (mainListN02.Count > 0)
            {
                mainDataViewModel.N02_TD_GLRY = mainListN02[0]["TD_GLRY"].ToString();
                mainDataViewModel.N02_ADM_DIST = mainListN02[0]["ADM_DIST"].ToString();
                mainDataViewModel.N02_FYR_FUND = ConvertDoubleValue(mainListN02[0]["FYR_FUND"].ToString());
                mainDataViewModel.N02_TD_FUND = ConvertDoubleValue(mainListN02[0]["TD_FUND"].ToString());
                mainDataViewModel.N02_YR_FUND = ConvertDoubleValue(mainListN02[0]["YR_FUND"].ToString());
                mainDataViewModel.N02_REV_FUND = ConvertDoubleValue(mainListN02[0]["REV_FUND"].ToString());
                mainDataViewModel.N02_MN_DIST_FUND = ConvertDoubleValue(mainListN02[0]["MN_DIST_FUND"].ToString());
                mainDataViewModel.N02_REV_MN_DIST_FUND = ConvertDoubleValue(mainListN02[0]["REV_MN_DIST_FUND"].ToString());
                mainDataViewModel.N02_COST_FEE = ConvertDoubleValue(mainListN02[0]["COST_FEE"].ToString());
                mainDataViewModel.N02_REV_COST_FEE = ConvertDoubleValue(mainListN02[0]["REV_COST_FEE"].ToString());
                mainDataViewModel.N02_PPAY_FEE = ConvertDoubleValue(mainListN02[0]["PPAY_FEE"].ToString());
                mainDataViewModel.N02_REV_PPAY_FEE = ConvertDoubleValue(mainListN02[0]["REV_PPAY_FEE"].ToString());
                mainDataViewModel.N02_FIN_SAVE = ConvertDoubleValue(mainListN02[0]["FIN_SAVE"].ToString());
                mainDataViewModel.N02_SUMYR_FUND = ConvertDoubleValue(mainListN02[0]["SUMYR_FUND"].ToString());
                mainDataViewModel.N02_SUMMN_DIST_FUND = ConvertDoubleValue(mainListN02[0]["SUMMN_DIST_FUND"].ToString());
                mainDataViewModel.N02_SUMCOST_FEE = ConvertDoubleValue(mainListN02[0]["SUMCOST_FEE"].ToString());
                mainDataViewModel.N02_SUMPPAY_FEE = ConvertDoubleValue(mainListN02[0]["SUMPPAY_FEE"].ToString());
                mainDataViewModel.N02_COST_CN_FEE = ConvertDoubleValue(mainListN02[0]["COST_CN_FEE"].ToString());
                mainDataViewModel.N02_COST_LD_FEE = ConvertDoubleValue(mainListN02[0]["COST_LD_FEE"].ToString());
                mainDataViewModel.N02_LD_CNT = ConvertDoubleValue(mainListN02[0]["LD_CNT"].ToString());
                mainDataViewModel.N02_LD_UG_CNT = ConvertDoubleValue(mainListN02[0]["LD_UG_CNT"].ToString());
                mainDataViewModel.N02_LD_HG_CNT = ConvertDoubleValue(mainListN02[0]["LD_HG_CNT"].ToString());
                mainDataViewModel.N02_DG_CNT = ConvertDoubleValue(mainListN02[0]["DG_CNT"].ToString());
                mainDataViewModel.N02_DG_UD_CNT = ConvertDoubleValue(mainListN02[0]["DG_UD_CNT"].ToString());
                mainDataViewModel.N02_DG_HD_CNT = ConvertDoubleValue(mainListN02[0]["DG_HD_CNT"].ToString());
                mainDataViewModel.N02_CN_CNT = ConvertDoubleValue(mainListN02[0]["CN_CNT"].ToString());
                mainDataViewModel.N02_CN_UD_CNT = ConvertDoubleValue(mainListN02[0]["CN_UD_CNT"].ToString());
                mainDataViewModel.N02_CN_UB_CNT = ConvertDoubleValue(mainListN02[0]["CN_UB_CNT"].ToString());
                mainDataViewModel.N02_CN_CN_CNT = ConvertDoubleValue(mainListN02[0]["CN_CN_CNT"].ToString());
                mainDataViewModel.N02_CN_HF_CNT = ConvertDoubleValue(mainListN02[0]["CN_HF_CNT"].ToString());
                mainDataViewModel.N02_M3_PF_CNT = ConvertDoubleValue(mainListN02[0]["M3_PF_CNT"].ToString());
                mainDataViewModel.N02_M4_PF_CNT = ConvertDoubleValue(mainListN02[0]["M4_PF_CNT"].ToString());
                mainDataViewModel.N02_M4_PU_CNT = ConvertDoubleValue(mainListN02[0]["M4_PU_CNT"].ToString());

            }
            //南區02
            if (mainListS02.Count > 0)
            {
                mainDataViewModel.S02_TD_GLRY = mainListS02[0]["TD_GLRY"].ToString();
                mainDataViewModel.S02_ADM_DIST = mainListS02[0]["ADM_DIST"].ToString();
                mainDataViewModel.S02_FYR_FUND = ConvertDoubleValue(mainListS02[0]["FYR_FUND"].ToString());
                mainDataViewModel.S02_TD_FUND = ConvertDoubleValue(mainListS02[0]["TD_FUND"].ToString());
                mainDataViewModel.S02_YR_FUND = ConvertDoubleValue(mainListS02[0]["YR_FUND"].ToString());
                mainDataViewModel.S02_REV_FUND = ConvertDoubleValue(mainListS02[0]["REV_FUND"].ToString());
                mainDataViewModel.S02_MN_DIST_FUND = ConvertDoubleValue(mainListS02[0]["MN_DIST_FUND"].ToString());
                mainDataViewModel.S02_REV_MN_DIST_FUND = ConvertDoubleValue(mainListS02[0]["REV_MN_DIST_FUND"].ToString());
                mainDataViewModel.S02_COST_FEE = ConvertDoubleValue(mainListS02[0]["COST_FEE"].ToString());
                mainDataViewModel.S02_REV_COST_FEE = ConvertDoubleValue(mainListS02[0]["REV_COST_FEE"].ToString());
                mainDataViewModel.S02_PPAY_FEE = ConvertDoubleValue(mainListS02[0]["PPAY_FEE"].ToString());
                mainDataViewModel.S02_REV_PPAY_FEE = ConvertDoubleValue(mainListS02[0]["REV_PPAY_FEE"].ToString());
                mainDataViewModel.S02_FIN_SAVE = ConvertDoubleValue(mainListS02[0]["FIN_SAVE"].ToString());
                mainDataViewModel.S02_SUMYR_FUND = ConvertDoubleValue(mainListS02[0]["SUMYR_FUND"].ToString());
                mainDataViewModel.S02_SUMMN_DIST_FUND = ConvertDoubleValue(mainListS02[0]["SUMMN_DIST_FUND"].ToString());
                mainDataViewModel.S02_SUMCOST_FEE = ConvertDoubleValue(mainListS02[0]["SUMCOST_FEE"].ToString());
                mainDataViewModel.S02_SUMPPAY_FEE = ConvertDoubleValue(mainListS02[0]["SUMPPAY_FEE"].ToString());
                mainDataViewModel.S02_COST_CN_FEE = ConvertDoubleValue(mainListS02[0]["COST_CN_FEE"].ToString());
                mainDataViewModel.S02_COST_LD_FEE = ConvertDoubleValue(mainListS02[0]["COST_LD_FEE"].ToString());
                mainDataViewModel.S02_LD_CNT = ConvertDoubleValue(mainListS02[0]["LD_CNT"].ToString());
                mainDataViewModel.S02_LD_UG_CNT = ConvertDoubleValue(mainListS02[0]["LD_UG_CNT"].ToString());
                mainDataViewModel.S02_LD_HG_CNT = ConvertDoubleValue(mainListS02[0]["LD_HG_CNT"].ToString());
                mainDataViewModel.S02_DG_CNT = ConvertDoubleValue(mainListS02[0]["DG_CNT"].ToString());
                mainDataViewModel.S02_DG_UD_CNT = ConvertDoubleValue(mainListS02[0]["DG_UD_CNT"].ToString());
                mainDataViewModel.S02_DG_HD_CNT = ConvertDoubleValue(mainListS02[0]["DG_HD_CNT"].ToString());
                mainDataViewModel.S02_CN_CNT = ConvertDoubleValue(mainListS02[0]["CN_CNT"].ToString());
                mainDataViewModel.S02_CN_UD_CNT = ConvertDoubleValue(mainListS02[0]["CN_UD_CNT"].ToString());
                mainDataViewModel.S02_CN_UB_CNT = ConvertDoubleValue(mainListS02[0]["CN_UB_CNT"].ToString());
                mainDataViewModel.S02_CN_CN_CNT = ConvertDoubleValue(mainListS02[0]["CN_CN_CNT"].ToString());
                mainDataViewModel.S02_CN_HF_CNT = ConvertDoubleValue(mainListS02[0]["CN_HF_CNT"].ToString());
                mainDataViewModel.S02_M3_PF_CNT = ConvertDoubleValue(mainListS02[0]["M3_PF_CNT"].ToString());
                mainDataViewModel.S02_M4_PF_CNT = ConvertDoubleValue(mainListS02[0]["M4_PF_CNT"].ToString());
                mainDataViewModel.S02_M4_PU_CNT = ConvertDoubleValue(mainListS02[0]["M4_PU_CNT"].ToString());

            }

            return mainDataViewModel;
        }

        public static double ConvertDoubleValue(string value)
        {
            if (value == "&nbsp;")
            {
                return 0.0;
            }

            if (value.Trim() != "")
            {
                return Convert.ToDouble(value);
            }

            return 0.0;
        }

        private MainDataViewModel ReplaceMainDataViewModel(MainDataViewModel gmdvm, string[] PJYear, string PID, string SYR, string SMN, string aDV_UID)
        {
            string[] aryPid = PJYear;
            string Pids = "";
            for (int i = 0; i < aryPid.Length; i++)
            {
                if (Pids == "")
                {
                    Pids = aryPid[i];
                }
                else
                {
                    Pids += "," + aryPid[i];
                }
            }


            string pid = PID;


            // 自辦 - 計畫總經費
            gmdvm.TD_FUND_01 = GetTdFund(Pids, "<>") / 1000;
            // 補助地方工程 - 計畫總經費
            gmdvm.TD_FUND_02 = GetTdFund(Pids, "=") / 1000;
            // 自辦 - 本年度預算數
            gmdvm.FYR_FUND_01 = GetFsyrFund(pid, "<>") / 1000;
            // 補助地方工程 - 本年度預算數
            gmdvm.FYR_FUND_02 = GetFsyrFund(pid, "=") / 1000;
            // 自辦 - 以前年度保留款
            gmdvm.REV_FUND_01 = GetRevFund(Pids, SYR, "<>") / 1000;
            // 補助地方工程 - 以前年度保留款
            gmdvm.REV_FUND_02 = GetRevFund(Pids, SYR, "=") / 1000;
            // 本年度 - 截至本月底累計支用數
            gmdvm.COST_FEE_01 = GetCostFee(pid, SYR, SMN, "<>") / 1000;
            // 本年度 - 截至本月底累計支用數
            gmdvm.COST_FEE_02 = GetCostFee(pid, SYR, SMN, "=") / 1000;
            // 以前年度保留款 - 截至本月底累計支用數
            gmdvm.REV_COST_FEE_01 = GetRevCostFee(Pids, SYR, SMN, "<>") / 1000;
            // 以前年度保留款 - 截至本月底累計支用數
            gmdvm.REV_COST_FEE_02 = GetRevCostFee(Pids, SYR, SMN, "=") / 1000;

            // 自辦工程 - 本年度可支用預算 - 北中南區
            var dt1 = GetSumYrFund(pid, Pids, SYR, "<>", aDV_UID);
            if (dt1 != null)
            {
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    if (dt1.Rows[i]["AREA"].ToString() == "北區") { gmdvm.N01_SUMYR_FUND = Convert.ToDouble(dt1.Rows[i]["FUND"].ToString()) / 1000; }
                    else if (dt1.Rows[i]["AREA"].ToString() == "中區") { gmdvm.C01_SUMYR_FUND = Convert.ToDouble(dt1.Rows[i]["FUND"].ToString()) / 1000; }
                    else if (dt1.Rows[i]["AREA"].ToString() == "南區") { gmdvm.S01_SUMYR_FUND = Convert.ToDouble(dt1.Rows[i]["FUND"].ToString()) / 1000; }
                }
            }
            // 補助地方工程 - 本年度可支用預算 - 北中南區
            dt1 = GetSumYrFund(pid, Pids, SYR, "=", aDV_UID);
            if (dt1 != null)
            {
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    if (dt1.Rows[i]["AREA"].ToString() == "北區") { gmdvm.N02_SUMYR_FUND = Convert.ToDouble(dt1.Rows[i]["FUND"].ToString()) / 1000; }
                    else if (dt1.Rows[i]["AREA"].ToString() == "中區") { gmdvm.C02_SUMYR_FUND = Convert.ToDouble(dt1.Rows[i]["FUND"].ToString()) / 1000; }
                    else if (dt1.Rows[i]["AREA"].ToString() == "南區") { gmdvm.S02_SUMYR_FUND = Convert.ToDouble(dt1.Rows[i]["FUND"].ToString()) / 1000; }
                }
            }



            return gmdvm;

        }

        public double GetTdFund(string pids, string flag)
        {
            double result = 0;
            string connectionString = connectionString2;
            string sqlQuery = string.Format("select sum(FUND) as FUND from CMEPMS.dbo.CTDITFD where TID in (select a.TID from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDMFORM b where a.TD_CODE=b.TD_CODE and a.PID in ({0}) and b.CATY_CODE{1}'L')", pids, flag);

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sqlQuery, conn))
                {
                    cmd.Parameters.AddWithValue("@pids", pids);
                    cmd.Parameters.AddWithValue("@flag", flag);

                    try
                    {
                        conn.Open();
                        SqlDataReader reader = cmd.ExecuteReader();
                        if (reader.Read() && !reader.IsDBNull(reader.GetOrdinal("FUND")))
                        {
                            result = Convert.ToDouble(reader["FUND"]);
                        }
                    }
                    catch (Exception)
                    {

                    }
                }
            }

            return result;
        }

        public double GetFsyrFund(string pid, string flag)
        {
            string FSQLStr = "";
            double result = 0;
            FSQLStr = string.Format("select sum(FUND) as FUND from CMEPMS.dbo.CTDITFD where TID in (select a.TID from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDMFORM b where a.TD_CODE=b.TD_CODE and a.PID={0} and b.CATY_CODE{1}'L')", pid, flag);
            using (SqlConnection conn = new SqlConnection(connectionString2))
            {
                SqlCommand cmd = new SqlCommand(FSQLStr, conn);
                conn.Open();
                var sum = cmd.ExecuteScalar();
                if (sum != DBNull.Value)
                {
                    result = Convert.ToDouble(sum);
                }
            }
            return result;
        }

        public static void GetDtTitleName(string SEQ_NO, B1DtFooter b1DtFooter)
        {
            if (SEQ_NO == "0")
            {
                b1DtFooter.TITLE2 = "原預算 工作計畫項目";
                b1DtFooter.TITLE1 = "";//預設前一期標題
            }
            else
            {
                b1DtFooter.TITLE2 = "第" + SEQ_NO + "次修正工作計畫項目";
                string sSEQ_NO = Convert.ToString(Convert.ToInt32(SEQ_NO) - 1);
                if (sSEQ_NO == "0")
                {
                    b1DtFooter.TITLE1 = "原預算 工作計畫項目";
                }
                else
                {
                    b1DtFooter.TITLE1 = "第" + sSEQ_NO + "次修正工作計畫項目";
                }
            }
        }

        public double GetRevFund(string pids, string syr, string flag)
        {
            double revFund = 0;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = string.Format("select sum(a.REV_FUND) as REV_FUND from CMEPMS.dbo.CRDMNDS a,CMETMS.dbo.TDMFORM b where a.PID in({0}) and a.SYR={1} and a.TD_CODE1=b.TD_CODE and b.CATY_CODE {2} 'L'", pids, syr, flag);
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@pids", pids);
                    command.Parameters.Add("@syr", SqlDbType.VarChar, 3).Value = syr;
                    command.Parameters.AddWithValue("@flag", flag);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            if (!reader.IsDBNull(reader.GetOrdinal("REV_FUND")))
                            {
                                revFund = ObjectToDouble(reader.GetOrdinal("REV_FUND"));
                            }
                            else
                            {
                                revFund = 0;
                            }
                        }
                    }
                }
            }

            double fund = 0;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = string.Format("select sum(FUND) as FUND from CMEPMS.dbo.CTDITFD where TID in (select a.TID from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDMFORM b where a.PID in ({1}) and a.TD_CODE=b.TD_CODE and b.CATY_CODE{2}'L') and FIT_CODE='00' and SYR<{1}", pids, syr, flag);
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@pids", pids);
                    command.Parameters.Add("@syr", SqlDbType.VarChar, 3).Value = syr;
                    command.Parameters.AddWithValue("@flag", flag);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        var a = reader.GetOrdinal("FUND");
                        if (reader.Read())
                        {
                            if (!reader.IsDBNull(reader.GetOrdinal("FUND")))
                            {
                                fund = ObjectToDouble(reader.GetOrdinal("FUND"));
                            }
                            else

                            {
                                fund = 0;
                            }

                        }
                    }
                }
            }

            double fee = 0;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = string.Format("select sum(COST_FEE+PPAY_FEE) as FEE from CMEPMS.dbo.TDDSTCF where FD_CODE in (select a.FD_CODE from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDMFORM b where a.PID in ({0}) and a.TD_CODE=b.TD_CODE and b.CATY_CODE{1}'L') and SYR={2} and SMN='12'", pids, flag, syr);
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@pids", pids);
                    command.Parameters.Add("@syr", SqlDbType.VarChar, 3).Value = syr;
                    command.Parameters.AddWithValue("@flag", flag);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {

                            if (!reader.IsDBNull(reader.GetOrdinal("FEE")))
                            {
                                fee = Convert.ToDouble(reader["FEE"]);
                            }
                            else
                            {
                                fee = 0;
                            }
                        }
                    }
                }
            }
            double result = revFund + fund - fee;
            return result;
        }

        public double GetCostFee(string pid, string syr, string smn, string flag)
        {
            double result = 0;
            using (SqlConnection connection = new SqlConnection(connectionString2))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("select sum(COST_FEE + PPAY_FEE) as FUND from CMEPMS.dbo.TDDSTCF where FD_CODE in (select a.FD_CODE from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDMFORM b where a.TD_CODE=b.TD_CODE and a.PID=@PID and b.CATY_CODE" + flag + "'L') and SYR=@SYR and SMN=@SMN", connection))
                {
                    command.Parameters.AddWithValue("@PID", pid);
                    command.Parameters.AddWithValue("@SYR", syr);
                    command.Parameters.AddWithValue("@SMN", smn);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read() && reader["FUND"] != DBNull.Value)
                        {
                            result = Convert.ToDouble(reader["FUND"]);
                        }
                    }
                }
            }
            return result;
        }

        public double CalDiv(double x, double y)
        {
            if (y != 0)
            {
                return x / y;
            }
            else
            {
                return 0;
            }
        }

        public static double CalMinus(string x, string y)
        {
            if (string.IsNullOrEmpty(x))
            {
                return string.IsNullOrEmpty(y) ? 0 : 0;
            }
            else
            {
                if (string.IsNullOrEmpty(y))
                {
                    return Convert.ToDouble(x);
                }
                else
                {
                    double k = Convert.ToDouble(x);
                    double h = Convert.ToDouble(y);
                    return k - h;
                }
            }
        }

        public static double CalPlus(string x, string y)
        {
            if (string.IsNullOrEmpty(x))
            {
                return string.IsNullOrEmpty(y) ? 0 : Convert.ToDouble(y);
            }
            else
            {
                double h = Convert.ToDouble(x);
                double i = string.IsNullOrEmpty(y) ? 0 : Convert.ToDouble(y);
                return h + i;
            }
        }

        public double GetRevCostFee(string pids, string syr, string smn, string flag)
        {
            double rFee = 0;

            using (SqlConnection connection = new SqlConnection(connectionString2))
            {
                connection.Open();
                string query = string.Format("SELECT SUM(COST_FEE+PPAY_FEE) as FEE FROM CMEPMS.dbo.TDDSTRF WHERE FD_CODE IN (SELECT a.FD_CODE FROM CMEPMS.dbo.CRDMNDS a, CMETMS.dbo.TDMFORM b WHERE a.PID IN ({0}) AND a.SYR='{1}' AND a.TD_CODE1=b.TD_CODE AND b.CATY_CODE{2}'L') AND SYR='{1}' AND SMN='{3}'", pids, syr, flag, smn);
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(reader.GetOrdinal("FEE")))
                        {
                            rFee = reader.GetDouble(reader.GetOrdinal("FEE"));

                        }
                    }
                }
                reader.Close();
            }

            double fee = 0;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = string.Format("SELECT SUM(COST_FEE+PPAY_FEE) as Fee FROM CMEPMS.dbo.TDDSTCF WHERE FD_CODE IN (SELECT a.FD_CODE FROM CMEPMS.dbo.CTDMNDS a, CMETMS.dbo.TDMFORM b WHERE a.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PID IN ({0}) AND SYR<'{1}') AND a.TD_CODE=b.TD_CODE AND b.CATY_CODE{2}'L') AND SYR='{1}' AND SMN='{3}'", pids, syr, flag, smn);
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(reader.GetOrdinal("Fee")))
                        {
                            fee = reader.GetDouble(reader.GetOrdinal("Fee"));
                        }
                    }
                }
                reader.Close();
            }

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = string.Format("SELECT SUM(COST_FEE+PPAY_FEE) as Fee FROM CMEPMS.dbo.TDDSTCF WHERE FD_CODE IN (SELECT a.FD_CODE FROM CMEPMS.dbo.CTDMNDS a, CMETMS.dbo.TDMFORM b WHERE a.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PID IN ({0}) AND SYR<'{1}') AND a.TD_CODE=b.TD_CODE AND b.CATY_CODE{2}'L') AND SYR='{3}' AND SMN='12'", pids, syr, flag, Convert.ToString(Convert.ToInt16(syr) - 1).PadLeft(3, '0'));
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(reader.GetOrdinal("FEE")))
                        {
                            rFee = reader.GetDouble(reader.GetOrdinal("FEE"));
                        }
                    }
                }
                reader.Close();
            }

            double result = rFee + fee;
            return result;
        }

        public DataTable GetSumYrFund(string pid, string pids, string syr, string flag, string aDV_UID)
        {
            DataTable dt = null;
            using (SqlConnection conn = new SqlConnection(connectionString2))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = string.Format("select sum(c.FUND) as FUND,d.AREA from CMEPMS.dbo.CTDITFD c,(select b.TID, case TD_CITY when '不分區' then '北區' else (select AREA from CMBASE.dbo.BAS_SMCPOST_CITY_TOWN where CITY+isnull(TOWN,'')=a.TD_CITY+isnull(a.TD_TOWN,'')) end as AREA from CMETMS.dbo.TDMFORM a,CMEPMS.dbo.CTDMNDS b where b.PID in ({0}) and a.CATY_CODE{1}'L' and a.TD_CODE=b.TD_CODE AND EXISTS (SELECT 1 FROM CMBASE.dbo.SECCITY WHERE USR_ID ='{2}'  AND (CITY = a.TD_CITY AND TOWN = a.TD_TOWN))) d where c.TID=d.TID group by d.AREA", pids, flag, aDV_UID);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                dt = new DataTable();
                adapter.Fill(dt);
                cmd.CommandText = string.Format("select sum(c.COST_FEE+c.PPAY_FEE) as FEE,d.AREA from CMEPMS.dbo.TDDSTCF c,(select b.FD_CODE, case TD_CITY when '不分區' then '北區' else (select AREA from CMBASE.dbo.BAS_SMCPOST_CITY_TOWN where CITY+isnull(TOWN,'')=a.TD_CITY+isnull(a.TD_TOWN,'')) end as AREA from CMETMS.dbo.TDMFORM a,CMEPMS.dbo.CTDMNDS b where b.PID in ({0}) and a.CATY_CODE{1}'L' and a.TD_CODE=b.TD_CODE AND EXISTS (SELECT 1 FROM CMBASE.dbo.SECCITY WHERE USR_ID ='{3}'  AND (CITY = a.TD_CITY AND TOWN = a.TD_TOWN))) d where c.FD_CODE=d.FD_CODE and c.SYR='{2}' and c.SMN='12' group by d.AREA", pids, flag, Convert.ToString(Convert.ToInt16(syr) - 1).PadLeft(3, '0'), aDV_UID);
                adapter.SelectCommand = cmd;
                DataTable dtTemp = new DataTable();
                adapter.Fill(dtTemp);

                for (int i = 0; i < dtTemp.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        if (dtTemp.Rows[i]["AREA"].ToString() == dt.Rows[j]["AREA"].ToString())
                        {
                            dt.Rows[j]["FUND"] = Convert.ToDouble(dt.Rows[j]["FUND"].ToString()) - Convert.ToDouble(dtTemp.Rows[i]["FEE"].ToString());
                        }
                    }
                }
            }
            return dt;
        }

        private double chkdouble(string val)
        {
            double ddt = 0;
            //bool tempbool = double.TryParse(val, System.Globalization.NumberStyles.Integer, null, out ddt);
            bool tempbool = double.TryParse(val, out ddt);

            if (tempbool)
            {
                return ddt;
            }
            else
            {
                return 0;
            }

        }

        private string formatDate(string sDate, string split)
        {
            string rtn = "";
            if (sDate != null && sDate.Length == 7)
            {
                rtn = sDate.Substring(0, 3) + split + sDate.Substring(3, 2) + split + sDate.Substring(5, 2);
            }

            return rtn;
        }

        public string GetTDDMSCD_SCHD(string SYR, string SMN, string TD_CODE, string Sel_NAME)
        {
            var result = (from t in db3.TDDMSCD
                          where t.SYR == SYR && string.Compare(t.SMN, SMN) <= 0 && t.TD_CODE == TD_CODE
                          select t.PRE_SCHD).Max();
            string resultAsString = result.ToString();
            return resultAsString;
        }

        private void SheetSub(Worksheet Ws, MainDataViewModel gmdvm, string PJNAME)
        {
            int row = 0;
            string title = "內政部國土管理署" + Ws.Name + "一億元以上工程計畫執行情形";
            string area = "";
            switch (Ws.Name)
            {
                case "北區分處":
                    area = "N";
                    break;
                case "中區分處":
                    area = "C";
                    break;
                case "南區分處":
                    area = "S";
                    break;
                default:
                    break;
            }
            Ws.Cells[0, 0].Value = title;
            Ws.Cells[1, 0].Value = DateTime.Now.ToString("yyyyMMdd");
            row = 4;
            Ws.Cells[row, 0].Value = PJNAME;
            Ws.Cells[row, 6].Value = gmdvm.PJ_DIST_FUND;
            row = 5;
            for (int k = 1; k < 3; k++)
            {
                double TD_FUND = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_TD_FUND").GetValue(gmdvm, null);
                double FYR_FUND = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_FYR_FUND").GetValue(gmdvm, null);
                double REV_FUND = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_REV_FUND").GetValue(gmdvm, null);
                double SUMYR_FUND = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_SUMYR_FUND").GetValue(gmdvm, null);
                double COST_FEE = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_COST_FEE").GetValue(gmdvm, null);
                double REV_COST_FEE = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_REV_COST_FEE").GetValue(gmdvm, null);
                double SUMPPAY_FEE = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_SUMPPAY_FEE").GetValue(gmdvm, null);
                double FIN_SAVE = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_FIN_SAVE").GetValue(gmdvm, null);
                Ws.Cells[1, 0].Value = DateTime.Now.ToString("yyyyMMdd");

                //中區分處自辦工程統計

                Ws.Cells[row, 1].Value = TD_FUND;
                Ws.Cells[row, 2].Value = FYR_FUND;
                Ws.Cells[row, 3].Value = 0;
                Ws.Cells[row, 4].Value = REV_FUND;
                Ws.Cells[row, 5].Value = FYR_FUND + 0 + REV_FUND;
                Ws.Cells[row, 6].Value = SUMYR_FUND;
                Ws.Cells[row, 7].Value = COST_FEE;
                Ws.Cells[row, 8].Value = 0;
                Ws.Cells[row, 9].Value = REV_COST_FEE;
                Ws.Cells[row, 10].Value = COST_FEE + 0 + REV_COST_FEE;
                Ws.Cells[row, 11].Value = SUMPPAY_FEE;
                Ws.Cells[row, 12].Value = FIN_SAVE;
                Ws.Cells[row, 13].Value = CalDiv(COST_FEE + 0 + REV_COST_FEE, SUMYR_FUND);
                Ws.Cells[row, 14].Value = CalDiv(COST_FEE + 0 + REV_COST_FEE + SUMPPAY_FEE + FIN_SAVE, SUMYR_FUND);
                Ws.Cells[row, 15].Value = CalDiv(COST_FEE + 0 + REV_COST_FEE + SUMPPAY_FEE + FIN_SAVE, FYR_FUND);
                row++;
                //中區分處補助地方工程統計
            }
            //加總
            row = row - 3;
            for (int n = 1; n < 13; n++)
            {
                Ws.Cells[row, n].Value = ObjectToDouble(Ws.Cells[row + 1, n].Value) + ObjectToDouble(Ws.Cells[row + 2, n].Value);
            }
            Ws.Cells[row, 14].Value = CalDiv(ObjectToDouble(Ws.Cells[row, 11].Value), ObjectToDouble(Ws.Cells[row, 7].Value));
            Ws.Cells[row, 15].Value = CalDiv(ObjectToDouble(Ws.Cells[row, 11].Value) + ObjectToDouble(Ws.Cells[row, 12].Value) + ObjectToDouble(Ws.Cells[row, 13].Value), ObjectToDouble(Ws.Cells[row, 7].Value));
            Ws.Cells[row, 16].Value = CalDiv(ObjectToDouble(Ws.Cells[row, 11].Value) + ObjectToDouble(Ws.Cells[row, 12].Value) + ObjectToDouble(Ws.Cells[row, 13].Value), ObjectToDouble(Ws.Cells[row, 3].Value));

            //中區分處自辦工程中區

            switch (area)
            {
                case "N":
                    row = 13;
                    break;
                case "C":
                    row = 14;
                    break;
                case "S":
                    row = 15;
                    break;
            }
            for (int k = 1; k < 3; k++)
            {
                double SUMYR_FUND = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_SUMYR_FUND").GetValue(gmdvm, null);
                double SUMMN_DIST_FUND = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_SUMMN_DIST_FUND").GetValue(gmdvm, null);
                double SUMCOST_FEE = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_SUMCOST_FEE").GetValue(gmdvm, null);
                double SUMPPAY_FEE = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_SUMPPAY_FEE").GetValue(gmdvm, null);
                double LD_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_LD_CNT").GetValue(gmdvm, null);
                double LD_UG_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_LD_UG_CNT").GetValue(gmdvm, null);
                double LD_HG_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_LD_HG_CNT").GetValue(gmdvm, null);
                double DG_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_DG_CNT").GetValue(gmdvm, null);
                double DG_UD_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_DG_UD_CNT").GetValue(gmdvm, null);
                double DG_HD_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_DG_HD_CNT").GetValue(gmdvm, null);
                double CN_UD_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_CN_UD_CNT").GetValue(gmdvm, null);
                double CN_UB_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_CN_UB_CNT").GetValue(gmdvm, null);
                double CN_CN_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_CN_CN_CNT").GetValue(gmdvm, null);
                double CN_HF_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_CN_HF_CNT").GetValue(gmdvm, null);
                double M3_PF_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_M3_PF_CNT").GetValue(gmdvm, null);
                double M4_PF_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_M4_PF_CNT").GetValue(gmdvm, null);
                double M4_PU_CNT = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_M4_PU_CNT").GetValue(gmdvm, null);
                double COST_CN_FEE = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_COST_CN_FEE").GetValue(gmdvm, null);
                double COST_LD_FEE = (double)gmdvm.GetType().GetProperty(area + k.ToString("D2") + "_COST_LD_FEE").GetValue(gmdvm, null);


                Ws.Cells[row, 1].Value = SUMYR_FUND;
                Ws.Cells[row, 2].Value = SUMMN_DIST_FUND;
                Ws.Cells[row, 3].Value = SUMCOST_FEE;
                Ws.Cells[row, 4].Value = SUMPPAY_FEE;
                Ws.Cells[row, 5].Value = CalDiv(SUMCOST_FEE + SUMPPAY_FEE, SUMMN_DIST_FUND);
                Ws.Cells[row, 6].Value = LD_CNT + DG_CNT + CN_UD_CNT + CN_UB_CNT + CN_CN_CNT + CN_HF_CNT;
                Ws.Cells[row, 7].Value = LD_CNT;
                Ws.Cells[row, 8].Value = LD_UG_CNT;
                Ws.Cells[row, 9].Value = LD_HG_CNT;
                Ws.Cells[row, 10].Value = DG_CNT;
                Ws.Cells[row, 11].Value = DG_UD_CNT;
                Ws.Cells[row, 12].Value = DG_HD_CNT;
                Ws.Cells[row, 13].Value = CN_UD_CNT + CN_UB_CNT + CN_CN_CNT + CN_HF_CNT;
                Ws.Cells[row, 14].Value = CN_UD_CNT;
                Ws.Cells[row, 15].Value = CN_UB_CNT;
                Ws.Cells[row, 16].Value = CN_CN_CNT;
                Ws.Cells[row, 17].Value = CN_HF_CNT;
                if (k == 1)
                {
                    Ws.Cells[row, 18].Value = M3_PF_CNT;
                    Ws.Cells[row, 19].Value = M4_PF_CNT;
                    Ws.Cells[row, 20].Value = M4_PU_CNT;
                }
                else
                {
                    Ws.Cells[row, 18].Value = COST_CN_FEE;
                    Ws.Cells[row, 19].Value = COST_LD_FEE;
                }
                int row1 = 0;
                if (k == 1)
                {
                    row1 = 16;
                }
                else
                {
                    row1 = 22;
                }
                for (int l = 1; l < 20; l++)
                {
                    Ws.Cells[row1, l].Value = ObjectToDouble(Ws.Cells[row, l].Value);
                }
                row = row + 6;
            }
            row = 23;
            for (int m = 1; m < 18; m++)
            {
                Ws.Cells[row, m].Value = ObjectToDouble(Ws.Cells[17, m].Value) + ObjectToDouble(Ws.Cells[23, m].Value);
            }
            Ws.Cells[row, 5].Value = CalDiv(ObjectToDouble(Ws.Cells[row, 3].Value) + ObjectToDouble(Ws.Cells[row, 4].Value), ObjectToDouble(Ws.Cells[row, 2].Value));

        }

        public ActionResult Get_PID(string Type)
        {
            var PID = db2.PJMFORM.Where(m => m.Water_Type == Type && m.AN_PID != null).FirstOrDefault()?.PID ?? 0;
            string PID_STRING = PID.ToString();
            return Json(PID_STRING);

        }

    }
}