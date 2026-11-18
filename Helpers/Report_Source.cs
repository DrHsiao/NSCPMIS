using Aspose.Cells;
using PagedList.Mvc;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using WCals.Controllers;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEPMSModels;

namespace WCals.Source
{
    public class Report_Source
    {
        static public CMBASEEntities db = new CMBASEEntities();
        static public CMEPMSEntities db2 = new CMEPMSEntities();

        public static void getProject09(ref DataTable TDMFORM, ref DataTable TDDSPRS, string PID, string DATE_START, string DATE_END, string connectionString)
        {
            //在建工程==>已開工，未完工
            DataTable dt = new DataTable();
            string sql = "";
            DataRow[] rr;
            int i;

            sql = $@"
                    SELECT 
                        TD_CODE,
                        SP_DT,
                        CATEGORY,
                        CATEGORY_N,
                        SP_DEPT,
                        SP_ALL_SCORE,
                        SP_SCORE,
                        HIATUS,
                        HIATUS2,
                        CASE WHEN ISNULL(HIATUS, 0) = 1 THEN '是' ELSE '否' END AS HIATUS_DESC,
                        CASE WHEN ISNULL(HIATUS2, 0) = 1 THEN '是' ELSE '否' END AS HIATUS2_DESC,
                        IMP_DT,
                        FIN_DT,
                        EXA_UNIT
                    FROM 
                        CMETMS.dbo.TDV_TDDSPRS
                    WHERE 
                        TD_CODE IN (
                            SELECT TD_CODE 
                            FROM CMEPMS.dbo.CTDMNDS 
                            WHERE PID IN ( {PID} )
                        )";

            if (DATE_START != "")
            {
                if (DATE_END != "")
                {
                    sql = sql + " and CMETMS.dbo.getCHTDate(SP_DT,'') between '" + DATE_START + "' and '" + DATE_END + "'";
                }
                else
                {
                    sql = sql + " and CMETMS.dbo.getCHTDate(SP_DT,'') >= '" + DATE_START + "'";
                }
            }
            else
            {
                if (DATE_END != "")
                {
                    sql = sql + " and CMETMS.dbo.getCHTDate(SP_DT,'') <= '" + DATE_END + "'";
                }
            }

            sql = sql + " order by TD_CODE,SP_DT";
            TDDSPRS = SourceFunction.DataTableGet(sql, connectionString);
            sql = "select T.TD_CODE,T.TD_NAME,T.IPJ_CODE, S1.CODE_DESC as IPJ_DESC,B1.REL_EDT as TK106_EDT,B2.REL_EDT as TK107_EDT";
            sql = sql + ",'0' as SP_03,'0' as SP_02,'0' as SP_01";//查核，抽查，督導
            sql = sql + " from CMETMS.dbo.TDMFORM T";
            sql = sql + " left join CMETMS.dbo.SMCCODE S1 on S1.SYS_ID = 'SM' and S1.CATY_ID = '02' and S1.CATY_COD = T.IPJ_CODE";
            sql = sql + " left join CMETMS.dbo.TDDSDWB B1 on B1.TD_CODE = T.TD_CODE and B1.IT_CODE = 'TK106'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB B2 on B2.TD_CODE = T.TD_CODE and B2.IT_CODE = 'TK107'";
            sql = sql + " where isnull(B1.REL_EDT,'')<>'' and isnull(B2.REL_EDT,'')=''";
            sql = sql + $@" and T.TD_CODE in (select TD_CODE from CMEPMS.dbo.CTDMNDS where PID IN ({PID}))";
            sql = sql + " order by T.TD_CODE";
            TDMFORM = SourceFunction.DataTableGet(sql, connectionString);
            for (i = 0; i < TDMFORM.Rows.Count; i++)
            {
                rr = TDDSPRS.Select("TD_CODE='" + TDMFORM.Rows[i]["TD_CODE"].ToString() + "' and CATEGORY='01'");
                TDMFORM.Rows[i]["SP_01"] = rr.Length.ToString();
                rr = TDDSPRS.Select("TD_CODE='" + TDMFORM.Rows[i]["TD_CODE"].ToString() + "' and CATEGORY='02'");
                TDMFORM.Rows[i]["SP_02"] = rr.Length.ToString();
                rr = TDDSPRS.Select("TD_CODE='" + TDMFORM.Rows[i]["TD_CODE"].ToString() + "' and CATEGORY='03'");
                TDMFORM.Rows[i]["SP_03"] = rr.Length.ToString();
            }
        }

        public static string NumberToColumnLetter(int num)
        {
            string columnLetter = "";
            while (num > 0)
            {
                int remainder = (num - 1) % 26;
                columnLetter = (char)(65 + remainder) + columnLetter;
                num = (num - 1) / 26;
            }
            return columnLetter;
        }

        public static void CreateSheet(Workbook workbook, Aspose.Cells.Worksheet sheet1, DataTable dt, DataTable ditdt, string SYR, string PSEQ_NO, string SEQ_NO, string PJ_NAME)
        {
            Worksheet sheet1_1 = workbook.Worksheets["Sheet1_1"];

            int i;
            int j;
            int start_col;
            int line;
            int cols;
            int sid;

            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");

            double sum_FUND = 0;
            double diff_FUND = 0;

            if (SYR != "")
            {
                sSYR = SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }

            sheet1.Cells.CreateRange("A1", "E5").Copy(sheet1_1.Cells.CreateRange(0, 0, 5, 5));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_1.Cells.GetRowHeight(row));
            }
            cols = 5;
            start_col = 5;

            /// 前期以前年度
            sheet1.Cells.CreateRange("A1", "E5").Copy(sheet1_1.Cells.CreateRange("A1", "E5"));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_1.Cells.GetRowHeight(row));
            }

            /// 前期以前年度
            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_1.Cells.CreateRange(0, 5, 5, 1));
                sheet1.Cells.SetColumnWidth(cols, sheet1_1.Cells.GetColumnWidth(5));
                for (int row = 0; row < 5; row++)
                {
                    sheet1.Cells.SetRowHeight(row, sheet1_1.Cells.GetRowHeight(row + 5));
                }

                cols = cols + 1;
            }

            /// 前期本年度
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_1.Cells.CreateRange(0, 6, 5, 1));
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_1.Cells.CreateRange(0, 6, 5, 1));
            sheet1.Cells.SetColumnWidth(cols, sheet1_1.Cells.GetColumnWidth(6));

            for (int row = 0; row <= 4; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_1.Cells.GetRowHeight(row));
            }

            sheet1.Cells.Merge(2, start_col, 1, cols - start_col + 1);
            cols = cols + 1;

            /// 前期小計
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_1.Cells.CreateRange(0, 7, 5, 1));
            sheet1.Cells.SetColumnWidth(cols, sheet1_1.Cells.GetColumnWidth(7));

            for (int row = 0; row <= 4; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_1.Cells.GetRowHeight(row));
            }
            cols = cols + 1;
            start_col = cols;

            /// 本期以前年度
            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_1.Cells.CreateRange(0, 8, 5, 1));
                sheet1.Cells.SetColumnWidth(cols, sheet1_1.Cells.GetColumnWidth(8));

                for (int row = 0; row <= 4; row++)
                {
                    sheet1.Cells.SetRowHeight(row, sheet1_1.Cells.GetRowHeight(row));
                }

                cols = cols + 1;
            }

            /// 本期本年度
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_1.Cells.CreateRange(0, 9, 5, 1));
            sheet1.Cells.SetColumnWidth(cols, sheet1_1.Cells.GetColumnWidth(9));

            for (int row = 0; row <= 4; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_1.Cells.GetRowHeight(row));
            }

            sheet1.Cells.Merge(2, start_col, 1, cols - start_col + 1);
            cols = cols + 1;

            /// 本期小計
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_1.Cells.CreateRange(0, 10, 5, 1));
            sheet1.Cells.SetColumnWidth(cols, sheet1_1.Cells.GetColumnWidth(10));

            for (int row = 0; row <= 4; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_1.Cells.GetRowHeight(row));
            }
            cols = cols + 1;

            /// 其他欄位
            sheet1.Cells.CreateRange(0, cols, 5, 3).Copy(sheet1_1.Cells.CreateRange(0, 11, 5, 3));
            sheet1.Cells.SetColumnWidth(cols, sheet1_1.Cells.GetColumnWidth(11));
            sheet1.Cells.SetColumnWidth(cols + 1, sheet1_1.Cells.GetColumnWidth(12));
            sheet1.Cells.SetColumnWidth(cols + 2, sheet1_1.Cells.GetColumnWidth(13));
            sheet1.Cells.SetRowHeight(0, sheet1_1.Cells.GetRowHeight(0));
            sheet1.Cells.SetRowHeight(1, sheet1_1.Cells.GetRowHeight(1));
            sheet1.Cells.SetRowHeight(2, sheet1_1.Cells.GetRowHeight(2));
            sheet1.Cells.SetRowHeight(3, sheet1_1.Cells.GetRowHeight(3));
            sheet1.Cells.SetRowHeight(4, sheet1_1.Cells.GetRowHeight(4));
            cols = cols + 2;
            line = 5;

            for (i = 0; i < dt.Rows.Count - 1; i++)
            {
                sheet1.Cells.CreateRange(line + i, 0, 1, cols + 1).Copy(sheet1.Cells.CreateRange(4, 0, 1, cols + 1));
                sheet1.Cells.SetRowHeight(line + i, sheet1.Cells.GetRowHeight(4));
            }

            sheet1.Cells[0, 0].PutValue(PJ_NAME + "經費調整表");
            // Create a new style
            Style style = sheet1.Cells[0, 0].GetStyle();

            // Set the horizontal alignment in the style
            style.HorizontalAlignment = TextAlignmentType.CenterAcross;

            // Apply the style to the cell
            sheet1.Cells[0, 0].SetStyle(style);
            sheet1.Cells.SetRowHeight(0, sheet1.Cells.GetRowHeight(1));
            sheet1.Cells.Merge(1, 0, 1, 3);
            sheet1.Cells[1, 0].PutValue("更新日期" + DateTime.Now);
            cols = 5;

            switch (PSEQ_NO)
            {
                case "":
                    //sheet1.Cells["F3"].PutValue("第" + SEQ_NO + "次核定預算");
                    sheet1.Cells["F3"].PutValue("原預算");
                    sheet1.Cells[2, cols + ditdt.Rows.Count + 1].PutValue("小計");
                    sheet1.Cells[2, cols + ditdt.Rows.Count + 2].PutValue("修正後");
                    break;
                default:
                    sheet1.Cells["F3"].PutValue("第" + PSEQ_NO + "期核定預算");
                    sheet1.Cells[2, cols + ditdt.Rows.Count + 1].PutValue("小計");
                    sheet1.Cells[2, cols + ditdt.Rows.Count + 2].PutValue("第" + SEQ_NO + "期核定預算");
                    break;
            }

            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                sheet1.Cells[3, cols].PutValue(ditdt.Rows[i]["PID_SYR"]);
                cols = cols + 1;
            }

            //目前年度
            sheet1.Cells[3, cols].PutValue(SYR);
            cols = cols + 2;
            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                sheet1.Cells[3, cols].PutValue(ditdt.Rows[i]["PID_SYR"]);
                cols = cols + 1;
            }


            sheet1.Cells[3, cols].PutValue(SYR);
            line = 4;
            int line_A = 5;
            sid = 0;

            for (i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i]["TD_NAME"].ToString() != "小計")
                {
                    sid++;
                    sheet1.Cells["A" + (line_A + i).ToString()].PutValue(sid);
                }
                else
                {
                    sheet1.Cells["A" + (line_A + i).ToString()].PutValue(string.Empty);
                }
                sheet1.Cells["B" + (line_A + i)].PutValue(dt.Rows[i]["CITY"].ToString());
                sheet1.Cells["C" + (line_A + i)].PutValue(dt.Rows[i]["SIP_TD_CODE"].ToString());
                sheet1.Cells["D" + (line_A + i)].PutValue(dt.Rows[i]["TOWN"].ToString());
                sheet1.Cells["E" + (line_A + i)].PutValue(dt.Rows[i]["TD_NAME"].ToString());

                switch (SEQ_NO)
                {
                    case "0":

                        /// 前期
                        cols = 5;
                        sum_FUND = 0;
                        for (j = 0; j < ditdt.Rows.Count; j++)
                        {
                            sheet1.Cells[line + i, cols].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"]].ToString()));
                            cols = cols + 1;
                            sum_FUND = sum_FUND + Convert.ToDouble(Tools.StringToDouble(dt.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString()));
                        }

                        // 目前年度
                        sheet1.Cells[line + i, cols].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["YR_FUND"].ToString()));
                        cols = cols + 1;
                        sum_FUND = sum_FUND + Convert.ToDouble(Tools.StringToDouble(dt.Rows[i]["YR_FUND"].ToString()));

                        /// 前期小計
                        sheet1.Cells[line + i, cols].PutValue(sum_FUND);
                        diff_FUND = sum_FUND;

                        /// 本期
                        cols = cols + 1;
                        sum_FUND = 0;
                        for (j = 0; j < ditdt.Rows.Count; j++)
                        {
                            sheet1.Cells[line + i, cols].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"]].ToString()));
                            cols = cols + 1;
                            sum_FUND = sum_FUND + Convert.ToDouble(Tools.StringToDouble(dt.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString()));
                        }

                        /// 目前年度
                        sheet1.Cells[line + i, cols].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["YR_FUND"].ToString()));
                        cols = cols + 1;
                        sum_FUND = sum_FUND + Convert.ToDouble(Tools.StringToDouble(dt.Rows[i]["YR_FUND"].ToString()));

                        /// 本期小計
                        sheet1.Cells[line + i, cols].PutValue(sum_FUND);
                        cols = cols + 1;

                        /// 增減
                        diff_FUND = diff_FUND - sum_FUND;
                        sheet1.Cells[line + i, cols].PutValue(diff_FUND);
                        cols = cols + 1;

                        /// 備註
                        sheet1.Cells[line + i, cols].PutValue(string.Empty);

                        /// 委託地方代辦
                        cols = cols + 1;
                        switch (dt.Rows[i]["N_IPJ_CODE"].ToString())
                        {
                            case "04":
                                sheet1.Cells[line + i, cols].PutValue(dt.Rows[i]["IPJ_DESC"]);
                                break;
                            default:
                                sheet1.Cells[line + i, cols].PutValue(string.Empty);
                                break;
                        }
                        break;

                    default:

                        /// 前期
                        cols = 5;
                        sum_FUND = 0;
                        for (j = 0; j < ditdt.Rows.Count; j++)
                        {
                            sheet1.Cells[line + i, cols].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString()));
                            cols = cols + 1;
                            sum_FUND = sum_FUND + Convert.ToDouble(Tools.StringToDouble(dt.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString()));
                        }

                        /// 目前年度
                        sheet1.Cells[line + i, cols].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PYR_FUND"].ToString()));
                        cols = cols + 1;
                        sum_FUND = sum_FUND + Convert.ToDouble(Tools.StringToDouble(dt.Rows[i]["PYR_FUND"].ToString()));

                        /// 前期小計
                        sheet1.Cells[line + i, cols].PutValue(sum_FUND);
                        diff_FUND = sum_FUND;

                        /// 本期
                        cols = cols + 1;
                        sum_FUND = 0;
                        for (j = 0; j < ditdt.Rows.Count; j++)
                        {
                            sheet1.Cells[line + i, cols].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString()));
                            cols = cols + 1;
                            sum_FUND = sum_FUND + Convert.ToDouble(Tools.StringToDouble(dt.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString()));
                        }

                        /// 目前年度
                        sheet1.Cells[line + i, cols].PutValue(Source.Tools.StringToDouble(dt.Rows[i]["YR_FUND"].ToString()));
                        cols = cols + 1;
                        sum_FUND = sum_FUND + Convert.ToDouble(Tools.StringToDouble(dt.Rows[i]["YR_FUND"].ToString()));

                        /// 本期小計
                        sheet1.Cells[line + i, cols].PutValue(sum_FUND);
                        cols = cols + 1;

                        /// 增減
                        diff_FUND = diff_FUND - sum_FUND;
                        sheet1.Cells[line + i, cols].PutValue(diff_FUND);
                        cols = cols + 1;

                        /// 備註
                        sheet1.Cells[line + i, cols].PutValue(string.Empty);

                        /// 委託地方代辦
                        cols = cols + 1;
                        switch (dt.Rows[i]["N_IPJ_CODE"].ToString())
                        {
                            case "04":
                                sheet1.Cells[line + i, cols].PutValue(dt.Rows[i]["IPJ_DESC"]);
                                break;
                            default:
                                sheet1.Cells[line + i, cols].PutValue(string.Empty);
                                break;
                        }

                        break;
                }
            }

            #region 20241008 => 新增總計

            // 先固定 MaxRow 的值，避免插入行後範圍變化
            int MaxRow = sheet1.Cells.MaxDataRow;
            int MaxCol = sheet1.Cells.MaxDataColumn;

            // 在最後一行下方插入一行，複製最後一行的樣式到新行
            sheet1.Cells.InsertRow(MaxRow + 1);
            sheet1.Cells.CopyRows(sheet1.Cells, MaxRow, MaxRow + 1, 1);

            // 清除新行的內容，保留格式
            int newRowIndex = MaxRow + 1;
            for (int col = 0; col <= MaxCol; col++)
            {
                sheet1.Cells[newRowIndex, col].PutValue(string.Empty);  // 清除值
            }

            // 設置新行的內容，插入 "總計" 字樣
            sheet1.Cells[newRowIndex, 4].PutValue("總計");

            // 初始化加總變量
            double sumF = 0;
            double sumG = 0;
            double sumH = 0;
            double sumI = 0;
            double sumJ = 0;

            // 遍歷每一行，對 F, G, H, I 列進行加總
            for (int row = 4; row <= MaxRow; row++)  // 假設數據從第5行（索引 4）開始
            {
                // 檢查並確保單元格內容是數值
                if (double.TryParse(sheet1.Cells[row, 5].Value?.ToString(), out double valueF))
                {
                    sumF += valueF;
                }
                if (double.TryParse(sheet1.Cells[row, 6].Value?.ToString(), out double valueG))
                {
                    sumG += valueG;
                }
                if (double.TryParse(sheet1.Cells[row, 7].Value?.ToString(), out double valueH))
                {
                    sumH += valueH;
                }
                if (double.TryParse(sheet1.Cells[row, 8].Value?.ToString(), out double valueI))
                {
                    sumI += valueI;
                }
                if (double.TryParse(sheet1.Cells[row, 9].Value?.ToString(), out double valueJ))
                {
                    sumJ += valueJ;
                }
            }

            // 將加總結果寫入相應的列
            sheet1.Cells[newRowIndex, 5].PutValue(sumF / 2);  // F 列
            sheet1.Cells[newRowIndex, 6].PutValue(sumG / 2);  // G 列
            sheet1.Cells[newRowIndex, 7].PutValue(sumH / 2);  // H 列
            sheet1.Cells[newRowIndex, 8].PutValue(sumI / 2);  // I 列
            sheet1.Cells[newRowIndex, 9].PutValue(sumJ / 2);  // I 列

            #endregion
        }

        public static void CreateSheet2(Workbook workbook, Worksheet sheet1, DataTable dt, DataTable ditdt, string SYR, string PSEQ_NO, string SEQ_NO, string PJ_NAME)
        {
            // 取得工作表 "Sheet1_2"
            Worksheet sheet1_2 = workbook.Worksheets["Sheet1_2"];
            int i, j, k, m, start_col, last_col, line, cols;

            /// 目前年與月（民國年）
            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");

            /// 前一年民國年
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");

            /// 暫存變數，用來存放科目 ID、名稱等資料
            string tempACC_ID = "";
            string tempACC_ID_NAME = "";
            string tempCATY_TYPE = "";
            string tempCATY_DESC = "";

            /// 篩選 DataRow 的結果
            DataRow[] rr;

            if (SYR != "")
            {
                sSYR = SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }

            cols = 2;
            start_col = 2;

            /// 複製 Sheet1 的前 5 行資料到 Sheet1_2
            sheet1.Cells.CreateRange(0, 0, 5, 2).Copy(sheet1_2.Cells.CreateRange(0, 0, 5, 2));

            /// 複製前 5 行的行高
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }

            /// 處理以前年度資料
            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 2, 5, 1));

                /// 設置列寬和行高
                sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(2));
                sheet1.Cells.SetRowHeight(0, sheet1_2.Cells.GetRowHeight(0));
                sheet1.Cells.SetRowHeight(1, sheet1_2.Cells.GetRowHeight(1));
                sheet1.Cells.SetRowHeight(2, sheet1_2.Cells.GetRowHeight(2));
                sheet1.Cells.SetRowHeight(3, sheet1_2.Cells.GetRowHeight(3));
                sheet1.Cells.SetRowHeight(4, sheet1_2.Cells.GetRowHeight(4));

                cols = cols + 1;
                sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 3, 5, 1));

                /// 設置列寬和行高
                sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(3));
                for (int row = 0; row < 5; row++)
                {
                    sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
                }
                cols = cols + 1;
                sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 4, 5, 1));

                /// 設置列寬和行高
                sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(4));
                for (int row = 0; row < 5; row++)
                {
                    sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
                }

                /// 合併儲存格
                sheet1.Cells.Merge(2, start_col, 1, cols - start_col + 1);
                cols = cols + 1;
                start_col = cols;
            }

            /// 前期本年度
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 2, 5, 1));

            /// 設置列寬和行高
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(2));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }
            cols = cols + 1;
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 3, 5, 1));

            /// 設置列寬和行高
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(3));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }
            cols = cols + 1;
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 4, 5, 1));

            /// 設置列寬和行高
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(4));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }
            sheet1.Cells.Merge(2, start_col, 1, cols - start_col + 1);

            cols = cols + 1;
            start_col = cols;

            /// 增減小計
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 8, 5, 1));

            /// 設置列寬和行高
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(8));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }
            cols = cols + 1;

            /// 調整後總計
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 9, 5, 1));

            /// 設置列寬和行高
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(9));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }

            last_col = cols + 1;

            /// Create New FormData Start
            Style style = sheet1.Cells[0, 0].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.CenterAcross;
            sheet1.Cells[0, 0].SetStyle(style);

            sheet1.Cells.Merge(0, 0, 1, cols);
            sheet1.Cells[0, 0].PutValue(PJ_NAME + "經費調整表");

            sheet1.Cells.Merge(1, 0, 1, 3);
            sheet1.Cells[1, 0].PutValue("更新日期" + DateTime.Now);

            cols = cols + 1;
            line = 5;
            cols = 2;

            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                sheet1.Cells[2, cols].PutValue(ditdt.Rows[i]["PID_SYR"].ToString());
                cols = cols + 3;
            }

            /// 目前年度
            sheet1.Cells[2, cols].PutValue(SYR);
            line = 4;
            int line_A = 5;

            /// 設備及投資(署辦工程) / 獎補助費(地方辦理) / 業務費
            for (m = 0; m < 3; m++)
            {
                switch (m)
                {
                    case 0:
                        tempACC_ID = "0300";
                        tempACC_ID_NAME = "設備及投資(署辦工程)";
                        break;
                    case 1:
                        tempACC_ID = "0400";
                        tempACC_ID_NAME = "獎補助費(地方辦理)";
                        break;
                    case 2:
                        tempACC_ID = "0200";
                        tempACC_ID_NAME = "業務費";
                        break;
                }

                /// 項目別小計
                sheet1.Cells.CreateRange(line, 0, 2, last_col).Copy(sheet1_2.Cells.CreateRange(4, 0, 1, last_col));
                sheet1.Cells.SetRowHeight(line, sheet1_2.Cells.GetRowHeight(4));

                cols = 2;
                sheet1.Cells.Merge(line, 0, 1, 2);
                sheet1.Cells["A" + (line_A).ToString()].PutValue(tempACC_ID_NAME);

                // 設置背景以及Style
                int rgbValue = Color.FromArgb(217, 217, 217).ToArgb();
                Color color = Color.FromArgb(rgbValue);
                Style style1 = sheet1.Cells["A" + (line_A).ToString()].GetStyle();
                style1.ForegroundColor = color;
                style1.Pattern = BackgroundType.Solid;
                sheet1.Cells["A" + (line_A).ToString()].SetStyle(style1);

                for (j = 0; j < ditdt.Rows.Count; j++)
                {
                    if (PSEQ_NO != "")
                    {
                        /// 目前年度調整前顯示數值
                        sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "ACC_ID='" + tempACC_ID + "'"));

                        style1 = sheet1.Cells[line, cols].GetStyle();
                        style1.ForegroundColor = color;
                        style1.Pattern = BackgroundType.Solid;
                        sheet1.Cells[line, cols].SetStyle(style1);
                        cols = cols + 1;

                        /// 目前年度調整後的值
                        sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "ACC_ID='" + tempACC_ID + "'"));

                        style1 = sheet1.Cells[line, cols].GetStyle();
                        style1.ForegroundColor = color;
                        style1.Pattern = BackgroundType.Solid;
                        sheet1.Cells[line, cols].SetStyle(style1);
                        cols = cols + 2;
                    }
                    else
                    {
                        sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "ACC_ID='" + tempACC_ID + "'"));
                        style1 = sheet1.Cells[line, cols].GetStyle();
                        style1.ForegroundColor = color;
                        style1.Pattern = BackgroundType.Solid;
                        sheet1.Cells[line, cols].SetStyle(style1);
                        cols = cols + 1;

                        sheet1.Cells[line, cols].PutValue("0");
                        style1 = sheet1.Cells[line, cols].GetStyle();
                        style1.ForegroundColor = color;
                        style1.Pattern = BackgroundType.Solid;
                        sheet1.Cells[line, cols].SetStyle(style1);
                        cols = cols + 2;
                    }
                }

                if (PSEQ_NO != "")
                {
                    // 目前年度調整前
                    sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_FUND", "ACC_ID='" + tempACC_ID + "'"));

                    style1 = sheet1.Cells[line, cols].GetStyle();
                    style1.ForegroundColor = color;
                    style1.Pattern = BackgroundType.Solid;
                    sheet1.Cells[line, cols].SetStyle(style1);
                    cols = cols + 1;

                    // 目前年度調整後
                    sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "YR_FUND", "ACC_ID='" + tempACC_ID + "'"));

                    style1 = sheet1.Cells[line, cols].GetStyle();
                    style1.ForegroundColor = color;
                    style1.Pattern = BackgroundType.Solid;
                    sheet1.Cells[line, cols].SetStyle(style1);
                    cols = cols + 2;
                }
                else
                {
                    // 目前年度調整前
                    sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "YR_FUND", "ACC_ID='" + tempACC_ID + "'"));

                    style1 = sheet1.Cells[line, cols].GetStyle();
                    style1.ForegroundColor = color;
                    style1.Pattern = BackgroundType.Solid;
                    sheet1.Cells[line, cols].SetStyle(style1);
                    cols = cols + 1;

                    /// 目前年度調整後
                    sheet1.Cells[line, cols].PutValue("0");

                    style1 = sheet1.Cells[line, cols].GetStyle();
                    style1.ForegroundColor = color;
                    style1.Pattern = BackgroundType.Solid;
                    sheet1.Cells[line, cols].SetStyle(style1);
                    cols = cols + 2;
                }

                /// 增減小計
                style1 = sheet1.Cells[line, cols].GetStyle();
                style1.ForegroundColor = color;
                style1.Pattern = BackgroundType.Solid;
                sheet1.Cells[line, cols].SetStyle(style1);
                line++;
                line_A++;

                /// 經常門 / 經常門 / 總計
                for (k = 0; k < 2; k++)
                {
                    cols = 2;

                    switch (k)
                    {
                        case 0:
                            tempCATY_TYPE = "02";
                            tempCATY_DESC = "經常門";
                            break;

                        case 1:
                            tempCATY_TYPE = "01";
                            tempCATY_DESC = "資本門";
                            break;
                    }

                    rr = dt.Select("ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'");

                    if (rr.Length > 0)
                    {
                        sheet1.Cells.CreateRange(line, 0, 1, last_col).Copy(sheet1_2.Cells.CreateRange(4, 0, 1, last_col));
                        sheet1.Cells.SetRowHeight(line, sheet1_2.Cells.GetRowHeight(4));
                        sheet1.Cells.Merge(line, 0, 1, 2);
                        sheet1.Cells["A" + (line + 1).ToString()].PutValue(tempCATY_DESC);

                        for (j = 0; j < ditdt.Rows.Count; j++)
                        {

                            var PYRYR_FUND = RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'");

                            if (PSEQ_NO != "")
                            {
                                sheet1.Cells[line, cols].Value = PYRYR_FUND;
                                cols = cols + 1;

                                sheet1.Cells[line, cols].Value = PYRYR_FUND;
                                cols = cols + 2;
                            }
                            else
                            {
                                sheet1.Cells.CreateRange(line, cols, 1, 2).Value = new object[] { PYRYR_FUND, "0" };
                                cols = cols + 3;
                            }

                        }

                        var PYRFUND = RainReport_Source.getDtColumns_SUM(dt, "PYR_FUND", "ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'");
                        var YRFUND = RainReport_Source.getDtColumns_SUM(dt, "YR_FUND", "ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'");

                        if (PSEQ_NO != "")
                        {
                            sheet1.Cells.CreateRange(line, cols, 1, 2).Value = new object[] { PYRFUND, YRFUND };
                            cols = cols + 2;
                        }
                        else
                        {
                            sheet1.Cells.CreateRange(line, cols, 1, 2).Value = new object[] { YRFUND, "0" };
                            cols = cols + 2;
                        }

                        line++;
                        line_A++;

                        /// 項目別明細
                        /// 經費調整表-明細
                        /// ex. 052 公建-都市總合治水建設計畫(設)
                        for (i = 0; i < rr.Length; i++)
                        {
                            sheet1.Cells.CreateRange(line, 0, 1, last_col).Copy(sheet1_2.Cells.CreateRange(4, 0, 1, last_col));
                            sheet1.Cells.SetRowHeight(line, sheet1_2.Cells.GetRowHeight(4));
                            cols = 1;

                            sheet1.Cells["A" + (line + 1).ToString()].PutValue(rr[i]["CATY_CODE"].ToString());
                            sheet1.Cells["B" + (line + 1).ToString()].PutValue(rr[i]["CATY_NAME"].ToString());
                            cols = 2;

                            for (j = 0; j < ditdt.Rows.Count; j++)
                            {
                                if (PSEQ_NO != "")
                                {
                                    var range = sheet1.Cells.CreateRange(line, cols, 1, 2);
                                    range[0, 0].PutValue(rr[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString());
                                    range[0, 1].PutValue(rr[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString());
                                    cols = cols + 2;
                                }
                                else
                                {
                                    var range = sheet1.Cells.CreateRange(line, cols, 1, 2);
                                    range[0, 0].PutValue(rr[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString());
                                    range[0, 1].PutValue("0");
                                    cols = cols + 2;
                                }
                            }

                            if (PSEQ_NO != "")
                            {
                                var range = sheet1.Cells.CreateRange(line, cols, 1, 2);
                                range[0, 0].PutValue(rr[i]["PYR_FUND"].ToString());
                                range[0, 1].PutValue(rr[i]["YR_FUND"].ToString());
                                cols = cols + 2;
                            }
                            else
                            {
                                var range = sheet1.Cells.CreateRange(line, cols, 1, 2);
                                range[0, 0].PutValue(rr[i]["YR_FUND"].ToString());
                                range[0, 1].PutValue("0");
                                cols = cols + 2;
                            }

                            line++;
                            line_A++;
                        }

                        #region 總計 = 經常門 + 資本門

                        decimal Total_YRFUND = 0;  // 調整前總和 (YR_FUND)
                        decimal Total_PYRFUND = 0;  // 調整後總和 (PYR_FUND)

                        foreach (DataRow row in dt.Rows)
                        {
                            // 取出 YR_FUND 和 PYR_FUND 字串
                            string yr_fund_str = row["YR_FUND"].ToString();
                            string pyr_fund_str = row["PYR_FUND"].ToString();

                            decimal yr_fund = 0;
                            decimal pyr_fund = 0;

                            if (!string.IsNullOrEmpty(yr_fund_str) && decimal.TryParse(yr_fund_str, out yr_fund))
                            {
                                Total_YRFUND += yr_fund; // 累加 YR_FUND 的值
                            }

                            if (!string.IsNullOrEmpty(pyr_fund_str) && decimal.TryParse(pyr_fund_str, out pyr_fund))
                            {
                                Total_PYRFUND += pyr_fund; // 累加 PYR_FUND 的值
                            }
                        }


                        sheet1.Cells.CreateRange(line, 0, 1, last_col).Copy(sheet1_2.Cells.CreateRange(4, 0, 1, last_col));
                        sheet1.Cells.SetRowHeight(line, sheet1_2.Cells.GetRowHeight(4));
                        sheet1.Cells.Merge(line, 0, 1, 2);
                        sheet1.Cells["A" + (line + 1).ToString()].PutValue("總計");
                        sheet1.Cells[line, 2].PutValue(Total_YRFUND);   // 調整前總和
                        sheet1.Cells[line, 3].PutValue(Total_PYRFUND);  // 調整後總和
                        cols = cols + 2;  // 確保後續儲存格的正確處理

                        #endregion
                    }
                }
            }
        }

        public static void CreateSheet3(Workbook workbook, Worksheet sheet1, DataTable dt, DataTable ditdt, string SYR, string PSEQ_NO, string SEQ_NO, string PJ_NAME)
        {
            Worksheet sheet1_2 = workbook.Worksheets["Sheet1_2"];

            int i;          /// 迴圈控制
            int j;          /// 迴圈控制
            int k;          /// 迴圈控制 (經資門處理)
            int m;          /// 處理不同類型的 ACC_ID
            int start_col;  /// 追蹤目前開始的欄位
            int last_col;   /// 追蹤目前最後的欄位
            int line;       /// 追蹤目前處理的行數
            int cols;       /// 計算目前欄位位置

            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");     /// 目前民國年
            string sSMN = (System.DateTime.Now.Month).ToString("00");            /// 目前月份
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");/// 前一年

            string tempACC_ID = "";     /// 暫存當前 ACC_ID
            string tempACC_ID_NAME = "";/// 暫存當前 ACC_ID 名稱
            string tempCATY_TYPE = "";  /// 暫存類型 "經常門" 或 "資本門"
            string tempCATY_DESC = "";  /// 類型描述
            string ACC_ID_WHERE = "";   /// 用來形成 SQL 的 WHERE 條件

            DataRow[] rr;               /// 用來存放篩選結果的資料列陣列

            // 如果傳入 SYR (目前年度) 不為空，則設定 sSYR 和 sPSYR
            if (SYR != "")
            {
                sSYR = SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }

            cols = 2;
            start_col = 2;

            /// 複製 sheet1_2 的前 5 行的樣式到 sheet1
            sheet1.Cells.CreateRange(0, 0, 5, 2).Copy(sheet1_2.Cells.CreateRange(0, 0, 5, 2));

            /// 設定前 5 行的高度
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }

            /// 調整前以前年度
            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                /// 複製樣式到對應欄位
                sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 2, 5, 1));
                sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(2));

                /// 設定列高
                for (int row = 0; row < 5; row++)
                {
                    sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
                }
                cols = cols + 1;

                /// 重複樣式和欄位寬度設定
                sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 3, 5, 1));
                sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(3));

                for (int row = 0; row < 5; row++)
                {
                    sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
                }
                cols = cols + 1;

                /// 設定新的一年欄位
                sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 4, 5, 1));
                sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(4));
                for (int row = 0; row < 5; row++)
                {
                    sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
                }

                /// 合併欄位
                sheet1.Cells.Merge(2, start_col, 1, cols - start_col + 1);
                cols = cols + 1;
                start_col = cols;
            }

            /// 前期本年度
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 2, 5, 1));
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(2));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }
            cols = cols + 1;

            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 3, 5, 1));
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(3));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }
            cols = cols + 1;

            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 4, 5, 1));
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(4));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }

            /// 合併欄位
            sheet1.Cells.Merge(2, start_col, 1, cols - start_col + 1);
            cols = cols + 1;
            start_col = cols;

            /// 增加"增減小計"欄位
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 8, 5, 1));
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(8));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }
            cols = cols + 1;

            /// 調整後總計
            sheet1.Cells.CreateRange(0, cols, 5, 1).Copy(sheet1_2.Cells.CreateRange(0, 9, 5, 1));
            sheet1.Cells.SetColumnWidth(cols, sheet1_2.Cells.GetColumnWidth(9));
            for (int row = 0; row < 5; row++)
            {
                sheet1.Cells.SetRowHeight(row, sheet1_2.Cells.GetRowHeight(row));
            }
            last_col = cols + 1;

            /// 合併標體行，並設置表名
            sheet1.Cells.Merge(0, 0, 1, cols);
            sheet1.Cells[0, 0].PutValue($"{PJ_NAME}經費調整表");

            /// Create a new style
            /// Set the horizontal alignment in the style
            /// Apply the style to the cell
            Style style = sheet1.Cells[0, 0].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.CenterAcross;
            sheet1.Cells[0, 0].SetStyle(style);

            /// 設定更新日期行
            sheet1.Cells.Merge(1, 0, 1, 3);
            sheet1.Cells[1, 0].PutValue("更新日期" + DateTime.Now);

            cols = cols + 1;
            line = 5;
            cols = 2;

            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                sheet1.Cells[2, cols].PutValue(ditdt.Rows[i]["PID_SYR"].ToString());
                cols = cols + 3;
            }

            /// 目前年度
            sheet1.Cells[2, cols].PutValue(SYR);
            line = 4;

            /// (設備及投資、獎補助費、業務費、總計)
            for (m = 0; m < 4; m++)
            {
                switch (m)
                {
                    case 0:
                        tempACC_ID = "0300"; /// 設備及投資
                        tempACC_ID_NAME = "設備及投資(署辦工程)";
                        ACC_ID_WHERE = " ACC_ID='" + tempACC_ID + "'";
                        break;
                    case 1:
                        tempACC_ID = "0400";
                        tempACC_ID_NAME = "獎補助費(地方辦理)";
                        ACC_ID_WHERE = " ACC_ID='" + tempACC_ID + "'";
                        break;
                    case 2:
                        tempACC_ID = "0200";
                        tempACC_ID_NAME = "業務費";
                        ACC_ID_WHERE = " ACC_ID='" + tempACC_ID + "'";
                        break;
                    case 3:
                        tempACC_ID = "";
                        tempACC_ID_NAME = "總計";
                        ACC_ID_WHERE = "";
                        break;
                }

                /// 項目別小計
                sheet1.Cells.CreateRange(line, 0, 1, last_col).Copy(sheet1_2.Cells.CreateRange(4, 0, 1, last_col));
                sheet1.Cells.SetRowHeight(line, sheet1_2.Cells.GetRowHeight(4));

                cols = 2;
                sheet1.Cells.Merge(line, 0, 1, 2);
                sheet1.Cells["A" + (line + 1).ToString()].PutValue(tempACC_ID_NAME);

                for (j = 0; j < ditdt.Rows.Count; j++)
                {
                    if (PSEQ_NO != "")
                    {
                        var range = sheet1.Cells.CreateRange(line, cols, 1, 3);
                        range[0, 0].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), ACC_ID_WHERE));
                        range[0, 1].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), ACC_ID_WHERE));
                        style = sheet1.Cells.GetCellStyle(0, cols + 2);
                        style.ForegroundColor = Color.FromArgb(255, 255, 255);
                        range[0, 2].SetStyle(style);
                        cols = cols + 3;
                    }
                    else
                    {
                        var range = sheet1.Cells.CreateRange(line, cols, 1, 3);
                        range[0, 0].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), ACC_ID_WHERE));
                        range[0, 1].PutValue("0");

                        style = sheet1.Cells.GetCellStyle(0, cols + 2);
                        style.ForegroundColor = Color.FromArgb(255, 255, 255);
                        range[0, 2].SetStyle(style);

                        cols = cols + 3;
                    }
                }

                if (PSEQ_NO != "")
                {
                    var range = sheet1.Cells.CreateRange(line, cols, 1, 3);
                    range[0, 0].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_FUND", ACC_ID_WHERE));/// 目前年度調整前
                    range[0, 1].PutValue(RainReport_Source.getDtColumns_SUM(dt, "YR_FUND", ACC_ID_WHERE)); /// 目前年度調整後

                    style = sheet1.Cells.GetCellStyle(0, cols + 2);
                    style.ForegroundColor = Color.FromArgb(255, 255, 255);
                    range.SetStyle(style);

                    cols = cols + 3;
                }
                else
                {
                    sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "YR_FUND", ACC_ID_WHERE));/// 目前年度調整前
                    cols = cols + 1;

                    sheet1.Cells[line, cols].PutValue("0");/// 目前年度調整後
                    cols = cols + 1;

                    style = sheet1.Cells.GetCellStyle(line, cols);
                    style.ForegroundColor = Color.FromArgb(255, 255, 255);
                    sheet1.Cells[line, cols].SetStyle(style);

                    cols = cols + 1;
                }

                /// 增減小計
                style = sheet1.Cells.GetCellStyle(line, cols);
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.FromArgb(255, 255, 255);
                sheet1.Cells[line, cols].SetStyle(style);
                line++;

                if (tempACC_ID == "0200")
                {
                    /// case 0:(經常門) / case 1:(資本門)
                    for (k = 0; k < 2; k++)
                    {
                        cols = 2;
                        switch (k)
                        {
                            //case 0:
                            //    tempCATY_TYPE = "02";
                            //    tempCATY_DESC = "經常門";
                            //    break;
                            case 1:
                                tempCATY_TYPE = "01";
                                tempCATY_DESC = "資本門";
                                break;
                        }
                        rr = dt.Select("ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'");
                        if (rr.Length > 0)
                        {
                            sheet1.Cells.CreateRange(line, 0, 1, last_col).Copy(sheet1_2.Cells.CreateRange(4, 0, 1, last_col));
                            sheet1.Cells.SetRowHeight(line, sheet1_2.Cells.GetRowHeight(4));
                            sheet1.Cells.Merge(line, 0, 1, 2);
                            sheet1.Cells["A" + (line + 1).ToString()].PutValue(tempCATY_DESC);

                            for (j = 0; j < ditdt.Rows.Count; j++)
                            {
                                if (PSEQ_NO != "")
                                {
                                    sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'"));
                                    cols = cols + 1;

                                    sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'"));
                                    cols = cols + 1;

                                    style = sheet1.Cells.GetCellStyle(line, cols);
                                    style.ForegroundColor = Color.FromArgb(255, 255, 255);
                                    sheet1.Cells[line, cols].SetStyle(style);
                                    cols = cols + 1;
                                }
                                else
                                {
                                    sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'"));
                                    cols = cols + 1;

                                    sheet1.Cells[line, cols].PutValue("0");
                                    cols = cols + 1;

                                    style = sheet1.Cells.GetCellStyle(line, cols);
                                    style.ForegroundColor = Color.FromArgb(255, 255, 255);
                                    sheet1.Cells[line, cols].SetStyle(style);
                                    cols = cols + 1;
                                }
                            }
                            if (PSEQ_NO != "")
                            {
                                /// 目前年度調整前
                                sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "PYR_FUND", "ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'"));
                                cols = cols + 1;

                                /// 目前年度調整後
                                sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "YR_FUND", "ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'"));
                                cols = cols + 1;

                                style = sheet1.Cells.GetCellStyle(line, cols);
                                style.ForegroundColor = Color.FromArgb(255, 255, 255);
                                sheet1.Cells[line, cols].SetStyle(style);
                                cols = cols + 1;
                            }
                            else
                            {
                                /// 目前年度調整前
                                sheet1.Cells[line, cols].PutValue(RainReport_Source.getDtColumns_SUM(dt, "YR_FUND", "ACC_ID='" + tempACC_ID + "' and CATY_TYPE='" + tempCATY_TYPE + "'"));
                                cols = cols + 1;

                                /// 目前年度調整後
                                sheet1.Cells[line, cols].PutValue("0");
                                cols = cols + 1;

                                style = sheet1.Cells.GetCellStyle(line, cols);
                                style.ForegroundColor = Color.FromArgb(255, 255, 255);
                                sheet1.Cells[line, cols].SetStyle(style);
                                cols = cols + 1;
                            }
                            line++;
                        }
                    }
                }
            }
        }

        public static DataTable getPJDFDCR_SEQ_NO(string PID, string connectionString)
        {
            //CMBASEEntities db = new CMBASEEntities();

            if (string.IsNullOrEmpty(PID) == false)
            {
                string FSQLStr = "";
                FSQLStr = string.Format("select '第' + Convert(varchar,SEQ_NO) +'期' As SEQ_NON, SEQ_NO from CMEPMS.dbo.PJDFDCR where PID='{0}' Order By SEQ_NO Desc", PID);
                return SourceFunction.DataTableGet(FSQLStr, connectionString);
            }
            return null;
        }

        public static DataTable getBillionProject06Single(string PID, string SYR, string SMN, string connectionString)
        {
            string FSQLStr = "";
            FSQLStr = "select sum(REV_FUND) as OTHER_REV_FUND,sum(FUND) as OTHER_FUND,sum(REV_SMN_FUND) as OTHER_REV_SMN_FUND,sum(SMN_FUND) as OTHER_SMN_FUND,";
            FSQLStr += "sum(REV_COST_FEE) as OTHER_REV_COST_FEE,sum(COST_FEE) as OTHER_COST_FEE,sum(REV_PPAY_FEE) as OTHER_REV_PPAY_FEE,sum(PPAY_FEE) as OTHER_PPAY_FEE from ";
            FSQLStr += "(select a.TD_CODE,0 as REV_FUND,isnull(c.FUND,0) as FUND,0 as REV_SMN_FUND,e.SMN_FUND,";
            FSQLStr += "0 as REV_COST_FEE,isnull(f.COST_FEE,0) as COST_FEE,0 as REV_PPAY_FEE,isnull(f.PPAY_FEE,0) as PPAY_FEE from ";
            FSQLStr += string.Format("(select TD_CODE,TID,FD_CODE from CMEPMS.dbo.CTDMNDS where PID={0}) a ", PID);
            FSQLStr += "inner join (select TD_CODE,case isnull(PD_FG,'') when 'D01' then 'P' else CATY_CODE end YR_WORK from CMETMS.dbo.TDMFORM) b on b.YR_WORK not in ('C','L') and a.TD_CODE=b.TD_CODE and 1=2 ";
            FSQLStr += "inner join (select TID,FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00') c on a.TID=c.TID ";
            FSQLStr += "left join (select d.TID,sum(d.SMN_FUND) as SMN_FUND from  ";
            FSQLStr += string.Format("(select TID,'01' as SMN,isnull(M_01,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'02' as SMN,isnull(M_02,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'03' as SMN,isnull(M_03,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'04' as SMN,isnull(M_04,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'05' as SMN,isnull(M_05,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'06' as SMN,isnull(M_06,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'07' as SMN,isnull(M_07,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'08' as SMN,isnull(M_08,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'09' as SMN,isnull(M_09,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'10' as SMN,isnull(M_10,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'11' as SMN,isnull(M_11,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'12' as SMN,isnull(M_12,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00') d ", SYR);
            FSQLStr += string.Format("where d.SMN<='{0}' group by d.TID) e on a.TID=e.TID ", SMN);
            FSQLStr += string.Format("left join (select FD_CODE,COST_FEE,PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR='{0}' and SMN='{1}') f on a.FD_CODE=f.FD_CODE) z", SYR, SMN);
            DataTable dt = SourceFunction.DataTableGet(FSQLStr, connectionString);
            dt.Columns.Add("NO_FUND");
            FSQLStr = string.Format("select (select PJ_FUND from CMEPMS.dbo.PJMFORM where PID={0}) - ", PID);
            FSQLStr += string.Format("(select sum(FUND) as FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00' and TID in (select TID from CMEPMS.dbo.CTDMNDS where PID={0})) ", PID);
            FSQLStr += "as NO_FUND";
            dt.Rows[0]["NO_FUND"] = SourceFunction.DataTableGet(FSQLStr, connectionString).Rows[0]["NO_FUND"].ToString();
            return dt;
        }

        public static DataTable getBillionProject06Base(string PID, string SYR, string SMN, string connectionString)
        {
            string FSQLStr = "";
            FSQLStr = "select sum(c.FUND)/1000 as OTHER_FUND,sum(e.SMN_FUND)/1000 as OTHER_SMN_FUND,sum(isnull(f.COST_FEE,0)+isnull(f.PPAY_FEE,0))/1000 as OTHER_FEE from ";
            FSQLStr += string.Format("(select TD_CODE,TID,FD_CODE from CMEPMS.dbo.CTDMNDS where PID={0}) a ", PID);
            FSQLStr += "inner join (select TD_CODE,IPJ_CODE,case isnull(PD_FG,'') when 'D01' then 'P' else CATY_CODE end YR_WORK from CMETMS.dbo.TDMFORM) b on b.YR_WORK in ('C','L') and a.TD_CODE=b.TD_CODE ";
            FSQLStr += "inner join (select TID,FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00') c on a.TID=c.TID ";
            FSQLStr += "left join (select d.TID,sum(d.SMN_FUND) as SMN_FUND from ";
            FSQLStr += string.Format("(select TID,'01' as SMN,isnull(M_01,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'02' as SMN,isnull(M_02,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'03' as SMN,isnull(M_03,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'04' as SMN,isnull(M_04,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'05' as SMN,isnull(M_05,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'06' as SMN,isnull(M_06,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'07' as SMN,isnull(M_07,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'08' as SMN,isnull(M_08,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'09' as SMN,isnull(M_09,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'10' as SMN,isnull(M_10,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'11' as SMN,isnull(M_11,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'12' as SMN,isnull(M_12,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00') d ", SYR);
            FSQLStr += string.Format("where d.SMN<='{0}' group by d.TID) e on a.TID=e.TID ", SMN);
            FSQLStr += string.Format("left join (select FD_CODE,COST_FEE,PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR='{0}' and SMN='{1}') f on a.FD_CODE=f.FD_CODE ", SYR, SMN);
            DataTable dt = SourceFunction.DataTableGet(FSQLStr, connectionString);
            dt.Columns.Add("NO_FUND");
            FSQLStr = string.Format("select (select PJ_FUND from CMEPMS.dbo.PJMFORM where PID={0}) - ", PID);
            FSQLStr += string.Format("(select sum(FUND) as FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00' and TID in (select TID from CMEPMS.dbo.CTDMNDS where PID={0})) ", PID);
            FSQLStr += "as NO_FUND";
            dt.Rows[0]["NO_FUND"] = SourceFunction.DataTableGet(FSQLStr, connectionString).Rows[0]["NO_FUND"].ToString();
            return dt;
        }

        public static DataTable getBillionProject06Multi(string USR_ID, string DEPT_NO, string PID, string SYR, string SMN, string connectionString)
        {
            string FSQLStr = "";
            string constr = connectionString;
            if (PID == "38")
            {
                FSQLStr = string.Format("select * from CMEPMS.dbo.PJDFDCR where PID={0} and SEQ_NO=(select max(SEQ_NO) from CMEPMS.dbo.PJDFDCR where PID={0})", PID);
                SqlConnection con1 = new SqlConnection();
                //constr = @"Data Source=10.1.1.117\SEPMIS;initial catalog=CMEPMS;user id=***;password=***;" + "MultipleActiveResultSets=True;" + "initial catalog=CMETMS;";
                con1.ConnectionString = connectionString;
                SqlDataAdapter adp = new SqlDataAdapter(FSQLStr, con1);
                DataSet ds = new DataSet();
                adp.Fill(ds);
                DataTable dtPJDFDCR = ds.Tables[0];
                string sMNDS = "";
                string sSEQ = "";
                if (dtPJDFDCR.Rows.Count == 0)
                {
                    sMNDS = "CTDMNDS";
                }
                else if (dtPJDFDCR.Rows[0]["APROV_DT"].ToString() != "")
                {
                    sMNDS = "CTDMNDS";
                }
                else
                {
                    sMNDS = "CTHMNDS";
                    sSEQ = Convert.ToString(Convert.ToInt16(dtPJDFDCR.Rows[0]["SEQ_NO"].ToString()) - 1);
                }

                FSQLStr = "select z.TD_CODE,z.TD_NAME,z.TD_CITY,z.TD_TOWN,z.IPJ_CODE,z.CODE_DESC,z.YR_WORK,z.TIT_CODE,";
                FSQLStr += "z.REV_FUND,z.FUND,z.REV_SMN_FUND,z.SMN_FUND,z.REV_COST_FEE,z.COST_FEE,";
                FSQLStr += "z.REV_PPAY_FEE,z.PPAY_FEE,z.REV_ECC_UPAY_FEE,z.ECC_UPAY_FEE,";
                FSQLStr += "z.PRE_TK102,z.REL_TK102,z.PRE_TK105,z.REL_TK105,z.PRE_TK107,z.REL_TK107,";
                FSQLStr += "z.PRE_SCHD,z.REL_SCHD,z.EXECT_DESC,z.MEMO,z.AREA,case z.ORD when 1 then case when (charindex('業務費',z.TD_NAME)>0) then 1 else 2 end else 0 end as ORD from ";
                FSQLStr += "(select a.TD_CODE,b.TD_NAME,b.TD_CITY,b.TD_TOWN,b.IPJ_CODE,c.CODE_DESC,b.YR_WORK,d.CODE_DESC as TIT_CODE,";
                FSQLStr += "0 as REV_FUND,e.FUND,0 as REV_SMN_FUND,g.SMN_FUND,0 as REV_COST_FEE,isnull(h.COST_FEE,0) as COST_FEE,";
                FSQLStr += "0 as REV_PPAY_FEE,isnull(h.PPAY_FEE,0) as PPAY_FEE,0 as REV_ECC_UPAY_FEE,isnull(m.ECC_UPAY_FEE,0) as ECC_UPAY_FEE,";
                FSQLStr += "isnull(i.PRE_EDT,'') as PRE_TK102,isnull(i.REL_EDT,'') as REL_TK102,";
                FSQLStr += "isnull(j.PRE_EDT,'') as PRE_TK105,isnull(j.REL_EDT,'') as REL_TK105,";
                FSQLStr += "isnull(k.PRE_EDT,'') as PRE_TK107,isnull(k.REL_EDT,'') as REL_TK107,";
                FSQLStr += "isnull(m.PRE_SCHD,0) as PRE_SCHD,isnull(m.REL_SCHD,0) as REL_SCHD,m.EXECT_DESC,m.DELAY_RESN as MEMO,";
                FSQLStr += "n.AREA,n.CITY_ID,o.ORD from ";

                if (sMNDS == "CTDMNDS")
                {
                    FSQLStr += string.Format("(select distinct TD_CODE,TID,FD_CODE from CMEPMS.dbo.CTDMNDS where PID={0}) a ", PID);
                }
                else
                {
                    FSQLStr += string.Format("(select distinct TD_CODE,TID,FD_CODE from CMEPMS.dbo.CTHMNDS where PID={0} and SEQ_NO={1}) a ", PID, sSEQ);
                }

                FSQLStr += "inner join (select TD_CODE,TD_NAME,TD_CITY,TD_TOWN,IPJ_CODE,case isnull(PD_FG,'') when 'D01' then 'P' else CATY_CODE end YR_WORK from CMETMS.dbo.TDMFORM) b on a.TD_CODE=b.TD_CODE ";
                //FSQLStr += "inner join (select TD_CODE,TD_NAME,TD_CITY,IPJ_CODE,case isnull(PD_FG,'') when 'D01' then 'P' else CATY_CODE end YR_WORK from CMETMS.dbo.TDMFORM) b on b.YR_WORK in ('C','L') and a.TD_CODE=b.TD_CODE ";

                FSQLStr += "inner join (select CATY_COD,CODE_DESC from CMETMS.dbo.SMCCODE where SYS_ID='SM' and CATY_ID='02') c on b.IPJ_CODE=c.CATY_COD ";
                FSQLStr += "left join (select CATY_COD,CODE_DESC from CMETMS.dbo.SMCCODE where SYS_ID='TT' and CATY_ID='04') d on b.YR_WORK=d.CATY_COD ";

                if (sMNDS == "CTDMNDS")
                {
                    FSQLStr += "inner join (select TID,FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00') e on a.TID=e.TID ";
                    FSQLStr += "left join (select f.TID,sum(f.SMN_FUND) as SMN_FUND from ";
                    FSQLStr += string.Format("(select TID,'01' as SMN,isnull(M_01,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'02' as SMN,isnull(M_02,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'03' as SMN,isnull(M_03,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'04' as SMN,isnull(M_04,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'05' as SMN,isnull(M_05,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'06' as SMN,isnull(M_06,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'07' as SMN,isnull(M_07,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'08' as SMN,isnull(M_08,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'09' as SMN,isnull(M_09,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'10' as SMN,isnull(M_10,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'11' as SMN,isnull(M_11,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                    FSQLStr += string.Format("union select TID,'12' as SMN,isnull(M_12,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00') f ", SYR);
                    FSQLStr += string.Format("where f.SMN<='{0}' group by f.TID) g on a.TID=g.TID ", SMN);
                }
                else
                {
                    FSQLStr += string.Format("inner join (select TID,FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}) e on a.TID=e.TID ", sSEQ);
                    FSQLStr += "left join (select f.TID,sum(f.SMN_FUND) as SMN_FUND from ";
                    FSQLStr += string.Format("(select TID,'01' as SMN,isnull(M_01,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'02' as SMN,isnull(M_02,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'03' as SMN,isnull(M_03,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'04' as SMN,isnull(M_04,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'05' as SMN,isnull(M_05,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'06' as SMN,isnull(M_06,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'07' as SMN,isnull(M_07,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'08' as SMN,isnull(M_08,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'09' as SMN,isnull(M_09,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'10' as SMN,isnull(M_10,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'11' as SMN,isnull(M_11,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1} ", SYR, sSEQ);
                    FSQLStr += string.Format("union select TID,'12' as SMN,isnull(M_12,0) as SMN_FUND from CMEPMS.dbo.CTHITFD where SYR='{0}' and FIT_CODE='00' and SEQ_NO={1}) f ", SYR, sSEQ);
                    FSQLStr += string.Format("where f.SMN<='{0}' group by f.TID) g on a.TID=g.TID ", SMN);
                }

                FSQLStr += string.Format("left join (select FD_CODE,COST_FEE,PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR='{0}' and SMN='{1}') h on a.FD_CODE=h.FD_CODE ", SYR, SMN);
                FSQLStr += "left join (select TD_CODE,PRE_EDT,REL_EDT from CMETMS.dbo.TDDSDWB where IT_CODE='TK103') i on b.TD_CODE=i.TD_CODE ";
                FSQLStr += "left join (select TD_CODE,PRE_EDT,REL_EDT from CMETMS.dbo.TDDSDWB where IT_CODE='TK105') j on b.TD_CODE=j.TD_CODE ";
                FSQLStr += "left join (select TD_CODE,PRE_EDT,REL_EDT from CMETMS.dbo.TDDSDWB where IT_CODE='TK107') k on b.TD_CODE=k.TD_CODE ";
                FSQLStr += string.Format("left join (select TD_CODE,ECC_UPAY_FEE,PRE_SCHD,REL_SCHD,EXECT_DESC,DELAY_RESN from CMETMS.dbo.TDDMSCD where SYR='{0}' and SMN='{1}') m on b.TD_CODE=m.TD_CODE ", SYR, SMN);
                //FSQLStr += "left join (select distinct case AREA when '北區' then 1 when '中區' then 2 else 3 end as AREA,CITY_ID,CITY from CMBASE.dbo.SMCPOST) n on b.TD_CITY=n.CITY ";
                FSQLStr += "left join (select AREA_ID As AREA,CITY,CITY_ID from CMBASE.dbo.BAS_SMCPOST_AREA) n on b.TD_CITY=n.CITY ";
                FSQLStr += "left join (select CITY,case CITY when '不分區' then 1 else 0 end as ORD from CMBASE.dbo.BAS_SMCPOST_AREA) o on b.TD_CITY=o.CITY) z ";
                FSQLStr += "where z.REV_FUND+z.FUND>0 ";
                if (db.SECCITY.Where(m => m.USR_ID == USR_ID).Any())
                {
                    FSQLStr += string.Format("and TD_CITY IN (SELECT CITY FROM CMBASE.dbo.SECCITY where USR_ID = '{0}')", USR_ID);
                }
                else
                {
                    FSQLStr += string.Format("and TD_CITY IN (SELECT CITY FROM CMBASE.dbo.SECCITY where DEPT_NO = '{0}')", DEPT_NO);
                }
                FSQLStr += "order by z.ORD,z.AREA,z.CITY_ID,z.TD_TOWN,z.TD_CODE";
            }
            else if (PID == "41")
            {
                FSQLStr = "select z.TD_CODE,z.TD_NAME,z.TD_CITY,z.TD_TOWN,z.IPJ_CODE,z.CODE_DESC,z.YR_WORK,z.TIT_CODE,";
                FSQLStr += "z.REV_FUND,z.FUND,z.REV_SMN_FUND,z.SMN_FUND,z.REV_COST_FEE,z.COST_FEE,";
                FSQLStr += "z.REV_PPAY_FEE,z.PPAY_FEE,z.REV_ECC_UPAY_FEE,z.ECC_UPAY_FEE,";
                FSQLStr += "z.PRE_TK102,z.REL_TK102,z.PRE_TK105,z.REL_TK105,z.PRE_TK107,z.REL_TK107,";
                FSQLStr += "z.PRE_SCHD,z.REL_SCHD,z.EXECT_DESC,z.MEMO,z.AREA,case z.ORD when 1 then case when (charindex('業務費',z.TD_NAME)>0) then 1 else 2 end else 0 end as ORD from ";
                FSQLStr += "(select a.TD_CODE,b.TD_NAME,b.TD_CITY,b.TD_TOWN,b.IPJ_CODE,c.CODE_DESC,b.YR_WORK,d.CODE_DESC as TIT_CODE,";
                FSQLStr += "0 as REV_FUND,e.FUND,0 as REV_SMN_FUND,g.SMN_FUND,0 as REV_COST_FEE,isnull(h.COST_FEE,0) as COST_FEE,";
                FSQLStr += "0 as REV_PPAY_FEE,isnull(h.PPAY_FEE,0) as PPAY_FEE,0 as REV_ECC_UPAY_FEE,isnull(m.ECC_UPAY_FEE,0) as ECC_UPAY_FEE,";
                FSQLStr += "isnull(i.PRE_EDT,'') as PRE_TK102,isnull(i.REL_EDT,'') as REL_TK102,";
                FSQLStr += "isnull(j.PRE_EDT,'') as PRE_TK105,isnull(j.REL_EDT,'') as REL_TK105,";
                FSQLStr += "isnull(k.PRE_EDT,'') as PRE_TK107,isnull(k.REL_EDT,'') as REL_TK107,";
                FSQLStr += "isnull(m.PRE_SCHD,0) as PRE_SCHD,isnull(m.REL_SCHD,0) as REL_SCHD,m.EXECT_DESC,m.DELAY_RESN as MEMO,";
                FSQLStr += "n.AREA,n.CITY_ID,o.ORD from ";
                FSQLStr += string.Format("(select distinct TD_CODE1 as TD_CODE,TID,FD_CODE from CMEPMS.dbo.CRDMNDS where PID={0}) a ", PID);

                FSQLStr += "inner join (select TD_CODE,TD_NAME,TD_CITY,TD_TOWN,IPJ_CODE,case isnull(PD_FG,'') when 'D01' then 'P' else CATY_CODE end YR_WORK from CMETMS.dbo.TDMFORM) b on a.TD_CODE=b.TD_CODE ";
                //FSQLStr += "inner join (select TD_CODE,TD_NAME,TD_CITY,IPJ_CODE,case isnull(PD_FG,'') when 'D01' then 'P' else CATY_CODE end YR_WORK from CMETMS.dbo.TDMFORM) b on b.YR_WORK in ('C','L') and a.TD_CODE=b.TD_CODE ";

                FSQLStr += "inner join (select CATY_COD,CODE_DESC from CMETMS.dbo.SMCCODE where SYS_ID='SM' and CATY_ID='02') c on b.IPJ_CODE=c.CATY_COD ";
                FSQLStr += "inner join (select CATY_COD,CODE_DESC from CMETMS.dbo.SMCCODE where SYS_ID='TT' and CATY_ID='04') d on b.YR_WORK=d.CATY_COD ";
                FSQLStr += "inner join (select TID,REV_FUND as FUND from CMEPMS.dbo.CRDMNDS) e on a.TID=e.TID ";
                FSQLStr += "left join (select f.TID,sum(f.SMN_FUND) as SMN_FUND from ";
                FSQLStr += string.Format("(select TID,'01' as SMN,isnull(M_01,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'02' as SMN,isnull(M_02,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'03' as SMN,isnull(M_03,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'04' as SMN,isnull(M_04,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'05' as SMN,isnull(M_05,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'06' as SMN,isnull(M_06,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'07' as SMN,isnull(M_07,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'08' as SMN,isnull(M_08,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'09' as SMN,isnull(M_09,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'10' as SMN,isnull(M_10,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'11' as SMN,isnull(M_11,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}'", SYR);
                FSQLStr += string.Format("union select TID,'12' as SMN,isnull(M_12,0) as SMN_FUND from CMEPMS.dbo.CRDMNDS where SYR='{0}') f ", SYR);
                FSQLStr += string.Format("where f.SMN<='{0}' group by f.TID) g on a.TID=g.TID ", SMN);
                FSQLStr += string.Format("left join (select FD_CODE,COST_FEE,PPAY_FEE from CMEPMS.dbo.TDDSTRF where SYR='{0}' and SMN='{1}') h on a.FD_CODE=h.FD_CODE ", SYR, SMN);
                FSQLStr += "left join (select TD_CODE,PRE_EDT,REL_EDT from CMETMS.dbo.TDDSDWB where IT_CODE='TK103') i on b.TD_CODE=i.TD_CODE ";
                FSQLStr += "left join (select TD_CODE,PRE_EDT,REL_EDT from CMETMS.dbo.TDDSDWB where IT_CODE='TK105') j on b.TD_CODE=j.TD_CODE ";
                FSQLStr += "left join (select TD_CODE,PRE_EDT,REL_EDT from CMETMS.dbo.TDDSDWB where IT_CODE='TK107') k on b.TD_CODE=k.TD_CODE ";
                FSQLStr += string.Format("left join (select TD_CODE,ECC_UPAY_FEE,PRE_SCHD,REL_SCHD,EXECT_DESC,DELAY_RESN from CMETMS.dbo.TDDMSCD where SYR='{0}' and SMN='{1}') m on b.TD_CODE=m.TD_CODE ", SYR, SMN);
                //FSQLStr += "left join (select distinct case AREA when '北區' then 1 when '中區' then 2 else 3 end as AREA,CITY_ID,CITY from CMBASE.dbo.SMCPOST) n on b.TD_CITY=n.CITY ";
                FSQLStr += "left join (select AREA_ID As AREA,CITY,CITY_ID from CMBASE.dbo.BAS_SMCPOST_AREA) n on b.TD_CITY=n.CITY ";
                FSQLStr += "left join (select CITY,case CITY when '不分區' then 1 else 0 end as ORD from CMBASE.dbo.BAS_SMCPOST_AREA) o on b.TD_CITY=o.CITY) z ";
                FSQLStr += "where z.REV_FUND+z.FUND>0 ";
                if (db.SECCITY.Where(m => m.USR_ID == USR_ID).Any())
                {
                    FSQLStr += string.Format("and TD_CITY IN (SELECT CITY FROM CMBASE.dbo.SECCITY where USR_ID = '{0}')", USR_ID);
                }
                else
                {
                    FSQLStr += string.Format("and TD_CITY IN (SELECT CITY FROM CMBASE.dbo.SECCITY where DEPT_NO = '{0}')", DEPT_NO);
                }
                FSQLStr += "order by z.ORD,z.AREA,z.CITY_ID,z.TD_TOWN,z.TD_CODE";
            }
            else
            {
                FSQLStr = "select z.TD_CODE,z.TD_NAME,z.TD_CITY,z.TD_TOWN,z.IPJ_CODE,z.CODE_DESC,z.YR_WORK,z.TIT_CODE,";
                FSQLStr += "z.REV_FUND,z.FUND,z.REV_SMN_FUND,z.SMN_FUND,z.REV_COST_FEE,z.COST_FEE,";
                FSQLStr += "z.REV_PPAY_FEE,z.PPAY_FEE,z.REV_ECC_UPAY_FEE,z.ECC_UPAY_FEE,";
                FSQLStr += "z.PRE_TK102,z.REL_TK102,z.PRE_TK105,z.REL_TK105,z.PRE_TK107,z.REL_TK107,";
                FSQLStr += "z.PRE_SCHD,z.REL_SCHD,z.EXECT_DESC,z.MEMO,z.AREA,case z.ORD when 1 then case when (charindex('業務費',z.TD_NAME)>0) then 1 else 2 end else 0 end as ORD from ";
                FSQLStr += "(select a.TD_CODE,b.TD_NAME,b.TD_CITY,b.TD_TOWN,b.IPJ_CODE,c.CODE_DESC,b.YR_WORK,d.CODE_DESC as TIT_CODE,";
                FSQLStr += "0 as REV_FUND,e.FUND,0 as REV_SMN_FUND,g.SMN_FUND,0 as REV_COST_FEE,isnull(h.COST_FEE,0) as COST_FEE,";
                FSQLStr += "0 as REV_PPAY_FEE,isnull(h.PPAY_FEE,0) as PPAY_FEE,0 as REV_ECC_UPAY_FEE,isnull(m.ECC_UPAY_FEE,0) as ECC_UPAY_FEE,";
                FSQLStr += "isnull(i.PRE_EDT,'') as PRE_TK102,isnull(i.REL_EDT,'') as REL_TK102,";
                FSQLStr += "isnull(j.PRE_EDT,'') as PRE_TK105,isnull(j.REL_EDT,'') as REL_TK105,";
                FSQLStr += "isnull(k.PRE_EDT,'') as PRE_TK107,isnull(k.REL_EDT,'') as REL_TK107,";
                FSQLStr += "isnull(m.PRE_SCHD,0) as PRE_SCHD,isnull(m.REL_SCHD,0) as REL_SCHD,m.EXECT_DESC,m.DELAY_RESN as MEMO,";
                FSQLStr += "n.AREA,n.CITY_ID,o.ORD from ";
                FSQLStr += string.Format("(select distinct TD_CODE,TID,FD_CODE from CMEPMS.dbo.CTDMNDS where PID={0}) a ", PID);

                FSQLStr += "inner join (select TD_CODE,TD_NAME,TD_CITY,TD_TOWN,IPJ_CODE,case isnull(PD_FG,'') when 'D01' then 'P' else CATY_CODE end YR_WORK from CMETMS.dbo.TDMFORM) b on a.TD_CODE=b.TD_CODE ";
                //FSQLStr += "inner join (select TD_CODE,TD_NAME,TD_CITY,IPJ_CODE,case isnull(PD_FG,'') when 'D01' then 'P' else CATY_CODE end YR_WORK from CMETMS.dbo.TDMFORM) b on b.YR_WORK in ('C','L') and a.TD_CODE=b.TD_CODE ";

                FSQLStr += "inner join (select CATY_COD,CODE_DESC from CMETMS.dbo.SMCCODE where SYS_ID='SM' and CATY_ID='02') c on b.IPJ_CODE=c.CATY_COD ";
                FSQLStr += "inner join (select CATY_COD,CODE_DESC from CMETMS.dbo.SMCCODE where SYS_ID='TT' and CATY_ID='04') d on b.YR_WORK=d.CATY_COD ";
                FSQLStr += "inner join (select TID,FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00') e on a.TID=e.TID ";
                FSQLStr += "left join (select f.TID,sum(f.SMN_FUND) as SMN_FUND from ";
                FSQLStr += string.Format("(select TID,'01' as SMN,isnull(M_01,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'02' as SMN,isnull(M_02,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'03' as SMN,isnull(M_03,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'04' as SMN,isnull(M_04,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'05' as SMN,isnull(M_05,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'06' as SMN,isnull(M_06,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'07' as SMN,isnull(M_07,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'08' as SMN,isnull(M_08,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'09' as SMN,isnull(M_09,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'10' as SMN,isnull(M_10,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'11' as SMN,isnull(M_11,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
                FSQLStr += string.Format("union select TID,'12' as SMN,isnull(M_12,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00') f ", SYR);
                FSQLStr += string.Format("where f.SMN<='{0}' group by f.TID) g on a.TID=g.TID ", SMN);
                FSQLStr += string.Format("left join (select FD_CODE,COST_FEE,PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR='{0}' and SMN='{1}') h on a.FD_CODE=h.FD_CODE ", SYR, SMN);
                FSQLStr += "left join (select TD_CODE,PRE_EDT,REL_EDT from CMETMS.dbo.TDDSDWB where IT_CODE='TK103') i on b.TD_CODE=i.TD_CODE ";
                FSQLStr += "left join (select TD_CODE,PRE_EDT,REL_EDT from CMETMS.dbo.TDDSDWB where IT_CODE='TK105') j on b.TD_CODE=j.TD_CODE ";
                FSQLStr += "left join (select TD_CODE,PRE_EDT,REL_EDT from CMETMS.dbo.TDDSDWB where IT_CODE='TK107') k on b.TD_CODE=k.TD_CODE ";
                FSQLStr += string.Format("left join (select TD_CODE,ECC_UPAY_FEE,PRE_SCHD,REL_SCHD,EXECT_DESC,DELAY_RESN from CMETMS.dbo.TDDMSCD where SYR='{0}' and SMN='{1}') m on b.TD_CODE=m.TD_CODE ", SYR, SMN);
                //FSQLStr += "left join (select distinct case AREA when '北區' then 1 when '中區' then 2 else 3 end as AREA,CITY_ID,CITY from CMBASE.dbo.SMCPOST) n on b.TD_CITY=n.CITY ";
                FSQLStr += "left join (select AREA_ID As AREA,CITY,CITY_ID from CMBASE.dbo.BAS_SMCPOST_AREA) n on b.TD_CITY=n.CITY ";
                FSQLStr += "left join (select CITY,case CITY when '不分區' then 1 else 0 end as ORD from CMBASE.dbo.BAS_SMCPOST_AREA) o on b.TD_CITY=o.CITY) z ";

                if (db.SECCITY.Where(m => m.USR_ID == USR_ID).Any())
                {
                    FSQLStr += string.Format("where TD_CITY IN (SELECT CITY FROM CMBASE.dbo.SECCITY where USR_ID = '{0}')", USR_ID);
                }
                else
                {
                    FSQLStr += string.Format("where TD_CITY IN (SELECT CITY FROM CMBASE.dbo.SECCITY where DEPT_NO = '{0}')", DEPT_NO);
                }
                FSQLStr += "and z.REV_FUND+z.FUND>0 ";
                FSQLStr += "order by z.ORD,z.AREA,z.CITY_ID,z.TD_TOWN,z.TD_CODE";

            }
            SqlConnection con = new SqlConnection();
            con.ConnectionString = connectionString;
            SqlDataAdapter adp1 = new SqlDataAdapter(FSQLStr, connectionString);
            DataSet ds1 = new DataSet();
            adp1.Fill(ds1);
            DataTable dt = ds1.Tables[0];
            return dt;
        }

        public static DataTable getBillionProject06Cities(string USR_ID, string DEPT_NO, string PID, string SYR, string SMN, string connectionString)
        {
            string FSQLStr = "";
            FSQLStr = "select b.TD_CITY,sum(c.FUND) as YR_FUND,sum(e.SMN_FUND) as SMN_FUND,sum(isnull(f.COST_FEE,0)+isnull(f.PPAY_FEE,0)) as FEE from ";
            FSQLStr += string.Format("(select TD_CODE,TID,FD_CODE from CMEPMS.dbo.CTDMNDS where PID={0}) a ", PID);
            FSQLStr += "inner join (select TD_CODE,TD_CITY,IPJ_CODE,case isnull(PD_FG,'') when 'D01' then 'P' else CATY_CODE end YR_WORK from CMETMS.dbo.TDMFORM) b on b.YR_WORK in ('C','L') and a.TD_CODE=b.TD_CODE ";
            FSQLStr += "inner join (select TID,FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00') c on a.TID=c.TID ";
            FSQLStr += "left join (select d.TID,sum(d.SMN_FUND) as SMN_FUND from ";
            FSQLStr += string.Format("(select TID,'01' as SMN,isnull(M_01,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'02' as SMN,isnull(M_02,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'03' as SMN,isnull(M_03,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'04' as SMN,isnull(M_04,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'05' as SMN,isnull(M_05,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'06' as SMN,isnull(M_06,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'07' as SMN,isnull(M_07,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'08' as SMN,isnull(M_08,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'09' as SMN,isnull(M_09,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'10' as SMN,isnull(M_10,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'11' as SMN,isnull(M_11,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00' ", SYR);
            FSQLStr += string.Format("union select TID,'12' as SMN,isnull(M_12,0) as SMN_FUND from CMEPMS.dbo.CTDITFD where SYR='{0}' and FIT_CODE='00') d ", SYR);
            FSQLStr += string.Format("where d.SMN<='{0}' group by d.TID) e on a.TID=e.TID ", SMN);
            FSQLStr += string.Format("left join (select FD_CODE,COST_FEE,PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR='{0}' and SMN='{1}') f on a.FD_CODE=f.FD_CODE ", SYR, SMN);

            if (db.SECCITY.Where(m => m.USR_ID == USR_ID).Any())
            {
                FSQLStr += string.Format("WHERE TD_CITY IN (SELECT CITY FROM CMBASE.dbo.SECCITY where USR_ID = '{0}')", USR_ID);
            }
            else
            {
                FSQLStr += string.Format("WHERE TD_CITY IN (SELECT CITY FROM CMBASE.dbo.SECCITY where DEPT_NO = '{0}')", DEPT_NO);
            }
            FSQLStr += "　group by b.TD_CITY ";
            FSQLStr += "order by b.TD_CITY";
            SqlConnection con = new SqlConnection();
            //string constr = @"Data Source=10.1.1.117\SEPMIS;initial catalog=CMEPMS;user id=***;password=***;" + "MultipleActiveResultSets=True;" + "initial catalog=CMETMS;";
            con.ConnectionString = connectionString;
            SqlDataAdapter adp = new SqlDataAdapter(FSQLStr, con);
            DataSet ds = new DataSet();
            adp.Fill(ds);
            DataTable dt = ds.Tables[0];
            return dt;
        }

        public static DataTable getBillionProject08(string PID, string SYR, string SMN, string CITY, string FUND0, string connectionString)
        {

            string FSQLStr = "";
            string PSYR = (Convert.ToInt32(SYR) - 1).ToString("000");

            // 20240910 修改項目 => 查詢語法修改 SQL (IN)
            FSQLStr = string.Format("SELECT MAX(SEQ_NO) as SEQ_NO from CMEPMS.dbo.PJDFDCR WHERE PID IN ({0})", PID);

            SqlConnection con = new SqlConnection();
            con.ConnectionString = connectionString;

            SqlDataAdapter adp = new SqlDataAdapter(FSQLStr, con);
            DataSet ds = new DataSet();
            adp.Fill(ds);
            DataTable dt = ds.Tables[0];

            int seq = 0;
            int i;
            int j;

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["SEQ_NO"].ToString() != "")
                {
                    seq = Convert.ToInt16(dt.Rows[0]["SEQ_NO"].ToString());
                }
            }

            // 20240910 修改項目 => 查詢語法修改 SQL (IN)
            FSQLStr = string.Format("select APROV_DT+APROV_FN as APROV from CMEPMS.dbo.PJDFDCR where PID IN ({0}) and SEQ_NO={1}", PID, seq.ToString());
            SqlDataAdapter adp1 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds1 = new DataSet();
            adp1.Fill(ds1);
            dt = ds1.Tables[0];
            bool aprov = true;

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["APROV"].ToString() == "")
                {
                    aprov = false;
                }
            }

            if (aprov == true)
            {
                Debug.WriteLine("執行了 true 區塊");
                FSQLStr = "select a.CATY_NAME,a.CATY_CODE,a.CATY_FUND,d.CATY_CODE as TD_CATY,S8.CODE_DESC as CATY_CODE_DESC,S9.CODE_DESC as IPJ_DESC";
                FSQLStr = string.Format("{0},b.TD_CODE,b.TD_NAME,b.FD_CODE", FSQLStr);
                FSQLStr = string.Format("{0},c.YR_FUND", FSQLStr);
                FSQLStr = string.Format("{0},d.TD_CITY,d.TD_TOWN,d.TK105_TYPE,d.BUSS_NAME,d.DSG_BUSS_NAME", FSQLStr);
                FSQLStr = string.Format("{0},isnull(d.SAVE_FEE,0) as SAVE_FEE", FSQLStr);
                FSQLStr = string.Format("{0},isnull(m.EXECT_DESC,'') as EXECT_DESC", FSQLStr); /// 新增辦理情形
                FSQLStr = string.Format("{0},isnull(m.PRE_SCHD,0) as PRE_SCHD", FSQLStr);      /// 新增預定進度
                FSQLStr = string.Format("{0},isnull(m.REL_SCHD,0) as REL_SCHD", FSQLStr);      /// 新增實際進度
                FSQLStr = string.Format("{0},S1.PRE_EDT as PRE_TK103,S1.REL_EDT as REL_TK103,S2.PRE_EDT as PRE_TK105,S2.REL_EDT as REL_TK105,S6.PRE_EDT as PRE_TK102,S6.REL_EDT as REL_TK102,S7.PRE_EDT as PRE_TK1031,S7.REL_EDT as REL_TK1031", FSQLStr);
                //FSQLStr = string.Format("{0},case d.CATY_CODE when 'P' then S3.PRE_EDT else S4.PRE_EDT end as PRE_TK107", FSQLStr);
                //FSQLStr = string.Format("{0},case d.CATY_CODE when 'P' then S3.REL_EDT else S4.REL_EDT end as REL_TK107", FSQLStr);
                FSQLStr = string.Format("{0},case when d.CATY_CODE = 'P' and S3.PRE_EDT is not null then S3.PRE_EDT else S4.PRE_EDT end as PRE_TK107", FSQLStr);
                FSQLStr = string.Format("{0},case when d.CATY_CODE = 'P' and S3.REL_EDT is not null then S3.REL_EDT else S4.REL_EDT end as REL_TK107", FSQLStr);
                FSQLStr = string.Format("{0},S10.PRE_EDT as PRE_TK108,S10.REL_EDT as REL_TK108", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) from CMETMS.dbo.TDDBGLS where TD_CODE=b.TD_CODE and REC_TP='C' and SEQ_NO=(select max(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=b.TD_CODE and REC_TP='C')) as BID_FEE", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) from CMETMS.dbo.TDDBGLS where TD_CODE=b.TD_CODE and REC_TP='B' and SEQ_NO=(select max(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=b.TD_CODE and REC_TP='B')) as BG_FEE", FSQLStr);
                FSQLStr = string.Format("{0},e.AREA,e.CITY_ID", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B')) as B_FUND", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE='0101') as B_0101_FUND", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE in ('0114','0199')) as B_0114_FUND", FSQLStr);//間接費用=設備費-間接費用 + 設備費-準備金
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE like '07%') as B_FUND_07", FSQLStr);//用戶接管總經費
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE='0701') as B_0701_FUND", FSQLStr);//用戶接管工程費
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE in ('0702','0703')) as B_0702_FUND", FSQLStr);//用戶接管間接費 = 用戶接管費-間接費用 + 用戶接管費-準備金
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C')) as C_FUND", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C') and AC_ITEM_CODE='0101') C_0101_FUND", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C') and AC_ITEM_CODE like '07%') as C_FUND_07", FSQLStr);//用戶接管總經費
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C') and AC_ITEM_CODE='0701') as C_0701_FUND", FSQLStr);//用戶接管工程費
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C') and AC_ITEM_CODE in ('0702','0703')) as C_0702_FUND", FSQLStr);//用戶接管間接費
                FSQLStr = string.Format("{0} from CMEPMS.dbo.PJDCATY a", FSQLStr);
                FSQLStr = string.Format("{0} inner join CMEPMS.dbo.CTDMNDS b on a.PID=b.PID and a.CATY_CODE=b.CATY_CODE and isnull(b.NOCTRL_FG,0)=0", FSQLStr);
                FSQLStr = string.Format("{0} inner join (select TID,sum(FUND) as YR_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE like '0%' group by TID) c on b.TID=c.TID", FSQLStr);

                if (CITY == "請選擇" || CITY == "") { FSQLStr = string.Format("{0} inner join CMETMS.dbo.TDMFORM d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "新北市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('新北市','台北縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "台中市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('台中市','台中縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "台南市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('台南市','台南縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "高雄市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('高雄市','高雄縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "桃園市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('桃園市','桃園縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY='{1}') d on b.TD_CODE=d.TD_CODE", FSQLStr, CITY); }

                FSQLStr = string.Format("{0} inner join (select AREA_ID As AREA,CITY,CITY_ID from CMBASE.dbo.BAS_SMCPOST_AREA) e on d.TD_CITY=e.CITY", FSQLStr);
                FSQLStr = string.Format("{0} left join CMETMS.dbo.TDDSDWB S1 on S1.IT_CODE='TK103' and S1.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} left join CMETMS.dbo.TDDSDWB S2 on S2.IT_CODE='TK105' and S2.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} left join CMETMS.dbo.TDDSDWB S3 on S3.IT_CODE='DG106' and S3.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} left join CMETMS.dbo.TDDSDWB S4 on S4.IT_CODE='TK107' and S4.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} left join CMETMS.dbo.TDDSDWB S6 on S6.IT_CODE='TK102' and S6.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} left join CMETMS.dbo.TDDSDWB S7 on S7.IT_CODE='TK1031' and S7.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} left join CMETMS.dbo.TDDSDWB S10 on S10.IT_CODE='TK108' and S10.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} left join CMETMS.dbo.SMCCODE S8 on S8.SYS_ID='TT' and S8.CATY_ID='03' and S8.CATY_COD=d.CATY_CODE", FSQLStr);
                FSQLStr = string.Format("{0} left join CMETMS.dbo.SMCCODE S9 on S9.SYS_ID ='SM' and S9.CATY_ID = '02'  and S9.CATY_COD = d.IPJ_CODE", FSQLStr);
                FSQLStr = string.Format("{0} left join (select TD_CODE, EXECT_DESC, PRE_SCHD, REL_SCHD from CMETMS.dbo.TDDMSCD where SYR = '{1}' and SMN = '{2}') m on m.TD_CODE = b.TD_CODE", FSQLStr, SYR, SMN);

                // 20240910 修改項目 => 查詢語法修改 SQL (IN)
                FSQLStr = string.Format("{0} where a.PID IN ({1})", FSQLStr, PID);

                // 1071113 修改項目 => 修改回原本的資料
                FSQLStr = string.Format("{0} group by a.CATY_NAME,a.CATY_CODE,a.CATY_FUND,S8.CODE_DESC,S9.CODE_DESC", FSQLStr);
                FSQLStr = string.Format("{0},b.TD_CODE,b.TD_NAME,b.FD_CODE", FSQLStr);
                FSQLStr = string.Format("{0},c.YR_FUND", FSQLStr);
                FSQLStr = string.Format("{0},d.TD_CITY,d.TD_TOWN,d.CATY_CODE,d.TK105_TYPE,d.BUSS_NAME,d.DSG_BUSS_NAME", FSQLStr);
                FSQLStr = string.Format("{0},isnull(d.SAVE_FEE,0)", FSQLStr);
                FSQLStr = string.Format("{0},e.AREA,e.CITY_ID,S1.PRE_EDT,S1.REL_EDT,S2.PRE_EDT,S2.REL_EDT,S3.PRE_EDT,S3.REL_EDT,S4.PRE_EDT,S4.REL_EDT,S6.PRE_EDT,S6.REL_EDT,S7.PRE_EDT,S7.REL_EDT,S10.PRE_EDT,S10.REL_EDT,m.EXECT_DESC,m.PRE_SCHD,m.REL_SCHD", FSQLStr);
                FSQLStr = string.Format("{0} order by e.AREA,substring(b.FD_CODE,12,4),e.CITY_ID,d.TD_CITY,a.CATY_CODE,d.TD_TOWN,substring(b.TD_CODE,len(b.TD_CODE)-3,4)", FSQLStr);
            }
            else
            {
                Debug.WriteLine("執行了 false 區塊");
                FSQLStr = "select a.CATY_NAME,a.CATY_CODE,a.CATY_FUND,S8.CODE_DESC as CATY_CODE_DESC,S9.CODE_DESC as IPJ_DESC";
                FSQLStr = string.Format("{0},b.TD_CODE,b.TD_NAME,b.FD_CODE", FSQLStr);
                FSQLStr = string.Format("{0},c.YR_FUND", FSQLStr);
                FSQLStr = string.Format("{0},d.TD_CITY,d.TD_TOWN,d.TK105_TYPE,d.BUSS_NAME,d.DSG_BUSS_NAME", FSQLStr);
                // 1071113 修改項目 => 修改回原本的資料
                FSQLStr = string.Format("{0},isnull(d.SAVE_FEE,0) as SAVE_FEE", FSQLStr);
                FSQLStr = string.Format("{0},ISNULL(m.EXECT_DESC, '') AS EXECT_DESC", FSQLStr); /// 新增辦理情形
                FSQLStr = string.Format("{0},ISNULL(m.PRE_SCHD, 0) AS PRE_SCHD", FSQLStr);      /// 新增預定進度
                FSQLStr = string.Format("{0},ISNULL(m.REL_SCHD, 0) AS REL_SCHD", FSQLStr);      /// 新增實際進度
                FSQLStr = string.Format("{0},S1.PRE_EDT as PRE_TK103,S1.REL_EDT as REL_TK103,S2.PRE_EDT as PRE_TK105,S2.REL_EDT as REL_TK105,S6.PRE_EDT as PRE_TK102,S6.REL_EDT as REL_TK102,S7.PRE_EDT as PRE_TK1031,S7.REL_EDT as REL_TK1031", FSQLStr);
                //FSQLStr = string.Format("{0},case d.CATY_CODE when 'P' then S3.PRE_EDT else S4.PRE_EDT end as PRE_TK107", FSQLStr);
                //FSQLStr = string.Format("{0},case d.CATY_CODE when 'P' then S3.REL_EDT else S4.REL_EDT end as REL_TK107", FSQLStr);
                FSQLStr = string.Format("{0},case when d.CATY_CODE = 'P' and S3.PRE_EDT is not null then S3.PRE_EDT else S4.PRE_EDT end as PRE_TK107", FSQLStr);
                FSQLStr = string.Format("{0},case when d.CATY_CODE = 'P' and S3.REL_EDT is not null then S3.REL_EDT else S4.REL_EDT end as REL_TK107", FSQLStr);
                FSQLStr = string.Format("{0},S10.PRE_EDT as PRE_TK108,S10.REL_EDT as REL_TK108", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) from CMETMS.dbo.TDDBGLS where TD_CODE=b.TD_CODE and REC_TP='C' and SEQ_NO=(select max(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=b.TD_CODE and REC_TP='C')) as BID_FEE", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) from CMETMS.dbo.TDDBGLS where TD_CODE=b.TD_CODE and REC_TP='B' and SEQ_NO=(select max(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=b.TD_CODE and REC_TP='B')) as BG_FEE", FSQLStr);
                FSQLStr = string.Format("{0},e.AREA,e.CITY_ID", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B')) as B_FUND", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE='0101') as B_0101_FUND", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE in ('0114','0199')) as B_0114_FUND", FSQLStr);//間接費用=設備費-間接費用 + 設備費-準備金
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE like '07%') as B_FUND_07", FSQLStr);//用戶接管總經費
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE='0701') as B_0701_FUND", FSQLStr);//用戶接管工程費
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='B' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='B') and AC_ITEM_CODE in ('0702','0703')) as B_0702_FUND", FSQLStr);//用戶接管間接費 = 用戶接管費-間接費用 + 用戶接管費-準備金
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C')) as C_FUND", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C') and AC_ITEM_CODE='0101') C_0101_FUND", FSQLStr);
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C') and AC_ITEM_CODE like '07%') as C_FUND_07", FSQLStr);//用戶接管總經費
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C') and AC_ITEM_CODE='0701') as C_0701_FUND", FSQLStr);//用戶接管工程費
                FSQLStr = string.Format("{0},(select sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP='C' and TD_CODE=b.TD_CODE and SEQ_NO=(select MAX(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=CMETMS.dbo.TDDBGLS.TD_CODE and REC_TP='C') and AC_ITEM_CODE in ('0702','0703')) as C_0702_FUND", FSQLStr);//用戶接管間接費
                FSQLStr = string.Format("{0} from CMEPMS.dbo.PJDCATY a", FSQLStr);
                FSQLStr = string.Format("{0} inner join CMEPMS.dbo.CTDMNDS b on a.PID=b.PID and a.CATY_CODE=b.CATY_CODE and isnull(b.NOCTRL_FG,0)=0", FSQLStr);
                FSQLStr = string.Format("{0} inner join (select TID,sum(FUND) as YR_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE like '0%' group by TID) c on b.TID=c.TID", FSQLStr);

                if (CITY == "請選擇" || CITY == "") { FSQLStr = string.Format("{0} inner join CMETMS.dbo.TDMFORM d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "新北市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('新北市','台北縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "台中市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('台中市','台中縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "台南市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('台南市','台南縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "高雄市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('高雄市','高雄縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else if (CITY == "桃園市") { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY in ('桃園市','桃園縣')) d on b.TD_CODE=d.TD_CODE", FSQLStr); }
                else { FSQLStr = string.Format("{0} inner join (select * from CMETMS.dbo.TDMFORM where TD_CITY='{1}') d on b.TD_CODE=d.TD_CODE", FSQLStr, CITY); }

                FSQLStr = string.Format("{0} inner join (select AREA_ID As AREA,CITY,CITY_ID from CMBASE.dbo.BAS_SMCPOST_AREA) e on d.TD_CITY=e.CITY", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN CMETMS.dbo.TDDSDWB S1 on S1.IT_CODE='TK103' and S1.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN CMETMS.dbo.TDDSDWB S2 on S2.IT_CODE='TK105' and S2.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN CMETMS.dbo.TDDSDWB S3 on S3.IT_CODE='DG106' and S3.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN CMETMS.dbo.TDDSDWB S4 on S4.IT_CODE='TK107' and S4.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN CMETMS.dbo.TDDSDWB S6 on S6.IT_CODE='TK102' and S6.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN CMETMS.dbo.TDDSDWB S7 on S7.IT_CODE='TK1031' and S7.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN CMETMS.dbo.TDDSDWB S10 on S10.IT_CODE='TK108' and S10.TD_CODE=b.TD_CODE", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN CMETMS.dbo.SMCCODE S8 on S8.SYS_ID='TT' and S8.CATY_ID='03' and S8.CATY_COD=d.CATY_CODE", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN CMETMS.dbo.SMCCODE S9 on S9.SYS_ID ='SM' and S9.CATY_ID = '02'  and S9.CATY_COD = d.IPJ_CODE", FSQLStr);
                FSQLStr = string.Format("{0} LEFT JOIN (SELECT TD_CODE, EXECT_DESC, PRE_SCHD, REL_SCHD from CMETMS.dbo.TDDMSCD where SYR = '{1}' and SMN = '{2}') m on m.TD_CODE = b.TD_CODE", FSQLStr, SYR, SMN);

                // 20240910 修改項目 => 查詢語法修改 SQL (IN)
                FSQLStr = string.Format("{0} where a.PID IN ({1})", FSQLStr, PID);

                // 1071113 修正 ==> 修改回原本的資料
                FSQLStr = string.Format("{0} group by a.CATY_NAME,a.CATY_CODE,a.CATY_FUND,S8.CODE_DESC,S9.CODE_DESC", FSQLStr);
                FSQLStr = string.Format("{0},b.TD_CODE,b.TD_NAME,b.FD_CODE", FSQLStr);
                FSQLStr = string.Format("{0},c.YR_FUND", FSQLStr);
                FSQLStr = string.Format("{0},d.TD_CITY,d.TD_TOWN,d.CATY_CODE,d.TK105_TYPE,d.BUSS_NAME,d.DSG_BUSS_NAME", FSQLStr);
                FSQLStr = string.Format("{0},isnull(d.SAVE_FEE,0)", FSQLStr);
                FSQLStr = string.Format("{0},e.AREA,e.CITY_ID,S1.PRE_EDT,S1.REL_EDT,S2.PRE_EDT,S2.REL_EDT,S3.PRE_EDT,S3.REL_EDT,S4.PRE_EDT,S4.REL_EDT,S6.PRE_EDT,S6.REL_EDT,S7.PRE_EDT,S7.REL_EDT,S10.PRE_EDT,S10.REL_EDT,m.EXECT_DESC,m.PRE_SCHD,m.REL_SCHD", FSQLStr);
                FSQLStr = string.Format("{0} order by e.AREA,substring(b.FD_CODE,12,4),e.CITY_ID,d.TD_CITY,a.CATY_CODE,d.TD_TOWN,substring(b.TD_CODE,len(b.TD_CODE)-3,4)", FSQLStr);
            }

            Debug.WriteLine("區塊執行結果");
            Debug.WriteLine($"SQL => {FSQLStr}");

            SqlDataAdapter adp2 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds2 = new DataSet();
            adp2.Fill(ds2);
            dt = ds2.Tables[0];

            // 1070410 修正項目 => 判斷是否包含經費歸零案
            if (FUND0 == "Y")
            {
                // 包含經費歸零案
            }
            else
            {
                // 不包含經費歸零案
                for (i = dt.Rows.Count - 1; i > -1; i--)
                {
                    if (dt.Rows[i]["YR_FUND"].ToString() == "" || Convert.ToDouble(dt.Rows[i]["YR_FUND"].ToString()) == 0)
                    {
                        // 1061031 修正 經費歸零案也要列出 ==> 交辦人2894
                        dt.Rows.Remove(dt.Rows[i]);
                    }
                }
            }

            dt.Columns.Add("ALL_PPAY_FEE");// 歷年累積撥款金額==>跨年度累計金額
            dt.Columns.Add("PPAY_FEE");    // 撥款金額==>當年度累計金額
            dt.Columns.Add("L_PLN_COST");  // 中央款累積分配數==>1071106來源改為TDDPCOS
            dt.Columns.Add("L_REL_COST");  // 中央款實際支用數==>1071106來源改為TDDPCOS
            dt.Columns.Add("L_UNPAY_COST");// 中央款應付未付數==>1071106來源改為TDDPCOS
            //dt.Columns.Add("PRE_SCHD");
            //dt.Columns.Add("REL_SCHD");
            //dt.Columns.Add("EXECT_DESC");
            dt.Columns.Add("DELAY_RESN");
            dt.Columns.Add("CTDELAY_RESN");

            // 20111017 執行單位主管要印在備註欄位(署辦、署代辦)
            FSQLStr = "select";
            FSQLStr += " b.TD_CODE,isnull(h.USR_NAME,'') as USR_NAME";
            FSQLStr += " from";
            // 20240910 修改項目 => 查詢語法修改 SQL (IN)
            FSQLStr += string.Format(" (select PID from CMEPMS.dbo.PJMFORM where PID IN ({0})) a", PID);
            FSQLStr += " inner join (select PID,TD_CODE from CMEPMS.dbo.CTDMNDS) b on a.PID=b.PID";
            FSQLStr += " inner join (select TD_CODE,IPJ_CODE from CMETMS.dbo.TDMFORM where IPJ_CODE in ('01','03')) c on b.TD_CODE=c.TD_CODE";
            FSQLStr += " left join (select TD_CODE,ADM_CODE from CPAPMS.CMETMS.dbo.TDMFORM) d on b.TD_CODE=d.ADM_CODE";
            FSQLStr += " left join (select KEY_CODE,ADV_DP from CPAPMS.CMBASE.dbo.FLWDATA where DATA_TYPE='TDMFORM') e on d.TD_CODE=e.KEY_CODE";
            FSQLStr += " left join (select KEY_CODE,ADV_DP from CPAPMS.CMBASE.dbo.FLWDATA where DATA_TYPE='DAMFORM') f on d.TD_CODE=f.KEY_CODE";
            FSQLStr += " left join (select DEPT_NO,DEPT_NAME,DEPT_CHIEF from CPAPMS.CMBASE.dbo.OGCDEPT) g on isnull(e.ADV_DP,f.ADV_DP)=g.DEPT_NO";
            FSQLStr += " left join (select USR_ID,USR_NAME from CPAPMS.CMBASE.dbo.OGCUSER) h on g.DEPT_CHIEF=h.USR_ID";

            SqlDataAdapter adp3 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds3 = new DataSet();
            adp3.Fill(ds3);
            DataTable dtTemp = ds3.Tables[0];

            for (i = 0; i < dtTemp.Rows.Count; i++)
            {
                if (dtTemp.Rows[i]["USR_NAME"].ToString() != "")
                {
                    for (j = 0; j < dt.Rows.Count; j++)
                    {
                        if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                        {
                            dt.Rows[j]["CTDELAY_RESN"] = dtTemp.Rows[i]["USR_NAME"].ToString();
                        }
                    }
                }
            }

            // 20140321 執行單位主管要印基本資料的本署監造督導管理人員欄位
            FSQLStr = "Select";
            FSQLStr += " b.TD_CODE,c.USR_NAME";
            FSQLStr += " from";
            // 20240910 修改項目 => 查詢語法修改 SQL (IN)
            FSQLStr += string.Format(" (select PID from CMEPMS.dbo.PJMFORM where PID IN ({0})) a", PID);
            FSQLStr += " inner join (select PID,TD_CODE from CMEPMS.dbo.CTDMNDS) b on a.PID=b.PID";
            FSQLStr += " inner join (select TD_CODE,IPJ_CODE,ADV_CHIEF as USR_NAME from CMETMS.dbo.TDMFORM where IPJ_CODE in ('02')) c on b.TD_CODE=c.TD_CODE";

            SqlDataAdapter adp4 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds4 = new DataSet();
            adp4.Fill(ds4);
            dtTemp = ds4.Tables[0];

            for (i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["CTDELAY_RESN"] = dtTemp.Rows[i]["USR_NAME"].ToString();
                    }
                }
            }

            for (i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["ALL_PPAY_FEE"] = 0;
                dt.Rows[i]["PPAY_FEE"] = 0;
                dt.Rows[i]["L_PLN_COST"] = 0;
                dt.Rows[i]["L_REL_COST"] = 0;
                dt.Rows[i]["L_UNPAY_COST"] = 0;
            }

            FSQLStr = "select C.TD_CODE";
            FSQLStr += ",case when datepart(d,getdate())>(select convert(int,DESC_COD) from CMEPMS.dbo.SMCCODE where SYS_ID='SC' and CATY_ID='01' and CATY_COD='ACDT') then";
            FSQLStr += string.Format(" isnull((select top 1 (isnull(COST_FEE,0)+isnull(PPAY_FEE,0)) from TDDSTCF where FD_CODE=C.FD_CODE and SYR+SMN='{0}{1}'),0) - isnull((select top 1 (isnull(COST_FEE,0)+isnull(PPAY_FEE,0)) from TDDSTCF where FD_CODE=C.FD_CODE and SYR+SMN='{2}'),0)", SYR, SMN, PSYR + "12");
            FSQLStr += string.Format(" else isnull((select top 1 (isnull(COST_FEE,0)+isnull(PPAY_FEE,0)) from TDDSTCF where FD_CODE=C.FD_CODE and SYR+SMN='{0}{1}'),0)", SYR, SMN);
            FSQLStr += " end PPAY_FEE";
            // 20240910 修改項目 => 查詢語法修改 SQL (IN) <前一年度以前歷年(實支數+保留數)>
            FSQLStr += string.Format(",isnull((select sum(isnull(COST_FEE,0)+isnull(PPAY_FEE,0)) from TDDSTCF where  TD_CODE1=C.TD_CODE and FD_CODE in (select FD_CODE from CMEPMS.dbo.CTDMNDS where PID in (select PID from CMEPMS.dbo.PJMFORM where PJ_NAME in (select PJ_NAME from CMEPMS.dbo.PJMFORM where PID IN (" + PID + ")))) and SYR<='{0}' and SMN='12'),0) as ALL_PPAY_FEE", PSYR, SMN); //前一年度以前歷年(實支數+保留數)
            FSQLStr += " from CMEPMS.dbo.CTDMNDS C,CMETMS.dbo.TDMFORM T,CMEPMS.dbo.PJMFORM P";
            // 20240910 修改項目 => 查詢語法修改 SQL (IN)
            FSQLStr = string.Format("{0} where C.PID IN ({1}) and P.SYR='{2}' and C.TD_CODE=T.TD_CODE and C.PID=P.PID", FSQLStr, PID, SYR);
            // 20240910 修改項目 => 查詢語法修改 SQL (IN)
            FSQLStr += " and C.CATY_CODE in (select CATY_CODE from CMEPMS.dbo.PJDCATY where PID IN (" + PID + "))";

            SqlDataAdapter adp5 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds5 = new DataSet();
            adp5.Fill(ds5);
            dtTemp = ds5.Tables[0];
            for (i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["PPAY_FEE"] = dtTemp.Rows[i]["PPAY_FEE"].ToString();
                        dt.Rows[j]["ALL_PPAY_FEE"] = (Convert.ToDouble(dtTemp.Rows[i]["ALL_PPAY_FEE"].ToString()) + Convert.ToDouble(dtTemp.Rows[i]["PPAY_FEE"].ToString())).ToString();//歷年累積撥發金額 = 歷年(實支數+保留數) = 前一年度以前歷年(實支數+保留數) + 今年(實支數+保留數)
                    }
                }
            }

            // 1071113修正 => 讀取TDDMSCD的值 ==> 交辦人2894
            // 20240910 修改項目 => 查詢語法修改 SQL (IN)
            FSQLStr = "select TD_CODE,isnull(L_PLN_COST,0) as L_PLN_COST,isnull(L_REL_COST,0) as L_REL_COST,isnull(L_UNPAY_COST,0) as L_UNPAY_COST from CMETMS.dbo.TDDMSCD where SYR='" + SYR + "' and SMN='" + SMN + "' and TD_CODE in (select TD_CODE from CMEPMS.dbo.CTDMNDS where PID IN (" + PID + "))";
            SqlDataAdapter adp6 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds6 = new DataSet();
            adp6.Fill(ds6);
            dtTemp = ds6.Tables[0];

            for (i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["L_PLN_COST"] = dtTemp.Rows[i]["L_PLN_COST"].ToString();
                        dt.Rows[j]["L_REL_COST"] = dtTemp.Rows[i]["L_REL_COST"].ToString();
                        dt.Rows[j]["L_UNPAY_COST"] = dtTemp.Rows[i]["L_UNPAY_COST"].ToString();
                    }
                }
            }

            FSQLStr = "select b.TD_CODE,c.PRE_SCHD,c.REL_SCHD,c.EXECT_DESC,c.DELAY_RESN,c.CTDELAY_RESN";
            FSQLStr = string.Format("{0} from CMEPMS.dbo.PJDCATY a,CMEPMS.dbo.CTDMNDS b,CMETMS.dbo.TDDMSCD c", FSQLStr);
            // 20240910 修改項目 => 查詢語法修改 SQL (IN)
            FSQLStr = string.Format("{0} where a.PID IN ({1}) and a.PID=b.PID and a.CATY_CODE=b.CATY_CODE and b.TD_CODE=c.TD_CODE and c.SYR='{2}' and c.SMN='{3}'", FSQLStr, PID, SYR, SMN);

            SqlDataAdapter adp7 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds7 = new DataSet();
            adp7.Fill(ds7);
            dtTemp = ds7.Tables[0];

            for (i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        // dt.Rows[j]["PRE_SCHD"] = dtTemp.Rows[i]["PRE_SCHD"].ToString();
                        dt.Rows[j]["PRE_SCHD"] = decimal.TryParse(dtTemp.Rows[i]["PRE_SCHD"].ToString(), out var preSchdValue) ? preSchdValue : 0;
                        // dt.Rows[j]["REL_SCHD"] = dtTemp.Rows[i]["REL_SCHD"].ToString();
                        dt.Rows[j]["REL_SCHD"] = decimal.TryParse(dtTemp.Rows[i]["REL_SCHD"].ToString(), out var relSchdValue) ? relSchdValue : 0;
                        dt.Rows[j]["EXECT_DESC"] = dtTemp.Rows[i]["EXECT_DESC"].ToString();
                        dt.Rows[j]["DELAY_RESN"] = dtTemp.Rows[i]["DELAY_RESN"].ToString();
                        if (dt.Rows[j]["CTDELAY_RESN"].ToString() != "")
                        {
                            dt.Rows[j]["CTDELAY_RESN"] += "\n";
                        }
                        dt.Rows[j]["CTDELAY_RESN"] += dtTemp.Rows[i]["CTDELAY_RESN"].ToString();
                    }
                }
            }

            string d1 = (Convert.ToInt16(DateTime.Now.Year) - 1911).ToString("000")
                        + Convert.ToInt16(DateTime.Now.Month).ToString("00")
                        + Convert.ToInt16(DateTime.Now.Day).ToString("00");

            string d2 = SYR + SMN + Convert.ToInt16(DateTime.Now.Day).ToString("00");

            if (Convert.ToInt32(d1) < Convert.ToInt32(d2))
            {
                // 20240910 修改項目 => 查詢語法修改 SQL (IN)
                FSQLStr = string.Format("select TD_CODE,sum(isnull(IT_WGT,0)) as PRE_SCHD from CMETMS.dbo.TDDSDWB where TD_CODE in (select a.TD_CODE from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDDSDWB b where a.PID IN ({0}) and a.TD_CODE=b.TD_CODE and b.IT_CODE='TK105' and isnull(b.REL_EDT,'')='') and PRE_EDT is not null and PRE_EDT<>'' and substring(PRE_EDT,1,5)<='{1}' group by TD_CODE", PID, d1);
            }
            else
            {
                // 20240910 修改項目 => 查詢語法修改 SQL (IN)
                FSQLStr = string.Format("select TD_CODE,sum(isnull(IT_WGT,0)) as PRE_SCHD from CMETMS.dbo.TDDSDWB where TD_CODE in (select a.TD_CODE from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDDSDWB b where a.PID IN ({0}) and a.TD_CODE=b.TD_CODE and b.IT_CODE='TK105' and isnull(b.REL_EDT,'')='') and PRE_EDT is not null and PRE_EDT<>'' and substring(PRE_EDT,1,5)<='{1}' group by TD_CODE", PID, d2);
            }

            SqlDataAdapter adp8 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds8 = new DataSet();
            adp8.Fill(ds8);
            dtTemp = ds8.Tables[0];

            for (i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["PRE_SCHD"] = dtTemp.Rows[i]["PRE_SCHD"].ToString();
                    }
                }
            }

            // 20240910 修改項目 => 查詢語法修改 SQL (IN)
            FSQLStr = string.Format("select TD_CODE,sum(isnull(IT_WGT,0)) as REL_SCHD from CMETMS.dbo.TDDSDWB where TD_CODE in (select a.TD_CODE from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDDSDWB b where a.PID IN ({0}) and a.TD_CODE=b.TD_CODE and b.IT_CODE='TK105' and isnull(b.REL_EDT,'')='') and REL_EDT is not null and REL_EDT<>'' group by TD_CODE", PID);
            SqlDataAdapter adp9 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds9 = new DataSet();
            adp9.Fill(ds9);
            dtTemp = ds9.Tables[0];

            for (i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["REL_SCHD"] = dtTemp.Rows[i]["REL_SCHD"].ToString();
                    }
                }
            }

            dt.Columns.Add("SEQ");
            dt.Columns.Add("CATY_FUND_");
            string catyName = "";
            double catyFund = 0;
            string tdCity = "";
            int index = -1;

            for (i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["SEQ"] = i + 1;
                if (catyName == dt.Rows[i]["CATY_NAME"].ToString())
                {
                    if (dt.Rows[i]["YR_FUND"].ToString() != "")
                    {
                        catyFund = catyFund + Convert.ToDouble(dt.Rows[i]["YR_FUND"].ToString());
                    }
                }
                else
                {
                    if (index > -1)
                    {
                        dt.Rows[index]["CATY_FUND_"] = catyFund;
                    }
                    if (dt.Rows[i]["YR_FUND"].ToString() != "")
                    {
                        catyFund = Convert.ToDouble(dt.Rows[i]["YR_FUND"].ToString());
                    }
                    else
                    {
                        catyFund = 0;
                    }
                    index = i;
                    catyName = dt.Rows[i]["CATY_NAME"].ToString();
                }

                if (tdCity == dt.Rows[i]["TD_CITY"].ToString() || (tdCity == "新北市" && dt.Rows[i]["TD_CITY"].ToString() == "台北縣") || (tdCity == "台北縣" && dt.Rows[i]["TD_CITY"].ToString() == "新北市"))
                {

                }
                else
                {
                    tdCity = dt.Rows[i]["TD_CITY"].ToString();
                }
            }

            if (index > -1)
            {
                dt.Rows[index]["CATY_FUND_"] = catyFund;
            }
            return dt;
        }

        public static DataTable getBillionProject05(string DEPT_ID, string USR_ID, string PID, string[] aryPid, string SYR, string SMN, string connectionString)
        {
            string FSQLStr;
            string sPids = "";
            for (int i = 0; i < aryPid.Length; i++)
            {
                if (sPids == "")
                {
                    sPids = aryPid[i];
                }
                else
                {
                    sPids += "," + aryPid[i];
                }
            }

            // 判斷 列印日期是否已超過會計結帳日 (smccode:SYS_ID  ='SC' and CATY_ID='01' and CATY_COD='ACDT')
            // 如果在結帳日之前的話，則抓前一個月的資料
            FSQLStr = string.Format("select max(SEQ_NO) as SEQ_NO from PJDFDCR where PID={0}", PID);
            ////////////////////////////////////////////這邊到時候需要確認是否需要上述的效果///////////////////////////////////////////////////
            SqlConnection con = new SqlConnection();
            //string constr = @"Data Source=10.1.1.117\SEPMIS;initial catalog=CMEPMS;user id=***;password=***;" + "MultipleActiveResultSets=True;" + "initial catalog=CMETMS;";
            con.ConnectionString = connectionString;

            FSQLStr = "select d.PID,d.SYR,e.TD_CODE,f.TD_CITY,g.AREA,g.CITY_ID,h.CATY_NAME,h.CATY_FUND,d.TD_CODE,f.TD_NAME,f.TD_CITY,isnull(f.SAVE_FEE,0) as SAVE_FEE";
            FSQLStr += ",(select sum(FUND) from CMEPMS.dbo.CTDITFD where TID=d.TID and FIT_CODE like '0%') as YR_FUND";
            FSQLStr += ",(select PRE_EDT from CMETMS.dbo.TDDSDWB where TD_CODE=d.TD_CODE and IT_CODE='TK103') as PRE_TK103";
            FSQLStr += ",(select REL_EDT from CMETMS.dbo.TDDSDWB where TD_CODE=d.TD_CODE and IT_CODE='TK103') as REL_TK103";
            FSQLStr += ",(select PRE_EDT from CMETMS.dbo.TDDSDWB where TD_CODE=d.TD_CODE and IT_CODE='TK105') as PRE_TK105";
            FSQLStr += ",(select REL_EDT from CMETMS.dbo.TDDSDWB where TD_CODE=d.TD_CODE and IT_CODE='TK105') as REL_TK105";
            FSQLStr += ",(select PRE_EDT from CMETMS.dbo.TDDSDWB where TD_CODE=d.TD_CODE and IT_CODE='TK107') as PRE_TK107";
            FSQLStr += ",(select REL_EDT from CMETMS.dbo.TDDSDWB where TD_CODE=d.TD_CODE and IT_CODE='TK107') as REL_TK107";
            FSQLStr += ",(select PRE_EDT from CMETMS.dbo.TDDSDWB where TD_CODE=d.TD_CODE and IT_CODE='TK108') as PRE_TK_108";
            FSQLStr += ",(select REL_EDT from CMETMS.dbo.TDDSDWB where TD_CODE=d.TD_CODE and IT_CODE='TK108') as REL_TK_108";
            FSQLStr += ",(select sum(BUDGET_FEE) from CMETMS.dbo.TDDBGLS where TD_CODE=d.TD_CODE and REC_TP='C' and SEQ_NO=(select max(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=d.TD_CODE and REC_TP='C')) as BID_FEE";
            FSQLStr += ",(select sum(BUDGET_FEE) from CMETMS.dbo.TDDBGLS where TD_CODE=d.TD_CODE and REC_TP='B' and SEQ_NO=(select max(SEQ_NO) from CMETMS.dbo.TDDBGRC where TD_CODE=d.TD_CODE and REC_TP='B')) as BG_FEE";
            FSQLStr += " from";
            FSQLStr += string.Format(" (select a.PID,a.SYR,b.TD_CODE,b.CATY_CODE,b.TID from CMEPMS.dbo.PJMFORM a,CMEPMS.dbo.CTDMNDS b where a.PID in ({0}) and a.PID=b.PID and isnull(b.NOCTRL_FG,0)<>1) d,", sPids);
            FSQLStr += string.Format("(select max(c.SYR) as SYR,c.TD_CODE from (select a.SYR,b.TD_CODE from CMEPMS.dbo.PJMFORM a,CMEPMS.dbo.CTDMNDS b where a.PID in ({0}) and a.PID=b.PID and isnull(b.NOCTRL_FG,0)<>1) c group by c.TD_CODE) e,", sPids);
            FSQLStr += "CMETMS.dbo.TDMFORM f,";
            //FSQLStr += "(select AREA_ID As AREA,CITY,CITY_ID from CMBASE.dbo.BAS_SMCPOST_AREA) g,";
            if (db.SECCITY.Where(m => m.USR_ID == USR_ID).Count() > 0)
            {
                FSQLStr += string.Format("(SELECT A.AREA_ID as AREA, A.CITY_ID, A.CITY FROM CMBASE.dbo.BAS_SMCPOST_AREA A WHERE EXISTS (SELECT 1 FROM CMBASE.dbo.SECCITY S WHERE A.CITY = S.CITY AND S.USR_ID = '{0}')) g,", USR_ID);
            }
            else
            {
                FSQLStr += string.Format("(SELECT A.AREA_ID as AREA, A.CITY_ID, A.CITY FROM CMBASE.dbo.BAS_SMCPOST_AREA A WHERE EXISTS (SELECT 1 FROM CMBASE.dbo.SECCITY S WHERE A.CITY = S.CITY AND S.DEPT_NO = '{0}')) g,", DEPT_ID);
            }
            FSQLStr += "CMEPMS.dbo.PJDCATY h";
            FSQLStr += " where d.SYR=e.SYR and d.TD_CODE=e.TD_CODE and d.TD_CODE=f.TD_CODE and f.TD_CITY=g.CITY and d.PID=h.PID and d.CATY_CODE=h.CATY_CODE";
            FSQLStr += " order by g.AREA,g.CITY_ID,f.TD_CITY,d.CATY_CODE,substring(d.TD_CODE,len(d.TD_CODE)-3,4)";

            SqlDataAdapter adp = new SqlDataAdapter(FSQLStr, con);
            DataSet ds = new DataSet();
            adp.Fill(ds);
            DataTable dt = ds.Tables[0];

            dt.Columns.Add("PPAY_FEE");
            dt.Columns.Add("L_PLN_COST");
            dt.Columns.Add("L_REL_COST");
            dt.Columns.Add("L_UNPAY_COST");
            dt.Columns.Add("PRE_SCHD");
            dt.Columns.Add("REL_SCHD");
            dt.Columns.Add("EXECT_DESC");
            dt.Columns.Add("DELAY_RESN");
            dt.Columns.Add("CTDELAY_RESN");

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["PPAY_FEE"] = 0;
                dt.Rows[i]["L_PLN_COST"] = 0;
                dt.Rows[i]["L_REL_COST"] = 0;
                dt.Rows[i]["L_UNPAY_COST"] = 0;
            }

            FSQLStr = "select C.TD_CODE";
            FSQLStr += ",case when datepart(d,getdate())>(select convert(int,DESC_COD) from CMEPMS.dbo.SMCCODE where SYS_ID='SC' and CATY_ID='01' and CATY_COD='ACDT') then";
            //FSQLStr += " isnull((select sum(isnull(COST_FEE,0)+isnull(PPAY_FEE,0)) from TDDSTCF where FD_CODE=C.FD_CODE and SYR+SMN=substring(dbo.getCHTDate(getdate(),''),1,5)),0)";
            FSQLStr += string.Format(" isnull((select sum(isnull(COST_FEE,0)+isnull(PPAY_FEE,0)) from TDDSTCF where FD_CODE=C.FD_CODE and SYR+SMN='{0}{1}'),0)", SYR, SMN);
            //FSQLStr += " else isnull((select sum(isnull(COST_FEE,0)+isnull(PPAY_FEE,0)) from TDDSTCF where FD_CODE=C.FD_CODE and SYR+SMN=substring(dbo.getCHTDate(dateadd(m,-1,getdate()),''),1,5)),0)";
            FSQLStr += string.Format(" else isnull((select sum(isnull(COST_FEE,0)+isnull(PPAY_FEE,0)) from TDDSTCF where FD_CODE=C.FD_CODE and SYR+SMN='{0}{1}'),0)", SYR, SMN);
            FSQLStr += " end PPAY_FEE";
            FSQLStr += " from CMEPMS.dbo.CTDMNDS C,CMETMS.dbo.TDMFORM T,CMEPMS.dbo.PJMFORM P";
            FSQLStr = string.Format("{0} where C.PID={1} and P.SYR='{2}' and C.TD_CODE=T.TD_CODE and C.PID=P.PID", FSQLStr, PID, SYR);

            SqlDataAdapter adp1 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds2 = new DataSet();
            adp1.Fill(ds2);
            DataTable dtTemp = ds2.Tables[0];
            for (int i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["PPAY_FEE"] = dtTemp.Rows[i]["PPAY_FEE"].ToString();
                    }
                }
            }

            FSQLStr = "select b.TD_CODE,isnull(c.L_PLN_COST,0) as L_PLN_COST,isnull(c.L_REL_COST,0) as L_REL_COST,isnull(c.L_UNPAY_COST,0) as L_UNPAY_COST,isnull(c.SAVE_FEE,0) as SAVE_FEE";
            FSQLStr = string.Format("{0} from CMEPMS.dbo.PJDCATY a,CMEPMS.dbo.CTDMNDS b,CMETMS.dbo.TDDMSCD c", FSQLStr);
            FSQLStr = string.Format("{0} where a.PID='{1}' and a.PID=b.PID and a.CATY_CODE=b.CATY_CODE and b.TD_CODE=c.TD_CODE and c.SYR='{2}' and c.SMN='{3}'", FSQLStr, PID, SYR, SMN);

            SqlDataAdapter adp2 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds3 = new DataSet();
            adp2.Fill(ds3);
            dtTemp = ds3.Tables[0];
            for (int i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["L_PLN_COST"] = dtTemp.Rows[i]["L_PLN_COST"].ToString();
                        dt.Rows[j]["L_REL_COST"] = dtTemp.Rows[i]["L_REL_COST"].ToString();
                        dt.Rows[j]["L_UNPAY_COST"] = dtTemp.Rows[i]["L_UNPAY_COST"].ToString();
                        //dt.Rows[j]["SAVE_FEE"] = dtTemp.Rows[i]["SAVE_FEE"].ToString();
                    }
                }
            }

            FSQLStr = "select";
            FSQLStr += " a.PID,c.TD_CODE";
            FSQLStr += " from";
            FSQLStr += string.Format(" (select PID from CMEPMS.dbo.PJMFORM where PID={0}) a", PID);
            FSQLStr += " inner join (select PID,TD_CODE from CMEPMS.dbo.CTDMNDS) b on a.PID=b.PID";
            FSQLStr += " inner join (select TD_CODE,IPJ_CODE from CMETMS.dbo.TDMFORM where IPJ_CODE in ('01','03')) c on b.TD_CODE=c.TD_CODE";

            SqlDataAdapter adp3 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds4 = new DataSet();
            adp3.Fill(ds4);
            dtTemp = ds4.Tables[0];
            for (int i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["L_UNPAY_COST"] = 0;
                    }
                }
            }

            FSQLStr = "select b.TD_CODE,c.PRE_SCHD,c.REL_SCHD,c.EXECT_DESC,c.DELAY_RESN,c.CTDELAY_RESN";
            FSQLStr = string.Format("{0} from CMEPMS.dbo.PJDCATY a,CMEPMS.dbo.CTDMNDS b,CMETMS.dbo.TDDMSCD c", FSQLStr);
            FSQLStr = string.Format("{0} where a.PID='{1}' and a.PID=b.PID and a.CATY_CODE=b.CATY_CODE and b.TD_CODE=c.TD_CODE and c.SYR='{2}' and c.SMN='{3}'", FSQLStr, PID, SYR, SMN);

            SqlDataAdapter adp4 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds5 = new DataSet();
            adp4.Fill(ds5);
            dtTemp = ds5.Tables[0];
            for (int i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["PRE_SCHD"] = dtTemp.Rows[i]["PRE_SCHD"].ToString();
                        dt.Rows[j]["REL_SCHD"] = dtTemp.Rows[i]["REL_SCHD"].ToString();
                        dt.Rows[j]["EXECT_DESC"] = dtTemp.Rows[i]["EXECT_DESC"].ToString();
                        dt.Rows[j]["DELAY_RESN"] = dtTemp.Rows[i]["DELAY_RESN"].ToString();
                        dt.Rows[j]["CTDELAY_RESN"] = dtTemp.Rows[i]["CTDELAY_RESN"].ToString();
                    }
                }
            }
            FSQLStr = string.Format("select TD_CODE,sum(isnull(IT_WGT,0)) as PRE_SCHD from CMETMS.dbo.TDDSDWB where TD_CODE in (select a.TD_CODE from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDDSDWB b where a.PID={0} and a.TD_CODE=b.TD_CODE and b.IT_CODE='TK105' and isnull(b.REL_EDT,'')='') and PRE_EDT is not null and PRE_EDT<>'' and substring(PRE_EDT,1,5)<='{1}{2}' group by TD_CODE", PID, SYR, SMN);

            SqlDataAdapter adp5 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds6 = new DataSet();
            adp5.Fill(ds6);
            dtTemp = ds6.Tables[0];
            for (int i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["PRE_SCHD"] = dtTemp.Rows[i]["PRE_SCHD"].ToString();
                    }
                }
            }
            FSQLStr = string.Format("select TD_CODE,sum(isnull(IT_WGT,0)) as REL_SCHD from CMETMS.dbo.TDDSDWB where TD_CODE in (select a.TD_CODE from CMEPMS.dbo.CTDMNDS a,CMETMS.dbo.TDDSDWB b where a.PID={0} and a.TD_CODE=b.TD_CODE and b.IT_CODE='TK105' and isnull(b.REL_EDT,'')='') and REL_EDT is not null and REL_EDT<>'' group by TD_CODE", PID);

            SqlDataAdapter adp6 = new SqlDataAdapter(FSQLStr, con);
            DataSet ds7 = new DataSet();
            adp6.Fill(ds7);
            dtTemp = ds7.Tables[0];
            for (int i = 0; i < dtTemp.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    if (dtTemp.Rows[i]["TD_CODE"].ToString() == dt.Rows[j]["TD_CODE"].ToString())
                    {
                        dt.Rows[j]["REL_SCHD"] = dtTemp.Rows[i]["REL_SCHD"].ToString();
                    }
                }
            }
            dt.Columns.Add("SEQ");
            dt.Columns.Add("CATY_FUND_");
            string catyName = "";
            double catyFund = 0;
            string tdCity = "";
            int index = -1;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["SEQ"] = i + 1;
                if (catyName == dt.Rows[i]["CATY_NAME"].ToString())
                {
                    dt.Rows[i]["CATY_NAME"] = "";
                    if (dt.Rows[i]["YR_FUND"].ToString() != "")
                    {
                        catyFund = catyFund + Convert.ToDouble(dt.Rows[i]["YR_FUND"].ToString());
                    }
                }
                else
                {
                    if (index > -1)
                    {
                        dt.Rows[index]["CATY_FUND_"] = catyFund;
                    }
                    if (dt.Rows[i]["YR_FUND"].ToString() != "")
                    {
                        catyFund = Convert.ToDouble(dt.Rows[i]["YR_FUND"].ToString());
                    }
                    else
                    {
                        catyFund = 0;
                    }
                    index = i;
                    catyName = dt.Rows[i]["CATY_NAME"].ToString();
                }
                if (tdCity == dt.Rows[i]["TD_CITY"].ToString())
                {
                    dt.Rows[i]["TD_CITY"] = "";
                }
                else
                {
                    tdCity = dt.Rows[i]["TD_CITY"].ToString();
                }
            }

            if (index > -1)
            {
                dt.Rows[index]["CATY_FUND_"] = catyFund;
            }
            return dt;
        }

        public static string GetBillionProject07(string PID, string SYR, string SMN, string connectionString, string Usr_ID, string Dept)
        {

            string FSQLStr;
            FSQLStr = string.Format("select * from CMEPMS.dbo.PJDFDCR where PID={0} order by SEQ_NO desc", PID);
            SqlConnection con = new SqlConnection();
            //string constr = @"Data Source=10.1.1.117\SEPMIS;initial catalog=CMEPMS;user id=***;password=***;" + "MultipleActiveResultSets=True;" + "initial catalog=CMETMS;";
            con.ConnectionString = connectionString;
            SqlDataAdapter adp = new SqlDataAdapter(FSQLStr, con);
            DataSet ds = new DataSet();
            adp.Fill(ds);
            DataTable dt2 = ds.Tables[0];

            string ITFD = "CTDITFD";
            int SEQ = 0;
            if (dt2.Rows.Count > 0)
            {
                if (dt2.Rows[0]["APROV_DT"].ToString() == "")
                {
                    ITFD = "CTHITFD";
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        if (dt2.Rows[i]["APROV_DT"].ToString() != "")
                        {
                            SEQ = Convert.ToInt16(dt2.Rows[i]["SEQ_NO"].ToString());
                            break;
                        }
                    }
                }
            }
            FSQLStr = "select distinct a.TD_CODE, a.TD_CITY,a.TD_TOWN,a.COST_FEE,a.DELAY_RESN,a.EXECT_DESC,a.FD_CODE,a.FUND,a.PPAY_FEE,a.PRE_SCHD,a.PRE_TK102,a.PRE_TK105,a.PRE_TK107,a.REL_SCHD,a.REL_TK102,a.REL_TK105,a.REL_TK107,a.REV_COST_FEE,a.REV_FUND,a.REV_PPAY_FEE,a.REV_SMN_FUND\r\n,a.SMN_FUND,a.TD_NAME from\r\n(select a.TD_CITY,a.TD_TOWN,b.COST_FEE,b.DELAY_RESN,b.EXECT_DESC,b.FD_CODE,b.FUND,b.PPAY_FEE,b.PRE_SCHD,b.PRE_TK102,b.PRE_TK105,b.PRE_TK107,b.REL_SCHD,b.REL_TK102,b.REL_TK105,b.REL_TK107,b.REV_COST_FEE,b.REV_FUND,b.REV_PPAY_FEE,b.REV_SMN_FUND\r\n,b.SMN_FUND,b.TD_CODE,b.TD_NAME from CMETMS.dbo.TDMFORM a,";
            FSQLStr += "(select a.*,isnull(b.PRE_TK102,'') as PRE_TK102,isnull(b.REL_TK102,'') as REL_TK102,isnull(c.PRE_TK105,'') as PRE_TK105,isnull(c.REL_TK105,'') as REL_TK105,isnull(d.PRE_TK107,'') as PRE_TK107,isnull(d.REL_TK107,'') as REL_TK107,isnull(e.PRE_SCHD,0) as PRE_SCHD,isnull(e.REL_SCHD,0) as REL_SCHD,isnull(e.EXECT_DESC,'') as EXECT_DESC,isnull(e.DELAY_RESN,'') as DELAY_RESN";
            FSQLStr += " from (";

            FSQLStr += "select zz.TD_CODE,zz.TD_NAME,zz.FD_CODE,sum(zz.FUND) as FUND,sum(zz.REV_FUND) as REV_FUND,sum(zz.SMN_FUND) as SMN_FUND,sum(zz.REV_SMN_FUND) as REV_SMN_FUND,sum(zz.COST_FEE) as COST_FEE,sum(zz.REV_COST_FEE) as REV_COST_FEE,sum(PPAY_FEE) as PPAY_FEE,sum(REV_PPAY_FEE) as REV_PPAY_FEE from (";

            FSQLStr += "select isnull(i.TD_CODE,j.TD_CODE) as TD_CODE,isnull(i.TD_NAME,j.TD_NAME) as TD_NAME,substring(isnull(i.FD_CODE,j.FD_CODE),12,4) as FD_CODE,isnull(i.FUND,0) as FUND,isnull(j.REV_FUND,0) as REV_FUND";
            FSQLStr += ",isnull(i.SMN_FUND,0) as SMN_FUND,isnull(j.REV_SMN_FUND,0) as REV_SMN_FUND,isnull(i.COST_FEE,0) as COST_FEE,isnull(j.REV_COST_FEE,0) as REV_COST_FEE";
            FSQLStr += ",isnull(i.PPAY_FEE,0) as PPAY_FEE,isnull(j.REV_PPAY_FEE,0) as REV_PPAY_FEE";
            FSQLStr += " from";
            FSQLStr += " (select a.TD_CODE,a.TD_NAME,a.FD_CODE,isnull(b.FUND,0) as FUND,isnull(c.SMN_FUND,0) as SMN_FUND,isnull(d.COST_FEE,0) as COST_FEE,isnull(d.PPAY_FEE,0) as PPAY_FEE";
            FSQLStr += " from";
            FSQLStr += string.Format(" (select TID,TD_CODE,TD_NAME,FD_CODE from CMEPMS.dbo.CTDMNDS where PID={0}) a", PID);
            if (ITFD == "CTDITFD")
            {
                FSQLStr += " left join (select TID,FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00') b on a.TID=b.TID";
                FSQLStr += " left join (select x.TID,sum(isnull(x.SMN_FUND,0)) as SMN_FUND from";
                FSQLStr += " (select TID,'01' as SMN,M_01 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'02' as SMN,M_02 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'03' as SMN,M_03 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'04' as SMN,M_04 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'05' as SMN,M_05 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'06' as SMN,M_06 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'07' as SMN,M_07 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'08' as SMN,M_08 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'09' as SMN,M_09 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'10' as SMN,M_10 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'11' as SMN,M_11 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00'";
                FSQLStr += " union select TID,'12' as SMN,M_12 as SMN_FUND from CMEPMS.dbo.CTDITFD where FIT_CODE='00') x";
            }
            else
            {
                FSQLStr += string.Format(" left join (select TID,FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}) b on a.TID=b.TID", SEQ);
                FSQLStr += " left join (select x.TID,sum(isnull(x.SMN_FUND,0)) as SMN_FUND from";
                FSQLStr += string.Format(" (select TID,'01' as SMN,M_01 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'02' as SMN,M_02 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'03' as SMN,M_03 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'04' as SMN,M_04 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'05' as SMN,M_05 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'06' as SMN,M_06 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'07' as SMN,M_07 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'08' as SMN,M_08 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'09' as SMN,M_09 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'10' as SMN,M_10 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'11' as SMN,M_11 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}", SEQ);
                FSQLStr += string.Format(" union select TID,'12' as SMN,M_12 as SMN_FUND from CMEPMS.dbo.CTHITFD where FIT_CODE='00' and SEQ_NO={0}) x", SEQ);
            }
            FSQLStr += string.Format(" where x.SMN<'{0}'", SMN);
            FSQLStr += " group by x.TID) c on a.TID=c.TID";
            FSQLStr += string.Format(" left join (select TD_CODE1 as TD_CODE,FD_CODE,isnull(COST_FEE,0) as COST_FEE,isnull(PPAY_FEE,0) as PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR='{0}' and SMN='{1}') d on a.TD_CODE=d.TD_CODE and a.FD_CODE=d.FD_CODE) i", SYR, SMN);
            FSQLStr += " full join";
            //FSQLStr += " (select a.TD_CODE1 as TD_CODE,a.TD_NAME,a.FD_CODE,isnull(a.REV_FUND,0) as REV_FUND,isnull(b.SMN_FUND,0) as REV_SMN_FUND,isnull(c.COST_FEE,0) as REV_COST_FEE,isnull(c.PPAY_FEE,0) as REV_PPAY_FEE from";
            FSQLStr += " (select a.TD_CODE1 as TD_CODE,a.TD_NAME,a.FD_CODE,isnull(a.REV_FUND,0)-isnull(d.PPAY_FEE,0) as REV_FUND,isnull(b.SMN_FUND,0) as REV_SMN_FUND,isnull(c.COST_FEE,0) as REV_COST_FEE,isnull(c.PPAY_FEE,0) as REV_PPAY_FEE from";
            //FSQLStr += string.Format(" (select PID,SYR,TD_CODE1,TD_NAME,FD_CODE,REV_FUND from CMEPMS.dbo.CRDMNDS where PID={0} and TD_CODE1 in (select TD_CODE from CMEPMS.dbo.CTDMNDS where TID in (select TID from CMEPMS.dbo.CTDITFD where SYR='{1}'))) a", PID,SYR);
            FSQLStr += string.Format(" (select PID,SYR,TD_CODE1,TD_NAME,FD_CODE,REV_FUND from CMEPMS.dbo.CRDMNDS where PID={0}) a", PID);
            FSQLStr += " left join (select x.PID,x.SYR,x.TD_CODE1,x.FD_CODE,sum(isnull(SMN_FUND,0)) as SMN_FUND from";
            FSQLStr += " (select PID,SYR,TD_CODE1,FD_CODE,'01' as SMN,M_01 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'02' as SMN,M_02 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'03' as SMN,M_03 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'04' as SMN,M_04 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'05' as SMN,M_05 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'06' as SMN,M_06 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'07' as SMN,M_07 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'08' as SMN,M_08 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'09' as SMN,M_09 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'10' as SMN,M_10 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'11' as SMN,M_11 as SMN_FUND from CMEPMS.dbo.CRDMNDS";
            FSQLStr += " union select PID,SYR,TD_CODE1,FD_CODE,'12' as SMN,M_12 as SMN_FUND from CMEPMS.dbo.CRDMNDS) x";
            FSQLStr += string.Format(" where x.SYR='{0}' and x.SMN<='{1}'", SYR, SMN);
            FSQLStr += " group by x.PID,x.SYR,x.TD_CODE1,x.FD_CODE) b on a.PID=b.PID and a.SYR=b.SYR and a.TD_CODE1=b.TD_CODE1 and a.FD_CODE=b.FD_CODE";
            FSQLStr += string.Format(" left join (select TD_CODE1,FD_CODE,isnull(COST_FEE,0) as COST_FEE,isnull(PPAY_FEE,0) as PPAY_FEE from CMEPMS.dbo.TDDSTRF where SYR='{0}' and SMN='{1}') c on a.TD_CODE1=c.TD_CODE1 and a.FD_CODE=c.FD_CODE", SYR, SMN);

            FSQLStr += string.Format(" left join (select isnull(x.FD_CODE,y.FD_CODE) as FD_CODE,isnull(x.PPAY_FEE,y.PPAY_FEE) as PPAY_FEE from (select FD_CODE,PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR='{0}' and SMN='12') x full join (select FD_CODE,PPAY_FEE from CMEPMS.dbo.TDDSTRF where SYR='{0}' and SMN='12') y on x.FD_CODE=y.FD_CODE) d on a.FD_CODE=d.FD_CODE", Convert.ToInt16(Convert.ToInt16(SYR) - 1).ToString("000"));

            FSQLStr += ") j on i.TD_CODE=j.TD_CODE";

            FSQLStr += ") zz group by zz.TD_CODE,zz.TD_NAME,zz.FD_CODE";

            FSQLStr += ") a";
            FSQLStr += " left join (select TD_CODE,PRE_EDT as PRE_TK102,REL_EDT as REL_TK102 from CMETMS.dbo.TDDSDWB where IT_CODE='TK103') b on a.TD_CODE=b.TD_CODE";
            FSQLStr += " left join (select TD_CODE,PRE_EDT as PRE_TK105,REL_EDT as REL_TK105 from CMETMS.dbo.TDDSDWB where IT_CODE='TK105') c on a.TD_CODE=c.TD_CODE";
            FSQLStr += " left join (select TD_CODE,PRE_EDT as PRE_TK107,REL_EDT as REL_TK107 from CMETMS.dbo.TDDSDWB where IT_CODE='TK107') d on a.TD_CODE=d.TD_CODE";
            FSQLStr += string.Format(" left join (select TD_CODE,PRE_SCHD,REL_SCHD,EXECT_DESC,DELAY_RESN from CMETMS.dbo.TDDMSCD where SYR='{0}' and SMN='{1}') e on a.TD_CODE=e.TD_CODE", SYR, SMN);
            FSQLStr += " where a.FUND<>0 or a.REV_FUND<>0) b";
            FSQLStr += " where a.TD_CODE=b.TD_CODE) a,CMBASE.dbo.SECCITY b ";
            if (db.SECCITY.Where(m => m.USR_ID == Usr_ID).Any())
            {
                FSQLStr += string.Format("where a.TD_CITY=b.CITY and b.DEPT_NO='{0}'", Dept);
            }
            else
            {
                FSQLStr += string.Format("where a.TD_CITY=b.CITY and b.USR_ID='{0}'", Usr_ID);
            }
            FSQLStr += " order by a.FD_CODE,a.TD_CODE";
            return FSQLStr;
        }

        /// <summary>
        /// B-1.修正預算調整表(資料來源)
        /// </summary>
        /// <param name="pID"> 工程編號 </param>
        /// <param name="SEQ_NO"> 本期期數 </param>
        /// <param name="sACC_ID"></param>
        /// <param name="ADV_UID"></param>
        /// <param name="aDV_DP"></param>
        /// <param name="Dept"></param>
        /// <param name="Choose"></param>
        /// <returns></returns>
        public static DataSet SP_RCD0010_UDF003(string pID, string SEQ_NO, string SYR, string sACC_ID, string ADV_UID, string aDV_DP, string Dept, string Choose, string connectionString)
        {
            Debug.WriteLine($"pID:{pID}");
            Debug.WriteLine($"SEQ_NO:{SEQ_NO}");
            Debug.WriteLine($"SYR:{SYR}");
            Debug.WriteLine($"sACC_ID:{sACC_ID}");
            Debug.WriteLine($"ADV_UID:{ADV_UID}");
            Debug.WriteLine($"aDV_DP:{aDV_DP}");
            Debug.WriteLine($"Dept:{Dept}");

            string presentSql = "";
            string prophaseSql = "";

            /// 計算前一期數
            int sSEQ_NO = Convert.ToInt32(SEQ_NO) - 1;
            string PSEQ_NO = sSEQ_NO.ToString();

            decimal PPID = Convert.ToDecimal(pID);

            DataSet ds = new DataSet();
            DataTable dt1 = new DataTable(); // 本期
            DataTable dt2 = new DataTable(); // 前期

            #region 本期資料

            if (Convert.ToInt32(SEQ_NO) == 0)
            {
                presentSql = $@"
                            WITH modifyBudgetManagement AS (
                                    SELECT 
                                        T.PID,
                                        T.PID_SYR,
                                        T.TID,
		                                T.SEQ_NO,
                                        T.FD_CODE,
                                        T.TD_CODE,
                                        T.TD_NAME,
                                        T.SIP_TD_CODE,
                                        CASE T.CITY
                                            WHEN '台北縣' THEN '新北市'
                                            WHEN '台中縣' THEN '台中市'
                                            WHEN '台南縣' THEN '台南市'
                                            WHEN '高雄縣' THEN '高雄市'
                                            WHEN '桃園縣' THEN '桃園市'
                                            ELSE T.CITY
                                        END AS CITY,
                                        T.TOWN,
                                        T.ADM_DIST,
                                        T.CEN_SP_RATIO,
                                        T.CATY_CODE,
                                        CASE ISNULL(T.BOT_FG, 0)
                                            WHEN 1 THEN
                                                CASE T.CATY_CODE
                                                    WHEN 'L' THEN '用地取得'
                                                    WHEN 'J' THEN '先期規劃'
                                                    WHEN 'M' THEN '專案管理'
                                                    ELSE '政府應辦事項'
                                                END
                                            ELSE
                                                CASE T.CATY_CODE
                                                    WHEN 'C' THEN '工程類'
                                                    ELSE '規劃類'
                                                END
                                        END AS CATY_DESC,
                                        T.IPJ_CODE,
                                        CASE T.N_IPJ_CODE
                                            WHEN '01' THEN '01'
                                            WHEN '02' THEN '03'
                                            WHEN '03' THEN '02'
                                            WHEN '04' THEN '04'
                                            ELSE T.N_IPJ_CODE
                                        END AS IPJ_ORDER,
                                        T.IPJ_DESC,
                                        T.YR_FUND,
                                        T.CEN_FUND,
                                        T.LOC_FUND,
                                        T.YR_WORK,
                                        T.EMG_CATY,
                                        T.BOT_FG,
                                        T.ACC_ID,
                                        T.ACC_DESC,
                                        T.ADV_DP,
                                        T.FCAL_FG,
                                        T.UCNT_FG,
                                        CASE T.CATY_CODE
                                            WHEN 'C' THEN 'C'
                                            ELSE 'P'
                                        END AS FUND_TYPE,
                                        CASE T.CATY_CODE
                                            WHEN 'L' THEN 'L'
                                            WHEN 'J' THEN 'J'
                                            WHEN 'M' THEN 'M'
                                            ELSE 'C'
                                        END AS PROJECT_TYPE,
                                        SUM(TD.COST_FEE + TD.PPAY_FEE) AS TOTAL_COST,
                                        CMEPMS.dbo.getCNTFDS(T.TID) AS FUND_DESC,
                                        MAX(SP.CPA_CITY_ID) AS CITY_ORDER,
                                        SUM(CF.M_01) AS M_01,
                                        SUM(CF.M_02) AS M_02,
                                        SUM(CF.M_03) AS M_03,
                                        SUM(CF.M_04) AS M_04,
                                        SUM(CF.M_05) AS M_05,
                                        SUM(CF.M_06) AS M_06,
                                        SUM(CF.M_07) AS M_07,
                                        SUM(CF.M_08) AS M_08,
                                        SUM(CF.M_09) AS M_09,
                                        SUM(CF.M_10) AS M_10,
                                        SUM(CF.M_11) AS M_11,
                                        SUM(CF.M_12) AS M_12,
                                        CASE T.BOT_FG
                                            WHEN 1 THEN 'C5'
                                            ELSE
                                                CASE T.CATY_CODE
                                                    WHEN 'P' THEN 'C1'
                                                    WHEN 'C' THEN 'C2'
                                                    WHEN 'L' THEN 'C3'
                                                    ELSE 'C4'
                                                END
                                        END AS PROJECT_CATEGORY,
                                        MAX(D.EXECT_DESC) AS EXECT_DESC,
                                        ROW_NUMBER() OVER (
                                            PARTITION BY T.FD_CODE 
                                            ORDER BY D.SYR DESC, D.SMN DESC
                                        ) AS RowNum
                                    FROM 
                                        CMEPMS.dbo.TDV_TDDHFTS AS T
                                    LEFT JOIN TDDSTCF AS TD
                                        ON T.FD_CODE = TD.FD_CODE AND TD.SYR = '{SYR}' 
                                    LEFT JOIN CMBASE.dbo.SMCPOST AS SP
                                        ON SP.CITY = CASE T.CITY
                                                        WHEN '新北市' THEN '台北縣'
                                                        WHEN '台中市' THEN '台中縣'
                                                        WHEN '台南市' THEN '台南縣'
                                                        WHEN '高雄市' THEN '高雄縣'
                                                        WHEN '桃園市' THEN '桃園縣'
                                                        ELSE T.CITY
                                                     END
                                    LEFT JOIN CTDITFD AS CF
                                        ON CF.TID = T.TID
                                    LEFT JOIN CMETMS.dbo.TDDMSCD AS D
                                        ON D.TD_CODE = T.TD_CODE AND D.SYR = '{SYR}' 
                                    WHERE 
                                        T.PID IN ({PPID})
		                                AND T.SEQ_NO = '{SEQ_NO}'
                                    GROUP BY 
                                        T.PID, T.PID_SYR,T.SEQ_NO, T.TID, T.FD_CODE, T.TD_CODE, T.TD_NAME, T.SIP_TD_CODE,
                                        T.CITY, T.TOWN, T.ADM_DIST, T.CEN_SP_RATIO, T.CATY_CODE, T.BOT_FG,
                                        T.IPJ_CODE, T.N_IPJ_CODE, T.IPJ_DESC, T.YR_FUND, T.CEN_FUND, T.LOC_FUND,
                                        T.YR_WORK, T.EMG_CATY, T.ACC_ID, T.ACC_DESC, T.ADV_DP, T.FCAL_FG, T.UCNT_FG,
                                        D.SYR, D.SMN
                                )
                                SELECT *
                                FROM modifyBudgetManagement
                                WHERE RowNum = 1
                                ORDER BY FD_CODE; ";
            }
            else
            {
                presentSql = $@"
                            WITH modifyBudgetManagement AS (
                                    SELECT 
                                        T.PID,
                                        T.PID_SYR,
                                        T.TID,
		                                T.SEQ_NO,
                                        T.FD_CODE,
                                        T.TD_CODE,
                                        T.TD_NAME,
                                        T.SIP_TD_CODE,
                                        CASE T.CITY
                                            WHEN '台北縣' THEN '新北市'
                                            WHEN '台中縣' THEN '台中市'
                                            WHEN '台南縣' THEN '台南市'
                                            WHEN '高雄縣' THEN '高雄市'
                                            WHEN '桃園縣' THEN '桃園市'
                                            ELSE T.CITY
                                        END AS CITY,
                                        T.TOWN,
                                        T.ADM_DIST,
                                        T.CEN_SP_RATIO,
                                        T.CATY_CODE,
                                        CASE ISNULL(T.BOT_FG, 0)
                                            WHEN 1 THEN
                                                CASE T.CATY_CODE
                                                    WHEN 'L' THEN '用地取得'
                                                    WHEN 'J' THEN '先期規劃'
                                                    WHEN 'M' THEN '專案管理'
                                                    ELSE '政府應辦事項'
                                                END
                                            ELSE
                                                CASE T.CATY_CODE
                                                    WHEN 'C' THEN '工程類'
                                                    ELSE '規劃類'
                                                END
                                        END AS CATY_DESC,
                                        T.IPJ_CODE,
                                        CASE T.N_IPJ_CODE
                                            WHEN '01' THEN '01'
                                            WHEN '02' THEN '03'
                                            WHEN '03' THEN '02'
                                            WHEN '04' THEN '04'
                                            ELSE T.N_IPJ_CODE
                                        END AS IPJ_ORDER,
                                        T.IPJ_DESC,
                                        T.YR_FUND,
                                        T.CEN_FUND,
                                        T.LOC_FUND,
                                        T.YR_WORK,
                                        T.EMG_CATY,
                                        T.BOT_FG,
                                        T.ACC_ID,
                                        T.ACC_DESC,
                                        T.ADV_DP,
                                        T.FCAL_FG,
                                        T.UCNT_FG,
                                        CASE T.CATY_CODE
                                            WHEN 'C' THEN 'C'
                                            ELSE 'P'
                                        END AS FUND_TYPE,
                                        CASE T.CATY_CODE
                                            WHEN 'L' THEN 'L'
                                            WHEN 'J' THEN 'J'
                                            WHEN 'M' THEN 'M'
                                            ELSE 'C'
                                        END AS PROJECT_TYPE,
                                        SUM(TD.COST_FEE + TD.PPAY_FEE) AS TOTAL_COST,
                                        dbo.getCNTFDS(T.TID) AS FUND_DESC,
                                        MAX(SP.CPA_CITY_ID) AS CITY_ORDER,
                                        SUM(CF.M_01) AS M_01,
                                        SUM(CF.M_02) AS M_02,
                                        SUM(CF.M_03) AS M_03,
                                        SUM(CF.M_04) AS M_04,
                                        SUM(CF.M_05) AS M_05,
                                        SUM(CF.M_06) AS M_06,
                                        SUM(CF.M_07) AS M_07,
                                        SUM(CF.M_08) AS M_08,
                                        SUM(CF.M_09) AS M_09,
                                        SUM(CF.M_10) AS M_10,
                                        SUM(CF.M_11) AS M_11,
                                        SUM(CF.M_12) AS M_12,
                                        CASE T.BOT_FG
                                            WHEN 1 THEN 'C5'
                                            ELSE
                                                CASE T.CATY_CODE
                                                    WHEN 'P' THEN 'C1'
                                                    WHEN 'C' THEN 'C2'
                                                    WHEN 'L' THEN 'C3'
                                                    ELSE 'C4'
                                                END
                                        END AS PROJECT_CATEGORY,
                                        MAX(D.EXECT_DESC) AS EXECT_DESC,
                                        ROW_NUMBER() OVER (
                                            PARTITION BY T.FD_CODE 
                                            ORDER BY D.SYR DESC, D.SMN DESC
                                        ) AS RowNum
                                    FROM 
                                        CMEPMS.dbo.TDV_TDDCFTS AS T
                                    LEFT JOIN TDDSTCF AS TD
                                        ON T.FD_CODE = TD.FD_CODE AND TD.SYR = '{SYR}' 
                                    LEFT JOIN CMBASE.dbo.SMCPOST AS SP
                                        ON SP.CITY = CASE T.CITY
                                                        WHEN '新北市' THEN '台北縣'
                                                        WHEN '台中市' THEN '台中縣'
                                                        WHEN '台南市' THEN '台南縣'
                                                        WHEN '高雄市' THEN '高雄縣'
                                                        WHEN '桃園市' THEN '桃園縣'
                                                        ELSE T.CITY
                                                     END
                                    LEFT JOIN CTDITFD AS CF
                                        ON CF.TID = T.TID
                                    LEFT JOIN CMETMS.dbo.TDDMSCD AS D
                                        ON D.TD_CODE = T.TD_CODE AND D.SYR = '{SYR}' 
                                    WHERE 
                                        T.PID IN ({PPID})
		                                AND T.SEQ_NO = '{SEQ_NO}'
                                    GROUP BY 
                                        T.PID, T.PID_SYR,T.SEQ_NO, T.TID, T.FD_CODE, T.TD_CODE, T.TD_NAME, T.SIP_TD_CODE,
                                        T.CITY, T.TOWN, T.ADM_DIST, T.CEN_SP_RATIO, T.CATY_CODE, T.BOT_FG,
                                        T.IPJ_CODE, T.N_IPJ_CODE, T.IPJ_DESC, T.YR_FUND, T.CEN_FUND, T.LOC_FUND,
                                        T.YR_WORK, T.EMG_CATY, T.ACC_ID, T.ACC_DESC, T.ADV_DP, T.FCAL_FG, T.UCNT_FG,
                                        D.SYR, D.SMN
                                )
                                SELECT *
                                FROM modifyBudgetManagement
                                WHERE RowNum = 1
                                ORDER BY FD_CODE; ";
            }

            dt1 = SourceFunction.DataTableGet(presentSql, connectionString);

            #endregion

            #region 前期資料

            prophaseSql = $@"
                            WITH modifyBudgetManagement AS (
                                    SELECT 
                                        T.PID,
                                        T.PID_SYR,
                                        T.TID,
		                                T.SEQ_NO,
                                        T.FD_CODE,
                                        T.TD_CODE,
                                        T.TD_NAME,
                                        T.SIP_TD_CODE,
                                        CASE T.CITY
                                            WHEN '台北縣' THEN '新北市'
                                            WHEN '台中縣' THEN '台中市'
                                            WHEN '台南縣' THEN '台南市'
                                            WHEN '高雄縣' THEN '高雄市'
                                            WHEN '桃園縣' THEN '桃園市'
                                            ELSE T.CITY
                                        END AS CITY,
                                        T.TOWN,
                                        T.ADM_DIST,
                                        T.CEN_SP_RATIO,
                                        T.CATY_CODE,
                                        CASE ISNULL(T.BOT_FG, 0)
                                            WHEN 1 THEN
                                                CASE T.CATY_CODE
                                                    WHEN 'L' THEN '用地取得'
                                                    WHEN 'J' THEN '先期規劃'
                                                    WHEN 'M' THEN '專案管理'
                                                    ELSE '政府應辦事項'
                                                END
                                            ELSE
                                                CASE T.CATY_CODE
                                                    WHEN 'C' THEN '工程類'
                                                    ELSE '規劃類'
                                                END
                                        END AS CATY_DESC,
                                        T.IPJ_CODE,
                                        CASE T.N_IPJ_CODE
                                            WHEN '01' THEN '01'
                                            WHEN '02' THEN '03'
                                            WHEN '03' THEN '02'
                                            WHEN '04' THEN '04'
                                            ELSE T.N_IPJ_CODE
                                        END AS IPJ_ORDER,
                                        T.IPJ_DESC,
                                        T.YR_FUND,
                                        T.CEN_FUND,
                                        T.LOC_FUND,
                                        T.YR_WORK,
                                        T.EMG_CATY,
                                        T.BOT_FG,
                                        T.ACC_ID,
                                        T.ACC_DESC,
                                        T.ADV_DP,
                                        T.FCAL_FG,
                                        T.UCNT_FG,
                                        CASE T.CATY_CODE
                                            WHEN 'C' THEN 'C'
                                            ELSE 'P'
                                        END AS FUND_TYPE,
                                        CASE T.CATY_CODE
                                            WHEN 'L' THEN 'L'
                                            WHEN 'J' THEN 'J'
                                            WHEN 'M' THEN 'M'
                                            ELSE 'C'
                                        END AS PROJECT_TYPE,
                                        SUM(TD.COST_FEE + TD.PPAY_FEE) AS TOTAL_COST,
                                        CMEPMS.dbo.getCNTFDS(T.TID) AS FUND_DESC,
                                        MAX(SP.CPA_CITY_ID) AS CITY_ORDER,
                                        SUM(CF.M_01) AS M_01,
                                        SUM(CF.M_02) AS M_02,
                                        SUM(CF.M_03) AS M_03,
                                        SUM(CF.M_04) AS M_04,
                                        SUM(CF.M_05) AS M_05,
                                        SUM(CF.M_06) AS M_06,
                                        SUM(CF.M_07) AS M_07,
                                        SUM(CF.M_08) AS M_08,
                                        SUM(CF.M_09) AS M_09,
                                        SUM(CF.M_10) AS M_10,
                                        SUM(CF.M_11) AS M_11,
                                        SUM(CF.M_12) AS M_12,
                                        CASE T.BOT_FG
                                            WHEN 1 THEN 'C5'
                                            ELSE
                                                CASE T.CATY_CODE
                                                    WHEN 'P' THEN 'C1'
                                                    WHEN 'C' THEN 'C2'
                                                    WHEN 'L' THEN 'C3'
                                                    ELSE 'C4'
                                                END
                                        END AS PROJECT_CATEGORY,
                                        MAX(D.EXECT_DESC) AS EXECT_DESC,
                                        ROW_NUMBER() OVER (
                                            PARTITION BY T.FD_CODE 
                                            ORDER BY D.SYR DESC, D.SMN DESC
                                        ) AS RowNum
                                    FROM 
                                        CMEPMS.dbo.TDV_TDDHFTS AS T
                                    LEFT JOIN TDDSTCF AS TD
                                        ON T.FD_CODE = TD.FD_CODE AND TD.SYR = '{SYR}' 
                                    LEFT JOIN CMBASE.dbo.SMCPOST AS SP
                                        ON SP.CITY = CASE T.CITY
                                                        WHEN '新北市' THEN '台北縣'
                                                        WHEN '台中市' THEN '台中縣'
                                                        WHEN '台南市' THEN '台南縣'
                                                        WHEN '高雄市' THEN '高雄縣'
                                                        WHEN '桃園市' THEN '桃園縣'
                                                        ELSE T.CITY
                                                     END
                                    LEFT JOIN CTDITFD AS CF
                                        ON CF.TID = T.TID
                                    LEFT JOIN CMETMS.dbo.TDDMSCD AS D
                                        ON D.TD_CODE = T.TD_CODE AND D.SYR = '{SYR}' 
                                    WHERE 
                                        T.PID IN ({PPID})
		                                AND T.SEQ_NO = '{PSEQ_NO}'
                                    GROUP BY 
                                        T.PID, T.PID_SYR,T.SEQ_NO, T.TID, T.FD_CODE, T.TD_CODE, T.TD_NAME, T.SIP_TD_CODE,
                                        T.CITY, T.TOWN, T.ADM_DIST, T.CEN_SP_RATIO, T.CATY_CODE, T.BOT_FG,
                                        T.IPJ_CODE, T.N_IPJ_CODE, T.IPJ_DESC, T.YR_FUND, T.CEN_FUND, T.LOC_FUND,
                                        T.YR_WORK, T.EMG_CATY, T.ACC_ID, T.ACC_DESC, T.ADV_DP, T.FCAL_FG, T.UCNT_FG,
                                        D.SYR, D.SMN
                                )
                                SELECT *
                                FROM modifyBudgetManagement
                                WHERE RowNum = 1
                                ORDER BY FD_CODE; ";

            dt2 = SourceFunction.DataTableGet(prophaseSql, connectionString);

            #endregion

            DataTable dt1Copy = dt1.Copy();
            dt1Copy.TableName = "dt1";
            ds.Tables.Add(dt1Copy);

            DataTable dt2Copy = dt2.Copy();
            dt2Copy.TableName = "dt2";
            ds.Tables.Add(dt2Copy);

            return ds;
        }

        /// <summary>
        /// G-1.雨水工程控管表_資料來源
        /// </summary>
        /// <param name="PLAN"></param>
        /// <param name="CITY"></param>
        /// <param name="YEAR1"></param>
        /// <param name="YEAR2"></param>
        /// <param name="MON"></param>
        /// <param name="TYPE"></param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        public static DataTable GetBillionData11(string PLAN, string CITY, string YEAR1, string YEAR2, string month, string TYPE, string connectionString)
        {
            Trace.WriteLine("G-1.雨水工程控管報表-資料來源");
            Trace.WriteLine($"PLAN:{PLAN}");
            Trace.WriteLine($"CITY:{CITY}");
            Trace.WriteLine($"YEAR1:{YEAR1}");
            Trace.WriteLine($"YEAR2:{YEAR2}");
            Trace.WriteLine($"MON:{month}");
            Trace.WriteLine($"TYPE:{TYPE}");
            Trace.WriteLine($"connectionString:{connectionString}");

            string SQL1 = "";
            string SQL2 = "";
            List<int> PIDList = new List<int>();
            DataTable ResultTable = new DataTable();

            /// 資料來源:SQL查詢
            #region =================================================================================================
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                /// Step 1. 取 PID List
                SQL1 = $@" SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PJ_CODE IN ({PLAN}) 
                        AND SYR BETWEEN '{YEAR1}' 
                        AND '{YEAR2}' 
                        AND ( ADV_UID = 'A270400' OR ADV_UID = 'B150400' ) 
                        AND ( (PJ_CODE IN ('01','02','03') AND SIP_W_CODE = 'G302') OR (PJ_CODE IN ('04','05') AND SIP_W_CODE = '4000') OR (PJ_CODE IN ('07') AND SIP_W_CODE = '4000') )";

                using (SqlCommand cmd = new SqlCommand(SQL1, con))

                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        int pid = Convert.ToInt32(dr["PID"]);
                        PIDList.Add(pid);
                    }
                }

                string PIDString = string.Join(",", PIDList);

                Trace.WriteLine($"PID : {PIDString}");
                Trace.WriteLine($"SQL1 : {SQL1}");

                string cityIn = string.Join(",", CITY.Split(',').Select(c => $"'{c.Trim()}'"));
                Trace.WriteLine($"CITY LIST : {cityIn}");

                SQL2 = $@"
                        SELECT
                            A.PID
                            ,E.TID
                            ,E.PID_SYR
                            ,E.SEQ_NO
                            ,E.SIP_TD_CODE          
                            ,A.TD_CODE
                            ,A.FD_CODE
                            ,B.CATY_CODE
                            ,D.TD_CITY -- 縣市
                            ,E.CITY
                            ,E.TOWN
                            ,S9.CODE_DESC AS IPJ_DESC -- 辦理層級
                            ,CAST(ROUND(ISNULL(E.CEN_SP_RATIO, 0), 0) AS INT) AS CEN_RATIO -- 補助比例
                            ,B.CATY_NAME -- 計畫說明
                            ,A.TD_NAME   -- 工作名稱
                            ,IIF(d.TD_FINISH = 1, N'已結案', IIF(d.TD_FINISH = 0, N'未結案', N'未填報')) AS TD_FINISH_DESC -- 結案情形
                            ,ISNULL((SELECT SUM(AE_CTM + PP_CTM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = {YEAR1} AND SMN = '12' AND FD_CODE = A.FD_CODE AND b.PID IN ({PIDString} ) ), 0) AS LAST_PPAY_FEE -- 當年累積撥發金額(仟元)
                            ,ISNULL((SELECT SUM(AE_CTM + PP_CTM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = {YEAR2} AND SMN = {month} AND FD_CODE = A.FD_CODE AND b.PID IN ( {PIDString} ) ), 0) AS ALL_PPAY_FEE -- 當年累積撥發金額(仟元)
                            ,D.START_FUND
                            ,B.CATY_FUND -- 分支計畫核定經費總計(仟元)
                            ,( ( SELECT SUM(BUDGET_FEE) FROM CMETMS.dbo.TDDBGLS WHERE REC_TP = 'B' AND TD_CODE = A.TD_CODE AND SEQ_NO = ( SELECT MAX(SEQ_NO) FROM CMETMS.dbo.TDDBGRC WHERE TD_CODE = CMETMS.dbo.TDDBGLS.TD_CODE AND REC_TP = 'B' ) ) * 0.95 * ISNULL(m.REL_SCHD, 0) / 100.0 * CAST(ROUND(ISNULL(E.CEN_SP_RATIO, 0), 0) AS INT) / 100.0 ) AS L_FUND -- 依進度可撥金費(李)
                            ,(SELECT SUM(FUND) FROM CMEPMS.dbo.CTDITFD WHERE TID = A.TID AND FIT_CODE = '00') AS CEN_FUND
                            ,ISNULL(F.REV_FUND, 0) AS REV_FUND
	                        ,ISNULL(F.OCR_FUND, 0) AS OCR_FUND
                            ,(SELECT SUM(BUDGET_FEE) FROM CMETMS.dbo.TDDBGLS WHERE REC_TP = 'B' AND TD_CODE = A.TD_CODE AND SEQ_NO = (SELECT MAX(SEQ_NO) FROM CMETMS.dbo.TDDBGRC WHERE TD_CODE = CMETMS.dbo.TDDBGLS.TD_CODE AND REC_TP = 'B')) AS B_FUND
                            -----------------------------------------------------------------------------------------------------------------------
                            ,STUFF(STUFF(S1.PRE_EDT,4,0,'.'),7,0,'.') AS PRE_TK103  -- 預定細設完成日期 
                            ,STUFF(STUFF(S1.REL_EDT,4,0,'.'),7,0,'.') AS REL_TK103  -- 實際細設完成日期
                            ,STUFF(STUFF(S2.PRE_EDT,4,0,'.'),7,0,'.') AS PRE_TK105  -- 預定發包日期
                            ,STUFF(STUFF(S2.REL_EDT,4,0,'.'),7,0,'.') AS REL_TK105  -- 實際發包日期
                            ,STUFF(STUFF(S6.REL_EDT,4,0,'.'),7,0,'.') AS REL_TK102  -- 基設核定日期
                            ,STUFF(STUFF(S10.PRE_EDT,4,0,'.'),7,0,'.') AS PRE_TK108 -- 預定正驗日期
                            ,STUFF(STUFF(S10.REL_EDT,4,0,'.'),7,0,'.') AS REL_TK108 -- 實際正驗日期
                            ,ISNULL(M.PRE_SCHD, 0) AS PRE_SCHD -- 預定工程進度
                            ,ISNULL(M.REL_SCHD, 0) AS REL_SCHD -- 實際工程進度
                            ,(ISNULL(M.REL_SCHD, 0) - ISNULL(M.PRE_SCHD, 0)) AS SCHD_DIFF -- 進度差異
                            ,CASE WHEN D.CATY_CODE = 'P' AND S3.PRE_EDT IS NOT NULL THEN STUFF(STUFF(S3.PRE_EDT,4,0,'.'),7,0,'.') ELSE STUFF(STUFF(S4.PRE_EDT,4,0,'.'),7,0,'.') END AS PRE_TK107 -- 預定完工日期
                            ,CASE WHEN D.CATY_CODE = 'P' AND S3.REL_EDT IS NOT NULL THEN STUFF(STUFF(S3.REL_EDT,4,0,'.'),7,0,'.') ELSE STUFF(STUFF(S4.REL_EDT,4,0,'.'),7,0,'.') END AS REL_TK107 -- 實際完工日期
	                        -- ,S3.PRE_EDT
                            -- ,S3.REL_EDT
                            -- ,S4.PRE_EDT
                            -- ,S4.REL_EDT
	                        -- ,S6.PRE_EDT AS PRE_TK102  
                            -- ,S7.PRE_EDT AS PRE_TK1031
                            -- ,S7.REL_EDT AS REL_TK1031
	                        -----------------------------------------------------------------------------------------------------------------------
	                        ,ISNULL(M.EXECT_DESC, '') AS EXECT_DESC -- 辦理情形或落後原因
                            ,M.DELAY_RESN							-- 落後改善對策
                            ,M.CTDELAY_RESN							-- 備註
                            ,D.DSG_BUSS_NAME						-- 設計單位
                            ,D.BUSS_NAME							-- 廠商名稱
                        FROM CMEPMS.dbo.CTDMNDS A
                        LEFT JOIN CMEPMS.dbo.PJDCATY B ON B.PID = A.PID AND B.CATY_CODE = A.CATY_CODE
                        LEFT JOIN CMEPMS.dbo.TDV_TDDHFTS C ON C.TD_CODE = A.TD_CODE AND C.TID = A.TID AND C.SEQ_NO = ( SELECT MAX(SEQ_NO) FROM CMEPMS.dbo.TDV_TDDHFTS WHERE FD_CODE = A.FD_CODE )
                        LEFT JOIN CMETMS.dbo.TDMFORM D ON D.TD_CODE = A.TD_CODE
                        LEFT JOIN CMEPMS.dbo.TDV_TDDCFTS E ON E.TD_CODE = A.TD_CODE AND E.TID = A.TID AND E.SEQ_NO = ( SELECT MAX(SEQ_NO) FROM CMEPMS.dbo.TDV_TDDCFTS WHERE FD_CODE = A.FD_CODE )
                        LEFT JOIN CMEPMS.dbo.CRDMNDS F ON F.TD_CODE = A.TD_CODE
                        --
                        LEFT JOIN CMETMS.dbo.TDDSDWB S1 ON S1.IT_CODE = 'TK103'  AND S1.TD_CODE = A.TD_CODE
                        LEFT JOIN CMETMS.dbo.TDDSDWB S2 ON S2.IT_CODE = 'TK105'  AND S2.TD_CODE = A.TD_CODE
                        LEFT JOIN CMETMS.dbo.TDDSDWB S3 ON S3.IT_CODE = 'DG106'  AND S3.TD_CODE = A.TD_CODE
                        LEFT JOIN CMETMS.dbo.TDDSDWB S4 ON S4.IT_CODE = 'TK107'  AND S4.TD_CODE = A.TD_CODE
                        LEFT JOIN CMETMS.dbo.TDDSDWB S6 ON S6.IT_CODE = 'TK102'  AND S6.TD_CODE = A.TD_CODE
                        LEFT JOIN CMETMS.dbo.TDDSDWB S7 ON S7.IT_CODE = 'TK1031' AND S7.TD_CODE = A.TD_CODE
                        LEFT JOIN CMETMS.dbo.TDDSDWB S10 ON S10.IT_CODE = 'TK108' AND S10.TD_CODE = A.TD_CODE
                        --
                        LEFT JOIN CMETMS.dbo.SMCCODE S8 ON S8.SYS_ID = 'TT' AND S8.CATY_ID = '03' AND S8.CATY_COD = A.CATY_CODE
                        LEFT JOIN CMETMS.dbo.SMCCODE S9 ON S9.SYS_ID = 'SM' AND S9.CATY_ID = '02' AND S9.CATY_COD = E.IPJ_CODE
                        LEFT JOIN ( SELECT TD_CODE, EXECT_DESC, PRE_SCHD, REL_SCHD, DELAY_RESN, CTDELAY_RESN FROM CMETMS.dbo.TDDMSCD WHERE SYR = '{YEAR2}' AND SMN = '{month}') M ON M.TD_CODE = A.TD_CODE
                        LEFT JOIN ( SELECT TD_CODE, EXECT_DESC, PRE_SCHD, REL_SCHD, DELAY_RESN, CTDELAY_RESN FROM CMETMS.dbo.TDDMSCD WHERE SYR = '{YEAR2}' AND SMN = '{month}') O ON O.TD_CODE = A.TD_CODE
                        WHERE A.PID IN ( {PIDString} ) AND D.TD_CITY IN ( {cityIn} ) ";

                Trace.WriteLine($"SQL:{SQL2}");

                using (SqlCommand cmd2 = new SqlCommand(SQL2, con))
                using (SqlDataAdapter da = new SqlDataAdapter(cmd2))
                {
                    da.Fill(ResultTable);
                }
            }

            #endregion ==============================================================================================

            return ResultTable;
        }

    }
}