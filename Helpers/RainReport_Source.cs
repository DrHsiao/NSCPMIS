using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Web.UI.WebControls;
using WCals.Controllers;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEPMSModels;

namespace WCals.Source
{
    public class RainReport_Source
    {
        static public CMBASEEntities db = new CMBASEEntities();
        static public CMEPMSEntities db2 = new CMEPMSEntities();

        /// <summary>
        /// =============== 舊版(已無使用) ===============
        /// </summary>
        /// <param name="AN_PID"></param>
        /// <param name="PJ_SYR"></param>
        /// <param name="PJ_SMN"></param>
        /// <param name="connectionString"></param>
        /// <param name="DEPT"></param>
        /// <returns></returns>
        public static DataTable GetRPT0003Data(decimal AN_PID, string PJ_SYR, string PJ_SMN, string connectionString, string DEPT)
        {
            string sql = "";
            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");
            DataTable dt = new DataTable();

            if (PJ_SYR != "")
            {
                sSYR = PJ_SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }
            if (PJ_SMN != "")
            {
                sSMN = PJ_SMN;
            }

            sql = "select P.SYR,C.TD_CODE";
            sql = sql + ",case len(C.TD_CODE)";
            sql = sql + "   when 19 then SUBSTRING(C.TD_CODE, 1, 3) +'-' + SUBSTRING(C.TD_CODE, 4, 4) + '-' + SUBSTRING(C.TD_CODE, 8, 4) + '-' + SUBSTRING(C.TD_CODE, 12, 4) + '-' + SUBSTRING(C.TD_CODE, 16, 4)";
            sql = sql + "   when 20 then SUBSTRING(C.TD_CODE, 1, 3) +'-' + SUBSTRING(C.TD_CODE, 4, 4) + '-' + SUBSTRING(C.TD_CODE, 8, 4) + '-' + SUBSTRING(C.TD_CODE, 12, 4) + '-' + SUBSTRING(C.TD_CODE, 16, 4)";
            sql = sql + "   when 11 then SUBSTRING(C.TD_CODE, 1, 2) +'-' + SUBSTRING(C.TD_CODE, 3, 4) + '-' + SUBSTRING(C.TD_CODE, 7, 3) + SUBSTRING(C.TD_CODE, 10, 5)";
            sql = sql + "   else SUBSTRING(C.TD_CODE, 1, 2) + '-' + SUBSTRING(C.TD_CODE, 3, 4) + '-' + SUBSTRING(C.TD_CODE, 7, 3) + '-' + SUBSTRING(C.TD_CODE, 10, 5)";
            sql = sql + "end SIP_TD_CODE";//--工程編號
            sql = sql + ",C.FD_CODE";
            sql = sql + ",case len(C.FD_CODE)";
            sql = sql + "   when 19 then SUBSTRING(C.FD_CODE, 1, 3) +'-' + SUBSTRING(C.FD_CODE, 4, 4) + '-' + SUBSTRING(C.FD_CODE, 8, 4) + '-' + SUBSTRING(C.FD_CODE, 12, 4) + '-' + SUBSTRING(C.FD_CODE, 16, 4)";
            sql = sql + "   when 20 then SUBSTRING(C.FD_CODE, 1, 3) +'-' + SUBSTRING(C.FD_CODE, 4, 4) + '-' + SUBSTRING(C.FD_CODE, 8, 4) + '-' + SUBSTRING(C.FD_CODE, 12, 4) + '-' + SUBSTRING(C.FD_CODE, 16, 4)";
            sql = sql + "   when 11 then SUBSTRING(C.TD_CODE, 1, 2) +'-' + SUBSTRING(C.FD_CODE, 3, 4) + '-' + SUBSTRING(C.FD_CODE, 7, 3) + SUBSTRING(C.FD_CODE, 10, 5)";
            sql = sql + "   else SUBSTRING(C.FD_CODE, 1, 2) + '-' + SUBSTRING(C.FD_CODE, 3, 4) + '-' + SUBSTRING(C.FD_CODE, 7, 3) + '-' + SUBSTRING(C.FD_CODE, 10, 5)";
            sql = sql + "end SIP_FD_CODE";//--經費編號
            sql = sql + ",T.TD_CITY,T.TD_TOWN,C.TD_NAME";
            sql = sql + ",isnull(PC.SUP_RATIO, (select case (F00.CEN_FUND + F01.LOC_FUND) when 0 then 100 else F00.CEN_FUND / (F00.CEN_FUND + F01.LOC_FUND) * 100.00 end)) as CEN_SP_RATIO_S ";//--補助比例
            sql = sql + ",F00.CEN_FUND as ALL_CEN_FUND";//--核定補助經費 ==> 歷年中央款合計
            sql = sql + ",F01.LOC_FUND as ALL_LOC_FUND";//--歷年地方款合計
            sql = sql + ",F02.CEN_FUND as PRE_CEN_FUND";//--分配預算 ==> 前一年度中央款
            sql = sql + ",isnull((SF1.COST_FEE + SF1.PPAY_FEE), 0) as PRE_PAY_FUND";//--撥款數 ==> 實支數 + 已撥保留款
            sql = sql + ",isnull(SF1.COST_FEE, 0) as PRE_COST_FEE";//--實支數
            sql = sql + ",isnull(SF1.PPAY_FEE, 0) as PRE_PPAY_FEE";//--已撥保留款
            sql = sql + ",isnull(R.OCR_FUND, 0) as PRE_OCR_FUND";//--發生權責數
            sql = sql + ",isnull(R.REV_FUND, 0) as PRE_REV_FUND";//--保留款中的權責數
            sql = sql + ",isnull(F04.CEN_FUND, 0) as CEN_FUND";//--年度分配預算 ==> 本年度中央款
            sql = sql + ",isnull(SF2.COST_FEE, 0) as COST_FEE";//--本年度實支數
            sql = sql + ",isnull(W1.REL_EDT, W1.PRE_EDT) as TK105_DT";//--發包後請款期程 ==> 實際發包日期(若無則為預訂)
            sql = sql + ",isnull(W2.PRE_EDT, '') as TK107_PRE_DT";//--預定完工日期
            sql = sql + ",isnull(W2.REL_EDT, '') as TK107_REL_DT";//--實際完工日期
            sql = sql + ",isnull(W4.REL_EDT, W3.REL_EDT) as TK110_REL_DT";//--結案 ==> 決算日期(若無則結算日期)
            sql = sql + ",isnull(M.PRE_SCHD, 0) as PRE_SCHD";
            sql = sql + ",isnull(M.REL_SCHD, 0) as REL_SCHD";
            sql = sql + ",isnull(G_B0.BUDGET_FEE, 0) as BUDGET_FEE_B0";//--總工程經費 ==> 預算款項的合計
            sql = sql + ",isnull(G_C0.BUDGET_FEE, 0) as BUDGET_FEE_C0";//--契約金額 ==> 契約款項的合計
            sql = sql + ",isnull(M.EXECT_DESC, '') as EXECT_DESC";//--辦理情形 ==> 輸出月份辦理情形
            sql = sql + ",case isnull(T.TD_CITY,'')";
            sql = sql + "   when '' then '00'";
            sql = sql + "   when '不分區' then '00'";
            sql = sql + "   when '台北市' then '01'";
            sql = sql + "   when '臺北市' then '01'";
            sql = sql + "   when '新北市' then '02'";
            sql = sql + "   when '台北縣' then '02'";
            sql = sql + "   when '臺北縣' then '02'";
            sql = sql + "   when '桃園市' then '03'";
            sql = sql + "   when '桃園縣' then '03'";
            sql = sql + "   when '基隆市' then '04'";
            sql = sql + "   when '宜蘭縣' then '05'";
            sql = sql + "   when '花蓮縣' then '06'";
            sql = sql + "   when '新竹縣' then '07'";
            sql = sql + "   when '連江縣' then '08'";
            sql = sql + "   when '臺中市' then '09'";
            sql = sql + "   when '台中市' then '09'";
            sql = sql + "   when '臺中縣' then '09'";
            sql = sql + "   when '台中縣' then '09'";
            sql = sql + "   when '新竹市' then '10'";
            sql = sql + "   when '苗栗縣' then '11'";
            sql = sql + "   when '南投縣' then '12'";
            sql = sql + "   when '彰化縣' then '13'";
            sql = sql + "   when '雲林縣' then '14'";
            sql = sql + "   when '臺南市' then '15'";
            sql = sql + "   when '台南市' then '15'";
            sql = sql + "   when '臺南縣' then '15'";
            sql = sql + "   when '台南縣' then '15'";
            sql = sql + "   when '高雄市' then '16'";
            sql = sql + "   when '高雄縣' then '16'";
            sql = sql + "   when '嘉義縣' then '17'";
            sql = sql + "   when '嘉義市' then '18'";
            sql = sql + "   when '屏東縣' then '19'";
            sql = sql + "   when '金門縣' then '20'";
            sql = sql + "   when '澎湖縣' then '21'";
            sql = sql + "   when '臺東縣' then '22'";
            sql = sql + "   when '台東縣' then '22'";
            sql = sql + "   else '99'";
            sql = sql + " end as TD_CITY_ORDER";
            sql = sql + " from CMETMS.dbo.TDMFORM T";
            sql = sql + " inner join CMEPMS.dbo.CTDMNDS C on C.TD_CODE = T.TD_CODE";
            sql = sql + " left join CMEPMS.dbo.PJDCTSP PC on PC.PID = C.PID and PC.CITY = T.TD_CITY";
            sql = sql + " left join CMEPMs.dbo.PJMFORM P on P.PID = C.PID";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' group by C.TD_CODE) F00 on F00.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as LOC_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE <> '00' group by C.TD_CODE) F01 on F01.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' and F.SYR = '" + sPSYR + "' group by C.TD_CODE) F02 on F02.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as LOC_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE <> '00'and F.SYR = '" + sPSYR + "' group by C.TD_CODE) F03 on F03.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' and F.SYR = '" + sSYR + "' group by C.TD_CODE) F04 on F04.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as LOC_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE <> '00'and F.SYR = '" + sSYR + "' group by C.TD_CODE) F05 on F05.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select PID, TD_CODE1 as TD_CODE,sum(REV_FUND) as REV_FUND,sum(OCR_FUND) as OCR_FUND from CMEPMS.dbo.CRDMNDS where SYR = '107' group by PID, TD_CODE1)  R on R.TD_CODE = C.TD_CODE";
            sql = sql + " left join (select TD_CODE1, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR = '" + sPSYR + "' and SMN = '12' group by TD_CODE1) SF1 on SF1.TD_CODE1 = C.TD_CODE ";
            sql = sql + " left join (select TD_CODE1, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.TDDSTCF F10 where SYR+SMN = (select MAX(SYR+SMN) from CMEPMS.dbo.TDDSTCF where TD_CODE = F10.TD_CODE and SYR+SMN <= '" + sSYR + sSMN + "') group by TD_CODE1) SF2 on SF2.TD_CODE1 = C.TD_CODE";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W1 on W1.TD_CODE = C.TD_CODE and W1.IT_CODE = 'TK105'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W2 on W2.TD_CODE = C.TD_CODE and W2.IT_CODE = 'TK107'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W3 on W3.TD_CODE = C.TD_CODE and W3.IT_CODE = 'TK109'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W4 on W4.TD_CODE = C.TD_CODE and W4.IT_CODE = 'TK110'";
            sql = sql + " left join(select TD_CODE, sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP = 'B' and SEQ_NO = '0' group by TD_CODE) G_B0 on G_B0.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select TD_CODE, sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP = 'C' and SEQ_NO = '0' group by TD_CODE) G_C0 on G_C0.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select TD_CODE, EXECT_DESC,PRE_SCHD,REL_SCHD from CMETMS.dbo.TDDMSCD where SYR= '" + sSYR + "' and SMN = '" + sSMN + "') M on M.TD_CODE = C.TD_CODE";
            /// 加上部門帳號查詢條件
            sql = sql + " where C.PID in (select PID from CMEPMS.dbo.PJMFORM where AN_PID='" + AN_PID + "' and SYR='" + sSYR + "'" + " and ADV_DP='" + DEPT + "'" + " AND Water_Type = '雨水使用' " + ")";

            dt = SourceFunction.DataTableGet(sql, connectionString); ;

            dt.DefaultView.Sort = "TD_CITY_ORDER";

            return dt.DefaultView.ToTable();
        }

        /// <summary> 取得前一年度的資料 </summary>
        /// <param name="PJ_NAMEs"></param>
        /// <param name="PJ_SYR"></param>
        /// <param name="PJ_SMN"></param>
        /// <param name="connectionString"></param>
        /// <param name="DEPT"></param>
        /// <returns></returns>
        public static DataTable GetRPT0003PrevData(string[] PJ_NAMEs, string PJ_SYR, string PJ_SMN, string connectionString, string DEPT)
        {
            Debug.WriteLine($"GetRPT0003PrevData");
            foreach (var item in PJ_NAMEs)
            {
                Debug.WriteLine($"PJ_NAMEs:{item} ");
            }

            Debug.WriteLine($"PJ_SYR:{PJ_SYR}");
            Debug.WriteLine($"PJ_SMN:{PJ_SMN}");
            Debug.WriteLine($"connectionString:{connectionString}");
            Debug.WriteLine($"DEPT:{DEPT}");

            if (PJ_NAMEs == null || PJ_NAMEs.Length == 0)
            {
                return new DataTable();
            }

            string PJ_NAME = string.Join(" OR ", PJ_NAMEs.Select(term => $"PJ_NAME LIKE '%{term}%'"));
            string sql = "";
            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sNSYR = (System.DateTime.Now.Year - 1911 + 1).ToString("000");
            DataTable dt = new DataTable();

            if (PJ_SYR != "")
            {
                sSYR = PJ_SYR;
                sNSYR = (Convert.ToInt32(sSYR) + 1).ToString("000");
            }
            if (PJ_SMN != "")
            {
                sSMN = PJ_SMN;
            }

            sql = "select P.SYR,C.TD_CODE";
            sql = sql + ",case len(C.TD_CODE)";
            sql = sql + "   when 19 then SUBSTRING(C.TD_CODE, 1, 3) +'-' + SUBSTRING(C.TD_CODE, 4, 4) + '-' + SUBSTRING(C.TD_CODE, 8, 4) + '-' + SUBSTRING(C.TD_CODE, 12, 4) + '-' + SUBSTRING(C.TD_CODE, 16, 4)";
            sql = sql + "   when 20 then SUBSTRING(C.TD_CODE, 1, 3) +'-' + SUBSTRING(C.TD_CODE, 4, 4) + '-' + SUBSTRING(C.TD_CODE, 8, 4) + '-' + SUBSTRING(C.TD_CODE, 12, 4) + '-' + SUBSTRING(C.TD_CODE, 16, 4)";
            sql = sql + "   when 11 then SUBSTRING(C.TD_CODE, 1, 2) +'-' + SUBSTRING(C.TD_CODE, 3, 4) + '-' + SUBSTRING(C.TD_CODE, 7, 3) + SUBSTRING(C.TD_CODE, 10, 5)";
            sql = sql + "   else SUBSTRING(C.TD_CODE, 1, 2) + '-' + SUBSTRING(C.TD_CODE, 3, 4) + '-' + SUBSTRING(C.TD_CODE, 7, 3) + '-' + SUBSTRING(C.TD_CODE, 10, 5)";
            sql = sql + "end SIP_TD_CODE"; /// 工程編號
            sql = sql + ",C.FD_CODE";
            sql = sql + ",case len(C.FD_CODE)";
            sql = sql + "   when 19 then SUBSTRING(C.FD_CODE, 1, 3) +'-' + SUBSTRING(C.FD_CODE, 4, 4) + '-' + SUBSTRING(C.FD_CODE, 8, 4) + '-' + SUBSTRING(C.FD_CODE, 12, 4) + '-' + SUBSTRING(C.FD_CODE, 16, 4)";
            sql = sql + "   when 20 then SUBSTRING(C.FD_CODE, 1, 3) +'-' + SUBSTRING(C.FD_CODE, 4, 4) + '-' + SUBSTRING(C.FD_CODE, 8, 4) + '-' + SUBSTRING(C.FD_CODE, 12, 4) + '-' + SUBSTRING(C.FD_CODE, 16, 4)";
            sql = sql + "   when 11 then SUBSTRING(C.TD_CODE, 1, 2) +'-' + SUBSTRING(C.FD_CODE, 3, 4) + '-' + SUBSTRING(C.FD_CODE, 7, 3) + SUBSTRING(C.FD_CODE, 10, 5)";
            sql = sql + "   else SUBSTRING(C.FD_CODE, 1, 2) + '-' + SUBSTRING(C.FD_CODE, 3, 4) + '-' + SUBSTRING(C.FD_CODE, 7, 3) + '-' + SUBSTRING(C.FD_CODE, 10, 5)";
            sql = sql + "end SIP_FD_CODE"; /// 經費編號
            sql = sql + ",T.TD_CITY,T.TD_TOWN,C.TD_NAME";
            //sql = sql + ",isnull(PC.SUP_RATIO, (select case (F00.CEN_FUND + F01.LOC_FUND) when 0 then 100 else F00.CEN_FUND / (F00.CEN_FUND + F01.LOC_FUND) * 100.00 end)) as CEN_SP_RATIO_S ";//--補助比例
            sql = sql + ",CASE D.YR_FUND WHEN 0 THEN 100 ELSE CONVERT(decimal(10, 2), (D.CEN_FUND) / D.YR_FUND) * 100 END AS CEN_SP_RATIO_S"; //v241218修正：補助比例抓取中央款比例(運算)
            sql = sql + ",F00.CEN_FUND as ALL_CEN_FUND";//--核定補助經費 ==> 歷年中央款合計
            sql = sql + ",F01.LOC_FUND as ALL_LOC_FUND";//--歷年地方款合計
            sql = sql + ",isnull(F02.CEN_FUND, 0) as PRE_CEN_FUND";//--分配預算 ==> 前一年度中央款 //-k 2412-1
            sql = sql + ",isnull((SF1.COST_FEE + SF1.PPAY_FEE), 0) as PRE_PAY_FUND";//--撥款數 ==> 實支數 + 已撥保留款
            sql = sql + ",isnull(SF1.COST_FEE, 0) as PRE_COST_FEE";//--實支數
            sql = sql + ",isnull(SF1.PPAY_FEE, 0) as PRE_PPAY_FEE";//--已撥保留款
            sql = sql + ",isnull(R.OCR_FUND, 0) as PRE_OCR_FUND";//--發生權責數
            sql = sql + ",isnull(R.REV_FUND, 0) as PRE_REV_FUND";//--保留款中的權責數
            //sql = sql + ",isnull(F04.CEN_FUND, 0) as CEN_FUND";//--年度分配預算 ==> 本年度中央款
            sql = sql + ",isnull(SF2.COST_FEE, 0) as COST_FEE";//--本年度實支數
            sql = sql + ",isnull(W1.REL_EDT, W1.PRE_EDT) as TK105_DT";//--發包後請款期程 ==> 實際發包日期(若無則為預訂)
            sql = sql + ",isnull(W2.PRE_EDT, '') as TK107_PRE_DT";//--預定完工日期
            sql = sql + ",isnull(W2.REL_EDT, '') as TK107_REL_DT";//--實際完工日期
            sql = sql + ",isnull(W4.REL_EDT, W3.REL_EDT) as TK110_REL_DT";//--結案 ==> 決算日期(若無則結算日期)
            sql = sql + ",isnull(W5.REL_EDT, W5.PRE_EDT) as TK141_DT"; //v2412-3修正：加入「進度達40%期程」資料 => 若無實際則為預定
            sql = sql + ",isnull(M.PRE_SCHD, 0) as PRE_SCHD";
            sql = sql + ",isnull(M.REL_SCHD, 0) as REL_SCHD";
            sql = sql + ",isnull(G_B0.BUDGET_FEE, 0) as BUDGET_FEE_B0";//--總工程經費 ==> 預算款項的合計
            sql = sql + ",isnull(G_C0.BUDGET_FEE, 0) as BUDGET_FEE_C0";//--契約金額 ==> 契約款項的合計
            sql = sql + ",isnull(M.EXECT_DESC, '') as EXECT_DESC";//--辦理情形 ==> 輸出月份辦理情形
            sql = sql + ",case isnull(T.TD_CITY,'')";
            sql = sql + "   when '' then '00'";
            sql = sql + "   when '不分區' then '00'";
            sql = sql + "   when '台北市' then '01'";
            sql = sql + "   when '臺北市' then '01'";
            sql = sql + "   when '新北市' then '02'";
            sql = sql + "   when '台北縣' then '02'";
            sql = sql + "   when '臺北縣' then '02'";
            sql = sql + "   when '桃園市' then '03'";
            sql = sql + "   when '桃園縣' then '03'";
            sql = sql + "   when '基隆市' then '04'";
            sql = sql + "   when '宜蘭縣' then '05'";
            sql = sql + "   when '花蓮縣' then '06'";
            sql = sql + "   when '新竹縣' then '07'";
            sql = sql + "   when '連江縣' then '08'";
            sql = sql + "   when '臺中市' then '09'";
            sql = sql + "   when '台中市' then '09'";
            sql = sql + "   when '臺中縣' then '09'";
            sql = sql + "   when '台中縣' then '09'";
            sql = sql + "   when '新竹市' then '10'";
            sql = sql + "   when '苗栗縣' then '11'";
            sql = sql + "   when '南投縣' then '12'";
            sql = sql + "   when '彰化縣' then '13'";
            sql = sql + "   when '雲林縣' then '14'";
            sql = sql + "   when '臺南市' then '15'";
            sql = sql + "   when '台南市' then '15'";
            sql = sql + "   when '臺南縣' then '15'";
            sql = sql + "   when '台南縣' then '15'";
            sql = sql + "   when '高雄市' then '16'";
            sql = sql + "   when '高雄縣' then '16'";
            sql = sql + "   when '嘉義縣' then '17'";
            sql = sql + "   when '嘉義市' then '18'";
            sql = sql + "   when '屏東縣' then '19'";
            sql = sql + "   when '金門縣' then '20'";
            sql = sql + "   when '澎湖縣' then '21'";
            sql = sql + "   when '臺東縣' then '22'";
            sql = sql + "   when '台東縣' then '22'";
            sql = sql + "   else '99'";
            sql = sql + " end as TD_CITY_ORDER";
            sql = sql + " ,C.CATY_CODE"; /// 分項內容編號
            sql = sql + " from CMETMS.dbo.TDMFORM T";
            sql = sql + " inner join CMEPMS.dbo.CTDMNDS C on C.TD_CODE = T.TD_CODE";
            sql = sql + " left join CMEPMS.dbo.PJDCTSP PC on PC.PID = C.PID and PC.CITY = T.TD_CITY";
            sql = sql + " left join CMEPMs.dbo.PJMFORM P on P.PID = C.PID";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' group by C.TD_CODE) F00 on F00.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as LOC_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE <> '00' group by C.TD_CODE) F01 on F01.TD_CODE = C.TD_CODE";
            //sql = sql + " left join(select C.TD_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' and F.SYR = '" + sSYR + "' group by C.TD_CODE) F02 on F02.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, C.CATY_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' and F.SYR = '" + sSYR + "' group by C.TD_CODE, C.CATY_CODE) F02 on F02.TD_CODE = C.TD_CODE AND F02.CATY_CODE = C.CATY_CODE"; //加入分項內容編號區別不同經費 //-k 2412-1
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as LOC_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE <> '00'and F.SYR = '" + sSYR + "' group by C.TD_CODE) F03 on F03.TD_CODE = C.TD_CODE";
            //sql = sql + " left join(select C.TD_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' and F.SYR = '" + sNSYR + "' group by C.TD_CODE) F04 on F04.TD_CODE = C.TD_CODE";
            //sql = sql + " left join(select C.TD_CODE, sum(FUND) as LOC_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE <> '00'and F.SYR = '" + sNSYR + "' group by C.TD_CODE) F05 on F05.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select PID, TD_CODE1 as TD_CODE,sum(REV_FUND) as REV_FUND,sum(OCR_FUND) as OCR_FUND from CMEPMS.dbo.CRDMNDS where SYR = '107' group by PID, TD_CODE1)  R on R.TD_CODE = C.TD_CODE";
            //sql = sql + " left join (select TD_CODE1, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR = '" + sSYR + "' and SMN = '12' group by TD_CODE1) SF1 on SF1.TD_CODE1 = C.TD_CODE ";
            sql = sql + " left join (select C.TD_CODE, C.CATY_CODE, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.TDDSTCF F on C.FD_CODE = F.FD_CODE and F.SYR = '" + sSYR + "' and F.SMN = '12' group by C.TD_CODE, C.CATY_CODE) SF1 on SF1.TD_CODE = C.TD_CODE AND SF1.CATY_CODE = C.CATY_CODE"; //加入分項內容編號區別不同經費 //-k 2412-1
            //sql = sql + " left join (select TD_CODE1, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.TDDSTCF F10 where SYR+SMN = (select MAX(SYR+SMN) from CMEPMS.dbo.TDDSTCF where TD_CODE = F10.TD_CODE and SYR+SMN <= '" + sNSYR + sSMN + "') group by TD_CODE1) SF2 on SF2.TD_CODE1 = C.TD_CODE";
            sql = sql + " left join (select C.TD_CODE, C.CATY_CODE, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.TDDSTCF F10 on C.FD_CODE = F10.FD_CODE and F10.SYR + F10.SMN = (select MAX(SYR + SMN) from CMEPMS.dbo.TDDSTCF where TD_CODE = F10.TD_CODE and SYR+SMN <= '" + sNSYR + sSMN + "') group by C.TD_CODE, C.CATY_CODE) SF2 on SF2.TD_CODE = C.TD_CODE  AND SF2.CATY_CODE = C.CATY_CODE";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W1 on W1.TD_CODE = C.TD_CODE and W1.IT_CODE = 'TK105'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W2 on W2.TD_CODE = C.TD_CODE and W2.IT_CODE = 'TK107'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W3 on W3.TD_CODE = C.TD_CODE and W3.IT_CODE = 'TK109'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W4 on W4.TD_CODE = C.TD_CODE and W4.IT_CODE = 'TK110'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W5 on W5.TD_CODE = C.TD_CODE and W5.IT_CODE = 'TK141'"; //v2412-3修正：加入「進度達40%期程」資料
            sql = sql + " left join(select TD_CODE, sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP = 'B' and SEQ_NO = '0' group by TD_CODE) G_B0 on G_B0.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select TD_CODE, sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP = 'C' and SEQ_NO = '0' group by TD_CODE) G_C0 on G_C0.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select TD_CODE, EXECT_DESC,PRE_SCHD,REL_SCHD from CMETMS.dbo.TDDMSCD where SYR= '" + sSYR + "' and SMN = '" + sSMN + "') M on M.TD_CODE = C.TD_CODE";
            sql = sql + " left join(SELECT TID, SUM(YR_FUND) AS YR_FUND, SUM(CEN_FUND) AS CEN_FUND, SUM(LOC_FUND) AS LOC_FUND FROM (SELECT TID, SUM(FUND) AS YR_FUND, CASE WHEN LEFT(DSum.FIT_CODE, 1) = '0' THEN SUM(IsNull(DSum.FUND, 0)) ELSE 0 END AS CEN_FUND, CASE WHEN LEFT(DSum.FIT_CODE, 1) = '1' THEN SUM(IsNull(DSum.FUND, 0)) ELSE 0 END AS LOC_FUND FROM cmepms.dbo.CTDITFD AS DSum GROUP BY TID, FIT_CODE) AS D1 GROUP BY TID) AS D ON D.TID = C.TID";
            /// 加上部門帳號查詢條件
            /// 20241218 修改項目 => 補助比例抓取中央款比例(運算)
            sql = sql + " where C.PID in (select PID from CMEPMS.dbo.PJMFORM where (" + PJ_NAME + ") and SYR='" + sSYR + "'" + " and ADV_DP='" + DEPT + "'" + ")";

            Debug.WriteLine($"GetRPT0003PrevData:{sql}");

            dt = SourceFunction.DataTableGet(sql, connectionString); ;

            dt.DefaultView.Sort = "TD_CITY_ORDER";



            return dt.DefaultView.ToTable();
        }

        /// <summary> 取得當年度的資料 </summary>
        /// <param name="AN_PID"></param>
        /// <param name="PJ_SYR"></param>
        /// <param name="PJ_SMN"></param>
        /// <param name="connectionString"></param>
        /// <param name="DEPT"></param>
        /// <returns></returns>
        public static DataTable GetRPT0003CurData(decimal AN_PID, string PJ_SYR, string PJ_SMN, string connectionString, string DEPT)
        {
            Debug.WriteLine($"GetRPT0003CurData");
            Debug.WriteLine($"AN_PID:{AN_PID}");
            Debug.WriteLine($"PJ_SYR:{PJ_SYR}");
            Debug.WriteLine($"PJ_SMN:{PJ_SMN}");
            Debug.WriteLine($"connectionString:{connectionString}");
            Debug.WriteLine($"DEPT:{DEPT}");

            string sql = "";
            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");
            DataTable dt = new DataTable();

            if (PJ_SYR != "")
            {
                sSYR = PJ_SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }
            if (PJ_SMN != "")
            {
                sSMN = PJ_SMN;
            }

            sql = "select P.SYR,C.TD_CODE";
            sql = sql + ",case len(C.TD_CODE)";
            sql = sql + "   when 19 then SUBSTRING(C.TD_CODE, 1, 3) +'-' + SUBSTRING(C.TD_CODE, 4, 4) + '-' + SUBSTRING(C.TD_CODE, 8, 4) + '-' + SUBSTRING(C.TD_CODE, 12, 4) + '-' + SUBSTRING(C.TD_CODE, 16, 4)";
            sql = sql + "   when 20 then SUBSTRING(C.TD_CODE, 1, 3) +'-' + SUBSTRING(C.TD_CODE, 4, 4) + '-' + SUBSTRING(C.TD_CODE, 8, 4) + '-' + SUBSTRING(C.TD_CODE, 12, 4) + '-' + SUBSTRING(C.TD_CODE, 16, 4)";
            sql = sql + "   when 11 then SUBSTRING(C.TD_CODE, 1, 2) +'-' + SUBSTRING(C.TD_CODE, 3, 4) + '-' + SUBSTRING(C.TD_CODE, 7, 3) + SUBSTRING(C.TD_CODE, 10, 5)";
            sql = sql + "   else SUBSTRING(C.TD_CODE, 1, 2) + '-' + SUBSTRING(C.TD_CODE, 3, 4) + '-' + SUBSTRING(C.TD_CODE, 7, 3) + '-' + SUBSTRING(C.TD_CODE, 10, 5)";
            sql = sql + "end SIP_TD_CODE";//--工程編號
            sql = sql + ",C.FD_CODE";
            sql = sql + ",case len(C.FD_CODE)";
            sql = sql + "   when 19 then SUBSTRING(C.FD_CODE, 1, 3) +'-' + SUBSTRING(C.FD_CODE, 4, 4) + '-' + SUBSTRING(C.FD_CODE, 8, 4) + '-' + SUBSTRING(C.FD_CODE, 12, 4) + '-' + SUBSTRING(C.FD_CODE, 16, 4)";
            sql = sql + "   when 20 then SUBSTRING(C.FD_CODE, 1, 3) +'-' + SUBSTRING(C.FD_CODE, 4, 4) + '-' + SUBSTRING(C.FD_CODE, 8, 4) + '-' + SUBSTRING(C.FD_CODE, 12, 4) + '-' + SUBSTRING(C.FD_CODE, 16, 4)";
            sql = sql + "   when 11 then SUBSTRING(C.TD_CODE, 1, 2) +'-' + SUBSTRING(C.FD_CODE, 3, 4) + '-' + SUBSTRING(C.FD_CODE, 7, 3) + SUBSTRING(C.FD_CODE, 10, 5)";
            sql = sql + "   else SUBSTRING(C.FD_CODE, 1, 2) + '-' + SUBSTRING(C.FD_CODE, 3, 4) + '-' + SUBSTRING(C.FD_CODE, 7, 3) + '-' + SUBSTRING(C.FD_CODE, 10, 5)";
            sql = sql + "end SIP_FD_CODE";//--經費編號
            sql = sql + ",T.TD_CITY,T.TD_TOWN,C.TD_NAME";
            //sql = sql + ",isnull(PC.SUP_RATIO, (select case (F00.CEN_FUND + F01.LOC_FUND) when 0 then 100 else F00.CEN_FUND / (F00.CEN_FUND + F01.LOC_FUND) * 100.00 end)) as CEN_SP_RATIO_S ";//--補助比例
            sql = sql + ",CASE D.YR_FUND WHEN 0 THEN 100 ELSE CONVERT(decimal(10, 2), (D.CEN_FUND) / D.YR_FUND) * 100 END AS CEN_SP_RATIO_S"; //v2412-3修正：補助比例 => 抓取中央款比例(運算)
            sql = sql + ",F00.CEN_FUND as ALL_CEN_FUND";//--核定補助經費 ==> 歷年中央款合計
            sql = sql + ",F01.LOC_FUND as ALL_LOC_FUND";//--歷年地方款合計
            sql = sql + ",isnull(F02.CEN_FUND, 0) as PRE_CEN_FUND";//--分配預算 ==> 前一年度中央款 //-k 2412-1
            sql = sql + ",isnull((SF1.COST_FEE + SF1.PPAY_FEE), 0) as PRE_PAY_FUND";//--撥款數 ==> 實支數 + 已撥保留款
            sql = sql + ",isnull(SF1.COST_FEE, 0) as PRE_COST_FEE";//--實支數
            sql = sql + ",isnull(SF1.PPAY_FEE, 0) as PRE_PPAY_FEE";//--已撥保留款
            sql = sql + ",isnull(R.OCR_FUND, 0) as PRE_OCR_FUND";//--發生權責數
            sql = sql + ",isnull(R.REV_FUND, 0) as PRE_REV_FUND";//--保留款中的權責數
            sql = sql + ",isnull(F04.CEN_FUND, 0) as CEN_FUND";//--年度分配預算 ==> 本年度中央款
            sql = sql + ",isnull(SF2.COST_FEE, 0) as COST_FEE";//--本年度實支數
            sql = sql + ",isnull(W1.REL_EDT, W1.PRE_EDT) as TK105_DT";//--發包後請款期程 ==> 實際發包日期(若無則為預訂)
            sql = sql + ",isnull(W2.PRE_EDT, '') as TK107_PRE_DT";//--預定完工日期
            sql = sql + ",isnull(W2.REL_EDT, '') as TK107_REL_DT";//--實際完工日期
            sql = sql + ",isnull(W4.REL_EDT, W3.REL_EDT) as TK110_REL_DT";//--結案 ==> 決算日期(若無則結算日期)
            sql = sql + ",isnull(W5.REL_EDT, W5.PRE_EDT) as TK141_DT"; //v2412-3修正：加入「進度達40%期程」資料 => 若無實際則為預定
            sql = sql + ",isnull(M.PRE_SCHD, 0) as PRE_SCHD";
            sql = sql + ",isnull(M.REL_SCHD, 0) as REL_SCHD";
            sql = sql + ",isnull(G_B0.BUDGET_FEE, 0) as BUDGET_FEE_B0";//--總工程經費 ==> 預算款項的合計
            sql = sql + ",isnull(G_C0.BUDGET_FEE, 0) as BUDGET_FEE_C0";//--契約金額 ==> 契約款項的合計
            sql = sql + ",isnull(M.EXECT_DESC, '') as EXECT_DESC";//--辦理情形 ==> 輸出月份辦理情形
            sql = sql + ",case isnull(T.TD_CITY,'')";
            sql = sql + "   when '' then '00'";
            sql = sql + "   when '不分區' then '00'";
            sql = sql + "   when '台北市' then '01'";
            sql = sql + "   when '臺北市' then '01'";
            sql = sql + "   when '新北市' then '02'";
            sql = sql + "   when '台北縣' then '02'";
            sql = sql + "   when '臺北縣' then '02'";
            sql = sql + "   when '桃園市' then '03'";
            sql = sql + "   when '桃園縣' then '03'";
            sql = sql + "   when '基隆市' then '04'";
            sql = sql + "   when '宜蘭縣' then '05'";
            sql = sql + "   when '花蓮縣' then '06'";
            sql = sql + "   when '新竹縣' then '07'";
            sql = sql + "   when '連江縣' then '08'";
            sql = sql + "   when '臺中市' then '09'";
            sql = sql + "   when '台中市' then '09'";
            sql = sql + "   when '臺中縣' then '09'";
            sql = sql + "   when '台中縣' then '09'";
            sql = sql + "   when '新竹市' then '10'";
            sql = sql + "   when '苗栗縣' then '11'";
            sql = sql + "   when '南投縣' then '12'";
            sql = sql + "   when '彰化縣' then '13'";
            sql = sql + "   when '雲林縣' then '14'";
            sql = sql + "   when '臺南市' then '15'";
            sql = sql + "   when '台南市' then '15'";
            sql = sql + "   when '臺南縣' then '15'";
            sql = sql + "   when '台南縣' then '15'";
            sql = sql + "   when '高雄市' then '16'";
            sql = sql + "   when '高雄縣' then '16'";
            sql = sql + "   when '嘉義縣' then '17'";
            sql = sql + "   when '嘉義市' then '18'";
            sql = sql + "   when '屏東縣' then '19'";
            sql = sql + "   when '金門縣' then '20'";
            sql = sql + "   when '澎湖縣' then '21'";
            sql = sql + "   when '臺東縣' then '22'";
            sql = sql + "   when '台東縣' then '22'";
            sql = sql + "   else '99'";
            sql = sql + " end as TD_CITY_ORDER";
            sql = sql + " ,C.CATY_CODE"; //分項內容編號 //-k 2412-1
            sql = sql + " from CMETMS.dbo.TDMFORM T";
            sql = sql + " inner join CMEPMS.dbo.CTDMNDS C on C.TD_CODE = T.TD_CODE";
            sql = sql + " left join CMEPMS.dbo.PJDCTSP PC on PC.PID = C.PID and PC.CITY = T.TD_CITY";
            sql = sql + " left join CMEPMs.dbo.PJMFORM P on P.PID = C.PID";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' group by C.TD_CODE) F00 on F00.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as LOC_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE <> '00' group by C.TD_CODE) F01 on F01.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, C.CATY_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' and F.SYR = '" + sPSYR + "' group by C.TD_CODE, C.CATY_CODE) F02 on F02.TD_CODE = C.TD_CODE AND F02.CATY_CODE = C.CATY_CODE"; //加入分項內容編號區別不同經費 //-k 2412-1
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as LOC_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE <> '00'and F.SYR = '" + sPSYR + "' group by C.TD_CODE) F03 on F03.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select C.TD_CODE, C.CATY_CODE, sum(FUND) as CEN_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE = '00' and F.SYR = '" + sSYR + "' group by C.TD_CODE, C.CATY_CODE) F04 on F04.TD_CODE = C.TD_CODE AND F04.CATY_CODE = C.CATY_CODE"; //加入分項內容編號區別不同經費 //-k 2412-1
            sql = sql + " left join(select C.TD_CODE, sum(FUND) as LOC_FUND from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.CTDITFD F on C.TID = F.TID and F.FIT_CODE <> '00'and F.SYR = '" + sSYR + "' group by C.TD_CODE) F05 on F05.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select PID, TD_CODE1 as TD_CODE,sum(REV_FUND) as REV_FUND,sum(OCR_FUND) as OCR_FUND from CMEPMS.dbo.CRDMNDS where SYR = '107' group by PID, TD_CODE1)  R on R.TD_CODE = C.TD_CODE";
            //sql = sql + " left join (select TD_CODE1, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.TDDSTCF where SYR = '" + sPSYR + "' and SMN = '12' group by TD_CODE1) SF1 on SF1.TD_CODE1 = C.TD_CODE ";
            sql = sql + " left join (select C.TD_CODE, C.CATY_CODE, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.TDDSTCF F on C.FD_CODE = F.FD_CODE and F.SYR = '" + sPSYR + "' and F.SMN = '12' group by C.TD_CODE, C.CATY_CODE) SF1 on SF1.TD_CODE = C.TD_CODE AND SF1.CATY_CODE = C.CATY_CODE"; //加入分項內容編號區別不同經費 //-k 2412-1
            //sql = sql + " left join (select TD_CODE1, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.TDDSTCF F10 where SYR+SMN = (select MAX(SYR+SMN) from CMEPMS.dbo.TDDSTCF where TD_CODE = F10.TD_CODE and SYR+SMN <= '" + sSYR + sSMN + "') group by TD_CODE1) SF2 on SF2.TD_CODE1 = C.TD_CODE";
            sql = sql + " left join (select C.TD_CODE, C.CATY_CODE, sum(COST_FEE) as COST_FEE,sum(PPAY_FEE) as PPAY_FEE from CMEPMS.dbo.CTDMNDS C left join CMEPMS.dbo.TDDSTCF F10 on C.FD_CODE = F10.FD_CODE and F10.SYR + F10.SMN = (select MAX(SYR + SMN) from CMEPMS.dbo.TDDSTCF where TD_CODE = F10.TD_CODE1 and SYR+SMN <= '" + sSYR + sSMN + "') group by C.TD_CODE, C.CATY_CODE) SF2 on SF2.TD_CODE = C.TD_CODE  AND SF2.CATY_CODE = C.CATY_CODE";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W1 on W1.TD_CODE = C.TD_CODE and W1.IT_CODE = 'TK105'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W2 on W2.TD_CODE = C.TD_CODE and W2.IT_CODE = 'TK107'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W3 on W3.TD_CODE = C.TD_CODE and W3.IT_CODE = 'TK109'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W4 on W4.TD_CODE = C.TD_CODE and W4.IT_CODE = 'TK110'";
            sql = sql + " left join CMETMS.dbo.TDDSDWB W5 on W5.TD_CODE = C.TD_CODE and W5.IT_CODE = 'TK141'"; //v2412-3修正：加入「進度達40%期程」資料
            sql = sql + " left join(select TD_CODE, sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP = 'B' and SEQ_NO = '0' group by TD_CODE) G_B0 on G_B0.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select TD_CODE, sum(BUDGET_FEE) as BUDGET_FEE from CMETMS.dbo.TDDBGLS where REC_TP = 'C' and SEQ_NO = '0' group by TD_CODE) G_C0 on G_C0.TD_CODE = C.TD_CODE";
            sql = sql + " left join(select TD_CODE, EXECT_DESC,PRE_SCHD,REL_SCHD from CMETMS.dbo.TDDMSCD where SYR= '" + sSYR + "' and SMN = '" + sSMN + "') M on M.TD_CODE = C.TD_CODE";
            /// 20250422 修改項目 => 補助比例:抓取中央款比例(運算)
            sql = sql + " left join(SELECT TID, SUM(YR_FUND) AS YR_FUND, SUM(CEN_FUND) AS CEN_FUND, SUM(LOC_FUND) AS LOC_FUND FROM (SELECT TID, SUM(FUND) AS YR_FUND, CASE WHEN LEFT(DSum.FIT_CODE, 1) = '0' THEN SUM(IsNull(DSum.FUND, 0)) ELSE 0 END AS CEN_FUND, CASE WHEN LEFT(DSum.FIT_CODE, 1) = '1' THEN SUM(IsNull(DSum.FUND, 0)) ELSE 0 END AS LOC_FUND FROM cmepms.dbo.CTDITFD AS DSum GROUP BY TID, FIT_CODE) AS D1 GROUP BY TID) AS D ON D.TID = C.TID";

            /// 加上部門帳號查詢條件
            sql = sql + " where C.PID in (select PID from CMEPMS.dbo.PJMFORM where AN_PID='" + AN_PID + "' and SYR='" + sSYR + "'" + " and ADV_DP='" + DEPT + "'" + ")";

            Debug.WriteLine($"GetRPT0003CurData:{sql}");

            dt = SourceFunction.DataTableGet(sql, connectionString); ;

            dt.DefaultView.Sort = "TD_CITY_ORDER, TD_CODE";



            return dt.DefaultView.ToTable();
        }

        public static string getDtColumns_SUM(DataTable dt, string columns, string sWhere)
        {
            string rtn = "0";
            double sVal = 0;
            int i;
            DataRow[] rr;

            rr = dt.Select(sWhere);
            for (i = 0; i < rr.Length; i++)
            {
                try
                {
                    sVal = sVal + Convert.ToDouble(rr[i][columns].ToString());
                }
                catch
                {

                }
            }
            rtn = sVal.ToString();
            return rtn;
        }

        /// <summary>
        /// I-1.雨水例行性管考報表(資料來源)
        /// </summary>
        /// <param name="AN_PID"> 工程編號 </param>
        /// <param name="PJ_SYR"> 年份 </param>
        /// <param name="PJ_SMN"> 月份 </param>
        /// <param name="connectionString"> 連接字串 </param>
        /// <param name="DEPT"> 部門 </param>
        public static DataSet GetRPT0001Data(decimal AN_PID, string PJ_SYR, string PJ_SMN, string connectionString, string DEPT)
        {
            Trace.WriteLine($"AN_PID:{AN_PID}");
            Trace.WriteLine($"PJ_SYR:{PJ_SYR}");
            Trace.WriteLine($"PJ_SMN:{PJ_SMN}");
            Trace.WriteLine($"DEPT:{DEPT}");
            Trace.WriteLine($"ConnectionString:{connectionString}");

            string OpID = DEPT;
            string sql = "";

            DataSet ds = new DataSet();          // 儲存最終結果的 DataSet
            DataTable dt = new DataTable();      // 一般用途的 DataTable
            DataTable dtp = new DataTable();     // 多用途的 DataTable
            DataTable dtp1 = new DataTable();    // 署辦(IPJ_CODE='01' and ACC_ID<>'0200')
            DataTable dtp2 = new DataTable();    // 地辦(IPJ_CODE='02')
            DataTable dtp3 = new DataTable();    // 業務費(IPJ_CODE='01' and ACC_ID='0200')
            DataTable dtp4 = new DataTable();    // 經費調整表-明細
            DataTable dtp5 = new DataTable();    // 經費調整表
            DataTable PJDCATY = new DataTable(); // 計畫類別資料表
            DataTable dt1 = new DataTable();     // 本期資料
            DataTable dt2 = new DataTable();     // 前期資料
            DataTable dt3 = new DataTable();     // 前一年度資料
            DataTable dt4 = new DataTable();     // 未來使用的 DataTable
            DataTable ditdt = new DataTable();   // 儲存臨時資料的 DataTable

            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");      // 取得民國年
            string sSMN = (System.DateTime.Now.Month).ToString("00");             // 取得月份
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000"); // 取得前一年度
            string sSEQ_NO = "0"; // 預設最大序列號
            string sPSEQ_NO = ""; // 預設次大序列號

            DataRow[] rr;
            DataRow[] rr1;
            DataRow dr;
            DataRow r;

            int i;
            int j;
            int k;

            string tempCITY = "";
            string sWhere = "";

            #region ---------------- 1.PJ_SYR(年度) PJ_SMN(月份)
            // 傳入參數 PJ_SYR 不為 null，該值覆蓋當前年度，並取得前一年度
            // 傳入參數 PJ_SMN 不為 null，該值覆蓋當前月份

            if (PJ_SYR != "")
            {
                sSYR = PJ_SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }

            if (PJ_SMN != "")
            {
                sSMN = PJ_SMN;
            }

            #endregion -------------

            #region ---------------- 2.Plan = AN_PID 查詢 PJMFORM 結果
            // 使用 AN_PID 查詢 PJMFORM 資料表，取得對應的計畫數據。

            var plan = db2.PJMFORM.Where(m => m.AN_PID == AN_PID).FirstOrDefault();

            #endregion -------------

            #region ---------------- 3.PJDCATY 和 DT

            /// 如果是分支計畫的情況
            if (plan is null)
            {
                sql = "select PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC from PJV_PJDCATY where PID in (select PID from PJMFORM where PID=" + AN_PID + ")";
                PJDCATY = SourceFunction.DataTableGet(sql, connectionString);

                sql = "select max(SEQ_NO) as MAX_SEQ_NO,max(SEQ_NO)-1 as pMAX_SEQ_NO,min(SEQ_NO) as MIN_SEQ_NO from PJDFDCR where PID in (select PID from PJMFORM where PID = " + AN_PID + " and ADV_DP = '" + OpID + "' and SYR='" + sSYR + "') and case isnull(APROV_DT, isnull(AP_DT, '')) when '' then '' else isnull(APROV_DT, isnull(AP_DT, '')) end <> '' and isnull(APROV_DT, isnull(AP_DT,''))< '" + sSYR + sSMN + "32'";
                dt = SourceFunction.DataTableGet(sql, connectionString);
            }
            else
            {
                sql = "select PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC from PJV_PJDCATY where PID in (select PID from PJMFORM where PID=" + AN_PID + ")";
                PJDCATY = SourceFunction.DataTableGet(sql, connectionString);

                sql = "select max(SEQ_NO) as MAX_SEQ_NO,max(SEQ_NO)-1 as pMAX_SEQ_NO,min(SEQ_NO) as MIN_SEQ_NO from PJDFDCR where PID in (select PID from PJMFORM where PID = " + AN_PID + " and ADV_DP = '" + OpID + "' and SYR='" + sSYR + "') and case isnull(APROV_DT, isnull(AP_DT, '')) when '' then '' else isnull(APROV_DT, isnull(AP_DT, '')) end <> '' and isnull(APROV_DT, isnull(AP_DT,''))< '" + sSYR + sSMN + "32'";
                dt = SourceFunction.DataTableGet(sql, connectionString);
            }

            #endregion -------------

            #region ---------------- 4.MAX_SEQ_NO(最大) / pMAX_SEQ_NO(次大)
            /// sSEQ_NO  => MAX_SEQ_NO(最大)
            /// sPSEQ_NO => pMAX_SEQ_NO(次大)

            if (dt != null && dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["MAX_SEQ_NO"].ToString() != "")
                {
                    sSEQ_NO = dt.Rows[0]["MAX_SEQ_NO"].ToString();
                }

                if (dt.Rows[0]["pMAX_SEQ_NO"].ToString() != "")
                {
                    sPSEQ_NO = dt.Rows[0]["pMAX_SEQ_NO"].ToString();
                }
            }

            #endregion -------------

            #region  DT1 DataTable 本期資料

            dt1 = new DataTable();

            if (plan is null)
            {
                #region SQL 查詢語句(舊版查詢)
                // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDCFTS where PID in (select PID from PJMFORM where PID=" + AN_PID + ") and PID_SYR='" + PJ_SYR + "' and SEQ_NO='" + sSEQ_NO + "'";
                #endregion

                sql = $@"
                        SELECT 
                            TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                            TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                            TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                            TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                            TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                            TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type
                        FROM CMEPMS.dbo.TDV_TDDCFTS TDV
                        LEFT JOIN CMEPMS.dbo.PJMFORM PJM ON TDV.PID = PJM.PID
                        WHERE TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PID = {AN_PID})
                        AND TDV.PID_SYR = '{PJ_SYR}' 
                        AND TDV.SEQ_NO = '{sSEQ_NO}' 
                        AND PJM.Water_Type = '雨水使用'
                        ORDER BY TDV.SIP_TD_CODE DESC";
            }
            else
            {
                #region SQL 查詢語句(舊版查詢)
                // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDCFTS where PID in (select PID from PJMFORM where AN_PID=" + AN_PID + ") and PID_SYR='" + PJ_SYR + "' and SEQ_NO='" + sSEQ_NO + "'";
                #endregion

                sql = $@"
                        SELECT 
                            TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                            TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                            TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                            TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                            TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                            TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type
                        FROM CMEPMS.dbo.TDV_TDDCFTS TDV
                        LEFT JOIN CMEPMS.dbo.PJMFORM PJM ON TDV.PID = PJM.PID
                        WHERE TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE AN_PID = {AN_PID})
                        AND TDV.PID_SYR = '{PJ_SYR}' 
                        AND TDV.SEQ_NO = '{sSEQ_NO}' 
                        AND PJM.Water_Type = '雨水使用'
                        ORDER BY TDV.SIP_TD_CODE,TDV.TID DESC";
            }

            dt1 = SourceFunction.DataTableGet(sql, connectionString);

            #endregion -------------

            #region DT2 DataTable 前期資料

            if (dt1 != null && dt1.Rows.Count > 0)
            {
                if (sPSEQ_NO != "")
                {
                    if (plan is null) // 包含"分支計畫"
                    {
                        #region SQL 查詢語句(舊版查詢)
                        // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where PID=" + AN_PID + ") and SEQ_NO='" + sPSEQ_NO + "'";
                        #endregion

                        sql = $@"
                                SELECT 
                                    TDV.PID,TDV.PID_SYR,TDV.SEQ_NO,TDV.TID, 
                                    TDV.FD_CODE,TDV.TD_CODE,TDV.TD_NAME,TDV.SIP_TD_CODE, 
                                    TDV.CITY,TDV.TOWN,TDV.ADM_DIST,TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC,TDV.N_IPJ_CODE, 
                                    TDV.IPJ_DESC,TDV.CATY_TYPE,TDV.YR_FUND,TDV.CEN_FUND, 
                                    TDV.LOC_FUND,TDV.PJV0050ord01,TDV.PJV0050Ord02,TDV.ACC_ID,PJM.Water_Type
                                FROM CMEPMS.dbo.TDV_TDDCFTS TDV
                                LEFT JOIN CMEPMS.dbo.PJMFORM PJM 
                                ON TDV.PID = PJM.PID
                                WHERE 
                                TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PID = {AN_PID})
                                AND TDV.PID_SYR = '{PJ_SYR}' 
                                AND TDV.SEQ_NO = '{sSEQ_NO}' 
                                AND PJM.Water_Type = '雨水使用'
                                ORDER BY TDV.SIP_TD_CODE DESC";
                    }
                    else
                    {
                        #region SQL 查詢語句(舊版查詢)
                        // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where AN_PID=" + AN_PID + ") and SEQ_NO='" + sPSEQ_NO + "'";
                        #endregion

                        sql = $@"
                                SELECT 
                                    TDV.PID,TDV.PID_SYR,TDV.SEQ_NO,TDV.TID, 
                                    TDV.FD_CODE,TDV.TD_CODE,TDV.TD_NAME,TDV.SIP_TD_CODE, 
                                    TDV.CITY,TDV.TOWN,TDV.ADM_DIST,TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC,TDV.N_IPJ_CODE, 
                                    TDV.IPJ_DESC,TDV.CATY_TYPE,TDV.YR_FUND,TDV.CEN_FUND, 
                                    TDV.LOC_FUND,TDV.PJV0050ord01,TDV.PJV0050Ord02,TDV.ACC_ID,PJM.Water_Type
                                FROM CMEPMS.dbo.TDV_TDDCFTS TDV
                                LEFT JOIN CMEPMS.dbo.PJMFORM PJM 
                                ON TDV.PID = PJM.PID
                                WHERE 
                                TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE AN_PID = {AN_PID})
                                AND TDV.PID_SYR = '{PJ_SYR}' 
                                AND TDV.SEQ_NO = '{sSEQ_NO}' 
                                AND PJM.Water_Type = '雨水使用'
                                ORDER BY TDV.SIP_TD_CODE DESC";
                    }
                }
                else
                {
                    if (plan is null) // 如果是分支計畫的情況
                    {
                        #region SQL 查詢語句(舊版查詢)
                        // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where PID=" + AN_PID + " and ADV_DP='" + OpID + "') and SEQ_NO='999'";
                        #endregion

                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type 
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE PID = {AN_PID} AND ADV_DP = '{OpID}') 
                                AND TDV.SEQ_NO = '999'
                                AND PJM.Water_Type = '雨水使用'
                                ORDER BY TDV.SIP_TD_CODE DESC ";
                    }
                    else
                    {
                        #region SQL 查詢語句(舊版查詢)
                        // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where AN_PID=" + AN_PID + " and ADV_DP='" + OpID + "') and SEQ_NO='999'";
                        #endregion

                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type 
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE AN_PID = {AN_PID} AND ADV_DP = '{OpID}') 
                                AND TDV.SEQ_NO = '999'
                                AND PJM.Water_Type = '雨水使用'
                                ORDER BY TDV.SIP_TD_CODE DESC ";
                    }
                }

                dt2 = SourceFunction.DataTableGet(sql, connectionString);
            }
            else
            {
                dt1 = new DataTable();
                if (plan is null) /// 如果是分支計畫的情況
                {
                    #region SQL 查詢語句(舊版查詢)
                    // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where PID=" + AN_PID + ") and SEQ_NO='" + sSEQ_NO + "'";
                    #endregion

                    sql = $@"
                            SELECT 
                                TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type -- 加入 Water_Type 欄位
                            FROM CMEPMS.dbo.TDV_TDDHFTS TDV
                            LEFT JOIN CMEPMS.dbo.PJMFORM PJM ON TDV.PID = PJM.PID -- 進行 LEFT JOIN
                            WHERE TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PID = {AN_PID}) 
                            AND TDV.SEQ_NO = '{sSEQ_NO}' 
                            AND PJM.Water_Type = '雨水使用'
                            ORDER BY TDV.SIP_TD_CODE DESC ";
                }
                else
                {
                    #region SQL 查詢語句(舊版查詢)
                    // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where AN_PID=" + AN_PID + ") and SEQ_NO='" + sSEQ_NO + "'";
                    #endregion

                    sql = $@"
                            SELECT 
                                TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type 
                            FROM TDV_TDDHFTS TDV
                            LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID -- 加入 LEFT JOIN
                            WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE AN_PID = {AN_PID}) 
                            AND TDV.SEQ_NO = '{sSEQ_NO}' 
                            AND PJM.Water_Type = '雨水使用'
                            ORDER BY TDV.SIP_TD_CODE DESC ";
                }
                dt1 = SourceFunction.DataTableGet(sql, connectionString);

                if (sPSEQ_NO != "")
                {
                    if (plan is null)
                    {
                        #region SQL 查詢語句(舊版查詢)
                        // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where PID=" + AN_PID + ") and SEQ_NO='" + sPSEQ_NO + "'";
                        #endregion

                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE PID = {AN_PID}) 
                                AND TDV.SEQ_NO = '{sPSEQ_NO}'
                                AND PJM.Water_Type = '雨水使用' 
                                ORDER BY TDV.SIP_TD_CODE DESC ";
                    }
                    else
                    {
                        #region SQL 查詢語句(舊版查詢)
                        // sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where AN_PID=" + AN_PID + ") and SEQ_NO='" + sPSEQ_NO + "'";
                        #endregion

                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type 
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE AN_PID = {AN_PID})
                                AND TDV.SEQ_NO = '{sPSEQ_NO}'
                                AND PJM.Water_Type = '雨水使用' 
                                ORDER BY TDV.SIP_TD_CODE DESC ";
                    }
                }
                else
                {
                    if (plan is null)
                    {
                        #region SQL 查詢語句(舊版查詢)
                        //sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where PID=" + AN_PID + ") and SEQ_NO='999'";
                        #endregion

                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE PID = {AN_PID} AND ADV_DP = '{OpID}') 
                                AND TDV.SEQ_NO = '999'
                                AND PJM.Water_Type = '雨水使用' 
                                ORDER BY TDV.SIP_TD_CODE DESC ";
                    }
                    else
                    {
                        #region SQL 查詢語句(舊版查詢)
                        //sql = "select PID,PID_SYR,SEQ_NO,TID,FD_CODE,TD_CODE,TD_NAME,SIP_TD_CODE,CITY,TOWN,ADM_DIST,IPJ_CODE,IPJ_CODE_DESC,N_IPJ_CODE,IPJ_DESC,CATY_TYPE,YR_FUND,CEN_FUND,LOC_FUND,PJV0050ord01,PJV0050Ord02,ACC_ID from TDV_TDDHFTS where PID in (select PID from PJMFORM where AN_PID=" + AN_PID + ") and SEQ_NO='999'";
                        #endregion

                        sql = $@"
                                SELECT 
                                    TDV.PID,TDV.PID_SYR,TDV.SEQ_NO,TDV.TID, 
                                    TDV.FD_CODE,TDV.TD_CODE,TDV.TD_NAME,TDV.SIP_TD_CODE, 
                                    TDV.CITY,TDV.TOWN,TDV.ADM_DIST,TDV.IPJ_CODE,TDV.IPJ_CODE_DESC, 
                                    TDV.N_IPJ_CODE,TDV.IPJ_DESC,TDV.CATY_TYPE,TDV.YR_FUND,TDV.CEN_FUND, 
                                    TDV.LOC_FUND,TDV.PJV0050ord01,TDV.PJV0050Ord02,TDV.ACC_ID,PJM.Water_Type
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE AN_PID = {AN_PID}) 
                                AND TDV.SEQ_NO = '999'
                                AND PJM.Water_Type = '雨水使用' 
                                ORDER BY TDV.SIP_TD_CODE DESC ";
                    }
                }
                dt2 = SourceFunction.DataTableGet(sql, connectionString);
            }

            #endregion -------------

            #region  DT3 DataTable

            /// 判斷 Plan 是否為 Null
            /// == NULL => 執行查詢結果，將結果存至 DT3
            /// != NULL => 執行查詢結果，將結果存至 DT3

            dt3 = new DataTable();

            if (plan is null)
            {
                sql = $@"
                        SELECT 
                            TDV.TD_CODE,TDV.FD_CODE,TDV.PID,TDV.SEQ_NO,TDV.TID,
                            TDV.TD_NAME,TDV.PID_SYR,TDV.SIP_TD_CODE,TDV.CITY,TDV.TOWN,
                            TDV.ADM_DIST,TDV.IPJ_CODE,TDV.IPJ_CODE_DESC,TDV.N_IPJ_CODE,TDV.IPJ_DESC,
                            TDV.CATY_TYPE,TDV.YR_FUND,TDV.CEN_FUND,TDV.LOC_FUND,TDV.PJV0050ord01,
                            TDV.PJV0050Ord02,TDV.ACC_ID,PJM.Water_Type
                        FROM TDV_TDDCFTS TDV
                        LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID
                        WHERE TDV.PID IN (
                            SELECT PID 
                            FROM PJMFORM 
                            WHERE SYR <= '{sPSYR}' 
                              AND PID = {AN_PID}
                              AND PJM.Water_Type = '雨水使用'
                        )
                    ";

            }
            else
            {
                sql = $@"
                        SELECT 
                            TDV.TD_CODE,TDV.FD_CODE,TDV.PID,TDV.SEQ_NO,TDV.TID,TDV.TD_NAME,
                            TDV.PID_SYR,TDV.SIP_TD_CODE,TDV.CITY,TDV.TOWN,TDV.ADM_DIST,TDV.IPJ_CODE,
                            TDV.IPJ_CODE_DESC,TDV.N_IPJ_CODE,TDV.IPJ_DESC,TDV.CATY_TYPE,TDV.YR_FUND,TDV.CEN_FUND,
                            TDV.LOC_FUND,TDV.PJV0050ord01,TDV.PJV0050Ord02,TDV.ACC_ID,PJM.Water_Type
                        FROM TDV_TDDCFTS TDV
                        LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID
                        WHERE TDV.PID IN (
                            SELECT PID 
                            FROM PJMFORM 
                            WHERE SYR <= '{sPSYR}' 
                              AND AN_PID = {AN_PID}
                              AND PJM.Water_Type = '雨水使用'
                        )
                        ORDER BY TDV.PID_SYR, TDV.SIP_TD_CODE";
            }

            dt3 = SourceFunction.DataTableGet(sql, connectionString);

            #endregion -------------

            #region

            ditdt = dt3.DefaultView.ToTable(true, new string[] { "PID_SYR" });
            ditdt.TableName = "ditdt";

            dt = new DataTable();
            dt.Columns.Add("PID");
            dt.Columns.Add("PID_SYR");
            dt.Columns.Add("SEQ_NO");
            dt.Columns.Add("TID");
            dt.Columns.Add("FD_CODE");
            dt.Columns.Add("TD_CODE");
            dt.Columns.Add("TD_NAME");
            dt.Columns.Add("SIP_TD_CODE");
            dt.Columns.Add("CITY");
            dt.Columns.Add("DATACITY");
            dt.Columns.Add("TOWN");
            dt.Columns.Add("ADM_DIST");
            dt.Columns.Add("IPJ_CODE");
            dt.Columns.Add("IPJ_CODE_DESC");
            dt.Columns.Add("N_IPJ_CODE");
            dt.Columns.Add("IPJ_DESC");
            dt.Columns.Add("CATY_TYPE");
            dt.Columns.Add("YR_FUND");
            dt.Columns.Add("CEN_FUND");
            dt.Columns.Add("LOC_FUND");
            dt.Columns.Add("PYR_FUND"); /// 前一期
            dt.Columns.Add("PCEN_FUND");
            dt.Columns.Add("PLOC_FUND");
            dt.Columns.Add("FROM");


            /// 以前年度經費欄位
            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                dt.Columns.Add("PYR_YR_FUND_" + ditdt.Rows[i]["PID_SYR"].ToString());  /// 以前年度
                dt.Columns.Add("PYR_CEN_FUND_" + ditdt.Rows[i]["PID_SYR"].ToString());
                dt.Columns.Add("PYR_LOC_FUND_" + ditdt.Rows[i]["PID_SYR"].ToString());
            }

            dt.Columns.Add("PJV0050ord01"); /// 計畫項目別
            dt.Columns.Add("PJV0050ord02");
            dt.Columns.Add("ACC_ID");

            /// 本期資料
            for (i = 0; i < dt1.Rows.Count; i++)
            {
                dr = dt.NewRow();
                dr["PID"] = dt1.Rows[i]["PID"].ToString();
                dr["PID_SYR"] = dt1.Rows[i]["PID_SYR"].ToString();
                dr["SEQ_NO"] = dt1.Rows[i]["SEQ_NO"].ToString();
                dr["TID"] = dt1.Rows[i]["TID"].ToString();
                dr["FD_CODE"] = dt1.Rows[i]["FD_CODE"].ToString();
                dr["TD_CODE"] = dt1.Rows[i]["TD_CODE"].ToString();
                dr["TD_NAME"] = dt1.Rows[i]["TD_NAME"].ToString();
                dr["SIP_TD_CODE"] = dt1.Rows[i]["SIP_TD_CODE"].ToString();
                dr["CITY"] = dt1.Rows[i]["CITY"].ToString();
                dr["TOWN"] = dt1.Rows[i]["TOWN"].ToString();
                dr["ADM_DIST"] = dt1.Rows[i]["ADM_DIST"].ToString();
                dr["IPJ_CODE"] = dt1.Rows[i]["IPJ_CODE"].ToString();
                dr["IPJ_CODE_DESC"] = dt1.Rows[i]["IPJ_CODE_DESC"].ToString();
                dr["N_IPJ_CODE"] = dt1.Rows[i]["N_IPJ_CODE"].ToString();
                dr["IPJ_DESC"] = dt1.Rows[i]["IPJ_DESC"].ToString();
                dr["CATY_TYPE"] = dt1.Rows[i]["CATY_TYPE"].ToString();
                dr["YR_FUND"] = dt1.Rows[i]["YR_FUND"].ToString();
                dr["CEN_FUND"] = dt1.Rows[i]["CEN_FUND"].ToString();
                dr["LOC_FUND"] = dt1.Rows[i]["LOC_FUND"].ToString();
                dr["PYR_FUND"] = "0";
                dr["PCEN_FUND"] = "0";
                dr["PLOC_FUND"] = "0";

                for (j = 0; j < ditdt.Rows.Count; j++)
                {
                    dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = "0";
                    dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = "0";
                    dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = "0";
                }

                dr["FROM"] = "dt1";
                dr["PJV0050ord01"] = dt1.Rows[i]["PJV0050ord01"].ToString();
                dr["PJV0050ord02"] = dt1.Rows[i]["PJV0050ord02"].ToString();
                dr["ACC_ID"] = dt1.Rows[i]["ACC_ID"].ToString();
                dt.Rows.Add(dr);
            }

            /// 前期有，本期沒有
            for (i = 0; i < dt2.Rows.Count; i++)
            {
                rr = dt2.Select("TD_CODE='" + dt2.Rows[i]["TD_CODE"].ToString() + "'");

                if (rr.Length == 0)
                {
                    dr = dt.NewRow();

                    dr["PID"] = dt2.Rows[i]["PID"].ToString();
                    dr["PID_SYR"] = dt2.Rows[i]["PID_SYR"].ToString();
                    dr["SEQ_NO"] = dt2.Rows[i]["SEQ_NO"].ToString();
                    dr["TID"] = dt2.Rows[i]["TID"].ToString();
                    dr["FD_CODE"] = dt2.Rows[i]["FD_CODE"].ToString();
                    dr["TD_CODE"] = dt2.Rows[i]["TD_CODE"].ToString();
                    dr["TD_NAME"] = dt2.Rows[i]["TD_NAME"].ToString();
                    dr["SIP_TD_CODE"] = dt2.Rows[i]["SIP_TD_CODE"].ToString();
                    dr["CITY"] = dt2.Rows[i]["CITY"].ToString();
                    dr["TOWN"] = dt2.Rows[i]["TOWN"].ToString();
                    dr["ADM_DIST"] = dt2.Rows[i]["ADM_DIST"].ToString();
                    dr["IPJ_CODE"] = dt2.Rows[i]["IPJ_CODE"].ToString();
                    dr["IPJ_CODE_DESC"] = dt2.Rows[i]["IPJ_CODE_DESC"].ToString();
                    dr["N_IPJ_CODE"] = dt2.Rows[i]["N_IPJ_CODE"].ToString();
                    dr["IPJ_DESC"] = dt2.Rows[i]["IPJ_DESC"].ToString();
                    dr["CATY_TYPE"] = dt2.Rows[i]["CATY_TYPE"].ToString();
                    dr["YR_FUND"] = "0";  /// 本期經費為 0
                    dr["CEN_FUND"] = "0";
                    dr["LOC_FUND"] = "0";
                    dr["PYR_FUND"] = dt2.Rows[i]["YR_FUND"].ToString();   /// 前期經費
                    dr["PCEN_FUND"] = dt2.Rows[i]["CEN_FUND"].ToString();
                    dr["PLOC_FUND"] = dt2.Rows[i]["LOC_FUND"].ToString();

                    for (j = 0; j < ditdt.Rows.Count; j++)
                    {
                        dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = "0";
                        dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = "0";
                        dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = "0";
                    }

                    dr["FROM"] = "dt2"; /// 資料來自 DT2
                    dr["PJV0050ord01"] = dt2.Rows[i]["PJV0050ord01"].ToString();
                    dr["PJV0050ord02"] = dt2.Rows[i]["PJV0050ord02"].ToString();
                    dr["ACC_ID"] = dt2.Rows[i]["ACC_ID"].ToString();
                    dt.Rows.Add(dr);
                }
            }

            for (i = 0; i < dt.Rows.Count; i++)
            {
                /// 查找 DT2 中是否存在相同的 TD_CODE
                rr = dt2.Select("TD_CODE='" + dt.Rows[i]["TD_CODE"].ToString() + "'");

                if (rr.Length > 0)
                {
                    /// 更新前期經費
                    for (j = 0; j < rr.Length; j++)
                    {
                        dt.Rows[i]["PYR_FUND"] = rr[j]["YR_FUND"].ToString();
                        dt.Rows[i]["PCEN_FUND"] = rr[j]["CEN_FUND"].ToString();
                        dt.Rows[i]["PLOC_FUND"] = rr[j]["LOC_FUND"].ToString();
                    }
                }
            }

            /// 前一年度有，本期沒有
            for (i = 0; i < dt3.Rows.Count; i++)
            {
                rr = dt1.Select("TD_CODE='" + dt3.Rows[i]["TD_CODE"].ToString() + "'");

                if (rr.Length == 0)
                {
                    rr = dt.Select("TD_CODE='" + dt3.Rows[i]["TD_CODE"].ToString() + "'");

                    if (rr.Length == 0)
                    {
                        dr = dt.NewRow();
                        dr["PID"] = dt3.Rows[i]["PID"].ToString();
                        dr["PID_SYR"] = "";
                        dr["SEQ_NO"] = dt3.Rows[i]["SEQ_NO"].ToString();
                        dr["TID"] = dt3.Rows[i]["TID"].ToString();
                        dr["FD_CODE"] = dt3.Rows[i]["FD_CODE"].ToString();
                        dr["TD_CODE"] = dt3.Rows[i]["TD_CODE"].ToString();
                        dr["TD_NAME"] = dt3.Rows[i]["TD_NAME"].ToString();
                        dr["SIP_TD_CODE"] = dt3.Rows[i]["SIP_TD_CODE"].ToString();
                        dr["CITY"] = dt3.Rows[i]["CITY"].ToString();
                        dr["TOWN"] = dt3.Rows[i]["TOWN"].ToString();
                        dr["ADM_DIST"] = dt3.Rows[i]["ADM_DIST"].ToString();
                        dr["IPJ_CODE"] = dt3.Rows[i]["IPJ_CODE"].ToString();
                        dr["IPJ_CODE_DESC"] = dt3.Rows[i]["IPJ_CODE_DESC"].ToString();
                        dr["N_IPJ_CODE"] = dt3.Rows[i]["N_IPJ_CODE"].ToString();
                        dr["IPJ_DESC"] = dt3.Rows[i]["IPJ_DESC"].ToString();
                        dr["CATY_TYPE"] = dt3.Rows[i]["CATY_TYPE"].ToString();
                        dr["YR_FUND"] = "0";   /// 本期經費為 0
                        dr["CEN_FUND"] = "0";  /// 中央經費為 0
                        dr["LOC_FUND"] = "0";  /// 地方經費為 0
                        dr["PYR_FUND"] = "0";  /// 前期經費為 0
                        dr["PCEN_FUND"] = "0"; /// 前期中央經費為 0
                        dr["PLOC_FUND"] = "0"; /// 前期地方經費為 0
                        dr["PYR_YR_FUND_" + dt3.Rows[i]["PID_SYR"].ToString()] = dt3.Rows[i]["YR_FUND"].ToString();  /// 填入前一年度的經費數據
                        dr["PYR_CEN_FUND_" + dt3.Rows[i]["PID_SYR"].ToString()] = dt3.Rows[i]["CEN_FUND"].ToString();/// 填入前一年度的經費數據
                        dr["PYR_LOC_FUND_" + dt3.Rows[i]["PID_SYR"].ToString()] = dt3.Rows[i]["LOC_FUND"].ToString();/// 填入前一年度的經費數據
                        dr["PJV0050ord01"] = dt3.Rows[i]["PJV0050ord01"].ToString(); /// 填入其他計畫項目和來源資訊
                        dr["PJV0050ord02"] = dt3.Rows[i]["PJV0050ord02"].ToString(); /// 填入其他計畫項目和來源資訊
                        dr["FROM"] = "dt3"; /// 來源是 DT3
                        dr["ACC_ID"] = dt3.Rows[i]["ACC_ID"].ToString();
                        dt.Rows.Add(dr);
                    }
                }
            }

            for (i = 0; i < dt.Rows.Count; i++)
            {
                rr = dt3.Select("TD_CODE='" + dt.Rows[i]["TD_CODE"].ToString() + "'");

                for (j = 0; j < rr.Length; j++)
                {
                    /// 將前一年度的經費資料從 DT3 更新到 DT 中
                    dt.Rows[i]["PYR_YR_FUND_" + rr[j]["PID_SYR"].ToString()] = rr[j]["YR_FUND"].ToString();
                    dt.Rows[i]["PYR_CEN_FUND_" + rr[j]["PID_SYR"].ToString()] = rr[j]["CEN_FUND"].ToString();
                    dt.Rows[i]["PYR_LOC_FUND_" + rr[j]["PID_SYR"].ToString()] = rr[j]["LOC_FUND"].ToString();
                }
            }

            for (i = 0; i < dt.Rows.Count; i++)
            {
                switch (dt.Rows[i]["CITY"].ToString())
                {
                    case "台北縣":
                    case "新北市":
                        dt.Rows[i]["DATACITY"] = "新北市";
                        break;
                    case "台中市":
                    case "台中縣":
                        dt.Rows[i]["DATACITY"] = "台中市";
                        break;
                    case "桃園市":
                    case "桃園縣":
                        dt.Rows[i]["DATACITY"] = "桃園市";
                        break;
                    case "台南市":
                    case "台南縣":
                        dt.Rows[i]["DATACITY"] = "台南市";
                        break;
                    case "高雄縣":
                    case "高雄市":
                        dt.Rows[i]["DATACITY"] = "高雄市";
                        break;
                    default:
                        dt.Rows[i]["DATACITY"] = dt.Rows[i]["CITY"].ToString();
                        break;
                }
            }

            dt.DefaultView.Sort = "DATACITY";
            dt = dt.DefaultView.ToTable();

            #endregion

            #region ---------------- 署辦IPJ_CODE='01' and ACC_ID<>'0200'

            /// 指定為署辦 IPJ_CODE='01'
            /// 經費編號不為 '0200'
            sWhere = " IPJ_CODE='01' and ACC_ID<>'0200'";

            /// Clone DT 資料，利用 SWhere 篩選條件
            dtp1 = dt.Clone();
            dtp1.TableName = "dtp1";
            rr1 = dt.Select(sWhere);

            tempCITY = "";

            if (rr1.Length > 0)
            {
                tempCITY = rr1[0]["DATACITY"].ToString();
            }

            for (i = 0; i < rr1.Length; i++)
            {
                r = rr1[i];

                if (tempCITY != "" && tempCITY != r["DATACITY"].ToString())
                {
                    dr = dtp1.NewRow();

                    /// 縣市變更時，新增一筆小計
                    dr["PID"] = "";
                    dr["PID_SYR"] = "";
                    dr["SEQ_NO"] = "";
                    dr["TID"] = "";
                    dr["FD_CODE"] = "";
                    dr["TD_CODE"] = "";
                    dr["TD_NAME"] = "小計";
                    dr["SIP_TD_CODE"] = "";
                    dr["CITY"] = "";
                    dr["TOWN"] = "";
                    dr["ADM_DIST"] = "";
                    dr["IPJ_CODE"] = "";
                    dr["IPJ_CODE_DESC"] = "";
                    dr["N_IPJ_CODE"] = "";
                    dr["IPJ_DESC"] = "";
                    dr["YR_FUND"] = getDtColumns_SUM(dt, "YR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["CEN_FUND"] = getDtColumns_SUM(dt, "CEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["LOC_FUND"] = getDtColumns_SUM(dt, "LOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PYR_FUND"] = getDtColumns_SUM(dt, "PYR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PCEN_FUND"] = getDtColumns_SUM(dt, "PCEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PLOC_FUND"] = getDtColumns_SUM(dt, "PLOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);

                    /// 針對各年份進行累計
                    for (j = 0; j < ditdt.Rows.Count; j++)
                    {
                        dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                        dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                        dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                    }

                    dr["FROM"] = "";
                    dtp1.Rows.Add(dr);
                    tempCITY = r["DATACITY"].ToString();
                }

                rr = dtp1.Select("TD_CODE='" + rr1[i]["TD_CODE"].ToString() + "'");

                if (rr.Length == 0)
                {
                    /// 如果該專案尚未存在於 `DTP1`，則新增一筆資料
                    dr = dtp1.NewRow();

                    dr["PID"] = r["PID"].ToString();
                    dr["PID_SYR"] = r["PID_SYR"].ToString();
                    dr["SEQ_NO"] = r["SEQ_NO"].ToString();
                    dr["TID"] = r["TID"].ToString();
                    dr["FD_CODE"] = r["FD_CODE"].ToString();
                    dr["TD_CODE"] = r["TD_CODE"].ToString();
                    dr["TD_NAME"] = r["TD_NAME"].ToString();
                    dr["SIP_TD_CODE"] = r["SIP_TD_CODE"].ToString();
                    dr["CITY"] = r["CITY"].ToString();
                    dr["TOWN"] = r["TOWN"].ToString();
                    dr["ADM_DIST"] = r["ADM_DIST"].ToString();
                    dr["IPJ_CODE"] = r["IPJ_CODE"].ToString();
                    dr["IPJ_CODE_DESC"] = r["IPJ_CODE_DESC"].ToString();
                    dr["N_IPJ_CODE"] = r["N_IPJ_CODE"].ToString();
                    dr["IPJ_DESC"] = r["IPJ_DESC"].ToString();
                    dr["CATY_TYPE"] = r["CATY_TYPE"].ToString();
                    dr["YR_FUND"] = r["YR_FUND"].ToString();
                    dr["CEN_FUND"] = r["CEN_FUND"].ToString();
                    dr["LOC_FUND"] = r["LOC_FUND"].ToString();
                    dr["PYR_FUND"] = r["PYR_FUND"].ToString();
                    dr["PCEN_FUND"] = r["PCEN_FUND"].ToString();
                    dr["PLOC_FUND"] = r["PLOC_FUND"].ToString();

                    for (j = 0; j < ditdt.Rows.Count; j++)
                    {
                        dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                        dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                        dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                    }

                    dr["FROM"] = r["FROM"].ToString();
                    dr["PJV0050ord01"] = r["PJV0050ord01"].ToString();
                    dr["PJV0050ord02"] = r["PJV0050ord02"].ToString();
                    dtp1.Rows.Add(dr);
                }
                else
                {
                    /// 相同工程編號，不同經費編號
                    rr[0]["YR_FUND"] = r["YR_FUND"].ToString();
                    rr[0]["CEN_FUND"] = r["CEN_FUND"].ToString();
                    rr[0]["LOC_FUND"] = r["LOC_FUND"].ToString();
                    rr[0]["PYR_FUND"] = r["PYR_FUND"].ToString();
                    rr[0]["PCEN_FUND"] = r["PCEN_FUND"].ToString();
                    rr[0]["PLOC_FUND"] = r["PLOC_FUND"].ToString();
                }
            }

            if (rr1.Length > 0)
            {
                tempCITY = rr1[rr1.Length - 1]["DATACITY"].ToString();

                /// 最後一個縣市
                dr = dtp1.NewRow();
                dr["PID"] = "";
                dr["PID_SYR"] = "";
                dr["SEQ_NO"] = "";
                dr["TID"] = "";
                dr["FD_CODE"] = "";
                dr["TD_CODE"] = "";
                dr["TD_NAME"] = "小計";
                dr["SIP_TD_CODE"] = "";
                dr["CITY"] = "";
                dr["TOWN"] = "";
                dr["ADM_DIST"] = "";
                dr["IPJ_CODE"] = "";
                dr["IPJ_CODE_DESC"] = "";
                dr["N_IPJ_CODE"] = "";
                dr["IPJ_DESC"] = "";
                dr["YR_FUND"] = getDtColumns_SUM(dt, "YR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["CEN_FUND"] = getDtColumns_SUM(dt, "CEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["LOC_FUND"] = getDtColumns_SUM(dt, "LOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["PYR_FUND"] = getDtColumns_SUM(dt, "PYR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["PCEN_FUND"] = getDtColumns_SUM(dt, "PCEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["PLOC_FUND"] = getDtColumns_SUM(dt, "PLOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);

                for (j = 0; j < ditdt.Rows.Count; j++)
                {
                    dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                }

                dr["FROM"] = "";

                dtp1.Rows.Add(dr);
            }

            ds.Tables.Add(dtp1);

            #endregion -------------

            #region ---------------- 地辦IPJ_CODE='02'

            sWhere = " IPJ_CODE='02'";

            dtp2 = dt.Clone();

            dtp2.TableName = "dtp2";

            rr1 = dt.Select(sWhere);

            tempCITY = "";

            if (rr1.Length > 0)
            {
                tempCITY = rr1[0]["DATACITY"].ToString();
            }

            for (i = 0; i < rr1.Length; i++)
            {
                r = rr1[i];

                if (tempCITY != "" && tempCITY != r["DATACITY"].ToString())
                {
                    dr = dtp2.NewRow();
                    dr["PID"] = "";
                    dr["PID_SYR"] = "";
                    dr["SEQ_NO"] = "";
                    dr["TID"] = "";
                    dr["FD_CODE"] = "";
                    dr["TD_CODE"] = "";
                    dr["TD_NAME"] = "小計";
                    dr["SIP_TD_CODE"] = "";
                    dr["CITY"] = "";
                    dr["TOWN"] = "";
                    dr["ADM_DIST"] = "";
                    dr["IPJ_CODE"] = "";
                    dr["IPJ_CODE_DESC"] = "";
                    dr["N_IPJ_CODE"] = "";
                    dr["IPJ_DESC"] = "";
                    dr["YR_FUND"] = getDtColumns_SUM(dt, "YR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["CEN_FUND"] = getDtColumns_SUM(dt, "CEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["LOC_FUND"] = getDtColumns_SUM(dt, "LOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PYR_FUND"] = getDtColumns_SUM(dt, "PYR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PCEN_FUND"] = getDtColumns_SUM(dt, "PCEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PLOC_FUND"] = getDtColumns_SUM(dt, "PLOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);

                    for (j = 0; j < ditdt.Rows.Count; j++)
                    {
                        dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                        dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                        dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                    }

                    dr["FROM"] = "";
                    dtp2.Rows.Add(dr);
                    tempCITY = r["DATACITY"].ToString();
                }

                rr = dtp2.Select("TD_CODE='" + dt.Rows[i]["TD_CODE"].ToString() + "'");

                if (rr.Length == 0)
                {
                    dr = dtp2.NewRow();
                    dr["PID"] = r["PID"].ToString();
                    dr["PID_SYR"] = r["PID_SYR"].ToString();
                    dr["SEQ_NO"] = r["SEQ_NO"].ToString();
                    dr["TID"] = r["TID"].ToString();
                    dr["FD_CODE"] = r["FD_CODE"].ToString();
                    dr["TD_CODE"] = r["TD_CODE"].ToString();
                    dr["TD_NAME"] = r["TD_NAME"].ToString();
                    dr["SIP_TD_CODE"] = r["SIP_TD_CODE"].ToString();
                    dr["CITY"] = r["CITY"].ToString();
                    dr["TOWN"] = r["TOWN"].ToString();
                    dr["ADM_DIST"] = r["ADM_DIST"].ToString();
                    dr["IPJ_CODE"] = r["IPJ_CODE"].ToString();
                    dr["IPJ_CODE_DESC"] = r["IPJ_CODE_DESC"].ToString();
                    dr["N_IPJ_CODE"] = r["N_IPJ_CODE"].ToString();
                    dr["IPJ_DESC"] = r["IPJ_DESC"].ToString();
                    dr["CATY_TYPE"] = r["CATY_TYPE"].ToString();
                    dr["YR_FUND"] = r["YR_FUND"].ToString();
                    dr["CEN_FUND"] = r["CEN_FUND"].ToString();
                    dr["LOC_FUND"] = r["LOC_FUND"].ToString();
                    dr["PYR_FUND"] = r["PYR_FUND"].ToString();
                    dr["PCEN_FUND"] = r["PCEN_FUND"].ToString();
                    dr["PLOC_FUND"] = r["PLOC_FUND"].ToString();

                    for (j = 0; j < ditdt.Rows.Count; j++)
                    {
                        dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                        dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                        dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                    }

                    dr["FROM"] = r["FROM"].ToString();
                    dr["PJV0050ord01"] = r["PJV0050ord01"].ToString();
                    dr["PJV0050ord02"] = r["PJV0050ord02"].ToString();
                    dtp2.Rows.Add(dr);
                }
                else
                {
                    /// 相同工程編號，不同經費編號
                    rr[0]["YR_FUND"] = r["YR_FUND"].ToString();
                    rr[0]["CEN_FUND"] = r["CEN_FUND"].ToString();
                    rr[0]["LOC_FUND"] = r["LOC_FUND"].ToString();
                    rr[0]["PYR_FUND"] = r["PYR_FUND"].ToString();
                    rr[0]["PCEN_FUND"] = r["PCEN_FUND"].ToString();
                    rr[0]["PLOC_FUND"] = r["PLOC_FUND"].ToString();
                }
            }

            if (rr1.Length > 0)
            {
                tempCITY = rr1[rr1.Length - 1]["DATACITY"].ToString();

                //最後一個縣市
                dr = dtp2.NewRow();
                dr["PID"] = "";
                dr["PID_SYR"] = "";
                dr["SEQ_NO"] = "";
                dr["TID"] = "";
                dr["FD_CODE"] = "";
                dr["TD_CODE"] = "";
                dr["TD_NAME"] = "小計";
                dr["SIP_TD_CODE"] = "";
                dr["CITY"] = "";
                dr["TOWN"] = "";
                dr["ADM_DIST"] = "";
                dr["IPJ_CODE"] = "";
                dr["IPJ_CODE_DESC"] = "";
                dr["N_IPJ_CODE"] = "";
                dr["IPJ_DESC"] = "";
                dr["YR_FUND"] = getDtColumns_SUM(dt, "YR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["CEN_FUND"] = getDtColumns_SUM(dt, "CEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["LOC_FUND"] = getDtColumns_SUM(dt, "LOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["PYR_FUND"] = getDtColumns_SUM(dt, "PYR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["PCEN_FUND"] = getDtColumns_SUM(dt, "PCEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["PLOC_FUND"] = getDtColumns_SUM(dt, "PLOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);

                for (j = 0; j < ditdt.Rows.Count; j++)
                {
                    dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                }

                dr["FROM"] = "";
                dtp2.Rows.Add(dr);
            }

            ds.Tables.Add(dtp2);

            #endregion -------------

            #region ---------------- 業務費 IPJ_CODE='01' and ACC_ID='0200'

            sWhere = "IPJ_CODE='01' and ACC_ID='0200'";
            dtp3 = dt.Clone();
            dtp3.TableName = "dtp3";
            rr1 = dt.Select(sWhere);
            tempCITY = "";

            if (rr1.Length > 0)
            {
                tempCITY = rr1[0]["DATACITY"].ToString();
            }

            for (i = 0; i < rr1.Length; i++)
            {
                r = rr1[i];
                if (tempCITY != "" && tempCITY != r["DATACITY"].ToString())
                {
                    dr = dtp3.NewRow();
                    dr["PID"] = "";
                    dr["PID_SYR"] = "";
                    dr["SEQ_NO"] = "";
                    dr["TID"] = "";
                    dr["FD_CODE"] = "";
                    dr["TD_CODE"] = "";
                    dr["TD_NAME"] = "小計";
                    dr["SIP_TD_CODE"] = "";
                    dr["CITY"] = "";
                    dr["TOWN"] = "";
                    dr["ADM_DIST"] = "";
                    dr["IPJ_CODE"] = "";
                    dr["IPJ_CODE_DESC"] = "";
                    dr["N_IPJ_CODE"] = "";
                    dr["IPJ_DESC"] = "";
                    dr["YR_FUND"] = getDtColumns_SUM(dt, "YR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["CEN_FUND"] = getDtColumns_SUM(dt, "CEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["LOC_FUND"] = getDtColumns_SUM(dt, "LOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PYR_FUND"] = getDtColumns_SUM(dt, "PYR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PCEN_FUND"] = getDtColumns_SUM(dt, "PCEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PLOC_FUND"] = getDtColumns_SUM(dt, "PLOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);

                    for (j = 0; j < ditdt.Rows.Count; j++)
                    {
                        dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                        dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                        dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                    }

                    dr["FROM"] = "";
                    dtp3.Rows.Add(dr);
                    tempCITY = r["DATACITY"].ToString();
                }

                dr = dtp3.NewRow();
                dr["PID"] = r["PID"].ToString();
                dr["PID_SYR"] = r["PID_SYR"].ToString();
                dr["SEQ_NO"] = r["SEQ_NO"].ToString();
                dr["TID"] = r["TID"].ToString();
                dr["FD_CODE"] = r["FD_CODE"].ToString();
                dr["TD_CODE"] = r["TD_CODE"].ToString();
                dr["TD_NAME"] = r["TD_NAME"].ToString();
                dr["SIP_TD_CODE"] = r["SIP_TD_CODE"].ToString();
                dr["CITY"] = r["CITY"].ToString();
                dr["TOWN"] = r["TOWN"].ToString();
                dr["ADM_DIST"] = r["ADM_DIST"].ToString();
                dr["IPJ_CODE"] = r["IPJ_CODE"].ToString();
                dr["IPJ_CODE_DESC"] = r["IPJ_CODE_DESC"].ToString();
                dr["N_IPJ_CODE"] = r["N_IPJ_CODE"].ToString();
                dr["IPJ_DESC"] = r["IPJ_DESC"].ToString();
                dr["CATY_TYPE"] = r["CATY_TYPE"].ToString();
                dr["YR_FUND"] = r["YR_FUND"].ToString();
                dr["CEN_FUND"] = r["CEN_FUND"].ToString();
                dr["LOC_FUND"] = r["LOC_FUND"].ToString();
                dr["PYR_FUND"] = r["PYR_FUND"].ToString();
                dr["PCEN_FUND"] = r["PCEN_FUND"].ToString();
                dr["PLOC_FUND"] = r["PLOC_FUND"].ToString();

                for (j = 0; j < ditdt.Rows.Count; j++)
                {
                    dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                    dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                    dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                }

                dr["FROM"] = r["FROM"].ToString();
                dr["PJV0050ord01"] = r["PJV0050ord01"].ToString();
                dr["PJV0050ord02"] = r["PJV0050ord02"].ToString();
                dtp3.Rows.Add(dr);
            }

            if (rr1.Length > 0)
            {
                tempCITY = rr1[rr1.Length - 1]["DATACITY"].ToString();

                /// 最後一個縣市
                dr = dtp3.NewRow();
                dr["PID"] = "";
                dr["PID_SYR"] = "";
                dr["SEQ_NO"] = "";
                dr["TID"] = "";
                dr["FD_CODE"] = "";
                dr["TD_CODE"] = "";
                dr["TD_NAME"] = "小計";
                dr["SIP_TD_CODE"] = "";
                dr["CITY"] = "";
                dr["TOWN"] = "";
                dr["ADM_DIST"] = "";
                dr["IPJ_CODE"] = "";
                dr["IPJ_CODE_DESC"] = "";
                dr["N_IPJ_CODE"] = "";
                dr["IPJ_DESC"] = "";
                dr["YR_FUND"] = getDtColumns_SUM(dt, "YR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["CEN_FUND"] = getDtColumns_SUM(dt, "CEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["LOC_FUND"] = getDtColumns_SUM(dt, "LOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["PYR_FUND"] = getDtColumns_SUM(dt, "PYR_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["PCEN_FUND"] = getDtColumns_SUM(dt, "PCEN_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);
                dr["PLOC_FUND"] = getDtColumns_SUM(dt, "PLOC_FUND", "DATACITY='" + tempCITY + "' and " + sWhere);

                for (j = 0; j < ditdt.Rows.Count; j++)
                {
                    dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                    dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and " + sWhere);
                }

                dr["FROM"] = "";
                dtp3.Rows.Add(dr);
            }

            ds.Tables.Add(dtp3);

            #endregion -------------

            #region ---------------- 經費調整表-明細

            if (plan is null)
            {
                #region
                // sql = "select PID,PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC,isnull((select sum(YR_FUND) from TDV_TDDCFTS where PID = P.PID and PJV0050ord01 = P.PJ_CODE + P.CATY_CODE and SEQ_NO = '" + sSEQ_NO + "'),0) as YR_FUND from PJV_PJDCATY P where P.PID in (select PID from PJMFORM where SYR='" + PJ_SYR + "' and PID='" + AN_PID + "')";
                #endregion

                sql = $@"
                        SELECT 
                            P.PID, 
                            P.PJ_CODE + P.CATY_CODE AS CATY_CODE, 
                            P.CATY_NAME, 
                            P.ACC_ID, 
                            P.ACC_DESC, 
                            P.CATY_TYPE, 
                            P.CATY_DESC, 
                            ISNULL(
                                (SELECT SUM(YR_FUND) 
                                 FROM TDV_TDDCFTS 
                                 WHERE PID = P.PID 
                                 AND PJ_CODE = P.PJ_CODE 
                                 AND PJ_CATY_CODE = P.CATY_CODE
                                 AND SEQ_NO = '{sSEQ_NO}'
                                ), 0
                            ) AS YR_FUND
                        FROM PJV_PJDCATY P
                        LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID -- 添加 LEFT JOIN
                        WHERE P.PID IN 
                            (SELECT PID 
                             FROM PJMFORM 
                             WHERE SYR = '{PJ_SYR}' 
                             AND PID = '{AN_PID}'
                            )
                        AND PJM.Water_Type = '雨水使用'";
            }
            else
            {
                #region
                // sql = "select PID,PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC,isnull((select sum(YR_FUND) from TDV_TDDCFTS where PID = P.PID and PJV0050ord01 = P.PJ_CODE + P.CATY_CODE and SEQ_NO = '" + sSEQ_NO + "'),0) as YR_FUND from PJV_PJDCATY P where P.PID in (select PID from PJMFORM where SYR='" + PJ_SYR + "' and AN_PID='" + AN_PID + "')";
                #endregion

                sql = $@"
                        SELECT 
                            P.PID, 
                            P.PJ_CODE + P.CATY_CODE AS CATY_CODE, 
                            P.CATY_NAME, 
                            P.ACC_ID, 
                            P.ACC_DESC, 
                            P.CATY_TYPE, 
                            P.CATY_DESC, 
                            ISNULL(
                                (SELECT SUM(YR_FUND) 
                                 FROM TDV_TDDCFTS 
                                 WHERE PID = P.PID 
                                 AND PJ_CODE = P.PJ_CODE 
                                 AND PJ_CATY_CODE = P.CATY_CODE 
                                 AND SEQ_NO = '{sSEQ_NO}'
                                ), 0
                            ) AS YR_FUND
                        FROM PJV_PJDCATY P
                        LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID -- 添加 LEFT JOIN 到 PJMFORM
                        WHERE P.PID IN 
                            (SELECT PID 
                             FROM PJMFORM 
                             WHERE SYR = '{PJ_SYR}' 
                             AND AN_PID = '{AN_PID}' -- 保持 AN_PID 條件
                            )
                        AND PJM.Water_Type = '雨水使用'";
            }

            dt4 = new DataTable();
            dt4 = SourceFunction.DataTableGet(sql, connectionString);

            dtp4 = new DataTable();
            dtp4.TableName = "dtp4";
            dtp4.Columns.Add("PID");
            dtp4.Columns.Add("CATY_CODE");
            dtp4.Columns.Add("CATY_NAME");
            dtp4.Columns.Add("ACC_ID");
            dtp4.Columns.Add("ACC_DESC");
            dtp4.Columns.Add("CATY_TYPE");
            dtp4.Columns.Add("CATY_DESC");
            dtp4.Columns.Add("YR_FUND");
            dtp4.Columns.Add("PYR_FUND");

            for (i = 0; i < ditdt.Rows.Count; i++)
            {
                dtp4.Columns.Add("PYR_YR_FUND_" + ditdt.Rows[i]["PID_SYR"].ToString()); //以前年度
                dtp4.Columns.Add("PYR_PYR_FUND_" + ditdt.Rows[i]["PID_SYR"].ToString());//以前年度
            }

            for (i = 0; i < dt4.Rows.Count; i++)
            {
                dr = dtp4.NewRow();
                dr["PID"] = dt4.Rows[i]["PID"].ToString();
                dr["CATY_CODE"] = dt4.Rows[i]["CATY_CODE"].ToString();
                dr["CATY_NAME"] = dt4.Rows[i]["CATY_NAME"].ToString();
                dr["ACC_ID"] = dt4.Rows[i]["ACC_ID"].ToString();
                dr["ACC_DESC"] = dt4.Rows[i]["ACC_DESC"].ToString();
                dr["CATY_TYPE"] = dt4.Rows[i]["CATY_TYPE"].ToString();
                dr["CATY_DESC"] = dt4.Rows[i]["CATY_DESC"].ToString();
                dr["YR_FUND"] = dt4.Rows[i]["YR_FUND"].ToString();
                dtp4.Rows.Add(dr);
            }

            if (sPSEQ_NO != "")
            {
                if (plan is null)
                {
                    #region
                    // sql = "select PID,PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC,isnull((select sum(YR_FUND) from TDV_TDDHFTS where PID = P.PID and PJV0050ord01 = P.PJ_CODE + P.CATY_CODE and SEQ_NO = '" + sPSEQ_NO + "'),0) as YR_FUND from PJV_PJDCATY P where P.PID in (select PID from PJMFORM where SYR='" + PJ_SYR + "' and PID='" + AN_PID + "')";
                    #endregion

                    sql = $@"
                            SELECT 
                                P.PID, 
                                P.PJ_CODE + P.CATY_CODE AS CATY_CODE, 
                                P.CATY_NAME, 
                                P.ACC_ID, 
                                P.ACC_DESC, 
                                P.CATY_TYPE, 
                                P.CATY_DESC, 
                                ISNULL(
                                    (SELECT SUM(YR_FUND) 
                                     FROM TDV_TDDHFTS 
                                     WHERE PID = P.PID 
                                     AND PJ_CODE = P.PJ_CODE 
                                     AND PJ_CATY_CODE = P.CATY_CODE 
                                     AND SEQ_NO = '{sPSEQ_NO}'
                                    ), 0
                                ) AS YR_FUND
                            FROM PJV_PJDCATY P
                            LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID 
                            WHERE P.PID IN 
                                (SELECT PID 
                                 FROM PJMFORM 
                                 WHERE SYR = '{PJ_SYR}' 
                                 AND PID = '{AN_PID}'
                                )
                            AND PJM.Water_Type = '雨水使用' ";

                }
                else
                {
                    #region
                    // sql = "select PID,PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC,isnull((select sum(YR_FUND) from TDV_TDDHFTS where PID = P.PID and PJV0050ord01 = P.PJ_CODE + P.CATY_CODE and SEQ_NO = '" + sPSEQ_NO + "'),0) as YR_FUND from PJV_PJDCATY P where P.PID in (select PID from PJMFORM where SYR='" + PJ_SYR + "' and AN_PID='" + AN_PID + "')";
                    #endregion

                    sql = $@"
                            SELECT 
                                P.PID, 
                                P.PJ_CODE + P.CATY_CODE AS CATY_CODE, 
                                P.CATY_NAME, 
                                P.ACC_ID, 
                                P.ACC_DESC, 
                                P.CATY_TYPE, 
                                P.CATY_DESC, 
                                ISNULL(
                                    (SELECT SUM(YR_FUND) 
                                     FROM TDV_TDDHFTS 
                                     WHERE PID = P.PID 
                                     AND PJ_CODE = P.PJ_CODE 
                                     AND PJ_CATY_CODE = P.CATY_CODE 
                                     AND SEQ_NO = '{sPSEQ_NO}'
                                    ), 0
                                ) AS YR_FUND
                            FROM PJV_PJDCATY P
                            LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID 
                            WHERE P.PID IN 
                                (SELECT PID 
                                 FROM PJMFORM 
                                 WHERE SYR = '{PJ_SYR}' 
                                 AND AN_PID = '{AN_PID}'
                                )
                            AND PJM.Water_Type = '雨水使用' ";
                }
            }
            else
            {
                if (plan is null)
                {
                    #region
                    //sql = "select PID,PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC,isnull((select sum(YR_FUND) from TDV_TDDHFTS where PID = P.PID and PJV0050ord01 = P.PJ_CODE + P.CATY_CODE and SEQ_NO = '999'),0) as YR_FUND from PJV_PJDCATY P where P.PID in (select PID from PJMFORM where SYR='" + PJ_SYR + "' and PID='" + AN_PID + "')";
                    #endregion

                    sql = $@"
                            SELECT 
                                P.PID, 
                                P.PJ_CODE + P.CATY_CODE AS CATY_CODE, 
                                P.CATY_NAME, 
                                P.ACC_ID, 
                                P.ACC_DESC, 
                                P.CATY_TYPE, 
                                P.CATY_DESC, 
                                ISNULL(
                                    (SELECT SUM(YR_FUND) 
                                     FROM TDV_TDDHFTS 
                                     WHERE PID = P.PID 
                                     AND PJ_CODE = P.PJ_CODE 
                                     AND PJ_CATY_CODE = P.CATY_CODE 
                                     AND SEQ_NO = '999'), 0
                                ) AS YR_FUND
                            FROM PJV_PJDCATY P
                            LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID 
                            WHERE P.PID IN 
                                (SELECT PID 
                                 FROM PJMFORM 
                                 WHERE SYR = '{PJ_SYR}' 
                                 AND PID = '{AN_PID}' 
                                )
                            AND PJM.Water_Type = '雨水使用' ";
                }
                else
                {
                    #region
                    // sql = "select PID,PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC,isnull((select sum(YR_FUND) from TDV_TDDHFTS where PID = P.PID and PJV0050ord01 = P.PJ_CODE + P.CATY_CODE and SEQ_NO = '999'),0) as YR_FUND from PJV_PJDCATY P where P.PID in (select PID from PJMFORM where SYR='" + PJ_SYR + "' and AN_PID='" + AN_PID + "')";
                    #endregion

                    sql = $@"
                            SELECT 
                                P.PID, 
                                P.PJ_CODE + P.CATY_CODE AS CATY_CODE, 
                                P.CATY_NAME, 
                                P.ACC_ID, 
                                P.ACC_DESC, 
                                P.CATY_TYPE, 
                                P.CATY_DESC, 
                                ISNULL(
                                    (SELECT SUM(YR_FUND) 
                                     FROM TDV_TDDHFTS 
                                     WHERE PID = P.PID 
                                     AND PJ_CODE = P.PJ_CODE 
                                     AND PJ_CATY_CODE = P.CATY_CODE 
                                     AND SEQ_NO = '999'), 0
                                ) AS YR_FUND
                            FROM PJV_PJDCATY P
                            LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID 
                            WHERE P.PID IN 
                                (SELECT PID 
                                 FROM PJMFORM 
                                 WHERE SYR = '{PJ_SYR}' 
                                 AND AN_PID = '{AN_PID}' 
                                )
                            AND PJM.Water_Type = '雨水使用' ";
                }
            }

            dt4 = new DataTable();

            dt4 = SourceFunction.DataTableGet(sql, connectionString);

            for (i = 0; i < dtp4.Rows.Count; i++)
            {
                rr = dt4.Select("CATY_CODE = '" + dtp4.Rows[i]["CATY_CODE"].ToString() + "'");

                if (rr.Length > 0)
                {
                    dtp4.Rows[i]["PYR_FUND"] = rr[0]["YR_FUND"].ToString();
                }
                else
                {
                    dtp4.Rows[i]["PYR_FUND"] = "0";
                }
            }

            var TDV_TDDCFTS = (from a in db2.TDV_TDDCFTS
                               join b in db2.PJV_PJDCATY on a.PID + a.PJV0050Ord01 equals b.PID + b.PJ_CODE + b.CATY_CODE
                               select new
                               {
                                   PID = a.PID,
                                   TD_NAME = a.TD_NAME,
                                   YR_FUND = a.YR_FUND
                               }).ToList();

            /// 以前年度
            if (plan is null)
            {
                #region
                //sql = "select PID,PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC,isnull((select sum(YR_FUND) from TDV_TDDCFTS where PID = P.PID and PJV0050ord01 = P.PJ_CODE + P.CATY_CODE),0) as YR_FUND from PJV_PJDCATY P where P.PID in (select PID from PJMFORM where SYR<='" + sPSYR + "' and PID=" + AN_PID + ")";
                #endregion

                sql = $@"
                        SELECT 
                            P.PID, 
                            P.PJ_CODE + P.CATY_CODE AS CATY_CODE, 
                            P.CATY_NAME, 
                            P.ACC_ID, 
                            P.ACC_DESC, 
                            P.CATY_TYPE, 
                            P.CATY_DESC, 
                            ISNULL(
                                (SELECT SUM(YR_FUND) 
                                 FROM TDV_TDDCFTS 
                                 WHERE PID = P.PID 
                                 AND PJ_CODE = P.PJ_CODE 
                                 AND PJ_CATY_CODE = P.CATY_CODE
                                ), 0
                            ) AS YR_FUND
                        FROM PJV_PJDCATY P
                        LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID -- 加入 LEFT JOIN 條件
                        WHERE P.PID IN 
                            (SELECT PID 
                             FROM PJMFORM 
                             WHERE SYR <= '{sPSYR}' 
                             AND PID = {AN_PID}
                            )
                        AND PJM.Water_Type = '雨水使用' -- 加入 Water_Type 條件";
            }
            else
            {
                #region 
                //sql = "select PID,PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC,isnull((select sum(YR_FUND) from TDV_TDDCFTS where PID = P.PID and PJV0050ord01 = P.PJ_CODE + P.CATY_CODE),0) as YR_FUND from PJV_PJDCATY P where P.PID in (select PID from PJMFORM where SYR<='" + sPSYR + "' and AN_PID=" + AN_PID + ")";
                #endregion

                sql = $@"
                        SELECT 
                            P.PID, 
                            P.PJ_CODE + P.CATY_CODE AS CATY_CODE, 
                            P.CATY_NAME, 
                            P.ACC_ID, 
                            P.ACC_DESC, 
                            P.CATY_TYPE, 
                            P.CATY_DESC, 
                            ISNULL(
                                (SELECT SUM(YR_FUND) 
                                 FROM TDV_TDDCFTS 
                                 WHERE PID = P.PID 
                                 AND PJV0050ord01 = P.PJ_CODE + P.CATY_CODE
                                ), 0
                            ) AS YR_FUND 
                        FROM PJV_PJDCATY P
                        LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID -- 添加 LEFT JOIN PJMFORM
                        WHERE P.PID IN 
                            (SELECT PID 
                             FROM PJMFORM 
                             WHERE SYR <= '{sPSYR}' 
                             AND AN_PID = '{AN_PID}'
                            )
                        AND PJM.Water_Type = '雨水使用'";
            }

            dt4 = new DataTable();
            dt4 = SourceFunction.DataTableGet(sql, connectionString);

            for (i = 0; i < dtp4.Rows.Count; i++)
            {
                rr = dt4.Select("CATY_CODE = '" + dtp4.Rows[i]["CATY_CODE"].ToString() + "'");

                if (rr.Length > 0)
                {
                    for (k = 0; k < rr.Length; k++)
                    {
                        for (j = 0; j < ditdt.Rows.Count; j++)
                        {
                            dtp4.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = rr[k]["YR_FUND"].ToString();
                            dtp4.Rows[i]["PYR_PYR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = "0"; /// 調整後欄位設為 0
                        }
                    }
                }
                else
                {
                    for (j = 0; j < ditdt.Rows.Count; j++)
                    {
                        dtp4.Rows[i]["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = "0";
                        dtp4.Rows[i]["PYR_PYR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = "0";
                    }
                }
            }

            ds.Tables.Add(dtp4);

            #endregion -------------

            #region

            dtp = dt.Clone();

            tempCITY = "";

            if (dt.Rows.Count > 0)
            {
                tempCITY = dt.Rows[0]["DATACITY"].ToString();
            }

            /// I-1 最後生成資料
            for (i = 0; i < dt.Rows.Count; i++)
            {
                r = dt.Rows[i];

                if (tempCITY != "" && tempCITY != r["DATACITY"].ToString())
                {
                    dr = dtp.NewRow();
                    dr["PID"] = "";
                    dr["PID_SYR"] = "";
                    dr["SEQ_NO"] = "";
                    dr["TID"] = "";
                    dr["FD_CODE"] = "";
                    dr["TD_CODE"] = "";
                    //dr["TD_NAME"] = "小計";
                    dr["SIP_TD_CODE"] = "";
                    dr["CITY"] = "";
                    dr["TOWN"] = "";
                    dr["ADM_DIST"] = "";
                    dr["IPJ_CODE"] = "";
                    dr["IPJ_CODE_DESC"] = "";
                    dr["N_IPJ_CODE"] = "";
                    dr["IPJ_DESC"] = "";
                    dr["YR_FUND"] = getDtColumns_SUM(dt, "YR_FUND", "DATACITY='" + tempCITY + "'");
                    dr["CEN_FUND"] = getDtColumns_SUM(dt, "CEN_FUND", "DATACITY='" + tempCITY + "'");
                    dr["LOC_FUND"] = getDtColumns_SUM(dt, "LOC_FUND", "DATACITY='" + tempCITY + "'");
                    dr["PYR_FUND"] = getDtColumns_SUM(dt, "PYR_FUND", "DATACITY='" + tempCITY + "'");
                    dr["PCEN_FUND"] = getDtColumns_SUM(dt, "PCEN_FUND", "DATACITY='" + tempCITY + "'");
                    dr["PLOC_FUND"] = getDtColumns_SUM(dt, "PLOC_FUND", "DATACITY='" + tempCITY + "'");

                    for (j = 0; j < ditdt.Rows.Count; j++)
                    {
                        dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and PID_SYR='" + ditdt.Rows[j]["PID_SYR"].ToString() + "'");
                        dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and PID_SYR='" + ditdt.Rows[j]["PID_SYR"].ToString() + "'");
                        dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and PID_SYR='" + ditdt.Rows[j]["PID_SYR"].ToString() + "'");
                    }

                    dr["FROM"] = "";
                    dtp.Rows.Add(dr);
                    tempCITY = r["DATACITY"].ToString();
                }

                dr = dtp.NewRow();

                dr["PID"] = r["PID"].ToString();
                dr["PID_SYR"] = r["PID_SYR"].ToString();
                dr["SEQ_NO"] = r["SEQ_NO"].ToString();
                dr["TID"] = r["TID"].ToString();
                dr["FD_CODE"] = r["FD_CODE"].ToString();
                dr["TD_CODE"] = r["TD_CODE"].ToString();
                dr["TD_NAME"] = r["TD_NAME"].ToString();
                dr["SIP_TD_CODE"] = r["SIP_TD_CODE"].ToString();
                dr["CITY"] = r["CITY"].ToString();
                dr["TOWN"] = r["TOWN"].ToString();
                dr["ADM_DIST"] = r["ADM_DIST"].ToString();
                dr["IPJ_CODE"] = r["IPJ_CODE"].ToString();
                dr["IPJ_CODE_DESC"] = r["IPJ_CODE_DESC"].ToString();
                dr["N_IPJ_CODE"] = r["N_IPJ_CODE"].ToString();
                dr["IPJ_DESC"] = r["IPJ_DESC"].ToString();
                dr["CATY_TYPE"] = r["CATY_TYPE"].ToString();
                dr["YR_FUND"] = r["YR_FUND"].ToString();
                dr["CEN_FUND"] = r["CEN_FUND"].ToString();
                dr["LOC_FUND"] = r["LOC_FUND"].ToString();
                dr["PYR_FUND"] = r["PYR_FUND"].ToString();
                dr["PCEN_FUND"] = r["PCEN_FUND"].ToString();
                dr["PLOC_FUND"] = r["PLOC_FUND"].ToString();

                for (j = 0; j < ditdt.Rows.Count; j++)
                {
                    dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                    dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                    dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = r["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()].ToString();
                }

                dr["FROM"] = r["FROM"].ToString();
                dr["PJV0050ord01"] = r["PJV0050ord01"].ToString();
                dr["PJV0050ord02"] = r["PJV0050ord02"].ToString();
                dtp.Rows.Add(dr);
            }

            if (dt.Rows.Count > 0)
            {
                tempCITY = dt.Rows[dt.Rows.Count - 1]["DATACITY"].ToString();

                /// 最後一個縣市
                dr = dtp.NewRow();
                dr["PID"] = "";
                dr["PID_SYR"] = "";
                dr["SEQ_NO"] = "";
                dr["TID"] = "";
                dr["FD_CODE"] = "";
                dr["TD_CODE"] = "";
                dr["TD_NAME"] = "小計";
                dr["SIP_TD_CODE"] = "";
                dr["CITY"] = "";
                dr["TOWN"] = "";
                dr["ADM_DIST"] = "";
                dr["IPJ_CODE"] = "";
                dr["IPJ_CODE_DESC"] = "";
                dr["N_IPJ_CODE"] = "";
                dr["IPJ_DESC"] = "";
                dr["YR_FUND"] = getDtColumns_SUM(dt, "YR_FUND", "");
                dr["CEN_FUND"] = getDtColumns_SUM(dt, "CEN_FUND", "");
                dr["LOC_FUND"] = getDtColumns_SUM(dt, "LOC_FUND", "");
                dr["PYR_FUND"] = getDtColumns_SUM(dt, "PYR_FUND", "");
                dr["PCEN_FUND"] = getDtColumns_SUM(dt, "PCEN_FUND", "DATACITY='" + tempCITY + "'");
                dr["PLOC_FUND"] = getDtColumns_SUM(dt, "PLOC_FUND", "DATACITY='" + tempCITY + "'");

                for (j = 0; j < ditdt.Rows.Count; j++)
                {
                    dr["PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_YR_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and PID_SYR='" + ditdt.Rows[j]["PID_SYR"].ToString() + "'");
                    dr["PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_CEN_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and PID_SYR='" + ditdt.Rows[j]["PID_SYR"].ToString() + "'");
                    dr["PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString()] = getDtColumns_SUM(dt, "PYR_LOC_FUND_" + ditdt.Rows[j]["PID_SYR"].ToString(), "DATACITY='" + tempCITY + "' and PID_SYR='" + ditdt.Rows[j]["PID_SYR"].ToString() + "'");
                }
                dr["FROM"] = "";
                dtp.Rows.Add(dr);
            }

            dtp.TableName = "dtp";
            ditdt.TableName = "ditdt";

            ds.Tables.Add(dtp);
            ds.Tables.Add(ditdt);

            #endregion

            return ds;
        }

        public static DataView GetDvOnlyRead(string sql, string ConnStr)
        {
            string text = "";
            DataView result = new DataView();
            if (sql.Trim().Length == 0)
            {
                result = null;
            }
            ConnStr = "Provider=SQLOLEDB;" + ConnStr;
            OleDbConnection oleDbConnection = new OleDbConnection(ConnStr);
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDbConnection);
            oleDbCommand.CommandTimeout = 6000;
            DataTable dataTable = new DataTable();
            try
            {
                oleDbConnection.Open();
                dataTable.Load(oleDbCommand.ExecuteReader());
                result = new DataView(dataTable);
                dataTable.Dispose();
            }
            catch (Exception ex)
            {
                text = sql + " error-" + (ex.Message); //////錯誤訊息需自細看
            }
            finally
            {
                oleDbConnection.Close();
            }

            return result;
        }

        #region =============================================== 新版報表 ===============================================

        /// <summary>
        /// I-1.雨水例行性管考報表(資料來源)
        /// </summary>
        /// <param name="AN_PID"> 工程編號 </param>
        /// <param name="PJ_SYR"> 年份 </param>
        /// <param name="PJ_SMN"> 月份 </param>
        /// <param name="connectionString"> 連接字串 </param>
        /// <param name="DEPT"> 部門 </param>
        /// <see cref="RoutineRainwaterInspectionReport"/> 總表
        /// <see cref="FundingAdjustmentTable"/>           經費調整表
        /// <see cref="FundingAdjustmentTableDetails"/>    經費調整表_明細
        public static DataSet New_GetRPT0001Data(decimal AN_PID, string PJ_SYR, string PJ_SMN, string connectionString, string DEPT)
        {
            string OpID = DEPT;
            string sql = "";
            bool isMatch;

            DataSet ds = new DataSet();          // 儲存最終結果的 DataSet
            DataTable dt = new DataTable();      // 一般用途的 DataTable
            DataTable dtp = new DataTable();     // 多用途的 DataTable
            DataTable dtp1 = new DataTable();    // 署辦(IPJ_CODE='01' and ACC_ID<>'0200')
            DataTable dtp2 = new DataTable();    // 地辦(IPJ_CODE='02')
            DataTable dtp3 = new DataTable();    // 業務費(IPJ_CODE='01' and ACC_ID='0200')
            DataTable dtp4 = new DataTable();    // 經費調整表-明細
            DataTable dtp5 = new DataTable();    // 經費調整表
            DataTable dtp6 = new DataTable();    // NEW_經費調整表
            DataTable PJDCATY = new DataTable(); // 計畫類別資料表
            DataTable dt1 = new DataTable();     // 本期資料
            DataTable dt2 = new DataTable();     // 前期資料
            DataTable dt3 = new DataTable();     // 前一年度資料
            DataTable dt4 = new DataTable();     // 始終抓取原預算
            DataTable dt6 = new DataTable();     // NEW_經費調整表
            DataTable ditdt = new DataTable();   // 儲存臨時資料的 DataTable

            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");      // 取得民國年
            string sSMN = (System.DateTime.Now.Month).ToString("00");             // 取得月份
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000"); // 取得前一年度
            string sSEQ_NO = "0"; // 預設最大序列號
            string sPSEQ_NO = ""; // 預設次大序列號

            #region ---------------- 1.PJ_SYR(年度) PJ_SMN(月份)
            // 傳入參數 PJ_SYR 不為 null，該值覆蓋當前年度，並取得前一年度
            // 傳入參數 PJ_SMN 不為 null，該值覆蓋當前月份

            if (PJ_SYR != "")
            {
                sSYR = PJ_SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
            }

            if (PJ_SMN != "")
            {
                sSMN = PJ_SMN;
            }

            #endregion -------------

            #region ---------------- 2.Plan = AN_PID 查詢 PJMFORM 結果

            var plan = db2.PJMFORM.Where(m => m.AN_PID == AN_PID).FirstOrDefault();

            #endregion -------------

            #region ---------------- 3.PJDCATY 和 DT

            /// 如果是分支計畫的情況
            if (plan is null)
            {
                sql = "select PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC from PJV_PJDCATY where PID in (select PID from PJMFORM where PID=" + AN_PID + ")";
                PJDCATY = SourceFunction.DataTableGet(sql, connectionString);

                sql = "select max(SEQ_NO) as MAX_SEQ_NO,max(SEQ_NO)-1 as pMAX_SEQ_NO,min(SEQ_NO) as MIN_SEQ_NO from PJDFDCR where PID in (select PID from PJMFORM where PID = " + AN_PID + " and ADV_DP = '" + OpID + "' and SYR='" + sSYR + "') and case isnull(APROV_DT, isnull(AP_DT, '')) when '' then '' else isnull(APROV_DT, isnull(AP_DT, '')) end <> '' and isnull(APROV_DT, isnull(AP_DT,''))< '" + sSYR + sSMN + "32'";
                dt = SourceFunction.DataTableGet(sql, connectionString);
            }
            else
            {
                sql = "select PJ_CODE + CATY_CODE as CATY_CODE,CATY_NAME,ACC_ID,ACC_DESC,CATY_TYPE,CATY_DESC from PJV_PJDCATY where PID in (select PID from PJMFORM where AN_PID=" + AN_PID + ")";
                PJDCATY = SourceFunction.DataTableGet(sql, connectionString);

                sql = "select max(SEQ_NO) as MAX_SEQ_NO,max(SEQ_NO)-1 as pMAX_SEQ_NO,min(SEQ_NO) as MIN_SEQ_NO from PJDFDCR where PID in (select PID from PJMFORM where AN_PID = " + AN_PID + " and ADV_DP = '" + OpID + "' and SYR='" + sSYR + "') and case isnull(APROV_DT, isnull(AP_DT, '')) when '' then '' else isnull(APROV_DT, isnull(AP_DT, '')) end <> '' and isnull(APROV_DT, isnull(AP_DT,''))< '" + sSYR + sSMN + "32'";
                dt = SourceFunction.DataTableGet(sql, connectionString);
            }

            Debug.WriteLine($"sql = {sql}");

            #endregion -------------

            #region ---------------- 4.MAX_SEQ_NO(最大) / pMAX_SEQ_NO(次大)

            /// sSEQ_NO  => MAX_SEQ_NO(最大)
            /// sPSEQ_NO => pMAX_SEQ_NO(次大)

            if (dt != null && dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["MAX_SEQ_NO"].ToString() != "")
                {
                    sSEQ_NO = dt.Rows[0]["MAX_SEQ_NO"].ToString();
                }

                if (dt.Rows[0]["pMAX_SEQ_NO"].ToString() != "")
                {
                    sPSEQ_NO = dt.Rows[0]["pMAX_SEQ_NO"].ToString();
                }
            }

            #endregion -------------

            #region  DT1 DataTable 本期資料

            dt1 = new DataTable();

            if (plan is null)
            {
                sql = $@"
                        SELECT 
                            TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                            TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                            TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                            TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                            TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                            TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type
                        FROM CMEPMS.dbo.TDV_TDDCFTS TDV
                        LEFT JOIN CMEPMS.dbo.PJMFORM PJM ON TDV.PID = PJM.PID
                        WHERE TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PID = {AN_PID})
                        AND TDV.PID_SYR = '{PJ_SYR}' 
                        AND TDV.SEQ_NO = '{sSEQ_NO}' 
                        AND PJM.Water_Type = '雨水使用'
                        ORDER BY TDV.CITY ";
            }
            else
            {
                sql = $@"
                        SELECT 
                            TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                            TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                            TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                            TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                            TDV.YR_FUND, TDV.CEN_FUND,TDV.ACC_DESC, TDV.LOC_FUND, TDV.PJV0050ord01, 
                            TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type
                        FROM CMEPMS.dbo.TDV_TDDCFTS TDV
                        LEFT JOIN CMEPMS.dbo.PJMFORM PJM ON TDV.PID = PJM.PID
                        WHERE TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE AN_PID = {AN_PID})
                        AND TDV.PID_SYR = '{PJ_SYR}' 
                        AND TDV.SEQ_NO = '{sSEQ_NO}' 
                        AND PJM.Water_Type = '雨水使用'
                        ORDER BY TDV.CITY ";
            }

            dt1 = SourceFunction.DataTableGet(sql, connectionString);

            #endregion -------------

            #region DT2 DataTable 新本期/前期資料

            if (dt1 != null && dt1.Rows.Count > 0)
            {
                if (sPSEQ_NO != "")
                {
                    if (plan is null) /// 包含"分支計畫"
                    {
                        sql = $@"
                                SELECT 
                                    TDV.PID,TDV.PID_SYR,TDV.SEQ_NO,TDV.TID, 
                                    TDV.FD_CODE,TDV.TD_CODE,TDV.TD_NAME,TDV.SIP_TD_CODE, 
                                    TDV.CITY,TDV.TOWN,TDV.ADM_DIST,TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC,TDV.N_IPJ_CODE, 
                                    TDV.IPJ_DESC,TDV.CATY_TYPE,TDV.YR_FUND,TDV.CEN_FUND, 
                                    TDV.LOC_FUND,TDV.PJV0050ord01,TDV.PJV0050Ord02,TDV.ACC_ID,PJM.Water_Type
                                FROM CMEPMS.dbo.TDV_TDDCFTS TDV
                                LEFT JOIN CMEPMS.dbo.PJMFORM PJM 
                                ON TDV.PID = PJM.PID
                                WHERE 
                                TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PID = {AN_PID})
                                AND TDV.PID_SYR = '{PJ_SYR}' 
                                AND TDV.SEQ_NO = '{sSEQ_NO}' 
                                AND PJM.Water_Type = '雨水使用'
                                ORDER BY TDV.SIP_TD_CODE DESC";
                    }
                    else
                    {
                        /// 20241112 修改項目 => 新增查詢條件，修正至查詢前一期
                        if (sPSEQ_NO != null)
                        {
                            sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02,TDV.ACC_DESC, TDV.ACC_ID, PJM.Water_Type 
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE AN_PID = {AN_PID})
                                AND TDV.PID_SYR = '{PJ_SYR}' 
                                AND TDV.SEQ_NO = '{sPSEQ_NO}'
                                AND PJM.Water_Type = '雨水使用' 
                                ORDER BY TDV.CITY ";
                        }
                        else
                        {
                            sql = $@"
                                SELECT 
                                    TDV.PID,TDV.PID_SYR,TDV.SEQ_NO,TDV.TID, 
                                    TDV.FD_CODE,TDV.TD_CODE,TDV.TD_NAME,TDV.SIP_TD_CODE, 
                                    TDV.CITY,TDV.TOWN,TDV.ADM_DIST,TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC,TDV.N_IPJ_CODE,TDV.ACC_DESC, 
                                    TDV.IPJ_DESC,TDV.CATY_TYPE,TDV.YR_FUND,TDV.CEN_FUND, 
                                    TDV.LOC_FUND,TDV.PJV0050ord01,TDV.PJV0050Ord02,TDV.ACC_ID,PJM.Water_Type
                                FROM CMEPMS.dbo.TDV_TDDCFTS TDV
                                LEFT JOIN CMEPMS.dbo.PJMFORM PJM 
                                ON TDV.PID = PJM.PID
                                WHERE 
                                TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE AN_PID = {AN_PID})
                                AND TDV.PID_SYR = '{PJ_SYR}' 
                                AND TDV.SEQ_NO = '{sSEQ_NO}' 
                                AND PJM.Water_Type = '雨水使用'
                                ORDER BY TDV.CITY ";
                        }
                    }
                }
                else
                {
                    if (plan is null) /// 包含"分支計畫"
                    {
                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type 
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE PID = {AN_PID} AND ADV_DP = '{OpID}') 
                                AND TDV.SEQ_NO = '999'
                                AND PJM.Water_Type = '雨水使用'
                                ORDER BY TDV.SIP_TD_CODE DESC ";
                    }
                    else
                    {
                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type 
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE AN_PID = {AN_PID} AND ADV_DP = '{OpID}') 
                                AND TDV.SEQ_NO = '999'
                                AND PJM.Water_Type = '雨水使用'
                                ORDER BY TDV.CITY ";
                    }
                }
                dt2 = SourceFunction.DataTableGet(sql, connectionString);
            }
            else
            {
                dt1 = new DataTable();

                if (plan is null) /// 包含"分支計畫"
                {
                    sql = $@"
                            SELECT 
                                TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type -- 加入 Water_Type 欄位
                            FROM CMEPMS.dbo.TDV_TDDHFTS TDV
                            LEFT JOIN CMEPMS.dbo.PJMFORM PJM ON TDV.PID = PJM.PID -- 進行 LEFT JOIN
                            WHERE TDV.PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PID = {AN_PID}) 
                            AND TDV.SEQ_NO = '{sSEQ_NO}' 
                            AND PJM.Water_Type = '雨水使用'
                            ORDER BY TDV.CITY ";
                }
                else
                {

                    sql = $@"
                            SELECT 
                                TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type 
                            FROM TDV_TDDHFTS TDV
                            LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID -- 加入 LEFT JOIN
                            WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE AN_PID = {AN_PID}) 
                            AND TDV.PID_SYR = '{PJ_SYR}' 
                            AND TDV.SEQ_NO = '{sSEQ_NO}' 
                            AND PJM.Water_Type = '雨水使用'
                            ORDER BY TDV.CITY ";
                }

                dt1 = SourceFunction.DataTableGet(sql, connectionString);

                if (sPSEQ_NO != "")
                {
                    if (plan is null)
                    {
                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE PID = {AN_PID}) 
                                AND TDV.SEQ_NO = '{sPSEQ_NO}'
                                AND PJM.Water_Type = '雨水使用' 
                                ORDER BY TDV.CITY ";
                    }
                    else
                    {
                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type 
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE AN_PID = {AN_PID}) 
                                AND TDV.PID_SYR = '{PJ_SYR}' 
                                AND TDV.SEQ_NO = '{sPSEQ_NO}'
                                AND PJM.Water_Type = '雨水使用' 
                                ORDER BY TDV.CITY ";
                    }
                }
                else
                {
                    if (plan is null)
                    {
                        sql = $@"
                                SELECT 
                                    TDV.PID, TDV.PID_SYR, TDV.SEQ_NO, TDV.TID, 
                                    TDV.FD_CODE, TDV.TD_CODE, TDV.TD_NAME, TDV.SIP_TD_CODE, 
                                    TDV.CITY, TDV.TOWN, TDV.ADM_DIST, TDV.IPJ_CODE, 
                                    TDV.IPJ_CODE_DESC, TDV.N_IPJ_CODE, TDV.IPJ_DESC, TDV.CATY_TYPE, 
                                    TDV.YR_FUND, TDV.CEN_FUND, TDV.LOC_FUND, TDV.PJV0050ord01, 
                                    TDV.PJV0050Ord02, TDV.ACC_ID, PJM.Water_Type
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE PID = {AN_PID} AND ADV_DP = '{OpID}') 
                                AND TDV.SEQ_NO = '999'
                                AND PJM.Water_Type = '雨水使用' 
                                ORDER BY TDV.CITY ";
                    }
                    else
                    {
                        sql = $@"
                                SELECT 
                                    TDV.PID,TDV.PID_SYR,TDV.SEQ_NO,TDV.TID, 
                                    TDV.FD_CODE,TDV.TD_CODE,TDV.TD_NAME,TDV.SIP_TD_CODE, 
                                    TDV.CITY,TDV.TOWN,TDV.ADM_DIST,TDV.IPJ_CODE,TDV.IPJ_CODE_DESC, 
                                    TDV.N_IPJ_CODE,TDV.IPJ_DESC,TDV.CATY_TYPE,TDV.YR_FUND,TDV.CEN_FUND, 
                                    TDV.LOC_FUND,TDV.PJV0050ord01,TDV.PJV0050Ord02,TDV.ACC_ID,PJM.Water_Type
                                FROM TDV_TDDHFTS TDV
                                LEFT JOIN PJMFORM PJM ON TDV.PID = PJM.PID 
                                WHERE TDV.PID IN (SELECT PID FROM PJMFORM WHERE AN_PID = {AN_PID}) 
                                AND TDV.SEQ_NO = '999'
                                AND PJM.Water_Type = '雨水使用' 
                                ORDER BY TDV.CITY ";
                    }
                }

                dt2 = SourceFunction.DataTableGet(sql, connectionString);
            }

            #endregion -------------

            #region ---------------- 5.經費調整表 / 經費調整表-明細 ----------------

            /// 本期/當期
            if (plan is null)
            {
                sql = $@"
                        SELECT 
                            P.PID, P.PJ_CODE + P.FCATY_CODE AS CATY_CODE, P.CATY_NAME, P.ACC_ID, 
                            P.ACC_DESC, P.CATY_TYPE, P.CATY_DESC, 
                            ISNULL(
                                (SELECT SUM(YR_FUND) 
                                 FROM TDV_TDDCFTS 
                                 WHERE PID = P.PID 
                                 AND PJ_CODE = P.PJ_CODE 
                                 AND PJ_CATY_CODE = P.CATY_CODE
                                 AND SEQ_NO = '{sSEQ_NO}'
                                ), 0
                            ) AS YR_FUND
                        FROM PJV_PJDCATY P
                        LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID -- 添加 LEFT JOIN
                        WHERE P.PID IN 
                            (SELECT PID 
                             FROM PJMFORM 
                             WHERE SYR = '{PJ_SYR}' 
                             AND PID = '{AN_PID}'
                            )
                        AND PJM.Water_Type = '雨水使用'";
            }
            else
            {
                sql = $@"
                        SELECT 
                            P.PID, P.PJ_CODE + P.FCATY_CODE AS CATY_CODE, P.CATY_NAME, P.ACC_ID, 
                            P.ACC_DESC, P.CATY_TYPE, P.CATY_DESC, 
                            ISNULL(
                                (SELECT SUM(YR_FUND) 
                                 FROM TDV_TDDCFTS 
                                 WHERE PID = P.PID 
                                 AND PJ_CODE = P.PJ_CODE 
                                 AND PJ_CATY_CODE = P.CATY_CODE 
                                 AND SEQ_NO = '{sSEQ_NO}'
                                ), 0
                            ) AS YR_FUND,
                            '{sSEQ_NO}' AS SEQ_NO, 
                            ISNULL(
                                (SELECT SUM(CEN_FUND) 
                                 FROM TDV_TDDCFTS 
                                 WHERE PID = P.PID 
                                 AND PJ_CODE = P.PJ_CODE 
                                 AND PJ_CATY_CODE = P.CATY_CODE 
                                 AND SEQ_NO = '{sSEQ_NO}'
                                ), 0
                            ) AS CEN_FUND
                        FROM PJV_PJDCATY P
                        LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID
                        WHERE P.PID IN 
                            (SELECT PID 
                             FROM PJMFORM 
                             WHERE SYR = '{PJ_SYR}' 
                             AND PID = '{AN_PID}'
                            )
                        AND PJM.Water_Type = '雨水使用'";

                Debug.WriteLine($"經費調整表(2):{sql}");

                dt3 = new DataTable();
                dt3 = SourceFunction.DataTableGet(sql, connectionString);

                isMatch = dt3.AsEnumerable().All(row => Convert.ToInt64(row["CEN_FUND"]) == 0);

                if (isMatch)
                {
                    sql = $@"
                            SELECT 
                                P.PID, P.PJ_CODE + P.FCATY_CODE AS CATY_CODE, P.CATY_NAME, P.ACC_ID, 
                                P.ACC_DESC, P.CATY_TYPE, P.CATY_DESC, 
                                ISNULL(
                                    (SELECT SUM(YR_FUND) 
                                     FROM TDV_TDDHFTS 
                                     WHERE PID = P.PID 
                                     AND PJ_CODE = P.PJ_CODE 
                                     AND PJ_CATY_CODE = P.CATY_CODE 
                                     AND SEQ_NO = '{sSEQ_NO}'
                                    ), 0
                                ) AS YR_FUND,
                                '{sSEQ_NO}' AS SEQ_NO, 
                                ISNULL(
                                (SELECT SUM(CEN_FUND) 
                                 FROM TDV_TDDHFTS
                                 WHERE PID = P.PID 
                                 AND PJ_CODE = P.PJ_CODE 
                                 AND PJ_CATY_CODE = P.CATY_CODE 
                                 AND SEQ_NO = '{sSEQ_NO}'
                                ), 0
                            ) AS CEN_FUND
                            FROM PJV_PJDCATY P
                            LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID 
                            WHERE P.PID IN 
                                (SELECT PID 
                                 FROM PJMFORM 
                                 WHERE SYR = '{PJ_SYR}' 
                                 AND PID = '{AN_PID}'
                                )
                            AND PJM.Water_Type = '雨水使用' ";
                }

                dt3 = new DataTable();
                dt3 = SourceFunction.DataTableGet(sql, connectionString);
            }

            /// 前一期資料
            if (sPSEQ_NO != "")
            {
                if (plan is null)
                {
                    sql = $@"
                            SELECT 
                                P.PID, P.PJ_CODE + P.FCATY_CODE AS CATY_CODE, P.CATY_NAME, P.ACC_ID, 
                                P.ACC_DESC, P.CATY_TYPE, P.CATY_DESC, 
                                ISNULL(
                                    (SELECT SUM(YR_FUND) 
                                     FROM TDV_TDDHFTS 
                                     WHERE PID = P.PID 
                                     AND PJ_CODE = P.PJ_CODE 
                                     AND PJ_CATY_CODE = P.CATY_CODE 
                                     AND SEQ_NO = '{sPSEQ_NO}'
                                    ), 0
                                ) AS YR_FUND
                            FROM PJV_PJDCATY P
                            LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID 
                            WHERE P.PID IN 
                                (SELECT PID 
                                 FROM PJMFORM 
                                 WHERE SYR = '{PJ_SYR}' 
                                 AND PID = '{AN_PID}'
                                )
                            AND PJM.Water_Type = '雨水使用' ";

                }
                else
                {
                    sql = $@"
                            SELECT 
                                P.PID, P.PJ_CODE + P.FCATY_CODE AS CATY_CODE, P.CATY_NAME, P.ACC_ID, 
                                P.ACC_DESC, P.CATY_TYPE, P.CATY_DESC, 
                                ISNULL(
                                    (SELECT SUM(YR_FUND) 
                                     FROM TDV_TDDHFTS 
                                     WHERE PID = P.PID 
                                     AND PJ_CODE = P.PJ_CODE 
                                     AND PJ_CATY_CODE = P.CATY_CODE 
                                     AND SEQ_NO = '{sPSEQ_NO}'
                                    ), 0
                                ) AS YR_FUND,
                                '{sPSEQ_NO}' AS SEQ_NO, 
                                ISNULL(
                                    (SELECT SUM(CEN_FUND) 
                                     FROM TDV_TDDHFTS 
                                     WHERE PID = P.PID 
                                     AND PJ_CODE = P.PJ_CODE 
                                     AND PJ_CATY_CODE = P.CATY_CODE 
                                     AND SEQ_NO = '{sPSEQ_NO}'
                                    ), 0
                                ) AS CEN_FUND
                            FROM PJV_PJDCATY P
                            LEFT JOIN PJMFORM PJM ON P.PID = PJM.PID 
                            WHERE P.PID IN 
                                (SELECT PID 
                                 FROM PJMFORM 
                                 WHERE SYR = '{PJ_SYR}' 
                                 AND PID = '{AN_PID}'
                                )
                            AND PJM.Water_Type = '雨水使用' ";
                }

                dt4 = new DataTable();
                dt4 = SourceFunction.DataTableGet(sql, connectionString);
            }

            #endregion -------------

            dt = new DataTable();
            dt.Columns.Add("PID");
            dt.Columns.Add("PID_SYR");
            dt.Columns.Add("SEQ_NO");
            dt.Columns.Add("TID");
            dt.Columns.Add("FD_CODE");
            dt.Columns.Add("TD_CODE");
            dt.Columns.Add("TD_NAME");
            dt.Columns.Add("SIP_TD_CODE");
            dt.Columns.Add("CITY");
            dt.Columns.Add("DATACITY");
            dt.Columns.Add("TOWN");
            dt.Columns.Add("ADM_DIST");
            dt.Columns.Add("IPJ_CODE");
            dt.Columns.Add("IPJ_CODE_DESC");
            dt.Columns.Add("N_IPJ_CODE");
            dt.Columns.Add("IPJ_DESC");
            dt.Columns.Add("CATY_TYPE");
            dt.Columns.Add("YR_FUND");
            dt.Columns.Add("CEN_FUND");
            dt.Columns.Add("LOC_FUND");
            dt.Columns.Add("PYR_FUND");      /// 前一期
            dt.Columns.Add("PCEN_FUND");
            dt.Columns.Add("PLOC_FUND");
            dt.Columns.Add("FROM");

            dtp.TableName = "dtp";
            ditdt.TableName = "ditdt";

            ds.Tables.Add(dtp);
            ds.Tables.Add(ditdt);

            DataTable dt1Copy = dt1.Copy();
            dt1Copy.TableName = "dt1";
            ds.Tables.Add(dt1Copy);

            DataTable dt2Copy = dt2.Copy();
            dt2Copy.TableName = "dt2";
            ds.Tables.Add(dt2Copy);

            DataTable dt3Copy = dt3.Copy();
            dt3Copy.TableName = "dt3";
            ds.Tables.Add(dt3Copy);

            DataTable dt4Copy = dt4.Copy();
            dt4Copy.TableName = "dt4";
            ds.Tables.Add(dt4Copy);

            return ds;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="PJ_SYR"></param>
        /// <param name="PJ_SMN"></param>
        /// <param name="connectionString"></param>
        /// <param name="DEPT"></param>
        /// <returns></returns>
        public static DataSet I1REPORTDATA(decimal PID, string PJ_SYR, string PJ_SMN, string connectionString, string DEPT)
        {
            string SEQ_Search = "";
            string DT1_Search = "";
            string DT2_Search = "";
            string DT3_Search = "";
            string DT4_Search = "";
            string SEQ_NO = "";
            string PSEQ_NO = "";

            DataSet DataSet_DS = new DataSet();
            DataTable SEQTable = new DataTable();
            DataTable DT1Table = new DataTable();
            DataTable DT2Table = new DataTable();
            DataTable DT3Table = new DataTable();
            DataTable DT4Table = new DataTable();

            /// Step 0.查詢計畫
            var Plan = db2.PJMFORM.Where(m => m.PID == PID).FirstOrDefault();

            /// Step 1.查詢最新一期(SEQ_NO)與前一期(PSEQ_NO)
            SEQ_Search = $@"SELECT MAX(SEQ_NO) AS MAX_SEQ_NO, MAX(SEQ_NO)-1 AS PMAX_SEQ_NO,min(SEQ_NO) AS MIN_SEQ_NO FROM CMEPMS.dbo.PJDFDCR WHERE PID IN (SELECT PID FROM CMEPMS.dbo.PJMFORM WHERE PID IN ({PID}) AND SYR = {PJ_SYR} )";

            SEQTable = SourceFunction.DataTableGet(SEQ_Search, connectionString);

            Trace.WriteLine($"SEQ_Search = {SEQ_Search}");

            if (SEQTable != null && SEQTable.Rows.Count > 0)
            {
                if (SEQTable.Rows[0]["MAX_SEQ_NO"].ToString() != "")
                {
                    SEQ_NO = SEQTable.Rows[0]["MAX_SEQ_NO"].ToString();
                }

                if (SEQTable.Rows[0]["PMAX_SEQ_NO"].ToString() != "")
                {
                    PSEQ_NO = SEQTable.Rows[0]["PMAX_SEQ_NO"].ToString();
                }
            }

            #region ================================= 總表|署辦|地辦|業務費 =================================

            string BaseSQL = $@"
                                SELECT
                                    A.PID,
                                    PJM.SYR AS PID_SYR,
                                    A.TID,
                                    ISNULL(B.SEQ_NO, 0) AS SEQ_NO,
                                    A.TD_CODE,
                                    A.FD_CODE,
                                    A.TD_NAME,
                                    STUFF(STUFF(STUFF(STUFF(CAST(A.TD_CODE AS varchar(20)), 4, 0, '-'), 9, 0, '-'), 14, 0, '-'), 19, 0, '-') AS SIP_TD_CODE,
                                    C.TD_TOWN AS TOWN,
                                    C.TD_CITY AS CITY,
                                    ISNULL(FUND_SUM.FUND_TOTAL, 0) AS CEN_FUND,
                                    PT.ACC_ID,
                                    CASE C.IPJ_CODE
                                        WHEN '01' THEN '署辦'
                                        WHEN '02' THEN '地辦'
                                        WHEN '03' THEN '地辦'
                                        ELSE '署辦'
                                    END AS IPJ_CODE_DESC,
                                    SM.CODE_DESC AS IPJ_DESC
                                FROM CMEPMS.dbo.CTDMNDS A
                                LEFT JOIN CMEPMS.dbo.PJDFDCR B ON B.PID = A.PID
                                LEFT JOIN CMETMS.dbo.TDMFORM C ON C.TD_CODE = A.TD_CODE
                                LEFT JOIN CMEPMS.dbo.PJMFORM PJM ON PJM.PID = A.PID
                                LEFT JOIN CMEPMS.dbo.PJDCATY PT ON PT.PID = A.PID AND PT.CATY_CODE = A.CATY_CODE
                                LEFT JOIN CMETMS.dbo.SMCCODE SM ON SM.SYS_ID = 'SM' AND SM.CATY_ID = '02' AND SM.CATY_COD = C.IPJ_CODE
                                LEFT JOIN ( SELECT TID, SUM(FUND) AS FUND_TOTAL FROM CMEPMS.dbo.CTDITFD WHERE FIT_CODE = '00' GROUP BY TID ) FUND_SUM ON FUND_SUM.TID = A.TID
                                WHERE A.PID IN ({PID})";

            /// 假設 MAX_SEQ_NO | PMAX_SEQ_NO 皆不為空值
            if (!string.IsNullOrEmpty(SEQ_NO) && !string.IsNullOrEmpty(PSEQ_NO))
            {
                DT1_Search = BaseSQL + $@" AND B.SEQ_NO IN ('{SEQ_NO}')";
                DT2_Search = BaseSQL + $@" AND B.SEQ_NO IN ('{PSEQ_NO}')";

                DT1Table = SourceFunction.DataTableGet(DT1_Search, connectionString);
                DT2Table = SourceFunction.DataTableGet(DT2_Search, connectionString);
            }
            else
            {
                /// DT1_Search = BaseSQL + $@" AND B.SEQ_NO IN ('{SEQ_NO}')";
                DT1_Search = BaseSQL;
                DT1Table = SourceFunction.DataTableGet(DT1_Search, connectionString);

                if (DT1Table == null)
                {
                    DT1_Search = BaseSQL;
                    DT1Table = SourceFunction.DataTableGet(DT1_Search, connectionString);
                }
            }

            #endregion ===================================================================================

            #region ================================= 經費調整表|調整表_明細 ================================

            string FundSQL = $@"
                            SELECT
                                A.CATY_CODE,
                                ISNULL(PT.CATY_NAME, '') AS CATY_NAME,
                                CASE 
                                    WHEN PT.ACC_ID = '0300' THEN N'資本門'
                                    ELSE N'經常門'
                                END AS CATY_DESC,
                                ISNULL(PT.ACC_ID, '') AS ACC_ID,
                                ISNULL(AD.CODE_DESC, '') AS ACC_DESC,
                                SUM(ISNULL(FUND_SUM.FUND_TOTAL, 0)) AS CEN_FUND
                            FROM CMEPMS.dbo.CTDMNDS A
                            LEFT JOIN CMEPMS.dbo.PJDCATY PT ON PT.PID = A.PID AND PT.CATY_CODE = A.CATY_CODE
                            LEFT JOIN CMEPMS.dbo.SMCCODE AD ON AD.SYS_ID = 'FD' AND AD.CATY_ID = '02' AND AD.CATY_COD = PT.ACC_ID
                            LEFT JOIN ( SELECT TID,SUM(ISNULL(FUND, 0)) AS FUND_TOTAL FROM CMEPMS.dbo.CTDITFD WHERE LEFT(FIT_CODE, 1) = '0' GROUP BY TID) FUND_SUM ON FUND_SUM.TID = A.TID
                            WHERE A.PID IN ({PID}) ";

            if (!string.IsNullOrEmpty(SEQ_NO))
            {
                DT3_Search = FundSQL + $@" AND EXISTS ( SELECT 1 FROM CMEPMS.dbo.PJDFDCR B WHERE B.PID = A.PID AND B.SEQ_NO = {SEQ_NO} ) GROUP BY
                                A.CATY_CODE,
                                ISNULL(PT.CATY_NAME, ''),
                                CASE 
                                    WHEN PT.ACC_ID = '0300' THEN N'資本門'
                                    ELSE N'經常門'
                                END,
                                ISNULL(PT.ACC_ID, ''),
                                ISNULL(AD.CODE_DESC, '') ";
            }
            else
            {
                /// 沒有當期核定期數時，退回只查 PID 基礎資料
                DT3_Search = FundSQL;
            }

            DT3Table = SourceFunction.DataTableGet(DT3_Search, connectionString);

            if (!string.IsNullOrEmpty(PSEQ_NO))
            {
                DT4_Search = FundSQL + $@" AND EXISTS ( SELECT 1 FROM CMEPMS.dbo.PJDFDCR B WHERE B.PID = A.PID AND B.SEQ_NO = {PSEQ_NO} ) GROUP BY
                                A.CATY_CODE,
                                ISNULL(PT.CATY_NAME, ''),
                                CASE 
                                    WHEN PT.ACC_ID = '0300' THEN N'資本門'
                                    ELSE N'經常門'
                                END,
                                ISNULL(PT.ACC_ID, ''),
                                ISNULL(AD.CODE_DESC, '') ";

                DT4Table = SourceFunction.DataTableGet(DT4_Search, connectionString);
            }
            else
            {
                /// 沒有前一期時，給空表
                DT4_Search = "";
                DT4Table = new DataTable();
            }

            #endregion ===================================================================================

            Trace.WriteLine($"經費調整表|調整表_明細 SQL DT3 = {DT3_Search}");
            Trace.WriteLine($"經費調整表|調整表_明細 SQL DT4 = {DT4_Search}");

            DataTable CopyTable1 = DT1Table.Copy();
            CopyTable1.TableName = "dt1";
            DataSet_DS.Tables.Add(CopyTable1);

            DataTable CopyTable2 = DT2Table.Copy();
            CopyTable2.TableName = "dt2";
            DataSet_DS.Tables.Add(CopyTable2);

            DataTable CopyTable3 = DT3Table.Copy();
            CopyTable3.TableName = "dt3";
            DataSet_DS.Tables.Add(CopyTable3);

            DataTable CopyTable4 = DT4Table.Copy();
            CopyTable4.TableName = "dt4";
            DataSet_DS.Tables.Add(CopyTable4);

            return DataSet_DS;
        }

        /// <summary>
        /// I-1.雨水例行性管考報表 => (經費調整表)總表
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet8"></param>
        /// <param name="dt1"></param>
        /// <param name="dt2"></param>
        /// <param name="dt3"></param>
        /// <param name="ditdt"></param>
        /// <param name="SYR"></param>
        /// <param name="PSEQ_NO"></param>
        /// <param name="SEQ_NO"></param>
        /// <param name="PJ_NAME"></param>
        /// <see cref="New_GetRPT0001Data(decimal, string, string, string, string)"/> 資料來源
        public static void RoutineRainwaterInspectionReport(Workbook workbook, Worksheet sheet8, DataTable dt1, DataTable dt2, DataTable ditdt, string SYR, string PSEQ_NO, string SEQ_NO, string PJ_NAME)
        {
            Worksheet Sheet8work = workbook.Worksheets["Sheet8"];

            int ItemNum = 0;   // 項次
            int StartCow1 = 4; // 從第5列插入資料
            int maxRow = sheet8.Cells.MaxDataRow; // 宣告欄位最大值
            double TotalyrFund = 0;  // 總計(F)
            double TotalyrFund2 = 0; // 總計(G)

            /// 總表標題列
            sheet8.Cells.Merge(0, 0, 1, 12);
            sheet8.Cells[0, 0].PutValue($"{SYR}年度 {PJ_NAME} 總表");
            Aspose.Cells.Style SheetTitle = sheet8.Cells[0, 0].GetStyle();
            SheetTitle.HorizontalAlignment = TextAlignmentType.Center;
            sheet8.Cells[0, 0].SetStyle(SheetTitle);

            // 原預算(前一期)
            if (dt1.Rows.Count > 0 && dt2.Rows.Count == 0)
            {
                // 核定"期數" / 計畫年度
                int DT1_SEQNO = Convert.ToInt32(dt1.Rows[0]["SEQ_NO"]);
                string SEQ_Text1 = DT1_SEQNO == 0 ? "原預算" : $"第{DT1_SEQNO}次核定預算";
                sheet8.Cells["F3"].PutValue(SEQ_Text1);
                sheet8.Cells["F4"].PutValue($"{SYR}");
                sheet8.Cells["G3"].PutValue("暫無預算");
                sheet8.Cells["G4"].PutValue($"{SYR}");

                // 資料填入
                foreach (DataRow row1 in dt1.Rows)
                {
                    sheet8.Cells.CopyRow(sheet8.Cells, 4, StartCow1);

                    sheet8.Cells[StartCow1, 0].PutValue(++ItemNum);
                    sheet8.Cells[StartCow1, 1].PutValue(row1["CITY"].ToString());
                    sheet8.Cells[StartCow1, 2].PutValue(row1["SIP_TD_CODE"].ToString());
                    sheet8.Cells[StartCow1, 3].PutValue(row1["TOWN"].ToString());
                    sheet8.Cells[StartCow1, 4].PutValue(row1["TD_NAME"].ToString());
                    sheet8.Cells[StartCow1, 5].PutValue(Convert.ToDouble(row1["CEN_FUND"]));// 中央款
                    sheet8.Cells[StartCow1, 6].PutValue(0);
                    sheet8.Cells[StartCow1, 9].PutValue(row1["IPJ_CODE_DESC"].ToString());  // 署辦

                    if (double.TryParse(row1["CEN_FUND"].ToString(), out double yrFund))
                    {
                        TotalyrFund += yrFund;
                    }

                    StartCow1++;
                }
            }
            else if (dt1.Rows.Count > 0 && dt2.Rows.Count > 0)
            {
                // 核定"期數" / 計畫年度
                int DT1_SEQNO = Convert.ToInt32(dt1.Rows[0]["SEQ_NO"]);
                int DT2_SEQNO = Convert.ToInt32(dt2.Rows[0]["SEQ_NO"]);
                string SEQ_Text1 = DT1_SEQNO == 0 ? "原預算" : $"第{DT1_SEQNO}次核定預算";
                string SEQ_Text2 = DT2_SEQNO == 0 ? "原預算" : $"第{DT2_SEQNO}次核定預算";
                sheet8.Cells["G3"].PutValue(SEQ_Text1);
                sheet8.Cells["G4"].PutValue($"{SYR}");
                sheet8.Cells["F3"].PutValue(SEQ_Text2);
                sheet8.Cells["F4"].PutValue($"{SYR}");

                // 資料填入
                foreach (DataRow row2 in dt2.Rows)
                {
                    // 比對 DT2(前一期) 和 DT1(本期中) 的 TID
                    string TID = row2["TID"].ToString();
                    DataRow[] matchingRows = dt1.Select($"TID = '{TID}'");

                    sheet8.Cells.CopyRow(sheet8.Cells, 4, StartCow1); // 複製格式 Style
                    sheet8.Cells[StartCow1, 0].PutValue(++ItemNum);                          // 項次
                    sheet8.Cells[StartCow1, 1].PutValue(row2["CITY"].ToString());            // 縣市別
                    sheet8.Cells[StartCow1, 2].PutValue(row2["SIP_TD_CODE"].ToString());     // 工程編號
                    sheet8.Cells[StartCow1, 3].PutValue(row2["TOWN"].ToString());            // 鄉鎮別
                    sheet8.Cells[StartCow1, 4].PutValue(row2["TD_NAME"].ToString());         // 工程名稱
                    sheet8.Cells[StartCow1, 5].PutValue(Convert.ToDouble(row2["CEN_FUND"])); // 中央款
                    sheet8.Cells[StartCow1, 9].PutValue(row2["IPJ_CODE_DESC"].ToString());   // 署辦

                    if (double.TryParse(row2["CEN_FUND"].ToString(), out double yrFund1))
                    {
                        TotalyrFund += yrFund1;
                    }

                    if (matchingRows.Length > 0)
                    {
                        // 更新金額，填入新一期的 CEN_FUND 金額
                        double CenFUND = Convert.ToDouble(matchingRows[0]["CEN_FUND"]);      // 更新中央款
                        sheet8.Cells[StartCow1, 6].PutValue(CenFUND);
                        TotalyrFund2 += CenFUND;
                    }
                    StartCow1++; // 換下一筆資料
                }

                // 比對 DT2(前一期) 和 DT1(本期中)
                foreach (DataRow row1 in dt1.Rows)
                {
                    string TID = row1["TID"].ToString();
                    DataRow[] matchingRowsInDt2 = dt2.Select($"TID = '{TID}'");

                    if (matchingRowsInDt2.Length == 0 && row1["PID_SYR"].ToString() == SYR)
                    {
                        // TID 在 DT2(前一期) 中找不到，新增至 Excel
                        sheet8.Cells.CopyRow(sheet8.Cells, 4, StartCow1);
                        sheet8.Cells[StartCow1, 0].PutValue(++ItemNum);                         // 項次
                        sheet8.Cells[StartCow1, 1].PutValue(row1["CITY"].ToString());           // 縣市別
                        sheet8.Cells[StartCow1, 2].PutValue(row1["SIP_TD_CODE"].ToString());    // 工程編號
                        sheet8.Cells[StartCow1, 3].PutValue(row1["TOWN"].ToString());           // 鄉鎮別
                        sheet8.Cells[StartCow1, 4].PutValue(row1["TD_NAME"].ToString());        // 工程名稱
                        sheet8.Cells[StartCow1, 5].PutValue(0);                                 // 新增的前一期無預算
                        double newCenFund = Convert.ToDouble(row1["CEN_FUND"]);                 // 新一期(中央款)
                        sheet8.Cells[StartCow1, 6].PutValue(newCenFund);
                        sheet8.Cells[StartCow1, 9].PutValue(row1["IPJ_CODE_DESC"].ToString());  // 署辦

                        TotalyrFund2 += newCenFund;

                        StartCow1++; // 換下一筆資料
                    }
                }
            }

            // 原預算/期數(總計)
            sheet8.Cells.CopyRow(sheet8.Cells, 4, StartCow1);
            sheet8.Cells[StartCow1, 0].PutValue("");
            sheet8.Cells[StartCow1, 1].PutValue("");
            sheet8.Cells[StartCow1, 2].PutValue("");
            sheet8.Cells[StartCow1, 3].PutValue("");
            sheet8.Cells[StartCow1, 4].PutValue("總計");
            sheet8.Cells[StartCow1, 5].PutValue(TotalyrFund);
            sheet8.Cells[StartCow1, 6].PutValue(TotalyrFund2 != 0 ? TotalyrFund2 : 0);
            sheet8.Cells[StartCow1, 9].PutValue("");

            // 增減欄位
            // 從 StartCow1 開始逐行處理
            for (int i = StartCow1; i <= maxRow; i++)
            {
                var valueF = sheet8.Cells[i, 5].DoubleValue; // F列 (索引5)
                var valueG = sheet8.Cells[i, 6].DoubleValue; // G列 (索引6)
                double diff = valueG - valueF;               // 計算增減值 (G - F)
                sheet8.Cells[i, 7].PutValue(diff);           // 將結果放入 H 欄 (索引7)
            }
        }

        /// <summary>
        /// I-1.雨水例行性管考報表 => (經費調整表)經費調整表
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet9"></param>
        /// <param name="dt3"></param>
        /// <param name="dt4"></param>
        /// <param name="SYR"></param>
        /// <param name="PSEQ_NO"></param>
        /// <param name="SEQ_NO"></param>
        /// <param name="PJ_NAME"></param>
        /// /// <see cref="New_GetRPT0001Data(decimal, string, string, string, string)"/> 資料來源
        public static void FundingAdjustmentTable(Workbook workbook, Worksheet sheet9, DataTable dt3, DataTable dt4, string SYR, string PSEQ_NO, string SEQ_NO, string PJ_NAME)
        {
            Worksheet Sheet9work = workbook.Worksheets["Sheet9"];
            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");

            int maxRow = sheet9.Cells.MaxDataRow + 1; /// 宣告欄位最大值

            double deviceYRFUND = 0;        /// 設備及投資
            double businessYRFUND = 0;      /// 業務費
            double RedeviceYRFUND = 0;      /// 設備及投資
            double RebusinessYRFUND = 0;    /// 業務費
            double subsidyYRFUND = 0;
            double REsubsidyYRFUND = 0;

            /// 總表標題列
            sheet9.Cells.Merge(0, 0, 1, 7);
            sheet9.Cells[0, 0].PutValue($"{SYR}年度 {PJ_NAME} 經費調整表");
            Aspose.Cells.Style SheetTitle = sheet9.Cells[0, 0].GetStyle();
            SheetTitle.HorizontalAlignment = TextAlignmentType.Center;
            sheet9.Cells[0, 0].SetStyle(SheetTitle);
            sheet9.Cells["C3"].PutValue(SYR.ToString());

            /// 新增列
            sheet9.Cells.InsertRows(maxRow, 3);
            sheet9.Cells["A5"].PutValue(1);
            sheet9.Cells["B5"].PutValue("設備及投資(署辦工程)");
            sheet9.Cells["A6"].PutValue(2);
            sheet9.Cells["B6"].PutValue("獎補助費(地方辦理)");
            sheet9.Cells["A7"].PutValue(3);
            sheet9.Cells["B7"].PutValue("業務費");
            sheet9.Cells["A8"].PutValue("");
            sheet9.Cells["B8"].PutValue("總計");

            bool hasAccDescDt3 = dt3.Columns.Contains("ACC_DESC");
            bool hasAccDescDt4 = dt4.Columns.Contains("ACC_DESC");

            if (dt3.Rows.Count > 0 && dt4.Rows.Count == 0)
            {
                /// 資金調整標題
                sheet9.Cells["C4"].PutValue("調整前");
                sheet9.Cells["D4"].PutValue("暫無調整");

                foreach (DataRow row1 in dt3.Rows)
                {
                    double cenFund = row1["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(row1["CEN_FUND"]);

                    if (row1["ACC_ID"].ToString() == "0300") /// 設備及投資費用總計
                    {
                        deviceYRFUND += cenFund;
                    }

                    if (row1["ACC_ID"].ToString() == "0200") /// 業務費總計
                    {
                        businessYRFUND += cenFund;
                    }

                    if (hasAccDescDt3 && row1["ACC_DESC"].ToString() == "獎補助費") /// 獎補助費總計
                    {
                        subsidyYRFUND += cenFund;
                    }
                }

                /// 增減欄位
                for (int count = 4; count < 8; count++)
                {
                    object cellValue2 = sheet9.Cells[count, 2].Value;
                    object cellValue3 = sheet9.Cells[count, 3].Value;

                    double value2 = (cellValue2 is double) ? (double)cellValue2 : 0;
                    double value3 = (cellValue3 is double) ? (double)cellValue3 : 0;

                    sheet9.Cells[count, 4].PutValue(value3 - value2);
                }

                sheet9.Cells["C5"].PutValue(deviceYRFUND);
                sheet9.Cells["D5"].PutValue(0);
                sheet9.Cells["C6"].PutValue(subsidyYRFUND);
                sheet9.Cells["D6"].PutValue(0);
                sheet9.Cells["C7"].PutValue(businessYRFUND);
                sheet9.Cells["D7"].PutValue(0);
                sheet9.Cells["C8"].PutValue(deviceYRFUND + businessYRFUND + subsidyYRFUND);
                sheet9.Cells["D8"].PutValue(0);
            }
            else
            {
                /// 資金調整標題
                sheet9.Cells["C4"].PutValue("調整前");
                sheet9.Cells["D4"].PutValue("調整後");

                foreach (DataRow row1 in dt3.Rows)
                {
                    double cenFund = row1["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(row1["CEN_FUND"]);

                    if (row1["ACC_ID"].ToString() == "0300")
                    {
                        deviceYRFUND += cenFund;
                    }

                    if (row1["ACC_ID"].ToString() == "0200")
                    {
                        businessYRFUND += cenFund;
                    }

                    if (hasAccDescDt3 && row1["ACC_DESC"].ToString() == "獎補助費")
                    {
                        subsidyYRFUND += cenFund;
                    }
                }

                foreach (DataRow row2 in dt4.Rows)
                {
                    double cenFund = row2["CEN_FUND"] == DBNull.Value ? 0 : Convert.ToDouble(row2["CEN_FUND"]);

                    if (row2["ACC_ID"].ToString() == "0300")
                    {
                        RedeviceYRFUND += cenFund;
                    }

                    if (row2["ACC_ID"].ToString() == "0200")
                    {
                        RebusinessYRFUND += cenFund;
                    }

                    if (hasAccDescDt4 && row2["ACC_DESC"].ToString() == "獎補助費")
                    {
                        REsubsidyYRFUND += cenFund;
                    }
                }

                /// 增減欄位
                for (int count = 4; count < 8; count++)
                {
                    object adjustValue1 = sheet9.Cells[count, 2].Value; // C欄數值(調整前)
                    object adjustValue2 = sheet9.Cells[count, 3].Value; // D欄數值(調整後)

                    double PreAdjustment = (adjustValue1 is double) ? (double)adjustValue1 : 0;  // C欄數值(調整前)
                    double PostAdjustment = (adjustValue2 is double) ? (double)adjustValue2 : 0; // D欄數值(調整後)

                    sheet9.Cells[count, 4].PutValue(PostAdjustment - PreAdjustment);
                }

                double subtractionDevice = deviceYRFUND - RedeviceYRFUND;
                double subtractionBusiness = businessYRFUND - RebusinessYRFUND;
                double subtractionsubsidy = subsidyYRFUND - REsubsidyYRFUND;

                double total1 = RedeviceYRFUND + RebusinessYRFUND + REsubsidyYRFUND;
                double total2 = deviceYRFUND + businessYRFUND + subsidyYRFUND;

                sheet9.Cells["C5"].PutValue(RedeviceYRFUND);
                sheet9.Cells["C6"].PutValue(REsubsidyYRFUND);
                sheet9.Cells["C7"].PutValue(RebusinessYRFUND);
                sheet9.Cells["C8"].PutValue(RedeviceYRFUND + RebusinessYRFUND + REsubsidyYRFUND);

                sheet9.Cells["D5"].PutValue(deviceYRFUND);
                sheet9.Cells["D6"].PutValue(subsidyYRFUND);
                sheet9.Cells["D7"].PutValue(businessYRFUND);
                sheet9.Cells["D8"].PutValue(deviceYRFUND + businessYRFUND + subsidyYRFUND);

                sheet9.Cells["E5"].PutValue(subtractionDevice);
                sheet9.Cells["E6"].PutValue(subtractionsubsidy);
                sheet9.Cells["E7"].PutValue(subtractionBusiness);
                sheet9.Cells["E8"].PutValue(total2 - total1);
            }
        }

        /// <summary>
        /// I-1.雨水例行性管考報表 => (經費調整表)經費調整表_明細
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet10"></param>
        /// <param name="dt3"></param>
        /// <param name="dt4"></param>
        /// <param name="SYR"></param>
        /// <param name="PSEQ_NO"></param>
        /// <param name="SEQ_NO"></param>
        /// <param name="PJ_NAME"></param>
        /// <see cref="New_GetRPT0001Data(decimal, string, string, string, string)"/> 資料來源
        public static void FundingAdjustmentTableDetails(Workbook workbook, Worksheet sheet10, DataTable dt3, DataTable dt4, string SYR, string PSEQ_NO, string SEQ_NO, string PJ_NAME)
        {
            Worksheet Sheet10work = workbook.Worksheets["Sheet10"];

            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");

            Debug.WriteLine("DT3");
            foreach (DataRow rows in dt3.Rows)
            {
                foreach (DataColumn cols in dt3.Columns)
                {
                    Debug.Write($"{cols.ColumnName}:{rows[cols]} \t");
                }
                Debug.WriteLine("");
            }

            Debug.WriteLine("DT4");
            foreach (DataRow rows in dt4.Rows)
            {
                foreach (DataColumn cols in dt4.Columns)
                {
                    Debug.Write($"{cols.ColumnName}:{rows[cols]} \t");
                }
                Debug.WriteLine("");
            }

            int StartRow = 4; // 起始行（A5）

            double capitalTotal = 0;     // 資本門金額(調整前)
            double regularTotal = 0;     // 經常門金額(調整前)
            double businessTotal = 0;    // 業務費金額(調整前)
            double equipmentTotal = 0;   // 設備及投資金額(調整前)
            double subsidyTotal = 0;     // 獎補助費(調整前)
            double Total = 0;            // 總計(調整前)

            double afcapitalTotal = 0;   // 資本門金額(調整後)
            double afregularTotal = 0;   // 經常門金額(調整後)
            double afbusinessTotal = 0;  // 業務費金額(調整後)
            double afequipmentTotal = 0; // 設備及投資金額(調整後)
            double afsubsidyTotal = 0;   // 獎補助費(調整後)
            double afTotal = 0;          // 總計(調整後)

            /// 標題列
            sheet10.Cells.Merge(0, 0, 1, 6);
            sheet10.Cells["A1"].PutValue($"{SYR}年度 {PJ_NAME} 經費調整表明細");
            Aspose.Cells.Style SheetTitle = sheet10.Cells[0, 0].GetStyle();
            SheetTitle.HorizontalAlignment = TextAlignmentType.Center;
            sheet10.Cells[0, 0].SetStyle(SheetTitle);
            sheet10.Cells["C3"].PutValue(SYR.ToString());
            sheet10.Cells["C4"].PutValue("調整前");
            sheet10.Cells["D4"].PutValue("調整後");

            if (dt3.Rows.Count > 0 && dt4.Rows.Count == 0)
            {
                foreach (DataRow rows in dt3.Rows)
                {
                    sheet10.Cells.CopyRow(sheet10.Cells, 4, StartRow);
                    sheet10.Cells[StartRow, 0].PutValue(rows["CATY_CODE"].ToString());  // 項次
                    sheet10.Cells[StartRow, 1].PutValue(rows["CATY_NAME"].ToString());  // 項目別
                    double beforeAdjust = Convert.ToDouble(rows["CEN_FUND"]);
                    sheet10.Cells[StartRow, 2].PutValue(beforeAdjust);                  // 調整前金額
                    double afterAdjust = 0;
                    sheet10.Cells[StartRow, 3].PutValue(afterAdjust);                   // 調整後金額

                    double difference = afterAdjust - beforeAdjust;                     // 增減欄位
                    sheet10.Cells[StartRow, 4].PutValue(difference);
                    if (difference < 0)
                    {
                        Aspose.Cells.Style style = sheet10.Cells[StartRow, 4].GetStyle();
                        style.Font.Color = System.Drawing.Color.Red;
                        sheet10.Cells[StartRow, 4].SetStyle(style);
                    }

                    StartRow++;
                }

                sheet10.Cells.CopyRow(sheet10.Cells, 4, StartRow);
                sheet10.Cells.InsertRows(StartRow, 5);

                int nextCount = StartRow;
                int finalCount = nextCount + 6;

                foreach (DataRow rows in dt3.Rows)
                {
                    for (int i = nextCount; i < finalCount; i++)
                    {
                        sheet10.Cells.Merge(i, 0, 1, 2); // 獨立合併 (A 和 B)

                        if (i == finalCount - 6)
                        {
                            sheet10.Cells[i, 0].PutValue("資本門");

                            if (rows["CATY_DESC"].ToString() == "資本門")
                            {
                                capitalTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(capitalTotal);
                            sheet10.Cells[i, 3].PutValue(Convert.ToDouble(0));
                            sheet10.Cells[i, 4].PutValue(0 - capitalTotal);
                        }
                        else if (i == finalCount - 5)
                        {
                            sheet10.Cells[i, 0].PutValue("業務費");

                            if (rows["ACC_DESC"].ToString() == "業務費")
                            {
                                businessTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(businessTotal);
                            sheet10.Cells[i, 3].PutValue(Convert.ToDouble(0));
                            sheet10.Cells[i, 4].PutValue(0 - businessTotal);
                        }
                        else if (i == finalCount - 4)
                        {
                            sheet10.Cells[i, 0].PutValue("經常門");

                            if (rows["CATY_DESC"].ToString() == "經常門")
                            {
                                regularTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(regularTotal);
                            sheet10.Cells[i, 3].PutValue(Convert.ToDouble(0));
                            sheet10.Cells[i, 4].PutValue(0 - regularTotal);
                        }
                        else if (i == finalCount - 3)
                        {
                            sheet10.Cells[i, 0].PutValue("設備及投資(署辦工程)");

                            if (rows["ACC_DESC"].ToString() == "設備及投資")
                            {
                                equipmentTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(equipmentTotal);
                            sheet10.Cells[i, 3].PutValue(Convert.ToDouble(0));
                            sheet10.Cells[i, 4].PutValue(0 - equipmentTotal);
                        }
                        else if (i == finalCount - 2)
                        {
                            sheet10.Cells[i, 0].PutValue("獎補助費(地方辦理)");

                            if (rows["ACC_DESC"].ToString() == "獎補助費")
                            {
                                subsidyTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(subsidyTotal);
                            sheet10.Cells[i, 3].PutValue(Convert.ToDouble(0));
                            sheet10.Cells[i, 4].PutValue(0 - subsidyTotal);
                        }
                        else if (i == finalCount - 1)
                        {
                            Total = capitalTotal + regularTotal;
                            sheet10.Cells[i, 0].PutValue("總計");
                            sheet10.Cells[i, 2].PutValue(Total);
                            sheet10.Cells[i, 3].PutValue(Convert.ToDouble(0));
                            sheet10.Cells[i, 4].PutValue(0 - Total);
                        }
                    }
                }
            }
            else
            {
                foreach (DataRow rows in dt4.Rows)
                {
                    sheet10.Cells.CopyRow(sheet10.Cells, 4, StartRow);                  // 複製樣式和結構
                    sheet10.Cells[StartRow, 0].PutValue(rows["CATY_CODE"].ToString());  // 項次
                    sheet10.Cells[StartRow, 1].PutValue(rows["CATY_NAME"].ToString());  // 項目別
                    double beforeAdjust = Convert.ToDouble(rows["CEN_FUND"]);
                    sheet10.Cells[StartRow, 2].PutValue(beforeAdjust);                  // 調整前金額
                    double afterAdjust = 0;                                             // 預設調整後金額

                    string catyCode = rows["CATY_CODE"].ToString();                     // DT4 的 CATY_CODE
                    DataRow[] matchingRows = dt3.Select($"CATY_CODE = '{catyCode}'");   // 查詢 DT3 中匹配的 CATY_CODE 並比對

                    if (matchingRows.Length > 0)
                    {
                        afterAdjust = Convert.ToDouble(matchingRows[0]["CEN_FUND"]);     // 匹配 CATY_CODE 後更新 CEN_FUND
                    }

                    sheet10.Cells[StartRow, 3].PutValue(afterAdjust);                   // 更新 D 欄 (調整後金額)

                    /// 增減欄位計算
                    double difference = afterAdjust - beforeAdjust;
                    sheet10.Cells[StartRow, 4].PutValue(difference);
                    if (difference < 0)
                    {
                        Aspose.Cells.Style style = sheet10.Cells[StartRow, 4].GetStyle();
                        style.Font.Color = System.Drawing.Color.Black;
                        sheet10.Cells[StartRow, 4].SetStyle(style);
                    }

                    StartRow++;
                }

                foreach (DataRow row3 in dt3.Rows)
                {
                    if (row3["CATY_DESC"].ToString() == "資本門")
                    {
                        afcapitalTotal += Convert.ToDouble(row3["CEN_FUND"]);
                    }

                    if (row3["CATY_DESC"].ToString() == "經常門")
                    {
                        afregularTotal += Convert.ToDouble(row3["CEN_FUND"]);
                    }

                    if (row3["ACC_DESC"].ToString() == "設備及投資")
                    {
                        afequipmentTotal += Convert.ToDouble(row3["CEN_FUND"]);
                    }

                    if (row3["ACC_DESC"].ToString() == "業務費")
                    {
                        afbusinessTotal += Convert.ToDouble(row3["CEN_FUND"]);
                    }

                    if (row3["ACC_DESC"].ToString() == "獎補助費")
                    {
                        afsubsidyTotal += Convert.ToDouble(row3["CEN_FUND"]);
                    }
                }

                int nextCount = StartRow;
                int finalCount = nextCount + 6;

                /// 經常門/資本門/設備及投資/業務費
                sheet10.Cells.CopyRow(sheet10.Cells, 4, StartRow);
                sheet10.Cells.InsertRows(StartRow, 5);

                foreach (DataRow rows in dt4.Rows)
                {
                    for (int i = nextCount; i < finalCount; i++)
                    {
                        sheet10.Cells.Merge(i, 0, 1, 2); // 獨立合併 (A 和 B)

                        if (i == finalCount - 6)
                        {
                            sheet10.Cells[i, 0].PutValue("資本門");

                            if (rows["CATY_DESC"].ToString() == "資本門")
                            {
                                capitalTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(capitalTotal);
                            sheet10.Cells[i, 3].PutValue(afcapitalTotal);
                            sheet10.Cells[i, 4].PutValue(afcapitalTotal - capitalTotal);
                        }
                        else if (i == finalCount - 5)
                        {
                            sheet10.Cells[i, 0].PutValue("業務費");

                            if (rows["ACC_DESC"].ToString() == "業務費")
                            {
                                businessTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(businessTotal);
                            sheet10.Cells[i, 3].PutValue(afbusinessTotal);
                            sheet10.Cells[i, 4].PutValue(afbusinessTotal - businessTotal);
                        }
                        else if (i == finalCount - 4)
                        {
                            sheet10.Cells[i, 0].PutValue("經常門");

                            if (rows["CATY_DESC"].ToString() == "經常門")
                            {
                                regularTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(regularTotal);
                            sheet10.Cells[i, 3].PutValue(afregularTotal);
                            sheet10.Cells[i, 4].PutValue(afregularTotal - regularTotal);
                        }
                        else if (i == finalCount - 3)
                        {
                            sheet10.Cells[i, 0].PutValue("設備及投資(署辦工程)");

                            if (rows["ACC_DESC"].ToString() == "設備及投資")
                            {
                                equipmentTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(equipmentTotal);
                            sheet10.Cells[i, 3].PutValue(afequipmentTotal);
                            sheet10.Cells[i, 4].PutValue(afequipmentTotal - equipmentTotal);
                        }
                        else if (i == finalCount - 2)
                        {
                            sheet10.Cells[i, 0].PutValue("獎補助費(地方辦理)");

                            if (rows["ACC_DESC"].ToString() == "獎補助費")
                            {
                                subsidyTotal += Convert.ToDouble(rows["CEN_FUND"]);
                            }

                            sheet10.Cells[i, 2].PutValue(subsidyTotal);
                            sheet10.Cells[i, 3].PutValue(afsubsidyTotal);
                            sheet10.Cells[i, 4].PutValue(afsubsidyTotal - subsidyTotal);
                        }
                        else if (i == finalCount - 1)
                        {
                            Total = capitalTotal + regularTotal;
                            afTotal = afcapitalTotal + afregularTotal;
                            sheet10.Cells[i, 0].PutValue("總計");
                            sheet10.Cells[i, 2].PutValue(Total);
                            sheet10.Cells[i, 3].PutValue(afTotal);
                            sheet10.Cells[i, 4].PutValue(afTotal - Total);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// I-1.雨水例行性管考報表 => (經費調整表)署辦
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet11"></param>
        /// <param name="dt1"></param>
        /// <param name="dt2"></param>
        /// <param name="ditdt"></param>
        /// <param name="SYR"></param>
        /// <param name="PSEQ_NO"></param>
        /// <param name="SEQ_NO"></param>
        /// <param name="PJ_NAME"></param>
        /// <see cref="New_GetRPT0001Data(decimal, string, string, string, string)"/> 資料來源
        public static void DepartmentOfficeTable(Workbook workbook, Worksheet sheet11, DataTable dt1, DataTable dt2, string SYR, string PSEQ_NO, string SEQ_NO, string PJ_NAME)
        {
            Worksheet Sheet11work = workbook.Worksheets["Sheet11"];

            int itemCount = 0;                     // 項次
            int startRow = 4;                      // 插入資料列5
            int maxRow = sheet11.Cells.MaxDataRow; // 總計(列)位置
            double totalYrfund1 = 0;
            double totalYrfund2 = 0;
            string currentCity = string.Empty;
            int cityRowCount = 0;                  // 紀錄當前行政區的資料筆數
            double cityTotal = 0;                  // 紀錄當前行政區的金額總和
            double cityTotalFund1 = 0;             // 當前行政區中央款總和
            double cityTotalFund2 = 0;             // 當前行政區本期金額總和
            double grandTotalFund1 = 0;            // 全部中央款總和
            double grandTotalFund2 = 0;            // 全部本期金額總和

            /// 總表標題列
            sheet11.Cells.Merge(0, 0, 1, 10);
            /// sheet11.Cells[0, 0].PutValue(PJ_NAME + "總表");
            sheet11.Cells[0, 0].PutValue($"{SYR}年度 {PJ_NAME} 署辦");
            Aspose.Cells.Style SheetTitle = sheet11.Cells[0, 0].GetStyle();
            SheetTitle.HorizontalAlignment = TextAlignmentType.Center;
            sheet11.Cells[0, 0].SetStyle(SheetTitle);

            //Debug.WriteLine("DT1");
            //foreach (DataRow rows in dt1.Rows)
            //{
            //    foreach (DataColumn cols in dt1.Columns)
            //    {
            //        Debug.Write($"{cols.ColumnName}:{rows[cols]} \t");
            //    }
            //    Debug.WriteLine("");
            //}

            //Debug.WriteLine("DT2");
            //foreach (DataRow rows in dt2.Rows)
            //{
            //    foreach (DataColumn cols in dt2.Columns)
            //    {
            //        Debug.Write($"{cols.ColumnName}:{rows[cols]} \t");
            //    }
            //    Debug.WriteLine("");
            //}

            var subCity = new List<string>
            {
                "基隆市", "新北市", "台北市", "桃園市", "新竹縣", "新竹市", "苗栗縣",
                "台中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市", "台南市",
                "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣", "澎湖縣", "金門縣", "連江縣", "不分區"
            };

            if (dt1.Rows.Count > 0 && dt2.Rows.Count == 0)
            {
                // 核定"期數" / 計畫年度
                int DT1_SEQNO = Convert.ToInt32(dt1.Rows[0]["SEQ_NO"]);
                string SEQ_Text1 = DT1_SEQNO == 0 ? "原預算" : $"第{DT1_SEQNO}次核定預算";
                sheet11.Cells["F3"].PutValue(SEQ_Text1);
                sheet11.Cells["F4"].PutValue($"{SYR}");
                sheet11.Cells["G3"].PutValue("暫無預算");
                sheet11.Cells["G4"].PutValue($"{SYR}");

                foreach (DataRow rows in dt1.Rows)
                {
                    // 署辦 : IPJ_CODE='01' and ACC_ID<>'0200'
                    if (rows["IPJ_DESC"].ToString() == "署辦" ||
                         rows["IPJ_DESC"].ToString() == "署委辦" ||
                         rows["IPJ_DESC"].ToString() == "署代辦")
                    {
                        string sinCity = rows["CITY"].ToString();

                        // 如果行政區改變，檢查是否需要插入小計
                        if (!string.IsNullOrEmpty(currentCity) && currentCity != sinCity)
                        {
                            if (cityRowCount > 0) // 檢查當前行政區是否有資料
                            {
                                // 插入小計行
                                sheet11.Cells.CopyRow(sheet11.Cells, 4, startRow);
                                sheet11.Cells[startRow, 0].PutValue("");
                                sheet11.Cells[startRow, 1].PutValue("");
                                sheet11.Cells[startRow, 2].PutValue("");
                                sheet11.Cells[startRow, 3].PutValue("");
                                sheet11.Cells[startRow, 4].PutValue("小計");
                                sheet11.Cells[startRow, 5].PutValue(cityTotal);           // 行政區總計
                                sheet11.Cells[startRow, 6].PutValue(Convert.ToDouble(0));
                                startRow++;
                            }

                            cityTotal = 0;
                            cityRowCount = 0;
                        }

                        currentCity = sinCity; // 更新當前行政區

                        sheet11.Cells.CopyRow(sheet11.Cells, 4, startRow);
                        sheet11.Cells[startRow, 0].PutValue(++itemCount);
                        sheet11.Cells[startRow, 1].PutValue(rows["CITY"]);
                        sheet11.Cells[startRow, 2].PutValue(rows["SIP_TD_CODE"]);
                        sheet11.Cells[startRow, 3].PutValue(rows["TOWN"]);
                        sheet11.Cells[startRow, 4].PutValue(rows["TD_NAME"]);
                        double scenFund = Convert.ToDouble(rows["CEN_FUND"]);
                        sheet11.Cells[startRow, 5].PutValue(scenFund);
                        sheet11.Cells[startRow, 6].PutValue(Convert.ToDouble(0));

                        totalYrfund1 += scenFund;
                        startRow++;

                        // 累加行政區金額和資料筆數
                        cityTotal += scenFund;
                        cityRowCount++;
                    }
                }

                // 處理最後一個行政區的小計
                if (!string.IsNullOrEmpty(currentCity) && cityRowCount > 0)
                {
                    sheet11.Cells.CopyRow(sheet11.Cells, 4, startRow);
                    sheet11.Cells[startRow, 0].PutValue("");
                    sheet11.Cells[startRow, 1].PutValue("");
                    sheet11.Cells[startRow, 2].PutValue("");
                    sheet11.Cells[startRow, 3].PutValue("");
                    sheet11.Cells[startRow, 4].PutValue("小計");
                    sheet11.Cells[startRow, 5].PutValue(cityTotal);           // 行政區總計
                    sheet11.Cells[startRow, 6].PutValue(Convert.ToDouble(0));
                    sheet11.Cells[startRow, 9].PutValue("");
                    startRow++;
                }

                sheet11.Cells.CopyRow(sheet11.Cells, 4, startRow);
                sheet11.Cells[startRow, 0].PutValue("");
                sheet11.Cells[startRow, 1].PutValue("");
                sheet11.Cells[startRow, 2].PutValue("");
                sheet11.Cells[startRow, 3].PutValue("");
                sheet11.Cells[startRow, 4].PutValue("總計");
                sheet11.Cells[startRow, 5].PutValue(totalYrfund1);
                sheet11.Cells[startRow, 6].PutValue(totalYrfund2);
                sheet11.Cells[startRow, 9].PutValue("");
            }
            else if (dt1.Rows.Count > 0 && dt2.Rows.Count > 0)
            {
                // 核定"期數" / 計畫年度
                int DT1_SEQNO = Convert.ToInt32(dt1.Rows[0]["SEQ_NO"]);
                int DT2_SEQNO = Convert.ToInt32(dt2.Rows[0]["SEQ_NO"]);
                string SEQ_Text1 = DT1_SEQNO == 0 ? "原預算" : $"第{DT1_SEQNO}次核定預算";
                string SEQ_Text2 = DT2_SEQNO == 0 ? "原預算" : $"第{DT2_SEQNO}次核定預算";
                sheet11.Cells["G3"].PutValue(SEQ_Text1);
                sheet11.Cells["G4"].PutValue($"{SYR}");
                sheet11.Cells["F3"].PutValue(SEQ_Text2);
                sheet11.Cells["F4"].PutValue($"{SYR}");

                foreach (string city in subCity)
                {
                    bool hasCityData = false; // 行政區判斷

                    foreach (DataRow row2 in dt2.Rows)
                    {
                        if (row2["IPJ_DESC"].ToString() == "署辦" ||
                             row2["IPJ_DESC"].ToString() == "署委辦" ||
                             row2["IPJ_DESC"].ToString() == "署代辦")
                        {
                            if (row2["CITY"].ToString() == city) // 署辦/署代辦/署委辦,且行政區符合 subCity 內的行政區
                            {
                                hasCityData = true;

                                // 對比本期中(DT1)與前一期(DT2)的TID
                                string TID = row2["TID"].ToString();
                                DataRow[] matchingRows = dt1.Select($"TID = '{TID}'");

                                sheet11.Cells.CopyRow(sheet11.Cells, 4, startRow);
                                sheet11.Cells[startRow, 0].PutValue(++itemCount);                       // 項次
                                sheet11.Cells[startRow, 1].PutValue(row2["CITY"].ToString());           // 縣市別
                                sheet11.Cells[startRow, 2].PutValue(row2["SIP_TD_CODE"].ToString());    // 工程編號
                                sheet11.Cells[startRow, 3].PutValue(row2["TOWN"].ToString());           // 鄉鎮別
                                sheet11.Cells[startRow, 4].PutValue(row2["TD_NAME"].ToString());        // 工程名稱
                                double scenFund = Convert.ToDouble(row2["CEN_FUND"]);                   // 本期中央款
                                sheet11.Cells[startRow, 5].PutValue(scenFund);                          // 前一期中央款
                                sheet11.Cells[startRow, 9].PutValue(row2["IPJ_CODE_DESC"].ToString());

                                // 更新中央款(本期)中央款
                                if (matchingRows.Length > 0)
                                {
                                    double cenFund2 = Convert.ToDouble(matchingRows[0]["CEN_FUND"]);
                                    sheet11.Cells[startRow, 6].PutValue(cenFund2);                      // 本期中央款
                                    totalYrfund2 += cenFund2;
                                    cityTotalFund2 += cenFund2;
                                    grandTotalFund2 += cenFund2;
                                }
                                else
                                {
                                    sheet11.Cells[startRow, 6].PutValue(0);                             // 若無本期資料則填 0
                                }

                                totalYrfund1 += scenFund;
                                cityTotalFund1 += scenFund;
                                grandTotalFund1 += scenFund;
                                startRow++;
                            }
                        }
                    }

                    // 檢查 DT1 中是否有該行政區但 DT2 中沒有的資料
                    foreach (DataRow row1 in dt1.Rows)
                    {
                        if (row1["IPJ_DESC"].ToString() == "署辦" ||
                            row1["IPJ_DESC"].ToString() == "署委辦" ||
                            row1["IPJ_DESC"].ToString() == "署代辦")
                        {
                            if (row1["CITY"].ToString() == city && !dt2.Select($"TID = '{row1["TID"].ToString()}'").Any())
                            {
                                hasCityData = true;
                                sheet11.Cells.CopyRow(sheet11.Cells, 4, startRow);
                                sheet11.Cells[startRow, 0].PutValue(++itemCount);                      // 項次
                                sheet11.Cells[startRow, 1].PutValue(row1["CITY"].ToString());          // 縣市別
                                sheet11.Cells[startRow, 2].PutValue(row1["SIP_TD_CODE"].ToString());   // 工程編號
                                sheet11.Cells[startRow, 3].PutValue(row1["TOWN"].ToString());          // 鄉鎮別
                                sheet11.Cells[startRow, 4].PutValue(row1["TD_NAME"].ToString());       // 工程名稱
                                sheet11.Cells[startRow, 5].PutValue(0);                                // 前一期無預算
                                double cenFund2 = Convert.ToDouble(row1["CEN_FUND"]);
                                sheet11.Cells[startRow, 6].PutValue(cenFund2);                         // 本期中央款
                                sheet11.Cells[startRow, 9].PutValue(row1["IPJ_CODE_DESC"].ToString()); // 署辦

                                cityTotalFund2 += cenFund2;
                                grandTotalFund2 += cenFund2;
                                startRow++;
                            }
                        }
                    }

                    // 插入行政區小計
                    if (hasCityData)
                    {
                        sheet11.Cells.CopyRow(sheet11.Cells, 4, startRow);
                        sheet11.Cells[startRow, 0].PutValue("");
                        sheet11.Cells[startRow, 1].PutValue("");
                        sheet11.Cells[startRow, 2].PutValue("");
                        sheet11.Cells[startRow, 3].PutValue("");
                        sheet11.Cells[startRow, 4].PutValue("小計");
                        sheet11.Cells[startRow, 5].PutValue(cityTotalFund1);
                        sheet11.Cells[startRow, 6].PutValue(cityTotalFund2);
                        sheet11.Cells[startRow, 9].PutValue("");
                        cityTotalFund1 = 0;
                        cityTotalFund2 = 0;
                        startRow++;
                    }
                }

                sheet11.Cells.CopyRow(sheet11.Cells, 4, startRow);
                sheet11.Cells[startRow, 0].PutValue("");
                sheet11.Cells[startRow, 1].PutValue("");
                sheet11.Cells[startRow, 2].PutValue("");
                sheet11.Cells[startRow, 3].PutValue("");
                sheet11.Cells[startRow, 4].PutValue("總計");
                sheet11.Cells[startRow, 5].PutValue(totalYrfund1);
                sheet11.Cells[startRow, 6].PutValue(totalYrfund2);
                sheet11.Cells[startRow, 9].PutValue("");
            }
        }

        /// <summary>
        /// I-1.雨水例行性管考報表 => (經費調整表)地辦
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet12"></param>
        /// <param name="dt1"></param>
        /// <param name="dt2"></param>
        /// <param name="ditdt"></param>
        /// <param name="SYR"></param>
        /// <param name="PSEQ_NO"></param>
        /// <param name="SEQ_NO"></param>
        /// <param name="PJ_NAME"></param>
        /// <see cref="New_GetRPT0001Data(decimal, string, string, string, string)"/> 資料來源
        public static void LocalOfficeTable(Workbook workbook, Worksheet sheet12, DataTable dt1, DataTable dt2, string SYR, string PSEQ_NO, string SEQ_NO, string PJ_NAME)
        {
            Worksheet Sheet12work = workbook.Worksheets["Sheet12"];

            int itemCount = 0;                     // 項次
            int startRow = 4;                      // 插入資料列5
            int maxRow = sheet12.Cells.MaxDataRow; // 總計(列)位置
            double totalYrfund1 = 0;
            double totalYrfund2 = 0;
            string currentCity = string.Empty;
            int cityRowCount = 0;                  // 紀錄當前行政區的資料筆數
            double cityTotal = 0;                  // 紀錄當前行政區的金額總和
            double cityTotalFund1 = 0;             // 當前行政區中央款總和
            double cityTotalFund2 = 0;             // 當前行政區本期金額總和
            double grandTotalFund1 = 0;            // 全部中央款總和
            double grandTotalFund2 = 0;            // 全部本期金額總和

            // 總表標題列
            sheet12.Cells.Merge(0, 0, 1, 10);
            /// sheet12.Cells[0, 0].PutValue(PJ_NAME + "總表");
            sheet12.Cells[0, 0].PutValue($"{SYR}年度 {PJ_NAME} 地辦");
            Aspose.Cells.Style SheetTitle = sheet12.Cells[0, 0].GetStyle();
            SheetTitle.HorizontalAlignment = TextAlignmentType.Center;
            sheet12.Cells[0, 0].SetStyle(SheetTitle);

            var subCity = new List<string>
            {
                "基隆市", "新北市", "台北市", "桃園市", "新竹縣", "新竹市", "苗栗縣",
                "台中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市", "台南市",
                "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣", "澎湖縣", "金門縣", "連江縣", "不分區"
            };

            if (dt1.Rows.Count > 0 && dt2.Rows.Count == 0)
            {
                // 核定"期數" / 計畫年度
                int DT1_SEQNO = Convert.ToInt32(dt1.Rows[0]["SEQ_NO"]);
                string SEQ_Text1 = DT1_SEQNO == 0 ? "原預算" : $"第{DT1_SEQNO}次核定預算";
                sheet12.Cells["F3"].PutValue(SEQ_Text1);
                sheet12.Cells["F4"].PutValue($"{SYR}");
                sheet12.Cells["G3"].PutValue("暫無預算");
                sheet12.Cells["G4"].PutValue($"{SYR}");

                foreach (DataRow rows in dt1.Rows)
                {
                    if (rows["IPJ_DESC"].ToString() == "地辦" ||
                        rows["IPJ_DESC"].ToString() == "地方辦")
                    {
                        string sinCity = rows["CITY"].ToString();

                        // 如果行政區改變，檢查是否需要插入小計
                        if (!string.IsNullOrEmpty(currentCity) && currentCity != sinCity)
                        {
                            if (cityRowCount > 0) // 檢查當前行政區是否有資料
                            {
                                // 插入小計行
                                sheet12.Cells.CopyRow(sheet12.Cells, 4, startRow);
                                sheet12.Cells[startRow, 0].PutValue("");
                                sheet12.Cells[startRow, 1].PutValue("");
                                sheet12.Cells[startRow, 2].PutValue("");
                                sheet12.Cells[startRow, 3].PutValue("");
                                sheet12.Cells[startRow, 4].PutValue("小計");
                                sheet12.Cells[startRow, 5].PutValue(cityTotal);           // 行政區總計
                                sheet12.Cells[startRow, 6].PutValue(Convert.ToDouble(0));
                                startRow++;
                            }

                            cityTotal = 0;
                            cityRowCount = 0;
                        }

                        currentCity = sinCity; // 更新當前行政區

                        sheet12.Cells.CopyRow(sheet12.Cells, 4, startRow);
                        sheet12.Cells[startRow, 0].PutValue(++itemCount);
                        sheet12.Cells[startRow, 1].PutValue(rows["CITY"]);
                        sheet12.Cells[startRow, 2].PutValue(rows["SIP_TD_CODE"]);
                        sheet12.Cells[startRow, 3].PutValue(rows["TOWN"]);
                        sheet12.Cells[startRow, 4].PutValue(rows["TD_NAME"]);
                        double scenFund = Convert.ToDouble(rows["CEN_FUND"]);
                        sheet12.Cells[startRow, 5].PutValue(scenFund);
                        sheet12.Cells[startRow, 6].PutValue(Convert.ToDouble(0));

                        totalYrfund1 += scenFund;
                        startRow++;

                        // 累加行政區金額和資料筆數
                        cityTotal += scenFund;
                        cityRowCount++;
                    }
                }

                // 處理最後一個行政區的小計
                if (!string.IsNullOrEmpty(currentCity) && cityRowCount > 0)
                {
                    sheet12.Cells.CopyRow(sheet12.Cells, 4, startRow);
                    sheet12.Cells[startRow, 0].PutValue("");
                    sheet12.Cells[startRow, 1].PutValue("");
                    sheet12.Cells[startRow, 2].PutValue("");
                    sheet12.Cells[startRow, 3].PutValue("");
                    sheet12.Cells[startRow, 4].PutValue("小計");
                    sheet12.Cells[startRow, 5].PutValue(cityTotal);           // 行政區總計
                    sheet12.Cells[startRow, 6].PutValue(Convert.ToDouble(0));
                    sheet12.Cells[startRow, 9].PutValue("");
                    startRow++;
                }

                sheet12.Cells.CopyRow(sheet12.Cells, 4, startRow);
                sheet12.Cells[startRow, 0].PutValue("");
                sheet12.Cells[startRow, 1].PutValue("");
                sheet12.Cells[startRow, 2].PutValue("");
                sheet12.Cells[startRow, 3].PutValue("");
                sheet12.Cells[startRow, 4].PutValue("總計");
                sheet12.Cells[startRow, 5].PutValue(totalYrfund1);
                sheet12.Cells[startRow, 6].PutValue(totalYrfund2);
                sheet12.Cells[startRow, 9].PutValue("");
            }
            else if (dt1.Rows.Count > 0 && dt2.Rows.Count > 0)
            {
                // 核定"期數" / 計畫年度
                int DT1_SEQNO = Convert.ToInt32(dt1.Rows[0]["SEQ_NO"]);
                int DT2_SEQNO = Convert.ToInt32(dt2.Rows[0]["SEQ_NO"]);
                string SEQ_Text1 = DT1_SEQNO == 0 ? "原預算" : $"第{DT1_SEQNO}次核定預算";
                string SEQ_Text2 = DT2_SEQNO == 0 ? "原預算" : $"第{DT2_SEQNO}次核定預算";
                sheet12.Cells["G3"].PutValue(SEQ_Text1);
                sheet12.Cells["G4"].PutValue($"{SYR}");
                sheet12.Cells["F3"].PutValue(SEQ_Text2);
                sheet12.Cells["F4"].PutValue($"{SYR}");

                foreach (string city in subCity)
                {
                    bool hasCityData = false; // 行政區判斷

                    foreach (DataRow row2 in dt2.Rows)
                    {
                        if (row2["IPJ_DESC"].ToString() == "地辦" ||
                            row2["IPJ_DESC"].ToString() == "地方辦")
                        {
                            if (row2["CITY"].ToString() == city) // 署辦/署代辦/署委辦,且行政區符合 subCity 內的行政區
                            {
                                hasCityData = true;

                                // 對比本期中(DT1)與前一期(DT2)的TID
                                string TID = row2["TID"].ToString();
                                DataRow[] matchingRows = dt1.Select($"TID = '{TID}'");

                                sheet12.Cells.CopyRow(sheet12.Cells, 4, startRow);
                                sheet12.Cells[startRow, 0].PutValue(++itemCount);                       // 項次
                                sheet12.Cells[startRow, 1].PutValue(row2["CITY"].ToString());           // 縣市別
                                sheet12.Cells[startRow, 2].PutValue(row2["SIP_TD_CODE"].ToString());    // 工程編號
                                sheet12.Cells[startRow, 3].PutValue(row2["TOWN"].ToString());           // 鄉鎮別
                                sheet12.Cells[startRow, 4].PutValue(row2["TD_NAME"].ToString());        // 工程名稱
                                double scenFund = Convert.ToDouble(row2["CEN_FUND"]);                   // 本期中央款
                                sheet12.Cells[startRow, 5].PutValue(scenFund);                          // 前一期中央款
                                sheet12.Cells[startRow, 9].PutValue(row2["IPJ_CODE_DESC"].ToString());

                                // 更新中央款(本期)中央款
                                if (matchingRows.Length > 0)
                                {
                                    double cenFund2 = Convert.ToDouble(matchingRows[0]["CEN_FUND"]);
                                    sheet12.Cells[startRow, 6].PutValue(cenFund2);                      // 本期中央款
                                    totalYrfund2 += cenFund2;
                                    cityTotalFund2 += cenFund2;
                                    grandTotalFund2 += cenFund2;
                                }
                                else
                                {
                                    sheet12.Cells[startRow, 6].PutValue(0);                             // 若無本期資料則填 0
                                }

                                totalYrfund1 += scenFund;
                                cityTotalFund1 += scenFund;
                                grandTotalFund1 += scenFund;
                                startRow++;
                            }
                        }
                    }

                    // 檢查 DT1 中是否有該行政區但 DT2 中沒有的資料
                    foreach (DataRow row1 in dt1.Rows)
                    {
                        if (row1["IPJ_DESC"].ToString() == "地辦" ||
                            row1["IPJ_DESC"].ToString() == "地方辦")
                        {
                            if (row1["CITY"].ToString() == city && !dt2.Select($"TID = '{row1["TID"].ToString()}'").Any())
                            {
                                hasCityData = true;
                                sheet12.Cells.CopyRow(sheet12.Cells, 4, startRow);
                                sheet12.Cells[startRow, 0].PutValue(++itemCount);                      // 項次
                                sheet12.Cells[startRow, 1].PutValue(row1["CITY"].ToString());          // 縣市別
                                sheet12.Cells[startRow, 2].PutValue(row1["SIP_TD_CODE"].ToString());   // 工程編號
                                sheet12.Cells[startRow, 3].PutValue(row1["TOWN"].ToString());          // 鄉鎮別
                                sheet12.Cells[startRow, 4].PutValue(row1["TD_NAME"].ToString());       // 工程名稱
                                sheet12.Cells[startRow, 5].PutValue(0);                                // 前一期無預算
                                double cenFund2 = Convert.ToDouble(row1["CEN_FUND"]);
                                sheet12.Cells[startRow, 6].PutValue(cenFund2);                         // 本期中央款
                                sheet12.Cells[startRow, 9].PutValue(row1["IPJ_CODE_DESC"].ToString()); // 署辦

                                cityTotalFund2 += cenFund2;
                                grandTotalFund2 += cenFund2;
                                startRow++;
                            }
                        }
                    }

                    // 插入行政區小計
                    if (hasCityData)
                    {
                        sheet12.Cells.CopyRow(sheet12.Cells, 4, startRow);
                        sheet12.Cells[startRow, 0].PutValue("");
                        sheet12.Cells[startRow, 1].PutValue("");
                        sheet12.Cells[startRow, 2].PutValue("");
                        sheet12.Cells[startRow, 3].PutValue("");
                        sheet12.Cells[startRow, 4].PutValue("小計");
                        sheet12.Cells[startRow, 5].PutValue(cityTotalFund1);
                        sheet12.Cells[startRow, 6].PutValue(cityTotalFund2);
                        sheet12.Cells[startRow, 9].PutValue("");
                        cityTotalFund1 = 0;
                        cityTotalFund2 = 0;
                        startRow++;
                    }
                }

                sheet12.Cells.CopyRow(sheet12.Cells, 4, startRow);
                sheet12.Cells[startRow, 0].PutValue("");
                sheet12.Cells[startRow, 1].PutValue("");
                sheet12.Cells[startRow, 2].PutValue("");
                sheet12.Cells[startRow, 3].PutValue("");
                sheet12.Cells[startRow, 4].PutValue("總計");
                sheet12.Cells[startRow, 5].PutValue(totalYrfund1);
                sheet12.Cells[startRow, 6].PutValue(totalYrfund2);
                sheet12.Cells[startRow, 9].PutValue("");
            }

            // 增減欄位
            // 從 StartCow1 開始逐行處理
            for (int i = startRow; i <= maxRow; i++)
            {
                //double valueF = sheet12.Cells[i, 5].DoubleValue; // F列 (索引5)
                //double valueG = sheet12.Cells[i, 6].DoubleValue; // G列 (索引6)
                double valueF = sheet12.Cells[i, 5].Value != null && double.TryParse(sheet12.Cells[i, 5].Value.ToString(), out double f) ? f : 0;
                double valueG = sheet12.Cells[i, 6].Value != null && double.TryParse(sheet12.Cells[i, 6].Value.ToString(), out double g) ? g : 0;
                double diff = valueG - valueF;                   // 計算增減值 (G - F)
                sheet12.Cells[i, 7].PutValue(diff);              // 將結果放入 H 欄 (索引7)
            }
        }

        /// <summary>
        /// I-1.雨水例行性管考報表 => (經費調整表)業務費
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet13"></param>
        /// <param name="dt1"></param>
        /// <param name="dt2"></param>
        /// <param name="dt3"></param>
        /// <param name="ditdt"></param>
        /// <param name="SYR"></param>
        /// <param name="PSEQ_NO"></param>
        /// <param name="SEQ_NO"></param>
        /// <param name="PJ_NAME"></param
        /// <see cref="New_GetRPT0001Data(decimal, string, string, string, string)"/> 資料來源
        public static void BusinessExpensesTable(Workbook workbook, Worksheet sheet13, DataTable dt1, DataTable dt2, DataTable dt3, string SYR, string PSEQ_NO, string SEQ_NO, string PJ_NAME)
        {
            Worksheet Sheet13work = workbook.Worksheets["Sheet13"];

            int itemCount = 0;
            int startRow = 4;
            double TotalB_Fund = 0;
            double TotalB_Fund2 = 0;
            int maxRow = sheet13.Cells.MaxDataRow; // 總計(列)位置

            // 總表標題列
            sheet13.Cells.Merge(0, 0, 1, 10);
            /// sheet13.Cells[0, 0].PutValue(PJ_NAME + "總表");
            sheet13.Cells[0, 0].PutValue($"{SYR}年度 {PJ_NAME} 業務費");
            Aspose.Cells.Style SheetTitle = sheet13.Cells[0, 0].GetStyle();
            SheetTitle.HorizontalAlignment = TextAlignmentType.Center;
            sheet13.Cells[0, 0].SetStyle(SheetTitle);

            if (dt1.Rows.Count > 0 && dt2.Rows.Count == 0)
            {
                foreach (DataRow rows in dt1.Rows)
                {
                    // 核定"期數" / 計畫年度
                    int DT1_SEQNO = Convert.ToInt32(dt1.Rows[0]["SEQ_NO"]);
                    string SEQ_Text1 = DT1_SEQNO == 0 ? "原預算" : $"第{DT1_SEQNO}次核定預算";
                    sheet13.Cells["F3"].PutValue(SEQ_Text1);
                    sheet13.Cells["F4"].PutValue($"{SYR}");
                    sheet13.Cells["G3"].PutValue("暫無預算");
                    sheet13.Cells["G4"].PutValue($"{SYR}");

                    if (rows["ACC_ID"].ToString() == "0200")
                    {
                        sheet13.Cells.CopyRow(sheet13.Cells, 4, startRow);

                        sheet13.Cells[startRow, 0].PutValue(++itemCount);
                        sheet13.Cells[startRow, 1].PutValue(rows["CITY"]);
                        sheet13.Cells[startRow, 2].PutValue(rows["SIP_TD_CODE"]);
                        sheet13.Cells[startRow, 3].PutValue(rows["TOWN"]);
                        sheet13.Cells[startRow, 4].PutValue(rows["TD_NAME"]);
                        sheet13.Cells[startRow, 5].PutValue(Convert.ToDouble(rows["CEN_FUND"]));
                        sheet13.Cells[startRow, 6].PutValue(Convert.ToDouble(0));

                        if (double.TryParse(rows["CEN_FUND"].ToString(), out double yrFund))
                        {
                            TotalB_Fund += yrFund;
                        }

                        startRow++;
                    }
                }
            }
            else if (dt1.Rows.Count > 0 && dt2.Rows.Count > 0)
            {
                // 核定"期數" / 計畫年度
                int DT1_SEQNO = Convert.ToInt32(dt1.Rows[0]["SEQ_NO"]);
                int DT2_SEQNO = Convert.ToInt32(dt2.Rows[0]["SEQ_NO"]);
                string SEQ_Text1 = DT1_SEQNO == 0 ? "原預算" : $"第{DT1_SEQNO}次核定預算";
                string SEQ_Text2 = DT2_SEQNO == 0 ? "原預算" : $"第{DT2_SEQNO}次核定預算";
                sheet13.Cells["G3"].PutValue(SEQ_Text1);
                sheet13.Cells["G4"].PutValue($"{SYR}");
                sheet13.Cells["F3"].PutValue(SEQ_Text2);
                sheet13.Cells["F4"].PutValue($"{SYR}");

                // 資料填入
                foreach (DataRow row2 in dt2.Rows)
                {
                    if (row2["ACC_ID"].ToString() == "0200")
                    {
                        // 比對 DT2(前一期) 和 DT1(本期中) 的 TID
                        string TID = row2["TID"].ToString();
                        DataRow[] matchingRows = dt1.Select($"TID = '{TID}'");

                        sheet13.Cells.CopyRow(sheet13.Cells, 4, startRow); // 複製格式 Style
                        sheet13.Cells[startRow, 0].PutValue(++itemCount);                          // 項次
                        sheet13.Cells[startRow, 1].PutValue(row2["CITY"].ToString());            // 縣市別
                        sheet13.Cells[startRow, 2].PutValue(row2["SIP_TD_CODE"].ToString());     // 工程編號
                        sheet13.Cells[startRow, 3].PutValue(row2["TOWN"].ToString());            // 鄉鎮別
                        sheet13.Cells[startRow, 4].PutValue(row2["TD_NAME"].ToString());         // 工程名稱
                        sheet13.Cells[startRow, 5].PutValue(Convert.ToDouble(row2["CEN_FUND"])); // 中央款
                        sheet13.Cells[startRow, 9].PutValue(row2["IPJ_CODE_DESC"].ToString());   // 署辦

                        if (double.TryParse(row2["CEN_FUND"].ToString(), out double yrFund1))
                        {
                            TotalB_Fund += yrFund1;
                        }

                        if (matchingRows.Length > 0)
                        {
                            // 更新金額，填入新一期的 CEN_FUND 金額
                            double CenFUND = Convert.ToDouble(matchingRows[0]["CEN_FUND"]);      // 更新中央款
                            sheet13.Cells[startRow, 6].PutValue(CenFUND);
                            TotalB_Fund2 += CenFUND;
                        }
                        startRow++; // 換下一筆資料
                    }

                }

                // 比對 DT2(前一期) 和 DT1(本期中)
                foreach (DataRow row1 in dt1.Rows)
                {
                    if (row1["ACC_ID"].ToString() == "0200")
                    {
                        string TID = row1["TID"].ToString();
                        DataRow[] matchingRowsInDt2 = dt2.Select($"TID = '{TID}'");

                        if (matchingRowsInDt2.Length == 0 && row1["PID_SYR"].ToString() == SYR)
                        {
                            // TID 在 DT2(前一期) 中找不到，新增至 Excel
                            sheet13.Cells.CopyRow(sheet13.Cells, 4, startRow);
                            sheet13.Cells[startRow, 0].PutValue(++itemCount);                         // 項次
                            sheet13.Cells[startRow, 1].PutValue(row1["CITY"].ToString());           // 縣市別
                            sheet13.Cells[startRow, 2].PutValue(row1["SIP_TD_CODE"].ToString());    // 工程編號
                            sheet13.Cells[startRow, 3].PutValue(row1["TOWN"].ToString());           // 鄉鎮別
                            sheet13.Cells[startRow, 4].PutValue(row1["TD_NAME"].ToString());        // 工程名稱
                            sheet13.Cells[startRow, 5].PutValue(0);                                 // 新增的前一期無預算
                            double newCenFund = Convert.ToDouble(row1["CEN_FUND"]);                 // 新一期(中央款)
                            sheet13.Cells[startRow, 6].PutValue(newCenFund);
                            sheet13.Cells[startRow, 9].PutValue(row1["IPJ_CODE_DESC"].ToString());  // 署辦

                            TotalB_Fund2 += newCenFund;

                            startRow++; // 換下一筆資料
                        }
                    }

                }
            }

            // 增減欄位
            // 從 StartCow1 開始逐行處理
            for (int i = startRow; i <= maxRow; i++)
            {
                double valueF = sheet13.Cells[i, 5].Value != null && double.TryParse(sheet13.Cells[i, 5].Value.ToString(), out double f) ? f : 0;
                double valueG = sheet13.Cells[i, 6].Value != null && double.TryParse(sheet13.Cells[i, 6].Value.ToString(), out double g) ? g : 0;
                double diff = valueG - valueF;                   // 計算增減值 (G - F)
                sheet13.Cells[i, 7].PutValue(diff);              // 將結果放入 H 欄 (索引7)
            }

            // 原預算/期數(總計)
            sheet13.Cells.CopyRow(sheet13.Cells, 4, startRow);
            sheet13.Cells[startRow, 0].PutValue("");
            sheet13.Cells[startRow, 1].PutValue("");
            sheet13.Cells[startRow, 2].PutValue("");
            sheet13.Cells[startRow, 3].PutValue("");
            sheet13.Cells[startRow, 4].PutValue("總計");
            sheet13.Cells[startRow, 5].PutValue(TotalB_Fund);
            sheet13.Cells[startRow, 6].PutValue(TotalB_Fund2 != 0 ? TotalB_Fund2 : 0);
            sheet13.Cells[startRow, 9].PutValue("");
        }

        /// <summary>
        /// I-1.雨水例行性管考報表[執行情形表]-資料來源
        /// </summary>
        /// <param name="AN_PID"></param>
        /// <param name="PJ_SYR"></param>
        /// <param name="PJ_SMN"></param>
        /// <param name="connectionString"></param>
        /// <param name="DEPT"></param>
        /// <returns></returns>
        public static DataView GetRPT0002Data(decimal AN_PID, string PJ_SYR, string PJ_SMN, string connectionString, string DEPT)
        {
            string sql = "";
            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");
            string sPPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");
            string sPSMN = (System.DateTime.Now.AddMonths(-1).Month).ToString("00");
            string sPSYRMN = "";//前一個月分

            if (PJ_SYR != "")
            {
                sSYR = PJ_SYR;
                sPSYR = (Convert.ToInt32(sSYR) - 1).ToString("000");
                sPPSYR = (Convert.ToInt32(sSYR) - 2).ToString("000");

            }
            if (PJ_SMN != "")
            {
                sSMN = PJ_SMN;
                sPSMN = (Convert.ToInt32(sSMN) - 1).ToString("00");
            }

            switch (sSMN)
            {
                case "01":
                    sPSYRMN = sPSYR + "12";
                    break;
                default:
                    sPSYRMN = sSYR + sPSMN;
                    break;
            }

            string SYR = sSYR;
            string PSYR = sPSYR;

            sql = $@"
                    SELECT 
                        C.TD_NAME,C.CITY,
                        C.TOWN,C.TD_CODE,C.SIP_TD_CODE,
                        C.CEN_SP_RATIO / 100.0 AS CEN_SP_RATIO,
                        ISNULL(T.TD_FUND, 0) AS TD_FUND,
                        ISNULL((SELECT SUM(FUND) FROM CMEPMS.dbo.CTDITFD WHERE TID IN ( SELECT TID FROM CMEPMS.dbo.CTDMNDS WHERE TD_CODE = C.TD_CODE ) AND SYR = '{sPSYR}' AND FIT_CODE = '00' ), 0) AS CEN_FUND_1,
                        ISNULL((SELECT SUM(FUND) FROM CMEPMS.dbo.CTDITFD WHERE TID IN ( SELECT TID FROM CMEPMS.dbo.CTDMNDS WHERE TD_CODE = C.TD_CODE ) AND SYR = '{sSYR}' AND FIT_CODE = '00' ), 0) AS CEN_FUND_2,
                        ISNULL(C.CEN_FUND, 0) + CASE ISNULL(R1.R_COUNT, 0) WHEN 0 THEN ISNULL(( SELECT SUM(FUND) FROM CMEPMS.dbo.CTDITFD WHERE TID IN ( SELECT TID FROM CMEPMS.dbo.CTDMNDS WHERE TD_CODE = C.TD_CODE ) AND SYR = '{sPSYR}' AND FIT_CODE = '00' ), 0) - ISNULL(F1.STCF_FUND, 0) ELSE ISNULL(R.OCR_FUND, 0) END AS YR_FUND,
                        ISNULL(C.CEN_FUND, 0) AS CEN_FUND,
                        ISNULL(C.LOC_FUND, 0) AS LOC_FUND,
                        ISNULL(R.OCR_FUND, 0) AS OCR_FUND,  -- 發生權責數
                        ISNULL(R.REV_FUND, 0) AS REV_FUND,  -- 保留款
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.L_PLN_COST ELSE M.L_PLN_COST END, 0 ) AS L_PLN_COST,   -- 中央款累計分配數
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}12' THEN M2.L_PLN_COST ELSE M.L_PLN_COST END, 0 ) AS P_PLN_COST,       -- 前一年度中央款累計分配數
                        ISNULL(F2.STCF_FUND, 0) AS L_REL_COST,
            
                        -- 前一年度中央款累計實際支用數
                        ISNULL(F1.STCF_FUND, 0) - ISNULL(( SELECT SUM(ISNULL(COST_FEE, 0) + ISNULL(PPAY_FEE, 0)) FROM CMEPMS.dbo.TDDSTCF WHERE REPLACE(TD_CODE, '-', '') = M.TD_CODE AND SYR = '{sPSYR}' AND SMN = CONVERT(INT, '12') - 1 ), 0) AS PL_REL_COST,
                        
                        ISNULL(F2.STCF_FUND, 0) AS P_PLN_REL_COST,
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.L_UNPAY_COST ELSE M.L_UNPAY_COST END, 0 ) AS L_UNPAY_COST,     -- 中央款累計應付未付數
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.SAVE_FEE ELSE M.SAVE_FEE END, 0 ) AS SAVE_FEE,                 -- 節餘數
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.PRE_CON_LEN ELSE M.PRE_CON_LEN END, 0 ) AS PRE_CON_LEN,        -- 預定施作長度
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.REL_CON_LEN ELSE M.REL_CON_LEN END, 0 ) AS REL_CON_LEN,        -- 實際施作長度
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.PRE_CON_FLOOD ELSE M.PRE_CON_FLOOD END, 0 ) AS PRE_CON_FLOOD,  -- 預定施作滯洪池量
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.REL_CON_FLOOD ELSE M.REL_CON_FLOOD END, 0 ) AS REL_CON_FLOOD,  -- 實際施作滯洪池量
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.PRE_SCHD ELSE M.PRE_SCHD END, 0 ) / 100.0 AS PRE_SCHD,         -- 預定進度
                        ISNULL( CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.REL_SCHD ELSE M.REL_SCHD END, 0 ) / 100.0 AS REL_SCHD,         -- 實際進度
                        CASE WHEN M1.SYRMN <= '{sSYR}{sSMN}' THEN M2.EXECT_DESC ELSE M.EXECT_DESC END AS EXECT_DESC,                        -- 辦理情形說明(備註)

                        STUFF(STUFF(RIGHT('0000000' + CAST(S1.PRE_EDT AS VARCHAR(7)), 7),4, 0, '.'),7, 0, '.') AS PRE_TK_103,
	                    STUFF(STUFF(RIGHT('0000000' + CAST(S1.REL_EDT AS VARCHAR(7)), 7),4, 0, '.'),7, 0, '.') AS REL_TK_103,
	                    STUFF(STUFF(RIGHT('0000000' + CAST(S2.PRE_EDT AS VARCHAR(7)), 7),4, 0, '.'),7, 0, '.') AS PRE_TK_105,
	                    STUFF(STUFF(RIGHT('0000000' + CAST(S2.REL_EDT AS VARCHAR(7)), 7),4, 0, '.'),7, 0, '.') AS REL_TK_105,

                        CASE ISNULL(C.CATY_CODE, '') WHEN 'P' THEN S3.PRE_EDT WHEN 'L' THEN S4.PRE_EDT ELSE S5.PRE_EDT END AS PRE_TK_107,
                        CASE ISNULL(C.CATY_CODE, '') WHEN 'P' THEN S3.REL_EDT WHEN 'L' THEN S4.REL_EDT ELSE S5.REL_EDT END AS REL_TK_107,

                        C.BUDGET_FEE_B,         -- 經費(預算)千元 => 總經費
                        C.BUDGET_FEE_B_0101,    -- 經費(預算)千元 => 施工費
                        C.BUDGET_FEE_B_N0101,   -- 經費(預算)千元 => 間接費用
                        C.BUDGET_FEE_C,         -- 經費(決標)千元 => 總經費
                        C.BUDGET_FEE_C_0101,    -- 經費(決標)千元 => 施工費
                        C.BUDGET_FEE_C_N0101,   -- 經費(決標)千元 => 間接費用

                        STUFF(STUFF(RIGHT('0000000' + CAST(TDD.PRE_SDT AS VARCHAR(7)), 7),4, 0, '.'),7, 0, '.') AS EXTRA_PRE_SDT,
                        STUFF(STUFF(RIGHT('0000000' + CAST(TDD.PRE_EDT AS VARCHAR(7)), 7),4, 0, '.'),7, 0, '.') AS EXTRA_PRE_EDT

                    FROM CMEPMS.dbo.TDV_TDDCFTS C
                    LEFT JOIN CMETMS.dbo.TDMFORM T ON T.TD_CODE = C.TD_CODE
                    LEFT JOIN CMETMS.dbo.TDDMSCD M ON M.SYR + M.SMN = '{sSYR}{sSMN}' AND M.TD_CODE = C.TD_CODE
                    LEFT JOIN ( SELECT TD_CODE, MAX(SYR + SMN) AS SYRMN FROM CMETMS.dbo.TDDMSCD GROUP BY TD_CODE ) M1 ON M1.TD_CODE = C.TD_CODE
                    LEFT JOIN CMETMS.dbo.TDDMSCD M2 ON M2.SYR + M2.SMN = M1.SYRMN AND M2.TD_CODE = C.TD_CODE
                    LEFT JOIN ( SELECT PID, SYR, TD_CODE1 AS TD_CODE, SUM(REV_FUND) AS REV_FUND, SUM(OCR_FUND) AS OCR_FUND FROM CMEPMS.dbo.CRDMNDS GROUP BY PID, SYR, TD_CODE1 ) R ON R.TD_CODE = C.TD_CODE AND R.PID = C.PID
                    LEFT JOIN ( SELECT PID, COUNT(TD_CODE) AS R_COUNT FROM CMEPMS.dbo.CRDMNDS GROUP BY PID ) R1 ON R1.PID = C.PID
                    LEFT JOIN CMETMS.dbo.TDDSDWB S1 ON S1.TD_CODE = C.TD_CODE AND S1.IT_CODE = 'TK1041'
                    LEFT JOIN CMETMS.dbo.TDDSDWB S2 ON S2.TD_CODE = C.TD_CODE AND S2.IT_CODE = 'TK105'
                    LEFT JOIN CMETMS.dbo.TDDSDWB S3 ON S3.TD_CODE = C.TD_CODE AND S3.IT_CODE = 'BG107'
                    LEFT JOIN CMETMS.dbo.TDDSDWB S4 ON S4.TD_CODE = C.TD_CODE AND S4.IT_CODE = 'TF108'
                    LEFT JOIN CMETMS.dbo.TDDSDWB S5 ON S5.TD_CODE = C.TD_CODE AND S5.IT_CODE = 'TK107'
                    LEFT JOIN ( SELECT REPLACE(TD_CODE, '-', '') AS TD_CODE, SUM(ISNULL(COST_FEE, 0) + ISNULL(PPAY_FEE, 0)) AS STCF_FUND FROM CMEPMS.dbo.TDDSTCF WHERE SYR + SMN = '{sPSYR}12' GROUP BY REPLACE(TD_CODE, '-', '') ) F1 ON F1.TD_CODE = M.TD_CODE
                    LEFT JOIN ( SELECT REPLACE(TD_CODE, '-', '') AS TD_CODE, SUM(ISNULL(COST_FEE, 0) + ISNULL(PPAY_FEE, 0)) AS STCF_FUND FROM CMEPMS.dbo.TDDSTCF WHERE SYR + SMN = '{sSYR}{sSMN}' GROUP BY REPLACE(TD_CODE, '-', '') ) F2 ON F2.TD_CODE = M.TD_CODE
                    LEFT JOIN ( SELECT * FROM ( SELECT *, ROW_NUMBER() OVER (PARTITION BY TD_CODE ORDER BY PRE_SDT DESC) AS rn FROM CMETMS.dbo.TDDSDWB ) sub WHERE rn = 1 ) TDD ON TDD.TD_CODE = C.TD_CODE ";

            if (AN_PID == 277 || AN_PID == 260 || AN_PID == 249)
            {
                sql += " where C.PID ='282' or C.PID ='264' or C.PID ='253' ";
                sql += " And (C.PID in (select PID from CMEPMS.dbo.PJMFORM where AN_PID = '249' or AN_PID ='260' or AN_PID ='277'  and SYR in ('" + sPPSYR + "' ,'" + sPSYR + "','" + SYR + "') AND Water_Type = '雨水使用' ) ";
            }
            else
            {
                sql += " Where (C.PID in (select PID from CMEPMS.dbo.PJMFORM where PID = " + AN_PID + " and SYR in ('" + sPSYR + "','" + SYR + "') AND Water_Type = '雨水使用' ) ";
            }

            sql += "And (((isnull(C.YR_FUND,0) + isnull(R.REV_FUND,0))<>0 Or IsNull(C.FCAL_FG,0)=1) or ((isnull(C.YR_FUND,0) + isnull(R.REV_FUND,0))=0 And IsNull(C.FCAL_FG,0)<>1 ) ))";

            /// 20180305 修正 => 當年度經費零工程不要出現 => 2894
            /// 20180507 修正 => 配合本年度無經費, 但有前年度保留款工程, 調整為年度經費 + 保留款 <> 0
            sql += " and (isnull(C.YR_FUND,0) + isnull(R.REV_FUND,0))>0";
            sql += " Order By C.PJV0050CityOrd,C.PJV0050Ord01,C.PJV0050Ord02,C.TD_NAME,C.TID Desc";

            Trace.WriteLine($"SQL : {sql}");

            /// return GetDvOnlyRead(sql, connectionString);
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
            {
                cmd.CommandTimeout = 6000;

                DataTable dt = new DataTable();
                conn.Open();
                da.Fill(dt);

                return dt.DefaultView;
            }
        }

        /// <summary>
        /// I-1.雨水例行性管考報表[前瞻計畫執行情形表]-資料來源
        /// </summary>
        /// <param name="AN_PID"></param>
        /// <param name="PJ_SYR"></param>
        /// <param name="PJ_SMN"></param>
        /// <param name="connectionString"></param>
        /// <param name="DEPT"></param>
        /// <returns></returns>
        public static DataTable GetRPT0003NEWData(decimal AN_PID, string PJ_SYR, string PJ_SMN, string connectionString, string DEPT)
        {
            string sql = "";
            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");
            DataTable dt = new DataTable();

            /// 前一年度
            int PSYR_Count = Convert.ToInt32(PJ_SYR) - 1;
            string PSYR = PSYR_Count.ToString();

            sql = $@"
                    SELECT 
                        A.PID
	                    ,STUFF(STUFF(STUFF(STUFF(CAST(A.TD_CODE AS varchar(20)), 4, 0, '-'), 9, 0, '-'), 14, 0, '-'), 19, 0, '-') AS SIP_TD_CODE
	                    ,A.FD_CODE
	                    ,A.TD_NAME
	                    ,B.CITY
	                    ,CAST( ISNULL(B.CEN_SP_RATIO, 0) AS DECIMAL(18,2)) AS CEN_SP_RATIO
	                    ,ISNULL(FUND_SUM.FUND_TOTAL ,0) AS CEN_FUND
	                    ,ISNULL((SELECT SUM(BAM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = '{PSYR}' AND SMN = '{PJ_SMN}' AND REPLACE(TD_CODE, '-', '') = A.TD_CODE AND A.PID IN ( {AN_PID} ) ), 0) AS PCEN_FUND					-- 分配預算(前一年度中央款)
	                    ,ISNULL((SELECT SUM(AE_CTM + PP_CTM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = '{PSYR}' AND SMN = '{PJ_SMN}' AND REPLACE(TD_CODE, '-', '') = A.TD_CODE AND A.PID IN ({AN_PID} )　), 0) AS LAST_PPAY_FEE	-- 撥款數 = (實支數 + 暫付數)(前一年度)
	                    ,ISNULL((SELECT SUM(AE_CTM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = '{PSYR}' AND SMN = '{PJ_SMN}' AND REPLACE(TD_CODE, '-', '') = A.TD_CODE AND A.PID IN ( {AN_PID} ) ), 0) AS AE_CTM					-- 實支數(前一年度)
                        ,ISNULL((SELECT SUM(BAM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = '{PJ_SYR}' AND SMN = '{PJ_SMN}' AND REPLACE(TD_CODE, '-', '') = A.TD_CODE AND A.PID IN ( {AN_PID} ) ), 0) AS NCEN_FUND					-- 分配預算(今年度中央款)
	                    ,ISNULL((SELECT SUM(AE_CTM + PP_CTM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = '{PJ_SYR}' AND SMN = '{PJ_SMN}' AND REPLACE(TD_CODE, '-', '') = A.TD_CODE AND A.PID IN ({AN_PID} )　), 0) AS NOW_PPAY_FEE	-- 撥款數 = (實支數 + 暫付數)(今年度)
	                    ,ISNULL((SELECT SUM(AE_CTM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = '{PJ_SYR}' AND SMN = '{PJ_SMN}' AND REPLACE(TD_CODE, '-', '') = A.TD_CODE AND A.PID IN ( {AN_PID} ) ), 0) AS NAE_CTM				-- 實支數(今年度)
	                    ,ISNULL(W1.REL_EDT, W1.PRE_EDT) AS TK105_DT		-- 發包後請款期程
	                    ,ISNULL(W5.REL_EDT, W5.PRE_EDT) AS TK141_DT		-- 進度達40％期程
	                    ,ISNULL(W2.PRE_EDT, '') AS TK107_PRE_DT			-- 完工日期
	                    ,ISNULL(W2.REL_EDT, '') AS TK107_REL_DT			-- 完工日期(實際)
	                    ,ISNULL(W4.REL_EDT, W3.REL_EDT) AS TK110_REL_DT	-- 結案
	                    ,ISNULL(MM.PRE_SCHD, 0) AS PRE_SCHD				-- 工程進度(%)月底(預定)
	                    ,ISNULL(MM.REL_SCHD, 0) AS REL_SCHD				-- 工程進度(%)月底(實際)
	                    ,ISNULL(G_B0.BUDGET_FEE, 0) AS BUDGET_FEE_B0	-- 總工程經費
	                    ,ISNULL(G_C0.BUDGET_FEE, 0) AS BUDGET_FEE_C0	-- 契約金額
	                    ,ISNULL(M.EXECT_DESC, '') AS EXECT_DESC			-- 辦理情形
                    FROM CMEPMS.dbo.CTDMNDS A
	                    LEFT JOIN CMEPMS.dbo.TDV_TDDCFTS B ON A.TID = B.TID
	                    LEFT JOIN CMETMS.dbo.TDDSDWB W1 ON W1.TD_CODE = A.TD_CODE AND W1.IT_CODE = 'TK105'
	                    LEFT JOIN CMETMS.dbo.TDDSDWB W5 ON W5.TD_CODE = A.TD_CODE AND W5.IT_CODE = 'TK141'
	                    LEFT JOIN CMETMS.dbo.TDDSDWB W2 ON W2.TD_CODE = A.TD_CODE AND W2.IT_CODE = 'TK107'
	                    LEFT JOIN CMETMS.dbo.TDDSDWB W3 ON W3.TD_CODE = A.TD_CODE AND W3.IT_CODE = 'TK109'
	                    LEFT JOIN CMETMS.dbo.TDDSDWB W4 ON W4.TD_CODE = A.TD_CODE AND W4.IT_CODE = 'TK110'
	                    LEFT JOIN CMETMS.dbo.TDDMSCD MM ON MM.TD_CODE = A.TD_CODE AND MM.SYR = '{PJ_SYR}' and MM.SMN = '{PJ_SMN}'
	                    LEFT JOIN ( SELECT TD_CODE, SUM(BUDGET_FEE) AS BUDGET_FEE FROM CMETMS.dbo.TDDBGLS WHERE REC_TP = 'B' AND SEQ_NO = '0' GROUP BY TD_CODE ) G_B0 ON G_B0.TD_CODE = A.TD_CODE	    -- 總工程經費
	                    LEFT JOIN ( SELECT TD_CODE, SUM(BUDGET_FEE) AS BUDGET_FEE FROM CMETMS.dbo.TDDBGLS WHERE REC_TP = 'C' AND SEQ_NO = '0' GROUP BY TD_CODE ) G_C0 ON G_C0.TD_CODE = A.TD_CODE	    -- 契約金額
	                    LEFT JOIN ( SELECT TD_CODE, EXECT_DESC, PRE_SCHD, REL_SCHD FROM CMETMS.dbo.TDDMSCD WHERE SYR = '{PJ_SYR}' AND SMN = RIGHT('00' + '{PJ_SMN}', 2) ) M ON M.TD_CODE = A.TD_CODE    -- 辦理情形
	                    LEFT JOIN ( SELECT TID, SUM(FUND) AS FUND_TOTAL FROM CMEPMS.dbo.CTDITFD WHERE FIT_CODE = '00' GROUP BY TID ) FUND_SUM ON FUND_SUM.TID = A.TID                                   -- 核定補助經費(中央款) | 分配預算
                    WHERE A.PID IN ({AN_PID}) ";

            Trace.WriteLine("I-1.雨水例行性管考報表:");
            Trace.WriteLine(sql);

            dt = SourceFunction.DataTableGet(sql, connectionString);

            return dt;
        }

        /// <summary>
        /// A-2.細項執行情形表-資料來源
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="SYR"></param>
        /// <param name="SMN"></param>
        /// <param name="CityList"></param>
        /// <param name="funding"></param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        public static DataTable getBillionProject08_NEW(string PID, string SYR, string SMN, string[] CityList, string funding, string connectionString)
        {
            int seqNo = 0;
            DataTable ResultTable = new DataTable();

            /// 資料來源:SQL查詢
            #region =================================================================================================
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                /// Step 1.取得案件期別
                /// SEQ_NO : 期別
                string SearchSql1 = $"SELECT MAX(SEQ_NO) AS SEQ_NO FROM CMEPMS.dbo.PJDFDCR WHERE PID IN ({PID})";
                using (SqlCommand cmd1 = new SqlCommand(SearchSql1, con))
                {
                    object result1 = cmd1.ExecuteScalar();

                    if (result1 != null && result1 != DBNull.Value)
                        seqNo = Convert.ToInt32(result1);
                }

                /// Step 2.判斷是否為核定狀態(暫不判定 true/false)
                /// APROV_DT : 核定日期
                /// APROV_FN : 核定文號
                bool isApproved = true;
                string SearchSql2 = $"SELECT (APROV_DT + APROV_FN) AS APROV FROM CMEPMS.dbo.PJDFDCR WHERE PID IN ({PID}) AND SEQ_NO = {seqNo}";

                using (SqlCommand cmd2 = new SqlCommand(SearchSql2, con))
                {
                    object resultApprove = cmd2.ExecuteScalar();

                    if (resultApprove != null && resultApprove != DBNull.Value)
                    {
                        string aprovStr = resultApprove.ToString().Trim();
                        if (string.IsNullOrEmpty(aprovStr))
                            isApproved = false;
                    }
                }

                /// Step 3.主要資料來源
                /// funding : 判定移除經費為空或 0 的列 (若 FUND0 != "Y")
                string SearchSql3 = "";
                string citylist = "";
                string PSYR = (Convert.ToInt32(SYR) - 1).ToString();

                var safeCities = CityList
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Select(s => $"N'{s.Trim().Replace("'", "''")}'");
                citylist = $"({string.Join(",", safeCities)})";

                if (isApproved == true)
                {
                    #region 新版程式碼[]

                    SearchSql3 = $@"
                                  -- A-2.細項執行情形表
                                  SELECT
                                    A.PID
                                    ,C.PID_SYR
                                    ,C.SEQ_NO
                                    ,A.TD_CODE
                                    ,A.FD_CODE
                                    ,B.CATY_CODE
                                    ,D.TD_CITY -- 縣市
                                    ,C.CITY
                                    ,C.TOWN
                                    ,S9.CODE_DESC AS IPJ_DESC -- 辦理層級
                                    ,CAST(ROUND(ISNULL(E.CEN_SP_RATIO, 0), 0) AS INT) AS CEN_RATIO -- 補助比例
                                    ,B.CATY_NAME -- 計畫說明
                                    ,A.TD_NAME   -- 工作名稱
                                    ,IIF(d.TD_FINISH = 1, N'已結案', IIF(d.TD_FINISH = 0, N'未結案', N'未填報')) AS TD_FINISH_DESC -- 結案情形
                                    ,ISNULL((SELECT SUM(AE_CTM + PP_CTM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = {PSYR} AND SMN = '12' AND FD_CODE = A.FD_CODE AND b.PID IN ( {PID} ) ), 0) AS LAST_PPAY_FEE -- 當年累積撥發金額(仟元)
                                    ,ISNULL((SELECT SUM(AE_CTM + PP_CTM) FROM CMEPMS.dbo.TDDCTET WHERE SYR = {SYR} AND SMN = {SMN} AND FD_CODE = A.FD_CODE AND b.PID IN ( {PID} ) ), 0) AS ALL_PPAY_FEE -- 當年累積撥發金額(仟元)
                                    ,D.START_FUND
                                    ,B.CATY_FUND -- 分支計畫核定經費總計(仟元)
                                    ,( ( SELECT SUM(BUDGET_FEE) FROM CMETMS.dbo.TDDBGLS WHERE REC_TP = 'B' AND TD_CODE = A.TD_CODE AND SEQ_NO = ( SELECT MAX(SEQ_NO) FROM CMETMS.dbo.TDDBGRC WHERE TD_CODE = CMETMS.dbo.TDDBGLS.TD_CODE AND REC_TP = 'B' ) ) * 0.95 * ISNULL(m.REL_SCHD, 0) / 100.0 * CAST(ROUND(ISNULL(E.CEN_SP_RATIO, 0), 0) AS INT) / 100.0 ) AS L_FUND -- 依進度可撥金費(李)
                                    ,(SELECT SUM(FUND) FROM CMEPMS.dbo.CTDITFD WHERE TID = A.TID AND FIT_CODE = '00') AS CEN_FUND
                                    ,(SELECT SUM(BUDGET_FEE) FROM CMETMS.dbo.TDDBGLS WHERE REC_TP = 'B' AND TD_CODE = A.TD_CODE AND SEQ_NO = (SELECT MAX(SEQ_NO) FROM CMETMS.dbo.TDDBGRC WHERE TD_CODE = CMETMS.dbo.TDDBGLS.TD_CODE AND REC_TP = 'B')) AS B_FUND
                                    -----------------------------------------------------------------------------------------------------------------------
                                    ,S1.PRE_EDT AS PRE_TK103  -- 預定細設完成日期 
                                    ,S1.REL_EDT AS REL_TK103  -- 實際細設完成日期
                                    ,S2.PRE_EDT AS PRE_TK105  -- 預定發包日期
                                    ,S2.REL_EDT AS REL_TK105  -- 實際發包日期
                                    ,S6.REL_EDT AS REL_TK102  -- 基設核定日期
                                    ,S10.PRE_EDT AS PRE_TK108 -- 預定正驗日期
                                    ,S10.REL_EDT AS REL_TK108 -- 實際正驗日期
                                    ,ISNULL(M.PRE_SCHD, 0) AS PRE_SCHD -- 預定工程進度
                                    ,ISNULL(M.REL_SCHD, 0) AS REL_SCHD -- 實際工程進度
                                    ,(ISNULL(M.REL_SCHD, 0) - ISNULL(M.PRE_SCHD, 0)) AS SCHD_DIFF -- 進度差異
                                    ,CASE WHEN D.CATY_CODE = 'P' AND S3.PRE_EDT IS NOT NULL THEN S3.PRE_EDT ELSE S4.PRE_EDT END AS PRE_TK107 -- 預定完工日期
                                    ,CASE WHEN D.CATY_CODE = 'P' AND S3.REL_EDT IS NOT NULL THEN S3.REL_EDT ELSE S4.REL_EDT END AS REL_TK107 -- 實際完工日期
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
                                LEFT JOIN CMETMS.dbo.SMCCODE S9 ON S9.SYS_ID = 'SM' AND S9.CATY_ID = '02' AND S9.CATY_COD = C.IPJ_CODE
                                LEFT JOIN ( SELECT TD_CODE, EXECT_DESC, PRE_SCHD, REL_SCHD, DELAY_RESN, CTDELAY_RESN FROM CMETMS.dbo.TDDMSCD WHERE SYR = '{SYR}' AND SMN = '{SMN}') M ON M.TD_CODE = A.TD_CODE
                                LEFT JOIN ( SELECT TD_CODE, EXECT_DESC, PRE_SCHD, REL_SCHD, DELAY_RESN, CTDELAY_RESN FROM CMETMS.dbo.TDDMSCD WHERE SYR = '{SYR}' AND SMN = '{SMN}') O ON O.TD_CODE = A.TD_CODE
                                WHERE A.PID IN ( {PID} ) ";

                    #endregion

                }

                Trace.WriteLine("SearchSql3:");
                Trace.WriteLine(SearchSql3);

                using (SqlDataAdapter cmd3 = new SqlDataAdapter(SearchSql3, con))
                {
                    cmd3.Fill(ResultTable);
                }

            }

            #endregion ==============================================================================================

            #region 測試 CODE
            Trace.WriteLine("ResultTable:");

            foreach (DataRow rows in ResultTable.Rows)
            {
                foreach (DataColumn cols in ResultTable.Columns)
                {
                    Trace.Write($"{cols.ColumnName}:{rows[cols]} \t");
                }
                Trace.WriteLine("");
            }
            #endregion

            return ResultTable;
        }

        /// <summary>
        /// B-1.修正預算管考報表_總表資料來源
        /// </summary>
        /// <param name="PID"></param>
        /// <param name="SYR"></param>
        /// <param name="SEQNO"></param>
        /// <param name="Connect"></param>
        /// <returns></returns>
        public static DataTable B1RERPORTDATA(string PID, string SYR, string SEQNO, string Connect)
        {
            string sql = "";
            string sSYR = (System.DateTime.Now.Year - 1911).ToString("000");
            string sSMN = (System.DateTime.Now.Month).ToString("00");
            string sPSYR = (System.DateTime.Now.Year - 1911 - 1).ToString("000");

            int SEQ = Convert.ToInt32(SEQNO);
            int PSEQ = SEQ - 1;

            DataTable DT = new DataTable();

            /// 原預算
            if (SEQ == 0)
            {
                /// 以下 SQL 對應 "前瞻計畫基礎建設" 01 / 02 / 03
                sql = $@"
                        SELECT
	                        A.TID
                            ,A.PID
	                        ,D.PID_SYR
	                        ,B.SEQ_NO
	                        ,A.TD_CODE
	                        ,D.FD_CODE
	                        ,D.SIP_TD_CODE
	                        ,D.CITY
	                        ,D.TOWN
	                        ,A.TD_NAME
	                        ,D.CATY_TYPE
	                        ,D.BOT_FG
	                        ,D.ACC_ID
	                        ,D.ACC_DESC
                            ,ISNULL(FUND_SUM.FUND_TOTAL / 1000 ,0) AS CEN_FUND
	                        ,D.CATY_CODE
	                        ,D.IPJ_CODE
                            ,D.IPJ_DESC
	                        ,D.IPJ_CODE_DESC
                            ,D.CEN_SP_RATIO
                            ,D.LOC_FUND
                            ,ISNULL(M.EXECT_DESC,'') AS EXECT_DESC
	                        ,ISNULL(D.M_01,0) AS M_01
	                        ,ISNULL(D.M_02,0) AS M_02
	                        ,ISNULL(D.M_03,0) AS M_03
	                        ,ISNULL(D.M_04,0) AS M_04
	                        ,ISNULL(D.M_05,0) AS M_05
	                        ,ISNULL(D.M_06,0) AS M_06
	                        ,ISNULL(D.M_07,0) AS M_07
	                        ,ISNULL(D.M_08,0) AS M_08
	                        ,ISNULL(D.M_09,0) AS M_09
	                        ,ISNULL(D.M_10,0) AS M_10
	                        ,ISNULL(D.M_11,0) AS M_11
	                        ,ISNULL(D.M_12,0) AS M_12
                        FROM CMEPMS.dbo.CTDMNDS A
                        LEFT JOIN CMEPMS.dbo.PJDFDCR B ON B.PID = A.PID
                        LEFT JOIN CMEPMS.dbo.CTHMNDS C ON C.TID = A.TID AND C.PID = A.PID AND C.SEQ_NO = B.SEQ_NO
                        LEFT JOIN CMEPMS.dbo.TDV_TDDCFTS D ON D.PID  = A.PID AND D.TID = A.TID
                        LEFT JOIN ( SELECT TID, SUM(FUND) AS FUND_TOTAL FROM CMEPMS.dbo.CTDITFD WHERE FIT_CODE = '00' GROUP BY TID ) FUND_SUM ON FUND_SUM.TID = A.TID
                        LEFT JOIN CMETMS.dbo.TDDMSCD M ON M.TD_CODE = A.TD_CODE AND M.SYR = '{SYR}' AND M.SMN = ( SELECT MAX(M2.SMN) FROM CMETMS.dbo.TDDMSCD M2 WHERE M2.TD_CODE = A.TD_CODE AND M2.SYR = '{SYR}')
                        WHERE A.PID IN ({PID}) AND B.SEQ_NO IN ('{SEQ}') ORDER BY D.SEQ_NO, D.CITY ";

                DT = SourceFunction.DataTableGet(sql, Connect);

                if (DT.Rows.Count == 0)
                {
                    sql = $@"
                        SELECT
	                        A.TID
                            ,A.PID
	                        ,D.PID_SYR
	                        ,D.SEQ_NO
	                        ,A.TD_CODE
	                        ,D.FD_CODE
	                        ,D.SIP_TD_CODE
	                        ,D.CITY
	                        ,D.TOWN
	                        ,A.TD_NAME
	                        ,D.CATY_TYPE
	                        ,D.BOT_FG
	                        ,D.ACC_ID
	                        ,D.ACC_DESC
                            ,ISNULL(FUND_SUM.FUND_TOTAL / 1000 ,0) AS CEN_FUND
	                        ,D.CATY_CODE
	                        ,D.IPJ_CODE
                            ,D.IPJ_DESC
	                        ,D.IPJ_CODE_DESC
                            ,D.CEN_SP_RATIO
                            ,D.LOC_FUND
                            ,ISNULL(M.EXECT_DESC,'') AS EXECT_DESC
	                        ,ISNULL(D.M_01,0) AS M_01
	                        ,ISNULL(D.M_02,0) AS M_02
	                        ,ISNULL(D.M_03,0) AS M_03
	                        ,ISNULL(D.M_04,0) AS M_04
	                        ,ISNULL(D.M_05,0) AS M_05
	                        ,ISNULL(D.M_06,0) AS M_06
	                        ,ISNULL(D.M_07,0) AS M_07
	                        ,ISNULL(D.M_08,0) AS M_08
	                        ,ISNULL(D.M_09,0) AS M_09
	                        ,ISNULL(D.M_10,0) AS M_10
	                        ,ISNULL(D.M_11,0) AS M_11
	                        ,ISNULL(D.M_12,0) AS M_12
                        FROM CMEPMS.dbo.CTDMNDS A
                        LEFT JOIN CMEPMS.dbo.PJDFDCR B ON B.PID = A.PID
                        LEFT JOIN CMEPMS.dbo.CTHMNDS C ON C.TID = A.TID AND C.PID = A.PID AND C.SEQ_NO = B.SEQ_NO
                        LEFT JOIN CMEPMS.dbo.TDV_TDDCFTS D ON D.PID  = A.PID AND D.TID = A.TID
                        LEFT JOIN ( SELECT TID, SUM(FUND) AS FUND_TOTAL FROM CMEPMS.dbo.CTDITFD WHERE FIT_CODE = '00' GROUP BY TID ) FUND_SUM ON FUND_SUM.TID = A.TID
                        LEFT JOIN CMETMS.dbo.TDDMSCD M ON M.TD_CODE = A.TD_CODE AND M.SYR = '{SYR}' AND M.SMN = ( SELECT MAX(M2.SMN) FROM CMETMS.dbo.TDDMSCD M2 WHERE M2.TD_CODE = A.TD_CODE AND M2.SYR = '{SYR}')
                        WHERE A.PID IN ({PID}) AND D.SEQ_NO IN ('{SEQ}') ORDER BY D.SEQ_NO, D.CITY ";

                    DT = SourceFunction.DataTableGet(sql, Connect);
                }
            }
            /// 預算調整(SEQ_NO)
            else
            {
                /// 以下 SQL 對應 "前瞻計畫基礎建設" 01 / 02 / 03
                sql = $@"
                        SELECT 
	                        A.TID
                            ,A.PID
	                        ,D.PID_SYR
	                        ,B.SEQ_NO
	                        ,A.TD_CODE
	                        ,D.FD_CODE
	                        ,D.SIP_TD_CODE
	                        ,D.CITY
	                        ,D.TOWN
	                        ,A.TD_NAME
	                        ,D.CATY_TYPE
	                        ,D.BOT_FG
	                        ,D.ACC_ID
	                        ,D.ACC_DESC
                            ,ISNULL(FUND_SUM.FUND_TOTAL / 1000 ,0) AS CEN_FUND
	                        ,D.CATY_CODE
	                        ,D.IPJ_CODE
                            ,D.IPJ_DESC
	                        ,D.IPJ_CODE_DESC
                            ,D.CEN_SP_RATIO
                            ,D.LOC_FUND
                            ,ISNULL(M.EXECT_DESC,'') AS EXECT_DESC
	                        ,ISNULL(D.M_01,0) AS M_01
	                        ,ISNULL(D.M_02,0) AS M_02
	                        ,ISNULL(D.M_03,0) AS M_03
	                        ,ISNULL(D.M_04,0) AS M_04
	                        ,ISNULL(D.M_05,0) AS M_05
	                        ,ISNULL(D.M_06,0) AS M_06
	                        ,ISNULL(D.M_07,0) AS M_07
	                        ,ISNULL(D.M_08,0) AS M_08
	                        ,ISNULL(D.M_09,0) AS M_09
	                        ,ISNULL(D.M_10,0) AS M_10
	                        ,ISNULL(D.M_11,0) AS M_11
	                        ,ISNULL(D.M_12,0) AS M_12
                        FROM CMEPMS.dbo.CTDMNDS A
                        LEFT JOIN CMEPMS.dbo.PJDFDCR B ON B.PID = A.PID
                        LEFT JOIN CMEPMS.dbo.CTHMNDS C ON C.TID = A.TID AND C.PID = A.PID AND C.SEQ_NO = B.SEQ_NO
                        LEFT JOIN CMEPMS.dbo.TDV_TDDCFTS D ON D.PID  = A.PID AND D.TID = A.TID
                        LEFT JOIN ( SELECT TID, SUM(FUND) AS FUND_TOTAL FROM CMEPMS.dbo.CTDITFD WHERE FIT_CODE = '00' GROUP BY TID ) FUND_SUM ON FUND_SUM.TID = A.TID
                        LEFT JOIN CMETMS.dbo.TDDMSCD M ON M.TD_CODE = A.TD_CODE AND M.SYR = '{SYR}' AND M.SMN = ( SELECT MAX(M2.SMN) FROM CMETMS.dbo.TDDMSCD M2 WHERE M2.TD_CODE = A.TD_CODE AND M2.SYR = '{SYR}')
                        WHERE A.PID IN ({PID}) AND B.SEQ_NO IN ('{PSEQ}','{SEQ}') ORDER BY D.SEQ_NO, D.CITY ";

                DT = SourceFunction.DataTableGet(sql, Connect);

                if (DT.Rows.Count == 0)
                {
                    sql = $@"
                        SELECT
	                        A.TID
                            ,A.PID
	                        ,D.PID_SYR
	                        ,D.SEQ_NO
	                        ,A.TD_CODE
	                        ,D.FD_CODE
	                        ,D.SIP_TD_CODE
	                        ,D.CITY
	                        ,D.TOWN
	                        ,A.TD_NAME
	                        ,D.CATY_TYPE
	                        ,D.BOT_FG
	                        ,D.ACC_ID
	                        ,D.ACC_DESC
                            ,ISNULL(FUND_SUM.FUND_TOTAL / 1000 ,0) AS CEN_FUND
                            -- 20260327 新增
	                        ,D.CATY_CODE
	                        ,D.IPJ_CODE
                            ,D.IPJ_DESC
	                        ,D.IPJ_CODE_DESC
                            ,D.CEN_SP_RATIO
                            ,D.LOC_FUND
                            ,ISNULL(M.EXECT_DESC,'') AS EXECT_DESC
	                        ,ISNULL(D.M_01,0) AS M_01
	                        ,ISNULL(D.M_02,0) AS M_02
	                        ,ISNULL(D.M_03,0) AS M_03
	                        ,ISNULL(D.M_04,0) AS M_04
	                        ,ISNULL(D.M_05,0) AS M_05
	                        ,ISNULL(D.M_06,0) AS M_06
	                        ,ISNULL(D.M_07,0) AS M_07
	                        ,ISNULL(D.M_08,0) AS M_08
	                        ,ISNULL(D.M_09,0) AS M_09
	                        ,ISNULL(D.M_10,0) AS M_10
	                        ,ISNULL(D.M_11,0) AS M_11
	                        ,ISNULL(D.M_12,0) AS M_12
	                        -- 20260327 新增
                        FROM CMEPMS.dbo.CTDMNDS A
                        LEFT JOIN CMEPMS.dbo.PJDFDCR B ON B.PID = A.PID
                        LEFT JOIN CMEPMS.dbo.CTHMNDS C ON C.TID = A.TID AND C.PID = A.PID AND C.SEQ_NO = B.SEQ_NO
                        LEFT JOIN CMEPMS.dbo.TDV_TDDCFTS D ON D.PID  = A.PID AND D.TID = A.TID
                        LEFT JOIN ( SELECT TID, SUM(FUND) AS FUND_TOTAL FROM CMEPMS.dbo.CTDITFD WHERE FIT_CODE = '00' GROUP BY TID ) FUND_SUM ON FUND_SUM.TID = A.TID
                        LEFT JOIN CMETMS.dbo.TDDMSCD M ON M.TD_CODE = A.TD_CODE AND M.SYR = '{SYR}' AND M.SMN = ( SELECT MAX(M2.SMN) FROM CMETMS.dbo.TDDMSCD M2 WHERE M2.TD_CODE = A.TD_CODE AND M2.SYR = '{SYR}')
                        WHERE A.PID IN ({PID}) AND D.SEQ_NO IN ('{PSEQ}','{SEQ}') ORDER BY D.SEQ_NO, D.CITY ";

                    DT = SourceFunction.DataTableGet(sql, Connect);
                }

            }

            Trace.WriteLine("B-1.修正預算管考報表");
            Trace.WriteLine(sql);

            return DT;
        }

        #endregion ====================================================================================================
    }
}