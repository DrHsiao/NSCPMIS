using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models.CMEPMSModels
{
    public class CTDMNDSmergeTDV_TDDCFTS
    {
        public string SYR { get; set; }
        public string PJ_NAME { get; set; }
        public decimal PID { get; set; }
        public string TD_CITY { get; set; }
        public string TD_TOWN { get; set; }
        public string TD_CODE { get; set; }
        public string IPJ_CODE { get; set; }
        public decimal TOTAL_ALL_BUDGET_FEE { get; set; }
        public string ADV_DP { get; set; }
        public string IPJ_DESC { get; set; }
        public string TD_NOTE { get; set; }
        public string SMP_DEPT_NAME { get ; set; }
        public decimal TOTAL_FUND { get; set; }
        public decimal FUND { get; set; }
        public decimal PERCENT_FUND { get; set; }
        public decimal TID { get; set; }
        public string TD_NAME { get; set; }
        public string TD_LINE { get; set; }
        public decimal COST_FEE { get; set; }
        public decimal COST_FEE_TOTAL { get; set; }
        public string PRE_CONTRACT_DATE { get; set; }
        public string REL_CONTRACT_DATE { get; set; }
        public string PRE_START_DATE { get; set; }
        public string REL_START_DATE { get; set; }
        public string PRE_END_DATE { get; set;}
        public string REL_END_DATE { get; set; }
        public decimal PRE_SCHD { get ; set; }
        public decimal REL_SCHD { get;set; }
        public string EXECT_DESC { get; set; }
        public string DELAY_RESN { get; set; }
        public string CTDELAY_RESN { get; set; }
        public decimal PROGRESS_DIFFERENCE { get; set; }
        public string TD_MEMO { get; set; }

        /// 20241022 修改項目
        public string PRE_TK103_DATE { get; set; }  /// 預定(細部)設計完成日期

        public string REL_TK103_DATE { get; set; }  /// 實際(細部)設計完成日期

        public string ALL_OUTSRC_FG { get; set; }   /// 已發包的預算款項名稱

        public string DSG_BUSS_NAME { get; set; }   /// 設計單位

        public string SVY_BUSS_NAME { get; set; }   /// 監造單位

        public string BUSS_NAME { get; set; }       /// 廠商名稱

        public string FINISH_FG { get; set; }       /// 已結案

        public string START_FUND { get; set; }      /// 基設核定經費(千元)

        public string HOUSEHOLD { get; set; }       /// 計畫接管戶數

        public string TOTAL_HOUSEHOLD { get; set; } /// 實際累計接管戶數

        public string WTR_PL_NEW_LEN { get; set; }  /// [雨水用]雨水管線新(改)建長度(m)

        public string WTR_PL_REP_LEN { get; set; }  /// [雨水用]雨水管線修繕長度(m)

        public string WTR_FLOOD { get; set; }       /// [雨水用]滯洪量(m3)

        public string WTR_TOTAL_PUMP { get; set; }  /// [雨水用]總抽水量(CMS)

        public string WTR_GU_STA { get; set; }      /// [雨水用]水位計站數

        public string REL_TK108_DATE { get; set; }  /// 實際驗收完成日期

        public string FD_CODE { get; set; }         /// 經費編號
    }
}