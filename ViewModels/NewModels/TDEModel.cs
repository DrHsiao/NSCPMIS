using Aspose.Cells;
using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.EnterpriseServices.Internal;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WCals.Models.CMEDOCModels;
using WCals.Models.CMETMSModels;


namespace WCals.ViewModels.NewModels {
    public class TDEModel {
        public int page1 { get; set; }
        public string TD_CODE { get; set; } = string.Empty;
        public string REC_TP { get; set; } = string.Empty;
        public string TD_NAME { get; set; } = string.Empty;
        public string CITY { get; set; } = string.Empty;
        public string TOWN { get; set; } = string.Empty;
        public string IPJ_CODE { get; set; } = string.Empty;
        public string CATY_CODE { get; set; } = string.Empty;
        public short SEQ_NO { get; set; }
        public string TID { get; set; } = string.Empty;
        public IPagedList<DataRow> List1 { get; set; }
    }

    public class TDE0110Model:TDEModel{

    }

    /// <summary>
    /// TDE0051 工程進度監控 - 5個頁籤
    /// </summary>
    public class TDE0051Model : TDEModel {
        /// <summary>
        /// 年度計畫查詢條件 (0 = 不限)
        /// </summary>
        public decimal PID { get; set; }

        /// <summary>
        /// 當前頁籤 (1-5)
        /// </summary>
        public int tab { get; set; } = 1;

        // 各頁籤的分頁頁碼
        public int page2 { get; set; } = 1;
        public int page3 { get; set; } = 1;
        public int page4 { get; set; } = 1;
        public int page5 { get; set; } = 1;

        /// <summary>
        /// 頁籤1: 開/完工日填報不全
        /// SQL: select SIP_TD_CODE as S1_HLK,* from TDV_TDDCNYM
        /// </summary>
        public IPagedList<DataRow> List_TDDCNYM { get; set; }

        /// <summary>
        /// 頁籤2: 進度未填(前月)
        /// SQL: select SIP_TD_CODE as S1_HLK,* from TDV_TDDCNSD
        /// </summary>
        public IPagedList<DataRow> List_TDDCNSD { get; set; }

        /// <summary>
        /// 頁籤3: 進度異常(前月)
        /// SQL: select SIP_TD_CODE as S1_HLK,* from TDV_TDDCN
        /// </summary>
        public IPagedList<DataRow> List_TDDCN { get; set; }

        /// <summary>
        /// 頁籤4: 工程相關經費未填
        /// SQL: select SIP_TD_CODE as S1_HLK,* from TDV_TDDCNFD
        /// </summary>
        public IPagedList<DataRow> List_TDDCNFD { get; set; }

        /// <summary>
        /// 頁籤5: 連續三個月落後
        /// SQL: select SIP_TD_CODE as S1_HLK,* from CMETMS.dbo.TDV_TDDMSCD
        /// </summary>
        public IPagedList<DataRow> List_TDDMSCD { get; set; }
    }

    public class TDE0040Model:TDEModel{

        public string PID { get; set; } = string.Empty;
        public string SYR { get; set; } = string.Empty;
        public string TD_STAT_DESC { get; set; } = string.Empty;
        public string SYS_ID_FG { get; set; } = string.Empty;
        public string PJ_ID_FG { get; set; } = string.Empty;
        public string SYS_PJ_FG { get; set; } = string.Empty;
        

        public TDMFORM TDMFORM { get; set; }

        public TDV_TDMFORM TDVFORM { get; set; }
        public TDDDCOB TDDDCOB { get; set; }
        
        public TDDBGRC TDDBGRC { get; set; }
        
        public TDDBDRC TDDBDRC { get; set; }

        public TDDBGLS TDDBGLS { get; set; }
    }

    public class TDE0020Model: TDEModel {

        public IEnumerable<SelectListItem> pids { get; set; }
        public string pid { get; set; }

        public string city { get; set; }
        public string town { get; set; }
        public string year { get; set; }
        public string month { get; set; }
        public string td_code { get; set; }
        public string td_name { get; set; }
        public string type { get; set; }

        public string[] tdcodes { get; set; }
        public string[] years { get; set; }
        public string[] months { get; set; }
        public string[] TK101  { get; set; }
        public string[] TK102  { get; set; }
        public string[] TK103  { get; set; }
        public string[] TK1031 { get; set; }
        public string[] TK104  { get; set; }
        public string[] TK1041 { get; set; }
        public string[] TK105  { get; set; }
        public string[] TK106  { get; set; }
        public string[] TK107  { get; set; }
        public string[] TK1071 { get; set; }
        public string[] TK108  { get; set; }
        public string[] TK109  { get; set; }
        public string[] TK110  { get; set; }
        public string[] TK140  { get; set; }
        public string[] TK141 { get; set; }
    }

    /// <summary>
    /// TDE0070 污水下水道普及率統計 - 2個頁籤
    /// </summary>
    public class TDE0070Model : TDEModel {
        /// <summary>
        /// 當前頁籤 (1-2)
        /// </summary>
        public int tab { get; set; } = 1;

        /// <summary>
        /// 年度 (民國年 096-114)
        /// </summary>
        public string SYR { get; set; } = string.Empty;

        /// <summary>
        /// 月份 (01-12)
        /// </summary>
        public string SMN { get; set; } = string.Empty;

        /// <summary>
        /// 頁籤2分頁
        /// </summary>
        public int page2 { get; set; } = 1;

        /// <summary>
        /// 頁籤1: 統計資料列表
        /// </summary>
        public IPagedList<DataRow> List_Tab1 { get; set; }

        /// <summary>
        /// 頁籤2: 統計資料列表
        /// </summary>
        public IPagedList<DataRow> List_Tab2 { get; set; }
    }

    /// <summary>
    /// TDE0130 用戶接管及化糞池拆除統計 - 2個頁籤
    /// </summary>
    public class TDE0130Model : TDEModel {
        /// <summary>
        /// 當前頁籤 (1-2)
        /// </summary>
        public int tab { get; set; } = 1;

        /// <summary>
        /// 年度
        /// </summary>
        public string SYR { get; set; } = string.Empty;

        /// <summary>
        /// 月份
        /// </summary>
        public string SMN { get; set; } = string.Empty;

        /// <summary>
        /// 縣市
        /// </summary>
        public string TD_CITY { get; set; } = string.Empty;

        /// <summary>
        /// 頁籤2分頁
        /// </summary>
        public int page2 { get; set; } = 1;

        /// <summary>
        /// 頁籤1: 用戶接管
        /// </summary>
        public IPagedList<DataRow> List_Tab1 { get; set; }

        /// <summary>
        /// 頁籤2: 化糞池拆除
        /// </summary>
        public IPagedList<DataRow> List_Tab2 { get; set; }
    }
}