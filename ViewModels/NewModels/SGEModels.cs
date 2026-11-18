using PagedList;
using System;
using System.Data;
using System.Web;
using WCals.Models.CMPSGSModels;

namespace WCals.ViewModels.NewModels
{
    public class SGEModel
    {
        public string SEQ_NO { get; set; } = string.Empty;
        public string AP_STAT { get; set; }
        public string YRC_ID { get; set; }
        public string SYR { get; set; }
        public string END_SYR { get; set; }
        public string PJ_NAME { get; set; }
        public string PSG_NAME { get; set; }
        public string WK_CODE { get; set; }
        public string PJ_FUND { get; set; }
        public string ANPID { get; set; }
        public decimal PPID { get; set; }
        public IPagedList<DataRow> MainList { get; set; }
        public IPagedList<DataRow> List1 { get; set; }
        public IPagedList<DataRow> List2 { get; set; }
        public IPagedList<DataRow> List3 { get; set; }
        public int page { get; set; }
        public string keyword { get; set; }
        public PJMFORM_CMPSGS PJMFORM { get; set; }
        public decimal PID { get; set; }
        public string YRC { get; set; }
        /// <summary>
        /// 工程名稱搜尋條件
        /// </summary>
        public string keyword2 { get; set; }
    }



    public class SGE0010Model : SGEModel
    {

        public string AID { get; set; }
        public int tab { get; set; }

        public int page1 { get; set; }
        public int page2 { get; set; }

        public string city { get; set; }
        public APMFORM apmform { get; set; }

        public string FN_DT { get; set; }
        public string OF_DT { get; set; }

        public string[] DCS_FUND { get; set; }
        public string[] JUD_FUND { get; set; }
        public decimal[] FID { get; set; }
        public string[] FUND_DESC { get; set; }
    }

    public class SGE0000Model : SGEModel
    {
        public string seqno { get; set; } = string.Empty;
        public string AID { get; set; } = string.Empty;
        public string[] DEL_FG { get; set; } = Array.Empty<string>();
        public APMFORM APMFORM { get; set; }
        public string EXT_FG { get; set; } = string.Empty;
        public string START_DT { get; set; } = string.Empty;
        public string END_DT { get; set; } = string.Empty;
        public string CITY { get; set; } = string.Empty;
        public string SWW1_FG { get; set; } = string.Empty;
        public string SWW2_FG { get; set; } = string.Empty;
        public string SWW3_FG { get; set; } = string.Empty;
        public string SWW4_FG { get; set; } = string.Empty;
        public string SWW5_FG { get; set; } = string.Empty;
        public string SWW6_FG { get; set; } = string.Empty;
        public string SWW7_FG { get; set; } = string.Empty;
        public string SWW8_FG { get; set; } = string.Empty;
        public string SWW81_FG { get; set; } = string.Empty;
        public string SWW9_FG { get; set; } = string.Empty;
        public string SWW10_FG { get; set; } = string.Empty;
        public string SWW11_FG { get; set; } = string.Empty;
        public string SWWOT_DESC { get; set; } = string.Empty;
        public string TF106_DT { get; set; } = string.Empty;
        public string TF108_DT { get; set; } = string.Empty;
        public string TK103_DT { get; set; } = string.Empty;
        public string TK105_DT { get; set; } = string.Empty;
        public string TK106_DT { get; set; } = string.Empty;
        public string TK107_DT { get; set; } = string.Empty;
        public string SYS_CATY { get; set; } = string.Empty;
        public string IPJ_CODE { get; set; } = string.Empty;

        //public APDFUND[] APDFUND { get; set; }
        public string[] APDFUND_FID { get; set; } = Array.Empty<string>();
        public string[] SYRS { get; set; } = Array.Empty<string>();
        public string[] FIT_CODE { get; set; } = Array.Empty<string>();
        public string[] FUND { get; set; } = Array.Empty<string>();
        public string[] FUND_DESC { get; set; } = Array.Empty<string>();

        //public APDFILE[] APDFILE { get; set; }
        public HttpPostedFileBase[] files { get; set; }
        public string[] FC_DESC { set; get; } = Array.Empty<string>();
        public string[] FILE_MEMO { get; set; } = Array.Empty<string>();
        //public string[] EXT_FG { get; set; } = Array.Empty<string>();
        public string[] FID { get; set; } = Array.Empty<string>();

        //public APDLOCP[] APDLOCP { get; set; }        
        public string[] LID { get; set; } = Array.Empty<string>();
        public string[] TOWN { get; set; } = Array.Empty<string>();
        public string[] POS_SYB { get; set; } = Array.Empty<string>();
        public string[] GIS_X { get; set; } = Array.Empty<string>();
        public string[] GIS_Y { get; set; } = Array.Empty<string>();
        public string[] GIS_Z { get; set; } = Array.Empty<string>();
        public string[] GIS_SYS { get; set; } = Array.Empty<string>();


        //public APDWQTY[] APDWQTY { get; set; }
        public string[] MT_LID { get; set; } = Array.Empty<string>();
        public string[] MT_MID { get; set; } = Array.Empty<string>();
        public string[] MT_DESC { get; set; } = Array.Empty<string>();
        public string[] MT_NAME { get; set; } = Array.Empty<string>();
        public string[] MT_UNIT { get; set; } = Array.Empty<string>();
        public string[] MT_QTY { get; set; } = Array.Empty<string>();
        public string[] MT_SPEC { get; set; } = Array.Empty<string>();


    }


}