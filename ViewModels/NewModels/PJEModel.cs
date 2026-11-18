using PagedList;
using System;
using System.Data;
using System.Web;
using WCals.Models.CMEPMSModels;
using WCals.Models.CMPSGSModels;

namespace WCals.ViewModels.NewModels
{
    public class PJEModel
    {
        /// <summary>
        /// 目前所在的主計畫名稱(+分支計畫名稱)
        /// </summary>
        public string CurrentProjectName { get; set; }

        //public IPagedList<DataRow> PJV_PJMFORM { get; set; }
        public IPagedList<PJV_PJMFORM> PJV_PJMFORM { get; set; }
        public PJMFORM_CMEPMS PJMFORM { get; set; }
        public PJMFORM_CMEPMS PJMFORM_P { get; set; }
        public int page { get; set; }
        public int branchPage { get; set; }
        public int SEQ_NO { get; set; }
        public decimal PID { get; set; }
        public decimal EditPID { get; set; }
        public decimal DID { get; set; }
        public string SYR { get; set; }
        public string END_SYR { get; set; }
        public string PJ_NAME { get; set; }
        public string WK_CODE { get; set; }
        public string PJ_FUND { get; set; }
        public string ANPID { get; set; }
        public string EXE_DP { get; set; }
        public string EXE_UID { get; set; }
        public IPagedList<DataRow> MainList { get; set; }
        public IPagedList<DataRow> List1 { get; set; }
        public IPagedList<DataRow> List2 { get; set; }
        public IPagedList<DataRow> List3 { get; set; }
        public string Water_Type { get; set; } = string.Empty;

    }

    public class PJE0011Model
    {
        /// <summary>
        /// 工程名稱搜尋條件
        /// </summary>
        public string keyword2 { get; set; }
        public string city { get; set; }
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
        public string YRC_ID { get; set; }
        public string AID { get; set; }
        public int tab { get; set; }
        public string LEV1_DT { get; set; }
        public string LEV2_DT { get; set; }
        public string LEV3_DT { get; set; }

        public int page1 { get; set; }
        public int page2 { get; set; }
        public string DeleteId1 { get; set; } = string.Empty;
        public string DeleteId2 { get; set; } = string.Empty;
        public string DeleteId3 { get; set; } = string.Empty;
        public string[] DeleteIds { get; set; } = Array.Empty<string>();

        //PJDCATY;
        public string[] PJDCATY_CID { get; set; } = Array.Empty<string>();
        public string[] PJDCATY_CATY_CODE { get; set; } = Array.Empty<string>();
        public string[] PJDCATY_CATY_NAME { get; set; } = Array.Empty<string>();
        public string[] PJDCATY_EVRF_FG { get; set; } = Array.Empty<string>();
        public string[] PJDCATY_OTDSG_DT { get; set; } = Array.Empty<string>();
        public string[] PJDCATY_OTDSG_BUSS { get; set; } = Array.Empty<string>();
        public string[] PJDCATY_DSGRL_DT { get; set; } = Array.Empty<string>();
        public string[] PJDCATY_PCCJDG_DT { get; set; } = Array.Empty<string>();
        public string[] PJDCATY_PCCJDG_FUND { get; set; } = Array.Empty<string>();

        //PJDCMCT
        public string[] PJDCMCT_PPID { get; set; } = Array.Empty<string>();
        public string[] PJDCMCT_CMS_CODE { get; set; } = Array.Empty<string>();
        public string[] PJDCMCT_CMS_NAME { get; set; } = Array.Empty<string>();
        public string[] PJDCMCT_CMS_UNIT { get; set; } = Array.Empty<string>();
        public string[] PJDCMCT_CMS_NOTE { get; set; } = Array.Empty<string>();

        //PJDFUND
        public string[] PJDFUND_PPID { get; set; } = Array.Empty<string>();
        public string[] PJDFUND_SYR { get; set; } = Array.Empty<string>();
        public string[] PJDFUND_IPJ_CODE { get; set; } = Array.Empty<string>();
        public string[] PJDFUND_FUND { get; set; } = Array.Empty<string>();
        public string[] PJDFUND_FUND_DESC { get; set; } = Array.Empty<string>();
        public string[] PJDFUND_YBG_FUND { get; set; } = Array.Empty<string>();
        public string[] PJDFUND_PID { get; set; } = Array.Empty<string>();

        //PJDFDIS
        public HttpPostedFileBase[] files { get; set; } = Array.Empty<HttpPostedFileBase>();
        public string[] FID { get; set; } = Array.Empty<string>();
        public string[] FILE_DESC { get; set; } = Array.Empty<string>();
        public string[] FILE_NAME { get; set; } = Array.Empty<string>();

        //PJDSUPD
        //public string[] PJDSUPD_PPID { get; set; } = Array.Empty<string>();
        public string[] PJDSUPD_DEPT_NO { get; set; } = Array.Empty<string>();

        //PJDSYRC
        public string[] PJDSYRC_PSG_NAME { get; set; } = Array.Empty<string>();
        public string[] PJDSYRC_BGN_SYR { get; set; } = Array.Empty<string>();
        public string[] PJDSYRC_END_SYR { get; set; } = Array.Empty<string>();

    }

    public class PJE0020Model : PJEModel
    {
        public IPagedList<PJV_PJMFORM> PJV_PJMFORM { get; set; }

        public decimal CID { get; set; }
        public decimal CATY_FUND { get; set; }
        public string CATY_NAME { get; set; }
        //分項
        public string[] PIDS { get; set; }
        public string[] CIDS { get; set; } //序號 
        public string[] CATY_CODES { get; set; }//經費序號
        public string[] CATY_FUNDS { get; set; }//經費
        public string[] ACC_IDS { get; set; } //會計科目
        public string[] CATY_TYPES { get; set; } //資本門
        public string[] ADV_DPS { get; set; }//經管單位
        public string[] ADV_UIDS { get; set; }//經管人
        public string[] CATY_NAMES { get; set; }//內容
        //分配       
        public string[] SYRS { get; set; }
        public string[] M_01 { get; set; }
        public string[] M_02 { get; set; }
        public string[] M_03 { get; set; }
        public string[] M_04 { get; set; }
        public string[] M_05 { get; set; }
        public string[] M_06 { get; set; }
        public string[] M_07 { get; set; }
        public string[] M_08 { get; set; }
        public string[] M_09 { get; set; }
        public string[] M_10 { get; set; }
        public string[] M_11 { get; set; }
        public string[] M_12 { get; set; }



    }

    public class PJE0030Model : PJEModel
    {
        public decimal Limit { get; set; }
        public decimal DelTid { get; set; }
        public decimal DelFid { get; set; }
        public string SYR { get; set; }
        public string stype { get; set; }
        public string keyword { get; set; }
        public IPagedList<DataRow> MainList { get; set; }
        public int PPID { get; set; }
        public string PJDSYRC { get; set; }
        public string addkey { get; set; }
        public int prePage { get; set; }
        public PJDFDCR PJDFDCR { get; set; }
        public PJMFORM_CMEPMS PJMFORM { get; set; }
        public decimal TID { get; set; }
        //public CTHMNDS CTHMNDS { get; set; }
        public TDV_TDHCFTS TDV_TDHCFTS { get; set; }
        public decimal[] CTDITFD_TID { get; set; }
        public string[] CTDITFD_SYR { get; set; }
        public string[] CTDITFD_FD_CODE { get; set; }
        public decimal[] CTDITFD_FUND { get; set; }
        public string[] CTDITFD_FUND_DESC { get; set; }
        public decimal[] CTDITFD_FID { get; set; }
        public string AID { get; set; }
        public APMFORM APMFORM { get; set; }
        public string CITY { get; set; }
        public string CITY_SELECT { get; set; }
        public string TOWN { get; set; }
        public string TD_CATY { get; set; }
        public string EXE_TYPE { get; set; }
        public string IPJ_CODE { get; set; }
        public string N_IPJ_CODE { get; set; }
        public string EMG_CATY { get; set; }
        public string ACC_ID { get; set; }
        public string TD_NAME { get; set; }
        public string PJ_CATY_CODE { get; set; }
        public decimal YR_FUND { set; get; }
        public string TD_GROUP { get; set; }
        public string NOCTRL_FG { get; set; }
        public string FCAL_FG { get; set; } = string.Empty;
        public string UCNT_FG { get; set; }
        public string B_GROUP { get; set; }
        public string F_GROUP { get; set; }
        public CRDMNDS CRDMNDS { get; set; }
        //導已成案 APMFORM
        public string[] TD_CODE_P { get; set; }
        public string[] FD_CODE_P { get; set; }
        public string[] CITY_P { get; set; }
        public string[] TOWN_P { get; set; }
        public string[] PJ_CATY_CODE_P { get; set; }//工程編號
        public string[] CATY_CODE_P { get; set; }//案件類型
        public string[] IPJ_CODE_P { get; set; }//經費層級
        public string[] TD_CATY_P { get; set; }//工程類別
        public string[] YR_WORK_P { get; set; }//工作類別


        public string NO1 { get; set; }//經費編號
        public string NO2 { get; set; }//經費編號
        public string NO3 { get; set; }//經費編號
        public string NO4 { get; set; }//經費編號
        public string NO5 { get; set; }//經費編號
        public string NO6 { get; set; }//經費編號
    }

    public class PJE0040Model : PJEModel
    {

        public string SEQ_NO_M { get; set; }
        public string SEQ_NO_S { get; set; }


    }
}