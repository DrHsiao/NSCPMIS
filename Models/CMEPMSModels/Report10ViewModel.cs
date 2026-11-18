using Aspose.Cells.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models.CMEPMSModels
{
    public class Report10ViewModel
    {
        public decimal PID { get; set; }
        public decimal TID { get; set; }
        public string SYR { get; set; }
        public string PJ_NAME { get; set; }
        public string TD_CODE { get; set; }
        public string TD_NAME { get; set; }
        public string CITY { get; set; }
        public string TOWN { get; set; }
        public string FIT_CODE { get; set; }
        public int FUND { get; set; }
        public int COST_FEE { get; set; }
        public string TD_MEMO { get; set; }
    }

    public class Report11ViewModel
    {
        public decimal PID { get; set; }
        public decimal TID { get; set; }
        public string SYR { get; set; }
        public string PJ_NAME { get; set; }
        public string TD_CODE { get; set; }
        public string TD_NAME { get; set; }
        public string CITY { get; set; }
        public string TOWN { get; set; }
        public string FIT_CODE { get; set; }
        public decimal FUND { get; set; }
        public int COST_FEE { get; set; }
        public string TD_MEMO { get; set; }
        public decimal TD_FUND { get; set; }
    }
}