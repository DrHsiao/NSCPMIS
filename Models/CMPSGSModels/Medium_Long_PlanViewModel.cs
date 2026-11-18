using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models.CMPSGSModels
{
    public class Medium_Long_PlanViewModel
    {
        public string CATY_COD { get; set; }
        public decimal PPID { get; set; }
        public string SYR { get; set; }
        public string IPJ_CODE { get; set; }
        public string CODE_DESC { get; set; }
        public Nullable<decimal> FUND { get; set; }
        public Nullable<decimal> YBG_FUND { get; set; }
        public string FUND_DESC { get; set; }
        public Nullable<System.DateTime> ADD_DT { get; set; }
        public Nullable<System.DateTime> UPD_DT { get; set; }
        public string UPD_UID { get; set; }
    }
}