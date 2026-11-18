using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models.CMEPMSModels
{
    public class TDDSTCF_INSERT_CLOMNS
    {
        public decimal ID { get; set; }
        public string SYR { get; set; }
        public string SMN { get; set; }
        public string FD_CODE { get; set; }
        public string TD_CODE { get; set; }
        public string TD_CODE1 { get; set; }
        public Nullable<decimal> COST_FEE { get; set; }
        public Nullable<decimal> PPAY_FEE { get; set; }
        public Nullable<System.DateTime> ADD_DT { get; set; }
        public Nullable<System.DateTime> UPD_DT { get; set; }
        public string UPD_UID { get; set; }
        public decimal? MATH { get; internal set; }
    }
}