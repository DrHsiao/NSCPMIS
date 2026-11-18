using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models.CMPSGSModels
{
    public class FundList
    {
        public string[] SYR { get; set; }
        public string[] FIT_CODE { get; set; }
        public string[] FUND_DESC { get; set; }
        public Nullable<decimal>[] FUND { get; set; }
    }
}