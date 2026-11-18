using WCals.Models.CMPSGSModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models
{
    public class ProposalRequestViewModel
    {
        public decimal PPID { get; set; }
        public decimal YRC_ID { get; set; }
        public string PJ_NAME { get; set; }
        public string SYR { get ; set; }
    }
}