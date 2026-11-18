using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    public class AUTHViewModel
    {
        public IEnumerable<WCals.Models.CMBASEModels.SMCMENU3> ListSMCMENU { get; set; }

        public WCals.Models.CMBASEModels.SMCMENU3 DataSMCMENU { get; set; }

        public String jsonSMCAUTH_D { get; set; }

        public String jsonSMCAUTH_G { get; set; }

    }
}