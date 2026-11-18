using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    public class OGCDPUREditModel
    {
        public Models.CMBASEModels.OGCUSER OGCUSER { get; set; }

        public String jsonOGCDPUR_D { get; set; }

        public String jsonOGCDPUR_G { get; set; }
    }
}