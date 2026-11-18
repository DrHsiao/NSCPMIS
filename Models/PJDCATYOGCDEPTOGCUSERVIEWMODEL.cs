using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models
{
    public class PJDCATYOGCDEPTOGCUSERVIEWMODEL
    {
        public List<CMEPMSModels.PJDCATY_CMEPMS> PJDCATY { get; set; }
        public List<CMBASEModels.OGCDEPT> OGCDEPT { get; set; }
        public List<CMBASEModels.OGCUSER> OGCUSER { get; set; }
    }
}