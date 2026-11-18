using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WCals.Models.CMBASEModels;
using WCals.Models.CMEPMSModels;

namespace WCals.Models
{
    public class ReportViewModel
    {
        public List<SiteCityViewModel> SiteCityViewModel { get; set; }
        public List<PJMFORM_CMEPMS> pJMFORMs { get; set; }
        public List<PJMFORM_CMEPMS> PJMFORM { get; set; }
        public IEnumerable<SelectListItem> PJ_NAME { get; set; }

    }

    public class Report13ViewModel
    {
        public decimal PID { get; set; }
        public int AN_PID { get; set; }
        public string SYR { get; set; }
        public string PJ_NAME { get; set; }
    }
}