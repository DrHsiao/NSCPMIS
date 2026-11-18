using System.Collections.Generic;
using WCals.Models.CMBASEModels;

namespace WCals.ViewModels
{
    public class HomeIndexViewModel
    {
        public IEnumerable<SMCNEWS> ListSMCNEWS { get; set; }
        public WCals.Models.CMBASEModels.SMCNEWS dataSMCNEWS { get; set; }
        public IEnumerable<SMCNEWS> ListSMCDOCS { get; set; }
        public WCals.Models.CMBASEModels.SMCNEWS dataSMCDOCS { get; set; }
        public IEnumerable<SMCNEWSDOC> ListSMCNEWSDOC { get; set; }
        //public IEnumerable<SMCQADA> ListSMCQADA { get; set; }
        //public WCals.Models.SMCQADA dataSMCQADA { get; set; }        
    }
}