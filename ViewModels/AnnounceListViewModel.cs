using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    public class AnnounceListViewModel
    {
        public AnnounceSearchViewModel SearchParameter { get; set; }

        public List<WCals.Models.CMBASEModels.SMCNEWS> SMCNEWS { get; set; }

        public string ADV_DP { get; set; }
    }
}