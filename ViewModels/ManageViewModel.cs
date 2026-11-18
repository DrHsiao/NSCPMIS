using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using PagedList;

namespace WCals.ViewModels
{
    public class ManageViewModel
    {
        //public IPagedList<WCals.Models.CMBASEModels.SMCQADA> SMCQADAList { get; set; }

        public IPagedList<WCals.Models.CMBASEModels.SMCLOGS> SMCLOGSList { get; set; }

        public LogSearchViewModel SearchParameter { get; set; }

        public int PageNumber { get; set; }
    }
}