using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    public class OGCUSERListViewModel
    {
        public OGCUSERSearchViewModel SearchParameter { get; set; }

        public IPagedList<Models.CMBASEModels.OGCUSER> OGCUSER { get; set; }
    }
}