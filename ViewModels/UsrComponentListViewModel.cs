using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    public class UsrComponentListViewModel
    {
        public UsrComponentSearchViewModel SearchParameter { get; set; }

        public IPagedList<OGCUSERItemViewModel> OGCUSER { get; set; }
    }
}