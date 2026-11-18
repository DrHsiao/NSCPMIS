using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    public class OGCDEPTListViewModel
    {
        public OGCDEPTSearchViewModel SearchParameter { get; set; }

        public IPagedList<OGCDEPTItemViewModel> OGCDEPT_AVAILABLE { get; set; }

        public IPagedList<OGCDEPTItemViewModel> OGCDEPT_CANCELED { get; set; }
    }
}