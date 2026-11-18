using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models.CMBASEModels
{
    public class GroupChooseViewModel
    {
        public IEnumerable<OGCUSER>  USER { get; set; }
        
        public OGCDEPT ChooseDEPT { get; set; }
        public IEnumerable<OGCDPUR> ChooseDPUR { get; set; }
    }

    public class GroupUserViewModel
    {
        public OGCDPUR DPUR { get; set; }
        public OGCUSER USER { get; set; }
    }

}