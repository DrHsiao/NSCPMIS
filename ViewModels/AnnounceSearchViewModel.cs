using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    public class AnnounceSearchViewModel
    {
        public String SUBJECT { get; set; }

        public String CONTENT { get; set; }

        public Nullable<System.DateTime> POST_DT_START { get; set; }

        public Nullable<System.DateTime> POST_DT_END { get; set; }
    }
}