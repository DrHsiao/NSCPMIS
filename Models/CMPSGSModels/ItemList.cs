using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models.CMPSGSModels
{
    public class ItemList
    {
        public string[] MT_NAME { get; set;}
        public string[] MT_DESC { get; set;}
        public string[] MT_UNIT { get; set;}
        public decimal[] MT_QTY { get; set;}
        public string[] MT_SPEC { get; set;}
    }
}