using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.ExtensionMethods
{
    public static class DeciamlExtentions
    {
        public static decimal? Normalize(this decimal? value)
        {
            return (null == value) ? value : value.Value / 1.000000000000000000000000000000000m;
        }
    }

}