using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WCals.ViewModels
{
    public class OGCDEPTItemViewModel
    {
        public Models.CMBASEModels.OGCDEPT OGCDEPT { get; set; }

        public string AN_DEPT_NAME { get; set; }
    }
}