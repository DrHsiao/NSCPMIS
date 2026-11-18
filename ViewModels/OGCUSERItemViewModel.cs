using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace WCals.ViewModels
{
    public class OGCUSERItemViewModel
    {
        public Models.CMBASEModels.OGCUSER OGCUSER { get; set; }

        public string DEPT_NAME { get; set; }
    }
}