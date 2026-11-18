using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace WCals.ViewModels
{
    public class OGCDEPTEditViewModel
    {
        public OGCDEPTItemViewModel OGCDEPT_WITH_JOIN { get; set; }

        [Display(Name = "部門編號")]
        [Required(ErrorMessage = "請輸入部門編號")]
        public string DEPT_NO { get; set; }

        public string AN_DEPT_SHOW_VAL { get; set; }
    }
}