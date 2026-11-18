using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace WCals.ViewModels
{
    public class ADAccountViewModel
    {
        // 舊帳號 有填入則必須存在於目前OGCUSER
        [UsrIdValidate]
        [Display(Name = "帳號")]
        public string Account { get; set; }

        // AD帳號 必須不存在於目前OGCUSER，必須同時檢核USR_ID與AD_USR_ID
        [Required]
        [UsrIdAdValidate]
        [Display(Name = "AD帳號")]
        public string ADAccount { get; set; }

        [Required]
        [Display(Name = "申請部門")]
        public string DeptNo { get; set; }

        [Required]
        [Display(Name = "姓名")]
        public string UserName { get; set; }

        [Display(Name = "電子郵件")]
        public string Email { get; set; }

    }
}