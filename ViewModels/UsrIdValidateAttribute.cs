using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    sealed public class UsrIdValidateAttribute : ValidationAttribute
    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            WCals.Models.CMBASEModels.CMBASEEntities db = new WCals.Models.CMBASEModels.CMBASEEntities();
            if (null != value && !string.IsNullOrEmpty(value.ToString()))
            {
                string usrId = value as string;
                var OGCUSER = (from f in db.OGCUSER
                               where f.USR_ID == usrId
                               select f).FirstOrDefault();

                if (null == OGCUSER) {
                    return new ValidationResult("此舊系統帳號不存在。");
                }
            }

            return ValidationResult.Success;
        }
    }
}