using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    sealed public class OGCUSERValidateADAttribute : ValidationAttribute
    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            WCals.Models.CMBASEModels.CMBASEEntities db = new WCals.Models.CMBASEModels.CMBASEEntities();
            if (null != value && !string.IsNullOrEmpty(value.ToString()))
            {
                string AD_USR_ID = value as string;

                var OGCUSER = (from f in db.OGCUSER
                               where (f.USR_ID == AD_USR_ID)
                               select f).ToList();
                if (OGCUSER.Count > 0)
                {
                    return new ValidationResult("此AD帳號已經存在。");
                }
            }

            return ValidationResult.Success;
        }
    }
}