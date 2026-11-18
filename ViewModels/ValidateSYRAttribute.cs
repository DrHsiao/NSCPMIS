using System;
using System.ComponentModel.DataAnnotations;

namespace WCals.ViewModels
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    sealed public class ValidatSYRAttribute : ValidationAttribute

    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {            
            if (null != value)
            {
                string stringVal = value.ToString();
                // 檢核填寫格式
                if (stringVal.Length != 3 || !int.TryParse(stringVal, out int n))
                {
                    return new ValidationResult("輸入格式有誤，請輸入3位阿拉伯數字之民國年月，例如:099, 100,...");
                }
            }

            return ValidationResult.Success;
        }
    }
}