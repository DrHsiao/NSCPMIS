using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WCals.ViewModels
{
    public class LogonViewModel
    {
        //[Required]
        //[Display(Name = "所屬單位")]
        //public string sc_no { get; set; }

        [Required]
        [Display(Name = "帳號")]
        public string Account { get; set; }

        [Required]
        [DataType(DataType.Password)]
        [Display(Name ="密碼")]
        public string Password { get; set; }

        [Display(Name ="記住我")]
        public bool Remember { get; set; }
    }
}