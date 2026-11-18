using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace WCals.ViewModels
{
    public class OGCUSEREditViewModel
    {
        public Models.CMBASEModels.OGCUSER OGCUSER { get; set; }

        public String jsonOGCDPUR_D { get; set; }

        public String jsonOGCDPUR_G { get; set; }

        public int page { get; set; }

        [Required(ErrorMessage = "請輸入使用者編號")]
        [OGCUSERValidateID]
        public string VALIDATE_USR_ID { get; set; }

        [OGCUSERValidateAD]
        public string VALIDATE_AD_USR_ID { get; set; }
    }
}