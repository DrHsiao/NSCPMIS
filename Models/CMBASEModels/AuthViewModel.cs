using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models.CMBASEModels
{
    public class AuthViewModel
    {
        public List<SMCMENU3> menu1 { get; set; }
        public List<SMCMENU3> menu2 { get; set; }
        public List<SMCMENU3> menu3 { get; set; }
        //public List<AddViewModel> avm1 { get; set; }
}
    public class AddViewModel
    {
        public string SYS_ID { get; set; }
        public bool? CUS_FG { get; set; }
        public string DEPT_NAME { get; set; }
        public string DEPT_NO { get; set; }
        public string DATA_TP { get; set; }
        public bool SelectedAdd { get; set; }

        public int Index { get; set; }
    }
}