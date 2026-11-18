using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WCals.Models
{
    public class PualViewModel
    {
        public string PJ_CODE { get; set; }
        public string PJ_NAME { get; set; }
        public string BATC_NAME { get; set; }
        public string PJ_FUND { get; set; }
        public string APMFORM_APCODE { get; set; }
        public dynamic DID { get; set; }
        public IQueryable<DBO_WATERRANGE> DBO { get; set; }
        public IQueryable<string> SMV_SMCCITY { get; set; }
        public IQueryable<string> SMV_SMCTOWN { get; set; }
        public IQueryable<PJDFDND> PJDFDND_OBJ { get; set; }
        public IQueryable<SMCCODE> SMCCODE_OBJ { get; set; }
        public string START_YM { get; set; }
        public string END_YM { get; set; }


    }
}