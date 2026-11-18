using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WCals.ViewModels.NewModels {
    public class SYSModels {
        public string[] SYS_IDS { get; set; }
        public string[]CITYS       {get;set;}
        public string[]CITY_IDS    {get;set;}
        public string[]SYS_NAMES   {get;set;}
        public string[]CONST_TPS   {get;set;}
        public string[]IS_SUPS     {get;set;}
        public string[]ADD_DTS     {get;set;}
        public string[]UPD_DTS     {get;set;}
        public string[]UPD_UIDS    {get;set;}
        public string[]ADV_UIDS    {get;set;}
        public string[]ADV_DPS     {get;set;}
        public string[]STAT_NOS    {get;set;}

        public string SYS_ID { get; set; }
        public string CITY { get; set; }
        public string CITY_ID { get; set; }
        public string SYS_NAME { get; set; }
        public string CONST_TP { get; set; }
        public string IS_SUP { get; set; }
        public string ADD_DT { get; set; }
        public string UPD_DT { get; set; }
        public string UPD_UID { get; set; }
        public string ADV_UID { get; set; }
        public string ADV_DP { get; set; }
        public string[] STAT_NO { get; set; }




    }

    public class SYS0030Model:SYSModels{

        
        public string sys { get; set; }
        public string doc_type { get; set; }
        public string doc_name { get; set; }
        public IPagedList<DataRow> List1 { get; set; }
        public string PJ_ID { get; set; }
        public int page { get; set; }

    }
}