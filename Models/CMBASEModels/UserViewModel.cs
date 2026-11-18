using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WCals.Models.CMBASEModels
{
    public class UserViewModel
    {
        public List<OGCDEPT> OGCDEPT { get; set; }
        public List<SMCPOSTN> SMCPOSTN { get; set; }
        public List<OGCUSER> OGCUSER { get; set; }
        public List<SECCITY> SECCITY { get; set; }
        public List<SECCITY> ALLSECCITY { get; set; }
        public List<SMCPOSTN> ALLSMCPOSTN { get; set; }
    }

    public class SiteCityViewModel
    {
        public SMCPOSTN sMCPOSTN { get; set; }
        public SECCITY sECCITY { get; set; }
        public List<SECCITY> SECCITY { get; set; }
        public CMEPMSModels.PJMFORM_CMEPMS pJMFORM { get; set; }
    }
    public class state9ViewModel
    {
        public List<OGCTITLE> OGCTITLE { get; set; }
        public List<SMCPOSTN> SMCPOSTN { get; set; }

        public List<CMEPMSModels.PJV_PJDPCAT> PJV_PJDPCAT { get; set; }

        public List<Water_Type_Use> Water_Type { get; set; }
    }

    public class Water_Type_Use
    {
        public string SYR { get; set; }
        public string AN_PJ_NAME { get; set; }
        public decimal PID { get; set; }
        public string PJ_NAME { get; set; }
        public string Water_Type { get; set; }
    }


    public class ChooseUser1ViewModel
    {
        public OGCUSER oGCUSER { get; set; }
        public SECCITY sECCITY { get; set; }
    }

    public class ChooseUser2ViewModel
    {
        public OGCDEPT oGCDEPT { get; set; }

        public OGCDPUR oGCDPUR { get; set; }
    }

    public class ChooseUser3ViewModel
    {
        public ChooseUser1ViewModel chooseUser1ViewModel { get; set; }
        public ChooseUser2ViewModel ChooseUser2ViewModel { get; set; }
    }



    public class DeptCityViewModel
    {
        public OGCDEPT oGCDEPT { get; set; }
        public SECCITY sECCITY { get; set; }
    }
    public class UserCityViewModel
    {
        public OGCUSER oGCUSER { get; set; }
        public SECCITY sECCITY { get; set; }
    }

    public class CityViewModel
    {
        public bool City0 { get; set; }
        public bool City1 { get; set; }
        public bool City2 { get; set; }

        public bool City3 { get; set; }
        public bool City4 { get; set; }
        public bool City5 { get; set; }
        public bool City6 { get; set; }
        public bool City7 { get; set; }
        public bool City8 { get; set; }
        public bool City9 { get; set; }
        public bool City10 { get; set; }
        public bool City11 { get; set; }
        public bool City12 { get; set; }
        public bool City13 { get; set; }
        public bool City14 { get; set; }
        public bool City15 { get; set; }
        public bool City16 { get; set; }
        public bool City17 { get; set; }
        public bool City18 { get; set; }
        public bool City19 { get; set; }
        public bool City20 { get; set; }
        public bool City21 { get; set; }
        public bool City22 { get; set; }
    }
}