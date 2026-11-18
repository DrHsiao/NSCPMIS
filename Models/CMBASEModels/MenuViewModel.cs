using PagedList;
using System.Collections.Generic;

namespace WCals.Models.CMBASEModels
{
    public class DEPTGROUP_ViewModel
    {
        public SECCITY SECCITY { get; set; }
        public OGCDEPT DEPTNAME { get; set; }
        public OGCDPUR DEPTID { get; set; }
    }
    public class DEPT_NO_ViewModel
    {
        public OGCDPUR DEPTVM { get; set; }
        public SMCAUTH FUNCTIONVM { get; set; }
        public SMCMENU3 MENUVM { get; set; }
    }

    public class ALL
    {
        public List<SMCMENU3> Menu1 { get; set; }
        public List<SMCMENU3> Menu2 { get; set; }
        public List<SMCMENU3> Menu3 { get; set; }
        public List<DEPT_NO_ViewModel> Menu4 { get; set; }
        public IPagedList<OGCUSER> User { get; set; }
        public List<OGCUSER> Customer { get; set; }
    }

    public class DeptGroupViewModel
    {
        public List<OGCDPUR> DEPTID { get; set; }
        public List<OGCDEPT> DEPTNAME { get; set; }
    }
    public class MenuViewModel
    {
        public List<SMCMENU3> Menu1 { get; set; }
        public List<SMCMENU3> Menu2 { get; set; }
        public List<SMCMENU3> Menu3 { get; set; }

        public List<SMCMENU3> Menu { get; set; }
        public List<SMCAUTH> FUNCTION { get; set; }

        public List<OGCDPUR> DEPT { get; set; }
    }

    public class ChooseViewModel
    {
        public OGCUSER OGCUSER { get; set; }
        public OGCDPUR OGCDPUR { get; set; }
        public OGCDEPT OGCDEPT { get; set; }
    }


}