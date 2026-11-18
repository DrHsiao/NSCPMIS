using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace WCals.Models.CMBASEModels
{
    public class DeptViewModel
    {
        public List<OGCDEPT> Dept1 { get; set; }
        public List<OGCDEPT> Dept2 { get; set; }
        public IEnumerable< DeptUserViewModel> DeptUser { get; set; }
        public List<OGCUSER> User { get; set; }
        public string DEPT_NO { get; set; }
        public string DEPT_NAME { get; set; }
        public string DEPT_CHIEF { get; set; }
        public bool MGR_FG { get; set; }
        public string AN_DEPT_NO { get; set; }
        public string selected_user_id { get; set; }
        public string selected_an_user_id { get; set; }
        public IEnumerable<SelectListItem> id_name { get; set; }
        public IEnumerable<SelectListItem> an_id_name { get; set; }
        public List<ListItem> List { get; set; }


    }
    public class DeptUserViewModel
    {
        public string DEPTNO { get; set; }
        public string USRID { get; set; }
        public string USRNAME { get; set; }
        public string DEPTNAME { get; set; }
        public bool MGR_FG { get; set; }
        public string DEPT_CHIEF { get; set; }
        public IEnumerable<SelectListItem> id_name { get; set; }
        public string selected_user_id { get; set; }
    }
    public class ListItem
    {
        public string Id { get; set; }
    }
}