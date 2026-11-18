using Newtonsoft.Json;
using System;
using System.Linq;
using System.Web.Mvc;
using WCals.Models.CMBASEModels;

namespace WCals.Controllers
{
    public class UserController : Controller
    {
        CMBASEEntities db = new CMBASEEntities();

        // GET: User
        [HttpGet]
        public ActionResult Index(string ID, string Name)
        {
            var SECCITY = db.SECCITY.ToList();
            var OGCUSER = db.OGCUSER.ToList();
            var Query = from s in SECCITY
                        join st in OGCUSER on s.USR_ID equals st.USR_ID
                        select new UserCityViewModel { sECCITY = s, oGCUSER = st };
            if (ID != null)
            {
                Query = Query.Where(m => m.sECCITY.USR_ID.Contains(ID)).ToList();
            }
            if (Name != null)
            {
                Query = Query.Where(m => m.oGCUSER.USR_NAME.Contains(Name)).ToList();
            }
            Query = Query.DistinctBy(m => m.sECCITY.USR_ID).ToList();

            return View(Query);
        }

        public ActionResult ChooseCity(string City)
        {
            string result = City.Substring(4);
            int number = int.Parse(result);
            var ALL_CITY = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            string City_Name = ALL_CITY[number].CITY;
            var CITY_TOWN = db.SMCPOSTN.Where(m => m.CITY == City_Name).ToList();
            string json = JsonConvert.SerializeObject(CITY_TOWN);
            return Json(json);
        }

        public ActionResult UserEdit(string choose, string chooseName, string TestQuest /*, string Site*/)
        {
            ViewBag.Errorr = TestQuest;
            ViewBag.choose = choose;
            ViewBag.chooseName = chooseName;
            UserViewModel userViewModel = new UserViewModel();
            userViewModel.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            userViewModel.OGCDEPT = db.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            var query = db.OGCUSER.Where(m => m.USR_ID == choose).FirstOrDefault();
            userViewModel.SECCITY = db.SECCITY.Where(m => m.USR_ID == query.USR_ID /*&& m.AREA == "北區"*/).ToList();
            //if(Site != null)
            //{
            //    userViewModel.SECCITY = db.SECCITY.Where(m => m.USR_ID == query.USR_ID && m.AREA == Site).ToList();
            //}
            userViewModel.ALLSECCITY = db.SECCITY.Where(m => m.USR_ID == choose).ToList();
            userViewModel.ALLSMCPOSTN = db.SMCPOSTN.ToList();
            return View(userViewModel);
        }
        [HttpPost]
        public ActionResult UserEdit(string User, string UserName, string[] CityCheckbox, string[] SelectedboxTOWN/*, string Site*/ )
        {
            string TestQuest = "";
            SECCITY s = new SECCITY();
            UserViewModel userViewModel = new UserViewModel();
            userViewModel.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            userViewModel.OGCDEPT = db.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            s.USR_ID = User;

            for (int a = 0; a < userViewModel.SMCPOSTN.Count(); a++)
            {
                bool foundMatch = false;
                var city = userViewModel.SMCPOSTN[a].CITY;

                string newCity = "";

                switch (city)
                {
                    case "新北市":
                        newCity = "台北縣";
                        break;
                    case "高雄市":
                        newCity = "高雄縣";
                        break;
                    case "台南市":
                        newCity = "台南縣";
                        break;
                    case "台中市":
                        newCity = "台中縣";
                        break;
                    case "桃園市":
                        newCity = "桃園縣";
                        break;
                }

                if (!string.IsNullOrEmpty(newCity))
                {
                    var old_CITY = db.SECCITY.Where(m => m.CITY == newCity && m.USR_ID == User /*&& m.AREA == Site*/).ToList();
                    db.SECCITY.RemoveRange(old_CITY);
                    db.SaveChanges();
                }
                var g = db.SECCITY.Where(m => m.CITY == city && m.USR_ID == User/* && m.AREA ==Site*/).ToList();
                db.SECCITY.RemoveRange(g);
                db.SaveChanges();

                for (int b = 0; b < CityCheckbox.Count(); b++)
                {

                    if (userViewModel.SMCPOSTN[a].CITY == CityCheckbox[b])
                    {
                        var c = db.SECCITY.Where(m => m.CITY == city && m.USR_ID == User).FirstOrDefault();

                        if (c == null)
                        {
                            foreach (var item in SelectedboxTOWN)
                            {
                                string CityFirstName = item.Substring(0, 3);
                                string Town = item.Substring(3);
                                if (city == CityFirstName)
                                {
                                    s.ADD_DT = DateTime.Now;
                                    s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                                    s.CITY = userViewModel.SMCPOSTN[a].CITY.ToString();
                                    s.TOWN = Town;
                                    //s.AREA = Site;
                                    db.SECCITY.Add(s);
                                    db.SaveChanges();

                                    string newCity2 = "";

                                    switch (city)
                                    {
                                        case "新北市":
                                            newCity2 = "台北縣";
                                            break;
                                        case "高雄市":
                                            newCity2 = "高雄縣";
                                            break;
                                        case "台南市":
                                            newCity2 = "台南縣";
                                            break;
                                        case "台中市":
                                            newCity2 = "台中縣";
                                            break;
                                        case "桃園市":
                                            newCity2 = "桃園縣";
                                            break;
                                    }

                                    if (!string.IsNullOrEmpty(newCity2))
                                    {
                                        s.ADD_DT = DateTime.Now;
                                        s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                                        s.CITY = newCity2;
                                        s.TOWN = Town;
                                        //s.AREA = Site;
                                        db.SECCITY.Add(s);
                                        db.SaveChanges();
                                    }
                                }
                            }
                        }
                        foundMatch = true;
                    }
                }

                if (!foundMatch)
                {
                    try
                    {
                        var c = userViewModel.SMCPOSTN[a].CITY;
                        var t = db.SECCITY.Where(m => m.CITY == city && m.USR_ID == User/* && m.AREA ==Site*/).ToList();


                        db.SECCITY.RemoveRange(t);
                        db.SaveChanges();

                        string newCity2 = "";

                        switch (c)
                        {
                            case "新北市":
                                newCity2 = "台北縣";
                                break;
                            case "高雄市":
                                newCity2 = "高雄縣";
                                break;
                            case "台南市":
                                newCity2 = "台南縣";
                                break;
                            case "台中市":
                                newCity2 = "台中縣";
                                break;
                            case "桃園市":
                                newCity2 = "桃園縣";
                                break;
                        }
                        if (!string.IsNullOrEmpty(newCity2))
                        {
                            var R = db.SECCITY.Where(m => m.CITY == newCity2 && m.USR_ID == User /*&& m.AREA == Site*/).ToList();
                            db.SECCITY.RemoveRange(R);
                            db.SaveChanges();
                        }
                        TestQuest = "修改成功";
                    }
                    catch { }

                }
            }

            //for (int i = 0; i < 23; i++)
            //{
            //    bool propertyValue = (bool)typeof(CityViewModel).GetProperty("City" + i).GetValue(model, null);
            //    if (propertyValue)
            //    {
            //        for (int z = 0; z < 23; z++)
            //        {
            //            if (i == z)
            //            {
            //                var a = userViewModel.SMCPOSTN[z].CITY;
            //                var b = db.SECCITY.Where(m => m.CITY == a && m.USR_ID == User).FirstOrDefault();
            //                if (b == null)
            //                {
            //                    s.ADD_DT = DateTime.Now;
            //                    s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            //                    s.CITY = userViewModel.SMCPOSTN[z].CITY.ToString();
            //                    db.SECCITY.Add(s);
            //                    db.SaveChanges();
            //                }
            //            }
            //        }
            //    }
            //    else
            //    {
            //        for (int z = 0; z < 23; z++)
            //        {
            //            if (i == z)
            //            {

            //                try
            //                {
            //                    var a = userViewModel.SMCPOSTN[z].CITY;
            //                    var t = db.SECCITY.Where(m => m.CITY == a && m.USR_ID == User).ToList();
            //                    db.SECCITY.RemoveRange(t);
            //                    db.SaveChanges();
            //                    TestQuest = "修改成功";
            //                }
            //                catch { }
            //            }
            //        }
            //    }
            //}
            return RedirectToAction("UserEdit", "User", new { choose = User, chooseName = UserName, TestQuest = TestQuest });
        }
        [HttpGet]
        public ActionResult ChooseUser(string ID, string Name, string dropdownlist)
        {
            ViewBag.DEPT = db.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            var SECCITY = db.SECCITY.DistinctBy(m => m.USR_ID).ToList();
            var OGCDEPT = db.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            var OGCDPUR = db.OGCDPUR.Where(m => m.DATA_TP == "D").ToList();
            var OGCUSER = db.OGCUSER.ToList();
            var Query1 = from a in OGCUSER
                         join b in SECCITY on a.USR_ID equals b.USR_ID into abgroup
                         from b in abgroup.DefaultIfEmpty()
                         where b == null
                         select new ChooseUser1ViewModel { oGCUSER = a };
            var Query2 = from a in OGCDPUR
                         join b in OGCDEPT on a.DEPT_NO equals b.DEPT_NO
                         select new ChooseUser2ViewModel { oGCDPUR = a, oGCDEPT = b };
            var Query3 = from a in Query1
                         join b in Query2 on a.oGCUSER.USR_ID equals b.oGCDPUR.USR_ID
                         select new ChooseUser3ViewModel { chooseUser1ViewModel = a, ChooseUser2ViewModel = b };
            if (ID != null)
            {
                Query3 = Query3.Where(m => m.chooseUser1ViewModel.oGCUSER.USR_ID.Contains(ID)).ToList();
            }
            if (Name != null)
            {
                Query3 = Query3.Where(m => m.chooseUser1ViewModel.oGCUSER.USR_NAME.Contains(Name)).ToList();
            }
            if (dropdownlist == "請選擇")
            {
                dropdownlist = "";
            }
            if (dropdownlist != null)
            {
                Query3 = Query3.Where(m => m.ChooseUser2ViewModel.oGCDEPT.DEPT_NAME.Contains(dropdownlist)).ToList();
            }
            return View(Query3);
        }

        public ActionResult Choose(string ID, string NAME)
        {
            return RedirectToAction("UserCreate", "User", new { ID = ID, Name = NAME });
        }
        public ActionResult UserCreate(string ID, string Name, /*string Save,*/ string User)
        {
            //try
            //{
            //    if (db.SECCITY.Any(m => m.USR_ID == ID && m.AREA == "北區" && m.DEPT_NO == null))
            //    {
            //        ViewBag.North = "Y";
            //    }
            //    if (db.SECCITY.Any(m => m.USR_ID == ID && m.AREA == "中區" && m.DEPT_NO == null))
            //    {
            //        ViewBag.Central = "Y";
            //    }
            //    if (db.SECCITY.Any(m => m.USR_ID == ID && m.AREA == "南區" && m.DEPT_NO == null))
            //    {
            //        ViewBag.South = "Y";
            //    }
            //}
            //catch { }

            //if(Save != null)
            //{
            //    ViewBag.Save = Save + "新增成功";
            //}

            ViewBag.User = User;
            ViewBag.ID = ID;
            ViewBag.Name = Name;
            UserViewModel userViewModel = new UserViewModel();
            userViewModel.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            try
            {
                var SECCITY_CITY = db.SECCITY.Where(m => m.USR_ID == ID && m.DEPT_NO == null).DistinctBy(m => m.CITY).ToList();
                var includedCities = SECCITY_CITY.Select(m => m.CITY);
                ViewBag.includeCities = includedCities;
            }
            catch { }
            userViewModel.SECCITY = db.SECCITY.ToList();

            return View(userViewModel);
        }

        [HttpPost]
        public ActionResult UserCreate(string ID, CityViewModel model, string[] SelectedboxTOWN /*, string Site*/)
        {

            if (ID == "") { return RedirectToAction("Index", "User"); }
            SECCITY s = new SECCITY();
            s.USR_ID = ID;
            UserViewModel userViewModel = new UserViewModel();
            userViewModel.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            userViewModel.OGCDEPT = db.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            for (int i = 0; i < 23; i++)
            {
                bool propertyValue = (bool)typeof(CityViewModel).GetProperty("City" + i).GetValue(model, null);
                if (propertyValue)
                {
                    for (int z = 0; z < 23; z++)
                    {
                        if (i == z)
                        {
                            var a = userViewModel.SMCPOSTN[z].CITY;
                            var b = db.SECCITY.Where(m => m.CITY == a && m.USR_ID == ID).FirstOrDefault();
                            for (int m = 0; m < SelectedboxTOWN.Length; m++)
                            {
                                var Name = SelectedboxTOWN[m];
                                string CityNumber = new string(Name.Where(char.IsDigit).ToArray());
                                int CityID = int.Parse(CityNumber);
                                int endIndex = Name.IndexOf(CityNumber) + CityNumber.Length;
                                string textAfterNumber = Name.Substring(endIndex);

                                if (i == CityID)
                                {
                                    if (b == null)
                                    {
                                        s.CITY = userViewModel.SMCPOSTN[z].CITY.ToString();
                                        s.ADD_DT = DateTime.Now;
                                        s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                                        s.TOWN = textAfterNumber;
                                        //s.AREA = Site;
                                        db.SECCITY.Add(s);
                                        db.SaveChanges();
                                        string newCity = "";

                                        switch (s.CITY)
                                        {
                                            case "新北市":
                                                newCity = "台北縣";
                                                break;
                                            case "高雄市":
                                                newCity = "高雄縣";
                                                break;
                                            case "台南市":
                                                newCity = "台南縣";
                                                break;
                                            case "台中市":
                                                newCity = "台中縣";
                                                break;
                                            case "桃園市":
                                                newCity = "桃園縣";
                                                break;
                                        }

                                        if (!string.IsNullOrEmpty(newCity))
                                        {
                                            s.CITY = newCity;
                                            s.ADD_DT = DateTime.Now;
                                            s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                                            s.TOWN = textAfterNumber;
                                            //s.AREA = Site;
                                            db.SECCITY.Add(s);
                                            db.SaveChanges();
                                        }
                                    }

                                }
                            }

                        }
                    }
                }

            }
            return RedirectToAction("Index", "User"/*, new { Save = Site }*/);
        }
        [HttpGet]
        public ActionResult Dept(string DeptName)
        {
            DeptCityViewModel aa = new DeptCityViewModel();
            var SECCITY = db.SECCITY.ToList();
            var OGCDEPT = db.OGCDEPT.ToList();
            if (DeptName != null)
            {
                OGCDEPT = db.OGCDEPT.Where(m => m.DEPT_NAME.Contains(DeptName)).ToList();
            }
            var Query = from s in SECCITY
                        join st in OGCDEPT on s.DEPT_NO equals st.DEPT_NO
                        select new DeptCityViewModel { sECCITY = s, oGCDEPT = st };

            Query = Query.DistinctBy(m => m.sECCITY.DEPT_NO).ToList();
            return View(Query);
        }

        public ActionResult DeptEdit(string choose, string TestQuest/*, string Site*/)
        {
            ViewBag.Errorr = TestQuest;
            ViewBag.choose = choose;

            UserViewModel userViewModel = new UserViewModel();
            userViewModel.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            userViewModel.OGCDEPT = db.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            var query = db.OGCDEPT.Where(m => m.DEPT_NAME == choose).FirstOrDefault();
            userViewModel.SECCITY = db.SECCITY.Where(m => m.DEPT_NO == query.DEPT_NO/* && m.AREA == "北區"*/).ToList();
            //if (Site != null)
            //{
            //    userViewModel.SECCITY = db.SECCITY.Where(m => m.DEPT_NO == query.DEPT_NO && m.AREA == Site).ToList();
            //}
            userViewModel.ALLSECCITY = db.SECCITY.Where(m => m.USR_ID == choose).ToList();
            userViewModel.ALLSMCPOSTN = db.SMCPOSTN.ToList();
            return View(userViewModel);
        }

        [HttpPost]
        public ActionResult DeptEdit(string Dept, CityViewModel model, string[] CityCheckbox, string[] SelectedboxTOWN/*, string Site*/)
        {
            string TestQuest = "";
            SECCITY s = new SECCITY();
            var DEPT_Name = db.OGCDEPT.Where(m => m.DEPT_NAME == Dept).FirstOrDefault();
            var DEPT_NO = DEPT_Name.DEPT_NO.ToString();
            UserViewModel userViewModel = new UserViewModel();
            userViewModel.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            userViewModel.OGCDEPT = db.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            s.DEPT_NO = DEPT_NO;
            for (int a = 0; a < userViewModel.SMCPOSTN.Count(); a++)
            {
                bool foundMatch = false;
                var city = userViewModel.SMCPOSTN[a].CITY;

                string newCity = "";

                switch (city)
                {
                    case "新北市":
                        newCity = "台北縣";
                        break;
                    case "高雄市":
                        newCity = "高雄縣";
                        break;
                    case "台南市":
                        newCity = "台南縣";
                        break;
                    case "台中市":
                        newCity = "台中縣";
                        break;
                    case "桃園市":
                        newCity = "桃園縣";
                        break;
                }

                if (!string.IsNullOrEmpty(newCity))
                {
                    var old_CITY = db.SECCITY.Where(m => m.CITY == newCity && m.DEPT_NO == DEPT_NO /*&& m.AREA == Site*/).ToList();
                    db.SECCITY.RemoveRange(old_CITY);
                    db.SaveChanges();
                }
                var g = db.SECCITY.Where(m => m.CITY == city && m.DEPT_NO == DEPT_NO/* && m.AREA ==Site*/).ToList();
                db.SECCITY.RemoveRange(g);
                db.SaveChanges();

                for (int b = 0; b < CityCheckbox.Count(); b++)
                {

                    if (userViewModel.SMCPOSTN[a].CITY == CityCheckbox[b])
                    {
                        var c = db.SECCITY.Where(m => m.CITY == city && m.DEPT_NO == DEPT_NO).FirstOrDefault();

                        if (c == null)
                        {
                            foreach (var item in SelectedboxTOWN)
                            {
                                string CityFirstName = item.Substring(0, 3);
                                string Town = item.Substring(3);
                                if (city == CityFirstName)
                                {
                                    s.ADD_DT = DateTime.Now;
                                    s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                                    s.CITY = userViewModel.SMCPOSTN[a].CITY.ToString();
                                    s.TOWN = Town;
                                    //s.AREA = Site;
                                    db.SECCITY.Add(s);
                                    db.SaveChanges();

                                    string newCity2 = "";

                                    switch (city)
                                    {
                                        case "新北市":
                                            newCity2 = "台北縣";
                                            break;
                                        case "高雄市":
                                            newCity2 = "高雄縣";
                                            break;
                                        case "台南市":
                                            newCity2 = "台南縣";
                                            break;
                                        case "台中市":
                                            newCity2 = "台中縣";
                                            break;
                                        case "桃園市":
                                            newCity2 = "桃園縣";
                                            break;
                                    }

                                    if (!string.IsNullOrEmpty(newCity2))
                                    {
                                        s.ADD_DT = DateTime.Now;
                                        s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                                        s.CITY = newCity2;
                                        s.TOWN = Town;
                                        //s.AREA = Site;
                                        db.SECCITY.Add(s);
                                        db.SaveChanges();
                                    }
                                }
                            }
                        }
                        foundMatch = true;
                    }


                }
                if (!foundMatch)
                {
                    try
                    {
                        var c = userViewModel.SMCPOSTN[a].CITY;
                        var t = db.SECCITY.Where(m => m.CITY == city && m.DEPT_NO == DEPT_NO/* && m.AREA ==Site*/).ToList();


                        db.SECCITY.RemoveRange(t);
                        db.SaveChanges();

                        string newCity2 = "";

                        switch (c)
                        {
                            case "新北市":
                                newCity2 = "台北縣";
                                break;
                            case "高雄市":
                                newCity2 = "高雄縣";
                                break;
                            case "台南市":
                                newCity2 = "台南縣";
                                break;
                            case "台中市":
                                newCity2 = "台中縣";
                                break;
                            case "桃園市":
                                newCity2 = "桃園縣";
                                break;
                        }
                        if (!string.IsNullOrEmpty(newCity2))
                        {
                            var R = db.SECCITY.Where(m => m.CITY == newCity2 && m.DEPT_NO == DEPT_NO /*&& m.AREA == Site*/).ToList();
                            db.SECCITY.RemoveRange(R);
                            db.SaveChanges();
                        }
                        TestQuest = "修改成功";
                    }
                    catch { }

                }
            }

            //for (int i = 0; i < 23; i++)
            //{
            //    bool propertyValue = (bool)typeof(CityViewModel).GetProperty("City" + i).GetValue(model, null);
            //    if (propertyValue)
            //    {
            //        for (int z = 0; z < 23; z++)
            //        {
            //            if (i == z)
            //            {
            //                var a = userViewModel.SMCPOSTN[z].CITY;
            //                var b = db.SECCITY.Where(m => m.CITY == a && m.USR_ID == User).FirstOrDefault();
            //                if (b == null)
            //                {
            //                    s.ADD_DT = DateTime.Now;
            //                    s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
            //                    s.CITY = userViewModel.SMCPOSTN[z].CITY.ToString();
            //                    db.SECCITY.Add(s);
            //                    db.SaveChanges();
            //                }
            //            }
            //        }
            //    }
            //    else
            //    {
            //        for (int z = 0; z < 23; z++)
            //        {
            //            if (i == z)
            //            {

            //                try
            //                {
            //                    var a = userViewModel.SMCPOSTN[z].CITY;
            //                    var t = db.SECCITY.Where(m => m.CITY == a && m.USR_ID == User).ToList();
            //                    db.SECCITY.RemoveRange(t);
            //                    db.SaveChanges();
            //                    TestQuest = "修改成功";
            //                }
            //                catch { }
            //            }
            //        }
            //    }
            //}
            return RedirectToAction("DeptEdit", "User", new { choose = Dept, TestQuest = TestQuest });
        }

        public ActionResult UserDelete(string choose)
        {
            var d = db.SECCITY.Where(m => m.USR_ID == choose).ToList();
            db.SECCITY.RemoveRange(d);
            db.SaveChanges();
            return RedirectToAction("Index", "User");
        }
        public ActionResult DeptDelete(string choose)
        {
            var DEPT_Name = db.OGCDEPT.Where(m => m.DEPT_NAME == choose).FirstOrDefault();
            var DEPT_NO = DEPT_Name.DEPT_NO.ToString();
            var d = db.SECCITY.Where(m => m.DEPT_NO == DEPT_NO).ToList();
            db.SECCITY.RemoveRange(d);
            db.SaveChanges();
            return RedirectToAction("Dept", "User");
        }
        public ActionResult DeptCreate(string ID/*, string Save*/, string Dept)
        {

            UserViewModel userViewModel = new UserViewModel();
            userViewModel.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            var DeptID = "";

            //try
            //{
            //    DeptID = db.OGCDEPT.Where(m => m.DEPT_NAME == Dept).FirstOrDefault().DEPT_NO;

            //    if (db.SECCITY.Any(m => m.USR_ID == null && m.AREA == "北區" && m.DEPT_NO == DeptID))
            //    {
            //        ViewBag.North = "Y";
            //    }
            //    if (db.SECCITY.Any(m => m.USR_ID == null && m.AREA == "中區" && m.DEPT_NO == DeptID))
            //    {
            //        ViewBag.Central = "Y";
            //    }
            //    if (db.SECCITY.Any(m => m.USR_ID == null && m.AREA == "南區" && m.DEPT_NO == DeptID))
            //    {
            //        ViewBag.South = "Y";
            //    }
            //    
            //}
            ViewBag.Dept = Dept;

            //catch { }

            try
            {
                var SECCITY_CITY = db.SECCITY.Where(m => m.USR_ID == null && m.DEPT_NO == DeptID).DistinctBy(m => m.CITY).ToList();
                var includedCities = SECCITY_CITY.Select(m => m.CITY);
                ViewBag.includeCities = includedCities;

            }
            catch { }

            //if (Save != null)
            //{
            //    ViewBag.Save = Save + "新增成功";
            //}

            userViewModel.OGCDEPT = db.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            userViewModel.SECCITY = db.SECCITY.ToList();
            return View(userViewModel);
        }

        CityViewModel model = new CityViewModel();
        [HttpPost]
        public ActionResult DeptCreate(string Dept, CityViewModel model, string[] SelectedboxTown/*, string Site*/)
        {

            if (Dept == "") { return RedirectToAction("Index", "Dept"); }
            SECCITY s = new SECCITY();
            var DEPT_Name = db.OGCDEPT.Where(m => m.DEPT_NAME == Dept).FirstOrDefault();
            if (DEPT_Name == null)
            {
                return RedirectToAction("Dept", "User");
            }
            var DEPT_NO = DEPT_Name.DEPT_NO.ToString();
            UserViewModel userViewModel = new UserViewModel();
            userViewModel.SMCPOSTN = db.SMCPOSTN.DistinctBy(m => m.CITY).ToList();
            userViewModel.OGCDEPT = db.OGCDEPT.Where(m => m.DATA_TP == "D").ToList();
            s.DEPT_NO = DEPT_NO;

            for (int i = 0; i < 23; i++)
            {
                // 使用反射獲取屬性值
                bool propertyValue = (bool)typeof(CityViewModel).GetProperty("City" + i).GetValue(model, null);

                // 如果屬性值為true，則顯示屬性名稱
                if (propertyValue)
                {
                    for (int z = 0; z < 23; z++)
                    {
                        if (i == z)
                        {
                            var a = userViewModel.SMCPOSTN[z].CITY;
                            var b = db.SECCITY.Where(m => m.CITY == a && m.DEPT_NO == DEPT_NO).FirstOrDefault();
                            for (int m = 0; m < SelectedboxTown.Length; m++)
                            {
                                var Name = SelectedboxTown[m];
                                string CityNumber = new string(Name.Where(char.IsDigit).ToArray());
                                int CityID = int.Parse(CityNumber);
                                int endIndex = Name.IndexOf(CityNumber) + CityNumber.Length;
                                string textAfterNumber = Name.Substring(endIndex);

                                if (i == CityID)
                                {
                                    if (b == null)
                                    {
                                        s.CITY = userViewModel.SMCPOSTN[z].CITY.ToString();
                                        s.ADD_DT = DateTime.Now;
                                        s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                                        s.TOWN = textAfterNumber;
                                        //s.AREA = Site;
                                        db.SECCITY.Add(s);
                                        db.SaveChanges();

                                        string newCity = "";

                                        switch (s.CITY)
                                        {
                                            case "新北市":
                                                newCity = "台北縣";
                                                break;
                                            case "高雄市":
                                                newCity = "高雄縣";
                                                break;
                                            case "台南市":
                                                newCity = "台南縣";
                                                break;
                                            case "台中市":
                                                newCity = "台中縣";
                                                break;
                                            case "桃園市":
                                                newCity = "桃園縣";
                                                break;
                                        }

                                        if (!string.IsNullOrEmpty(newCity))
                                        {
                                            s.CITY = newCity;
                                            s.ADD_DT = DateTime.Now;
                                            s.UPD_UID = Helpers.WebSiteHelper.getUserId(); //(string)Session["UsrId"];
                                            s.TOWN = textAfterNumber;
                                            //s.AREA = Site;
                                            db.SECCITY.Add(s);
                                            db.SaveChanges();
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
            }
            return RedirectToAction("Dept", "User", new {/*Save = Site,*/ Dept = Dept });
        }
    }
}