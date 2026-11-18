using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Newtonsoft.Json;
using System.Web;

namespace WCals.Helpers
{
    public class TenderHelper
    {
        //public static String getPJDYPACJsonString(String PJ_CODE, String E_CODE, Boolean getHistory, int Version, String selYear)
        //{
        //    Models.WCalsEntities db = new Models.WCalsEntities();
        //    //string NewPJ_CODE = E_CODE+ PJ_CODE.Substring(2, 3);
        //    var pjdypac = (from f in db.PJV_PJDYPAC
        //                   where f.PJ_CODE == PJ_CODE && f.E_CODE== E_CODE
        //                   //where f.PJ_CODE == NewPJ_CODE
        //                   orderby f.SYR, f.ACC_ID
        //                   select new { f.AID, f.PJ_CODE, f.SYR, f.E_CODE, f.ACC_ID, f.ACC_NAME, f.ACC_FUND, f.USED_FUND, METHOD = "original" }).ToList();

        //    if (getHistory== true)
        //    {
        //        pjdypac = (from f in db.PJV_PJHYPAC
        //                   where f.PJ_CODE == PJ_CODE && f.E_CODE== E_CODE && f.SEQ_NO== Version
        //                   orderby f.SYR, f.ACC_ID
        //                   select new { f.AID, f.PJ_CODE, f.SYR, f.E_CODE, f.ACC_ID, f.ACC_NAME, f.ACC_FUND, f.USED_FUND, METHOD = "original" }).ToList();
        //    }

        //    if (selYear == null)
        //    {
        //        return Newtonsoft.Json.JsonConvert.SerializeObject(pjdypac);
        //    }
        //    else
        //    {
        //        return Newtonsoft.Json.JsonConvert.SerializeObject(pjdypac.Where(f => f.SYR == selYear));
        //        //model.jsonPJDFDND = JsonConvert.SerializeObject(pjdfdnd.Where(f => f.SYR == selYear));
        //    }

        //    //return Newtonsoft.Json.JsonConvert.SerializeObject(pjdypac);
        //}

        //public static String getWeatherName(String WeatherCode)
        //{
        //    Models.WCalsEntities db = new Models.WCalsEntities();

        //    var smccode = (from f in db.SMCCODE
        //                   where f.SYS_ID == "SM" && f.CATY_ID== WeatherCode
        //                   select f).FirstOrDefault();
        //    if (smccode!= null)
        //    {
        //        return smccode.CATY_COD;
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        /// <summary>
        /// 提供開始年月及結束年月的相減列表
        /// </summary>
        /// <returns>List</returns>
        public static List<SelectListItem> getSE_YRList(String SYR, String EYR)
        {
            List<SelectListItem> itemKindList = new List<SelectListItem>();
            if (String.IsNullOrEmpty(SYR) == false && String.IsNullOrEmpty(EYR) == false)
            {
                int syrs = Convert.ToInt16(SYR.Substring(0, 3));
                int eyrs = Convert.ToInt16(EYR.Substring(0, 3));
                for (int i = syrs; i <= eyrs; i++)
                {
                    itemKindList.Add(new SelectListItem()
                    {
                        Text = i >= 100 ? i.ToString() : "0" + i.ToString(),
                        Value = i >= 100 ? i.ToString() : "0" + i.ToString(),
                        Selected = false
                    });
                }
            }
            else
            {
                int yrs = DateTime.Now.Year;
                yrs = yrs - 1911;
                itemKindList.Add(new SelectListItem()
                {
                    Text = yrs.ToString(),
                    Value = yrs.ToString(),
                    Selected = false
                });
            }

            return itemKindList;
        }

        /// <summary>
        /// 提供代碼維護內容，新增類別後再進行系統修正
        /// </summary>
        /// <returns>List</returns>
        public static List<SelectListItem> getCodeList()
        {
            List<SelectListItem> itemKindList = new List<SelectListItem>();
            itemKindList.Add(new SelectListItem()
            {
                Text = "供水系統別",
                Value = "dbO_WATERRANGE",
                Selected = false
            });
            itemKindList.Add(new SelectListItem()
            {
                Text = "機關代碼",
                Value = "SMCOGLS",
                Selected = false
            });
            itemKindList.Add(new SelectListItem()
            {
                Text = "單位部門代碼維護",
                Value = "OGCDEPT",
                Selected = false
            });
            itemKindList.Add(new SelectListItem()
            {
                Text = "縣市代碼維護",
                Value = "SMCPOSN",
                Selected = false
            });
            itemKindList.Add(new SelectListItem()
            {
                Text = "科目代碼維護",
                Value = "SMCCODE@AC@02",
                Selected = false
            });
            return itemKindList;
        }

        //public static String getPJ_MODEName(String PJ_MODE)
        //{
        //    Models.WCalsEntities db = new Models.WCalsEntities();
        //    if (String.IsNullOrEmpty(PJ_MODE)== false)
        //    {
        //        return db.SMCCODE.Where(f => f.SYS_ID == "PJ").Where(f => f.CATY_ID == "00").Where(f => f.CATY_COD == PJ_MODE).FirstOrDefault().CODE_DESC;

        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        public static List<SelectListItem> getItemKind()
        {
            List<SelectListItem> itemKindList = new List<SelectListItem>();
            itemKindList.Add(new SelectListItem()
            {
                Text = "固定單價",
                Value = "G",
                Selected = true
            });
            itemKindList.Add(new SelectListItem()
            {
                Text = "單價分析",
                Value = "A",
                Selected = false
            });
            itemKindList.Add(new SelectListItem()
            {
                Text = "變動單價",
                Value = "V",
                Selected = false
            });

            return itemKindList;
        }

        public static List<SelectListItem> getCarry()
        {
            List<SelectListItem> carryList = new List<SelectListItem>();
            carryList.Add(new SelectListItem()
            {
                Text = "四捨五入",
                Value = "0",
                Selected = false
            });
            carryList.Add(new SelectListItem()
            {
                Text = "無條件進位",
                Value = "1",
                Selected = false
            });
            carryList.Add(new SelectListItem()
            {
                Text = "無條件捨去",
                Value = "2",
                Selected = false
            });
            return carryList;
        }

        public static List<SelectListItem> getCalculateDecimal()
        {
            List<SelectListItem> calculateDecimalList = new List<SelectListItem>();
            calculateDecimalList.Add(new SelectListItem()
            {
                Text = "分",
                Value = "0",
                Selected = false
            });

            calculateDecimalList.Add(new SelectListItem()
            {
                Text = "角",
                Value = "1",
                Selected = false
            });
            calculateDecimalList.Add(new SelectListItem()
            {
                Text = "元",
                Value = "2",
                Selected = false
            });

            calculateDecimalList.Add(new SelectListItem()
            {
                Text = "十",
                Value = "3",
                Selected = false
            });
            calculateDecimalList.Add(new SelectListItem()
            {
                Text = "百",
                Value = "4",
                Selected = false
            });
            calculateDecimalList.Add(new SelectListItem()
            {
                Text = "千",
                Value = "5",
                Selected = false
            });
            return calculateDecimalList;
        }

        public static List<SelectListItem> getCanFgItems()
        {
            List<SelectListItem> canFgItemList = new List<SelectListItem>();
            canFgItemList.Add(new SelectListItem()
            {
                Text = "是",
                Value = "是",
                Selected = false
            });
            return canFgItemList;
        }

        public static List<SelectListItem> getBinaryItems()
        {
            List<SelectListItem> binaryItemList = new List<SelectListItem>();
            binaryItemList.Add(new SelectListItem()
            {
                Text = "是",
                Value = "true",
                Selected = false
            });
            binaryItemList.Add(new SelectListItem()
            {
                Text = "否",
                Value = "false",
                Selected = false
            });

            return binaryItemList;
        }

        //public static List<SelectListItem> getDefaultYears()
        //{
        //    Models.WCalsEntities db = new Models.WCalsEntities();

        //    List<SelectListItem> result = new List<SelectListItem>();

        //    var apcyear = (from f in db.SMV_APCYEAR
        //                   orderby f.SYR
        //                   where f.SYR != "" && f.SYR != null && f.SYR.Length >2
        //                   select f).ToList();

        //    foreach (Models.SMV_APCYEAR item in apcyear)
        //    {
        //        result.Add(new SelectListItem()
        //        {
        //            Text = item.SYR,
        //            Value = item.SYR,
        //            Selected = (Convert.ToInt32(item.SYR) == DateTime.Now.Year) ? true : false
        //            //Selected = false
        //        });
        //    }
        //    return result;

        //}

        public static List<SelectListItem> getDefaultMonthes()
        {
            List<SelectListItem> defaultMonthes = new List<SelectListItem>();
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "1月",
                Value = "01",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "2月",
                Value = "02",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "3月",
                Value = "03",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "4月",
                Value = "04",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "5月",
                Value = "05",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "6月",
                Value = "06",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "7月",
                Value = "07",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "8月",
                Value = "08",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "9月",
                Value = "09",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "10月",
                Value = "10",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "11月",
                Value = "11",
                Selected = false
            });
            defaultMonthes.Add(new SelectListItem()
            {
                Text = "12月",
                Value = "12",
                Selected = false
            });

            return defaultMonthes;
        }

        /*public static List<SelectListItem> getSMCCATYOptions()
        {
            Models.WCalsEntities db = new Models.WCalsEntities();

            List<SelectListItem> result = new List<SelectListItem>();

            var smccaty = (from f in db.SMCCATY
                           select f).ToList();

            foreach (Models.SMCCATY item in smccaty)
            {
                result.Add(new SelectListItem()
                {
                    Text = item.CATY_DESC,
                    Value = item.SYS_ID + "@" + item.CATY_ID,
                    Selected = false
                });
            }

            if (result.Count > 0)
            {
                result[0].Selected = true;
            }

            return result;
        }*/

        public static List<SelectListItem> getAdministratorDepts()
        {
            Models.CMBASEModels.CMBASEEntities db = new Models.CMBASEModels.CMBASEEntities();

            List<SelectListItem> result = new List<SelectListItem>();

            var ogcdept = (from f in db.OGCDEPT
                           where f.DEPT_NO.StartsWith("000")
                           select f).ToList();

            foreach (Models.CMBASEModels.OGCDEPT item in ogcdept)
            {
                result.Add(new SelectListItem()
                {
                    Text = item.DEPT_NAME,
                    Value = item.DEPT_NO,
                    Selected = false
                });
            }

            return result;
        }

        public static List<SelectListItem> getItemAreasOptions()
        {
            List<SelectListItem> itemAreasOptions = new List<SelectListItem>();
            itemAreasOptions.Add(new SelectListItem()
            {
                Text = "北",
                Value = "北",
                Selected = false
            });
            itemAreasOptions.Add(new SelectListItem()
            {
                Text = "中",
                Value = "中",
                Selected = false
            });
            itemAreasOptions.Add(new SelectListItem()
            {
                Text = "南",
                Value = "南",
                Selected = false
            });
            itemAreasOptions.Add(new SelectListItem()
            {
                Text = "東",
                Value = "東",
                Selected = false
            });

            return itemAreasOptions;
        }

        public static List<SelectListItem> getItemActionTypeOptions()
        {
            List<SelectListItem> actionTypeOptions = new List<SelectListItem>();
            actionTypeOptions.Add(new SelectListItem()
            {
                Text = "登入",
                Value = "L",
                Selected = false
            });
            actionTypeOptions.Add(new SelectListItem()
            {
                Text = "功能",
                Value = "F",
                Selected = false
            });

            return actionTypeOptions;
        }

        public static List<SelectListItem> getCostKindOptionsForV()
        {
            List<SelectListItem> costKindList = new List<SelectListItem>();
            costKindList.Add(new SelectListItem()
            {
                Text = "變動單價",
                Value = "$",
                Selected = true
            });
            costKindList.Add(new SelectListItem()
            {
                Text = "百分比",
                Value = "%",
                Selected = false
            });
            costKindList.Add(new SelectListItem()
            {
                Text = "機具百分比",
                Value = "E%",
                Selected = false
            });
            costKindList.Add(new SelectListItem()
            {
                Text = "人工百分比",
                Value = "L%",
                Selected = false
            });
            costKindList.Add(new SelectListItem()
            {
                Text = "材料百分比",
                Value = "M%",
                Selected = false
            });
            costKindList.Add(new SelectListItem()
            {
                Text = "雜項",
                Value = "W",
                Selected = false
            });


            return costKindList;
        }

        public static List<SelectListItem> getCostKindOptionsForG()
        {
            List<SelectListItem> costKindOptions = new List<SelectListItem>();
            costKindOptions.Add(new SelectListItem()
            {
                Text = "機具",
                Value = "E",
                Selected = true
            });
            costKindOptions.Add(new SelectListItem()
            {
                Text = "人工",
                Value = "L",
                Selected = false
            });
            costKindOptions.Add(new SelectListItem()
            {
                Text = "材料",
                Value = "M",
                Selected = false
            });
            costKindOptions.Add(new SelectListItem()
            {
                Text = "雜項",
                Value = "W",
                Selected = false
            });

            return costKindOptions;
        }

        public static List<SelectListItem> getCostKindOptionsForA()
        {
            List<SelectListItem> costKindOptions = new List<SelectListItem>();
            costKindOptions.Add(new SelectListItem()
            {
                Text = "單價分析",
                Value = "A",
                Selected = true
            });

            return costKindOptions;
        }

        public static string FormatTD_CODE(string TD_CODE)
        {
            if (TD_CODE == null)
                return null;
            //if (

            //return TD_CODE.Substring(0,2) + '-' + ;
            if (TD_CODE.Length >= 13)
            {
                return TD_CODE.Substring(0, 2) +"-"+ TD_CODE.Substring(2, 3) +"-" + TD_CODE.Substring(5, 4) + "-"+ TD_CODE.Substring(9, 4);
            } else if (TD_CODE.Length == 10)
            {
                return TD_CODE.Substring(0, 2) + "-" + TD_CODE.Substring(2, 2) + "-" + TD_CODE.Substring(4, 4) + "-" + TD_CODE.Substring(8, 2);
            } else
            {
                return TD_CODE;
            }



            
        }
    }
}