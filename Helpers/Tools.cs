using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using WCals.classes.john;

namespace WCals.Source
{
    public class Tools
    {
        //public static double IfNull0(string i)
        //{
        //    if (!string.IsNullOrEmpty(i))
        //    {
        //        double d = double.Parse(i);
        //        return d;
        //    }
        //    else
        //    {
        //        return 0;
        //    }
        //}

        public static decimal ToDecimal(object value)
        {
            if (value.ToString() == "")   //////怕跑出警告所以在value後面加上了 ToString()，如果又跑出ToDecimal錯誤，請把 ToString()拿掉
            {
                return 0;
            }

            if (value != null)
            {
                return ((IConvertible)value).ToDecimal(null);
            }



            return 0m;
        }



        public static double IfFormula(string an, string ap, string ao)
        {
            double apValue = 0;
            double aoValue = 0;

            if (!string.IsNullOrEmpty(an) && double.TryParse(ao, out aoValue))
            {
                if (aoValue != 0)
                {
                    return aoValue / 1000;
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(ap) && double.TryParse(ap, out apValue))
                {
                    return apValue / 1000;
                }
            }

            return 0;
        }
        public static double StringToDouble(string i)
        {
            double rtn = 0;
            if (string.IsNullOrEmpty(i))
            {                
                return rtn;
            }
            else
            {
                rtn = double.Parse(i);
                return rtn;
            }

        }
        public static int StringToInt(string i)
        {
            int rtn = 0;
            if (string.IsNullOrEmpty(i))
            {
                return rtn;
            }
            else
            {
                rtn = int.Parse(i);
                return rtn;
            }

        }

        public static decimal StringToDecimal(string i)
        {
            decimal rtn = 0;
            if (Decimal.TryParse(i, out rtn))
            {
                return rtn;
            }
            else
            {
                return rtn;
            }

        }
        public static double thousandchange1(string i)
        {
            if (string.IsNullOrEmpty(i))
            {
                // 如果 AF5 是空的，返回空值
                return 0;
            }
            else
            {
                // 如果 AF5 不是空的，將其轉換為數字，除以 1000，並返回結果
                double result = double.Parse(i) / 1000;
                double roundedResult = Math.Round(result, MidpointRounding.AwayFromZero); 
                return roundedResult;
            }
        }

        public static double thousandchange3(string i)
        {
            if (string.IsNullOrEmpty(i))
            {
                // 如果 AF5 是空的，返回空值
                return 0;
            }
            else
            {
                // 如果 AF5 不是空的，將其轉換為數字，除以 1000，並返回結果
                double result = double.Parse(i) / 1000;
                
                return result;
            }
        }

        public static void SetCellstyle(Aspose.Cells.Worksheet ws, string cellAddress)
        {
            var cell = ws.Cells[cellAddress];
            Aspose.Cells.Style cellStyle = cell.GetStyle();
            cellStyle.Custom = "#,##0";
            cell.SetStyle(cellStyle);
        }

        public static double IsNumberOrNot(double i)
        {
            if (i == 0)
            {
                // 如果 AF5 是空的，返回空值
                return 0;
            }
            else
            {
                // 如果 AF5 不是空的，將其轉換為數字並返回結果
                double result = i;
                return result;
            }
        }

        public static string IsNumberOrNot_String(string i)
        {
            if( i == "")
            {
                return "0";
            }
            else if (i == "0")
            {
                // 如果 AF5 是空的，返回空值
                return "0";
            }
            else
            {
                // 如果 AF5 不是空的，將其轉換為數字並返回結果
                string result = i;
                return result;
            }
        }

        public static double thousandchange2(double i)
        {
            if (i == 0)
            {
                // 如果 AF5 是空的，返回空值
                return 0;
            }
            else
            {
                // 如果 AF5 不是空的，將其轉換為數字，除以 1000，並返回結果
                double result = i / 1000;
                return result;
            }
        }


        public static string FormatDate(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                // 如果 input 是空的，返回空值
                return "";
            }
            else if (input.Length == 7 && input.StartsWith("0"))
            {
                // 如果 input 的長度是 8 且第一個字符是 0，按照 "0xxyyzz" 的格式進行處理
                string year = input.Substring(1, 2);
                string month = input.Substring(3, 2);
                string day = input.Substring(5, 2);
                return string.Format("{0}.{1}.{2}", year, month, day);
            }
            else if (input.Length == 7)
            {
                // 如果 input 的長度是 7，按照 "xxyyzz" 的格式進行處理
                string year = input.Substring(0, 3);
                string month = input.Substring(3, 2);
                string day = input.Substring(5, 2);
                return string.Format("{0}.{1}.{2}", year, month, day);
            }
            else
            {
                // 如果 input 的長度不是 7，返回空值
                return "";
            }
        }

        public static string GetLoginUserCity(string userId)
        {
            // 1. 取得目前登入者的 ID (從 FormsAuthentication Cookie 解析出來的)
            //string userId = User.Identity.Name;

            if (string.IsNullOrEmpty(userId))
            {
                return null;
            }

            // 2. 查詢該使用者的單位名稱
            // 邏輯：OGCUSER -> OGCDPUR -> OGCDEPT -> DEPT_NAME
            string sql = @"
        SELECT TOP 1 d.SMP_DEPT_NAME 
        FROM OGCDEPT d
        JOIN OGCDPUR r ON d.DEPT_NO = r.DEPT_NO
        WHERE r.USR_ID = @USR_ID";

            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add(new SqlParameter("USR_ID", userId));

            DataTable dt = BASE_DAO.SqlDao.CMBASEQuery4List(cmd); // 注意：這裡是查 CMBASE

            if (dt.Rows.Count > 0)
            {
                string deptName = dt.Rows[0]["SMP_DEPT_NAME"].ToString();

                // 3. 判斷單位名稱是否包含「縣」或「市」，並回傳縣市名稱
                // 這裡需要根據貴單位的 DEPT_NAME 規則來寫
                // 例如：如果是 "宜蘭縣政府"，就回傳 "宜蘭縣"
                // 如果是 "營建署" (中央)，就回傳 null 或 ""

                if (deptName.Contains("縣政府"))
                {
                    return deptName.Replace("政府", ""); // "宜蘭縣政府" -> "宜蘭縣"
                }
                else if (deptName.EndsWith("市政府"))
                {
                    return deptName.Replace("政府", ""); // "高雄市政府" -> "高雄市"
                }
                else if (deptName.EndsWith("縣") || deptName.EndsWith("市"))
                {
                    // 如果單位名稱本身就是 "宜蘭縣" 或 "基隆市"
                    return deptName;
                }
            }

            // 如果是中央單位 (如營建署) 或查無資料，回傳 null
            return null;
        }

    }
}