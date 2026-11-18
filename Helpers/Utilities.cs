using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Globalization;
using System.Data;
using System.Net;
using System.Net.Mail;
using Newtonsoft.Json.Converters;
using System.Text;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Aspose.Cells;
using Aspose.Words;


namespace WCals.Helpers
{

    public static class DateTimeExtensions
    {
        /// <summary>
        /// To the full taiwan date.
        /// </summary>
        /// <param name="datetime">The datetime.</param>
        /// <returns></returns>
        public static string ToFullTaiwanDate(this DateTime datetime)
        {
            TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

            return string.Format("民國 {0} 年 {1} 月 {2} 日",
                taiwanCalendar.GetYear(datetime),
                datetime.Month,
                datetime.Day);
        }

        /// <summary>
        /// To the simple taiwan date.
        /// </summary>
        /// <param name="datetime">The datetime.</param>
        /// <returns></returns>
        public static string ToSimpleTaiwanDate(this DateTime datetime)
        {
            TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

            return string.Format("{0}/{1}/{2}",
                taiwanCalendar.GetYear(datetime),
                datetime.Month,
                datetime.Day);
        }

        /// <summary>
        /// To the simple taiwan date without slash.
        /// </summary>
        /// <param name="datetime">The datetime.</param>
        /// <returns></returns>
        public static string ToSimpleTaiwanDateWithoutSlash(this DateTime datetime)
        {
            TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

            return string.Format("{0}{1}{2}",
                taiwanCalendar.GetYear(datetime).ToString("000"),
                datetime.ToString("MM"),
                datetime.ToString("dd"));
        }

        public static String ToSimpleTaiwanDateYM(this DateTime datetime)
        {
            TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

            //return String.Format("{0:000}", taiwanCalendar.GetYear(datetime).ToString("000")) + String.Format("{0:00}", datetime.Month);
            return String.Format(taiwanCalendar.GetYear(datetime).ToString("000") + datetime.ToString("MM"));
        }

        public static String FormatTwDate(string twDate, string delim)
        {
            try
            {
                string formattedDate = twDate;
                if (twDate.Length == 7)
                {
                    formattedDate = String.Format("{0}"+ delim + "{1}" + delim  + "{2}", formattedDate.Substring(0, 3), formattedDate.Substring(3, 2), formattedDate.Substring(5, 2));
                }

                return formattedDate;
            }
            catch (Exception e)
            {
                return "";
            }
        }

        public static DateTime TwDateToUsDateTime(string twDate, bool isWithSlash)
        {
            try
            {
                string formattedDate = twDate;
                if (twDate.Length <= 5)
                {
                    formattedDate = formattedDate + "01";
                }
                if (!isWithSlash)
                {
                    formattedDate = String.Format("{0}/{1}/{2}", formattedDate.Substring(0, 3), formattedDate.Substring(3, 2), formattedDate.Substring(5, 2));
                }
                CultureInfo culture = new CultureInfo("zh-TW");
                culture.DateTimeFormat.Calendar = new TaiwanCalendar();

                return DateTime.Parse(formattedDate, culture);
            }
            catch (Exception e)
            {
                return DateTime.Now;
            }
        }

    }

    static class Utilities
    {
        public static string ConvertControlCharToEmpty(string input)
        {

            string output = "";

            foreach (char c in input)//走訪字串的每個字元
            {
                int asciiInt = Convert.ToInt32(c);
                if ((asciiInt >= 0 && asciiInt <= 31) == false)
                {//非控制字元
                    output += c;//就把合法的字元累加到output字串
                }

            }//end foreach

            return output;

        }

        public static string ReplaceLowOrderASCIICharacters(string tmp)
        {
            StringBuilder info = new StringBuilder();
            foreach (char cc in tmp)
            {
                int ss = (int)cc;
                if (((ss >= 0) && (ss <= 8)) || ((ss >= 11) && (ss <= 12)) || ((ss >= 14) && (ss <= 32)))
                    info.AppendFormat(" ", ss);//&#x{0:X};
                else info.Append(cc);
            }
            return info.ToString();
        }
        public static DataTable CopyGenericToDataTable<T>(this IEnumerable<T> items)
        {
            var properties = typeof(T).GetProperties();
            var result = new DataTable();

            //Build the columns
            foreach (var prop in properties)
            {
                Type colType = prop.PropertyType;
                if (colType.IsGenericType && colType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    colType = colType.GetGenericArguments()[0];
                }
                result.Columns.Add(prop.Name, colType);
            }

            //Fill the DataTable
            foreach (var item in items)
            {
                var row = result.NewRow();

                foreach (var prop in properties)
                {
                    var itemValue = prop.GetValue(item, new object[] { });
                    if (itemValue == null)
                    {
                        itemValue = DBNull.Value;
                    }
                    row[prop.Name] = itemValue;
                }

                result.Rows.Add(row);
            }

            return result;
        }

        public static string SafeTrim(this string str)
        {
            if (str == null)
                return null;

            return str.Trim();
        }
        public static decimal SafeDecimal(object data)
        {
             if (data== null)
            {
                return 0;
            }
            return Convert.ToDecimal(data);
        }
    }

    public class DateFormatConverter : IsoDateTimeConverter
    {
        public DateFormatConverter(string format)
        {
            DateTimeFormat = format;
        }
    }

    public class ReportFormatTransfer
    {
        public static void OutputODS(String ReportType, HttpContextBase HttpContext, String FileName, IWorkbook sourceContent)
        {
            if (ReportType == "M")
            {
                sourceContent.SaveAs(FileName + ".xlsx", HttpContext.ApplicationInstance.Response, ExcelDownloadType.Open);
            }
            else
            {
                Aspose.Cells.License license = new Aspose.Cells.License();
                license.SetLicense("Aspose.Total.lic");

                MemoryStream XlsStream = new MemoryStream();
                sourceContent.SaveAs(XlsStream);

                Aspose.Cells.Workbook aspWorkbook = new Aspose.Cells.Workbook(XlsStream);
                XlsStream.Flush();
                XlsStream.Close();
                XlsStream.Dispose();

                Aspose.Cells.HtmlSaveOptions saveOption = new Aspose.Cells.HtmlSaveOptions(Aspose.Cells.SaveFormat.ODS);
                aspWorkbook.Save(HttpContext.ApplicationInstance.Response, FileName + ".ods", Aspose.Cells.ContentDisposition.Attachment, saveOption);
            }
        }


    }
}