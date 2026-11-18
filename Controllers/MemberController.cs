using System;
using System.Collections.Generic;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;

namespace WCals.Controllers
{
    public class MemberController : Controller
    {

    }
    public static class DistinctByClass //Distinct 可以直接選取欄位的方法   ///////////////////通常如果這裡報錯請查看資料庫是否資料錯誤，連線資料庫等等  並不是此函式錯誤。  /////////////////
    {
        public static IEnumerable<TSource> DistinctBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            HashSet<TKey> seenKeys = new HashSet<TKey>();
            foreach (TSource element in source)
            {
                if (seenKeys.Add(keySelector(element)))
                {
                    yield return element;
                }
            }
        }
    }
    
    public static class SourceFunction
    {
        public static DataTable DataTableGet(string FSQLStr, string connectionString)
        {
            SqlConnection con1 = new SqlConnection();
            string constr = connectionString;
            con1.ConnectionString = constr;
            SqlDataAdapter adp = new SqlDataAdapter(FSQLStr, con1);
            DataSet ds = new DataSet();
            adp.Fill(ds);
            DataTable IAMDATATABLE = ds.Tables[0];
            return IAMDATATABLE;
        }
    }
}