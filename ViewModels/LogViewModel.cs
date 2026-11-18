using PagedList;
using System;

namespace WCals.ViewModels
{
    /// <summary>
    /// 系統操作紀錄-顯示項目
    /// </summary>
    public class LogListItemViewModel
    {
        /// <summary>
        /// 使用者帳號
        /// </summary>
        public string USR_ID { get; set; }
        /// <summary>
        /// 使用者名稱
        /// </summary>
        public string USR_NAME { get; set; }
        /// <summary>
        /// 操作說明
        /// </summary>
        public string ACTION_DESC { get; set; }
        /// <summary>
        /// 操作日期
        /// </summary>
        public DateTime? ACTION_DT { get; set; }
        /// <summary>
        /// 最後登入日期
        /// </summary>
        public DateTime? LASTEST_LOGIN_DT { get; set; }
    }

    /// <summary> 系統操作紀錄-查詢條件 </summary>
    public class LogListSearchViewModel
    {
        /// <summary>
        /// 使用者帳號
        /// </summary>
        public string USR_ID { get; set; }
        /// <summary>
        /// 使用者名稱
        /// </summary>
        public string USR_NAME { get; set; }
        /// <summary>
        /// 操作日期-起
        /// </summary>
        public string ACTION_DT_START { get; set; }
        /// <summary>
        /// 操作日期-迄
        /// </summary>
        public string ACTION_DT_END { get; set; }
    }

    /// <summary>
    /// 系統操作記錄-查詢資料與結果
    /// </summary>
    public class LogListViewModel
    {
        public IPagedList<LogListItemViewModel> LogList { get; set; }

        public LogListSearchViewModel LogSearch { get; set; }
    }
}