using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Web.Mvc;

namespace WCals.ViewModels.NewModels {
    /// <summary>
    /// SUG0010 建議案管理 - 2個頁籤
    /// 頁籤1: 列管中 (SUG_CLOSE=0)
    /// 頁籤2: 已結案 (SUG_CLOSE=1)
    /// </summary>
    public class SUG0010Model {
        /// <summary>
        /// 當前頁籤 (1-2)
        /// </summary>
        public int tab { get; set; } = 1;

        /// <summary>
        /// 頁籤1分頁
        /// </summary>
        public int page1 { get; set; } = 1;

        /// <summary>
        /// 頁籤2分頁
        /// </summary>
        public int page2 { get; set; } = 1;

        #region 查詢條件
        /// <summary>
        /// 交辦人
        /// </summary>
        public string INSTRUCT_NAME { get; set; } = string.Empty;

        /// <summary>
        /// 建議人
        /// </summary>
        public string SUG_NAME { get; set; } = string.Empty;

        /// <summary>
        /// 建議內容關鍵字
        /// </summary>
        public string SUG_CONTENT { get; set; } = string.Empty;

        /// <summary>
        /// 查詢起始日期
        /// </summary>
        public string START_DATE { get; set; } = string.Empty;

        /// <summary>
        /// 查詢結束日期
        /// </summary>
        public string END_DATE { get; set; } = string.Empty;
        #endregion

        #region 編輯欄位 (SUGMASTER)
        /// <summary>
        /// 案件ID
        /// </summary>
        public decimal SUG_ID { get; set; }

        /// <summary>
        /// 交辦日期
        /// </summary>
        public DateTime? INSTRUCT_DATE { get; set; }

        /// <summary>
        /// 建議日期
        /// </summary>
        public DateTime? SUG_DATE { get; set; }

        /// <summary>
        /// 建議人職稱
        /// </summary>
        public string SUG_JOB { get; set; } = string.Empty;

        /// <summary>
        /// 建議內容附件
        /// </summary>
        public string SUG_FILE { get; set; } = string.Empty;

        /// <summary>
        /// 最終處理內容
        /// </summary>
        public string FINAL_CONTENT { get; set; } = string.Empty;

        /// <summary>
        /// 是否滿意
        /// </summary>
        public bool? AGREE { get; set; }

        /// <summary>
        /// 關鍵字
        /// </summary>
        public string KEYWORDS { get; set; } = string.Empty;

        /// <summary>
        /// 結案狀態 (0=列管中, 1=已結案)
        /// </summary>
        public bool? SUG_CLOSE { get; set; }

        /// <summary>
        /// 建立部門
        /// </summary>
        public string CREATE_DEPT { get; set; } = string.Empty;
        #endregion

        #region 承辦單位 (SUGDETAIL)
        /// <summary>
        /// 主辦單位代碼
        /// </summary>
        public string MAIN_DEPT { get; set; } = string.Empty;

        /// <summary>
        /// 協辦單位代碼 (多選)
        /// </summary>
        public string[] CO_DEPTS { get; set; }

        /// <summary>
        /// 來源部門
        /// </summary>
        public string FROM_DP { get; set; } = string.Empty;

        /// <summary>
        /// 承辦部門
        /// </summary>
        public string TO_DP { get; set; } = string.Empty;

        /// <summary>
        /// 回覆內容
        /// </summary>
        public string REPLY_CONTENT { get; set; } = string.Empty;        public string RENEW_CONTENT { get; set; } = string.Empty;
        #endregion

        #region 資料列表
        /// <summary>
        /// 頁籤1: 列管中
        /// </summary>
        public IPagedList<DataRow> List_Tab1 { get; set; }

        /// <summary>
        /// 頁籤2: 已結案
        /// </summary>
        public IPagedList<DataRow> List_Tab2 { get; set; }

        /// <summary>
        /// 承辦單位列表
        /// </summary>
        public DataTable DetailList { get; set; }

        /// <summary>
        /// 附件列表
        /// </summary>
        public DataTable FileList { get; set; }
        #endregion
    }
}
