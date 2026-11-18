using PagedList;
using System;
using System.ComponentModel.DataAnnotations;
using System.Web;
using System.Web.Mvc;

namespace WCals.ViewModels
{
    /// <summary>
    /// 最新消息維護-顯示項目
    /// </summary>
    public class NewsListItemViewModel
    {
        /// <summary>
        /// 是否公告中
        /// </summary>
        public string IsAnnounced { get; set; }
        /// <summary>
        /// 編號
        /// </summary>
        public int NEWSID { get; set; }
        /// <summary>
        /// 訊息類別代碼
        /// </summary>
        public string TYPE { get; set; }
        /// <summary>
        /// 訊息類別名稱
        /// </summary>
        public string TYPE_NAME { get; set; }
        /// <summary>
        /// 標題
        /// </summary>
        public string HEADLINE { get; set; }
        /// <summary>
        /// 內容
        /// </summary>
        public string SUMMARY { get; set; }
        /// <summary>
        /// 開始日期
        /// </summary>
        public DateTime? STARTDATE { get; set; }
        /// <summary>
        /// 結束日期
        /// </summary>
        public DateTime? ENDDATE { get; set; }
    }

    /// <summary>
    /// 最新消息維護-查詢條件
    /// </summary>
    public class NewsSearchViewModel
    {
        /// <summary>
        /// 訊息類別
        /// </summary>
        public string TYPE { get; set; }
        /// <summary>
        /// 標題
        /// </summary>
        public string HEADLINE { get; set; }
        /// <summary>
        /// 內容
        /// </summary>
        public string SUMMARY { get; set; }
        /// <summary>
        /// 開始日期
        /// </summary>
        public string STARTDATE { get; set; }
        /// <summary>
        /// 結束日期
        /// </summary>
        public string ENDDATE { get; set; }
    }

    /// <summary> 
    /// 最新消息維護-查詢資料與結果
    /// </summary>
    public class NewsListViewModel
    {
        /// <summary>
        /// 最新消息維護-查詢結果清單
        /// </summary>
        public IPagedList<NewsListItemViewModel> NewsList { get; set; }

        /// <summary>
        /// 最新消息維護-查詢條件
        /// </summary>
        public NewsSearchViewModel NewsSearch { get; set; }
    }

    /// <summary>
    /// 最新消息維護-新增/修改項目
    /// </summary>
    public class NewsDataViewModel
    {
        public int NEWSID { get; set; }

        [Required()]
        [Display(Name = "訊息類別")]
        public string TYPE { get; set; }

        [Required()]
        [MaxLength(50)]
        [Display(Name = "標題")]
        public string HEADLINE { get; set; }

        [Required()]
        [MaxLength(1000)]
        [Display(Name = "內容")]
        [AllowHtml]
        public string BODY { get; set; }

        [Required()]
        [Display(Name = "開始日期")]
        public DateTime? STARTDATE { get; set; }

        [Required()]
        [Display(Name = "結束日期")]
        public DateTime? ENDDATE { get; set; }

        /// <summary>
        /// 上傳檔案清單
        /// </summary>
        public HttpPostedFileBase[] UploadedFiles { get; set; }

        /// <summary>
        /// 上傳檔案ID清單
        /// </summary>
        public string[] FID { get; set; } = Array.Empty<string>();
    }
}