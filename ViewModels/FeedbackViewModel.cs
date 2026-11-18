using PagedList;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web;
using WCals.Models.CMEMQAModels;

namespace WCals.ViewModels
{
    /// <summary>
    /// 使用者意見紀錄-顯示項目
    /// </summary>
    public class FeedbackListItemViewModel
    {
        /// <summary>
        /// 意見編號
        /// </summary>
        public string MAIN_ID { get; set; }
        /// <summary>
        /// 內容類別代碼
        /// </summary>
        public string APL_TYPE { get; set; }
        /// <summary>
        /// 內容類別顯示名稱
        /// </summary>
        public string APL_TYPE_NAME { get; set; }
        /// <summary>
        /// 問題主旨
        /// </summary>
        public string QUE_HEAD { get; set; }
        /// <summary>
        /// 問題內容
        /// </summary>
        public string QUE_DATA { get; set; }
        /// <summary>
        /// 文章人氣
        /// </summary>
        public decimal OPEN_PEP { get; set; }
        /// <summary>
        /// 回覆數
        /// </summary>
        public decimal RETURN_PEP { get; set; }
        /// <summary>
        /// 發問者
        /// </summary>
        public string QUE_USER { get; set; }
        /// <summary>
        /// 建立日期
        /// </summary>
        public DateTime? ADD_DT { get; set; }
    }

    /// <summary>
    /// 使用者意見紀錄-查詢條件
    /// </summary>
    public class FeedbackSearchViewModel
    {
        /// <summary>
        /// 內容類別
        /// </summary>
        public string APL_TYPE { get; set; }
        /// <summary>
        /// 問題主旨
        /// </summary>
        public string QUE_HEAD { get; set; }
        /// <summary>
        /// 問題內容
        /// </summary>
        public string QUE_DATA { get; set; }
        /// <summary>
        /// 發問日期–起
        /// </summary>
        public string STARTDATE { get; set; }
        /// <summary>
        /// 發問日期–迄
        /// </summary>
        public string ENDDATE { get; set; }
    }

    /// <summary>
    /// 使用者意見紀錄-查詢資料與結果
    /// </summary>
    public class FeedbackListViewModel
    {
        /// <summary>
        /// 使用者意見紀錄-查詢結果清單
        /// </summary>
        public IPagedList<FeedbackListItemViewModel> FeedbackList { get; set; }
        /// <summary>
        /// 使用者意見紀錄-查詢條件
        /// </summary>
        public FeedbackSearchViewModel FeedbackSearch { get; set; }
    }

    /// <summary>
    /// 使用者意見紀錄-詳細資料
    /// </summary>
    public class FeedbackDetailsViewModel
    {
        /// <summary>
        /// 使用者意見紀錄-留言資料
        /// </summary>
        public FeedbackListItemViewModel FeedbackDetails { get; set; }
        /// <summary>
        /// 使用者意見紀錄-檔案資料
        /// </summary>
        public List<MQAFILE> FeedbackFileList { get; set; }
        /// <summary>
        /// 使用者意見紀錄-回覆資料
        /// </summary>
        public List<MQAFORM> FeedbackReplyList { get; set; }
        /// <summary>
        /// 目前登入帳號是否有權限回覆意見
        /// </summary>
        public bool UserCanReply { get; set; }
    }

    /// <summary>
    /// 使用者意見紀錄-建立留言項目
    /// </summary>
    public class FeedbackDataViewModel
    {
        [Required()]
        [Display(Name = "內容類別")]
        public string APL_TYPE { get; set; }

        [Required()]
        [MaxLength(50)]
        [Display(Name = "問題主旨")]
        public string QUE_HEAD { get; set; }

        [Required()]
        [MaxLength(1000)]
        [Display(Name = "發問內容")]
        public string QUE_DATA { get; set; }

        [Required()]
        [Display(Name = "E-mail")]
        [RegularExpression(@"^[^@\s]+@[^@\s]+\.[^@\s]+$", ErrorMessage = "請輸入正確的電子郵件地址")]
        public string USER_MAIL { get; set; }

        [Required()]
        [Display(Name = "連絡電話")]
        [RegularExpression(@"^(09\d{8}|0\d{1,2}-\d{6,8}(#\d{1,5})?)$", ErrorMessage = "請輸入有效的電話號碼")]
        public string USER_TEL { get; set; }

        public List<FeedbackFileUploadViewModel> FileUploads { get; set; } = new List<FeedbackFileUploadViewModel>();
    }

    /// <summary>
    /// 使用者意見紀錄-檔案上傳項目
    /// </summary>
    public class FeedbackFileUploadViewModel
    {
        /// <summary>
        /// 檔案資料
        /// </summary>
        public HttpPostedFileBase File { get; set; }
        /// <summary>
        /// 檔案描述
        /// </summary>
        public string FILE_DESC { get; set; }
    }
}