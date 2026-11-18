using System;

namespace WCals.Exceptions
{
    public class BusinessException : Exception
    {
        /// <summary>
        /// 初始化 BusinessException 類別的新執行個體
        /// </summary>
        public BusinessException() { }

        /// <summary>
        /// 驗證邏輯錯誤異常類別
        /// </summary>
        /// <param name="message"> 自訂錯誤訊息 </param>
        public BusinessException(string message) : base(message) { }

        /// <summary>
        /// 驗證邏輯錯誤異常類別，並保留原始例外資訊以利追蹤問題
        /// </summary>
        /// <param name="message"> 描述錯誤原因的訊息 </param>
        /// <param name="innerException"> 原始發生的例外 </param>
        public BusinessException(string message, Exception innerException)
        : base(message, innerException) { }
    }
}