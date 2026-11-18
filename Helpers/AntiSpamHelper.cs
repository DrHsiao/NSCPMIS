using System;
using System.Web.Mvc;

namespace WCals.Helpers
{
    public static class AntiSpamHelper
    {
        /// <summary>
        /// 驗證是否為機器人提交表單(初步防止)
        /// </summary>
        /// <param name="controller"></param>
        /// <param name="honeypotValue"> 隱藏欄位 </param>
        /// <param name="formLoadTicks"> 填表時間 </param>
        /// <param name="minDelaySeconds"> 延遲檢查時間(預設3秒) </param>
        /// <returns></returns>
        public static bool IsBot(this Controller controller, string honeypotValue, long formLoadTicks, int minDelaySeconds = 3)
        {
            // 1. 若隱藏欄位(Honeypot)有值，視為 bot
            if (!string.IsNullOrWhiteSpace(honeypotValue))
                return true;

            // 2. 若填表時間過短，視為 bot
            var now = DateTime.UtcNow;
            var loadedTime = new DateTime(formLoadTicks);
            var elapsedSeconds = (now - loadedTime).TotalSeconds;

            return elapsedSeconds < minDelaySeconds;
        }
    }
}