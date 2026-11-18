using System;
using System.Configuration; // 必須引用此命名空間

namespace NSEPMIS.Helpers
{
    public static class EnvHelper
    {
        // 快取結果，提升效能
        private static bool? _isTestMode = null;

        /// <summary>
        /// 判斷是否為測試環境
        /// </summary>
        public static bool IsTestMode
        {
            get
            {
                if (_isTestMode.HasValue)
                {
                    return _isTestMode.Value;
                }

                _isTestMode = CheckIsTestMode();
                return _isTestMode.Value;
            }
        }

        private static bool CheckIsTestMode()
        {
            try
            {
                // 直接讀取 connections.config 裡面的 EnvironmentMode
                var envConfig = ConfigurationManager.ConnectionStrings["EnvironmentMode"];

                if (envConfig != null && !string.IsNullOrWhiteSpace(envConfig.ConnectionString))
                {
                    // 若設定為 TEST，則回傳 true
                    return envConfig.ConnectionString.Trim().ToUpper() == "TEST";
                }
            }
            catch (Exception)
            {
                // 讀取失敗時，預設為正式機以防萬一
            }

            // 若沒設定或讀不到，預設為 false (正式機)
            return false;
        }
    }
}