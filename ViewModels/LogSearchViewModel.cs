using System;

namespace WCals.ViewModels
{
    public class LogSearchViewModel
    {
        public String ACTION_TYPE { get; set; }

        public String USR_ID { get; set; }

        public String LOGIN_IP { get; set; }

        public DateTime? LOGIN_DT_START { get; set; }

        public DateTime? LOGIN_DT_END { get; set; }
    }
}