using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OneDriveDataRobot.Models
{
    public class DataRobotSetup
    {
        public string SubscriptionId { get; set; }

        public bool Success { get; set; }

        public string Error { get; set; }
        public DateTimeOffset? ExpirationDateTime { get; internal set; }
    }
}