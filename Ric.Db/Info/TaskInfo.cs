using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Model;

namespace Ric.Db.Info
{
    public class TaskInfo
    {
        public int TaskId { get; set; }
        public string TaskName { get; set; }
        public int MarketId { get; set; }
        public string MarketName { get; set; }
        public string MailToRecipients { get; set; }
        public string MailCcRecipients { get; set; }
        public string ConfigFile { get; set; }
        public string ConfigType { get; set; }
        public string GeneratorType { get; set; }

        public string GroupName { get; set; }
        public string Description { get; set; }
        public TaskStatus Status { get; set; }
    }

    public class DocumentElement
    {
        public List<TaskInfo> taskInfo { get; set; }
    }
}
