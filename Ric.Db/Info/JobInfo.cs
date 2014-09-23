using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Info
{
    public class JobInfo
    {
        public int JobId { get; set; }
        public string JobName { get; set; }
        public string JobSequence { get; set; }
        public string AssignTo { get; set; }
        public string MailCCRecipients { get; set; }
    }
}
