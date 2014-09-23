using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Info
{
    public class EmailAccountInfo
    {
        public int AccountId { get; set; }
        public string AccountName { get; set; }
        public string Domain { get; set; }
        public string MailAddress { get; set; }
        public string Password { get; set; }
        public AccountStatus Status { get; set; }
        public string Description { get; set; }
    }

    public enum AccountStatus
    {
        Active,
        Disabled
    }
}
