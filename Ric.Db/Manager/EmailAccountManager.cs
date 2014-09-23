using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Info;
using System.Data;

namespace Ric.Db.Manager
{
    public class EmailAccountManager : ManagerBase
    {
        private const string ETI_EMAILACCOUNT_TABLE_NAME = "ETI_AutoEmailAccount";

        public static EmailAccountInfo SelectEmailAccountByAccountName(string accountName)
        {
            try
            {
                if (accountName == null || accountName.Trim() == "")
                {
                    return null;
                }

                string where = string.Format("where AccountName = '{0}' and Status = 'Active'", accountName);
                DataTable dt = ManagerBase.Select(ETI_EMAILACCOUNT_TABLE_NAME, new string[] { "*" }, where);

                if (dt == null || dt.Rows.Count == 0)
                {
                    return null;
                }

                //Column AccountName is unique so dt.Row.Count==1
                DataRow dr = dt.Rows[0];
                EmailAccountInfo emailAccount = new EmailAccountInfo();
                emailAccount.AccountName = Convert.ToString(dr["AccountName"]);
                emailAccount.Domain = Convert.ToString(dr["Domain"]);
                emailAccount.MailAddress = Convert.ToString(dr["MailAddress"]);
                emailAccount.Password = Convert.ToString(dr["Password"]);
                emailAccount.Status = (AccountStatus)Enum.Parse(typeof(AccountStatus), Convert.ToString(dr["Status"]));

                return emailAccount;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static int UpdateEmailAccount(EmailAccountInfo emailAccount)
        {
            try
            {
                if (emailAccount == null || emailAccount.AccountName.Trim() == "")
                {
                    return 0;
                }

                string where = string.Format("where AccountName = '{0}'", emailAccount.AccountName);
                DataTable dt = Select(ETI_EMAILACCOUNT_TABLE_NAME, new string[] { "*" }, where);

                if (dt == null)
                {
                    return 0;
                }

                if (dt.Rows.Count > 0)
                {
                    //Column AccountName is unique so dt.Row.Count==1
                    DataRow dr = dt.Rows[0];

                    dr["AccountName"] = emailAccount.AccountName;
                    dr["Domain"] = emailAccount.Domain;
                    dr["MailAddress"] = emailAccount.MailAddress;
                    dr["Password"] = emailAccount.Password;
                    dr["Status"] = emailAccount.Status.ToString();
                    dr["Description"] = emailAccount.Description;
                }
                else
                {
                    DataRow dr = dt.NewRow();

                    dr["AccountName"] = emailAccount.AccountName;
                    dr["Domain"] = emailAccount.Domain;
                    dr["MailAddress"] = emailAccount.MailAddress;
                    dr["Password"] = emailAccount.Password;
                    dr["Status"] = emailAccount.Status.ToString();
                    dr["Description"] = emailAccount.Description;

                    dt.Rows.Add(dr);
                }

                return UpdateDbTable(dt, ETI_EMAILACCOUNT_TABLE_NAME);
            }
            catch (Exception)
            {
                return 0;
            }
        }
    }
}
