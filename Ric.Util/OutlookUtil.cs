using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using System.ComponentModel;

namespace Ric.Util
{
    [TypeConverter(typeof(ExpandableObjectConverter))] 
    public class MailSearchQuery
    {
        public string AccountName { get; set; }
        public string Sender { get; set; }
        public List<string> ReceiverList { get; set; }
        public string MailFolderPath { get; set; }
        public List<string> SubjectKeywordList { get; set; }
        public List<string> SubjectExcludedwordList { get; set; }
        public List<string> bodyKeywordList { get; set; }
        public List<string> AttachFileList { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public MailSearchQuery(string accountName, string sender, string mailFolderPath, DateTime startDate, DateTime endDate)
        {
            this.AccountName = accountName;
            this.Sender = sender;
            this.MailFolderPath = mailFolderPath;
            this.StartDate = startDate;
            this.EndDate = endDate;
            this.ReceiverList = new List<string>();
            this.SubjectKeywordList = new List<string>();
            this.SubjectExcludedwordList = new List<string>();
            this.bodyKeywordList = new List<string>();
            this.AttachFileList = new List<string>();
        }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class MailToSend
    {
        public List<string> ToReceiverList { get; set; }
        public List<string> CCReceiverList { get; set; }
        public List<string> BCCReceiverList { get; set; }
        public string MailSubject { get; set; }
        public string MailBody { get; set; }
        public string MailHtmlBody { get; set; }
        public List<string> AttachFileList { get; set; }

        public MailToSend()
        {
            this.ToReceiverList = new List<string>();
            this.CCReceiverList = new List<string>();
            this.BCCReceiverList = new List<string>();
            this.AttachFileList = new List<string>();
        }
    }

    public class OutlookApp:IDisposable
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private Application outlookAppInstance = null;

        public Application OutlookAppInstance
        {
            get { return outlookAppInstance; }
            set { outlookAppInstance = value; }
        }

         public OutlookApp()
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            outlookAppInstance = new Microsoft.Office.Interop.Outlook.Application();
        }

         #region IDisposable Members

         public void Dispose()
         {
             //if (outlookAppInstance != null)
             //{
             //    outlookAppInstance.Quit();
             //}
             //Thread.Sleep(1000);
         }

         #endregion
    }


    public class OutlookUtil
    {
        [DllImport("user32.dll")]
        private static extern void GetWindowThreadProcessId(IntPtr hWnd, out int k);

        public static void CreateAndSendMail(OutlookApp outlookApp, MailToSend mail, out string err)
        {           
            CreateAndSendMail(outlookApp, mail.AttachFileList, mail.MailSubject, mail.ToReceiverList, mail.CCReceiverList, mail.MailBody, mail.MailHtmlBody, out err);
        }

        public static void CreateAndSendMail(OutlookApp outlookApp, List<string> attachedFileList, string subject, List<string> toTypeRecipients, List<string> ccTypeRecipient, string mailBody, out string err)
        {
            CreateAndSendMail(outlookApp, attachedFileList, subject, null, toTypeRecipients, ccTypeRecipient, null, mailBody, null, out err);
        }

        public static void CreateAndSendMail(OutlookApp outlookApp, List<string> attachedFileList, string subject, List<string> originatorRecipientList, List<string> toTypeRecipients,List<string>ccTypeRecipient,List<string>bccTypeRecipient,string mailBody, string mailHtmlBody, out string err)
        {
            err = string.Empty;
            MailItem mail = outlookApp.OutlookAppInstance.CreateItem(OlItemType.olMailItem) as MailItem;
            mail.Subject = subject;
            mail.Body = mailBody;
            if (mailHtmlBody != null)
            {                
                mail.HTMLBody = mailHtmlBody;
            }
            AddressEntry currentUser = outlookApp.OutlookAppInstance.Session.CurrentUser.AddressEntry;
            if (toTypeRecipients != null && toTypeRecipients.Count > 0)
            {
                foreach (string recipient in toTypeRecipients)
                {
                    if (!string.IsNullOrEmpty(recipient))
                    {
                        Recipient toRecipient = mail.Recipients.Add(recipient);
                        toRecipient.Type = (int)OlMailRecipientType.olTo;
                    }
                }
            }

            if (originatorRecipientList != null && originatorRecipientList.Count > 0)
            {
                foreach (string recipient in originatorRecipientList)
                {
                    Recipient originalRecipient = mail.Recipients.Add(recipient);
                    originalRecipient.Type = (int)OlMailRecipientType.olOriginator;
                }
            }

            if (ccTypeRecipient != null && ccTypeRecipient.Count > 0)
            {
                foreach (string recipient in ccTypeRecipient)
                {
                    if (!string.IsNullOrEmpty(recipient))
                    {
                        Recipient ccRecipient = mail.Recipients.Add(recipient);
                        ccRecipient.Type = (int)OlMailRecipientType.olCC;
                    }
                }
            }

            if (bccTypeRecipient != null && bccTypeRecipient.Count > 0)
            {
                foreach (string recipient in bccTypeRecipient)
                {
                    Recipient bccRecipient = mail.Recipients.Add(recipient);
                    bccRecipient.Type = (int)OlMailRecipientType.olBCC;
                }
            }

            bool isResolveSuccess = mail.Recipients.ResolveAll();
            
            if (!isResolveSuccess)
            {
                err += "Warning: There's error when resolve the recipients' names";
            }            

            if(attachedFileList!=null)
            {
                try
                {
                    foreach (string attachedFilePath in attachedFileList)
                    {
                        mail.Attachments.Add(attachedFilePath, OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }
                }
                catch (System.Exception ex)
                {
                    err += "Error: There's error when adding attach file to the mail " + ex.Message;
                    Console.WriteLine("There's error when adding attach file to the mail " + ex.Message);
                }
            }
            mail.Send();
        }

        public static void CreateAndSendMail(OutlookApp outlookApp, List<string> attachedFileList, string subject, List<string> toTypeRecipients, List<string> ccTypeRecipient, string mailBody, string mailHtmlBody, out string err)
        {
             CreateAndSendMail(outlookApp, attachedFileList, subject, null, toTypeRecipients, ccTypeRecipient, null, mailBody, mailHtmlBody, out err);
        }

        //Send mail without attachment
        public static void CreateAndSendMail(OutlookApp outlookApp, string subject, List<string> toTypeRecipients, List<string> ccTypeRecipient, string mailBody, out string err)
        {
            CreateAndSendMail(outlookApp, null, subject, toTypeRecipients, ccTypeRecipient, mailBody, null, out err);
        }

        public static List<MailItem> SearchMail(OutlookApp outlookApp, MailSearchQuery query)
        {
            List<MailItem> mailList = new List<MailItem>();
            MAPIFolder folder = SearchTargetMailAccount(outlookApp, query.AccountName);
            folder = GetFolderPath(folder, query.MailFolderPath);
            //mailList = SearchEmailInSpecifiedFolder(outlookApp, query.AccountName, query.MailFolderPath);
            //mailList = GetMailsSendInCertainTime(mailList,query.StartDate,query.EndDate);
            mailList = GetMailsSendInCertainTime(folder, query.StartDate, query.EndDate);
            mailList = GetMailsFromCertainAddress(mailList, query.Sender);
            mailList = GetMailsWithCertainSubject(mailList, query.SubjectKeywordList);
            mailList = GetMailsWithoutCertainSubject(mailList, query.SubjectExcludedwordList);
            return mailList;
        }


        //Search mails in specified sub folders under a specified account
        public static List<MailItem> SearchEmailInSpecifiedFolder(OutlookApp outlookApp, string accountName, string subFolderPath)
        {
            MAPIFolder mapiFolder = SearchTargetMailAccount(outlookApp, accountName);
            List<MailItem> mailList = SearchEmailInSpecifiedFolder(mapiFolder, subFolderPath);
            return mailList;
        }
        //Search mails in a specified account
        public static MAPIFolder SearchTargetMailAccount(OutlookApp outlookApp, string accountName)
        {
            _NameSpace ns = outlookApp.OutlookAppInstance.GetNamespace("MAPI");
            ns.Logon(null, null, false, false);
            MAPIFolder folder = null;
            try
            {
                Folders folders = outlookApp.OutlookAppInstance.Application.Session.Folders;

                foreach (MAPIFolder item in folders)
                {
                    if (item.Name.Contains(accountName))
                    {
                        folder = item;
                        break;
                    }
                }
            }

            catch (System.Exception ex)
            {
                throw new System.Exception(string.Format("Error found when searching the target mail account {0}. The exception message is: {1}", accountName,ex.Message));
            }
            return folder;
        }

        //Search mails in specified sub folders
        public static List<MailItem> SearchEmailInSpecifiedFolder(MAPIFolder mapiFolder, string folderPath)
        {
            List<MailItem> itemList = new List<MailItem>();
   
            try
            {
                mapiFolder = GetFolderPath(mapiFolder, folderPath);
                Items items = mapiFolder.Items;

                for (int i = 1; i <= items.Count; i++)
                {
                    try
                    {
                        MailItem item = (MailItem)items[i];
                        itemList.Add(item);
                    }
                    catch (System.Exception ex)
                    {
                        string err = ex.ToString();
                    }
                }
            }

            catch (System.Exception ex)
            {
                string errInfo = ex.ToString();
                throw new System.Exception(string.Format("Error found when searching mails in {0} under the account {1}",folderPath,mapiFolder.Name));
            }
            return itemList;
        }

        public static MAPIFolder GetFolderPath(MAPIFolder mapiFolder, string folderPath)
        {
            string[] folderPathArr = folderPath.Split('\\');
            for (int i = 0; i < folderPathArr.Length; i++)
            {
                mapiFolder = mapiFolder.Folders[folderPathArr[i]];
            }

            return mapiFolder;
        }

        //Search and get emails by email address
        public static List<MailItem> GetMailsFromCertainAddress(List<MailItem> itemList, string senderEmailAddress)
        {
            List<MailItem> mailList = new List<MailItem>();
            foreach (MailItem item in itemList)
            {
                if (item.SenderEmailAddress.ToLower().Contains(senderEmailAddress.ToLower()))
                {
                    mailList.Add(item);
                }
            }
            return mailList;
        }

        // Search emails which subject contains the subject keyword
        public static List<MailItem> GetMailsWithCertainSubject(List<MailItem> itemList, string subjectKeyword)
        {
            if (string.IsNullOrEmpty(subjectKeyword))
            {
                return itemList;
            }
            else
            {
                List<MailItem> mailList = new List<MailItem>();
                foreach (MailItem item in itemList)
                {
                    if (item.Subject.Contains(subjectKeyword))
                    {
                        mailList.Add(item);
                    }
                }
                return mailList;
            }
        }

        // Search emails which subject contains the subject keyword
        public static List<MailItem> GetMailsWithCertainSubject(List<MailItem> itemList, List<string> subjectKeywordList)
        {
            List<MailItem> mailList = new List<MailItem>();
            if (subjectKeywordList == null || subjectKeywordList.Count == 0)
            {
                mailList = itemList;
            }

            else
            {
                foreach (MailItem item in itemList)
                {
                    string subject = item.Subject;
                    bool isRequired = true;

                    foreach (string keyword in subjectKeywordList)
                    {
                        if (isRequired == true)
                        {
                            isRequired = subject.Contains(keyword);
                        }
                        else
                            break;
                    }

                    if (isRequired == true)
                    {
                        mailList.Add(item);
                    }
                }
            }
            return mailList;
        }

        // Search emails which subject contains the subject keyword
        public static List<MailItem> GetMailsWithoutCertainSubject(List<MailItem> itemList, List<string> excludedKeywordList)
        {
            List<MailItem> mailList = new List<MailItem>();
            if (excludedKeywordList == null || excludedKeywordList.Count == 0)
            {
                mailList = itemList;
            }

            else
            {
                foreach (MailItem item in itemList)
                {
                    string subject = item.Subject;
                    bool isRequired = true;
                    foreach (string keyword in excludedKeywordList)
                    {
                        if (isRequired == true)
                        {
                            isRequired = !subject.Contains(keyword);
                        }
                        else
                            break;
                    }

                    if (isRequired == true)
                    {
                        mailList.Add(item);
                    }
                }
            }
            return mailList;
        }

        //Search the mail which was sent between start time and end time
        public static List<MailItem> GetMailsSendInCertainTime(List<MailItem> itemList, DateTime startTime, DateTime endTime)
        {
            List<MailItem> mailList = new List<MailItem>();
            foreach (MailItem item in itemList)
            {
                if (item.ReceivedTime.Date >= startTime.Date && item.ReceivedTime.Date <= endTime.Date)
                {
                    mailList.Add(item);
                }
            }
            return mailList;
        }

        public static List<MailItem> GetMailsSendInCertainTime(MAPIFolder folder, DateTime startTime, DateTime endTime)
        {
            List<MailItem> mailList = new List<MailItem>();
            try
            {
                Items items = folder.Items;

                for (int i = 1; i <= items.Count; i++)
                {
                    try
                    {
                        MailItem item = (MailItem)items[i];
                        if (item.ReceivedTime.Date >= startTime.Date && item.ReceivedTime.Date <= endTime.Date)
                        {
                            mailList.Add(item);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        string err = ex.ToString();
                    }
                }
            }

            catch (System.Exception ex)
            {
                string err = ex.ToString();
                throw new System.Exception(string.Format("Error found when searching mails between {0} and {1}", startTime.ToString("yyyyMMdd"), endTime.ToString("yyyyMMdd")));
            }
            return mailList;
            

        }


        //Download the required attachment file
        public static List<string> DownloadAttachments(MailItem item, List<string> attachmentKeywordList, List<string> excludedWordList, string savedDir)
        {
            List<string> attachmentList = new List<string>();
            if (item == null)
            {
                return null;
            }
            foreach (Attachment attachment in item.Attachments)
            {
                bool isRequired = true;
                string fileName = attachment.FileName;
                if (attachmentKeywordList != null && attachmentKeywordList.Count != 0)
                {
                    foreach (string keyword in attachmentKeywordList)
                    {
                        if (isRequired == true)
                        {
                            if (!fileName.Contains(keyword))
                            {
                                isRequired = false;
                                break;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                //if (isRequired == false)
                //{
                //    break;
                //}

                else
                {
                    if (excludedWordList != null && excludedWordList.Count != 0)
                    {
                        foreach (string excludedWord in excludedWordList)
                        {
                            if (isRequired == true)
                            {
                                if (fileName.Contains(excludedWord))
                                {
                                    isRequired = false;
                                    break;
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }

                if (isRequired == true)
                {
                    string filePath = Path.Combine(savedDir, attachment.FileName);
                    attachment.SaveAsFile(filePath);
                    attachmentList.Add(filePath);
                }
            }

            return attachmentList;
        }

        //Download the required attachment file
        public static List<string> DownloadAttachments(MailItem item, string keyword, string excludedWord, string savedDir)
        {
            List<string> attachmentList = new List<string>();
            List<string> keywordList = new List<string>();
            if (!string.IsNullOrEmpty(keyword))
            {
                keywordList.Add(keyword);
            }
            List<string> excludedWordList = new List<string>();
            if (!string.IsNullOrEmpty(excludedWord))
            {
                excludedWordList.Add(excludedWord);
            }
            return DownloadAttachments(item, keywordList, excludedWordList, savedDir);
        }

        public static void SaveMail(OutlookApp outlookApp, MailToSend mail, out string err, string path)
        {
            SaveMail(outlookApp, mail.AttachFileList, mail.MailSubject, mail.ToReceiverList, mail.CCReceiverList, mail.MailBody, mail.MailHtmlBody, out err, path);
        }

        public static void SaveMail(OutlookApp outlookApp, List<string> attachedFileList, string subject, List<string> originatorRecipientList, List<string> toTypeRecipients, List<string> ccTypeRecipient, List<string> bccTypeRecipient, string mailBody, string mailHtmlBody, out string err, string path)
        {
            err = string.Empty;
            MailItem mail = outlookApp.OutlookAppInstance.CreateItem(OlItemType.olMailItem) as MailItem;
            mail.Subject = subject;
            mail.Body = mailBody;
            if (mailHtmlBody != null)
            {
                mail.HTMLBody = mailHtmlBody;
            }
            AddressEntry currentUser = outlookApp.OutlookAppInstance.Session.CurrentUser.AddressEntry;
            if (toTypeRecipients != null && toTypeRecipients.Count > 0)
            {
                foreach (string recipient in toTypeRecipients)
                {
                    if (!string.IsNullOrEmpty(recipient))
                    {
                        Recipient toRecipient = mail.Recipients.Add(recipient);
                        toRecipient.Type = (int)OlMailRecipientType.olTo;
                    }
                }
            }

            if (originatorRecipientList != null && originatorRecipientList.Count > 0)
            {
                foreach (string recipient in originatorRecipientList)
                {
                    Recipient originalRecipient = mail.Recipients.Add(recipient);
                    originalRecipient.Type = (int)OlMailRecipientType.olOriginator;
                }
            }

            if (ccTypeRecipient != null && ccTypeRecipient.Count > 0)
            {
                foreach (string recipient in ccTypeRecipient)
                {
                    if (!string.IsNullOrEmpty(recipient))
                    {
                        Recipient ccRecipient = mail.Recipients.Add(recipient);
                        ccRecipient.Type = (int)OlMailRecipientType.olCC;
                    }
                }
            }

            if (bccTypeRecipient != null && bccTypeRecipient.Count > 0)
            {
                foreach (string recipient in bccTypeRecipient)
                {
                    Recipient bccRecipient = mail.Recipients.Add(recipient);
                    bccRecipient.Type = (int)OlMailRecipientType.olBCC;
                }
            }

            bool isResolveSuccess = mail.Recipients.ResolveAll();
            
            if (!isResolveSuccess)
            {
                err += "Warning: There's error when resolve the recipients' names";
            }
            

            if (attachedFileList != null)
            {
                try
                {
                    foreach (string attachedFilePath in attachedFileList)
                    {
                        mail.Attachments.Add(attachedFilePath, OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }
                }
                catch (System.Exception ex)
                {
                    err += "Error: There's error when adding attach file to the mail " + ex.Message;
                    Console.WriteLine("There's error when adding attach file to the mail " + ex.Message);
                }
            }
            mail.SaveAs(path, Type.Missing);
        }

        public static void SaveMail(OutlookApp outlookApp, List<string> attachedFileList, string subject, List<string> toTypeRecipients, List<string> ccTypeRecipient, string mailBody, string mailHtmlBody, out string err, string path)
        {
            SaveMail(outlookApp, attachedFileList, subject, null, toTypeRecipients, ccTypeRecipient, null, mailBody, mailHtmlBody, out err, path);
        }

    }
}
