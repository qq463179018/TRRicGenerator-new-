using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using MSAD.Common.OfficeUtility;

namespace Ric.Util
{
    public class ExchangeUtil
    {
        private ExchangeService service;
        private EWSMailSearchQuery query;

        private string userName;
        private string password;
        private string domain;
        private string urlWSDL;

        private ExchangeUtil(ExchangeService service)
        {
            this.service = service;
        }

        public ExchangeUtil(string userName, string password, string domain, string urlWSDL = @"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx")
        {
            this.userName = userName;
            this.password = password;
            this.domain = domain;
            this.urlWSDL = urlWSDL;
            service = EWSUtility.CreateService(new System.Net.NetworkCredential(userName, password, domain), new Uri(urlWSDL));
        }

        public List<EmailMessage> GetSearchQueryEmailMessage(string mailbox, string subjectKeyword, DateTime startDate, DateTime endDate, string sendAddress = "", string mailFolderPath = @"Inbox", string bodyKeyword = "")
        {
            List<EmailMessage> emails = null;
            this.query = new EWSMailSearchQuery(sendAddress, mailbox, mailFolderPath, subjectKeyword, bodyKeyword, startDate, endDate);

            RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
            {
                emails = EWSMailSearchQuery.SearchMail(service, query);
            });

            return emails;
        }

        public List<string> DownloadAttachmentFile(string savePath, EmailMessage email, string keyWord = "", string excludedWord = "")
        {
            return EWSMailSearchQuery.DownloadAttachments(service, email, keyWord, excludedWord, savePath);
        }

        public List<string> DownloadAttachmentFile(string savePath, List<EmailMessage> emails, string keyWord = "", string excludedWord = "")
        {
            List<string> attachmentFilesPath = new List<string>();
            List<string> temp = null;

            foreach (var email in emails)
            {
                temp = EWSMailSearchQuery.DownloadAttachments(service, email, keyWord, excludedWord, savePath);
                if (temp != null && temp.Count != 0)
                    attachmentFilesPath.AddRange(temp);
            }

            return attachmentFilesPath;
        }

        public List<string> DownloadAttachmentFile(string savePath, string mailbox, string subjectKeyword, DateTime startDate, DateTime endDate, string sendAddress = "", string mailFolderPath = @"Inbox", string bodyKeyword = "")
        {
            List<EmailMessage> emails = GetSearchQueryEmailMessage(mailbox, subjectKeyword, startDate, endDate, sendAddress, mailFolderPath, bodyKeyword);
            if (emails == null)
                throw new Exception("searche query email error,please check query paramters.");

            return DownloadAttachmentFile(savePath, emails);
        }
    }
}
