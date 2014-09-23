using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Drawing.Design;
using Microsoft.Exchange.WebServices.Data;
using Ric.Db.Info;
using Ric.Db.Manager;
using System.Configuration;
using Ric.Db.Config;
using System.IO;
using MSAD.Common.OfficeUtility;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Ric.Core;
using System.Windows.Forms;
using Ric.Util;

namespace Ric.Tasks.China
{
    #region [Configuration]
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    class ChinaIpoQCMorningConfig
    {
        [StoreInDB]
        [Category("EmailAccount")]
        [DefaultValue("UC159450")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [Description("Mail folder path,like: Inbox\\XXXXX")]
        public string MailFolder { get; set; }

        [StoreInDB]
        [Category("OutputPath")]
        [Description("output result file path")]
        public string OutputPath { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }
    }
    #endregion

    class ChinaIpoQCMorning : GeneratorBase
    {
        #region [Description]
        private static ChinaIpoQCMorningConfig configObj = null;
        EmailAccountInfo emailAccount = null;
        private ExchangeService service;
        private string downloadFilePath = string.Empty;
        private string resultFilePath = string.Empty;
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private List<string> listRic = null;
        private List<string> attacheFileList = null;
        List<string> listDownloadFile = null;
        private string mailFolder = string.Empty;
        private DateTime startDate = DateTime.Now.ToUniversalTime().AddHours(+8).AddHours(-DateTime.Now.ToUniversalTime().AddHours(+8).Hour);
        private DateTime endDate = DateTime.Now.ToUniversalTime().AddHours(+8);
        private int emailCountInMailBox = 0;

        protected override void Initialize()
        {
            configObj = Config as ChinaIpoQCMorningConfig;
            emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());

            if (emailAccount == null)
            {
                MessageBox.Show("email account is not exist in DB. ");
                return;
            }

            service = MSAD.Common.OfficeUtility.EWSUtility.CreateService(new System.Net.NetworkCredential(emailAccount.AccountName, emailAccount.Password, emailAccount.Domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            downloadFilePath = Path.Combine(configObj.OutputPath, "downloadFile");
            resultFilePath = Path.Combine(configObj.OutputPath, "output");
            mailFolder = configObj.MailFolder.Trim();
        }
        #endregion

        protected override void Start()
        {
            #region [Get Download List]
            try
            {
                listDownloadFile = GetDownloadFilePathFromEmail();
            }
            catch (Exception ex)
            {
                string msg = string.Format("download file from email error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            #region [Get Ric From File]
            try
            {
                listRic = GetRicFromFile(listDownloadFile);
            }
            catch (Exception ex)
            {
                string msg = string.Format("get ric from file error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            #region [Generate Result File]
            try
            {
                attacheFileList = GenerateFile(listRic);
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate txt file to local error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            #region [Send Email]
            try
            {
                SendEmail();
            }
            catch (Exception ex)
            {
                string msg = string.Format("send email error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion
        }

        #region [Get Ric From File]
        private List<string> GetRicFromFile(List<string> list)
        {
            List<string> listRic = new List<string>();
            string filePath = string.Empty;

            if (list == null || list.Count == 0)
            {
                string msg = string.Format("no download file in this email .");
                Logger.Log(msg, Logger.LogType.Warning);
                return null;
            }

            foreach (var item in list)
            {
                filePath = item;
                break;
            }

            using (Ric.Util.ExcelApp app = new Ric.Util.ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                var worksheet = (Worksheet)workbook.Worksheets[1];
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Down))
                {
                    while (reader.Row <= lastUsedRow)
                    {
                        string key = reader.ReadLineCellText();

                        if (key.Equals("CHANGE"))
                            reader.PlaceNext(reader.Row + 11, reader.Col);

                        if (key.Equals("DROP"))
                            reader.PlaceNext(reader.Row + 8, reader.Col);

                        if (key.Equals("RIC"))
                        {
                            int lastUsedCol = worksheet.UsedRange.Columns.Count;
                            string value = string.Empty;
                            string ricSS = string.Empty;

                            using (ExcelLineWriter readerCol = new ExcelLineWriter(worksheet, reader.Row - 1, reader.Col + 1, ExcelLineWriter.Direction.Right))
                            {
                                while (readerCol.Col <= lastUsedCol)
                                {
                                    value = readerCol.ReadLineCellText();
                                    if (!string.IsNullOrEmpty(value) && !listRic.Contains(value))
                                        listRic.Add(value);

                                    if (value.EndsWith(".SS") && value.StartsWith("6"))
                                    {
                                        ricSS = string.Format("{0}.SH", value.Substring(0, value.Length - 3));
                                        if (!string.IsNullOrEmpty(ricSS) && !listRic.Contains(ricSS))
                                            listRic.Add(ricSS);
                                    }
                                }
                            }
                        }
                    }
                    reader.PlaceNext(reader.Row, 1);
                }
                workbook.Close(false, workbook.FullName, false);
            }

            return listRic;
        }
        #endregion

        #region [Send Email]
        private void SendEmail()
        {
            string subject = string.Empty;
            string content = string.Empty;

            if (listDownloadFile == null)
            {
                subject = "China IPO QC Morning - Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy");
                content = "<center>********Generate Unknown Exception About Email Account.********</center><br /><br />";
                SendEmail(service, subject, content, new List<string>());
                return;
            }

            if (attacheFileList == null || attacheFileList.Count == 0)
            {
                subject = "China IPO QC Morning - Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy");
                content = "<center>********No IPO Ric Today .********</center><br /><br />";
                SendEmail(service, subject, content, new List<string>());
                return;
            }

            subject = "China IPO QC Morning - Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy");
            content = "<center>********Ipo Ric Today.********</center><br /><br />";
            SendEmail(service, subject, content, attacheFileList);//ok
        }

        private void SendEmail(ExchangeService service, string subject, string content, List<string> attacheFileList)
        {
            StringBuilder bodyBuilder = new StringBuilder();
            bodyBuilder.Append(content);
            bodyBuilder.Append("<p>");

            foreach (string signatureLine in configObj.MailSignature)
            {
                bodyBuilder.AppendFormat("{0}<br />", signatureLine);
            }

            bodyBuilder.Append("</p>");
            content = bodyBuilder.ToString();

            if (configObj.MailCC.Count > 1 || (configObj.MailCC.Count == 1 && configObj.MailCC[0] != ""))
            {
                listMailCC = configObj.MailCC;
            }

            MSAD.Common.OfficeUtility.EWSUtility.CreateAndSendMail(service, listMailTo, listMailCC, new List<string>(), subject, content, attacheFileList);
        }
        #endregion

        #region [Generate Result File]
        private List<string> GenerateFile(List<string> list)
        {
            List<string> listAttachement = new List<string>();

            if (list == null || list.Count == 0)
            {
                string msg = string.Format("no ric in the pdf files!");
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }

            string path = Path.Combine(resultFilePath, string.Format("RIC_{0}.txt", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyy-MM-dd")));
            StringBuilder sb = new StringBuilder();
            sb.Append("RIC\t\r\n");

            foreach (var item in list)
            {
                sb.AppendFormat("{0}\t\r\n", item);
            }

            try
            {
                if (!Directory.Exists(resultFilePath))
                    Directory.CreateDirectory(resultFilePath);

                File.WriteAllText(path, sb.ToString());
                TaskResultList.Add(new TaskResultEntry("China Ipo Qc", "ric list", path));
                listAttachement.Add(path);

                return listAttachement;
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating file. Ex: {0} .", ex.Message));

                return null;
            }
        }
        #endregion

        #region [Get Download File From Emial]
        private List<string> GetDownloadFilePathFromEmail()
        {
            List<string> list = new List<string>();

            try
            {
                EWSMailSearchQuery query = new EWSMailSearchQuery("", emailAccount.MailAddress, @mailFolder, "China FM for", "", startDate, endDate);
                List<EmailMessage> mailList = null;
                EmailMessage mail = null;

                for (int i = 0; i < 5; i++)
                {
                    try
                    {
                        mailList = EWSMailSearchQuery.SearchMail(service, query);
                        break;
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(2000);

                        if (i == 4)
                        {
                            throw new Exception("[EWSMailSearchQuery.SearchMail(service, query)] error.msg:" + ex.ToString());
                        }
                    }
                }

                if (mailList == null)
                {
                    string msg = string.Format("can't get email from mailbox. ");
                    Logger.Log(msg, Logger.LogType.Error);
                    return list;//list.count==0 //no ipo today
                }

                emailCountInMailBox = mailList.Count;//email count received

                if (!(emailCountInMailBox > 0))
                {
                    string msg = string.Format("no email in the mailbox. ");
                    Logger.Log(msg, Logger.LogType.Warning);
                    return list;//list.count==0 //no ipo today
                }

                if (!Directory.Exists(downloadFilePath))
                    Directory.CreateDirectory(downloadFilePath);

                mail = mailList[0];
                mail.Load();
                List<string> attachments = EWSMailSearchQuery.DownloadAttachments(service, mail, "", "", downloadFilePath);

                if (attachments == null && attachments.Count == 0)
                {
                    return list;// emailcount!=0 but no attachement so list.count==0 //no ipo today
                }

                int start = 0;
                string fileName = string.Empty;

                foreach (var str in attachments)
                {
                    start = str.LastIndexOf("\\");
                    fileName = str.Substring(start + 1, str.Length - start - 1);

                    if (!fileName.Contains(".xls"))
                        continue;

                    list.Add(str);
                }

                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("get email and download attachement error. :{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }
        #endregion
    }
}