using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;

namespace Ric.Tasks
{
    #region Configuration

    [ConfigStoredInDB]
    public class DownloadNDAFileConfig
    {
        [StoreInDB]
        [DisplayName("Account name")]
        [Category("EmailAccount")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Recipients")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Recipients (Cc)")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Signature")]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }

        [StoreInDB]
        [Category("FilePath")]
        [DisplayName("File path")]
        [Description("GeneratedFilePath")]
        public string FilePath { get; set; }

        [StoreInDB]
        [Category("DateOfFileList")]
        [Description(" Default download today's NDA file like: 2014-03-03")]
        public string DateTime { get; set; }

        [StoreInDB]
        [Category("SimpleFiles")]
        [Description("SimpleNDAFile list \nNDA files name on ftp\nlike: EM010223.M")]
        public List<string> ListSimpleFileName { get; set; }
    }
    #endregion

    public class DownloadNDAFile : GeneratorBase
    {
        #region Description

        private static DownloadNDAFileConfig configObj;
        private string accountName = string.Empty;//UC169XXX
        private string password = string.Empty;//********
        private string domain = string.Empty;//TEN
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private string strFilePath = string.Empty;//store file path
        private string strDateTime = string.Empty;//download file this day default:download today's file
        private List<string> listNDAFileSimple = new List<string>();//list file to get the file type
        private List<string> listFileNameError = new List<string>();//list fileName of error when download file from ftp
        private ExchangeService service;
        private string strFtp = string.Empty;

        protected override void Initialize()
        {
            configObj = Config as DownloadNDAFileConfig;
            strFilePath = configObj.FilePath;
            listNDAFileSimple = configObj.ListSimpleFileName;
            strDateTime = configObj.DateTime;
            strFtp = @"ftp://ASIA2:ASIA2@ds1.rds.reuters.com//";
            accountName = configObj.AccountName.Trim();
            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());
            accountName = emailAccount.AccountName;
            password = emailAccount.Password;
            domain = emailAccount.Domain;
            listMailTo = configObj.MailTo;
            service = MSAD.Common.OfficeUtility.EWSUtility.CreateService(new NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
        }
        #endregion

        protected override void Start()
        {
            List<string> listNDAFile = FormatDownloadConfig(strDateTime, listNDAFileSimple);
            DownloadListFile(strFilePath, listNDAFile);
            SendEmail(listNDAFile, listFileNameError);
            CleanOldFile(strFilePath);
        }

        /// <summary>
        /// delete file more than ten days
        /// </summary>
        /// <param name="strFilePath">Config</param>
        private void CleanOldFile(string strFilePath)
        {
            if (!Directory.Exists(strFilePath))
            {
                Logger.Log(string.Format("the directory:[{0}] doesn't exist", strFilePath));
                return;
            }
            List<string> listDateCode = new List<string>
            {
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(+1).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-1).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-2).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-3).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-4).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-5).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-6).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-7).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-8).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-9).ToString("MMdd"),
                DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-10).ToString("MMdd")
            };
            string[] arrFilesPath = Directory.GetFiles(strFilePath);
            for (int i = 0; i < arrFilesPath.Length; i++)
            {
                string strFileCode = arrFilesPath[i].Substring(arrFilesPath[i].Length - 6, 4);
                if (!listDateCode.Contains(strFileCode))
                {
                    if (File.Exists(arrFilesPath[i]))
                    {
                        File.Delete(arrFilesPath[i]);
                    }
                }
            }
        }

        /// <summary>
        /// send email step:one
        /// </summary>
        /// <param name="listNDAFile">download file path</param>
        /// <param name="listFileNameError">download file error path</param>
        private void SendEmail(List<string> listNDAFile, List<string> listFileNameError)
        {
            string subject = string.Empty;
            string content = string.Empty;
            if (listNDAFile == null || listNDAFile.Count == 0)
            {
                subject = "From Task ETI-297 QC DSE FTP file sourcing " + DateTime.Now.ToString("dd-MMM-yyyy") + "[error information]";
                content = "<center>*******Error happened when load configuration,Can't get file simple type.*********</center><br />";
                SendMail(service, subject, content, new List<string>());
                return;
            }
            if (listFileNameError == null || listFileNameError.Count == 0)
            {
                return;
            }
            subject = "From Task ETI-297 QC DSE FTP file sourcing " + DateTime.Now.ToString("dd-MMM-yyyy") + "[error information]";
            content = listFileNameError.Aggregate("<center>*******Error happened when download following files.*********</center><br />", (current, str) => current + (str + "<br />"));
            SendMail(service, subject, content, new List<string>());
        }

        /// <summary>
        /// download files step:one
        /// </summary>
        /// <param name="strFilePath">file path</param>
        /// <param name="listNDAFile">download file list</param>
        private void DownloadListFile(string strFilePath, List<string> listNDAFile)
        {
            if (listNDAFile == null || listNDAFile.Count == 0)
            {
                Logger.Log(string.Format("Config error, no simple list file, ListNDAFile==null "));
                return;
            }
            if (!Directory.Exists(strFilePath))
            {
                Directory.CreateDirectory(strFilePath);
            }
            foreach (string strUrl in listNDAFile)
            {
                DownloadFile(strUrl, strFilePath);
            }
        }

        /// <summary>
        /// download files step:two
        /// </summary>
        /// <param name="url">file url</param>
        /// <param name="strFilePath">path local</param>
        public void DownloadFile(string url, string strFilePath)
        {
            string strFileName = string.Empty;
            try
            {
                int start = url.LastIndexOf("/");
                strFileName = url.Substring(start + 1, url.Length - start - 1);
                strFileName = Path.Combine(strFilePath, strFileName);
                FtpWebRequest Myrq = (FtpWebRequest)WebRequest.Create(url);
                WebProxy proxy = new WebProxy("10.40.14.56", 80);
                Myrq.Proxy = proxy;
                FtpWebResponse myrp = (FtpWebResponse)Myrq.GetResponse();
                using (Stream st = myrp.GetResponseStream())
                {
                    using (Stream so = new FileStream(strFileName, FileMode.Create))
                    {
                        long totalDownloadedByte = 0;
                        byte[] by = new byte[128];
                        int osize = st.Read(by, 0, @by.Length);
                        while (osize > 0)
                        {
                            totalDownloadedByte = osize + totalDownloadedByte;
                            so.Write(by, 0, osize);
                            osize = st.Read(by, 0, @by.Length);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                listFileNameError.Add(url);
                Logger.Log(string.Format("error when try to download file :{0} Exception:{1}", strFileName, ex));
            }
        }

        /// <summary>
        /// get download file path step:two
        /// </summary>
        /// <param name="strDateTime">file name</param>
        /// <param name="listNDAFileSimple">file type</param>
        /// <returns></returns>
        private List<string> FormatDownloadConfig(string strDateTime, List<string> listNDAFileSimple)
        {
            List<string> listNDAFile = new List<string>();
            DateTime date;
            string strFirstCode = string.Empty;
            string strLastCode = string.Empty;
            string fileName = string.Empty;

            if (listNDAFileSimple == null || listNDAFileSimple.Count == 0)
            {
                Logger.Log(string.Format("Config error, no simple list file : "));
                return null;
            }
            if (listNDAFileSimple.Count == 1 && listNDAFileSimple[0].Trim() == "")
            {
                Logger.Log(string.Format("Config error, no simple list file : "));
                return null;
            }
            if (!string.IsNullOrEmpty(strDateTime.Trim()))
            {
                try
                {
                    date = DateTime.Parse(configObj.DateTime);
                    strDateTime = date.ToUniversalTime().AddHours(+8).ToString("MMdd");
                }
                catch (Exception ex)
                {
                    Logger.Log(string.Format("Config error,can't convert{0} to DateTime Type :{1}", strDateTime, ex.ToString()));
                    return null;
                }
            }
            else
            {
                strDateTime = DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd");
            }
            foreach (string str in listNDAFileSimple.Where(str => !string.IsNullOrEmpty(str.Trim())))
            {
                strFirstCode = str.Substring(0, 4);
                strLastCode = str.Substring(8, 2);
                fileName = Path.Combine(strFtp, string.Format("{0}{1}{2}", strFirstCode, strDateTime, strLastCode));

                if (!listNDAFile.Contains(fileName))
                {
                    listNDAFile.Add(fileName);
                }
            }
            return listNDAFile;
        }

        /// <summary>
        /// send email step:two
        /// </summary>
        /// <param name="service">logging</param>
        /// <param name="subject">subject</param>
        /// <param name="content">body</param>
        /// <param name="attacheFileList">attachement</param>
        private void SendMail(ExchangeService service, string subject, string content, List<string> attacheFileList)
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
    }
}
