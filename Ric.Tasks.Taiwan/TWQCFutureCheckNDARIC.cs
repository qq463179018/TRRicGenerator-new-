using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Exchange.WebServices.Data;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;

namespace Ric.Tasks.Taiwan
{
    #region Configuration
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class TWQCFutureCheckNDARICConfig
    {
        [StoreInDB]
        [Category("Email Account")]
        [DisplayName("Folder path")]
        [Description("Mail folder path,like: Inbox/XXXXX")]
        public string MailFolderPath { get; set; }

        [StoreInDB]
        [Category("Email Account")]
        [DisplayName("Name")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("File")]
        [DisplayName("File path")]
        [Description("Generated Files Path,like: G:\\China")]
        public string FilePath { get; set; }

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
    }
    #endregion

    #region Description
    class TWQCFutureCheckNDARIC : GeneratorBase
    {
        public TWQCFutureCheckNDARICConfig configObj = null;
        private string filePath = string.Empty;//G://XX
        private string accountName = string.Empty;//UC169XXX
        private string password = string.Empty;//********
        private string domain = string.Empty;//TEN
        private string mailAdress = string.Empty;//eti.XXXXXX@thomsonreuters.com
        private string mailFolder = string.Empty;//Inbox/XXXXX
        private ExchangeService service;
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private List<string> attacheFileList = new List<string>();
        private string strFilePathFromFtpOfFutRic = string.Empty;
        private string strFilePathFromFtpOfSpdRic = string.Empty;
        List<string> listFutRic = null;
        List<string> listSpdRic = null;
        Dictionary<string, string> dicFutRicFtp = new Dictionary<string, string>();
        Dictionary<string, string> dicSpdRicFtp = new Dictionary<string, string>();
        private string strFutPattern = string.Empty;
        private string strSpdPattern = string.Empty;
        private string strFutRicFilePath = string.Empty;
        private string strSpdRicFilePath = string.Empty;
        private bool isExistFirstFileOnFtp = true;
        private bool isExistSecondFileOnFtp = true;

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as TWQCFutureCheckNDARICConfig;
            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());
            accountName = emailAccount.AccountName;
            password = emailAccount.Password;
            domain = emailAccount.Domain;
            mailAdress = emailAccount.MailAddress;
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            filePath = configObj.FilePath;
            strFutPattern = @"^\b(?<RIC>\w+)\b.+?";
            strSpdPattern = @"^\b(?<RIC>\S+)\b\s+?";
            strFutRicFilePath = Path.Combine(filePath, DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-1).ToString("MMdd") + "FutRic.txt");
            strSpdRicFilePath = Path.Combine(filePath, DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-1).ToString("MMdd") + "SpdRic.txt");
            strFilePathFromFtpOfFutRic = "5225" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M";
            strFilePathFromFtpOfSpdRic = "5673" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M";
            service = MSAD.Common.OfficeUtility.EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));

        }
    #endregion

        protected override void Start()
        {
            GetDataToDicFromFtp(ref isExistFirstFileOnFtp, strFilePathFromFtpOfFutRic, strFutPattern, dicFutRicFtp);
            GetDataToDicFromFtp(ref isExistSecondFileOnFtp, strFilePathFromFtpOfSpdRic, strSpdPattern, dicSpdRicFtp);
            GetDataToFutListFromLocalFile(strFutRicFilePath);
            GetDataToSpdListFromLocalFile(strSpdRicFilePath);
            RemoveExistRicAndSendEmail(listFutRic, dicFutRicFtp, listSpdRic, dicSpdRicFtp);
        }

        #region RemoveExistRicAndSendEmail
        /// <summary>
        /// remove exist and send email
        /// </summary>
        /// <param name="listFutRic">futRic</param>
        /// <param name="dicFutRicFtp">futDic</param>
        /// <param name="listSpdRic">spdRic</param>
        /// <param name="dicSpdRicFtp">spdDic</param>
        private void RemoveExistRicAndSendEmail(List<string> listFutRic, Dictionary<string, string> dicFutRicFtp, List<string> listSpdRic, Dictionary<string, string> dicSpdRicFtp)
        {
            if (isExistFirstFileOnFtp && isExistSecondFileOnFtp)
            {
                string subject = string.Empty;
                string content = string.Empty;
                if ((listFutRic != null) && (listSpdRic != null))
                {
                    if (dicFutRicFtp.Count > 0)
                    {
                        int indexFut = listFutRic.Count;
                        for (int i = indexFut - 1; i >= 0; i--)
                        {
                            if (dicFutRicFtp.ContainsKey(listFutRic[i]))
                            {
                                listFutRic.Remove(listFutRic[i]);
                            }
                        }
                        int indexSpd = listSpdRic.Count;
                        for (int i = indexSpd - 1; i >= 0; i--)
                        {
                            if (dicSpdRicFtp.ContainsKey(listSpdRic[i]))
                            {
                                listSpdRic.Remove(listSpdRic[i]);
                            }
                        }
                        if ((listFutRic.Count == indexFut) && (listSpdRic.Count == indexSpd)) //send Email with no missing Ric on the ftp 2
                        {
                            subject = "TW QC Future Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[FutRic.txt and SpdRic.txt No Exist Missing Ric]";
                            content = "<center>********All FutRic********</center><br /><br /><p>RIC</p>";
                            foreach (string str in listFutRic)
                            {
                                content += string.Format("{0}", str);
                                content += "<br />";
                            }
                        }
                        else if ((listFutRic.Count != indexFut) && (listSpdRic.Count != indexSpd)) //send Email with exisi missing ric from ftp 2
                        {
                            subject = "TW QC Future Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[FutRic.txt and SpdRic.txt Exist Missing Ric]";
                            content = "<center>********Missing FutRic********</center><br /><br /><p>RIC</p>";
                            foreach (string str in listFutRic)
                            {
                                content += string.Format("{0}", str);
                                content += "<br />";
                            }
                        }
                        else if ((listFutRic.Count == indexFut) && (listSpdRic.Count != indexSpd)) //send Email with FutRic no missing SpdRic exist missing
                        {
                            subject = "TW QC Future Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[FutRic.txt No Missing Ric And SpdRic.txt Exist Missing Ric]";
                            content = "<center>********All FutRic********</center><br /><br /><p>RIC</p>";
                            foreach (string str in listFutRic)
                            {
                                content += string.Format("{0}", str);
                                content += "<br />";
                            }
                        }
                        else  //send Email with FutRic exist missing SpdRic no missing
                        {
                            subject = "TW QC Future Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[FutRic.txt Missing Ric And SpdRic.txt Exist Missing Ric]";
                            content = "<center>********All FutRic********</center><br /><br /><p>RIC</p>";
                            foreach (string str in listFutRic)
                            {
                                content += string.Format("{0}", str);
                                content += "<br />";
                            }
                        }
                    }
                    else
                    {
                        subject = "TW QC Future Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[FutRic.txt and SpdRic.txt No Exist Missing Ric]";
                        content = "<center>********All FutRic********</center><br /><br /><p>RIC</p>";
                        foreach (string str in listFutRic)
                        {
                            content += string.Format("{0}", str);
                            content += "<br />";
                        }
                    }
                    SendMail(service, subject, content, attacheFileList);
                }
            }
        }
        #endregion

        #region ReadTxtFileToFut
        /// <summary>
        /// read txt file get data to list
        /// </summary>
        /// <param name="listFutRic">list</param>
        /// <param name="strFilePathFromFtpOfFutRic">path</param>
        private void GetDataToFutListFromLocalFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                attacheFileList.Add(filePath);
                FileStream fs = new FileStream(filePath, FileMode.Open);
                StreamReader sr = new StreamReader(fs);
                listFutRic = new List<string>(sr.ReadToEnd().Split(','));
                sr.Close();
                fs.Close();
            }
        }
        #endregion

        #region ReadTxtFileToSpd
        /// <summary>
        /// read txt file get data to list
        /// </summary>
        /// <param name="listFutRic">list</param>
        /// <param name="strFilePathFromFtpOfFutRic">path</param>
        private void GetDataToSpdListFromLocalFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                attacheFileList.Add(filePath);
                FileStream fs = new FileStream(filePath, FileMode.Open);
                StreamReader sr = new StreamReader(fs);
                listSpdRic = new List<string>(sr.ReadToEnd().Split(','));
                sr.Close();
                fs.Close();
            }
        }
        #endregion

        #region GetDataFromFtp
        /// <summary>
        /// Get Data from ftp
        /// </summary>
        /// <param name="filePath">path</param>
        /// <param name="pattern">regular expression</param>
        /// <param name="dic">dic store data</param>
        private void GetDataToDicFromFtp(ref bool isExistFileOnFtp, string filePath, string pattern, Dictionary<string, string> dic)
        {
            try
            {
                string filePathFromFtp = string.Empty;
                filePathFromFtp = @"ftp://ASIA2:ASIA2@ds1.rds.reuters.com//" + filePath;
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(filePathFromFtp);
                WebProxy proxy = new WebProxy("10.40.14.56", 80);
                request.Proxy = proxy;
                WebResponse res = request.GetResponse();
                StreamReader sr = new StreamReader(res.GetResponseStream());
                string tmp = null;
                string strRic = string.Empty;
                while ((tmp = sr.ReadLine()) != null)
                {
                    Regex r = new Regex(pattern);
                    MatchCollection mc = r.Matches(tmp.Trim());
                    if (mc.Count == 1)
                    {
                        if (filePath.Substring(0, 4).Equals("5225"))
                        {
                            strRic = mc[0].Groups["RIC"].Value.Substring(2, mc[0].Groups["RIC"].Value.Length - 2);
                            if (!dic.ContainsKey(strRic))
                            {
                                dic.Add(strRic, "");
                            }
                        }
                        else
                        {
                            strRic = mc[0].Groups["RIC"].Value.Substring(2, mc[0].Groups["RIC"].Value.Length - 2);
                            if (!dic.ContainsKey(strRic))
                            {
                                if (!strRic.Substring(strRic.Length - 2, 2).Equals("^1"))
                                {
                                    dic.Add(strRic, "");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                isExistFileOnFtp = false;
                string subject = string.Empty;
                string content = string.Empty;
                subject = "TTW QC Option Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[No file on ftp]";
                content = "<center>*******Can not find the file:[" + filePath + "] on ftp*********</center><br />";
                SendMail(service, subject, content, attacheFileList);
                Logger.Log("error when read file on ftp!");
            }
        }
        #endregion

        #region SendMail
        /// <summary>
        /// SendMail
        /// </summary>
        /// <param name="service">Login Email</param>
        /// <param name="subject">subject</param>
        /// <param name="content">Body</param>
        /// <param name="attacheFileList">Attachements</param>
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
        #endregion
    }
}
