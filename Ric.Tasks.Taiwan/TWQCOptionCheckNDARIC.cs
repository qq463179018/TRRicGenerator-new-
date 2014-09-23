using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Exchange.WebServices.Data;
using MSAD.Common.OfficeUtility;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;

namespace Ric.Tasks.Taiwan
{
    #region Configuration
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    class TWQCOptionCheckNDARICConfig
    {
        [StoreInDB]
        [Category("File")]
        [DisplayName("File path")]
        [Description("Generated Files Path,like: G:\\China")]
        public string FilePath { get; set; }

        [StoreInDB]
        [Category("Email Account")]
        [DisplayName("Name")]
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
    }
    #endregion

    #region Description
    class TWQCOptionCheckNDARIC : GeneratorBase
    {
        public TWQCOptionCheckNDARICConfig configObj = null;
        private string filePath = string.Empty;//G://XX
        private string accountName = string.Empty;//UC169XXX
        private string password = string.Empty;//********
        private string domain = string.Empty;//TEN
        private string mailAdress = string.Empty;//eti.XXXXXX@thomsonreuters.com
        private string mailFolder = string.Empty;//Inbox/XXXXX
        private ExchangeService service = null;
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private List<string> attacheFileList = new List<string>();
        private string strFilePathFromFtpOfRic = string.Empty;
        List<string> listRic = null;
        Dictionary<string, string> dicRicFtp = new Dictionary<string, string>();
        private bool isExistOnFtp = true;
        private string strPattern = string.Empty;
        private string strRicFilePath = string.Empty;
        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as TWQCOptionCheckNDARICConfig;
            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());
            accountName = emailAccount.AccountName;
            password = emailAccount.Password;
            domain = emailAccount.Domain;
            mailAdress = emailAccount.MailAddress;
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            filePath = configObj.FilePath;
            strPattern = @"^\b\w{2}(?<RIC>\w+\.[A-Z]{2,3})\b\s+?";
            strRicFilePath = Path.Combine(filePath, "OptionRic.txt");
            strFilePathFromFtpOfRic = "7137" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M";
        }
    #endregion

        protected override void Start()
        {
            GetDataToDicFromFtp(strFilePathFromFtpOfRic, strPattern, dicRicFtp);
            GetDataToListFromLocalFile(strRicFilePath);
            RemoveExistRicAndSendEmail(listRic, dicRicFtp);
        }

        #region RemoveExistRicAndSendEmail
        /// <summary>
        /// RemoveExistOptionRicAndSendEmail
        /// </summary>
        /// <param name="listRic">Ric from txt</param>
        /// <param name="dicRicFtp">Ric from ftp</param>
        private void RemoveExistRicAndSendEmail(List<string> listRic, Dictionary<string, string> dicRicFtp)
        {
            if (isExistOnFtp)
            {
                string subject = string.Empty;
                string content = string.Empty;
                if (listRic != null)
                {
                    if (dicRicFtp.Count > 0)
                    {
                        int indexFut = listRic.Count;
                        for (int i = indexFut - 1; i >= 0; i--)//remove exist ric from list
                        {
                            if (dicRicFtp.ContainsKey(listRic[i]))
                            {
                                listRic.Remove(listRic[i]);
                            }
                        }
                        if ((listRic.Count == 1) && string.IsNullOrEmpty(listRic[0].Trim()))//send Email with empty txt file 
                        {
                            subject = "TW QC Option Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[OptionRic.txt is Empty]";
                            content = "<center>********All OptionRic********</center><br /><br /><p>RIC</p>";
                        }
                        else if (listRic.Count == indexFut) //send Email with no missing OptionRic on the ftp
                        {
                            subject = "TW QC Option Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[OptionRic.txt No Missing Ric]";
                            content = "<center>********All OptionRic********</center><br /><br /><p>RIC</p>";
                            foreach (string str in listRic)
                            {
                                content += string.Format("{0}", str);
                                content += "<br />";
                            }
                        }
                        else if (listRic.Count != indexFut)  //send Email with exisi missing OptionRic from ftp
                        {
                            subject = "TW QC Option Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[OptionRic.txt Exist Missing Ric]";
                            content = "<center>********Missing OptionRic********</center><br /><br /><p>RIC</p>";
                            foreach (string str in listRic)
                            {
                                content += string.Format("{0}", str);
                                content += "<br />";
                            }
                        }
                    }
                    else
                    {
                        subject = "TW QC Option Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[OptionRic.txt Exist No Missing Ric]";
                        content = "<center>********All OptionRic********</center><br /><br /><p>RIC</p>";
                        foreach (string str in listRic)
                        {
                            content += string.Format("{0}", str);
                            content += "<br />";
                        }
                    }
                    SendMail(subject, content, attacheFileList);
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
        private void GetDataToListFromLocalFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                attacheFileList.Add(filePath);
                FileStream fs = new FileStream(filePath, FileMode.Open);
                StreamReader sr = new StreamReader(fs);
                listRic = new List<string>(sr.ReadToEnd().Split(','));
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
        private void GetDataToDicFromFtp(string filePath, string pattern, Dictionary<string, string> dic)
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
                        strRic = mc[0].Groups["RIC"].Value;
                        if (!dic.ContainsKey(strRic))
                        {
                            dic.Add(strRic, "");
                        }
                    }
                }
            }
            catch (Exception)
            {
                string subject = string.Empty;
                string content = string.Empty;
                isExistOnFtp = false;
                subject = "TTW QC Option Check NDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[No file on ftp]";
                content = "<center>*******Can not find the file:[" + filePath + "] on ftp*********</center><br />";
                SendMail(subject, content, attacheFileList);
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
        private void SendMail(string subject, string content, List<string> attacheFileList)
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
            service = EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
            if (configObj.MailCC.Count > 1 || (configObj.MailCC.Count == 1 && configObj.MailCC[0] != ""))
            {
                listMailCC = configObj.MailCC;
            }
            EWSUtility.CreateAndSendMail(service, listMailTo, listMailCC, new List<string>(), subject, content, attacheFileList);
        }
        #endregion
    }
}
