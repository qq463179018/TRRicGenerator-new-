using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;
using MSAD.Common.OfficeUtility;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks.Taiwan
{
    #region Configuration
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class TWQCOptionCheckGEDARICConfig
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
        [DisplayName("Txt file path")]
        [Description("GeneratedFilePath")]
        public string TxtFilePath { get; set; }

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
    class TWQCOptionCheckGEDARIC : GeneratorBase
    {
        public TWQCOptionCheckGEDARICConfig configObj = null;
        private string accountName = string.Empty;//UC169XXX
        private string password = string.Empty;//********
        private string domain = string.Empty;//TEN
        private string mailAdress = string.Empty;//eti.XXXXXX@thomsonreuters.com
        private string mailFolder = string.Empty;//Inbox/XXXXX
        private string txtFilePath = string.Empty;
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private List<string> listRic = new List<string>();//20
        private int countListRic = 0;
        private Dictionary<string, string> dicRicExistGATS = new Dictionary<string, string>();
        private List<string> listRicMissing = new List<string>();//20
        private ExchangeService service = null;
        private DateTime startDate = DateTime.Now.ToUniversalTime().AddHours(+8).AddHours(-DateTime.Now.ToUniversalTime().AddHours(+8).Hour);
        private DateTime endDate = DateTime.Now.ToUniversalTime().AddHours(+8);
        private string strEmailRic = string.Empty;
        private string strRicFirstEmailPattern = string.Empty;
        private string strRicSecondEmailPattern = string.Empty;
        private string strRicGATSPattern = string.Empty;
        private string strRicFileName = string.Empty;
        private string strFirstEmailKeyWordListRic = string.Empty;
        private string strSecondEmailKeyWordListRic = string.Empty;
        private bool isExistFirstEmail = false;
        private bool isExistSecondEmail = false;
        private bool isExistFirstEmptyEmail = false;
        private bool isExistSeconEmptydEmail = false;
        private List<string> attacheFileList = new List<string>();

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as TWQCOptionCheckGEDARICConfig;
            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());
            accountName = emailAccount.AccountName;
            password = emailAccount.Password;
            domain = emailAccount.Domain;
            mailAdress = emailAccount.MailAddress;
            mailFolder = configObj.MailFolderPath.Trim();
            txtFilePath = configObj.TxtFilePath.Trim();
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            strRicFirstEmailPattern = @"RIC\s+\[(?<RIC>\w+\.[A-Z]{2,3})\]\s+for\s+Exch\s+symbol";
            strRicSecondEmailPattern = @"\b(?<RIC>\w+[A-Z]{1}\d{1}\.TM)\b";
            strRicGATSPattern = @"\r\n(?<RIC>\w+\.[A-Z]{2,3})\b\s+\bOFFCL_CODE\b\s+\b(?<Value>\w+)";
            strRicFileName = "OptionRic.txt";
            strFirstEmailKeyWordListRic = "TAIFO IFFM - EDA Automation Report";
            strSecondEmailKeyWordListRic = "TAIFEX_OPT_IFFM Report for Weekly Option of Taifex option iffm feed";
        }
    #endregion

        protected override void Start()
        {
            ReadEmailTolistRicChain(listRic, strRicFirstEmailPattern, strFirstEmailKeyWordListRic);
            ReadEmailTolistRicChain(listRic, strRicSecondEmailPattern, strSecondEmailKeyWordListRic);
            GenerateFile(listRic, strRicFileName);
            GetRicMissingFromGATS(listRic, listRicMissing);
            SendEmailThreeCase(listRic, listRicMissing);
        }

        #region SendEmailWithFourCase
        /// <summary>
        /// SendEmail
        /// </summary>
        /// <param name="listRic">list from email</param>
        /// <param name="listRicMissing">list from GATS</param>
        private void SendEmailThreeCase(List<string> listRic, List<string> listRicMissing)
        {
            string subject = string.Empty;
            string content = string.Empty;
            if (isExistFirstEmail && isExistSecondEmail)//exist+exist
            {
                if (isExistFirstEmptyEmail && isExistSeconEmptydEmail)//empty+empty
                {
                    subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Two Empty Email]";
                    content = "<center>********Empty Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********Empty Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********No New Ric********</center><br /><br />";
                }
                else if ((!isExistFirstEmptyEmail) && (!isExistSeconEmptydEmail))//yes+yes
                {
                    if (listRicMissing.Count > 0)//yes+yes + missing ric
                    {
                        subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Two Email]";
                        content = "<center>********Found Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********Found Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********Missing Ric********</center><br /><br /><p>RIC</p>";
                        foreach (string str in listRicMissing)
                        {
                            content += string.Format("{0}", str);
                            content += "<br />";
                        }
                    }
                    else//yes + yes + no missing ric
                    {
                        subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Two Email]";
                        content = "<center>********Found Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********Found Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********No Missing Ric********</center><br /><br /><p>RIC</p>";
                    }
                }
                else if (isExistFirstEmptyEmail)//empty+yes
                {
                    if (listRicMissing.Count > 0)//empty+yes + missing ric
                    {
                        subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Two Email]";
                        content = "<center>********Found Empty Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********Found Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********Missing Ric********</center><br /><br /><p>RIC</p>";
                        foreach (string str in listRicMissing)
                        {
                            content += string.Format("{0}", str);
                            content += "<br />";
                        }
                    }
                    else//empty + yes + no missing ric
                    {
                        subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Two Email]";
                        content = "<center>********Found Empty Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********Found Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********No Missing Ric********</center><br /><br /><p>RIC</p>";
                    }
                }
                else if (isExistSeconEmptydEmail)//yes+empty
                {
                    if (listRicMissing.Count > 0)//yes+empty + missing ric
                    {
                        subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Two Email]";
                        content = "<center>********Found Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********Found Empty Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********Missing Ric********</center><br /><br /><p>RIC</p>";
                        foreach (string str in listRicMissing)
                        {
                            content += string.Format("{0}", str);
                            content += "<br />";
                        }
                    }
                    else//yes+empty + no missing ric
                    {
                        subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Two Email]";
                        content = "<center>********Found Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********Found Empty Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********No Missing Ric********</center><br /><br /><p>RIC</p>";
                    }
                }
            }
            else if ((!isExistFirstEmail) && (!isExistSecondEmail))//no + no
            {
                subject = "No TAIFO IFFM - EDA Automation Report for today  " + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Neither Email]";
                content = "<center>********No Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********No Email of " + strSecondEmailKeyWordListRic + "********</center><br />";
            }
            else if (isExistFirstEmail)//first exist
            {
                if (isExistFirstEmptyEmail)//empty+no
                {
                    subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found One Empty Email]";
                    content = "<center>********Found Empty Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********No Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********Missing Ric********</center><br /><br /><p>RIC</p>";
                }
                else//yes +no
                {
                    if (listRicMissing.Count > 0)//yes +no + missing ric
                    {
                        subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found One Email]";
                        content = "<center>********Found Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********No Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********Missing Ric********</center><br /><br /><p>RIC</p>";
                        foreach (string str in listRicMissing)
                        {
                            content += string.Format("{0}", str);
                            content += "<br />";
                        }
                    }
                    else//yes + no + no missing ric
                    {
                        subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found One Email]";
                        content = "<center>********Found Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********No Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********No Missing Ric********</center><br /><br /><p>RIC</p>";
                    }
                }
            }
            else if (isExistSecondEmail)//second exist
            {
                if (listRicMissing.Count > 0)//no +yes + missing ric
                {
                    subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found One Email]";
                    content = "<center>********Found Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********No Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********Missing Ric********</center><br /><br /><p>RIC</p>";
                    foreach (string str in listRicMissing)
                    {
                        content += string.Format("{0}", str);
                        content += "<br />";
                    }
                }
                else//no + yes + no missing ric
                {
                    subject = "TW QC Option Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found One Email]";
                    content = "<center>********Found Empty Email of " + strSecondEmailKeyWordListRic + "********</center><br /><center>********No Email of " + strFirstEmailKeyWordListRic + "********</center><br /><center>********No Missing Ric********</center><br /><br /><p>RIC</p>";
                }
            }
            SendMail(service, subject, content, attacheFileList);//sendEmail()
        }
        #endregion

        #region ReadEmailGetData
        /// <summary>
        /// ReadEmailAndGetData
        /// </summary>
        /// <param name="listRic">listRic</param>
        /// <param name="strPattern">regular expression</param>
        /// <param name="strEmailKeyWord">subject of email</param>
        private void ReadEmailTolistRicChain(List<string> listRic, string strPattern, string strEmailKeyWord)
        {
            try
            {
                service = EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
                EWSMailSearchQuery query = null;
                query = new EWSMailSearchQuery("", mailAdress, mailFolder, strEmailKeyWord, "", startDate, endDate);
                List<EmailMessage> mailList = null;
                for (int i = 0; i < 5; i++)
                {
                    try
                    {
                        mailList = EWSMailSearchQuery.SearchMail(service, query);
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(5000);
                        if (i == 4)
                        {
                            throw;
                        }
                    }
                }
                if (mailList.Count > 0)
                {
                    if (strEmailKeyWord.Equals("TAIFO IFFM - EDA Automation Report"))
                    {
                        isExistFirstEmail = true;
                    }
                    else
                    {
                        isExistSecondEmail = true;
                    }
                    EmailMessage mail = mailList[0];
                    mail.Load();
                    strEmailRic = TWHelper.ClearHtmlTags(mail.Body.ToString());
                    Regex regex = new Regex(strPattern);
                    MatchCollection matches = regex.Matches(strEmailRic);
                    foreach (Match match in matches)
                    {
                        listRic.Add(match.Groups["RIC"].Value);
                    }
                    if (strEmailKeyWord.Equals("TAIFO IFFM - EDA Automation Report"))
                    {
                        countListRic = listRic.Count;
                        if (listRic.Count == 0)
                        {
                            isExistFirstEmptyEmail = true;
                        }
                    }
                    else
                    {
                        if (listRic.Count == countListRic)
                        {
                            isExistSeconEmptydEmail = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Get Data from mail failed. Ex: " + ex.Message);
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
            EWSUtility.CreateAndSendMail(service, listMailTo, listMailCC, new List<string>(), subject, content, attacheFileList);
        }
        #endregion

        #region GenerateFile
        /// <summary>
        /// GenerateFile
        /// </summary>
        /// <param name="list">DateType</param>
        /// <param name="fileName">fileName</param>
        private void GenerateFile(List<string> list, string fileName)
        {
            string filePath = Path.Combine(txtFilePath, fileName);
            if (list.Count > 0)
            {
                string content = string.Empty;
                foreach (var str in list)
                {
                    content += string.Format(",{0}", str);
                }
                content = content.Remove(0, 1);
                try
                {
                    File.WriteAllText(filePath, content);
                    attacheFileList.Add(filePath);
                    AddResult("TW QC - Option New RIC (CQF 403)", filePath, "file");
                    //TaskResultList.Add(new TaskResultEntry("TW QC - Option New RIC (CQF 403)", "RicTxtFile", filePath));
                }
                catch (Exception ex)
                {
                    Logger.Log(string.Format("Error happens when generating file. Ex: {0} .", ex.Message));
                }
            }
            else
            {
                File.Delete(filePath);
            }
        }
        #endregion

        #region GetRicMissingFromGATS
        /// <summary>
        /// GetRicMissingFromGATS
        /// </summary>
        /// <param name="listRicChain">GetFromEmail</param>
        /// <param name="listRicChainMissing">RicNotExistInGATS</param>
        private void GetRicMissingFromGATS(List<string> listRic, List<string> listRicMissing)
        {
            if (listRic.Count > 0)
            {
                string content = string.Empty;
                foreach (var str in listRic)
                {
                    content += string.Format(",{0}", str);
                }
                content = content.Remove(0, 1);
                string offclCode = "OFFCL_CODE";
                GetDataFromGATSToExistDic(strRicGATSPattern, content, offclCode, dicRicExistGATS);
                foreach (var str in listRic)
                {
                    if (!dicRicExistGATS.ContainsKey(str))
                    {
                        listRicMissing.Add(str);
                    }
                }
            }
        }
        #endregion

        #region GetDataFromGATSToExistDic
        /// <summary>
        /// GetDataFromGATSToExistDic
        /// </summary>
        /// <param name="pattern">Regex</param>
        /// <param name="rics">Parms</param>
        /// <param name="fids">Parms</param>
        /// <param name="dicExist">Store Data From GATS</param>
        private void GetDataFromGATSToExistDic(string pattern, string rics, string fids, Dictionary<string, string> dicExist)
        {
            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(rics, fids);
            Regex regex = new Regex(pattern);
            MatchCollection matches = regex.Matches(response);
            foreach (Match match in matches)
            {
                if (!dicExist.ContainsKey(match.Groups["RIC"].Value))
                {
                    dicExist.Add(match.Groups["RIC"].Value, match.Groups["Value"].Value);
                }
            }
        }
        #endregion
    }
}
