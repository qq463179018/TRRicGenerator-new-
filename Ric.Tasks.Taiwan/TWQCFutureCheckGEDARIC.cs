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
    public class TWQCFutureCheckGEDARICConfig
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
    class TWQCFutureCheckGEDARIC : GeneratorBase
    {
        public TWQCFutureCheckGEDARICConfig configObj = null;
        private string accountName = string.Empty;//UC169XXX
        private string password = string.Empty;//********
        private string domain = string.Empty;//TEN
        private string mailAdress = string.Empty;//eti.XXXXXX@thomsonreuters.com
        private string mailFolder = string.Empty;//Inbox/XXXXX
        private string txtFilePath = string.Empty;
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private List<string> listRicChain = new List<string>();//20
        private Dictionary<string, string> dicRicChainExistGATS = new Dictionary<string, string>();
        private List<string> listRicChainMissing = new List<string>();//20
        private List<string> listNewRicChain = new List<string>();//20
        private Dictionary<string, string> dicRic = new Dictionary<string, string>();//2000
        private List<string> listNewRic = new List<string>();//3000
        private ExchangeService service;
        private DateTime startDate = DateTime.Now.ToUniversalTime().AddHours(+8).AddHours(-DateTime.Now.ToUniversalTime().AddHours(+8).Hour);
        private DateTime endDate = DateTime.Now.ToUniversalTime().AddHours(+8);
        private string strEmailRicChain = string.Empty;
        private string strRicChainPattern = string.Empty;
        private string strRicChainFileName = string.Empty;
        private string strRicChainMissingPattern = string.Empty;
        private string strRicDicPattern = string.Empty;
        private string strRicFileName = string.Empty;
        private List<string> listHKGRic = new List<string>();
        private string strRicMTXW = string.Empty;
        private string strRicMTX = string.Empty;
        private List<string> listMTXW = new List<string>();
        private List<string> listMTX = new List<string>();
        private string strHKGRicPattern = string.Empty;
        private string strEmailKeyWordListRicChain = string.Empty;
        private string strEmailKeyWordHKGRic = string.Empty;
        private string strMTXWRicPattern = string.Empty;
        private string strMTXRicPattern = string.Empty;
        private bool isExistEmail = false;
        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as TWQCFutureCheckGEDARICConfig;
            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());
            accountName = emailAccount.AccountName;
            password = emailAccount.Password;
            domain = emailAccount.Domain;
            mailAdress = emailAccount.MailAddress;
            mailFolder = configObj.MailFolderPath.Trim();
            txtFilePath = configObj.TxtFilePath.Trim();
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            strRicChainPattern = @"\](?<RIC>[A-Za-z0-9]+)\b\s+\b\d{2}\-[A-Z]{3}\-\d{4}\b";
            strRicChainFileName = DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + "FutRic.txt";
            strRicChainMissingPattern = @"\r\n(?<RIC>[A-Za-z0-9]+)\b\s+\bOFFCL_CODE\b\s+\b(?<Value>[A-Za-z0-9]+)\b";
            strRicDicPattern = @"\bLONGLINK\d{1,2}\b\s+\b(?<RIC>[A-Za-z0-9]+)\r\n\S+\b";
            strRicFileName = DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + "SpdRic.txt";
            strRicMTXW = "0#MTXW:";
            strRicMTX = "0#MTX:";
            strEmailKeyWordListRicChain = "Futures Job log file for RA(Region-HKG|Exchanges-TM|TAIFEX)";
            strEmailKeyWordHKGRic = "Futures Job log file for RA(Region-HKG)";
            strHKGRicPattern = @"\*\s+\*\s+\b(?<RIC>[A-Za-z0-9]+)\b";
            service = MSAD.Common.OfficeUtility.EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
        }
    #endregion

        protected override void Start()
        {
            ReadEmailTolistRicChain(listRicChain, strRicChainPattern, strEmailKeyWordListRicChain);
            GenerateFile(listRicChain, strRicChainFileName);
            GetRicChainMissingFromGATS(listRicChain, listRicChainMissing);
            GenerateNewRicChainByRicChain(listRicChain, listNewRicChain);
            GetGroupedDicRicFromGATSByNewRicChain(listNewRicChain, dicRic);
            GetDicNewRic(dicRic, listNewRic);
            ReadEmailTolistRicChain(listHKGRic, strHKGRicPattern, strEmailKeyWordHKGRic);
            GetDataListMTXWAndListMTX(listMTXW, listMTX, listHKGRic, listNewRic);
            GenerateFile(listNewRic, strRicFileName);
            SendEmailThreeCase(listRicChain, listRicChainMissing, listNewRic);
        }

        #region GetDataListMTXWAndListMTX
        /// <summary>
        /// GetDataListMTXWAndListMTX
        /// </summary>
        /// <param name="listMTXW">get data from gats</param>
        /// <param name="listMTX">get data from gats</param>
        /// <param name="listHKGRic">get from email</param>
        private void GetDataListMTXWAndListMTX(List<string> listMTXW, List<string> listMTX, List<string> listHKGRic, List<string> listNewRic)
        {
            string fids = null;
            GetDataFromGATSToExistList(strRicDicPattern, strRicMTXW, fids, listMTXW);//fill listMTXW
            GetDataFromGATSToExistList(strRicDicPattern, strRicMTX, fids, listMTX);//fill listMTX
            if (listMTXW.Contains("MTXW3F4") && listMTXW.Count == 2)
            {
                listNewRic.Add("MTXW4F4-3F4");
            }
            foreach (string str in listMTX)
            {
                if (listHKGRic.Count == 1)
                {
                    string strHKGRic = listHKGRic[0].ToString();
                    if ((str.Substring(str.Length - 2, 2)).Equals(strHKGRic.Substring(str.Length - 2, 2)))
                    {
                        listNewRic.Add(str + "-" + strHKGRic.Substring(strHKGRic.Length - 4, 4));
                    }
                    else
                    {
                        listNewRic.Add(strHKGRic + "-" + str.Substring(str.Length - 2, 2));
                    }
                }
            }
        }
        #endregion

        #region SendEmailThreeCase
        /// <summary>
        /// SendEmail with three cases
        /// </summary>
        /// <param name="listRicChain">FromEmail</param>
        /// <param name="listRicChainMissing">MissingRic</param>
        /// <param name="listNewRic">NewRic</param>
        private void SendEmailThreeCase(List<string> listRicChain, List<string> listRicChainMissing, List<string> listNewRic)
        {
            string subject = string.Empty;
            string content = string.Empty;
            List<string> attacheFileList = new List<string>();
            if (listRicChain.Count > 0)
            {
                string attachedRicChainFile = Path.Combine(txtFilePath, strRicChainFileName);
                attacheFileList.Add(attachedRicChainFile);
                if (listRicChainMissing.Count == 0)
                {
                    subject = "TW QC Future Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Email AND no Missing FutRic]";
                    content = "<center>********All the FutRic are exist in the GATS********</center><br /><br />";
                }
                else
                {
                    string attachedRicFile = Path.Combine(txtFilePath, strRicFileName);
                    attacheFileList.Add(attachedRicFile);
                    subject = "TW QC Future Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Email AND Exist Missing FutRic]";
                    content = "<center>********Missing FutRic********</center><br /><br /><p>FutRIC</p>";
                    foreach (string str in listRicChainMissing)
                    {
                        content += string.Format("{0}", str);
                        content += "<br />";
                    }
                }
            }
            else if (isExistEmail)
            {
                subject = "TW QC Future Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Email BUT Email Is Empty]";
                content = "<center>********Empty Email********</center><br /><br />";
            }
            else
            {
                subject = "TW QC Future Check GEDA RIC Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[no  Email]";
                content = "<center>********No Email********</center><br /><br />";
            }
            SendMail(service, subject, content, attacheFileList);
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

        #region GetDicNewRic
        /// <summary>
        /// GetDicNewRic
        /// </summary>
        /// <param name="dicRic">GroupedRic</param>
        /// <param name="dicNewRic">NewGeneratdRicByGroupedRic</param>
        private void GetDicNewRic(Dictionary<string, string> dicRic, List<string> listNewRic)
        {
            List<string> listTmpLinq = new List<string>();
            Dictionary<string, string>.ValueCollection valueCol = dicRic.Values;
            foreach (string value in valueCol)// get All grouped Value in the dicRic
            {
                if (!listTmpLinq.Contains(value))
                {
                    listTmpLinq.Add(value);
                }
            }
            foreach (string value in listTmpLinq)//Generate listNewRic
            {
                List<string> listTmp = new List<string>();
                foreach (KeyValuePair<string, string> kvp in dicRic)//Get the key with the same value
                {
                    if (kvp.Value == value)
                    {
                        listTmp.Add(kvp.Key);
                    }
                }
                for (int i = 0; i < listTmp.Count; i++)
                {
                    for (int j = i; j < listTmp.Count - 1; j++)
                    {
                        listNewRic.Add(listTmp[i].ToString() + "-" + listTmp[j + 1].Substring(listTmp[j + 1].Length - 2, 2));//generate newRic 
                    }
                }
            }
        }
        #endregion

        #region GetGroupedDicRicFromGATSByNewRicChain
        /// <summary>
        /// GetRicFromGATSByNewRicChain 
        /// </summary>
        /// <param name="listNewRicChain">RicChain to NewRicChain</param>
        /// <param name="dicRic">Group By Value</param>
        private void GetGroupedDicRicFromGATSByNewRicChain(List<string> listNewRicChain, Dictionary<string, string> dicRic)//in the Dictory<key,value> group by value
        {
            if (listNewRicChain.Count > 0)
            {
                string content = string.Empty;
                foreach (var str in listNewRicChain)
                {
                    content += string.Format(",{0}", str);
                }
                content = content.Remove(0, 1);
                string offclCode = null;
                GetDataFromGATSToExistDic(strRicDicPattern, content, offclCode, dicRic);
                if (dicRic.Count > 0)
                {
                    string[] KeyCol = new string[dicRic.Keys.Count];
                    dicRic.Keys.CopyTo(KeyCol, 0);
                    foreach (string key in KeyCol)
                    {
                        dicRic[key] = key.Substring(0, key.Length - 2);
                    }
                }
            }
        }
        #endregion

        #region GenerateNewRicChainByRicChain
        /// <summary>
        /// GenerateNewRicChainByRicChain
        /// </summary>
        /// <param name="listRicChain">GetFromEmail</param>
        /// <param name="listNewRicChain">Generate NewRicChain</param>
        private void GenerateNewRicChainByRicChain(List<string> listRicChain, List<string> listNewRicChain)
        {
            if (listRicChain.Count > 0)
            {
                string strNewRic = string.Empty;
                foreach (var str in listRicChain)
                {
                    strNewRic = "0#" + str.Substring(0, str.Length - 2) + ":";
                    listNewRicChain.Add(strNewRic);
                }
            }
        }
        #endregion

        #region GetRicChainMissingFromGATS
        /// <summary>
        /// GetRicChainMissingFromGATS
        /// </summary>
        /// <param name="listRicChain">GetFromEmail</param>
        /// <param name="listRicChainMissing">RicNotExistInGATS</param>
        private void GetRicChainMissingFromGATS(List<string> listRicChain, List<string> listRicChainMissing)
        {
            if (listRicChain.Count > 0)
            {
                string content = string.Empty;
                foreach (var str in listRicChain)
                {
                    content += string.Format(",{0}", str);
                }
                content = content.Remove(0, 1);
                string offclCode = "OFFCL_CODE";
                GetDataFromGATSToExistDic(strRicChainMissingPattern, content, offclCode, dicRicChainExistGATS);
                foreach (var str in listRicChain)
                {
                    if (!dicRicChainExistGATS.ContainsKey(str))
                    {
                        listRicChainMissing.Add(str);
                    }
                }
            }
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
                    TaskResultList.Add(new TaskResultEntry("TWQCFutureNewRIC", "RicTxtFile", filePath));
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

        #region ReadEmailTolistRicChain
        /// <summary>
        /// ReadEmailTolistRicChain
        /// </summary>
        /// <param name="listRicChain"></param>
        /// <param name="strPattern">pattern</param>
        /// <param name="strEmailKeyWord">Email title</param>
        private void ReadEmailTolistRicChain(List<string> listRicChain, string strPattern, string strEmailKeyWord)
        {
            try
            {
                EWSMailSearchQuery query = new EWSMailSearchQuery("", mailAdress, mailFolder, strEmailKeyWord, "", startDate, endDate);
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
                    isExistEmail = true;
                    EmailMessage mail = mailList[0];
                    mail.Load();
                    strEmailRicChain =TWHelper.ClearHtmlTags(mail.Body.ToString());
                    Regex regex = new Regex(strPattern);
                    MatchCollection matches = regex.Matches(strEmailRicChain);
                    foreach (Match match in matches)
                    {
                        listRicChain.Add(match.Groups["RIC"].Value);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Get Data from mail failed. Ex: " + ex.Message);
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

        #region GetDataFromGATSToExistList
        /// <summary>
        /// GetDataFromGATSToExistList
        /// </summary>
        /// <param name="pattern">Regex</param>
        /// <param name="rics">Parms</param>
        /// <param name="fids">Parms</param>
        /// <param name="listExist">Store Data From GATS</param>
        private void GetDataFromGATSToExistList(string pattern, string rics, string fids, List<string> listExist)
        {
            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(rics, fids);
            Regex regex = new Regex(pattern);
            MatchCollection matches = regex.Matches(response);
            foreach (Match match in matches)
            {
                listExist.Add(match.Groups["RIC"].Value);
            }
        }
        #endregion
    }
}