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

namespace Ric.Tasks.HongKong
{
    #region Configuration
    [ConfigStoredInDB]
    class HKIPONDAConfig
    {
        [StoreInDB]
        [Category("EmailAccount")]
        [DisplayName("Name")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("File")]
        [DisplayName("Txt file path")]
        [Description("GeneratedFilePath")]
        public string TxtFilePath { get; set; }

        [StoreInDB]
        [Category("File")]
        [DisplayName("Ftp file path")]
        [Description("GetDownloadFilePath")]
        public string FtpFilePath { get; set; }

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

    class HKIPONDA : GeneratorBase
    {
        #region Description
        private static HKIPONDAConfig configObj = null;
        private string accountName = string.Empty;//UC169XXX
        private string password = string.Empty;//********
        private string domain = string.Empty;//TEN
        private string mailAdress = string.Empty;//eti.XXXXXX@thomsonreuters.com
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private List<string> attacheFileList = new List<string>();
        private ExchangeService service;
        private List<string> listHKIPO = null;//get data from local
        private string strTxtFilePath = string.Empty;
        private string strFileNameOnFtpToday = string.Empty;//Data From Ftp
        private string strFileNameOnFtpYesterday = string.Empty;//Data From Ftp
        private Dictionary<string, string> dicHKIPOFromFtp = new Dictionary<string, string>();//get ipo from .M on FTP
        private string strPatternFTP = string.Empty;
        private bool isExistYesterdayFileOnFtp = true;
        private bool isExistTodayFileOnFtp = true;
        private string strGetYesdayTxtFile = string.Empty;
        private string strGeneratedTodayTxtFile = string.Empty;
        private string strDownloadFilePath = string.Empty;
        private string strFileNameReadOnFtpYesterday = string.Empty;
        private string strFileNameReadOnFtpToday = string.Empty;

        private bool isExistYesterdayFileOnFtpHS = true;
        private bool isExistTodayFileOnFtpHS = true;
        private string strFileNameReadOnFtpYesterdayHS = string.Empty;
        private string strFileNameReadOnFtpTodayHS = string.Empty;

        protected override void Initialize()
        {
            configObj = Config as HKIPONDAConfig;
            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());
            accountName = emailAccount.AccountName;
            password = emailAccount.Password;
            domain = emailAccount.Domain;
            mailAdress = emailAccount.MailAddress;
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            strTxtFilePath = configObj.TxtFilePath;
            strDownloadFilePath = configObj.FtpFilePath.Trim();
            service = EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
            strGetYesdayTxtFile = Path.Combine(strTxtFilePath, DateTime.Now.ToUniversalTime().AddHours(+8).AddDays(-1).ToString("MMdd") + "All_HK_IPO_IDN.txt");
            strGeneratedTodayTxtFile = Path.Combine(strTxtFilePath, DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + "Missing_HK_IPO_NDA.txt");
            strFileNameOnFtpYesterday = "EM01" + DateTime.Now.ToUniversalTime().AddDays(-1).AddHours(+8).ToString("MMdd") + ".M";
            strFileNameOnFtpToday = "EM01" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M";
            strPatternFTP = @"^\S{2}(?<RIC>\S{4,10}\.(HK|HS))\b\s+";
            strFileNameReadOnFtpYesterday = "0001" + DateTime.Now.ToUniversalTime().AddDays(-1).AddHours(+8).ToString("MMdd") + ".M";
            strFileNameReadOnFtpToday = "0001" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M";

            strFileNameReadOnFtpYesterdayHS = "3200" + DateTime.Now.ToUniversalTime().AddDays(-1).AddHours(+8).ToString("MMdd") + ".M";
            strFileNameReadOnFtpTodayHS = "3200" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M";
        }
        #endregion

        protected override void Start()
        {
            GetDataToFutListFromLocalFile(strGetYesdayTxtFile);
            GetHKIPOFromFtp(strFileNameOnFtpYesterday, strPatternFTP, dicHKIPOFromFtp, ref isExistYesterdayFileOnFtp);
            GetHKIPOFromFtp(strFileNameOnFtpToday, strPatternFTP, dicHKIPOFromFtp, ref isExistTodayFileOnFtp);

            GetHKIPOFromFtpToDic(strFileNameReadOnFtpYesterday, strPatternFTP, dicHKIPOFromFtp);
            GetHKIPOFromFtpToDic(strFileNameReadOnFtpToday, strPatternFTP, dicHKIPOFromFtp);

            GetHKIPOFromFtpToDic(strFileNameReadOnFtpYesterdayHS, strPatternFTP, dicHKIPOFromFtp);
            GetHKIPOFromFtpToDic(strFileNameReadOnFtpTodayHS, strPatternFTP, dicHKIPOFromFtp);

            RemoveExistHKIPO(listHKIPO, dicHKIPOFromFtp);
            GenerateFile(listHKIPO, strGeneratedTodayTxtFile);//ipo removed
            SendEmail(listHKIPO);
        }

        private void GetHKIPOFromFtpToDic(string path, string pattern, Dictionary<string, string> dic)
        {
            try
            {
                string filePathFromFtp = string.Empty;
                filePathFromFtp = @"ftp://ASIA2:ASIA2@ds1.rds.reuters.com//" + path;
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
                    if (!(mc.Count > 0))
                        continue;
                    for (int i = 0; i < mc.Count; i++)
                    {
                        if (dic.ContainsKey(mc[i].Groups["RIC"].Value.Trim()))
                            continue;
                        dic.Add(mc[i].Groups["RIC"].Value.Trim(), "");
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("error when read file on ftp.{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        #region SendEmail(step:one)
        /// <summary>
        /// send email with diff case
        /// </summary>
        /// <param name="listHKIPO">ipo ric</param>
        private void SendEmail(List<string> listHKIPO)
        {
            string subject = string.Empty;
            string content = string.Empty;
            if (listHKIPO != null)
            {
                if (isExistYesterdayFileOnFtp || isExistTodayFileOnFtp)
                {
                    if (listHKIPO.Count > 0)//ok
                    {
                        subject = "HK IPO - NDA Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Txt and *.M File]";
                        content = "<center>*******IPO RIC Missing*********</center><br />";
                    }
                    else if (listHKIPO.Count == 0)//ok
                    {
                        subject = "HK IPO - NDA Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Found Txt and *.M File]";
                        content = "<center>*******No Missing IPO*********</center><br />";
                    }
                }
                else//ok
                {
                    subject = "HK IPO - NDA Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[No *.M File On Ftp]";
                    content = "<center>*******Failed To Extract DataScope File*********</center><br />";
                }
            }
            else//ok
            {
                subject = "HK IPO - NDA Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[No Txt File]";
                content = "<center>*******No IPO Today*********</center><br />";

            }
            SendMail(service, subject, content, attacheFileList);
        }
        #endregion

        #region GenerateTxtFileToLocal
        /// <summary>
        /// Generate txt file
        /// </summary>
        /// <param name="listHKIPO">ipo ric</param>
        /// <param name="txtFilePath">fileNamePath</param>
        private void GenerateFile(List<string> listHKIPO, string txtFilePath)
        {
            if ((isExistYesterdayFileOnFtp || isExistTodayFileOnFtp) && listHKIPO != null && listHKIPO.Count > 0)
            {
                string content = string.Empty;
                foreach (var str in listHKIPO)
                {
                    content += string.Format(",{0}", str);
                }
                content = content.Remove(0, 1);
                try
                {
                    File.WriteAllText(txtFilePath, content);
                    attacheFileList.Add(txtFilePath);
                    AddResult("HK IPO QC NDA File", txtFilePath, "nda");
                    //TaskResultList.Add(new TaskResultEntry("HKIPOQCNDAFile", "HKIPONDA", txtFilePath));
                }
                catch (Exception ex)
                {
                    Logger.Log(string.Format("Error happens when generating file. Ex: {0} .", ex.Message));
                }
            }
        }
        #endregion

        #region CleanRicFromListByDic
        /// <summary>
        /// Clean Exist Ric In dic
        /// </summary>
        /// <param name="listHKIPO">ric</param>
        /// <param name="dicHKIPOFromGATS">dic</param>
        private void RemoveExistHKIPO(List<string> listHKIPO, Dictionary<string, string> dicHKIPOFromGATS)
        {
            if ((isExistTodayFileOnFtp || isExistYesterdayFileOnFtp) && listHKIPO != null && listHKIPO.Count > 0)
            {
                int index = listHKIPO.Count;
                for (int i = index - 1; i >= 0; i--)
                {
                    if (dicHKIPOFromGATS.ContainsKey(listHKIPO[i]))
                    {
                        listHKIPO.Remove(listHKIPO[i]);
                    }
                }
            }
        }
        #endregion

        #region GetDataFromFtp
        /// <summary>
        /// GetDataFromFtp
        /// </summary>
        /// <param name="filePath">fileNameOnFtp</param>
        /// <param name="pattern">pattern</param>
        /// <param name="dic">input to dic</param>
        /// <param name="isExistFileOnFtp">bool</param>
        private void GetHKIPOFromFtp(string filePath, string pattern, Dictionary<string, string> dic, ref bool isExistFileOnFtp)
        {
            if (listHKIPO != null && listHKIPO.Count > 0)
            {
                try
                {
                    string filePathFromFtp = Path.Combine(strDownloadFilePath, filePath);
                    if (!File.Exists(filePathFromFtp))
                    {
                        isExistFileOnFtp = false;
                        return;
                    }
                    using (FileStream fs = new FileStream(filePathFromFtp, FileMode.Open))
                    {
                        using (StreamReader sr = new StreamReader(fs))
                        {
                            string tmp = null;
                            string strRic = string.Empty;
                            while ((tmp = sr.ReadLine()) != null)
                            {
                                Regex r = new Regex(pattern);
                                MatchCollection mc = r.Matches(tmp.Trim());
                                if (mc.Count > 0)
                                {
                                    strRic = mc[0].Groups["RIC"].Value;
                                    if (!dic.ContainsKey(strRic))
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
                    Logger.Log("error when read file on ftp!");
                }
            }
        }
        #endregion

        #region SendEmail(step:two)
        /// <summary>
        /// send email function
        /// </summary>
        /// <param name="service">login</param>
        /// <param name="subject">subject of email</param>
        /// <param name="content">content of email</param>
        /// <param name="attacheFileList">attachement of email</param>
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

        #region GetDataFromTxtFile
        /// <summary>
        /// GetDataFromTxt
        /// </summary>
        /// <param name="filePath">fileNamePath</param>
        private void GetDataToFutListFromLocalFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                FileStream fs = new FileStream(filePath, FileMode.Open);
                StreamReader sr = new StreamReader(fs);
                listHKIPO = new List<string>(sr.ReadToEnd().Split(','));
                sr.Close();
                fs.Close();
            }
        }
        #endregion
    }
}
