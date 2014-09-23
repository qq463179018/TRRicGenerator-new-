using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Drawing.Design;
using Microsoft.Exchange.WebServices.Data;
using Ric.Db.Info;
using Ric.Db.Manager;
using MSAD.Common.OfficeUtility;
using System.Threading;
using System.Text.RegularExpressions;
using System.IO;
using System.Net;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Taiwan
{
    #region Configuration
    [ConfigStoredInDB]
    class TWIpoQCConfig
    {
        [StoreInDB]
        [Category("EmailAccount")]
        [DefaultValue("Inbox\\")]
        [Description("Mail folder path,like: Inbox\\XXXXX")]
        public string MailFolderPath { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [DefaultValue("UC159450")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("FilePath")]
        [Description("GeneratedFilePath")]
        public string OutPutPath { get; set; }

        [StoreInDB]
        [Category("PublicDownloadPath")]
        [Description("Get EM01MMdd.M File")]
        public string PublicDownloadPath { get; set; }

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

    class TWIpoQC : GeneratorBase
    {
        #region Description
        private static TWIpoQCConfig configObj = null;

        private string accountName = string.Empty;//UC169XXX
        private string password = string.Empty;//********
        private string domain = string.Empty;//TEN
        private string mailAdress = string.Empty;//eti.XXXXXX@thomsonreuters.com
        private string mailFolder = string.Empty;//Inbox/XXXXX
        private string outPutFilePath = string.Empty;

        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private List<string> attacheFileList = new List<string>();

        private ExchangeService service;
        private DateTime startDate = DateTime.Now.ToUniversalTime().AddHours(+8).AddHours(-DateTime.Now.ToUniversalTime().AddHours(+8).Hour);
        private DateTime endDate = DateTime.Now.ToUniversalTime().AddHours(+8);

        private string emailPattern = string.Empty;
        private string emailKeyWord = string.Empty;
        private int emailCountInMailBox = 0;

        private List<string> listRic = null;
        private List<string> listRicMissing = null;
        private List<string> listRicMissingInFtp = null;
        private Dictionary<string, string> dicExistInGATS = new Dictionary<string, string>();
        private List<string> listFileCodeInFtp = new List<string>();

        private int fileCountInFtp = 0;
        private string fileCodeOfToday = DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd");
        private string ftpPattern = string.Empty;

        private List<string> listRicMissingInGATS = null;
        private Dictionary<string, string> dicExistInFtp = new Dictionary<string, string>();

        private string gatsPattern = string.Empty;
        private string strFileNameOnLocalToday = string.Empty;//Data From Ftp
        private string publicDownloadPath = string.Empty;

        protected override void Initialize()
        {
            configObj = Config as TWIpoQCConfig;

            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());

            accountName = emailAccount.AccountName;
            password = emailAccount.Password;
            domain = emailAccount.Domain;
            mailAdress = emailAccount.MailAddress;
            mailFolder = configObj.MailFolderPath.Trim();
            outPutFilePath = Path.Combine(configObj.OutPutPath.Trim(), "MissingRic.txt");
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;

            emailKeyWord = "*IPO* RIC Add";//there are two email need to search
            emailPattern = @"For\sTQSEffective\sDate\s+\:\s+\w+\s+\:\s+(?<RIC>\w+\.TWO{0,1})";
            //emailPattern = @"\s(?<RIC>\w+\.[A-Z]{2,3})Displayname";
            gatsPattern = @"(?<RIC>\w+\.[A-Z]{2,3})\s+PROD_PERM\s+\d+\s+";

            listFileCodeInFtp.Add("0126" + fileCodeOfToday + ".M");
            listFileCodeInFtp.Add("1127" + fileCodeOfToday + ".M");
            listFileCodeInFtp.Add("0645" + fileCodeOfToday + ".M");
            //listFileCodeInFtp.Add("01260319.M");
            //listFileCodeInFtp.Add("11270322.M");
            //listFileCodeInFtp.Add("06450319.M");

            ftpPattern = @"(?<RIC>\w+\.[A-Z]{2,3})\s+ENNOCONN ORD\s+";
            strFileNameOnLocalToday = "EM01" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M";
            publicDownloadPath = configObj.PublicDownloadPath;
        }
        #endregion

        protected override void Start()
        {
            listRic = GetRicFromEmail(emailKeyWord);
            listRicMissingInGATS = GetMissingRicFromGATS(listRic);
            listRicMissingInFtp = GetMissingRicFromFtp(listRic);
            listRicMissing = GenerateMissingRicFile(listRicMissingInGATS, listRicMissingInFtp);
            SendEmail();
        }

        private List<string> GetMissingRicFromFtp(List<string> listRic)
        {
            List<string> list = new List<string>();//Missing ric

            if (listRic == null)
                return null;

            if (listRic.Count == 0)
                return list;

            foreach (var str in listFileCodeInFtp)
            {
                GetDataToDicFromFtp(str, ftpPattern, dicExistInFtp);
            }

            GetDataToDicFromLocal(strFileNameOnLocalToday, ftpPattern, dicExistInFtp);//EM01MMdd.M is exist on local

            return FoundMissingRic(listRic, dicExistInFtp);
        }

        private void GetDataToDicFromLocal(string fileName, string pattern, Dictionary<string, string> dic)
        {
            try
            {
                string filePathFromLocal = Path.Combine(publicDownloadPath, fileName);

                if (!File.Exists(filePathFromLocal))
                {
                    string msg = string.Format("EM01MMdd.M file is not exist on local.{0}");
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                using (FileStream fs = new FileStream(filePathFromLocal, FileMode.Open))
                {
                    using (StreamReader sr = new StreamReader(fs))
                    {
                        string tmp = null;
                        string strRic = string.Empty;

                        while ((tmp = sr.ReadLine()) != null)
                        {
                            Regex r = new Regex(pattern);
                            MatchCollection mc = r.Matches(tmp.Trim());

                            if (!(mc.Count > 0))
                                continue;

                            strRic = mc[0].Groups["RIC"].Value.ToString().Trim();

                            if (dic.ContainsKey(strRic))
                                continue;

                            dic.Add(strRic, "");
                        }
                    }
                }

                fileCountInFtp++;
            }
            catch (Exception ex)
            {
                string msg = string.Format("error when read EM01MMdd.M file on local.{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GetDataToDicFromFtp(string path, string pattern, Dictionary<string, string> dic)
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

                fileCountInFtp++;
            }
            catch (Exception ex)
            {
                string msg = string.Format("error when read file on ftp.{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void SendEmail()
        {
            string subject = string.Empty;
            string content = string.Empty;

            if (emailCountInMailBox != 0)
            {
                if (listRic.Count == 0)//ok 
                {
                    subject = "TW IPO QC" + DateTime.Now.ToString("dd-MMM-yyyy") + "[OK]";
                    content = "<center>********No IPO Today********</center><br />";
                }
                else
                {
                    if (fileCountInFtp == 4)
                    {
                        if (listRicMissing.Count > 0)//ok
                        {
                            subject = "TW IPO QC" + DateTime.Now.ToString("dd-MMM-yyyy") + "[OK]";
                            content = "<center>********IPO RIC Missing********</center><br />";
                            SendMail(subject, content, attacheFileList);
                            return;
                        }
                        else if (listRicMissing.Count == 0)//ok
                        {
                            subject = "TW IPO QC" + DateTime.Now.ToString("dd-MMM-yyyy") + "[OK]";
                            content = "<center>********No Missing IPO********</center><br />";
                        }
                    }
                    else//ok
                    {
                        subject = "TW IPO QC" + DateTime.Now.ToString("dd-MMM-yyyy") + "[error]";
                        content = "<center>********found  " + fileCountInFtp + "  files in ftp (4 file is right)********</center>";
                    }
                }
            }
            else//ok
            {
                subject = "TW IPO QC" + DateTime.Now.ToString("dd-MMM-yyyy") + "[OK]";
                content = "<center>********No IPO Today********</center><br />";
                //subject = "TW IPO QC" + DateTime.Now.ToString("dd-MMM-yyyy") + "[error]";
                //content = "<center>********received  " + emailCountInMailBox + "  emails in mail box(2 email is right)********</center>";
            }

            SendMail(subject, content, new List<string>());
        }

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

            if (configObj.MailCC.Count > 1 || (configObj.MailCC.Count == 1 && configObj.MailCC[0] != ""))
            {
                listMailCC = configObj.MailCC;
            }

            EWSUtility.CreateAndSendMail(service, listMailTo, listMailCC, new List<string>(), subject, content, attacheFileList);
        }

        private List<string> GenerateMissingRicFile(List<string> listRicMissingInGATS, List<string> listRicMissingInFtp)
        {
            List<string> listRicMissing = new List<string>();

            if (listRicMissingInGATS == null || listRicMissingInFtp == null)
            {
                string msg = string.Format("the object of listRicMissingInGATS or listRicMissingInFtp is not exist!");
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }

            foreach (var str in listRicMissingInGATS)
            {
                //it must be mising ric wehen found in listRicMissingInGATS and listRicMissingInFtp
                if (listRicMissingInFtp.Contains(str))
                    listRicMissing.Add(str);
            }

            if (listRicMissing.Count == 0)
            {
                string msg = string.Format("no missing ric in listMissingRIC,no need generate txt missing ric file.");
                Logger.Log(msg, Logger.LogType.Warning);
                return listRicMissing;
            }

            GenerateFile(listRicMissing, outPutFilePath);
            attacheFileList.Add(outPutFilePath);

            return listRicMissing;
        }

        private void GenerateFile(List<string> list, string path)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("MissingRIC\r\n");

            foreach (var str in list)
            {
                sb.Append(str);
                sb.Append("\r\n");
            }

            try
            {
                File.WriteAllText(path, sb.ToString());
                TaskResultList.Add(new TaskResultEntry("TWipoQC", "MissingRIC", path));
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error happens when generating file. Ex: {0} .", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private List<string> GetMissingRicFromGATS(List<string> listRic)
        {
            List<string> list = new List<string>();//Missing ric

            if (listRic == null)
                return null;

            if (listRic.Count == 0)
                return list;

            string strQuery = string.Empty;
            int count = listRic.Count;
            int fenMu = 2000;
            int qiuYu = count % fenMu;
            int qiuShang = count / fenMu;

            if (qiuShang > 0)
            {
                for (int i = 0; i < qiuShang; i++)
                {
                    for (int j = 0; j < fenMu; j++)
                    {
                        string strTmp = listRic[i * fenMu + j].ToString().Trim();

                        if (!string.IsNullOrEmpty(strTmp))
                        {
                            strQuery += string.Format(",{0}", strTmp);
                        }
                    }

                    strQuery = strQuery.Remove(0, 1);
                    GetDataFromGATSToExistDic(strQuery, gatsPattern, dicExistInGATS);
                    strQuery = string.Empty;
                }
            }

            for (int i = qiuShang * fenMu; i < count; i++)
            {
                string strTmp = listRic[i].ToString().Trim();

                if (!string.IsNullOrEmpty(strTmp))
                {
                    strQuery += string.Format(",{0}", strTmp);
                }
            }

            strQuery = strQuery.Remove(0, 1);
            GetDataFromGATSToExistDic(strQuery, gatsPattern, dicExistInGATS);

            return FoundMissingRic(listRic, dicExistInGATS);
        }

        private List<string> FoundMissingRic(List<string> listRic, Dictionary<string, string> dicExistRIC)
        {
            List<string> list = new List<string>();

            if (listRic == null || listRic.Count == 0)
            {
                string msg = string.Format("no ric in listRic.");
                Logger.Log(msg, Logger.LogType.Warning);
                return null;
            }

            foreach (var str in listRic)
            {
                if (dicExistRIC.ContainsKey(str.Trim()))
                    continue;

                if (list.Contains(str.Trim()))
                    continue;

                list.Add(str.Trim());
            }

            return list;
        }

        private void GetDataFromGATSToExistDic(string strQuery, string pattern, Dictionary<string, string> dic)
        {
            try
            {
                GatsUtil gats = new GatsUtil();
                string response = gats.GetGatsResponse(strQuery, "PROD_PERM");
                Regex regex = new Regex(pattern);
                MatchCollection matches = regex.Matches(response);
                string tmp = string.Empty;

                foreach (Match match in matches)
                {
                    if (dic.ContainsKey(match.Groups["RIC"].Value.ToString().Trim()))
                        continue;

                    dic.Add(match.Groups["RIC"].Value.ToString().Trim(), "");
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("get exist ric in the gats error.:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

        }

        private List<string> GetRicFromEmail(string emailKeyWord)
        {
            List<string> list = new List<string>();
            try
            {
                service = EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
                EWSMailSearchQuery query = query = new EWSMailSearchQuery("", mailAdress, @mailFolder, emailKeyWord, "", startDate, endDate);
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

                emailCountInMailBox = mailList.Count;

                if (!(emailCountInMailBox > 0))
                {
                    string msg = string.Format("no email in the mailbox. ");
                    Logger.Log(msg, Logger.LogType.Warning);
                    return null;
                }

                string strEmailBody = string.Empty;

                for (int i = 0; i < emailCountInMailBox; i++)
                {
                    EmailMessage mail = mailList[i];
                    mail.Load();
                    strEmailBody = TWHelper.ClearHtmlTags(mail.Body.ToString());
                    Regex regex = new Regex(emailPattern);
                    MatchCollection matches = regex.Matches(strEmailBody);

                    foreach (Match match in matches)
                    {
                        if (list.Contains(match.Groups["RIC"].Value))
                            continue;

                        list.Add(match.Groups["RIC"].Value.ToString().Trim());
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("get ric from email error.:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }
        }
    }
}
