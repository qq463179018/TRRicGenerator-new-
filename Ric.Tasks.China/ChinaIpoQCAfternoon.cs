using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Drawing.Design;
using Microsoft.Exchange.WebServices.Data;
using Ric.Db.Info;
using Ric.Db.Manager;
using System.IO;
using System.Text.RegularExpressions;
using System.Net;
using Ric.Core;
using System.Windows.Forms;
using Ric.Util;

namespace Ric.Tasks.China
{
    #region Configuration
    [ConfigStoredInDB]
    class ChinaIpoQCAfternoonConfig
    {
        [StoreInDB]
        [Category("EmailAccount")]
        [DefaultValue("UC159450")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("InputPath")]
        [Description("same as ChinaIpoQCMorning's OutputPath")]
        public string InputPath { get; set; }

        [StoreInDB]
        [Category("OutputPath")]
        [Description("Save Missing Ric File Path ")]
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

    class ChinaIpoQCAfternoon : GeneratorBase
    {
        #region Description

        private static ChinaIpoQCAfternoonConfig configObj = null;
        EmailAccountInfo emailAccount = null;
        private ExchangeService service;
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private List<string> listRic = null;//get ric from local file
        private List<string> listRicMissing = null;//save ric not exist in IND and NDA
        private List<string> attacheFileList = null;
        private string resultFilePath = string.Empty;
        private List<string> listRicIDN = null;//GATS
        private string patternIDN = string.Empty;
        private List<string> listRicNDA = null;//FTP
        private string patternNDA = string.Empty;
        private List<string> listCodeNDA = null;//eg:0001
        private int fileCountInFtp = 0;
        private string inputPath = string.Empty;
        private string patternTxtRic = string.Empty;

        protected override void Initialize()
        {
            configObj = Config as ChinaIpoQCAfternoonConfig;
            emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());

            if (emailAccount == null)
            {
                MessageBox.Show("email account is not exist in DB. ");
                return;
            }

            service = MSAD.Common.OfficeUtility.EWSUtility.CreateService(new System.Net.NetworkCredential(emailAccount.AccountName, emailAccount.Password, emailAccount.Domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            inputPath = Path.Combine(configObj.InputPath.Trim(), string.Format(@"output\RIC_{0}.txt", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyy-MM-dd")));
            patternIDN = @"\r\n(?<RIC>\w{5,15}[\.|\=]\w{1,10})\s+PROD_PERM\s+";
            patternTxtRic = @"(?<RIC>\w{5,15}[\.|\=]\w{1,10})";
            patternNDA = @"^\b\S+(?<RIC>\w{5,15}[\.|\=]\w{1,10})";
            listCodeNDA = new List<string>();
            listCodeNDA.Add("0163" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M");
            listCodeNDA.Add("0179" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M");
            listCodeNDA.Add("3201" + DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + ".M");
            //listCodeNDA.Add("0163" + "0408" + ".M");
            //listCodeNDA.Add("0179" + "0408" + ".M");
            resultFilePath = Path.Combine(configObj.OutputPath, "output");
        }
        #endregion

        protected override void Start()
        {
            #region [Get Ric List]
            try
            {
                listRic = GetRic(inputPath);
            }
            catch (Exception ex)
            {
                string msg = string.Format("get ric from loacl txt file generated by morning task error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            #region [Get Ric In IDN]
            try
            {
                listRicIDN = GetRicFromIDN(listRic);//gats 
            }
            catch (Exception ex)
            {
                string msg = string.Format("query exist ric in gats error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            #region [Get Ric In NDA]
            try
            {
                listRicNDA = GetRicFromNDA(listRic);//ftp 
            }
            catch (Exception ex)
            {
                string msg = string.Format("get exist ric from ftp error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            #region [Get Missing Ric]
            try
            {
                listRicMissing = GetMissingRic(listRicIDN, listRicNDA);
            }
            catch (Exception ex)
            {
                string msg = string.Format("traversal list ric and find ric not in idn and nda error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            #region [Generate Result File]
            try
            {
                attacheFileList = GenerateFile(listRicMissing);
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate missing ric txt file to local error. msg:{0}", ex.ToString());
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
                string msg = string.Format("send result email to user error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion
        }

        #region [Send Email]
        private void SendEmail()
        {
            string subject = string.Empty;
            string content = string.Empty;
            if (listRic == null || listRic.Count == 0)//no txt file in morning //ok
            {
                subject = "China IPO QC Afternoon - Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy");
                content = "<center>********No Missing IPO Ric Today .********</center><br /><br />";
                SendEmail(service, subject, content, new List<string>());
                return;
            }
            if (listRicIDN != null && fileCountInFtp == 2)//no error//ok
            {
                if (attacheFileList != null && attacheFileList.Count > 0)
                {
                    subject = "China IPO QC Afternoon - Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy");
                    content = "<center>********Missing IPO Ric Today .********</center><br /><br />";
                    SendEmail(service, subject, content, attacheFileList);
                    return;
                }
                else//all ric find in the idn and nda//ok
                {
                    subject = "China IPO QC Afternoon - Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy");
                    content = "<center>********No Missing IPO Ric Today .********</center><br /><br />";
                    SendEmail(service, subject, content, new List<string>());
                    return;
                }
            }
            else//error in gats or ftp
            {
                if (listRicIDN == null) //gats error
                {
                    subject = "China IPO QC Afternoon - Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy");
                    content = "<center>********Unknown Exception: Get Ric From GATS(IDN) Error.********</center><br /><br />";
                    SendEmail(service, subject, content, new List<string>());
                    return;
                }
                if (fileCountInFtp < 2) //ftp error//ok
                {
                    subject = "China IPO QC Afternoon - Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy");
                    content = "<center>********Unknown Exception: The File On Ftp Count = " + fileCountInFtp + " [2 file is right!].********</center><br /><br />";
                    SendEmail(service, subject, content, new List<string>());
                    return;
                }

            }
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

            string path = Path.Combine(resultFilePath, string.Format("RIC_Missing_{0}.txt", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyy-MM-dd")));
            StringBuilder sb = new StringBuilder();
            sb.Append("RIC_Missing\t\r\n");

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

        #region [Get Missing Ric]
        private List<string> GetMissingRic(List<string> listRicIDN, List<string> listRicNDA)
        {
            List<string> list = new List<string>();
            List<string> listFirst = new List<string>();

            if (listRicIDN == null || listRicIDN.Count == 0)
            {
                string msg = string.Format("no ric in the local or get ric from GATS error .");
                Logger.Log(msg, Logger.LogType.Warning);
            }

            if (listRicNDA == null || listRicNDA.Count == 0)
            {
                string msg = string.Format("no ric in the local or get ric from FTP error .");
                Logger.Log(msg, Logger.LogType.Warning);
            }

            foreach (var item in listRic)
            {
                if (!listRicIDN.Contains(item))
                    listFirst.Add(item);
            }

            foreach (var item in listFirst)
            {
                if (!listRicNDA.Contains(item))
                    list.Add(item);
            }

            return list;
        }
        #endregion

        #region [Get Ric In NDA]
        private List<string> GetRicFromNDA(List<string> listRic)
        {
            List<string> list = new List<string>();//Missing ric

            if (listRic == null)
                return null;

            if (listRic.Count == 0)
                return list;

            foreach (var str in listCodeNDA)
            {
                GetRicFromFtp(str, patternNDA, list);
            }
            return list;
        }

        private void GetRicFromFtp(string path, string pattern, List<string> list)
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
                        if (list.Contains(mc[i].Groups["RIC"].Value.Trim()))
                            continue;

                        list.Add(mc[i].Groups["RIC"].Value.Trim());
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
        #endregion

        #region [Get Ric In IDN]
        private List<string> GetRicFromIDN(List<string> listRic)
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
                    GetRicFromGATS(strQuery, patternIDN, list);
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
            GetRicFromGATS(strQuery, patternIDN, list);
            return list;
        }

        private void GetRicFromGATS(string strQuery, string pattern, List<string> list)
        {
            try
            {
                GatsUtil gats = new GatsUtil();
                string response = gats.GetGatsResponse(strQuery, "PROD_PERM");//*******************
                Regex regex = new Regex(pattern);
                MatchCollection matches = regex.Matches(response);
                string tmp = string.Empty;

                foreach (Match match in matches)
                {
                    if (list.Contains(match.Groups["RIC"].Value.ToString().Trim()))
                        continue;

                    list.Add(match.Groups["RIC"].Value.ToString().Trim());
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("get exist ric in the gats error.:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }
        #endregion

        #region [Get Ric List]
        private List<string> GetRic(string path)
        {
            try
            {
                List<string> list = new List<string>();

                if (!File.Exists(path))
                {
                    string msg = string.Format("The file: {0} is not exist in the local.", inputPath);
                    Logger.Log(msg, Logger.LogType.Info);
                    return list;//send email with no ipo ric missing today
                }

                using (FileStream fs = new FileStream(path, FileMode.Open))
                {
                    using (StreamReader sr = new StreamReader(fs))
                    {
                        //list = new List<string>(sr.ReadToEnd().Replace("\t\r\n", ",").Split(','));
                        string tmp = null;
                        while ((tmp = sr.ReadLine()) != null)
                        {
                            Regex r = new Regex(patternTxtRic);
                            MatchCollection mc = r.Matches(tmp.Trim());

                            if (!(mc.Count > 0))
                                continue;

                            for (int i = 0; i < mc.Count; i++)
                            {
                                if (list.Contains(mc[i].Groups["RIC"].Value.Trim()))
                                    continue;

                                list.Add(mc[i].Groups["RIC"].Value.Trim());
                            }
                        }
                        return list;//to next step
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("read txt file from local error . msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return null;//send email with unknown exception 
            }
        }
        #endregion
    }
}
