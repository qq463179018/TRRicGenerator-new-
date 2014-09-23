using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Ric.Db.Manager;
using Ric.Db.Info;
using System.IO;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using Ric.Core;


namespace Ric.Tasks
{
    class DeutscheBasicPackageFileDownloading : GeneratorBase
    {
        #region Fields
        private string mFileFolder = string.Empty;
        private DataStreamRicCreationWithFileDownloadConfig ConfigObj = null;
        private Dictionary<string, string> downloadUrl = new Dictionary<string, string>();
        private CookieContainer cookies = new CookieContainer();
        private string usrName = string.Empty;
        private string passWord = string.Empty;
        //email
        private string email_password = string.Empty;
        private string domain = string.Empty;
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private string accountName = string.Empty;
        private ExchangeService service;
        //
        private List<string> listFileNameError = new List<string>();//list fileName of error when downloading
        #endregion

        #region Initialize and Start

        private void InitializeDownloadUrlDirectory()
        {
            downloadUrl.Add("CDAX_IWB", "");
            downloadUrl.Add("ClassicAllShare_IWB", "");
            downloadUrl.Add("DAX_IWB", "");
            downloadUrl.Add("EntryStandard_IWB", "");
            downloadUrl.Add("GEX_IWB", "");
            downloadUrl.Add("HDAX_IWB", "");
            downloadUrl.Add("MDAX_IWB", "");
            downloadUrl.Add("MidcapMarketIndex_IWB", "");
            downloadUrl.Add("PrimeAllShare_IWB", "");
            downloadUrl.Add("SDAX_IWB", "");
            downloadUrl.Add("TecDAX_IWB", "");
            downloadUrl.Add("TechnologyAllShare_IWB", "");
        
        }
        protected override void Initialize()
        {
            ConfigObj = Config as DataStreamRicCreationWithFileDownloadConfig;
            TaskResultList.Add(new TaskResultEntry("LOG File", "LOG File", Logger.FilePath));

            usrName = ConfigObj.Username;
            passWord = ConfigObj.Password;

            usrName = usrName.Replace("@", "%40");
            passWord = passWord.Replace("@", "%40");

            accountName = ConfigObj.AccountName.Trim();
            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(ConfigObj.AccountName.Trim());
            if (emailAccount != null)
            {
                accountName = emailAccount.AccountName;
                email_password = emailAccount.Password;
                domain = emailAccount.Domain;
            }

            listMailTo = ConfigObj.MailTo;
            listMailSignature = ConfigObj.MailSignature;

            service = MSAD.Common.OfficeUtility.EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, email_password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));

            InitializeDownloadUrlDirectory();
            InitializeFileDirectory();

            string msg = "Initialize...OK!";
            Logger.Log(msg);
        }
        private void InitializeFileDirectory()
        {
           // mFileFolder = Path.Combine(ConfigObj.OutputPath, DateTime.Today.ToString("yyyy-MM-dd"));
            mFileFolder = ConfigObj.OutputPath;
            mFileFolder = mFileFolder + "\\";

            if (!Directory.Exists(mFileFolder))
            {
                Directory.CreateDirectory(mFileFolder);
            }

            TaskResultList.Add(new TaskResultEntry("FILES", "FILES", mFileFolder));
        }
        protected override void Start()
        {
            StartJob();
        }
        private void StartJob()
        {
            GetFilesUrl();
            List<string> result = DownLoadFiles();
            SendEmail(result, listFileNameError);
        }
        private void GetFilesUrl()
        {
            try
            {
                StreamReader SrResult = DataStreamRicCreationWithFileDownload.LoginWebSite(usrName, passWord, cookies);

                HtmlDocument doc = new HtmlDocument();
                doc.Load(SrResult);
                if (doc == null)
                {
                    return;
                }
                HtmlNode node = doc.DocumentNode.SelectSingleNode("//input[@id='ajaxDynaToken']");
                if (node == null)
                {
                    node = doc.DocumentNode.SelectSingleNode("//input[@name='org.apache.struts.taglib.html.TOKEN']");
                }
                string token = node.Attributes["value"].Value.Trim();


                node = doc.DocumentNode.SelectSingleNode("//form[@name='mainForm']");
                string actionStr = null;
                if (node != null)
                {
                    actionStr = node.Attributes["action"].Value.Trim();

                }
                string StResult = DataStreamRicCreationWithFileDownload.ChangeTab(actionStr, token, "1", cookies);

                Regex regex = new Regex(@"<!--delta.+-->");
                MatchCollection matches = regex.Matches(StResult);
                token = matches[0].Value;

                int i = token.IndexOf(";");
                token = token.Substring(i + 1, 32);

                StResult = DataStreamRicCreationWithFileDownload.ExpandDropDownBox(actionStr, token, @"37_1%3DExpand%3D-10", cookies);

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+CDAX_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["CDAX_IWB"] = matches[0].Value;

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+ClassicAllShare_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["ClassicAllShare_IWB"] = matches[0].Value;

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+DAX_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["DAX_IWB"] = matches[0].Value;

                //page 2
                regex = new Regex(@"<!--delta.+-->");
                matches = regex.Matches(StResult);
                token = matches[0].Value;

                //goto next page
                i = token.IndexOf(";");
                token = token.Substring(i + 1, 32);
                StResult = DataStreamRicCreationWithFileDownload.ChangePage(actionStr, token, @"37_1%3DPage%3D1", cookies);

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+EntryStandard_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["EntryStandard_IWB"] = matches[0].Value;

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+GEX_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["GEX_IWB"] = matches[0].Value;

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+HDAX_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["HDAX_IWB"] = matches[0].Value;

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+MDAX_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["MDAX_IWB"] = matches[0].Value;

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+MidcapMarketIndex_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["MidcapMarketIndex_IWB"] = matches[0].Value;

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+PrimeAllShare_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["PrimeAllShare_IWB"] = matches[0].Value;

                //page 3
                regex = new Regex(@"<!--delta.+-->");
                matches = regex.Matches(StResult);
                token = matches[0].Value;

                //goto next page
                i = token.IndexOf(";");
                token = token.Substring(i + 1, 32);
                StResult = DataStreamRicCreationWithFileDownload.ChangePage(actionStr, token, @"37_1%3DPage%3D2", cookies);

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+SDAX_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["SDAX_IWB"] = matches[0].Value;

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+TecDAX_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["TecDAX_IWB"] = matches[0].Value;

                regex = new Regex(@"download\/[0-9A-Z]+\.vong_00[a-z0-9_\/]+TechnologyAllShare_IWB\.[0-9]+\.xls");
                matches = regex.Matches(StResult);
                downloadUrl["TechnologyAllShare_IWB"] = matches[0].Value;
            }
            catch (System.Exception ex)
            {
                Logger.Log(string.Format("error when getting files url,exception is {0}", ex.ToString()));
            }
           
        }
        private List<string> DownLoadFiles()
        {
            string url = null;
            string fileName = null;
            List<string> result = new List<string>();
            try
            {
                foreach (string key in downloadUrl.Keys)
                {
                    Regex regex = new Regex(key + @"\.[0-9]+\.xls");
                    MatchCollection matches = regex.Matches(downloadUrl[key]);
                    fileName = mFileFolder + matches[0].Value;
                    url = @"https://contracts.deutsche-boerse.com/indexdata/" + downloadUrl[key];
                    DataStreamRicCreationWithFileDownload.DownLoadFiles(url, fileName, cookies);
                    result.Add(fileName);
                }
            }
            catch (System.Exception ex)
            {
                listFileNameError.Add(url);
                Logger.Log(string.Format("error when try to download file :{0} Exception:{1}", fileName, ex.ToString()));
                return result;
            }
            return result;

        }
        private void SendMail(ExchangeService service, string subject, string content, List<string> attacheFileList)
        {
            try
            {
                StringBuilder bodyBuilder = new StringBuilder();
                bodyBuilder.Append(content);
                bodyBuilder.Append("<p>");
                foreach (string signatureLine in ConfigObj.MailSignature)
                {
                    bodyBuilder.AppendFormat("{0}<br />", signatureLine);
                }
                bodyBuilder.Append("</p>");
                content = bodyBuilder.ToString();
                if (ConfigObj.MailCC.Count > 1 || (ConfigObj.MailCC.Count == 1 && ConfigObj.MailCC[0] != ""))
                {
                    listMailCC = ConfigObj.MailCC;
                }
                MSAD.Common.OfficeUtility.EWSUtility.CreateAndSendMail(service, listMailTo, listMailCC, new List<string>(), subject, content, attacheFileList);
            }
            catch (System.Exception ex)
            {
                Logger.Log(string.Format("error when Send result mail,exception is {0}", ex.ToString()));
            }
           
        }
        private void SendEmail(List<string> listFile, List<string> listFileNameError)
        {
            try
            {
                string subject = string.Empty;
                string content = string.Empty;
                if (listFile == null || listFile.Count == 0)
                {
                    //subject = "From Task ETI-278 Deutsche Borse Indices  Basic Package file downloading " + DateTime.Now.ToString("dd-MMM-yyyy") + "[error information]";
                    //content = "<center>*******Error happened when load files,Can't get file .*********</center><br />";
                    //SendMail(service, subject, content, new List<string>());//ok
                    Logger.Log("Error happened when load files,Can't get file");
                    LogMessage("Error happened when load files,Can't get file");
                    return;
                }
                else
                {
                    subject = "From Task ETI-278 Deutsche Borse Indices  Basic Package file downloading " + DateTime.Now.ToString("dd-MMM-yyyy") + "[task complete]";
                    content = "<center>*******download files list.*********</center><br />";
                    foreach (string str in listFile)
                    {
                        content += str + "<br />";
                    }
                    SendMail(service, subject, content,listFile);//ok
                }
                if (listFileNameError == null || listFileNameError.Count == 0)
                {
                    return;
                }
                //subject = "From Task ETI-278 Deutsche Borse Indices  Basic Package file downloading " + DateTime.Now.ToString("dd-MMM-yyyy") + "[error information]";
                //content = "<center>*******Error happened when download following files.*********</center><br />";
                //foreach (string str in listFileNameError)
                //{
                //    content += str + "<br />";
                //}
                //SendMail(service, subject, content, new List<string>());//ok
                Logger.Log("Error happened when download following files.");
                LogMessage("Error happened when download following files.");
            }
            catch (System.Exception ex)
            {
                Logger.Log(string.Format("error when Send error mail,exception is {0}", ex.ToString()));
            }
            
        }
        #endregion
    }
}
