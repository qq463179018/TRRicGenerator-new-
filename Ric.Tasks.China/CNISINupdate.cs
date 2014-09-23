using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Microsoft.Exchange.WebServices.Data;
using MSAD.Common.OfficeUtility;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks.China
{
    #region Configuration

    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class CNISINupdateConfig
    {
        [StoreInDB]
        [Category("File")]
        [DisplayName("File path")]
        [Description("Generated Files Path,like: G\\China")]
        public string FilePath { get; set; }

        [StoreInDB]
        [Category("Account")]
        [DisplayName("Account Name")]
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

    #region Declaration
    class WebsitTableEntity
    {
        public string RIC { get; set; }
        public string OFFC_CODE2 { get; set; }
    }

    class CNISINupdate : GeneratorBase
    {
        List<string> listErrorRic = new List<string>();
        List<List<string>> listListSS = new List<List<string>>();
        List<List<string>> listListSSMissing = new List<List<string>>();
        private string strSSOrdinaryshares = "0#A1.SS,0#A2.SS,0#A3.SS,0#A4.SS,0#A5.SS,0#A6.SS,0#A7.SS,0#A8.SS,0#A9.SS,0#B.SS";
        List<string> listStrSSOrdinaryshares = new List<string>();
        private string strSSFunds = "0#FUND.SS,0#ETF.SS";
        List<string> listStrSSFunds = new List<string>();
        private string strSSBonds = "0#CNTSY=SS,0#CNCOR=SS,0#CNMUN=SS,0#CNCOV=SS";
        List<string> listStrSSBonds = new List<string>();
        List<List<string>> listListSZ = new List<List<string>>();
        List<List<string>> listListMissing = new List<List<string>>();
        private string strSHOrdinaryshares = "0#A1.SZ,0#A2.SZ,0#A3.SZ,0#A4.SZ,0#A5.SZ,0#A6.SZ,0#A7.SZ,0#A8.SZ,0#SME.SZ,0#CHINEXT.SZ,0#B.SZ";
        List<string> listStrSHOrdinaryshares = new List<string>();
        private string strSHFunds = "0#FUND.SZ,0#LOF.SZ,0#ETF.SZ";
        List<string> listStrSHFunds = new List<string>();
        private string strSHBonds = "0#CNTSY=SZ,0#CNCOR=SZ,0#CNCOV=SZ,0#CNMUN=SZ";
        List<string> listStrSHBonds = new List<string>();
        private string strPatternRIC = @"\bLONGLINK\d{1,2}\b\s+\b(?<RIC>C{0,1}N{0,1}\d{6}[\.|\=]S[S|Z])\r\n\S*\b";
        private string strPatternValue = @"\bOFFC_CODE2\b\s+\b(?<Value>\S*)\r\n(?<RIC>C{0,1}N{0,1}\d{6}[\.|\=]S[S|Z])\b";
        private string filePath = string.Empty;
        private string taskResultGEDAName = "GEDA ISIN Update";
        private string strSSfileName = "SS_ISIN.txt";//SS
        private string strSZfileName = "SZ_ISIN.txt";
        private string taskResultMissing = "Missing ISIN List";
        private string strMissingSSfileName = "SS_ISIN_Missing.txt";//SS
        private string strMissingSZfileName = "SZ_ISIN_Missing.txt";
        private string taskResultCSVName = "NDA ISIN Update";
        private string strSSCSVfileName = "SS_ISIN.csv";//SS
        private string strSZCSVfileName = "SZ_ISIN.csv";
        Dictionary<string, string> dicExist = new Dictionary<string, string>();
        private static CNISINupdateConfig configObj = null;
        private string querySecuritiesCode = string.Empty;
        private string securitiesNameCI = string.Empty;
        private string securitiesNameEl = string.Empty;
        List<WebsitTableEntity> listStrSSOrdinarysharesEntity = new List<WebsitTableEntity>();//*.SS
        List<WebsitTableEntity> listStrSSFundsEntity = new List<WebsitTableEntity>();//*.SS
        List<WebsitTableEntity> listStrSSBondsEntity = new List<WebsitTableEntity>();//CN*=SS
        List<WebsitTableEntity> listStrSZOrdinarysharesEntity = new List<WebsitTableEntity>();//*.SZ
        List<WebsitTableEntity> listStrSZFundsEntity = new List<WebsitTableEntity>();//*.SZ
        List<WebsitTableEntity> listStrSZBondsEntity = new List<WebsitTableEntity>();//CN*=SZ
        Dictionary<string, string> dicSSGuoZhai = new Dictionary<string, string>();
        Dictionary<string, string> dicSSZhaiJuan = new Dictionary<string, string>();
        Dictionary<string, string> dicSZGuoZhai = new Dictionary<string, string>();
        Dictionary<string, string> dicSZZhaiJuan = new Dictionary<string, string>();
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
        private string subject = string.Empty;
        private string content = string.Empty;

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as CNISINupdateConfig;
            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());
            accountName = emailAccount.AccountName;
            password = emailAccount.Password;
            domain = emailAccount.Domain;
            mailAdress = emailAccount.MailAddress;
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            filePath = configObj.FilePath;
            dicSSGuoZhai.Add("010", "国债");
            dicSSGuoZhai.Add("019", "国债");
            dicSSGuoZhai.Add("020", "国债");
            dicSSZhaiJuan.Add("12", "债券");
            dicSSZhaiJuan.Add("110", "债券");
            dicSSZhaiJuan.Add("113", "债券");
            dicSZGuoZhai.Add("110", "国债");
            dicSZGuoZhai.Add("101", "国债");
            dicSZGuoZhai.Add("108", "国债");
            dicSZZhaiJuan.Add("11", "债券");
            dicSZZhaiJuan.Add("12", "债券");
        }
    #endregion

        protected override void Start()
        {
            LogMessage("Getting data from GATS");
            GetDataFromGATSTolistList(listListSS, listListSZ);
            LogMessage("Loading data from GATS (SS)");
            GetExistDicFromGATS(listListSS);
            LogMessage("Loading data from GATS (SZ)");
            GetExistDicFromGATS(listListSZ);
            LogMessage("Remove existing RICs (SS)");
            RemoveExistRIC(listListSS, dicExist);
            LogMessage("Remove existing RICs (SZ)");
            RemoveExistRIC(listListSZ, dicExist);
            LogMessage("Compare data with GATS");
            IsAllExistInGATS(listListSS, listListSZ);
            LogMessage("Sending Email");
            SendEmail();
        }

        #region SendEmail

        private void SendEmail()
        {
            if (listErrorRic.Count == 0)
            {
                subject = "CN ISIN update" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Files]";
                content = "<center>********CN ISIN update-Generated file********</center><br /><br />";
                SendMail(subject, content, attacheFileList);
            }
            else
            {
                subject = "CN ISIN update" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Exist Missing Ric]";
                content = "<center>********There are some Ric need you query on: http://www.csisc.cn/isin-webapp/isinbiz/SearchQuery.do?m=userEnter ********</center><br /><br />Missing RIC<br />";
                foreach (string ric in listErrorRic)
                {
                    content += ric + "<br />";
                }
                SendMail(subject, content, attacheFileList);
            }
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
            service = EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
            if (configObj.MailCC.Count > 1 || (configObj.MailCC.Count == 1 && configObj.MailCC[0] != ""))
            {
                listMailCC = configObj.MailCC;
            }
            EWSUtility.CreateAndSendMail(service, listMailTo, listMailCC, new List<string>(), subject, content, attacheFileList);
        }
        #endregion

        private void IsAllExistInGATS(List<List<string>> listListSS, List<List<string>> listListSZ)
        {
            if ((listListSS[0].Count + listListSS[1].Count + listListSS[2].Count) > 0)
            {
                GetDataFromWebsiteToListEntity(listListSS, listStrSSOrdinarysharesEntity, listStrSSFundsEntity, listStrSSBondsEntity, dicSSGuoZhai, dicSSZhaiJuan);
                GenerateTxtFile(listStrSSOrdinarysharesEntity, listStrSSFundsEntity, listStrSSBondsEntity, strSSfileName, taskResultGEDAName);//.txt
                GenerateTxtFile(listStrSSOrdinarysharesEntity, listStrSSFundsEntity, listStrSSBondsEntity, strMissingSSfileName, taskResultMissing);//Missing.txt
                GenerateCsvFile(listStrSSOrdinarysharesEntity, listStrSSFundsEntity, listStrSSBondsEntity, strSSCSVfileName, taskResultCSVName);//.csv
            }
            if ((listListSZ[0].Count + listListSZ[1].Count + listListSZ[2].Count) > 0)
            {
                GetDataFromWebsiteToListEntity(listListSZ, listStrSZOrdinarysharesEntity, listStrSZFundsEntity, listStrSZBondsEntity, dicSZGuoZhai, dicSZZhaiJuan);
                GenerateTxtFile(listStrSZOrdinarysharesEntity, listStrSZFundsEntity, listStrSZBondsEntity, strSZfileName, taskResultGEDAName);//.txt
                GenerateTxtFile(listStrSZOrdinarysharesEntity, listStrSZFundsEntity, listStrSZBondsEntity, strMissingSZfileName, taskResultMissing);//Missing.txt
                GenerateCsvFile(listStrSZOrdinarysharesEntity, listStrSZFundsEntity, listStrSZBondsEntity, strSZCSVfileName, taskResultCSVName);//.csv
            }
        }

        private void GenerateTxtFile(List<WebsitTableEntity> listStrOrdinarysharesEntity, List<WebsitTableEntity> listStrFundsEntity, List<WebsitTableEntity> listStrBondsEntity, string fileName, string taskResultName)
        {
            string cNISINupdateTxtFilePath = Path.Combine(filePath, fileName);
            string content = "RIC\t#INSTMOD_OFFC_CODE2\r\n";
            string strEnd = string.Empty;
            string strBondStart = "CN";
            string strBondEnd = string.Empty;
            if (fileName.Substring(0, 2).Equals("SS"))
            {
                strEnd = ".SS";
                strBondEnd = "=SS";
            }
            else
            {
                strEnd = ".SZ";
                strBondEnd = "=SZ";
            }
            foreach (WebsitTableEntity webTable in listStrOrdinarysharesEntity.Where(webTable => webTable.RIC.Length == 6))
            {
                if (fileName.Length == 11)//.txt
                {
                    if (!string.IsNullOrEmpty(webTable.OFFC_CODE2.Trim()))
                    {
                        content += string.Format("{0}\t", webTable.RIC + strEnd);
                        content += string.Format("{0}\t", webTable.OFFC_CODE2);
                        content += "\r\n";
                    }
                }
                else  //Missing.txt
                {
                    if (string.IsNullOrEmpty(webTable.OFFC_CODE2.Trim()))
                    {
                        content += string.Format("{0}\t", webTable.RIC + strEnd);
                        content += string.Format("{0}\t", webTable.OFFC_CODE2);
                        content += "\r\n";
                    }
                }
            }
            foreach (WebsitTableEntity webTable in listStrFundsEntity.Where(webTable => webTable.RIC.Length == 6))
            {
                if (fileName.Length == 11)//.txt
                {
                    if (!string.IsNullOrEmpty(webTable.OFFC_CODE2.Trim()))
                    {
                        content += string.Format("{0}\t", webTable.RIC + strEnd);
                        content += string.Format("{0}\t", webTable.OFFC_CODE2);
                        content += "\r\n";
                    }
                }
                else     //Missing.txt
                {
                    if (string.IsNullOrEmpty(webTable.OFFC_CODE2.Trim()))
                    {
                        content += string.Format("{0}\t", webTable.RIC + strEnd);
                        content += string.Format("{0}\t", webTable.OFFC_CODE2);
                        content += "\r\n";
                    }
                }
            }
            foreach (WebsitTableEntity webTable in listStrBondsEntity.Where(webTable => webTable.RIC.Length == 6))
            {
                if (fileName.Length == 11)//.txt
                {
                    if (!string.IsNullOrEmpty(webTable.OFFC_CODE2.Trim()))
                    {
                        content += string.Format("{0}\t", strBondStart + webTable.RIC + strBondEnd);
                        content += string.Format("{0}\t", webTable.OFFC_CODE2);
                        content += "\r\n";
                    }
                }
                else      //Missing.txt
                {
                    if (string.IsNullOrEmpty(webTable.OFFC_CODE2.Trim()))
                    {
                        content += string.Format("{0}\t", strBondStart + webTable.RIC + strBondEnd);
                        content += string.Format("{0}\t", webTable.OFFC_CODE2);
                        content += "\r\n";
                    }
                }
            }
            try
            {
                File.WriteAllText(cNISINupdateTxtFilePath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LogMessage(string.Format("Error happens when generating txt file. Ex: {0} .", ex.Message), Logger.LogType.Error);
            }
            attacheFileList.Add(cNISINupdateTxtFilePath);
            AddResult(taskResultName, cNISINupdateTxtFilePath, "file");
        }

        private void GenerateCsvFile(List<WebsitTableEntity> listStrOrdinarysharesEntity, List<WebsitTableEntity> listStrFundsEntity, List<WebsitTableEntity> listStrBondsEntity, string fileName, string taskResultName)
        {
            string cNISINupdateTxtFilePath = Path.Combine(filePath, fileName);
            string content = "RIC,OFFC_CODE2\n";
            string strEnd = string.Empty;
            string strBondStart = "CN";
            string strBondEnd = string.Empty;
            if (fileName.Substring(0, 2).Equals("SS"))
            {
                strEnd = ".SS";
                strBondEnd = "=SS";
            }
            else
            {
                strEnd = ".SZ";
                strBondEnd = "=SZ";
            }
            foreach (WebsitTableEntity webTable in listStrOrdinarysharesEntity.Where(webTable => webTable.RIC.Length == 6)
                                                                              .Where(webTable => !string.IsNullOrEmpty(webTable.OFFC_CODE2.Trim())))
            {
                content += string.Format("{0}\t", webTable.RIC + strEnd + ",");
                content += string.Format("{0}\t", webTable.OFFC_CODE2);
                content += "\r\n";
                if (strEnd.Equals(".SS") && webTable.RIC.StartsWith("6"))
                {
                    content += string.Format("{0}\t", webTable.RIC + ".SH" + ",");
                    content += string.Format("{0}\t", webTable.OFFC_CODE2);
                    content += "\r\n";
                }
            }
            foreach (WebsitTableEntity webTable in listStrFundsEntity.Where(webTable => webTable.RIC.Length == 6)
                                                                     .Where(webTable => !string.IsNullOrEmpty(webTable.OFFC_CODE2.Trim())))
            {
                content += string.Format("{0}\t", webTable.RIC + strEnd + ",");
                content += string.Format("{0}\t", webTable.OFFC_CODE2);
                content += "\r\n";
                if (strEnd.Equals(".SS") && webTable.RIC.StartsWith("6"))
                {
                    content += string.Format("{0}\t", webTable.RIC + ".SH" + ",");
                    content += string.Format("{0}\t", webTable.OFFC_CODE2);
                    content += "\r\n";
                }
            }
            foreach (WebsitTableEntity webTable in listStrBondsEntity.Where(webTable => webTable.RIC.Length == 6)
                                                                     .Where(webTable => !string.IsNullOrEmpty(webTable.OFFC_CODE2.Trim())))
            {
                content += string.Format("{0}\t", strBondStart + webTable.RIC + strBondEnd + ",");
                content += string.Format("{0}\t", webTable.OFFC_CODE2);
                content += "\r\n";
            }
            try
            {
                File.WriteAllText(cNISINupdateTxtFilePath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LogMessage(string.Format("Error happens when generating txt file. Ex: {0} .", ex.Message), Logger.LogType.Error);
            }
            attacheFileList.Add(cNISINupdateTxtFilePath);
            AddResult(taskResultName, cNISINupdateTxtFilePath, "file");
        }

        private void GetDataFromWebsiteToListEntity(List<List<string>> listList, List<WebsitTableEntity> listStrOrdinarysharesEntity, List<WebsitTableEntity> listStrFundsEntity, List<WebsitTableEntity> listStrBondsEntity, Dictionary<string, string> dicGuoZhai, Dictionary<string, string> dicZhaiJuan)
        {
            for (int i = 0; i < listList.Count; i++)
            {
                switch (i.ToString())
                {
                    case "0":
                        foreach (string str in listList[i])
                        {
                            querySecuritiesCode = str.Substring(0, 6);
                            securitiesNameCI = "%E6%99%AE%E9%80%9A%E8%82%A1";//普通股
                            securitiesNameEl = "";
                            GetDataFromWebsiteByStockCode(listStrOrdinarysharesEntity, querySecuritiesCode, securitiesNameCI, securitiesNameEl);
                        }
                        break;
                    case "1":
                        foreach (string str in listList[i])
                        {
                            querySecuritiesCode = str.Substring(0, 6);
                            securitiesNameCI = "";
                            securitiesNameEl = "";
                            GetDataFromWebsiteByStockCode(listStrFundsEntity, querySecuritiesCode, securitiesNameCI, securitiesNameEl);
                        }
                        break;
                    case "2":
                        foreach (string str in listList[i])
                        {
                            querySecuritiesCode = str.Substring(2, 6);
                            securitiesNameCI = "";
                            if (dicGuoZhai.ContainsKey(querySecuritiesCode.Substring(0, 3)))
                            {
                                securitiesNameCI = "%E5%9B%BD%E5%80%BA";//国债
                            }
                            else if (dicZhaiJuan.ContainsKey(querySecuritiesCode.Substring(0, 2)) || dicZhaiJuan.ContainsKey(querySecuritiesCode.Substring(0, 3)))
                            {
                                securitiesNameCI = "%E5%80%BA%E5%88%B8";//债券
                            }
                            securitiesNameEl = "BOND";
                            GetDataFromWebsiteByStockCode(listStrBondsEntity, querySecuritiesCode, securitiesNameCI, securitiesNameEl);
                        }
                        break;
                }
            }
        }

        private void RemoveExistRIC(List<List<string>> listList, Dictionary<string, string> dicExist)
        {
            foreach (List<string> list in listList)
            {
                for (int i = 0; i < list.Count; i++)
                {
                    if (dicExist.ContainsKey(list[i]))
                    {
                        list.Remove(list[i]);
                        i--;
                    }
                }
            }
        }

        private void GetExistDicFromGATS(List<List<string>> listList)
        {
            foreach (List<string> list in listList)
            {
                string strQuery = list.Aggregate(string.Empty, (current, str) => current + (str + ","));
                strQuery = strQuery.Remove(strQuery.Length - 1, 1);
                GetDataFromGATSToExistDic(strQuery, dicExist);
            }
        }

        private void GetDataFromGATSTolistList(List<List<string>> listListSS, List<List<string>> listListSZ)
        {
            AutoInputStockCodeChainToGATS(strSSOrdinaryshares, listStrSSOrdinaryshares);
            listListSS.Add(listStrSSOrdinaryshares);
            AutoInputStockCodeChainToGATS(strSSFunds, listStrSSFunds);
            listListSS.Add(listStrSSFunds);
            AutoInputStockCodeChainToGATS(strSSBonds, listStrSSBonds);
            listListSS.Add(listStrSSBonds);
            AutoInputStockCodeChainToGATS(strSHOrdinaryshares, listStrSHOrdinaryshares);
            listListSZ.Add(listStrSHOrdinaryshares);
            AutoInputStockCodeChainToGATS(strSHFunds, listStrSHFunds);
            listListSZ.Add(listStrSHFunds);
            AutoInputStockCodeChainToGATS(strSHBonds, listStrSHBonds);
            listListSZ.Add(listStrSHBonds);
        }

        private void AutoInputStockCodeChainToGATS(string parmGATS, List<string> listStockCode)
        {
            try
            {
                int low = 0;
                int listRunAfter = listStockCode.Count;
                int listRunBefor = -1;
                while (listRunAfter != listRunBefor)
                {
                    listRunBefor = listStockCode.Count;
                    GetDataFromGATSToList(parmGATS, listStockCode);
                    parmGATS = parmGATS.Replace(low + "#", (low + 1) + "#");
                    low++;
                    listRunAfter = listStockCode.Count;
                }
            }
            catch (Exception e)
            {
                LogMessage("error has generated : " + e.Message, Logger.LogType.Error);
            }
        }

        private void GetDataFromWebsiteByStockCode(List<WebsitTableEntity> listEntity, string querySecuritiesCode, string securitiesNameCI, string securitiesNameEl)
        {
            string uri = string.Format(@"http://www.csisc.cn/isin-webapp/isinbiz/SearchQuery.do");
            string postDataStr = string.Format("m=userEnter&applyId=&currentPage=1&querySecuritiesCode={0}&queryIsin=&securitiesNameCs=&securitiesNameCl={1}&securitiesNameEl={2}&applyForm=%E6%9F%A5+%E8%AF%A2&jumpToPage=", querySecuritiesCode, securitiesNameCI, securitiesNameEl);
            WebsitTableEntity websiteTableEntity = null;
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
                request.CookieContainer = new CookieContainer();
                CookieContainer cookie = request.CookieContainer;
                request.Referer = @"http://www.csisc.cn/isin-webapp/isinbiz/SearchQuery.do?m=userEnter";
                request.Accept = "Accept:text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                request.Headers["Accept-Language"] = "zh-CN,zh;q=0.";
                request.Headers["Accept-Charset"] = "GBK,utf-8;q=0.7,*;q=0.3";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1500.95 Safari/537.36";
                request.KeepAlive = false;
                request.Timeout = 5000;
                request.ContentType = "application/x-www-form-urlencoded";
                request.Method = "POST";
                Encoding encoding = Encoding.UTF8;
                byte[] postData = encoding.GetBytes(postDataStr);
                request.ContentLength = postData.Length;
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(postData, 0, postData.Length);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream responseStream = response.GetResponseStream();
                if (response.Headers["Content-Encoding"] != null && response.Headers["Content-Encoding"].ToLower().Contains("gzip"))
                {
                    responseStream = new GZipStream(responseStream, CompressionMode.Decompress);
                }
                StreamReader streamReader = new StreamReader(responseStream, encoding);
                HtmlDocument htmlDoc = new HtmlDocument();
                htmlDoc.Load(streamReader);
                HtmlNodeCollection tables = htmlDoc.DocumentNode.SelectNodes(".//table");
                HtmlNode table = tables[10];
                HtmlNodeCollection trs = table.SelectNodes(".//tr");
                if (trs.Count == 1)
                {
                    websiteTableEntity = new WebsitTableEntity();
                    websiteTableEntity.OFFC_CODE2 = "            ";
                    websiteTableEntity.RIC = querySecuritiesCode;
                    listEntity.Add(websiteTableEntity);
                }
                else if (trs.Count >= 2)
                {
                    for (int i = 1; i < trs.Count; i++)
                    {
                        string ric = trs[i].SelectNodes(".//td")[2].InnerText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();
                        if (ric.Length == 6)
                        {
                            websiteTableEntity = new WebsitTableEntity
                            {
                                OFFC_CODE2 =
                                    trs[i].SelectNodes(".//td")[1].InnerText.Replace("&nbsp;", "")
                                        .Replace("\n", "")
                                        .Replace("\r", "")
                                        .Replace("\t", "")
                                        .Trim(),
                                RIC = ric
                            };
                            listEntity.Add(websiteTableEntity);
                        }
                        else
                        {
                            websiteTableEntity = new WebsitTableEntity
                            {
                                OFFC_CODE2 =
                                    trs[i].SelectNodes(".//td")[1].InnerText.Replace("&nbsp;", "")
                                        .Replace("\n", "")
                                        .Replace("\r", "")
                                        .Replace("\t", "")
                                        .Trim(),
                                RIC = ric
                            };
                            listEntity.Add(websiteTableEntity);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                listErrorRic.Add(querySecuritiesCode);
                Logger.Log(string.Format("Error found during task: {0}. Exception message: {1}", "Error generated by GetDataFromWebsiteUrlToDRExchangeEntity", ex.Message));
            }
        }

        private void GetDataFromGATSToList(string strGats, List<string> list)
        {
            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(strGats, null);
            Regex regex = new Regex(strPatternRIC);
            MatchCollection matches = regex.Matches(response);
            foreach (Match match in matches.Cast<Match>().Where(match => !list.Contains(match.Groups["RIC"].Value)))
            {
                list.Add(match.Groups["RIC"].Value);
            }
        }

        private void GetDataFromGATSToExistDic(string strGats, Dictionary<string, string> dicExist)
        {
            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(strGats, null);
            Regex regex = new Regex(strPatternValue);
            MatchCollection matches = regex.Matches(response);
            foreach (Match match in matches.Cast<Match>().Where(match => !dicExist.ContainsKey(match.Groups["RIC"].Value)))
            {
                dicExist.Add(match.Groups["RIC"].Value, match.Groups["Value"].Value);
            }
        }
    }
}
