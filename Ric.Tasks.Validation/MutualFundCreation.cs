using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Design;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Office.Interop.Excel;
//using MSAD.Common.OfficeUtility;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Ric.Core;
using Ric.Util;
using MSAD.Common.OfficeUtility;

namespace Ric.Tasks.Validation
{
    #region Config

    [ConfigStoredInDB]
    public class MutualFundCreationConfig
    {
        [StoreInDB]
        [Category("Source File")]
        [Description("Path of the result excel file")]
        public string ResultFilePath { get; set; }

        [StoreInDB]
        [Category("Cusip Account")]
        [Description("Username on the cusip website")]
        public string CusipUsername { get; set; }

        [StoreInDB]
        [Category("Cusip Account")]
        [Description("Password for the cusip website")]
        public string CusipPassword { get; set; }

        [StoreInDB]
        [Category("Mail Account")]
        [Description("Username of the email account")]
        public string MailUsername { get; set; }

        [StoreInDB]
        [Category("Mail Account")]
        [Description("Password of the email account")]
        public string MailPassword { get; set; }

        [StoreInDB]
        [Category("Mail Account")]
        [Description("Domain of the email account")]
        public string MailDomain { get; set; }

        [StoreInDB]
        [Category("Mail Account")]
        [Description("Email address")]
        public string Email { get; set; }

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

    /// <summary>
    /// Mutual Fund representation
    /// </summary>
    public class MutualFund
    {
        public string Mnem { get; set; }
        public string Name { get; set; }
        public string FundFamily { get; set; }
        public string Loc { get; set; }
        public string Isin { get; set; }
        public DateTime BDate { get; set; }
    }

    #endregion

    #region Task

    public class MutualFundCreation : GeneratorBase
    {
        #region Declaration

        private static MutualFundCreationConfig configObj;
        private List<string> titles;
        private List<MutualFund> funds;
        private Ric.Util.ExcelApp app;
        private ExchangeService service;
        private string excelPath;
        private string resultFileName;
        private CookieContainer cookies;
        Dictionary<string, string> abbr = new Dictionary<string, string>();

        #endregion

        #region Interface implementation

        protected override void Initialize()
        {
            base.Initialize();
            funds = new List<MutualFund>();
            cookies = new CookieContainer();
            configObj = Config as MutualFundCreationConfig;
            try
            {
                service = EWSUtility.CreateService(new System.Net.NetworkCredential(configObj.MailUsername, configObj.MailPassword, configObj.MailDomain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
                app = new Ric.Util.ExcelApp(false, false);
                if (app.ExcelAppInstance == null)
                {
                    Logger.Log("Excel cannot be started", Logger.LogType.Error);
                }
                titles = new List<string>() { "MNEM", "NAME", "FUND FAMILY", "SEQN", "LOC", "SECD", "ISIN", "BDATE" };
            }
            catch (Exception ex)
            {
                Logger.Log("Initialization failed. Ex: " + ex.Message);
                throw new Exception("Cannot initialize");
            }
        }

        protected override void Start()
        {
            Logger.Log("Starting Task");
            try
            {
                excelPath = DownladExcel();
                GetAbbreviation();
                ReadExcel();
                LoginToWebsite();
                foreach (MutualFund fund in funds)
                {
                    if (fund.Loc.Length == 9)
                    {
                        Search(fund.Loc);
                        GetSummary();
                        fund.Isin = GetIsinFromIssueDetail(fund.Loc);
                    }
                }
                CreateExcel();
                SendEmail();
                Logger.Log("Task finished");
            }
            catch (Exception ex)
            {
                Logger.Log("Error: " + ex.Message, Logger.LogType.Error);
                throw new Exception("Task failed: " + ex.Message, ex);
            }
            finally
            {
                app.Dispose();
            }
        }

        #endregion

        #region Download Excel

        /// <summary>
        /// Find the latest Email with the title "Nasdaq MF corp file"
        /// Download the attachments and put it in the result directory
        /// </summary>
        /// <returns>The name of the downloaded file</returns>
        private string DownladExcel()
        {
            try
            {
                EWSMailSearchQuery query = new EWSMailSearchQuery("", configObj.Email, "Inbox\\", "Mutual Fund", "", DateTime.Now.AddDays(-4), DateTime.Now);
                List<EmailMessage> mailList = EWSMailSearchQuery.SearchMail(service, query);
                string path = Path.GetDirectoryName(configObj.ResultFilePath);

                foreach (EmailMessage mail in mailList)
                {
                    mail.Load();
                    List<string> attachments = EWSMailSearchQuery.DownloadAttachments(service, mail, "", "", path);
                    return attachments[0];
                }
                return null;
            }
            catch (Exception ex)
            {
                Logger.Log("Cannot get excel from email");
                throw new Exception("Get excel file from mail failed.\r\nCheck your Email login credentials.\n\r: " + ex.Message, ex);
            }
        }

        #endregion

        #region Reading Excel

        /// <summary>
        /// Get worksheet with source information
        /// </summary>
        /// <returns>the worksheet</returns>
        private Worksheet GetWorksheet(out Workbook workbook)
        {
            workbook = ExcelUtil.CreateOrOpenExcelFile(app, excelPath);
            Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
            return worksheet;
        }

        /// <summary>
        /// Getting infos from excel
        /// </summary>
        public void ReadExcel()
        {
            Workbook workbook;
            Worksheet worksheet = GetWorksheet(out workbook);
            try
            {
                for (int line = 1; worksheet.Range["A" + line].Value2 != null; line++)
                {
                    if (worksheet.Range["I" + line].Value2 != null && worksheet.Range["I" + line].Value2.ToString() == "Z")
                    {
                        funds.Add(GetMutualFundFromLine(worksheet.Range["A" + line, "BK" + line]));
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception("Cannot read workbook");
            }
            finally
            {
                workbook.SaveAs(workbook.FullName, XlFileFormat.xlWorkbookNormal);
                workbook.Close();
                //File.Delete(excelPath);
            }
        }

        /// <summary>
        /// Return a MutualFund object from Excel line
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        private MutualFund GetMutualFundFromLine(Range range)
        {
            MutualFund fund = new MutualFund();

            fund.BDate = DateTime.FromOADate(Convert.ToDouble(range.Range["C1"].Value2.ToString()));
            fund.FundFamily = range.Range["Q1"].Value2.ToString();
            fund.Loc = range.Range["M1"].Value2.ToString();
            fund.Mnem = range.Range["K1"].Value2.ToString();
            fund.Name = range.Range["O1"].Value2.ToString();

            return fund;
        }

        #endregion

        #region Parsing Website

        /// <summary>
        /// To validate everytime HTTPS Certificate
        /// </summary>
        /// <param name="senter"></param>
        /// <param name="certificate"></param>
        /// <param name="chain"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private static bool CheckValidationResult(object senter, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="postData"></param>
        /// <param name="url"></param>
        /// <param name="referer"></param>
        /// <returns></returns>
        private Stream MakeRequest(string type, string postData, string uri, string referer)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
            HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;

            try
            {
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.ContentType = "application/x-www-form-urlencoded";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
                request.Method = type;
                request.KeepAlive = true;
                request.AllowAutoRedirect = true;
                request.Referer = referer;
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                request.CookieContainer = cookies;
                request.Credentials = CredentialCache.DefaultCredentials;

                if (type == "POST")
                {
                    byte[] buf = Encoding.UTF8.GetBytes(postData);
                    request.ContentLength = buf.Length;
                    request.GetRequestStream().Write(buf, 0, buf.Length);
                }

                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();
                return (httpResponse.GetResponseStream());
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                throw new WebException("Cannot send request to " + uri);
            }
        }

        /// <summary>
        /// Connect to the website and get the Cookie
        /// </summary>
        private void LoginToWebsite()
        {
            string uri = string.Format("https://www.access.cusip.com/cusipaccess/accessLogin.htm");
            string postData = string.Format("userFrom=CUSIP&action=login&productCode=UU1UP&userId={0}&password={1}&forcedLogin=Y&rememberMe=true&_rememberMe=on", configObj.CusipUsername, configObj.CusipPassword);
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
            try
            {
                HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.ContentType = "application/x-www-form-urlencoded";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
                request.Method = "POST";
                request.KeepAlive = true;
                request.AllowAutoRedirect = true;
                request.Referer = "https://www.access.cusip.com/cusipaccess/home.htm";
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                request.CookieContainer = cookies;
                request.Credentials = CredentialCache.DefaultCredentials;

                byte[] buf = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);

                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();

                Stream page = httpResponse.GetResponseStream();

                using (Stream file = File.OpenWrite(configObj.ResultFilePath + "login.html"))
                {
                    byte[] buffer = new byte[8 * 1024];
                    int len;
                    while ((len = page.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        file.Write(buffer, 0, len);
                    }
                }

            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex.ToString();
                throw new WebException("Cannot send request to " + uri);
            }
        }

        /// <summary>
        /// Send a request to quickSearch.htm with the Loc as parameter
        /// </summary>
        private void Search(string loc)
        {
            string uri = string.Format("https://www.access.cusip.com/cusipaccess/quickSearch.htm");
            string postData = string.Format("submitSearch=Y&issuerDescription=&identifierValue={0}", loc);
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
            try
            {
                HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.ContentType = "application/x-www-form-urlencoded";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
                request.Method = "POST";
                request.KeepAlive = true;
                request.AllowAutoRedirect = true;
                request.Referer = "https://www.access.cusip.com/cusipaccess/descriptionSearchHome.htm";
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                request.CookieContainer = cookies;
                request.Credentials = CredentialCache.DefaultCredentials;

                byte[] buf = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);

                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();

                Stream page = httpResponse.GetResponseStream();

                using (Stream file = File.OpenWrite(configObj.ResultFilePath + "search.html"))
                {

                    byte[] buffer = new byte[8 * 1024];
                    int len;
                    while ((len = page.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        file.Write(buffer, 0, len);
                    }
                }

            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex.ToString();
                throw new WebException("Cannot send request to " + uri);
            }
        }

        private void GetSummary()
        {
            string url = String.Format("https://www.access.cusip.com/cusipaccess/issueSummary.htm");
            try
            {
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
                request.Method = "GET";
                request.KeepAlive = true;
                request.AllowAutoRedirect = true;
                request.Referer = "https://www.access.cusip.com/cusipaccess/quickSearch.htm";
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                request.CookieContainer = cookies;
                request.Credentials = CredentialCache.DefaultCredentials;

                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();

                Stream page = httpResponse.GetResponseStream();

                using (Stream file = File.OpenWrite(configObj.ResultFilePath + "summary.html"))
                {

                    byte[] buffer = new byte[8 * 1024];
                    int len;
                    while ((len = page.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        file.Write(buffer, 0, len);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex.ToString();
                throw new WebException("Cannot get response from Url : " + url);
            }
        }

        /// <summary>
        /// Search the ISIN in the Table of the Issue Detail page
        /// </summary>
        private string GetIsinFromIssueDetail(string loc)
        {
            string isin = "";
            string url = String.Format("https://www.access.cusip.com/cusipaccess/issueDetail.htm?issuerNumber={0}&issueNumber={1}&issueCheck={2}", loc.Substring(0, 6), loc.Substring(6, 2), loc.Substring(8, 1));
            try
            {
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
                request.Method = "GET";
                request.KeepAlive = true;
                request.Referer = "https://www.access.cusip.com/cusipaccess/issueSummary.htm";
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
                request.CookieContainer = cookies;
                request.Credentials = CredentialCache.DefaultCredentials;
                request.Headers["Cache-Control"] = "max-age=0";
                request.Headers["Accept-Encoding"] = "gzip,deflate,sdch";
                request.Headers["Accept-Language"] = "en-US,en;q=0.8";
                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();

                Stream page = httpResponse.GetResponseStream();
                HtmlDocument html = new HtmlDocument();
                html.Load(page);

                HtmlNodeCollection tables = html.DocumentNode.SelectNodes(".//table");
                if (tables.Count > 4)
                {
                    isin = tables[6].SelectSingleNode(".//tr[4]/td[2]").InnerText;
                }
                return isin;
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                throw new WebException("Cannot get response from Url : " + url);
            }
        }

        #endregion

        #region Abbreviation

        /// <summary>
        /// Get abbreviation array from website
        /// </summary>
        private void GetAbbreviation()
        {
            HtmlDocument htc = new HtmlDocument();
            string uri = String.Format("http://dataops.datastream.com/cgi-bin/readfile.pl?filename=H:/Production/Loaders/Global/DataStream/Tools/Abbreviation/Mload/abbrev.txt&warnold=1");
            htc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
            HtmlNode fullFile = htc.DocumentNode.SelectSingleNode(".//xmp[1]");
            string abbrFile = fullFile.InnerText;

            Regex rgxSpace = new Regex(@"\s+");
            string[] stringSeparators = new string[] { "\r\n" };
            char[] stringSeparators2 = new char[] { ' ' };

            string[] lines = abbrFile.Split(stringSeparators, StringSplitOptions.None);
            foreach (string line in lines)
            {
                string formattedLine = rgxSpace.Replace(line.Trim(), " ");
                string[] lineTab = formattedLine.Split(stringSeparators2);
                if (lineTab.Length >= 2)
                {
                    string key = "";
                    for (int i = 0; i < lineTab.Length - 1; i++)
                    {
                        key += lineTab[i] + " ";
                    }
                    key = key.Trim();
                    if (!abbr.ContainsKey(key))
                    {
                        abbr[key] = lineTab[lineTab.Length - 1];
                    }
                }

            }
        }

        /// <summary>
        /// Abbreviate a string
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private string Abbreviate(string name)
        {
            char[] stringSeparators2 = new char[] { ' ' };
            string[] lineTab = name.Replace("-", " ").Replace("  ", " ").Replace("   ", "").Split(stringSeparators2);
            int wordNb = lineTab.Length - 1;
            int wordsInString;
            List<string> keysFound = new List<string>();

            for (wordsInString = wordNb; wordsInString > 0; wordsInString--)
            {
                for (int startindex = 0; startindex + wordsInString <= wordNb; startindex++)
                {
                    string keyToSearch = "";
                    for (int index = startindex; index < startindex + wordsInString; index++)
                    {
                        if (lineTab[index].StartsWith("-") || lineTab[index].EndsWith("-"))
                        {
                            keyToSearch += lineTab[index + 1].TrimStart().Replace("-", "") + " ";
                        }
                        else
                        {
                            keyToSearch += lineTab[index + 1].TrimStart() + " ";
                        }
                    }
                    keyToSearch = keyToSearch.TrimEnd();
                    if (abbr.ContainsKey(keyToSearch))
                    {
                        keysFound.Add(keyToSearch);
                    }
                }
            }
            foreach (var key in keysFound)
            {
                name = name.Replace(key, abbr[key]);
            }
            name = name.Replace(" ", ".").Replace("..", ".").Replace(".-.", " ").Replace("-.", " ").Replace(".-", " ").Replace("-", " ").Replace(".POR", " POR").Replace(lineTab[0] + ".", lineTab[0] + "_");
            if (name.EndsWith("."))
            {
                return name.Substring(0, (name.Length - 1));
            }
            return name;
        }
        #endregion

        #region Creating Excel

        /// <summary>
        /// Filling titles in the new excel
        /// </summary>
        /// <param name="worksheet"></param>
        private void FillTitles(ref Worksheet worksheet)
        {
            if (worksheet != null)
            {
                for (int column = 0; column < titles.Count; column++)
                {
                    worksheet.Cells[1, column + 1] = titles[column];
                }
            }
        }

        /// <summary>
        /// Initialize the workbook and worksheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="worksheet"></param>
        private void InitializeExcel(out Workbook workbook, out Worksheet worksheet)
        {
            resultFileName = DateTime.Now.ToString("dMMMyyyy_HHmm");
            string title = String.Format("{0}Results_{1}.xls", configObj.ResultFilePath, resultFileName);
            workbook = ExcelUtil.CreateOrOpenExcelFile(app, title);
            worksheet = null;
            if (workbook != null)
            {
                if ((worksheet = (Worksheet)workbook.Worksheets[1]) != null)
                {
                    worksheet.Name = "MUFD";
                }
            }
        }

        /// <summary>
        /// Excel creation 
        /// </summary>
        public void CreateExcel()
        {
            Worksheet worksheet = null;
            Workbook workbook = null;
            int line = 2;

            try
            {
                InitializeExcel(out workbook, out worksheet);
                FillTitles(ref worksheet);
                foreach (MutualFund fund in funds)
                {
                    worksheet.Cells[line, 1] = fund.Mnem;
                    worksheet.Cells[line, 2] = Abbreviate(fund.Name.ToUpper());
                    worksheet.Cells[line, 3] = Abbreviate(fund.FundFamily.ToUpper());
                    if (fund.Isin == null)
                    {

                        if (fund.Loc.Length >= 8)
                        {
                            worksheet.Cells[line, 5] = "U" + fund.Loc;
                        }
                        else
                        {
                            worksheet.Cells[line, 5] = "U" + fund.Loc.Substring(0, 8);
                        }
                        worksheet.Cells[line, 7] = "";
                    }
                    else
                    {
                        worksheet.Cells[line, 5] = "U" + fund.Isin.Substring(2, 8);
                        worksheet.Cells[line, 7] = fund.Isin;
                    }
                    worksheet.Cells[line, 8] = fund.BDate;
                    Range rg3 = (Range)worksheet.Cells[line, 8];
                    rg3.NumberFormat = "dd/mm/yy";

                    rg3.Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    line++;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error while creating new Excel file: " + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception("Error while creating new Excel file");
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.SaveAs(workbook.FullName, XlFileFormat.xlWorkbookNormal);
                    workbook.Close();
                }
            }
        }

        #endregion

        #region Send Email

        /// <summary>
        /// Sending Email with the newly created CSV attached
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        private void SendEmail()
        {
            try
            {
                List<string> mailCC = new List<string>();
                List<string> attachement = new List<string>();
                StringBuilder bodyBuilder = new StringBuilder();
                string body = String.Empty;
                string title = String.Format("MUTUAL FUND {0}", DateTime.Now.ToString("dd/MM/yy"));

                bodyBuilder.Append("<p>");
                foreach (string signatureLine in configObj.MailSignature)
                {
                    bodyBuilder.AppendFormat("{0}<br />", signatureLine);
                }
                bodyBuilder.Append("</p>");
                body = bodyBuilder.ToString();
                body = bodyBuilder.ToString();
                attachement.Add(String.Format("{0}Results_{1}.xls", configObj.ResultFilePath, resultFileName));
                if (configObj.MailCC.Count > 1 || (configObj.MailCC.Count == 1 && configObj.MailCC[0] != ""))
                {
                    mailCC = configObj.MailCC;
                }
                MSAD.Common.OfficeUtility.EWSUtility.CreateAndSendMail(service, configObj.MailTo, mailCC, new List<string>(), title, body, attachement);
            }
            catch (Exception ex)
            {
                string msg = "Cannot send mail :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        #endregion
    }

    #endregion
}
