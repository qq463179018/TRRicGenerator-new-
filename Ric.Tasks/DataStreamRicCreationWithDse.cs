using System;
using System.ComponentModel;
using Ric.Core;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Text;
using System.Collections.Generic;
using HtmlAgilityPack;
using Ric.Util;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Drawing.Design;


namespace Ric.Tasks
{
    #region Configuration

    [ConfigStoredInDB]
    public class DataStreamRicCreationWithDseConfig
    {
        [StoreInDB]
        [Category("Path")]
        [DefaultValue("D:\\DataStream\\RIC_Creation\\")]
        [DisplayName("Output path")]
        [Description("Path to save generated output file. E.g.D:\\DataStream\\RIC_Creation\\")]
        public string OutputPath { get; set; }

        [Category("Announcement Date")]
        [Description("Date format: MMdd. E.g. 1206")]
        public string Date { get; set; }

        [StoreInDB]
        [Category("SEDOL Login Info")]
        [DefaultValue("xxx.xxx@thomsonreuters.com")]
        [Description("Username for login to the SEDOL source website.")]
        public string Username { get; set; }

        [StoreInDB]
        [Category("SEDOL Login Info")]
        [Description("Password for login to the SEDOL source website.")]
        public string Password { get; set; }

        [StoreInDB]
        [Category("Proxy")]
        [DefaultValue("10.40.14.56")]
        [Description("Proxy IP address for assess to DSE.")]
        public string IP { get; set; }

        [StoreInDB]
        [Category("Proxy")]
        [DefaultValue("80")]
        [Description("Proxy port for assess to DSE.")]
        public string Port { get; set; }

        [StoreInDB]
        [Category("Whether treat the abnormal record")]
        [DefaultValue("false")]
        [Description("If set true,program will treat the abnormal record")]
        public bool isTreatAbnormalRecord { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

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

        public DataStreamRicCreationWithDseConfig()
        {
            Date = DateTime.Today.AddDays(-1).ToString("MMdd");
        }

    }

    #endregion

    #region Data Model

    public class DataStreamRicCreationInfo
    {
        public string Ticker { get; set; }
        public string Sedol { get; set; }
        public string CompanyName { get; set; }
        public string FirstTradingDate { get; set; }
        public string ExchangeCode { get; set; }
        public string Isin { get; set; }
        public string SecurityDescription { get; set; }
        public string AssetCategory { get; set; }
        public string SecurityLongDescription { get; set; }
        public string ThomsonReutersClassificationScheme { get; set; }
        public string CUSIP { get; set; }
        public string ReutersEditorialRIC { get; set; }
        public string RIC { get; set; }
        public string CurrencyCode { get; set; }
        public string fileType { get; set; }
        public string roundLotSize { get; set; }
        public string marketSegmentName { get; set; }
        public string originalRecord { get; set; }
        public string malaysiaCode { get; set; }
    }

    public class DataStreamRicCreationBulkTemplate
    {
        public string Seqn { get; set; }
        public string Mnem { get; set; }
        public string Sedol { get; set; }
        public string Name1 { get; set; }
        public string Name2 { get; set; }
        public string Bdate { get; set; }
        public string DefMki { get; set; }
        public string Nosh { get; set; }
        public string Mnem2 { get; set; }
        public string Isin { get; set; }
        public string CName { get; set; }
        public string NameSuffix { get; set; }

        //For HongKong. Mark if this record is an H-share
        public bool IsHShare { get; set; }

        //For HongKong. Two added field.
        public string Geog { get; set; }
        public string Dcur { get; set; }

        //For US
        public string Remk { get; set; }
        public string Secty { get; set; }
        public string Grpi { get; set; }
        public string Fname1 { get; set; }
        public string Mic { get; set; }
        public string Fname2 { get; set; }
        public string Cusip { get; set; }
        public string Qci { get; set; }
        public string Adr_Gdr { get; set; }
        public string PARENT_RIC { get; set; }

        public DataStreamRicCreationBulkTemplate()
        {
            Seqn = Mnem = Sedol = Name1 = Name2 = Bdate = DefMki = Nosh = Mnem2 = Isin = CName = NameSuffix = Geog = Dcur = string.Empty;
            IsHShare = false;
        }
    }

    public class DseFieldPosition
    {
        public int StartPosition { get; set; }
        public int EndPosition { get; set; }

        public DseFieldPosition(int start, int end)
        {
            StartPosition = start;
            EndPosition = end;
        }
    }

    public class DseRecord
    {
        public string DseContent { get; set; }
        public string DseFileType { get; set; }

        public DseRecord(string content, string fileType)
        {
            DseContent = content;
            DseFileType = fileType;
        }
    }
    #endregion


    public class DataStreamCommon
    {
        public static Dictionary<string, string> DownloadNameRules(Logger logger)
        {
            Dictionary<string, string> namesAbbs = new Dictionary<string, string>();
            string url = @"http://dataops.datastream.com/cgi-bin/readfile.pl?filename=H:/Production/Loaders/Global/DataStream/Tools/Abbreviation/Mload/abbreviations.rep&warnold=1";
            HtmlDocument doc = null;
            int retry = 3;
            while (doc == null && retry-- > 0)
            {
                try
                {
                    string pageSource = WebClientUtil.GetPageSource(null, url, 180000, null, Encoding.GetEncoding("ISO-8859-1"));
                    if (pageSource != null)
                    {
                        doc = new HtmlDocument();
                        doc.LoadHtml(pageSource);
                    }
                }
                catch (Exception ex)
                {
                    string msg = "At DownloadNameRules(). Error found when downloading the name abbreviations file. " + ex.Message;
                    logger.Log(msg);
                }
            }
            HtmlNode xmpNode = doc.DocumentNode.SelectSingleNode("//xmp");
            if (xmpNode == null)
            {
                string msg = string.Format("At DownloadNameRules(). Can not get the name abbreviation in file.");
                logger.Log(msg, Logger.LogType.Warning);
                return namesAbbs;
            }
            string content = xmpNode.InnerText;

            string[] lines = content.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            int startLine = 0;
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("MEANING") && lines[i].Contains("ABBREVATION"))
                {
                    startLine = i + 1;
                    break;
                }
            }
            for (int i = startLine; i < lines.Length; i++)
            {
                string singleLine = lines[i];
                if (!singleLine.Contains("   "))
                {
                    continue;
                }
                string[] names = singleLine.Split(new string[] { "   " }, StringSplitOptions.RemoveEmptyEntries);
                if (names.Length != 2)
                {
                    string msg = string.Format("At DownloadNameRules(). Irregular name and abbrevation at line: {0} in 'Abbrevation file'. Ignore it.\r\n\t\t\t\t Line content:{1}.", i, singleLine);
                    //logger.Log(msg);
                    continue;
                }
                if (!namesAbbs.ContainsKey(names[0].Trim()))
                {
                    namesAbbs.Add(names[0].Trim(), names[1].Trim());
                }
                else
                {
                    string msg = string.Format("At DownloadNameRules(). Repeated name at line: {0} in Abbrevation file. Line content:{1}.", i, singleLine);
                    //logger.Log(msg);
                }
            }
            return namesAbbs;
        }
        private static bool CheckValidationResult(object senter, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }
        private static string Encode(string encodeText)
        {
            return Regex.Replace(encodeText, "[^a-zA-Z0-9]", delegate(Match match) { return "%" + BitConverter.ToString(Encoding.GetEncoding("euc-kr").GetBytes(match.Value)).Replace("-", "%"); });
        }
        public static void LoginToWebsite(string Username, string password, CookieContainer cookies, string jSessionId, Logger logger)
        {
            string username = Encode(Username);
            string uri = string.Format("https://www.unavista.londonstockexchange.com/datasolutions/login.html");
            string postData = string.Format("callCount=1\r\nwindowName=unavista_datasolutions\r\nc0-scriptName=LoginHelper\r\nc0-methodName=doLogin\r\nc0-id=0\r\nc0-param0=string:{0}\r\nc0-param1=null:null\r\nc0-param2=string:{1}\r\nc0-param3=boolean:false\r\nc0-param4=string:Google%20Chrome\r\nc0-param5=null:null\r\nc0-param6=boolean:false\r\nc0-param7=null:null\r\nc0-param8=null:null\r\nc0-param9=string:11.9.0.23449\r\nbatchId=2\r\npage=%2Fdatasolutions%2Flogin.html\r\nhttpSessionId=\r\nscriptSessionId=", username, password);
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
            try
            {
                HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.ContentType = "text/plain";
                request.UserAgent = "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
                request.Method = "POST";
                request.KeepAlive = true;
                request.AllowAutoRedirect = true;
                request.Referer = "https://www.unavista.londonstockexchange.com/datasolutions/login.html";
                request.Accept = "*/*";
                request.ContentLength = 2080;

                request.CookieContainer = cookies;
                request.Credentials = CredentialCache.DefaultCredentials;


                byte[] buf = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);

                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();
                string jsessionid = httpResponse.GetResponseHeader("Set-Cookie");

                Regex regex = new Regex("JSESSIONID=(?<JSessionId>.*?); Path");
                MatchCollection matches = regex.Matches(jsessionid);
                foreach (Match match in matches)
                {
                    jSessionId = match.Groups["JSessionId"].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error in LoginToWebsite():" + ex.ToString();
                logger.Log(msg, Logger.LogType.Error);
            }
        }

        public static string QuerySedol(string isin, string jSessionId, CookieContainer cookies, Logger logger)
        {
            string uri = "https://www.unavista.londonstockexchange.com/datasolutions/dwr/call/plaincall/ClusterListHelper.loadClusterList.dwr";
            string postData = "";

            string[] postContent = new string[] {             
                                                "callCount=1", 
                                                "windowName=unavista_datasolutions", 
                                                "c0-scriptName=ClusterListHelper", 
                                                "c0-methodName=loadClusterList", 
                                                "c0-id=0", 
                                                "c0-param0=number:2199194", 
                                                "c0-param1=number:2199194", 
                                                "c0-param2=number:1085", 
                                                "c0-param3=number:10", 
                                                "c0-param4=boolean:false", 
                                                "c0-param5=null:null", 
                                                "c0-param6=null:null", 
                                                "c0-param7=array:[]", 
                                                "c0-param8=array:[]", 
                                                "c0-param9=null:null", 
                                                "c0-e2=string:(", 
                                                "c0-e3=string:ts.%5BISIN%5D", 
                                                "c0-e4=string:OR", 
                                                "c0-e5=string:false", 
                                                "c0-e6=string:" + isin, 
                                                "c0-e7=string:string", 
                                                "c0-e8=null:null", 
                                                "c0-e9=string:%3D", 
                                                "c0-e1=Object_Object:{openparen:reference:c0-e2, field:reference:c0-e3, logicaloperator:reference:c0-e4, unicode:reference:c0-e5, value:reference:c0-e6, datatype:reference:c0-e7, closeparen:reference:c0-e8, operator:reference:c0-e9}", 
                                                "c0-e11=null:null", 
                                                "c0-e12=string:ts.%5BSEDOL%5D", 
                                                "c0-e13=string:OR", 
                                                "c0-e14=string:false", 
                                                "c0-e15=string:" + isin, 
                                                "c0-e16=string:string", 
                                                "c0-e17=null:null", 
                                                "c0-e18=string:%3D", 
                                                "c0-e10=Object_Object:{openparen:reference:c0-e11, field:reference:c0-e12, logicaloperator:reference:c0-e13, unicode:reference:c0-e14, value:reference:c0-e15, datatype:reference:c0-e16, closeparen:reference:c0-e17, operator:reference:c0-e18}", 
                                                "c0-e20=null:null", 
                                                "c0-e21=string:ts.%5BprevSEDOL%5D", 
                                                "c0-e22=null:null", 
                                                "c0-e23=string:false", 
                                                "c0-e24=string:" + isin, 
                                                "c0-e25=string:string", 
                                                "c0-e26=string:)", 
                                                "c0-e27=string:%3D", 
                                                "c0-e19=Object_Object:{openparen:reference:c0-e20, field:reference:c0-e21, logicaloperator:reference:c0-e22, unicode:reference:c0-e23, value:reference:c0-e24, datatype:reference:c0-e25, closeparen:reference:c0-e26, operator:reference:c0-e27}", 
                                                "c0-param10=array:[reference:c0-e1,reference:c0-e10,reference:c0-e19]", 
                                                "c0-param11=boolean:false", 
                                                "c0-param12=boolean:false", 
                                                "c0-param13=boolean:false", 
                                                "c0-param14=null:null", 
                                                "c0-param15=null:null", 
                                                "c0-param16=null:null", 
                                                "c0-param17=array:[]", 
                                                "c0-param18=boolean:false", 
                                                "c0-param19=boolean:false", 
                                                "c0-param20=null:null", 
                                                "c0-param21=null:null", 
                                                "c0-param22=null:null", 
                                                "c0-param23=null:null", 
                                                "batchId=15", 
                                                "page=%2Fdatasolutions%2Funavistalanding.html", 
                                                "httpSessionId=" + jSessionId, 
                                                "scriptSessionId=" 
                                                 };
            postData = string.Join("\r\n", postContent);
            try
            {
                HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.ContentType = "text/plain";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
                request.Method = "POST";
                request.KeepAlive = true;
                request.AllowAutoRedirect = true;
                request.Referer = "https://www.unavista.londonstockexchange.com/datasolutions/unavistalanding.html";
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                request.CookieContainer = cookies;
                request.Credentials = CredentialCache.DefaultCredentials;

                byte[] buf = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);

                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();

                StreamReader sr = new StreamReader(httpResponse.GetResponseStream());

                string content = sr.ReadToEnd();

                return content;

            }
            catch (Exception ex)
            {
                string msg = "Error found in QuerySedol():" + ex.ToString();
                logger.Log(msg, Logger.LogType.Error);
                return "";
            }

        }

        public static string ParseSedol(string response, string isin, Logger logger)
        {
            string pattern = string.Format(@"data:\[\[.*?{0}.*?\]\]", isin);
            Regex regex = new Regex(pattern);
            Match match = regex.Match(response);
            if (match.Success)
            {
                string[] dataList = match.Value.Split(new string[] { "\",\"" }, StringSplitOptions.RemoveEmptyEntries);
                if (dataList.Length > 8)
                {
                    return dataList[7];
                }
            }
            logger.Log("At ParseSedol(). Can not get SEDOL from website.", Logger.LogType.Error);
            return "";
        }  
    
        
    }

    public enum nameInputType
    {
        Name,
        Fname,
    }
}
