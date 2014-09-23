using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.Net;
using System.IO;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Ric.Core;
using Ric.Tasks;
using Ric.Util;

namespace Ric.Tasks
{
    class DataStreamRicCreationWithDseVN : GeneratorBase
    {
        #region Fields

        private DataStreamRicCreationWithDseConfig ConfigObj = null;
        private Dictionary<string, string> namesAbbs = new Dictionary<string, string>();
        private CookieContainer cookies = new CookieContainer();
        private bool ifDownNameRule = false;
        private bool loginSuccess = false;
        private bool noshSessioned = false;
        private string jSessionId = string.Empty;

        private string mFileFolder = string.Empty;

        #endregion

        #region Initialize and Start

        protected override void Initialize()
        {
            ConfigObj = Config as DataStreamRicCreationWithDseConfig;
            TaskResultList.Add(new TaskResultEntry("LOG File", "LOG File", Logger.FilePath));

            InitializeFileDirectory();
            string msg = "Initialize...OK!";
            Logger.Log(msg);
        }

        private void InitializeFileDirectory()
        {
            string outputFolder = Path.Combine(ConfigObj.OutputPath, DateTime.Today.ToString("yyyy-MM-dd"));

            mFileFolder = Path.Combine(outputFolder, "DSE_FILES");

            if (!Directory.Exists(mFileFolder))
            {
                Directory.CreateDirectory(mFileFolder);
            }

            TaskResultList.Add(new TaskResultEntry("DSE_FILES", "DSE_FILES", mFileFolder));
        }

        protected override void Start()
        {
            StartJob();
        }

        private void StartJob()
        {
            DownloadFtpFiles();
            List<DseRecord> records = GetRecords();
            if (records == null || records.Count == 0)
            {
                Logger.Log("No RIC Creation today.");
                return;
            }
            List<DataStreamRicCreationInfo> ricCreations = FormatRecords(records);

            if (ifDownNameRule)
            {
                DownloadNameRules();
            }

            string lineToFile = FormatOutputLine(ricCreations);
            if (!string.IsNullOrEmpty(lineToFile))
            {
                GenerateFile(lineToFile);
            }

        }
        #endregion

        private void DownloadFtpFiles()
        {
            string currentDate = String.Empty;

            string[] fileStartArr = new string[] { "1068", "1435", "1779","EM01","EM11" };

            foreach (string fileStart in fileStartArr)
            {
                currentDate = ConfigObj.Date;
                string fileName = string.Format("{0}{1}.M", fileStart, currentDate);

                string ftpfullpath = @"ftp://ASIA2:ASIA2@ds1.rds.reuters.com/" + fileName;

                try
                {
                    WebClient request = new WebClient();

                    if (!string.IsNullOrEmpty(ConfigObj.IP) && !string.IsNullOrEmpty(ConfigObj.Port))
                    {
                        WebProxy proxy = new WebProxy(ConfigObj.IP, Convert.ToInt32(ConfigObj.Port));
                        request.Proxy = proxy;
                    }

                    request.Credentials = new NetworkCredential("ASIA2", "ASIA2");

                    string mfilePath = Path.Combine(mFileFolder, fileName);

                    request.DownloadFile(ftpfullpath, mfilePath);
                   
                    Logger.Log(string.Format("Download FTP File {0}... OK!", fileName));
                }
                catch (Exception ex)
                {
                    string msg = string.Format("Can not download file: {0} from FTP. Response:{1}", fileName, ex.Message);
                    Logger.Log(msg, Logger.LogType.Error);
                    LogMessage(string.Format("Can not download file: {0}", fileName));
                }
            }

        }

        private List<DseRecord> GetRecords()
        {
            List<DseRecord> xeRecord = new List<DseRecord>();
            string currentDate = String.Empty;

            string[] fileStartArr = new string[] { "1068", "1435", "1779","EM01","EM11" };


            foreach (string fileStart in fileStartArr)
            {
                currentDate = ConfigObj.Date;
                string fileName = string.Format("{0}{1}.M", fileStart, currentDate);

                string mfilePath = Path.Combine(mFileFolder, fileName);

                if (!File.Exists(mfilePath))
                {
                    continue;
                }

                using (StreamReader sr = new StreamReader(mfilePath))
                {
                    string tmp = null;
                    while ((tmp = sr.ReadLine()) != null)
                    {
                        if (!tmp.StartsWith("XE"))
                        {
                            continue;
                        }

                        xeRecord.Add(new DseRecord(tmp, fileStart));

                        string msg = string.Format("Get 1 record from file: {0}.", fileName);
                        Logger.Log(msg);

                    }
                }
            }


            return xeRecord;
        }

        private List<DataStreamRicCreationInfo> FormatRecords(List<DseRecord> records)
        {
            Dictionary<string, DseFieldPosition> dseFields = InitializeDseFields();
            List<DataStreamRicCreationInfo> ricCreations = new List<DataStreamRicCreationInfo>();
            foreach (DseRecord record in records)
            {
                DataStreamRicCreationInfo item = new DataStreamRicCreationInfo();
                item.Ticker = FormatField(record, "Ticker", dseFields);
                item.AssetCategory = FormatField(record, "AssetCategory", dseFields);
                item.Sedol = FormatField(record, "Sedol", dseFields);
                item.SecurityLongDescription = FormatField(record, "SecurityLongDescription", dseFields);
                item.CompanyName = FormatField(record, "CompanyName", dseFields);
                if (!string.IsNullOrEmpty(item.CompanyName))
                {
                    ifDownNameRule = true;
                }
                item.FirstTradingDate = FormatField(record, "FirstTradingDate", dseFields);
                item.ExchangeCode = FormatField(record, "ExchangeCode", dseFields);
                item.Isin = FormatField(record, "Isin", dseFields);
                item.ThomsonReutersClassificationScheme = FormatField(record, "ThomsonReutersClassificationScheme", dseFields);
                item.CUSIP = FormatField(record, "CUSIP", dseFields);
                item.ReutersEditorialRIC = FormatField(record, "ReutersEditorialRIC", dseFields);
                item.RIC = FormatField(record, "RIC", dseFields);
                item.fileType = record.DseFileType;

                if (item.fileType.StartsWith("EM"))
                {
                    string exchangeCode = item.ExchangeCode;
                    if (exchangeCode.CompareTo("UPC") != 0 && exchangeCode.CompareTo("HSX") != 0 && exchangeCode.CompareTo("HNX") != 0)
                    {
                        continue;
                    }
                }

                ricCreations.Add(item);
            }
            return ricCreations;
        }

        private Dictionary<string, DseFieldPosition> InitializeDseFields()
        {
            Dictionary<string, DseFieldPosition> dseFields = new Dictionary<string, DseFieldPosition>();
            dseFields.Add("Ticker", new DseFieldPosition(412, 436));
            dseFields.Add("AssetCategory", new DseFieldPosition(407, 410));
            dseFields.Add("Sedol", new DseFieldPosition(68, 74));
            dseFields.Add("SecurityLongDescription", new DseFieldPosition(437, 546));
            dseFields.Add("CompanyName", new DseFieldPosition(105, 184));
            dseFields.Add("FirstTradingDate", new DseFieldPosition(1022, 1029));
            dseFields.Add("ExchangeCode", new DseFieldPosition(98, 100));
            dseFields.Add("Isin", new DseFieldPosition(84, 95));
            dseFields.Add("ThomsonReutersClassificationScheme", new DseFieldPosition(806, 815));
            dseFields.Add("CUSIP", new DseFieldPosition(59, 67));
            dseFields.Add("ReutersEditorialRIC", new DseFieldPosition(347, 363));
            dseFields.Add("RIC", new DseFieldPosition(3, 22));

            return dseFields;
        }

        private string FormatField(DseRecord record, string field, Dictionary<string, DseFieldPosition> dseFields)
        {
            if (!dseFields.ContainsKey(field))
            {
                return "";
            }
            DseFieldPosition dseField = dseFields[field];
            if (record.DseContent.Length < dseField.EndPosition + 1)
            {
                string msg = string.Format("At FormatField(). Can not get field:{0}. Out of length.", field);
                Logger.Log(msg, Logger.LogType.Error);
                return "";
            }
            string result = record.DseContent.Substring(dseField.StartPosition - 1, dseField.EndPosition - dseField.StartPosition + 1).Trim().ToUpper();
            return result;
        }

        private void DownloadNameRules()
        {
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
                    Logger.Log(msg);
                }
            }
            HtmlNode xmpNode = doc.DocumentNode.SelectSingleNode("//xmp");
            if (xmpNode == null)
            {
                string msg = string.Format("At DownloadNameRules(). Can not get the name abbreviation in file.");
                Logger.Log(msg, Logger.LogType.Warning);
                return;
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
                    //Logger.Log(msg);
                    continue;
                }
                if (!namesAbbs.ContainsKey(names[0].Trim()))
                {
                    namesAbbs.Add(names[0].Trim(), names[1].Trim());
                }
                else
                {
                    string msg = string.Format("At DownloadNameRules(). Repeated name at line: {0} in Abbrevation file. Line content:{1}.", i, singleLine);
                    //Logger.Log(msg);
                }
            }
        }

        private string FormatOutputLine(List<DataStreamRicCreationInfo> ricCreations)
        {
            List<string> lines = new List<string>();
            foreach (DataStreamRicCreationInfo item in ricCreations)
            {
                string temp = "";
                string name1 = string.Empty;
                string name2 = string.Empty;
                string suffix = string.Empty;
                DataStreamRicCreationBulkTemplate bulkTemp = new DataStreamRicCreationBulkTemplate();
                bulkTemp.Seqn = temp.PadRight(7, ' ');
                bulkTemp.Mnem = GetMnem(item.Ticker);
                if (string.IsNullOrEmpty(item.Sedol))
                {
                    string sedol = GetSedol(item.Isin);
                    if (string.IsNullOrEmpty(sedol))
                    {
                        sedol = temp.PadRight(9, ' ');
                    }
                    else
                    {
                        sedol = "UK" + sedol;
                    }
                    bulkTemp.Sedol = sedol;
                }
                else
                {
                    bulkTemp.Sedol = "UK" + item.Sedol;
                }

                if (item.AssetCategory.CompareTo("PRF") == 0 || item.AssetCategory.CompareTo("CPR") == 0)
                {
                    FormatBulkNames(item.SecurityLongDescription, 24, 24, nameInputType.Name, ref name1, ref name2, ref suffix);
                    bulkTemp.NameSuffix = suffix;
                    bulkTemp.Name1 = name1;
                    bulkTemp.Name2 = name2;

                    FormatBulkNames(item.SecurityLongDescription, 19, 24, nameInputType.Fname, ref name1, ref name2, ref suffix);
                    bulkTemp.Fname1 = name1;
                    bulkTemp.Fname2 = name2;
                }
                else
                {
                    FormatBulkNames(item.CompanyName, 24, 24, nameInputType.Name, ref name1, ref name2, ref suffix);
                    bulkTemp.NameSuffix = suffix;
                    bulkTemp.Name1 = name1;
                    bulkTemp.Name2 = name2;

                    FormatBulkNames(item.CompanyName, 19, 24, nameInputType.Fname, ref name1, ref name2, ref suffix);
                    bulkTemp.Fname1 = name1;
                    bulkTemp.Fname2 = name2;
                }

                if (item.AssetCategory.CompareTo("PRF") == 0 || item.AssetCategory.CompareTo("CPR") == 0)
                {
                    FormatCName(bulkTemp, item.CompanyName);
                }
                else
                {
                    name1 = bulkTemp.Name1.Trim();
                    name2 = bulkTemp.Name2.Trim();

                    string names = string.Empty;
                    if (!name1.EndsWith("."))
                    {
                        names = name1 + " " + name2;
                    }
                    else
                    {
                        names = name1 + name2;
                    }
                    FormatCName(bulkTemp, names);
                }


                if (item.FirstTradingDate.Equals("-9999999"))
                {
                    bulkTemp.Bdate = temp.PadRight(8, ' ');
                }
                else
                {
                    //string bdate = DateTime.ParseExact(item.FirstTradingDate, "ddMMyyyy", CultureInfo.InvariantCulture).ToString("dd/MM/yy", DateTimeFormatInfo.InvariantInfo);
                    string bdate = item.FirstTradingDate.Substring(6, 2) + "/" + item.FirstTradingDate.Substring(4, 2) + "/" + item.FirstTradingDate.Substring(2, 2);
                    bulkTemp.Bdate = bdate;
                }

                if (item.ExchangeCode.CompareTo("UPC") == 0)
                {
                     bulkTemp.DefMki = "HO";
                }
                else if (item.ExchangeCode.CompareTo("HNX") == 0)
                {
                    bulkTemp.DefMki = "HS";
                }
                else if (item.ExchangeCode.CompareTo("HSX") == 0)
                {
                    bulkTemp.DefMki = "HC";
                }
                else
                {
                    bulkTemp.DefMki = string.Empty.PadRight(2,' ');
                }
                bulkTemp.Nosh = "1";
                bulkTemp.Mnem2 = "VT"+item.Ticker.PadRight(10,' ');
                bulkTemp.Isin = item.Isin;

                bulkTemp.Geog = "103";
                bulkTemp.Dcur = "007";
                bulkTemp.Remk = string.Empty.PadRight(4,' ');
              
                bulkTemp.Secty = "EQ";
                bulkTemp.Grpi = "116";
              


               
                bulkTemp.Mic = string.Empty.PadRight(3, ' ');

                bulkTemp.Cusip = string.Empty.PadRight(12, ' ');
                
                bulkTemp.Qci = string.Empty.PadRight(2,' ');
                  

                if (string.Compare(item.RIC, item.ReutersEditorialRIC) == 0)
                {
                    bulkTemp.PARENT_RIC = string.Empty;
                }
                else
                {
                    bulkTemp.PARENT_RIC = item.ReutersEditorialRIC;
                }

                string line = string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",\"{12}\",\"{13}\",\"{14}\",\"{15}\",\"{16}\",\"{17}\",\"{18}\",\"{19}\"",
                                            bulkTemp.Mnem, bulkTemp.Sedol, bulkTemp.Name1, bulkTemp.Name2,
                                            bulkTemp.Bdate, bulkTemp.DefMki, bulkTemp.Mnem2, bulkTemp.Isin, bulkTemp.CName,
                                            bulkTemp.Geog, bulkTemp.Dcur, bulkTemp.Remk, bulkTemp.Secty, bulkTemp.Grpi, bulkTemp.Fname1, bulkTemp.Mic,
                                            bulkTemp.Fname2, bulkTemp.Cusip, bulkTemp.Qci, bulkTemp.PARENT_RIC);
                lines.Add(line);

                System.Threading.Thread.Sleep(5000);

            }
            string outputLine = string.Join("\r\n", lines.ToArray());
            return outputLine;
        }

        private string GetMnem(string ticker)
        {
            string mnem = "VT:"+ticker.PadRight(3,' ');

            return mnem;
        }

        private string GetSedol(string isin)
        {
            if (string.IsNullOrEmpty(isin))
            {
                return "";
            }
            try
            {
                if (!loginSuccess)
                {
                    LoginToWebsite();
                    loginSuccess = true;
                }
                string response = QuerySedol(isin);
                return ParseSedol(response, isin);
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GetSedol(). Error found in getting SEDOL for ISIN:{0}. Error Message: {1}. ", isin, ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return "";
            }
        }

        public static string Encode(string encodeText)
        {
            return Regex.Replace(encodeText, "[^a-zA-Z0-9]", delegate(Match match) { return "%" + BitConverter.ToString(Encoding.GetEncoding("euc-kr").GetBytes(match.Value)).Replace("-", "%"); });
        }

        private static bool CheckValidationResult(object senter, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }

        private void LoginToWebsite()
        {
            string username = Encode(ConfigObj.Username);
            string uri = string.Format("https://www.unavista.londonstockexchange.com/datasolutions/login.html");
            string postData = string.Format("callCount=1\r\nwindowName=unavista_datasolutions\r\nc0-scriptName=LoginHelper\r\nc0-methodName=doLogin\r\nc0-id=0\r\nc0-param0=string:{0}\r\nc0-param1=null:null\r\nc0-param2=string:{1}\r\nc0-param3=boolean:false\r\nc0-param4=string:Google%20Chrome\r\nc0-param5=null:null\r\nc0-param6=boolean:false\r\nc0-param7=null:null\r\nc0-param8=null:null\r\nc0-param9=string:11.9.0.23449\r\nbatchId=2\r\npage=%2Fdatasolutions%2Flogin.html\r\nhttpSessionId=\r\nscriptSessionId=", username, ConfigObj.Password);
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
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private string QuerySedol(string isin)
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
                Logger.Log(msg, Logger.LogType.Error);
                return "";
            }

        }

        private string ParseSedol(string response, string isin)
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
            Logger.Log("At ParseSedol(). Can not get SEDOL from website.", Logger.LogType.Error);
            return "";
        }

        //private void FormatBulkNames(string companyName, DataStreamRicCreationBulkTemplate bulkTemp)
        //{
        //    //Rules:
        //    //1. 提取第一个单词后trim
        //    //2. 去掉5个后缀 if 有
        //    //3. 去掉company后缀， if 有
        //    //4. 如果Joint。。。Company出现在中间， 用JSC.替换
        //    //5. 开始找缩写
        //    //6. 开始format两个name
        //    //7. if > 48, user input
        //    string temp = string.Empty;
        //    if (string.IsNullOrEmpty(companyName))
        //    {
        //        bulkTemp.Name1 = temp.PadRight(24, ' ');
        //        bulkTemp.Name2 = temp.PadRight(24, ' ');
        //        return;
        //    }

        //    string inputConpanyName = companyName;
        //    companyName = companyName.ToUpper();

        //    string[] nameWords = companyName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        //    //Company Name only contains one word.
        //    if (nameWords.Length == 1)
        //    {
        //        string abbr = companyName;
        //        if (namesAbbs.ContainsKey(companyName))
        //        {
        //            abbr = namesAbbs[companyName];
        //        }
        //        bulkTemp.Name1 = abbr.PadRight(24, ' ');
        //        bulkTemp.Name2 = temp.PadRight(24, ' ');
        //        return;
        //    }

        //    string nameSuffix = string.Empty;
        //    string nameLeft = TrimCompanyEndings(companyName, ref nameSuffix);
        //    bulkTemp.NameSuffix = nameSuffix;

        //    List<string> abbreviations = GetNameAbbreviations(nameLeft);

        //    string namePart1 = string.Empty;
        //    string namePart2 = string.Empty;

        //    bool reFormat = false;
        //    do
        //    {
        //        if (reFormat)
        //        {
        //            if (abbreviations[abbreviations.Count - 1] == "CORP.")
        //            {
        //                bulkTemp.NameSuffix = "CORP.";
        //            }

        //            if (abbreviations.Contains("CORP."))
        //            {
        //                abbreviations.Remove("CORP.");
        //            }
        //            if (abbreviations.Contains("COMPANY"))
        //            {
        //                abbreviations.Remove("COMPANY");
        //            }
        //        }
        //        string formatedName = string.Join("", abbreviations.ToArray()).Trim();
        //        if (formatedName.Length > 48)
        //        {
        //            if (reFormat)
        //            {
        //                string[] namesInput = DataStreamRicCreationNamesInput.Prompt(inputConpanyName, formatedName);
        //                if (namesInput != null)
        //                {
        //                    namePart1 = namesInput[0];
        //                    namePart2 = namesInput[1];
        //                }
        //            }
        //            else
        //            {
        //                reFormat = true;
        //            }
        //        }
        //        else if (formatedName.Length <= 24)
        //        {
        //            namePart1 = formatedName;
        //        }

        //        else
        //        {
        //            int index = 24;
        //            int index2 = 24;
        //            char[] charArr = formatedName.ToCharArray();
        //            if (charArr[index] == ' ')
        //            {
        //                index2 = index;
        //            }
        //            else
        //            {
        //                while (true)
        //                {
        //                    char middleChar = charArr[index];
        //                    if (middleChar == ' ' || middleChar == '.')
        //                    {
        //                        break;
        //                    }
        //                    index--;
        //                }
        //                index2 = index + 1;
        //            }
        //            namePart1 = formatedName.Substring(0, index2);
        //            namePart2 = formatedName.Substring(index2).Trim();
        //            if (namePart2.Length > 24)
        //            {
        //                if (reFormat)
        //                {
        //                    string[] namesInput = DataStreamRicCreationNamesInput.Prompt(inputConpanyName, formatedName);
        //                    if (namesInput != null)
        //                    {
        //                        namePart1 = namesInput[0];
        //                        namePart2 = namesInput[1];
        //                    }
        //                }
        //                else
        //                {
        //                    reFormat = true;
        //                }
        //            }
        //        }
        //    }
        //    while (reFormat);

        //    bulkTemp.Name1 = namePart1.PadRight(24, ' ');
        //    bulkTemp.Name2 = namePart2.PadRight(24, ' ');
        //}

        private void FormatBulkNames(string companyName, int nameLength1, int nameLength2, nameInputType type, ref string name1, ref string name2, ref string suffix)
        {
            string temp = string.Empty;
            if (string.IsNullOrEmpty(companyName))
            {
                name1 = temp.PadRight(nameLength1, ' ');
                name2 = temp.PadRight(nameLength2, ' ');
                return;
            }

            string inputConpanyName = companyName;
            companyName = companyName.ToUpper();

            string[] nameWords = companyName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            //Company Name only contains one word.
            if (nameWords.Length == 1)
            {
                string abbr = companyName;
                if (namesAbbs.ContainsKey(companyName))
                {
                    abbr = namesAbbs[companyName];
                }
                name1 = abbr.PadRight(nameLength1, ' ');
                name2 = temp.PadRight(nameLength2, ' ');
                return;
            }

            string nameSuffix = string.Empty;
            string nameLeft = TrimCompanyEndings(companyName, ref nameSuffix);
            suffix = nameSuffix;

            List<string> abbreviations = GetNameAbbreviations(nameLeft, nameLength1, nameLength2);

            string namePart1 = string.Empty;
            string namePart2 = string.Empty;

            bool reFormat = false;
            do
            {
                if (reFormat)
                {
                    if (abbreviations[abbreviations.Count - 1] == "CORP.")
                    {
                        suffix = "CORP.";
                    }

                    if (abbreviations.Contains("CORP."))
                    {
                        abbreviations.Remove("CORP.");
                    }
                    if (abbreviations.Contains("COMPANY"))
                    {
                        abbreviations.Remove("COMPANY");
                    }
                }
                string formatedName = string.Join("", abbreviations.ToArray()).Trim();
                if (formatedName.Length > (nameLength1 + nameLength2))
                {
                    if (reFormat)
                    {
                        string[] namesInput = DataStreamRicCreationNamesInput.Prompt(inputConpanyName, formatedName, type);
                        if (namesInput != null)
                        {
                            namePart1 = namesInput[0];
                            namePart2 = namesInput[1];
                        }
                        reFormat = false;
                    }
                    else
                    {
                        reFormat = true;
                    }
                }
                else if (formatedName.Length <= nameLength1)
                {
                    namePart1 = formatedName;
                }

                else
                {
                    int index = nameLength1;
                    int index2 = nameLength2;
                    char[] charArr = formatedName.ToCharArray();
                    if (charArr[index] == ' ')
                    {
                        index2 = index;
                    }
                    else
                    {
                        while (true)
                        {
                            char middleChar = charArr[index];
                            if (middleChar == ' ' || middleChar == '.')
                            {
                                break;
                            }
                            index--;
                        }
                        index2 = index + 1;
                    }
                    namePart1 = formatedName.Substring(0, index2);
                    namePart2 = formatedName.Substring(index2).Trim();
                    if (namePart2.Length > 24)
                    {
                        if (reFormat)
                        {
                            string[] namesInput = DataStreamRicCreationNamesInput.Prompt(inputConpanyName, formatedName, type);
                            if (namesInput != null)
                            {
                                namePart1 = namesInput[0];
                                namePart2 = namesInput[1];
                            }
                            reFormat = false;
                        }
                        else
                        {
                            reFormat = true;
                        }
                    }
                }
            }
            while (reFormat);

            name1 = namePart1.PadRight(nameLength1, ' ');
            name2 = namePart2.PadRight(nameLength2, ' ');
        }

        private string TrimCompanyEndings(string nameLeft, ref string nameSuffix)
        {
            string[] endings = new string[] { "JSC", "JOINT STOCK COMPANY", "JOINT STOCK CO", "JOINT STOCK CORPORATION", "JOINT-STOCK CORPORATION", "PUBLIC LIMITED COMPANY", "INCORPORATED", "COMPANY", "LIMITED", "LTD", "CORPORATION", "CO.", "CORP", "INC", "LLC", "PLC.", "PLC", "Public Limited Company", "Public Limited Co", "Public Limited Co." };
            foreach (string ending in endings)
            {
                if ((nameLeft.Length > ending.Length) && ((nameLeft.Substring(nameLeft.Length - ending.Length - 1).Equals(" " + ending)) || (nameLeft.Substring(nameLeft.Length - ending.Length - 1).Equals("." + ending))))
                {
                    nameLeft = nameLeft.Replace(ending, "");
                    if (ending.Equals("JSC") || ending.Equals("JOINT STOCK COMPANY") || ending.Equals("JOINT STOCK CO") || ending.Equals("JOINT STOCK CORPORATION"))
                    {
                        nameSuffix = "JSC.";
                    }
                    else if (ending.Equals("CORPORATION"))
                    {
                        nameSuffix = "CORP.";
                    }
                    else if (ending.Equals("LIMITED") || ending.Equals("LTD"))
                    {
                        nameSuffix = "LTD.";
                    }
                    else if (ending.Equals("INCORPORATED") || ending.Equals("INC"))
                    {
                        nameSuffix = "INCO.";
                    }
                    else if (ending.Equals("LLC"))
                    {
                        nameSuffix = "LLC.";
                    }
                    else if (ending.Equals("PLC.") || ending.Equals("PLC") || ending.Equals("Public Limited Company") || ending.Equals("Public Limited Co") || ending.Equals("Public Limited Co."))
                    {
                        nameSuffix = "PLC.";
                    }
                }
            }

            if (nameLeft.Contains(" JOINT STOCK COMPANY "))
            {
                nameLeft = nameLeft.Replace("JOINT STOCK COMPANY", "JSC.");
            }

            return nameLeft.Trim();
        }

        //private List<string> GetNameAbbreviations(string nameLeft)
        //{

        //    List<string> abbrevations = new List<string>();

        //    List<string> nameLeftArr = nameLeft.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();

        //    //Rule 14 in requirement v3.0. Abbreviate from right to left.            

        //    if (nameLeft.Length <= 24)
        //    {
        //        foreach (string subName in nameLeftArr)
        //        {
        //            abbrevations.Add(subName + " ");
        //        }
        //        return abbrevations;
        //    }

        //    string firstWord = nameLeftArr[0] + " ";
        //    nameLeftArr.RemoveAt(0);

        //    while (nameLeftArr.Count > 0)
        //    {
        //        //string namePart = nameLeftArr[i];
        //        string nameToFind = string.Join(" ", nameLeftArr.ToArray());
        //        for (int j = 0; j <= nameLeftArr.Count - 1; j++)
        //        {
        //            if (!namesAbbs.ContainsKey(nameToFind) && j != nameLeftArr.Count - 1)
        //            {
        //                nameToFind = nameToFind.Replace(nameLeftArr[j], "").Trim();
        //                continue;
        //            }
        //            string abbr = string.Empty;
        //            if (j == nameLeftArr.Count - 1)
        //            {
        //                if (namesAbbs.ContainsKey(nameToFind))
        //                {
        //                    abbr = namesAbbs[nameToFind] + ".";
        //                }
        //                else
        //                {
        //                    abbr = nameLeftArr[j];
        //                    if (!abbr.Contains("."))
        //                    {
        //                        abbr = abbr + " ";
        //                    }
        //                }
        //                if (abbr.Equals("&."))
        //                {
        //                    abbr = "&";
        //                }

        //            }
        //            else
        //            {
        //                abbr = namesAbbs[nameToFind] + ".";
        //            }
        //            //for (int k = 0; k <= nameLeftArr.Count - 1 - j; )
        //            //{
        //            //    nameLeftArr.RemoveAt(nameLeftArr.Count - 1);
        //            //}

        //            int k = nameLeftArr.Count - 1 - j;
        //            while (k-- >= 0)
        //            {
        //                nameLeftArr.RemoveAt(nameLeftArr.Count - 1);
        //            }

        //            if (!string.IsNullOrEmpty(abbr))
        //            {
        //                abbrevations.Add(abbr);
        //                string nameFormatTemp = firstWord + string.Join(" ", nameLeftArr.ToArray()) + " " + string.Join("", abbrevations.ToArray());
        //                if (nameFormatTemp.Length <= 24)
        //                {
        //                    for (int i = nameLeftArr.Count - 1; i >= 0; i--)
        //                    {
        //                        abbrevations.Add(nameLeftArr[i] + " ");
        //                    }

        //                    abbrevations.Add(firstWord);
        //                    abbrevations.Reverse();
        //                    return abbrevations;
        //                }
        //            }
        //        }
        //    }
        //    abbrevations.Add(firstWord);
        //    abbrevations.Reverse();
        //    return abbrevations;
        //}

        private List<string> GetNameAbbreviations(string nameLeft, int nameLength1, int nameLength2)
        {

            List<string> abbrevations = new List<string>();

            List<string> nameLeftArr = nameLeft.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            //Rule 14 in requirement v3.0. Abbreviate from right to left.            

            if (nameLeft.Length <= nameLength1)
            {
                foreach (string subName in nameLeftArr)
                {
                    abbrevations.Add(subName + " ");
                }
                return abbrevations;
            }

            string firstWord = nameLeftArr[0] + " ";
            nameLeftArr.RemoveAt(0);

            while (nameLeftArr.Count > 0)
            {
                //string namePart = nameLeftArr[i];
                string nameToFind = string.Join(" ", nameLeftArr.ToArray());
                for (int j = 0; j <= nameLeftArr.Count - 1; j++)
                {
                    if (!namesAbbs.ContainsKey(nameToFind) && j != nameLeftArr.Count - 1)
                    {
                        nameToFind = nameToFind.Replace(nameLeftArr[j], "").Trim();
                        continue;
                    }
                    string abbr = string.Empty;
                    if (j == nameLeftArr.Count - 1)
                    {
                        if (namesAbbs.ContainsKey(nameToFind))
                        {
                            abbr = namesAbbs[nameToFind] + ".";
                        }
                        else
                        {
                            abbr = nameLeftArr[j];
                            if (!abbr.Contains("."))
                            {
                                abbr = abbr + " ";
                            }
                        }
                        if (abbr.Equals("&."))
                        {
                            abbr = "&";
                        }

                    }
                    else
                    {
                        abbr = namesAbbs[nameToFind] + ".";
                    }
                    //for (int k = 0; k <= nameLeftArr.Count - 1 - j; )
                    //{
                    //    nameLeftArr.RemoveAt(nameLeftArr.Count - 1);
                    //}

                    int k = nameLeftArr.Count - 1 - j;
                    while (k-- >= 0)
                    {
                        nameLeftArr.RemoveAt(nameLeftArr.Count - 1);
                    }

                    if (!string.IsNullOrEmpty(abbr))
                    {
                        abbrevations.Add(abbr);
                        string nameFormatTemp = firstWord + string.Join(" ", nameLeftArr.ToArray()) + " " + string.Join("", abbrevations.ToArray());
                        if (nameFormatTemp.Length <= nameLength1)
                        {
                            for (int i = nameLeftArr.Count - 1; i >= 0; i--)
                            {
                                abbrevations.Add(nameLeftArr[i] + " ");
                            }

                            abbrevations.Add(firstWord);
                            abbrevations.Reverse();
                            return abbrevations;
                        }
                    }
                }
            }
            abbrevations.Add(firstWord);
            abbrevations.Reverse();
            return abbrevations;
        }

        private void FormatCName(DataStreamRicCreationBulkTemplate bulkTemp, string names)
        {

            if (names.EndsWith(".") && (names.Length + bulkTemp.NameSuffix.Length) <= 24)
            {
                names += bulkTemp.NameSuffix;
            }
            else if ((!names.EndsWith(".")) && (names.Length + 1 + bulkTemp.NameSuffix.Length) <= 24)
            {
                if (names.EndsWith(" "))
                {
                    names += bulkTemp.NameSuffix;
                }
                else
                {
                    names += " " + bulkTemp.NameSuffix;
                }

            }
            else if (names.Length > 24)
            {
                names = names.Substring(0, 24);
            }

            bulkTemp.CName = names.PadRight(24, ' ');
        }

        private void GenerateFile(string lineToFile)
        {
            try
            {
                string outputFolder = Path.Combine(ConfigObj.OutputPath, DateTime.Today.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                string template = InitializeMacTemplate();
                string dataLine = "[DataStreamRicCreationInfoLine]";
                template = template.Replace(dataLine, lineToFile);
                string fileName = string.Format("VN_{0}.MAC", ConfigObj.Date);
                string filePath = Path.Combine(outputFolder, fileName);
                File.WriteAllText(filePath, template, Encoding.ASCII);
                TaskResultList.Add(new TaskResultEntry("Output Folder", "Output Folder", outputFolder));
                TaskResultList.Add(new TaskResultEntry("MAC File", "MAC File", filePath));

                string msg = string.Format("Generate normal RIC creation output file...OK!");
                Logger.Log(msg);
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GenerateFile(). Error found in generating MAC file. Error message:{0}.", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private string InitializeMacTemplate()
        {
            string template1 = string.Empty;
            string templateFilePath = Directory.GetCurrentDirectory() + @"\\DataStreamTemplate\\00.MAC";
            if (File.Exists(templateFilePath))
            {
                string[] comtent = File.ReadAllLines(templateFilePath);
                template1 = string.Join("\r\n", comtent);
            }
            else
            {
                template1 = ";=================================================================================================================\r\n\r\n;Start on Primary Menu\r\n;Make sure DLIVE06 is already running to avoid password issues\r\n\r\n;DESCRIPTION: For Common Type of Creation\r\n\r\n;=================================================================================================================\r\n\r\nLOOP:\r\nIF &ENDOFDATA=FALSE THEN\r\nINPUT MNEM,SEDOL,NAME1,NAME2,BDATE,DEF_MKT,MNEM2,ISIN,CNAME,GEOG,DCUR,REMK,SECTY,GRPI,FNAME1,MIC,FNAME2,CUSIP,QCI,PARENT_RIC\r\n\r\nSET NOSH TO \"1\"\r\n\r\nIF (mid$(BDATE,3,1)=\"/\") THEN\r\n\tSET CODE3 TO \"   \"\r\nELSE\r\n\tSET CODE3 TO BDATE\r\n\tSET BDATE TO \"[TAB]\"\r\nENDIF\r\n\r\n;=================================================================================================================\r\n\r\n;SEQUENCE NUMBER (SEQN)\r\n\r\nSend( \"[CLEAR]\" )\r\n\r\nSET DIRCHECK TO mid$(NAME1,2,1)\r\nIF (DIRCHECK <\"9\") AND (DIRCHECK<>\" \") THEN\r\n\tSET DIRCHECK TO \"Z\"\r\nENDIF\t\r\nIF (DIRCHECK<\"M\") THEN\r\n\tSET DIR TO \"BEGIN\"\r\n\tSET DIR2 TO \"[PF8]\"\r\n\tSET DIR3 TO \"FORWARD\"\r\nELSE\r\n\tSET DIR TO \"END\"\r\n\tSET DIR2 TO \"[PF7]\"\r\n\tSET DIR3 TO \"BACKWARD\"\r\nENDIF\r\n\r\n\r\n\tSET TYPE TO \"FAMER\"\r\n\tSET WHAT TO \"PAGE\"\t\r\n\r\nPREP:\r\n\r\n\tIF (TYPE=\"FAMER\")AND (WHAT=\"PAGE\") THEN\r\n\t\tSET NameStr TO left$(NAME1,1)\r\n\t\tSend (\"15\")\r\n\t\tSend (\"7\")\r\n\t\tSend (\"FAMER\"+NameStr+\"[TAB]\"+DIR)\r\n\r\n\t\tSET NameStr TO left$(NAME1,7)\r\n\t\tIF ((right$(NAME1,1)<>\" \")AND(left$(NAME2,1)<>\" \")) THEN\r\n\t\t\tSET str to NAME1+\" \"+NAME2\r\n\t\tELSE \r\n\t\t\tSET str to NAME1+NAME2\r\n\t\tENDIF\r\n\t\r\n\tENDIF\r\n\t\r\n\tIF (WHAT=\"LINE\") THEN\r\n\t\tSET nl TO 0\r\n\t\tSET CNT TO 1 \r\n\t\t\t\t\r\n\t\tLOOPORDERLINE:\r\n\t\tIF (CNT>15) THEN\r\n\t\t\tSET nl TO 0\r\n\t\t\tIF TYPE=\"FAMER\" THEN\r\n\t\t\t\tSET PREVIOUS TO mid$(&screen,1602,6)\r\n\t\t\tELSE\r\n\t\t\t\t;TYPE=\"ALLEQ\r\n\t\t\t\tSET PREVIOUS TO mid$(&screen,1674,7)\t\t\t\t\r\n\t\t\tENDIF\r\n\t\t\t\tCHECKBOTTOM:\r\n\t\t\t\tIF left$(PREVIOUS,1)=\" \" THEN\r\n\t\t\t\t\tSET nl TO nl+1\r\n\t\t\t\t\tIF TYPE=\"FAMER\" THEN\r\n\t\t\t\t\t\tSET PREVIOUS TO mid$(&screen,1602-(80*nl),6)\r\n\t\t\t\t\tELSE\r\n\t\t\t\t\t\t;TYPE=\"ALLEQ\r\n\t\t\t\t\t\tSET PREVIOUS TO mid$(&screen,1674-(80*nl),7)\t\t\t\t\r\n\t\t\t\t\tENDIF\r\n\t\t\t\t\r\n\t\t\t\t\tGOTO CHECKBOTTOM\r\n\t\t\t\tENDIF\r\n\t\t\t\r\n\t\t\tSET DIR3 TO \"FORWARD\"\r\n\t\t\tSET DIR2 TO \"[PF8]\"\t\t\t\t\t\t\t\t\r\n\t\t\t\r\n\t\t\tGOTO ENDPAGECHECK\r\n\t\tENDIF\r\n\r\n\t\tSET str2 TO mid$(&screen,489+(80*nl),26)\r\n\t\tIF (left$(str,1)<>left$(str2,1)) AND (left$(str,1)>left$(str2,1)) THEN\r\n\t\t\tGOTO GETNEXT1\r\n\t\tENDIF\r\n\t\t\t\r\n\t\tGOTO CHARCHECK\r\n\tENDIF\r\n\t\r\n\tIF (TYPE=\"ALLEQ\") AND (WHAT=\"PAGE\") THEN\r\n\t\tSend (\"15\")\r\n\t\tSend (\"7\")\r\n\t\tSend (\"ALLEQ[TAB]\"+SEQN_BASE)\r\n\tENDIF\r\n\r\n\tIF (WHAT=\"PAGE\") THEN\r\n\t\tPAGECHECK1:\r\n\t\r\n\t\t\tIF (DIR3=\"FORWARD\") THEN\r\n\t\t\t\tSET str2 TO mid$(&screen,1609,50)\r\n\t\t\tELSE\r\n\t\t\t\t;DIR3=\"BAcKWARD\"\r\n\t\t\t\tSET str2 TO mid$(&screen,489,50)\r\n\t\t\tENDIF\r\n\t\t\r\n\t\t\tGOTO CHARCHECK\r\n\tENDIF\r\n\r\nCHARCHECK:\r\n\r\n\t\t\tSET xx TO 1\r\n\t\t\tLOOPCHAR1:\r\n\t\t\t\tSET char1 TO mid$(str,xx,1)\r\n\t\t\t\tSET char2 TO mid$(str2,xx,1)\r\n\t\t\t\t\r\n\t\t\t\t;==CHECK FOR NUMBER AND SYMBOL\r\n\r\n\t\t\t\t\tIF ((char1>\"9\") AND (char2>\"9\")) OR (char1=\" \" OR char2=\" \" OR char1=\"\" OR char2=\"\" OR char1=\"-\" OR char2=\"-\" OR char1=\"&\" OR char2=\"&\" OR char1=\".\" OR char2=\".\" OR char1=\"+\" OR char2=\"+\" OR char1=\"/\" OR char2=\"/\" OR char1=\"'\" OR char2=\"'\") THEN\r\n\t\t\t\t\t\tGOTO EXITCHECK1\r\n\t\t\t\t\tENDIF\r\n\r\n\t\t\t\t\tIF ((char1<=\"9\") AND (char1<>\" \") AND (char2<=\"9\") AND (char2<>\" \")) THEN\r\n\t\t\t\t\t\t;NOTHING\r\n\t\t\t\t\t\t\r\n\t\t\t\t\tELSE \r\n\t\t\t\t\t\tIF (char1<=\"9\") THEN\t\t\r\n\t\t\t\t\t\t\tSET char1 TO \"ZZ\"\r\n\t\t\t\t\t\tENDIF\t\r\n\r\n\t\t\t\t\t\tIF (char2<=\"9\") THEN\t\t\t\t\t\t\t\r\n\t\t\t\t\t\t\tSET char2 TO \"ZZ\"\r\n\t\t\t\t\t\tENDIF\t\r\n\t\r\n\t\t\t\t\tENDIF\r\n\r\n\t\t\t\t\tEXITCHECK1:\r\n\t\t\t\t;==END CHECK FOR NUMBER AND SYMBOL\r\n\t\t\t\t\r\n\r\nEVAL:\r\n\r\n\t\t\tIF (WHAT=\"PAGE\") AND (DIR3=\"FORWARD\") THEN\r\n\t\t\t\t\tIF (char1<char2) THEN\r\n\t\t\t\t\t\tGOTO ENDLOOPCHAR1\r\n\r\n\t\t\t\t\tELSE \r\n\t\t\t\t\t\tIF (char1=char2) THEN\r\n\t\t\t\t\t\t\tSET xx TO xx +1\r\n\t\t\t\t\t\t\tGOTO LOOPCHAR1\r\n\t\t\t\t\t\tELSE\r\n\t\t\t\t\t\t\tIF (TYPE=\"FAMER\") THEN\r\n\t\t\t\t\t\t\t\tSET PREVIOUS TO mid$(&screen,1602,6)\t\t\t\t\r\n\t\t\t\t\t\t\tELSE\r\n\t\t\t\t\t\t\t\t;TYPE ALLEQ\r\n\t\t\t\t\t\t\t\tSET PREVIOUS TO mid$(&screen,1674,7)\r\n\t\t\t\t\t\t\tENDIF\r\n\t\t\t\t\t\t\t\r\n\t\t\t\t\t\t\tSend (DIR2)\r\n\t\t\t\t\t\t\tGOTO PAGECHECK1\r\n\t\t\t\t\t\tENDIF\r\n\t\t\t\t\tENDIF\r\n\t\t\tENDIF\r\n\r\n\t\t\tIF (WHAT=\"PAGE\") AND (DIR3=\"BACKWARD\") THEN\r\n\t\t\t\t\tIF (char1>char2) THEN\r\n\t\t\t\t\t\tGOTO ENDLOOPCHAR1\r\n\r\n\t\t\t\t\tELSE \r\n\t\t\t\t\t\tIF (char1=char2) THEN\r\n\t\t\t\t\t\t\tSET xx TO xx +1\r\n\t\t\t\t\t\t\tGOTO LOOPCHAR1\r\n\t\t\t\t\t\tELSE\r\n\t\t\t\t\t\t\tIF (TYPE=\"FAMER\") THEN\r\n\t\t\t\t\t\t\t\tSET PREVIOUS TO mid$(&screen,482,6)\r\n\t\t\t\t\t\t\tELSE\r\n\t\t\t\t\t\t\t\t;TYPE ALLEQ\r\n\t\t\t\t\t\t\t\tSET PREVIOUS TO mid$(&screen,554,7)\r\n\t\t\t\t\t\t\tENDIF\r\n\t\t\t\t\t\t\tSend (DIR2)\r\n\t\t\t\t\t\t\tGOTO PAGECHECK1\r\n\t\t\t\t\t\tENDIF\r\n\t\t\t\t\tENDIF\r\n\t\t\t\tENDLOOPCHAR1:\r\n\t\t\t\tSET WHAT TO \"LINE\"\r\n\t\t\t\tGOTO PREP\r\n\t\t\tENDIF\r\n\t\t\r\n\t\t\tIF (WHAT=\"LINE\") THEN\r\n\t\t\t\tIF (char1<char2) THEN\r\n\t\t\t\t\tSET DIR3 TO \"FORWARD\"\r\n\t\t\t\t\tSET DIR2 TO \"[PF8]\"\r\n\r\n\t\t\t\t\tCHECKTOP:\r\n\t\t\t\t\tIF (nl <> 0) THEN\r\n\t\t\t\t\t\tSET nl TO nl-1\t\r\n\t\t\t\t\t\tIF (TYPE=\"FAMER\") THEN\r\n\t\t\t\t\t\t\tSET PREVIOUS TO mid$(&screen,482+80*nl,6)\r\n\t\t\t\t\t\tELSE\r\n\t\t\t\t\t\t\t;TYPE=\"ALLEQ\"\r\n\t\t\t\t\t\t\tSET PREVIOUS TO mid$(&screen,554+80*nl,7)\r\n\r\n\t\t\t\t\t\tENDIF\r\n\t\t\t\t\tENDIF\r\n\t\r\n\t\t\t\t\tIF left$(PREVIOUS,1)=\" \" THEN\r\n\t\t\t\t\t\tSET nl TO nl+2\r\n\t\t\t\t\t\tSET DIR3 TO \"BACKWARD\"\r\n\t\t\t\t\t\tSET DIR2 TO \"[PF7]\"\r\n\t\t\t\t\t\tGOTO CHECKTOP\r\n\t\t\t\t\tENDIF\r\n\r\n\t\t\t\t\tGOTO ENDLOOPLINE\r\n\t\t\t\tELSE \r\n\t\t\t\t\tIF (char1=char2) THEN\r\n\t\t\t\t\t\tSET xx TO xx +1\r\n\t\t\t\t\t\tGOTO LOOPCHAR1\r\n\t\t\t\t\tELSE\t\t\t\t\t\t\r\n\t\t\t\t\t\tGOTO GETNEXT1\r\n\t\t\t\t\tENDIF\r\n\t\t\t\tENDIF\r\n\t\t\t\t\tGETNEXT1:\r\n\t\t\t\t\tSET CNT TO CNT+1\r\n\t\t\t\t\tSET nl TO nl+1\r\n\t\t\t\t\tGOTO LOOPORDERLINE\r\n\r\n\t\t\tENDIF\r\n\t\t\tENDLOOPLINE:\r\n\t\t\tENDPAGECHECK:\r\n\t\t\t\r\n\t\t\tIF (TYPE=\"FAMER\") AND (WHAT=\"LINE\") THEN\r\n\t\t\t\tSET SEQN_BASE TO PREVIOUS\r\n\t\t\t\tSend (\"[CLEAR]\")\r\n\t\t\t\tSET TYPE TO \"ALLEQ\"\r\n\t\t\t\tSET WHAT TO \"PAGE\"\r\n\t\t\t\tGOTO PREP\r\n\t\t\tENDIF\r\n\r\n\t\t\tIF (TYPE=\"ALLEQ\") AND (WHAT=\"LINE\") THEN\r\n\t\t\t\tSET GET_SEQN TO PREVIOUS\r\n\t\t\t\tSET K TO 0\r\n\t\t\t\tSET LAST TO \"\"\r\n\r\n\t\t\t\tLOOP_SEQNX:\r\n\t\t\t\tIF (K=7) THEN\r\n\t\t\t\t\tMESSAGE(\"Sequence number has exceeded 7 digits\")\r\n\t\t\t\t\tGOTO END\r\n\t\t\t\tENDIF\r\n\r\n\r\n\t\t\t\tSET DIGIT TO val(mid$(GET_SEQN,7-K,1))+1\r\n\t\t\t\tIF (DIGIT<=9) THEN\r\n\t\t\t\t\tIF (K=0) THEN\r\n\t\t\t\t\t\tSET LAST TO DIGIT\r\n\t\t\t\t\t\tSET SEQN TO left$(GET_SEQN,6+K)+str$(LAST)\r\n\t\t\t\t\t\tGOTO ENDLOOPSEQNX\r\n\t\t\t\t\tELSE\r\n\t\t\t\t\t\tSET SEQN TO left$(GET_SEQN,6-K)+str$(DIGIT)+LAST\r\n\t\t\t\t\t\tGOTO ENDLOOPSEQNX\r\n\t\t\t\t\tENDIF\r\n\t\t\t\tELSE\t\r\n\t\t\t\t\tSET ENDDIGIT TO \"0\"\t\t\r\n\t\t\t\t\tSET LAST TO ENDDIGIT+LAST\r\n\t\t\t\t\tSET K TO K+1\r\n\r\n\t\t\t\t\tGOTO LOOP_SEQNX\r\n\t\t\t\tENDIF\r\n\r\n\t\t\t\tENDLOOPSEQNX:\r\n\t\t\t\tSend(\"[CLEAR]\")\r\n\t\t\tENDIF\r\n\r\n;END GET SEQN\r\n\r\n;=================================================================================================================\r\n;COMMON STOCK, NORMAL STOCK\r\n\r\nCOMMON:\r\n\r\n\tSend( \"[CLEAR]\" )\r\n\tSend (\"14\")\r\n\tSend (\"S\")  \r\n\t\r\n\tIF (DEF_MKT=\"HK\") THEN\r\n\t\tSend (\"[DOWN][DOWN][DOWN]048YR\")\r\n\tELSE\r\n\t\tSend (\"[DOWN][DOWN][DOWN]\"+GEOG+\"YR\")\r\n\tENDIF\r\n\t\r\n\tSend (\"[ENTER]\")\r\n\tSET DSCD TO mid$(&screen,183,6)\r\n\t\r\n\tIF (DEF_MKT=\"HK\") AND (GEOG=\"048\") AND (DCUR=\"079\") THEN\r\n\t\tSend (SEQN+MNEM+\"            \"+REMK+\"[TAB]\"+NAME1+NAME2+\"[DOWN]\"+DCUR+\"[TAB]*[DOWN]\"+DEF_MKT+\"[DOWN]D\"+BDATE)\r\n\tELSE\r\n\t\tSend (SEQN+MNEM+\"            \"+REMK+\"[TAB]\"+NAME1+NAME2+\"[DOWN][DOWN][DOWN]\"+DEF_MKT+\"[DOWN]D\"+BDATE)\r\n\r\n\tENDIF\r\n\tSend (\"[ENTER]\")\r\n\r\n\r\n\r\n\tLOOP_SEQN3:\r\n\tIF (mid$(&screen,1761,8) = \"DC955092\") THEN\r\n\t\tSET GET_SEQN TO mid$(&screen,223,7)\r\n\t\tSET K TO 0\r\n\t\tSET LAST TO \"\"\r\n\r\n\t\tLOOP_SEQN1:\r\n\t\tIF (K=7) THEN\r\n\t\t\tMESSAGE(\"Sequence number has exceeded 7 digits\")\r\n\t\t\tGOTO END\r\n\t\tENDIF\r\n\r\n\r\n\t\tSET DIGIT TO val(mid$(GET_SEQN,7-K,1))+1\r\n\t\tIF (DIGIT<=9) THEN\r\n\t\t\tIF (K=0) THEN\r\n\t\t\t\tSET LAST TO DIGIT\r\n\t\t\t\tSET SEQN TO left$(GET_SEQN,6+K)+str$(LAST)\r\n\t\t\t\tGOTO ENDLOOPSEQN1\r\n\t\t\tELSE\r\n\t\t\t\tSET SEQN TO left$(GET_SEQN,6-K)+str$(DIGIT)+LAST\r\n\t\t\t\tGOTO ENDLOOPSEQN1\r\n\t\t\tENDIF\r\n\t\tELSE\r\n\t\t\tSET ENDDIGIT TO \"0\"\t\t\r\n\t\t\tSET LAST TO ENDDIGIT+LAST\r\n\t\t\tSET K TO K+1\r\n\r\n\t\t\tGOTO LOOP_SEQN1\r\n\t\tENDIF\r\n\t\tENDLOOPSEQN1:\r\n\t\tSEND (SEQN)\r\n\t\tGOTO LOOP_SEQN3\r\n\tENDIF\r\n\tENDLOOP_SEQN3:\r\n\r\n\r\n\tSend (\"C\")\r\n\t\r\n\tIF (DEF_MKT=\"HK\") AND (GEOG=\"048\") AND (DCUR=\"079\") THEN\r\n\t\tSend (\"[TAB][TAB]\"+DCUR+\"[TAB]*\")\r\n\tELSE\r\n\t\tSend (\"[ENTER]\")\r\n\tENDIF\r\n\r\n\tSend (NOSH+\"[ENTER]\")\r\n\tSend (\"C\")\r\n\tSend (CNAME+\"[DOWN]116[ENTER]\")\r\n\tSend (\"Y\")\r\n\r\n\tSend (\"[HOME]QFLAG\")\r\n\tSend ( MNEM )\r\n\tSend (\"YY\")\r\n\r\n\tSend (\"[HOME]MAJOR\")\r\n\tSend ( MNEM )\r\n\tSend (\"Y\")\r\n\tSend (\"S\")\r\n\tSend (\"Y\")\r\n\r\n\tSend (\"[PF3]\")\r\n\tSend (\"[TAB][TAB]\"+MNEM)\r\n\tSend (\"Y\")\r\n\tSend (\"S\")\r\n\tSend (\"S\")\r\n\tSend (\"Y\")\r\n\r\n\tSend (\"[PF3]\")\r\n\tSend (\"[PF3]\")\r\n\tSend (\"[HOME]CTREE\")\r\n\tSend ( MNEM+\"[DOWN]Y\" )\r\n\tSend (\"Y\")\r\n\r\n\r\n\tIF (DEF_MKT=\"HK\") OR (SECTY =\"ET\") OR (left$(REMK,1)<>\" \") THEN\r\n\t;==ONLY FOR H-SHARES, DUAL CURRENCY, ETF AND REMK\r\n\r\n\t\tIF (DEF_MKT=\"HK\") AND (GEOG=\"051\") AND (DCUR=\"079\") THEN\r\n\t\t\t;H-share whether ETF or not\r\n\t\t\t\tSend (\"[HOME]BASIC\") \r\n\t\t\t\tSend ( MNEM )\r\n\t\t\t\tSend ( GEOG+\"[DOWN][DOWN][DOWN]\"+REMK+REMK+\" [DOWN]\"+DEF_MKT+DCUR+\"[DOWN][DOWN][DOWN][RIGHT][RIGHT][RIGHT]\"+SECTY )\r\n\t\t\t\tSend (\"[Enter]\")\r\n\t\t\t\tSend (\"Y\")\r\n\t\t\t\tSend (\"[Enter]\")\r\n\t\t\t\tGOTO ENDBASICCHECK\r\n\t\tENDIF\r\n\t\tIF (DEF_MKT=\"HK\") AND (GEOG=\"048\") AND (DCUR=\"079\") THEN\r\n\t\t\t;Dual currency whether ETF or not\r\n\t\t\t\tSend (\"[HOME]BASIC\") \r\n\t\t\t\tSend ( MNEM )\r\n\t\t\t\tSend ( GEOG+\"[DOWN][DOWN][DOWN]\"+REMK+REMK+\" [DOWN]\"+DEF_MKT+DCUR+\"[TAB]\"+DCUR+\"[TAB]\"+DCUR+\"[DOWN][RIGHT][RIGHT][RIGHT]\"+SECTY )\r\n\t\t\t\tSend (\"[Enter]\")\r\n\t\t\t\tSend (\"Y\")\r\n\t\t\t\tSend (\"[Enter]\")\r\n\t\t\t\tGOTO ENDBASICCHECK\r\n\t\tENDIF\r\n\t\t\t;ETF or with REMK that are not HK stocks\r\n\t\r\n\t\tIF (SECTY =\"ET\") OR (left$(REMK,1)<>\" \") THEN\r\n\t\t\t\tSend (\"[HOME]BASIC\") \r\n\t\t\t\tSend ( MNEM )\r\n\t\t\t\tSend ( GEOG+\"[DOWN][DOWN][DOWN]\"+REMK+REMK+\" [DOWN]\"+DEF_MKT+DCUR+\"[DOWN][DOWN][DOWN][RIGHT][RIGHT][RIGHT]\"+SECTY )\r\n\t\t\t\tSend (\"[Enter]\")\r\n\t\t\t\tSend (\"Y\")\r\n\t\t\t\tSend (\"[Enter]\")\r\n\t\t\t\tGOTO ENDBASICCHECK\t\r\n\t\tENDIF\r\n\r\n\tENDIF\r\n\tENDBASICCHECK:\r\n\r\n\r\nENDIF\r\nENDER:\r\n\r\n\tIF left$(SEDOL,1)<>\" \" THEN\r\n\t\tSET PUT_SEDOL TO 'I'+SEDOL+'E'\r\n\tELSE\r\n\t\tSET PUT_SEDOL TO '[DOWN]'\r\n\tENDIF\r\n\r\n\tIF left$(ISIN,1)<>\" \" THEN\r\n\t\tSET PUT_ISIN TO 'I'+ISIN+'V E'\r\n\tELSE\r\n\t\tSET PUT_ISIN TO '[DOWN]'\r\n\tENDIF\r\n\r\n\tIF left$(CUSIP,1)<>\" \" THEN\r\n\t\tSET PUT_CUSIP TO 'I'+CUSIP+'E'\r\n\tELSE\t\r\n\t\tSET PUT_CUSIP TO '[DOWN]'\r\n\tENDIF\r\n\r\n\tIF left$(MNEM2,1)<>\" \" THEN\r\n\t\tSET PUT_MNEM2 TO 'I'+MNEM2+'E'\r\n\tELSE\r\n\t\tSET PUT_MNEM2 TO '[DOWN]'\r\n\tENDIF\r\n\r\n\tIF left$(CODE3,1)<>\" \" THEN\r\n\t\tSET PUT_CODE3 TO 'I'+CODE3+'M E'\r\n\tELSE\r\n\t\tSET PUT_CODE3 TO '[DOWN]'\r\n\tENDIF\r\n\r\n\tSET INDEX tO '[DOWN][DOWN]'+PUT_SEDOL+PUT_MNEM2+PUT_CUSIP+'[TAB][TAB][TAB][TAB][TAB][TAB]'+PUT_CODE3+PUT_ISIN\r\n\r\n\r\n\r\n\tSend (\"[HOME]INDEX\")\r\n\tSend ( DSCD )\r\n\tSend (INDEX)\r\n\tSend ( \"Y\" )\r\n\r\n;=================================================================================================================\r\n\r\nSend (\"[CLEAR]\")\r\nGOTO LOOP\r\nENDIF\r\nEND:\r\n\r\nLIST:\r\nDATA\r\n[DataStreamRicCreationInfoLine]\r\nENDDATA";
            }

            return template1;
        }
    }
}
