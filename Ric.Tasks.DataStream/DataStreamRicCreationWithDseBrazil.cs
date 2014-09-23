using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Text.RegularExpressions;
using Ric.Tasks.DataStream;
using HtmlAgilityPack;
using Ric.Util;
using System.Security.Cryptography.X509Certificates;

namespace Ric.Tasks.DataStream
{
    #region [Config]
    [ConfigStoredInDB]
    class DataStreamRicCreationWithDseBrazilConfig
    {
        [StoreInDB]
        [Category("Path")]
        [DefaultValue("D:\\DataStream\\RIC_Creation\\")]
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

        public DataStreamRicCreationWithDseBrazilConfig()
        {
            Date = DateTime.Today.AddDays(-1).ToString("MMdd");
        }
    }
    #endregion

    class DataStreamRicCreationWithDseBrazil : GeneratorBase
    {
        private static DataStreamRicCreationWithDseBrazilConfig configObj = null;
        private bool loginSuccess = false;
        private string jSessionId = string.Empty;
        private CookieContainer cookies = new CookieContainer();
        private Dictionary<string, string> namesAbbs = new Dictionary<string, string>();
        private List<string> listExchangeCode = new List<string>();

        protected override void Initialize()
        {
            configObj = Config as DataStreamRicCreationWithDseBrazilConfig;
        }

        protected override void Start()
        {
            LogMessage("start prepearing config...");

            if (!PrepearConfig())
                return;

            LogMessage("start get exchange code...");
            listExchangeCode = GetListExchangeCode();
            LogMessage(string.Format("exchange code count:{0}",
                listExchangeCode == null ? "0" : listExchangeCode.Count.ToString()));

            if (listExchangeCode == null || listExchangeCode.Count == 0)
                return;

            LogMessage("start get exchange code templates...");
            List<BrazilTemplateExchangeCode> listBrazilTemplateExchangeCode = FormatTemplatesExchangeCode(listExchangeCode);
            LogMessage(string.Format("exchange code template count:{0}",
                listBrazilTemplateExchangeCode == null ? "0" : listBrazilTemplateExchangeCode.Count.ToString()));

            if (listBrazilTemplateExchangeCode == null || listBrazilTemplateExchangeCode.Count == 0)
                return;

            LogMessage("start download rules...");
            DownloadNameRules();

            LogMessage("start format template by download rules ...");
            List<BrazilTemplate> listBrazilTemplate = FormatTemplates(listBrazilTemplateExchangeCode);
            LogMessage(string.Format("templates records count:{0}",
                listBrazilTemplate == null ? "0" : listBrazilTemplate.Count.ToString()));

            if (listBrazilTemplate == null || listBrazilTemplate.Count == 0)
                return;

            LogMessage("start generate output file ...");
            GenetateFile(listBrazilTemplate);
            LogMessage("task finished!");
        }

        #region [format templates exchange code]
        private List<BrazilTemplateExchangeCode> FormatTemplatesExchangeCode(List<string> listExchangeCode)
        {
            List<BrazilTemplateExchangeCode> list = new List<BrazilTemplateExchangeCode>();

            if (listExchangeCode == null || listExchangeCode.Count == 0)
                return list;

            try
            {
                foreach (var item in listExchangeCode)
                {
                    BrazilTemplateExchangeCode brazilTemplateExchangeCode = GetBrazilTemplateExchangeCode(item);

                    if (brazilTemplateExchangeCode == null)
                        continue;

                    if (string.IsNullOrEmpty(brazilTemplateExchangeCode.LotSize))
                        continue;

                    list.Add(brazilTemplateExchangeCode);
                }

                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private BrazilTemplateExchangeCode GetBrazilTemplateExchangeCode(string item)
        {
            BrazilTemplateExchangeCode brazilTemplateExchangeCode = new BrazilTemplateExchangeCode();
            string urlGet = @"http://www.bmfbovespa.com.br/Cias-Listadas/Empresas-Listadas/BuscaEmpresaListada.aspx?Nome={0}&idioma=en-us";//KROT3
            HtmlDocument htc = null;
            string pageSourceSecond = null;
            HtmlDocument htcThird = null;
            brazilTemplateExchangeCode.ExchangeCode = item;
            brazilTemplateExchangeCode.Url = string.Format(urlGet, item);

            try
            {
                RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
                {
                    htc = WebClientUtil.GetHtmlDocument(brazilTemplateExchangeCode.Url, 10000);
                });

                brazilTemplateExchangeCode.UrlSecond = GetSecondUrl(htc);
                brazilTemplateExchangeCode.FullName = GetFullName(htc);

                RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
                {
                    pageSourceSecond = WebClientUtil.GetPageSource(brazilTemplateExchangeCode.UrlSecond, 10000);
                });

                brazilTemplateExchangeCode.UrlThird = GetThirdUrl(pageSourceSecond);

                RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
                {
                    htcThird = WebClientUtil.GetHtmlDocument(brazilTemplateExchangeCode.UrlThird, 10000);
                });

                return GetBrazilTemplateExchangeCode(htcThird, brazilTemplateExchangeCode);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }

        }

        private string GetFullName(HtmlDocument htc)
        {
            string fullName = string.Empty;

            try
            {
                HtmlNodeCollection tr = htc.DocumentNode.SelectNodes(".//table")[0].SelectNodes(".//tr");

                if (tr == null || tr.Count <= 1)
                    return fullName;

                fullName = tr[1].SelectNodes(".//td")[0].SelectSingleNode(".//a").InnerText;

                if (string.IsNullOrEmpty(fullName))
                    return fullName;

                return fullName.Trim();
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private BrazilTemplateExchangeCode GetBrazilTemplateExchangeCode(HtmlDocument htc, BrazilTemplateExchangeCode brazilTemplateExchangeCode)
        {
            if (htc == null)
                return brazilTemplateExchangeCode;

            try
            {
                HtmlNodeCollection table = htc.DocumentNode.SelectNodes(".//table");
                brazilTemplateExchangeCode.TradingName = GetTradingName(table);
                brazilTemplateExchangeCode.ISIN = GetISIN(table);
                brazilTemplateExchangeCode.Type = GetExchangeCodeType(table);
                brazilTemplateExchangeCode.LotSize = GetLotSize(brazilTemplateExchangeCode);

                return brazilTemplateExchangeCode;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string GetLotSize(BrazilTemplateExchangeCode brazilTemplateExchangeCode)
        {
            string lotSize = string.Empty;           //"0001";
            string udsUrl = string.Empty;
            string targetPath = string.Empty;
            string findUdsLine = string.Empty;
            try
            {
                udsUrl = GetudsUrl();

                if (string.IsNullOrEmpty(udsUrl))
                {
                    LogMessage("no uds csv file today .");
                    return null;
                }

                targetPath = downloadUdsUrl(udsUrl);

                if (string.IsNullOrEmpty(targetPath))
                {
                    LogMessage("can not download uds csv file .");
                    return null;
                }

                findUdsLine = GetUdsLine(targetPath, brazilTemplateExchangeCode.TradingName);

                if (string.IsNullOrEmpty(findUdsLine))
                {
                    LogMessage("can not find Trading Name in csv file.");
                    brazilTemplateExchangeCode.Type = string.Empty;
                    return null;
                }


                if (findUdsLine.Contains("ON"))
                    brazilTemplateExchangeCode.Type = "ON";

                if (findUdsLine.Contains("PN"))
                    brazilTemplateExchangeCode.Type = "PN";

                if (findUdsLine.Contains("*"))
                    lotSize = "1000";
                else
                    lotSize = "0001";

                return lotSize;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string GetUdsLine(string targetPath, string tradingName)
        {
            string Udsline = string.Empty;
            string udsString = string.Empty;
            //string pattern = string.Format(@"{0}\s*(?<Type>(PN|ON)\s{0,2}\*?)", tradingName);
            string pattern = tradingName + @"\s*(?<Type>(PN|ON)\s{0,2}\*?)";

            try
            {
                if (string.IsNullOrEmpty(targetPath) && !File.Exists(targetPath))
                    return Udsline;

                udsString = File.ReadAllText(targetPath);

                Regex regex = new Regex(pattern);
                MatchCollection matches = regex.Matches(udsString);

                foreach (Match match in matches)
                {
                    Udsline = match.Groups["Type"].Value;
                    break;
                }

                return Udsline;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string downloadUdsUrl(string udsUrl)
        {
            string targetFolder = Path.Combine(Path.Combine(configObj.OutputPath, DateTime.Now.ToString("yyyy-MM-dd")), "DownloadUds");
            string targetPath = string.Empty;
            try
            {
                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                targetPath = Path.Combine(targetFolder, udsUrl.Substring(udsUrl.Length - 11, 11));

                if (File.Exists(targetPath))
                    return targetPath;

                RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
                {
                    WebClientUtil.DownloadFile(udsUrl, 10000, targetPath);
                });

                if (!File.Exists(targetPath))
                {
                    LogMessage(string.Format("download uds csv file error. url:{0}", udsUrl));
                    return string.Empty;
                }

                AddResult("DownloadUds", targetFolder, "DownloadUds");

                return targetPath;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }

        }

        private string GetudsUrl()
        {
            string url = @"http://dataops.datastream.com/cgi-bin/readdir.pl?dir=H:/Production/Data_Archive/Brazil/Valoriza/Equities/Price&resubmit=1&edit=1&reasons=archive&arc=1";
            string udsUrl = string.Empty;
            HtmlDocument htc = new HtmlDocument();
            string findUrl = string.Empty;

            try
            {
                RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
                {
                    htc = WebClientUtil.GetHtmlDocument(url, 10000);
                });

                HtmlNodeCollection trs = htc.DocumentNode.SelectNodes(".//table")[0].SelectNodes(".//tr");

                for (int i = 0; i < trs.Count; i++)
                {
                    try
                    {
                        findUrl = trs[i].SelectNodes(".//td")[0].SelectSingleNode(".//a[@href]").Attributes["href"].Value;

                        if (!string.IsNullOrEmpty(findUrl) && IsValidFindUrl(findUrl))
                        {
                            udsUrl = string.Format(@"http://dataops.datastream.com{0}", findUrl);
                            break;
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                return udsUrl;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private bool IsValidFindUrl(string findUrl)
        {
            bool result = false;
            string udsDate = DateTime.Now.ToString("ddMM");
            string pattern = "uds" + udsDate + @"\.csv";

            try
            {
                Match ma = (new Regex(pattern)).Match(findUrl);

                if (ma.Success)
                    result = true;

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return false;
            }
        }

        private string GetExchangeCodeType(HtmlNodeCollection table)
        {
            string type = string.Empty;
            string columnType = string.Empty;

            try
            {
                for (int i = 0; i < table.Count; i++)
                {
                    try
                    {
                        columnType = table[i].SelectNodes(".//thead")[0].SelectNodes(".//tr")[0].SelectNodes(".//th")[1].InnerText.Replace("\r\n", "").Trim();

                        if (columnType.Equals("Type") || columnType.Equals("Espec."))
                        {
                            type = table[i].SelectNodes(".//tr")[1].SelectNodes(".//td")[1].InnerText;
                            break;
                        }

                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                if (string.IsNullOrEmpty(type))
                    return type;

                return type.Replace("\r\n", "").Trim();
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string GetISIN(HtmlNodeCollection table)
        {
            string isin = string.Empty;

            try
            {
                isin = table[1].SelectNodes(".//tr")[1].SelectNodes(".//td")[1].InnerText;

                if (string.IsNullOrEmpty(isin))
                    return isin;

                return isin.Replace("\r\n", "").Trim();
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string GetTradingName(HtmlNodeCollection table)
        {
            string tradingName = string.Empty;

            try
            {
                tradingName = table[0].SelectNodes(".//tr")[0].SelectNodes(".//td")[1].InnerText;

                if (string.IsNullOrEmpty(tradingName))
                    return tradingName;

                return tradingName.Replace("\r\n", "").Trim();
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string GetThirdUrl(string pageSource)
        {
            string url = string.Empty;
            string pattern = "ctl00_contentPlaceHolderConteudo_iframeCarregadorPaginaExterna\"\\s*src\\=\"\\.\\.\\/\\.\\.\\/(?<url>[^\"]+)\\#a\"";

            if (string.IsNullOrEmpty(pageSource))
                return url;

            try
            {
                Match ma = (new Regex(pattern)).Match(pageSource);

                if (ma.Success)
                    url = string.Format(@"http://www.bmfbovespa.com.br/{0}", ma.Groups["url"].Value.Replace("amp;", ""));

                return url;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string GetSecondUrl(HtmlDocument htc)
        {
            string url = string.Empty;

            if (htc == null)
                return url;

            try
            {
                HtmlNodeCollection tr = htc.DocumentNode.SelectNodes(".//table")[0].SelectNodes(".//tr");

                if (tr == null || tr.Count <= 1)
                    return url;

                url = tr[1].SelectNodes(".//td")[0].SelectSingleNode(".//a[@href]").Attributes["href"].Value;

                return string.Format(@"http://www.bmfbovespa.com.br/Cias-Listadas/Empresas-Listadas/{0}", url);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private List<string> GetListExchangeCode()
        {
            List<string> listExchangeCode = null;
            string urlListingStatus = @"http://dataops.datastream.com/cgi-bin/readfile.pl?filename=H:/Production/Loaders/Brazil/Valoriza/Equities/Non-Price/Mload/brazil.rep&warnold=1";
            string pageSource = string.Empty;
            string[] remkListingData;

            try
            {
                RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
                {
                    pageSource = WebClientUtil.GetPageSource(null, urlListingStatus, 180000, null, Encoding.GetEncoding("ISO-8859-1"));
                });

                if (pageSource.Contains("Listing Status"))
                {
                    remkListingData = pageSource.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                    listExchangeCode = GetExchangeCode(remkListingData);
                }

                return listExchangeCode;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                       System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                       System.Reflection.MethodBase.GetCurrentMethod().Name,
                       ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }
        #endregion

        #region [generate bulk file]
        private void GenetateFile(List<BrazilTemplate> listBrazilTemplate)
        {
            if (listBrazilTemplate == null || listBrazilTemplate.Count == 0)
            {
                Logger.Log("listBrazilTemplate.count==0, no data need to ouput. ", Logger.LogType.Warning);
                return;
            }

            var sb = new StringBuilder();
            try
            {
                string fileFolder = Path.Combine(configObj.OutputPath, DateTime.Today.ToString("yyyy-MM-dd"));
                string filePath = Path.Combine(fileFolder, string.Format("Brazil_{0}.MAC", configObj.Date));

                foreach (var item in listBrazilTemplate)
                {
                    sb.AppendFormat("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",\"{12}\",\"{13}\",\"{14}\",\"{15}\",\"{16}\",\"{17}\",\"{18}\",\"{19}\"",
                                     item.Mnem, item.Sedol, item.Name1, item.Name2,
                                     item.Bdate, item.DefMkt, item.Mnem2, item.Isin, item.Cname,
                                     item.Geog, item.Dcur, item.Remk, item.Secty, item.Grpi, item.Fname1, item.Mic,
                                     item.Fname2, item.Cusip, item.Qci, item.ParentRic);
                    sb.AppendLine();
                }

                if (!Directory.Exists(fileFolder))
                    Directory.CreateDirectory(fileFolder);

                string templateString = File.ReadAllText(@"Config\DataStream\Brazil.MAC", Encoding.ASCII);
                string resultString = string.Format(templateString, sb.ToString());
                File.WriteAllText(filePath, resultString, Encoding.ASCII);
                AddResult("Output Folder", fileFolder, "Output Folder");
                AddResult("MAC File", filePath, "MAC File");
                Logger.Log("Generate normal RIC creation output file...OK!");
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }
        #endregion

        #region download rules
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

            Logger.Log(string.Format("download the rules count:{0}", namesAbbs == null ? "0" : namesAbbs.Keys.Count.ToString()));
        }
        #endregion

        #region format template
        private List<BrazilTemplate> FormatTemplates(List<BrazilTemplateExchangeCode> listBrazilTemplateExchangeCode)
        {
            List<BrazilTemplate> list = new List<BrazilTemplate>();
            string sedol = string.Empty;
            string name1 = string.Empty;
            string name2 = string.Empty;
            string suffix = string.Empty;

            try
            {
                if (listBrazilTemplateExchangeCode == null || listBrazilTemplateExchangeCode.Count == 0)
                    return list;

                foreach (var item in listBrazilTemplateExchangeCode)
                {
                    BrazilTemplate template = new BrazilTemplate(item);

                    //Sedol 
                    sedol = GetSedol(template.Isin);

                    if (string.IsNullOrEmpty(sedol))
                        sedol = "         ";//9 space
                    else
                        sedol = "UK" + sedol;

                    template.Sedol = sedol;

                    //Cname 
                    //FormatBulkNames(item.CompanyName.Trim(), 24, 24, nameInputType.Name,ref suffix);
                    //template.NameSuffix = suffix;
                    //FormatCName(template, item.FullName);
                    if (string.IsNullOrEmpty(item.FullName))
                        continue;

                    if (item.FullName.Length > 24)
                        template.Cname = item.FullName.Substring(0, 24);
                    else
                        template.Cname = item.FullName.PadRight(24, ' ');

                    list.Add(template);
                }
                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private void FormatBulkNames(string companyName, int nameLength1, int nameLength2, nameInputType type, ref string suffix)
        {
            string temp = string.Empty;
            if (string.IsNullOrEmpty(companyName))
            {
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
                        char middleChar = charArr[index];
                        if (middleChar == ' ' || middleChar == '.')
                        {
                            index--;
                        }
                        while (true)
                        {
                            middleChar = charArr[index];
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
        }

        private List<string> GetExchangeCode(string[] remkListingData)
        {
            List<string> listExchangeCode = new List<string>();
            List<int> listIndexExchangeCode = GetlistIndexExchangeCode(remkListingData);
            string pattern = @"\s{12}(?<ExchangeCode>[A-Z0-9]{5})\s+";
            string exchangeCode = string.Empty;

            try
            {
                foreach (var item in listIndexExchangeCode)
                {
                    Match ma = (new Regex(pattern)).Match(remkListingData[item]);

                    if (!ma.Success)
                        continue;

                    exchangeCode = ma.Groups["ExchangeCode"].Value.Trim();

                    if (string.IsNullOrEmpty(exchangeCode) || listExchangeCode.Contains(exchangeCode))
                        continue;

                    listExchangeCode.Add(exchangeCode);
                }

                return listExchangeCode;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private List<int> GetlistIndexExchangeCode(string[] remkListingData)
        {
            List<int> listIndexExchangeCode = new List<int>();
            bool isUnderListingStatus = false;

            if (remkListingData == null || remkListingData.Length <= 0)
                return listIndexExchangeCode;

            try
            {
                for (int i = 0; i < remkListingData.Length - 1; i++)
                {
                    if (remkListingData[i].Contains("Listing Status") || isUnderListingStatus)
                    {
                        listIndexExchangeCode.Add(i + 1);
                        isUnderListingStatus = true;
                    }

                    if (remkListingData[i].Contains("Page:") && isUnderListingStatus)
                        break;
                }

                return listIndexExchangeCode;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private void FormatCName(BrazilTemplate bulkTemp, string names)
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

            bulkTemp.Cname = names.PadRight(24, ' ');
        }

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
                    else if (ending.Equals("INCORPORATED"))
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

        private string GetSedol(string isin)
        {
            string result = string.Empty;
            try
            {
                if ((isin + "").Trim().Length == 0)
                    return result;

                if (!loginSuccess)
                {
                    LoginToWebsite();
                    loginSuccess = true;
                }

                string response = QuerySedol(isin.Trim());
                result = ParseSedol(response, isin.Trim());

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return result;
            }
        }

        private string ParseSedol(string response, string isin)
        {
            string result = string.Empty;
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

            return result;
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

        private void LoginToWebsite()
        {
            string username = Encode(configObj.Username);
            string uri = string.Format("https://www.unavista.londonstockexchange.com/datasolutions/login.html");
            string postData = string.Format("callCount=1\r\nwindowName=unavista_datasolutions\r\nc0-scriptName=LoginHelper\r\nc0-methodName=doLogin\r\nc0-id=0\r\nc0-param0=string:{0}\r\nc0-param1=null:null\r\nc0-param2=string:{1}\r\nc0-param3=boolean:false\r\nc0-param4=string:Google%20Chrome\r\nc0-param5=null:null\r\nc0-param6=boolean:false\r\nc0-param7=null:null\r\nc0-param8=null:null\r\nc0-param9=string:11.9.0.23449\r\nbatchId=2\r\npage=%2Fdatasolutions%2Flogin.html\r\nhttpSessionId=\r\nscriptSessionId=", username, configObj.Password);
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

        public static string Encode(string encodeText)
        {
            return Regex.Replace(encodeText, "[^a-zA-Z0-9]", delegate(Match match) { return "%" + BitConverter.ToString(Encoding.GetEncoding("euc-kr").GetBytes(match.Value)).Replace("-", "%"); });
        }

        private static bool CheckValidationResult(object senter, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }

        private bool IsValidBrazilTemplate(DataStreamRicCreationEntity item)
        {
            bool result = true;
            try
            {
                if ((item.ExchangeCode + "").Trim().Equals("RTS"))
                    result = false;

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return false;
            }
        }
        #endregion


        #region [format records]
        private List<DataStreamRicCreationEntity> FormatRecords(List<string> listExtractedRecords)
        {
            List<DataStreamRicCreationEntity> list = new List<DataStreamRicCreationEntity>();
            try
            {
                if (listExtractedRecords == null || listExtractedRecords.Count == 0)
                    return null;

                foreach (var record in listExtractedRecords)
                {
                    DataStreamRicCreationEntity dsInfo = new DataStreamRicCreationEntity(record);

                    if (IsValidDataStreamRicCreationInfo(dsInfo))
                        list.Add(dsInfo);
                }

                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private bool IsValidDataStreamRicCreationInfo(DataStreamRicCreationEntity dsInfo)
        {
            try
            {
                bool result = true;

                if (dsInfo == null)
                    result = false;

                if (dsInfo.ThomsonReutersClassificationScheme.Trim().Equals("RTS"))
                    result = false;

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return false;
            }
        }
        #endregion

        #region [extract records]
        private List<string> ExtractRecordsFromFiles(List<string> listDownloadFtpFilePath)
        {
            List<string> list = new List<string>();
            try
            {
                if (listDownloadFtpFilePath == null || listDownloadFtpFilePath.Count == 0)
                {
                    string msg = string.Format("no ftp file need to extract .");
                    Logger.Log(msg, Logger.LogType.Warning);
                    return null;
                }

                foreach (var item in listDownloadFtpFilePath)
                {
                    ExtractRecord(item, list);
                }
                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private void ExtractRecord(string item, List<string> list)
        {
            try
            {
                if (!File.Exists(item))
                    return;

                int start = item.LastIndexOf("\\");
                string fileName = item.Substring(start + 1, item.Length - start - 1);

                using (StreamReader sr = new StreamReader(item))
                {
                    string line = null;
                    string exchangeCode = string.Empty;

                    while ((line = sr.ReadLine().ToUpper()) != null)
                    {
                        if (!line.StartsWith("XE"))
                            continue;

                        if (line.Length < 1031)
                            continue;

                        list.Add(line + fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }
        #endregion

        #region [download files]
        private List<string> DownloadFilesFromFtp(List<string> listDownloadFtpCode)
        {
            List<string> list = new List<string>();

            try
            {
                string fileName = string.Empty;
                string fileUrl = string.Empty;
                string targetFolder = Path.Combine(
                    Path.Combine(configObj.OutputPath, DateTime.Now.ToString("yyyy-MM-dd")),
                    "DSE_FILES");
                string targetPath = string.Empty;

                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                AddResult("DSE_FILES", targetFolder, "DSE_FILES");

                foreach (var item in listDownloadFtpCode)
                {
                    fileName = string.Format("{0}{1}.M", item, configObj.Date);
                    fileUrl = @"ftp://ASIA2:ASIA2@ds1.rds.reuters.com/" + fileName;
                    targetPath = Path.Combine(targetFolder, fileName);

                    if (!File.Exists(targetPath))
                        DownloadFileFromFtp(targetPath, fileUrl);

                    if (File.Exists(targetPath))
                        list.Add(targetPath);
                }
                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private void DownloadFileFromFtp(string targetPath, string fileUrl)
        {
            try
            {
                WebClient request = new WebClient();
                WebProxy proxy = new WebProxy(configObj.IP, Convert.ToInt32(configObj.Port));
                request.Proxy = proxy;
                request.Credentials = new NetworkCredential("ASIA2", "ASIA2");
                request.DownloadFile(fileUrl, targetPath);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                LogMessage(string.Format("Can not download file: {0}", targetPath));
            }
        }
        #endregion


        #region [check fonfig]
        private bool PrepearConfig()
        {
            try
            {
                bool result = true;

                if ((configObj.OutputPath + "").Trim().Length == 0)
                    return false;

                if (!Directory.Exists(configObj.OutputPath))
                    Directory.CreateDirectory(configObj.OutputPath);

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return false;
            }
        }
        #endregion
    }

    #region [entity]
    public class BrazilTemplateExchangeCode
    {
        public string ExchangeCode { get; set; }
        public string FullName { get; set; }
        public string LotSize { get; set; }
        public string ISIN { get; set; }
        public string SEDOL { get; set; }
        public string TradingName { get; set; }
        public string Type { get; set; }
        public string Url { get; set; }
        public string UrlSecond { get; set; }
        public string UrlThird { get; set; }
    }

    public class BrazilTemplate
    {
        public string Mnem { get; set; }
        public string Sedol { get; set; }
        public string Name1 { get; set; }
        public string Name2 { get; set; }
        public string Bdate { get; set; }
        public string DefMkt { get; set; }
        public string Mnem2 { get; set; }
        public string Isin { get; set; }
        public string Cname { get; set; }
        public string Geog { get; set; }
        public string Dcur { get; set; }
        public string Remk { get; set; }
        public string Secty { get; set; }
        public string Grpi { get; set; }
        public string Fname1 { get; set; }
        public string Mic { get; set; }
        public string Fname2 { get; set; }
        public string Cusip { get; set; }
        public string Qci { get; set; }
        public string ParentRic { get; set; }
        public string NameSuffix { get; set; }

        public BrazilTemplate(BrazilTemplateExchangeCode brazilTemplate)
        {
            //Mnem    
            if ((brazilTemplate.ExchangeCode + "").Trim().Length >= 3)
            {
                int exchangeCodeLength = brazilTemplate.ExchangeCode.Trim().Length;
                this.Mnem = "BR:" + brazilTemplate.ExchangeCode.Trim().Substring(0, 2) + brazilTemplate.ExchangeCode.Trim().Substring(exchangeCodeLength - 1, 1);
            }
            else
            {
                this.Mnem = "      ";
            }

            //Sedol 

            //Name1 
            //Name
            Name1 = Name2 = string.Format("{0} {1}", brazilTemplate.TradingName, brazilTemplate.Type).PadRight(24, ' ');

            //Bdate 
            this.Bdate = "        ";//8

            //DefMkt 
            this.DefMkt = "SP";

            //Mnem2 
            this.Mnem2 = "            ";//12

            //Isin 
            if ((brazilTemplate.ISIN + "").Trim().Length > 0)
                this.Isin = brazilTemplate.ISIN;
            else
                this.Isin = "            ";//12

            //Cname 

            //Geog
            this.Geog = "062";

            //Dcur 
            this.Dcur = "058";

            //Remk 
            if (brazilTemplate.LotSize.Equals("0001"))
                Remk = "1";
            else if (brazilTemplate.LotSize.Equals("1000"))
                Remk = "1000";

            //Secty 
            //Grpi 
            this.Secty = "EQ";
            this.Grpi = "116";

            //Fname1 
            this.Fname1 = "                   ";//19

            //Mic
            this.Mic = "   ";//3

            //Fname2
            this.Fname2 = "                        ";//24

            //Cusip
            this.Cusip = ("BR:" + brazilTemplate.ExchangeCode).PadRight(12, ' ');

            //Qci 
            this.Qci = "  ";//2

            //ParentRic
            this.ParentRic = "";
        }
    }
    #endregion
}
