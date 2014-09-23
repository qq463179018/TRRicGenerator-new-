using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;
using Ric.Db.Manager;
using Ric.Db.Config;
using MySql.Data.MySqlClient;
using System.Net;
using System.IO;
using HtmlAgilityPack;
using Ric.Util;
using Microsoft.Office.Interop.Excel;
using pdftron;
using pdftron.PDF;
using System.Text.RegularExpressions;
using Microsoft.International.Converters.PinYinConverter;
using System.Collections.ObjectModel;
using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;
using Word = Microsoft.Office.Interop.Word; 
using System.Reflection;
//using ipo.data;

namespace Ric.Tasks.China
{
    public class ChinaIPOInfo
    {
        public string UpdateDate { get; set; }
        public string EffectiveDate { get; set; }
        public string RIC { get; set; }
        public string FM { get; set; }
        public string IDNName { get; set; }
        public string OfficialCode { get; set; }
        public string COI_DSPLY { get; set; }
        public string BCAST_REF { get; set; }
        public string LONGLINK { get; set; }
        public string BKGD_REF { get; set; }
        public string Type { get; set; }
        public string ChainRic { get; set; }
        public string EnglishName { get; set; }
        public string ListingShare { get; set; }
        public string IPOPrice { get; set; }

    }

    public class BondAddInfo
    {
        public string UpdatedOn { get; set; }
        public string EffectiveDate { get; set; }
        public string Ric { get; set; }
        public string IDNDisplayName { get; set; }
        public string OfficialCode { get; set; }
        public string Matur_Date { get; set; }
        public string Cope_Rate { get; set; }
        public string Coupn_Date { get; set; }
        public string Issue_Date { get; set; }
        public string BondType { get; set; }
        public string CoiDisplayNmll { get; set; }
        public string BKGD_Ref { get; set; }
        public string IssueAmt { get; set; }
        public string PayFreq { get; set; }
        public string ProvSymb { get; set; }
    }

    [ConfigStoredInDB]
    public class ChinaFMConfig
    {
        [Category("Date")]
        [DisplayName("Date")]
        [Description("Date format: yyyyMMdd. E.g. 2014-12-06")]
        public string Date { get; set; }

        [StoreInDB]
        [Category("TargetPath")]
        [DisplayName("Output path")]
        public string OutputPath { get; set; }

        [StoreInDB]
        [Category("TemplateFilePath")]
        [DisplayName("Template file path")]
        public string TemplateFilePath { get; set; }

        public ChinaFMConfig()
        {
            Date = DateTime.Now.ToString("yyyy-MM-dd");
        }
    }
    //KZZ:可转债
    //JRZ:金融债
    //GSZ:公司债
    class ChinaFM : GeneratorBase 
    {
        private ExtendDataContext extendContext = null;
        private List<ChinaIPOInfo> IPOinfo = null;
        private List<BondAddInfo> BondADDInfo = null;
        private ChinaFMConfig ConfigObj = null;
        private CookieContainer cookies = new CookieContainer();
        private CookieContainer cookiesSH = new CookieContainer();
        private List<string> PDFPath = new List<string>();
        private List<string> SZtmpFilePath = new List<string>();
        private List<string> KZZpdfPath = new List<string>();
        private List<string> KZZtmpFilePath = new List<string>();
        private List<string> JRZWordPath = new List<string>();
        private List<string> JRZtmpFilePath = new List<string>();
        private List<string> GSZpdfPath = new List<string>();
        private List<string> GSZtmpFilePath = new List<string>();
        private ExcelApp app;
        private Word.Application WordApp = new Word.Application();

        protected override void Initialize()
        {
            ConfigObj = Config as ChinaFMConfig;
            extendContext = new ExtendDataContext(new MySqlConnection(AceDbConfig.DataSourceStringDeals));
            IPOinfo = new List<ChinaIPOInfo>();
            BondADDInfo = new List<BondAddInfo>();
            app = new ExcelApp(false, false);
        }

        protected override void Start()
        {
            GetIPOIDfromDBbyDate();
            GetPDFfromSzWebSite();
            DownloadPDFforSZ();
            AnalyseSZpdf();
            GetPDFfromShangHaiWebSite();
            DownloadPDFforShanghai();
            AnalyseInfoForShanghai();
            GenerateFMFile();
        }


        
        private void GetGSZpdf()
        {
            try
            {
                string st = WebClientUtil.GetPageSource(@"http://www.sse.com.cn/disclosure/bond/corporate/s_latest.htm", 300000);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(st);
                HtmlNode node = doc.DocumentNode.SelectSingleNode(@".//ul[@class='list_ul']");
                HtmlNodeCollection ss = node.SelectNodes(@".//li");
                foreach (HtmlNode item in ss)
                {
                    try
                    {
                        HtmlNode dateNode = item.SelectSingleNode(".//span");
                        string date = null;
                        if (dateNode != null)
                        {
                            date = dateNode.InnerText;
                        }
                        if (!string.IsNullOrEmpty(date) && date.Equals(ConfigObj.Date))
                        {
                            HtmlNode a = item.SelectSingleNode(@".//a");
                            string title = a.InnerText;
                            string url = a.Attributes["href"].Value.ToString();
                            if (title.Contains("上市公告书"))
                            {
                                url = @"http://www.sse.com.cn/" + url;
                                GSZpdfPath.Add(url);
                                string tmpfilepath = ConfigObj.OutputPath + "\\" + title + ".pdf";
                                GSZtmpFilePath.Add(tmpfilepath);
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        string msg = string.Format("Get GSZ PDF for shang hai error. ex:{0}", ex.ToString());
                        Logger.Log(msg, Logger.LogType.Error);
                        continue;
                    }
                    
                }
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get GSZ PDF for shang hai error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void DownloadPDFforGSZ()
        {
            try
            {
                int index = 0;
                foreach (string path in GSZpdfPath)
                {
                    try
                    {
                        string url = path;
                        HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                        request.Timeout = 300000;
                        request.Method = "GET";
                        request.CookieContainer = cookies;
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        Stream content = response.GetResponseStream();
                        string outPutfile = GSZtmpFilePath[index];
                        index++;
                        using (Stream file = File.Create(outPutfile))
                        {

                            byte[] buffer = new byte[8 * 1024];
                            int len;
                            int offset = 0;

                            while ((len = content.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                file.Write(buffer, 0, len);
                                offset += len;
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        string msg = string.Format("Get GSZ PDF for shang hai error. ex:{0}", ex.ToString());
                        Logger.Log(msg, Logger.LogType.Error);
                        continue;
                    }
                    
                }
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get GSZ PDF for shang hai error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GetKZZpdf()
        {
            try
            {
                //得到可转债列表
                string st = WebClientUtil.GetPageSource(@"http://www.sse.com.cn/js/common/detachableConvertibleBulletin/convertiblebulletin_new.js", 300000);
                string[] split = new string[] { "_t.push" };
                split = st.Split(split, StringSplitOptions.None);
                foreach (string item in split)
                {
                    string[] info1 = item.Split(new char[] { '"' });
                    if (item.Contains(DateTime.Now.ToString("yyyy-MM-dd")))
                    {
                        if (item.Contains(@"上市公告书"))
                        {
                            string[] info = item.Split(new char[] { '"' });
                            if (info.Length > 6)
                            {
                                KZZpdfPath.Add(@"http://www.sse.com.cn/" + info[5]);
                                string tmpfilepath = ConfigObj.OutputPath + "\\"  + info[1] + ".pdf";
                                KZZtmpFilePath.Add(tmpfilepath);
                            }

                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get KZZ PDF for shang hai error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void DownloadPDFforKZZ()
        {
            try
            {
                int index = 0;
                foreach (string path in KZZpdfPath)
                {
                    string url = path;
                    HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                    request.Timeout = 300000;
                    request.Method = "GET";
                    request.CookieContainer = cookies;
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    Stream content = response.GetResponseStream();
                    string outPutfile = KZZtmpFilePath[index];
                    index++;
                    using (Stream file = File.Create(outPutfile))
                    {

                        byte[] buffer = new byte[8 * 1024];
                        int len;
                        int offset = 0;

                        while ((len = content.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            file.Write(buffer, 0, len);
                            offset += len;
                        }
                    }

                }
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Download KZZ PDF error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GetPDFfromShangHaiWebSite()
        {
            GetKZZpdf();
            GetJRZWord();
            GetGSZpdf();
        }

        private void DownloadPDFforShanghai()
        {
            DownloadWordforJRZ();
            DownloadPDFforGSZ();
            DownloadPDFforKZZ();
        }

        private void AnalyseInfoForShanghai()
        {
            AnalyseJRZword();
            AnalyseKZZpdf();
            AnalyseGSZpdf();
        }

        private void GetJRZWord()
        {
            try
            {
                string st = WebClientUtil.GetPageSource(@"http://www.sse.com.cn/disclosure/bond/financial/", 300000);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(st);
                HtmlNode node = doc.DocumentNode.SelectSingleNode(@".//ul[@class='list_ul']");
                HtmlNodeCollection ss = node.SelectNodes(@".//li");
                foreach (HtmlNode item in ss)
                {
                    
                    string date = item.SelectSingleNode(@".//span").InnerText;
                    if (!string.IsNullOrEmpty(date) && date.Equals(ConfigObj.Date))
                    {
                        string title = item.SelectSingleNode(@".//a").InnerText;
                        string url = item.SelectSingleNode(@".//a").Attributes["href"].Value.ToString();
                        if (!title.Contains("增发") && title.Contains("发行情况公告"))
                        {
                            url = @"http://www.sse.com.cn/" + url;
                            JRZWordPath.Add(url);
                            string tmpfilepath = ConfigObj.OutputPath + "\\" + title + ".doc";
                            JRZtmpFilePath.Add(tmpfilepath);
                        }
                    }

                    
                }
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get JRZ WORD for shang hai error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }
        
        private void DownloadWordforJRZ()
        {
            try
            {
                int index = 0;
                foreach (string path in JRZWordPath)
                {
                    string url = path;
                    HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                    request.Timeout = 300000;
                    request.Method = "GET";
                    request.CookieContainer = cookies;
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    Stream content = response.GetResponseStream();
                    string outPutfile = JRZtmpFilePath[index];
                    index++;
                    using (Stream file = File.Create(outPutfile))
                    {

                        byte[] buffer = new byte[8 * 1024];
                        int len;
                        int offset = 0;

                        while ((len = content.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            file.Write(buffer, 0, len);
                            offset += len;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Download word for JRZ error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
                
        }

        private void GetPDFfromSzWebSite()
        {
            try
            {
                //登陆
                string url = @"http://www.szse.cn/";
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 300000;
                request.Method = "GET";
                request.CookieContainer = cookies;
                request.Accept = @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
                request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
                request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
                request.Headers["Cache-Control"] = @"max-age=0";
                request.Host = @"www.szse.cn";
                request.KeepAlive = true;
                request.Referer = @"http://www.szse.cn/";
                request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader sr = new StreamReader(response.GetResponseStream());
                string st = sr.ReadToEnd();

                //消息披露
                url = @"http://www.szse.cn/main/disclosure/";
                request = WebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 300000;
                request.Method = "GET";
                request.CookieContainer = cookies;
                request.Accept = @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
                request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
                request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
                request.Host = @"www.szse.cn";
                request.KeepAlive = true;
                request.Referer = @"http://www.szse.cn/";
                request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
                response = (HttpWebResponse)request.GetResponse();
                sr = new StreamReader(response.GetResponseStream());
                st = sr.ReadToEnd();

                //债券信息-债券公告
                url = @"http://disclosure.szse.cn/m/zqgg.htm";
                request = WebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 300000;
                request.Method = "GET";
                request.CookieContainer = cookies;
                response = (HttpWebResponse)request.GetResponse();
                sr = new StreamReader(response.GetResponseStream());
                st = sr.ReadToEnd();

                url = @"http://disclosure.szse.cn//disclosure/fulltext/plate/szbondlatest_1m.js?ver=" + DateTime.Now.ToString("yyyyMMddHHmm");
                request = WebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 300000;
                request.Method = "GET";
                request.CookieContainer = cookies;
                request.Accept = @"*/*";
                request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
                request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
                request.Host = @"disclosure.szse.cn";
                request.KeepAlive = true;
                request.Referer = @"http://disclosure.szse.cn/m/unit/xxpllist.html?s=%2Fdisclosure%2Ffulltext%2Fplate%2Fszbondlatest_1m.js";
                request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
                response = (HttpWebResponse)request.GetResponse();
                sr = new StreamReader(response.GetResponseStream());
                st = sr.ReadToEnd();

                url = @"http://disclosure.szse.cn/m/search_wl.jsp?stockType=04";
                string postData = @"leftid=1&lmid=drgg&stockCode=&keyword=%C9%CF%CA%D0%B9%AB%B8%E6%CA%E9&startTime=" + ConfigObj.Date + "&endTime=" + ConfigObj.Date + "&imageField.x=-863&imageField.y=-199";
                request = WebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 300000;
                request.Method = "POST";
                request.CookieContainer = cookies;
                request.Accept = @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
                request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
                request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
                request.Headers["Cache-Control"] = @"max-age=0";
                request.ContentType = @"application/x-www-form-urlencoded";
                request.Referer = @"http://disclosure.szse.cn/m/zqgg.htm";
                request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
                request.KeepAlive = true;
                request.Host = @"disclosure.szse.cn";
                request.Headers["Origin"] = @"http://disclosure.szse.cn";
                byte[] buf = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);
                response = (HttpWebResponse)request.GetResponse();

                sr = new StreamReader(response.GetResponseStream());
                st = sr.ReadToEnd();

                HtmlDocument html = new HtmlDocument();
                html.LoadHtml(st);
                HtmlNodeCollection nodes = html.DocumentNode.SelectNodes("//a");
                foreach(HtmlNode node in nodes)
                {
                    string href = node.Attributes["href"].Value.Trim();
                    if (href.EndsWith(".PDF"))
                    {
                        PDFPath.Add(href);
                    }
                    
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("Get PDF from shenzhen web site error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void DownloadPDFforSZ()
        {
            try
            {
                foreach (string path in PDFPath)
                {
                    string url = @"http://disclosure.szse.cn/m/" + path;
                    HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                    request.Timeout = 300000;
                    request.Method = "GET";
                    request.CookieContainer = cookies;
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    Stream content = response.GetResponseStream();
                   string outPutfile = ConfigObj.OutputPath + "\\" + path.Substring(path.LastIndexOf('/') + 1);
                   using (Stream file = File.Create(outPutfile))
                   {

                       byte[] buffer = new byte[8 * 1024];
                       int len;
                       int offset = 0;

                       while ((len = content.Read(buffer, 0, buffer.Length)) > 0)
                       {
                           file.Write(buffer, 0, len);
                           offset += len;
                       }
                   }
                   SZtmpFilePath.Add(outPutfile);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("Download PDF for SZ error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private List<PdfString> FindListPdfString(string pattern, string path)
        {
            PDFDoc doc = null;

            try
            {
                if (!File.Exists(path))
                {
                    string msg = string.Format("the file {0} is not exist.", path);
                    Logger.Log(msg, Logger.LogType.Error);
                    return null;
                }

                //PDFNet.Initialize();
                PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(path);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", path);
                    Logger.Log(msg, Logger.LogType.Error);
                    return null;
                }

                PdfAnalyzer pa = new PdfAnalyzer();
                return pa.RegexSearchAllPages(doc, pattern);
            }
            catch (Exception ex)
            {
                string msg = string.Format("Find list PDF string error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
            finally
            {
                doc.Dispose();
            }
        }

        private List<string> FindValueFromListPdfString(string pattern, List<PdfString> listPdfString, List<string> groupName)
        {
            List<string> result = new List<string>();

            try
            {
                foreach (var item in listPdfString)
                {
                    Match match = (new Regex(pattern)).Match(item.ToString().Replace(" ", ""));

                    if (match.Success)
                    {
                        foreach (var name in groupName)
                            result.Add(match.Groups[name].Value);

                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("Find value from list pdf string error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }

            return null;
        }

        private string GetSZEffectiveDate(string path)
        {
            try
            {
                string patternPdfExDate = @".{0,10}(?<RecordDate>\d\s*\d\s*\d?\s*\d?\s*年\s*\d\s*\d?\s*月\s*\d\s*\d?\s*日).{0,30}";
                List<PdfString> list = FindListPdfString(patternPdfExDate, path);
                List<string> groupName = new List<string>() { "year", "month", "day" };
                string patternExDate = @"上市时间\D*(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日";
                List<string> effectiveDate = FindValueFromListPdfString(patternExDate, list, groupName);
                if (effectiveDate == null || effectiveDate.Count == 0)
                {
                    return string.Empty;
                }
                string year = effectiveDate[0].Trim().Length == 2 ? "20" + effectiveDate[0].Trim() : effectiveDate[0].Trim();//year
                string month = effectiveDate[1].Trim().Length == 1 ? "0" + effectiveDate[1].Trim() : effectiveDate[1].Trim();//month
                string day = effectiveDate[2].Trim().Length == 1 ? "0" + effectiveDate[2].Trim() : effectiveDate[2].Trim();//month
                string date = year + month + day;
                date = DateTime.ParseExact(date, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd-MMM-yyyy");
                return date;
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get Shenzhen effective date error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return string.Empty;
            }
        }

        private string GetSZRic(string path)
        {
            try
            {
                string patternPdfRic = @".{15,20}(?<Ric>\d+\s*)";
                List<PdfString> list = FindListPdfString(patternPdfRic, path);
                List<string> groupName = new List<string>() { "Ric" };
                string patternRic = @"证券代码\D*(?<Ric>\d+\s*)";
                List<string> Ric = FindValueFromListPdfString(patternRic, list, groupName);
                if (Ric == null || Ric.Count == 0)
                {
                    return string.Empty;
                }
                return Ric[0];
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get Shenzhen RIC error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return string.Empty;
            }
        }

        private string GetSZIssueAmt(string path)
        {
            try
            {
                string patternPdfIssueAmt = @".{15,20}(?<IssueAmt>\d+\s*\d*).{2,5}";
                List<PdfString> list = FindListPdfString(patternPdfIssueAmt, path);
                List<string> groupName = new List<string>() { "IssueAmt" };
                string patternIssueAmt = @"发行规模\D*(?<IssueAmt>\d+\.*\d*)";
                List<string> IssueAmt = FindValueFromListPdfString(patternIssueAmt, list, groupName);
                if (IssueAmt == null || IssueAmt.Count == 0)
                {
                    return string.Empty;
                }
                double d = System.Convert.ToDouble(IssueAmt[0]);

                return (d * 100).ToString();
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get Shenzhen issue amt error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return string.Empty;
            }
        }

        private string GetSZName(string path)
        {
            try
            {
                string patternPdfName = @"\D\D\D\D(：|:)\s*(?<name>.+)";
                List<PdfString> list = FindListPdfString(patternPdfName, path);
                List<string> groupName = new List<string>() { "name" };
                string patternName = @".{20,30}证券简称\s*(：|:)\s*(?<name>.+)\s*证券代码";
                List<string> name = FindValueFromListPdfString(patternName, list, groupName);
                if (name == null || name.Count == 0)
                {
                    return string.Empty;
                }
                
                return name[0];
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get Shenzhen name error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return string.Empty;
            }
        }

        private string GetKZZName(string path)
        {
            try
            {
                string patternPdfName = @"\D\D\D\D(：|:)\s*(?<name>.+)";
                List<PdfString> list = FindListPdfString(patternPdfName, path);
                List<string> groupName = new List<string>() { "name" };
                string patternName = @"债券简称\s*(：|:)\s*(?<name>.+?转债)";
                List<string> name = FindValueFromListPdfString(patternName, list, groupName);
                if (name == null || name.Count == 0)
                {
                    return string.Empty;
                }

                return name[0];
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get KZZ name error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return string.Empty;
            }
        }

        private string GetGSZName(string path)
        {
            try
            {
                string patternPdfName = @"\D\D\D\D(：|:)\s*(?<name>.+)";
                List<PdfString> list = FindListPdfString(patternPdfName, path);
                List<string> groupName = new List<string>() { "name" };
                string patternName = @"证券简称\s*(：|:)\s*(?<name>.+)\s*证券代码";
                List<string> name = FindValueFromListPdfString(patternName, list, groupName);
                if (name == null || name.Count == 0)
                {
                    return string.Empty;
                }

                return name[0];
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Get GSZ name error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return string.Empty;
            }
        }

        private void AnalyseJRZword()
        {
            try
            {
                foreach (string item in JRZtmpFilePath)
                {
                    if (!File.Exists(item))
                    {
                        continue;
                    }
                    BondAddInfo tmp = new BondAddInfo();
                    object path = item;
                    object oMissing = Missing.Value;
                    Word.Document doc = WordApp.Documents.Open(ref path, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    int tableIndex = 1;
                    Word.Table table = doc.Tables[tableIndex];
                    string name = table.Cell(1, 4).Range.Text.Trim(new char[] { '\r', '\a' });
                    string ric = table.Cell(2, 2).Range.Text.Trim(new char[] { '\r', '\a' });
                    string effectiveDay = table.Cell(3, 4).Range.Text.Trim(new char[] { '\r', '\a' });
                    string issueAmt = table.Cell(4, 4).Range.Text.Trim(new char[] { '\r', '\a' });
                    issueAmt = issueAmt.Replace("亿","");
                    string copeRate = table.Cell(5, 4).Range.Text.Trim(new char[] { '\r', '\a' });
                    copeRate = copeRate.Replace("%", "");
                    string[] array = effectiveDay.Split(new char[] { '年','月','日' });
                    if (array[1].Length == 1)
                    {
                        array[1] = "0" + array[1];
                    }
                    if (array[2].Length == 1)
                    {
                        array[2] = "0" + array[2];
                    }
                    string date = string.Format("{0}-{1}-{2}", array[0], array[1], array[2]);
                    date = DateTime.ParseExact(date, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd-MMM-yyyy");
                    tmp.EffectiveDate = date;
                    tmp.IssueAmt = (Double.Parse(issueAmt) * 100).ToString();
                    tmp.Cope_Rate = copeRate;
                    tmp.UpdatedOn = DateTime.ParseExact(ConfigObj.Date, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd-MMM-yy");
                    tmp.Ric = "CN" + ric + "=SZ";
                    tmp.OfficialCode = ric;
                    tmp.BKGD_Ref = "CN" + ric + "=SZ1";
                    tmp.ProvSymb = "CN" + ric + "=SZ";
                    tmp.PayFreq = "ANNUAL";
                    char[] nameChar = name.ToCharArray();
                    string namePinYin = null;
                    for (int i = 0; i < nameChar.Length; i++)
                    {
                        if (nameChar[i] >= 0x4e00 && nameChar[i] <= 0x9fbb)
                        {
                            ChineseChar cc = new ChineseChar(nameChar[i]);
                            ReadOnlyCollection<string> ss = cc.Pinyins;
                            if (ss.Count > 0)
                            {
                                string py = ss[0];
                                namePinYin += py.Substring(0, py.Length - 1);
                            }
                        }
                        else
                        {
                            namePinYin += nameChar[i];
                        }
                    }
                    if (namePinYin.Length > 16)
                    {
                        namePinYin = namePinYin.Substring(0, 16);
                    }
                    tmp.IDNDisplayName = namePinYin;
                    tmp.CoiDisplayNmll = ChineseConverter.Convert(name, ChineseConversionDirection.SimplifiedToTraditional);
                    BondADDInfo.Add(tmp);
                }
                
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Analyse JRZ WORD error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void AnalyseKZZpdf()
        {
            try
            {
                foreach (string item in KZZtmpFilePath)
                {
                    if (!File.Exists(item))
                    {
                        continue;
                    }
                    BondAddInfo tmp = new BondAddInfo();
                    tmp.UpdatedOn = DateTime.ParseExact(ConfigObj.Date, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd-MMM-yy");
                    tmp.EffectiveDate = GetSZEffectiveDate(item);
                    string ric = GetSZRic(item);
                    tmp.Ric = "CN" + ric + "=SZ";
                    tmp.OfficialCode = ric;
                    tmp.BKGD_Ref = "CN" + ric + "=SZ1";
                    tmp.ProvSymb = "CN" + ric + "=SZ";
                    tmp.PayFreq = "ANNUAL";
                    tmp.IssueAmt = GetSZIssueAmt(item);
                    string name = GetKZZName(item);
                    name = name.Replace("转债", "");
                    char[] nameChar = name.ToCharArray();
                    string namePinYin = null;
                    for (int i = 0; i < nameChar.Length; i++)
                    {
                        if (nameChar[i] >= 0x4e00 && nameChar[i] <= 0x9fbb)
                        {
                            ChineseChar cc = new ChineseChar(nameChar[i]);
                            ReadOnlyCollection<string> ss = cc.Pinyins;
                            if (ss.Count > 0)
                            {
                                string py = ss[0];
                                namePinYin += py.Substring(0, py.Length - 1);
                            }
                        }
                        else
                        {
                            namePinYin += nameChar[i];
                        }
                    }
                    if (namePinYin.Length > 16)
                    {
                        namePinYin = namePinYin.Substring(0, 16);
                    }
                    tmp.IDNDisplayName = namePinYin+" ZZ";
                    tmp.CoiDisplayNmll = ChineseConverter.Convert(name, ChineseConversionDirection.SimplifiedToTraditional);
                    BondADDInfo.Add(tmp);
                }
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Analyse KZZ pdf error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }
        
        private void AnalyseSZpdf()
        {
            try
            {
                foreach(string item in SZtmpFilePath)
                {
                    if (!File.Exists(item))
                    {
                        continue;
                    }
                    BondAddInfo tmp = new BondAddInfo();
                    tmp.UpdatedOn = DateTime.ParseExact(ConfigObj.Date, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd-MMM-yy");
                    tmp.EffectiveDate = GetSZEffectiveDate(item);
                    string ric = GetSZRic(item);
                    tmp.Ric = "CN" + ric + "=SZ";
                    tmp.OfficialCode = ric;
                    tmp.BKGD_Ref = "CN" + ric + "=SZ1";
                    tmp.ProvSymb = "CN" + ric + "=SZ";
                    tmp.PayFreq = "ANNUAL";
                    tmp.IssueAmt = GetSZIssueAmt(item);
                    string name = GetSZName(item);
                    char[] nameChar = name.ToCharArray();
                    string namePinYin = null;
                    for (int i = 0; i < nameChar.Length; i++ )
                    {
                        if (nameChar[i] >= 0x4e00 && nameChar[i] <= 0x9fbb)
                        {
                            ChineseChar cc = new ChineseChar(nameChar[i]);
                            ReadOnlyCollection<string> ss = cc.Pinyins;
                            if (ss.Count > 0)
                            {
                                string py = ss[0];
                                namePinYin += py.Substring(0,py.Length - 1);
                            }
                        }
                        else
                        {
                            namePinYin += nameChar[i];
                        }
                    }
                    if (namePinYin.Length > 16)
                    {
                        namePinYin = namePinYin.Substring(0, 16);
                    }
                    tmp.IDNDisplayName = namePinYin;
                    tmp.CoiDisplayNmll = ChineseConverter.Convert(name, ChineseConversionDirection.SimplifiedToTraditional);
                    BondADDInfo.Add(tmp);
                }
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Analyse Shenzhen pdf error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void AnalyseGSZpdf()
        {
            try
            {
                foreach (string item in GSZtmpFilePath)
                {
                    if (!File.Exists(item))
                    {
                        continue;
                    }
                
                    BondAddInfo tmp = new BondAddInfo();
                    tmp.UpdatedOn = DateTime.ParseExact(ConfigObj.Date, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture).ToString("dd-MMM-yy");
                    tmp.EffectiveDate = GetSZEffectiveDate(item);
                    string ric = GetSZRic(item);
                    tmp.Ric = "CN" + ric + "=SZ";
                    tmp.OfficialCode = ric;
                    tmp.BKGD_Ref = "CN" + ric + "=SZ1";
                    tmp.ProvSymb = "CN" + ric + "=SZ";
                    tmp.PayFreq = "ANNUAL";
                    tmp.IssueAmt = GetSZIssueAmt(item);
                    string name = GetGSZName(item);
                    char[] nameChar = name.ToCharArray();
                    string namePinYin = null;
                    for (int i = 0; i < nameChar.Length; i++)
                    {
                        if (nameChar[i] >= 0x4e00 && nameChar[i] <= 0x9fbb)
                        {
                            ChineseChar cc = new ChineseChar(nameChar[i]);
                            ReadOnlyCollection<string> ss = cc.Pinyins;
                            if (ss.Count > 0)
                            {
                                string py = ss[0];
                                namePinYin += py.Substring(0, py.Length - 1);
                            }
                        }
                        else
                        {
                            namePinYin += nameChar[i];
                        }
                    }
                    if (namePinYin.Length > 16)
                    {
                        namePinYin = namePinYin.Substring(0, 16);
                    }
                    tmp.IDNDisplayName = namePinYin;
                    tmp.CoiDisplayNmll = ChineseConverter.Convert(name, ChineseConversionDirection.SimplifiedToTraditional);
                    BondADDInfo.Add(tmp);
                }
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Analyse GSZ pdf error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private ChinaIPOInfo FormatIPOinfo(IPOSource ipoInfo,int type)
        {

            ChinaIPOInfo tmp = new ChinaIPOInfo();
            tmp.UpdateDate = System.Convert.ToDateTime(ConfigObj.Date).ToString("dd-MMM-yy");
            if (ipoInfo.Ticker.StartsWith("6") || ipoInfo.Ticker.StartsWith("9"))
            {
                tmp.RIC = ipoInfo.Ticker + ".SS";
            }
            else if (ipoInfo.Ticker.StartsWith("0") || ipoInfo.Ticker.StartsWith("2") || ipoInfo.Ticker.StartsWith("3"))
            {
                tmp.RIC = ipoInfo.Ticker + ".SZ";
            }
            else
            {
                tmp.RIC = ipoInfo.Ticker;
            }

            if (type == 1)
            {
                tmp.FM = type.ToString();
            }
            else if (type == 2)
            {
                tmp.FM = type.ToString();
                tmp.EffectiveDate = ((DateTime)ipoInfo.ListingDate).ToString("dd-MMM-yyyy");
            }

            if (ipoInfo.LongName.Length > 16)
            {
                tmp.IDNName = ipoInfo.LongName.Substring(0,16);
            }
            else
            {
                tmp.IDNName = ipoInfo.LongName;
            }
            tmp.OfficialCode = ipoInfo.Ticker;
            tmp.COI_DSPLY = ipoInfo.CompanyName;
            tmp.BCAST_REF = tmp.RIC;
            tmp.LONGLINK = "t" + tmp.RIC;
            tmp.BKGD_REF = tmp.RIC + "B2";
            tmp.Type = "Ordinary Share";
            tmp.EnglishName = ipoInfo.EnglishName;

            GetListingShareAndIpoPriceFromWebSite(ipoInfo.Ticker, ref tmp);
            return tmp;
        }

        private void GetIPOinfoFromDBbyID(List<IPOAnnOUCe> list)
        {
            if (list == null)
            {
                string msg = string.Format("IPO id list is null .");
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            try
            {
                string lastIPOid = string.Empty;
                foreach (IPOAnnOUCe item in list)
                {
                    if (item.IPoiD.ToString().Equals(lastIPOid))
                    {
                        continue;
                    }
                    else
                    {
                        lastIPOid = item.IPoiD.ToString();
                        var sqlString = string.Format("SELECT * FROM ipo_source WHERE ID = '{0}'",item.IPoiD);
                        List<IPOSource> ipoInfoList = extendContext.ExecuteQuery<IPOSource>(sqlString, 60).ToList();
                        if (ipoInfoList.Count > 0)
                        {
                            if (item.Type.Equals("Intention"))
                            {
                                IPOinfo.Add(FormatIPOinfo(ipoInfoList[0], 1));
                            }
                            else if(item.Type.Equals("Listing"))
                            {
                                IPOinfo.Add(FormatIPOinfo(ipoInfoList[0], 2));
                            }
                            
                        }
                        
                    }
                    
                }

                
            }
            catch (Exception ex)
            {
                string msg = string.Format("Get IPO info from Ace DB error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GetIPOIDfromDBbyDate()
        {
            if (extendContext == null)
            {
                string msg = string.Format("DataSourceString is null .");
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                var sqlString = string.Format("SELECT * FROM ipo_annouce WHERE (Type = 'Intention' OR Type = 'Listing') AND (CreatedDate > '2014-07-18' AND CreatedDate < '2014-07-19')");
                List<IPOAnnOUCe> list = extendContext.ExecuteQuery<IPOAnnOUCe>(sqlString, 60).ToList();
                GetIPOinfoFromDBbyID(list);
            }
            catch (Exception ex)
            {
                string msg = string.Format("Get IPO id from Ace DB error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GetListingShareAndIpoPriceFromWebSite(string ticker, ref ChinaIPOInfo info)
        {
            try
            {
                string url = @"http://www.cninfo.com.cn/  ";
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 300000;
                request.Method = "GET";
                request.CookieContainer = cookies;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                StreamReader sr = new StreamReader(response.GetResponseStream());
                string st = sr.ReadToEnd();

                url = @"http://www.cninfo.com.cn/" + "information/brief/shmb" + ticker + ".html";
                request = WebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 300000;
                request.Method = "GET";
                request.CookieContainer = cookies;
                response = (HttpWebResponse)request.GetResponse();

                sr = new StreamReader(response.GetResponseStream());
                st = sr.ReadToEnd();

                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(st);
                HtmlNodeCollection td = doc.DocumentNode.SelectNodes(".//td[@class='zx_data2']");

                info.ListingShare = (System.Convert.ToDouble(td[14].InnerText.Trim())*10000.0).ToString();
                info.IPOPrice = td[15].InnerText.Trim();
            }
            catch (Exception ex)
            {
                string msg = string.Format("Get listing share and IPO price from web site error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GenerateFMFile()
        {
            try
            {
                if (!File.Exists(ConfigObj.TemplateFilePath))
                {
                    Logger.Log("Can't find template file!", Logger.LogType.Error);
                    return;
                }

                string date = System.Convert.ToDateTime(ConfigObj.Date).ToString("dd-MMM-yy");
                date = date.Replace("-"," ");
                string targetFile = ConfigObj.OutputPath + "\\" + "China FM for " + date + ".xlsx"; 

                if (File.Exists(targetFile))
                {
                    File.Delete(targetFile);
                }
                File.Copy(ConfigObj.TemplateFilePath, targetFile);

                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, targetFile);
                Worksheet sheet = workbook.Sheets[1];

                int i = 2;
                sheet.Cells[1, 1] = date.Replace(" ", "-");
                foreach (ChinaIPOInfo item in IPOinfo)
                {
                    sheet.Cells[4, i] = item.UpdateDate;
                    sheet.Cells[5, i] = item.EffectiveDate;
                    sheet.Cells[6, i] = item.RIC;
                    sheet.Cells[7, i] = item.FM;
                    sheet.Cells[8, i] = item.IDNName;
                    sheet.Cells[9, i] = item.OfficialCode;
                    sheet.Cells[10,i] = item.COI_DSPLY;
                    sheet.Cells[11,i] = item.BCAST_REF;
                    sheet.Cells[12,i] = item.LONGLINK;
                    sheet.Cells[13,i] = item.BKGD_REF;
                    sheet.Cells[14,i] = item.Type;
                    sheet.Cells[16,i] = item.EnglishName;
                    sheet.Cells[19,i] = item.ListingShare;
                    sheet.Cells[20,i] = item.IPOPrice;
                    i++;
                }
                i = 2;
                foreach (BondAddInfo item in BondADDInfo)
                {
                    sheet.Cells[28, i] = item.UpdatedOn;
                    sheet.Cells[29, i] = item.EffectiveDate;
                    sheet.Cells[30, i] = item.Ric;
                    sheet.Cells[31, i] = item.IDNDisplayName;
                    sheet.Cells[32, i] = item.OfficialCode;
                    sheet.Cells[33, i] = item.Matur_Date;
                    sheet.Cells[34, i] = item.Cope_Rate;
                    sheet.Cells[35, i] = item.Coupn_Date;
                    sheet.Cells[36, i] = item.Issue_Date;
                    sheet.Cells[37, i] = item.BondType;
                    sheet.Cells[38, i] = item.CoiDisplayNmll;
                    sheet.Cells[39, i] = item.BKGD_Ref;
                    sheet.Cells[40, i] = item.IssueAmt;
                    sheet.Cells[41, i] = item.PayFreq;
                    sheet.Cells[42, i] = item.ProvSymb;
                }

                workbook.Save();
            }
            catch (Exception ex)
            {
                string msg = string.Format("Generate FM file error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }
    }
}
