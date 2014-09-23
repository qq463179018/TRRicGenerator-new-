using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.IO;
using System.IO.Compression;

namespace Ric.Util
{
    /// <summary>
    /// Operations for web page,like get page sources
    /// </summary>
    public class WebClientUtil
    {
        //Load page source as Html document 
        public static HtmlDocument GetHtmlDocument(string uri, int timeout)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(GetPageSource(uri, timeout));
            return doc;
        }

        public static HtmlDocument GetHtmlDocument(string uri, int timeout, string postData)
        {
            HtmlDocument doc = new HtmlDocument();
            String pageSource = GetDynamicPageSource(uri, timeout, postData);
            if (!String.IsNullOrEmpty(pageSource))
                doc.LoadHtml(pageSource);
            else
                return null;
            return doc;
        }

        public static HtmlDocument GetHtmlDocument(string uri, int timeout, string postData, Encoding encoding)
        {
            HtmlDocument doc = new HtmlDocument();
            String pageSource = GetPageSource(null, uri, timeout, postData, encoding);
            if (!String.IsNullOrEmpty(pageSource))
                doc.LoadHtml(pageSource);
            else
                return null;
            return doc;
        }

        public static void DownloadFile(string uri, int timeout, string targetFilePath)
        {
            DownloadFile(null, uri, timeout, targetFilePath, "");
        }

        /// <summary>
        /// Download file from website to local 
        /// </summary>
        /// <param name="uri">Url of the file</param>
        /// <param name="timeout">Timeout waiting for a response from server, in seconds.</param>
        /// <param name="targetFilePath">Local path where the file will be saved.</param>
        public static void DownloadFile(string uri, int timeout, string targetFilePath, string postData)
        {
            DownloadFile(null, uri, timeout, targetFilePath, postData);
        }

        /// <summary>
        /// Download file from website
        /// </summary>
        /// <param name="wc">Current webclient</param>
        /// <param name="uri">Website url</param>
        /// <param name="timeout">Timeout waiting for a response from server, in seconds.</param>
        /// <param name="targetFilePath">The local file path</param>
        public static void DownloadFile(AdvancedWebClient wc, string uri, int timeout, string targetFilePath, string postData)
        {
            if (wc == null)
            {
                wc = new AdvancedWebClient();
            }

            Byte[] buf = null;
            wc.Timeout = timeout;

            if (!string.IsNullOrEmpty(postData))
            {
                wc.PostData = postData;
            }

            int retriesLeft = 5;
            Exception innerException = null;

            while ((buf == null || buf.Length == 0) && retriesLeft-- > 0)
            {
                try
                {
                    buf = wc.DownloadData(uri);
                }
                catch (Exception ex)
                {
                    innerException = ex;
                }
            }

            if (buf == null)
            {
                throw new Exception(string.Format("Cannot download file from [{0}].", uri), innerException);
            }
            else
            {
                string targetDir = Path.GetDirectoryName(targetFilePath);
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }
                using (var fs = new FileStream(targetFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    fs.Write(buf, 0, buf.Length);
                }
            }
        }

        public static string GetDynamicPageSource(string uri, int timeout, string postData)
        {
            string pageSource = string.Empty;
            Encoding encoding = Encoding.UTF8;
            byte[] buf = null;
            int retriesLeft = 5;
            if (!string.IsNullOrEmpty(postData))
                buf = encoding.GetBytes(postData);

            while (string.IsNullOrEmpty(pageSource) && retriesLeft-- > 0)
            {
                var request = WebRequest.Create(uri) as HttpWebRequest;
                request.Timeout = timeout;
                //request.AllowAutoRedirect = false;
                request.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:6.0.2) Gecko/20100101 Firefox/6.0.2";
                request.Method = "GET";
                request.ContentType = "application/x-www-form-urlencoded";
                if (!string.IsNullOrEmpty(postData))
                {
                    request.Method = "POST";
                    request.ContentLength = buf.Length;
                    request.GetRequestStream().Write(buf, 0, buf.Length);
                }
                using (WebResponse response = request.GetResponse())
                {
                    var sr = new StreamReader(response.GetResponseStream());
                    pageSource = sr.ReadToEnd();
                }
            }
            if (pageSource == string.Empty)
            {
                throw new Exception(string.Format("Cannot download page {0} with post data {1}", uri, postData));
            }
            return pageSource;
        }

        public static string GetDynamicPageSource(HttpWebRequest request, string postData, Encoding encode)
        {
            string pageSource = string.Empty;
            Encoding encoding = encode;
            byte[] buf = null;
            int retriesLeft = 5;
            if (!string.IsNullOrEmpty(postData))
                buf = encoding.GetBytes(postData);

            while (string.IsNullOrEmpty(pageSource) && retriesLeft-- > 0)
            {
                if (!string.IsNullOrEmpty(postData))
                {
                    request.ContentLength = buf.Length;
                    request.GetRequestStream().Write(buf, 0, buf.Length);
                }
                using (WebResponse response = request.GetResponse())
                {
                    var sr = new StreamReader(response.GetResponseStream(), encode);
                    pageSource = sr.ReadToEnd();
                }
            }
            if (pageSource == string.Empty)
            {
                throw new Exception(string.Format("Cannot download page with post data {0}", postData));
            }
            return pageSource;
        }

        public static string GetPageSource(string uri, int timeout)
        {
            return GetPageSource(uri, timeout, null);
        }

        public static string GetPageSource(string uri, int timeout, string postData)
        {
            return GetPageSource(null, uri, timeout, postData);
        }

        //Get web page source from the certain URI
        public static string GetPageSource(AdvancedWebClient wc, string uri, int timeout, string postData)
        {
            return GetPageSource(wc, uri, timeout, postData,null);
        }

        //Get web page source from the certain URI
        public static byte[] Decompress(Byte[] bytes)
        {
            using (var tempMs = new MemoryStream())
            {
                using (var ms = new MemoryStream(bytes))
                {
                    var decompress = new GZipStream(ms, CompressionMode.Decompress);

                    byte[] buffer = new byte[512];
                    int count = 0;
                    while ((count = decompress.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        tempMs.Write(buffer, 0, count);
                    }

                    decompress.Close();
                    return tempMs.ToArray();
                }
            }
        }
        public static string GetPageSource(AdvancedWebClient wc, string uri, int timeout, string postData, Encoding encoding)
        {
            if (wc == null)
            {
                wc = new AdvancedWebClient();
            }

            byte[] buf = null;
            wc.Timeout = timeout;
            if (!string.IsNullOrEmpty(postData))
            {
                wc.PostData = postData;
            }

            int retriesLeft = 5;
            while ((buf == null || buf.Length == 0) && retriesLeft-- > 0)
            {
                try { buf = wc.DownloadData(uri); }
                catch (Exception ex) { string errInfo = ex.ToString(); }
            }

            if (buf == null)
            {
                throw new Exception(string.Format("Cannot download page {0}", uri));
            }

            ////Encoding encoding = Encoding.UTF8;
            //Encoding encoding = Encoding.GetEncoding("gb2312");
            string contentEncoding = wc.ResponseHeaders["Content-Encoding"];
            if (!string.IsNullOrEmpty(contentEncoding))
            {
                if (contentEncoding.Equals("gzip"))
                {
                    buf = Decompress(buf);
                }
            }

            string html = Encoding.UTF8.GetString(buf);
            if (encoding == null)
            {
                Encoding encoding2 = GetEncoding(html);
                if (encoding2 != null && encoding2 != encoding)
                {
                    return encoding2.GetString(buf);

                }

            }
            else
            {
                html = encoding.GetString(buf);
            }
            return html;

            //Encoding encoding2 = GetEncoding(Encoding.UTF8.GetString(buf));
            //if ((encoding2 == null || encoding2 != encoding) && encoding != null)
            //{
            //    string html = encoding.GetString(buf);
            //    return html;
            //}
            //else
            //{
            //    return encoding2.GetString(buf);
            //}
        }
        public static string GetPageSourceCompressed(AdvancedWebClient wc, string uri, int timeout, string postData, Encoding encoding)
        {
            if (wc == null)
            {
                wc = new AdvancedWebClient();
            }

            byte[] buf = null;
            wc.Timeout = timeout;
            if (!string.IsNullOrEmpty(postData))
            {
                wc.PostData = postData;
            }

            int retriesLeft = 5;
            while ((buf == null || buf.Length == 0) && retriesLeft-- > 0)
            {
                try { buf = wc.DownloadData(uri); }
                catch (Exception ex) { string errInfo = ex.ToString(); }
            }

            if (buf == null)
            {
                throw new Exception(string.Format("Cannot download page {0}", uri));
            }

            ////Encoding encoding = Encoding.UTF8;
            //Encoding encoding = Encoding.GetEncoding("gb2312");
            string contentEncoding = wc.ResponseHeaders["Content-Encoding"];
            buf = Decompress(buf);
            string html = encoding.GetString(buf);
            //string test = Encoding.GetEncoding("EUC-KR").GetString(buf);
            Encoding encoding2 = GetEncoding(html);
            if (encoding2 == null || encoding2 != encoding)
            {
                return html;
            }
            return encoding2.GetString(buf);
        }

        //Get the encoding type of a web page
        public static Encoding GetEncoding(string html)
        {
            try
            {
                string pattern = @"(?i)\bcharset=(?<charset>[-a-zA-Z_0-9]+)";
                string charset = Regex.Match(html, pattern).Groups["charset"].Value;
                return Encoding.GetEncoding(charset);
            }
            catch (Exception)
            {
                return null;
            }
        }

    }

    public class AdvancedWebClient : WebClient
    {
        private string _userAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; .NET CLR 2.0.50727)";

        public string UserAgent
        {
            get { return _userAgent; }
            set { _userAgent = value; }
        }
        private int _timeout = -1;

        public int Timeout
        {
            get { return _timeout; }
            set { _timeout = value; }
        }
        private CookieContainer _cookieContainer = new CookieContainer();

        public CookieContainer CookieContainer
        {
            get { return _cookieContainer; }
            set { _cookieContainer = value; }
        }

        private string _postData = string.Empty;

        public string PostData
        {
            get { return _postData; }
            set { _postData = value; }
        }

        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest request = base.GetWebRequest(address);
            if (request is HttpWebRequest)
            {
                (request as HttpWebRequest).CachePolicy = new System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore);
                (request as HttpWebRequest).Proxy = WebRequest.DefaultWebProxy;
                (request as HttpWebRequest).CookieContainer = _cookieContainer;
                (request as HttpWebRequest).UserAgent = _userAgent;
                (request as HttpWebRequest).Timeout = _timeout;

                if (!string.IsNullOrEmpty(PostData))
                {
                    byte[] buf = Encoding.UTF8.GetBytes(PostData);
                    (request as HttpWebRequest).Method = "POST";
                    (request as HttpWebRequest).ContentType = "application/x-www-form-urlencoded";
                    (request as HttpWebRequest).ContentLength = buf.Length;
                    (request as HttpWebRequest).GetRequestStream().Write(buf, 0, buf.Length);
                }
            }
            return request;
        }

    }
}
