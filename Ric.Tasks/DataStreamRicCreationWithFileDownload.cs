using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using HtmlAgilityPack;



namespace Ric.Tasks
{
    public class DataStreamRicCreationWithFileDownload
    {
        
        public static StreamReader LoginWebSite( string usrName, string passWord, CookieContainer cookies)
        {
            string url = @"https://sso.deutsche-boerse.com/cas/activateAndLogin";
            string lt = null;
  
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "GET";
            request.CookieContainer = cookies;
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            StreamReader sr = new StreamReader(response.GetResponseStream());
            HtmlDocument doc = new HtmlDocument();

            doc.Load(sr);
            if (doc == null)
            {
                return null;
            }
            HtmlNode inputlt = doc.DocumentNode.SelectSingleNode("//input[@name='lt']");

            if (inputlt == null)
            {
                return null;
            }
            lt = inputlt.Attributes["value"].Value.Trim();
            url = @"https://sso.deutsche-boerse.com/cas/login";
            string postData = @"lt=" + lt + @"&_eventId=submit&sso_u="+usrName+"&password="+passWord;
            request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "POST";
            request.CookieContainer = cookies;
            request.Accept = @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
            request.Headers["Accept-Language"] = @"zh-CN,zh;q=0.8";
            request.Headers["Cache-Control"] = @"max-age=0";
            request.ContentType = @"application/x-www-form-urlencoded";
            request.Referer = @"https://sso.deutsche-boerse.com/cas/activateAndLogin";
            request.UserAgent = @"Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/33.0.1750.154 Safari/537.36";
            byte[] buf = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = buf.Length;
            request.GetRequestStream().Write(buf, 0, buf.Length);
            response = (HttpWebResponse)request.GetResponse();

            doc.Load(new StreamReader(response.GetResponseStream()));
            HtmlNode node = doc.DocumentNode.SelectSingleNode("//a[@href='https://contracts.deutsche-boerse.com/indexdata']");
            if (node != null)
            {
                url = @"https://contracts.deutsche-boerse.com/indexdata";
                request = WebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 300000;
                request.Method = "GET";
                request.CookieContainer = cookies;
                response = (HttpWebResponse)request.GetResponse();
            }
            
            sr = new StreamReader(response.GetResponseStream());
            return sr;
        }

        public static string ExpandDropDownBox(string actionStr, string token, string ctrla, CookieContainer cookies)
        {
            string url = @"https://contracts.deutsche-boerse.com/" + actionStr;
            long time = DateTime.Now.Ticks;
            time = (time - 621355968000000000) / 10000;
            
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "POST";
            request.CookieContainer = cookies;
            request.Accept = @"*/*";
            request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
            request.Headers["Accept-Language"] = @"en-US,en;q=0.8";
            request.KeepAlive = true;
            request.ContentType = @"application/x-www-form-urlencoded";
            request.Referer = @"//contracts.deutsche-boerse.com/indexdata/main.do";
            request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
            string postData = @"org.apache.struts.taglib.html.TOKEN=" + token + @"&ctrla="+ctrla + @"&AjaxRequestUniqueId=" + time + @"%20form.submit";
            byte[] buf = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = buf.Length;
            request.GetRequestStream().Write(buf, 0, buf.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(response.GetResponseStream());
            string ss = sr.ReadToEnd();
            return ss;
        }

        public static string ChangePage(string actionStr, string token, string ctrla, CookieContainer cookies)
        {
            string url = @"https://contracts.deutsche-boerse.com/" + actionStr;
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "POST";
            request.CookieContainer = cookies;
            request.Accept = @"*/*";
            request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
            request.Headers["Accept-Language"] = @"en-US,en;q=0.8";
            request.KeepAlive = true;
            request.ContentType = @"application/x-www-form-urlencoded";
            request.Referer = @"//contracts.deutsche-boerse.com/indexdata/main.do";
            request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";

            long time = DateTime.Now.Ticks;
            time = (time - 621355968000000000) / 10000;
            string postData = @"org.apache.struts.taglib.html.TOKEN=" + token + @"&ctrla=" + ctrla + @"&AjaxRequestUniqueId=" + time + @"%20form.submit";
            byte[]buf = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = buf.Length;
            request.GetRequestStream().Write(buf, 0, buf.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(response.GetResponseStream());
            string ss = sr.ReadToEnd();
            return ss;
        }

        public static string ChangeTab(string actionStr, string token,string param ,CookieContainer cookies)
        {
            string url = @"https://contracts.deutsche-boerse.com/" + actionStr;
            long time = DateTime.Now.Ticks;
            time = (time - 621355968000000000) / 10000;
            url = url + @"?ctrl=32_1&action=TabClick&param="+param+@"&org.apache.struts.taglib.html.TOKEN=" + token + @"&AjaxRequestUniqueId=" + time;
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "GET";
            request.CookieContainer = cookies;
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(response.GetResponseStream());
            string ss = sr.ReadToEnd();
            return ss;
        }

        public static void DownLoadFiles(string url, string fileName, CookieContainer cookies)
        {
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "GET";
            request.CookieContainer = cookies;
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream content = response.GetResponseStream();
            using (Stream file = File.Create(fileName))
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
}
