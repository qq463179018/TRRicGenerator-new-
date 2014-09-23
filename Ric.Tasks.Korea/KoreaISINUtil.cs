using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using HtmlAgilityPack;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Ric.Db.Info;
using Ric.Util;
using System.IO;
using System.Net;

namespace Ric.Tasks.Korea
{
    public class KoreaISINUtil
    {
        private static string queryURL = "http://isin.krx.co.kr/jsp/realBoard99.jsp";
        private static string detailURLFormat = "http://isin.krx.co.kr/jsp/BA_VW010.jsp?isu_cd={0}&modi=f&req_no=";
        private static string detailURLFormatKDR = "http://isin.krx.co.kr/jsp/BA_VW013.jsp?isu_cd={0}&modi=f&req_no=";
        private static CookieContainer cookies = new CookieContainer();
        private static string GetPageSource(string uri, string postData)
        {
            HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
            request.ProtocolVersion = HttpVersion.Version11;
            request.Timeout = 100000;
            request.Method = "POST";
            request.Accept = @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
            request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
            request.Headers["Cache-Control"] = @"max-age=0";
            request.ContentType = @"multipart/form-data; boundary=----WebKitFormBoundaryupl1XUpxfQWZVyQB";
            request.Host = @"isin.krx.co.kr";
            request.Headers["Origin"] = @"http://isin.krx.co.kr";
            request.Referer = @"http://isin.krx.co.kr/srch/srch.do?method=srchList";
            request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
            request.CookieContainer = cookies;
            byte[] buf = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = buf.Length;
            request.GetRequestStream().Write(buf, 0, buf.Length);
            HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(httpResponse.GetResponseStream());
            return sr.ReadToEnd();
        }

        public static HtmlNodeCollection SearchISIN(string companyName, bool onlyEquity)
        {
            //string securityScope = onlyEquity ? "01" : "99";
            //onlyNonListing = onlyEquity ? onlyNonListing : false; // no choice for all security scope
            //string listScope = onlyNonListing ? "lst_yn2=N" : "lst_yn_all=on&lst_yn1=Y&lst_yn2=N&lst_yn3=D";
            //companyName = HttpUtility.UrlEncode(companyName, Encoding.GetEncoding("euc-kr"));

            //string postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun={0}"
            //    + "&{1}&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on"
            //    + "&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1"
            //    + "&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word={2}"
            //    + "&ef_iss_inst_cd=&ef_isu_nm=&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=",
            //    securityScope,
            //    listScope,
            //    companyName);
            string uri = @"http://isin.krx.co.kr/srch/srch.do?method=srchList";
            HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "GET";
            request.CookieContainer = cookies;
            request.Host = @"isin.krx.co.kr";
            request.Accept = @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
            request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
            request.KeepAlive = true;
            request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(response.GetResponseStream());
            string st = sr.ReadToEnd();

            string postDataPath;
            if (onlyEquity)
            {
                if (!File.Exists(@"Config\Korea\EquityISINPostData.txt"))
                {
                    System.Windows.Forms.MessageBox.Show(string.Format("The file {0} missing.", @"Config\Korea\EquityISINPostData.txt"));
                    return null;
                }
                postDataPath = @"Config\Korea\EquityISINPostData.txt";   
            }
            else
            {
                if (!File.Exists(@"Config\Korea\KDRISINPostData.txt"))
                {
                    System.Windows.Forms.MessageBox.Show(string.Format("The file {0} missing.", @"Config\Korea\KDRISINPostData.txt"));
                    return null;
                }
                postDataPath = @"Config\Korea\KDRISINPostData.txt";
            }

            
            string postData = string.Format(File.ReadAllText(postDataPath, Encoding.UTF8), companyName, companyName);

            //AdvancedWebClient wc = new AdvancedWebClient();
            //string pageSource = WebClientUtil.GetPageSource(wc, uri, 300000, postData, Encoding.GetEncoding("euc-kr"));
            string pageSource = GetPageSource(uri,postData);
            if (string.IsNullOrEmpty(pageSource))
            {
                return null;
            }

            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(pageSource);

            if (doc == null)
            {
                return null;
            }

            if (doc.DocumentNode.SelectNodes("//table").Count < 2)
            {
                return null;
            }

            HtmlNodeCollection records = doc.DocumentNode.SelectNodes("//table")[2].SelectNodes(".//tr");

            if (records.Count == 1)
            {
                return null;
            }

            records.RemoveAt(0);

            return records;
        }

        public static bool UpdateEquityISINReport(KoreaEquityInfo item)
        {
            HtmlNodeCollection records = SearchISIN(item.KoreaName, true);

            if (records == null)
            {
                return false;
            }

            string isin = null;
            string ticker = null;
            string copName = null;

            foreach(HtmlNode n in records)
            {
                HtmlNode td = n.SelectSingleNode(".//td[2]");
                isin = td.InnerText.Trim();                
                HtmlNode aNode = td.SelectSingleNode(".//a");
                string checkMod = aNode.Attributes["href"].Value.Trim();
                string type = checkMod.Split(',')[2].Trim(new char[] { '\'', ')', ';' });
                if (type.Equals("STK"))
                {
                    ticker = GetTickerByISIN(isin, 1);
                }
               
                if (item.Ticker.Equals(ticker))
                {
                    copName = n.SelectSingleNode(".//td[3]").InnerText.Trim();

                    if (copName.Contains("보통주"))
                    {
                        item.Type = "ORD";
                    }
                    else if (copName.Contains("우선주"))
                    {
                        item.Type = "PRF";
                    }

                    item.ISIN = isin;

                    return true;
                }
            }

            return false;
        }

        
        /// <summary>
        /// Created on 2014-02-17, kind web page changed on 2014-01-20
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static bool UpdateEquityISINReportNew(KoreaEquityInfo item)
        {
            HtmlNodeCollection records = SearchISIN(item.KoreaName, true);

            if (records == null)
            {
                return false;
            }

            string isin = null;            
            string copName = null;
                        
            foreach (HtmlNode n in records)
            {
                HtmlNode td = n.SelectSingleNode(".//td[2]");
                isin = td.InnerText.Trim();
                HtmlNode aNode = td.SelectSingleNode(".//a");
                string checkMod = aNode.Attributes["href"].Value.Trim();
               // string type = checkMod.Split(',')[2].Trim(new char[] { '\'', ')', ';' });
                copName = n.SelectSingleNode(".//td[3]").InnerText.Trim();
                if (copName.Contains("보통주"))
                {
                    item.Type = "ORD";
                }
                else if (copName.Contains("우선주"))
                {
                    continue;
                    //item.Type = "PRF";
                }

                item.ISIN = isin;

                return true;
                
            }

            return false;
        }

        public static bool UpdateKDRISINReport(KoreaEquityInfo item)
        {
            HtmlNodeCollection records = SearchISIN(item.KoreaName, false);

            if (records == null)
            {
                return false;
            }

            string isin = null;
            string ticker = null;

            foreach (HtmlNode n in records)
            {
                if (!n.SelectSingleNode(".//td[4]").InnerText.Trim().Equals("예탁증서"))
                {
                    continue;
                }

                isin = n.SelectSingleNode(".//td[2]").InnerText.Trim();
                ticker = GetTickerByISIN(isin, 2);

                if (item.Ticker.Equals(ticker))
                {
                    item.ISIN = isin;
                    item.Type = "KDR";

                    return true;
                }
            }

            return false;
        }

        /// <summary>   
        /// Created on 2014-02-17, kind web page changed on 2014-01-20
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static bool UpdateKDRISINReportNew(KoreaEquityInfo item)
        {
            HtmlNodeCollection records = SearchISIN(item.KoreaName, false);

            if (records == null || records.Count < 1)
            {
                return false;
            }

            string isin = null;
            string copName = null;
            //foreach (HtmlNode n in records)
            //{
            //    if (!n.SelectSingleNode(".//td[4]").InnerText.Trim().Equals("예탁증서"))
            //    {
            //        continue;
            //    }

            //    isin = n.SelectSingleNode(".//td[2]").InnerText.Trim();               
                               
            //    item.ISIN = isin;
            //    item.Type = "KDR";

            //    return true;
                
            //}

            //return false;

            foreach (HtmlNode n in records)
            {
                HtmlNode td = n.SelectSingleNode(".//td[2]");
                isin = td.InnerText.Trim();
                HtmlNode aNode = td.SelectSingleNode(".//a");
                string checkMod = aNode.Attributes["href"].Value.Trim();
                // string type = checkMod.Split(',')[2].Trim(new char[] { '\'', ')', ';' });
                copName = n.SelectSingleNode(".//td[3]").InnerText.Trim();
                if (copName.Contains("보통주"))
                {
                    item.Type = "ORD";
                }
                else if (copName.Contains("우선주"))
                {
                    continue;
                    //item.Type = "PRF";
                }

                item.ISIN = isin;

                return true;

            }

            return false;
        }

        public static string GetTickerByISIN(string isin, int type)
        {
            string url = null;
            if (type == 1)
            {
                url = string.Format(detailURLFormat, isin);
            }
            else if (type == 2)
            {
                url = string.Format(detailURLFormatKDR, isin);
            }
            HtmlDocument detailDoc = WebClientUtil.GetHtmlDocument(url, 180000, null, Encoding.GetEncoding("EUC-KR"));
            if (detailDoc == null)
            {
                return null;
            }

            String ticker = String.Empty;

            if (detailDoc.DocumentNode.SelectNodes("//table").Count < 2)
            {
                return null;
            }

            HtmlNode tickerNode = detailDoc.DocumentNode.SelectNodes("//table")[2].SelectSingleNode(".//tr[3]/td[4]");

            if (tickerNode == null)
            {
                return null;
            }

            return tickerNode.InnerText.Trim().Substring(1).Trim();
        }

        public static void GetTickerAndLegalNameByISIN(KoreaEquityInfo item)
        {
            string isin = item.ISIN;
            string type = item.Type;
            if (string.IsNullOrEmpty(isin) || string.IsNullOrEmpty(type))
            {
                return;
            }
            string url = @"http://isin.krx.co.kr/srch/srch.do?method=srchPopup2";
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "POST";
            request.CookieContainer = cookies;
            request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
            request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
            request.Headers["Cache-Control"] = @"max-age=0";
            request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
            request.Accept = @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.ContentType = @"application/x-www-form-urlencoded";
            request.Headers["Origin"] = @"http://isin.krx.co.kr";
            request.Referer = @"http://isin.krx.co.kr/srch/srch.do?method=srchList";
            request.Host = @"isin.krx.co.kr";
            request.KeepAlive = true;
            string postData = @"stdcd_type=2&std_cd=" + item.ISIN;
            byte[] buf = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = buf.Length;
            request.GetRequestStream().Write(buf, 0, buf.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(response.GetResponseStream());
            string ss = sr.ReadToEnd();

            HtmlDocument detailDoc = new HtmlDocument();
            detailDoc.LoadHtml(ss);
            //string url = null;
            //if (type.Equals("ORD") || type.Equals("PRF"))
            //{
            //    url = string.Format(detailURLFormat, isin);
            //}
            //else if (type.Equals("KDR"))
            //{
            //    url = string.Format(detailURLFormatKDR, isin);
            //}

            //HtmlDocument detailDoc = WebClientUtil.GetHtmlDocument(url, 180000, null, Encoding.GetEncoding("EUC-KR"));
            //if (detailDoc == null)
            //{
            //    return;
            //}

            String ticker = String.Empty;

            if (detailDoc.DocumentNode.SelectNodes("//table").Count < 2)
            {
                return;
            }

            HtmlNode tickerNode = detailDoc.DocumentNode.SelectNodes("//table")[1].SelectSingleNode(".//tr[2]/td[2]");
            
            HtmlNode legalNameNode = null;
            legalNameNode = detailDoc.DocumentNode.SelectNodes("//table")[1].SelectSingleNode(".//tr[3]/td[2]");
            //if (type.Equals("ORD") || type.Equals("PRF"))
            //{
            //    legalNameNode = detailDoc.DocumentNode.SelectNodes("//table")[2].SelectSingleNode(".//tr[11]/td[4]");
            //}
            //else
            //{
            //   legalNameNode = detailDoc.DocumentNode.SelectNodes("//table")[2].SelectSingleNode(".//tr[5]/td[4]");
            //}
            
            if (tickerNode == null)
            {
                return;
            }

            item.Ticker = tickerNode.InnerText.Trim().Substring(1).Trim();
            item.LegalName = legalNameNode.InnerText.Trim();            
        }        
    }
}
