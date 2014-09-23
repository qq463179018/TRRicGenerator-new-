using System.Linq;
using System.Text;
using System.Web;
using HtmlAgilityPack;
using Ric.Db.Info;
using Ric.Util;

namespace Ric.Tasks
{
    public class KoreaISINUtil
    {
        private const string queryURL = "http://isin.krx.co.kr/jsp/realBoard99.jsp";
        private const string detailURLFormat = "http://isin.krx.co.kr/jsp/BA_VW010.jsp?isu_cd={0}&modi=f&req_no=";
        private const string detailURLFormatKDR = "http://isin.krx.co.kr/jsp/BA_VW013.jsp?isu_cd={0}&modi=f&req_no=";

        public static HtmlNodeCollection SearchISIN(string companyName, bool onlyEquity, bool onlyNonListing)
        {
            string securityScope = onlyEquity ? "01" : "99";
            onlyNonListing = onlyEquity && onlyNonListing; // no choice for all security scope
            string listScope = onlyNonListing ? "lst_yn2=N" : "lst_yn_all=on&lst_yn1=Y&lst_yn2=N&lst_yn3=D";
            companyName = HttpUtility.UrlEncode(companyName, Encoding.GetEncoding("euc-kr"));

            string postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun={0}"
                + "&{1}&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on"
                + "&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1"
                + "&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word={2}"
                + "&ef_iss_inst_cd=&ef_isu_nm=&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=",
                securityScope,
                listScope,
                companyName);

            AdvancedWebClient wc = new AdvancedWebClient();
            string pageSource = WebClientUtil.GetPageSource(wc, queryURL, 300000, postData, Encoding.GetEncoding("euc-kr"));

            if (string.IsNullOrEmpty(pageSource))
            {
                return null;
            }

            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(pageSource);

            if (doc.DocumentNode.SelectNodes("//table").Count < 2)
            {
                return null;
            }

            HtmlNodeCollection records = doc.DocumentNode.SelectNodes("//table")[1].SelectNodes(".//tr");

            if (records.Count == 1)
            {
                return null;
            }

            records.RemoveAt(0);

            return records;
        }

        public static bool UpdateEquityISINReport(KoreaEquityInfo item)
        {
            HtmlNodeCollection records = SearchISIN(item.KoreaName, true, true);

            if (records == null)
            {
                return false;
            }

            string isin;
            string ticker = null;
            string copName;

            foreach(HtmlNode n in records)
            {
                HtmlNode td = n.SelectSingleNode(".//td[2]");
                isin = td.InnerText.Trim();                
                HtmlNode aNode = td.SelectSingleNode(".//a");
                string checkMod = aNode.Attributes["href"].Value.Trim();
                string type = checkMod.Split(',')[2].Trim(new[] { '\'', ')', ';' });
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
            HtmlNodeCollection records = SearchISIN(item.KoreaName, true, true);

            if (records == null || records.Count > 1)
            {
                return false;
            }

            string isin;            
            string copName;
                        
            foreach (HtmlNode n in records)
            {
                HtmlNode td = n.SelectSingleNode(".//td[2]");
                isin = td.InnerText.Trim();
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

            return false;
        }

        public static bool UpdateKDRISINReport(KoreaEquityInfo item)
        {
            HtmlNodeCollection records = SearchISIN(item.KoreaName, false, false);

            if (records == null)
            {
                return false;
            }

            string isin;
            string ticker;

            foreach (HtmlNode n in records.Where(n => n.SelectSingleNode(".//td[4]").InnerText.Trim().Equals("예탁증서")))
            {
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
            HtmlNodeCollection records = SearchISIN(item.KoreaName, false, false);

            if (records == null || records.Count > 1)
            {
                return false;
            }

            string isin;           

            foreach (HtmlNode n in records.Where(n => n.SelectSingleNode(".//td[4]").InnerText.Trim().Equals("예탁증서")))
            {
                isin = n.SelectSingleNode(".//td[2]").InnerText.Trim();               
                               
                item.ISIN = isin;
                item.Type = "KDR";

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

            if (detailDoc.DocumentNode.SelectNodes("//table").Count < 2)
            {
                return null;
            }

            HtmlNode tickerNode = detailDoc.DocumentNode.SelectNodes("//table")[2].SelectSingleNode(".//tr[3]/td[4]");

            return tickerNode == null ? null : tickerNode.InnerText.Trim().Substring(1).Trim();
        }

        public static void GetTickerAndLegalNameByISIN(KoreaEquityInfo item)
        {
            string isin = item.ISIN;
            string type = item.Type;
            if (string.IsNullOrEmpty(isin) || string.IsNullOrEmpty(type))
            {
                return;
            }
            string url = null;
            if (type.Equals("ORD") || type.Equals("PRF"))
            {
                url = string.Format(detailURLFormat, isin);
            }
            else if (type.Equals("KDR"))
            {
                url = string.Format(detailURLFormatKDR, isin);
            }

            HtmlDocument detailDoc = WebClientUtil.GetHtmlDocument(url, 180000, null, Encoding.GetEncoding("EUC-KR"));
            if (detailDoc == null)
            {
                return;
            }

            if (detailDoc.DocumentNode.SelectNodes("//table").Count < 2)
            {
                return;
            }

            HtmlNode tickerNode = detailDoc.DocumentNode.SelectNodes("//table")[2].SelectSingleNode(".//tr[3]/td[4]");
            HtmlNode legalNameNode;
            if (type.Equals("ORD") || type.Equals("PRF"))
            {
                legalNameNode = detailDoc.DocumentNode.SelectNodes("//table")[2].SelectSingleNode(".//tr[11]/td[4]");
            }
            else
            {
               legalNameNode = detailDoc.DocumentNode.SelectNodes("//table")[2].SelectSingleNode(".//tr[5]/td[4]");
            }
            
            if (tickerNode == null)
            {
                return;
            }

            item.Ticker = tickerNode.InnerText.Trim().Substring(1).Trim();
            item.LegalName = legalNameNode.InnerText.Trim();            
        }        
    }
}
