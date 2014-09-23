using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Net;
using System.Xml;
using System.Globalization;
using Microsoft.Win32;
using System.Collections;
using HtmlAgilityPack;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.Web;
using System.Text.RegularExpressions;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks.Korea
{
    public class ISINTemp
    {
        public string Issuer { get; set; }
        public string ISIN { get; set; }
        public string ISINLink { get; set; }
        public string ItemName { get; set; }
        public string Category { get; set; }
        public string Status { get; set; }
        public string IssueDate { get; set; }
        public string ListingDate { get; set; }
        public string ISINCreateDate { get; set; }
    }

    public class ISINQuery
    {
        public string StartDate { get; set; } //yyyyMMdd
        public string EndDate { get; set; } //yyyyMMdd
        public string Category { get; set; }
        public string IssueCompany { get; set; }
        public string Code { get; set; }
        public ISINQuery(string startDate, string endDate, string category, string issueCompany, string code)
        {
            this.StartDate = startDate;
            this.EndDate = endDate;
            this.Category = category;//유가증권선택
            this.IssueCompany = issueCompany;//발행기관 
            this.Code = code;//종목명 및 종목코드 
        }
    }

    public class KoreaCompany
    {
        public string KoreaName { get; set; }
        public string LegalName { get; set; }
        public string ISIN { get; set; }
        public string Market { get; set; }
        public string ListingDate { get; set; }
    }

    public class KoreaNames
    {
        public string KoreaName { get; set; }
        public string IDNDisplayName { get; set; }
        public string LegalName { get; set; }
    }


    public class Common
    {
        [DllImport("user32.dll")]
        private static extern void GetWindowThreadProcessId(IntPtr hWnd, out int k);

        private static List<ISINTemp> getISINListFromSinglePage(string pageSource)
        {
            List<ISINTemp> isinList = new List<ISINTemp>();
            HtmlDocument htc = new HtmlDocument();
            if (!String.IsNullOrEmpty(pageSource))
                htc.LoadHtml(pageSource);
            if (htc != null)
            {
                var nodeList = htc.DocumentNode.SelectNodes("//table/tr/td/table/tr");
                if (nodeList != null)
                {
                    for (int x = 1; x < nodeList.Count - 1; x++)
                    {
                        ISINTemp temp = new ISINTemp();
                        temp.ISIN = MiscUtil.GetCleanTextFromHtml(nodeList[x].SelectSingleNode(".//td/a").InnerText);
                        temp.Issuer = MiscUtil.GetCleanTextFromHtml(nodeList[x].SelectNodes(".//td")[0].InnerText);
                        String attribute = MiscUtil.GetCleanTextFromHtml(nodeList[x].SelectSingleNode(".//td/a").Attributes["href"].Value);
                        String param1 = attribute.Split('(')[1].Split(',')[0].Trim(new Char[] { ' ', '\'' }).ToString();
                        String param2 = attribute.Split('(')[1].Split(',')[1].Trim(new Char[] { ' ', '\'' }).ToString();
                        temp.ISINLink = String.Format("http://isin.krx.co.kr/jsp/BA_VW021.jsp?isu_cd={0}&modi=t&req_no={1}", param2, param1);
                        temp.ItemName = MiscUtil.GetCleanTextFromHtml(nodeList[x].SelectNodes(".//td")[2].InnerText);
                        temp.Category = MiscUtil.GetCleanTextFromHtml(nodeList[x].SelectNodes(".//td")[3].InnerText);
                        temp.Status = MiscUtil.GetCleanTextFromHtml(nodeList[x].SelectNodes(".//td")[4].InnerText);
                        temp.IssueDate = MiscUtil.GetCleanTextFromHtml(nodeList[x].SelectNodes(".//td")[5].InnerText);
                        temp.ListingDate = MiscUtil.GetCleanTextFromHtml(nodeList[x].SelectNodes(".//td")[6].InnerText);
                        temp.ISINCreateDate = MiscUtil.GetCleanTextFromHtml(nodeList[x].SelectNodes(".//td")[7].InnerText);
                        isinList.Add(temp);
                    }
                }
            }
            return isinList;
        }
        private static List<ISINTemp> getISINListFromSinglePage(string uri, string postData)
        {
            List<ISINTemp> isinList = new List<ISINTemp>();
            AdvancedWebClient wc = new AdvancedWebClient();
            HtmlDocument htc = new HtmlDocument();
            string pageSource = WebClientUtil.GetPageSource(wc, uri, 300000, postData);
            isinList = getISINListFromSinglePage(pageSource);
            return isinList;
        }

        //baseUrl ="http://isin.krx.co.kr/jsp/realBoard07.jsp";
        public static List<ISINTemp> getISINListFromISINWebPage(ISINQuery query)
        {
            List<ISINTemp> isinList = new List<ISINTemp>();
            //string startDate = "";
            //string endDate = "";

            //if (query.StartDate != null)
            //{
            //    startDate = query.StartDate.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US"));

            //}

            //if(query.EndDate!=null)
            //{
            //    endDate = query.EndDate.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US"));
            //}
            try
            {
                int pageCount = 0;
                String issuername = HttpUtility.UrlEncode(query.IssueCompany, Encoding.GetEncoding("euc-kr"));
                String num = HttpUtility.UrlEncode(query.Code.EndsWith("호") ? query.Code : string.Format(query.Code, "호"), Encoding.GetEncoding("euc-kr"));
                string postData = string.Format("kind=&ef_std_cd_grnt_dt_from={0}&ef_std_cd_grnt_dt_to={1}&secuGubun={2}&lst_yn_all=on&lst_yn1=Y&lst_yn2=N&lst_yn3=D&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word={3}&ef_iss_inst_cd=&ef_isu_nm={4}&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=",
                                                query.StartDate, query.EndDate, query.Category, issuername, num);

                //String uri = "http://isin.krx.co.kr/jsp/BA_LT113.jsp"; Website Change
                string uri = string.Format("http://isin.krx.co.kr/jsp/realBoard{0}.jsp", query.Category);
                postData = "kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun=06&lst_yn_all=on&lst_yn1=Y&lst_yn2=N&lst_yn3=D&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word=&ef_iss_inst_cd=&ef_isu_nm=%B4%EB%BE%E7%B1%DD%BC%D3+2WR&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=";
                AdvancedWebClient wc = new AdvancedWebClient();
                string pageSource = WebClientUtil.GetPageSource(wc, uri, 300000, postData);
                pageCount = GetTotalPageCount(postData);
                isinList.AddRange(getISINListFromSinglePage(pageSource));

                if (pageCount == 0)
                {
                    pageCount += 1;
                }
                for (var i = 2; i <= pageCount; i++)
                {
                    postData = string.Format("pg_no={0}&lst_yn1=Y&lst_yn2=N&lst_yn3=D&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word={1}&ef_isu_nm={2}&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=&ef_std_cd_grnt_dt_from={3}&ef_std_cd_grnt_dt_to={4}", i.ToString("D2"), issuername, num, query.StartDate, query.EndDate);
                    HtmlDocument htc = new HtmlDocument();
                    pageSource = WebClientUtil.GetPageSource(uri, 300000, postData);
                    isinList.AddRange(getISINListFromSinglePage(pageSource));
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in GrabDataFromWebpage()    : \r\n" + ex.ToString();
            }
            return isinList;
        }

        private static int GetTotalPageCount(string pageSource)
        {
            int pageCount = 0;
            HtmlDocument htc = new HtmlDocument();
            try
            {
                if (!String.IsNullOrEmpty(pageSource))
                    htc.LoadHtml(pageSource);
                if (htc != null)
                {
                    HtmlNode node = htc.DocumentNode.SelectSingleNode("//table/tr/td[@class='contents']/font");

                    String totalPage = String.Empty;
                    if (node != null)
                        totalPage = node.InnerText.ToString().Split('/')[1].Trim(new Char[] { ' ', ';', ')' });
                    if (!String.IsNullOrEmpty(totalPage))
                        pageCount = Convert.ToInt32(totalPage);
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in GetTotalPageCount()   : \r\n" + ex.ToString();
            }
            return pageCount;
        }

        public void KillExcelProcess(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            IntPtr t = new IntPtr(excelApp.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();
        }

        public int CompareTicker(WarrantTemplate x, WarrantTemplate y)
        {
            if (x.Ticker == null)
            {
                if (y.Ticker == null)
                    return 0;
                else
                    return -1;
            }
            else
            {
                if (y.Ticker == null)
                    return 1;
                else
                {
                    int retval = x.Ticker.CompareTo(y.Ticker);
                    if (retval != 0)
                        return retval;
                    else
                        return x.Ticker.CompareTo(y.Ticker);
                }
            }
        }

        //true = use automatically detect settings
        //false = use proxy server 10.40.14.23 : 80
        public void SetUpAutoConfigScript(bool enable)
        {
            //enable automatically detect settings
            if (enable)
            {
                RegistryKey rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings", true);
                rk.SetValue("ProxyEnable", 0);
                rk.SetValue("AutoConfigURL", "http://wtd.apac.ime.truters.com/proxy.pac");
                rk.Close();
            }
            //enable proxy server
            else
            {
                RegistryKey rk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings", true);
                rk.SetValue("ProxyServer", "10.40.14.23:80");
                rk.SetValue("ProxyEnable", 1);
                rk.SetValue("AutoConfigURL", "<local>");
                rk.Close();
            }
        }


        public static void DeleteUselessFileAndFolder(string dir)
        {

            try
            {
                if (Directory.Exists(dir))
                {
                    if (Directory.GetDirectories(dir).Length == 0 && Directory.GetFiles(dir).Length == 0)
                    {
                        Directory.Delete(dir);
                        return;
                    }
                    foreach (string var in Directory.GetDirectories(dir))
                    {
                        DeleteUselessFileAndFolder(var);
                    }
                    foreach (string var in Directory.GetFiles(dir))
                    {

                        File.SetAttributes(var, FileAttributes.Normal);
                        File.Delete(var);
                    }
                    Directory.Delete(dir);
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in Delete()   : \r\n" + ex.ToString();
                throw;
            }
        }

        //public static void DeleteUselessFileAndFolder(String ipath)
        //{
        //    if (!String.IsNullOrEmpty(ipath))
        //    {
        //        String extension = Path.GetExtension(ipath);
        //        if (!String.IsNullOrEmpty(extension))
        //        {
        //            DeleteFile(ipath);
        //        }
        //        else
        //        {
        //            DeleteFolder(ipath);
        //        }
        //    }
        //}

        //private static void DeleteFile(String ipath)
        //{
        //    if (File.Exists(ipath))
        //    {
        //        File.Delete(ipath);
        //    }
        //}

        //private static void DeleteFolder(String ipath)
        //{
        //    if (Directory.Exists(ipath))
        //    {
        //        Delete(ipath);
        //    }
        //}

        //private static void Delete(String path)
        //{
        //    try
        //    {
        //        String[] collection = Directory.GetFileSystemEntries(path);
        //        int count = collection.Length;
        //        if (count > 0)
        //        {
        //            foreach (var item in collection)
        //            {
        //                if (File.Exists(item))
        //                {
        //                    File.Delete(item);
        //                }
        //                else
        //                {
        //                    if (Directory.Exists(item))
        //                    {
        //                        //Delete(item);
        //                        Directory.Delete(item);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        String msg = "Error found in Delete()   : \r\n" + ex.ToString();
        //        throw;
        //    }
        //}

        public static void CreateDirectory(String ipath)
        {
            if (!String.IsNullOrEmpty(ipath))
            {
                if (!Directory.Exists(ipath))
                {
                    Directory.CreateDirectory(ipath);
                }
                //else
                //{
                //    DeleteFolder(ipath);
                //    CreateDirectory(ipath);
                //}
            }
        }

        /// <summary>
        /// Get company infomation from page: companysummary.do
        /// </summary>
        /// <param name="attribute">onclick attribute of node</param>
        /// <returns>company infomation</returns>
        public static KoreaCompany GetCompanyInfoBak(string attribute)
        {
            KoreaCompany company = null;
            try
            {
                //http://kind.krx.co.kr/common/companysummary.do?method=searchCompanySummary&strIsurCd=00347&lstCd=undefined
                string url = string.Format("http://kind.krx.co.kr/common/companysummary.do?method=searchCompanySummary&strIsurCd={0}&lstCd=undefined", attribute);
                HtmlDocument doc = WebClientUtil.GetHtmlDocument(url, 300000, null);

                if (doc != null)
                {
                    company = new KoreaCompany();
                    HtmlNode table = doc.DocumentNode.SelectNodes("//table")[0];
                    string koreaName = table.SelectSingleNode(".//tr[1]/td[1]").InnerText.Trim();
                    string legalName = table.SelectSingleNode(".//tr[1]/td[2]").InnerText.Trim();
                    string isin = table.SelectSingleNode(".//tr[2]/td[1]").InnerText.Trim(); ;
                    string market = table.SelectSingleNode(".//tr[2]/td[2]").InnerText.Trim();
                    market = GetMarketCode(market);
                    string listingDate = table.SelectSingleNode(".//tr[3]/td[2]").InnerText.Trim();

                    if (!string.IsNullOrEmpty(market))
                    {
                        company.Market = market;
                    }
                    if (!string.IsNullOrEmpty(isin))
                    {
                        company.ISIN = isin;
                    }
                    if (!string.IsNullOrEmpty(legalName))
                    {
                        company.LegalName = legalName;
                    }
                    if (!string.IsNullOrEmpty(koreaName))
                    {
                        company.KoreaName = koreaName;
                    }
                    if (!string.IsNullOrEmpty(listingDate))
                    {
                        company.ListingDate = listingDate;
                    }
                }
            }
            catch
            {
                return null;
            }
            return company;
        }

        public static KoreaCompany GetCompanyInfo(string attribute)
        {

            KoreaCompany company = null;
            try
            {

                string uri = "http://kind.krx.co.kr/common/companysummary.do";
                string companyPostData = string.Format("method=searchCompanySummaryOvrvwDetail&menuIndex=0&strIsurCd={0}&lstCd=undefined&taskDd=&spotIsuTrdMktTpCd=&methodType=0", attribute);

                string companyPageSource = null;
                int retry = 3;

                while (string.IsNullOrEmpty(companyPageSource) && retry-- > 0)
                {
                    try
                    {
                        companyPageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, companyPostData);
                    }
                    catch
                    {
                        System.Threading.Thread.Sleep(3000);
                    }
                }

                if (string.IsNullOrEmpty(companyPageSource))
                { 
                    return null;
                }

                companyPageSource = MiscUtil.GetCleanTextFromHtml(companyPageSource);
                HtmlDocument companyDoc = new HtmlDocument();
                if (!string.IsNullOrEmpty(companyPageSource))
                    companyDoc.LoadHtml(companyPageSource);


                if (companyDoc != null)
                {
                    company = new KoreaCompany();
                    HtmlNode table = companyDoc.DocumentNode.SelectNodes("//table")[0];
                    string koreaName = table.SelectSingleNode(".//tr[1]/td[1]").InnerText.Trim();
                    string legalName = table.SelectSingleNode(".//tr[1]/td[2]").InnerText.Trim();
                    string isin = table.SelectSingleNode(".//tr[2]/td[1]").InnerText.Trim(); ;
                    string market = table.SelectSingleNode(".//tr[3]/td[2]").InnerText.Trim();
                    market = GetMarketCode(market);
                    string listingDate = table.SelectSingleNode(".//tr[3]/td[1]").InnerText.Trim();

                    if (!string.IsNullOrEmpty(market))
                    {
                        company.Market = market;
                    }
                    if (!string.IsNullOrEmpty(isin))
                    {
                        company.ISIN = isin;
                    }
                    if (!string.IsNullOrEmpty(legalName))
                    {
                        company.LegalName = legalName;
                    }
                    if (!string.IsNullOrEmpty(koreaName))
                    {
                        company.KoreaName = koreaName;
                    }
                    if (!string.IsNullOrEmpty(listingDate))
                    {
                        company.ListingDate = listingDate;
                    }
                }
            }
            catch
            {
                return null;
            }
            return company;

        }

        /// <summary>
        /// Change exchange board code from exchange board name.
        /// </summary>
        /// <param name="MarketName">name</param>
        /// <returns>exchange board code</returns>
        public static string GetMarketCode(string MarketName)
        {
            if (MarketName.Contains("코스닥"))
            {
                return "KQ";
            }
            else if (MarketName.Contains("유가증권"))
            {
                return "KS";
            }
            else if (MarketName.Contains("코넥스"))
            {
                return "KN";
            }
            return null;
        }

        /// <summary>
        /// ISIN Website. Get different table by isin.
        /// </summary>
        /// <param name="isin">isin</param>
        /// <returns>table html node</returns>
        public static HtmlNode GetTargetTableByIsin(string isin)
        {
            string url = GetUrlByIsin(isin);
            return GetTargetTableByUrl(url);
        }

        /// <summary>
        /// ISIN Website. Get different table by url.
        /// </summary>
        /// <param name="url">url</param>
        /// <returns>table html node</returns>
        public static HtmlNode GetTargetTableByUrl(string url)
        {
            if (string.IsNullOrEmpty(url))
            {
                return null;
            }
            string pageSource = null;
            int retry = 3;
            while (pageSource == null && retry-- > 0)
            {
                try
                {
                    pageSource = WebClientUtil.GetPageSource(null, url, 18000, "", Encoding.GetEncoding("EUC-KR"));
                }
                catch
                {
                    continue;
                }
            }
            if (pageSource == null)
            {
                return null;
            }
            HtmlDocument isinRoot = new HtmlDocument();
            isinRoot.LoadHtml(pageSource);

            if (isinRoot == null)
            {
                return null;
            }

            HtmlNode isinTable = isinRoot.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[1]/tr[3]/td[1]/table[1]");
            if (isinTable == null)
            {
                return null;
            }
            return isinTable;
        }

        /// <summary>
        /// ISIN Website. Get target url by isin.
        /// </summary>
        /// <param name="isin">isin</param>
        /// <returns>url</returns>
        private static string GetUrlByIsin(string isin)
        {
            string postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun=99&lst_yn_all=on&lst_yn1=Y&lst_yn2=N&lst_yn3=D&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word=&ef_iss_inst_cd=&ef_isu_nm={0}&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=", isin);
            string uri = "http://isin.krx.co.kr/jsp/realBoard99.jsp";
            string pageSource = null;
            int retries = 3;
            while (pageSource == null && retries-- > 0)
            {
                try
                {
                    AdvancedWebClient wc = new AdvancedWebClient();
                    pageSource = WebClientUtil.GetPageSource(wc, uri, 180000, postData, Encoding.GetEncoding("EUC-KR"));
                }
                catch
                {
                    System.Threading.Thread.Sleep(5000);
                }
            }
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(pageSource);

            HtmlNodeCollection trs = null;
            try
            {
                trs = doc.DocumentNode.SelectNodes("/html[1]/body[1]/table[1]/tr[2]/td[1]/table[1]/tr");
            }
            catch
            {
                return null;
            }
            HtmlNode tr = trs[1];
            HtmlNode td = tr.SelectNodes("./td")[1];
            HtmlNode aNode = td.SelectSingleNode(".//a");
            string checkMod = aNode.Attributes["href"].Value.Trim();
            string type = checkMod.Split(',')[2].Trim(new char[] { '\'', ')', ';' });
            string targetPage = GetTargetPageCode(type);
            string url = string.Format("http://isin.krx.co.kr/jsp/{0}?isu_cd={1}&modi=f&req_no=", targetPage, isin);
            return url;
        }

        /// <summary>
        /// ISIN Website. Get different url part by instrument type.
        /// </summary>
        /// <param name="type">instrument type</param>
        /// <returns>url part</returns>
        public static string GetTargetPageCode(string type)
        {
            string url = string.Empty;
            if (type == "BND")
            { // 국채,지방채, 특수채, 회사채, 외국채권
                url = "BA_VW011.jsp";
            }
            else if (type == "BC")
            { // 수익증권/개발시탁
                url = "BA_VW012.jsp";
            }
            else if (type == "BW")
            { // 신주인수권
                url = "BA_VW016.jsp";
            }
            else if (type == "ELW")
            { // 주식워런트
                url = "BA_VW021.jsp";
            }
            else if (type == "STK")
            { // 주권
                url = "BA_VW010.jsp";
            }
            else if (type == "CP")
            { // CP
                url = "BA_VW014.jsp";
            }
            else if (type == "DR")
            { // 예탁증서
                url = "BA_VW013.jsp";
            }
            else if (type == "INDX")
            { // 지표
                url = "BA_VW017.jsp";
            }
            else if (type == "STRIP")
            { // 스트립
                url = "BA_VW015.jsp";
            }
            else if (type == "CD")
            { // CD
                url = "BA_VW020.jsp";
            }
            else if (type == "FUT_OPT")
            { // 선물옵션
                url = "BA_VW019.jsp";
            }
            else if (type == "SHORTBOND")
            { // 전자단기채권
                url = "BA_VW022.jsp";
            }
            return url;
        }

        /// <summary>
        /// ISIN Website. Get target page source by isin.
        /// </summary>
        /// <param name="isin">isin</param>
        /// <returns>page source</returns>
        public static string GetTargetPageSourceByIsin(string isin)
        {
            string url = GetUrlByIsin(isin);
            if (string.IsNullOrEmpty(url))
            {
                return null;
            }
            string pageSource = null;
            int retry = 3;
            while (pageSource == null && retry-- > 0)
            {
                try
                {
                    pageSource = WebClientUtil.GetPageSource(null, url, 18000, "", Encoding.GetEncoding("EUC-KR"));
                }
                catch
                {
                    continue;
                }
            }
            return pageSource;
        }

        public static KoreaNames GetKoreaNamesByISIN(string isin)
        {
            try
            {
                string url = GetUrlByIsin(isin);
                if (!(url.Contains("BA_VW010.jsp") || url.Contains("BA_VW012.jsp") || url.Contains("BA_VW013.jsp")))
                {
                    return null;
                }

                HtmlNode table = GetTargetTableByUrl(url);
                if (table == null)
                {
                    return null;
                }
                HtmlNodeCollection isinTrs = table.SelectNodes("./tr");
                if (isinTrs == null)
                {
                    return null;
                }
                string koreaName = string.Empty;
                string legalName = string.Empty;
                string displayName = string.Empty;
                KoreaNames names = new KoreaNames();
                if (url.Contains("BA_VW010.jsp")) //ORD PRF ETF REIT
                {
                    koreaName = isinTrs[9].SelectNodes("./td")[1].InnerText.Trim();
                    if (koreaName.Contains("보통주"))
                    {
                        koreaName = koreaName.Replace("보통주", "");
                    }
                    if (koreaName.Contains("상장지수투자신탁"))
                    {
                        koreaName = koreaName.Substring(0, koreaName.IndexOf("상장지수투자신탁"));
                    }

                    if (koreaName.Length > 14)
                    {
                        koreaName = isinTrs[9].SelectNodes("./td")[3].InnerText.Trim();
                    }
                    legalName = isinTrs[10].SelectNodes("./td")[1].InnerText.Trim();
                    displayName = isinTrs[10].SelectNodes("./td")[3].InnerText.Trim();
                }
                else  //BC //KDR
                {
                    koreaName = isinTrs[3].SelectNodes("./td")[1].InnerText.Trim();
                    if (koreaName.Length > 14)
                    {
                        koreaName = isinTrs[3].SelectNodes("./td")[3].InnerText.Trim();
                    }
                    legalName = isinTrs[4].SelectNodes("./td")[1].InnerText.Trim();
                    displayName = isinTrs[4].SelectNodes("./td")[3].InnerText.Trim();
                }

                names.KoreaName = koreaName;
                names.LegalName = legalName;
                names.IDNDisplayName = displayName;
                return names;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Get PEO Add type by ISIN. ORD/KDR/PRF 
        /// </summary>
        /// <param name="isin">ISIN</param>
        /// <returns>PEO type</returns>
        public static string GetPeoTypeByISIN(string isin)
        {
            string uri = "http://isin.krx.co.kr/jsp/realBoard99.jsp";
            string postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun=99&lst_yn_all=on&lst_yn1=Y&lst_yn2=N&lst_yn3=D&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word=&ef_iss_inst_cd=&ef_isu_nm={0}&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=", isin);
            string pageSource = null;
            int retries = 3;
            while (pageSource == null && retries-- > 0)
            {
                try
                {
                    pageSource = WebClientUtil.GetPageSource(null, uri, 180000, postData, Encoding.GetEncoding("EUC-KR"));
                }
                catch
                {
                    System.Threading.Thread.Sleep(5000);
                }
            }

            if (pageSource == null)
            {
                return null;
            }

            string peoType = string.Empty;
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(pageSource);
            HtmlNodeCollection trs = null;
            try
            {
                trs = doc.DocumentNode.SelectNodes("/html[1]/body[1]/table[1]/tr[2]/td[1]/table[1]/tr");
                HtmlNode tr = trs[1];
                HtmlNodeCollection tds = tr.SelectNodes("./td");
                string companyName = tds[2].InnerText.Trim();
                if (companyName.Contains("보통주"))
                {
                    peoType = "ORD";
                }
                else if (companyName.Contains("우선주"))
                {
                    peoType = "PRF";
                }
                else
                {
                    string instrumentType = tds[3].InnerText.Trim();
                    if (instrumentType.Equals("예탁증서"))
                    {
                        peoType = "KDR";
                    }
                    else
                    {
                        peoType = @"#N/A";
                    }
                }
            }
            catch
            {
                string msg = "At GetPeoTypeByISIN(string isin). Error found in searching peo type infos for " + isin;
            }
            return peoType;
        }


        /*==========================================================================================================*/

        #region useless code

        /*public void CreateDir(String fullPath)
        {
            if (!Directory.Exists(fullPath))
            {
                DirectoryInfo dir = new DirectoryInfo(fullPath);
                dir.Create();
            }
            else
            {
                DeleteTempDir(fullPath);
                DirectoryInfo dir = new DirectoryInfo(fullPath);
                dir.Create();
            }
        }*/

        /*public void DeleteTempDir(String dir)
        {

            try
            {
                if (Directory.GetDirectories(dir).Length == 0 && Directory.GetFiles(dir).Length == 0)
                {
                    Directory.Delete(dir);
                    return;
                }

                string[] dirArr = Directory.GetDirectories(dir);
                foreach (string var in dirArr  )//Directory.GetDirectories(dir)
                {
                    DeleteTempDir(var);
                }

                string[] fileArr = Directory.GetFiles(dir);
                foreach (string var in fileArr )//Directory.GetFiles(dir)
                {
                    if (var != "C:\\Korea FM\\ELW FM1\\Korea ELW FM Masterfile.xls")
                    {
                        File.SetAttributes(var, FileAttributes.Normal);
                        File.Delete(var);
                    }
                    else
                        continue;
                }
                //Directory.Delete(dir);
            }
            catch (Exception ex)
            {
                String errInfo = ex.ToString();
            }
        }*/

        /*public String PDFDownload(String url, String ricCode)
        {
            DateTime today = DateTime.Today;
            WebClient pdfClient = new WebClient();
            String pdfFilePath = pdf_path + "\\" + subFolder + "\\" + ricCode + ".pdf";
            try
            {
                pdfClient.DownloadFile(url, pdfFilePath);
                return PDFToTxt(ricCode, pdfFilePath);
            }
            catch (Exception ex)
            {
                String logerror = ex.Message.ToString();
                return "PDFDownload Error";
            }
        }*/

        /*private String PDFToTxt(String ricCode, String pdfFilePath)
        {
            String command = "pdftotext.exe";
            String txtPath = pdf_path + "\\" + subFolder + "\\" + ricCode + ".txt";
            String parameters = "-layout -enc UTF-8 -q " + pdfFilePath + " " + txtPath;
            System.Diagnostics.Process.Start(command, parameters);
            return txtPath;
        }*/

        #endregion
    }
}
