using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Selenium;
using System.Text.RegularExpressions;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Drawing;
using System.Threading;
using System.Web;
using HtmlAgilityPack;
using Ric.Db.Info;
using Ric.Db.Manager;
using System.IO;
using System.Net;
using Ric.Db.Config;
using MySql.Data.MySqlClient;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
{
    class ADD : GeneratorBase
    {
        private List<KoreaEquityInfo> paList = new List<KoreaEquityInfo>();
        private List<KoreaEquityInfo> raList = new List<KoreaEquityInfo>();
        private List<KoreaEquityInfo> eaList = new List<KoreaEquityInfo>();
        private List<KoreaEquityInfo> bcList = new List<KoreaEquityInfo>();
        private Hashtable equityDate = new Hashtable();

        private Hashtable taskResult = new Hashtable();
        private KOREA_ADDGeneratorConfig configObj = null;
        private List<string> errors = new List<string>();
        private bool peoChanged = false;

        #region [CoraxKoreaIPO]
        private ExtendDataContext extendContext = null;
        private List<KORProcessItem> historicalAnnouncements = null;
        private Dictionary<string, List<string>> dicListXls1 = null;//xlsx file 
        private Dictionary<string, List<string>> dicListXls2 = null;//xlsx file 
        private List<string> listTitle1 = null;
        private List<string> listTitle2 = null;
        private string file1 = string.Empty;
        private string file2 = string.Empty;

        private string sourcingEngine = string.Empty;
        private string marketName = string.Empty;
        private string startTime = string.Empty;
        private string endTime = string.Empty;
        private string outPutPath = string.Empty;
        #endregion

        protected override void Start()
        {
            StartCoraxKoreaIpoJob();
            StartADDJob();    //  Read from this part
        }

        private void StartCoraxKoreaIpoJob()
        {
            GetDownloadListFromDB(historicalAnnouncements);

            if (historicalAnnouncements == null || historicalAnnouncements.Count == 0)
            {
                Logger.Log("No Data In DB!", Logger.LogType.Info);
                return;
            }

            dicListXls1 = ExtractData1(historicalAnnouncements);
            dicListXls2 = ExtractData2(historicalAnnouncements);
            GenerateFile(dicListXls1, file1);
            AddResult("CoraxKoreaIPO", file1, "Issue_Details_NV_VR_Bulkloader");
            GenerateFile(dicListXls2, file2);
            AddResult("CoraxKoreaIPO", file2, "TSO_Bulkloader");
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREA_ADDGeneratorConfig;

            #region [CoraxKoreaIPO]
            historicalAnnouncements = new List<KORProcessItem>();
            extendContext = new ExtendDataContext(new MySqlConnection(AceDbConfig.DataSourceStringDeals));
            sourcingEngine = "CORAX KOR";
            marketName = "Korea";
            startTime = configObj.StartDate;
            endTime = DateTime.Parse(configObj.EndDate).AddDays(+1).ToString("yyyy-MM-dd");
            listTitle1 = new List<string>() { "RIC", "CHANGE_DATE", "NOMINAL_VALUE", "NO_PAR_VALUE", "CLA_CUR_VAL", "CAC_MA_COMMENTS", "VOTING_RIGHTS_PER_SHARE", "VRI_CHANGE_DATE", "VOTING_RIGHTS_DESCRIPTION", "CONVERSION_RATIO", "CRA_CHANGE_DATE", "SOURCE_ID", "SOURCE_TYPE", "RELEASE_DATE", "LOCAL_DATE", "TIMEZONE_NAME", "SOURCE_PROVIDER", "SOURCE_LINK", "SOURCE_DESCRIPTION" };//file 1
            listTitle2 = new List<string>() { "RIC", "TSO_TYPE", "TSO_DATE", "ISSUE_SPEC_SHARES", "UNLISTED_SHARES", "CAC_MA_COMMENTS", "EXTERNAL_DESCRIPTION", "SOURCE_ID", "SOURCE_TYPE", "RELEASE_DATE", "LOCAL_DATE", "TIMEZONE_NAME", "SOURCE_PROVIDER", "SOURCE_LINK", "SOURCE_DESCRIPTION" };
            outPutPath = configObj.CoraxKoreaBulkFile;
            file1 = Path.Combine(outPutPath, string.Format(@"{0}\Issue_Details_NV_VR_Bulkloader.csv", DateTime.Now.ToString("yyyy-MM-dd")));
            file2 = Path.Combine(outPutPath, string.Format(@"{0}\TSO_Bulkloader.csv", DateTime.Now.ToString("yyyy-MM-dd")));
            #endregion
        }

        //from  this part  wzy 
        public void StartADDJob()
        {
            try
            {
                //Sourcing 
                int count = GrabDataFromKindWebSite();

                if (count == 0)
                {
                    return;
                }

                //Formatting
                FormatPEOADDTemplate();
                FormatETFADDTemplate();
                FormatREITADDTemplate();
                //wzy:format?

                count = paList.Count + eaList.Count + raList.Count + bcList.Count;
                if (count == 0)
                {
                    return;
                }

                //Distributing
                GenerateFMFilesByGroup();
                GenerateGEDAFiles();
                GenerateNDAFiles();
                GenerateNDATickLotFiles();
                // GenerateNDAFutureDating();


                CreateTaskResult();
                DeletePeoFmOne();
            }
            catch (Exception ex)
            {
                string msg = "At StartADDJob()." + ex.Message + "\r\n" + ex.StackTrace;
                errors.Add(msg);
            }
            finally
            {
                LogErrors();
            }
        }

        #region Sourcing

        /// <summary>
        /// Grab data from kind web site.
        /// </summary>
        private int GrabDataFromKindWebSite()
        {
            string startDate = configObj.StartDate.Trim();
            string endDate = configObj.EndDate.Trim();
            if (string.IsNullOrEmpty(startDate))
                startDate = DateTime.Today.ToString("yyyy-MM-dd");
            if (string.IsNullOrEmpty(endDate))
                endDate = DateTime.Today.ToString("yyyy-MM-dd");

            if ((DateTime.Parse(startDate)).CompareTo(DateTime.Parse(endDate)) > 0)
            {
                string temp = startDate;
                startDate = endDate;
                endDate = temp;
            }

            GrabAddData(startDate, endDate);
            GrabPEONewListing(startDate, endDate);
            int count = paList.Count + eaList.Count + raList.Count + bcList.Count;
            string dataMsg = string.Format("From {0} to {1}, we grabbed {2} announcements. As follows:\r\n\t\t\t\t\t\t" +
                                            "PEO: {3}\tETF: {4}\tREIT: {5}\t BC: {6}", startDate, endDate, count.ToString(), paList.Count, eaList.Count, raList.Count, bcList.Count);
            Logger.Log(dataMsg, Logger.LogType.Info);
            return count;
        }

        #region Grab IPO data

        #region Common Methods

        /// <summary>
        /// Grab add data use the key word: IPO(신규상장)
        /// </summary>
        /// <param name="startDate">start date</param>
        /// <param name="endDate">end date</param>
        private void GrabAddData(string startDate, string endDate)
        {
            string dataStartDate = null;
            DateTime startDay = DateTime.Parse(startDate);
            DateTime twoMonthEarly = DateTime.Parse(endDate).AddMonths(-2);
            if (startDay.CompareTo(twoMonthEarly) < 0)
            {
                dataStartDate = startDate;
            }
            else
            {
                dataStartDate = twoMonthEarly.ToString("yyyy-MM-dd");
            }

            string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%8B%A0%EA%B7%9C%EC%83%81%EC%9E%A5&repIsuSrtCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=300&scn=mktact&srchFd=2&kwd=%EC%8B%A0%EA%B7%9C%EC%83%81%EC%9E%A5&fromData={0}&toData={1}", dataStartDate, endDate);
            string uri = "http://kind.krx.co.kr/disclosure/searchtotalinfo.do";

            HtmlDocument htc = new HtmlDocument();
            string pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, postData);
            if (!string.IsNullOrEmpty(pageSource))
                htc.LoadHtml(pageSource);

            if (htc != null)
            {
                HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes("//dl/dt");
                HtmlNodeCollection ddCollections = htc.DocumentNode.SelectNodes("//dl/dd");
                int count = nodeCollections.Count;

                for (var i = 0; i < count; i++)
                {
                    HtmlNode ddNode = ddCollections[i].SelectSingleNode(".//span");
                    HtmlNode node = nodeCollections[i].SelectSingleNode(".//span/a");
                    string title = string.Empty;
                    if (node != null)
                        title = node.InnerText.Trim().ToString();
                    HtmlNode nodeDate = nodeCollections[i].SelectSingleNode("./em");
                    if (nodeDate != null)
                    {
                        DateTime anouncementDate = new DateTime();
                        anouncementDate = DateTime.Parse(nodeDate.InnerText.Trim(), new CultureInfo("en-US"));
                        if (anouncementDate < startDay)
                        {
                            return;
                        }
                    }
                    if (!string.IsNullOrEmpty(title))
                    {
                        HtmlDocument doc = null;
                        string KSorKQorKN = string.Empty;

                        if (title.Contains("신규상장"))
                        {
                            //상장지수투자신탁(ETF)             
                            //위탁관리부동산투자회사(REIT)
                            //수익증권신규상장(BC)
                            //first 4 chars 신규상장(PEO)
                            //상장지수자투자신탁[주식-파생형], 상장일 2013.1.21(월)

                            if (title.Replace(" ", "").Contains("상장지수투자신탁") || title.Replace(" ", "").Contains("상장지수자투자신탁"))
                            {
                                KSorKQorKN = "KS";
                                doc = GetTargetPageDocument(node);
                                bool isGlobal = JudgeIfGloblaETF(nodeDate.InnerText.Trim(), title);
                                GrabETFDataAction(doc, isGlobal);
                            }
                            else if (title.Replace(" ", "").Contains("부동산투자회사"))
                            {
                                KSorKQorKN = "KS";
                                doc = GetTargetPageDocument(node);
                                GrabREITsDataAction(doc);
                            }
                            else if (title.Replace(" ", "").Contains("수익증권신규상장"))
                            {
                                KSorKQorKN = "KS";
                                doc = GetTargetPageDocument(node);
                                GrabBCDataAction(doc);
                            }
                            else
                            {
                                if (!title.Substring(0, "신규상장".Length).Equals("신규상장")) continue;

                                KSorKQorKN = GetDutyCode(ddNode);
                                doc = GetTargetPageDocument(node);
                                if (KSorKQorKN.Equals("KS"))
                                {
                                    HtmlNode companyNode = nodeCollections[i].SelectSingleNode(".//strong/a");
                                    KoreaCompany company = GetCompanyInfo(companyNode);
                                    GrabKSPeoData(doc, company);
                                }
                                else
                                {
                                    GrabKQDataAction(doc, KSorKQorKN);
                                    //  GrabPeoKQKNData(doc, KSorKQorKN, company);
                                }
                            }
                        }
                        else if (title.Replace(" ", "").Contains("상장지수투자신탁"))
                        {
                            HtmlNode companyNode = node.SelectSingleNode(".//strong/a");
                            KoreaCompany company = GetCompanyInfo(companyNode);
                            KSorKQorKN = "KS";
                            doc = GetTargetPageDocument(node);
                            bool isGlobal = JudgeIfGloblaETF(nodeDate.InnerText.Trim(), title);
                            GrabETFDataAction(doc, isGlobal);
                        }
                        else if (title.Replace(" ", "").Contains("부동산투자회사"))
                        {
                            HtmlNode companyNode = node.SelectSingleNode(".//strong/a");
                            KoreaCompany company = GetCompanyInfo(companyNode);
                            KSorKQorKN = "KS";
                            doc = GetTargetPageDocument(node);
                            GrabREITsDataAction(doc);
                        }
                        else
                            continue;
                    }
                }
            }
        }

        private KoreaCompany GetCompanyInfo(HtmlNode companyNode)
        {
            try
            {
                string judgeText = companyNode.Attributes["onclick"].Value.Trim().ToString();
                judgeText = judgeText.Split('(')[1].Split(')')[0].Trim(new char[] { '\'', ' ' }).ToString();
                KoreaCompany company = Common.GetCompanyInfo(judgeText);
                company.KoreaName = company.KoreaName.Replace("(주)", "");
                return company;
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GetCompanyInfo(). Error found when getting company information. Error message: {0}.", ex.Message);
                errors.Add(msg);
                return null;
            }

        }

        /// <summary>
        /// Get the onclick target page document.
        /// </summary>
        /// <param name="node">title node</param>
        /// <param name="attribute"></param>
        /// <returns></returns>
        private HtmlDocument GetTargetPageDocument(HtmlNode node)
        {
            HtmlDocument doc = null;
            try
            {
                string attribute = node.Attributes["onclick"].Value.Trim().ToString();
                if (!string.IsNullOrEmpty(attribute))
                    attribute = attribute.Split('(')[1].Split(')')[0].Trim(new Char[] { ' ', ',', '\'' }).ToString();
                string param = attribute.Split(',')[1].Trim(new Char[] { ' ', '\'', ',' }).ToString();
                string url = GetTargetUrl(param);
                doc = WebClientUtil.GetHtmlDocument(url, 600000);
            }
            catch (Exception ex)
            {
                string msg = "" + ex.StackTrace + "  : ----> \n\r" + ex.ToString();
                errors.Add(msg);
            }
            return doc;
        }

        /// <summary>
        /// Get target page's url
        /// </summary>
        /// <param name="docno">doc no get from attribute</param>
        /// <returns>url</returns>
        private string GetTargetUrl(string docno)
        {
            try
            {
                string url = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=searchContents&docNo={0}", docno);
                string pageSource = WebClientUtil.GetPageSource(null, url, 180000, null, Encoding.GetEncoding("EUC-KR"));
                Regex regex = new Regex(@"http://kind.krx.co.kr/.+htm");
                Match match = regex.Match(pageSource);
                return match.Value;
            }
            catch (Exception ex)
            {
                string msg = "Error found in getting target url." + ex.Message;
                errors.Add(msg);
                return null;
            }
        }

        /// <summary>
        /// Fommat viriable string from given position
        /// </summary>
        /// <param name="pos">position</param>
        /// <param name="strPre">string to cut</param>
        /// <returns>formatted string</returns>
        private string FormatDataWithPos(int pos, string strPre)
        {
            char[] tempArr = strPre.ToCharArray();
            string result = "";
            while (tempArr[pos] != '\n')    //||tempArr[pos] != '\r'
            {
                result += tempArr[pos].ToString();
                if ((pos + 1) < tempArr.Length)
                    pos++;
                else
                    break;
            }
            result = result.Trim(new char[] { ' ', ':' });
            return result;
        }

        /// <summary>
        /// Format date
        /// </summary>
        /// <param name="effectiveDate"></param>
        private void FormatDate(ref string effectiveDate)
        {
            DateTime dt;
            if (DateTime.TryParse(effectiveDate, out dt))
            {
                effectiveDate = dt.ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
            }
            else
            {
                Regex regex = new Regex("[0-9]+");
                MatchCollection matches = regex.Matches(effectiveDate);
                if (matches.Count == 3)
                {
                    if (matches[0].Value.Length == 4)
                    {
                        effectiveDate = string.Format("{0}-{1}-{2}", matches[0].Value, matches[1].Value, matches[2].Value);
                    }
                    else
                    {
                        effectiveDate = string.Format("{2}-{1}-{0}", matches[0].Value, matches[1].Value, matches[2].Value);
                    }
                    effectiveDate = DateTime.Parse(effectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                }
                else
                {
                    effectiveDate = "1900-Jan-01";
                }
            }
        }

        #endregion

        #region Grab PEO data

        /// <summary>
        /// Get exchange board code from announcement.
        /// </summary>
        /// <param name="ddNode">announcement node</param>
        /// <returns>exchange board code</returns>
        private string GetDutyCode(HtmlNode ddNode)
        {
            string dutyName = ddNode.InnerText.Split(':')[1].Trim();
            if (dutyName.Contains("코스닥"))
            {
                return "KQ";
            }
            else if (dutyName.Contains("유가증권"))
            {
                return "KS";
            }
            else if (dutyName.Contains("코넥스"))
            {
                return "KN";
            }
            return "";
        }

        /// <summary>
        /// Grab PEO add KS data(data in text).
        /// </summary>
        /// <param name="doc">page document</param>
        private void GrabKSPeoData(HtmlDocument doc, KoreaCompany company)
        {
            try
            {
                HtmlNode pre = doc.DocumentNode.SelectSingleNode(".//pre");
                string document = pre.InnerText;

                int effectiveDatePos = document.IndexOf("상장일") + "상장일".Length;
                string effectiveDate = FormatDataWithPos(effectiveDatePos, document);

                if (!string.IsNullOrEmpty(effectiveDate))
                {
                    FormatDate(ref effectiveDate);
                }
                int pricePos = 0;
                string priceKey = string.Empty;
                string price = string.Empty;
                if (document.Contains("1주의 발행(공모)가액"))
                {
                    priceKey = "1주의 발행(공모)가액";
                }
                //1주의 발행가액(무액면)
                else if (document.Contains("1주의 액면가액(공모가액)"))
                {
                    priceKey = "1주의 액면가액(공모가액)";
                }

                else if (document.Contains("1주의 발행가액(무액면)"))
                {
                    priceKey = "1주의 발행가액(무액면)";
                }
                else if (document.Contains("1주의 액면가액"))
                {
                    priceKey = "1주의 액면가액";
                }

                if (!string.IsNullOrEmpty(priceKey))
                {
                    pricePos = document.IndexOf(priceKey) + priceKey.Length;
                    price = FormatDataWithPos(pricePos, document);
                    if (price.Contains("(") && price.Contains(")"))
                    {
                        int startPos = price.IndexOf("(");
                        int endPos = price.IndexOf(")");
                        price = price.Substring(0, startPos) + price.Substring(endPos + 1);
                    }

                    price = Regex.Replace(price, ",| ", "");

                    Regex priceRegex = new Regex("[0-9]+");
                    Match priceMatch = priceRegex.Match(price);
                    if (priceMatch.Success)
                    {
                        price = priceMatch.Value;
                    }
                }

                List<string> isins = new List<string>();
                List<string> tickers = new List<string>();

                string pattern = @":\W*(?<ISIN>\w+)\W*단축코드\W*(?<Ticker>\w+)";
                Regex r1 = new Regex(pattern);
                MatchCollection matches = r1.Matches(document);
                for (int i = 0; i < matches.Count; i++)
                {
                    Match m1 = matches[i];
                    string isinLine = m1.Groups["ISIN"].ToString().Trim();
                    string tickerLine = m1.Groups["Ticker"].ToString().Trim();
                    isins.Add(isinLine);
                    tickers.Add(tickerLine.Substring(1));
                }

                for (int i = 0; i < isins.Count; i++)
                {

                    string isinItem = isins[i];
                    KoreaEquityInfo peo = new KoreaEquityInfo();
                    peo.ISIN = isinItem;
                    peo.Ticker = tickers[i];
                    if (isinItem == company.ISIN)
                    {
                        peo.LegalName = company.LegalName;
                        peo.KoreaName = company.KoreaName;
                    }
                    else
                    {
                        int legalNamePos = document.IndexOf("영문종목명") + "영문종목명".Length;
                        string legalName = FormatDataWithPos(legalNamePos, document);
                        int koreaNamePos = document.IndexOf("한글종목명") + "영문종목명".Length;
                        string koreaName = FormatDataWithPos(koreaNamePos, document);
                        peo.LegalName = legalName;
                        peo.KoreaName = koreaName;
                    }
                    string idnDisplayName = string.Empty;
                    pattern = @"종목약명.*?영문.*";
                    r1 = new Regex(pattern);
                    Match match = r1.Match(document);
                    if (match.Success)
                    {
                        idnDisplayName = match.Value.Substring(match.Value.IndexOf("영문") + 2).TrimStart(new char[] { ' ', ')', ':', '(' });
                    }
                    peo.IDNDisplayName = idnDisplayName.ToUpper();

                    if ((isins.Count - i) > 1)
                    {
                        int endPos = document.IndexOf("종목약명") + "종목약명".Length;
                        document = document.Substring(endPos);
                    }

                    peo.EffectiveDate = effectiveDate;
                    peo.Type = GetPeoTypeByISIN(peo.ISIN);
                    peo.Lotsize = price;
                    peo.Market = "KS";
                    peo.RIC = peo.Ticker + ".KS";
                    paList.Add(peo);

                    // GetPeoDetailByISIN(isinItem, effectiveDate, company);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in grabbing KS data. GrabKSPeoData(HtmlDocument doc)" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Get PEO Add type by ISIN. ORD/KDR/PRF 
        /// </summary>
        /// <param name="isin">ISIN</param>
        /// <returns>PEO type</returns>
        private string GetPeoTypeByISIN(string isin)
        {
            //string uri = "http://isin.krx.co.kr/jsp/realBoard99.jsp";
            //string postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun=99&lst_yn_all=on&lst_yn1=Y&lst_yn2=N&lst_yn3=D&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word=&ef_iss_inst_cd=&ef_isu_nm={0}&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=", isin);

            if (!File.Exists(@"Config\Korea\ADDPostData.txt"))
                System.Windows.Forms.MessageBox.Show(string.Format("The file {0} missing.", @"Config\Korea\ADDPostData.txt"));

            string uri = @"http://isin.krx.co.kr/srch/srch.do?method=srchList";
            string postData = string.Format(File.ReadAllText(@"Config\Korea\ADDPostData.txt", Encoding.UTF8), isin);

            string pageSource = null;
            int retries = 3;
            while (pageSource == null && retries-- > 0)
            {
                try
                {
                    //pageSource = WebClientUtil.GetPageSource(null, uri, 180000, postData, Encoding.GetEncoding("EUC-KR"));
                    pageSource = GetPageSource(uri, postData);
                }
                catch
                {
                    System.Threading.Thread.Sleep(5000);
                }
            }

            if (pageSource == null)
            {
                string msg = "At GetPeoTypeByISIN(string isin). Cannot find peo type infos for " + isin;
                errors.Add(msg);
                return null;
            }

            string peoType = string.Empty;
            //HtmlDocument doc = new HtmlDocument();
            //doc.LoadHtml(pageSource);
            //HtmlNodeCollection trs = null;
            try
            {
                //trs = doc.DocumentNode.SelectNodes("/html[1]/body[1]/table[1]/tr[2]/td[1]/table[1]/tr");
                //HtmlNode tr = trs[1];
                //HtmlNodeCollection tds = tr.SelectNodes("./td");
                //string companyName = tds[2].InnerText.Trim();

                //trs = doc.DocumentNode.SelectNodes(".//table");
                //HtmlNode tr = trs[1];
                //HtmlNodeCollection tds = tr.SelectNodes(".//tr");
                //string companyName = tds[2].InnerText.Trim();

                HtmlDocument htc = new HtmlDocument();
                htc.LoadHtml(pageSource);
                HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                HtmlNode table = tables[1];
                HtmlNodeCollection trs = table.SelectNodes(".//tr");
                string companyName = string.Empty;
                companyName = trs[1].SelectNodes(".//td")[2].InnerText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();


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
                    //string instrumentType = tds[3].InnerText.Trim();
                    string instrumentType = trs[1].SelectNodes(".//td")[3].InnerText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();
                    if (instrumentType.Equals("예탁증서"))
                    {
                        peoType = "KDR";
                    }
                    else
                    {
                        peoType = "ORD";
                        errors.Add("At GetPeoTypeByISIN(). Can not determine the PEO type by isin." + isin + " Set it default as ORD.");
                    }
                }
            }
            catch
            {
                string msg = "At GetPeoTypeByISIN(string isin). Error found in searching peo type infos for " + isin;
                errors.Add(msg);
            }
            return peoType;
        }

        private string GetPageSource(string uri, string postData)
        {
            HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
            request.ProtocolVersion = HttpVersion.Version11;
            request.Timeout = 100000;
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1500.95 Safari/537.36";
            request.Method = "POST";
            request.KeepAlive = true;
            request.Headers["Accept-Language"] = "zh-CN,zh;q=0.8,en;q=0.6";
            request.Accept = "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*";
            //request.ContentType = "multipart/form-data; boundary=---------------------------7de29a81c0e16";
            request.ContentType = "multipart/form-data; boundary=----WebKitFormBoundaryh2aBgmqWp5ezUvqh";
            byte[] buf = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = buf.Length;
            request.GetRequestStream().Write(buf, 0, buf.Length);
            HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(httpResponse.GetResponseStream());
            return sr.ReadToEnd();
        }

        /// <summary>
        /// Grab PEO add kq or kn data from page(data in table).
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="dutyCode"></param>
        private void GrabKQDataAction(HtmlDocument doc, string dutyCode)
        {
            try
            {
                HtmlNodeCollection tables = doc.DocumentNode.SelectNodes("//table");
                int cou = tables.Count;
                HtmlNode table = tables[(cou - 1)];

                string strIdnDisplayName = table.SelectSingleNode(".//tr[4]/td[2]").InnerText.Trim();
                string strTicker = table.SelectSingleNode(".//tr[6]/td[3]").InnerText;
                string strKorreaname = table.SelectSingleNode(".//tr[1]/td[4]").InnerText;
                string strLegalname = table.SelectSingleNode(".//tr[2]/td[2]").InnerText;
                // string strPrice = table.SelectSingleNode(".//tr[7]/td[2]").InnerText;
                string strIsin = table.SelectSingleNode(".//tr[6]/td[2]").InnerText;
                string strEffectiveDate = string.Empty;
                if (dutyCode.Equals("KQ"))
                {
                    strEffectiveDate = table.SelectSingleNode(".//tr[15]/td[2]").InnerText.Trim();
                }
                else
                {
                    strEffectiveDate = table.SelectSingleNode(".//tr[14]/td[2]").InnerText.Trim();
                }

                FormatDate(ref strEffectiveDate);

                // strPrice = Regex.Replace(strPrice, "[^0-9]", "");

                KoreaEquityInfo peoTemp = new KoreaEquityInfo();

                peoTemp.Ticker = strTicker.Trim().ToString().Substring(1);
                peoTemp.RIC = peoTemp.Ticker + "." + dutyCode;
                peoTemp.KoreaName = strKorreaname.Trim().Contains("(주)") ? strKorreaname.Trim().Replace("(주)", "") : strKorreaname.Trim();
                peoTemp.LegalName = strLegalname.Trim().ToString();
                peoTemp.ISIN = strIsin.Trim().ToString();
                peoTemp.EffectiveDate = strEffectiveDate;
                //  peoTemp.Lotsize = strPrice;                
                peoTemp.Market = dutyCode;
                peoTemp.IDNDisplayName = strIdnDisplayName.ToUpper();
                peoTemp.Type = GetPeoTypeByISIN(peoTemp.ISIN);
                paList.Add(peoTemp);
            }
            catch (Exception ex)
            {
                string msg = "At GrabKQDataAction()." + ex.Message;
                errors.Add(msg);
            }
        }

        #endregion

        #region Grab ETF data

        /// <summary>
        /// Grab ETF data.
        /// </summary>
        /// <param name="doc">page document</param>
        private void GrabETFDataAction(HtmlDocument doc, bool isGlobal)
        {
            if (doc != null)
            {
                HtmlNode pre = doc.DocumentNode.SelectSingleNode(".//pre");
                string document = pre.InnerText;
                if (string.IsNullOrEmpty(document))
                    return;

                if (document.Contains("상장종목명"))
                {
                    string koreanName = string.Empty;
                    string legalName = string.Empty;
                    string effectiveDate = string.Empty;
                    string isinRic = string.Empty;

                    int effectiveDatePos = document.IndexOf("상장일") + ("상장일".Length);
                    int isinRicPos = document.IndexOf("표준코드") + ("표준코드".Length);
                    isinRic = FormatDataWithPos(isinRicPos, document);
                    string isin = isinRic.Split('(')[0].Trim().ToString();
                    string ticker = isinRic.Split(':')[1].Trim(new Char[] { ' ', ':', ')' }).ToString().Substring(1);

                    int koreanNamePos = document.IndexOf("한글종목명") + ("한글종목명".Length);
                    int legalNamePos = document.IndexOf("영문종목명") + ("영문종목명".Length);
                    koreanName = FormatDataWithPos(koreanNamePos, document);
                    legalName = FormatDataWithPos(legalNamePos, document);


                    effectiveDate = FormatDataWithPos(effectiveDatePos, document);
                    FormatDate(ref effectiveDate);
                    string[] names = GetETFDetailByIsin(isin);

                    KoreaEquityInfo etf = new KoreaEquityInfo();
                    etf.ISIN = isin;
                    etf.Ticker = ticker;
                    etf.RIC = ticker + ".KS";
                    etf.EffectiveDate = effectiveDate;
                    etf.KoreaName = koreanName;
                    etf.LegalName = legalName;
                    etf.Market = "KS";
                    etf.Type = "ETF";
                    etf.IsGlobalETF = isGlobal;

                    etf.IDNDisplayName = KoreaEquityCommon.ClearCoLtdForName(legalName);

                    if (names != null)
                    {
                        if (!string.IsNullOrEmpty(names[0]) && koreanName.Length > 14)
                        {
                            etf.KoreaName = names[0];
                        }
                        if (!string.IsNullOrEmpty(names[1]) && etf.IDNDisplayName.Length > 16)
                        {
                            etf.IDNDisplayName = names[1].ToUpper();
                        }
                    }
                    eaList.Add(etf);
                }
            }
        }

        /// <summary>
        /// Get ETF detail infomation from ISIN website.
        /// </summary>
        /// <param name="isin">isin</param>
        /// <returns>korea name and legal name</returns>
        private string[] GetETFDetailByIsin(string isin)
        {
            string[] names = new string[2];
            string uri = string.Format("http://isin.krx.co.kr/jsp/BA_VW010.jsp?isu_cd={0}&modi=f&req_no=", isin);
            string pageSource = null;
            int retry = 3;
            while (pageSource == null && retry-- > 0)
            {
                try
                {
                    pageSource = WebClientUtil.GetPageSource(null, uri, 6000, "", Encoding.GetEncoding("EUC-KR"));
                }
                catch
                {
                    continue;
                }
            }
            if (pageSource == null)
            {
                string msg = "Can not get the ETF equity infos in ISIN webpage. For ISIN:" + isin + ". please check if the webpage can be accessed!";
                return null;
            }

            try
            {
                HtmlDocument isinRoot = new HtmlDocument();
                isinRoot.LoadHtml(pageSource);
                HtmlNode isinTable = isinRoot.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[1]/tr[3]/td[1]/table[1]");
                HtmlNodeCollection isinTrs = isinTable.SelectNodes("./tr");

                string koreaName = isinTrs[9].SelectNodes("./td")[1].InnerText.Trim();
                if (koreaName.Contains("상장지수투자신"))
                {
                    koreaName = koreaName.Substring(0, koreaName.IndexOf("상장지수투자신"));
                }
                // 미래에셋TIGER합성-MSCIUS리츠부동산상장지수투자신탁(파생형)(H)
                // 미래에셋TIGER합성-MSCIUS리츠부동산상장지수투자신탁(파생형)(H)
                //                                 상장지수투자신탁
                // 미래에셋TIGER합성-MSCIUS리츠부동산상장지수투자신

                if (koreaName.Length > 14)
                {
                    koreaName = isinTrs[9].SelectNodes("./td")[3].InnerText.Trim();
                }
                names[0] = koreaName;
                names[1] = isinTrs[10].SelectNodes("./td")[3].InnerText.Trim();
                return names;
            }

            catch (Exception ex)
            {
                string msg = string.Format("At GetETFDetailByIsin(). Error found in getting korea name and legal name for ISIN:{0}. Error message: {1}", isin, ex.Message);
                errors.Add(msg);
                return null;
            }
        }

        /// <summary>
        /// For ETF, User need to judge if it is a global ETF.
        /// </summary>
        /// <param name="time">announcement time</param>
        /// <param name="title">announcement titile</param>
        /// <returns>true or false</returns>
        private bool JudgeIfGloblaETF(string time, string title)
        {
            string isGlobal = null;
            while (isGlobal == null)
            {
                isGlobal = JudgeGlobalETF.Prompt(time, title);
            }
            if (isGlobal.Equals("Y"))
            {
                return true;
            }
            return false;
        }

        #endregion

        #region Grab REIT data

        /// <summary>
        /// Grab REIT data.
        /// </summary>
        /// <param name="doc">page document</param>
        private void GrabREITsDataAction(HtmlDocument doc)
        {
            if (doc != null)
            {
                HtmlNode pre = doc.DocumentNode.SelectSingleNode(".//pre");
                string strPre = pre.InnerText;
                if (string.IsNullOrEmpty(strPre))
                    return;

                int strKoreaNamePos = strPre.IndexOf("- 한글종목명") + ("- 한글종목명".Length);
                if (strKoreaNamePos < ("- 한글종목명".Length))
                    strKoreaNamePos = strPre.IndexOf("- 한글명") + ("- 한글명".Length);

                int strLegalNamePos = strPre.IndexOf("- 영문종목명") + ("- 영문종목명".Length);
                if (strLegalNamePos < ("- 영문종목명".Length))
                    strLegalNamePos = strPre.IndexOf("- 영문명") + ("- 영문명".Length);

                int strEffectivePos = strPre.IndexOf("상장일") + ("상장일".Length);
                if (strEffectivePos < ("상장일".Length))
                    strEffectivePos = strPre.IndexOf("상장(예정)일") + ("상장(예정)일".Length);

                int strIsinTickerPos = strPre.IndexOf("표준코드") + ("표준코드".Length);

                int strStrikePos = strPre.IndexOf("1주의 발행가액") + ("1주의 발행가액".Length);
                if (strStrikePos < ("1주의 발행가액".Length))
                    strStrikePos = strPre.IndexOf("1주의 발행(공모)가액") + ("1주의 발행(공모)가액".Length);
                if (strStrikePos < ("1주의 발행(공모)가액".Length))
                    strStrikePos = strPre.IndexOf("1주의 액면가액(공모가액)") + ("1주의 액면가액(공모가액)".Length);


                string koreanNameTemp = FormatDataWithPos(strKoreaNamePos, strPre);
                string legalNameTemp = FormatDataWithPos(strLegalNamePos, strPre);
                string effectiveTemp = FormatDataWithPos(strEffectivePos, strPre);
                string isinRicTemp = FormatDataWithPos(strIsinTickerPos, strPre);
                string strikePriceTemp = FormatDataWithPos(strStrikePos, strPre);

                string koreanName = koreanNameTemp.Trim(new Char[] { ' ', ':', ')' }).ToString();
                koreanName = koreanName.Contains("(약명") ? koreanName.Substring(0, (koreanName.IndexOf("(약명"))).Trim().ToString() : koreanName;
                string legalName = legalNameTemp.Trim(new Char[] { ' ', ':', ')' }).ToString();
                legalName = legalNameTemp.Contains("(약명") ? legalName.Substring(0, (legalName.IndexOf("(약명"))) : legalName;
                string effectiveDate = effectiveTemp.Trim(new Char[] { ' ', ':', ')', '목', '(' }).ToString();
                FormatDate(ref effectiveDate);
                string isin = isinRicTemp.Split('(')[0].Trim(new Char[] { ' ', ':', '(' }).ToString();
                string ticker = Regex.Split(isinRicTemp, "단축코드", RegexOptions.IgnoreCase)[1].Trim(new Char[] { ' ', ':', '(', ')' }).ToString();
                string strikePrice = strikePriceTemp.Trim(new Char[] { ' ', ':' }).ToString();
                strikePrice = strikePrice.Contains('원') ? strikePrice.Substring(0, (strikePrice.IndexOf('원'))) : strikePrice;
                strikePrice = strikePrice.Contains(',') ? strikePrice.Replace(",", "") : strikePrice;

                KoreaEquityInfo ra = new KoreaEquityInfo();
                ra.ISIN = isin;
                ra.Ticker = ticker.Substring(1).Trim().ToString();
                ra.RIC = ra.Ticker + ".KS";
                ra.EffectiveDate = effectiveDate;
                ra.KoreaName = koreanName;
                ra.LegalName = legalName;
                ra.Lotsize = strikePrice;
                ra.IDNDisplayName = legalName.Substring(0, 16).ToUpper();
                ra.Market = "KS";
                ra.Type = "REIT";
                raList.Add(ra);
            }
        }

        #endregion

        #region Grab BC data and format

        /// <summary>
        /// Grab BC data. 
        /// Get ISIN in ks page and get the detail infos in ISIN website.
        /// </summary>
        /// <param name="doc">ks page</param>
        private void GrabBCDataAction(HtmlDocument doc)
        {
            HtmlNode pre = doc.DocumentNode.SelectSingleNode(".//pre");
            string document = pre.InnerText;
            if (string.IsNullOrEmpty(document))
            {
                int isinPos = document.IndexOf("표준코드") + 4;
                string isinLine = FormatDataWithPos(isinPos, document);
                string isin = string.Empty;
                string ticker = string.Empty;
                Regex regex = new Regex("([a-zA-Z0-9]+)");
                Match match = regex.Match(isinLine);
                isin = match.Value;

                if (string.IsNullOrEmpty(isin))
                {
                    string msg = "The ISIN to find detail infomation is null for BC in GrabBCDataAction(HtmlDocument doc)!";
                    errors.Add(msg);
                    return;
                }

                int effectiveDatePos = document.IndexOf("상장일") + "상장일".Length;
                string effectiveDate = FormatDataWithPos(effectiveDatePos, document);
                FormatDate(ref effectiveDate);

                GetBCDetailByIsin(isin, effectiveDate);
            }
        }

        /// <summary>
        /// Get BC detail infomation from isin website.
        /// </summary>
        /// <param name="isin">isin</param>
        private void GetBCDetailByIsin(string isin, string effectiveDate)
        {
            string uri = string.Format("http://isin.krx.co.kr/jsp/BA_VW012.jsp?isu_cd={0}&modi=f&req_no=", isin);

            string pageSource = null;
            int retry = 3;
            while (pageSource == null && retry-- > 0)
            {
                try
                {
                    pageSource = WebClientUtil.GetPageSource(null, uri, 6000, "", Encoding.GetEncoding("EUC-KR"));
                }
                catch
                {
                    continue;
                }
            }
            if (pageSource == null)
            {
                string msg = "Can not get the New Underlying infos in ISIN webpage. For ISIN:" + isin + ". please check if the webpage can be accessed!";
                return;
            }
            HtmlDocument isinRoot = new HtmlDocument();
            isinRoot.LoadHtml(pageSource);
            HtmlNode isinTable = isinRoot.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[1]/tr[3]/td[1]/table[1]");
            HtmlNodeCollection isinTrs = isinTable.SelectNodes("./tr");

            string ticker = string.Empty;
            string koreaName = string.Empty;
            string koreaFullName = string.Empty;
            string legalName = string.Empty;
            string idnDisplayName = string.Empty;
            foreach (var tr in isinTrs)
            {
                if (tr.InnerText.Contains("단축코드") && string.IsNullOrEmpty(ticker))
                {
                    ticker = tr.SelectSingleNode("./td[4]").InnerText.Trim().Substring(1); ;
                    continue;
                }
                if (tr.InnerText.Contains("종목약명") && string.IsNullOrEmpty(koreaName))
                {
                    koreaFullName = tr.SelectSingleNode("./td[2]").InnerText.Replace("(주)", "").Trim();
                    koreaName = tr.SelectSingleNode("./td[4]").InnerText.Trim();
                    continue;
                }
                if (tr.InnerText.Contains("종목영문명") && string.IsNullOrEmpty(koreaName))
                {
                    legalName = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                    idnDisplayName = tr.SelectSingleNode("./td[4]").InnerText.Trim();
                    continue;
                }
            }

            string ric = ticker + ".KS";

            if (KoreaEquityManager.ExsitsFMTwo(ric, isin))
            {
                string msg = string.Format("RIC:{0}, ISIN:{1} already exsited in database.", ric, isin);
                Logger.Log(msg);
                return;
            }

            KoreaEquityInfo equity = new KoreaEquityInfo();
            equity.Status = "Active";
            equity.RIC = ticker + ".KS";
            equity.ISIN = isin;
            equity.Ticker = ticker;
            equity.Lotsize = "10";
            equity.UpdateDate = DateTime.Today.ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
            equity.Type = "CEF";
            equity.RecordType = "112";
            equity.FM = "2";
            equity.Category = "CEF";
            equity.BcastRef = equity.RIC;
            equity.KoreaName = koreaFullName;
            if (equity.KoreaName.Length > 14)
            {
                equity.KoreaName = equity.KoreaName.Substring(0, 14);
            }
            equity.LegalName = legalName;
            equity.EffectiveDate = effectiveDate;
            equity.IDNDisplayName = idnDisplayName;

            KoreaEquityCommon.FormatEQIdnDisplayName(equity);

            try
            {
                KoreaEquityManager.UpdateEquity(equity);
            }
            catch (Exception ex)
            {
                string msg = "At GetBCDetailByIsin(). Error found in update bc equity to database." + ex.Message;
                errors.Add(msg);
            }
            AddEquityToGroup(equity, "BC");

            bcList.Add(equity);
        }

        #endregion

        #endregion

        #region Grab re-listing data

        /// <summary>
        /// Grab PEO new listing data. Key word:재상장
        /// </summary>
        /// <param name="startDate">start date</param>
        /// <param name="endDate">end date</param>
        private void GrabPEONewListing(string startDate, string endDate)
        {
            string dataStartDate = null;
            DateTime startDay = DateTime.Parse(startDate);
            DateTime twoMonthEarly = DateTime.Parse(endDate).AddMonths(-2);
            if (startDay.CompareTo(twoMonthEarly) < 0)
            {
                dataStartDate = startDate;
            }
            else
            {
                dataStartDate = twoMonthEarly.ToString("yyyy-MM-dd");
            }
            string uri = "http://kind.krx.co.kr/disclosure/searchtotalinfo.do";
            string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%9E%AC%EC%83%81%EC%9E%A5&repIsuSrtCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=500&scn=mktact&srchFd=2&kwd=%EC%9E%AC%EC%83%81%EC%9E%A5&fromData={0}&toData={1}", dataStartDate, endDate);
            HtmlDocument htc = new HtmlDocument();
            string pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, postData);
            if (!string.IsNullOrEmpty(pageSource))
                htc.LoadHtml(pageSource);

            if (htc != null)
            {
                HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes("//dl/dt");
                HtmlNodeCollection ddCollections = htc.DocumentNode.SelectNodes("//dl/dd");
                if (nodeCollections != null)
                {
                    int count = nodeCollections.Count;

                    for (var i = 0; i < count; i++)
                    {
                        HtmlNode ddNode = ddCollections[i].SelectSingleNode(".//span");
                        HtmlNode node = nodeCollections[i].SelectSingleNode(".//span/a");
                        HtmlNode header = nodeCollections[i].SelectSingleNode(".//strong/a");
                        string title = string.Empty;
                        if (node != null)
                            title = node.InnerText.Trim().ToString();
                        HtmlNode nodeDate = nodeCollections[i].SelectSingleNode("./em");
                        if (nodeDate != null)
                        {
                            DateTime anouncementDate = new DateTime();
                            anouncementDate = DateTime.Parse(nodeDate.InnerText.Trim(), new CultureInfo("en-US"));
                            if (anouncementDate < startDay)
                            {
                                return;
                            }
                        }
                        if (!string.IsNullOrEmpty(title))
                        {
                            HtmlDocument doc = null;
                            string KSorKQorKN = string.Empty;

                            if (title.Contains("재상장"))
                            {
                                if (!title.Substring(0, "재상장".Length).Equals("재상장")) continue;

                                HtmlNode companyNode = nodeCollections[i].SelectSingleNode(".//strong/a");
                                KoreaCompany company = GetCompanyInfo(companyNode);
                                KSorKQorKN = GetDutyCode(ddNode);
                                doc = GetTargetPageDocument(node);
                                if (KSorKQorKN.Equals("KS"))
                                {                                       // --------------> PEO ADD for KS 
                                    GrabKSPeoData(doc, company);
                                }
                                else                                   // --------------> PEO ADD for KQ and KN
                                {
                                    GrabKQDataAction(doc, KSorKQorKN);
                                    // GrabPeoKQKNData(doc, KSorKQorKN);
                                }
                            }
                            else
                                continue;
                        }
                    }

                }
            }
        }

        #endregion

        #endregion

        #region Formatting

        #region Format PEO data

        /// <summary>
        /// Format PEO add data.
        /// Group the equity with same effective date and same (exsitsFM1). Stored in hashtable
        /// if exsits FM1 . key: PEO+effective date. eg. PEO2013-09-09
        /// if not , key:PEOADD+effective date. eg. PEOADD2013-09-10
        /// </summary>
        private void FormatPEOADDTemplate()
        {
            for (int i = 0; i < paList.Count; i++)
            {
                KoreaEquityInfo item = paList[i];
                if (KoreaEquityManager.ExsitsFMTwo(item.RIC, item.ISIN))
                {
                    string msg = string.Format("RIC:{0}, ISIN:{1} already exsited in database.", item.RIC, item.ISIN);
                    Logger.Log(msg);
                    paList.RemoveAt(i);
                    i--;
                    continue;
                }

                item.Status = "Active";
                item.UpdateDate = DateTime.Today.ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                item.FM = "2";
                item.BcastRef = item.RIC;
                try
                {
                    if (item.Type.Equals("ORD"))
                    {
                        item.Category = "ORD";
                        item.RecordType = "113";
                        if (item.Market.Equals("KS"))
                        {
                            item.Lotsize = (Convert.ToInt32(item.Lotsize) >= 50000) ? "1" : "10";
                        }
                    }
                    else if (item.Type.Equals("PRF"))
                    {
                        item.Category = "PRF";
                        item.RecordType = "97";
                        if (item.Market.Equals("KS"))
                        {
                            item.Lotsize = (Convert.ToInt32(item.Lotsize) >= 50000) ? "1" : "10";
                        }
                        if (item.KoreaName.Contains("우선주") && item.KoreaName.Length > 14)
                        {
                            string tempName = item.KoreaName.Replace("우선주", "");
                            if (tempName.Length > 14)
                            {
                                item.KoreaName = tempName.Substring(0, 13) + "우";
                            }
                            else
                            {
                                item.KoreaName = tempName + "우";
                            }
                        }
                    }
                    else if (item.Type.Equals("KDR"))
                    {
                        item.Category = "DRC";
                        item.RecordType = "113";
                        item.Lotsize = "10";
                    }
                }
                catch
                {
                    item.Lotsize = "N";
                    string msg = "At FormatPEOADDTemplate().Error found in format Lotsize for RIC:" + item.RIC;
                    errors.Add(msg);
                }

                if (item.Market.Equals("KS"))
                {
                    item.Exchange = "KSC";
                }
                else if (item.Market.Equals("KQ"))
                {
                    item.Exchange = "KOE";
                    item.Lotsize = "1";
                }
                else
                {
                    item.Exchange = "KNX";
                    item.Lotsize = "100";
                }

                if (item.KoreaName.Length > 14)
                {
                    item.KoreaName = item.KoreaName.Substring(0, 14);
                }

                string groupName = "PEOADD";
                item.ExistsFM1 = KoreaEquityManager.ExistsFMOne(item.Ticker);
                if (item.ExistsFM1)
                {
                    groupName = "PEOFM";
                    CheckPeoForChanges(item);
                }
                else
                {
                    KoreaEquityCommon.FormatEQIdnDisplayName(item);
                }

                try
                {
                    KoreaEquityManager.UpdateEquity(item);
                }
                catch (Exception ex)
                {
                    string msg = "At FormatPEOADDTemplate(). Error found in update peo equity to database." + ex.Message;
                    errors.Add(msg);
                }

                if (string.IsNullOrEmpty(item.EffectiveDate))
                {
                    string msg = string.Format("Error found when formatting peo data. Effective Date for RIC:{0} is null.", item.RIC);
                    errors.Add(msg);
                    continue;
                }

                AddEquityToGroup(item, groupName);
            }
        }

        /// <summary>
        /// Check and mark the changes
        /// </summary>
        /// <param name="item">equity</param>
        private void CheckPeoForChanges(KoreaEquityInfo item)
        {
            KoreaEquityInfo itemFm1 = KoreaEquityManager.SelectEquityFMOne(item.RIC);
            if (itemFm1 == null)
            {
                return;
            }

            if (!item.LegalName.Equals(itemFm1.LegalName))
            {
                item.ChangeItems.Add(KoreaAddFMColumn.LegalName);
                if (string.IsNullOrEmpty(item.IDNDisplayName) || item.IDNDisplayName.Length > 16)
                {
                    KoreaEquityCommon.FormatEQIdnDisplayName(item);
                }
                if (!item.IDNDisplayName.Equals(itemFm1.IDNDisplayName))
                {
                    item.ChangeItems.Add(KoreaAddFMColumn.IDNDisplayName);
                    peoChanged = true;
                }
            }
            else
            {
                item.IDNDisplayName = itemFm1.IDNDisplayName;
            }

            if (item.KoreaName != itemFm1.KoreaName)
            {
                item.ChangeItems.Add(KoreaAddFMColumn.KoreaName);
            }
            if (item.ISIN != itemFm1.ISIN)
            {
                item.ChangeItems.Add(KoreaAddFMColumn.ISIN);
            }
        }

        /// <summary>
        /// Add the equities with same type and effective date to group. Stored in hash table.
        /// </summary>
        /// <param name="item">equity item</param>
        private void AddEquityToGroup(KoreaEquityInfo item, string groupName)
        {
            string key = groupName + "," + item.EffectiveDate;
            List<KoreaEquityInfo> equities = new List<KoreaEquityInfo>();
            if (equityDate.Contains(key))
            {
                equities = (List<KoreaEquityInfo>)equityDate[key];
                equities.Add(item);
                equityDate[key] = equities;
                return;
            }

            equities.Add(item);
            equityDate.Add(key, equities);
        }

        #endregion

        #region Format ETF data

        /// <summary>
        /// Format ETF add data.
        /// </summary>
        private void FormatETFADDTemplate()
        {
            for (int i = 0; i < eaList.Count; i++)
            {
                KoreaEquityInfo etf = eaList[i];
                if (KoreaEquityManager.ExsitsFMTwo(etf.RIC, etf.ISIN))
                {
                    string msg = string.Format("RIC:{0}, ISIN:{1} already exsited in database.", etf.RIC, etf.ISIN);
                    Logger.Log(msg);
                    paList.RemoveAt(i);
                    i--;
                    continue;
                }
                etf.Status = "Active";
                etf.FM = "2";
                etf.BcastRef = etf.RIC;
                etf.Type = "ETF";
                etf.RecordType = "112";
                etf.Lotsize = "1";
                etf.Category = "ECL";
                etf.UpdateDate = DateTime.Today.ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();

                if (etf.KoreaName.Length > 14)
                {
                    etf.KoreaName = etf.KoreaName.Substring(0, 14);
                }
                KoreaEquityCommon.FormatEQIdnDisplayName(etf);
                try
                {
                    KoreaEquityManager.UpdateEquity(etf);
                }
                catch (Exception ex)
                {
                    string msg = "At FormatETFADDTemplate(). Error found in update etf equity to database." + ex.Message;
                    errors.Add(msg);
                }
                AddEquityToGroup(etf, "ETF");
            }
        }

        #endregion

        #region Format REIT data

        /// <summary>
        /// Format REIT add data.
        /// </summary>
        private void FormatREITADDTemplate()
        {
            for (int i = 0; i < raList.Count; i++)
            {
                KoreaEquityInfo ra = raList[i];
                if (KoreaEquityManager.ExsitsFMTwo(ra.RIC, ra.ISIN))
                {
                    string msg = string.Format("RIC:{0}, ISIN:{1} already exsited in database.", ra.RIC, ra.ISIN);
                    Logger.Log(msg);
                    paList.RemoveAt(i);
                    i--;
                    continue;
                }
                ra.Status = "Active";
                ra.UpdateDate = DateTime.Today.ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                ra.RIC = ra.Ticker.Trim().ToString() + ".KS";
                ra.Type = "REIT";
                ra.RecordType = "112";
                ra.FM = "2";
                ra.Category = "REI";
                ra.BcastRef = ra.RIC;
                try
                {
                    ra.Lotsize = Convert.ToInt32(ra.Lotsize) >= 50000 ? "1" : "10";
                }
                catch
                {
                    ra.Lotsize = "N";
                    string msg = "At FormatREITADDTemplate(). Error found in format Lotsize for RIC:" + ra.RIC;
                    errors.Add(msg);
                }

                if (ra.KoreaName.Length > 14)
                {
                    ra.KoreaName = ra.KoreaName.Substring(0, 14);
                }
                KoreaEquityCommon.FormatEQIdnDisplayName(ra);
                try
                {
                    KoreaEquityManager.UpdateEquity(ra);
                }
                catch (Exception ex)
                {
                    string msg = "At FormatREITADDTemplate(). Error found in update reit equity to database." + ex.Message;
                    errors.Add(msg);
                }
                AddEquityToGroup(ra, "REIT");
            }
        }

        #endregion

        #endregion

        #region Distributing

        #region Generate FM File

        /// <summary>
        /// Generate FM Files.
        /// </summary>
        private void GenerateFMFiles()
        {
            List<KoreaEquityInfo> addList = new List<KoreaEquityInfo>();
            addList.AddRange(paList);
            addList.AddRange(eaList);
            addList.AddRange(raList);
            addList.AddRange(bcList);
            if (addList.Count > 0)
            {
                Logger.Log("Generate FM files.");
            }
            foreach (KoreaEquityInfo item in addList)
            {
                ExcelApp excelApp = new ExcelApp(false, false);
                if (excelApp.ExcelAppInstance == null)
                {
                    string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                    errors.Add(msg);
                    return;
                }
                try
                {
                    string filePath = string.Empty;
                    string fileName = string.Empty;
                    string mailType = string.Empty;
                    string today = DateTime.Today.ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                    if (item.Type.Equals("ORD") || item.Type.Equals("PRF") || item.Type.Equals("KDR"))
                    {
                        filePath = configObj.PEOFM;
                        fileName = "KR FM (PEO ADD)_" + item.RIC + " (wef " + today + ").xls";
                        if (!item.ExistsFM1)
                        {
                            fileName = "KR FM (PEO & PEO ADD)_" + item.RIC + " (wef " + today + ").xls";
                        }
                        mailType = "PEO";
                    }
                    else if (item.Type.Equals("ETF"))
                    {
                        filePath = configObj.ETFFM;
                        fileName = "KR FM (ETF & ETF ADD)_" + item.RIC + " (wef " + today + ").xls";
                        mailType = "ETF";
                    }
                    else if (item.Type.Equals("REIT"))
                    {
                        filePath = configObj.REITFM;
                        fileName = "KR FM (REIT ADD)_" + item.RIC + " (wef " + today + ").xls";
                        mailType = "REIT";
                    }
                    else if (item.Type.Equals("CEF"))
                    {
                        filePath = configObj.BCFM;
                        fileName = "KR FM (BC ADD)_" + item.RIC + " (wef " + today + ").xls";
                        mailType = "BC";
                    }

                    filePath = Path.Combine(filePath, fileName);
                    Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                    if (wSheet == null)
                    {
                        string msg = "Worksheet could not be started. Check that your office installation and project reference are correct!";
                        errors.Add(msg);
                        return;
                    }
                    List<string> fmTitle = new List<string>() {"Updated Date", "Effective Date", "RIC", "Type", "Record Type", "FM", "IDN Display Name", 
                                                               "ISIN", "Ticker", "BCAST REF", "Legal Name", "Korea Name", "Lotsize" };

                    ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 13;
                    ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 13;
                    ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 20;
                    ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 16;
                    ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 20;
                    ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 20;

                    ((Range)wSheet.Cells[3, 1]).Font.Underline = System.Drawing.FontStyle.Underline;
                    ((Range)wSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                    wSheet.Cells[3, 1] = "EQUITY ADD";
                    for (int j = 1; j <= fmTitle.Count; j++)
                    {
                        wSheet.Cells[4, j] = fmTitle[j - 1];
                    }
                    ((Range)wSheet.Rows[4, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    wSheet.Cells[5, 1] = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    wSheet.Cells[5, 2] = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    wSheet.Cells[5, 3] = item.RIC;
                    wSheet.Cells[5, 4] = item.Type;
                    wSheet.Cells[5, 5] = item.RecordType;
                    wSheet.Cells[5, 6] = item.FM;
                    wSheet.Cells[5, 7] = item.IDNDisplayName;
                    wSheet.Cells[5, 8] = item.ISIN;
                    ((Range)wSheet.Cells[5, 9]).NumberFormat = "@";
                    wSheet.Cells[5, 9] = item.Ticker;
                    wSheet.Cells[5, 10] = item.BcastRef;
                    wSheet.Cells[5, 11] = item.LegalName;
                    wSheet.Cells[5, 12] = item.KoreaName;
                    wSheet.Cells[5, 13] = item.Lotsize;
                    wSheet.Cells[8, 1] = "- End -";
                    ((Range)wSheet.Cells[8, 1]).Font.Bold = System.Drawing.FontStyle.Bold;

                    excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                    wBook.Save();

                    MailToSend mail = new MailToSend();
                    mail.ToReceiverList.AddRange(configObj.MailTo);
                    mail.CCReceiverList.AddRange(configObj.MailCC);
                    mail.MailSubject = Path.GetFileNameWithoutExtension(fileName);
                    mail.AttachFileList.Add(filePath);
                    mail.MailBody = mailType + " ADD:\t\t" + item.RIC + "\r\n\r\n"
                                    + "Effective Date:\t\t" + item.EffectiveDate + "\r\n\r\n\r\n\r\n";
                    string signature = string.Join("\r\n", configObj.MailSignature.ToArray());
                    mail.MailBody += signature;

                    // KoreaEquityManager.UpdateEquity(item);
                    AddResult(fileName, filePath, item.Type + " FM File");
                    Logger.Log("Generate" + item.Type + " ADD FM file successfully. Filepath is " + filePath);
                }
                catch (Exception ex)
                {
                    string msg = "At GenerateFMFiles(). Error found in Generate FM file." + ex.Message;
                    errors.Add(msg);
                }
                finally
                {
                    excelApp.Dispose();
                }
            }
        }

        private void GenerateFMFilesByGroup()
        {
            Logger.Log("Generate FM Files.");
            ArrayList fmFiles = new ArrayList(equityDate.Keys);
            fmFiles.Sort();
            foreach (string fm in fmFiles)
            {
                List<KoreaEquityInfo> equity = (List<KoreaEquityInfo>)equityDate[fm];
                string ricPart = CombineAllRics(equity);
                string fileName = string.Empty;
                string filePath = string.Empty;
                string mailType = string.Empty;
                string mailAdd = string.Empty;
                string effectiveDate = fm.Split(',')[1];
                if (fm.Contains("PEOFM"))
                {
                    fileName = "KR FM (PEO ADD)_" + ricPart + " (wef " + effectiveDate.ToUpper() + ").xls";
                    filePath = configObj.PEOFM;
                    mailType = "PEO";

                }
                else if (fm.Contains("PEOADD"))
                {
                    fileName = "KR FM (PEO & PEO ADD)_" + ricPart + " (wef " + effectiveDate.ToUpper() + ").xls";
                    filePath = configObj.PEOFM;
                    mailType = "PEO & PEO";
                }
                else if (fm.Contains("ETF"))
                {
                    fileName = "KR FM (ETF & ETF ADD)_" + ricPart + " (wef " + effectiveDate.ToUpper() + ").xls";
                    filePath = configObj.ETFFM;
                    mailType = "ETF & ETF";
                }
                else if (fm.Contains("REIT"))
                {
                    fileName = "KR FM (REIT ADD)_" + ricPart + " (wef " + effectiveDate.ToUpper() + ").xls";
                    filePath = configObj.REITFM;
                    mailType = "REIT";
                }
                else if (fm.Contains("BC"))
                {
                    fileName = "KR FM (BC ADD)_" + ricPart + " (wef " + effectiveDate.ToUpper() + ").xls";
                    filePath = configObj.BCFM;
                    mailType = "BC";
                }
                string path = Path.Combine(filePath, fileName);
                if (path.Length >= 260)
                {
                    fileName = fileName.Replace(ricPart, "RIC");
                }
                GenerateFMFileForEach(filePath, fileName, equity, mailType, ricPart, effectiveDate);
            }
        }


        //  private void GenerateFutruedating





        private void GenerateFMFileForEach(string filePath, string fileName, List<KoreaEquityInfo> equity, string mailType, string rics, string effectiveDate)
        {
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                errors.Add(msg);
                return;
            }
            try
            {
                filePath = Path.Combine(filePath, fileName);
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Worksheet could not be started. Check that your office installation and project reference are correct!";
                    errors.Add(msg);
                    return;
                }
                List<string> mailAdd = new List<string>();

                List<string> fmTitle = new List<string>() {"Updated Date", "Effective Date", "RIC", "Type", "Record Type", "FM", "IDN Display Name", 
                                                               "ISIN", "Ticker", "BCAST REF", "Legal Name", "Korea Name", "Lotsize" };

                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 13;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 13;
                ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 16;
                ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 20;

                ((Range)wSheet.Cells[3, 1]).Font.Underline = System.Drawing.FontStyle.Underline;
                ((Range)wSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                wSheet.Cells[3, 1] = "EQUITY ADD";
                for (int j = 1; j <= fmTitle.Count; j++)
                {
                    wSheet.Cells[4, j] = fmTitle[j - 1];
                }
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;

                for (int k = 0; k < equity.Count; k++)
                {
                    KoreaEquityInfo item = equity[k];

                    wSheet.Cells[5 + k, 1] = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    wSheet.Cells[5 + k, 2] = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    wSheet.Cells[5 + k, 3] = item.RIC;
                    wSheet.Cells[5 + k, 4] = item.Type;
                    wSheet.Cells[5 + k, 5] = item.RecordType;
                    wSheet.Cells[5 + k, 6] = item.FM;
                    wSheet.Cells[5 + k, 7] = item.IDNDisplayName;
                    wSheet.Cells[5 + k, 8] = item.ISIN;
                    ((Range)wSheet.Cells[5 + k, 9]).NumberFormat = "@";
                    wSheet.Cells[5 + k, 9] = item.Ticker;
                    wSheet.Cells[5 + k, 10] = item.BcastRef;
                    wSheet.Cells[5 + k, 11] = item.LegalName;
                    wSheet.Cells[5 + k, 12] = item.KoreaName;
                    wSheet.Cells[5 + k, 13] = item.Lotsize;

                    if (item.Type.Equals("PRF"))
                    {
                        mailAdd.Add(string.Format("Please notice {0} is a Preference Share", item.RIC));
                    }
                    else if (item.Type.Equals("KDR"))
                    {
                        mailAdd.Add(string.Format("Please notice {0} is a Korea Depositary Receipt", item.RIC));
                    }

                    if (item.IsGlobalETF)
                    {
                        mailAdd.Add(string.Format("Please notice {0} is a Global ETF", item.RIC));
                    }

                    foreach (int col in item.ChangeItems)
                    {
                        ((Range)wSheet.Cells[5 + k, col]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                    }
                }

                wSheet.Cells[7 + equity.Count, 1] = "- End -";
                ((Range)wSheet.Cells[7 + equity.Count, 1]).Font.Bold = System.Drawing.FontStyle.Bold;

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();


                MailToSend mail = new MailToSend();
                mail.ToReceiverList.AddRange(configObj.MailTo);
                mail.CCReceiverList.AddRange(configObj.MailCC);
                mail.MailSubject = Path.GetFileNameWithoutExtension(fileName);
                mail.AttachFileList.Add(filePath);
                mail.MailBody = mailType + " ADD:\t\t" + rics + "\r\n\r\n"
                                + "Effective Date:\t\t" + effectiveDate + "\r\n\r\n";
                if (mailAdd.Count > 0)
                {
                    string addMsg = string.Join("\r\n", mailAdd.ToArray());
                    mail.MailBody += addMsg;
                }

                mail.MailBody += "\r\n\r\n";

                string signature = string.Join("\r\n", configObj.MailSignature.ToArray());
                mail.MailBody += signature;

                AddResult(fileName, filePath, mailType + " ADD FM File");
                Logger.Log("Generate " + mailType + " ADD FM file successfully. Filepath is " + filePath);
            }
            catch (Exception ex)
            {
                string msg = "At GenerateFMFileForEach(). Error found in Generate FM file." + ex.Message;
                errors.Add(msg);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private string CombineAllRics(List<KoreaEquityInfo> items)
        {
            List<string> rics = new List<string>();

            foreach (KoreaEquityInfo e in items)
            {
                rics.Add(e.RIC);
            }

            return string.Join(",", rics.ToArray());
        }

        #endregion

        #region Generate GEDA File

        /// <summary>
        /// Generate GEDA files.
        /// </summary>
        private void GenerateGEDAFiles()
        {
            Logger.Log("Generate GEDA files.");
            foreach (KoreaEquityInfo peoItem in paList)
            {
                if (peoItem.ExistsFM1 && peoItem.ChangeItems.Count > 0)
                {
                    GenerateGedaPeoChange(peoItem);
                    continue;
                }

                if (peoItem.ExistsFM1 && peoItem.ChangeItems.Count == 0)
                {
                    continue;
                }

                if (peoItem.Type.Equals("ORD"))
                {
                    GenerateGedaOrdFile(peoItem);
                }
                else if (peoItem.Type.Equals("PRF"))
                {
                    GenerateGedaPrfFile(peoItem);
                }
                else if (peoItem.Type.Equals("KDR"))
                {
                    GenerateGedaKdrFile(peoItem);
                }
            }

            foreach (KoreaEquityInfo etfItem in eaList)
            {
                if (etfItem.IsGlobalETF)
                {
                    GenerateGedaEtfGlobalFile(etfItem);
                }
                else
                {
                    GenerateGedaEtfFile(etfItem);
                }
            }

            foreach (KoreaEquityInfo reitItem in raList)
            {
                GenerateGedaReitFile(reitItem);
            }

            foreach (KoreaEquityInfo bcItem in bcList)
            {
                GenerateGedaBcFile(bcItem);
            }
            Logger.Log("Successfully generated GEDA files.");
        }

        private void GenerateGedaPeoChange(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC" };
            List<string> data = new List<string>() { item.RIC };

            if (item.ChangeItems.Contains(KoreaAddFMColumn.IDNDisplayName))
            {
                title.Add("DSPLY_NAME");
                data.Add(item.IDNDisplayName);
            }
            if (item.ChangeItems.Contains(KoreaAddFMColumn.ISIN))
            {
                title.Add("EX_SYMBOL");
                title.Add("#INSTMOD_#ISIN");
                data.Add(item.ISIN);
                data.Add(item.ISIN);
            }
            if (item.ChangeItems.Contains(KoreaAddFMColumn.KoreaName))
            {
                title.Add("DSPLY_NMLL");
                data.Add(item.KoreaName);
            }
            if (item.ChangeItems.Contains(KoreaAddFMColumn.LegalName))
            {
                title.Add("#INSTMOD_TDN_ISSUER_NAME");
                data.Add(item.LegalName);
            }

            string fileName = "KR_GEDA_Bulk_Change_" + item.RIC + ".txt";
            string filePath = Path.Combine(configObj.PEOGEDA, fileName);
            try
            {
                FileUtil.WriteSingleLine(filePath, data, title);
                AddResult(fileName, filePath, "GEDA Change File (PEO)");
                string msg = string.Format("Generate GEDA change File successfully. FilePath:{0}", filePath);
                Logger.Log(msg, Logger.LogType.Info);
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GenerateGedaPeoChange(). Error found in generating GEDA change file. {0}", ex.Message);
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Use the ETF's legal name to check if it is a global ETF.
        /// </summary>
        /// <param name="legalName">ETF's legal name</param>
        /// <returns>true or false</returns>
        private bool CheckGlobalETF(string legalName)
        {
            legalName = KoreaEquityCommon.ClearCoLtdForName(legalName);
            List<string> subNames = legalName.Split(' ').ToList();
            string tableName = "ETI_Country_Name";

            foreach (string name in subNames)
            {
                string where = string.Format("where EnglishNameUpper = '{0}' or CountryCode = '{0}' or Remark = '{0}'", name);
                System.Data.DataTable dt = ManagerBase.Select(tableName, new string[] { "*" }, where);
                if (dt == null || dt.Rows.Count == 0)
                {
                    continue;
                }
                else
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Generate BC GEDA file for an bc equity.
        /// </summary>
        /// <param name="item">bc equity</param>
        private void GenerateGedaBcFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>();
            data.Add(item.RIC);
            data.Add(item.IDNDisplayName);
            data.Add(item.RIC);
            data.Add(item.Ticker);
            data.Add(item.ISIN);
            data.Add("****");
            data.Add(item.KoreaName);
            data.Add(item.RIC);
            data.Add(item.ISIN);
            data.Add("F" + item.Ticker);
            data.Add("115");
            data.Add(item.LegalName.ToUpper());
            data.Add(item.Ticker);
            data.Add("KSE_EQB_BC");

            string fileName = "KR_BC_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
            string filePath = Path.Combine(configObj.BCGEDA, fileName);
            try
            {
                FileUtil.WriteSingleLine(filePath, data, title);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "GEDA ADD File (BC)", filePath, FileProcessType.GEDA_BULK_RIC_CREATION);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate BC ADD ADD File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "At GenerateGedaBcFile(). Error found in generate BC ADD GEDA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate REIT GEDA file for a REIT equity.
        /// </summary>
        /// <param name="item">REIT equity</param>
        private void GenerateGedaReitFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_RECORDTYPE", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>();
            data.Add(item.RIC);
            data.Add(item.IDNDisplayName);
            data.Add(item.RIC);
            data.Add(item.Ticker);
            data.Add(item.ISIN);
            data.Add("****");
            data.Add(item.KoreaName);
            data.Add(item.RIC);
            data.Add(item.ISIN);
            data.Add("A" + item.Ticker);
            data.Add("112");
            data.Add("115");
            data.Add(item.LegalName.ToUpper());
            data.Add(item.Ticker);
            data.Add("KSE_EQB_3");

            string fileName = "KR_REIT_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
            string filePath = Path.Combine(configObj.REITGEDA, fileName);
            try
            {
                FileUtil.WriteSingleLine(filePath, data, title);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "GEDA ADD File (REIT)", filePath, FileProcessType.GEDA_BULK_RIC_CREATION);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate REIT ADD GEDA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "At GenerateGedaReitFile(). Error found in generate REIT ADD GEDA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate ETF GEDA file for an ETF equity.
        /// </summary>
        /// <param name="item">an ETF equity</param>
        private void GenerateGedaEtfFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>();
            data.Add(item.RIC);
            data.Add(item.IDNDisplayName);
            data.Add(item.RIC);
            data.Add(item.Ticker);
            data.Add(item.ISIN);
            data.Add("****");
            data.Add(item.KoreaName);
            data.Add(item.RIC);
            data.Add(item.ISIN);
            data.Add("A" + item.Ticker);
            data.Add("115");
            data.Add(item.LegalName.ToUpper());
            data.Add(item.Ticker);
            data.Add("KSE_EQB_ETF");

            string fileName = "KR_ETF_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
            string filePath = Path.Combine(configObj.ETFGEDA, fileName);
            try
            {
                FileUtil.WriteSingleLine(filePath, data, title);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "GEDA ADD File (ETF)", filePath, FileProcessType.GEDA_BULK_RIC_CREATION);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate ETF ADD GEDA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "At GenerateGedaEtfFile(). Error found in generate ETF ADD GEDA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate ETF GEDA file for an ETF global equity.
        /// </summary>
        /// <param name="item">an ETF global equity</param>
        private void GenerateGedaEtfGlobalFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_PROD_PERM", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>();
            data.Add(item.RIC);
            data.Add(item.IDNDisplayName);
            data.Add(item.RIC);
            data.Add(item.Ticker);
            data.Add(item.ISIN);
            data.Add("****");
            data.Add(item.KoreaName);
            data.Add(item.RIC);
            data.Add(item.ISIN);
            data.Add("A" + item.Ticker);
            data.Add("115");
            data.Add(item.LegalName.ToUpper());
            data.Add("6688");
            data.Add(item.Ticker);
            data.Add("KSE_EQB_ETF");

            string fileName = "KR_ETF_GLOBAL_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
            string filePath = Path.Combine(configObj.ETFGEDA, fileName);
            try
            {
                FileUtil.WriteSingleLine(filePath, data, title);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "GEDA ADD File (ETF Global)", filePath, FileProcessType.GEDA_BULK_RIC_CREATION);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate ETF GLOBAL ADD GEDA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "At GenerateGedaEtfGlobalFile(). Error found in generate ETF GLOBAL ADD GEDA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate KDR GEDA file for a KDR equity.
        /// </summary>
        /// <param name="item">a KDR equity</param>
        private void GenerateGedaKdrFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>();
            data.Add(item.RIC);
            data.Add(item.IDNDisplayName);
            data.Add(item.RIC);
            data.Add(item.Ticker);
            data.Add(item.ISIN);
            data.Add(item.KoreaName);
            data.Add(item.RIC);
            data.Add(item.ISIN);
            data.Add("A" + item.Ticker);
            data.Add(item.LegalName.ToUpper());
            data.Add(item.Ticker);
            if (item.Market.Equals("KS"))
            {
                data.Add("KSE_EQB_KDR");
            }
            else if (item.Market.Equals("KQ"))
            {
                data.Add("KOSDAQ_EQB_KDR");
            }
            else
            {
                data.Add("KONEX_EQB_KDR");
            }
            string fileName = "KR_KDR_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
            string filePath = Path.Combine(configObj.PEOGEDA, fileName);
            try
            {
                FileUtil.WriteSingleLine(filePath, data, title);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "GEDA ADD File (KDR)", filePath, FileProcessType.GEDA_BULK_RIC_CREATION);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate KDR ADD GEDA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "At GenerateGedaKdrFile(). Error found in generate KDR ADD GEDA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate GEDA file for a PRF equity.
        /// </summary>
        /// <param name="item">a PRF equity</param>
        private void GenerateGedaPrfFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>();
            data.Add(item.RIC);
            data.Add(item.IDNDisplayName);
            data.Add(item.RIC);
            data.Add(item.Ticker);
            data.Add(item.ISIN);
            data.Add("****");
            data.Add(item.KoreaName);
            data.Add(item.RIC);
            data.Add(item.ISIN);
            data.Add("A" + item.Ticker);
            data.Add("82");
            data.Add(item.LegalName.ToUpper());
            data.Add(item.Ticker);
            if (item.Market.Equals("KS"))
            {
                data.Add("KSE_EQLB");
            }
            else if (item.Market.Equals("KQ"))
            {
                data.Add("KOSDAQ_EQLB");
            }
            else
            {
                data.Add("KONEX_EQLB");
            }
            string fileName = "KR_PRF_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
            string filePath = Path.Combine(configObj.PEOGEDA, fileName);
            try
            {
                FileUtil.WriteSingleLine(filePath, data, title);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "GEDA ADD File (PRF)", filePath, FileProcessType.GEDA_BULK_RIC_CREATION);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate PRF ADD GEDA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "At GenerateGedaPrfFile(). Error found in generate PRF ADD GEDA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate KDR GEDA file for an ORD equity.
        /// </summary>
        /// <param name="item">an ORD equity</param>
        private void GenerateGedaOrdFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>();
            data.Add(item.RIC);
            data.Add(item.IDNDisplayName);
            data.Add(item.RIC);
            data.Add(item.Ticker);
            data.Add(item.ISIN);
            data.Add("****");
            data.Add(item.KoreaName);
            data.Add(item.RIC);
            if (item.Market.Equals("KN"))
            {
                data.Add("");
                data.Add("A" + item.Ticker);
                data.Add("");
                data.Add("");
                data.Add("KONEX_EQB_1");
            }
            else
            {
                data.Add(item.ISIN);
                data.Add("A" + item.Ticker);
                data.Add(item.LegalName.ToUpper());
                data.Add(item.Ticker);
                if (item.Market.Equals("KS"))
                {
                    //data.Add("KSE_EQB_3");
                    data.Add("KSE_EQB_2");
                }
                else if (item.Market.Equals("KQ"))
                {
                    data.Add("KOSDAQ_EQB_2");
                }
            }

            string fileName = "KR_ORD_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
            string filePath = Path.Combine(configObj.PEOGEDA, fileName);
            try
            {
                FileUtil.WriteSingleLine(filePath, data, title);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "GEDA ADD File (ORD)", filePath, FileProcessType.GEDA_BULK_RIC_CREATION);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate ORD ADD GEDA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "At GenerateGedaOrdFile(). Error found in generate ORD ADD GEDA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        #endregion

        #region Generate NDA File

        /// <summary>
        /// Generate NDA files.
        /// </summary>
        private void GenerateNDAFiles()
        {
            Logger.Log("Generate NDA files.");

            foreach (KoreaEquityInfo item in paList)
            {
                if (item.ExistsFM1)
                {
                    GenerateNdaPeoFM1File(item);
                }
                else
                {
                    GenerateNdaPeoFile(item);
                    GenerateNdaFutureDatingFile(item);
                }
            }

            foreach (KoreaEquityInfo item in eaList)
            {
                GenerateNdaEtfFile(item);
                GenerateNdaBGFile(item);
                GenerateNdaEtfFutureDatingFile(item);

            }

            foreach (KoreaEquityInfo item in raList)
            {
                GenerateNdaReitFile(item);
                GenerateNdaReitFutureDatingFile(item);
            }

            foreach (KoreaEquityInfo item in bcList)
            {
                GenerateNdaBcFile(item);
                GenerateNdaBcFutureDatingFile(item);
            }
            Logger.Log("Successfully generated NDA files.");
        }

        /// <summary>
        /// Generate NDA QA file for a PEO equity with FM1.
        /// </summary>
        /// <param name="item">a PEO equity</param>
        private void GenerateNdaPeoFM1File(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD" };
            string[] rics = new string[6] { ".", "F.", "S.", "stat.", "ta.", "bl." };
            string[] lotSizes = new string[6] { item.Lotsize, "", "1", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            List<List<string>> data = new List<List<string>>();

            if (peoChanged)
            {
                title.Add("ASSET SHORT NAME");
                title.Add("ASSET COMMON NAME");

                for (int i = 0; i < 6; i++)
                {
                    if (i == 2 && item.Market.Equals("KN"))
                    {
                        continue;
                    }
                    List<string> content = new List<string>();
                    content.Add(item.Ticker + rics[i] + item.Market);
                    content.Add(effectiveDate);
                    content.Add(lotSizes[i]);
                    content.Add("T+2");
                    content.Add(item.IDNDisplayName);
                    content.Add(item.IDNDisplayName + " " + item.Type);
                    data.Add(content);
                }
            }
            else
            {
                for (int i = 0; i < 6; i++)
                {
                    if (i == 2 && item.Market.Equals("KN"))
                    {
                        continue;
                    }
                    List<string> content = new List<string>();
                    content.Add(item.Ticker + rics[i] + item.Market);
                    content.Add(effectiveDate);
                    content.Add(lotSizes[i]);
                    content.Add("T+2");
                    data.Add(content);
                }
            }

            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAChg.csv";
            string filePath = Path.Combine(configObj.PEONDA, fileName);
            try
            {
                FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Append);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA QA Change File (PEO)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate PEO NDA QA Chg File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate PEO NDA QA Chg file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate NDA QA file for a PEO equity without FM1.
        /// </summary>
        /// <param name="item">a PEO equity</param>
        private void GenerateNdaFutureDatingFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC", "PROPERTY NAME", "PROPERTY VALUE", "EFFECTIVE FROM", "EFFECTIVE TO", "CHANGE OFFSET", "CHANGE TRIGGER", "CORAX PERMID" };

            string[] rics = new string[6] { ".", "F.", "S.", "stat.", "ta.", "bl." };

            //string[] tags = null;
            //if (item.Market.Equals("KQ"))
            //{
            //    tags = new string[6] { "673", "64399", "60673", "61287", "64380", "67094" };
            //}
            //else if (item.Market.Equals("KS"))
            //{
            //    tags = new string[6] { "184", "64398", "60184", "61286", "64379", "67093" };
            //}
            //else if (item.Market.Equals("KN"))
            //{
            //    tags = new string[6] { "67105", "67107", "*****", "67108", "64380", "67106" };
            //}
            //else
            //{
            //    string msg = "At GenerateNdaPeoFile(). Can not get the exchange board code.";
            //    errors.Add(msg);
            //    return;
            //}

            string[] lotSizes = new string[6] { item.Lotsize, "", "1", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            List<List<string>> data = new List<List<string>>();
            for (int i = 0; i < 6; i++)
            {
                if (i == 2 && item.Market.Equals("KN"))
                {
                    continue;
                }
                List<string> content1 = new List<string>();
                List<string> content2 = new List<string>();
                List<string> content3 = new List<string>();

                string ricname = item.Ticker + rics[i] + item.Market;
                //col1
                content1.Add(ricname);
                content2.Add(ricname);
                content3.Add(ricname);

                //col2
                content1.Add("");
                content2.Add("RIC");
                content3.Add("TICKER SYMBOL");
                //col3
                content1.Add("");
                content2.Add(ricname);
                content3.Add(item.Ticker);
                //col4
                content1.Add(effectiveDate);
                content2.Add(effectiveDate);
                content3.Add(effectiveDate);

                //col5
                content1.Add("");
                content2.Add("");
                content3.Add("");

                //col6
                content1.Add("");
                content2.Add("");
                content3.Add("");

                //col7
                content1.Add("");
                content2.Add("PEO");
                content3.Add("PEO");

                //col8
                content1.Add("");
                content2.Add("");
                content3.Add("");

                data.Add(content1);
                data.Add(content2);

                string lots = lotSizes[i];
                if (lots != null && lots != "")
                {
                    List<string> content = new List<string>();

                    content.Add(ricname);
                    content.Add("ROUND LOT SIZE");
                    content.Add(lots);
                    content.Add(effectiveDate);
                    content.Add("");
                    content.Add("");
                    content.Add("PEO");
                    content.Add("");
                    data.Add(content);
                }

                data.Add(content3);

                //content.Add(tags[i]);
                //content.Add("ISIN:" + item.ISIN);
                //content.Add(item.Ticker);
                //content.Add(item.IDNDisplayName);
                //content.Add(item.IDNDisplayName + " " + item.Type);
                //content.Add("EQUITY");
                //content.Add(item.Category);
                //content.Add("KRW");
                //content.Add(item.Exchange);
                //content.Add(effectiveDate);
                //content.Add(lotSizes[i]);
                //content.Add("T+2");
                //  data.Add(content);
            }

            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "PEOFutureDating.csv";
            string filePath = Path.Combine(configObj.PEONDA, fileName);
            try
            {
                FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Append);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA FutureDating File (PEO)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate PEO ADD NDA  FutureDating File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate PEO ADD NDA FutureDating  FutureDating file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        private void GenerateNdaPeoFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC", "TAG", "BASE ASSET","TICKER SYMBOL", "ASSET SHORT NAME", "ASSET COMMON NAME", "TYPE",	
                                                      "CATEGORY","CURRENCY", "EXCHANGE", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD","PRIMARY TRADABLE MARKET QUOTE"};
            string[] rics = new string[6] { ".", "F.", "S.", "stat.", "ta.", "bl." };
            string[] tags = null;
            if (item.Market.Equals("KQ"))
            {
                tags = new string[6] { "673", "64399", "60673", "61287", "64380", "67094" };
            }
            else if (item.Market.Equals("KS"))
            {
                tags = new string[6] { "184", "64398", "60184", "61286", "64379", "67093" };
            }
            else if (item.Market.Equals("KN"))
            {
                tags = new string[6] { "67105", "67107", "*****", "67108", "64380", "67106" };
            }
            else
            {
                string msg = "At GenerateNdaPeoFile(). Can not get the exchange board code.";
                errors.Add(msg);
                return;
            }

            string[] lotSizes = new string[6] { item.Lotsize, "", "1", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            List<List<string>> data = new List<List<string>>();
            for (int i = 0; i < 6; i++)
            {
                if (i == 2 && item.Market.Equals("KN"))
                {
                    continue;
                }
                List<string> content = new List<string>();
                content.Add(item.Ticker + rics[i] + item.Market);
                content.Add(tags[i]);
                content.Add("ISIN:" + item.ISIN);
                content.Add(item.Ticker);
                content.Add(item.IDNDisplayName);
                content.Add(item.IDNDisplayName + " " + item.Type);
                content.Add("EQUITY");
                content.Add(item.Category);
                content.Add("KRW");
                content.Add(item.Exchange);
                content.Add(effectiveDate);
                content.Add(lotSizes[i]);
                content.Add("T+2");

                if (i.CompareTo(0) == 0)
                    content.Add("Y");
                else
                    content.Add(string.Empty);

                data.Add(content);
            }

            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAAddPEO.csv";
            string filePath = Path.Combine(configObj.PEONDA, fileName);
            try
            {
                FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Append);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA QA ADD File (PEO)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate PEO ADD NDA QA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate PEO ADD NDA QA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }
        /// <summary>
        /// Generate NDA QA file for a ETF equity.
        /// </summary>
        /// <param name="item">a ETF equity</param>
        private void GenerateNdaEtfFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC", "TAG", "BASE ASSET","TICKER SYMBOL", "ASSET SHORT NAME", "ASSET COMMON NAME", "TYPE",	
                                                      "CATEGORY","CURRENCY", "EXCHANGE", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD","PRIMARY TRADABLE MARKET QUOTE"};
            List<List<string>> data = new List<List<string>>();
            string[] rics = new string[7] { ".KS", "F.KS", "S.KS", "stat.KS", "ta.KS", "bl.KS", "LP.KS" };
            string[] tags = new string[7] { "184", "64398", "60184", "61286", "64379", "67093", "64395" };
            string[] lotSizes = new string[7] { item.Lotsize, "", "1", "", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));

            for (int i = 0; i < 7; i++)
            {
                List<string> content = new List<string>();
                content.Add(item.Ticker + rics[i]);
                content.Add(tags[i]);
                content.Add("ISIN:" + item.ISIN);
                content.Add(item.Ticker);
                content.Add(item.IDNDisplayName);
                content.Add(item.IDNDisplayName + " ETF");
                content.Add("EQUITY");
                content.Add(item.Category);
                content.Add("KRW");
                content.Add("KSC");
                content.Add(effectiveDate);
                content.Add(lotSizes[i]);
                content.Add("T+2");

                if (i.CompareTo(0) == 0)
                    content.Add("Y");
                else
                    content.Add(string.Empty);

                data.Add(content);
            }

            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAAddETF.csv";
            string filePath = Path.Combine(configObj.ETFNDA, fileName);
            try
            {
                FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Append);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA QA ADD File (ETF)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate ETF ADD NDA QA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate ETF ADD NDA QA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        private void GenerateNdaEtfFutureDatingFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC", "PROPERTY NAME", "PROPERTY VALUE", "EFFECTIVE FROM", "EFFECTIVE TO", "CHANGE OFFSET", "CHANGE TRIGGER", "CORAX PERMID" };

            string[] rics = new string[7] { ".KS", "F.KS", "S.KS", "stat.KS", "ta.KS", "bl.KS", "LP.KS" };
            // string[] tags = new string[7] { "184", "64398", "60184", "61286", "64379", "67093", "64395" };
            string[] lotSizes = new string[7] { item.Lotsize, "", "1", "", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));

            List<List<string>> data = new List<List<string>>();
            for (int i = 0; i < 6; i++)
            {
                if (i == 2 && item.Market.Equals("KN"))
                {
                    continue;
                }
                List<string> content1 = new List<string>();
                List<string> content2 = new List<string>();
                List<string> content3 = new List<string>();

                string ricname = item.Ticker + rics[i] + item.Market;
                //col1
                content1.Add(ricname);
                content2.Add(ricname);
                content3.Add(ricname);

                //col2
                content1.Add("");
                content2.Add("RIC");
                content3.Add("TICKER SYMBOL");
                //col3
                content1.Add("");
                content2.Add(ricname);
                content3.Add(item.Ticker);
                //col4
                content1.Add(effectiveDate);
                content2.Add(effectiveDate);
                content3.Add(effectiveDate);

                //col5
                content1.Add("");
                content2.Add("");
                content3.Add("");

                //col6
                content1.Add("");
                content2.Add("");
                content3.Add("");

                //col7
                content1.Add("");
                content2.Add("PEO");
                content3.Add("PEO");

                //col8
                content1.Add("");
                content2.Add("");
                content3.Add("");

                data.Add(content1);
                data.Add(content2);

                string lots = lotSizes[i];
                if (lots != null && lots != "")
                {
                    List<string> content = new List<string>();

                    content.Add(ricname);
                    content.Add("ROUND LOT SIZE");
                    content.Add(lots);
                    content.Add(effectiveDate);
                    content.Add("");
                    content.Add("");
                    content.Add("PEO");
                    content.Add("");
                    data.Add(content);
                }
                data.Add(content3);
            }
            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "ETFFutureDating.csv";
            string filePath = Path.Combine(configObj.ETFNDA, fileName);
            try
            {
                FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Append);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA FutureDating File (ETF)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate ETF  NDA FutureDating File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate ETF  NDA FutureDating file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate NDA BG file for a ETF equity.
        /// </summary>
        /// <param name="item">a ETF equity</param>
        private void GenerateNdaBGFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "XXLEGAL NAMEXX", "XXBODY GROUP COMMON NAMEXX", "XXBODY GROUP SHORT NAMEXX", "XXORGANISATION TYPEXX", "XXCOUNTRY HEADQUARTERSXX" };
            List<string> data = new List<string>();
            string legalName = item.LegalName.Trim();
            if (!legalName.Contains("ETF"))
            {
                legalName += " ETF";
            }

            data.Add(legalName);
            data.Add(legalName);
            data.Add(item.IDNDisplayName);
            data.Add("FD");
            data.Add("KOR");

            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "BGAdd.csv";
            string filePath = Path.Combine(configObj.ETFNDA, fileName);
            try
            {
                FileUtil.WriteSingleLine(filePath, data, title);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA BG ADD File (ETF)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate ETF ADD NDA BG file successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate ETF ADD NDA BG file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate NDA QA file for a REIT equity.
        /// </summary>
        /// <param name="item">a REIT equity</param>
        private void GenerateNdaReitFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC", "TAG", "BASE ASSET","TICKER SYMBOL", "ASSET SHORT NAME", "ASSET COMMON NAME", "TYPE",	
                                                      "CATEGORY","CURRENCY", "EXCHANGE", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD","PRIMARY TRADABLE MARKET QUOTE"};
            List<List<string>> data = new List<List<string>>();
            string[] rics = new string[6] { ".KS", "F.KS", "S.KS", "stat.KS", "ta.KS", "bl.KS" };
            string[] tags = new string[6] { "184", "64398", "60184", "61286", "64379", "67093" };
            string[] lotSizes = new string[6] { item.Lotsize, "", "1", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));

            for (int i = 0; i < 6; i++)
            {
                List<string> content = new List<string>();
                content.Add(item.Ticker + rics[i]);
                content.Add(tags[i]);
                content.Add("ISIN:" + item.ISIN);
                content.Add(item.Ticker);
                content.Add(item.IDNDisplayName);
                content.Add(item.IDNDisplayName + " REIT");
                content.Add("EQUITY");
                content.Add("REI");
                content.Add("KRW");
                content.Add("KSC");
                content.Add(effectiveDate);
                content.Add(lotSizes[i]);
                content.Add("T+2");

                if (i.CompareTo(0) == 0)
                    content.Add("Y");
                else
                    content.Add(string.Empty);

                data.Add(content);
            }

            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAAddREIT.csv";
            string filePath = Path.Combine(configObj.REITNDA, fileName);
            try
            {
                FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Append);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA QA ADD File (REIT)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate REIT ADD NDA QA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate REIT ADD NDA QA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        private void GenerateNdaReitFutureDatingFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC", "PROPERTY NAME", "PROPERTY VALUE", "EFFECTIVE FROM", "EFFECTIVE TO", "CHANGE OFFSET", "CHANGE TRIGGER", "CORAX PERMID" };

            string[] rics = new string[6] { ".KS", "F.KS", "S.KS", "stat.KS", "ta.KS", "bl.KS" };
            // string[] tags = new string[6] { "184", "64398", "60184", "61286", "64379", "67093" };
            string[] lotSizes = new string[6] { item.Lotsize, "", "1", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            List<List<string>> data = new List<List<string>>();
            for (int i = 0; i < 6; i++)
            {
                if (i == 2 && item.Market.Equals("KN"))
                {
                    continue;
                }
                List<string> content1 = new List<string>();
                List<string> content2 = new List<string>();
                List<string> content3 = new List<string>();


                string ricname = item.Ticker + rics[i] + item.Market;
                //col1
                content1.Add(ricname);
                content2.Add(ricname);
                content3.Add(ricname);

                //col2
                content1.Add("");
                content2.Add("RIC");
                content3.Add("TICKER SYMBOL");
                //col3
                content1.Add("");
                content2.Add(ricname);
                content3.Add(item.Ticker);
                //col4
                content1.Add(effectiveDate);
                content2.Add(effectiveDate);
                content3.Add(effectiveDate);

                //col5
                content1.Add("");
                content2.Add("");
                content3.Add("");

                //col6
                content1.Add("");
                content2.Add("");
                content3.Add("");

                //col7
                content1.Add("");
                content2.Add("PEO");
                content3.Add("PEO");

                //col8
                content1.Add("");
                content2.Add("");
                content3.Add("");

                data.Add(content1);
                data.Add(content2);

                string lots = lotSizes[i];
                if (lots != null && lots != "")
                {
                    List<string> content = new List<string>();

                    content.Add(ricname);
                    content.Add("ROUND LOT SIZE");
                    content.Add(lots);
                    content.Add(effectiveDate);
                    content.Add("");
                    content.Add("");
                    content.Add("PEO");
                    content.Add("");
                    data.Add(content);
                }

                data.Add(content3);
            }

            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "REITFutureDating.csv";
            string filePath = Path.Combine(configObj.REITNDA, fileName);
            try
            {
                FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Append);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA FutureDating File (REIT)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate REIT ADD NDA QA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate REIT ADD NDA QA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }
        /// <summary>
        /// Generate NDA QA file for a BC equity.
        /// </summary>
        /// <param name="item">a BC equity</param>
        private void GenerateNdaBcFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC", "TAG", "BASE ASSET","TICKER SYMBOL", "ASSET SHORT NAME", "ASSET COMMON NAME", "TYPE",	
                                                      "CATEGORY","CURRENCY", "EXCHANGE", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD","PRIMARY TRADABLE MARKET QUOTE"};
            List<List<string>> data = new List<List<string>>();
            string[] rics = new string[2] { ".KS", "F.KS" };
            string[] tags = new string[2] { "184", "64398" };
            string[] lotSizes = new string[2] { "10", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));

            for (int i = 0; i < 2; i++)
            {
                List<string> content = new List<string>();
                content.Add(item.Ticker + rics[i]);
                content.Add(tags[i]);
                content.Add("ISIN:" + item.ISIN);
                content.Add(item.Ticker);
                content.Add(item.IDNDisplayName);
                content.Add(item.IDNDisplayName + " CEF");
                content.Add("EQUITY");
                content.Add("CEF");
                content.Add("KRW");
                content.Add("KSC");
                content.Add(effectiveDate);
                content.Add(lotSizes[i]);
                content.Add("T+2");

                if (i.CompareTo(0) == 0)
                    content.Add("Y");
                else
                    content.Add(string.Empty);

                data.Add(content);
            }

            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAAddBC.csv";
            string filePath = Path.Combine(configObj.BCNDA, fileName);
            try
            {
                FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Append);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA QA ADD File (BC)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate BC ADD NDA QA File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate BC ADD NDA QA file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        private void GenerateNdaBcFutureDatingFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>() { "RIC", "PROPERTY NAME", "PROPERTY VALUE", "EFFECTIVE FROM", "EFFECTIVE TO", "CHANGE OFFSET", "CHANGE TRIGGER", "CORAX PERMID" };

            string[] rics = new string[2] { ".KS", "F.KS" };
            // string[] tags = new string[2] { "184", "64398" };
            string[] lotSizes = new string[2] { "10", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            List<List<string>> data = new List<List<string>>();
            for (int i = 0; i < 6; i++)
            {
                if (i == 2 && item.Market.Equals("KN"))
                {
                    continue;
                }
                List<string> content1 = new List<string>();
                List<string> content2 = new List<string>();
                List<string> content3 = new List<string>();

                string ricname = item.Ticker + rics[i] + item.Market;
                //col1
                content1.Add(ricname);
                content2.Add(ricname);
                content3.Add(ricname);

                //col2
                content1.Add("");
                content2.Add("RIC");
                content3.Add("TICKER SYMBOL");
                //col3
                content1.Add("");
                content2.Add(ricname);
                content3.Add(item.Ticker);
                //col4
                content1.Add(effectiveDate);
                content2.Add(effectiveDate);
                content3.Add(effectiveDate);

                //col5
                content1.Add("");
                content2.Add("");
                content3.Add("");

                //col6
                content1.Add("");
                content2.Add("");
                content3.Add("");

                //col7
                content1.Add("");
                content2.Add("PEO");
                content3.Add("PEO");

                //col8
                content1.Add("");
                content2.Add("");
                content3.Add("");

                data.Add(content1);
                data.Add(content2);

                string lots = lotSizes[i];
                if (lots != null && lots != "")
                {
                    List<string> content = new List<string>();

                    content.Add(ricname);
                    content.Add("ROUND LOT SIZE");
                    content.Add(lots);
                    content.Add(effectiveDate);
                    content.Add("");
                    content.Add("");
                    content.Add("PEO");
                    content.Add("");
                }
                data.Add(content3);
            }

            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "BCFutureDating.csv";
            string filePath = Path.Combine(configObj.BCNDA, fileName);
            try
            {
                FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Append);
                if (!taskResult.Contains(fileName))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileName, "NDA FutureDatingFile (BC)", filePath, FileProcessType.NDA);
                    taskResult.Add(fileName, taskEntry);
                    string msg = string.Format("Generate BC ADD NDA FutureDating File successfully. FilePath:{0}", filePath);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in generate BC ADD NDA FutureDating file. File name:" + fileName + "\r\n\t\t" + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate NDA tick and lot files.
        /// </summary>
        private void GenerateNDATickLotFiles()
        {
            Logger.Log("Generate Tick and Lot NDA files.");
            string filePath = string.Empty;
            string tickLadderName = string.Empty;
            string lotLadderName = string.Empty;
            string fileNamePart = string.Empty;

            foreach (KoreaEquityInfo item in paList)
            {
                filePath = configObj.PEONDA;
                fileNamePart = "EQ";
                tickLadderName = "TICK_LADDER_KSC_1";
                lotLadderName = "LOT_LADDER_KSC_1";
                if (item.Market.Equals("KQ"))
                {
                    tickLadderName = "TICK_LADDER_KOE_1";
                    lotLadderName = "LOT_LADDER_EQTY_<1>";
                }
                else if (item.Market.Equals("KN"))
                {
                    tickLadderName = "TICK_LADDER_KOE_1";
                    lotLadderName = "LOT_LADDER_EQTY_<100>";
                }
                GenerateTickLotFile(item, tickLadderName, lotLadderName, filePath, fileNamePart);
            }

            foreach (KoreaEquityInfo item in eaList)
            {
                fileNamePart = "ETF";
                filePath = configObj.ETFNDA;
                tickLadderName = "TICK_LADDER_<5>";
                lotLadderName = "LOT_LADDER_EQTY_<1>";
                GenerateTickLotFile(item, tickLadderName, lotLadderName, filePath, fileNamePart);
            }

            foreach (KoreaEquityInfo item in raList)
            {
                fileNamePart = "REIT";
                filePath = configObj.REITNDA;
                tickLadderName = "TICK_LADDER_KSC_1";
                lotLadderName = "LOT_LADDER_KSC_1";
                GenerateTickLotFile(item, tickLadderName, lotLadderName, filePath, fileNamePart);
            }

            foreach (KoreaEquityInfo item in bcList)
            {
                fileNamePart = "BC";
                filePath = configObj.BCNDA;
                tickLadderName = "TICK_LADDER_KSC_1";
                lotLadderName = "LOT_LADDER_KSC_1";
                GenerateTickLotFile(item, tickLadderName, lotLadderName, filePath, fileNamePart);
            }

            Logger.Log("Generate Tick and Lot NDA files successfully.");
        }

        private void GenerateNDAFutureDating()
        {
            Logger.Log("Generate NDA Future Dating.");
        }

        /// <summary>
        /// Generate tick file and lot file for an equity item.
        /// </summary>
        /// <param name="item">an equity item</param>
        /// <param name="tickLadderName">value of tick ladder name column</param>
        /// <param name="lotLadderName">value of lot ladder name column</param>
        /// <param name="filePath">file path</param>
        private void GenerateTickLotFile(KoreaEquityInfo item, string tickLadderName, string lotLadderName, string filePath, string fileNamePart)
        {
            try
            {
                List<string> tickTitle = new List<string>(){"RIC", "TICK NOT APPLICABLE", "TICK LADDER NAME", 
                                                      "TICK EFFECTIVE FROM", "TICK EFFECTIVE TO", "TICK PRICE INDICATOR" };
                List<string> lotTitle = new List<string>(){"RIC", "LOT NOT APPLICABLE", "LOT LADDER NAME", 
                                                      "LOT EFFECTIVE FROM", "LOT EFFECTIVE TO", "LOT PRICE INDICATOR" };
                string today = DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US"));
                string fileNameTick = "TickAdd_" + fileNamePart + "_" + today + ".csv";
                string fileNameLot = "LotAdd_" + fileNamePart + "_" + today + ".csv";
                string filePathTick = Path.Combine(filePath, fileNameTick);
                string filePathLot = Path.Combine(filePath, fileNameLot);
                List<List<string>> tickContent = new List<List<string>>();
                List<List<string>> lotContent = new List<List<string>>();

                List<string> tickRecord = new List<string>();
                List<string> lotRecord = new List<string>();
                tickRecord.Add(item.RIC);
                lotRecord.Add(item.RIC);
                tickRecord.Add("N");
                lotRecord.Add("N");
                tickRecord.Add(tickLadderName);
                lotRecord.Add(lotLadderName);
                string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US"));
                tickRecord.Add(effectiveDate);
                lotRecord.Add(effectiveDate);
                tickRecord.Add("");
                lotRecord.Add("");
                tickRecord.Add("ORDER");
                lotRecord.Add("CLOSE");
                tickContent.Add(tickRecord);
                lotContent.Add(lotRecord);

                FileUtil.WriteOutputFile(filePathTick, tickContent, tickTitle, WriteMode.Append);
                FileUtil.WriteOutputFile(filePathLot, lotContent, lotTitle, WriteMode.Append);

                if (!taskResult.Contains(filePathTick))
                {
                    TaskResultEntry taskEntry = new TaskResultEntry(fileNameTick, "NDA Tick Ladder File", filePathTick, FileProcessType.NDA);
                    TaskResultEntry taskEntryLot = new TaskResultEntry(fileNameLot, "NDA Lot Ladder File", filePathLot, FileProcessType.NDA);
                    taskResult.Add(filePathTick, taskEntry);
                    taskResult.Add(filePathLot, taskEntryLot);
                    string msg = string.Format("Generate NDA Tick and Lot Ladder Files successfully. FilePath:\r\n{0}\r\n{1}", filePathTick, filePathLot);
                    Logger.Log(msg, Logger.LogType.Info);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in generating NDA Tick and Lot file. {0}", ex.Message);
                errors.Add(msg);
            }
        }

        #endregion

        #region Delete FM1 records in database
        /// <summary>
        /// Delete FM1 records for PEOs.
        /// </summary>
        private void DeletePeoFmOne()
        {
            foreach (KoreaEquityInfo item in paList)
            {
                if (!item.ExistsFM1)
                {
                    continue;
                }
                try
                {
                    KoreaEquityManager.DeleteEquityFMOne(item.RIC);
                    Logger.Log("Delete FM1 record for Equity:" + item.RIC);
                }
                catch (Exception ex)
                {
                    string msg = string.Format("At DeletePeoFmOne(). Error found in deleting equity FM1. RIC:{0}. Error message:{1}", item.RIC, ex.Message);
                    errors.Add(msg);
                }
            }
        }

        #endregion

        #endregion

        #region Results. For result list and log

        /// <summary>
        /// Create task result list for GEDA and NDA file.         
        /// </summary>
        private void CreateTaskResult()
        {
            ArrayList resultOrder = new ArrayList(taskResult.Keys);
            resultOrder.Sort();
            foreach (string fileName in resultOrder)
            {
                try
                {
                    TaskResultList.Add(taskResult[fileName] as TaskResultEntry);
                }
                catch (Exception ex)
                {
                    string msg = "Error found in CreateTaskResult()." + ex.Message;
                    errors.Add(msg);
                }
            }
        }

        /// <summary>
        /// Log errors in log files.
        /// </summary>
        private void LogErrors()
        {
            if (errors.Count > 0)
            {
                Logger.Log("\r\n\r\n ------------------------ERRORS-----------------------");
                foreach (string msg in errors)
                {
                    Logger.Log(msg, Logger.LogType.Error);
                }
            }
            AddResult("Log", Logger.FilePath, "LOG FILE");
        }

        #endregion

        #region [CoraxKoreaIPO]
        #region [Generate File]
        private void GenerateFile(Dictionary<string, List<string>> dicListXls, string file)
        {
            if (dicListXls == null || dicListXls.Count <= 1)
            {
                Logger.Log("No Data In DB!", Logger.LogType.Info);
                return;
            }

            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(file)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(file));
                }

                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("delete old the file :{0} error. msg:{1}", file, ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                string result = GetCsvString(dicListXls);
                File.WriteAllText(file, result);
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate the file :{0} error. msg:{1}", file, ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private string GetCsvString(Dictionary<string, List<string>> dicList)
        {
            string result = string.Empty;

            if (dicList == null || dicList.Count <= 1)//title must exist while no data in file
            {
                Logger.Log(string.Format("no data need to generate"), Logger.LogType.Error);
                return null;
            }

            StringBuilder sb = new StringBuilder();

            foreach (var item in dicList.Values.ToList())
            {
                foreach (var str in item)
                {
                    sb.AppendFormat("{0},", str.Replace(",", ""));
                }

                sb.Length = sb.Length - 1;
                sb.Append("\r\n");
            }

            result = sb.ToString();
            return result;
        }
        #endregion

        #region [Extract Data]
        private Dictionary<string, List<string>> ExtractData2(List<KORProcessItem> historicalAnnouncements)
        {
            Dictionary<string, List<string>> dicList = new Dictionary<string, List<string>>();
            dicList.Add("title", listTitle2);
            string pageSource = string.Empty;
            string votingRightsPerShare = string.Empty;
            string pageTicker = string.Empty;
            string pageSourceWithHtmlTag = string.Empty;

            if (historicalAnnouncements == null || historicalAnnouncements.Count == 0)
            {
                string msg = string.Format("no historicalAnnouncements in Ace DB from dateTime.");
                Logger.Log(msg, Logger.LogType.Warning);
                return null;
            }

            foreach (var item in historicalAnnouncements)
            {
                pageSourceWithHtmlTag = GetDataPage(item.SourceLink);
                pageSource = ClearHtml(pageSourceWithHtmlTag);
                pageTicker = GetPageTicker(pageSource);

                if (string.IsNullOrEmpty(pageSource))
                    continue;

                if (IsExistEFT(pageSource))//new requirements:if 영문 종목명 contain "EFT" invalid
                    continue;

                if (IsExistBond(pageSource))//new requirements:if 한글종목명 contain "채권" invalid
                    continue;

                if (IsExistKDR(pageSourceWithHtmlTag))
                    item.Content = "(KDR)";
                else
                    item.Content = string.Empty;

                if (!string.IsNullOrEmpty(pageTicker))//if exist pageticker
                {
                    if (item.Ticker.Contains(pageTicker))
                    {
                        dicList.Add(item.ID.ToString(), GetColumn2(item, pageSource));
                    }
                    else
                    {
                        item.Ticker = FormatPageTicker(item, pageTicker);
                        dicList.Add(item.ID.ToString(), GetColumn2(item, pageSource));
                    }
                }
                else
                {
                    //new requirements: if contains "보통주" and "우선주" generate two records with different votingRightsPerShare
                    if (IsDoubleRcords(pageSource))
                    {
                        //just to for look the same as the other file
                        item.TXT_LocalPath = "(preferred)";//use TXT_LocalPath to see if need to add "(preferred)"
                        dicList.Add(item.ID.ToString() + "0", GetColumn2(item, pageSource + "(preferred)"));
                        dicList.Add(item.ID.ToString() + "1", GetColumn2(item, pageSource));
                    }
                    else
                    {
                        dicList.Add(item.ID.ToString(), GetColumn2(item, pageSource));
                    }
                }
            }

            return dicList;
        }

        private List<string> GetColumn2(KORProcessItem item, string pageSource)
        {
            //string pageSource = GetDataPage(item.SourceLink);
            string provider = "Not Found";
            HtmlDocument htc = GetDataHtc(item.SourceLink);
            string ticker = string.Empty;

            if (item.Ticker.Contains(".KS"))
                provider = "KSC";
            else if (item.Ticker.Contains(".KQ"))
                provider = "KOE";
            else if (item.Ticker.Contains(".KN"))
                provider = "KNX";

            if (pageSource != null)
                pageSource = pageSource.Replace(" ", "");

            if (!string.IsNullOrEmpty(item.Content) && item.Content.Contains("(KDR)"))
                ticker = item.Ticker + "(KDR)";
            else
                ticker = item.Ticker;

            if (!string.IsNullOrEmpty(item.TXT_LocalPath) && item.TXT_LocalPath.Contains("(preferred)") && pageSource.Contains("(preferred)"))
                ticker = ticker + "(preferred)";

            List<string> list = new List<string>();
            //list.Add(string.IsNullOrWhiteSpace(item.Ticker) ? "Not Found" : item.Ticker); //RIC
            list.Add(ticker);//RIC
            list.Add("ISS");        //TSO_TYPE                                                                
            list.Add(GetListingDate(pageSource) == null ? "Not Found" : GetListingDate(pageSource));        //TSO_DATE                      listingdata                                                  
            list.Add((GetIssueSpecShares(htc, pageSource) == null ? "Not Found" : GetIssueSpecShares(htc, pageSource)));        //ISSUE_SPEC_SHARES                     tso                                                           
            list.Add("");        //UNLISTED_SHARES                                                                
            list.Add("IPO");        //CAC_MA_COMMENTS                                                                
            list.Add("");        //EXTERNAL_DESCRIPTION                                                                
            list.Add("");        //SOURCE_ID                                                                        
            list.Add("STX");        //SOURCE_TYPE                                                             
            list.Add(item.AnnouceDate.AddHours(-9).ToString("dd-MM-yyyy HH:mm").Replace("-", "/"));        //RELEASE_DATE                        ACE/Release date/time                                                 
            list.Add(item.AnnouceDate.ToString("dd-MM-yyyy HH:mm").Replace("-", "/"));        //LOCAL_DATE                        ACE/Local date/time                                               
            list.Add("KST");        //TIMEZONE_NAME                                                                
            list.Add(provider);        //SOURCE_PROVIDER                                                                
            list.Add("");        //SOURCE_LINK                                                                
            list.Add("");        //SOURCE_DESCRIPTION"      

            return list;
        }

        private Dictionary<string, List<string>> ExtractData1(List<KORProcessItem> historicalAnnouncements)
        {
            Dictionary<string, List<string>> dicList = new Dictionary<string, List<string>>();
            dicList.Add("title", listTitle1);
            string pageSource = string.Empty;
            string votingRightsPerShare = string.Empty;
            string pageTicker = string.Empty;
            string pageSourceWithHtmlTag = string.Empty;

            if (historicalAnnouncements == null || historicalAnnouncements.Count == 0)
            {
                string msg = string.Format("no historicalAnnouncements in Ace DB from dateTime.");
                Logger.Log(msg, Logger.LogType.Warning);
                return null;
            }

            foreach (var item in historicalAnnouncements)
            {
                pageSourceWithHtmlTag = GetDataPage(item.SourceLink);
                pageSource = ClearHtml(pageSourceWithHtmlTag);
                //pageSource = GetDataPage(item.SourceLink);
                pageTicker = GetPageTicker(pageSource);

                if (string.IsNullOrEmpty(pageSource))
                    continue;

                if (IsExistEFT(pageSource))//new requirements:if 영문 종목명 contain "EFT" invalid
                    continue;

                if (IsExistBond(pageSource))//new requirements:if 한글종목명 contain "채권" invalid
                    continue;

                if (IsExistKDR(pageSourceWithHtmlTag))
                    item.Content = "(KDR)";//use content to see if need to add "(KDR)"
                else
                    item.Content = string.Empty;

                if (!string.IsNullOrEmpty(pageTicker))// if exist pageticker
                {
                    if (item.Ticker.Contains(pageTicker))
                    {
                        dicList.Add(item.ID.ToString(), GetColumn1(item, pageSource, "1"));
                    }
                    else
                    {
                        item.Ticker = FormatPageTicker(item, pageTicker);
                        dicList.Add(item.ID.ToString(), GetColumn1(item, pageSource, "0"));
                    }
                }
                else
                {
                    if (IsDoubleRcords(pageSource))//new requirements: if contains "보통주" and "우선주" generate two records with different votingRightsPerShare
                    {
                        item.TXT_LocalPath = "(preferred)";//use TXT_LocalPath to see if need to add "(preferred)"
                        dicList.Add(item.ID.ToString() + "(preferred)", GetColumn1(item, pageSource, "0"));
                        dicList.Add(item.ID.ToString() + "1", GetColumn1(item, pageSource, "1"));
                    }
                    else
                    {
                        votingRightsPerShare = GetPreferenceShare(pageSource);
                        dicList.Add(item.ID.ToString(), GetColumn1(item, pageSource, votingRightsPerShare));
                    }
                }
            }

            return dicList;
        }

        private bool IsExistKDR(string pageSource)
        {
            bool result = false;
            string pageString = pageSource.Replace(" ", "");

            if (pageString.Contains(">KDR<") || pageString.Contains("(원주:KDR)"))
                result = true;

            return result;
        }

        private string FormatPageTicker(KORProcessItem item, string pageTicker)
        {
            string result = string.Empty;

            if (item.Ticker.Contains(".KS"))
                result = pageTicker + ".KS";
            else if (item.Ticker.Contains(".KQ"))
                result = pageTicker + ".KQ";
            else if (item.Ticker.Contains(".KN"))
                result = pageTicker + ".KN";
            else
                result = pageTicker;

            return result;
        }

        private bool IsDoubleRcords(string pageSource)
        {
            bool result = false;

            if (pageSource.Contains("보통주") && pageSource.Contains("우선주"))
                result = true;

            return result;
        }

        private string GetPreferenceShare(string p)//보통주
        {
            string result = null;

            if (string.IsNullOrWhiteSpace(p))
                return result;

            if (p.Contains("보통주") && p.Contains("우선주"))
            {
                int index1 = p.IndexOf("보통주");
                int index0 = p.IndexOf("우선주");

                if (index1 <= index0)
                    return "1";
                else
                    return "0";
            }

            if (p.Contains("우선주"))
                result = "0";

            if (p.Contains("보통주"))
                result = "1";

            return result;
        }

        private bool IsExistEFT(string pageSource)
        {
            bool result = false;
            string pageString = pageSource.Replace(" ", "");
            string englishName = string.Empty;
            int index = pageString.IndexOf("영문종목명");

            if (index < 0)
                return result;

            englishName = pageString.Substring(index, 100);

            if (englishName.Contains("ETF"))
                result = true;

            return result;
        }

        private bool IsExistBond(string pageSource)
        {
            bool result = false;
            string pageString = pageSource.Replace(" ", "");
            string englishName = string.Empty;
            int index = pageString.IndexOf("한글종목명");

            if (index < 0)
                return result;

            englishName = pageString.Substring(index, 100);

            if (englishName.Contains("채권"))
                result = true;

            return result;
        }

        private List<string> GetColumn1(KORProcessItem item, string pageSource, string votingRightsPerShare)
        {
            //string pageSource = GetDataPage(item.SourceLink);
            string provider = "Not Found";
            string ticker = string.Empty;

            if (item.Ticker.Contains(".KS"))
                provider = "KSC";
            else if (item.Ticker.Contains(".KQ"))
                provider = "KOE";
            else if (item.Ticker.Contains(".KN"))
                provider = "KNX";

            if (pageSource != null)
                pageSource = pageSource.Replace(" ", "");

            if (!string.IsNullOrEmpty(item.Content) && item.Content.Contains("(KDR)"))
                ticker = item.Ticker + "(KDR)";
            else
                ticker = item.Ticker;

            if (!string.IsNullOrEmpty(item.TXT_LocalPath) && item.TXT_LocalPath.Contains("(preferred)") && votingRightsPerShare.Equals("0"))
                ticker = ticker + "(preferred)";

            List<string> list = new List<string>();
            //list.Add(string.IsNullOrWhiteSpace(item.Ticker) ? "Not Found" : item.Ticker);//RIC
            list.Add(ticker);//RIC
            //list.Add(GetListingDate(pageSource) == null ? "Not Found" : DateFormat(GetListingDate(pageSource)));//CHANGE_DATE
            list.Add(GetListingDate(pageSource) == null ? "Not Found" : GetListingDate(pageSource));//CHANGE_DATE
            list.Add(GetNominalValue(pageSource) == null ? "Not Found" : GetNominalValue(pageSource));//NOMINAL_VALUE
            list.Add("N");//NO_PAR_VALUE
            list.Add("KRW");//CLA_CUR_VAL
            list.Add("IPO");//CAC_MA_COMMENTS
            list.Add(string.IsNullOrEmpty(votingRightsPerShare) ? "Not Found" : votingRightsPerShare);//VOTING_RIGHTS_PER_SHARE
            //list.Add(GetListingDate(pageSource) == null ? "Not Found" : GetListingDate(pageSource));//VRI_CHANGE_DATE
            list.Add(list[1]);//VRI_CHANGE_DATE
            list.Add("");//VOTING_RIGHTS_DESCRIPTION
            list.Add("");//CONVERSION_RATIO
            list.Add("");//CRA_CHANGE_DATE
            list.Add("");//SOURCE_ID
            list.Add("STX");//SOURCE_TYPE
            list.Add(item.AnnouceDate.AddHours(-9).ToString("dd-MM-yyyy HH:mm").Replace("-", "/"));//RELEASE_DATE
            list.Add(item.AnnouceDate.ToString("dd-MM-yyyy HH:mm").Replace("-", "/"));//LOCAL_DATE
            list.Add("KST");//TIMEZONE_NAME
            list.Add(provider);//SOURCE_PROVIDER
            list.Add("");//SOURCE_LINK
            list.Add("");//SOURCE_DESCRIPTION

            return list;
        }

        private string GetPageTicker(string p)
        {
            string result = null;

            if (string.IsNullOrWhiteSpace(p))
                return result;

            p = p.Replace(" ", "");
            List<string> listPattern = new List<string>();
            //8. 종목코드 : 표준코드 KR700781K015 (단축코드 : A00781K)
            listPattern.Add(@"단축코드\:?[A-Z](?<ticker>[0-9A-Z]{6})");

            try
            {
                foreach (var item in listPattern)
                {
                    Regex reDate = new Regex(item);
                    MatchCollection maDate = reDate.Matches(p);

                    if (maDate.Count < 1)
                        continue;

                    foreach (Match da in maDate)
                    {
                        return da.Groups["ticker"].Value;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in getting ticker. msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }
        }

        private string GetDataPage(string p)
        {
            try
            {
                string second = string.Empty;
                string code = string.Empty;
                Regex reCode = new Regex(@"acptno\=(?<Code>\d+)\&docno");
                MatchCollection maCodes = reCode.Matches(p);

                if (maCodes.Count < 1)
                    return null;

                foreach (Match item in maCodes)
                {
                    code = item.Groups["Code"].Value;
                    break;
                }

                string source = WebClientUtil.GetDynamicPageSource(p, 300000, null);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();

                if (string.IsNullOrWhiteSpace(source))
                    return null;

                doc.LoadHtml(source);

                if (doc == null)
                    return null;

                second = doc.DocumentNode.SelectSingleNode(".//select[@id='mainDoc']/option[2]").Attributes["value"].Value.Trim().ToString();
                second = second.Trim().ToString().Replace("|Y", "");
                string url = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=searchContents&docNo={0}", second);
                string pageSource = WebClientUtil.GetPageSource(null, url, 180000, null, Encoding.GetEncoding("UTF-8"));
                Regex regex = new Regex(@"http://kind.krx.co.kr/.+htm");
                Match match = regex.Match(pageSource);

                if (!match.Success)
                    return null;

                pageSource = WebClientUtil.GetPageSource(null, match.Value, 180000, null, Encoding.GetEncoding("UTF-8"));

                //return ClearHtml(pageSource);
                return pageSource;
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in getting target url. msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }
        }

        private HtmlDocument GetDataHtc(string p)
        {
            try
            {
                string second = string.Empty;
                string code = string.Empty;
                Regex reCode = new Regex(@"acptno\=(?<Code>\d+)\&docno");
                MatchCollection maCodes = reCode.Matches(p);

                if (maCodes.Count < 1)
                    return null;

                foreach (Match item in maCodes)
                {
                    code = item.Groups["Code"].Value;
                    break;
                }

                string source = WebClientUtil.GetDynamicPageSource(p, 300000, null);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();

                if (string.IsNullOrWhiteSpace(source))
                    return null;

                doc.LoadHtml(source);

                if (doc == null)
                    return null;

                second = doc.DocumentNode.SelectSingleNode(".//select[@id='mainDoc']/option[2]").Attributes["value"].Value.Trim().ToString();
                second = second.Trim().ToString().Replace("|Y", "");
                string url = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=searchContents&docNo={0}", second);
                string pageSource = WebClientUtil.GetPageSource(null, url, 180000, null, Encoding.GetEncoding("UTF-8"));
                Regex regex = new Regex(@"http://kind.krx.co.kr/.+htm");
                Match match = regex.Match(pageSource);

                if (!match.Success)
                    return null;

                HtmlDocument htc = new HtmlDocument();
                pageSource = WebClientUtil.GetPageSource(null, match.Value, 180000, null, Encoding.GetEncoding("UTF-8"));
                htc.LoadHtml(pageSource);

                return htc;
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in getting target url. msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }
        }

        public string ClearHtml(string str)
        {
            if (string.IsNullOrWhiteSpace(str))
                return null;

            str = Regex.Replace(str, @"<script[^>]*?>.*?</script>", "", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"-->", "", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"<!--.*", "", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&(quot|#34);", "\"", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&(nbsp|#160);", " ", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"&#(\d+);", "", RegexOptions.IgnoreCase);
            str = Regex.Replace(str, @"<img[^>]*>;", "", RegexOptions.IgnoreCase);
            str.Replace("<", "");
            str.Replace(">", "");
            str.Replace("\r\n", "");

            return str;
        }

        private string GetIssueSpecShares(HtmlDocument htc, string pageSource)
        {
            string result = null;

            if (htc == null)
                return result;

            try
            {
                HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");

                if (tables == null)//new requirement:new pages without tag <table/>
                {
                    List<string> listPattern = new List<string>();//주식수
                    //주식수(주)보통주KR7187790001A1877905,500,0003.액면가(원)1004.자본금(원)550,000,0005.
                    //주식수                            5,500,000
                    listPattern.Add(@"주식의종류와수\D*\d{0,4}\D*[^\d\,]+(?<money>[\d\,]{5,})");
                    foreach (var item in listPattern)
                    {
                        Regex reDate = new Regex(item);
                        MatchCollection maDate = reDate.Matches(pageSource);

                        if (maDate.Count < 1)
                            continue;

                        foreach (Match da in maDate)
                        {
                            result = da.Groups["money"].Value.Replace(",", "");
                        }
                    }
                }
                else
                {
                    HtmlNode table = tables[0];
                    HtmlNodeCollection trs = table.SelectNodes(".//tr");
                    string key = trs[4].SelectNodes(".//td")[4].InnerText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();
                    string value = trs[5].SelectNodes(".//td")[3].InnerText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();

                    if (key.Contains("주식수"))
                        result = value.Replace(",", "");
                }

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("get money(액면가) from web site error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return null;
                //throw new Exception(msg);
            }
        }


        private string GetNominalValue(string p)
        {
            string result = null;

            if (string.IsNullOrWhiteSpace(p))
                return result;

            string money = string.Empty;
            List<string> listPattern = new List<string>();//액면가
            //keyword: 1004. others
            //100
            listPattern.Add(@"(액면가|1주의발행가액)\D*(?<money>\d*0)\d{0,2}\D*");
            listPattern.Add(@"(액면가|1주의발행가액)\D*(?<money>[\d|\,]+0)");
            try
            {
                foreach (var item in listPattern)
                {
                    Regex reDate = new Regex(item);
                    MatchCollection maDate = reDate.Matches(p);

                    if (maDate.Count < 1)
                        continue;

                    foreach (Match da in maDate)
                    {
                        return da.Groups["money"].Value;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("get money(액면가) from web site error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return null;
                //throw new Exception(msg);
            }
        }

        private string GetListingDate(string p)
        {
            string result = null;

            if (string.IsNullOrWhiteSpace(p))
                return result;

            string year = string.Empty;
            string month = string.Empty;
            string day = string.Empty;
            List<string> listPattern = new List<string>();
            listPattern.Add(@"상장일\:*(?<year>\d{2,4})\D?(?<month>\d{1,2})\D?(?<day>\d{1,2})");
            listPattern.Add(@"상장일\(매매개시일\)(?<year>\d{2,4})년(?<month>\d{1,2})월(?<day>\d{1,2})일");
            try
            {
                foreach (var item in listPattern)
                {
                    Regex reDate = new Regex(item);
                    MatchCollection maDate = reDate.Matches(p);

                    if (maDate.Count < 1)
                        continue;

                    foreach (Match da in maDate)
                    {
                        year = da.Groups["year"].Value;
                        month = da.Groups["month"].Value;
                        day = da.Groups["day"].Value;
                        result = string.Format("{0}/{1}/{2}", day.Length == 1 ? "0" + day : day, month.Length == 1 ? "0" + month : month, year.Length == 2 ? "00" + year : year);
                        return result;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("get listing date from web site error. msg:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return null;
                //throw new Exception(msg);
            }
        }

        #endregion

        #region [Get Download Links]
        private void GetDownloadListFromDB(List<KORProcessItem> historicalAnnouncements)
        {
            List<KORProcessItem> list = new List<KORProcessItem>();
            List<int> listID = new List<int>();
            string keyWord1 = "%신규상장%";
            string keyWord2 = "%재상장%";

            if (extendContext == null)
            {
                string msg = string.Format("DataSourceString is null .");
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                var query1 = string.Format("SELECT * FROM source WHERE SourcingEngine = '{0}' AND ScopeType = 'InScope' AND MarketName = '{1}' AND AnnouceDate Between '{2}' AND '{3}' AND Title LIKE '{4}'", sourcingEngine, marketName, startTime, endTime, keyWord1);
                list.AddRange(extendContext.ExecuteQuery<KORProcessItem>(query1, 60).ToList());

                var query2 = string.Format("SELECT * FROM source WHERE SourcingEngine = '{0}' AND ScopeType = 'InScope' AND MarketName = '{1}' AND AnnouceDate Between '{2}' AND '{3}' AND Title LIKE '{4}'", sourcingEngine, marketName, startTime, endTime, keyWord2);
                list.AddRange(extendContext.ExecuteQuery<KORProcessItem>(query2, 60).ToList());

                foreach (var item in list)
                {
                    if (
                        item.Title.Contains("신주인수권증서 신규상장") ||
                        item.Title.Contains("주식선물 및 주식옵션 신규상장") ||
                        item.Title.Contains("ETF 신규상장 기준가격 안내") ||
                        item.Title.Contains("주권 신규상장 기준가격결정방법") ||
                        item.Title.Contains("재상장 기준가격결정방법") ||
                        item.Title.Contains("신주인수권증권 신규상장") ||
                        item.Title.Contains("주식선물 신규상장 안내") ||
                        item.Title.Contains("채권")
                        )
                        continue;

                    if (listID.Contains(item.ID))
                        continue;

                    historicalAnnouncements.Add(item);
                    listID.Add(item.ID);
                }

                Logger.Log(string.Format("{0} historical items loaded.", historicalAnnouncements.Count), Logger.LogType.Info);
            }
            catch (Exception ex)
            {
                string msg = string.Format("load historicalAnnouncements from Ace DB error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }
        #endregion
        #endregion
    }

    public class KoreaEquityCommon
    {
        /// <summary>
        /// Format Equity IDN Display Name.
        /// </summary>
        /// <param name="item">equity</param>
        public static void FormatEQIdnDisplayName(KoreaEquityInfo item)
        {
            if (item == null)
            {
                return;
            }
            if (string.IsNullOrEmpty(item.IDNDisplayName))
            {
                string legalName = ClearCoLtdForName(item.LegalName);
                item.IDNDisplayName = legalName;
            }
            else
            {
                item.IDNDisplayName = ClearCoLtdForName(item.IDNDisplayName);
            }

            if (item.IDNDisplayName.Length <= 16)
            {
                if (!KoreaEquityManager.ExistDisplayName(item.IDNDisplayName, item.RIC))
                {
                    return;
                }
            }
            else
            {
                if (item.Type.Equals("PRF"))
                {
                    FormatPrfEnding(item);
                }
                FormatKoreaIdnName(item);
            }
        }

        /// <summary>
        /// Format IDN display name if database contains a same name.
        /// </summary>
        /// <param name="item">equity</param>
        public static void FormatKoreaIdnName(KoreaEquityInfo item)
        {
            bool stop = false;
            int retry = 0;
            item.IDNDisplayName = GetIDNDisplayName(item.LegalName, item.PrfEnd, 16, 0, out stop);
            while (KoreaEquityManager.ExistDisplayName(item.IDNDisplayName, item.RIC) && !stop)
            {
                retry++;
                item.IDNDisplayName = GetIDNDisplayName(item.LegalName, item.PrfEnd, 16, retry, out stop);
            }
        }

        /// <summary>
        /// Get IDN Display Name.
        /// Keep the first word. SubString the other words.
        /// </summary>
        /// <param name="legalName">legal name</param>
        /// <param name="prfEnd">PRF end</param>
        /// <param name="targetLen">target length</param>
        /// <param name="retry">retry times</param>
        /// <param name="stop">if should stop</param>
        /// <returns>IDN Display Name</returns>
        public static string GetIDNDisplayName(string legalName, string prfEnd, int targetLen, int retry, out bool stop)
        {
            stop = false;
            string company = ClearCoLtdForName(legalName);

            if (string.IsNullOrEmpty(prfEnd))
            {
                prfEnd = "";
            }
            else
            {
                company = company.Replace(prfEnd, "").Trim(new char[] { ' ', '(', ')' });
                prfEnd = "(" + prfEnd + ")";
            }
            targetLen = targetLen - prfEnd.Length;

            List<string> subNames = company.Split(' ').ToList();
            string first = "";
            if (subNames.Count > 1)
            {
                first = subNames[0];
                subNames.RemoveAt(0);
            }
            int subLen = (targetLen - first.Length) / subNames.Count;
            int leftLen = (targetLen - first.Length) % subNames.Count;
            if (retry > subLen - 1 || subLen == 0)
            {
                first = company.Substring(0, targetLen);
                stop = true;
            }
            else
            {
                foreach (string name in subNames)
                {

                    if (name.Length > subLen)
                    {
                        if (leftLen > 0)
                        {
                            first += name.Substring(0, subLen + 1 - retry) + name.Substring(name.Length - retry);
                            leftLen--;
                        }
                        else
                        {
                            first += name.Substring(0, subLen - retry) + name.Substring(name.Length - retry);
                        }
                    }
                    else
                    {
                        first += name;
                    }
                }
            }
            first = (first + prfEnd).ToUpper();
            return first;
        }

        /// <summary>
        /// Get the PRF equity's ending. e.g. 1P 2P
        /// </summary>
        /// <param name="item">equity item</param>
        public static void FormatPrfEnding(KoreaEquityInfo item)
        {
            if (item == null || item.Type == null || (!item.Type.Equals("PRF")))
            {
                return;
            }
            string legalName = item.LegalName.Replace("(", "").Replace(")", "");
            Regex regex = new Regex(@"[0-9]+P");
            Match match = regex.Match(legalName);
            if (match.Success)
            {
                item.PrfEnd = legalName.Substring(match.Index);
            }
        }

        /// <summary>
        /// Remove the infos of company like CO LTD CORP INC CORPARATION
        /// </summary>
        /// <param name="underEngName">full name</param>
        /// <returns>name without company infos</returns>
        public static string ClearCoLtdForName(string underEngName)
        {
            underEngName = underEngName.ToUpper();
            List<string> names = underEngName.Split(new char[] { ' ', ',', '.' }).ToList();
            string result = "";
            names.Remove("CO");
            names.Remove("LTD");
            names.Remove("INC");
            names.Remove("CORP");
            names.Remove("COMPANY");
            names.Remove("LIMITED");
            names.Remove("CORPORATION");
            foreach (string name in names)
            {
                if (name == "" || name == " ")
                {
                    continue;
                }
                result += name + " ";
            }
            return result.TrimEnd();
        }
    }
}
