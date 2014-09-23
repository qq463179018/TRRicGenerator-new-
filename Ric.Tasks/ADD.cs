using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Util;
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
using Ric.Core;

namespace Ric.Tasks
{
    public class ADD : GeneratorBase
    {
        private List<KoreaEquityInfo> paList = new List<KoreaEquityInfo>();
        private List<KoreaEquityInfo> raList = new List<KoreaEquityInfo>();
        private List<KoreaEquityInfo> eaList = new List<KoreaEquityInfo>();
        private List<KoreaEquityInfo> bcList = new List<KoreaEquityInfo>();
        private Hashtable equityDate = new Hashtable();

        private Hashtable taskResult = new Hashtable();
        private KOREA_ADDGeneratorConfig configObj;
        private List<string> errors = new List<string>();
        private bool peoChanged;

        protected override void Start()
        {
            StartADDJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREA_ADDGeneratorConfig;
        }

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
            string dataStartDate;
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
                        HtmlDocument doc;
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
                                }
                            }
                        }
                        else if (title.Replace(" ", "").Contains("상장지수투자신탁"))
                        {
                            doc = GetTargetPageDocument(node);
                            bool isGlobal = JudgeIfGloblaETF(nodeDate.InnerText.Trim(), title);
                            GrabETFDataAction(doc, isGlobal);
                        }
                        else if (title.Replace(" ", "").Contains("부동산투자회사"))
                        {
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
                judgeText = judgeText.Split('(')[1].Split(')')[0].Trim(new[] { '\'', ' ' });
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
                    attribute = attribute.Split('(')[1].Split(')')[0].Trim(new[] { ' ', ',', '\'' });
                string param = attribute.Split(',')[1].Trim(new[] { ' ', '\'', ',' });
                string url = GetTargetUrl(param);
                doc = WebClientUtil.GetHtmlDocument(url, 600000);
            }
            catch (Exception ex)
            {
                string msg = "" + ex.StackTrace + "  : ----> \n\r" + ex;
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
            result = result.Trim(new[] { ' ', ':' });
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
            if (dutyName.Contains("유가증권"))
            {
                return "KS";
            }
            if (dutyName.Contains("코넥스"))
            {
                return "KN";
            }
            return "";
        }

        /// <summary>
        /// Grab PEO add KS data(data in text).
        /// </summary>
        /// <param name="doc">page document</param>
        /// <param name="company"></param>
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
                int pricePos;
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
                    KoreaEquityInfo peo = new KoreaEquityInfo
                    {
                        ISIN = isinItem, 
                        Ticker = tickers[i]
                    };
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
                        idnDisplayName = match.Value.Substring(match.Value.IndexOf("영문") + 2).TrimStart(new[] { ' ', ')', ':', '(' });
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
                    Thread.Sleep(5000);
                }
            }

            if (pageSource == null)
            {
                string msg = "At GetPeoTypeByISIN(string isin). Cannot find peo type infos for " + isin;
                errors.Add(msg);
                return null;
            }

            string peoType = string.Empty;
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(pageSource);
            HtmlNodeCollection trs;
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
                KoreaEquityInfo peoTemp = new KoreaEquityInfo
                {
                    Ticker = strTicker.Trim().Substring(1),
                    KoreaName =
                        strKorreaname.Trim().Contains("(주)")
                            ? strKorreaname.Trim().Replace("(주)", "")
                            : strKorreaname.Trim(),
                    LegalName = strLegalname.Trim(),
                    ISIN = strIsin.Trim(),
                    EffectiveDate = strEffectiveDate,
                    Market = dutyCode,
                    IDNDisplayName = strIdnDisplayName.ToUpper()
                };

                peoTemp.Type = GetPeoTypeByISIN(peoTemp.ISIN);
                peoTemp.RIC = peoTemp.Ticker + "." + dutyCode;
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
                    string isin = isinRic.Split('(')[0].Trim();
                    string ticker = isinRic.Split(':')[1].Trim(new[] { ' ', ':', ')' }).Substring(1);

                    int koreanNamePos = document.IndexOf("한글종목명") + ("한글종목명".Length);
                    int legalNamePos = document.IndexOf("영문종목명") + ("영문종목명".Length);
                    koreanName = FormatDataWithPos(koreanNamePos, document);
                    legalName = FormatDataWithPos(legalNamePos, document);


                    effectiveDate = FormatDataWithPos(effectiveDatePos, document);
                    FormatDate(ref effectiveDate);
                    string[] names = GetETFDetailByIsin(isin);

                    KoreaEquityInfo etf = new KoreaEquityInfo
                    {
                        ISIN = isin,
                        Ticker = ticker,
                        RIC = ticker + ".KS",
                        EffectiveDate = effectiveDate,
                        KoreaName = koreanName,
                        LegalName = legalName,
                        Market = "KS",
                        Type = "ETF",
                        IsGlobalETF = isGlobal,
                        IDNDisplayName = KoreaEquityCommon.ClearCoLtdForName(legalName)
                    };
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
            return isGlobal.Equals("Y");
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

                string koreanName = koreanNameTemp.Trim(new[] { ' ', ':', ')' });
                koreanName = koreanName.Contains("(약명") ? koreanName.Substring(0, (koreanName.IndexOf("(약명"))).Trim() : koreanName;
                string legalName = legalNameTemp.Trim(new[] { ' ', ':', ')' });
                legalName = legalNameTemp.Contains("(약명") ? legalName.Substring(0, (legalName.IndexOf("(약명"))) : legalName;
                string effectiveDate = effectiveTemp.Trim(new[] { ' ', ':', ')', '목', '(' }).ToString();
                FormatDate(ref effectiveDate);
                string isin = isinRicTemp.Split('(')[0].Trim(new[] { ' ', ':', '(' });
                string ticker = Regex.Split(isinRicTemp, "단축코드", RegexOptions.IgnoreCase)[1].Trim(new[] { ' ', ':', '(', ')' });
                string strikePrice = strikePriceTemp.Trim(new[] { ' ', ':' });
                strikePrice = strikePrice.Contains('원') ? strikePrice.Substring(0, (strikePrice.IndexOf('원'))) : strikePrice;
                strikePrice = strikePrice.Contains(',') ? strikePrice.Replace(",", "") : strikePrice;

                KoreaEquityInfo ra = new KoreaEquityInfo
                {
                    ISIN = isin,
                    Ticker = ticker.Substring(1).Trim(),
                    EffectiveDate = effectiveDate,
                    KoreaName = koreanName,
                    LegalName = legalName,
                    Lotsize = strikePrice,
                    IDNDisplayName = legalName.Substring(0, 16).ToUpper(),
                    Market = "KS",
                    Type = "REIT"
                };
                ra.RIC = ra.Ticker + ".KS";
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

            KoreaEquityInfo equity = new KoreaEquityInfo
            {
                Status = "Active",
                RIC = ticker + ".KS",
                ISIN = isin,
                Ticker = ticker,
                Lotsize = "10",
                UpdateDate = DateTime.Today.ToString("yyyy-MMM-dd", new CultureInfo("en-US")),
                Type = "CEF",
                RecordType = "112",
                FM = "2",
                Category = "CEF",
                KoreaName = koreaFullName,
                LegalName = legalName,
                EffectiveDate = effectiveDate,
                IDNDisplayName = idnDisplayName
            };
            equity.BcastRef = equity.RIC;
            
            if (equity.KoreaName.Length > 14)
            {
                equity.KoreaName = equity.KoreaName.Substring(0, 14);
            }
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
            string dataStartDate;
            DateTime startDay = DateTime.Parse(startDate);
            DateTime twoMonthEarly = DateTime.Parse(endDate).AddMonths(-2);
            dataStartDate = startDay.CompareTo(twoMonthEarly) < 0 ? startDate : twoMonthEarly.ToString("yyyy-MM-dd");
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
                            HtmlDocument doc;
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
                    List<string> fmTitle = new List<string>
                    {"Updated Date", "Effective Date", "RIC", "Type", "Record Type", "FM", "IDN Display Name", 
                                                               "ISIN", "Ticker", "BCAST REF", "Legal Name", "Korea Name", "Lotsize" };

                    ((Range)wSheet.Columns["A", Type.Missing]).ColumnWidth = 13;
                    ((Range)wSheet.Columns["B", Type.Missing]).ColumnWidth = 13;
                    ((Range)wSheet.Columns["G", Type.Missing]).ColumnWidth = 20;
                    ((Range)wSheet.Columns["H", Type.Missing]).ColumnWidth = 16;
                    ((Range)wSheet.Columns["K", Type.Missing]).ColumnWidth = 20;
                    ((Range)wSheet.Columns["L", Type.Missing]).ColumnWidth = 20;

                    ((Range)wSheet.Cells[3, 1]).Font.Underline = FontStyle.Underline;
                    ((Range)wSheet.Cells[3, 1]).Font.Bold = FontStyle.Bold;
                    wSheet.Cells[3, 1] = "EQUITY ADD";
                    for (int j = 1; j <= fmTitle.Count; j++)
                    {
                        wSheet.Cells[4, j] = fmTitle[j - 1];
                    }
                    ((Range)wSheet.Rows[4, Type.Missing]).Font.Bold = FontStyle.Bold;
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
                    ((Range)wSheet.Cells[8, 1]).Font.Bold = FontStyle.Bold;

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

                    TaskResultList.Add(new TaskResultEntry(fileName, item.Type + " FM File", filePath, mail));
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

                List<string> fmTitle = new List<string>
                {"Updated Date", "Effective Date", "RIC", "Type", "Record Type", "FM", "IDN Display Name", 
                                                               "ISIN", "Ticker", "BCAST REF", "Legal Name", "Korea Name", "Lotsize" };

                ((Range)wSheet.Columns["A", Type.Missing]).ColumnWidth = 13;
                ((Range)wSheet.Columns["B", Type.Missing]).ColumnWidth = 13;
                ((Range)wSheet.Columns["G", Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["H", Type.Missing]).ColumnWidth = 16;
                ((Range)wSheet.Columns["K", Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["L", Type.Missing]).ColumnWidth = 20;

                ((Range)wSheet.Cells[3, 1]).Font.Underline = FontStyle.Underline;
                ((Range)wSheet.Cells[3, 1]).Font.Bold = FontStyle.Bold;
                wSheet.Cells[3, 1] = "EQUITY ADD";
                for (int j = 1; j <= fmTitle.Count; j++)
                {
                    wSheet.Cells[4, j] = fmTitle[j - 1];
                }
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Bold = FontStyle.Bold;

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
                        ((Range)wSheet.Cells[5 + k, col]).Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                    }
                }

                wSheet.Cells[7 + equity.Count, 1] = "- End -";
                ((Range)wSheet.Cells[7 + equity.Count, 1]).Font.Bold = FontStyle.Bold;

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

                TaskResultList.Add(new TaskResultEntry(fileName, mailType + " ADD FM File", filePath, mail));
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

        private string CombineAllRics(IEnumerable<KoreaEquityInfo> items)
        {
            return string.Join(",", items.Select(e => e.RIC).ToArray());
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
            List<string> title = new List<string> { "RIC" };
            List<string> data = new List<string> { item.RIC };

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
                TaskResultList.Add(new TaskResultEntry(fileName, "GEDA Change File (PEO)", filePath, FileProcessType.GEDA_BULK_RIC_CHANGE));
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
        /// Generate BC GEDA file for an bc equity.
        /// </summary>
        /// <param name="item">bc equity</param>
        private void GenerateGedaBcFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>
            { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>
            {
                item.RIC,
                item.IDNDisplayName,
                item.RIC,
                item.Ticker,
                item.ISIN,
                "****",
                item.KoreaName,
                item.RIC,
                item.ISIN,
                "F" + item.Ticker,
                "115",
                item.LegalName.ToUpper(),
                item.Ticker,
                "KSE_EQB_BC"
            };
            string fileName = "KR_BC_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + ".txt";
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
            List<string> title = new List<string>
            { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_RECORDTYPE", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>
            {
                item.RIC,
                item.IDNDisplayName,
                item.RIC,
                item.Ticker,
                item.ISIN,
                "****",
                item.KoreaName,
                item.RIC,
                item.ISIN,
                "A" + item.Ticker,
                "112",
                "115",
                item.LegalName.ToUpper(),
                item.Ticker,
                "KSE_EQB_3"
            };

            string fileName = "KR_REIT_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + ".txt";
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
            List<string> title = new List<string>
            { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>
            {
                item.RIC,
                item.IDNDisplayName,
                item.RIC,
                item.Ticker,
                item.ISIN,
                "****",
                item.KoreaName,
                item.RIC,
                item.ISIN,
                "A" + item.Ticker,
                "115",
                item.LegalName.ToUpper(),
                item.Ticker,
                "KSE_EQB_ETF"
            };

            string fileName = "KR_ETF_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + ".txt";
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
            List<string> title = new List<string>
            { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_PROD_PERM", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>
            {
                item.RIC,
                item.IDNDisplayName,
                item.RIC,
                item.Ticker,
                item.ISIN,
                "****",
                item.KoreaName,
                item.RIC,
                item.ISIN,
                "A" + item.Ticker,
                "115",
                item.LegalName.ToUpper(),
                "6688",
                item.Ticker,
                "KSE_EQB_ETF"
            };
            string fileName = "KR_ETF_GLOBAL_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + ".txt";
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
            List<string> title = new List<string>
            { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>
            {
                item.RIC,
                item.IDNDisplayName,
                item.RIC,
                item.Ticker,
                item.ISIN,
                item.KoreaName,
                item.RIC,
                item.ISIN,
                "A" + item.Ticker,
                item.LegalName.ToUpper(),
                item.Ticker
            };
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
            string fileName = "KR_KDR_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + ".txt";
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
            List<string> title = new List<string>
            { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ASSET_SUB_TYPE", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>
            {
                item.RIC,
                item.IDNDisplayName,
                item.RIC,
                item.Ticker,
                item.ISIN,
                "****",
                item.KoreaName,
                item.RIC,
                item.ISIN,
                "A" + item.Ticker,
                "82",
                item.LegalName.ToUpper(),
                item.Ticker
            };
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
            string fileName = "KR_PRF_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + ".txt";
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
            List<string> title = new List<string>
            { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL","BCAST_REF", 
                                                         "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_TDN_ISSUER_NAME", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            List<string> data = new List<string>
            {
                item.RIC,
                item.IDNDisplayName,
                item.RIC,
                item.Ticker,
                item.ISIN,
                "****",
                item.KoreaName,
                item.RIC
            };
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
                    data.Add("KSE_EQB_3");
                }
                else if (item.Market.Equals("KQ"))
                {
                    data.Add("KOSDAQ_EQB_2");
                }
            }

            string fileName = "KR_ORD_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + ".txt";
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
                }
            }
            foreach (KoreaEquityInfo item in eaList)
            {
                GenerateNdaEtfFile(item);
                GenerateNdaBGFile(item);
            }
            foreach (KoreaEquityInfo item in raList)
            {
                GenerateNdaReitFile(item);
            }
            foreach (KoreaEquityInfo item in bcList)
            {
                GenerateNdaBcFile(item);
            }
            Logger.Log("Successfully generated NDA files.");
        }

        /// <summary>
        /// Generate NDA QA file for a PEO equity with FM1.
        /// </summary>
        /// <param name="item">a PEO equity</param>
        private void GenerateNdaPeoFM1File(KoreaEquityInfo item)
        {
            List<string> title = new List<string> { "RIC", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD" };
            string[] rics = { ".", "F.", "S.", "stat.", "ta.", "bl." };
            string[] lotSizes = { item.Lotsize, "", "1", "", "", "" };
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
                    List<string> content = new List<string>
                    {
                        item.Ticker + rics[i] + item.Market,
                        effectiveDate,
                        lotSizes[i],
                        "T+2",
                        item.IDNDisplayName,
                        item.IDNDisplayName + " " + item.Type
                    };
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
                    List<string> content = new List<string>
                    {
                        item.Ticker + rics[i] + item.Market,
                        effectiveDate,
                        lotSizes[i],
                        "T+2"
                    };
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
        private void GenerateNdaPeoFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>
            { "RIC", "TAG", "BASE ASSET","TICKER SYMBOL", "ASSET SHORT NAME", "ASSET COMMON NAME", "TYPE",	
                                                      "CATEGORY","CURRENCY", "EXCHANGE", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD"};
            string[] rics = { ".", "F.", "S.", "stat.", "ta.", "bl." };
            string[] tags;
            if (item.Market.Equals("KQ"))
            {
                tags = new[] { "673", "64399", "60673", "61287", "64380", "67094" };
            }
            else if (item.Market.Equals("KS"))
            {
                tags = new[] { "184", "64398", "60184", "61286", "64379", "67093" };
            }
            else if (item.Market.Equals("KN"))
            {
                tags = new[] { "67105", "67107", "*****", "67108", "64380", "67106" };
            }
            else
            {
                string msg = "At GenerateNdaPeoFile(). Can not get the exchange board code.";
                errors.Add(msg);
                return;
            }

            string[] lotSizes = { item.Lotsize, "", "1", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            List<List<string>> data = new List<List<string>>();
            for (int i = 0; i < 6; i++)
            {
                if (i == 2 && item.Market.Equals("KN"))
                {
                    continue;
                }
                List<string> content = new List<string>
                {
                    item.Ticker + rics[i] + item.Market,
                    tags[i],
                    "ISIN:" + item.ISIN,
                    item.Ticker,
                    item.IDNDisplayName,
                    item.IDNDisplayName + " " + item.Type,
                    "EQUITY",
                    item.Category,
                    "KRW",
                    item.Exchange,
                    effectiveDate,
                    lotSizes[i],
                    "T+2"
                };
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
            List<string> title = new List<string>
            { "RIC", "TAG", "BASE ASSET","TICKER SYMBOL", "ASSET SHORT NAME", "ASSET COMMON NAME", "TYPE",	
                                                      "CATEGORY","CURRENCY", "EXCHANGE", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD"};
            List<List<string>> data = new List<List<string>>();
            string[] rics = { ".KS", "F.KS", "S.KS", "stat.KS", "ta.KS", "bl.KS", "LP.KS" };
            string[] tags = { "184", "64398", "60184", "61286", "64379", "67093", "64395" };
            string[] lotSizes = { item.Lotsize, "", "1", "", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            for (int i = 0; i < 7; i++)
            {
                List<string> content = new List<string>
                {
                    item.Ticker + rics[i],
                    tags[i],
                    "ISIN:" + item.ISIN,
                    item.Ticker,
                    item.IDNDisplayName,
                    item.IDNDisplayName + " ETF",
                    "EQUITY",
                    item.Category,
                    "KRW",
                    "KSC",
                    effectiveDate,
                    lotSizes[i],
                    "T+2"
                };
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

        /// <summary>
        /// Generate NDA BG file for a ETF equity.
        /// </summary>
        /// <param name="item">a ETF equity</param>
        private void GenerateNdaBGFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string> { "XXLEGAL NAMEXX", "XXBODY GROUP COMMON NAMEXX", "XXBODY GROUP SHORT NAMEXX", "XXORGANISATION TYPEXX", "XXCOUNTRY HEADQUARTERSXX" };
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
            List<string> title = new List<string>
            { "RIC", "TAG", "BASE ASSET","TICKER SYMBOL", "ASSET SHORT NAME", "ASSET COMMON NAME", "TYPE",	
                                                      "CATEGORY","CURRENCY", "EXCHANGE", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD"};
            List<List<string>> data = new List<List<string>>();
            string[] rics = { ".KS", "F.KS", "S.KS", "stat.KS", "ta.KS", "bl.KS" };
            string[] tags = { "184", "64398", "60184", "61286", "64379", "67093" };
            string[] lotSizes = { item.Lotsize, "", "1", "", "", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            for (int i = 0; i < 6; i++)
            {
                List<string> content = new List<string>
                {
                    item.Ticker + rics[i],
                    tags[i],
                    "ISIN:" + item.ISIN,
                    item.Ticker,
                    item.IDNDisplayName,
                    item.IDNDisplayName + " REIT",
                    "EQUITY",
                    "REI",
                    "KRW",
                    "KSC",
                    effectiveDate,
                    lotSizes[i],
                    "T+2"
                };
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

        /// <summary>
        /// Generate NDA QA file for a BC equity.
        /// </summary>
        /// <param name="item">a BC equity</param>
        private void GenerateNdaBcFile(KoreaEquityInfo item)
        {
            List<string> title = new List<string>
            { "RIC", "TAG", "BASE ASSET","TICKER SYMBOL", "ASSET SHORT NAME", "ASSET COMMON NAME", "TYPE",	
                                                      "CATEGORY","CURRENCY", "EXCHANGE", "EQUITY FIRST TRADING DAY", "ROUND LOT SIZE", "SETTLEMENT PERIOD"};
            List<List<string>> data = new List<List<string>>();
            string[] rics = { ".KS", "F.KS" };
            string[] tags = { "184", "64398" };
            string[] lotSizes = { "10", "" };
            string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));

            for (int i = 0; i < 2; i++)
            {
                List<string> content = new List<string>
                {
                    item.Ticker + rics[i],
                    tags[i],
                    "ISIN:" + item.ISIN,
                    item.Ticker,
                    item.IDNDisplayName,
                    item.IDNDisplayName + " CEF",
                    "EQUITY",
                    "CEF",
                    "KRW",
                    "KSC",
                    effectiveDate,
                    lotSizes[i],
                    "T+2"
                };
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
                List<string> tickTitle = new List<string>
                {"RIC", "TICK NOT APPLICABLE", "TICK LADDER NAME", 
                                                      "TICK EFFECTIVE FROM", "TICK EFFECTIVE TO", "TICK PRICE INDICATOR" };
                List<string> lotTitle = new List<string>
                {"RIC", "LOT NOT APPLICABLE", "LOT LADDER NAME", 
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
                string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
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
            TaskResultList.Add(new TaskResultEntry("Log", "LOG FILE", Logger.FilePath));
        }

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
                company = company.Replace(prfEnd, "").Trim(new[] { ' ', '(', ')' });
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
            List<string> names = underEngName.Split(new[] { ' ', ',', '.' }).ToList();
            names.Remove("CO");
            names.Remove("LTD");
            names.Remove("INC");
            names.Remove("CORP");
            names.Remove("COMPANY");
            names.Remove("LIMITED");
            names.Remove("CORPORATION");
            string result = names.Where(name => name != "" && name != " ")
                                 .Aggregate("", (current, name) => current + (name + " "));
            return result.TrimEnd();
        }
    }
}
