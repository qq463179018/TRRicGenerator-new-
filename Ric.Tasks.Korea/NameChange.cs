using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Selenium;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
{

    public class NameChange : GeneratorBase
    {
        private List<DateTime> holidayList = null;        
        private KOREA_NameChangeGeneratorConfig configObj = null;
        private List<string> errors = new List<string>();
        private List<KoreaEquityInfo> nameChanges = new List<KoreaEquityInfo>();
        private List<KoreaEquityInfo> nameRevised = new List<KoreaEquityInfo>();
        private Hashtable equityDate = new Hashtable();
        private Hashtable filesResultList = new Hashtable();

        protected override void Start()
        {
            StartNameChangeJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREA_NameChangeGeneratorConfig;
            holidayList = HolidayManager.SelectHoliday(MarketId);//ConfigUtil.ReadConfig(HOLIDAY_LIST_FILE_PATH, typeof(List<DateTime>)) as List<DateTime>;
        }

        public void StartNameChangeJob()
        {
            try
            {
                //Sourcing 
                GrabDataFromWebpage();

                //Formating
                FormatNameChange();
 
                //Distributing
                GenerateFMFiles();
                if (nameChanges.Count > 0)
                {
                    GenerateGEDANDAFiles();
                    CreateTaskResult();
                }
            }
            catch (Exception ex)
            {
                string msg = "At StartNameChangeJob()." + ex.Message + "\r\n" + ex.StackTrace;
                errors.Add(msg);
            }
            finally
            {
                LogErrors();
            }
        }

        #region Sourcing

        private void GrabDataFromWebpage()
        {
            string startDate = configObj.StartDate;
            string endDate = configObj.EndDate;
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
            string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%83%81%ED%98%B8%EB%B3%80%EA%B2%BD&repIsuSrtCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=100&scn=mktact&srchFd=2&kwd=%EC%83%81%ED%98%B8%EB%B3%80%EA%B2%BD&fromData={0}&toData={1}", dataStartDate, endDate);
            string uri = "http://kind.krx.co.kr/disclosure/searchtotalinfo.do";
            string pageSource = string.Empty;

            pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, postData);
            HtmlDocument htc = new HtmlDocument();
            if (!string.IsNullOrEmpty(pageSource))
            {
                htc.LoadHtml(pageSource);
            }
            if (htc != null)
            {
                HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes(".//dl/dt");
                HtmlNodeCollection ddCollections = htc.DocumentNode.SelectNodes("//dl/dd");
                int count = nodeCollections.Count;
                for (var i = 0; i < count; i++)
                {

                    try
                    {
                        HtmlNode ddNode = ddCollections[i].SelectSingleNode(".//span");
                        HtmlNode node = nodeCollections[i];
                        HtmlNode dtNode = node.SelectSingleNode(".//span/a");
                        HtmlNode nodeDate = node.SelectSingleNode("./em");
                        if (nodeDate != null)
                        {
                            DateTime anouncementDate = new DateTime();
                            anouncementDate = DateTime.Parse(nodeDate.InnerText.Trim(), new CultureInfo("en-US"));
                            if (anouncementDate < startDay)
                            {
                                return;
                            }
                        }

                        string title = string.Empty;
                        if (node != null)
                            title = node.SelectSingleNode(".//span/a").InnerText.Trim().ToString();
                        if (!string.IsNullOrEmpty(title))
                        {
                            if (title.Contains("상장채권"))
                            {
                                continue;
                            }

                            if (title.Contains("영문명") || title.Contains("상호변경"))
                            {
                                HtmlNode companyNode = node.SelectSingleNode(".//strong/a");
                                KoreaEquityInfo item = GrabIsinInfo(companyNode);
                                if (item == null)
                                {
                                    string msg = string.Format("Can not get company information for announcement at {0}.", nodeDate.InnerText.Trim());
                                    errors.Add(msg);
                                    continue;
                                }
                                item.AnnouncementTime = nodeDate.InnerText.Trim();
                                string ticker = GetTickerByAttribute(dtNode);
                                // to record the 본문's value in announcement detail page
                                // e.g.  http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno=20130410000135&docno=20130410000330&viewerhost=&viewerport=

                                if (string.IsNullOrEmpty(ticker))
                                {
                                    string msg = string.Format("Can not get ticker for announcement at {0}.", nodeDate.InnerText.Trim());
                                    errors.Add(msg);
                                    continue;
                                }
                                item.Ticker = ticker;
                                if (string.IsNullOrEmpty(item.Market))
                                {
                                    item.Market = GetDutyCode(ddNode);
                                    if (string.IsNullOrEmpty(item.Market))
                                    {
                                        string msg = string.Format("Can not get market information for announcement at {0}.", nodeDate.InnerText.Trim());
                                        errors.Add(msg);
                                        continue;
                                    }
                                }
                                item.RIC = ticker + "." + item.Market;

                                HtmlDocument sourceNode = GetTargetPageDocument(dtNode);
                                if (sourceNode == null)
                                {
                                    string msg = string.Format("Can not get detail information for announcement at {0}.", item.AnnouncementTime);
                                    errors.Add(msg);
                                    continue;
                                }

                                GetNameChangeInfo(sourceNode, title, item);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string msg = "Error found in GrabDataFromWebpage(). " + ex.Message + "\r\n"+ ex.ToString();
                        Logger.Log(msg, Logger.LogType.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Grab ISIN infos from the first link node in anouncement.
        /// </summary>
        /// <param name="judgeNode">company link node</param>
        /// <returns>korea equity info</returns>
        private KoreaEquityInfo GrabIsinInfo(HtmlNode judgeNode)
        {
            try
            {
                string judgeText = judgeNode.Attributes["onclick"].Value.Trim().ToString();
                judgeText = judgeText.Split('(')[1].Split(')')[0].Trim(new char[] { '\'', ' ' }).ToString();
                KoreaCompany company = Common.GetCompanyInfo(judgeText);
                if (company == null)
                {
                    return null;
                }
                KoreaEquityInfo item = new KoreaEquityInfo();
                item.ISIN = company.ISIN;
                item.Market = company.Market;
                item.OldKoreaName = company.KoreaName.Replace("(주)", "");
                item.OldLegalName = company.LegalName;
                return item;
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabIsinInfo(). Error message:{0}", ex.Message);
                errors.Add(msg);
                return null;
            }
        }       

        /// <summary>
        /// Get ticker in target document. E.g. 6 digits in (123456)
        /// </summary>
        /// <param name="node">html node</param>
        /// <returns>ticker</returns>
        private string GetTickerByAttribute(HtmlNode node)
        {
            string ticker = string.Empty;
            try
            {
                string attribute = node.Attributes["onclick"].Value.Trim().ToString();
                attribute = attribute.Split('(')[1].Split(')')[0].Trim(new char[] { ' ', '\'', '(', ')' }).ToString();
                string param1 = attribute.Split(',')[0].Trim(new char[] { ' ', ',', '\'' }).ToString();
                string param2 = attribute.Split(',')[1].Trim(new char[] { ' ', ',', '\'' }).ToString();
                string url = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno={1}&viewerhost=&viewerport=", param1, param2);
                HtmlDocument doc = WebClientUtil.GetHtmlDocument(url, 300000, null);
                ticker = doc.DocumentNode.SelectSingleNode(".//header/h1").InnerText;
               // string category = string.Empty;
               // HtmlNode categoryNode = doc.DocumentNode.SelectSingleNode(".//div[@id='pWrapper']");//div[@id='pContArea']/form[@name='frm']/div[@id='vHeader']/div[1]/div[1]/select[1]/option[2]");

                if (!string.IsNullOrEmpty(ticker))
                {
                    Match m = Regex.Match(ticker, @"\(([0-9a-zA-Z]+)\)");
                    if (m == null)
                    {
                        string msg = "At GetTickerByAttribute(). Cannot get ticker numbers in ." + ticker;
                        errors.Add(msg);
                        return null;
                    }
                    ticker = m.Groups[1].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GetTickerByAttribute(). Cannot get ticker." + ex.Message);
                errors.Add(msg);
            }
            return ticker;
        }

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
        /// Grab anouncement detail infomation with different market.
        /// </summary>
        /// <param name="sourceNode">anouncement page source node</param>
        /// <param name="title">anouncement title</param>
        /// <param name="item">equity item</param>
        private void GetNameChangeInfo(HtmlDocument sourceNode, string title, KoreaEquityInfo item)
        {
            if (title.Contains("상호변경"))
            {
                //KQ and KN's source page are similar.
                if (item.Market.Equals("KQ"))
                {
                    GrabCommonKQDataChangeAction(item, sourceNode);
                }
                else if (item.Market.Equals("KN"))
                {
                    Logger.Log("KN anouncement name change occured.", Logger.LogType.Warning);
                    GrabCommonKQDataChangeAction(item, sourceNode);
                }
                else if (item.Market.Equals("KS"))
                {
                    if (title.Equals("종목명 변경 영문상호변경"))
                    {
                        GrabEnglishCompanyChangeData(item, sourceNode);
                    }
                    else
                    {
                        GrabCommonKSData(sourceNode, item);
                    }
                }
            }

            else if (title.Contains("영문명"))
            {
                //Until now, we only find KQ page. If KS or KN occured, the code need to rewrite.
                if (item.Market.Equals("KQ"))
                {
                    GrabLegalNameDataChangeAction(item, sourceNode);
                }
                else if (item.Market.Equals("KN"))
                {
                    Logger.Log("KN anouncement for English Name change occured.", Logger.LogType.Warning);
                    GrabLegalNameDataChangeAction(item, sourceNode);
                }
                else if (item.Market.Equals("KS"))
                {
                    Logger.Log("KS anouncement for English Name change occured.", Logger.LogType.Warning);
                }
            }

            if (title.Contains("[정정]"))
            {
                nameRevised.Add(item);
            }
            else
            {
                if (item.Market.Equals("KS") && title.Contains("상호변경"))
                {
                    return;
                }
                nameChanges.Add(item);
            }
        }

        /// <summary>
        /// Grab KQ data. title contains "상호변경"
        /// </summary>
        /// <param name="item">equity item</param>
        /// <param name="sourceNode">source node</param>
        private void GrabCommonKQDataChangeAction(KoreaEquityInfo item, HtmlDocument sourceNode)
        {
            try
            {
                HtmlNodeCollection tables = sourceNode.DocumentNode.SelectNodes(".//table");
                HtmlNode table = null;
                foreach (var node in tables)
                {
                    HtmlNode knameNode = node.SelectSingleNode(".//tr[3]/td[3]");
                    if (knameNode != null)
                        if (knameNode.InnerText.Trim().ToString().Equals("한글"))
                            table = node;
                }

                string koreaName = table.SelectSingleNode(".//tr[3]/td[4]").InnerText.Trim();
                string legalName = table.SelectSingleNode(".//tr[4]/td[2]").InnerText.Trim();
                string effectiveDate = table.SelectSingleNode(".//tr[7]/td[2]").InnerText.Trim();
                string displayName = table.SelectSingleNode(".//tr[6]/td[2]").InnerText.Trim();

                koreaName = koreaName.Contains("(주)") ? koreaName.Replace("(주)", "") : koreaName;
                effectiveDate = Convert.ToDateTime(effectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));

                if (!string.IsNullOrEmpty(koreaName))
                {
                    item.KoreaName = koreaName;
                }
                if (!string.IsNullOrEmpty(legalName))
                {
                    item.LegalName = legalName;
                }
                if (!string.IsNullOrEmpty(effectiveDate))
                {
                    item.EffectiveDate = effectiveDate;
                }
                if (!string.IsNullOrEmpty(displayName))
                {
                    item.IDNDisplayName = displayName.ToUpper();
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabCommonKQDataChangeAction     : \r\n" + ex.ToString();
                errors.Add(msg);
            }
        }

        /// <summary>
        /// An unique announcement occured on 2013.04.10. For S-1 Corporation
        /// http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno=20130410000135&docno=20130410000330&viewerhost=&viewerport=
        /// </summary>
        /// <param name="item">name change equity</param>
        /// <param name="sourceNode">page source node</param>
        private void GrabEnglishCompanyChangeData(KoreaEquityInfo item, HtmlDocument sourceNode)
        {
            try
            {
                string strPre = sourceNode.DocumentNode.SelectSingleNode(".//pre").InnerText.Trim().ToString();

                int englishNamePos = strPre.IndexOf("영문종목명") + "영문종목명".Length;
                int usNamePos = strPre.IndexOf("영문종목약명") + "영문종목약명".Length;
                int effectivePos = strPre.IndexOf("변경일자") + "변경일자".Length;

                string englishNameLine = FormatDataWithPos(englishNamePos, strPre);
                string usNameLine = FormatDataWithPos(usNamePos, strPre);
                string effectiveDateLine = FormatDataWithPos(effectivePos, strPre);

                if (englishNameLine.Contains("→"))
                {
                    item.LegalName = englishNameLine.Split('→')[1].Trim();
                }
                if (usNameLine.Contains("→"))
                {
                    item.IDNDisplayName = usNameLine.Split('→')[1].Trim().ToUpper();
                }
                item.KoreaName = item.OldKoreaName;
                item.EffectiveDate = DateTime.Parse(effectiveDateLine.Trim()).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                nameChanges.Add(item);
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabEnglishCompanyChangeData(). Can not grab detail announcement for {0}. Error message:{1}", item.RIC, ex.Message);
                errors.Add(msg);
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
            result = result.Trim(new char[] { ' ', '\n', '\r', '(', ')', '▶', ':' });
            return result;
        }

        private string FormatPrfNameWithPos(int pos, string strPre)
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
            result = result.Trim(new char[] { ' ', '\n', '\r', '(', '▶', ':' });
            return result;
        }

        /// <summary>
        /// Grab KS data for different pages. Add them to name changes list.
        /// </summary>
        /// <param name="sourceNode">page source node</param>
        private void GrabCommonKSData(HtmlDocument sourceNode, KoreaEquityInfo equity)
        {
            try
            {
                string strPre = sourceNode.DocumentNode.SelectSingleNode(".//pre").InnerText.Trim().ToString();

                string strKeyWords = "";
                string[] itemArr = strPre.Split('◎');
                bool contain_key_word = false;
                int no = strPre.Split('◎').Length;
                //step1 find the key item which contains 변경상장(상호변경) and catch date ,name....and so on.
                for (int i = 0; i < no; i++)
                {
                    string title = "";
                    if (itemArr[i] != "")
                    {
                        title = itemArr[i].Split('①')[0];
                    }
                    if (title.Contains("변경상장") && title.Contains("상호변경"))
                    {
                        contain_key_word = true;
                        strKeyWords = itemArr[i];
                        break;
                    }
                }
                if (contain_key_word == false)
                {
                    return;
                }

                List<string> isins = new List<string>();
                List<string> tickers = new List<string>();
                int strIsinStartPos = strPre.IndexOf(" 코드") + (" 코드".Length);
                int strIsinEndPos = strPre.IndexOf("업종코드");
                if (strIsinEndPos == -1 || (strIsinEndPos - strIsinStartPos) < 10) //为处理出现两项的情况。
                {
                    strIsinEndPos = strPre.IndexOf("심볼");
                }

                string strIsinTemp = strPre.Substring((strIsinStartPos), (strIsinEndPos - strIsinStartPos)).Trim(new char[] { ' ', '\r', '\n', '▶' }).ToString();
                string[] strIsinArr = strIsinTemp.Split('▶');

                string pattern = @"[0-9A-Z]+";
                Regex regex = new Regex(pattern);
                for (var i = 0; i < strIsinArr.Length; i++)
                {
                    string temp = strIsinArr[i].Trim().ToString();
                    if (temp.Contains("표준코드") && temp.Contains("단축코드"))
                    {
                        temp = temp.Substring(temp.IndexOf("표준코드"));
                        MatchCollection match = regex.Matches(temp);
                        if (match.Count == 2)
                        {
                            isins.Add(match[0].Value);
                            tickers.Add(match[1].Value.Substring(1));
                        }
                    }
                }

                if (isins.Count == 0)
                {
                    string msg = string.Format("At GrabCommonKSData(). Can not get detail information for KS announcement at {0}.", equity.AnnouncementTime);
                    return;
                }

                int effectiveStartPos = strKeyWords.IndexOf("변경상장일") + ("변경상장일".Length);
                string effectiveDate = FormatDataWithPos(effectiveStartPos, strKeyWords);
                if (!string.IsNullOrEmpty(effectiveDate))
                {
                    effectiveDate = Convert.ToDateTime(effectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                }

                if (strKeyWords.Contains("㉮"))
                {
                    int startPos = strKeyWords.IndexOf("-");
                    int endPos = strKeyWords.IndexOf("㉯");
                    string legalNamePart = strKeyWords.Substring(startPos, endPos - startPos);
                    int itemCount = 0;
                    while (legalNamePart.Contains("→"))
                    {
                        KoreaEquityInfo item = new KoreaEquityInfo();
                        int koreaNamePos = legalNamePart.IndexOf("→");
                        legalNamePart = legalNamePart.Substring(koreaNamePos + 1);
                        int legalNamePos = legalNamePart.IndexOf("(영문명");
                        string koreaName = legalNamePart.Substring(0, legalNamePos).Trim();
                        string legalName = FormatPrfNameWithPos(legalNamePos + 4, legalNamePart);
                        if (legalName.Substring(legalName.Length - 1).Equals(")"))
                        {
                            legalName = legalName.Substring(0, legalName.Length - 1);
                        }
                        item.KoreaName = koreaName;
                        item.LegalName = legalName;
                        if (isins.Count >= itemCount)
                        {
                            item.ISIN = isins[itemCount];
                            item.Ticker = tickers[itemCount];
                        }
                        item.AnnouncementTime = equity.AnnouncementTime;
                        item.Market = equity.Market;
                        item.RIC = item.Ticker + "." + item.Market;
                        item.EffectiveDate = effectiveDate;
                        itemCount++;
                        nameChanges.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabCommonKSData().    : \r\n" + ex.ToString();
                errors.Add(msg);
            }
        }
               
        /// <summary>
        /// Grab KQ data. title contains "영문명"
        /// </summary>
        /// <param name="item">equity item</param>
        /// <param name="sourceNode">source node</param>
        private void GrabLegalNameDataChangeAction(KoreaEquityInfo item, HtmlDocument sourceNode)
        {
            try
            {
                HtmlNodeCollection tables = sourceNode.DocumentNode.SelectNodes(".//table");
                HtmlNode table = null;
                foreach (var node in tables)
                {
                    HtmlNode knameNode = node.SelectSingleNode(".//tr[2]/td[1]");
                    if (knameNode != null)
                    {
                        if (knameNode.InnerText.Trim().Contains("변경내역"))
                        {
                            table = node;
                            break;
                        }
                    }
                }

                HtmlNodeCollection trs = table.SelectNodes("//tr");
                string koreaName = string.Empty;
                string effectiveDate = string.Empty;
                string legalName = string.Empty;
                for (int i = 0; i < trs.Count; i++)
                {
                    HtmlNode tr = trs[i];
                    string tdTitle = tr.SelectSingleNode("./td[1]").InnerText.Trim();
                    if (tdTitle.Contains("종목명") && koreaName == string.Empty)
                    {
                        koreaName = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        continue;
                    }
                    if (tdTitle.Contains("변경내역"))
                    {
                        tr = trs[i + 1];
                        HtmlNodeCollection tds = tr.SelectNodes(".//td");
                        for (int j = 0; j < tds.Count; j++)
                        {
                            HtmlNode td = tds[j];
                            if (td.InnerText.Trim().Contains("변경후"))
                            {
                                legalName = trs[i + 2].SelectNodes(".//td")[j].InnerText.Trim();
                                i = i + 2;
                                break;
                            }
                        }
                        continue;
                    }
                    if (tdTitle.Contains("상장일"))
                    {
                        effectiveDate = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        break;
                    }
                }

                item.LegalName = legalName;
                item.EffectiveDate = effectiveDate;
                item.KoreaName = koreaName.Replace("(주)", "");
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error fount in GrabLegalNameDataChangeAction(). For announcement at {0}. RIC:{1}. Error message:{2}", item.AnnouncementTime, item.RIC, ex.Message);
                errors.Add(msg);
            }
        }

        #endregion

        #region Formatting

        /// <summary>
        /// Format name changes data.
        /// </summary>
        private void FormatNameChange()
        {
            if (nameChanges.Count > 0)
            {
                FormatNameChangeTemplate(nameChanges);
            }
            if (nameRevised.Count > 0)
            {
                FormatNameChangeTemplate(nameRevised);
            }
        }

        /// <summary>
        /// Format name change for equity list.
        /// </summary>
        /// <param name="list">equity list</param>
        private void FormatNameChangeTemplate(List<KoreaEquityInfo> list)
        {
            foreach (var item in list)
            {
                try
                {
                    item.UpdateDate = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    item.EffectiveDate = DateTime.Parse(item.EffectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                    if (item.KoreaName.Contains("보통주"))
                    {
                        item.KoreaName = item.KoreaName.Replace("보통주", "");
                    }
                    item.KoreaName = item.KoreaName.Replace("(", "").Replace(")", "").Replace(" ", "").Trim();
                    if (item.KoreaName.Length > 14)
                    {
                        item.KoreaName = item.KoreaName.Substring(0, 14);
                    }

                    KoreaEquityInfo equity = KoreaEquityManager.SelectEquityByIsin(item.ISIN);
                    if (equity != null)
                    {
                        item.RIC = equity.RIC;
                        item.OldIDNDisplayName = equity.IDNDisplayName.ToUpper();
                        item.OldKoreaName = equity.KoreaName;
                        item.OldLegalName = equity.LegalName.Trim();
                        item.Type = equity.Type;
                    }
                    else
                    {
                        string msg = string.Format("Please notice that the database does not contains an equity record for RIC:{0}, ISIN:{1}\r\n", item.RIC, item.ISIN);
                        Logger.Log(msg, Logger.LogType.Warning);
                    }
                    KoreaEquityCommon.FormatEQIdnDisplayName(item);
                    UpdateChanges(item);

                    if (!item.IsRevised)
                    {
                        AddEquityToGroup(item);
                    }
                }
                catch (Exception ex)
                {
                    string msg = string.Format("At FormatNameChangeTemplate(). Error found in formatting name change template. RIC:{0}. Error message:{1}", item.RIC, ex.Message);
                    errors.Add(msg);
                }
            }
        }

        /// <summary>
        /// Add the equities with same type and effective date to group. Stored in hash table.
        /// </summary>
        /// <param name="item">equity item</param>
        private void AddEquityToGroup(KoreaEquityInfo item)
        {
            string key = item.EffectiveDate;
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

        /// <summary>
        /// Mark the changes and update equity in DB.
        /// </summary>
        /// <param name="item">equity item</param>
        private void UpdateChanges(KoreaEquityInfo item)
        {
            string updateSql = "";

            updateSql += string.Format(" OldIDNDisplayName = '{0}', IDNDisplayName = '{1}',", item.OldIDNDisplayName, item.IDNDisplayName);

            updateSql += string.Format(" OldKoreaName = N'{0}', KoreaName = N'{1}',", item.OldKoreaName, item.KoreaName);

            updateSql += string.Format(" OldLegalName = N'{0}', LegalName = N'{1}',", item.OldLegalName, item.LegalName);

            updateSql += string.Format(" EffectiveDateChange = '{0}', UpdateDateChange = '{1}' where ISIN = '{2}' and Status = 'Active'", item.EffectiveDate, item.UpdateDate, item.ISIN);

            try
            {
                int row = KoreaEquityManager.UpdateNameChange(updateSql);
                if (row > 0)
                {
                    Logger.Log(string.Format("Update {0} record in Korea Equity Table. RIC:{1}", row.ToString(), item.RIC));
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("At UpdateChanges(). Error message:{0}", ex.Message);
                errors.Add(msg);
            }
        }
        
        #endregion

        #region Distributing

        /// <summary>
        /// Generate FM Files.
        /// </summary>
        private void GenerateFMFiles()
        {
            Logger.Log("Generate FM Files.");
            if (nameRevised.Count > 0)
                GenerateNameChangeFMFile(nameRevised, "(Revised)");
            if (nameChanges.Count > 0)
                GenerateNameChangeFMByGroup(nameChanges);
        }

        /// <summary>
        /// Generate FM Files by the same effective date.
        /// </summary>
        /// <param name="list">name change equities list</param>
        private void GenerateNameChangeFMByGroup(List<KoreaEquityInfo> list)
        {
            ArrayList fmFiles = new ArrayList(equityDate.Keys);
            fmFiles.Sort();
            foreach (string date in fmFiles)
            {
                List<KoreaEquityInfo> equity = (List<KoreaEquityInfo>)equityDate[date];
                GenerateNameChangeFMFile(equity, null);
            }
        }

        /// <summary>
        /// Generate name change FM file for each equity list.
        /// If it is a revised equity, add revised to the end of FM file name.
        /// </summary>
        /// <param name="list">equity list</param>
        /// <param name="revised">if revised</param>
        private void GenerateNameChangeFMFile(List<KoreaEquityInfo> list, string revised)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                errors.Add(msg);
                return;
            }

            try
            {
                DateTime effectiveDate = Convert.ToDateTime(list[0].EffectiveDate);
                string effectiveDateFill = effectiveDate.ToString("dd-MMM-yy");
                string oneDayBefore = effectiveDate.ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                string fileName = "KR FM(Name Change)Korea FM_";
                string rics = CombineAllRics(list);
                fileName = fileName + rics + " (wef " + oneDayBefore + ").xls";
                if (!string.IsNullOrEmpty(revised))
                {
                    fileName = revised + fileName;
                }
                string filePath = Path.Combine(configObj.FM, fileName);  //"C:\\Korea_Auto\\Equity_Warrant\\Name_Change\\" + filename;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                    errors.Add(msg);
                    return;
                }

                GenerateExcelFileTitle(wSheet);
                int row = 2;
                foreach (var item in list)
                {
                    ((Range)wSheet.Cells[row, 1]).NumberFormat = "@";
                    wSheet.Cells[row, 1] = item.UpdateDate;
                    ((Range)wSheet.Cells[row, 2]).NumberFormat = "@";
                    wSheet.Cells[row, 2] = effectiveDateFill;
                    wSheet.Cells[row, 3] = item.RIC;
                    wSheet.Cells[row, 4] = item.RIC;
                    wSheet.Cells[row, 5] = item.ISIN;
                    wSheet.Cells[row, 5] = item.ISIN;
                    if (item.ISIN.Length != 12)
                        ((Range)wSheet.Cells[row, 6]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
                    wSheet.Cells[row, 6] = item.ISIN;
                    ((Range)wSheet.Cells[row, 7]).NumberFormat = "@";
                    wSheet.Cells[row, 7] = item.Ticker;
                    ((Range)wSheet.Cells[row, 8]).NumberFormat = "@";
                    wSheet.Cells[row, 8] = item.Ticker;

                    wSheet.Cells[row, 9] = item.OldLegalName;
                    wSheet.Cells[row, 10] = item.LegalName;
                    wSheet.Cells[row, 11] = item.OldIDNDisplayName;
                    wSheet.Cells[row, 12] = item.IDNDisplayName;
                    wSheet.Cells[row, 13] = item.KoreaName;
                    row++;
                }

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();

                MailToSend mail = new MailToSend();
                mail.ToReceiverList.AddRange(configObj.MailTo);
                mail.CCReceiverList.AddRange(configObj.MailCC);
                mail.MailSubject = Path.GetFileNameWithoutExtension(fileName);
                mail.AttachFileList.Add(filePath);
                mail.MailBody = "Name Change:\t\t" + rics + "\r\n\r\n"
                                + "Effective Date:\t\t" + list[0].EffectiveDate + "\r\n\r\n\r\n\r\n";
                string signature = string.Join("\r\n", configObj.MailSignature.ToArray());
                mail.MailBody += signature;

                AddResult(fileName,filePath,"FM File");
                Logger.Log("Generate FM file. Filepath is " + filePath);
            }
            catch (Exception ex)
            {
                string msg = "Error found in GenerateNameChangeFMFile()   : \r\n" + ex.ToString();
                errors.Add(msg);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        /// <summary>
        /// Combine all RICs to a string contactate with ','
        /// </summary>
        /// <param name="items">equity items</param>
        /// <returns>RICs</returns>
        private string CombineAllRics(List<KoreaEquityInfo> items)
        {
            List<string> rics = new List<string>();

            foreach (KoreaEquityInfo e in items)
            {
                rics.Add(e.RIC);
            }

            return string.Join(",", rics.ToArray());
        }

        /// <summary>
        /// Generate FM file title.
        /// </summary>
        /// <param name="wSheet">worksheet</param>
        private void GenerateExcelFileTitle(Worksheet wSheet)
        {
            if (wSheet.get_Range("C1", Type.Missing).Value2 == null)
            {
                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 30;
                ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 30;
                ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["A:M", System.Type.Missing]).Font.Name = "Arial";

                //fill the cells with color 
                ((Range)wSheet.Cells[1, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                ((Range)wSheet.Cells[1, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                ((Range)wSheet.Cells[1, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                ((Range)wSheet.Cells[1, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                ((Range)wSheet.Cells[1, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                ((Range)wSheet.Cells[1, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                ((Range)wSheet.Cells[1, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                ((Range)wSheet.Cells[1, 8]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                ((Range)wSheet.Cells[1, 9]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                ((Range)wSheet.Cells[1, 10]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                ((Range)wSheet.Cells[1, 11]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                ((Range)wSheet.Cells[1, 12]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                ((Range)wSheet.Cells[1, 13]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

                wSheet.Cells[1, 1] = "Updated Date";
                wSheet.Cells[1, 2] = "Effective Date";
                wSheet.Cells[1, 3] = "Old RIC";
                wSheet.Cells[1, 4] = "New RIC";
                wSheet.Cells[1, 5] = "Old ISIN";
                wSheet.Cells[1, 6] = "NEW ISIN";
                wSheet.Cells[1, 7] = "Old Ticker";
                wSheet.Cells[1, 8] = "New Ticker";
                wSheet.Cells[1, 9] = "Old Legal Name";
                wSheet.Cells[1, 10] = "New Legal Name";
                wSheet.Cells[1, 11] = "Old Display Name";
                wSheet.Cells[1, 12] = "New Display Name";
                wSheet.Cells[1, 13] = "Korea Name";
            }
        }

        /// <summary>
        /// Generate GEDA NDA Files.
        /// </summary>
        private void GenerateGEDANDAFiles()
        {
            Logger.Log("Generate GEDA and NDA files.");


            #region For GEDA and NDA File put into a same folder named with effective date for all korean folder.

            /*
            string folderGEDA = ConfigureOperator.GetGedaFileSaveDir();
            if (string.IsNullOrEmpty(folderGEDA))
            {
                 folderGEDA = CreateTempDirectory(configObj.FM, "GEDA" + DateTime.Today.ToString("yyyyMMdd"));
                 Logger.Log("Please notice that I can not get GEDA File folder. Please check the config in DB. The back up files is in folder " + folderGEDA);              
            }
            string folderNDA = ConfigureOperator.GetNdaFileSaveDir();
            if (string.IsNullOrEmpty(folderNDA))
            {
                folderNDA = CreateTempDirectory(configObj.FM, "NDA" + DateTime.Today.ToString("yyyyMMdd"));
                Logger.Log("Please notice that I can not get NDA File folder. Please check the config in DB. The back up files is in folder " + folderNDA);  
            } 

            for (int i = 0; i < nameChanges.Count; i++)
            {
                KoreaEquityInfo item = nameChanges[i];
                DateTime effectiveDay = DateTime.Parse(item.EffectiveDate);
                DateTime dayBefore = MiscUtil.GetLastTradingDay(effectiveDay, holidayList, 1);
                string oneDayBefore = dayBefore.ToString("yyyy-MM-dd", new CultureInfo("en-US")) 
             
                string gedaFolder = Path.Combine(folderGEDA, oneDayBefore);
                string ndaFolder = Path.Combine(folderNDA, oneDayBefore);
                CreateDirectory(gedaFolder);
                CreateDirectory(ndaFolder);
                
                CreateGEDAFile(nameChanges[i], gedaFolder);
                CreateNDAIAFile(nameChanges[i], ndaFolder);
                CreateNDAQAFile(nameChanges[i], ndaFolder);
            }

            AddResult("GEDA Folder",folderGEDA,"GEDA Folder");
            AddResult("NDA Folder",folderNDA,"NDA Folder");
            
             */
            #endregion


            //For GEDA and NDA file in config path.
            for (int i = 0; i < nameChanges.Count; i++)
            {
                CreateGEDAFile(nameChanges[i], configObj.GEDA);
                CreateNDAIAFile(nameChanges[i], configObj.NDA);
                CreateNDAQAFile(nameChanges[i], configObj.NDA);
            }
        }

        /// <summary>
        /// Create a temp folder if can not get GEDA NDA or EMA file from database.
        /// </summary>
        /// <param name="fmFolder">FM folder path</param>
        /// <param name="desFolder">destination folder name</param>
        /// <returns>combined path</returns>
        private string CreateTempDirectory(string fmFolder, string desFolder)
        {
            if (string.IsNullOrEmpty(fmFolder))
            {
                return Path.Combine(GetOutputFilePath(), "Temp");
            }
            if (string.IsNullOrEmpty(desFolder))
            {
                desFolder = "Temp";
            }

            string upFolder = Path.GetDirectoryName(fmFolder);
            upFolder = upFolder.Substring(0, upFolder.LastIndexOf('\\'));
            return Path.Combine(upFolder, desFolder);
        }

        /// <summary>
        /// Generate GEDA file for an equity.
        /// </summary>
        /// <param name="item">equity</param>
        private void CreateGEDAFile(KoreaEquityInfo item, string dir)
        {
            try
            {
                //string dir = configObj.GEDA;
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                DateTime effectiveDate = DateTime.Parse(item.EffectiveDate);
                //DateTime theDayBeforeEffictiveDate = effectiveDate.AddDays(-1);
                DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
                string fileName = "KR_NameChange_Bulk_Change_" + theDayBeforeEffictiveDate.ToString("yyyyMMdd") + ".txt";
                string path = Path.Combine(dir, fileName);


                if (!File.Exists(@path))
                {
                    StringBuilder new_Legal_Name = new StringBuilder(item.LegalName);
                    for (int i = 0; i < new_Legal_Name.Length; i++)
                    {
                        if (char.IsLower(new_Legal_Name[i]))
                        {
                            new_Legal_Name[i] = char.ToUpper(new_Legal_Name[i]);
                        }
                    }

                    StringBuilder strData = new StringBuilder();
                    strData.Append("RIC\tDSPLY_NAME\tDSPLY_NMLL\t#INSTMOD_TDN_ISSUER_NAME\r\n");
                    strData.Append(item.RIC);
                    strData.Append("\t");
                    strData.Append(item.IDNDisplayName);
                    strData.Append("\t");
                    strData.Append(item.KoreaName);
                    strData.Append("\t");
                    strData.Append(new_Legal_Name);
                    strData.Append("\r\n");
                    File.WriteAllText(path, strData.ToString(), Encoding.UTF8);
                }
                else
                {
                    StreamReader readFileAll = new StreamReader(path);
                    string strDataAll = readFileAll.ReadToEnd();
                    readFileAll.Close();

                    StreamReader readFile = new StreamReader(path);
                    bool exist = false;
                    // readFile.Close();
                    StringBuilder strData = new StringBuilder(strDataAll);
                    string line = readFile.ReadLine();
                    while (line != null)
                    {
                        string temp = line.Split('\t')[0].ToString();
                        if (temp == item.RIC)
                        {
                            exist = true;
                        }
                        line = readFile.ReadLine();
                    }
                    readFile.Close();
                    if (exist == false)
                    {
                        StringBuilder new_Legal_Name = new StringBuilder(item.LegalName);
                        for (int i = 0; i < new_Legal_Name.Length; i++)
                        {
                            if (char.IsLower(new_Legal_Name[i]))
                            {
                                new_Legal_Name[i] = char.ToUpper(new_Legal_Name[i]);
                            }
                        }
                        StreamWriter writeFile = new StreamWriter(path);
                        //StringBuilder strData = new StringBuilder();
                        strData.Append(item.RIC);
                        strData.Append("\t");
                        strData.Append(item.IDNDisplayName);
                        strData.Append("\t");
                        strData.Append(item.KoreaName);
                        strData.Append("\t");
                        strData.Append(new_Legal_Name);
                        strData.Append("\r\n");
                        writeFile.Write(strData);
                        writeFile.Close();
                    }
                }

                if (!filesResultList.Contains(fileName))
                {
                    TaskResultEntry entry = new TaskResultEntry(fileName, "GEDA File", path, FileProcessType.GEDA_BULK_RIC_CHANGE);
                    filesResultList.Add(fileName, entry);
                    Logger.Log("Generate GEDA file successfully. Filepath is " + path);
                }


            }
            catch (Exception ex)
            {
                string msg = string.Format("Error happens when trying to create the GEDA . Ex: {0}", ex.Message);
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate NDA IA file for an equity.
        /// </summary>
        /// <param name="item">an equity</param>
        private void CreateNDAIAFile(KoreaEquityInfo item, string dir)
        {
            try
            {
                //string dir = configObj.NDA;
                string suffix = FormatNdaSuffix(item);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                DateTime effectiveDate = DateTime.Parse(item.EffectiveDate);
                DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
                string fileName = "KR_EQName" + theDayBeforeEffictiveDate.ToString("yyyyMMdd") + "IAChg.csv";
                string path = Path.Combine(dir, fileName);
                string assetCommonName = FormatAssetCommonName(item);
                if (!File.Exists(@path))
                {
                    StringBuilder strData = new StringBuilder();
                    strData.Append("ISIN,ASSET COMMON NAME\r\n");
                    strData.Append(item.ISIN + "," + assetCommonName + suffix + "\r\n");
                    File.WriteAllText(path, strData.ToString(), Encoding.UTF8);
                }
                else
                {
                    StreamReader readFileAll = new StreamReader(path);
                    string strDataAll = readFileAll.ReadToEnd();
                    readFileAll.Close();

                    StreamReader readFile = new StreamReader(path);
                    bool exist = false;
                    StringBuilder strData = new StringBuilder(strDataAll);
                    string line = readFile.ReadLine();
                    while (line != null)
                    {
                        string temp = line.Split(',')[0].ToString();
                        if (temp == item.ISIN)
                        {
                            exist = true;
                        }
                        line = readFile.ReadLine();
                    }
                    readFile.Close();
                    if (exist == false)
                    {
                        StreamWriter writeFile = new StreamWriter(path);
                        strData.Append(item.ISIN + "," + assetCommonName + suffix + "\r\n");
                        writeFile.Write(strData);
                        writeFile.Close();
                    }
                }

                if (!filesResultList.Contains(fileName))
                {
                    TaskResultEntry entry = new TaskResultEntry(fileName, "NDA IA File", path, FileProcessType.NDA);
                    filesResultList.Add(fileName, entry);
                    Logger.Log("Generate NDA IA file successfully. Filepath is " + path);
                }


            }
            catch (Exception ex)
            {
                string msg = string.Format("Error happens when trying to create the NDA IA . Ex: {0}", ex.Message);
                errors.Add(msg);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private string FormatAssetCommonName(KoreaEquityInfo item)
        {
            string company = KoreaEquityCommon.ClearCoLtdForName(item.LegalName);
            if (item.Type.Equals("PRF"))
            {
                int index = 0;
                string ending = null;
                if (string.IsNullOrEmpty(item.PrfEnd))
                {
                    Regex regex = new Regex("[0-9]+P");
                    Match match = regex.Match(company);
                    if (match.Success)
                    {
                        index = match.Index;
                        ending = match.Value;
                    }
                }
                else
                {
                    index = company.IndexOf(item.PrfEnd);
                    ending = item.PrfEnd;
                }
                company = company.Replace(ending, "");
                FormatAssetNameNoEnd(ref company);
                if (!string.IsNullOrEmpty(ending))
                {
                    company = company.Insert(index, ending);
                }
            }
            else
            {
                FormatAssetNameNoEnd(ref company);
            }
            return company;

        }

        private void FormatAssetNameNoEnd(ref string company)
        {
            StringBuilder newDisPlayName = new StringBuilder(company);

            for (int i = 0; i < newDisPlayName.Length; i++)
            {
                if (i == 0)
                {
                    if (char.IsLower(newDisPlayName[i]))
                        newDisPlayName[i] = char.ToUpper(newDisPlayName[i]);
                }
                else if (newDisPlayName[i - 1] == ' ')
                {
                    if (char.IsLower(newDisPlayName[i]))
                        newDisPlayName[i] = char.ToUpper(newDisPlayName[i]);

                }
                else
                {
                    if (char.IsUpper(newDisPlayName[i]))
                        newDisPlayName[i] = char.ToLower(newDisPlayName[i]);
                }
            }
            company = newDisPlayName.ToString();
        }

        /// <summary>
        /// Generate NDA QA file for an equity.
        /// </summary>
        /// <param name="item">an equity</param>
        private void CreateNDAQAFile(KoreaEquityInfo item, string dir)
        {
            try
            {
                //string dir = configObj.NDA;
                string suffix = item.Type != null ? (" " + item.Type) : "";
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                DateTime effectiveDate = DateTime.Parse(item.EffectiveDate);
                //DateTime theDayBeforeEffictiveDate = effectiveDate.AddDays(-1);
                DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
                string fileName = "KR_EQName" + theDayBeforeEffictiveDate.ToString("yyyyMMdd") + "QAChg.csv";
                string path = Path.Combine(dir, fileName);

                if (!File.Exists(@path))
                {
                    //StringBuilder strColumn = new StringBuilder();
                    StringBuilder strData = new StringBuilder();
                    strData.Append("RIC,ASSET SHORT NAME,ASSET COMMON NAME\r\n");
                    string temp = item.RIC;
                    string tempS;
                    string tailS = "";
                    int j;
                    for (j = 0; j < temp.Length; j++)
                    {
                        if (temp[j] == '.')
                            break;
                    }
                    tempS = temp.Substring(0, j);
                    // tempS = tempS + "F";
                    for (int k = j; k < temp.Length; k++)
                    {
                        tailS = tailS + temp[k];
                    }
                    //temp.Insert(temp.Length - 4, "F");
                    strData.Append(item.RIC + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                    strData.Append(tempS + "F" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                    strData.Append(tempS + "S" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                    strData.Append(tempS + "stat" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                    strData.Append(tempS + "ta" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                    strData.Append(tempS + "bl" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                    File.WriteAllText(path, strData.ToString(), Encoding.UTF8);
                }
                else
                {
                    //StreamReader readFile = new StreamReader(pathNDA_QA);
                    StreamReader readFileAll = new StreamReader(path);
                    string strDataAll = readFileAll.ReadToEnd();
                    readFileAll.Close();
                    // strDataAll = null;
                    StreamReader readFile = new StreamReader(path);
                    bool exist = false;
                    // readFile.Close();
                    StringBuilder strData = new StringBuilder(strDataAll);
                    string line = readFile.ReadLine();
                    while (line != null)
                    {
                        string temp = line.Split(',')[0].ToString();
                        if (temp == item.RIC)
                        {
                            exist = true;
                        }
                        line = readFile.ReadLine();
                    }
                    readFile.Close();
                    if (exist == false)
                    {
                        StreamWriter writeFile = new StreamWriter(path);
                        //StringBuilder strData = new StringBuilder();
                        string temp = item.RIC;
                        string tempS;
                        string tailS = "";
                        int j;
                        for (j = 0; j < temp.Length; j++)
                        {
                            if (temp[j] == '.')
                                break;
                        }
                        tempS = temp.Substring(0, j);
                        // tempS = tempS + "F";
                        for (int k = j; k < temp.Length; k++)
                        {
                            tailS = tailS + temp[k];
                        }
                        //temp.Insert(temp.Length - 4, "F");
                        strData.Append(item.RIC + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                        strData.Append(tempS + "F" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                        strData.Append(tempS + "S" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                        strData.Append(tempS + "stat" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                        strData.Append(tempS + "ta" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                        strData.Append(tempS + "bl" + tailS + "," + item.IDNDisplayName + "," + item.IDNDisplayName + suffix + "\r\n");
                        writeFile.Write(strData);
                        writeFile.Close();
                    }

                }

                if (!filesResultList.Contains(fileName))
                {
                    TaskResultEntry entry = new TaskResultEntry(fileName, "NDA QA File", path, FileProcessType.NDA);
                    filesResultList.Add(fileName, entry);
                    Logger.Log("Generate NDA QA file successfully. Filepath is " + path);
                }


            }
            catch (Exception ex)
            {
                string msg = string.Format("Error happens when trying to create the NDA QA. Ex: {0}", ex.Message);
                errors.Add(msg);
            }

        }

        /// <summary>
        /// Format NDA IA suffix of column "ASSET COMMON NAME"
        /// </summary>
        /// <param name="type">equity type</param>
        /// <returns>suffix</returns>
        private string FormatNdaSuffix(KoreaEquityInfo item)
        {
            string type = item.Type;
            if (string.IsNullOrEmpty(type))
            {
                string msg = string.Format("Name change item {0} do not have a type. Please confirm.", item.RIC);
                Logger.Log(msg, Logger.LogType.Warning);
                return "";
            }
            string suffix = "";

            if (type.Equals("ORD"))
            {
                suffix = " Ord Shs";
            }
            else if (type.Equals("PRF"))
            {
                suffix = " Prf Shs";
            }
            else if (type.Equals("KDR"))
            {
                suffix = " KDR";
            }
            else
            {
                string msg = string.Format("Name change for {0} occurs. RIC:{1}. Please format the suffix of column: 'ASSET COMMON NAME' in NDA IA file. ", item.Type, item.RIC);
                Logger.Log(msg, Logger.LogType.Warning);
            }
            return suffix;
        }

        /// <summary>
        /// If file folder not exsits. Create the folder.
        /// </summary>
        /// <param name="directory"></param>
        private void CreateDirectory(string directory)
        {
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }

        #endregion

        #region Results. For result list and log

        /// <summary>
        /// Create task result list for GEDA and NDA file.         
        /// </summary>
        private void CreateTaskResult()
        {
            if (filesResultList.Count == 0)
            {
                return;
            }
            ArrayList resultOrder = new ArrayList(filesResultList.Keys);
            resultOrder.Sort();
            foreach (string fileName in resultOrder)
            {
                try
                {
                    TaskResultList.Add(filesResultList[fileName] as TaskResultEntry);
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
            AddResult("Log",Logger.FilePath,"LOG FILE");
        }

        #endregion

    }
}
