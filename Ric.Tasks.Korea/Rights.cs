using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Threading;
using System.Drawing;
using HtmlAgilityPack;
using System.IO;
using System.Data;
using Ric.Db.Manager;
using Ric.Core;
using Ric.Util;
using System.Net;
using Ric.Db.Info;

namespace Ric.Tasks.Korea
{
    public class Rights : GeneratorBase
    {
        private List<RightsTemplate> raList = new List<RightsTemplate>();
        private KOREARightsGeneratorConfig configObj = null;

        protected override void Start()
        {
            StartRightsJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREARightsGeneratorConfig;
            if (string.IsNullOrEmpty(configObj.StartDate))
            {
                configObj.StartDate = DateTime.Now.ToString("yyyy-MM-dd");
            }
            if (string.IsNullOrEmpty(configObj.EndDate))
            {
                configObj.EndDate = DateTime.Now.ToString("yyyy-MM-dd");
            }
            if ((DateTime.Parse(configObj.StartDate)).CompareTo(DateTime.Parse(configObj.EndDate)) > 0)
            {
                string temp = configObj.StartDate;
                configObj.StartDate = configObj.EndDate;
                configObj.EndDate = temp;
            }
        }

        public void StartRightsJob()
        {
            GrabRightsDataFromKindWebpage();
            if (raList.Count > 0)
            {
                GrabDataFromISINWebpage();
                FormatRightsData();
                GenerateRightsAddFileXls();
                GenerateRightsDropFileXls();
                GenerateGEDAFiles();
                GenerateNDAQAFile();
                GenerateNdaFutureDatingFile();
                GenerateNDATickLotFile();
                UpdateRightsToDb();
            }
            else
            {
                Logger.Log("There is no data grabbed.", Logger.LogType.Info);
            }
            AddResult("LOG", Logger.FilePath, "Log File");
        }

        /// <summary>
        /// Grab rights data from kind webpage using the date input. Judge the data is KQ or KS.
        /// </summary>
        private void GrabRightsDataFromKindWebpage()
        {
            Logger.Log("Grab Rights data from Kind Webpage");
            string startDate = configObj.StartDate;
            string endDate = configObj.EndDate;

            string dataStartDate = (DateTime.Parse(endDate).AddMonths(-2)).ToString("yyyy-MM-dd");
            string uri = string.Format("http://kind.krx.co.kr/disclosure/searchtotalinfo.do");
            try
            {
                //string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EC%84%9C&repIsuSrtCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=300&scn=mktact&srchFd=2&kwd=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EC%84%9C&fromData={0}&toData={1}", dataStartDate, endDate);
                //2014-02-10 kind page changed
                string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EC%84%9C&repIsuSrtCd=&isurCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=300&scn=mktact&srchFd=2&kwd=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EC%84%9C&fromData={0}&toData={1}", dataStartDate, endDate);

                string pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, postData);
                HtmlDocument htc = new HtmlDocument();
                if (!string.IsNullOrEmpty(pageSource))
                    htc.LoadHtml(pageSource);
                if (htc != null)
                {
                    HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes("//dl/dt");
                    HtmlNodeCollection ddCollections = htc.DocumentNode.SelectNodes("//dl/dd");

                    int count = 0;
                    if (nodeCollections != null)
                        count = nodeCollections.Count;
                    for (var i = 0; i < count; i++)
                    {

                        HtmlNode ddNode = ddCollections[i].SelectSingleNode(".//span");
                        var dtnode = nodeCollections[i] as HtmlNode;
                        string title = dtnode.SelectSingleNode(".//span/a").InnerText;
                        if (title.Contains("신주인수권증서"))
                        {
                            HtmlNode nodeDate = dtnode.SelectSingleNode("./em");
                            if (nodeDate != null)
                            {
                                DateTime anouncementDate = new DateTime();
                                anouncementDate = DateTime.Parse(nodeDate.InnerText.Trim(), new CultureInfo("en-US"));
                                if (anouncementDate < DateTime.Parse(startDate))
                                {
                                    return;
                                }
                            }
                            //string parameter = dtnode.SelectSingleNode(".//strong/a").Attributes["onclick"].Value.Trim().ToString();
                            //parameter = parameter.Split('(')[1].Split(')')[0].Trim(new char[] { ' ', '\'', '(', ')' }).ToString();

                            // Get company information

                            //string companyPostData = string.Format("method=searchCompanySummaryOvrvwDetail&menuIndex=0&strIsurCd={0}&lstCd=undefined&taskDd=&spotIsuTrdMktTpCd=&methodType=0", parameter);
                            //string companyPageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, companyPostData);
                            //HtmlDocument companyDoc = new HtmlDocument();
                            //if (!string.IsNullOrEmpty(companyPageSource))
                            //    companyDoc.LoadHtml(companyPageSource);

                            string judge = GetDutyCode(ddNode);

                            //judge 是判断 KS or KQ 的
                            //string url = string.Format("http://kind.krx.co.kr/common/companysummary.do?method=searchCompanySummary&strIsurCd={0}&lstCd=undefined", parameter);
                            //AdvancedWebClient wc = new AdvancedWebClient();
                            //HtmlDocument doc = WebClientUtil.GetHtmlDocument(url, 300000, null);
                            //string judge = string.Empty;
                            //if (doc != null)
                            //    judge = doc.DocumentNode.SelectSingleNode(".//tr[2]/td[2]").InnerText.Trim().ToString();
                            //if (string.IsNullOrEmpty(judge))
                            //    continue;
                            string url = string.Empty;
                            // string judge = string.Empty;
                            HtmlDocument doc = null;

                            if (string.IsNullOrEmpty(judge))
                                continue;

                            string attribute = dtnode.SelectSingleNode(".//span/a").Attributes["onclick"].Value.Trim().ToString();
                            attribute = attribute.Split('(')[1].Split(')')[0].Trim(new char[] { ' ', '\'', ',' }).ToString();
                            string param1 = attribute.Split(',')[0].Trim(new char[] { ' ', '\'', ',' }).ToString();
                            string param2 = attribute.Split(',')[1].Trim(new char[] { ' ', '\'', ',' }).ToString();
                            url = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno={1}&viewerhost=&viewerport=", param1, param2);

                            string source = WebClientUtil.GetDynamicPageSource(url, 300000, null);
                            if (!string.IsNullOrEmpty(source))
                            {
                                doc = new HtmlDocument();
                                doc.LoadHtml(source);
                            }
                            string ticker = doc.DocumentNode.SelectSingleNode(".//header/h1").InnerText.Trim().ToString();
                            if (!string.IsNullOrEmpty(ticker))
                            {
                                Match m = Regex.Match(ticker, @"\(([0-9a-zA-Z]{6})\)");
                                if (m == null)
                                {
                                    string msg = "Cannot get ticker numbers in ." + ticker;
                                    Logger.Log(msg, Logger.LogType.Error);
                                    return;
                                }
                                ticker = m.Groups[1].Value;
                            }
                            param1 = param1.Insert(4, "/").Insert(7, "/").Insert(10, "/");
                            string param3 = judge.Equals("KS") ? "68915" : (judge.Equals("KQ") ? "70926" : null);
                            url = string.Format("http://kind.krx.co.kr/external/{0}/{1}/{2}.htm", param1, param2, param3);
                            doc = WebClientUtil.GetHtmlDocument(url, 300000, null);
                            if (doc != null)
                            {
                                if (judge.Equals("KS"))
                                {
                                    GrabKSDataAction(doc, ticker);
                                }

                                if (judge.Equals("KQ"))
                                {
                                    GrabKQDataAction(doc, ticker);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in Grab Rights Data From KindWebpage: \r\n" + ex.ToString() + "   InnerException:    " + ex.InnerException;
                Logger.Log(msg, Logger.LogType.Error);
            }
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
        /// Grab KS data from html.
        /// </summary>
        /// <param name="doc">html</param>
        /// <param name="ticker">ticker number</param>
        private void GrabKSDataAction(HtmlDocument doc, string ticker)
        {
            Logger.Log(string.Format("Grab KS Data For Ticker: {0}", ticker));
            try
            {
                string strNodePre = doc.DocumentNode.SelectSingleNode(".//pre").InnerText.Trim().ToString();
                if (!string.IsNullOrEmpty(strNodePre))
                {
                    RightsTemplate ra = new RightsTemplate();
                    RightsTemplate rd = new RightsTemplate();
                    string[] lines = strNodePre.Split("\n".ToArray());
                    int koreanamePos = strNodePre.IndexOf("1. 상장종목") + ("1. 상장종목".Length);
                    int addEffectivePos = strNodePre.IndexOf("2. 상장일") + ("2. 상장일".Length);
                    int dropEffectivePos = strNodePre.IndexOf("3. 상장폐지일") + ("3. 상장폐지일".Length);
                    int quantityPos = strNodePre.IndexOf("4. 신주인수권 증서의 수") + ("4. 신주인수권 증서의 수".Length);
                    int QuantityNextPos = strNodePre.IndexOf("5. 목적주권 주식의 종류");
                    int tempVarPos = strNodePre.IndexOf("5. 목적주권 주식의 종류") + ("5. 목적주권 주식의 종류".Length);
                    int tempVarNextPos = strNodePre.IndexOf("6. 목적주권 1주의 발행가액");
                    int strikePos = strNodePre.IndexOf("6. 목적주권 1주의 발행가액") + ("6. 목적주권 1주의 발행가액".Length);
                    int strikeNextPos = strNodePre.IndexOf("7. 목적주권 청약개시일");
                    int isinKoreaCodePos = strNodePre.IndexOf("9. 코드") + ("9. 코드\n".Length);

                    string koreaName = VariableFormat(strNodePre, koreanamePos);
                    koreaName = koreaName.Contains("신주인수권증서") ? koreaName.Replace("신주인수권증서", "").Trim(new char[] { ' ', ':', ',' }).ToString() : koreaName.Trim().ToString();
                    string addEffectiveDate = VariableFormat(strNodePre, addEffectivePos);
                    addEffectiveDate = Convert.ToDateTime(addEffectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                    string dropEffectiveDate = VariableFormat(strNodePre, dropEffectivePos);
                    dropEffectiveDate = Convert.ToDateTime(dropEffectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();


                    string quantityTemp = strNodePre.Substring(quantityPos, (QuantityNextPos - quantityPos));
                    string tempVar = strNodePre.Substring(tempVarPos, (tempVarNextPos - tempVarPos));
                    string strikeTemp = strNodePre.Substring(strikePos, (strikeNextPos - strikePos));
                    string isinCodeTemp = strNodePre.Substring(isinKoreaCodePos);

                    if ((quantityTemp.Contains("신주인수권증서")) && (strikeTemp.Contains("신주인수권증서")) && (tempVar.Contains("신주인수권증서")))
                    {
                        string[] quantityTempArr = quantityTemp.Split('-');
                        string[] strikeTempArr = strikeTemp.Split('-');
                        string[] tempVarArr = tempVar.Split('-');
                        string[] isinCodeTempArr = isinCodeTemp.Split('-');
                        if ((quantityTempArr.Length == strikeTempArr.Length) && (strikeTempArr.Length == isinCodeTempArr.Length) && (isinCodeTempArr.Length == tempVarArr.Length))
                        {
                            for (var i = 1; i < quantityTempArr.Length; i++)
                            {
                                RightsTemplate raTemp = new RightsTemplate();
                                RightsTemplate rdTemp = new RightsTemplate();
                                string quantity = quantityTempArr[i].Split(':')[1].Trim(new char[] { ' ', ',', '증', '서', '\n' }).ToString();
                                quantity = quantity.Contains(',') ? quantity.Replace(",", "") : quantity;
                                string strike = strikeTempArr[i].Split(':')[1].Trim().ToString();
                                strike = strike.Substring(0, strike.IndexOf('원')).Trim().ToString();
                                strike = strike.Contains(',') ? strike.Replace(",", "") : strike;
                                string isinCodeToSpilt = isinCodeTempArr[i].Split(':')[1].Trim().ToString();
                                string isinReal = isinCodeToSpilt.Split('(')[0].Trim(new char[] { ' ', '(', ':' }).ToString();
                                string codeReal = Regex.Split(isinCodeToSpilt, "단축코드", RegexOptions.IgnoreCase)[1].Trim(new char[] { ' ', ')' }).ToString();
                                string tempVarReal = tempVarArr[i].Split(':')[1].Replace("기명식", "").Trim(new char[] { ' ', ',', '\n' });
                                raTemp.KoreaName = koreaName;
                                raTemp.AddEffectiveDate = addEffectiveDate;
                                raTemp.DropEffectiveDate = dropEffectiveDate;
                                raTemp.Edcoid = ".KS";
                                raTemp.ISIN = isinReal;
                                raTemp.KoreaCode = codeReal;
                                raTemp.QuantityOfRights = quantity;
                                raTemp.StrikePrice = strike;
                                raTemp.TempVar = tempVarReal;
                                raList.Add(raTemp);
                            }
                        }
                    }
                    else
                    {
                        string quantityOfRights = VariableFormat(strNodePre, quantityPos);
                        quantityOfRights = quantityOfRights.Replace("증서", "").ToString();
                        string strikePrice = VariableFormat(strNodePre, strikePos);
                        strikePrice = strikePrice.Substring(0, strikePrice.IndexOf('원')).Trim().ToString();
                        string isinKoreaCode = VariableFormat(strNodePre, isinKoreaCodePos);
                        string isin = isinKoreaCode.Split(':')[1].Split('(')[0].Trim().ToString();
                        string koreaCode = Regex.Split(isinKoreaCode, "단축코드", RegexOptions.IgnoreCase)[1].Trim(new char[] { ' ', ')', '(' }).ToString();

                        ra.Ticker = ticker;
                        ra.KoreaName = koreaName;
                        ra.AddEffectiveDate = addEffectiveDate;
                        ra.DropEffectiveDate = dropEffectiveDate;
                        ra.QuantityOfRights = quantityOfRights;
                        ra.StrikePrice = strikePrice;
                        ra.ISIN = isin;
                        ra.KoreaCode = koreaCode;
                        ra.Edcoid = ticker + ".KS";
                        ra.RIC = ticker + "_r.KS";
                        raList.Add(ra);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in Grab KQ Data for Ticker: {0}: \r\n {1}", ticker, ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Grab KQ data from html.
        /// </summary>
        /// <param name="doc">html</param>
        /// <param name="ticker">ticker number</param>
        private void GrabKQDataAction(HtmlDocument doc, string ticker)
        {
            Logger.Log(string.Format("Grab KQ Data For Ticker: {0}", ticker));
            try
            {
                HtmlNodeCollection tableNodes = doc.DocumentNode.SelectNodes("//table");
                HtmlNode table = null;
                foreach (var item in tableNodes)
                {
                    string strIsin = item.SelectSingleNode(".//tr[3]/td[2]").InnerText.Trim().ToString();
                    string strCode = item.SelectSingleNode(".//tr[3]/td[3]").InnerText.Trim().ToString();
                    if (strIsin.Equals("표준코드") && strCode.Equals("단축코드"))
                        table = item;
                }

                RightsTemplate ra = new RightsTemplate();
                RightsTemplate rd = new RightsTemplate();

                string strKoreaName = "";
                string isin = "";
                string strKoreaCode = "";
                string strQuantityOfRights = "";
                string strStrikePrice = "";
                bool isinFlag = false;
                string strAddEffectiveDate = "";
                string strDropEffectiveDate = "";
                var trs = table.SelectNodes("//tr");
                bool strikeDone = false;

                foreach (var tr in trs)
                {
                    if (isinFlag == true)
                    {
                        isin = tr.SelectSingleNode("./td[1]").InnerText.Trim().ToString();
                        strKoreaCode = tr.SelectSingleNode("./td[2]").InnerText.Trim().ToString();
                        isinFlag = false;
                        continue;
                    }
                    Regex regex = new Regex("상장종목명");
                    Match match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        strKoreaName = tr.SelectSingleNode("./td[2]").InnerText.Trim().ToString();
                        continue;
                    }
                    regex = new Regex("코드명");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        isinFlag = true;
                        continue;
                    }
                    regex = new Regex("신주인수권 증서의");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        strQuantityOfRights = tr.SelectSingleNode("./td[2]").InnerText.Trim().ToString();
                        continue;
                    }
                    regex = new Regex("1주당 발행가액");
                    match = regex.Match(tr.InnerText);
                    if (match.Success && !strikeDone)
                    {
                        strStrikePrice = tr.SelectSingleNode("./td[2]").InnerText.Trim().ToString();
                        strikeDone = true;
                        continue;
                    }
                    regex = new Regex("상장일");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        strAddEffectiveDate = tr.SelectSingleNode("./td[2]").InnerText.Trim().ToString();
                        continue;
                    }
                    regex = new Regex("상장폐지일");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        strDropEffectiveDate = tr.SelectSingleNode("./td[2]").InnerText.Trim().ToString();
                        continue;
                    }
                }
                string koreaName = strKoreaName.Trim().ToString();
                koreaName = koreaName.Contains("(주)") ? koreaName.Replace("(주)", "") : koreaName;

                string koreaCode = strKoreaCode.Trim().ToString();
                string quatityOfRights = strQuantityOfRights.Contains(",") ? strQuantityOfRights.Replace(",", "") : strQuantityOfRights;
                string strikePrice = strStrikePrice.Contains(",") ? strStrikePrice.Replace(",", "") : strStrikePrice;
                string addEffectiveDate = Convert.ToDateTime(strAddEffectiveDate.Trim()).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                string dropEffectiveDate = Convert.ToDateTime(strDropEffectiveDate.Trim()).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();

                ra.Ticker = ticker;
                ra.KoreaName = koreaName;
                ra.KoreaCode = koreaCode;
                ra.ISIN = isin;
                ra.QuantityOfRights = quatityOfRights;
                ra.StrikePrice = strikePrice;
                ra.AddEffectiveDate = addEffectiveDate;
                ra.DropEffectiveDate = dropEffectiveDate;
                ra.Edcoid = ticker + ".KQ";
                ra.RIC = ticker + "_r.KQ";
                raList.Add(ra);
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in Grab KQ Data Action() for Ticker: {0}    : \r\n {1}", ticker, ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Cut viriable string from given position
        /// </summary>
        /// <param name="strPre">string to cut</param>
        /// <param name="strPos">position</param>
        /// <returns></returns>
        private string VariableFormat(string strPre, int strPos)
        {
            string result = "";
            char[] charArr = strPre.ToCharArray();
            while (charArr[strPos] != '\n')
            {
                result += charArr[strPos];
                if ((strPos + 1) != charArr.Length)
                    strPos++;
                else
                    break;
            }
            result = result.Trim(new char[] { ' ', ':', ',', '-' }).ToString();
            result = result.Contains(",") ? result.Replace(",", "") : result;
            return result;
        }

        /// <summary>
        /// Grab legal names from ISIN Webpage.
        /// Format QA Short name and QA Common name, because they are generated from legal name.
        /// </summary>
        private void GrabDataFromISINWebpage()
        {
            Logger.Log("Grab legal names from ISIN WebPage.");
            string url = @"http://isin.krx.co.kr/srch/srch.do?method=srchPopup55";
            string postData = string.Empty;
            HtmlDocument doc = null;
            string legalName = string.Empty;

            foreach (var item in raList)
            {
                try
                {
                    postData = string.Format("stdcd_type=55&std_cd={0}", item.ISIN);//신주인수권증권/증서
                    doc = GetDocRetryMechanism(url, postData, 5, 3000);

                    if (doc == null)
                        continue;

                    var table = doc.DocumentNode.SelectNodes(".//table")[0].SelectNodes(".//table")[0];
                    legalName = table.SelectNodes(".//tr")[2].SelectNodes(".//td")[1].InnerText.ToString();//tr2td1   MIRAE ASSET SECURITIES ELW 4220

                    if (string.IsNullOrEmpty(legalName))
                        Logger.Log(string.Format("get isin:{0} 's legaName error.", item.ISIN), Logger.LogType.Warning);

                    item.LegalName = legalName.Trim().ToUpper();
                    item.QAShortName = GetQAShortName(item.LegalName);
                    item.QACommonName = GetQACommonName(item.QAShortName);
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
        }

        private string GetQACommonName(string p)
        {
            return p + " RTS";
        }

        private string GetQAShortName(string legalName)
        {
            int pLength = legalName.Length;

            if (pLength <= 12)
                return legalName;

            string endString = legalName.Substring(pLength - 3, 3);
            string endLegalName = string.Empty;

            if (endString[2].Equals('R') && char.IsDigit(endString[1]) && char.IsDigit(endString[0]))
                return string.Format("{0} {1}", legalName.Substring(0, 12).Trim(), endString);

            if (endString[2].Equals('R') && char.IsDigit(endString[1]))
                return string.Format("{0} {1}", legalName.Substring(0, 12).Trim(), endString.Substring(1, 2));

            return string.Format("{0} {1}", legalName.Substring(0, 12).Trim(), "![digit+R]");
        }

        private HtmlDocument GetDocRetryMechanism(string url, string postData, int retryTimes, int waitSecond)
        {
            HtmlDocument doc = new HtmlDocument();

            try
            {
                for (int i = 0; i < retryTimes; i++)
                {
                    try
                    {
                        doc = GetDocByISIN(url, postData);

                        if (doc != null)
                            break;
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(waitSecond);

                        if (i == retryTimes - 1)
                        {
                            string msg = string.Format("url:{0}     retryTimes:{1}      msg:{3}", url, retryTimes.ToString(), ex.Message);
                            Logger.Log(msg, Logger.LogType.Error);
                        }
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

            return doc;
        }

        private HtmlDocument GetDocByISIN(string url, string post)
        {
            try
            {
                HtmlDocument doc = new HtmlDocument();
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)";
                request.Method = "POST";
                request.Referer = @"http://isin.krx.co.kr/srch/srch.do?method=srchList";
                request.KeepAlive = true;
                request.Headers["Accept-Language"] = "zh-CN,zh;q=0.8,en;q=0.6";
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                request.ContentType = "application/x-www-form-urlencoded";
                byte[] buf = Encoding.UTF8.GetBytes(post);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);
                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();
                doc.Load(httpResponse.GetResponseStream(), Encoding.UTF8);
                return doc;
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

        /// <summary>
        /// Right Data Format Method
        /// </summary>
        private void FormatRightsData()
        {
            Logger.Log("Format Rights data");

            if (raList.Count > 0)
            {
                for (var i = 0; i < raList.Count; i++)
                {
                    try
                    {
                        var item = raList[i] as RightsTemplate;
                        item.KoreaCode = item.KoreaCode.Substring(1).ToString();
                        item.Currency = "KRW";
                        item.CountryHeadquarters = "KOR";
                        item.OldMSCI = "62";
                        item.RecordType = "97";
                        item.IssueClassification = "RT";
                        item.LotSize = "1";
                        GenerateNewRic(item);  //Generate the new RIC
                        //FormatQANames(item);
                    }
                    catch (Exception ex)
                    {
                        string msg = "Error found in RightsDataTemplateFormat()   : \r\n" + ex.ToString();
                        Logger.Log(msg, Logger.LogType.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Update rights list to database.
        /// </summary>
        private void UpdateRightsToDb()
        {
            try
            {
                int row = KoreaRightsManager.UpdateRights(raList);
                string msge = string.Format("Updated {0} Rights records in database.", row);
                Logger.Log(msge);
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in update Rights records to database.", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Format QA Short Name and QA Common Name
        /// QA Short Name: length less than 16
        /// QA Common Name: legal name length less than 36
        /// </summary>
        /// <param name="item"></param>
        private void FormatQANames(RightsTemplate item)
        {
            try
            {
                if (item.LegalName.Length > 32)
                {
                    item.QACommonName = item.LegalName.Substring(0, 32).ToUpper() + " RTS";
                }
                else
                {
                    item.QACommonName = item.LegalName.ToUpper() + " RTS";
                }

                if (string.IsNullOrEmpty(item.QAShortName))
                {
                    string legalName = item.LegalName.ToUpper();
                    Regex regex = new Regex("[0-9]+R");
                    Match match = regex.Match(legalName);
                    if (!match.Success)
                    {
                        item.QAShortName = item.LegalName.Substring(0, 16).ToUpper();
                    }
                    else
                    {
                        string ending = match.Value;
                        string prefix = item.LegalName.ToUpper().Replace(ending, "").Trim();
                        item.QAShortName = prefix.Substring(0, 15 - ending.Length) + " " + ending;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found when formatting QA Short Name For Rights:{0}, Error message:{1}", item.RIC, ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Ric Generate Method
        /// </summary>
        /// <param name="ra">Rights item</param>
        private void GenerateNewRic(RightsTemplate ra)
        {
            try
            {
                string result = "";
                if ((ra.RIC == null) && (ra.TempVar != string.Empty))
                {
                    ISINQuery query = new ISINQuery("", "", "99", "", ra.ISIN);
                    List<ISINTemp> isinList = Common.getISINListFromISINWebPage(query);
                    if (isinList == null || isinList.Count == 0)
                    {
                        Logger.Log(string.Format("Cannot find ISIN {0} in ISIN webpage", ra.ISIN));
                    }
                    else if (isinList.Count > 1)
                    {
                        Logger.Log(string.Format("Find two items for ISIN {0} in ISIN webpage. Choose the first one.", ra.ISIN));
                    }
                    else
                    {
                        HtmlDocument doc = WebClientUtil.GetHtmlDocument(isinList[0].ISINLink, 300000);
                        result = result.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in Generating new RIC    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Generate Rights ADD FM file(xls). 
        /// </summary>
        private void GenerateRightsAddFileXls()
        {
            Logger.Log("Generate Rights ADD FM File.");
            foreach (var item in raList)
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                ExcelApp excelApp = new ExcelApp(false, false);
                if (excelApp.ExcelAppInstance == null)
                {
                    string msg = "";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                try
                {
                    string fileName = "KR FM(Right Add)_" + item.RIC + " (wef " + item.AddEffectiveDate + ").xls";
                    string filePath = Path.Combine(configObj.RightsAddFM, fileName);
                    Workbook workBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                    Worksheet workSheet = workBook.Worksheets[1] as Worksheet;
                    if (workSheet == null)
                    {
                        string msg = "";
                        Logger.Log(msg, Logger.LogType.Error);
                    }

                    ((Range)workSheet.Columns["A", System.Type.Missing]).ColumnWidth = 24;
                    ((Range)workSheet.Columns["B", System.Type.Missing]).ColumnWidth = 3;
                    ((Range)workSheet.Columns["C", System.Type.Missing]).ColumnWidth = 30;
                    ((Range)workSheet.Columns["A:C", System.Type.Missing]).Font.Name = "Arial";
                    ((Range)workSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)workSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    workSheet.Cells[1, 1] = "FM Request";
                    workSheet.Cells[1, 2] = "";
                    workSheet.Cells[1, 3] = "Right RIC Add";
                    ((Range)workSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)workSheet.Cells[3, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    workSheet.Cells[3, 1] = "Effective Date";
                    workSheet.Cells[3, 2] = ":";
                    ((Range)workSheet.Cells[3, 3]).NumberFormat = "@";
                    workSheet.Cells[3, 3] = item.AddEffectiveDate;
                    ((Range)workSheet.Cells[4, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)workSheet.Cells[4, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    workSheet.Cells[4, 1] = "RIC";
                    workSheet.Cells[4, 2] = ":";
                    workSheet.Cells[4, 3] = item.RIC;
                    workSheet.Cells[5, 1] = "Currency";
                    workSheet.Cells[5, 2] = ":";
                    workSheet.Cells[5, 3] = item.Currency;
                    workSheet.Cells[6, 1] = "QA Common Name";
                    workSheet.Cells[6, 2] = ":";
                    ((Range)workSheet.Cells[6, 3]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                    workSheet.Cells[6, 3] = item.QACommonName;
                    workSheet.Cells[7, 1] = "QA Short Name";
                    workSheet.Cells[7, 2] = ":";
                    ((Range)workSheet.Cells[7, 3]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                    ((Range)workSheet.Cells[7, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                    workSheet.Cells[7, 3] = item.QAShortName;
                    workSheet.Cells[9, 1] = "Korea Code";
                    workSheet.Cells[9, 2] = ":";
                    ((Range)workSheet.Cells[9, 3]).NumberFormat = "@";
                    workSheet.Cells[9, 3] = item.KoreaCode;
                    workSheet.Cells[10, 1] = "ISIN";
                    workSheet.Cells[10, 2] = ":";
                    workSheet.Cells[10, 3] = item.ISIN;
                    workSheet.Cells[12, 1] = "Country Headquarters";
                    workSheet.Cells[12, 2] = ":";
                    workSheet.Cells[12, 3] = item.CountryHeadquarters;
                    workSheet.Cells[13, 1] = "Legal Name";
                    workSheet.Cells[13, 2] = ":";
                    workSheet.Cells[13, 3] = item.LegalName;
                    workSheet.Cells[14, 1] = "Korean Name";
                    workSheet.Cells[14, 2] = ":";
                    workSheet.Cells[14, 3] = item.KoreaName;
                    workSheet.Cells[15, 1] = "Edcoid";
                    workSheet.Cells[15, 2] = ":";
                    workSheet.Cells[15, 3] = item.Edcoid;
                    workSheet.Cells[16, 1] = "Old  MSCI";
                    workSheet.Cells[16, 2] = ":";
                    workSheet.Cells[16, 3] = item.OldMSCI;
                    workSheet.Cells[17, 1] = "RBSS";
                    workSheet.Cells[17, 2] = ":";
                    workSheet.Cells[17, 3] = "";
                    workSheet.Cells[18, 1] = "Korea Scheme";
                    workSheet.Cells[18, 2] = ":";
                    workSheet.Cells[18, 3] = "";
                    workSheet.Cells[19, 1] = "Quantity of Rights";
                    workSheet.Cells[19, 2] = ":";
                    workSheet.Cells[19, 3] = item.QuantityOfRights;
                    workSheet.Cells[20, 1] = "Strike Price ";
                    workSheet.Cells[20, 2] = ":";
                    workSheet.Cells[20, 3] = item.StrikePrice;
                    workSheet.Cells[21, 1] = "Record Type";
                    workSheet.Cells[21, 2] = ":";
                    workSheet.Cells[21, 3] = item.RecordType;
                    workSheet.Cells[22, 1] = "KOSPI Chain RIC";
                    workSheet.Cells[22, 2] = ":";
                    workSheet.Cells[22, 3] = "";
                    workSheet.Cells[23, 1] = "Position in Chain";
                    workSheet.Cells[23, 2] = ":";
                    workSheet.Cells[23, 3] = "";
                    workSheet.Cells[24, 1] = "Issue classification";
                    workSheet.Cells[24, 2] = ":";
                    workSheet.Cells[24, 3] = item.IssueClassification;
                    workSheet.Cells[25, 1] = "Lot Size";
                    workSheet.Cells[25, 2] = ":";
                    workSheet.Cells[25, 3] = item.LotSize;
                    workSheet.Cells.get_Range("A26", "C26").MergeCells = true;
                    workSheet.Cells[26, 1] = "'-------------------------------------------";

                    excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                    workBook.Save();

                    MailToSend mail = new MailToSend();
                    mail.ToReceiverList.AddRange(configObj.MailTo);
                    mail.CCReceiverList.AddRange(configObj.MailCC);
                    mail.MailSubject = Path.GetFileNameWithoutExtension(fileName);
                    mail.AttachFileList.Add(filePath);
                    mail.MailBody = "Right Add:\t\t" + item.RIC + "\r\n\r\n"
                                  + "Effective Date:\t\t" + item.AddEffectiveDate + "\r\n\r\n"
                                  + "And the estimate delist day is " + item.DropEffectiveDate + ".\r\n\r\n\r\n\r\n";
                    string signature = string.Join("\r\n", configObj.MailSignature.ToArray());
                    mail.MailBody += signature;

                    AddResult(fileName, filePath, "FM File");
                    Logger.Log("Generate Rights ADD FM file successfully. Filepath is " + filePath);
                }
                catch (Exception ex)
                {
                    string msg = "Error found in Generate_Korea_Rights ADD FM file :" + ex.StackTrace + " :\r\n " + ex.ToString();
                    Logger.Log(msg, Logger.LogType.Error);
                }
                finally
                {
                    excelApp.Dispose();
                }
            }
        }

        /// <summary>    
        /// Generate Rights DROP FM file(xls).       
        /// </summary>
        private void GenerateRightsDropFileXls()
        {
            Logger.Log("Generate Rights DROP File.");
            foreach (var item in raList)
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                ExcelApp excelApp = new ExcelApp(false, false);
                if (excelApp.ExcelAppInstance == null)
                {
                    string msg = "";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }
                try
                {
                    string fileName = "KR FM (Right DROP) Request_ " + item.RIC + "(wef " + item.DropEffectiveDate + ").xls";
                    string filePath = Path.Combine(configObj.RightsDropFM, fileName);
                    Workbook workBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                    Worksheet workSheet = workBook.Worksheets[1] as Worksheet;
                    if (workSheet == null)
                    {
                        string msg = "";
                        Logger.Log(msg, Logger.LogType.Error);
                    }

                    ((Range)workSheet.Columns["A", System.Type.Missing]).ColumnWidth = 24;
                    ((Range)workSheet.Columns["B", System.Type.Missing]).ColumnWidth = 3;
                    ((Range)workSheet.Columns["C", System.Type.Missing]).ColumnWidth = 30;
                    ((Range)workSheet.Columns["A:C", System.Type.Missing]).Font.Name = "Arial";
                    ((Range)workSheet.Columns["A", Type.Missing]).Font.Italic = System.Drawing.FontStyle.Italic;
                    ((Range)workSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)workSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    workSheet.Cells[1, 1] = "FM Request";
                    workSheet.Cells[1, 2] = "";
                    workSheet.Cells[1, 3] = "Right RIC Drop";
                    ((Range)workSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)workSheet.Cells[3, 1]).Font.Italic = System.Drawing.FontStyle.Italic;
                    ((Range)workSheet.Cells[3, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    workSheet.Cells[3, 1] = "Effective Date";
                    workSheet.Cells[3, 2] = ":";
                    workSheet.Cells[3, 3] = item.DropEffectiveDate;
                    ((Range)workSheet.Cells[4, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)workSheet.Cells[4, 1]).Font.Italic = System.Drawing.FontStyle.Italic;
                    ((Range)workSheet.Cells[4, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    workSheet.Cells[4, 1] = "RIC";
                    workSheet.Cells[4, 2] = ":";
                    workSheet.Cells[4, 3] = item.RIC;
                    workSheet.Cells[5, 1] = "ISIN";
                    workSheet.Cells[5, 2] = ":";
                    ((Range)workSheet.Cells[5, 3]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
                    workSheet.Cells[5, 3] = item.ISIN;
                    workSheet.Cells[6, 1] = "QA Short Name";
                    workSheet.Cells[6, 2] = ":";
                    ((Range)workSheet.Cells[6, 3]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                    ((Range)workSheet.Cells[6, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                    workSheet.Cells[6, 3] = item.QAShortName;
                    workSheet.Cells[7, 1] = "Legal Name";
                    workSheet.Cells[7, 2] = ":";
                    workSheet.Cells[7, 3] = item.LegalName;
                    workSheet.Cells.get_Range("A8", "C8").MergeCells = true;
                    workSheet.Cells[8, 1] = "'-------------------------------------------";

                    excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                    workBook.Save();

                    MailToSend mail = new MailToSend();
                    mail.ToReceiverList.AddRange(configObj.MailTo);
                    mail.CCReceiverList.AddRange(configObj.MailCC);
                    mail.MailSubject = Path.GetFileNameWithoutExtension(fileName);
                    mail.AttachFileList.Add(filePath);
                    mail.MailBody = "Right Drop:\t\t" + item.RIC + "\r\n\r\n"
                                  + "Effective Date:\t\t" + item.DropEffectiveDate + "\r\n\r\n\r\n\r\n";
                    string signature = string.Join("\r\n", configObj.MailSignature.ToArray());
                    mail.MailBody += signature;

                    AddResult(fileName, filePath, "FM File");
                    Logger.Log("Generate Rights DROP FM file successfully. Filepath is " + filePath);
                }
                catch (Exception ex)
                {
                    string msg = "Error found in Generate_Korea_Rights DROP FM file :\r\n" + ex.StackTrace + "      : \r\n" + ex.ToString();
                    Logger.Log(msg, Logger.LogType.Error);
                }
                finally
                {
                    excelApp.Dispose();
                }
            }
        }

        /// <summary>
        /// Generate GEDA Files. 
        /// For ADD, the name should be KR_RIGHTS_ADD_YYYYMMDD.txt (YYYYMMDD: System Date)
        /// For DROP, the name should be KR_RIGHTS_DROP_YYYYMMDD.txt (YYYYMMDD: Delisting Effective Date)
        /// </summary>
        private void GenerateGEDAFiles()
        {
            Logger.Log("Generate GEDA Files");
            List<string> addGEDATitle = new List<string>(){ "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL",
                                                            "BCAST_REF", "#INSTMOD_#ISIN", "#INSTMOD_MNEMONIC", "#INSTMOD_STRIKE_PRC", 
                                                            "#INSTMOD_TDN_SYMBOL", "EXL_NAME", "BCU" };

            Hashtable dropData = new Hashtable();
            System.Data.DataTable dt = GenerateTitle(addGEDATitle);
            for (int i = 0; i < raList.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = raList[i].RIC;
                dr[1] = raList[i].QAShortName;
                dr[2] = raList[i].RIC;
                dr[3] = raList[i].KoreaCode;
                dr[4] = raList[i].ISIN;
                dr[5] = "****";
                dr[6] = raList[i].KoreaName;
                dr[7] = raList[i].Edcoid;
                dr[8] = raList[i].ISIN;
                dr[9] = "J" + raList[i].KoreaCode;
                dr[10] = raList[i].StrikePrice;
                //dr[11] = "81";
                if (raList[i].RIC.Substring(raList[i].RIC.IndexOf(".") + 1) == "KS")
                {
                    dr[11] = raList[i].KoreaCode;
                    dr[12] = "KSE_EQB_RTS";
                    dr[13] = "KSE_EQ_RIGHTS";
                }
                else
                {
                    dr[11] = raList[i].RIC.Substring(0, 6) + "R";
                    dr[12] = "KOSDAQ_EQB_RTS";
                    dr[13] = "KOSDAQ_EQ_RIGHTS";
                }
                dt.Rows.Add(dr);

                if (dropData.Contains(raList[i].DropEffectiveDate))
                {
                    List<string> item = new List<string>();
                    item.AddRange((List<string>)dropData[raList[i].DropEffectiveDate]);
                    item.Add(raList[i].RIC);
                    dropData[raList[i].DropEffectiveDate] = item;
                }
                else
                {
                    List<string> newItem = new List<string>();
                    newItem.Add(raList[i].RIC);
                    dropData.Add(raList[i].DropEffectiveDate, newItem);
                }
            }

            WriteAddGEDAFile(dt);
            WriteDropGEDAFile(dropData);
        }

        /// <summary>
        /// Generate title for datatable using given title name list.
        /// </summary>
        /// <param name="titleName">title name list</param>
        /// <returns>datatable</returns>
        protected System.Data.DataTable GenerateTitle(List<string> titleName)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            foreach (string title in titleName)
            {
                dt.Columns.Add(title);
            }

            DataRow dr = dt.NewRow();
            for (int i = 0; i < titleName.Count; i++)
            {
                dr[i] = titleName[i];
            }
            dt.Rows.Add(dr);

            return dt;
        }

        /// <summary>
        /// Write GEDA file using the datatable.
        /// </summary>
        /// <param name="dt">data to write to txt</param>
        private void WriteAddGEDAFile(System.Data.DataTable dt)
        {
            string fileName = "KR_RIGHTS_ADD_" + DateTime.Now.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
            string filePath = Path.Combine(configObj.RightsAddGEDA, fileName);
            string[] content = new string[dt.Rows.Count];

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sb.Append(dt.Rows[i][j].ToString() + "\t");
                }
                content[i] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            WriteTxtFile(filePath, content);
            Logger.Log("Generate ADD GEDA file Successfully. Filepath is " + filePath);
            AddResult(fileName, filePath, "ADD GEDA FILE");
            AddResult(fileName, filePath, "ADD GEDA FILE");
        }

        /// <summary>
        /// Write DROP File using given data. Item with different DROP Effective dates should be seperated.
        /// </summary>
        /// <param name="dropData">data with date and RICs</param>
        private void WriteDropGEDAFile(Hashtable dropData)
        {
            foreach (DictionaryEntry de in dropData)
            {
                string fileName = "KR_DROP_" + DateTime.Parse(de.Key.ToString()).ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
                string filePath = Path.Combine(configObj.RightsDropGEDA, fileName);
                List<string> ricList = (List<string>)de.Value;
                WriteDropGEDAFileEachDay(filePath, ricList);
                Logger.Log("Generate DROP GEDA file Successfully. Filepath is " + filePath);
                AddResult(fileName, filePath, "DROP GEDA FILE");
            }

        }

        /// <summary>
        /// Write one date's Rights RIC for DROP.
        /// </summary>
        /// <param name="filePath">output file path</param>
        /// <param name="ricList">data to write to txt</param>
        private void WriteDropGEDAFileEachDay(string filePath, List<string> ricList)
        {
            List<List<string>> data = new List<List<string>>();
            List<string> title = new List<string>();
            title.Add("RIC");
            data.Add(title);
            foreach (string ric in ricList)
            {
                List<string> item = new List<string>();
                item.Add(ric);
                data.Add(item);
            }
            WriteToTxt(filePath, data, false);
        }

        /// <summary>
        /// Write TXT file using given data.
        /// </summary>
        /// <param name="fullpath">output file path</param>
        /// <param name="content">data to write</param>
        private void WriteTxtFile(string fullpath, string[] content)
        {
            if (!Directory.Exists(Path.GetDirectoryName(fullpath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(fullpath));
            }
            try
            {
                File.WriteAllLines(fullpath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                string errInfo = "Error found in writing Txt file :" + fullpath + "\r\n" + ex.ToString();
                Logger.Log(errInfo, Logger.LogType.Error);
            }
        }

        private void WriteToTxt(string filePath, List<List<string>> content, bool overWrite)
        {
            try
            {
                FileStream fs = null;
                if (overWrite)
                {
                    CreateDirectory(filePath);
                    fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                }
                else
                {
                    if (File.Exists(filePath))
                    {
                        content.RemoveAt(0);
                    }
                    else
                    {
                        CreateDirectory(filePath);
                    }
                    fs = new FileStream(filePath, FileMode.Append, FileAccess.Write);
                }
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
                for (int i = 0; i < content.Count; i++)
                {
                    for (int j = 0; j < content[i].Count; j++)
                    {
                        sw.Write(content[i][j]);
                    }
                    sw.Write("\r\n");
                }
                sw.Close();
                fs.Close();
            }
            catch
            {
                string errInfo = "Error found in writing Txt file :" + filePath;
                throw new Exception(errInfo);
            }
        }

        private void CreateDirectory(string directory)
        {
            if (Path.HasExtension(directory))
            {
                directory = Path.GetDirectoryName(directory);
            }
            if (!Directory.Exists(directory) && directory != null)
            {
                Directory.CreateDirectory(directory);
            }
        }

        private void GenerateNdaFutureDatingFile()
        {
            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "PEOFutureDating.csv";
            string filePath = Path.Combine(configObj.RightsAddNDA, fileName);
            List<List<string>> listReadFile = ReadExcelFile();

            if (listReadFile == null || listReadFile.Count <= 1)
            {
                LogMessage("no futhure dataing need to generate.");
                return;
            }

            List<List<string>> listListOutputFile = ForamtReadFile(listReadFile);

            if (listListOutputFile == null || listListOutputFile.Count <= 1)
            {
                LogMessage("formate data from KRyyyymmddQAAddRTS.csv error.");
                return;
            }

            XlsOrCsvUtil.GenerateStringCsv(filePath, listListOutputFile);
            AddResult("PEOFutureDating.csv", filePath, "CSV bulk file");
        }

        private List<List<string>> ForamtReadFile(List<List<string>> listReadFile)
        {
            List<List<string>> listListOutputFile = new List<List<string>>();

            try
            {
                List<string> listTitle = new List<string>(){
                "RIC",	
                "PROPERTY NAME",	
                "PROPERTY VALUE",	
                "EFFECTIVE FROM",	
                "EFFECTIVE TO",	
                "CHANGE OFFSET",	
                "CHANGE TRIGGER",	
                "CORAX PERMID"};
                listListOutputFile.Add(listTitle);

                List<string> listPropertyName = new List<string>(){
                "",                
                "RIC",
                "ROUND LOT SIZE",
                "TICKER SYMBOL"};
                int start = 0;
                while (start++ < listReadFile.Count - 1)
                {
                    if (listReadFile[start].Count < 13)
                        continue;

                    //if ((listReadFile[start][12] + "").Trim().Length != 0)
                    foreach (var strPro in listPropertyName)
                    {
                        if ((listReadFile[start][12] + "").Trim().Length == 0 && strPro.Equals("ROUND LOT SIZE"))
                            continue;

                        List<string> line = new List<string>();
                        line.Add(listReadFile[start][0]);//Ric
                        line.Add(strPro);//PROPERTY NAME

                        switch (strPro)//PROPERTY VALUE
                        {
                            case "":
                                line.Add("");
                                break;
                            case "RIC":
                                line.Add(listReadFile[start][0]);
                                break;
                            case "ROUND LOT SIZE":
                                line.Add(listReadFile[start][12]);
                                break;
                            case "TICKER SYMBOL":
                                line.Add(listReadFile[start][3]);
                                break;
                        }
                        //04-Aug-2014
                        //1-Jul-14
                        line.Add(ConvertDateTime(listReadFile[start][10], "dd-MMM-yyyy"));                               //"EFFECTIVE FROM",	
                        line.Add("");                               //"EFFECTIVE TO",	
                        line.Add("");                               //"CHANGE OFFSET",	
                        line.Add(strPro.Equals("") ? "" : "PEO");   //"CHANGE TRIGGER",	
                        line.Add("");                               //"CORAX PERMID"
                        listListOutputFile.Add(line);
                    }
                }

                return listListOutputFile;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return null;
        }

        private string ConvertDateTime(string str, string afterStr)
        {
            try
            {
                DateTime dt;
                int days;
                if (DateTime.TryParseExact("1-1-1900", "d-d-yyyy", new CultureInfo("en-US"), DateTimeStyles.None, out dt) && Int32.TryParse(str, out days))
                    return dt.AddDays(days - 2).ToString(afterStr);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return string.Empty;
        }

        private string ConvertDateTime(string str, string beforStr, string afterStr)
        {
            try
            {
                DateTime dt;
                if (DateTime.TryParseExact(str, beforStr, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
                    return dt.ToString(afterStr);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return string.Empty;
        }

        private List<List<string>> ReadExcelFile()
        {
            string filePath = string.Empty;

            try
            {
                filePath = Path.Combine(configObj.RightsAddNDA, "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAAddRTS.csv");
                if (!File.Exists(filePath))
                    LogMessage(string.Format("the input file: {0} is not exist,generate futhure dataing file failed. ", filePath));

                using (ExcelApp excelApp = new ExcelApp(false, false))
                    return ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath).ToList();
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return null;
        }

        /// <summary>
        /// Generate NDA file.
        /// The name should be KRYYYYMMDDQAAddRTS_KOSPI.csv YYYYMMDD is SystemDate.
        /// </summary>
        private void GenerateNDAQAFile()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            List<string> addTitle = new List<string>() { "RIC", "TAG", "BASE ASSET", "TICKER SYMBOL", "ASSET SHORT NAME","ASSET COMMON NAME", 
                                                         "TYPE", "CATEGORY", "CURRENCY", "EXCHANGE", "EQUITY FIRST TRADING DAY", 
                                                         "RETIRE DATE", "ROUND LOT SIZE", "SETTLEMENT PERIOD","PRIMARY TRADABLE MARKET QUOTE"};

            ExcelApp app = new ExcelApp(false, false);
            if (app.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            try
            {
                string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAAddRTS.csv";
                string filePath = Path.Combine(configObj.RightsAddNDA, fileName);
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;

                for (int i = 0; i < addTitle.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = addTitle[i];
                }
                for (int j = 0; j < raList.Count; j++)
                {
                    string ricSuffix = raList[j].RIC.Substring(raList[j].RIC.IndexOf(".") + 1);
                    string ricPerffix = raList[j].RIC.Substring(0, raList[j].RIC.IndexOf("."));
                    worksheet.Cells[4 * j + 2, 1] = raList[j].RIC;
                    worksheet.Cells[4 * j + 3, 1] = ricPerffix + "F." + ricSuffix;
                    //worksheet.Cells[6 * j + 4, 1] = ricPerffix + "S." + ricSuffix;
                    worksheet.Cells[4 * j + 4, 1] = ricPerffix + "stat." + ricSuffix;
                    worksheet.Cells[4 * j + 5, 1] = ricPerffix + "ta." + ricSuffix;
                    //worksheet.Cells[6 * j + 7, 1] = ricPerffix + "bl." + ricSuffix;
                    worksheet.Cells[4 * j + 2, 13] = raList[j].LotSize;
                    // worksheet.Cells[6 * j + 4, 13] = raList[j].LotSize;

                    if (ricSuffix == "KQ")
                    {
                        worksheet.Cells[4 * j + 2, 2] = "673";
                        worksheet.Cells[4 * j + 3, 2] = "64399";
                        // worksheet.Cells[6 * j + 4, 2] = "60673";
                        worksheet.Cells[4 * j + 4, 2] = "61287";
                        worksheet.Cells[4 * j + 5, 2] = "64380";
                        // worksheet.Cells[6 * j + 7, 2] = "67094";
                        for (int k = 0; k < 4; k++)
                        {
                            worksheet.Cells[4 * j + 2 + k, 10] = "KOE";
                        }
                    }
                    else
                    {
                        worksheet.Cells[4 * j + 2, 2] = "184";
                        worksheet.Cells[4 * j + 3, 2] = "64398";
                        // worksheet.Cells[6 * j + 4, 2] = "60184";
                        worksheet.Cells[4 * j + 4, 2] = "61286";
                        worksheet.Cells[4 * j + 5, 2] = "64379";
                        //worksheet.Cells[6 * j + 7, 2] = "67093";
                        for (int k = 0; k < 4; k++)
                        {
                            worksheet.Cells[4 * j + 2 + k, 10] = "KSC";
                        }
                    }
                    for (int k = 0; k < 4; k++)
                    {
                        worksheet.Cells[4 * j + 2 + k, 3] = "ISIN:" + raList[j].ISIN;
                        ((Range)worksheet.Cells[4 * j + 2 + k, 4]).NumberFormat = "@";
                        worksheet.Cells[4 * j + 2 + k, 4] = "'" + raList[j].KoreaCode;
                        worksheet.Cells[4 * j + 2 + k, 5] = raList[j].QAShortName;
                        worksheet.Cells[4 * j + 2 + k, 6] = raList[j].QACommonName;
                        worksheet.Cells[4 * j + 2 + k, 7] = "EQUITY";
                        worksheet.Cells[4 * j + 2 + k, 8] = "RTS";
                        worksheet.Cells[4 * j + 2 + k, 9] = "KRW";
                        string addDateFormat = DateTime.Parse(raList[j].AddEffectiveDate).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US"));
                        string dropDateFormat = DateTime.Parse(raList[j].DropEffectiveDate).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US"));
                        ((Range)worksheet.Cells[4 * j + 2 + k, 11]).NumberFormat = "@";
                        ((Range)worksheet.Cells[4 * j + 2 + k, 12]).NumberFormat = "@";
                        worksheet.Cells[4 * j + 2 + k, 11] = addDateFormat;
                        worksheet.Cells[4 * j + 2 + k, 12] = dropDateFormat;
                        worksheet.Cells[4 * j + 2 + k, 14] = "T+2";

                        if (k.CompareTo(0) == 0)
                            worksheet.Cells[4 * j + 2 + k, 15] = "Y";
                        else
                            worksheet.Cells[4 * j + 2 + k, 15] = string.Empty;
                    }
                }
                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbook.SaveAs(workbook.FullName, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                Logger.Log("Generate NDA QA file Successfully. Filepath is " + filePath);
                AddResult(fileName, filePath, "NDA QA File");

            }
            catch (Exception ex)
            {
                string msg = "Error found in generate NDA file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                app.Dispose();
            }
        }

        /// <summary>
        /// Generate NDA Tick and Lot Add file.
        /// </summary>
        private void GenerateNDATickLotFile()
        {
            try
            {
                List<string> tickTitle = new List<string>(){"RIC", "TICK NOT APPLICABLE", "TICK LADDER NAME", 
                                                      "TICK EFFECTIVE FROM", "TICK EFFECTIVE TO", "TICK PRICE INDICATOR" };
                List<string> lotTitle = new List<string>(){"RIC", "LOT NOT APPLICABLE", "LOT LADDER NAME", 
                                                      "LOT EFFECTIVE FROM", "LOT EFFECTIVE TO", "LOT PRICE INDICATOR" };
                string today = DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US"));
                string filePathTick = Path.Combine(configObj.RightsAddNDA, "TickAdd_RTS_" + today + ".csv");
                string filePathLot = Path.Combine(configObj.RightsAddNDA, "LotAdd_RTS_" + today + ".csv");
                List<List<string>> tickContent = new List<List<string>>();
                List<List<string>> lotContent = new List<List<string>>();
                foreach (RightsTemplate item in raList)
                {
                    string ricSuffix = item.RIC.Split('.')[1];
                    List<string> tickRecord = new List<string>();
                    List<string> lotRecord = new List<string>();
                    //List<string> 
                    tickRecord.Add(item.RIC);
                    lotRecord.Add(item.RIC);
                    tickRecord.Add("N");
                    lotRecord.Add("N");
                    if (ricSuffix == "KQ")
                    {
                        tickRecord.Add("TICK_LADDER_KOE_1");
                    }
                    else
                    {
                        tickRecord.Add("TICK_LADDER_KSC_1");
                    }
                    lotRecord.Add("LOT_LADDER_EQTY_<1>");
                    string effectiveDate = DateTime.Parse(item.AddEffectiveDate).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US"));
                    tickRecord.Add(effectiveDate);
                    lotRecord.Add(effectiveDate);
                    tickRecord.Add("");
                    lotRecord.Add("");
                    tickRecord.Add("ORDER");
                    lotRecord.Add("CLOSE");
                    tickContent.Add(tickRecord);
                    lotContent.Add(lotRecord);
                }
                FileUtil.WriteOutputFile(filePathTick, tickContent, tickTitle, WriteMode.Overwrite);
                FileUtil.WriteOutputFile(filePathLot, lotContent, lotTitle, WriteMode.Overwrite);
                AddResult(Path.GetFileName(filePathTick), filePathTick, "NDA Tick Add File");
                AddResult(Path.GetFileName(filePathLot), filePathLot, "NDA Lot Add File");

                Logger.Log("Generate NDA Tick and Lot add file files successfully.", Logger.LogType.Info);
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in generating NDA Tick and Lot file. \r\n" + ex.ToString(), Logger.LogType.Error);
            }
        }
    }
}
