using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Ric.Db.Manager;
using Selenium;
using System.ComponentModel;
using System.Drawing.Design;
using System.Data;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
{    
    public class WarrantChange : GeneratorBase
    {       
        private List<DateTime> holidayList = null;       
        private Hashtable kqPrice = new Hashtable();
        private Hashtable filesResultList = new Hashtable();
        private List<CompanyWarrantTemplate> wcList = new List<CompanyWarrantTemplate>(); 
        private KOREACompanyWarrantChangeGeneratorConfig configObj = null;
        private string folderEMA = "";
        List<CompanyWarrantTemplate> priceChange = new List<CompanyWarrantTemplate>();
        List<CompanyWarrantTemplate> priceNotChange = new List<CompanyWarrantTemplate>();
        
        private const string ETI_KOREA_COMPANYWARRANT_TABLE_NAME = "ETI_Korea_CompanyWarrant";

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREACompanyWarrantChangeGeneratorConfig;
            holidayList = HolidayManager.SelectHoliday(MarketId);
            if (holidayList == null)
            {
                holidayList = new List<DateTime>();
            }
        }       

        protected override void Start()
        {          
            GrabDataFromKindWebpage();
            if (wcList.Count > 0)
            {
                GetPriceList();
                FormatData();
                GenerateFiles();
            }
            else
            {
                Logger.Log("There is no Company Warrant CHANGE data grabbed");
            }
            AddResult("LOG file",Logger.FilePath,"LOG file");            
        }
      
        /// <summary>
        /// Grab warrant change data from website.
        /// </summary>
        private void GrabDataFromKindWebpage()
        {
            try
            {
                HtmlDocument htc = GetAnnouncementDocument();

                if (htc == null)
                {
                    return;
                }

                HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes(".//dl/dt");
                HtmlNodeCollection ddCollections = htc.DocumentNode.SelectNodes("//dl/dd");

               // HtmlNodeCollection nodeCollections = GetHtmlNodeCollection();
                if (nodeCollections == null || nodeCollections.Count == 0)
                {
                    Logger.Log("There is no Company Warrant Change data grabbed.");
                    return;
                }
                for (int i = nodeCollections.Count - 1; i >= 0; i--)
                {
                    HtmlNode ddNode = ddCollections[i].SelectSingleNode(".//span");

                    HtmlNode dtNode = nodeCollections[i];               
                    HtmlNode titleNode = dtNode.SelectSingleNode(".//span/a");
                    if (titleNode == null)
                    {
                        continue;
                    }
                    string title = titleNode.InnerText.Trim().ToString();
                    if (title.Contains("신주인수권증권 변경상장"))
                    {
                        HtmlNode nodeDate = dtNode.SelectSingleNode("./em");
                        if (nodeDate != null)
                        {
                            string startDate = configObj.StartDate;
                            DateTime anouncementDate = new DateTime();
                            anouncementDate = DateTime.Parse(nodeDate.InnerText.Trim(), new CultureInfo("en-US"));
                            if (anouncementDate < DateTime.Parse(startDate))
                            {
                                continue;
                            }
                        }

                        string type = GetDutyCode(ddNode);                       
                        
                        Object[] obj = ReturnTargetObjArray(titleNode, type);
                        if (type == "KQ")
                        {
                            GrabKQDataAction(obj);
                        }
                        else if(type == "KS")
                        {
                            GrabKSDataAction(obj);
                        }
                    }                
                }               
            }
            catch (Exception ex)
            {
                string msg = "Error found in Grabbing Data From KindWebpage. \r\n" + ex.ToString();
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
        ///  Get records search with startDate and endDate.
        /// </summary>
        /// <returns>Html nodes</returns>
        private HtmlNodeCollection GetHtmlNodeCollection()
        {
            HtmlNodeCollection nodeCollections = null;
            try
            {
                string today = DateTime.Today.ToString("yyyy-MM-dd");
                string startDate = configObj.StartDate.Trim();
                string endDate = configObj.EndDate.Trim();
                if (string.IsNullOrEmpty(startDate))
                    startDate = today;
                if (string.IsNullOrEmpty(endDate))
                    endDate = today;
                string dataStartDate = (DateTime.Parse(endDate).AddMonths(-2)).ToString("yyyy-MM-dd");
                string uri = string.Format("http://kind.krx.co.kr/disclosure/searchtotalinfo.do");
                string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EA%B6%8C&repIsuSrtCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=300&scn=mktact&srchFd=2&kwd=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EA%B6%8C&fromData={0}&toData={1}", dataStartDate, endDate);
                string pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, postData);
                HtmlDocument htc = new HtmlDocument();
                if (!string.IsNullOrEmpty(pageSource))
                {
                    htc.LoadHtml(pageSource);
                }
                if (htc != null)
                {
                    nodeCollections = htc.DocumentNode.SelectNodes("//dl/dt");
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetHtmlNodeCollection()     : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return nodeCollections;
        }

        /// <summary>
        ///  Get records search with startDate and endDate.
        /// </summary>
        /// <returns>Html nodes</returns>
        private HtmlDocument GetAnnouncementDocument()
        {
            HtmlDocument htc = null;
            try
            {
                string today = DateTime.Today.ToString("yyyy-MM-dd");
                string startDate = configObj.StartDate.Trim();
                string endDate = configObj.EndDate.Trim();
                if (string.IsNullOrEmpty(startDate))
                    startDate = today;
                if (string.IsNullOrEmpty(endDate))
                    endDate = today;
                string dataStartDate = (DateTime.Parse(endDate).AddMonths(-2)).ToString("yyyy-MM-dd");
                string uri = string.Format("http://kind.krx.co.kr/disclosure/searchtotalinfo.do");
                string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EA%B6%8C&repIsuSrtCd=&isurCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=300&scn=mktact&srchFd=2&kwd=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EA%B6%8C&fromData={0}&toData={1}", dataStartDate, endDate);
                string pageSource = null;
                int retry = 3;

                while (string.IsNullOrEmpty(pageSource) && retry-- > 0)
                {
                    try
                    {
                        pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, postData);
                    }
                    catch
                    {
                        System.Threading.Thread.Sleep(3000);
                    }
                }
                
                htc = new HtmlDocument();
                if (!string.IsNullOrEmpty(pageSource))
                {
                    htc.LoadHtml(pageSource);
                }              
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetAnnouncementDocument()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return htc;
        }

        /// <summary>
        /// Get the detail document of warrants.
        /// </summary>
        /// <param name="titleNode">Anouncement title</param>
        /// <param name="type">warrant type</param>
        /// <returns>warrant ticker and html document</returns>
        private Object[] ReturnTargetObjArray(HtmlNode titleNode, string type)
        {
            Object[] obj = new Object[2];
            try
            {
                string pageSource = null;
                string attribute = titleNode.Attributes["onclick"].Value.Trim().ToString();
                char[] trimChar = new char[] { ' ', '\'', ',', '(', ')' };
                if (!string.IsNullOrEmpty(attribute))
                {
                    attribute = attribute.Split('(')[1].Split(')')[0].Trim(trimChar).ToString();
                    string param1 = attribute.Split(',')[0].Trim(trimChar).ToString();
                    string param2 = attribute.Split(',')[1].Trim(trimChar).ToString();
                    string url = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno={1}&viewerhost=&viewerport=", param1, param2);
                    while (pageSource == null)
                    {
                        try
                        {
                            pageSource = WebClientUtil.GetDynamicPageSource(url, 180000, null);
                        }
                        catch
                        {
                            System.Threading.Thread.Sleep(5000);
                        }
                    }
                    HtmlDocument doc = new HtmlDocument();
                    if (!string.IsNullOrEmpty(pageSource))
                    {
                        doc.LoadHtml(pageSource);
                    }                       
                    string ticker = string.Empty;
                    if (doc != null)
                    {
                        ticker = doc.DocumentNode.SelectSingleNode(".//header/h1").InnerText.Trim().ToString();
                        if (!string.IsNullOrEmpty(ticker))
                        {
                            Match m = Regex.Match(ticker, @"\(([0-9a-zA-Z]{6})\)");
                            if (m == null)
                            {
                                string msg = "Cannot get ticker numbers in ." + ticker;
                                Logger.Log(msg, Logger.LogType.Error);
                                return null;
                            }
                            ticker = m.Groups[1].Value;
                        }
                    }
                    param1 = param1.Insert(4, "/").Insert(7, "/").Insert(10, "/");
                    string param3 = type.Equals("KS") ? "68151" : (type.Equals("KQ") ? "70921" : null);
                    url = string.Format("http://kind.krx.co.kr/external/{0}/{1}/{2}.htm", param1, param2, param3);
                    string pageSourceEx = null;
                    while (pageSourceEx == null)
                    {
                        try
                        {
                            pageSourceEx = WebClientUtil.GetDynamicPageSource(url, 180000, null);
                        }
                        catch
                        {
                            System.Threading.Thread.Sleep(5000);
                        }
                    }                    
                    if (!string.IsNullOrEmpty(pageSourceEx))
                    {
                        doc = new HtmlDocument();
                        doc.LoadHtml(pageSourceEx);
                    }
                    obj[0] = ticker;
                    obj[1] = doc;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in ReturnTargetObjArray()      : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return obj;
        }

        /// <summary>
        /// Grab data for KS for an item of company warrant list.
        /// </summary>
        /// <param name="obj">document and ticker</param>
        private void GrabKSDataAction(Object[] obj)
        {
            try
            {
                HtmlDocument doc = obj[1] as HtmlDocument;
                if (doc != null)
                {
                    string strPre = doc.DocumentNode.SelectSingleNode(".//pre").InnerText.Trim().ToString();
                    if (!string.IsNullOrEmpty(strPre))
                    {
                        string[] preArr = strPre.Split('◎');
                        if (preArr.Length > 1)
                        {
                            for (var i = 2; i < preArr.Length; i++)
                            {
                                strPre = preArr[i].ToString();
                                CompanyWarrantTemplate cw = new CompanyWarrantTemplate();
                                int strOnePos = strPre.IndexOf("①");//상장종목
                                int strTwoPos = strPre.IndexOf("②");//상장신주인수권                               
                                int strThreePos = strPre.IndexOf("③");//변경후 신주인수권
                                int strFourPos = strPre.IndexOf("④");//변경상장일
                                int strFivePos = strPre.IndexOf("⑤");//변경사유
                                int strSixPos = strPre.IndexOf("⑥");//변경후 신주인수권
                                int strSevenPos = strPre.IndexOf("⑦");//변경후 신주인수권                    
                                int strEightPos = strPre.IndexOf("⑧");//코드

                                string koreaName = strPre.Substring(strOnePos, (strTwoPos - strOnePos)).Trim().ToString();    //split by ':'[1]
                                koreaName = koreaName.Split(':')[1].Trim().ToString();
                                if (koreaName.Contains("(주)"))
                                    koreaName = koreaName.Replace("(주)", "");
                                string quantity = strPre.Substring(strTwoPos, (strThreePos - strTwoPos)).Trim().ToString();  //split by  '\n'[1]
                                quantity = quantity.Split('\n')[1].Trim().ToString();
                                string effectiveDate = strPre.Substring(strFourPos, (strFivePos - strFourPos)).Trim().ToString();     //split by ':'[1]
                                effectiveDate = effectiveDate.Split(':')[1].Trim().ToString();
                                string exercisePrice = strPre.Substring(strSixPos, (strSevenPos - strSixPos)).Trim().ToString().Replace("원", "").Replace(",", "");     //split by ':'[1]
                                exercisePrice = exercisePrice.Split(':')[1].Trim().ToString();
                                string expiryDate = strPre.Substring(strSevenPos, (strEightPos - strSevenPos)).Trim().ToString();             //split by ':'[1]
                                expiryDate = expiryDate.Split(':')[1].Trim().ToString();
                                string isinKoreaCode = strPre.Substring((strEightPos + 8)).Trim().ToString();
                                string ticker = obj[0].ToString();
                                string ric = ticker + "W.KS";


                                char[] trimChar = new char[] { ' ', ',', '▶', '\n', ')', ':' };
                                string tempQuantity = quantity.Contains("증권") ? quantity.Replace("증권", "").Trim(trimChar) : quantity.Trim(trimChar);
                                string oldQuantity = tempQuantity.Split('→')[0].Trim(trimChar).ToString();
                                oldQuantity = oldQuantity.Contains(",") ? oldQuantity.Replace(",", "") : oldQuantity;
                                string quantityOfWarrant = GetQuantityOfWarrant(tempQuantity);
                                quantityOfWarrant = quantityOfWarrant.Contains(",") ? quantityOfWarrant.Replace(",", "") : quantityOfWarrant;

                                string isin = string.Empty;
                                string koreaCode = string.Empty;
                                if (isinKoreaCode.Contains("표준코드") || isinKoreaCode.Contains("단축코드"))
                                {
                                    int isinPos = isinKoreaCode.IndexOf("표준코드") + ("표준코드".Length);
                                    int koreacodePos = isinKoreaCode.IndexOf("단축코드");
                                    isin = isinKoreaCode.Substring((isinPos), (koreacodePos - isinPos)).Trim(new char[] { '(', ':', ' ', ',' }).ToString();
                                    koreaCode = isinKoreaCode.Substring((koreacodePos + "단축코드".Length)).Trim(new char[] { ')', ':', ' ', ',' }).ToString();
                                }

                                string fromDate = string.Empty;
                                string endDate = string.Empty;
                                if (expiryDate.Length > 22)
                                {
                                    fromDate = expiryDate.Split('~')[0].Trim().ToString();
                                    endDate = expiryDate.Split('~')[1].Trim().ToString();
                                    fromDate = Convert.ToDateTime(fromDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                                    endDate = Convert.ToDateTime(endDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                                }
                                cw.Ticker = ticker;
                                cw.RIC = ric;
                                cw.OldQuantity = oldQuantity;
                                cw.QuantityOfWarrants = quantityOfWarrant.Trim();
                                cw.ISIN = isin.Trim();
                                cw.KoreaCode = koreaCode.Replace("J","").Trim();
                                cw.ExercisePeriod = "From " + fromDate + " to " + endDate;
                                cw.ExpiryDate = Convert.ToDateTime(endDate).ToString("yyyy-MM-dd");
                                cw.EffectiveDate = Convert.ToDateTime(effectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                                cw.ExercisePrice = exercisePrice.Trim();
                                cw.KoreanName = koreaName;
                                wcList.Add(cw);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in Grab KS DataAction()    : \r\n" + ex.ToString();
                Logger.Log(msg);
            }
        }

        /// <summary>
        /// Get quantity of warrant in KS txt document.
        /// </summary>
        /// <param name="rawSource">quantity relative text</param>
        /// <returns>warrant's quantity</returns>
        private string GetQuantityOfWarrant(string rawSource)
        {
            string strQualityOfWarrant = string.Empty;
            string temp = rawSource.Split('→')[1].Trim(' ');
            foreach (char c in temp)
            {
                if ((c >= '0' && c <= '9') || c == ',')
                {
                    strQualityOfWarrant += c;
                }
                else
                {
                    break;
                }
            }
            return strQualityOfWarrant;
        }

        /// <summary>
        /// Grab data for KQ for an item of company warrant list.
        /// </summary>
        /// <param name="obj">document and ticker</param>
        private void GrabKQDataAction(Object[] obj)
        {
            try
            {
                HtmlDocument doc = obj[1] as HtmlDocument;
                if (doc != null)
                {
                    CompanyWarrantTemplate cw = new CompanyWarrantTemplate();
                    HtmlNodeCollection tableNodes = doc.DocumentNode.SelectNodes("//table");
                    if (tableNodes.Count > 0)
                    {
                        foreach (var item in tableNodes)
                        {
                            string strIsin = item.SelectSingleNode(".//tr[2]/td[2]").InnerText.Trim().ToString();
                            string strCode = item.SelectSingleNode(".//tr[2]/td[3]").InnerText.Trim().ToString();

                            if (strIsin.Equals("표준코드") && strCode.Equals("단축코드"))
                            {
                                string koreanName = item.SelectSingleNode(".//tr[1]/td[2]").InnerText.Trim().ToString();
                                string isin = item.SelectSingleNode(".//tr[3]/td[1]").InnerText.Trim().ToString();
                                string koreaCode = item.SelectSingleNode(".//tr[3]/td[2]").InnerText.Trim().ToString();
                                string quantityOfWarrant = item.SelectSingleNode(".//tr[5]/td[2]").InnerText.Trim().ToString();
                                string effectiveDate = item.SelectSingleNode(".//tr[8]/td[2]").InnerText.Trim().ToString();
                                string kqPre = item.SelectSingleNode(".//tr[10]/td[2]").InnerText.Trim().ToString();

                                koreanName = koreanName.Contains("(주)") ? koreanName.Replace("(주)", "") : koreanName;
                                string ticker = obj[0].ToString();
                                string strRic = ticker + "W.KQ";
                                string previousExercisePrice = string.Empty;
                                string exercisePrice = string.Empty;
                                if (kqPre.Contains("행사가"))
                                {
                                    int strPos = kqPre.IndexOf("행사가") + "행사가".Length;
                                    string strExercise = kqPre.Substring(strPos).Trim(new char[] { ')', ' ', ':' }).ToString();
                                    previousExercisePrice = Regex.Split(strExercise, "->", RegexOptions.IgnoreCase)[0].Trim().ToString().Replace("원", "");
                                    exercisePrice = Regex.Split(strExercise, "->", RegexOptions.IgnoreCase)[1].Trim().ToString().Replace("원", "");
                                }

                                cw.Ticker = ticker;
                                cw.RIC = strRic;
                                cw.KoreanName = koreanName;
                                cw.ISIN = isin;
                                cw.KoreaCode = koreaCode.Replace("J","").Trim();
                                cw.QuantityOfWarrants = quantityOfWarrant;
                                cw.EffectiveDate = Convert.ToDateTime(effectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                                cw.ExercisePrice = exercisePrice;
                                cw.PreviousExercisePrice = previousExercisePrice;
                                wcList.Add(cw);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabKQDataAction()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Get all the change infos include korean name and price. 
        /// </summary>
        private void GetPriceList()
        {
            try
            {
                string url = string.Format("http://www.krx.co.kr/por_kor/corelogic/process/ptc/etc_l_003.xhtml?data-only=true");
                string postData = "searchBtn=&searchBtn2=%EC%A1%B0%ED%9A%8C&_=";
                string pageSource = WebClientUtil.GetDynamicPageSource(url, 300000, postData);
                if (pageSource != null)
                {
                    HtmlAgilityPack.HtmlDocument htc = new HtmlAgilityPack.HtmlDocument();
                    htc.LoadHtml(pageSource);
                    HtmlNodeCollection trs = htc.DocumentNode.SelectNodes("//tr");
                    foreach (HtmlNode tr in trs)
                    {
                        try
                        {
                            string koreanName = tr.SelectSingleNode("./td[1]").InnerText.Trim().Replace(" ", "");
                            string exercisePrice = tr.SelectSingleNode("./td[4]").InnerText.Trim().Replace(",", "");
                            kqPrice.Add(koreanName, exercisePrice);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in Getting price list from KRX website." + ex.ToString(), Logger.LogType.Error);
            }
        }

        /// <summary>
        /// For warrant whose price is null, get the price from website.
        /// Find the warrant infos in DB, check if the quantity of warrant or exercise price changed.
        /// </summary>
        private void FormatData()
        {
            foreach (CompanyWarrantTemplate item in wcList)
            {
                try
                {
                    item.ChangeItems = new List<string>();
                    if (item.ExercisePrice == "")
                    {
                        item.ExercisePrice = GetKQPriceByEdcoid(item.Edcoid, item.KoreanName);
                    }
                    string condition = " where ISIN = '" + item.ISIN.Trim() + "'";
                    System.Data.DataTable dt = ManagerBase.Select(ETI_KOREA_COMPANYWARRANT_TABLE_NAME, new string[] { "*" }, condition);
                    if (dt == null || dt.Rows.Count == 0)
                    {
                        Logger.Log("Error found in selecting data from DB. Check if there is error in DB or the table of company warrant!", Logger.LogType.Error);
                        continue;
                    }
                    DataRow row = dt.Rows[0];
                    item.RIC = Convert.ToString(row["RIC"]).Trim();
                    item.OldQuantity = Convert.ToString(row["QuantityOfWarrants"]).Trim();
                    item.PreviousExercisePrice = Convert.ToString(row["ExercisePrice"]).Trim();
                    item.QACommonName = Convert.ToString(row["QACommonName"]).Trim();
                    item.ConversionRatio = Convert.ToString(row["ConversionRatio"]).Trim();
                    item.QAShortName = Convert.ToString(row["QAShortName"]).Trim();
                    item.ExercisePeriod = Convert.ToString(row["ExercisePeriod"]).Trim();
                    item.ExpiryDate = Convert.ToString(row["ExpiryDate"]).Trim();
                    item.LegalName = Convert.ToString(row["LegalName"]).Trim();
                    item.Edcoid = Convert.ToString(row["Edcoid"]).Trim();
                    string commonNameTemp = Convert.ToString(row["ForIACommonName"]).Trim();
                    string expiryDate = DateTime.Parse(item.ExpiryDate).ToString("ddMMMyy", new CultureInfo("en-US"));
                    item.ForIACommonName = commonNameTemp + " Call " + item.ExercisePrice + " KRW " + commonNameTemp + " " + expiryDate;
                    item.Currency = "KRW";
                    item.CountryHeadquarters = "KOR";
                    item.WarrantStyle = "American Style";
                    item.RecordType = "97";
                    item.IssueClassification = "WT";
                    item.SettlementType = "Cash";
                    item.LotSize = "1";
                    string quantity = item.QuantityOfWarrants.Replace(",", "").Trim();
                    row["QuantityOfWarrants"] = quantity;

                    row["UpdateDateChange"] = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
                    row["EffectiveDateChange"] = item.EffectiveDate;

                    //Price Change--QACommonName ConvertionRatio ExercisePrice
                    if (item.ExercisePrice != item.PreviousExercisePrice)
                    {
                        item.QACommonName = item.QACommonName.Replace(item.PreviousExercisePrice, item.ExercisePrice);
                        GenerateNewConvertionRatio(item);
                        row["ExercisePrice"] = item.ExercisePrice;
                        row["ConversionRatio"] = item.ConversionRatio;
                        row["QACommonName"] = item.QACommonName;
                        row["ChangeType"] = "1";
                        item.ChangeItems.Add("12");
                        item.ChangeItems.Add("6");
                        item.ChangeItems.Add("8");
                        item.ChangeItems.Add("11");
                        priceChange.Add(item);
                    }
                    //only Quantity Change 
                    else if (quantity != item.OldQuantity.Replace(",","").Trim())
                    {
                        item.ChangeItems.Add("12");
                        row["ChangeType"] = "0";
                        priceNotChange.Add(item);
                    }
                    else
                    {
                        string msg = "Both quantity and price has no changes. Please check the infomation. RIC: " + item.RIC;
                        Logger.Log(msg, Logger.LogType.Warning);
                        continue;
                    }
                    if (dt.Rows.Count > 1)
                    {
                        string warnMsg = "Found mutilple records for RIC: " + item.RIC + "\t ISIN:" + item.ISIN;
                        Logger.Log(warnMsg, Logger.LogType.Warning);
                    }
                    ManagerBase.UpdateDbTable(dt, ETI_KOREA_COMPANYWARRANT_TABLE_NAME);
                    Logger.Log("1 record updated. RIC =" + item.RIC);
                }
                catch (Exception ex)
                {
                    string msg = "Error found in formatting warrant data for RIC: " + item.RIC + "\r\n     " + ex.ToString();
                    Logger.Log(msg, Logger.LogType.Error);
                }
            }
        }

        /// <summary>
        /// Use the korean name to find the same item in kqPrice(HashTable)
        /// if not exists, use the company name from website to search. 
        /// else, foreach the hashtable to check if contain part of the name.
        /// Get the price of this item.
        /// </summary>
        /// <param name="ticker"></param>
        /// <param name="kname"></param>
        /// <returns></returns>
        private string GetKQPriceByEdcoid(string ticker, string kname)
        {
            string price = "";
            kname = kname.Replace(" ", "").Trim();
            if (kqPrice.Contains(kname))
            {
                price = kqPrice[kname].ToString().Trim();
                kqPrice.Remove(kname);
            }
            else
            {
                string url1 = string.Format("http://kind.krx.co.kr/common/corpList.do?dummytf_for_preventing_error=1&method=searchCorpList&pageIndex=1&beginIndex=&currentPageSize=10&delistFlag=Y&konexFlag=Y&comnm_tffield=AKCKwd&corpName={0}&corpNameTmp={0}&marketType=all", ticker);
                string ex = "<a id=\"CorpInfo\".*?>(?<name>.*?)</a>";
                string pageSource1 = WebClientUtil.GetDynamicPageSource(url1, 300000, null);
                Regex r = new Regex(ex);
                Match m = r.Match(pageSource1);
                string name = m.Groups["name"].Value.Trim();
                name = Regex.Replace(name, @"</?img[^>]*>", "", RegexOptions.IgnoreCase).Trim();
                name = name.Replace(" ", "");

                r = new Regex("\\d+WR$");
                m = r.Match(kname);
                string suffixName = m.Groups[0].Value.Trim();
                name += suffixName;
                if (kqPrice.Contains(name))
                {
                    price = kqPrice[name].ToString().Trim();
                    kqPrice.Remove(name);
                }
                else
                {
                    foreach (DictionaryEntry de in kqPrice)
                    {
                        name = name.Replace(suffixName, "");
                        string listName = de.Key.ToString().Trim();
                        if (listName.Contains(suffixName))
                        {
                            listName = listName.Replace(suffixName, "");
                            if (listName.Contains(name) || name.Contains(listName))
                            {
                                price = de.Value.ToString().Trim();
                                break;
                            }
                        }
                    }
                    if (price == "")
                    {
                        string msg = "Can not find price for " + kname;
                        Logger.Log(msg, Logger.LogType.Error);
                        price = "******";
                    }
                }
            }

            return price;
        }

        /// <summary>
        /// For price change. the ConvertionRatio changes as well.
        /// Forluma:oldPrice * oldConvertionRatio / newPrice
        /// </summary>
        /// <param name="item">new ConvertionRatio</param>
        private void GenerateNewConvertionRatio(CompanyWarrantTemplate item)
        {
            try
            {
                string oldRatio = item.ConversionRatio;
                double newPrice = Convert.ToDouble(item.ExercisePrice);
                double oldPrice = Convert.ToDouble(item.PreviousExercisePrice);

                if (oldRatio.Contains("%"))
                {
                    oldRatio = oldRatio.Replace("%", "");
                    double tempRatio = Convert.ToDouble(oldRatio) / 100;
                    string result = ((oldPrice * tempRatio) / newPrice).ToString("##.#####%").TrimEnd(new char[] { '0', ' ' });
                    item.ConversionRatio = result;
                }
                else
                {
                    double ratio = Convert.ToDouble(oldRatio) * oldPrice / newPrice;
                    string result = ratio.ToString("##.#####%").TrimEnd(new char[] { '0', ' ' });
                    item.ConversionRatio = result;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in Generate New ConvertionRatio for RIC : " + item.RIC + " \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }

        }

        /// <summary>
        /// For Quantity Change. Generate FM and EMA files 
        /// For Price&Quantity Change. Generate FM GEDA NDA EMA files.
        /// </summary>
        private void GenerateFiles()
        {
            GenerateFMFiles();
            InitialEMAFolder();
            foreach (CompanyWarrantTemplate item in priceNotChange)
            {
                CreateEMAQUAFile(item);
            }
            if (priceChange.Count > 0)
            {
                string folderGEDA = configObj.GEDA;
                string folderNDA = configObj.NDA;
                CreateDirectory(folderGEDA);
                CreateDirectory(folderNDA);
                foreach (CompanyWarrantTemplate priceItem in priceChange)
                {
                    CreateGEDAFile(priceItem, folderGEDA);
                    CreateNDAQAFile(priceItem, folderNDA);
                    CreateNDAIAFile(priceItem, folderNDA);
                    CreateEMAQUAFile(priceItem);
                    CreateEMAPRCFile(priceItem);
                }
            }
            if (filesResultList.Count > 0)
            {
                SetTaskResultList();
            }
        }

        private void InitialEMAFolder()
        {
            folderEMA = ConfigureOperator.GetEmaFileSaveDir();
            if (string.IsNullOrEmpty(folderEMA))
            {
                Logger.Log("Can not get EMA File folder. Please check the config in DB. The back up files is in folder FM\\EMA_FILE");               
                folderEMA = Path.Combine(configObj.FM, "EMA_FILE");
            }            
            CreateDirectory(folderEMA);
        }

        /// <summary>
        /// If file folder not exsits. Create the folder.
        /// </summary>
        /// <param name="directory"></param>
        private void CreateDirectory(string directory)
        {
            if (!Directory.Exists(directory) && directory != null)
            {
                Directory.CreateDirectory(directory);
            }
        }

        /// <summary>
        /// Generate Company Warrant Change FM file.
        /// </summary>
        private void GenerateFMFiles()
        {
            Logger.Log("Generate FM files."); 
            foreach (CompanyWarrantTemplate item in wcList)
            {
                string filePath = configObj.FM;
                string fileName = "";
                string mailFirstLine = "";
                DateTime effectiveDate = DateTime.ParseExact(item.EffectiveDate, "yyyy-MMM-dd", new CultureInfo("en-US"));
                DateTime dayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
                ExcelApp excelApp = new ExcelApp(false, false);
                if (excelApp.ExcelAppInstance == null)
                {
                    string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }
                try
                {                  
                    if (item.ChangeItems.Count > 1)
                    {
                        fileName = "KR FM(Company Warrant Change)Request_Warrant_" + item.RIC + "(wef " + item.EffectiveDate + ").xls";
                        mailFirstLine = "Company Warrant Change";
                    }
                    else if (item.ChangeItems.Count == 1)
                    {
                        fileName = "KR FM(Quantity Change)Request_Warrant_" + item.RIC + "(wef " + item.EffectiveDate + ").xls";
                        mailFirstLine = "Quantity of Warrant Change";
                    }
                    else
                    {
                        fileName = "KR FM(No Change)Request_Warrant_" + item.RIC + "(wef " + item.EffectiveDate + ").xls";
                        mailFirstLine = "No Change(Warning)";
                    }
                    filePath = Path.Combine(filePath, fileName);
                    Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                    if (wSheet == null)
                    {
                        string msg = "Worksheet could not be started. Check that your office installation and project reference are correct!";
                        Logger.Log(msg, Logger.LogType.Error);
                        return;
                    }
                    ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 28;
                    ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 3;
                    ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 45;
                    ((Range)wSheet.Columns["A:C", System.Type.Missing]).Font.Name = "Arial";
                    ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    wSheet.Cells[1, 1] = "FM Request";
                    wSheet.Cells[1, 2] = " ";
                    wSheet.Cells[1, 3] = "Company Warrant Change";
                    CompanyWarrantTemplate cwTemp = item as CompanyWarrantTemplate;
                    ((Range)wSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)wSheet.Cells[3, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    wSheet.Cells[3, 1] = "Effective Date";
                    wSheet.Cells[3, 2] = ":";
                    ((Range)wSheet.Cells[3, 3]).NumberFormat = "@";
                    wSheet.Cells[3, 3] = cwTemp.EffectiveDate;
                    ((Range)wSheet.Cells[4, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)wSheet.Cells[4, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    wSheet.Cells[4, 1] = "RIC";
                    wSheet.Cells[4, 2] = ":";
                    wSheet.Cells[4, 3] = cwTemp.RIC;
                    wSheet.Cells[5, 1] = "Currency";
                    wSheet.Cells[5, 2] = ":";
                    wSheet.Cells[5, 3] = cwTemp.Currency;
                    wSheet.Cells[6, 1] = "QA Common Name";
                    wSheet.Cells[6, 2] = ":";
                    ((Range)wSheet.Cells[6, 3]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                    wSheet.Cells[6, 3] = cwTemp.QACommonName;
                    wSheet.Cells[7, 1] = "QA Short Name";
                    wSheet.Cells[7, 2] = ":";
                    ((Range)wSheet.Cells[7, 3]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                    wSheet.Cells[7, 3] = cwTemp.QAShortName;
                    wSheet.Cells[8, 1] = "Conversion Ratio";
                    wSheet.Cells[8, 2] = ":";
                    ((Range)wSheet.Cells[8, 3]).NumberFormat = "@";
                    wSheet.Cells[8, 3] = cwTemp.ConversionRatio;
                    wSheet.Cells[9, 1] = "Korea Code";
                    wSheet.Cells[9, 2] = ":";
                    ((Range)wSheet.Cells[9, 3]).NumberFormat = "@";
                    wSheet.Cells[9, 3] = cwTemp.KoreaCode.Trim().ToString();
                    wSheet.Cells[10, 1] = "ISIN";
                    wSheet.Cells[10, 2] = ":";
                    ((Range)wSheet.Cells[10, 3]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
                    wSheet.Cells[10, 3] = cwTemp.ISIN;
                    wSheet.Cells[11, 1] = "Exercise Price";
                    wSheet.Cells[11, 2] = ":";
                    wSheet.Cells[11, 3] = cwTemp.ExercisePrice;
                    wSheet.Cells[12, 1] = "Quantity of Warrants";
                    wSheet.Cells[12, 2] = ":";
                    wSheet.Cells[12, 3] = cwTemp.QuantityOfWarrants;
                    wSheet.Cells[13, 1] = "Exercise Period";
                    wSheet.Cells[13, 2] = ":";
                    wSheet.Cells[13, 3] = cwTemp.ExercisePeriod;
                    wSheet.Cells[14, 1] = "Expiry Date";
                    wSheet.Cells[14, 2] = ":";
                    wSheet.Cells[14, 3] = Convert.ToDateTime(cwTemp.ExpiryDate).ToString("yyyy-MM-dd");
                    wSheet.Cells[15, 1] = "Country Headquarters";
                    wSheet.Cells[15, 2] = ":";
                    wSheet.Cells[15, 3] = cwTemp.CountryHeadquarters;
                    wSheet.Cells[16, 1] = "Legal Name";
                    wSheet.Cells[16, 2] = ":";
                    wSheet.Cells[16, 3] = cwTemp.LegalName;
                    wSheet.Cells[17, 1] = "Korean Name";
                    wSheet.Cells[17, 2] = ":";
                    wSheet.Cells[17, 3] = cwTemp.KoreanName;
                    wSheet.Cells[18, 1] = "Warrant Style";
                    wSheet.Cells[18, 2] = ":";
                    wSheet.Cells[18, 3] = cwTemp.WarrantStyle;
                    wSheet.Cells[19, 1] = "Edcoid";
                    wSheet.Cells[19, 2] = ":";
                    wSheet.Cells[19, 3] = cwTemp.Edcoid;
                    wSheet.Cells[20, 1] = "Record Type";
                    wSheet.Cells[20, 2] = ":";
                    wSheet.Cells[20, 3] = cwTemp.RecordType;
                    wSheet.Cells[21, 1] = "Issue classification";
                    wSheet.Cells[21, 2] = ":";
                    wSheet.Cells[21, 3] = cwTemp.IssueClassification;
                    wSheet.Cells[22, 1] = "Settlement Type";
                    wSheet.Cells[22, 2] = ":";
                    wSheet.Cells[22, 3] = cwTemp.SettlementType;
                    wSheet.Cells[23, 1] = "Lot size";
                    wSheet.Cells[23, 2] = ":";
                    wSheet.Cells[23, 3] = cwTemp.LotSize;
                    wSheet.get_Range("A25", "C25").MergeCells = true;
                    wSheet.Cells[25, 1] = "'---------------------------------------------------------------------------------------------------------";
                    if (item.ChangeItems.Count != 0)
                    {
                        foreach (string index in item.ChangeItems)
                        {
                            ExcelUtil.GetRange("A" + index + ":C" + index, wSheet).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                        }
                    }
                    excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                    wBook.Save();

                    MailToSend mail = new MailToSend();
                    mail.ToReceiverList.AddRange(configObj.MailTo);
                    mail.CCReceiverList.AddRange(configObj.MailCC);
                    mail.MailSubject = Path.GetFileNameWithoutExtension(fileName);
                    mail.AttachFileList.Add(filePath);
                    mail.MailBody = mailFirstLine + ":\t\t" + item.RIC + "\r\n\r\n"
                                    + "Effective Date:\t\t" + item.EffectiveDate + "\r\n\r\n\r\n\r\n";
                    string signature = string.Join("\r\n", configObj.MailSignature.ToArray());
                    mail.MailBody += signature;

                    AddResult(fileName,filePath,"FM File"); 
                    Logger.Log("Generate FM file. Filepath is " + filePath);
                }
                catch (Exception ex)
                {
                    string msg = "Error found in Generate Company Warrant ADD FM file.      : \r\n" + ex.ToString();
                    Logger.Log(msg, Logger.LogType.Error);
                }
                finally
                {
                    excelApp.Dispose();
                }
            }
        }

        /// <summary>
        /// Create EMA Quantity change file. 
        /// File name: "WRT_QUA_" + theDayBeforeEffictiveDate.ToString("ddMMMyyyy") + "_Korea.csv"; 
        /// </summary>
        /// <param name="item">one warrant</param>      
        private void CreateEMAQUAFile(CompanyWarrantTemplate item)
        {
            Logger.Log("Generate EMA QUA File for :" + item.RIC);
            DateTime effectiveDate = DateTime.Parse(item.EffectiveDate);//, "yyyy-MMM-dd", new CultureInfo("en-US"));
            DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
            string fileName = "WRT_QUA_" + theDayBeforeEffictiveDate.ToString("ddMMMyyyy") + "_Korea.csv";
            string dirToSaveFile = Path.Combine(folderEMA, theDayBeforeEffictiveDate.ToString("yyyy-MM-dd"));
            CreateDirectory(dirToSaveFile);
            string path = Path.Combine(dirToSaveFile, fileName);
            string index = "1";          

            try
            {
                StringBuilder strData = new StringBuilder();
                if (!File.Exists(@path))
                {
                    strData.AppendLine("Logical_Key,Secondary_ID,SecondaryID_Type,EH_Issue_Quantity,Issue_Quantity");
                    strData.AppendLine(index + "," + item.ISIN + ",ISIN,N," + item.QuantityOfWarrants.Replace(",", ""));             
                }
                else
                {
                    string record = "," + item.ISIN + ",ISIN,N," + item.QuantityOfWarrants.Replace(",", "");             
                    strData = ReadExsitEmaRecords(path, item.ISIN, record); 
                }   
                File.WriteAllText(path, strData.ToString(), Encoding.UTF8);
                if (!filesResultList.Contains(fileName))
                {
                    filesResultList.Add(fileName, new List<string>() { path, "EMA QUA FILE" });
                }
                Logger.Log("Generate EMA QUA file successfully. Filepath is " + path);
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when trying to create the EMA QUA . Ex: {0}", ex.Message));
            }
        }
        
        /// <summary>
        /// Create EMA Price change file. 
        /// File name: "WRT_PRC_" + theDayBeforeEffictiveDate.ToString("ddMMMyyyy") + "_Korea.csv"
        /// </summary>
        /// <param name="item">one warrant</param>      
        private void CreateEMAPRCFile(CompanyWarrantTemplate item)
        {
            Logger.Log("Generate EMA PRC File for :" + item.RIC);
            DateTime effectiveDate = DateTime.ParseExact(item.EffectiveDate, "yyyy-MMM-dd", new CultureInfo("en-US"));
            DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
            string fileName = "WRT_PRC_" + theDayBeforeEffictiveDate.ToString("ddMMMyyyy") + "_Korea.csv";
            string dirToSaveFile = Path.Combine(folderEMA, theDayBeforeEffictiveDate.ToString("yyyy-MM-dd"));
            CreateDirectory(dirToSaveFile);
            string path = Path.Combine(dirToSaveFile, fileName);
            string index = "1";
            try
            {
                string warrantPerUnderlying = CalculateWarrantsPerUnderlying(item.ConversionRatio);
                StringBuilder strData = new StringBuilder();
                if (!File.Exists(@path))
                {
                    strData.AppendLine("Logical_Key,Secondary_ID,Secondary_ID_Type,Action,Exotic1_Parameter,Exotic1_Value,EH_Exercise_Price,Exercise_Price,EH_Warrants_Per_Underlying,Warrants_Per_Underlying");
                    strData.AppendLine(index + "," + item.ISIN + ",ISIN,,,,," + item.ExercisePrice + ",N," + warrantPerUnderlying);               
                }
                else
                {
                    string record = "," + item.ISIN + ",ISIN,,,,," + item.ExercisePrice + ",N," + warrantPerUnderlying;                  
                    strData = ReadExsitEmaRecords(path, item.ISIN, record);
                }
                File.WriteAllText(path, strData.ToString(), Encoding.UTF8);
                if (!filesResultList.Contains(fileName))
                {
                    filesResultList.Add(fileName, new List<string>() { path, "EMA PRC FILE" });
                }
                Logger.Log("Generate EMA PRC file successfully. Filepath is " + path);
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when trying to create the EMA PRC file . Ex: {0}", ex.Message));
            }
        }

        /// <summary>
        /// Calculate column:Warrants_Per_Underlying in EMA PRC file. Use warrant's ConvertionRatio
        /// </summary>
        /// <param name="convertionRatio">warrant's ConvertionRatio</param>
        /// <returns>value of Warrants_Per_Underlying</returns>
        private string CalculateWarrantsPerUnderlying(string convertionRatio)
        {
            string result = "";
            double division = 0;
            double tempRatio = 0;
            if (convertionRatio == null)
            {
                return result;
            }
            else
            {
                if (convertionRatio.Contains("%"))
                {
                    convertionRatio = convertionRatio.Replace("%", "");
                    if (Double.TryParse(convertionRatio, out tempRatio))
                    {
                        tempRatio = tempRatio / 100;
                    }
                }
                else if (!Double.TryParse(convertionRatio, out tempRatio))
                {
                    return result;
                }
               
                if (tempRatio != 0)
                {
                    division = 1 / tempRatio;
                    division = Math.Round(division, 5);

                    result = Convert.ToString(division);
                    return result;
                }
            }
            return result;
        }

        /// <summary>
        /// Create GEDA file. 
        /// File name: "KR_CWRTS_Bulk_Change_" + theDayBeforeEffictiveDate.ToString("yyyyMMdd") + ".txt";
        /// </summary>
        /// <param name="item">one warrant</param>
        /// <param name="dir">GEDA folder</param>
        private void CreateGEDAFile(CompanyWarrantTemplate item, string dir)
        {
            Logger.Log("Generate GEDA File for :" + item.RIC);
            DateTime effectiveDate = DateTime.ParseExact(item.EffectiveDate, "yyyy-MMM-dd", new CultureInfo("en-US"));
            DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
            string fileName = "KR_CWRTS_Bulk_Change_" + theDayBeforeEffictiveDate.ToString("yyyyMMdd") + ".txt";
            string path = Path.Combine(dir, fileName);
            try
            {
                StringBuilder strData = new StringBuilder();
                if (!File.Exists(@path))
                {
                    strData.Append("RIC\t#INSTMOD_STRIKE_PRC\r\n");
                }
                else
                {
                    strData = ReadExistRecords(path, item.RIC, '\t');
                }
                strData.Append(item.RIC);
                strData.Append("\t");
                strData.Append(item.ExercisePrice);
                strData.Append("\r\n");
                File.WriteAllText(path, strData.ToString(), Encoding.UTF8);
                if (!filesResultList.Contains(fileName))
                {
                    filesResultList.Add(fileName, new List<string>() { path, "GEDA FILE" });
                }
                Logger.Log("Generate GEDA file successfully. Filepath is " + path);
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when trying to create the GEDA files . Ex: {0}", ex.Message));
            }
        }

        /// <summary>
        /// Create NDA IA File for warrant. Warrants have same effective day will write to the same file.
        /// file name: "KR" + (EffectiveDay-1).ToString("yyyyMMdd") + "IAChg.csv";
        /// </summary>
        /// <param name="item">one warrant</param>
        /// <param name="dir">NDA file path</param>
        private void CreateNDAIAFile(CompanyWarrantTemplate item, string dir)
        {
            Logger.Log("Generate NDA IA File for :" + item.RIC);
            DateTime effectiveDate = DateTime.ParseExact(item.EffectiveDate, "yyyy-MMM-dd", new CultureInfo("en-US"));
            DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
            string fileName = "KR" + theDayBeforeEffictiveDate.ToString("yyyyMMdd") + "IAChg.csv";
            string path = Path.Combine(dir, fileName);
            try
            {
                StringBuilder strData = new StringBuilder();
                if (!File.Exists(@path))
                {
                    strData.Append("ISIN,ASSET COMMON NAME\r\n");
                }
                else
                {
                    strData = ReadExistRecords(path, item.RIC, ',');
                }
                strData.Append(item.ISIN + "," + item.ForIACommonName + "\r\n");
                File.WriteAllText(path, strData.ToString(), Encoding.UTF8);
                if (!filesResultList.Contains(fileName))
                {
                    filesResultList.Add(fileName, new List<string>() { path, "NDA IA FILE" });
                }
                Logger.Log("Generate NDA IA file successfully. Filepath is " + path);
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when trying to create the NDA for RIC:{0}.\r\n Ex: {1}", item.RIC, ex.Message), Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Create NDA QA File for warrant. Warrants have same effective day will write to the same file.
        /// file name: "KR" + (EffectiveDay-1).ToString("yyyyMMdd") + "QAChg.csv";
        /// </summary>
        /// <param name="item">one warrant</param>
        /// <param name="dir">NDA file path</param>
        private void CreateNDAQAFile(CompanyWarrantTemplate item, string dir)
        {
            Logger.Log("Generate NDA QA File for :" + item.RIC);
            DateTime effectiveDate = DateTime.ParseExact(item.EffectiveDate, "yyyy-MMM-dd", new CultureInfo("en-US"));
            DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
            string fileName = "KR" + theDayBeforeEffictiveDate.ToString("yyyyMMdd") + "QAChg.csv";
            string path = Path.Combine(dir, fileName);
            try
            {
                StringBuilder strData = new StringBuilder();
                if (!File.Exists(@path))
                {
                    strData.Append("RIC,ASSET COMMON NAME,STRIKE PRICE\r\n");
                }
                else
                {
                    strData = ReadExistRecords(path, item.RIC, ',');                  
                }
                string ricF = item.RIC.Split('.')[0] + "F." + item.RIC.Split('.')[1];
                strData.Append(item.RIC + "," + item.QACommonName + "," + item.ExercisePrice.ToString() + "\r\n");
                strData.Append(ricF + "," + item.QACommonName + "," + item.ExercisePrice.ToString() + "\r\n");
                File.WriteAllText(path, strData.ToString(), Encoding.UTF8);
                if (!filesResultList.Contains(fileName))
                {
                    filesResultList.Add(fileName, new List<string>() { path, "NDA QA FILE" });
                }
                Logger.Log("Generate NDA QA file successfully. Filepath is " + path);
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when trying to create the NDA for RIC:{0}.\r\n Ex: {1}", item.RIC, ex.Message), Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Read the Exsiting Records. if exsits the same RIC, delete it.  
        /// </summary>
        /// <param name="path">file path</param>
        /// <param name="ric">warrant ric</param>
        /// <param name="spiltChar">character to spilt</param>
        /// <returns></returns>
        private StringBuilder ReadExistRecords(string path, string ric, char spiltChar)
        {
            Logger.Log("File exsits. Write data after the last record. File:" + path);
            StreamReader readFileAll = new StreamReader(path);
            string strDataAll = readFileAll.ReadToEnd();
            readFileAll.Close();
            StreamReader readFile = new StreamReader(path);
            StringBuilder strData = new StringBuilder();
            string line = readFile.ReadLine();
            while (line != null)
            {
                strData.AppendLine(line);
                line = readFile.ReadLine();
            }
            readFile.Close();
            return strData;
        }

        /// <summary>
        /// Read the exsiting EMA file data and get the logical key index.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="isin"></param>
        private StringBuilder ReadExsitEmaRecords(string path, string isin, string record)
        {
            Logger.Log("EMA File exsits. Write data after the last record. File:" + path);
            StreamReader readFileAll = new StreamReader(path);
            string strDataAll = readFileAll.ReadToEnd();
            readFileAll.Close();
            StreamReader readFile = new StreamReader(path);
            StringBuilder strData = new StringBuilder();
            string line = readFile.ReadLine();
            int logicalKey = 0;            
            while (line != null)
            {
                strData.AppendLine(line);               
                line = readFile.ReadLine();
                logicalKey++;
            }            
            readFile.Close();           
            record = logicalKey.ToString() + record;
            strData.Append(record);            
            return strData;
        }

        /// <summary>
        /// Put the task results of GEDA NDA EMA files in Result window of task.
        /// </summary>
        private void SetTaskResultList()
        {
            ArrayList keysArr = new ArrayList(filesResultList.Keys);
            keysArr.Sort();
            foreach (string keyRusult in keysArr)
            {
                List<string> filePathAndType = filesResultList[keyRusult] as List<string>;
                string filePath = filePathAndType[0];
                string resultType = filePathAndType[1].Trim();
                if (resultType == "GEDA FILE")
                {
                    AddResult(keyRusult,filePath,resultType);
                }
                else if (resultType == "NDA IA FILE" || resultType == "NDA QA FILE")
                {
                    AddResult(keyRusult,filePath,resultType);
                }
                else
                {
                    AddResult(keyRusult,filePath,resultType);
                }
            }                   
        }

        public string folderEMAWithDate { get; set; }
    }
}
