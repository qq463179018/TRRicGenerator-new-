using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks
{
    public class WarrantAdd : GeneratorBase
    {       
        private Hashtable revisedWarrant = new Hashtable();       
        private List<CompanyWarrantTemplate> waList = new List<CompanyWarrantTemplate>();
        private KOREACompanyWarrantAddGeneratorConfig configObj = null;
        private const string ETI_KOREA_COMPANYWARRANT_TABLE_NAME = "ETI_Korea_CompanyWarrant";

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREACompanyWarrantAddGeneratorConfig;
            if (string.IsNullOrEmpty(configObj.StartDate))
            {
                configObj.StartDate = DateTime.Now.ToString("yyyy-MM-dd");
            }
            if (string.IsNullOrEmpty(configObj.EndDate))
            {
                configObj.EndDate = DateTime.Now.ToString("yyyy-MM-dd");
            }
          }     

        protected override void Start()
        {
            GrabDataFromKindWebpage();
            if (waList.Count > 0)
            {
                GrabDataFromISINWebpage();
                FormatCompanyWarrantAddData();
                GenerateFMFiles();
                GenerateGEDAFile();
                GenerateNDAQAFile();
                GenerateNDAIAFile();
                GenerateNDATickLotFile();
                GenerateEMAFile();
            }
            else
            {
                Logger.Log("There is no Company Warrant ADD data grabbed.", Logger.LogType.Info);
            }
            TaskResultList.Add(new TaskResultEntry("LOG", "Log File", Logger.FilePath));
        }

        /// <summary>
        /// Grab rights data from kind webpage using the date input(if exsit). Judge the data is KQ or KS.
        /// </summary>
        private void GrabDataFromKindWebpage()
        {
            Logger.Log("Grab data from kind webpage.");
            string startDate = configObj.StartDate.Trim();
            string endDate = configObj.EndDate.Trim();
            if (string.IsNullOrEmpty(startDate))
                startDate = DateTime.Today.ToString("yyyy-MM-dd");
            if (string.IsNullOrEmpty(endDate))
                endDate = DateTime.Today.ToString("yyyy-MM-dd");
            string dataStartDate = (DateTime.Parse(endDate).AddMonths(-2)).ToString("yyyy-MM-dd");
            try
            {
                string uri = string.Format("http://kind.krx.co.kr/disclosure/searchtotalinfo.do");
                string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EA%B6%8C&repIsuSrtCd=&isurCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=300&scn=mktact&srchFd=2&kwd=%EC%8B%A0%EC%A3%BC%EC%9D%B8%EC%88%98%EA%B6%8C%EC%A6%9D%EA%B6%8C&fromData={0}&toData={1}", dataStartDate, endDate);
                string pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, postData);
                HtmlDocument htc = new HtmlDocument();
                if (!string.IsNullOrEmpty(pageSource))
                    htc.LoadHtml(pageSource);
                if (htc != null)
                {
                    HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes(".//dl/dt");
                    HtmlNodeCollection ddCollections = htc.DocumentNode.SelectNodes("//dl/dd");

                    int count = nodeCollections.Count - 1;
                    for (int i = count; i >= 0; i--)
                    {
                        var dtNode = nodeCollections[i] as HtmlNode;
                        if (dtNode != null)
                        {
                            HtmlNode ddNode = ddCollections[i].SelectSingleNode(".//span");

                            HtmlNode node = dtNode.SelectSingleNode(".//span/a");
                            string title = string.Empty;
                            if (node == null)
                            {
                                continue;
                            }
                            title = node.InnerText;
                            if (title.Contains("신주인수권증권 신규상장") || title.Equals("신주인수권증권 상장"))
                            {
                                DateTime anouncementDate = new DateTime();
                                HtmlNode nodeDate = dtNode.SelectSingleNode("./em");
                                if (nodeDate != null)
                                {                                    
                                    anouncementDate = DateTime.Parse(nodeDate.InnerText.Trim(), new CultureInfo("en-US"));
                                    if (anouncementDate < DateTime.Parse(startDate))
                                    {
                                        continue;
                                    }
                                }

                                /* 2014-02-11 kind webpage changed
                                 * 
                                //string parameter = dtNode.SelectSingleNode(".//strong/a").Attributes["onclick"].Value.Trim().ToString();
                                //parameter = parameter.Split('(')[1].Split(')')[0].Trim(new char[] { ' ', '\'', '(', ')' }).ToString();
                                //string url = string.Format("http://kind.krx.co.kr/common/companysummary.do?method=searchCompanySummary&strIsurCd={0}&lstCd=undefined", parameter);

                                //HtmlDocument doc = WebClientUtil.GetHtmlDocument(url, 300000, null);
                                //string judge = string.Empty;
                                //if (doc != null)
                                //{
                                //    judge = doc.DocumentNode.SelectNodes("//table")[0].SelectSingleNode(".//tr[2]/td[2]").InnerText;
                                //}
                                 * 
                                 */

                                string judge = GetDutyCode(ddNode);                                
                                string url = string.Empty;
                                // string judge = string.Empty;
                                                            

                                if (!string.IsNullOrEmpty(judge))
                                {
                                    string attribute = node.Attributes["onclick"].Value.Trim().ToString();
                                    attribute = attribute.Split('(')[1].Split(')')[0].Trim(new char[] { ' ', '\'', '(', ')' }).ToString();
                                    string param1 = attribute.Split(',')[0].Trim(new char[] { ' ', ',', '\'' }).ToString();
                                    string param2 = attribute.Split(',')[1].Trim(new char[] { ' ', ',', '\'' }).ToString();
                                    url = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno={1}&viewerhost=&viewerport=", param1, param2);

                                    HtmlDocument doc = new HtmlDocument();
                                    doc = WebClientUtil.GetHtmlDocument(url, 300000, null);
                                    
                                    string ticker = doc.DocumentNode.SelectSingleNode(".//header/h1").InnerText;
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
                                    string param3 = judge.Equals("KS") ? "68912" : ((judge.Equals("KQ")) ? "70920" : null);
                                    url = string.Format("http://kind.krx.co.kr/external/{0}/{1}/{2}.htm", param1, param2, param3);
                                    doc = WebClientUtil.GetHtmlDocument(url, 300000, null);
                                    if (doc != null)
                                    {
                                        if (judge.Equals("KS"))
                                        {
                                            GrabKSData(doc, ticker, anouncementDate);
                                        }
                                        if (judge.Equals("KQ"))
                                        {
                                            GrabKQData(doc, ticker, anouncementDate);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in Grab Rights Data From KindWebpage: \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            string message = "The number of company warrants between " + startDate + " and " + endDate + " is : " + waList.Count;
            Logger.Log(message, Logger.LogType.Info);
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
        /// Grab data for KQ for an item of company warrant list.
        /// </summary>       
        /// <param name="doc">html to get data</param>
        /// <param name="ticker">warrant ticker</param>
        private void GrabKQData(HtmlDocument doc, string ticker, DateTime anounceDate)
        {
            try
            {
                Logger.Log("Get detail data for KQ RIC:" + ticker);
                HtmlNodeCollection tableNodes = doc.DocumentNode.SelectNodes("//table");
                HtmlNode table = null;
                foreach (var item in tableNodes)
                {
                    string strIsin = "";
                    string strCode = "";
                    try
                    {
                        strIsin = item.SelectSingleNode(".//tr[3]/td[2]").InnerText.Trim().ToString();
                        strCode = item.SelectSingleNode(".//tr[3]/td[3]").InnerText.Trim().ToString();
                    }
                    catch
                    {
                        continue;
                    }
                    if (strIsin.Equals("표준코드") && strCode.Equals("단축코드"))
                    {
                        table = item;
                        break;
                    }
                }
                if (table == null)
                {
                    Logger.Log("Can not get detail information for annoucement for ticker." + ticker);
                    return;
                }

                string koreaName = "";
                string isin = "";
                string koreaCode = "";
                string quantityOfWarrants = "";
                string exercisePrice = "";
                string exercisePeriodFromDate = "";
                string exercisePeriodEndDate = "";
                string effectiveDate = "";
                string issueDate = "";
                bool isinFlag = false;
                bool priceDone = false;
                HtmlNodeCollection lineNodes = table.SelectNodes("//tr");
                CompanyWarrantTemplate cw = new CompanyWarrantTemplate();

                foreach (HtmlNode tr in lineNodes)
                {
                    if (isinFlag == true)
                    {
                        isin = tr.SelectSingleNode("./td[1]").InnerText.Trim();
                        koreaCode = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        isinFlag = false;
                        continue;
                    }
                    Regex regex = new Regex("상장종목명");
                    Match match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        koreaName = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        continue;
                    }
                    regex = new Regex("코드명");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        isinFlag = true;
                        continue;
                    }
                    regex = new Regex("신주인수권 증권의");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        quantityOfWarrants = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        continue;
                    }
                    regex = new Regex("발행일");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        issueDate = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        continue;
                    }
                    regex = new Regex("신주인수권 행사가격");
                    match = regex.Match(tr.InnerText);
                    if (match.Success && !priceDone)
                    {
                        exercisePrice = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        priceDone = true;
                        continue;
                    }
                    regex = new Regex("시작일");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        exercisePeriodFromDate = tr.SelectSingleNode("./td[3]").InnerText.Trim();
                        continue;
                    }
                    regex = new Regex("만료일");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        exercisePeriodEndDate = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        continue;
                    }
                    regex = new Regex("상장일");
                    match = regex.Match(tr.InnerText);
                    if (match.Success)
                    {
                        effectiveDate = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        continue;
                    }
                }

                cw.Ticker = ticker.Trim().ToString();
                cw.RIC = ticker + "W.KQ";
                cw.KoreanName = koreaName.Trim().ToString();
                cw.ISIN = isin.Trim().ToString();
                cw.KoreaCode = koreaCode.Substring(1).Trim().ToString();
                cw.QuantityOfWarrants = quantityOfWarrants.Contains(",") ? quantityOfWarrants.Trim().ToString().Replace(",", "") : quantityOfWarrants;
                cw.ExercisePrice = exercisePrice.Contains(",") ? exercisePrice.Trim().ToString().Replace(",", "") : exercisePrice;
                exercisePeriodFromDate = Convert.ToDateTime(exercisePeriodFromDate.Trim()).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                exercisePeriodEndDate = Convert.ToDateTime(exercisePeriodEndDate.Trim()).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                cw.ExercisePeriod = string.Format("From {0} to {1}", exercisePeriodFromDate, exercisePeriodEndDate);
                cw.ExpiryDate = exercisePeriodEndDate;
                cw.EffectiveDate = Convert.ToDateTime(effectiveDate.Trim()).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                cw.IssueDate = DateTime.Parse(issueDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                cw.Edcoid = cw.Ticker + ".KQ";
                cw.AnouncementDate = anounceDate;
                waList.Add(cw);
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in get detail data for KQ Ticker:" + ticker + "\r\n" + ex.ToString(), Logger.LogType.Error);
            }

        }

        /// <summary>
        /// Grab data for KS for an item of company warrant list.
        /// </summary>
        /// <param name="doc">html to grab data</param>
        /// <param name="ticker">warrant ticker<</param>
        private void GrabKSData(HtmlDocument doc, string ticker, DateTime anounceDate)
        {
            Logger.Log("Get detail data for KS RIC:" + ticker);
            try
            {
                CompanyWarrantTemplate cw = new CompanyWarrantTemplate();
                string strPre = doc.DocumentNode.SelectSingleNode(".//pre").InnerText.Trim();
                int koreanamePos = strPre.IndexOf(". 상장종목 :") + (". 상장종목 :".Length);
                int effectiveDatePos = strPre.IndexOf(". 상장일 :") + (". 상장일 :".Length);
                int quantityWarrantPos = strPre.IndexOf(". 상장신주인수권 증권의 수 :") + (". 상장신주인수권 증권의 수 :".Length);
                int exercisePricePos = strPre.IndexOf(". 신주인수권 행사가격 :") + (". 신주인수권 행사가격 :".Length);
                int exercisePeriodPos = strPre.IndexOf(". 신주인수권 행사기간 :") + (". 신주인수권 행사기간 :".Length);
                int isinKoreacodePos = strPre.IndexOf(". 코드") + (". 코드".Length + 1);
                int issueDatePos = strPre.IndexOf(". 발행일자 :") + (". 발행일자 :".Length);
                string strKoreaname = FormatDataWithPos(koreanamePos, strPre);
                string koreaName = strKoreaname.Substring(0, (strKoreaname.IndexOf("WR") + 2)).Trim().ToString();
                cw.KoreanName = koreaName;
                string strEffectiveDate = FormatDataWithPos(effectiveDatePos, strPre);
                string effectiveDate = Convert.ToDateTime(strEffectiveDate.Trim()).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                cw.EffectiveDate = effectiveDate;
                string strQuantityOfWarrant = FormatDataWithPos(quantityWarrantPos, strPre);
                string quantityOfWarrant = strQuantityOfWarrant.Replace("증권", "").Replace(",", "").Trim();
                cw.QuantityOfWarrants = quantityOfWarrant;
                string strExercisePrice = FormatDataWithPos(exercisePricePos, strPre);
                string exercisePrice = strExercisePrice.Replace("원", "").Trim();
                exercisePrice = exercisePrice.Contains(",") ? exercisePrice.Replace(",", "") : exercisePrice;
                cw.ExercisePrice = exercisePrice;
                string strExercisePeriod = FormatDataWithPos(exercisePeriodPos, strPre);
                string exercisePeriodFrom = Convert.ToDateTime(strExercisePeriod.Split('~')[0].Trim()).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                string exercisePeriodEnd = Convert.ToDateTime(strExercisePeriod.Split('~')[1].Trim()).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                cw.ExercisePeriod = string.Format("From {0} to {1}", exercisePeriodFrom, exercisePeriodEnd);
                cw.ExpiryDate = exercisePeriodEnd.Trim();
                if (isinKoreacodePos < ("8. 코드".Length + 1))
                    isinKoreacodePos = strPre.IndexOf("8. 표준코드") + ("8. 표준코드".Length);
                string strIsinKoreaCode = FormatDataWithPos(isinKoreacodePos, strPre);
                string isin = Regex.Split(strIsinKoreaCode, "단축코드", RegexOptions.IgnoreCase)[0].Split(':')[1].Trim(new char[] { '(', ')', ' ', ':' });
                string koreaCode = Regex.Split(strIsinKoreaCode, "단축코드", RegexOptions.IgnoreCase)[1].Trim(new char[] { '(', ')', ' ', ':' }).Substring(1);
                string issueDate = FormatDataWithPos(issueDatePos, strPre);
                issueDate = Convert.ToDateTime(issueDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                cw.ISIN = isin.Trim();
                cw.KoreaCode = koreaCode.Trim();
                cw.Ticker = ticker;
                cw.RIC = ticker + "W.KS";
                cw.Edcoid = cw.Ticker + ".KS";
                cw.IssueDate = DateTime.Parse(issueDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                cw.AnouncementDate = anounceDate;
                waList.Add(cw);
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in get detail data for KS Ticker:" + ticker + "\r\n" + ex.ToString(), Logger.LogType.Error);
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
            string temp = strPre.Trim(new char[] { '\r', '\n', ' ', ',' }).ToString();

            char[] tempArr = temp.ToCharArray();
            string result = "";
            while (tempArr[pos] != '\n')    //||tempArr[pos] != '\r'
            {
                result += tempArr[pos].ToString();
                if ((pos + 1) < tempArr.Length)
                    pos++;
                else
                    break;
            }
            return result;
        }
        
        /// <summary>
        /// Get legal name from ISIN Webpage
        /// </summary>
        private void GrabDataFromISINWebpage()
        {
            try
            {
                HtmlDocument doc = null;
                foreach (var item in waList)
                {
                    Logger.Log("Grabbing legal name from ISIN Webpage for Ticker:" + item.Ticker);
                    string uri = string.Format("http://isin.krx.co.kr/jsp/BA_VW016.jsp?isu_cd={0}&modi=f&req_no=", item.ISIN);
                    AdvancedWebClient wc = new AdvancedWebClient();
                    string source = WebClientUtil.GetPageSource(wc, uri, 300000, null);
                    if (!string.IsNullOrEmpty(source))
                    {
                        doc = new HtmlDocument();
                        doc.LoadHtml(source);
                    }
                    if (doc != null)
                    {
                        HtmlNode tableNode = doc.DocumentNode.SelectNodes("//table")[2];
                        if (tableNode != null)
                        {
                            string legalName = tableNode.SelectSingleNode(".//tr[6]/td[2]").InnerText.Trim().ToString();
                            item.LegalName = legalName;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in grabbing legal name from ISIN Webpage. \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }
                
        /// <summary>
        /// Format  Company Warrant ADD Data. 
        /// Generate QACommonName, QAShortName. 
        /// If ric has exsited, generate new RIC,  Insert item to DB.
        /// </summary>
        private void FormatCompanyWarrantAddData()
        {
            Logger.Log("Formating company warrant data and update Database.");
            foreach (var item in waList)
            {
                item.Currency = "KRW";
                item.ConversionRatio = "100%";
                item.CountryHeadquarters = "KOR";
                item.WarrantStyle = "American Style";
                item.RecordType = "97";
                item.IssueClassification = "WT";
                item.SettlementType = "Cash";
                item.LotSize = "1";
                item.IssuerORGID = "";
                item.ForIACommonName = "";
                Regex r = new Regex("\\d+WR$");
                Match m1 = r.Match(item.LegalName);
                Match m2 = r.Match(item.KoreanName);

                if (m1.Success && m2.Success && (m1.Value != m2.Value))
                {
                    item.LegalName = item.LegalName.Replace(m1.Value, m2.Value);
                }
                else if (!m1.Success && m2.Success)
                { 
                    item.LegalName += " " + m2.Value;                   
                }
                else if(m1.Success && !m2.Success)
                {
                    item.KoreanName += " " + m1.Value; 
                }

                item.ChangeItems = new List<string>();
                //item.RIC = GenerateRic(item);               
                FormatQANames(item);                
                FormatRIC(item);
            }
        }

        /// <summary>
        /// Format the RIC
        /// Check if RIC is in DB. If exsits and not the same warrant, generate new RIC for this warrant and insert it into DB.
        /// Also mark the revised warrant.
        /// </summary>
        /// <param name="item">warrant item</param>
        private void FormatRIC(CompanyWarrantTemplate item)
        {
            try
            {
                bool isChecked = false;
                string ric = item.RIC;
                string isin = item.ISIN;
                string searchRic = ric.Substring(0, ric.IndexOf('.'));
                string condition = " where ric like '" + searchRic + "%'";
                System.Data.DataTable dt = ManagerBase.Select(ETI_KOREA_COMPANYWARRANT_TABLE_NAME , new string[] { "*" }, condition);
                if (dt == null)
                {
                    Logger.Log("Error found in selecting data from DB. Check if the DB can be attached!", Logger.LogType.Error);
                }
                if (dt.Rows.Count == 0)
                {
                    Logger.Log("The Database doesn't contains the record for RIC:" + item.RIC + "\t Please input the reference fields!");               
                    GetIssuerORGIDFromUser(item);
                    InsertNewRICToDB(item);
                    return;
                }
                foreach (DataRow row in dt.Rows)
                {
                    if (!isChecked)
                    {
                        item.IssuerORGID = Convert.ToString(row["IssuerORGID"]);
                        item.ForIACommonName = Convert.ToString(row["ForIACommonName"]);
                        isChecked = true;
                    }
                    if (isin.Trim() == Convert.ToString(row["ISIN"]).Trim())
                    {
                        item.RIC = Convert.ToString(row["RIC"]).Trim();
                        item.IssuerORGID = Convert.ToString(row["IssuerORGID"]).Trim();
                        MarkWarrantChanges(item, row, dt);                        
                        return;
                    }                    
                }
                item.RIC = searchRic + dt.Rows.Count.ToString() + ric.Substring(ric.IndexOf('.'));
                InsertNewRICToDB(item);
            }
            catch(Exception ex)
            {
                Logger.Log("Error found in searching RIC:" + item.RIC + " in Database .\r\n" + ex.ToString() ,Logger.LogType.Error);               
            }
        }

       /// <summary>
       /// Check the differences between warrant item and DB row. Add the differences in to List ChangeItems.
       /// Update the changes into DB        
       /// </summary>
       /// <param name="item">warrant</param>
       /// <param name="row">warrant info in DB</param>
       /// <param name="dt">warrants table</param>
        private void MarkWarrantChanges(CompanyWarrantTemplate item, DataRow row, System.Data.DataTable dt)
        {           
            string msg = "";           
            string koreaCodeDB = Convert.ToString(row["KoreaCode"]).Trim();
            string exercisePriceDB = Convert.ToString(row["ExercisePrice"]).Trim();
            string quantityOfWarrantsDB = Convert.ToString(row["QuantityOfWarrants"]).Trim();
            string exercisePeriodDB = Convert.ToString(row["ExercisePeriod"]).Trim();
            string koreanNameDB = Convert.ToString(row["KoreanName"]).Trim();
            string legalNameDB = Convert.ToString(row["LegalName"]).Trim();
            string qaCommonNameDB = Convert.ToString(row["QACommonName"]).Trim();
            string qaShortNameDB = Convert.ToString(row["QAShortName"]).Trim();
     
            if (item.KoreaCode.Trim() != koreaCodeDB)
            {
                item.ChangeItems.Add("9");
                row["KoreaCode"] = item.KoreaCode;
                item.ChangeItems.Add("7");
                row["QAShortName"] = item.QAShortName;
                msg += "KoreaCode:" + koreaCodeDB + "->" + item.KoreaCode + "\r\n\t\t\t\t\t\t";
            }            
            if (item.ExercisePrice.Trim() != exercisePriceDB)
            {
                item.ChangeItems.Add("11");                
                row["ExercisePrice"] = item.ExercisePrice;
                item.ChangeItems.Add("6");
                row["QACommonName"] = item.QACommonName;
                msg += "ExercisePrice:" + exercisePriceDB + "->" + item.ExercisePrice + "\r\n\t\t\t\t\t\t"; ;
            }
            if (item.QuantityOfWarrants.Trim() != quantityOfWarrantsDB)
            {
                item.ChangeItems.Add("12");
                row["QuantityOfWarrants"] = item.QuantityOfWarrants;
                msg += "QuantityOfWarrants:" + quantityOfWarrantsDB + "->" + item.QuantityOfWarrants + "\r\n\t\t\t\t\t\t";
            }
            if (item.ExercisePeriod.Trim() != exercisePeriodDB)
            {
                item.ChangeItems.Add("13");
                row["ExercisePeriod"] = item.ExercisePeriod;
                msg += "ExercisePeriod:" + exercisePeriodDB + "->" + item.ExercisePeriod + "\r\n\t\t\t\t\t\t";
            }
            DateTime expiryDateDB = DateTime.Parse(Convert.ToString(row["ExpiryDate"]));
            DateTime expiryDate = DateTime.Parse(item.ExpiryDate);
            if (!expiryDate.Equals(expiryDateDB))
            {
                item.ChangeItems.Add("14");
                row["ExpiryDate"] = expiryDate.Date;
                item.ChangeItems.Add("7");
                row["QAShortName"] = item.QAShortName;
                msg += "ExpiryDate:" + expiryDateDB.ToString("yyyy-MMM-dd").ToUpper() + "->" + item.ExpiryDate + "\r\n\t\t\t\t\t\t";
            }
            if (item.LegalName != legalNameDB)
            {
                item.ChangeItems.Add("16");
                row["LegalName"] = item.LegalName;
                item.ChangeItems.Add("7");
                row["QAShortName"] = item.QAShortName;
                msg += "LegalName:" + legalNameDB + "->" + item.LegalName + "\r\n\t\t\t\t\t\t";            
            }
            if (item.KoreanName.Trim() != koreanNameDB)
            {
                item.ChangeItems.Add("17");
                row["KoreanName"] = item.KoreanName;
                msg += "KoreanName:" + koreanNameDB + "->" + item.KoreanName + "\r\n\t\t\t\t\t\t";
            }           
            DateTime issueDateDB = DateTime.Parse(Convert.ToString(row["IssueDate"]));
            DateTime issueDate = DateTime.Parse(item.IssueDate);
            if (!issueDate.Equals(issueDateDB))
            {
                item.ChangeItems.Add("24");
                row["IssueDate"] = issueDate.Date;
                msg += "IssueDate:" + issueDateDB.ToString("yyyy-MMM-dd").ToUpper() + "->" + item.IssueDate + "\r\n\t\t\t\t\t\t";
            }
            if (item.ChangeItems.Count != 0)
            {
                ManagerBase.UpdateDbTable(dt, ETI_KOREA_COMPANYWARRANT_TABLE_NAME);
                if (revisedWarrant.ContainsKey(item.RIC))
                {
                    revisedWarrant[item.RIC] = item.AnouncementDate;
                }
                else
                {
                    revisedWarrant.Add(item.RIC, item.AnouncementDate);
                }
                msg = "REVISED! 1 record updated. RIC:" + Convert.ToString(row["RIC"]) + "\r\n\t\t\t\t\t\tColumns: " + msg;
                Logger.Log(msg, Logger.LogType.Warning);
            }           
        }

        /// <summary>
        /// If the RIC(or Edcoid) not in DB. User should input IssuerORGID and Issuer LegalName.
        /// </summary>
        /// <param name="item">warrant</param>
        private void GetIssuerORGIDFromUser(CompanyWarrantTemplate item)
        {  
            List<string> input = null;
            input = InputReferenceFields.Prompt(item.Edcoid);          
            if (input == null)
            {
                Logger.Log("User cancelled inputing Issuer ORGID and Issuer LegalName. Colunms in NDA IA and EMA files will leave blank");                 
                return; 
            }
            item.IssuerORGID = input[0];
            item.ForIACommonName = input[1];                    
        }

        /// <summary>
        /// Insert warrant item into DB. 
        /// </summary>
        /// <param name="item">warrant</param>
        private void InsertNewRICToDB(CompanyWarrantTemplate item)
        {
            try
            {
                System.Data.DataTable dt = ManagerBase.Select(ETI_KOREA_COMPANYWARRANT_TABLE_NAME, new string[] { "*" }, string.Format("where RIC='{0}'", item.RIC));
                if (dt.Rows.Count != 0)
                {
                    return;
                }
                DataRow dr = dt.Rows.Add();
                dr["RIC"] = item.RIC;
                dr["UpdateDateAdd"] = DateTime.Today.ToString("yyyy-MM-dd");
                if (!string.IsNullOrEmpty(item.EffectiveDate))
                {
                    dr["EffectiveDateAdd"] = item.EffectiveDate;
                }
                dr["QACommonName"] = item.QACommonName;
                dr["QAShortName"] = item.QAShortName;
                dr["ConversionRatio"] = item.ConversionRatio;
                dr["KoreaCode"] = item.KoreaCode;
                dr["ISIN"] = item.ISIN;
                dr["ExercisePrice"] = item.ExercisePrice;
                dr["QuantityOfWarrants"] = item.QuantityOfWarrants;
                dr["ExercisePeriod"] = item.ExercisePeriod;
                dr["ExpiryDate"] = item.ExpiryDate;
                dr["EffectiveDateDrop"] = item.ExpiryDate;
                dr["LegalName"] = item.LegalName;
                dr["KoreanName"] = item.KoreanName;
                dr["Edcoid"] = item.Edcoid;
                dr["IssueDate"] = item.IssueDate;
                dr["IssuerORGID"] = item.IssuerORGID;
                dr["ForIACommonName"] = item.ForIACommonName;
                dr["Status"] = "Active";
                ManagerBase.UpdateDbTable(dt, ETI_KOREA_COMPANYWARRANT_TABLE_NAME);
                Logger.Log("1 record inserted. RIC:" + item.RIC);
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in Insert new company warrant to Database for RIC:" + item.RIC +"\r\n" + ex.ToString(), Logger.LogType.Error);
            }
        }       

        /// <summary>
        /// Generate QA names for warrant.
        /// </summary>
        /// <param name="item">warrant</param>
        private void FormatQANames(CompanyWarrantTemplate item)
        {
            string prefix = "";
            string suffix = "";
            string legalName = item.LegalName;
            Regex r = new Regex("\\d+WR$");
            Match m = r.Match(legalName);
            if (m.Success)
            {
                suffix = m.Value.Trim();
                legalName = legalName.Replace(m.Groups[0].Value.Trim(), "").Trim();
                prefix = legalName;
            }

            int prefixMaxLength = 19 - item.ExercisePrice.Length - suffix.Length;
            if (prefix.Length > prefixMaxLength)
            {
                prefix = prefix.Substring(0, prefixMaxLength);
            } 

            legalName = prefix.Replace(" ", "");
            string twoletter = string.Empty;
            string fourletter = string.Empty;
            if (legalName.Length <= 4)
            {
                fourletter = legalName;
                twoletter = legalName.Substring(0, 2);
            }
            else
            {
                char[] characters = legalName.ToCharArray();
                for (int i = 0; i < characters.Length && twoletter.Length < 2; i++)
                {
                    if (char.IsLetter(characters[i]))
                    {
                        twoletter += characters[i].ToString();
                    }
                }
                fourletter = twoletter;
                for (int i = characters.Length / 2; i < characters.Length && fourletter.Length < 4; i++)
                {
                    if (char.IsLetter(characters[i]))
                    {
                        fourletter += characters[i].ToString();
                    }
                }
            }

            prefix += " " + suffix;

            string qaCommonName = prefix + " " + fourletter + " " + Convert.ToDateTime(item.ExpiryDate).ToString("MMM-yy",
                                      new CultureInfo("en-US")).Replace("-", "") + " " + item.ExercisePrice + "CWNT";
            item.QACommonName = qaCommonName.ToUpper();
            string qaShortName = twoletter + " " + item.KoreaCode + fourletter + "C";
            item.QAShortName = qaShortName.ToUpper();
        }

        /// <summary>
        /// Generate Company Warrant ADD FM file.
        /// </summary>
        private void GenerateFMFiles()
        {
            Logger.Log("Generate FM files.");
            string folderEMA = ConfigureOperator.GetEmaFileSaveDir();
            if (string.IsNullOrEmpty(folderEMA))
            {
                Logger.Log("Can not get EMA File folder. Please check the Config in DB. The back up files is in folder FM\\EMA_FILE");
                folderEMA = Path.Combine(configObj.FM, "EMA_FILE");
            }
            folderEMA = Path.Combine(folderEMA, DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US")));
            CreateDirectory(folderEMA);
            for (int i = 0; i < waList.Count; i++)
            {               
                ExcelApp excelApp = new ExcelApp(false, false);
                if (excelApp.ExcelAppInstance == null)
                {
                    string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }
                try
                {
                    string filePath = configObj.FM;
                    CompanyWarrantTemplate item = waList[i];
                    string fileName = "KR FM(Company Warrant ADD)Request_" + item.RIC + " (wef " + item.EffectiveDate + ").xls";
                    if (item.ChangeItems.Count != 0)
                    {                        
                        fileName = "KR FM(Company Warrant ADD)Request_" + item.RIC + " (wef " + item.EffectiveDate + ")(Revised).xls";                        
                        Logger.Log("Revised warrant occured!", Logger.LogType.Warning);                        
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
                    wSheet.Cells[1, 3] = "Company Warrant RIC Add";
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
                    ((Range)wSheet.Cells[6, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                    wSheet.Cells[6, 3] = cwTemp.QACommonName;
                    wSheet.Cells[7, 1] = "QA Short Name";
                    wSheet.Cells[7, 2] = ":";
                    ((Range)wSheet.Cells[7, 3]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                    ((Range)wSheet.Cells[7, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
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
                    wSheet.Cells[24, 1] = "Issuer Date";
                    wSheet.Cells[24, 2] = ":";
                    wSheet.Cells[24, 3] = cwTemp.IssueDate;
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
                    wBook.Close();

                    MailToSend mail = new MailToSend();
                    mail.ToReceiverList.AddRange(configObj.MailTo);
                    mail.CCReceiverList.AddRange(configObj.MailCC);
                    mail.MailSubject = Path.GetFileNameWithoutExtension(fileName);
                    mail.AttachFileList.Add(filePath);
                    mail.MailBody = "Company Warrant Add:\t\t" + item.RIC + "\r\n\r\n"
                                    + "Effective Date:\t\t" + item.EffectiveDate + "\r\n\r\n\r\n\r\n";
                    string signature = string.Join("\r\n", configObj.MailSignature.ToArray());
                    mail.MailBody += signature;

                    if (revisedWarrant.ContainsKey(item.RIC) && item.AnouncementDate < (DateTime)revisedWarrant[item.RIC])
                    {
                        waList.Remove(item);
                        i--;
                    }
                    else
                    {
                        File.Copy(filePath, Path.Combine(folderEMA, fileName), true);  
                    }

                    TaskResultList.Add(new TaskResultEntry(fileName, "FM File", filePath, mail)); 
                    Logger.Log("Generate FM file and copy to EMA folder successfully. Filepath is " + filePath);
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
        /// Generate EMA File to EMA file folder 
        /// </summary>
        private void GenerateEMAFile()
        {
            Logger.Log("Generat EMA file.");
            try
            {
                string fileName = "WRT_ADD_" + DateTime.Today.ToString("ddMMyyyy") + "_Korea_Company Warrant.csv";
                string folderEMA = Path.Combine(ConfigureOperator.GetEmaFileSaveDir(), DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US")));
                string filePath = Path.Combine(folderEMA, fileName);
                List<List<string>> content = new List<List<string>>();
                List<string> emaTitle = new List<string>(){ "Logical_Key", "Secondary_ID", "Secondary_ID_Type", "Warrant_Title", "Issuer_OrgId",
                                                        "Issue_Date", "Country_Of_Issue", "Governing_Country", "Announcement_Date", "Payment_Date", 
                                                        "Underlying_Type", "Clearinghouse1_OrgId", "Clearinghouse2_OrgId", "Clearinghouse3_OrgId", 
                                                        "Guarantor", "Guarantor_Type", "Guarantee_Type", "Incr_Exercise_Lot", "Min_Exercise_Lot", 
                                                        "Max_Exercise_Lot", "Rt_Page_Range", "Underwriter1_OrgId", "Underwriter1_Role", 
                                                        "Underwriter2_OrgId", "Underwriter2_Role", "Underwriter3_OrgId", "Underwriter3_Role", 
                                                        "Underwriter4_OrgId", "Underwriter4_Role", "Exercise_Style", "Warrant_Type", "Expiration_Date", 
                                                        "Registered_Bearer_Code", "Price_Display_Type", "Private_Placement", "Coverage_Type", 
                                                        "Warrant_Status", "Status_Date", "Redemption_Method", "Issue_Quantity", "Issue_Price", 
                                                        "Issue_Currency", "Issue_Price_Type", "Issue_Spot_Price", "Issue_Spot_Currency", 
                                                        "Issue_Spot_FX_Rate", "Issue_Delta", "Issue_Elasticity", "Issue_Gearing", "Issue_Premium", 
                                                        "Issue_Premium_PA", "Denominated_Amount", "Exercise_Begin_Date", "Exercise_End_Date", 
                                                        "Offset_Number", "Period_Number", "Offset_Frequency", "Offset_Calendar", "Period_Calendar", 
                                                        "Period_Frequency", "RAF_Event_Type", "Exercise_Price", "Exercise_Price_Type", 
                                                        "Warrants_Per_Underlying", "Underlying_FX_Rate", "Underlying_RIC", "Underlying_Item_Quantity", 
                                                        "Units", "Cash_Currency", "Delivery_Type", "Settlement_Type", "Settlement_Currency", 
                                                        "Underlying_Group", "Country1_Code", "Coverage1_Type", "Country2_Code", "Coverage2_Type", 
                                                        "Country3_Code", "Coverage3_Type", "Country4_Code", "Coverage4_Type", "Country5_Code", 
                                                        "Coverage5_Type", "Note1_Type", "Note1", "Note2_Type", "Note2", "Note3_Type", "Note3", 
                                                        "Note4_Type", "Note4", "Note5_Type", "Note5", "Note6_Type", "Note6", "Exotic1_Parameter", 
                                                        "Exotic1_Value", "Exotic1_Begin_Date", "Exotic1_End_Date", "Exotic2_Parameter", "Exotic2_Value", 
                                                        "Exotic2_Begin_Date", "Exotic2_End_Date", "Exotic3_Parameter", "Exotic3_Value", "Exotic3_Begin_Date", 
                                                        "Exotic3_End_Date", "Exotic4_Parameter", "Exotic4_Value", "Exotic4_Begin_Date", "Exotic4_End_Date", 
                                                        "Exotic5_Parameter", "Exotic5_Value", "Exotic5_Begin_Date", "Exotic5_End_Date", "Exotic6_Parameter", 
                                                        "Exotic6_Value", "Exotic6_Begin_Date", "Exotic6_End_Date", "Event_Type1", "Event_Period_Number1", 
                                                        "Event_Calendar_Type1", "Event_Frequency1", "Event_Type2", "Event_Period_Number2", 
                                                        "Event_Calendar_Type2", "Event_Frequency2", "Exchange_Code1", "Incr_Trade_Lot1", "Min_Trade_Lot1", 
                                                        "Min_Trade_Amount1", "Exchange_Code2", "Incr_Trade_Lot2", "Min_Trade_Lot2", "Min_Trade_Amount2", 
                                                        "Exchange_Code3", "Incr_Trade_Lot3", "Min_Trade_Lot3", "Min_Trade_Amount3", "Exchange_Code4", 
                                                        "Incr_Trade_Lot4", "Min_Trade_Lot4", "Min_Trade_Amount4", "Attached_To_Id", "Attached_To_Id_Type", 
                                                        "Attached_Quantity", "Attached_Code", "Detachable_Date", "Bond_Exercise", "Bond_Price_Percentage" };

               // content.Add(emaTitle);
                int logicalKey = 1;
                foreach (CompanyWarrantTemplate item in waList)
                {                    
                    string qaShortName = item.QAShortName;
                    string warrantTitle = item.LegalName;
                    Regex r = new Regex("\\d+WR$");
                    Match m = r.Match(item.LegalName);
                    if (m.Success)
                    {
                        warrantTitle = item.LegalName.Replace(m.Groups[0].Value.Trim(), "").Trim();
                    }
                    string warrantType = "";
                    string exerciseStyle = "E";
                    string exchange = "KSC";
                    DateTime expiryDate = DateTime.ParseExact(item.ExpiryDate, "yyyy-MMM-dd", new CultureInfo("en-US"));
                    DateTime startDate = DateTime.ParseExact(item.ExercisePeriod.Split(' ')[1], "yyyy-MMM-dd", new CultureInfo("en-US"));
                    DateTime endDate = DateTime.ParseExact(item.ExercisePeriod.Split(' ')[3], "yyyy-MMM-dd", new CultureInfo("en-US"));

                    if (qaShortName[qaShortName.Length - 1] == 'C')
                    {
                        warrantTitle = warrantTitle + " SHS CALL WTS ";
                        warrantType = "Call";
                    }
                    else
                    {
                        warrantTitle = warrantTitle + " SHS PUT WTS ";
                        warrantType = "Put";
                    }
                    if (item.WarrantStyle == "American Style")
                    {
                        exerciseStyle = "A";
                    }
                    if (item.RIC.Substring(item.RIC.IndexOf(".") + 1) == "KQ")
                    {
                        exchange = "KOE";
                    }
                    List<string> emaItem = new List<string>(150);
                    for (int i = 0; i < 127; i++)
                    {
                        emaItem.Add("");
                    }
                    emaItem.Insert(0, logicalKey.ToString());
                    emaItem.Insert(1, item.ISIN);
                    emaItem.Insert(2, "ISIN");
                    emaItem.Insert(3, warrantTitle.ToUpper() + expiryDate.ToString("dd-MMM-yyyy").ToUpper());
                    emaItem.Insert(4, item.IssuerORGID); 
                    emaItem.Insert(5, DateTime.Parse(item.IssueDate).ToString("dd/MM/yyyy").Replace("-", "/"));
                    emaItem.Insert(6, "KOR");
                    emaItem.Insert(7, "KOR");
                    emaItem.Insert(10, "STOCK");
                    emaItem.Insert(17, item.LotSize);
                    emaItem.Insert(18, item.LotSize);
                    emaItem.Insert(29, exerciseStyle);
                    emaItem.Insert(30, warrantType);
                    emaItem.Insert(31, expiryDate.ToString("dd-MM-yyyy").Replace("-", "/"));
                    emaItem.Insert(39, item.QuantityOfWarrants);
                    emaItem.Insert(51, item.LotSize);
                    emaItem.Insert(52, startDate.ToString("dd-MM-yyyy").Replace("-", "/"));
                    emaItem.Insert(53, endDate.ToString("dd-MM-yyyy").Replace("-", "/"));
                    emaItem.Insert(61, item.ExercisePrice);
                    emaItem.Insert(62, "A");
                    emaItem.Insert(63, CalculateWarrantsPerUnderlying(item.ConversionRatio));
                    emaItem.Insert(65, item.Edcoid);  //Underlying RIC 
                    emaItem.Insert(83, "T");
                    emaItem.Insert(84, "Last Trading Day is " + expiryDate.ToString("dd-MMM-yyyy").ToUpper() + ".");
                    emaItem.Insert(127, exchange);
                    emaItem.Insert(128, item.LotSize);
                    emaItem.Insert(129, item.LotSize);
                    content.Add(emaItem);                   
                    logicalKey++;
                }
                CreateDirectory(folderEMA);
                FileUtil.WriteOutputFile(filePath, content, emaTitle, WriteMode.Overwrite);
                //OperateExcel.WriteToCSV(filePath, content);
                TaskResultList.Add(new TaskResultEntry(fileName, "EMA File", filePath, FileProcessType.Other));
                TaskResultList.Add(new TaskResultEntry("FM Files Copies", "FM Files Copies Folder", folderEMA));
                Logger.Log("Generate EMA file successfully. Filepath is " + filePath);
            }
            catch(Exception ex)
            {
                Logger.Log("Error found in generating EMA file.\r\n" + ex.ToString(), Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Generate NDA QA File
        /// fileName : "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAAddCwnt.csv";
        /// </summary>
        private void GenerateNDAQAFile()
        {
            Logger.Log("Generat NDA QA file.");
            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAAddCwnt.csv";
            string filePath = Path.Combine(configObj.NDA, fileName);
            List<List<string>> content = new List<List<string>>();
            List<string> addTitle = new List<string>() { "RIC", "TAG", "TYPE", "CATEGORY", "EXCHANGE", "CURRENCY", "DERIVATIVES FIRST TRADING DAY", 
                                                         "ASSET COMMON NAME", "ASSET SHORT NAME", "CALL PUT OPTION", "STRIKE PRICE", 
                                                         "ROUND LOT SIZE", "EXPIRY DATE", "TICKER SYMBOL", "BASE ASSET" };
         
            foreach (var item in waList)
            {
                List<string> rics = new List<string>();
                List<string> tags = new List<string>();
                string exchange = "";
                if (item.RIC.Substring(item.RIC.IndexOf(".") + 1) == "KQ")
                {
                    rics.Add(item.RIC);
                    rics.Add(item.RIC.Split('.')[0] + "F.KQ");
                    tags.Add("47778");
                    tags.Add("44399");
                    exchange = "KOE";
                }
                else
                {
                    rics.Add(item.RIC);
                    rics.Add(item.RIC.Split('.')[0] + "F.KS");
                    tags.Add("46429");
                    tags.Add("44398");
                    exchange = "KSC";
                }
                for (int i = 0; i < 2; i++)
                {
                    List<string> qaItem = new List<string>();
                    qaItem.Add(rics[i]);
                    qaItem.Add(tags[i]);
                    qaItem.Add("DERIVATIVE");
                    qaItem.Add("EIW");
                    qaItem.Add(exchange);
                    qaItem.Add("KRW");
                    qaItem.Add(DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")).ToUpper());
                    qaItem.Add(item.QACommonName);
                    qaItem.Add(item.QAShortName);
                    qaItem.Add("CALL");
                    qaItem.Add(item.ExercisePrice);
                    qaItem.Add("1");
                    qaItem.Add(DateTime.Parse(item.ExpiryDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")).ToUpper());
                    qaItem.Add(item.KoreaCode);
                    qaItem.Add("ISIN:" + item.ISIN);
                    content.Add(qaItem);                
                }                
            }
            CreateDirectory(Path.GetDirectoryName(filePath));
            FileUtil.WriteOutputFile(filePath, content, addTitle, WriteMode.Append);
            TaskResultList.Add(new TaskResultEntry(fileName, "NDA QA File", filePath, FileProcessType.NDA));
            Logger.Log("Generate NDA QA file Successfully. Filepath is " + filePath);
        }

        /// <summary>
        /// Generate NDA IA File
        /// fileName : "KR" + DateTime.Now.ToString("yyyyMMdd") + "IAAddCwnt.csv";
        /// </summary>
        private void GenerateNDAIAFile()
        {
            Logger.Log("Generat NDA IA file.");
            string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "IAAddCwnt.csv";
            string filePath = Path.Combine(configObj.NDA, fileName);
            List<List<string>> content = new List<List<string>>();
            List<string> addTitle = new List<string>() { "ISIN", "TYPE", "CATEGORY", "WARRANT ISSUER", 
                                                         "RCS ASSET CLASS", "WARRANT ISSUE QUANTITY" };
            
            foreach (var item in waList)
            {
                List<string> rightItem = new List<string>();
                rightItem.Add(item.ISIN);
                rightItem.Add("DERIVATIVE");
                rightItem.Add("EIW");
                rightItem.Add(item.IssuerORGID); //Need to rename
                rightItem.Add("COWNT");
                rightItem.Add(item.QuantityOfWarrants);                
                content.Add(rightItem);
            }
            CreateDirectory(Path.GetDirectoryName(filePath));
            FileUtil.WriteOutputFile(filePath, content, addTitle, WriteMode.Append);    
            TaskResultList.Add(new TaskResultEntry(fileName, "NDA IA File", filePath, FileProcessType.NDA));
            Logger.Log("Generate NDA IA file Successfully. Filepath is " + filePath);
        }

        private void GenerateNDATickLotFile()
        {
            try
            {
                List<string> tickTitle = new List<string>(){"RIC", "TICK NOT APPLICABLE", "TICK LADDER NAME", 
                                                      "TICK EFFECTIVE FROM", "TICK EFFECTIVE TO", "TICK PRICE INDICATOR" };
                List<string> lotTitle = new List<string>(){"RIC", "LOT NOT APPLICABLE", "LOT LADDER NAME", 
                                                      "LOT EFFECTIVE FROM", "LOT EFFECTIVE TO", "LOT PRICE INDICATOR" };
                string today = DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US"));
                string filePathTick = Path.Combine(configObj.NDA, "TickAdd_CWNT_" + today + ".csv");
                string filePathLot = Path.Combine(configObj.NDA, "LotAdd_CWNT_" + today + ".csv");
                List<List<string>> tickContent = new List<List<string>>();
                List<List<string>> lotContent = new List<List<string>>();
                foreach (CompanyWarrantTemplate item in waList)
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
                    string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US"));
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
                TaskResultList.Add(new TaskResultEntry(Path.GetFileName(filePathTick), "NDA Tick Add File", filePathTick, FileProcessType.NDA));
                TaskResultList.Add(new TaskResultEntry(Path.GetFileName(filePathLot), "NDA Lot Add File", filePathLot, FileProcessType.NDA));

                Logger.Log("Generate NDA tick and lot add files successfully.", Logger.LogType.Info);
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in generating NDA Tick and Lot file. \r\n" + ex.ToString(), Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Generate GEDA file.
        /// </summary>
        private void GenerateGEDAFile()
        {
            Logger.Log("Generate GEDA File");
            List<string> warrantAddGEDATitle = new List<string>(){ "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "BCKGRNDPAG", "DSPLY_NMLL",
                                                                   "BCAST_REF", "CHAIN_FLAG", "#INSTMOD_#ISIN", "#INSTMOD_BUYER_ID", "#INSTMOD_MATUR_DATE", 
                                                                   "#INSTMOD_MNEMONIC", "#INSTMOD_STRIKE_PRC", "#INSTMOD_TDN_SYMBOL", "EXL_NAME"};
            System.Data.DataTable dt = GenerateTitle(warrantAddGEDATitle);
            for (int i = 0; i < waList.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = waList[i].RIC;
                dr[1] = waList[i].QAShortName;
                dr[2] = waList[i].RIC;
                dr[3] = waList[i].KoreaCode;
                dr[4] = waList[i].ISIN;
                dr[5] = "****";
                dr[6] = waList[i].KoreanName;
                dr[7] = waList[i].Edcoid;                
                dr[8] = "Y"; 
                dr[9] = waList[i].ISIN;
                dr[10] = "AME";
                dr[11] = DateTime.Parse(waList[i].ExpiryDate).ToString("dd/MM/yyyy");
                dr[12] = "J" + waList[i].KoreaCode;
                dr[13] = waList[i].ExercisePrice; 
                dr[14] = waList[i].KoreaCode;
                if (waList[i].RIC.Substring(waList[i].RIC.IndexOf(".") + 1) == "KQ")
                {
                    dr[15] = "KOSDAQ_EQB_WNT";
                }
                else
                {
                    dr[15] = "KSE_EQB_WNT";
                }                
                dt.Rows.Add(dr);
            }
            WriteAddGEDAFile(dt);
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
            try
            {
                string fileName = "KR_CWRTS_Bulk_Creation_" + DateTime.Now.ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
                string filePath = Path.Combine(configObj.GEDA, fileName);
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
                Logger.Log("Generate Company Warrant ADD GEDA file Successfully. Filepath is " + filePath);
                TaskResultList.Add(new TaskResultEntry(fileName, "GEDA FILE", filePath, FileProcessType.GEDA_BULK_RIC_CREATION));
            }
            catch(Exception ex)
            {
                Logger.Log("Error found in generating GEDA file. \r\n" + ex.ToString(), Logger.LogType.Error);
            }
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

        /// <summary>
        /// Calculate column:Warrants_Per_Underlying in EMA file. Use warrant's ConvertionRatio
        /// </summary>
        /// <param name="convertionRatio">warrant's ConvertionRatio</param>
        /// <returns>value of Warrants_Per_Underlying</returns>
        private string CalculateWarrantsPerUnderlying(string conversionRatio)
        {
            string result = "";
            double division = 0;
            double tempRatio = 0;
            if (conversionRatio == null)
            {
                return result;
            }
            else
            {
                if (conversionRatio.Contains("%"))
                {
                    conversionRatio = conversionRatio.Replace("%", "");
                    tempRatio = Convert.ToDouble(conversionRatio) / 100;
                }
                if (tempRatio != 0)
                {
                    division = 1 / tempRatio;
                    Math.Round(division, 5);

                    result = Convert.ToString(division);
                    return result;
                }
            }
            return result;
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
    }
}
