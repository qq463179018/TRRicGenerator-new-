using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using HtmlAgilityPack;
using System.Globalization;
using System.Text.RegularExpressions;
using Ric.Db.Info;
using Ric.Db.Manager;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Collections;
using Ric.Util;
using Ric.Core;

namespace Ric.Tasks.Korea
{
    public class EQDrop : GeneratorBase
    {
        private List<string> errors = new List<string>();
        private List<DropTemplate> edropList = new List<DropTemplate>();
        private KoreaDropConfig configObj = null;
        int idnN = 0;

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KoreaDropConfig;
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

        protected override void Start()
        {
            try
            {
                GrabDropData();
                if (edropList.Count > 0)
                {
                    FormatData();
                    GenerateFmFile();
                    GenerateGedaFile();
                    GenerateNdaFile();
                    UpdateEquityDB();
                    //InsertGedaInfoToDb();
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("At Start(). Error message:{0}/r/n/t/t {1}", ex.Message, ex.StackTrace);
                errors.Add(msg);
            }
            finally
            {
                LogErrors();
            }
        }

        private void UpdateEquityDB()
        {
            int row = 0;
            foreach (DropTemplate item in edropList)
            {
                row += KoreaEquityManager.UpdateDrop(item.RIC, item.EffectiveDate);
            }
            Logger.Log(string.Format("Update {0} records in database Equity table.", row.ToString()));
        }

        private void GrabDropData()
        {
            string dataStartDate = null;
            DateTime startDay = DateTime.Parse(configObj.StartDate);
            DateTime twoMonthEarly = DateTime.Parse(configObj.EndDate).AddMonths(-2);
            if (startDay.CompareTo(twoMonthEarly) < 0)
            {
                dataStartDate = configObj.StartDate;
            }
            else
            {
                dataStartDate = twoMonthEarly.ToString("yyyy-MM-dd");
            }

            string pageSource = GetAnnoucements(dataStartDate, configObj.EndDate);
            if (string.IsNullOrEmpty(pageSource))
            {
                string msg = "Can not get announcement web page source. Please check the internet.";
                errors.Add(msg);
                return;
            }
            HtmlDocument htc = new HtmlDocument();
            htc.LoadHtml(pageSource);
            GrabDropDetailData(htc, startDay);    

            string dataMsg = string.Format("From {0} to {1}, we grabbed {2} drop announcements.", configObj.StartDate, configObj.EndDate, edropList.Count);
            Logger.Log(dataMsg);
        }

        private string GetAnnoucements(string startDate, string endDate)
        {
            string uri = "http://kind.krx.co.kr/disclosure/searchtotalinfo.do";
            string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%83%81%EC%9E%A5%ED%8F%90%EC%A7%80&repIsuSrtCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=500&scn=mktact&srchFd=2&kwd=%EC%83%81%EC%9E%A5%ED%8F%90%EC%A7%80&fromData={0}&toData={1}", startDate, endDate);

            int retry = 3;
            string pageSource = null;
            while (pageSource == null && retry-- > 0)
            {
                pageSource = WebClientUtil.GetDynamicPageSource(uri, 180000, postData);
            }
            return pageSource;
        }

        private void GrabDropDetailData(HtmlDocument htc, DateTime startDay)
        {
            HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes("//dl/dt");
            HtmlNodeCollection ddCollections = htc.DocumentNode.SelectNodes("//dl/dd");
            int count = nodeCollections.Count;

            for (var i = 0; i < count; i++)
            {
                try
                {
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

                    HtmlNode node = nodeCollections[i].SelectSingleNode(".//span/a");
                    string title = string.Empty;
                    if (node != null)
                    {
                        title = node.InnerText.Trim();
                        if (title.Contains("["))
                        {
                            title = title.Replace("[", "").Replace("]", "");
                        }
                    }
                    if (string.IsNullOrEmpty(title))
                    {
                        continue;
                    }

                    if (title.Contains("상장폐지") || title.Contains("정정 상장폐지"))
                    {
                        if (title.Replace(" ", "").Contains("유가증권시장상장"))
                        {
                            continue;
                        }

                        //For BC drop
                        if (title.Contains("수익증권 상장폐지") && !title.Contains("예고안내"))
                        {
                            DropTemplate item = new DropTemplate();
                            item.AnnouncementTime = nodeDate.InnerText.Trim();
                            item.Market = "KS";
                            HtmlDocument doc = GetTargetPageDocument(node);
                            GrabBcDropData(doc, item);
                            continue;
                        }

                        string judge1 = title.Substring(0, ("상장폐지".Length)).Trim().ToString();
                        string judge2 = String.Empty;
                        if (title.Length > "정정 상장폐지".Length)
                        {
                            judge2 = title.Substring(0, ("정정 상장폐지".Length)).Trim().ToString();
                        }
                        //For PEO and ETF drop
                        if (judge1.Equals("상장폐지") || judge2.Equals("정정 상장폐지"))
                        {
                            //Console.WriteLine("EEE_Time: " + nodeDate.InnerText);
                            HtmlNode ddNode = ddCollections[i].SelectSingleNode(".//span");
                            HtmlNode companyNode = nodeCollections[i].SelectSingleNode(".//strong/a");
                            DropTemplate item = new DropTemplate();
                            item.AnnouncementTime = nodeDate.InnerText.Trim();
                            item.Market = GetDutyCode(ddNode);
                            HtmlDocument doc = GetTargetPageDocument(node);
                            item.isRevised = judge2.Equals("정정 상장폐지") ? true : false;
                            GrabIsinInfo(companyNode, item);

                            if (item.Type == null)
                            {
                                if (KoreaEquityManager.DelistedEquity(item.ISIN))
                                {
                                    Logger.Log(string.Format("Please notice the announcement at {0} is already dropped.", item.AnnouncementTime), Logger.LogType.Warning);
                                    continue;
                                }
                                string msg = string.Format("At GrabDropDetailData(). Can not get instrument type information from database for announcement at {0}", item.AnnouncementTime);
                                errors.Add(msg);
                                continue;
                            }
                            if (item.Type.Equals("ETF"))
                            {
                                item.RIC = GetTickerByAttribute(node) + "." + item.Market;
                                GrabDropDataText(doc, item);
                            }
                            //PEO
                            else
                            {
                                GrabPeoDropData(doc, item);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    string msg = string.Format("At GrabDropDetailData(). Error found in grab drop detail information. Error Message:{0}.", ex.Message);
                    errors.Add(msg);
                }
            }
        }

        private void GrabBcDropData(HtmlDocument doc, DropTemplate item)
        {
            try
            {
                int count = 0;
                if (doc.DocumentNode.SelectNodes("//table") != null)
                {
                    count = doc.DocumentNode.SelectNodes("//table").Count;
                }
                if (count == 0)
                {
                    string msg = string.Format("At GrabBcDropData(). Can not get BC drop detail infomation for RIC:{0}.", item.RIC);
                    errors.Add(msg);
                    return;
                }
                HtmlNode table = doc.DocumentNode.SelectNodes("//table")[(count - 1)];
                HtmlNodeCollection trs = table.SelectNodes("//tr");
                string effectiveDate = string.Empty;
                foreach (HtmlNode tr in trs)
                {
                    string tdTitle = tr.SelectSingleNode("./td[1]").InnerText.Trim();
                    if (Regex.Replace(tdTitle, @"[0-9. ]+", "", RegexOptions.IgnoreCase).Equals("상장폐지대상수익증권"))
                    {
                        string isin = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        Regex r = new Regex(@"\([0-9A-Za-z]+\)");
                        MatchCollection m = r.Matches(isin);
                        if (m.Count == 0)
                        {
                            string msg = "At GrabBcDropData(). Can not get BC ISIN from page.";
                            errors.Add(msg);
                            return;
                        }
                        else
                        {
                            item.ISIN = m[m.Count - 1].Value.Trim(new char[] { '(', ')' });
                            continue;
                        }
                    }
                    if (tdTitle.Contains("상장폐지일"))
                    {
                        effectiveDate = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        break;
                    }
                }
                if (!string.IsNullOrEmpty(effectiveDate))
                {                   
                    item.EffectiveDate = effectiveDate;
                }
                item.Type = "CEF";
                edropList.Add(item);
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabBcDropData(). Error found in grab bc drop detail information for RIC:{0}. Error Message:{1}.", item.RIC, ex.Message);
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Grab ISIN infos from the first link node in anouncement.
        /// </summary>
        /// <param name="judgeNode">company link node</param>
        /// <returns>korea equity info</returns>
        private void GrabIsinInfo(HtmlNode judgeNode, DropTemplate item)
        {
            try
            {
                string judgeText = judgeNode.Attributes["onclick"].Value.Trim().ToString();
                judgeText = judgeText.Split('(')[1].Split(')')[0].Trim(new char[] { '\'', ' ' }).ToString();
                KoreaCompany company = Common.GetCompanyInfo(judgeText);
                if (company == null)
                {
                    return;
                }
                item.ISIN = company.ISIN;
                KoreaEquityInfo equity = KoreaEquityManager.SelectEquityByIsin(item.ISIN);
                if (equity == null)
                {
                    return;
                }
                item.Type = equity.Type;
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabIsinInfo(). Error message:{0}", ex.Message);
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Grab the announcement that target data is in a text field 
        /// </summary>
        /// <param name="doc">page document</param>
        /// <param name="item">drop item</param>
        private void GrabDropDataText(HtmlDocument doc, DropTemplate item)
        {
            try
            {
                HtmlNode pre = doc.DocumentNode.SelectSingleNode("//pre");
                string strPre = string.Empty;
                if (pre != null)
                    strPre = pre.InnerText.ToString().Trim();

                int effectivePos = strPre.IndexOf("상장폐지일") + ("상장폐지일".Length);
                string sedate = FormatDataWithPos(effectivePos, strPre);
                if (!string.IsNullOrEmpty(sedate))
                {
                    item.EffectiveDate = sedate;
                    edropList.Add(item);
                } 
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabDropDataText(). Error found in grab drop detail information for RIC:{0}. Error Message:{1}.", item.RIC, ex.Message);
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Grab Peo drop data. Market: KQ KN
        /// </summary>
        /// <param name="doc">page document</param>
        /// <param name="item">drop item</param>
        private void GrabPeoDropData(HtmlDocument doc, DropTemplate item)
        {
            try
            {
                int count = 0;
                if (doc.DocumentNode.SelectNodes("//table") != null)
                {
                    count = doc.DocumentNode.SelectNodes("//table").Count;
                }
                if (count == 0)
                {
                    string msg = string.Format("At GrabPeoKQKNDropData(). Can not get peo drop detail infomation for RIC:{0}.", item.RIC);
                    errors.Add(msg);
                    return;
                }
                HtmlNode table = doc.DocumentNode.SelectNodes("//table")[(count - 1)];
                HtmlNodeCollection trs = table.SelectNodes("//tr");
                string effectiveDate = string.Empty;
                int tickerPos = -1;
                List<string> rics = new List<string>();
                foreach (HtmlNode tr in trs)
                {
                    HtmlNodeCollection tdNode = tr.SelectNodes(".//td");
                    if (tdNode.Count > 2 && tickerPos == -1)
                    {
                        for (int pos = 1; pos < tdNode.Count; pos++)
                        {
                            HtmlNode tickerNode = tdNode[pos];
                            if (tickerNode.InnerText.Contains("단축코드"))
                            {
                                tickerPos = pos - 1;
                                break;
                            }
                        }
                        if (tickerPos != -1 && tickerPos != 99)
                        {
                            continue;
                        }
                    }

                    if (tickerPos != -1 && tickerPos != 99)
                    {
                        if (tdNode.Count > tickerPos)
                        {
                            string ticker = tdNode[tickerPos].InnerText.Trim();
                            Regex regex = new Regex("[A-Za-z][0-9A-Za-z]+");
                            Match match = regex.Match(ticker);
                            if (match.Success)
                            {
                                rics.Add(ticker.Substring(1));
                                continue;
                            }
                            else
                            {
                                tickerPos = 99;
                            }
                        }
                        else
                        {
                            tickerPos = 99;
                        }
                    }
                    string tdTitle = tr.SelectSingleNode("./td[1]").InnerText.Trim();
                    if (tdTitle.Contains("상장폐지일"))
                    {
                        effectiveDate = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        break;
                    }
                }
                if (!string.IsNullOrEmpty(effectiveDate))
                {
                    foreach (string ric in rics)
                    {
                        DropTemplate dItem = new DropTemplate();
                        dItem.RIC = ric + "." + item.Market;
                        dItem.EffectiveDate = effectiveDate;
                        dItem.Market = item.Market;
                        dItem.AnnouncementTime = item.AnnouncementTime;
                        edropList.Add(dItem);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabPeoKQKNDropData(). Error found in grab drop detail information for RIC:{0}. Error Message:{1}.", item.RIC, ex.Message);
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Grab ETF drop data. Page type: text
        /// </summary>
        /// <param name="doc">page document</param>
        /// <param name="item">drop item</param>
        private void GrabEtfDropData(HtmlDocument doc, DropTemplate item)
        {
            try
            {
                HtmlNode pre = doc.DocumentNode.SelectSingleNode("//pre");
                string strPre = string.Empty;
                if (pre != null)
                    strPre = pre.InnerText.ToString().Trim();

                int effectivePos = strPre.IndexOf("상장폐지일") + ("상장폐지일".Length);
                string sedate = FormatDataWithPos(effectivePos, strPre);
                if (!string.IsNullOrEmpty(sedate))
                {
                    item.EffectiveDate = sedate;
                    edropList.Add(item);
                }                
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabBcDropData(). Error found in grab drop detail information for RIC:{0}. Error Message:{1}.", item.RIC, ex.Message);
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Grab Peo drop data. Page type: text
        /// If contains mutiple records. Add them orderly.
        /// </summary>
        /// <param name="doc">page document</param>
        /// <param name="item">drop item</param>
        private void GrabPeoDropDataText(HtmlDocument doc, DropTemplate item)
        {
            try
            {
                HtmlNode pre = doc.DocumentNode.SelectSingleNode("//pre");
                string strPre = string.Empty;
                if (pre != null)
                    strPre = pre.InnerText.ToString().Trim();

                int effectivePos = strPre.IndexOf("상장폐지일") + ("상장폐지일".Length);
                string effectiveDate = FormatDataWithPos(effectivePos, strPre);

                while (strPre.Contains("종목코드"))
                {
                    int isinPos = strPre.IndexOf("종목코드") + ("종목코드".Length);
                    string isin = FormatDataWithPos(isinPos, strPre);
                    strPre = strPre.Substring(isinPos);
                    DropTemplate dItem = new DropTemplate();
                    dItem.EffectiveDate = effectiveDate;
                    dItem.ISIN = isin;
                    dItem.Market = item.Market;
                    edropList.Add(dItem);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabPeoKSDropData(). Error found in grab drop detail information for RIC:{0}. Error Message:{1}.", item.RIC, ex.Message);
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
        /// Get ticker in target document. E.g. 6 digits in (123456)
        /// </summary>
        /// <param name="node">html node</param>
        /// <returns>ticker</returns>
        private string GetTickerByAttribute(HtmlNode node)
        {
            string attribute = node.Attributes["onclick"].Value.Trim().ToString();
            attribute = attribute.Split('(')[1].Split(')')[0].Trim(new char[] { ' ', '\'', '(', ')' }).ToString();
            string param1 = attribute.Split(',')[0].Trim(new char[] { ' ', ',', '\'' }).ToString();
            string param2 = attribute.Split(',')[1].Trim(new char[] { ' ', ',', '\'' }).ToString();
            string url = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno={1}&viewerhost=&viewerport=", param1, param2);
            HtmlDocument doc = WebClientUtil.GetHtmlDocument(url, 300000, null);
            string ticker = doc.DocumentNode.SelectSingleNode(".//div[@id='pHeader']/h2").InnerText;
            string category = string.Empty;
            HtmlNode categoryNode = doc.DocumentNode.SelectSingleNode(".//div[@id='pWrapper']");//div[@id='pContArea']/form[@name='frm']/div[@id='vHeader']/div[1]/div[1]/select[1]/option[2]");

            if (!string.IsNullOrEmpty(ticker))
            {
                Match m = Regex.Match(ticker, @"\(([0-9a-zA-Z]{6})\)");
                if (m == null)
                {
                    string msg = "At GetTickerByAttribute(). Cannot get ticker numbers in ." + ticker;
                    errors.Add(msg);
                    return null;
                }
                ticker = m.Groups[1].Value;
            }

            return ticker;
        }

        private void FormatData()
        {
            foreach (DropTemplate item in edropList)
            {
                KoreaEquityInfo equity = null;
                if (!string.IsNullOrEmpty(item.ISIN))
                {
                    equity = KoreaEquityManager.SelectEquityByIsin(item.ISIN);
                }
                else if(!string.IsNullOrEmpty(item.RIC))
                {
                    equity = KoreaEquityManager.SelectEquityByRic(item.RIC);
                }

                if (equity == null)
                {
                    string msg = string.Format("At FormatData(). Can not get listed information for Annoucement at {0}.", item.AnnouncementTime);
                    errors.Add(msg);
                    //edropList.Remove(item);
                    continue;
                }
                item.RIC = equity.RIC;
                item.QAShortName = equity.IDNDisplayName;
                item.LegalName = equity.LegalName;
                item.ISIN = equity.ISIN;
                item.Type = equity.Type;
                try
                {
                    item.EffectiveDate = DateTime.Parse(item.EffectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                }
                catch (Exception ex)
                {
                    string msg = string.Format("At FormatData(). Error found when converting effective date for RIC:{0}. Effective Date:{1}./r/n/t/t Error message:", item.RIC, item.EffectiveDate, ex.Message);
                    errors.Add(msg);
                }
            }
        }

        private void GenerateFmFile()
        {
            Logger.Log("Generate FM Files.");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                errors.Add(msg);
                return;
            }
            foreach (DropTemplate item in edropList)
            {
                try
                {
                    string effectiveDate = item.EffectiveDate.ToUpper();
                    string filePath = configObj.FM;
                    string fileName = "KR FM ({0}) Request_{1} (wef {2}).xls";
                    string dropType = string.Empty;
                    if (item.Type.Equals("ORD") || item.Type.Equals("PRF") || item.Type.Equals("KDR"))
                    {
                        dropType = "Drop";
                    }
                    else if (item.Type.Equals("ETF"))
                    {
                        dropType = "ETF Drop";
                    }
                    else if (item.Type.Equals("REIT"))
                    {
                        dropType = "REIT Drop";
                    }
                    else if (item.Type.Equals("CEF"))
                    {
                        dropType = "BC Drop";
                    }

                    fileName = string.Format(fileName, dropType, item.RIC, effectiveDate);
                    if (item.isRevised)
                    {
                        fileName = "(Revised) " + fileName;
                    }

                    filePath = Path.Combine(filePath, fileName);

                    Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                    if (wSheet == null)
                    {
                        string msg = "Worksheet could not be started. Check that your office installation and project reference are correct!";
                        errors.Add(msg);
                        continue;
                    }
                    GenerateEachFmFile(item, wSheet);
                    excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                    wBook.Save();

                    MailToSend mail = new MailToSend();
                    mail.ToReceiverList.AddRange(configObj.MailTo);
                    mail.CCReceiverList.AddRange(configObj.MailCC);
                    mail.MailSubject = Path.GetFileNameWithoutExtension(fileName);
                    mail.AttachFileList.Add(filePath);
                    mail.MailBody = "Drop:\t\t" + item.RIC + "\r\n\r\n"
                                    + "Effective Date:\t\t" + item.EffectiveDate + "\r\n\r\n\r\n\r\n";
                    string signature = string.Join("\r\n", configObj.MailSignature.ToArray());
                    mail.MailBody += signature;

                    AddResult(fileName,filePath,"FM File");
                    Logger.Log("Generate FM file successfully. Filepath is " + filePath);
                }
                catch (Exception ex)
                {
                    string msg = string.Format("At GenerateFmFile(). Error found when generating drop FM for RIC:{0}./r/n/t/t Error message:", item.RIC, ex.Message);
                    errors.Add(msg);
                }
            }
            excelApp.Dispose();
        }

        private void GenerateEachFmFile(DropTemplate item, Worksheet wSheet)
        {
            try
            {
                wSheet.Name = "DROP";
                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 2;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 30;
                ((Range)wSheet.Columns["A:C", System.Type.Missing]).Font.Name = "Arial";
                ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

                wSheet.Cells[1, 1] = "FM Request";
                wSheet.Cells[1, 2] = " ";
                wSheet.Cells[1, 3] = "Deletion";
                ((Range)wSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                wSheet.Cells[3, 1] = "Effective Date";
                wSheet.Cells[3, 2] = ":";
                ((Range)wSheet.Cells[3, 3]).NumberFormat = "@";
                wSheet.Cells[3, 3] = item.EffectiveDate;
                ((Range)wSheet.Cells[4, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Cells[4, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                wSheet.Cells[4, 1] = "RIC";
                wSheet.Cells[4, 2] = ":";
                wSheet.Cells[4, 3] = item.RIC;
                ((Range)wSheet.Cells[5, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Cells[5, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                wSheet.Cells[5, 1] = "ISIN";
                wSheet.Cells[5, 2] = ":";
                wSheet.Cells[5, 3] = item.ISIN;
                wSheet.Cells[6, 1] = "QA Short Name";
                wSheet.Cells[6, 2] = ":";
                ((Range)wSheet.Cells[6, 3]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                ((Range)wSheet.Cells[6, 3]).Font.Underline = true;
                wSheet.Cells[6, 3] = item.QAShortName;
                wSheet.Cells[7, 1] = "Legal Name";
                wSheet.Cells[7, 2] = ":";
                wSheet.Cells[7, 3] = item.LegalName;
                if (item.Type.Equals("ORD") || item.Type.Equals("PRF") || item.Type.Equals("KDR"))
                {
                    wSheet.Cells[1, 3] = "Equity Deletion";
                    wSheet.Cells[8, 1] = "ICW Index Drops";
                    wSheet.Cells[8, 2] = ":";
                    wSheet.Cells[8, 3] = "";
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GenerateEachFmFile(). Error found when writing information to drop FM for RIC:{0}./r/n/t/t Error message:", item.RIC, ex.Message);
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Generate Drop Geda files. 
        /// Write the records to the file named with effective date.
        /// </summary>
        private void GenerateGedaFile()
        {
            Logger.Log("Generate GEDA files.");
            Hashtable result = new Hashtable();
            List<string> gedaTitle = new List<string>() { "RIC" };
            foreach (DropTemplate item in edropList)
            {
                try
                {
                    if (item.isRevised)
                    {
                        continue;
                    }
                    string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US"));
                    string fileName = "KR_DROP_" + effectiveDate + ".txt";
                    string filePath = Path.Combine(configObj.GEDA, fileName);
                    List<List<string>> data = new List<List<string>>();
                    List<string> record = new List<string>();
                    record.Add(item.RIC);
                    data.Add(record);
                    FileUtil.WriteOutputFile(filePath, data, gedaTitle, WriteMode.Append);
                    if (!result.Contains(fileName))
                    {
                        TaskResultEntry entry = new TaskResultEntry(fileName, "GEDA DROP FILE", filePath, FileProcessType.GEDA_BULK_RIC_DELETE);
                        result.Add(fileName, entry);
                        Logger.Log("Generate GEDA file successfully. Filepath is " + filePath);
                    }
                }
                catch (Exception ex)
                {
                    string msg = "At GenerateGedaFile(). Error found in generating GEDA file for RIC:" + item.RIC + "\r\n " + ex.Message;
                    errors.Add(msg);
                }
            }
            SetTaskResultList(result);
        }

        /// <summary>
        /// For PEO/ETF/BC, the NDA file name are different.
        /// </summary>
        /// <param name="item">drop item</param>
        /// <returns></returns>
        private string[] GenerateNdaRicSuffix(DropTemplate item)
        {
            string[] rics = null;
            if (item.Type.Equals("ORD") || item.Type.Equals("PRF") || item.Type.Equals("KDR"))
            {
                rics = new string[6] { ".", "F.", "S.", "stat.", "ta.", "bl." };
            }
            else if (item.Type.Equals("ETF"))
            {
                rics = new string[7] { ".", "F.", "S.", "stat.", "ta.", "bl.", "LP." };
            }
            else if (item.Type.Equals("REIT"))
            {
                rics = new string[6] { ".", "F.", "S.", "stat.", "ta.", "bl." };
            }
            else if (item.Type.Equals("BC"))
            {
                rics = new string[2] { ".", "F." };
            }
            return rics;
        }

        /// <summary>
        /// Set the result list.
        /// </summary>
        /// <param name="result">results</param>
        private void SetTaskResultList(Hashtable result)
        {
            ArrayList keysArr = new ArrayList(result.Keys);
            keysArr.Sort();
            foreach (string keyRusult in keysArr)
            {
                TaskResultEntry entry = result[keyRusult] as TaskResultEntry;
                TaskResultList.Add(entry);
            }
        }

        private void GenerateNdaFile()
        {
            Logger.Log("Generate NDA QA file.");
            try
            {
                string fileName = "KR" + DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US")) + "QADrop.csv";
                string filePath = Path.Combine(configObj.NDA, fileName);
                List<string> ndaQaTitle = new List<string>() { "RIC", "RETIRE DATE" };
                List<List<string>> data = new List<List<string>>();
                foreach (DropTemplate item in edropList)
                {
                    if (item.isRevised)
                    {
                        continue;
                    }
                    string[] suffix = GenerateNdaRicSuffix(item);
                    for (int i = 0; i < suffix.Length; i++)
                    {
                        string ric = item.RIC.Split('.')[0] + suffix[i] + item.RIC.Split('.')[1];
                        if (i == 2 && item.Market.Equals("KN"))
                        {
                            continue;
                        }
                        string expiryDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
                        List<string> record = new List<string>();
                        record.Add(ric);
                        record.Add(expiryDate);
                        data.Add(record);
                    }
                }
                FileUtil.WriteOutputFile(filePath, data, ndaQaTitle, WriteMode.Overwrite);
                AddResult(fileName,filePath,"NDA QA FILE");
                Logger.Log("Generate NDA QA file successfully. Filepath is " + filePath);
            }
            catch (Exception ex)
            {
                string msg = "At GenerateNdaFile().Error found in generating NDA QA file.\r\n " + ex.Message;
                errors.Add(msg);
            }
        }

        /// <summary>
        /// Insert drop records into DB. Contains EffectiveDate, RIC and TaskId.
        /// </summary>
        private void InsertGedaInfoToDb()
        {
            int rows = 0;
            foreach (DropTemplate item in edropList)
            {
                try
                {
                    rows += DropGedaManager.UpdateDrop(item.EffectiveDate, item.RIC, TaskId);
                }
                catch (Exception ex)
                {
                    string msg = "Error found in insert drop record to DB. For RIC:" + item.RIC + "\r\n" + ex.Message;
                    errors.Add(msg);
                }
            }
            Logger.Log(string.Format("Update {0} drop records in DROP_GEDA table.", rows.ToString()));
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



        private void GrabDropDetailDataForTest(HtmlDocument htc, DateTime startDay)
        {
            HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes("//dl/dt");
            HtmlNodeCollection ddCollections = htc.DocumentNode.SelectNodes("//dl/dd");
            int count = nodeCollections.Count;

            for (var i = 0; i < count; i++)
            {
                try
                {
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

                    HtmlNode node = nodeCollections[i].SelectSingleNode(".//span/a");
                    string title = string.Empty;
                    if (node != null)
                    {
                        title = node.InnerText.Trim();
                        if (title.Contains("["))
                        {
                            title = title.Replace("[", "").Replace("]", "");
                        }
                    }
                    if (string.IsNullOrEmpty(title))
                    {
                        continue;
                    }

                    if (title.Contains("상장폐지") || title.Contains("정정 상장폐지"))
                    {
                        if (title.Replace(" ", "").Contains("유가증권시장상장"))
                        {
                            continue;
                        }

                        //For BC drop
                        if (title.Contains("수익증권 상장폐지") && !title.Contains("예고안내"))
                        {
                            Console.WriteLine("BC Time" + nodeDate.InnerText.Trim());
                            DropTemplate item = new DropTemplate();
                            item.Market = "KS";
                            item.RIC = GetTickerByAttribute(node) + "." + item.Market;
                            //HtmlDocument doc = GetTargetPageDocument(node);
                            item.Type = "CEF";
                            HtmlNode companyNode = nodeCollections[i].SelectSingleNode(".//strong/a");
                            GrabIsinInfoForTest(companyNode, item);
                            //GrabDropDataText(doc, item);
                            string targetUrl = GetTargetPageDocumentForTest(node);
                            HtmlDocument doc = WebClientUtil.GetHtmlDocument(targetUrl, 180000);
                            GrabPeoDropDataForTest(doc, item);
                            //InsertToDB(item);
                            continue;
                        }

                        string judge1 = title.Substring(0, ("상장폐지".Length)).Trim().ToString();
                        string judge2 = String.Empty;
                        if (title.Length > "정정 상장폐지".Length)
                        {
                            judge2 = title.Substring(0, ("정정 상장폐지".Length)).Trim().ToString();
                        }
                        //For PEO and ETF drop
                        if (judge1.Equals("상장폐지") || judge2.Equals("정정 상장폐지"))
                        {
                            Console.WriteLine("EE Time" + nodeDate.InnerText.Trim());
                            //Console.WriteLine("EEE_Time: " + nodeDate.InnerText);
                            HtmlNode ddNode = ddCollections[i].SelectSingleNode(".//span");
                            HtmlNode companyNode = nodeCollections[i].SelectSingleNode(".//strong/a");
                            DropTemplate item = new DropTemplate();
                            item.Market = GetDutyCode(ddNode);
                            item.RIC = GetTickerByAttribute(node) + "." + item.Market;
                            //HtmlDocument doc = GetTargetPageDocument(node);
                            string targetUrl = GetTargetPageDocumentForTest(node);

                            item.isRevised = judge2.Equals("정정 상장폐지") ? true : false;
                            GrabIsinInfoForTest(companyNode, item);

                            string ending = targetUrl.Substring(targetUrl.Length - 9);

                            HtmlDocument doc = WebClientUtil.GetHtmlDocument(targetUrl, 180000);
                            GrabPeoDropDataForTest(doc, item);

                        }
                    }
                }
                catch (Exception ex)
                {
                    string msg = string.Format("At GrabDropDetailData(). Error found in grab drop detail information. Error Message:{0}.", ex.Message);
                    errors.Add(msg);
                }
            }
        }

        private void InsertToDB(DropTemplate item)
        {
            KoreaEquityInfo eq = new KoreaEquityInfo();
            eq.RIC = item.RIC;
            eq.ISIN = !string.IsNullOrEmpty(item.ISIN) ? item.ISIN : "TOFIND";
            eq.Type = item.Type;
            eq.FM = "2";
            eq.IDNDisplayName = (idnN++).ToString() + "TESTIDNNAME";
            eq.LegalName = item.LegalName;
            eq.KoreaName = item.KoreaName;
            eq.Ticker = item.RIC.Split('.')[0];
            eq.BcastRef = item.RIC;
            eq.Status = "Active";
            Console.WriteLine("RIC:" + item.RIC);
            KoreaEquityManager.UpdateEquity(eq);
        }

        private void GrabIsinInfoForTest(HtmlNode judgeNode, DropTemplate item)
        {
            try
            {
                string judgeText = judgeNode.Attributes["onclick"].Value.Trim().ToString();
                judgeText = judgeText.Split('(')[1].Split(')')[0].Trim(new char[] { '\'', ' ' }).ToString();
                KoreaCompany company = Common.GetCompanyInfo(judgeText);
                if (company == null)
                {
                    return;
                }
                item.ISIN = company.ISIN;
                item.LegalName = company.LegalName;
                item.KoreaName = company.KoreaName.Replace("(주)", "");
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabIsinInfo(). Error message:{0}", ex.Message);
                errors.Add(msg);
            }
        }


        private string GetTargetPageDocumentForTest(HtmlNode node)
        {
            string url = null;
            try
            {
                string attribute = node.Attributes["onclick"].Value.Trim().ToString();
                if (!string.IsNullOrEmpty(attribute))
                    attribute = attribute.Split('(')[1].Split(')')[0].Trim(new Char[] { ' ', ',', '\'' }).ToString();
                string param = attribute.Split(',')[1].Trim(new Char[] { ' ', '\'', ',' }).ToString();
                url = GetTargetUrl(param);
            }
            catch (Exception ex)
            {
                string msg = "" + ex.StackTrace + "  : ----> \n\r" + ex.ToString();
                errors.Add(msg);
            }
            return url;
        }

        private void GrabPeoDropDataForTest(HtmlDocument doc, DropTemplate item)
        {
            try
            {
                int count = 0;
                if (doc.DocumentNode.SelectNodes("//table") != null)
                {
                    count = doc.DocumentNode.SelectNodes("//table").Count;
                }
                else
                {
                    HtmlNode pre = doc.DocumentNode.SelectSingleNode("//pre");
                    if (pre != null)
                    {
                        string strPre = string.Empty;
                        if (pre != null)
                            strPre = pre.InnerText.ToString().Trim();

                        int effectivePos = strPre.IndexOf("상장폐지일") + ("상장폐지일".Length);
                        string sedate = FormatDataWithPos(effectivePos, strPre);
                        if (!string.IsNullOrEmpty(sedate))
                        {
                            sedate = Convert.ToDateTime(sedate).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                        }

                        item.EffectiveDate = sedate;
                        item.Type = "ETF";
                        InsertToDB(item);
                        return;
                    }
                }


                if (count == 0)
                {
                    string msg = string.Format("At GrabPeoKQKNDropData(). Can not get peo drop detail infomation for RIC:{0}.", item.RIC);
                    errors.Add(msg);
                    return;
                }
                HtmlNode table = doc.DocumentNode.SelectNodes("//table")[(count - 1)];
                HtmlNodeCollection trs = table.SelectNodes("//tr");
                string effectiveDate = string.Empty;
                int tickerPos = -1;
                List<string> rics = new List<string>();
                foreach (HtmlNode tr in trs)
                {
                    HtmlNodeCollection tdNode = tr.SelectNodes(".//td");
                    if (tdNode.Count > 2 && tickerPos == -1)
                    {
                        for (int pos = 1; pos < tdNode.Count; pos++)
                        {
                            HtmlNode tickerNode = tdNode[pos];
                            if (tickerNode.InnerText.Contains("단축코드"))
                            {
                                tickerPos = pos - 1;
                                break;
                            }
                        }
                        if (tickerPos != -1 && tickerPos != 99)
                        {
                            continue;
                        }
                    }

                    if (tickerPos != -1 && tickerPos != 99)
                    {
                        if (tdNode.Count > tickerPos)
                        {
                            string ticker = tdNode[tickerPos].InnerText.Trim();
                            Regex regex = new Regex("[A-Za-z][0-9A-Za-z]+");
                            Match match = regex.Match(ticker);
                            if (match.Success)
                            {
                                rics.Add(ticker.Substring(1));
                                continue;
                            }
                            else
                            {
                                tickerPos = 99;
                            }
                        }
                        else
                        {
                            tickerPos = 99;
                        }
                    }
                    string tdTitle = tr.SelectSingleNode("./td[1]").InnerText.Trim();
                    if (tdTitle.Contains("상장폐지일"))
                    {
                        effectiveDate = tr.SelectSingleNode("./td[2]").InnerText.Trim();
                        break;
                    }
                }
                if (!string.IsNullOrEmpty(effectiveDate))
                {
                    foreach (string ric in rics)
                    {
                        DropTemplate dItem = new DropTemplate();
                        dItem.RIC = ric + "." + item.Market;
                        dItem.EffectiveDate = effectiveDate;
                        dItem.Market = item.Market;
                        dItem.KoreaName = item.KoreaName;
                        dItem.LegalName = item.LegalName;

                        if (ric.Substring(5) == "5")
                        {
                            dItem.Type = "PRF";
                        }
                        else
                        {
                            dItem.ISIN = item.ISIN;
                            dItem.Type = "ORD";
                        }
                        InsertToDB(dItem);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GrabPeoKQKNDropData(). Error found in grab drop detail information for RIC:{0}. Error Message:{1}.", item.RIC, ex.Message);
                errors.Add(msg);
            }
        }
    }
}
