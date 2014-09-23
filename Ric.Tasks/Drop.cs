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
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks
{
    public class Drop : GeneratorBase
    {
        private List<DateTime> holidayList;
        private static readonly string HOLIDAY_LIST_FILE_PATH = ".\\Config\\Korea\\Holiday.xml";
        private KOREA_DROP_Config configObj;

        private String searchUri = "http://kind.krx.co.kr/disclosure/searchtotalinfo.do";
        private String postData;

        private Hashtable cwlistingHash = new Hashtable();
        private Hashtable kskqlistingRicHash = new Hashtable();
        private Hashtable kskqlistingIsinHash = new Hashtable();
        private Hashtable etflistingHash = new Hashtable();
        private List<DropTemplate> wdropList = new List<DropTemplate>();    //Company Warrant Drop
        private List<DropTemplate> edropList = new List<DropTemplate>();    //Equity Drop
        private List<DropTemplate> bdropList = new List<DropTemplate>();    //BC Drop
        private List<DropTemplate> etfdropList = new List<DropTemplate>();  //ETF Drop

        protected override void Start()
        {
            StartDropJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            holidayList = ConfigUtil.ReadConfig(HOLIDAY_LIST_FILE_PATH, typeof(List<DateTime>)) as List<DateTime>;
            configObj = Config as KOREA_DROP_Config;
        }

        public void StartDropJob()
        {
            String startDate = configObj.KoreaDropStartDate;
            String endDate = configObj.KoreaDropEndDate;
            if (String.IsNullOrEmpty(startDate)) startDate = DateTime.Today.ToString("yyyy-MM-dd");
            if (String.IsNullOrEmpty(endDate)) endDate = DateTime.Today.ToString("yyyy-MM-dd");

            //the first date in postData means the date start from and the second date means the date to the end  so need use the user's input as the first date and use the today to as the second date
            postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%83%81%EC%9E%A5%ED%8F%90%EC%A7%80&repIsuSrtCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=300&scn=mktact&srchFd=2&kwd=%EC%83%81%EC%9E%A5%ED%8F%90%EC%A7%80&fromData={0}&toData={1}", startDate, endDate);
            ReadDataFromExcel();
            SearchTheWebpageToGrabData();
            GrabDataForCompanyWarrantDropFromISINWebpage();
            GetISINFromKSorKQListingItemsList();
            SearchLegalNameFromISINWebpage();
            GenerateDropFile_xls();
        }

        private void ReadDataFromExcel()
        {
            ReadKQorKSListingItemsList();
            ReadCompanyWarrantListingItemsList();
            ReadETFListingItemsList();
        }

        private void ReadCompanyWarrantListingItemsList()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            try
            {
                if (excelApp.ExcelAppInstance == null)
                {
                    String msg = "Excel could not be created ! please check your office installation and refence correct.";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                String ipath = configObj.KoreaDropCompanyWarrantReadFileConfig.WorkbookPath;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = ExcelUtil.GetWorksheet(configObj.KoreaDropCompanyWarrantReadFileConfig.WorksheetName, wBook);
                if (wSheet == null)
                {
                    String msg = "Excel Worksheet could not be created ! please check your office installation and refence correct.";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                int startLine = 2;
                while (wSheet.Range["A" + startLine, Type.Missing].Value2 != null && wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                {
                    CompanyWarrantList cwlisting = new CompanyWarrantList
                    {
                        Ric = wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim()
                    };
                    if (wSheet.Range["B" + startLine, Type.Missing].Value2 != null && wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                        cwlisting.Display_Name = wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim();
                    if (wSheet.Range["C" + startLine, Type.Missing].Value2 != null && wSheet.Range["C" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                        cwlisting.ISIN = wSheet.Range["C" + startLine, Type.Missing].Value2.ToString().Trim();
                    if (wSheet.Range["D" + startLine, Type.Missing].Value2 != null && wSheet.Range["D" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                        cwlisting.Conversion_ratio = wSheet.Range["D" + startLine, Type.Missing].Value2.ToString().Trim();
                    if (wSheet.Range["E" + startLine, Type.Missing].Value2 != null && wSheet.Range["E" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                        cwlisting.Exercise_Price = wSheet.Range["E" + startLine, Type.Missing].Value2.ToString().Trim();

                    if (!cwlistingHash.Contains(cwlisting.ISIN))
                        cwlistingHash.Add(cwlisting.ISIN, cwlisting);
                    startLine++;
                }
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                String msg = "Error found in ReadCompanyWarrantListingItemsList()    : \r\n" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void ReadKQorKSListingItemsList()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            try
            {
                if (excelApp.ExcelAppInstance == null)
                {
                    String msg = "Excel could not be created . Please check your office installation and refence corrected!";
                    Logger.Log(msg, Logger.LogType.Error);
                }

                String ipath = configObj.KoreaDropKQorKsListReadFilePathConfig.WorkbookPath;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = ExcelUtil.GetWorksheet(configObj.KoreaDropKQorKsListReadFilePathConfig.WorksheetName, wBook);
                if (wSheet == null)
                {
                    String msg = "Excel Worksheet could not be created . Please check your office installation and refence corrected!";
                    Logger.Log(msg, Logger.LogType.Error);
                }

                int startLine = 2;
                while (wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null && wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                {
                    KSorKQListingList listing = new KSorKQListingList
                    {
                        Ric = wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim()
                    };
                    if (wSheet.Range["B" + startLine, Type.Missing].Value2 != null && wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                        listing.IDNDisplayName = wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim();
                    if (wSheet.Range["C" + startLine, Type.Missing].Value2 != null && wSheet.Range["C" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                        listing.ISIN = wSheet.Range["C" + startLine, Type.Missing].Value2.ToString().Trim();

                    if (!kskqlistingRicHash.Contains(listing.Ric))
                        kskqlistingRicHash.Add(listing.Ric, listing);
                    if (!kskqlistingIsinHash.Contains(listing.ISIN))
                        kskqlistingIsinHash.Add(listing.ISIN, listing);
                    startLine++;
                }
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                String msg = "Error found in Read data from ETF Listing Items list      :\r\n" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void ReadETFListingItemsList()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            try
            {
                if (excelApp.ExcelAppInstance == null)
                {
                    String msg = "Excel could not be created . Please check your office installtion and refence correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                String ipath = configObj.KoreaDropEtfListingItemsReadFilePathConfig.WorkbookPath;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = ExcelUtil.GetWorksheet(configObj.KoreaDropEtfListingItemsReadFilePathConfig.WorksheetName, wBook);
                if (wSheet == null)
                {
                    String msg = "Excel Worksheet could not be created . Please check your office installtion and refence correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                }

                int startLine = 2;
                while (wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null && wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                {
                    ETFListingList listing = new ETFListingList
                    {
                        RIC = wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim()
                    };
                    if (wSheet.Range["B" + startLine, Type.Missing].Value2 != null && wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                        listing.IDNDisplayName = wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim();
                    if (wSheet.Range["C" + startLine, Type.Missing].Value2 != null && wSheet.Range["C" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                        listing.ISIN = wSheet.Range["C" + startLine, Type.Missing].Value2.ToString().Trim();

                    if (!String.IsNullOrEmpty(listing.RIC) && !etflistingHash.Contains(listing.RIC))
                        etflistingHash.Add(listing.RIC, listing);
                    startLine++;
                }
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                String msg = "Error found in Read data from ETF Listing Items list      :\r\n" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void SearchTheWebpageToGrabData()
        {
            HtmlDocument htc = new HtmlDocument();
            var content = WebClientUtil.GetDynamicPageSource(searchUri, 300000, postData);
            if (!String.IsNullOrEmpty(content))
                htc.LoadHtml(content);

            HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes("//dt");
            try
            {
                if (nodeCollections.Count > 0)
                {
                    foreach (HtmlNode selectNode in nodeCollections)
                    {
                        HtmlNode node = selectNode.SelectSingleNode(".//span[@class='subject']/a");
                        String tnode = String.Empty;
                        if (node != null)
                            tnode = node.InnerText;

                        tnode = tnode.Contains("[") ? tnode.Replace("[", "").Replace("]", "").Trim() : tnode;
                        //Company Warrant Drop
                        if (tnode.Contains("신주인수권증권 상장폐지"))
                        {
                            String judge = tnode.Substring(0, "신주인수권증권 상장폐지".Length).Trim();
                            if (judge.Equals("신주인수권증권 상장폐지"))
                            {
                                DropTemplate wdrop = new DropTemplate();
                                HtmlNode header = selectNode.SelectSingleNode(".//strong/a");
                                String attribute = String.Empty;
                                if (header != null)
                                    attribute = header.Attributes["onclick"].Value.Trim();
                                if (!String.IsNullOrEmpty(attribute))
                                    attribute = attribute.Split('(')[1].Split(')')[0].Trim(new[] { ' ', '\'', ';' });
                                String str_uri = String.Format("http://kind.krx.co.kr/common/companysummary.do?method=searchCompanySummary&strIsurCd={0}&lstCd=undefined", attribute);
                                String KsOrKq = String.Empty;
                                HtmlDocument doc = WebClientUtil.GetHtmlDocument(str_uri, 120000, null);
                                if (doc != null)
                                {
                                    HtmlNode docnode = doc.DocumentNode.SelectSingleNode("//div[@id='pContents']/table/tbody/tr[2]/td[2]");
                                    if (docnode != null)
                                        KsOrKq = docnode.InnerText;
                                }

                                String parameters = node.Attributes["onclick"].Value.Trim();
                                parameters = parameters.Split('(')[1].Split(')')[0].Trim(new[] { ' ', '\'', ';' });
                                String param1 = parameters.Split(',')[0].Trim(new[] { ' ', '\'', ',' });
                                String param2 = parameters.Split(',')[1].Trim(new[] {' ', '\'', ','});
                                String uri = String.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno={1}&viewerhost=&viewerport=", param1, param2);

                                doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                                String ticker = String.Empty;
                                if (doc != null)
                                    ticker = doc.DocumentNode.SelectSingleNode("//div[@id='pHeader']/h2").InnerText;
                                if (!String.IsNullOrEmpty(ticker))
                                    ticker = ticker.Split('(')[1].Trim(new[] { ' ', ')', '(' });

                                String param3 = KsOrKq.Equals("유가증권") ? "68913" : (KsOrKq.Equals("코스닥") ? "70925" : null);
                                if (String.IsNullOrEmpty(param3))
                                    return;
                                param1 = param1.Insert(4, "/").Insert(7, "/").Insert(10, "/");
                                uri = String.Format("http://kind.krx.co.kr/external/{0}/{1}/{2}.htm", param1, param2, param3);

                                doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                                if (doc == null)
                                    return;
                                // KQ
                                if (KsOrKq.Equals("코스닥"))
                                {
                                    HtmlNode koreaName = doc.DocumentNode.SelectSingleNode("//tr[1]/td[2]");
                                    HtmlNode effective = doc.DocumentNode.SelectSingleNode("//tr[5]/td[2]");
                                    String kname = String.Empty;
                                    String edate = String.Empty;
                                    if (koreaName != null)
                                        kname = koreaName.InnerText.Trim();
                                    if (effective != null)
                                        edate = effective.InnerText.Trim();
                                    kname = kname.Trim();
                                    if (!String.IsNullOrEmpty(kname))
                                        kname = kname.Contains("(주)") ? kname.Replace("(주)", "").Trim() : kname;
                                    edate = edate.Trim();
                                    if (!String.IsNullOrEmpty(edate))
                                        edate = Convert.ToDateTime(edate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                                    wdrop.KoreaName = kname;
                                    wdrop.EffectiveDate = edate;
                                    wdrop.RIC = ticker + ".KQ";
                                    wdropList.Add(wdrop);
                                }//KS
                                else if (KsOrKq.Equals("유가증권"))
                                {
                                    string skName = MiscUtil.GetCleanTextFromHtml(doc.DocumentNode.SelectSingleNode("//tr[1]/td[2]").InnerText);
                                    skName = Regex.Split(skName, "   ", RegexOptions.IgnoreCase)[0].Trim(new[] { ' ', ':' });

                                    DateTime dt = DateTime.Now;

                                    string seDate=MiscUtil.GetCleanTextFromHtml(doc.DocumentNode.SelectSingleNode("//tr[8]/td[2]").InnerText);

                                    wdrop.KoreaName = skName;
                                    wdrop.EffectiveDate = seDate;
                                    wdrop.RIC = ticker + ".KS";
                                    wdropList.Add(wdrop);
                                }//Error
                                else
                                {
                                    Logger.Log("Get the wrong node innerText .");
                                }
                            }
                        }//Equity Drop
                        else if (tnode.Contains("상장폐지") || tnode.Contains("정정 상장폐지"))
                        {
                            if (tnode.Replace(" ", "").Contains("유가증권시장상장"))
                                continue;
                            String judge1 = tnode.Substring(0, ("상장폐지".Length)).Trim();
                            String judge2 = String.Empty;
                            if (tnode.Length > "정정 상장폐지".Length)
                                judge2 = tnode.Substring(0, ("정정 상장폐지".Length)).Trim();
                            if (judge1.Equals("상장폐지") || judge2.Equals("정정 상장폐지"))
                            {
                                DropTemplate edrop = new DropTemplate
                                {
                                    isRevised = judge2.Equals("정정 상장폐지")
                                };
                                HtmlNode header = selectNode.SelectSingleNode(".//strong/a");
                                String attribute = String.Empty;
                                if (header != null)
                                    attribute = header.Attributes["onclick"].Value.Trim();
                                if (!String.IsNullOrEmpty(attribute))
                                    attribute = attribute.Split('(')[1].Split(')')[0].Trim(new[] { ' ', '\'', ';' });
                                String str_uri = String.Format("http://kind.krx.co.kr/common/companysummary.do?method=searchCompanySummary&strIsurCd={0}&lstCd=undefined", attribute);
                                String KsOrKq = String.Empty;
                                HtmlDocument doc = WebClientUtil.GetHtmlDocument(str_uri, 120000, null);
                                if (doc != null)
                                {
                                    HtmlNode docnode = doc.DocumentNode.SelectSingleNode("//div[@id='pContents']/table/tbody/tr[2]/td[2]");
                                    if (docnode != null)
                                        KsOrKq = docnode.InnerText;
                                }

                                String parameters = node.Attributes["onclick"].Value.Trim();
                                parameters = parameters.Split('(')[1].Split(')')[0].Trim(new[] { ' ', '\'', ';' });
                                String param1 = parameters.Split(',')[0].Trim(new[] { ' ', '\'', ',' });
                                String param2 = parameters.Split(',')[1].Trim(new[] { ' ', '\'', ',' });
                                String uri = String.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno={1}&viewerhost=&viewerport=", param1, param2);
                                doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                                String ticker = String.Empty;
                                if (doc != null)
                                    ticker = doc.DocumentNode.SelectSingleNode("//div[@id='pHeader']/h2").InnerText;
                                if (!String.IsNullOrEmpty(ticker))
                                    ticker = ticker.Split('(')[1].Trim(new[] { ' ', ')', '(' });

                                String param3 = KsOrKq.Equals("유가증권") ? "68051" : (KsOrKq.Equals("코스닥") ? "70769" : null);
                                if (String.IsNullOrEmpty(param3))
                                    return;
                                param1 = param1.Insert(4, "/").Insert(7, "/").Insert(10, "/");
                                uri = String.Format("http://kind.krx.co.kr/external/{0}/{1}/{2}.htm", param1, param2, param3);

                                doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                                if (doc == null)
                                    return;

                                // KQ
                                if (KsOrKq.Equals("코스닥"))
                                {
                                    int count = 0;
                                    if (doc.DocumentNode.SelectNodes("//table") != null)
                                        count = doc.DocumentNode.SelectNodes("//table").Count;
                                    HtmlNode effective = doc.DocumentNode.SelectNodes("//table")[(count - 1)].SelectSingleNode("//tr[9]/td[2]");
                                    String edate = String.Empty;
                                    if (effective != null)
                                        edate = effective.InnerText.Trim();
                                    edate = Convert.ToDateTime(edate).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                                    edrop.EffectiveDate = edate;
                                    edrop.RIC = ticker + ".KQ";
                                    edropList.Add(edrop);
                                }//KS
                                else if (KsOrKq.Equals("유가증권"))
                                {
                                    HtmlNode pre = doc.DocumentNode.SelectSingleNode("//pre");
                                    String str_pre = String.Empty;
                                    if (pre != null)
                                        str_pre = pre.InnerText.Trim();

                                    int effective_pos = str_pre.IndexOf("상장폐지일 ") + ("상장폐지일 ".Length);
                                    String sedate = VarFormat(effective_pos, str_pre);
                                    if (!String.IsNullOrEmpty(sedate))
                                        sedate = Convert.ToDateTime(sedate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));

                                    edrop.EffectiveDate = sedate;
                                    edrop.RIC = ticker + ".KS";
                                    edropList.Add(edrop);
                                }//Error
                                else
                                {
                                    Logger.Log("Get the wrong node innerText .");
                                }
                            }
                        }    //BC Drop
                        else if (tnode.Contains("수익증권 상장폐지") && !tnode.Contains("예고안내"))
                        {
                            String judge = tnode.Substring(0, ("수익증권 상장폐지".Length)).Trim();
                            if (judge.Equals("수익증권 상장폐지"))
                            {
                                DropTemplate bdrop = new DropTemplate();
                                String parameters = node.Attributes["onclick"].Value.Trim();
                                parameters = parameters.Split('(')[1].Split(')')[0].Trim(new[] { ' ', '\'', ';' });
                                String param1 = parameters.Split(',')[0].Trim(new[] { ' ', '\'', ',' });
                                String param2 = parameters.Split(',')[1].Trim(new[] { ' ', '\'', ',' });

                                param1 = param1.Insert(4, "/").Insert(7, "/").Insert(10, "/");
                                String uri = String.Format("http://kind.krx.co.kr/external/{0}/{1}/68909.htm", param1, param2);
                                HtmlDocument doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                                if (doc == null) continue;
                                HtmlNode pre = doc.DocumentNode.SelectSingleNode("//pre");
                                String str_pre = String.Empty;
                                if (pre != null)
                                    str_pre = pre.InnerText.Trim();

                                int effective_pos = str_pre.IndexOf("상장폐지일 ") + ("상장폐지일 ".Length);
                                String sedate = VarFormat(effective_pos, str_pre);
                                if (!String.IsNullOrEmpty(sedate))
                                    sedate = Convert.ToDateTime(sedate).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();

                                int isin_pos = str_pre.IndexOf("표준코드") + ("표준코드".Length);
                                String sisin = String.Empty;
                                if (isin_pos < "표준코드".Length)
                                {
                                    isin_pos = str_pre.IndexOf("코");
                                    sisin = VarFormat(isin_pos, str_pre);
                                    sisin = sisin.Replace(" ", "").Split(':')[0].Equals("코드") ? sisin.Split(':')[1].Trim() : null;
                                }
                                else
                                    sisin = VarFormat(isin_pos, str_pre);

                                bdrop.ISIN = sisin;
                                bdrop.EffectiveDate = sedate;
                                bdropList.Add(bdrop);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in SearchTheWebpageToGrabData()      : \r\n" + ex.InnerException + "  :  \r\n" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private String VarFormat(int pos, String pre)
        {
            String result = String.Empty;
            try
            {
                Char[] arr = pre.ToCharArray();
                while (arr[pos] != '\n')
                {
                    result += arr[pos].ToString();
                    if ((pos + 1) <= arr.Length)
                        pos++;
                    else break;
                }
                result = result.Trim(new[] { ' ', ':' });
            }
            catch (Exception ex)
            {
                String msg = "Error found in VarFormat()    : \r\n" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
            return result;
        }

        private void GrabDataForCompanyWarrantDropFromISINWebpage()
        {
            try
            {
                foreach (var item in wdropList)
                {
                    ISINQuery query = new ISINQuery("", "", "06", "", item.KoreaName);
                    List<ISINTemp> isinList = Common.getISINListFromISINWebPage(query);
                    if (isinList == null || isinList.Count == 0)
                    {
                        Logger.Log(string.Format("Cannot find result in ISIN webpage"));
                    }

                    else if (isinList.Count > 1)
                    {
                        Logger.Log(string.Format("Find two results in ISIN webpage. Choose the first one."));
                    }

                    else
                    {
                        HtmlDocument doc=WebClientUtil.GetHtmlDocument(isinList[0].ISINLink, 300000, null);
                        String slegalname = String.Empty;
                        if (doc != null)
                        {
                            HtmlNode legalname = doc.DocumentNode.SelectSingleNode(".//tr[6]/td[2]");
                            if (legalname != null)
                                slegalname = legalname.InnerText.Trim().ToString();
                        }
                        item.LegalName = slegalname;
                        item.ISIN = isinList[0].ISIN;
                    }
 
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in  GrabDataForCompanyWarrantDropFromISINWebpage()     : \r\n" + ex.InnerException + "  :  \r\n" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GetISINFromKSorKQListingItemsList()
        {
            if (edropList.Count > 0)
            {
                foreach (var item in edropList)
                {
                    if (kskqlistingRicHash.Contains(item.RIC))
                    {
                        var listing = kskqlistingRicHash[item.RIC] as KSorKQListingList;
                        item.ISIN = listing.ISIN;
                        item.QAShortName = listing.IDNDisplayName;
                    }
                    if (etflistingHash.Contains(item.RIC))
                    {
                        var listing = etflistingHash[item.RIC] as ETFListingList;
                        DropTemplate drop = new DropTemplate
                        {
                            RIC = item.RIC,
                            ISIN = listing.ISIN,
                            QAShortName = listing.IDNDisplayName
                        };
                        etfdropList.Add(drop);
                    }
                }
            }
            if (bdropList.Count > 0)
            {
                foreach (var item in bdropList)
                {
                    if (kskqlistingIsinHash.Contains(item.ISIN))
                    {
                        var listing = kskqlistingIsinHash[item.ISIN] as KSorKQListingList;
                        item.RIC = listing.Ric;
                        item.QAShortName = listing.IDNDisplayName;
                    }
                }
            }
            if (wdropList.Count > 0)
            {
                foreach (var item in wdropList)
                {
                    if (cwlistingHash.Contains(item.ISIN))
                    {
                        var listing = cwlistingHash[item.ISIN] as CompanyWarrantList;
                        item.RIC = listing.Ric;
                        item.QAShortName = listing.Display_Name;
                    }
                }
            }
        }

        private void SearchLegalNameFromISINWebpage()
        {
            try
            {
                //Legal Name search key word<종목영문명>
                /*     Equity Drop uri      */
                //http://isin.krx.co.kr/jsp/BA_VW010.jsp?isu_cd=KR7026220004&modi=f&req_no=
                //http://isin.krx.co.kr/jsp/BA_VW010.jsp?isu_cd=KR7039130000&modi=f&req_no=

                /*     BC Drop uri      */
                //http://isin.krx.co.kr/jsp/BA_VW012.jsp?isu_cd=KR5701016SB0&modi=f&req_no=
                //http://isin.krx.co.kr/jsp/BA_VW012.jsp?isu_cd=KR574001AV58&modi=f&req_no=
                HtmlDocument doc = new HtmlDocument();
                String uri = String.Empty;
                // Equity Drop Legal Name
                if (edropList.Count > 0)
                {
                    foreach (var item in edropList.Where(item => !String.IsNullOrEmpty(item.ISIN)))
                    {
                        uri = String.Format("http://isin.krx.co.kr/jsp/BA_VW010.jsp?isu_cd={0}&modi=f&req_no=", item.ISIN);
                        doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                        if (doc != null)
                            item.LegalName = doc.DocumentNode.SelectSingleNode("//tr[11]//td[2]").InnerText.Trim();
                    }
                }
                //ETF Drop Legal Name   :   select all
                if (etfdropList.Count > 0)
                {
                    foreach (var item in etfdropList.Where(item => !String.IsNullOrEmpty(item.ISIN)))
                    {
                        uri = String.Format("http://isin.krx.co.kr/jsp/BA_VW010.jsp?isu_cd={0}&modi=f&req_no=", item.ISIN);
                        doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                        if (doc != null)
                            item.LegalName = doc.DocumentNode.SelectSingleNode("//tr[11]//td[2]").InnerText.Trim();
                    }
                }
                //BC Drop Legal Name
                if (bdropList.Count > 0)
                {
                    foreach (var item in bdropList.Where(item => !String.IsNullOrEmpty(item.ISIN)))
                    {
                        uri = String.Format("http://isin.krx.co.kr/jsp/BA_VW012.jsp?isu_cd={0}&modi=f&req_no=", item.ISIN);
                        doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                        if (doc != null)
                            item.LegalName = doc.DocumentNode.SelectSingleNode("//tr[5]//td[2]").InnerText.Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in SearchLegalNameFromISINWebpage()   : \r\n" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GenerateDropFile_xls()
        {
            GenerateDropFile_xls(wdropList, "wdrop");
            GenerateDropFile_xls(edropList, "edrop");
            GenerateDropFile_xls(bdropList, "bdrop");
            GenerateDropFile_xls(etfdropList, "etfdrop");
            generateNDAandGEDAFileForCWChgDrop(wdropList);
        }

        private void GenerateDropFile_xls(ICollection<DropTemplate> list, String droptype)
        {
            if (list.Count > 0)
            {
                foreach (var item in list)
                {
                    if (String.IsNullOrEmpty(item.ISIN)) continue;
                    ExcelApp excelApp = new ExcelApp(false, false);
                    if (excelApp.ExcelAppInstance == null) { }
                    String filename = string.Empty;
                    String ipath = string.Empty;
                    try
                    {
                        switch (droptype)
                        {
                            case "wdrop":  //Company Warrant DROP
                                filename = "KR FM (Company Warrant Drop) Request_" + item.RIC + " (wef " + item.EffectiveDate + ").xls";
                                ipath = configObj.KoreaCompanyWarrantDropGeneratorConfig.WORKBOOK_PATH + filename;
                                break;
                            case "edrop":  //Equity DROP   KR FM (Drop) Request _ 052810.KQ (wef 2011-Nov-22).xls
                                filename = "KR FM (Drop) Request_" + item.RIC + " (wef " + item.EffectiveDate + ").xls";
                                ipath = configObj.KoreaEquityDropGeneratorConfig.WORKBOOK_PATH + filename;
                                break;
                            case "etfdrop":  //KR FM (ETF Drop) Request_110550.KS, 124090.KS(wef 2011-Dec-19).xls
                                filename = "KR FM (ETF Drop) Request_" + item.RIC + " (wef " + item.EffectiveDate + ").xls";
                                ipath = configObj.KoreaEtfDropGeneratorConfig.WORKBOOK_PATH + filename;
                                break;
                            case "bdrop":  //BC DROP
                                filename = "KR FM (BC Drop) Request_" + item.RIC + " (wef " + item.EffectiveDate + ").xls";
                                ipath = configObj.KoreaBcDropGeneratorConfig.WORKBOOK_PATH + filename;
                                break;
                        }
                        if (item.isRevised)
                            filename = "(Revised) " + filename;

                        Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                        Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet1", wBook);
                        if (wSheet == null) 
                        {
                            Logger.LogErrorAndRaiseException(string.Format("There's no such worksheet {0}in workbook {1}","Sheet1", wBook.FullName));
                        }
                        wSheet.Name = "DROP";

                        ((Range)wSheet.Columns["A", Type.Missing]).ColumnWidth = 20;
                        ((Range)wSheet.Columns["B", Type.Missing]).ColumnWidth = 2;
                        ((Range)wSheet.Columns["C", Type.Missing]).ColumnWidth = 30;
                        ((Range)wSheet.Columns["A:C", Type.Missing]).Font.Name = "Arial";

                        ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = FontStyle.Bold;
                        ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = ColorTranslator.ToOle(Color.Black);

                        wSheet.Cells[1, 1] = "FM Request";
                        wSheet.Cells[1, 2] = " ";

                        wSheet.Cells[1, 3] = droptype.Equals("edrop") ? "Equity Deletion" : "Deletion";

                        ((Range)wSheet.Cells[3, 1]).Font.Bold = FontStyle.Bold;
                        ((Range)wSheet.Cells[3, 1]).Font.Color = ColorTranslator.ToOle(Color.Black);
                        wSheet.Cells[3, 1] = "Effective Date";
                        wSheet.Cells[3, 2] = ":";
                        ((Range)wSheet.Cells[3, 3]).NumberFormat = "@";
                        wSheet.Cells[3, 3] = item.EffectiveDate;
                        ((Range)wSheet.Cells[4, 1]).Font.Bold = FontStyle.Bold;
                        ((Range)wSheet.Cells[4, 1]).Font.Color = ColorTranslator.ToOle(Color.Black);
                        wSheet.Cells[4, 1] = "RIC";
                        wSheet.Cells[4, 2] = ":";
                        wSheet.Cells[4, 3] = item.RIC;
                        ((Range)wSheet.Cells[5, 1]).Font.Bold = FontStyle.Bold;
                        ((Range)wSheet.Cells[5, 1]).Font.Color = ColorTranslator.ToOle(Color.Black);
                        wSheet.Cells[5, 1] = "ISIN";
                        wSheet.Cells[5, 2] = ":";
                        wSheet.Cells[5, 3] = item.ISIN;
                        wSheet.Cells[6, 1] = "QA Short Name";
                        wSheet.Cells[6, 2] = ":";
                        ((Range)wSheet.Cells[6, 3]).Font.Color = ColorTranslator.ToOle(Color.Blue);
                        ((Range)wSheet.Cells[6, 3]).Font.Underline = true;
                        wSheet.Cells[6, 3] = item.QAShortName;

                        if (!droptype.Equals("wdrop"))
                        {
                            wSheet.Cells[7, 1] = "Legal Name";
                            wSheet.Cells[7, 2] = ":";
                            wSheet.Cells[7, 3] = item.LegalName;
                            if (droptype.Equals("edrop"))
                            {
                                wSheet.Cells[8, 1] = "ICW Index Drops";
                                wSheet.Cells[8, 2] = ":";
                                wSheet.Cells[8, 3] = "";
                            }
                        }
                        excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                        wBook.Save();

                    }
                    catch (Exception ex)
                    {
                        String msg = "Error found in GenerateDropFile_xls()      : \r\n" + ex;
                        Logger.Log(msg, Logger.LogType.Error);
                    }
                    finally
                    {
                        excelApp.Dispose();
                    }
                }
            }
        }

        #region  //add by jackson to generate CWChg NDA and GEDA file
        
        private void generateNDAandGEDAFileForCWChgDrop(IEnumerable<DropTemplate> list)
        {
            try
            {
                createDirectory(configObj.KoreaCompanyWarrantDropGeneratorConfig.WORKBOOK_PATH);
                createDirectory(configObj.KoreaCompanyWarrantDropGeneratorConfig.GEDA_FILE_PATH);
                createDirectory(configObj.KoreaCompanyWarrantDropGeneratorConfig.NDA_FILE_PATH);
                foreach (DropTemplate template in list)
                {
                    createGEDAFileForCWChgDrop(template,configObj.KoreaCompanyWarrantDropGeneratorConfig.GEDA_FILE_PATH);
                    createNDAQAFileForCWChgDrop(template,configObj.KoreaCompanyWarrantDropGeneratorConfig.NDA_FILE_PATH);
                }
                upLoadNDAFile(configObj.KoreaCompanyWarrantDropGeneratorConfig.NDA_FILE_PATH);
                upLoadGEDAFile(configObj.KoreaCompanyWarrantDropGeneratorConfig.GEDA_FILE_PATH);
                sendFMFile(configObj.KoreaCompanyWarrantDropGeneratorConfig.WORKBOOK_PATH);
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }
        }
        private void createNDAQAFileForCWChgDrop(DropTemplate dInput,string dir)
        {
            createDirectory(dir);
            DateTime effectiveDate;
            effectiveDate = DateTime.ParseExact(dInput.EffectiveDate, dInput.EffectiveDate.Length == 10 ? "yyyy-MM-dd" : "yyyy-MMM-dd", new CultureInfo("en-US"));
            DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
            string fileName_QA = "KR_CWRTS" + theDayBeforeEffictiveDate.ToString("yyyyMMdd") + "QAChg.csv";
            string pathNDA_QA = Path.Combine(dir, fileName_QA);
            try
            {
                if (!File.Exists(@pathNDA_QA))
                {
                    StringBuilder strData = new StringBuilder();
                    strData.Append("RIC,EFFICTIVE DATE\r\n");
                    string temp = dInput.RIC;
                    int j;
                    for (j = 0; j < temp.Length; j++)
                    {
                        if (temp[j] == '.')
                            break;
                    }
                    string tempS = temp.Substring(0, j);
                    tempS = tempS + "F";
                    for (int k = j; k < temp.Length; k++)
                    {
                        tempS = tempS + temp[k];
                    }
                   
                    strData.Append(dInput.RIC + "," + effectiveDate.ToString("yyyy-MMM-dd") + "\r\n");
                    strData.Append(tempS + "," + effectiveDate.ToString("yyyy-MMM-dd") + "\r\n");
                    File.WriteAllText(pathNDA_QA, strData.ToString(), Encoding.UTF8);
                }
                else
                {
                    StreamReader readFileAll = new StreamReader(pathNDA_QA);
                    string strDataAll = readFileAll.ReadToEnd();
                    readFileAll.Close();
                    StreamReader readFile = new StreamReader(pathNDA_QA);
                    bool exist = false;
                    StringBuilder strData = new StringBuilder(strDataAll);
                    string line = readFile.ReadLine();
                    while (line != null)
                    {
                        string temp = line.Split(',')[0];
                        if (temp == dInput.RIC)
                        {
                            exist = true;
                        }
                        line = readFile.ReadLine();
                    }
                    readFile.Close();
                    if (exist == false)
                    {
                        StreamWriter writeFile = new StreamWriter(pathNDA_QA);
                        string temp = dInput.RIC;
                        int j;
                        for (j = 0; j < temp.Length; j++)
                        {
                            if (temp[j] == '.')
                                break;
                        }
                        string tempS = temp.Substring(0, j);
                        tempS = tempS + "F";
                        for (int k = j; k < temp.Length; k++)
                        {
                            tempS = tempS + temp[k];
                        }
                        strData.Append(dInput.RIC + "," + effectiveDate.ToString("yyyy-MMM-dd") + "\r\n");
                        strData.Append(tempS + "," + effectiveDate.ToString("yyyy-MMM-dd") + "\r\n");
                        writeFile.Write(strData);
                        writeFile.Close();
                    }

                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when trying to create the NDA . Ex: {0}", ex.Message));
            }
        }
        private void createGEDAFileForCWChgDrop(DropTemplate dInput,string dir)
        {
            createDirectory(dir);

            DateTime effectiveDate = DateTime.ParseExact(dInput.EffectiveDate, dInput.EffectiveDate.Length == 10 ? "yyyy-MM-dd" : "yyyy-MMM-dd", new CultureInfo("en-US"));
            DateTime theDayBeforeEffictiveDate = MiscUtil.GetLastTradingDay(effectiveDate, holidayList, 1);
            string fileName = "KR_CWRTS_Bulk_Change_Drop_" + theDayBeforeEffictiveDate.ToString("yyyyMMdd") + ".txt";
            string path = Path.Combine(dir, fileName);
            try
            {
                if (!File.Exists(@path))
                {
                    StringBuilder strData = new StringBuilder();
                    strData.Append("RIC\r\n");
                    strData.Append(dInput.RIC);
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
                    StringBuilder strData = new StringBuilder(strDataAll);
                    string line = readFile.ReadLine();
                    while (line != null)
                    {
                        string temp = line.Split('\t')[0];
                        if (temp == dInput.RIC)
                        {
                            exist = true;
                        }
                        line = readFile.ReadLine();
                    }
                    readFile.Close();
                    if (exist == false)
                    {
                        StreamWriter writeFile = new StreamWriter(path);
                        strData.Append(dInput.RIC);
                        strData.Append("\r\n");
                        writeFile.Write(strData);
                        writeFile.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when trying to create the GEDA . Ex: {0}", ex.Message));
            }
        }
        private void createDirectory(string directory)
        {
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }
        private void sendFMFile(string emaFileDir)
        {
            IEnumerable<string> fileList = GetAllTheEMAFile(emaFileDir);
            if (fileList == null)
            {
                return;
            }
            string mailSubject;
            string inscrubed = "";
            inscrubed = inscrubed + "\r\n\r\n\r\n";

            inscrubed = configObj.AlertMailSignatureInformationList.Aggregate(inscrubed, (current, t) => current + t + "\r\n");

            foreach (string file in from file in fileList let fileName = Path.GetFileName(file) where !fileName.Contains("(Revised)") let fi = new FileInfo(file) where fi.CreationTime.Date.Equals(DateTime.Today) select file)
            {
                mailSubject = Path.GetFileNameWithoutExtension(file);
                MailToSend mail = new MailToSend();
                mail.ToReceiverList.AddRange(configObj.AlertMailToList);

                mail.CCReceiverList.AddRange(configObj.AlertMailCCList);
                mail.MailSubject = mailSubject;
                mail.AttachFileList.Add(file);
                string changType = mailSubject.Split('(')[1];
                changType = changType.Split(')')[0];
                string effectiveDate = mailSubject.Split('(')[2];
                effectiveDate = effectiveDate.Split(')')[0];
                effectiveDate = effectiveDate.Split(' ')[1];
                DateTime effectiveD = DateTime.ParseExact(effectiveDate, effectiveDate.Length == 10 ? "yyyy-MM-dd" : "yyyy-MMM-dd", null);
                effectiveDate = effectiveD.ToString("dd-MMM-yyyy");
                string nameChange = mailSubject.Split('_')[1];
                nameChange = nameChange.Split('(')[0];
                string mailBody = changType + ":     " + nameChange + "\r\n\r\n" + "Effective Date:     " + effectiveDate + "\r\n";
                mailBody = mailBody + inscrubed;
                mail.MailBody = mailBody;

                TaskResultList.Add(new TaskResultEntry(effectiveDate, "FM File Path", file, mail));
            }
        }
        private void upLoadNDAFile(string dir)
        {
            IEnumerable<string> fileList = GetAllTheEMAFile(dir);
            if (fileList == null)
            {
                return;
            }
            foreach (string file in fileList)
            {

                string fileName = Path.GetFileName(file);
                if (!fileName.Contains("(Revised)"))
                {

                    FileInfo fi = new FileInfo(file);
                    string date = fileName.Split('S')[1].Split('Q')[0];
                    if (fi.CreationTime.Date.Equals(DateTime.Today))
                    {
                        TaskResultList.Add(new TaskResultEntry(date, "NDA File Path", fi.FullName, FileProcessType.NDA));
                    }
                }
            }
        }
        private IEnumerable<string> GetAllTheEMAFile(string dir)
        {
            if (!Directory.Exists(dir))
            {
                Logger.Log("No email file to send today !");
                return null;
            }
            return Directory.GetFiles(dir).ToList();
        }
        private void upLoadGEDAFile(string dir)
        {
            IEnumerable<string> fileList = GetAllTheEMAFile(dir);
            if (fileList == null)
            {
                return;
            }
            foreach (string file in fileList)
            {

                string fileName = Path.GetFileName(file);
                if (!fileName.Contains("(Revised)"))
                {

                    FileInfo fi = new FileInfo(file);
                    string date = fileName.Split('.')[0].Split('_')[5];
                    if (fi.CreationTime.Date.Equals(System.DateTime.Today))
                    {
                        TaskResultList.Add(new TaskResultEntry(date, "GEDA File Path", fi.FullName, FileProcessType.GEDA_BULK_RIC_CREATION));
                    }
                }
            }
        }
       
        #endregion
        
    }
}