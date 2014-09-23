using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using Ric.Db.Info;
using Ric.Db.Manager;
//using ETI.Core;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;
using Ric.Core;
using Ric.Util;
using System.Net;

namespace Ric.Tasks.Korea
{
    public class PEOFirstPart : GeneratorBase
    {
        private KOREA_PEOGeneratorConfig configObj = null;

        private List<KoreaEquityInfo> announcementList = new List<KoreaEquityInfo>();

        private CookieContainer cookies = new CookieContainer();

        protected override void Start()
        {
            try
            {
                List<List<string>> records = GrabTodayData();
                announcementList = FilterTodayAnnounce(records);

                if (announcementList.Count == 0)
                {
                    string msg = "No Equity Pre-IPO Announcement today.";
                    Logger.Log(msg);
                    return;
                }

                UpdateAnnouncementFromISINWebpage();
                FormatAnnouncement();

                AddDataToDb();
                GenerateFM();

                GenerateGEDAFile();
                GenerateNDAFile();


            }
            catch (Exception ex)
            {
                string msg = "Error found in Start." + ex.Message;
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                AddResult("Log",Logger.FilePath,"LOG FILE");
            }

        }

        protected override void Initialize()
        {
            configObj = Config as KOREA_PEOGeneratorConfig;
        }

        #region Sourcing

        private List<List<string>> GrabTodayData()
        {
            //string postData = string.Format("method=searchPubofrProgComSub&currentPageSize=50&pageIndex=1&orderMode=1&orderStat=D&searchMode=1&searchCodeType=&isurCd=&repIsuSrtCd=&bzProcsNo=&detailMarket=&forward=pubofrprogcom_sub&marketType=&searchCorpName=&fromDate={0}&toDate={1}",
            //    DateTime.Now.AddMonths(-6).ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd"));

            ////string source = WebClientUtil.GetPageSource(null, "http://kind.krx.co.kr/listinvstg/pubofrprogcom.do", 40000, postData, Encoding.UTF8);
            //string source = WebClientUtil.GetPageSource(null, "http://kind.krx.co.kr/listinvstg/pubofrprogcom.do?method=searchPubofrProgComMain", 40000, postData, Encoding.UTF8);
            //HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
            //document.LoadHtml(source);

            string url = @"http://kind.krx.co.kr/listinvstg/pubofrprogcom.do?method=searchPubofrProgComMain";
            HtmlDocument document = new HtmlDocument();
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "GET";
            request.CookieContainer = cookies;
            request.Host = @"kind.krx.co.kr";
            request.Accept = @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
            request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
            request.Headers["Cache-Control"] = @"max-age=0";
            request.KeepAlive = true;
            request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(response.GetResponseStream());
            string st = sr.ReadToEnd();


            url = @"http://203.235.1.91/tracker.jsp?u=4&XU=&EXEN=5&dr=&XDR=&dt=%EB%8C%80%ED%95%9C%EB%AF%BC%EA%B5%AD%20%EB%8C%80%ED%91%9C%20%EA%B8%B0%EC%97%85%EA%B3%B5%EC%8B%9C%EC%B1%84%EB%84%90%20KIND&du=http%3A%2F%2Fkind.krx.co.kr%2Flistinvstg%2Fpubofrprogcom.do%3Fmethod=searchPubofrProgComMain&SX=M&js=Y&ss=1280x1024&cd=32&ce=Y&je=Y&tzo=-480&tye=2014&tmo=7&tda=8&tho=16&tmi=4&tse=34";
            request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "GET";
            request.CookieContainer = cookies;
            request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
            request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
            request.Headers["Cache-Control"] = @"max-age=0";
            request.Host = @"203.235.1.91";
            request.Referer = @"http://kind.krx.co.kr/listinvstg/pubofrprogcom.do?method=searchPubofrProgComMain";
            request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
            request.Accept = @"image/webp,*/*;q=0.8";
            response = (HttpWebResponse)request.GetResponse();
            sr = new StreamReader(response.GetResponseStream());
            st = sr.ReadToEnd();

            string postData = string.Format("method=searchPubofrProgComSub&currentPageSize=50&pageIndex=1&orderMode=1&orderStat=D&searchMode=1&searchCodeType=&isurCd=&repIsuSrtCd=&bzProcsNo=&detailMarket=&forward=pubofrprogcom_sub&marketType=&searchCorpName=&fromDate={0}&toDate={1}",
                                             DateTime.Now.AddMonths(-6).ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd"));

            string source = WebClientUtil.GetPageSource(null, "http://kind.krx.co.kr/listinvstg/pubofrprogcom.do", 40000, postData, Encoding.UTF8);
            document.LoadHtml(source);



            List<List<string>> records = GetHtmlDataNeeded(document.DocumentNode);

            return records;
        }

        private List<List<string>> GetHtmlDataNeeded(HtmlNode rootNode)
        {

            List<List<string>> res = new List<List<string>>();

            HtmlNodeCollection nodeList = rootNode.SelectNodes("//table");
           // HtmlNodeCollection nodeList = rootNode.SelectNodes("//section[@id='section-talbe']");
            if (nodeList.Count < 1)
            {
                Logger.Log("No table HTML element found!", Logger.LogType.Error);
                return res;
            }

            HtmlNodeCollection rows = nodeList[0].SelectNodes("./tbody/tr");

            if (rows.Count < 1)
            {
                Logger.Log("No tr HTML element found!", Logger.LogType.Error);
                return res;
            }
            else if (rows.Count == 1)
            {
                if (rows[0].InnerText.Contains("조회된 결과값이 없습니다"))
                {
                    Logger.Log("No record found.", Logger.LogType.Info);
                    return res;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                try
                {

                    string effectiveDate = rows[i].SelectNodes("./td")[7].InnerText.Trim(); //Effective Date

                    DateTime dt;
                    if (!(DateTime.TryParse(effectiveDate, out dt)))
                    {
                        continue;
                    }

                    if (DateTime.Compare(dt, DateTime.Today) <= 0)
                    {
                        return res;
                    }

                    string market = rows[i].SelectNodes("./td")[0].SelectSingleNode("./img").Attributes["alt"].Value.Trim();

                    string onclick = rows[i].Attributes["onclick"].Value;

                    Regex reg = new Regex("(?<code>\\d{1,})");
                    Match codeMatch = reg.Match(onclick);
                    string code = codeMatch.Groups["code"].Value; //link paramater bzProcsNo . unique.

                    List<string> temp = new List<string>();
                    temp.Add(code);
                    temp.Add(rows[i].SelectNodes("./td")[0].InnerText.Trim());  //Name
                    //temp.Add(rows[i].SelectNodes("./td")[4].InnerText.Trim());  //Delivery date
                    temp.Add(effectiveDate);
                    temp.Add(market);

                    res.Add(temp);


                }
                catch (Exception ex)
                {
                    Logger.Log("Error: " + ex.Message, Logger.LogType.Error);
                }
            }

            return res;
        }

        private string GetCompanyName(string code)
        {
            string url = string.Format("http://kind.krx.co.kr/listinvstg/pubofrprogcomdetail.do?method=searchProgComDetailMain&bzProcsNo={0}", code);
            HtmlDocument document = new HtmlDocument();
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Timeout = 300000;
            request.Method = "GET";
            request.CookieContainer = cookies;
            request.Host = @"kind.krx.co.kr";
            request.Accept = @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.Headers["Accept-Encoding"] = @"gzip,deflate,sdch";
            request.Headers["Accept-Language"] = @"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4";
            
            request.KeepAlive = true;
            request.UserAgent = @"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36";
            request.Referer = @"http://kind.krx.co.kr/listinvstg/pubofrprogcom.do?method=searchPubofrProgComMain";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(response.GetResponseStream());
            string st = sr.ReadToEnd();
            document.LoadHtml(st);
            HtmlNodeCollection nodes = document.DocumentNode.SelectNodes("//tr[@class='first']");
            st = nodes[0].SelectNodes("./td")[0].InnerText.Trim();

            return st;
        }

        private List<KoreaEquityInfo> FilterTodayAnnounce(List<List<string>> records)
        {
            List<KoreaEquityInfo> anns = new List<KoreaEquityInfo>();

            //If code existed in DB. Ignore this announcement.

            foreach (List<string> record in records)
            {
                string code = record[0];
                if (KoreaEquityManager.ExistsFmOneCode(code))
                {
                    Logger.Log(string.Format("Announcement of {0} is already existed in database.", record[1]));
                    continue;
                }
                KoreaEquityInfo ann = new KoreaEquityInfo();
                ann.KoreaName = GetCompanyName(code);
                //ann.KoreaName = record[1];
                ann.EffectiveDate = record[2];
                ann.Market = GetDutyCode(record[3]);
                anns.Add(ann);
            }

            return anns;
        }

        private void UpdateAnnouncementFromISINWebpage()
        {
            if (announcementList.Count <= 0)
            {
                return;
            }

            foreach (KoreaEquityInfo item in announcementList)
            {
                try
                {
                    if (!KoreaISINUtil.UpdateEquityISINReportNew(item))
                    {
                        Logger.Log("Can't find equity item from ISNI web site", Logger.LogType.Info);

                        if (!KoreaISINUtil.UpdateKDRISINReportNew(item))
                        {
                            Logger.Log("Can't find KDR item from ISNI web site", Logger.LogType.Error);

                            // Update ISIN and Type manually
                            string input = null;

                            while (string.IsNullOrEmpty(input))
                            {
                                input = InputISINType.Prompt(item.KoreaName);
                                if (input == null)
                                {
                                    continue; // User cancelled input
                                }

                                if (input.IndexOf(",") == -1)
                                {
                                    System.Windows.Forms.MessageBox.Show("Bad format");
                                    input = null;
                                    continue;
                                }

                                int idx = input.IndexOf(",");
                                item.ISIN = input.Substring(0, idx).Trim();
                                item.Type = input.Substring(idx + 1).Trim();
                            }
                        }
                    }

                    //Get ticker and legal name
                    KoreaISINUtil.GetTickerAndLegalNameByISIN(item);

                }
                catch (Exception ex)
                {
                    string msg = string.Format("At UpdateAnnouncementFromISINWebpage(). Error message:{0}", ex.Message);
                    Logger.Log(msg, Logger.LogType.Error);
                }
            }
        }

        /// <summary>
        /// Get exchange board code from announcement.
        /// </summary>
        /// <param name="ddNode">announcement node</param>
        /// <returns>exchange board code</returns>
        private string GetDutyCode(string dutyName)
        {
            if (string.IsNullOrEmpty(dutyName))
            {
                return "";
            }
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

        #endregion

        #region Formatting

        private void FormatAnnouncement()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            for (var i = 0; i < announcementList.Count; i++)
            {
                KoreaEquityInfo peoTemp = announcementList[i];

                if (string.IsNullOrEmpty(peoTemp.Ticker))
                {
                    continue;
                }

                peoTemp.UpdateDate = DateTime.Today.ToString("dd-MMM-yy");
                peoTemp.EffectiveDate = DateTime.Today.Year.ToString();

                peoTemp.RecordType = "113";
                if (peoTemp.Type.Equals("PRF"))
                {
                    peoTemp.RecordType = "97";

                    if (peoTemp.KoreaName.Contains("우선주") && peoTemp.KoreaName.Length > 14)
                    {
                        string tempName = peoTemp.KoreaName.Replace("우선주", "");
                        if (tempName.Length > 14)
                        {
                            peoTemp.KoreaName = tempName.Substring(0, 13) + "우";
                        }
                        else
                        {
                            peoTemp.KoreaName = tempName + "우";
                        }
                    }
                }

                if (peoTemp.KoreaName.Length > 14)
                {
                    peoTemp.KoreaName = peoTemp.KoreaName.Substring(0, 14);
                }

                peoTemp.FM = "1";
                peoTemp.RIC = peoTemp.Ticker + "." + peoTemp.Market;
                peoTemp.BcastRef = peoTemp.RIC;

                KoreaEquityCommon.FormatEQIdnDisplayName(peoTemp);
            }
        }

        #endregion

        #region Distributing

        #region Update data to database

        private void AddDataToDb()
        {
            if (announcementList.Count > 0)
            {
                Logger.Log("Adding data to database.");
                int rows = KoreaEquityManager.UpdateEquity(announcementList);
                Logger.Log(string.Format("{0} records updated.", rows.ToString()));
            }
        }

        #endregion

        #region Generate FM File

        private void GenerateFM()
        {
            if (announcementList.Count < 1)
            {
                Logger.Log("No new added equity record", Logger.LogType.Info);
                return;
            }

            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference correct!!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            string rics = CombineAllRics(announcementList);
            string filename = string.Format("KR FM (PEO)_{0}.xls", rics);
            string fmFile = Path.Combine(configObj.FM, filename);

            try
            {
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, fmFile);
                Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet1", wBook);
                if (wSheet == null)
                {
                    string msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                int startLine = WriteFMHeader(wSheet, 3, true);

                foreach (var item in announcementList)
                {
                    WriteFMEquityItem(wSheet, startLine, item);

                    startLine++;
                }

                WriteFMFooter(wSheet, startLine + 1);

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();

                MailToSend mail = new MailToSend();
                mail.ToReceiverList.AddRange(configObj.MailTo);
                mail.MailSubject = Path.GetFileNameWithoutExtension(filename);
                mail.CCReceiverList.AddRange(configObj.MailCC);
                mail.AttachFileList.Add(fmFile);
                mail.MailBody = "PEO:\t" + rics + "\t\r\n\r\n\r\n";

                string signature = string.Join("\r\n", configObj.MailSignature.ToArray());

                mail.MailBody += signature;

                AddResult(filename,fmFile,"FM File");
            }
            catch (Exception ex)
            {
                Logger.Log("Error in GenerateFM(): " + ex.ToString(), Logger.LogType.Error);
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

        private int WriteFMHeader(Worksheet wSheet, int rowNum, bool writeAddition)
        {
            ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 12;
            ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 5;
            ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 12;
            ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 3;
            ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 25;
            ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 14;
            ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 8;
            ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 12;
            ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 30;
            ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 8;
            ((Range)wSheet.Columns["A:M", System.Type.Missing]).Font.Name = "Arial";

            ((Range)wSheet.Rows[rowNum, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            ((Range)wSheet.Rows[rowNum, Type.Missing]).Font.Underline = System.Drawing.FontStyle.Underline;
            ((Range)wSheet.Rows[rowNum + 1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            ((Range)wSheet.Rows[rowNum + 1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

            if (writeAddition)
            {
                wSheet.Cells[rowNum++, 1] = "EQUITY ADD";
            }

            wSheet.Cells[rowNum, 1] = "Updated Date";
            wSheet.Cells[rowNum, 2] = "Effective Date";
            wSheet.Cells[rowNum, 3] = "RIC";
            wSheet.Cells[rowNum, 4] = "Type";
            wSheet.Cells[rowNum, 5] = "Record Type";
            wSheet.Cells[rowNum, 6] = "FM";
            wSheet.Cells[rowNum, 7] = "IDN Display Name";
            wSheet.Cells[rowNum, 8] = "ISIN";
            wSheet.Cells[rowNum, 9] = "Ticker";
            wSheet.Cells[rowNum, 10] = "BCAST_REF";
            wSheet.Cells[rowNum, 11] = "Legal Name";
            wSheet.Cells[rowNum, 12] = "Korea Name";
            wSheet.Cells[rowNum, 13] = "Lotsize";

            return rowNum + 1;
        }

        private int WriteFMFooter(Worksheet wSheet, int lineNum)
        {
            ((Range)wSheet.Rows[lineNum, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            wSheet.Cells[lineNum, 1] = "- End -";

            return lineNum + 1;
        }

        private void WriteFMEquityItem(Worksheet wSheet, int startLine, KoreaEquityInfo item)
        {
            KoreaEquityInfo peoTemp = item;

            ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
            wSheet.Cells[startLine, 1] = peoTemp.UpdateDate;
            ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
            wSheet.Cells[startLine, 2] = peoTemp.EffectiveDate;
            if (peoTemp.RIC.Length < 9)
                ((Range)wSheet.Cells[startLine, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
            wSheet.Cells[startLine, 3] = peoTemp.RIC;
            wSheet.Cells[startLine, 4] = peoTemp.Type;
            wSheet.Cells[startLine, 5] = peoTemp.RecordType;
            wSheet.Cells[startLine, 6] = peoTemp.FM;
            wSheet.Cells[startLine, 7] = peoTemp.IDNDisplayName;
            if (peoTemp.IDNDisplayName.Length > 16)
            {
                ((Range)wSheet.Cells[startLine, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                Logger.Log("IDN Display Name is highlight because of the length is over 16 characters!");
            }
            wSheet.Cells[startLine, 8] = peoTemp.ISIN;
            ((Range)wSheet.Cells[startLine, 9]).NumberFormat = "@";
            if (String.IsNullOrEmpty(peoTemp.Ticker))
                ((Range)wSheet.Cells[startLine, 9]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
            wSheet.Cells[startLine, 9] = peoTemp.Ticker;
            wSheet.Cells[startLine, 10] = peoTemp.BcastRef;
            wSheet.Cells[startLine, 11] = peoTemp.LegalName;
            wSheet.Cells[startLine, 12] = peoTemp.KoreaName;
            if (peoTemp.RIC.EndsWith("KS"))
            {
                peoTemp.Lotsize = "1";
            }
            else if (peoTemp.RIC.EndsWith("KQ"))
            {
                peoTemp.Lotsize = "1";
            }
            else if (peoTemp.RIC.EndsWith("KN"))
            {
                peoTemp.Lotsize = "100";
            }
            wSheet.Cells[startLine, 13] = peoTemp.Lotsize;
        }

        #endregion

        #region Generate GEDA File

        private void GenerateGEDAFile()
        {
            if (announcementList.Count > 0)
            {
                if (!Directory.Exists(configObj.GEDA))
                {
                    Directory.CreateDirectory(configObj.GEDA);
                }

                string fileName = string.Format("KR_PEO_Bulk_Creation_KOSDAQ_{0}.txt", DateTime.Now.ToString("yyyyMMdd"));
                //string fileName = "KR_PEO_Bulk_Creation_KOSDAQ.txt";
                string gedaFile = Path.Combine(configObj.GEDA, fileName);

                StringBuilder kqBuilder = new StringBuilder();
                GEDAFileTemplate KOSDAQ = new GEDAFileTemplate();

                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Symbol);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Display_Name);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Ric);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Official_Code);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Ex_Symbol);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Background_Page);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Display_Nmll);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Bcast_Ref);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Instmod_ISIN);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Instmod_Mnemonic);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Instmod_Tdn_Issuer_Name);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Instmod_Tdn_Symbol);
                kqBuilder.AppendFormat("{0}\t", KOSDAQ.Geda_Exl_Name);
                kqBuilder.Append(KOSDAQ.Geda_BCU);
                kqBuilder.AppendLine();

                foreach (var item in announcementList)
                {
                    kqBuilder.AppendFormat("{0}\t", item.RIC);
                    kqBuilder.AppendFormat("{0}\t", item.IDNDisplayName);
                    kqBuilder.AppendFormat("{0}\t", item.RIC);
                    kqBuilder.AppendFormat("{0}\t", item.Ticker);

                    if (string.IsNullOrEmpty(item.ISIN))
                    {
                        kqBuilder.AppendFormat("==NOISIN==\t");
                    }
                    else
                    {
                        kqBuilder.AppendFormat("{0}\t", item.ISIN);
                    }

                    kqBuilder.AppendFormat("{0}\t", "****");
                    kqBuilder.AppendFormat("{0}\t", item.KoreaName);
                    kqBuilder.AppendFormat("{0}\t", item.RIC);
                    kqBuilder.AppendFormat("{0}\t", item.ISIN);
                    kqBuilder.AppendFormat("{0}\t", "A" + item.Ticker);
                    kqBuilder.AppendFormat("{0}\t", item.LegalName);
                    kqBuilder.AppendFormat("{0}\t", item.Ticker);

                    if (item.Type == null)
                    {
                        kqBuilder.Append("==NOTYPE==");
                    }
                    else if (item.Type.Equals("KDR"))
                    {
                        if (item.RIC.Contains(".KQ"))
                        {
                            kqBuilder.Append("KOSDAQ_EQB_KDR");
                        }
                        else if (item.RIC.Contains(".KS"))
                        {
                            kqBuilder.Append("KSE_EQB_KDR");
                        }
                    }
                    else
                    {
                        if (item.RIC.Contains(".KQ"))
                        {
                            kqBuilder.Append("KOSDAQ_EQB_2");
                        }
                        else if (item.RIC.Contains(".KS"))
                        {
                            kqBuilder.Append("KSE_EQB_3");
                        }

                        // BCU
                        if (item.Type.Equals("ORD"))
                        {
                            if (item.RIC.Contains(".KQ"))
                            {
                                kqBuilder.Append("\tKOSDAQ_EQ_IPO");
                            }
                            else if (item.RIC.Contains(".KS"))
                            {
                                kqBuilder.Append("\tKSE_EQ_IPO");
                            }
                        }
                    }

                    kqBuilder.AppendLine();
                }

                File.WriteAllText(gedaFile, kqBuilder.ToString(), Encoding.UTF8);
                AddResult(fileName,gedaFile,"GEDA File");
            }
        }

        #endregion

        #region Generate NDA File

        private void GenerateNDAFile()
        {
            if (announcementList.Count > 0)
            {
                if (!Directory.Exists(configObj.NDA))
                {
                    Directory.CreateDirectory(configObj.NDA);
                }
                string fileName = "KR" + DateTime.Now.ToString("yyyyMMdd") + "QAADD.csv";
                string ndaFile = Path.Combine(configObj.NDA, fileName);

                StringBuilder kqBuilder = new StringBuilder();
                NDAFileTemplate KOSDAQ = new NDAFileTemplate();

                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Ric);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Tag);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Base_Asset);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Ticker_Symbol);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Asset_Short_Name);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Asset_Common_Name);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Type);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Category);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Currency);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Exchange);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Equity_First_Trading_Day);
                kqBuilder.AppendFormat(",{0}", KOSDAQ.Nda_Round_Lot_Size);
                kqBuilder.Remove(0, 1);
                kqBuilder.AppendLine();

                foreach (var item in announcementList)
                {
                    //if (item.RIC.EndsWith("KS"))
                    //{
                    //    item.Lotsize = "1";
                    //}
                    //else if (item.RIC.EndsWith("KQ"))
                    //{
                    //    item.Lotsize = "1";
                    //}
                    //else if (item.RIC.EndsWith("KN"))
                    //{
                    //    item.Lotsize = "100";
                    //}
                    if (item.RIC.Contains("KQ"))
                    {
                        for (int i = 0; i < 6; i++)
                        {
                            switch (i)
                            {
                                case 0:
                                    kqBuilder.AppendFormat("{0}", item.RIC);
                                    item.Lotsize = "1";
                                    kqBuilder.AppendFormat(",{0}", "673");
                                    break;
                                case 1:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "F.KQ");
                                    kqBuilder.AppendFormat(",{0}", "64399");
                                    item.Lotsize = string.Empty;
                                    break;
                                case 2:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "S.KQ");
                                    item.Lotsize = "1";
                                    kqBuilder.AppendFormat(",{0}", "60673");
                                    break;
                                case 3:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "stat.KQ");
                                    kqBuilder.AppendFormat(",{0}", "61287");
                                    item.Lotsize = string.Empty;
                                    break;
                                case 4:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "ta.KQ");
                                    kqBuilder.AppendFormat(",{0}", "64380");
                                    item.Lotsize = string.Empty;
                                    break;
                                case 5:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "bl.KQ");
                                    kqBuilder.AppendFormat(",{0}", "67094");
                                    item.Lotsize = string.Empty;
                                    break;
                            }

                            AppendOtherColumnToNDAFile(kqBuilder, item, "KOE");
                        }
                    }

                    if (item.RIC.Contains("KS"))
                    {
                        for (int i = 0; i < 6; i++)
                        {
                            switch (i)
                            {
                                case 0:
                                    kqBuilder.AppendFormat("{0}", item.RIC);
                                    item.Lotsize = "1";
                                    kqBuilder.AppendFormat(",{0}", "184");
                                    break;
                                case 1:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "F.KS");
                                    kqBuilder.AppendFormat(",{0}", "64398");
                                    item.Lotsize = string.Empty;
                                    break;
                                case 2:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "S.KS");
                                    item.Lotsize = "1";
                                    kqBuilder.AppendFormat(",{0}", "60184");
                                    break;
                                case 3:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "stat.KS");
                                    kqBuilder.AppendFormat(",{0}", "61286");
                                    item.Lotsize = string.Empty;
                                    break;
                                case 4:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "ta.KS");
                                    kqBuilder.AppendFormat(",{0}", "64379");
                                    item.Lotsize = string.Empty;
                                    break;
                                case 5:
                                    kqBuilder.AppendFormat("{0}", item.RIC.Split('.')[0] + "bl.KS");
                                    kqBuilder.AppendFormat(",{0}", "67093");
                                    item.Lotsize = string.Empty;
                                    break;
                            }

                            AppendOtherColumnToNDAFile(kqBuilder, item, "KSC");
                        }
                    }
                }

                File.WriteAllText(ndaFile, kqBuilder.ToString(), Encoding.UTF8);
                AddResult(fileName,ndaFile,"NDA File");
            }
        }

        private void AppendOtherColumnToNDAFile(StringBuilder kqBuilder, KoreaEquityInfo item, string exchange)
        {
            if (string.IsNullOrEmpty(item.ISIN))
            {
                kqBuilder.AppendFormat(",{0}", "==NOISIN==");
            }
            else
            {
                kqBuilder.AppendFormat(",{0}", "ISIN:" + item.ISIN);
            }

            kqBuilder.AppendFormat(",{0}", item.Ticker);
            kqBuilder.AppendFormat(",{0}", item.IDNDisplayName);
            kqBuilder.AppendFormat(",{0}", item.IDNDisplayName + " " + item.Type);
            kqBuilder.AppendFormat(",{0}", "EQUITY");

            if (string.IsNullOrEmpty(item.Type))
            {
                kqBuilder.AppendFormat(",{0}", "==NOTYPE==");
            }
            else
            {
                if (item.Type.Equals("KDR"))
                {
                    kqBuilder.AppendFormat(",{0}", "DRC");
                }
                else
                {
                    kqBuilder.AppendFormat(",{0}", item.Type);
                }
            }

            kqBuilder.AppendFormat(",{0}", "KRW");
            kqBuilder.AppendFormat(",{0}", exchange);
            kqBuilder.AppendFormat(",{0}", string.Empty);
            kqBuilder.AppendFormat(",{0}", item.Lotsize);
            kqBuilder.AppendLine();
        }

        #endregion

        #endregion

    }
}
