using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using Selenium;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Xml;
using System.Globalization;
using System.IO;
//using Ric.Generator.Lib.Korea;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using HtmlAgilityPack;
using System.Reflection;
using Ric.Db.Manager;
using System.Text.RegularExpressions;
using System.Net;
//using ETI.Core;
using System.Xml.Linq;
using Ric.Db.Info;
using System.Threading;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
{
    public class ELWFMFirstPart : GeneratorBase
    {
        private List<WarrantTemplate> koreaList = new List<WarrantTemplate>();
        private List<KoreaUnderlyingInfo> newUnderLying = new List<KoreaUnderlyingInfo>();
        private KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig configObj = null;

        private void Initialize(KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig obj)
        {
            configObj = obj;
        }
        public ELWFMFirstPart(List<TaskResultEntry> resultList, Logger logger)
        {
            this.TaskResultList = resultList;
            this.Logger = logger;
        }
        public int StartELWFM1AndBulkFileJob(KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig obj)
        {

            Initialize(obj);
            int count = StartELWFMFirstPartJob();
            StartBulkFileGenerateJob();
            return count;
        }

        private int StartELWFMFirstPartJob()
        {
            GrabDataFromWebpage();
            if (newUnderLying.Count > 0)
            {
                GenerateNewUnderlyingFiles();
            }
            if (koreaList.Count > 0)
            {
                GenerateELWFM1File_xls();
                CopyELWFM1FileToEmaFile();
                UpdateELWTable();
            }
            int count = koreaList.Count;
            return count;
        }

        private void UpdateELWTable()
        {
            int row = KoreaELWManager.InsertELW(koreaList);
            string msge = string.Format("Updated {0} ELW FM1 record(s) in database.", row);
            Logger.Log(msge);
        }

        private void StartBulkFileGenerateJob()
        {
            if (koreaList.Count > 0)
            {
                string foldername = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
                string ipath = CreateFolder(foldername);

                GenerateNDATCFile_xls(ipath);
                GenerateNDAIAFileXls(ipath);
                GenerateNDAQAFileXls(ipath);
                GenerateKSEELWFileTxt(ipath);
            }
        }

        private void GrabDataFromWebpage()
        {
            try
            {
                int pageCount = 0;
                string startDate = configObj.StartDate.Replace("-", "");
                string endDate = configObj.EndDate.Replace("-", "");
                string postData = string.Empty;
                string pageSource = string.Empty;
                string uri = "http://isin.krx.co.kr/srch/srch.do?method=srchList";
                //pageCount = GetTotalPageCountNew(uri, startDate, endDate);

                for (int times = 0; times < 5; times++)
                {
                    try
                    {
                        pageCount = GetTotalPageCountNew(uri, startDate, endDate);

                        if (pageCount >= 0)
                            break;
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(2000);

                        if (times == 4)
                        {
                            string msg = string.Format("get page index error,retry 4 times.\r\nmsg:{0}", ex.Message);
                            Logger.Log(msg, Logger.LogType.Error);
                        }
                    }
                }

                for (var i = 1; i <= pageCount; i++)
                {
                    postData = File.ReadAllText(@"Config\Korea\ELWFMFirstPartPostData.txt", Encoding.UTF8);
                    postData = string.Format(postData, startDate, endDate, i.ToString());
                    Thread.Sleep(2000);
                    HtmlDocument htc = GetDocFromELW(uri, postData);

                    if (htc != null)
                    {
                        var nodeList = htc.DocumentNode.SelectNodes(".//table")[1].SelectNodes(".//tr");
                        if (nodeList != null)
                        {
                            for (int x = 1; x < nodeList.Count; x++)
                            {
                                WarrantTemplate warrantTemp = new WarrantTemplate();
                                string urlISIN = @"http://isin.krx.co.kr/srch/srch.do?method=srchPopup11";
                                string std_cd = nodeList[x].SelectNodes(".//td")[1].InnerText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();
                                string postDataISIN = string.Format("stdcd_type=11&std_cd={0}", std_cd);
                                //HtmlDocument doc = GetDocByISIN(urlISIN, postDataISIN);
                                HtmlDocument doc = null;

                                for (int times = 0; times < 5; times++)
                                {
                                    try
                                    {
                                        doc = GetDocByISIN(urlISIN, postDataISIN);

                                        if (doc != null)
                                            break;
                                    }
                                    catch (Exception ex)
                                    {
                                        Thread.Sleep(2000);

                                        if (times == 4)
                                        {
                                            string msg = string.Format("get table of :{0} error,retry 4 times.\r\nmsg:{0}", std_cd, ex.Message);
                                            Logger.Log(msg, Logger.LogType.Error);
                                        }
                                    }
                                }

                                if (doc != null)
                                {
                                    //var listNodes = doc.DocumentNode.SelectNodes(".//table")[1];
                                    var rootTable = doc.DocumentNode.SelectNodes(".//table");
                                    if (rootTable.Count < 2)
                                    {
                                        string msg = string.Format("extract isin:{0} failed.", std_cd);
                                        Logger.Log(msg, Logger.LogType.Warning);
                                        continue;
                                    }
                                    var listNodes = rootTable[1];
                                    //string tr5_td2 = table.SelectSingleNode(".//tr[5]/td[2]").InnerText.Trim().ToString();
                                    string tr5_td2 = listNodes.SelectNodes(".//tr")[10].SelectNodes(".//td")[0].InnerText.Replace(" ", "");
                                    if (tr5_td2.Equals("사모") || tr5_td2.Equals("기타"))
                                    {
                                        string msg = string.Format("ISIN:{0} equals 사모||기타. Ignore it.", std_cd);
                                        Logger.Log(msg);
                                        continue;
                                    }
                                    warrantTemp.ISIN = std_cd;
                                    Thread.Sleep(2000);
                                    GrabDataAndWarrantTemplateFormat(warrantTemp, doc);
                                    if (warrantTemp.IsKOBA)
                                        continue;
                                    else
                                        koreaList.Add(warrantTemp);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabDataFromWebpage()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private int GetTotalPageCountNew(string url, string startDate, string endDate)
        {
            int pageCount = 0;
            try
            {
                string page = string.Empty;
                string post = File.ReadAllText(@"Config\Korea\ELWFMFirstPartPostData.txt", Encoding.UTF8);
                post = string.Format(post, startDate, endDate, "1");
                HtmlDocument doc = GetDocFromELW(url, post);
                HtmlNodeCollection ListNode = doc.DocumentNode.SelectNodes(".//div")[21].SelectNodes(".//ul")[1].SelectNodes(".//span");
                page = ListNode[0].InnerText.Replace(" ", "").Replace(")", "");
                page = page.Substring(page.IndexOf('/') + 1, page.Length - page.IndexOf('/') - 1);

                if (int.TryParse(page, out pageCount))
                    return pageCount;

                return pageCount;
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetTotalPageCountNew() failed.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return -1;
            }
        }

        private HtmlDocument GetDocFromELW(string url, string post)
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
                request.ContentType = "multipart/form-data; boundary=----WebKitFormBoundarywi94Z59UxjiOxuTN";
                byte[] buf = Encoding.UTF8.GetBytes(post);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);
                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();
                doc.Load(httpResponse.GetResponseStream(), Encoding.UTF8);
                return doc;
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetDocFromELW() failed.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }
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
                string msg = string.Format("execute GetDocByISIN() failed.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }
        }

        private void GrabDataAndWarrantTemplateFormat(WarrantTemplate warrantTemp, HtmlDocument doc)
        {
            try
            {
                //string tr3_td2 = table.SelectSingleNode(".//tr[3]/td[1]").InnerText.Trim().ToString();
                //if (tr3_td2.Split(' ')[1].Equals("조기종료"))
                //{
                //    warrantTemp.IsKOBA = true;
                //    return;
                //}

                var table = doc.DocumentNode.SelectNodes(".//table")[0];
                var table0 = table.SelectNodes(".//table")[0];
                var table1 = table.SelectNodes(".//table")[1];
                //string tr5_td2 = table0.SelectNodes(".//tr")[10].SelectNodes(".//td")[0].InnerText.Replace(" ", "");

                string tr2_td4 = table0.SelectNodes(".//tr")[1].SelectNodes(".//td")[1].InnerText.ToString();//tr1td1   J594220
                string tr3_td4 = table0.SelectNodes(".//tr")[3].SelectNodes(".//td")[0].InnerText.ToString();//tr3td0   미래4220대림산업콜
                string tr4_td2 = table0.SelectNodes(".//tr")[2].SelectNodes(".//td")[1].InnerText.ToString();//tr2td1   MIRAE ASSET SECURITIES ELW 4220
                string tr4_td4 = table0.SelectNodes(".//tr")[3].SelectNodes(".//td")[1].InnerText.ToString().Replace(" ", "");//tr3td1   MIRAE4220DAELIM- C
                string tr6_td2 = table0.SelectNodes(".//tr")[7].SelectNodes(".//td")[1].InnerText.ToString();//tr7td1   2014-05-27
                string tr6_td4 = table0.SelectNodes(".//tr")[8].SelectNodes(".//td")[1].InnerText.ToString();//tr8td1   2015-01-20
                string tr7_td2 = table0.SelectNodes(".//tr")[8].SelectNodes(".//td")[0].InnerText.ToString();//tr8td0   200
                string tr7_td4 = table0.SelectNodes(".//tr")[7].SelectNodes(".//td")[0].InnerText.ToString();//tr7td0   5,000,000
                string tr8_td2 = table0.SelectNodes(".//tr")[10].SelectNodes(".//td")[1].InnerText.ToString();//tr10td1  0.0115
                string tr11_td4 = table1.SelectNodes(".//tr")[1].SelectNodes(".//td")[0].InnerText.ToString();//table1 tr1td0        KOSPI200  
                string tr13_td2 = table1.SelectNodes(".//tr")[0].SelectNodes(".//td")[1].InnerText.ToString();//table1 tr1td1        대림산업(주)  구성주식비율: 1
                //string tr20_td4 = table.SelectSingleNode(".//tr[20]/td[4]").InnerText.Trim().ToString();

                if (tr13_td2 == string.Empty && tr11_td4 != string.Empty)
                    tr13_td2 = tr11_td4;
                string koreaNameForNewUnderlying = tr13_td2;
                //tr13_td2 = tr13_td2.Contains("(주)") ? tr13_td2.Replace("(주)", "") : tr13_td2;
                //.Replace("&nbsp;", "")
                //(주)      구성주식비율:1
                tr13_td2 = tr13_td2.Replace("&nbsp;", "").Replace("(주)", "").Replace("구성주식비율: 1", "");//Replace("구성주식비율:1", "");

                warrantTemp.UnderlyingKoreaName = tr13_td2;

                string issuer_koreaname = tr3_td4;
                char[] str_issuer = issuer_koreaname.ToCharArray();
                string ikoreaname = string.Empty;
                foreach (var item in str_issuer)
                {
                    if (item > 47 && item < 58) break;
                    ikoreaname += item.ToString();
                }
                warrantTemp.IssuerKoreaName = ikoreaname;

                warrantTemp.UpdateDate = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                warrantTemp.EffectiveDate = DateTime.Today.Year.ToString();
                warrantTemp.FM = "1";
                warrantTemp.RIC = tr2_td4.Substring(1) + ".KS";
                warrantTemp.Ticker = tr2_td4.Substring(1);

                warrantTemp.OrgMatDate = tr6_td4.Trim().ToString();
                warrantTemp.MatDate = Convert.ToDateTime(tr6_td4).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                warrantTemp.StrikePrice = string.Empty;//tr20_td4.Contains(",") ? tr20_td4.Replace(",", "") : tr20_td4;
                warrantTemp.QuanityofWarrants = tr7_td4.Contains(",") ? tr7_td4.Replace(",", "") : tr7_td4;
                string issuer_price = tr7_td2.Contains(",") ? tr7_td2.Replace(",", "") : tr7_td2;
                warrantTemp.IssuePrice = issuer_price;
                warrantTemp.OrgIssueDate = tr6_td2;
                warrantTemp.IssueDate = Convert.ToDateTime(tr6_td2).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                warrantTemp.ConversionRatio = tr8_td2;
                warrantTemp.Issuer = tr4_td2.ToUpper();
                warrantTemp.KoreaWarrantName = tr3_td4;

                //combine data
                int accout = tr4_td4.Split('-').Length;
                string PorC = tr4_td4.Split('-')[(accout - 1)];
                warrantTemp.CallOrPut = PorC.Equals("C") ? "CALL" : "PUT";
                string underlying_ric = string.Empty;
                //IDN Display Name

                KoreaIssuerInfo issuer = KoreaIssuerManager.SelectIssuer(warrantTemp.IssuerKoreaName);

                if (issuer == null)
                {
                    Logger.Log("Can not find issuer info in database. Please check the table of issuer info", Logger.LogType.Warning);
                    return;
                }
                string sname = "***";
                string cname = "***";
                sname = issuer.IssuerCode4;

                string NumLen = tr2_td4.Substring(3, 4);
                string underly = string.Empty;
                if (!string.IsNullOrEmpty(tr13_td2))
                {
                    KoreaUnderlyingInfo underlying = KoreaUnderlyingManager.SelectUnderlying(tr13_td2);
                    if (underlying != null)
                    {
                        underly = underlying.IDNDisplayNamePart;
                        underlying_ric = underlying.UnderlyingRIC;
                        cname = underlying.QACommonNamePart;
                    }
                    else
                    {
                        string issuerName = issuer_koreaname.Replace(ikoreaname, "");
                        issuerName = issuerName.Substring(4, issuerName.Length - 5);
                        KoreaUnderlyingInfo temp = NewUnderlying.GrabNewUnderlyingInfo(tr13_td2);
                        if (temp != null)
                        {
                            temp.UnderlyingName = issuerName;
                            try
                            {
                                KoreaUnderlyingManager.UpdateUnderlying(temp);
                            }
                            catch (Exception ex)
                            {
                                Logger.Log("Error found in insert new underlying into Database.\r\n" + ex.ToString(), Logger.LogType.Error);
                            }
                            //temp.KoreaName = koreaNameForNewUnderlying;
                            newUnderLying.Add(temp);
                            underly = temp.IDNDisplayNamePart;
                            cname = temp.QACommonNamePart;
                            underlying_ric = temp.UnderlyingRIC;
                            Logger.Log("Insert new underlying to database successfully. RIC:" + temp.UnderlyingRIC);
                        }
                        else
                        {
                            Logger.Log("Can not get the new underlying info for " + tr13_td2, Logger.LogType.Error);
                            underly = "***";
                        }
                    }
                }
                warrantTemp.IDNDisplayName = sname + NumLen + underly + PorC;
                warrantTemp.BCASTREF = underlying_ric.ToUpper();

                string mtime = Convert.ToDateTime(warrantTemp.OrgMatDate).ToString("MMM-yy", new CultureInfo("en-US")).Replace("-", "").ToUpper();
                //string price = warrantTemp.StrikePrice.Contains(".") ? warrantTemp.StrikePrice.Split('.')[0] : warrantTemp.StrikePrice;
                //price = price.Length >= 4 ? price.Substring(0, 4) : price;
                string last = tr11_td4.Equals("KOSPI200") ? "IW" : "WNT";
                //string qacommonname = cname + " " + sname + " " + mtime + " " + price + " " + PorC + last;
                warrantTemp.QACommonName = "   "; //qacommonname.ToUpper();
                //Chain
                string chain = tr11_td4.Equals("KOSPI200") ? ("0#WARRANTS.KS, 0#ELW.KS, 0#.KS200W.KS") : "0#WARRANTS.KS, 0#ELW.KS, 0#CELW.KS, 0#" + warrantTemp.BCASTREF.Split('.')[0] + "W." + warrantTemp.BCASTREF.Split('.')[1];
                warrantTemp.Chain = chain.ToUpper();
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabDataAndWarrantTemplateFormat()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }
        }


        private void GenerateELWFM1File_xls()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                string filename = "Korea FM for " + DateTime.Today.ToString("dd-MMM-yyyy", new CultureInfo("en-US")).Replace("-", " ") + " (Morning).xls";
                ELWDropELWFM1ELWFileBulkGenerate.filename = filename;
                string ipath = Path.Combine(configObj.FM, filename);
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                GenerateExcelFileTitle(wSheet);
                koreaList.Sort(new Common().CompareTicker);

                int startLine = 5;
                AppendDataToFile(wSheet, startLine, "common");

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                string msg = "Error found in GenerateELWFM1File_xls()    : " + ex.StackTrace + " \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private string CreateFM1FileEmaDir()
        {
            string dir = Path.Combine(ConfigureOperator.GetEmaFileSaveDir(), DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US")));
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            return dir;
        }
        private void CopyELWFM1FileToEmaFile()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {

                if (string.IsNullOrEmpty(ELWDropELWFM1ELWFileBulkGenerate.filename))
                {
                    ELWDropELWFM1ELWFileBulkGenerate.filename = "Korea FM for " + DateTime.Today.ToString("dd-MMM-yyyy", new CultureInfo("en-US")).Replace("-", " ") + " (Morning).xls";
                }

                string emaPath = CreateFM1FileEmaDir() + "\\" + ELWDropELWFM1ELWFileBulkGenerate.filename;

                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, emaPath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                GenerateExcelEmaFileTitle(wSheet);
                koreaList.Sort(new Common().CompareTicker);

                int startLine = 5;
                AppendEmaDataToFile(wSheet, startLine, "common");
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
                AddResult(Path.GetFileNameWithoutExtension(emaPath), emaPath, "ADD FM File(EMA)");
            }
            catch (Exception ex)
            {
                string msg = "Error found in GenerateELWFM1File_xls()    : " + ex.StackTrace + " \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void GenerateExcelFileTitle(Worksheet wSheet)
        {
            if (wSheet.get_Range("C4", Type.Missing).Value2 == null)
            {
                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 5;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 35;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 60;

                ((Range)wSheet.Columns["Q", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["R", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["S", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["A:S", System.Type.Missing]).Font.Name = "Arial";



                //set bold font to the title(first row)
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                //fill the cells with color 
                //((Range)wSheet.Cells[2, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Green);
                //modify the font color that in standard cells
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

                ((Range)wSheet.Cells[1, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                wSheet.Cells[1, 1] = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                ((Range)wSheet.Cells[3, 1]).Font.Underline = System.Drawing.FontStyle.Underline;
                ((Range)wSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                wSheet.Cells[3, 1] = "WARRANT ADD 1";

                wSheet.Cells[4, 1] = "Updated Date";
                wSheet.Cells[4, 2] = "Effective Date";
                wSheet.Cells[4, 3] = "RIC";
                wSheet.Cells[4, 4] = "FM";
                wSheet.Cells[4, 5] = "IDN Display Name";
                wSheet.Cells[4, 6] = "ISIN";
                wSheet.Cells[4, 7] = "Ticker";
                wSheet.Cells[4, 8] = "BCAST_REF";
                wSheet.Cells[4, 9] = "QA Common Name";
                wSheet.Cells[4, 10] = "Mat Date";
                wSheet.Cells[4, 11] = "Strike Price";
                wSheet.Cells[4, 12] = "Quanity of Warrants";
                wSheet.Cells[4, 13] = "Issue Price";
                wSheet.Cells[4, 14] = "Issue Date";
                wSheet.Cells[4, 15] = "Conversion Ratio";
                wSheet.Cells[4, 16] = "Issuer";
                wSheet.Cells[4, 17] = "Chain";
                wSheet.Cells[4, 18] = "Last Trading Date";
                wSheet.Cells[4, 19] = "Knock-out Price";
            }
        }

        private void AppendDataToFile(Worksheet wSheet, int startLine, string type)
        {
            try
            {
                foreach (var item in koreaList)
                {
                    ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                    wSheet.Cells[(startLine), 1] = item.UpdateDate;
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine, 2] = item.EffectiveDate;
                    wSheet.Cells[startLine, 3] = item.RIC;
                    wSheet.Cells[startLine, 4] = item.FM;
                    wSheet.Cells[startLine, 5] = item.IDNDisplayName;
                    ((Range)wSheet.Cells[startLine, 6]).NumberFormat = "@";
                    wSheet.Cells[startLine, 6] = item.ISIN;
                    ((Range)wSheet.Cells[startLine, 7]).NumberFormat = "@";
                    wSheet.Cells[startLine, 7] = item.Ticker;
                    wSheet.Cells[startLine, 8] = item.BCASTREF;
                    wSheet.Cells[startLine, 9] = item.QACommonName;
                    ((Range)wSheet.Cells[startLine, 10]).NumberFormat = "@";
                    wSheet.Cells[startLine, 10] = item.MatDate;
                    wSheet.Cells[startLine, 11] = item.StrikePrice;
                    wSheet.Cells[startLine, 12] = item.QuanityofWarrants;
                    wSheet.Cells[startLine, 13] = item.IssuePrice;
                    ((Range)wSheet.Cells[startLine, 14]).NumberFormat = "@";
                    wSheet.Cells[startLine, 14] = item.IssueDate;
                    wSheet.Cells[startLine, 15] = item.ConversionRatio;
                    wSheet.Cells[startLine, 16] = item.Issuer;
                    if (type.Equals("master"))
                    {
                        wSheet.Cells[startLine, 17] = item.KoreaWarrantName;
                        wSheet.Cells[startLine, 18] = item.Chain;
                    }
                    else
                    {
                        wSheet.Cells[startLine, 17] = item.Chain;
                    }
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in AppendDataToFile()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }
        }


        private void GenerateExcelEmaFileTitle(Worksheet wSheet)
        {
            if (wSheet.get_Range("C4", Type.Missing).Value2 == null)
            {
                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 5;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 35;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 60;

                ((Range)wSheet.Columns["Q", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["R", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["S", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["A:S", System.Type.Missing]).Font.Name = "Arial";



                //set bold font to the title(first row)
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                //fill the cells with color 
                //((Range)wSheet.Cells[2, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Green);
                //modify the font color that in standard cells
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

                ((Range)wSheet.Cells[1, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                wSheet.Cells[1, 1] = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                ((Range)wSheet.Cells[3, 1]).Font.Underline = System.Drawing.FontStyle.Underline;
                ((Range)wSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                wSheet.Cells[3, 1] = "WARRANT ADD 1";

                wSheet.Cells[4, 1] = "Updated Date";
                wSheet.Cells[4, 2] = "Effective Date";
                wSheet.Cells[4, 3] = "RIC";
                wSheet.Cells[4, 4] = "ISIN";
                wSheet.Cells[4, 5] = "Mat Date";
                wSheet.Cells[4, 6] = "Issue Date";

            }
        }

        private void AppendEmaDataToFile(Worksheet wSheet, int startLine, string type)
        {
            try
            {
                foreach (var item in koreaList)
                {
                    ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                    wSheet.Cells[(startLine), 1] = item.UpdateDate;
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine, 2] = item.EffectiveDate;
                    wSheet.Cells[startLine, 3] = item.RIC;
                    wSheet.Cells[startLine, 4] = item.ISIN;
                    wSheet.Cells[startLine, 5] = item.MatDate;
                    wSheet.Cells[startLine, 6] = item.IssueDate;
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in AppendDataToFile()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }
        }

        /*================================================= Files ==================================================*/

        private string CreateFolder(string foldername)
        {
            string ipath = string.Empty;
            try
            {
                ipath = Path.Combine(configObj.BulkFile, foldername);
                Common.CreateDirectory(ipath);
            }
            catch (Exception ex)
            {
                string msg = "Connot create the folder for bulk file  : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return ipath;
        }

        //NDA IA File
        private void GenerateNDAIAFileXls(string ipath)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
            }

            try
            {
                string filename = "kr" + DateTime.Today.ToString("yyyy-MM-dd").Replace("-", "") + "IAWntAdd.csv";
                ipath = Path.Combine(ipath, filename);
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                //Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet1", wBook);
                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                }
                wSheet.Name = "kr" + DateTime.Today.ToString("yyyy-MM-dd").Replace("-", "") + "IAWntAdd";

                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 17;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 12;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 35;
                ((Range)wSheet.Columns["A:E", System.Type.Missing]).Font.Name = "Arial";
                ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                }
                wSheet.Cells[1, 1] = "ISIN";
                wSheet.Cells[1, 2] = "TYPE";
                wSheet.Cells[1, 3] = "CATEGORY";
                wSheet.Cells[1, 4] = "WARRANT ISSUER";
                wSheet.Cells[1, 5] = "RCS ASSET CLASS";
                wSheet.Cells[1, 6] = "WARRANT ISSUE QUANTITY";

                int startLine = 2;
                LoopPrintNDAIAFile(wSheet, startLine, koreaList);

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.SaveAs(wBook.FullName, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(ipath), ipath, "ADD NDA IA File");
            }
            catch (Exception ex)
            {
                string msg = "Error found in NDA IA file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void LoopPrintNDAIAFile(Worksheet wSheet, int startLine, List<WarrantTemplate> lists)
        {
            try
            {
                for (var i = 0; i < lists.Count; i++)
                {
                    wSheet.Cells[startLine, 1] = lists[i].ISIN;
                    wSheet.Cells[startLine, 2] = "DERIVATIVE";
                    wSheet.Cells[startLine, 3] = "EIW";
                    wSheet.Cells[startLine, 4] = GetNDAIssuerORGID(lists, i);
                    wSheet.Cells[startLine, 5] = "TRAD";
                    wSheet.Cells[startLine, 6] = lists[i].QuanityofWarrants;
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in LoopPrintNDAIAFile()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        //NDA QA File
        private void GenerateNDAQAFileXls(string ipath)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
            }

            try
            {
                string filename = "kr" + DateTime.Today.ToString("yyyy-MM-dd").Replace("-", "") + "QAWntAdd.csv";
                ipath = Path.Combine(ipath, filename);
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                //Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet1", wBook);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                }
                wSheet.Name = "kr" + DateTime.Today.ToString("yyyy-MM-dd").Replace("-", "") + "QAWntAdd";

                GenerateQAExcelTitle(wSheet);

                int startLine = 2;
                LoopPrintNDAQAFile(wSheet, startLine, koreaList, "QA");
                startLine += koreaList.Count;
                LoopPrintNDAQAFile(wSheet, startLine, koreaList, "QAF");
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.SaveAs(wBook.FullName, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(ipath), ipath, "ADD NDA QA File");
            }
            catch (Exception ex)
            {
                string msg = "Error found in NDA QA file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        //NDA QAF File
        //private void GenerateNDAQAFFile_xls(string ipath)
        //{
        //    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        //    ExcelApp excelApp = new ExcelApp(false, false);
        //    if (excelApp.ExcelAppInstance == null)
        //    {
        //        string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
        //        Logger.Log(msg, Logger.LogType.Error);
        //    }

        //    try
        //    {
        //        string filename = "kr" + DateTime.Today.ToString("yyyy-MM-dd").Replace("-", "") + "QAWntAddF.csv";
        //        ipath += filename;
        //        Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
        //        Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet1", wBook);
        //        if (wSheet == null)
        //        {
        //            string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
        //            Logger.Log(msg, Logger.LogType.Error);
        //        }

        //        wSheet.Name = "kr" + DateTime.Today.ToString("yyyy-MM-dd").Replace("-", "") + "QAWntAddF";

        //        GenerateQAExcelTitle(wSheet);

        //        int startLine = 2;
        //        LoopPrintNDAQAFile(wSheet, startLine, koreaList, "QAF");

        //        excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
        //        wBook.SaveAs(wBook.FullName, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
        //        AddResult(Path.GetFileNameWithoutExtension(ipath),ipath,"");
        //    }
        //    catch (Exception ex)
        //    {
        //        string msg = "Error found in NDA QA file :" + ex.ToString();
        //        Logger.Log(msg, Logger.LogType.Error);
        //    }
        //    finally
        //    {
        //        excelApp.Dispose();
        //    }
        //}

        private static void GenerateQAExcelTitle(Worksheet wSheet)
        {
            ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 13;
            ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 10;
            ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 12;
            ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 12;
            ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 12;
            ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 45;
            ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 25;
            ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 17;
            ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 13;
            ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["A:P", System.Type.Missing]).Font.Name = "Arial";
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

            wSheet.Cells[1, 1] = "RIC";
            wSheet.Cells[1, 2] = "TAG";
            wSheet.Cells[1, 3] = "TYPE";
            wSheet.Cells[1, 4] = "CATEGORY";
            wSheet.Cells[1, 5] = "EXCHANGE";
            wSheet.Cells[1, 6] = "CURRENCY";
            wSheet.Cells[1, 7] = "ASSET COMMON NAME";
            wSheet.Cells[1, 8] = "ASSET SHORT NAME";
            wSheet.Cells[1, 9] = "CALL PUT OPTION";
            wSheet.Cells[1, 10] = "STRIKE PRICE";
            wSheet.Cells[1, 11] = "WARRANT ISSUE PRICE";
            wSheet.Cells[1, 12] = "ROUND LOT SIZE";
            wSheet.Cells[1, 13] = "EXPIRY DATE";
            wSheet.Cells[1, 14] = "TICKER SYMBOL";
            wSheet.Cells[1, 15] = "BASE ASSET";
        }

        private void LoopPrintNDAQAFile(Worksheet wSheet, int startLine, List<WarrantTemplate> lists, string type)
        {
            try
            {
                for (var i = 0; i < lists.Count; i++)
                {
                    wSheet.Cells[startLine, 1] = lists[i].RIC;
                    if (type.Equals("QA"))
                    {
                        wSheet.Cells[startLine, 1] = lists[i].RIC;
                        wSheet.Cells[startLine, 2] = "46429";
                    }
                    else if (type.Equals("QAF"))
                    {
                        string ric = lists[i].RIC.Insert(6, "F");
                        wSheet.Cells[startLine, 1] = ric;
                        wSheet.Cells[startLine, 2] = "44398";
                    }
                    wSheet.Cells[startLine, 3] = "DERIVATIVE";
                    wSheet.Cells[startLine, 4] = "EIW";
                    wSheet.Cells[startLine, 5] = "KSC";
                    wSheet.Cells[startLine, 6] = "KRW";
                    wSheet.Cells[startLine, 7] = lists[i].QACommonName;
                    wSheet.Cells[startLine, 8] = lists[i].IDNDisplayName;
                    wSheet.Cells[startLine, 9] = lists[i].CallOrPut;
                    wSheet.Cells[startLine, 10] = lists[i].StrikePrice;
                    wSheet.Cells[startLine, 11] = lists[i].IssuePrice;
                    wSheet.Cells[startLine, 12] = "10";
                    ((Range)wSheet.Cells[startLine, 13]).NumberFormat = "@";
                    wSheet.Cells[startLine, 13] = Convert.ToDateTime(lists[i].OrgMatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
                    wSheet.Cells[startLine, 14] = lists[i].Ticker;
                    wSheet.Cells[startLine, 15] = "ISIN:" + lists[i].ISIN;
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in LoopPrintNDAQAFile()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        //NDA T&C File
        private void GenerateNDATCFile_xls(string ipath)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
            }

            try
            {
                string filename = "WRT_ADD_" + DateTime.Today.ToString("dd-MMM-yyyy", new CultureInfo("en-US")).Replace("-", "") + "_Korea.csv";
                ipath = Path.Combine(CreateFM1FileEmaDir(), filename);
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                }
                wSheet.Name = "WRT_ADD_" + DateTime.Today.ToString("dd-MMM-yyyy", new CultureInfo("en-US")).Replace("-", "") + "_Korea";

                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 16;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 70;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 12;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 12;
                ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 23;
                ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 12;
                ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["Q", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["R", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["S", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["T", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["U", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["V", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["W", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["X", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["Y", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["Z", System.Type.Missing]).ColumnWidth = 20;

                ((Range)wSheet.Columns["AA", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AB", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AC", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AD", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["AE", System.Type.Missing]).ColumnWidth = 15;
                //((Range)wSheet.Columns["AF", System.Type.Missing]).NumberFormat = "dd/MM/yyyy";
                ((Range)wSheet.Columns["AF", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["AG", System.Type.Missing]).ColumnWidth = 23;
                ((Range)wSheet.Columns["AH", System.Type.Missing]).ColumnWidth = 21;
                ((Range)wSheet.Columns["AI", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AJ", System.Type.Missing]).ColumnWidth = 17;
                ((Range)wSheet.Columns["AK", System.Type.Missing]).ColumnWidth = 17;
                ((Range)wSheet.Columns["AL", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AM", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AN", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["AO", System.Type.Missing]).ColumnWidth = 12;
                ((Range)wSheet.Columns["AP", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["AQ", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["AR", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AS", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AT", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AU", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AV", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AW", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AX", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AY", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["AZ", System.Type.Missing]).ColumnWidth = 24;

                ((Range)wSheet.Columns["BA", System.Type.Missing]).ColumnWidth = 22;
                ((Range)wSheet.Columns["BB", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BC", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["BD", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["BE", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["BF", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["BG", System.Type.Missing]).ColumnWidth = 22;
                ((Range)wSheet.Columns["BH", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["BI", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BJ", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["BK", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BL", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["BM", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BN", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["BO", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["BP", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["BQ", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BR", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["BS", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BT", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BU", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BV", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BW", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BX", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BY", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["BZ", System.Type.Missing]).ColumnWidth = 20;

                ((Range)wSheet.Columns["CA", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["CB", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["CC", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["CD", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["CE", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["CF", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["CG", System.Type.Missing]).ColumnWidth = 33;
                ((Range)wSheet.Columns["CH", System.Type.Missing]).ColumnWidth = 12;
                ((Range)wSheet.Columns["CI", System.Type.Missing]).ColumnWidth = 105;
                ((Range)wSheet.Columns["CJ", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CK", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CL", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["CM", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CN", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["CO", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CP", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CQ", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CR", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CS", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CT", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CU", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CV", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CW", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CX", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CY", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["CZ", System.Type.Missing]).ColumnWidth = 20;

                ((Range)wSheet.Columns["DA", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["DB", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["DC", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["DD", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["DE", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["DF", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["DG", System.Type.Missing]).ColumnWidth = 22;
                ((Range)wSheet.Columns["DH", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["DI", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DJ", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DK", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DL", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["DM", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DN", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["DO", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DP", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DQ", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DR", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DS", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DT", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DU", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DV", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DW", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["DX", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["DY", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["DZ", System.Type.Missing]).ColumnWidth = 18;

                ((Range)wSheet.Columns["EA", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["EB", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["EC", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["ED", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["EE", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["EF", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["EG", System.Type.Missing]).ColumnWidth = 45;
                ((Range)wSheet.Columns["EH", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["EI", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["EJ", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["EK", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["EL", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["EM", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["EN", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["EO", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["EP", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["EQ", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["ER", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["ES", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["ET", System.Type.Missing]).ColumnWidth = 20;

                ((Range)wSheet.Columns["A:ET", System.Type.Missing]).Font.Name = "Arail";
                ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                }
                wSheet.Cells[1, 1] = "Logical_Key";
                wSheet.Cells[1, 2] = "Secondary_ID";
                wSheet.Cells[1, 3] = "Secondary_ID_Type";
                wSheet.Cells[1, 4] = "Warrant_Title";
                wSheet.Cells[1, 5] = "Issuer_OrgId";
                wSheet.Cells[1, 6] = "Issue_Date";
                wSheet.Cells[1, 7] = "Country_Of_Issue";
                wSheet.Cells[1, 8] = "Governing_Country";
                wSheet.Cells[1, 9] = "Announcement_Date";
                wSheet.Cells[1, 10] = "Payment_Date";
                wSheet.Cells[1, 11] = "Underlying_Type";
                wSheet.Cells[1, 12] = "Clearinghouse1_OrgId";
                wSheet.Cells[1, 13] = "Clearinghouse2_OrgId";
                wSheet.Cells[1, 14] = "Clearinghouse3_OrgId";
                wSheet.Cells[1, 15] = "Guarantor";
                wSheet.Cells[1, 16] = "Guarantor_Type";
                wSheet.Cells[1, 17] = "Guarantee_Type";
                wSheet.Cells[1, 18] = "Incr_Exercise_Lot";
                wSheet.Cells[1, 19] = "Min_Exercise_Lot";
                wSheet.Cells[1, 20] = "Max_Exercise_Lot";
                wSheet.Cells[1, 21] = "Rt_Page_Range";
                wSheet.Cells[1, 22] = "Underwriter1_OrgId";
                wSheet.Cells[1, 23] = "Underwriter1_Role";
                wSheet.Cells[1, 24] = "Underwriter2_OrgId";
                wSheet.Cells[1, 25] = "Underwriter2_Role";
                wSheet.Cells[1, 26] = "Underwriter3_OrgId";
                wSheet.Cells[1, 27] = "Underwriter3_Role";
                wSheet.Cells[1, 28] = "Underwriter4_OrgId";
                wSheet.Cells[1, 29] = "Underwriter4_Role";
                wSheet.Cells[1, 30] = "Exercise_Style";
                wSheet.Cells[1, 31] = "Warrant_Type";
                wSheet.Cells[1, 32] = "Expiration_Date";
                wSheet.Cells[1, 33] = "Registered_Bearer_Code";
                wSheet.Cells[1, 34] = "Price_Display_Type";
                wSheet.Cells[1, 35] = "Private_Placement";
                wSheet.Cells[1, 36] = "Coverage_Type";

                wSheet.Cells[1, 37] = "Warrant_Status";
                wSheet.Cells[1, 38] = "Status_Date";
                wSheet.Cells[1, 39] = "Redemption_Method";
                wSheet.Cells[1, 40] = "Issue_Quantity";
                wSheet.Cells[1, 41] = "Issue_Price";
                wSheet.Cells[1, 42] = "Issue_Currency";
                wSheet.Cells[1, 43] = "Issue_Price_Type";
                wSheet.Cells[1, 44] = "Issue_Spot_Price";
                wSheet.Cells[1, 45] = "Issue_Spot_Currency";
                wSheet.Cells[1, 46] = "Issue_Spot_FX_Rate";
                wSheet.Cells[1, 47] = "Issue_Delta";
                wSheet.Cells[1, 48] = "Issue_Elasticity";
                wSheet.Cells[1, 49] = "Issue_Gearing";
                wSheet.Cells[1, 50] = "Issue_Premium";
                wSheet.Cells[1, 51] = "Issue_Premium_PA";
                wSheet.Cells[1, 52] = "Denominated_Amount";
                wSheet.Cells[1, 53] = "Exercise_Begin_Date";
                wSheet.Cells[1, 54] = "Exercise_End_Date";
                wSheet.Cells[1, 55] = "Offset_Number";
                wSheet.Cells[1, 56] = "Period_Number";
                wSheet.Cells[1, 57] = "Offset_Frequency";
                wSheet.Cells[1, 58] = "Offset_Calendar";
                wSheet.Cells[1, 59] = "Period_Calendar";
                wSheet.Cells[1, 60] = "Period_Frequency";
                wSheet.Cells[1, 61] = "RAF_Event_Type";
                wSheet.Cells[1, 62] = "Exercise_Price";
                wSheet.Cells[1, 63] = "Exercise_Price_Type";
                wSheet.Cells[1, 64] = "Warrants_Per_Underlying";
                wSheet.Cells[1, 65] = "Underlying_FX_Rate";
                wSheet.Cells[1, 66] = "Underlying_RIC";
                wSheet.Cells[1, 67] = "Underlying_Item_Quantity";
                wSheet.Cells[1, 68] = "Units";
                wSheet.Cells[1, 69] = "Cash_Currency";
                wSheet.Cells[1, 70] = "Delivery_Type";
                wSheet.Cells[1, 71] = "Settlement_Type";
                wSheet.Cells[1, 72] = "Settlement_Currency";
                wSheet.Cells[1, 73] = "Underlying_Group";
                wSheet.Cells[1, 74] = "Country1_Code";
                wSheet.Cells[1, 75] = "Coverage1_Type";
                wSheet.Cells[1, 76] = "Country2_Code";
                wSheet.Cells[1, 77] = "Coverage2_Type";
                wSheet.Cells[1, 78] = "Country3_Code";
                wSheet.Cells[1, 79] = "Coverage3_Type";
                wSheet.Cells[1, 80] = "Country4_Code";
                wSheet.Cells[1, 81] = "Coverage4_Type";
                wSheet.Cells[1, 82] = "Country5_Code";
                wSheet.Cells[1, 83] = "Coverage5_Type";
                wSheet.Cells[1, 84] = "Note1_Type";
                wSheet.Cells[1, 85] = "Note1";
                wSheet.Cells[1, 86] = "Note2_Type";
                wSheet.Cells[1, 87] = "Note2";
                wSheet.Cells[1, 88] = "Note3_Type";
                wSheet.Cells[1, 89] = "Note3";
                wSheet.Cells[1, 90] = "Note4_Type";
                wSheet.Cells[1, 91] = "Note4";
                wSheet.Cells[1, 92] = "Note5_Type";
                wSheet.Cells[1, 93] = "Note5";
                wSheet.Cells[1, 94] = "Note6_Type";
                wSheet.Cells[1, 95] = "Note6";
                wSheet.Cells[1, 96] = "Exotic1_Parameter";
                wSheet.Cells[1, 97] = "Exotic1_Value";
                wSheet.Cells[1, 98] = "Exotic1_Begin_Date";
                wSheet.Cells[1, 99] = "Exotic1_End_Date";
                wSheet.Cells[1, 100] = "Exotic2_Parameter";
                wSheet.Cells[1, 101] = "Exotic2_Value";
                wSheet.Cells[1, 102] = "Exotic2_Begin_Date";
                wSheet.Cells[1, 103] = "Exotic2_End_Date";
                wSheet.Cells[1, 104] = "Exotic3_Parameter";
                wSheet.Cells[1, 105] = "Exotic3_Value";
                wSheet.Cells[1, 106] = "Exotic3_Begin_Date";
                wSheet.Cells[1, 107] = "Exotic3_End_Date";
                wSheet.Cells[1, 108] = "Exotic4_Parameter";
                wSheet.Cells[1, 109] = "Exotic4_Value";
                wSheet.Cells[1, 110] = "Exotic4_Begin_Date";
                wSheet.Cells[1, 111] = "Exotic4_End_Date";
                wSheet.Cells[1, 112] = "Exotic5_Parameter";
                wSheet.Cells[1, 113] = "Exotic5_Value";
                wSheet.Cells[1, 114] = "Exotic5_Begin_Date";
                wSheet.Cells[1, 115] = "Exotic5_End_Date";
                wSheet.Cells[1, 116] = "Exotic6_Parameter";
                wSheet.Cells[1, 117] = "Exotic6_Value";
                wSheet.Cells[1, 118] = "Exotic6_Begin_Date";
                wSheet.Cells[1, 119] = "Exotic6_End_Date";
                wSheet.Cells[1, 120] = "Event_Type1";
                wSheet.Cells[1, 121] = "Event_Period_Number1";
                wSheet.Cells[1, 122] = "Event_Calendar_Type1";
                wSheet.Cells[1, 123] = "Event_Frequency1";
                wSheet.Cells[1, 124] = "Event_Type2";
                wSheet.Cells[1, 125] = "Event_Period_Number2";
                wSheet.Cells[1, 126] = "Event_Calendar_Type2";
                wSheet.Cells[1, 127] = "Event_Frequency2";
                wSheet.Cells[1, 128] = "Exchange_Code1";
                wSheet.Cells[1, 129] = "Incr_Trade_Lot1";
                wSheet.Cells[1, 130] = "Min_Trade_Lot1";
                wSheet.Cells[1, 131] = "Min_Trade_Amount1";
                wSheet.Cells[1, 132] = "Exchange_Code2";
                wSheet.Cells[1, 133] = "Incr_Trade_Lot2";
                wSheet.Cells[1, 134] = "Min_Trade_Lot2";
                wSheet.Cells[1, 135] = "Min_Trade_Amount2";
                wSheet.Cells[1, 136] = "Exchange_Code3";
                wSheet.Cells[1, 137] = "Incr_Trade_Lot3";
                wSheet.Cells[1, 138] = "Min_Trade_Lot3";
                wSheet.Cells[1, 139] = "Min_Trade_Amount3";
                wSheet.Cells[1, 140] = "Exchange_Code4";
                wSheet.Cells[1, 141] = "Incr_Trade_Lot4";
                wSheet.Cells[1, 142] = "Min_Trade_Lot4";
                wSheet.Cells[1, 143] = "Min_Trade_Amount4";
                wSheet.Cells[1, 144] = "Attached_To_Id";
                wSheet.Cells[1, 145] = "Attached_To_Id_Type";
                wSheet.Cells[1, 146] = "Attached_Quantity";
                wSheet.Cells[1, 147] = "Attached_Code";
                wSheet.Cells[1, 148] = "Detachable_Date";
                wSheet.Cells[1, 149] = "Bond_Exercise";
                wSheet.Cells[1, 150] = "Bond_Price_Percentage";

                int startLine = 2;
                LoopPrintNDATANDCFile(wSheet, startLine, koreaList);

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.SaveAs(wBook.FullName, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(ipath), ipath, "ADD NDA TC File(EMA)");
            }
            catch (Exception ex)
            {
                string msg = "Error found in NDA T&C file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void LoopPrintNDATANDCFile(Worksheet wSheet, int startLine, List<WarrantTemplate> lists)
        {
            try
            {
                for (var i = 0; i < lists.Count; i++)
                {
                    wSheet.Cells[startLine, 1] = (startLine - 1).ToString();
                    wSheet.Cells[startLine, 2] = lists[i].ISIN;
                    wSheet.Cells[startLine, 3] = "ISIN";

                    string warrant_title = CombineWarrantTitle(lists, i).ToUpper();
                    wSheet.Cells[startLine, 4] = warrant_title;

                    string nda_orgid = GetNDAIssuerORGID(lists, i);
                    wSheet.Cells[startLine, 5] = nda_orgid;

                    string issue_date = GetIssuerDate(lists, i);
                    ((Range)wSheet.Cells[startLine, 6]).NumberFormat = "@";
                    wSheet.Cells[startLine, 6] = issue_date;

                    wSheet.Cells[startLine, 7] = "KOR";
                    wSheet.Cells[startLine, 8] = "KOR";

                    string indexstock = lists[i].QACommonName.Substring((lists[i].QACommonName.Length - 2)) == "IW" ? "INDEX" : "STOCK";
                    wSheet.Cells[startLine, 11] = indexstock;

                    wSheet.Cells[startLine, 18] = "10";
                    wSheet.Cells[startLine, 19] = "10";
                    wSheet.Cells[startLine, 30] = "E";

                    wSheet.Cells[startLine, 31] = lists[i].CallOrPut == "CALL" ? "Call" : "Put";

                    //T&C file    column "AF"

                    string expiraton_date = GetExpirationDate(lists, i);
                    ((Range)wSheet.Cells[startLine, 32]).NumberFormat = "dd/MM/yyyy";
                    wSheet.Cells[startLine, 32] = expiraton_date;


                    wSheet.Cells[startLine, 33] = "R";
                    wSheet.Cells[startLine, 34] = "D";
                    wSheet.Cells[startLine, 40] = lists[i].QuanityofWarrants;
                    wSheet.Cells[startLine, 41] = lists[i].IssuePrice;
                    wSheet.Cells[startLine, 42] = "KRW";
                    wSheet.Cells[startLine, 43] = "A";
                    wSheet.Cells[startLine, 52] = "10";
                    //T&C file    column "BA"
                    ((Range)wSheet.Cells[startLine, 53]).NumberFormat = "dd/MM/yyyy";
                    wSheet.Cells[startLine, 53] = expiraton_date;
                    //T&C file    column "BB"
                    ((Range)wSheet.Cells[startLine, 54]).NumberFormat = "dd/MM/yyyy";
                    wSheet.Cells[startLine, 54] = expiraton_date;
                    wSheet.Cells[startLine, 62] = lists[i].StrikePrice;
                    wSheet.Cells[startLine, 63] = "A";

                    string warrant_per_underlying = GetWarrantPerUnderlying(lists, i);
                    wSheet.Cells[startLine, 64] = warrant_per_underlying;

                    ((Range)wSheet.Cells[startLine, 66]).NumberFormat = "@";
                    wSheet.Cells[startLine, 66] = lists[i].BCASTREF;
                    wSheet.Cells[startLine, 67] = "1";



                    wSheet.Cells[startLine, 68] = indexstock == "INDEX" ? "idx" : "shr";
                    wSheet.Cells[startLine, 70] = indexstock == "INDEX" ? "I" : "S";
                    wSheet.Cells[startLine, 71] = "C";
                    wSheet.Cells[startLine, 72] = "KRW";

                    wSheet.Cells[startLine, 84] = "T";

                    string date = GetNote1(lists, i);
                    wSheet.Cells[startLine, 85] = "Last Trading Day is " + date + ".";

                    wSheet.Cells[startLine, 86] = "S";

                    string note2 = GetNote2(lists, i);
                    wSheet.Cells[startLine, 87] = note2;

                    wSheet.Cells[startLine, 128] = "KSC";
                    wSheet.Cells[startLine, 129] = "10";
                    wSheet.Cells[startLine, 130] = "10";

                    startLine++;
                }

            }
            catch (Exception ex)
            {
                string msg = "Error found in LoopPrintNDAT&CFile()      : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private string GetNote2(List<WarrantTemplate> lists, int i)
        {
            string note2 = string.Empty;
            try
            {
                string cp = lists[i].CallOrPut == "CALL" ? "Call" : "Put";
                string bcasetref = lists[i].BCASTREF;

                string strCombine = bcasetref + " " + cp;
                switch (strCombine)
                {
                    case ".KS200 Call":
                        note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Closing level on last trade date - Strike level)* 100";
                        break;
                    case ".KS200 Put":
                        note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Strike level - Closing level on last trade date)* 100";
                        break;
                    case ".N225 Call":
                        note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Closing level on last trade date - Strike level)* Conversion ratio";
                        break;
                    case ".N225 Put":
                        note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Strike level - Closing level on last trade date)* Conversion ratio";
                        break;
                    case ".HSI Call":
                        note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Average closing level during the last 5 trading days(including last trading day) - Strike level)* Conversion ratio";
                        break;
                    case ".HSI Put":
                        note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Strike level - Average closing level during the last 5 trading days(including last trading day))* Conversion ratio";
                        break;
                    default:
                        note2 = "Average closing price of the underlying stock during the last 5 trading days(including last trading day) before expiry.";
                        break;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetNode2()  : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return note2;
        }

        private string GetNote1(List<WarrantTemplate> lists, int i)
        {
            string date = string.Empty;
            try
            {

                string matdate = lists[i].OrgMatDate;
                string weekday = Convert.ToDateTime(matdate).DayOfWeek.ToString();

                if (weekday == "Monday" || weekday == "Tuesday")
                    date = Convert.ToDateTime(matdate).AddDays(-4).ToString("dd-MMM-yyyy");
                else
                    date = Convert.ToDateTime(matdate).AddDays(-2).ToString("dd-MMM-yyyy");
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetNode1()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return date;
        }

        private string GetWarrantPerUnderlying(List<WarrantTemplate> lists, int i)
        {
            double ratio = Convert.ToDouble(lists[i].ConversionRatio);
            string warrant_per_underlying = (1.0 / ratio).ToString("0.000000000").TrimEnd('0');
            try
            {
                int len = warrant_per_underlying.Split('.')[0].Length;
                double d = Math.Round(Convert.ToDouble(warrant_per_underlying), (10 - len));
                warrant_per_underlying = d.ToString();
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetWarrant_Per_Underlying()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return warrant_per_underlying;
        }

        private string GetExpirationDate(List<WarrantTemplate> lists, int i)
        {
            string expiraton_date = lists[i].OrgMatDate;
            expiraton_date = Convert.ToDateTime(expiraton_date).ToString("dd-MM-yyyy").Replace('-', '/');
            return expiraton_date;
        }

        private string GetIssuerDate(List<WarrantTemplate> lists, int i)
        {
            string issue_date = lists[i].OrgIssueDate;
            issue_date = Convert.ToDateTime(issue_date).ToString("dd-MM-yyyy").Replace('-', '/');
            return issue_date;
        }

        private string GetNDAIssuerORGID(List<WarrantTemplate> lists, int i)
        {
            string ndaOrgid = string.Empty;
            try
            {
                KoreaIssuerInfo issuer = KoreaIssuerManager.SelectIssuer(lists[i].IssuerKoreaName);
                if (issuer != null)
                {
                    ndaOrgid = issuer.NDAIssuerOrgid;
                }
                else
                    ndaOrgid = "Issuer中不存在NDA_Issuer_ORGID";
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetNDA_Issuer_ORGID()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return ndaOrgid;
        }

        private string CombineWarrantTitle(List<WarrantTemplate> lists, int i)
        {
            string warrant_title = string.Empty;
            try
            {
                string issuer_nda_tc = string.Empty;
                KoreaIssuerInfo issuer = KoreaIssuerManager.SelectIssuer(lists[i].IssuerKoreaName);
                if (issuer != null)
                {
                    issuer_nda_tc = issuer.NDATCIssuerTitle;
                }
                else
                    issuer_nda_tc = "Issuer中不存在该NDA_T&C";

                string underlying_nda_tc = string.Empty;
                string underlyKoreaName = lists[i].UnderlyingKoreaName;
                KoreaUnderlyingInfo underlying = KoreaUnderlyingManager.SelectUnderlying(underlyKoreaName);
                if (underlying != null)
                {
                    underlying_nda_tc = underlying.NDATCUnderlyingTitle;
                }
                else
                {
                    underlying_nda_tc = "Underlying中不存在该NDA_T&C";
                }

                string charLen = lists[i].QACommonName.Substring(lists[i].QACommonName.Length - 2) == "IW" ? "INDEX" : "SHS";
                string callput = lists[i].CallOrPut == "CALL" ? "CALL WTS" : "PUT WTS";
                string mat = lists[i].OrgMatDate;
                mat = Convert.ToDateTime(mat).ToString("dd-MMM-yyyy");
                warrant_title = issuer_nda_tc + "/ " + underlying_nda_tc + " " + charLen + " " + callput + " " + mat;
            }
            catch (Exception ex)
            {
                string msg = "Error found in CombineWarrant_Title()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return warrant_title;
        }

        //KSE_ELW File
        private void GenerateKSEELWFileTxt(string ipath)
        {
            try
            {
                string filename = "KSE_ELW_" + DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US")).Replace("-", "") + ".txt";
                ipath = Path.Combine(ipath, filename);
                string[] content = new string[2];
                content[0] = "SYMBOL\tDSPLY_NAME\tRIC\tOFFCL_CODE\tEX_SYMBOL\tBCAST_REF\tEXL_NAME\t#INSTMOD_TDN_SYMBOL\t#INSTMOD_#ISIN\t#INSTMOD_MNEMONIC\tDSPLY_NMLL\tMATUR_DATE\t#INSTMOD_LONGLINK2\tBCU";

                StringBuilder sb = new StringBuilder();
                for (var i = 0; i < koreaList.Count; i++)
                {
                    WarrantTemplate warrTemp = koreaList[i] as WarrantTemplate;
                    sb.Append(warrTemp.RIC + "\t");
                    sb.Append(warrTemp.IDNDisplayName + "\t");
                    sb.Append(warrTemp.RIC + "\t");
                    sb.Append(warrTemp.Ticker + "\t");
                    sb.Append(warrTemp.ISIN + "\t");
                    sb.Append(warrTemp.BCASTREF + "\t");
                    sb.Append("KSE_EQB_ELW\t");
                    sb.Append(warrTemp.Ticker + "\t");
                    sb.Append(warrTemp.ISIN + "\t");
                    sb.Append("J" + warrTemp.Ticker + "\t");

                    string dsply = string.Empty;
                    string issue_dsply_nmll = warrTemp.IssuerKoreaName;
                    string _CharLen = warrTemp.IDNDisplayName.Substring(4, 4);

                    string underly_dsply_nmll = string.Empty;
                    string underlyingKoreaName = warrTemp.UnderlyingKoreaName;
                    KoreaUnderlyingInfo underlying = KoreaUnderlyingManager.SelectUnderlying(underlyingKoreaName);
                    if (underlying != null)
                    {
                        underly_dsply_nmll = underlying.UnderlyingName;
                    }
                    else
                    {
                        underly_dsply_nmll = "出现异常信息... ...";
                    }

                    //KRA5031A41C8
                    string _CharLast = warrTemp.IDNDisplayName.Substring((warrTemp.IDNDisplayName.Length - 1));
                    if (_CharLast == "C")
                        dsply = issue_dsply_nmll + _CharLen + underly_dsply_nmll + "콜";
                    else
                        dsply = issue_dsply_nmll + _CharLen + underly_dsply_nmll + "풋";

                    sb.Append(dsply + "\t");

                    sb.Append(Convert.ToDateTime(warrTemp.OrgMatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                    sb.Append("\t");
                    sb.Append("ELW/" + warrTemp.Ticker);
                    sb.Append("\t");
                    sb.Append("KSE_EQ_IPOELW");
                    sb.Append("\r\n");
                }
                content[1] = sb.ToString();
                File.WriteAllLines(ipath, content, Encoding.UTF8);
                //File.WriteAllText(ipath, sb.ToString(), Encoding.UTF8);
                AddResult(Path.GetFileNameWithoutExtension(ipath), ipath, "KSE ELW File");
            }
            catch (Exception ex)
            {
                string msg = "Error found in GenerateKSE_ELW_txtFile()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        #region For New Underlying: Grab new underlying data and generate 3 files and send email

        /// <summary>
        /// Encode text to bits
        /// </summary>
        /// <param name="encodeText">text to encode</param>
        /// <returns>encoded result</returns>
        private string Encode(string encodeText)
        {
            return Regex.Replace(encodeText, "[^a-zA-Z0-9]", delegate(Match match) { return "%" + BitConverter.ToString(Encoding.GetEncoding("euc-kr").GetBytes(match.Value)).Replace("-", "%"); });
        }

        /// <summary>
        /// Search the ISIN website for given korean name and key word: 보통주 first time: search with koreanName + 보통주. second time: search with koreanName
        /// If mutiple records or no record found. User need to input ISIN. 
        /// </summary>
        /// <param name="koreanName">korean name</param>
        /// <param name="times">searched times</param>
        /// <returns>isin</returns>
        private string GetIsinByName(string koreanName, int times)
        {
            string postData = "";
            string encodeName = Encode(koreanName);
            if (koreanName.Contains("-"))
            {
                encodeName = Encode(koreanName.Split('-')[1]);
            }

            if (times == 0)
            {
                postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun=01&lst_yn1=Y&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word={0}&ef_iss_inst_cd=&ef_isu_nm={0}%BA%B8%C5%EB%C1%D6&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=", encodeName);
            }
            else
            {
                postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun=01&lst_yn1=Y&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word={0}&ef_iss_inst_cd=&ef_isu_nm=%BA%B8%C5%EB%C1%D6&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=", encodeName);
            }
            string uri = "http://isin.krx.co.kr/jsp/realBoard01.jsp";
            HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
            request.Timeout = 100000;
            request.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:6.0.2) Gecko/20100101 Firefox/6.0.2";
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Referer = "http://isin.krx.co.kr/jsp/BA_LT113.jsp";
            request.Headers.Add("Accept-Encoding: gzip,deflate,sdch");

            string pageSource = null;
            int retries = 3;
            while (pageSource == null && retries-- > 0)
            {
                try
                {
                    pageSource = WebClientUtil.GetDynamicPageSource(request, postData, Encoding.GetEncoding("EUC-KR"));
                }
                catch
                {
                    System.Threading.Thread.Sleep(5000);
                }
            }
            string isin = "";
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(pageSource);

            HtmlNodeCollection trs = null;
            try
            {
                trs = doc.DocumentNode.SelectNodes("/html[1]/body[1]/table[1]/tr[2]/td[1]/table[1]/tr");
            }
            catch
            {
                string msg = "Error found in searching new underlying record found for " + koreanName;
                Logger.Log(msg, Logger.LogType.Error);
                isin = InputISIN.Prompt(koreanName, "Underlying Name");
            }
            //not found
            if (trs == null || trs.Count > 2)
            {
                isin = InputISIN.Prompt(koreanName, "Underlying Name");
            }
            else if (trs.Count == 1)
            {
                if (times == 0)
                {
                    isin = GetIsinByName(koreanName, 1);
                }
                else
                {
                    isin = InputISIN.Prompt(koreanName, "Underlying Name");
                }
            }
            //find only one record
            else if (trs.Count == 2)
            {
                HtmlNode tr = trs[1];
                HtmlNodeCollection tds = tr.SelectNodes("./td");
                isin = tds[1].InnerText.Trim();
            }
            return isin;
        }

        /// <summary>
        /// Grab new underlying info with korean name.
        /// </summary>
        /// <param name="koreaName">korean name</param>
        /// <returns>new underlying info</returns>
        private KoreaUnderlyingInfo GrabNewUnderlyingInfo(string koreaName)
        {
            string isin = GetIsinByName(koreaName, 0);
            if (string.IsNullOrEmpty(isin))
            {
                string msg1 = "No matched isin found for " + koreaName;
                Logger.Log(msg1, Logger.LogType.Warning);
                return null;
            }
            KoreaUnderlyingInfo newUnderlying = new KoreaUnderlyingInfo();
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
                string msg = "Can not get the New Underlying infos in ISIN webpage. For ISIN:" + isin + ". please check if the webpage can be accessed!";
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }
            HtmlDocument isinRoot = new HtmlDocument();
            isinRoot.LoadHtml(pageSource);
            HtmlNode isinTable = isinRoot.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[1]/tr[3]/td[1]/table[1]");
            HtmlNodeCollection isinTrs = isinTable.SelectNodes("./tr");

            string ric = isinTrs[2].SelectNodes("./td")[3].InnerText.TrimStart().TrimEnd();
            string sixDigit = ric.Substring(ric.Length - 6);
            string underEngName = isinTrs[10].SelectNodes("./td")[1].InnerText.TrimStart().TrimEnd();
            string suffix = string.IsNullOrEmpty(isinTrs[11].SelectNodes("./td")[2].InnerText.TrimStart().TrimEnd()) ? "KQ" : "KS";
            string usName = isinTrs[10].SelectNodes("./td")[3].InnerText.Trim();
            string symbol = isinTrs[3].SelectNodes("./td")[1].InnerText.Trim();

            newUnderlying.UnderlyingRIC = sixDigit + "." + suffix;

            sixDigit = "kr;" + sixDigit.TrimStart('0').TrimEnd('0');
            if (suffix == "KQ")
            {
                sixDigit += "K";
            }
            string ndaTc = ClearCoLtdForName(underEngName.ToUpper());
            newUnderlying.QACommonNamePart = ndaTc;
            newUnderlying.NDATCUnderlyingTitle = ndaTc;
            newUnderlying.BNDUnderlying = sixDigit;
            newUnderlying.KoreaName = koreaName;
            newUnderlying.KoreaNameFM2 = koreaName;
            //newUnderlying.KoreaNameDrop = koreaName;           
            newUnderlying.IDNDisplayNamePart = GetIDNDisplayName(symbol, usName, ndaTc);

            return newUnderlying;
        }

        /// <summary>
        /// Get IDN display name by symbol or usName or company name. 
        /// If the database contains a same display name, then change it.
        /// </summary>
        /// <param name="symbol"></param>
        /// <param name="usName"></param>
        /// <param name="ndaTc"></param>
        /// <returns></returns>
        private string GetIDNDisplayName(string symbol, string usName, string ndaTc)
        {
            //symbol usName<=7digits 取length长的
            //usName>7 symbol!= "" 取symbol  
            // symbol=null   1. 1 word 前四后三  2. >2 words  1四2三
            //判断是否重复 ，若重复，用ndatc生成
            usName = usName.ToUpper();
            string result = string.Empty;
            string nameNoBlank = Regex.Replace(usName, "([ ]+)", "");
            if (nameNoBlank.Length <= 7)
            {
                result = nameNoBlank.Length >= symbol.Length ? nameNoBlank : symbol;
            }
            else
            {
                if (!string.IsNullOrEmpty(symbol))
                {
                    result = symbol;
                }
                else
                {
                    nameNoBlank = Regex.Replace(nameNoBlank, "([^0-9A-Z]+)", "");
                    Regex regex = new Regex("([0-9]+)");
                    MatchCollection match = regex.Matches(nameNoBlank);
                    if (match.Count > 0)
                    {
                        string subNum = "";
                        Dictionary<int, string> numDe = new Dictionary<int, string>();
                        string nameNoDigit = Regex.Replace(nameNoBlank, "([^A-Z]+)", "");
                        int totalLength = 0;
                        for (int i = match.Count - 1; i >= 0; i--)
                        {
                            totalLength += match[i].Length;
                            if ((match[i].Index + match[i].Length) > 7)
                            {
                                subNum = match[i].Value.Trim() + subNum;
                            }
                            else
                            {
                                numDe.Add(match[i].Index, match[i].Value);
                            }
                        }
                        nameNoDigit = nameNoDigit.Substring(0, 7 - totalLength);
                        foreach (var item in numDe)
                        {
                            nameNoDigit = nameNoDigit.Insert(item.Key, item.Value);
                        }
                        result = nameNoDigit + subNum;
                    }
                    else
                    {
                        result = nameNoBlank.Substring(0, 7);
                    }
                }
            }

            int retry = 1;
            while (KoreaUnderlyingManager.ExsitDisplayName(result))
            {
                result = GetIDNDisplayNamePortion(ndaTc, retry++);
            }
            return result;
        }



        /// <summary>
        /// Use some rules to generate IDN display name for new underlying. And check the DB if it is unique.
        /// </summary>
        /// <param name="usName"></param>
        /// <returns></returns>
        private string GetUniqueIdnName(string usName, string companyName)
        {
            usName = usName.ToUpper();
            string result = "";
            string nameNoBlank = Regex.Replace(usName, "([ ]+)", "");
            if (nameNoBlank.Length <= 7)
            {
                result = nameNoBlank;
            }
            else
            {
                nameNoBlank = Regex.Replace(nameNoBlank, "([^0-9A-Z]+)", "");
                Regex regex = new Regex("([0-9]+)");
                MatchCollection match = regex.Matches(nameNoBlank);
                if (match.Count > 0)
                {
                    string subNum = "";
                    Dictionary<int, string> numDe = new Dictionary<int, string>();
                    string nameNoDigit = Regex.Replace(nameNoBlank, "([^A-Z]+)", "");
                    int totalLength = 0;
                    for (int i = match.Count - 1; i >= 0; i--)
                    {
                        totalLength += match[i].Length;
                        if ((match[i].Index + match[i].Length) > 7)
                        {
                            subNum = match[i].Value.Trim() + subNum;
                        }
                        else
                        {
                            numDe.Add(match[i].Index, match[i].Value);
                        }
                    }
                    nameNoDigit = nameNoDigit.Substring(0, 7 - totalLength);
                    foreach (var item in numDe)
                    {
                        nameNoDigit = nameNoDigit.Insert(item.Key, item.Value);
                    }
                    result = nameNoDigit + subNum;
                }
                else
                {
                    usName = ClearCoLtdForName(usName);
                    result = GetIDNDisplayNamePortion(usName, 0);
                }
            }
            int retry = 1;
            while (KoreaUnderlyingManager.ExsitDisplayName(result))
            {
                result = GetIDNDisplayNamePortion(companyName, retry++);
                //nameNoBlank = nameNoBlank.Substring(0, nameNoBlank.Length - 1);
            }
            return result;
        }

        /// <summary>
        /// Get IDNDisplay name(new underlying) (less than 7 characters) with underlying English campany name.
        /// </summary>
        /// <param name="companyName">underlying English campany name</param>
        /// <returns>IDN Display name</returns>
        private string GetIDNDisplayNamePortion(string companyName, int retry)
        {
            companyName = companyName.ToUpper();
            string cond = @"[A-Z0-9]+";
            Regex r = new Regex(cond);
            MatchCollection m = r.Matches(companyName);
            int n = m.Count;
            string idnName = "";
            if (n == 1)
            {
                return (m[0].Value.Substring(0, 7 - retry) + m[0].Value.Substring(m[0].Length - retry, retry));
            }

            idnName = m[0].Value.Length > 4 ? m[0].Value.Substring(0, 4) : m[0].Value;
            int subLength = (7 / n) == 0 ? 1 : (7 / n);
            for (int i = 1; i < n; i++)
            {
                Match item = m[i];
                if (item.Length >= (subLength + retry - 1))
                {
                    idnName += item.ToString().Substring(retry - 1, subLength);
                }
                else
                {
                    idnName += item.Value;
                }
            }
            idnName = idnName.Length > 7 ? idnName.Substring(0, 7) : idnName;
            return idnName;
        }

        /// <summary>
        /// Remove the infos of company like CO LTD CORP INC CORPARATION
        /// </summary>
        /// <param name="underEngName">full name</param>
        /// <returns>name without company infos</returns>
        public string ClearCoLtdForName(string underEngName)
        {
            underEngName = underEngName.ToUpper();
            underEngName = underEngName.Replace("CORPORATION", "");
            List<string> names = underEngName.Split(new char[] { ' ', ',', '.' }).ToList();
            string result = "";
            names.Remove("CO");
            names.Remove("LTD");
            names.Remove("INC");
            names.Remove("CORP");
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


        /// <summary>
        /// Generate three GEDA files for new underlying and two mails.
        /// e.g.
        ///	1.File name: UNDERLYING_CHAIN_UPLOAD_0#028150W.KQ.txt
        ///	2.File name: CHAIN_CONST_ADD_0#028150W.KQ.txt
        ///	3.File name: SUPERCHAIN_CONST_ADD_0#UNDLY.KQ.txt or KS
        /// </summary>
        private void GenerateNewUnderlyingFiles()
        {
            if (newUnderLying.Count == 0)
            {
                return;
            }
            //From config
            AddResult("New Underlying Folder", configObj.GEDA_NewUnderlying, "New Underlying Folder");

            string today = DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US"));
            string filePath = configObj.GEDA_NewUnderlying;
            string mailPath = Path.Combine(filePath, "MAIL" + today);
            FileUtil.CreateDirectory(filePath);
            FileUtil.CreateDirectory(mailPath);
            bool superKS = false;
            bool superKQ = false;
            ArrayList uploadContent = new ArrayList();
            ArrayList ksChain = new ArrayList();
            ArrayList kqChain = new ArrayList();
            ArrayList ksRic = new ArrayList();
            ArrayList kqRic = new ArrayList();
            string uploadFileName = "UNDERLYING_CHAIN_UPLOAD_" + today + ".txt";
            string superKSFileName = "SUPERCHAIN_CONST_ADD_0#UNDLY.KS.txt";
            string superKQFileName = "SUPERCHAIN_CONST_ADD_0#UNDLY.KQ.txt";
            string chainAddFileName = "CHAIN_CONST_ADD_" + today + ".txt";
            foreach (KoreaUnderlyingInfo newItem in newUnderLying)
            {
                string[] ricSpilt = newItem.UnderlyingRIC.Split('.');
                string modifyRic = newItem.UnderlyingRIC.Split('.')[0] + "W." + newItem.UnderlyingRIC.Split('.')[1];
                string fileName = Path.Combine(filePath, "UNDERLYING_CHAIN_UPLOAD_0#" + modifyRic + ".txt");

                string ricChainToFill = "KSE_STOCK_" + ricSpilt[0] + "W";
                string ksOrkqStrToFill = "STQS6\tKSE";
                string exchangeToFill = "KO";
                string mrnToFill = "287";
                string rdnExchidToFill = "156";
                string rdnExchd2ToFill = "156";
                string prodPermToFill = "3104";
                string superChain = "KSE_EQ_UNDLY";
                string ricToFill = ricSpilt[0];
                if (ricSpilt[1] == "KQ")
                {
                    ricChainToFill = "KOSDAQ_STOCK_" + ricSpilt[0] + "W";
                    ksOrkqStrToFill = "STQSR\tKOSDAQ";
                    exchangeToFill = "KQ";
                    mrnToFill = "144";
                    rdnExchidToFill = "0";
                    rdnExchd2ToFill = "380";
                    prodPermToFill = "4084";
                    superChain = "KOSDAQ_EQ_UNDLY";
                    superKQ = true;
                    kqChain.Add(ricChainToFill + "_CHAIN");
                    kqRic.Add(modifyRic);
                }
                else
                {
                    superKS = true;
                    ksChain.Add(ricChainToFill + "_CHAIN");
                    ksRic.Add(modifyRic);
                }

                string chainUploadData = "FILENAME\t" + ricChainToFill + "_CHAIN" + "\t" + ksOrkqStrToFill + "\r\n" +
                                       "CHAIN_RIC\t0#" + modifyRic + "\r\n" +
                                       "LINK_ROOT\t" + modifyRic + "\r\n" +
                                       "LOCAL_LANGUAGE\tLL_KOREAN\r\n" +
                                       "DISPLAY_NAME\t" + newItem.NDATCUnderlyingTitle + "\r\n" +
                                       "DISPLAY_NMLL\t" + newItem.CompanyName + "\r\n" +
                                       "RDNDISPLAY\t244\r\n" +
                                       "EXCHANGE\t" + exchangeToFill + "\r\n" +
                                       "ISSUE\tLINK\r\n" +
                                       "SEND\tIMSOUT\r\n" +
                                       "MRN\t" + mrnToFill + "\r\n" +
                                       "MRV\t0\r\n" +
                                       "INH_RANK\tTRUE\r\n" +
                                       "TPL_VER\t2.02\r\n" +
                                       "RDN_EXCHID\t" + rdnExchidToFill + "\r\n" +
                                       "RDN_EXCHD2\t" + rdnExchd2ToFill + "\r\n" +
                                       "TPL_NUM\t85\r\n" +
                                       "RECORDTYPE\t104\r\n" +
                                       "PROD_PERM\t" + prodPermToFill + "\r\n" +
                                       "CURRENCY\t410\r\n" +
                                       "RECKEY\tRK_SYMBOL\r\n" +
                                       "EMAIL_GROUP_ID\tgrpcntmarketdatastaff@thomsonreuters.com\r\nEND\r\n";

                string chainTitle = "RIC\tBCU\r\n";
                string chainConst = newItem.UnderlyingRIC + "\t" + ricChainToFill + "\r\n";
                string superChainData = "0#" + modifyRic + "\t" + superChain + "\r\n";

                File.AppendAllText(Path.Combine(filePath, uploadFileName), chainUploadData);

                fileName = Path.Combine(filePath, chainAddFileName);
                if (!File.Exists(fileName))
                {
                    chainConst = chainTitle + chainConst;
                }
                File.AppendAllText(fileName, chainConst);

                fileName = Path.Combine(filePath, "SUPERCHAIN_CONST_ADD_0#UNDLY." + ricSpilt[1] + ".txt");
                if (!File.Exists(fileName))
                {
                    superChainData = chainTitle + superChainData;
                }
                File.AppendAllText(fileName, superChainData);
            }
            AddResult(uploadFileName, Path.Combine(filePath, uploadFileName), @"GEDA File: UNDERLYING_CHAIN_UPLOAD");
            AddResult(chainAddFileName, Path.Combine(filePath, chainAddFileName), @"GEDA File: CHAIN_CONST_ADD");
            if (superKQ)
            {
                AddResult(superKQFileName, Path.Combine(filePath, superKQFileName), @"GEDA File: SUPERCHAIN_CONST_ADD");
            }
            if (superKS)
            {
                AddResult(superKSFileName, Path.Combine(filePath, superKSFileName), @"GEDA File: SUPERCHAIN_CONST_ADD");
            }
            SendNewUnderlyingMail(mailPath, ksChain, kqChain, ksRic, kqRic);
        }

        /// <summary>
        /// Generate two new underlying mails.
        /// 1. LXL update
        /// 2. Pls add Delay Chain  
        /// </summary>
        /// <param name="ksOrkqTE">TE mark for KQ or KS</param>
        /// <param name="chain">part of mail content</param>
        /// <param name="exchangeID">exchange ID for a new underlying</param>
        /// <param name="modifyRic">ric</param>
        private void SendNewUnderlyingMail(string filePath, ArrayList ksChain, ArrayList kqChain, ArrayList ksRic, ArrayList kqRic)
        {
            string mailBody = "<p>Hi BJG Central DBA,</p>";
            if (ksChain.Count > 0)
            {
                mailBody += "<p>Below chain has been created and TE success in STQS6"
                               + ".</p><p>Please update the following:</p><p>BCU to be added:</p>"
                               + "<table style=\"border-collapse:collapse;border:none; font-family: 'Arial';font-size: 12px;\"><tr><td style=\"border: solid #000 1px;\" >BCU</td><td style=\"border: solid #000 1px;\">Action Date</td></tr>";
                foreach (string chain in ksChain)
                {
                    mailBody += "<tr><td style=\"border: solid #000 1px;\">" + chain + "</td><td style=\"border: solid #000 1px;\">ASAP</td></tr>";
                }
                mailBody += "</table>";
            }
            if (kqChain.Count > 0)
            {
                mailBody += "<p>Below chain has been created and TE success in STQSR"
                               + ".</p><p>Please update the following:</p><p>BCU to be added:</p>"
                               + "<table style=\"border-collapse:collapse;border:none; font-family: 'Arial';font-size: 12px;\"><tr><td style=\"border: solid #000 1px;\" >BCU</td><td style=\"border: solid #000 1px;\">Action Date</td></tr>";
                foreach (string chain in kqChain)
                {
                    mailBody += "<tr><td style=\"border: solid #000 1px;\">" + chain + "</td><td style=\"border: solid #000 1px;\">ASAP</td></tr>";
                }
                mailBody += "</table>";
            }
            CreatMailAndSave("LXL update", mailBody, Path.Combine(filePath, "LXL update.msg"));

            mailBody = "<p>Hi BJG Central DBA,</p><p>Please help to build below delay chain under delay PE 5229.</p>";
            if (ksRic.Count > 0)
            {
                mailBody += "<p>The Exchange ID is 156.</p>";
                foreach (string ric in ksRic)
                {
                    mailBody += "<p>0#" + ric + "</p>";
                }
            }
            if (kqRic.Count > 0)
            {
                mailBody += "<p>The Exchange ID is 380.</p>";
                foreach (string ric in kqRic)
                {
                    mailBody += "<p>0#" + ric + "</p>";
                }
            }
            CreatMailAndSave("Pls add Delay Chain", mailBody, Path.Combine(filePath, "Pls add Delay Chain.msg"));

        }

        /// <summary>
        /// Create mail and save it to local disk. Users can check the mail content.
        /// </summary>
        /// <param name="mailSubject">mail subject</param>
        /// <param name="mailBody">mail body</param>
        /// <param name="filePath">path to save</param>
        private void CreatMailAndSave(string mailSubject, string mailBody, string filePath)
        {
            MailToSend mail = new MailToSend();
            mail.ToReceiverList.AddRange(configObj.NewUnderlying_MailTo);
            mail.CCReceiverList.AddRange(configObj.NewUnderlying_MailCC);
            mail.MailSubject = mailSubject;
            string signature = string.Join("<br>", configObj.NewUnderlying_Signature.ToArray());
            //mail.MailBody += signature;

            mail.MailHtmlBody = "<div style=\"font-family: 'Arial';font-size: 10pt;\">" + mailBody;
            mail.MailHtmlBody += "<br>" + signature + "</div>";
            string err = string.Empty;
            using (OutlookApp outlookApp = new OutlookApp())
            {
                OutlookUtil.SaveMail(outlookApp, mail, out err, filePath);
            }
            AddResult(mail.MailSubject + ".msg", filePath, mail.MailSubject);
        }

        #endregion


        #region BulkFileGenerator
        ////BDN ADD File
        //private void _PrintBDN_ADD_File()
        //{
        //    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    if (xlApp == null)
        //    {
        //        commonObj.WriteLogFile("Excel could not be started. Check that your office installation and project reference are correct !!!");
        //        return;
        //    }
        //    xlApp.Visible = false;
        //    try
        //    {
        //        Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        //        Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
        //        wSheet.Name = "BDN_ADD";

        //        ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 14;
        //        ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 35;
        //        ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 13;
        //        ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 15;
        //        ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 8;

        //        ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 7;
        //        ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 10;
        //        ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 10;
        //        ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 13;
        //        ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 14;
        //        ((Range)wSheet.Columns["A:B", System.Type.Missing]).Font.Name = "Courier";
        //        ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
        //        ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
        //        if (wSheet == null)
        //        {
        //            commonObj.WriteLogFile("Error found in Print BDN_ADD file ");
        //            return;
        //        }
        //        wSheet.Cells[1, 1] = "Symbol";
        //        wSheet.Cells[1, 2] = "Company Name";
        //        wSheet.Cells[1, 3] = "Listing Date";
        //        wSheet.Cells[1, 4] = "ISIN";
        //        wSheet.Cells[1, 5] = "1 or 2";
        //        wSheet.Cells[1, 6] = "Type";
        //        wSheet.Cells[1, 7] = "Exch Code";
        //        wSheet.Cells[1, 8] = "Und. Asset";
        //        wSheet.Cells[1, 9] = "Mat Date";
        //        wSheet.Cells[1, 10] = "Strike price";

        //        int startLine = 2;
        //        startLine = LoopPrintBDNADDFile(wSheet, startLine, koreaList);

        //        xlApp.DisplayAlerts = false;
        //        xlApp.AlertBeforeOverwriting = false;
        //        string fullPath = commonObj.Log_Path + "\\" + commonObj.SubFolder + "\\" + "BDN_ADD_" + DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US")).Replace("-", "");
        //        fullPath += ".xls";

        //        wBook.SaveCopyAs(fullPath);
        //    }
        //    catch (Exception ex)
        //    {
        //        commonObj.WriteLogFile("Error found in BDN ADD file :" + ex.ToString());
        //    }
        //    finally
        //    {
        //        xlApp.Quit();
        //        commonObj.KillExcelProcess(xlApp);
        //    }
        //}

        //private int LoopPrintBDNADDFile(Worksheet wSheet, int startLine, List<WarrantTemplate> lists)
        //{
        //    if (lists.Count > 0)
        //    {
        //        for (var i = 0; i < lists.Count; i++)
        //        {
        //            ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
        //            wSheet.Cells[startLine, 1] = "kr;" + lists[i].ISIN.Substring(3, 2) + lists[i].ISIN.Substring(7, 4) + "'";
        //            ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";

        //            string companyname = lists[i].Issuer;
        //            string[] company = companyname.Split(' ');
        //            string _new_company = "";
        //            for (var j = 0; j < company.Length; j++)
        //            {
        //                if (j < (company.Length - 2))
        //                {
        //                    string x = company[j].ToLower();
        //                    string y = x.Substring(0, 1).ToUpper();
        //                    string z = x.Substring(1);

        //                    x = y + z;
        //                    _new_company += (x + " ");
        //                }
        //                else
        //                    _new_company += (company[j] + " ");
        //            }
        //            _new_company = _new_company.Trim().ToString();
        //            wSheet.Cells[startLine, 2] = _new_company;
        //            wSheet.Cells[startLine, 3] = "";
        //            wSheet.Cells[startLine, 4] = lists[i].ISIN;
        //            wSheet.Cells[startLine, 5] = lists[i].CallOrPut == "CALL" ? "1" : "2";
        //            wSheet.Cells[startLine, 6] = lists[i].CallOrPut == "CALL" ? "called" : "put";
        //            wSheet.Cells[startLine, 7] = "J" + lists[i].Ticker;

        //            string _underlying_koreaname = lists[i].Underlying_Korea_Name;
        //            if (koreaUnderlyingHash.Contains(_underlying_koreaname))
        //                wSheet.Cells[startLine, 8] = ((_UnderlyingCode)koreaUnderlyingHash[_underlying_koreaname]).BDN.ToString();
        //            else
        //                wSheet.Cells[startLine, 8] = "出现异常信息... ...";

        //            ((Range)wSheet.Cells[startLine, 9]).NumberFormat = "@";
        //            wSheet.Cells[startLine, 9] = lists[i].Mat_Date;
        //            wSheet.Cells[startLine, 10] = lists[i].Strike_Price;
        //            startLine++;
        //        }
        //    }
        //    return startLine;
        //}
        #endregion


    }
}
