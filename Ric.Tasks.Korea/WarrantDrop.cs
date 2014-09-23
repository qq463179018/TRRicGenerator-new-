using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Globalization;
using Ric.Db.Manager;
using System.Data;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Collections;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
{
    public class WarrantDrop : GeneratorBase
    {
        private List<DateTime> holidayList = null;
        private static readonly string HOLIDAY_LIST_FILE_PATH = ".\\Config\\Korea\\Holiday.xml";
        private List<CompanyWarrantDropTemplate> waDrop = new List<CompanyWarrantDropTemplate>();
        private KoreaCompanyWarrantDropGeneratorConfig configObj = null;
        private const string ETI_KOREA_COMPANYWARRANT_TABLE_NAME = "ETI_Korea_CompanyWarrant";


        protected override void Initialize()
        {
            base.Initialize();
            holidayList = ConfigUtil.ReadConfig(HOLIDAY_LIST_FILE_PATH, typeof(List<DateTime>)) as List<DateTime>;
            configObj = Config as KoreaCompanyWarrantDropGeneratorConfig;
        }

        protected override void Start()
        {
            GrabDataFromKindWebpage();
            if (waDrop.Count > 0)
            {
                FormatData();
                GenerateFiles();
                //InsertGedaInfoToDb();
            }
            else
            {
                Logger.Log("There is no Company Warrant DROP data grabbed");
            }
            AddResult("LOG file",Logger.FilePath,"LOG file");  
        }

        /// <summary>
        /// Grab drop anouncement from kind webpage.
        /// </summary>
        private void GrabDataFromKindWebpage()
        {
            string startDate = configObj.StartDate;
            string endDate = configObj.EndDate;
            if (string.IsNullOrEmpty(startDate))
            {
                startDate = DateTime.Today.ToString("yyyy-MM-dd");
            }
            if (string.IsNullOrEmpty(endDate))
            {
                endDate = DateTime.Today.ToString("yyyy-MM-dd");
            }
            string dataStartDate = (DateTime.Parse(endDate).AddMonths(-2)).ToString("yyyy-MM-dd");
            string postData = string.Format("method=searchTotalInfoSub&forward=searchtotalinfo_detail&searchCodeType=&searchCorpName=%EC%83%81%EC%9E%A5%ED%8F%90%EC%A7%80&repIsuSrtCd=&isurCd=&fdName=all_mktact_idx&pageIndex=1&currentPageSize=300&scn=mktact&srchFd=2&kwd=%EC%83%81%EC%9E%A5%ED%8F%90%EC%A7%80&fromData={0}&toData={1}", dataStartDate, endDate);
            string searchUri = "http://kind.krx.co.kr/disclosure/searchtotalinfo.do";
            try
            {
                HtmlDocument htc = new HtmlDocument();
                var pageSource = WebClientUtil.GetDynamicPageSource(searchUri, 300000, postData);
                if (!string.IsNullOrEmpty(pageSource))
                {
                    htc.LoadHtml(pageSource);
                }
                HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes("//dt");
                HtmlNodeCollection ddCollections = htc.DocumentNode.SelectNodes("//dl/dd");

                if (nodeCollections.Count > 0)
                {
                    for (int i = nodeCollections.Count - 1; i >= 0; i--)
                    {
                        HtmlNode ddNode = ddCollections[i].SelectSingleNode(".//span");

                        HtmlNode dtNode = nodeCollections[i];
                        HtmlNode node = nodeCollections[i].SelectSingleNode(".//span/a");
                        string titleNode = string.Empty;
                        if (node != null)
                        {
                            titleNode = node.InnerText.ToString();
                        }
                        titleNode = titleNode.Contains("[") ? titleNode.Replace("[", "").Replace("]", "").Trim().ToString() : titleNode;

                        if (titleNode.Contains("신주인수권증권 상장폐지"))
                        {
                            HtmlNode nodeDate = dtNode.SelectSingleNode("./em");
                            if (nodeDate != null)
                            {                               
                                DateTime anouncementDate = new DateTime();
                                anouncementDate = DateTime.Parse(nodeDate.InnerText.Trim(), new CultureInfo("en-US"));
                                if (anouncementDate < DateTime.Parse(startDate))
                                {
                                    continue;
                                }
                            }
                            string title = titleNode.Substring(0, "신주인수권증권 상장폐지".Length).Trim().ToString();
                            if (title.Equals("신주인수권증권 상장폐지"))
                            {
                                CompanyWarrantDropTemplate item = new CompanyWarrantDropTemplate();
                                HtmlNode header = nodeCollections[i].SelectSingleNode(".//strong/a");
                                string attribute = string.Empty;
                                if (header != null)
                                {
                                    attribute = header.Attributes["onclick"].Value.ToString().Trim();
                                }
                                if (!string.IsNullOrEmpty(attribute))
                                {
                                    attribute = attribute.Split('(')[1].Split(')')[0].Trim(new Char[] { ' ', '\'', ';' }).ToString();
                                }

                                string judge = ddNode.InnerText.Split(':')[1].Trim();

                                //string attrituteUri = string.Format("http://kind.krx.co.kr/common/companysummary.do?method=searchCompanySummary&strIsurCd={0}&lstCd=undefined", attribute);
                                //string judge = string.Empty;
                                //HtmlDocument doc = WebClientUtil.GetHtmlDocument(attrituteUri, 120000, null);
                                //if (doc != null)
                                //{
                                //    HtmlNode docnode = doc.DocumentNode.SelectSingleNode("//div[@id='pContents']/table/tbody/tr[2]/td[2]");
                                //    if (docnode != null)
                                //    {
                                //        judge = docnode.InnerText;
                                //    }
                                //}


                                string parameters = node.Attributes["onclick"].Value.ToString().Trim();
                                parameters = parameters.Split('(')[1].Split(')')[0].ToString().Trim(new char[] { ' ', '\'', ';' });
                                string param1 = parameters.Split(',')[0].Trim(new char[] { ' ', '\'', ',' }).ToString();
                                string param2 = parameters.Split(',')[1].Trim(new char[] { ' ', '\'', ',' }).ToString();
                                string uri = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno={1}&viewerhost=&viewerport=", param1, param2);

                                HtmlDocument doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                                string ticker = string.Empty;
                                if (doc != null)
                                    ticker = doc.DocumentNode.SelectSingleNode("//header/h1").InnerText;

                                if (!string.IsNullOrEmpty(ticker))
                                {
                                    Match m = Regex.Match(ticker, @"\(([0-9a-zA-Z]{6})\)");
                                    if (m == null)
                                    {
                                        string msg = "Cannot get ticker numbers in ." + ticker;
                                        Logger.Log(msg, Logger.LogType.Error);
                                        continue;
                                    }
                                    ticker = m.Groups[1].Value;
                                }

                                //if (!string.IsNullOrEmpty(ticker))
                                //    ticker = ticker.Split('(')[1].Trim(new char[] { ' ', ')', '(' }).ToString();

                                string param3 = judge.Contains("유가증권") ? "68913" : (judge.Contains("코스닥") ? "70925" : null);
                                if (string.IsNullOrEmpty(param3))
                                    return;
                                param1 = param1.Insert(4, "/").Insert(7, "/").Insert(10, "/").ToString();
                                uri = string.Format("http://kind.krx.co.kr/external/{0}/{1}/{2}.htm", param1, param2, param3);

                                doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                                if (doc == null)
                                    return;
                                // For KQ
                                if (judge.Contains("코스닥"))
                                {
                                    HtmlNode koreaName = doc.DocumentNode.SelectSingleNode("//tr[1]/td[2]");
                                    HtmlNode effective = doc.DocumentNode.SelectSingleNode("//tr[5]/td[2]");
                                    string kname = string.Empty;
                                    string edate = string.Empty;
                                    if (koreaName != null)
                                        kname = koreaName.InnerText.Trim().ToString();
                                    if (effective != null)
                                        edate = effective.InnerText.Trim().ToString();
                                    kname = kname.Trim().ToString();
                                    if (!string.IsNullOrEmpty(kname))
                                        kname = kname.Contains("(주)") ? kname.Replace("(주)", "").Trim() : kname;
                                    edate = edate.Trim().ToString();
                                    if (!string.IsNullOrEmpty(edate))
                                        edate = Convert.ToDateTime(edate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                                    item.KoreanName = kname;
                                    item.EffectiveDate = edate;
                                    item.RIC = ticker + ".KQ";
                                    waDrop.Add(item);
                                }
                                //For KS
                                else if (judge.Contains("유가증권"))
                                {
                                    string skName = MiscUtil.GetCleanTextFromHtml(doc.DocumentNode.SelectSingleNode("//tr[1]/td[2]").InnerText);
                                    skName = Regex.Split(skName, "   ", RegexOptions.IgnoreCase)[0].Trim(new char[] { ' ', ':' }).ToString();
                                    string seDate = MiscUtil.GetCleanTextFromHtml(doc.DocumentNode.SelectSingleNode("//tr[8]/td[2]").InnerText);
                                    seDate = Convert.ToDateTime(seDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                                    item.KoreanName = skName;
                                    item.EffectiveDate = seDate;
                                    item.RIC = ticker + ".KS";
                                    waDrop.Add(item);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in SearchTheWebpageToGrabData()      : \r\n" + ex.InnerException + "  :  \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        /// <summary>
        /// Format drop data. Get infos from database.
        /// </summary>
        private void FormatData()
        {
            foreach (CompanyWarrantDropTemplate item in waDrop)
            {
                try
                {
                    string condition = "where REPLACE(KoreanName, ' ', '')  = N'" + item.KoreanName.Replace(" ", "") + "' and Status = 'Active'";
                    System.Data.DataTable dt = ManagerBase.Select(ETI_KOREA_COMPANYWARRANT_TABLE_NAME , new string[] { "*" }, condition);
                    if (dt == null)
                    {
                        Logger.Log("Error found in selecting data from DB. Check if there is error in DB", Logger.LogType.Error);
                        continue;
                    }
                    if (dt.Rows.Count == 0)
                    {
                        Logger.Log("The database does not contain record of RIC:" + item.RIC + "\tKoreanName:" + item.KoreanName + "! Please check the DB!", Logger.LogType.Error);
                        continue;
                    }
                    if (dt.Rows.Count > 1)
                    {
                        string warnMsg = "Found mutilple records for RIC: " + item.RIC + "\tKoreanName:" + item.KoreanName;
                        Logger.Log(warnMsg, Logger.LogType.Warning);
                    }
                    DataRow row = dt.Rows[0];
                    item.RIC = Convert.ToString(row["RIC"]).Trim();
                    item.ISIN = Convert.ToString(row["ISIN"]).Trim();
                    string qaCommonNameDb = Convert.ToString(row["QACommonName"]).Trim();
                    item.QAShortName = Convert.ToString(row["QAShortName"]).Trim();
                    string expiryDateDb = Convert.ToDateTime(row["ExpiryDate"]).ToString("MMMyy").ToUpper();
                    string commonNameTemp = Convert.ToString(row["ForIACommonName"]).Trim();
                    string exercisePrice = Convert.ToString(row["ExercisePrice"]).Trim();
                    string expiryDate = DateTime.Parse(item.EffectiveDate).ToString("ddMMMyy", new CultureInfo("en-US"));
                    string expiryDateSp = expiryDate.Substring(2).ToUpper();
                    item.ForIACommonName = commonNameTemp + " Call " + exercisePrice + " KRW " + commonNameTemp + " " + expiryDate;
                    item.QACommonName = qaCommonNameDb.Replace(expiryDateDb, expiryDateSp);

                    row["UpdateDateDrop"] = DateTime.Today.ToString("yyyy-MM-dd");
                    if (!string.IsNullOrEmpty(item.EffectiveDate))
                    {
                        row["EffectiveDateDrop"] = item.EffectiveDate;
                        row["ExpiryDate"] = item.EffectiveDate;
                    }
                    row["Status"] = "De-active";
                    row["QACommonName"] = item.QACommonName;
                    ManagerBase.UpdateDbTable(dt, ETI_KOREA_COMPANYWARRANT_TABLE_NAME);
                    Logger.Log("1 record status changed to 'De-active'. RIC:" + item.RIC);
                }
                catch (Exception ex)
                {
                    Logger.Log("Error found in formating DROP data from database. \r\n" + ex.ToString(), Logger.LogType.Error);
                }

            }
        }

        private void GenerateFiles()
        {
            GenerateFMFiles();
            GenerateGEDAFile();
            GenerateNDAQAFile();
            GenerateNDAIAFile();
        }

        /// <summary>
        /// Generate FM file for Drop.
        /// </summary>
        private void GenerateFMFiles()
        {
            foreach (CompanyWarrantDropTemplate item in waDrop)
            {
                ExcelApp excelApp = new ExcelApp(false, false);
                if (excelApp.ExcelAppInstance == null)
                {
                    string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }
                string fileName = "KR FM (Company Warrant Drop) Request_" + item.RIC + " (wef " + item.EffectiveDate + ").xls";
                string filePath = Path.Combine(configObj.FM, fileName);
                try
                {
                    Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                    if (wSheet == null)
                    {
                        string msg = "Worksheet could not be started. Check that your office installation and project reference are correct!";
                        Logger.Log(msg, Logger.LogType.Error);
                        return;
                    }
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
                    ((Range)wSheet.Cells[3, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
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
                    excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                    wBook.Save();

                    MailToSend mail = new MailToSend();
                    mail.ToReceiverList.AddRange(configObj.MailTo);
                    mail.CCReceiverList.AddRange(configObj.MailCC);
                    mail.MailSubject = Path.GetFileNameWithoutExtension(fileName);
                    mail.AttachFileList.Add(filePath);
                    mail.MailBody = "Company Warrant Drop:\t\t" + item.RIC + "\r\n\r\n"
                                    + "Effective Date:\t\t" + item.EffectiveDate + "\r\n\r\n\r\n\r\n";
                    string signature = string.Join("\r\n", configObj.MailSignature.ToArray());
                    mail.MailBody += signature;

                    AddResult(fileName,filePath,"FM File");
                    Logger.Log("Generate FM file successfully. Filepath is " + filePath);
                }
                catch (Exception ex)
                {
                    String msg = "Error found in generating FM file for RIC:" + item.RIC + " \r\n" + ex.Message;
                    Logger.Log(msg, Logger.LogType.Error);
                }
                finally
                {
                    excelApp.Dispose();
                }
            }
        }

        /// <summary>
        /// Generate Drop Geda files. 
        /// Write the records to the file named with effective date.
        /// </summary>
        private void GenerateGEDAFile()
        {
            Hashtable result = new Hashtable();
            List<string> gedaTitle = new List<string>() { "RIC" };
            foreach (CompanyWarrantDropTemplate item in waDrop)
            {
                try
                {
                    string fileName = "KR_DROP_" + DateTime.Parse(item.EffectiveDate).ToString("yyyyMMdd", new System.Globalization.CultureInfo("en-US")) + ".txt";
                    string filePath = Path.Combine(configObj.GEDA, fileName);
                    List<List<string>> data = new List<List<string>>();
                    List<string> record = new List<string>();
                    record.Add(item.RIC);
                    data.Add(record);
                    FileUtil.WriteOutputFile(filePath, data, gedaTitle, WriteMode.Append);
                    if (!result.Contains(fileName))
                    {
                        result.Add(fileName, filePath);
                    }
                    Logger.Log("Write to GEDA file successfully. Filepath is " + filePath);
                }
                catch (Exception ex)
                {
                    Logger.Log("Error found in generating GEDA file for RIC:" + item.RIC + "\r\n " + ex.Message, Logger.LogType.Error);
                }
            }
            SetTaskResultList(result);
        }

        /// <summary>
        /// Set the GEDA files result list.
        /// </summary>
        /// <param name="result">results</param>
        private void SetTaskResultList(Hashtable result)
        {
            ArrayList keysArr = new ArrayList(result.Keys);
            keysArr.Sort();
            foreach (string keyRusult in keysArr)
            {
                AddResult(keyRusult,result[keyRusult].ToString(),"GEDA DROP FILE");
            }
        }

        /// <summary>
        /// Generate NDA QA file.
        /// </summary>
        private void GenerateNDAQAFile()
        {
            try
            {
                string fileName = "KR" + DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US")) + "QADrop.csv";
                string filePath = Path.Combine(configObj.NDA, fileName);
                List<string> ndaQaTitle = new List<string>() { "RIC", "ASSET COMMON NAME", "EXPIRY DATE" };
                List<List<string>> data = new List<List<string>>();
                foreach (CompanyWarrantDropTemplate item in waDrop)
                {
                    string expiryDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
                    List<string> record = new List<string>();
                    record.Add(item.RIC);
                    record.Add(item.QACommonName);
                    record.Add(expiryDate);
                    data.Add(record);
                    List<string> recordF= new List<string>();
                    string ricF = item.RIC.Split('.')[0] + "F." + item.RIC.Split('.')[1];
                    recordF.Add(ricF);
                    recordF.Add(item.QACommonName);
                    recordF.Add(expiryDate);
                    data.Add(recordF);
                }
                FileUtil.WriteOutputFile(filePath, data, ndaQaTitle, WriteMode.Overwrite);
                AddResult(fileName,filePath,"NDA QA FILE");
                Logger.Log("Generate NDA QA file successfully. Filepath is " + filePath);
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in generating NDA QA file.\r\n " + ex.Message, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Generate NDA IA file.
        /// </summary>
        private void GenerateNDAIAFile()
        {
            try
            {
                string fileName = "KR" + DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US")) + "IADrop.csv";
                string filePath = Path.Combine(configObj.NDA, fileName);
                List<string> ndaQaTitle = new List<string>() { "ISIN", "ASSET COMMON NAME" };
                List<List<string>> data = new List<List<string>>();
                foreach (CompanyWarrantDropTemplate item in waDrop)
                {
                    List<string> record = new List<string>();
                    record.Add(item.ISIN);
                    record.Add(item.ForIACommonName);
                    data.Add(record);
                }
                FileUtil.WriteOutputFile(filePath, data, ndaQaTitle, WriteMode.Overwrite);
                AddResult(fileName,filePath,"NDA IA FILE");
                Logger.Log("Generate NDA IA file successfully. Filepath is " + filePath);
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in generating NDA IA file.\r\n " + ex.Message, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Insert drop records into DB. Contains EffectiveDate, RIC and TaskId.
        /// </summary>
        private void InsertGedaInfoToDb()
        {
            foreach (CompanyWarrantDropTemplate item in waDrop)
            {
                try
                {
                    DropGedaManager.UpdateDrop(item.EffectiveDate, item.RIC, TaskId);
                }
                catch(Exception ex)
                {
                    Logger.Log("Error found in insert drop record to DB. For RIC:" + item.RIC + "\r\n" + ex.Message, Logger.LogType.Error);
                }
            }
            Logger.Log("Update DROP_GEDA table successfully.");
        }
    }
}
