using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;
using System.Threading;

namespace Ric.Tasks.China
{
    #region Configuration

    [ConfigStoredInDB]
    public class ChinaIndexConstituentUpdateConfig
    {
        [StoreInDB]
        [Description("The path where the result will be written.\nEg: C:/Mydrive/")]
        public string ResultFolderPath { get; set; }

        [StoreInDB]
        [Description("The name of the result file ")]
        public string ResultFileName { get; set; }

        [StoreInDB]
        [Description("Excel path")]
        public string RicListExcelPath { get; set; }
    }

    #endregion

    #region Task

    class ChinaIndexConstituentUpdate : GeneratorBase
    {
        #region Declaration

        private static ChinaIndexConstituentUpdateConfig configObj = null;
        private ExcelApp app;
        private const string cnIndexUrl = "http://www.cnindex.com.cn/";
        private const string csIndexUrl = "http://www.csindex.com.cn";
        private List<string> pagesToCheck = new List<string> { "szxl", "jcxl", "cctv50" };
        private Dictionary<string, string> ricOfficialCode;
        private Dictionary<string, string> ricExl;
        private Dictionary<string, List<string>> newValues;
        private Dictionary<string, List<string>> gatsValues;
        private List<List<string>> toUpdate;

        #endregion

        #region Interface implementation

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as ChinaIndexConstituentUpdateConfig;
            app = new ExcelApp(false, false);
            if (app.ExcelAppInstance == null)
            {
                Logger.Log("Excel cannot be started", Logger.LogType.Error);
            }
            toUpdate = new List<List<string>>();
        }

        protected override void Start()
        {
            List<string> pages = new List<string>();
            List<string> excelLinks = new List<string>();

            try
            {
                ricOfficialCode = GetRicList();
                ricExl = GetExlList();


                //foreach (string pageToCheck in pagesToCheck)
                //{
                //    FindPagesToSearch(ref pages, pageToCheck);
                //}
                FindPagesToSearch2(ref pages);
                int i = 0;
                foreach (string page in pages)
                {
                    string tmpLink;
                    if ((tmpLink = FindExcelLink(page)) != null)
                    {
                        excelLinks.Add(tmpLink);
                    }
                    i++;
                }
                DownLoadExcels(excelLinks);
                newValues = GetNewValues();
                Thread.Sleep(5000);
                LogMessage("start querying in gats...");
                gatsValues = CallGats();
                Thread.Sleep(5000);
                CompareGatsWithNew();
                Thread.Sleep(5000);
                LogMessage("start sorting results...");
                ReOrderResults();
                Thread.Sleep(5000);
                LogMessage("start generating csv file ...");
                WriteResultsInCsv();
                AddResult("bulk file", configObj.ResultFileName, "result file");
                AddResult("bulk file folder", configObj.ResultFolderPath, "result folder");
            }
            catch
            {

            }
            finally
            {
                app.Dispose();
            }
        }

        #endregion

        #region Find Excels from CnIndex

        /// <summary>
        /// Parse the market page to find all listed indices pages
        /// </summary>
        /// <param name="pages">List with all the results</param>
        /// <param name="folder">The page folder name</param>
        private void FindPagesToSearch(ref List<string> pages, string folder)
        {
            HtmlDocument htc = new HtmlDocument();
            string uri = String.Format("{0}zstx/{1}/", cnIndexUrl, folder);
            try
            {
                htc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                int firstTable = 0;
                foreach (HtmlNode table in tables)
                {
                    if (firstTable == 0)
                    {
                        firstTable = 1;
                        continue;
                    }
                    int firstTr = 0;
                    HtmlNodeCollection trs = table.SelectNodes(".//tr");
                    foreach (HtmlNode tr in trs)
                    {
                        if (firstTr == 0)
                        {
                            firstTr = 1;
                            continue;
                        }
                        pages.Add(String.Format("{0}zstx/{1}{2}", cnIndexUrl, folder, tr.SelectSingleNode(".//td[1]/a").Attributes["href"].Value.Trim().ToString().Replace("./", "/")));
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Read the page and find the corresponding excel link
        /// </summary>
        /// <param name="page">Link of the page</param>
        /// <returns>The excel link</returns>
        private string FindExcelLink(string page)
        {
            HtmlDocument htc = new HtmlDocument();
            try
            {
                Thread.Sleep(1000);
                htc = WebClientUtil.GetHtmlDocument(page, 300000, null);
                if (page.Contains("cnindex"))
                {
                    HtmlNodeCollection uls = htc.DocumentNode.SelectNodes(".//ul");
                    HtmlNode excelLinkNode = uls[6];
                    if (excelLinkNode == null)
                    {
                        excelLinkNode = uls[5];
                    }
                    return (cnIndexUrl + excelLinkNode.SelectSingleNode(".//li[3]/a").Attributes["href"].Value.Trim().ToString().Replace("../", ""));
                }
                else
                {
                    HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                    foreach (HtmlNode tr in tables[1].SelectNodes(".//tr"))
                    {
                        string excelUrl = tr.SelectSingleNode(".//td[2]/a").Attributes["href"].Value.Trim().ToString();
                        if (excelUrl.EndsWith(".xls"))
                        {
                            if (excelUrl.StartsWith("http"))
                            {
                                return (excelUrl);
                            }
                            else
                            {
                                return (csIndexUrl + excelUrl);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return null;
        }

        #endregion

        #region Find Excels from CsIndex

        private void FindPagesToSearch2(ref List<string> pages)
        {
            HtmlDocument htc = new HtmlDocument();
            List<string> pagesToCheckCs = new List<string> { "1", "2", "4" };

            foreach (string pageNb in pagesToCheckCs)
            {
                string uri = String.Format("http://www.csindex.com.cn/sseportal/csiportal/zs/indexreport.do?type={0}", pageNb);
                try
                {
                    //htc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                    htc = GetDocRetryMechanism(uri, null, 5, 2000);
                    HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                    foreach (HtmlNode table in tables)
                    {
                        int firstTr = 0;
                        HtmlNodeCollection trs = table.SelectNodes(".//tr");
                        foreach (HtmlNode tr in trs)
                        {
                            if (firstTr == 0)
                            {
                                firstTr = 1;
                                continue;
                            }
                            pages.Add(csIndexUrl + tr.SelectSingleNode(".//td[1]/a").Attributes["href"].Value.Trim().ToString().Replace("./", "/"));
                        }
                    }
                }
                catch (Exception ex)
                {
                    string msg = "Error :" + ex.ToString();
                    Logger.Log(msg, Logger.LogType.Error);
                }
            }
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
                        doc = WebClientUtil.GetHtmlDocument(url, 300000, postData);

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
                            throw;
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
        #endregion

        #region Download Excels
        private void DownLoadExcels(List<string> excelLinks)
        {
            string savePath = string.Empty;
            LogMessage(string.Format("start to download .xls. count:{0}", excelLinks.Count));

            try
            {
                foreach (var item in excelLinks)
                {
                    savePath = Path.Combine(configObj.ResultFolderPath, Path.GetFileName(item));
                    RetryMechanism(item, savePath, 5, 2000);
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

            LogMessage("download all xls files finish!");
        }

        private void RetryMechanism(string downloadPath, string savePath, int retryTimes, int waitSecond)
        {
            for (int i = 0; i < retryTimes; i++)
            {
                try
                {
                    if (File.Exists(savePath))
                        break;

                    WebClientUtil.DownloadFile(downloadPath, 5000, savePath);

                    if (File.Exists(savePath))
                        break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(waitSecond);

                    if (i == retryTimes - 1)
                    {
                        string msg = string.Format("url:{0}     retryTimes:{1}      msg:{2}", downloadPath, retryTimes.ToString(), ex.Message);
                        Logger.Log(msg, Logger.LogType.Error);
                    }
                }
            }
        }
        #endregion

        #region Get values from Excels

        private Dictionary<string, string> GetRicList()
        {
            Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.RicListExcelPath);

            Dictionary<string, string> rics = new Dictionary<string, string>();
            for (int position = 1; position <= 3; position++)
            {
                List<List<string>> values = workbook.ToList(position);
                foreach (List<string> value in values)
                {
                    if (!rics.ContainsKey(value[1].Trim()))
                    {
                        rics.Add(value[1].Trim(), value[0]);
                    }
                }
            }
            return rics;
        }

        private Dictionary<string, string> GetExlList()
        {
            Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.RicListExcelPath);

            Dictionary<string, string> exls = new Dictionary<string, string>();
            for (int position = 1; position <= 3; position++)
            {
                int count = 0;
                List<List<string>> values = workbook.ToList(position);
                foreach (List<string> value in values)
                {
                    if (count == 0)
                    {
                        count++;
                        continue;
                    }
                    if (!exls.ContainsKey(value[0].Trim()))
                    {
                        exls.Add(value[0].Trim(), value[2]);
                    }
                }
            }
            return exls;
        }

        private List<string> GetDataFromExcel(Workbook workbook)
        {
            List<List<string>> values = workbook.ToList();
            List<string> data = new List<string>();

            if (values[0][0] == "指数代码")
            {
                foreach (List<string> value in values)
                {
                    if (!value[2].Trim().StartsWith("1"))
                    {
                        data.Add(value[2].Trim());
                    }
                }
                data[0] = values[1][0].Trim();
            }
            else
            {
                foreach (List<string> value in values)
                {
                    if (!value[2].Trim().StartsWith("1"))
                    {
                        data.Add(value[0].Trim());
                    }
                }
                data[0] = workbook.Name.Substring(0, 6);
            }
            return data;
        }

        private Dictionary<string, List<string>> GetNewValues()
        {
            Dictionary<string, List<string>> newValues = new Dictionary<string, List<string>>();
            string[] fileEntries = Directory.GetFiles(configObj.ResultFolderPath);
            foreach (string filename in fileEntries)
            {
                if (filename == configObj.RicListExcelPath)
                {
                    continue;
                }
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, filename);
                List<string> entry = GetDataFromExcel(workbook);
                if (ricOfficialCode.ContainsKey(entry[0]) && entry != null)
                {
                    string key = ricOfficialCode[entry[0]];
                    entry.RemoveAt(0);
                    newValues.Add(key, entry);
                }
            }
            return newValues;
        }

        #endregion

        #region Get values from gats

        private string CreateRicListForGats()
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                foreach (KeyValuePair<string, List<string>> value in newValues)
                {
                    string tmpRic = value.Key.Substring(1);
                    int pagesToSeek = (value.Value.Count / 14) + 1;
                    for (int i = 0; i < pagesToSeek; i++)
                    {
                        sb.Append(string.Format("{0}{1},", i.ToString(), tmpRic));
                    }
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Cannot generate Ric list for Gats: " + ex.Message);
            }
        }

        private Dictionary<string, List<string>> CallGats()
        {
            GatsUtil gats = new GatsUtil(GatsUtil.Server.Elektron);
            Dictionary<string, List<string>> gatsValues = new Dictionary<string, List<string>>();
            Regex rgxSpace = new Regex(@"\s+");
            Regex rgxLongLink = new Regex(@"^LONGLINK", RegexOptions.IgnoreCase);
            try
            {
                string rics = CreateRicListForGats();

                string[] stringSeparators = new string[] { "\r\n" };
                char[] stringSeparators2 = new char[] { ' ' };


                for (int i = 0; i < newValues.Count; i += 3000)
                {
                    StringBuilder sb = new StringBuilder();
                    int tmpi = 0;
                    foreach (KeyValuePair<string, List<string>> value in newValues)
                    {
                        string tmpRic = value.Key.Substring(1);
                        int pagesToSeek = (value.Value.Count / 14) + 1;
                        for (int j = 0; j < pagesToSeek; j++)
                        {
                            if (tmpi >= i && tmpi <= (i + 3000))
                            {
                                sb.Append(string.Format("{0}{1},", j.ToString(), tmpRic));
                            }
                        }
                        tmpi++;
                    }

                    string test = gats.GetGatsResponse(rics, "");

                    string[] lines = test.Split(stringSeparators, StringSplitOptions.None);
                    foreach (string line in lines)
                    {
                        string formattedLine = rgxSpace.Replace(line, " ");
                        string[] lineTab = formattedLine.Split(stringSeparators2);
                        if (lineTab.Length > 2 && rgxLongLink.IsMatch(lineTab[1]))
                        {
                            string tmpRic = "0" + lineTab[0].Substring(lineTab[0].IndexOf("#"));
                            if (gatsValues.ContainsKey(tmpRic))
                            {
                                if (lineTab[2] != "" && !lineTab[2].StartsWith("."))
                                {
                                    gatsValues[tmpRic].Add(lineTab[2]);
                                }
                            }
                            else
                            {
                                gatsValues.Add(tmpRic, new List<string>());
                                gatsValues[tmpRic].Add(lineTab[2]);
                            }
                        }
                    }

                }


                foreach (KeyValuePair<string, List<string>> value in gatsValues)
                {
                    value.Value.Sort();
                }
                return gatsValues;
            }
            catch (Exception ex)
            {
                throw new Exception("Error While using Gats: " + ex.Message);
            }
        }

        #endregion

        #region Compare

        private bool AreTheyTheSame(List<string> list1, List<string> list2)
        {
            try
            {
                if (list2[0].StartsWith("#") || list2[0].StartsWith("."))
                    list2.RemoveAt(0);
                if (list1.Count != list2.Count)
                {
                    return false;
                }
                else
                {
                    list1.Sort();
                    list2.Sort();
                    int index = 0;
                    foreach (string listEntry in list1)
                    {
                        if (listEntry != list2[index].Substring(0, 6))
                        {
                            return false;
                        }
                        index++;
                    }
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private void CompareGatsWithNew()
        {
            try
            {
                foreach (KeyValuePair<string, List<string>> value in newValues)
                {
                    List<string> resultEntry = new List<string>();
                    if (gatsValues.ContainsKey(value.Key))
                    {
                        if (!AreTheyTheSame(value.Value, gatsValues[value.Key]))
                        {
                            resultEntry.Add(value.Key.Substring(value.Key.IndexOf(".")));
                            resultEntry.AddRange(value.Value);
                            toUpdate.Add(resultEntry);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error happened when comparing new and Gats values: " + ex.Message);
            }
        }

        private static int CompareResultByLen(List<string> list1, List<string> list2)
        {
            if (list1.Count == list2.Count)
            {
                return 0;
            }
            else if (list2.Count > list1.Count)
            {
                return 1;
            }
            return -1;
        }

        #endregion

        #region Write Results in CSV

        private void ReOrderResults()
        {
            toUpdate.Sort(CompareResultByLen);
        }

        private void WriteResultsInCsv()
        {
            Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.ResultFileName);
            try
            {
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                ExcelLineWriter lineWriter = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Down);
                int col = 1;
                bool skip = false;
                foreach (List<string> updateEntry in toUpdate)
                {
                    List<string> tmp = updateEntry;
                    tmp.Sort();
                    skip = false;
                    lineWriter.PlaceNext(1, col);
                    string test = tmp[0];
                    if (ricExl.ContainsKey("0#" + tmp[0]))
                    {
                        lineWriter.WriteLine(ricExl["0#" + tmp[0]]);
                    }

                    foreach (string rowEntry in tmp)
                    {
                        if (rowEntry.StartsWith("1") || rowEntry.StartsWith("5"))
                        {
                            skip = true;
                        }
                        string suffix = "";
                        if (rowEntry.StartsWith("6"))
                        {
                            suffix = ".SS";
                        }
                        else if ((rowEntry.StartsWith("0") && !rowEntry.StartsWith("0#")) || rowEntry.StartsWith("2") || rowEntry.StartsWith("3"))
                        {
                            suffix = ".SZ";
                        }
                        if (!rowEntry.StartsWith("1"))
                        {
                            lineWriter.WriteLine(rowEntry + suffix);
                        }
                    }
                    if (skip == false)
                    {
                        col++;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error when creating result CSV: " + ex.Message);
            }
            finally
            {
                workbook.SaveAs(configObj.ResultFileName, XlFileFormat.xlCSV);
                workbook.Close();
            }
        }

        #endregion
    }

    #endregion
}
