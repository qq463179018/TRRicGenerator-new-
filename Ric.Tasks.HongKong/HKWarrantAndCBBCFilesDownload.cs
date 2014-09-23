using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Web;
using HtmlAgilityPack;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Ric.Util;
using Ric.Core;

namespace Ric.Tasks.HongKong
{
    public class FileList
    {
        public String DateTime { get; set; }
        public String FileName { get; set; }
    }

    public class Date
    {
        public String Year { get; set; }
        public String Month { get; set; }
        public String Day { get; set; }
    }

    public class FID
    {
        public string FileName { get; set; }
        public string WebLink { get; set; }
        public string LocalPath { get; set; }
    }

    public class HKWarrantAndCBBCFilesDownloadConfig
    {
        [Description("From this date the search begins.The Format Is : \"YYYY-MM-DD\", \"2012-02-01")]
        public String START_DATE { get; set; }
        [Description("To this date the search ends.The Format Is : \"YYYY-MM-DD\", \"2012-02-01")]
        public String END_DATE { get; set; }
        [Description("The folder in which to save the downloaded CBBC files")]
        public String CBBC_DOWNLOAD_FOLDER { get; set; }
        [Description("The folder in which to save the downloaded warrant files")]
        public String WARRANT_DOWNLOAD_FOLDER { get; set; }
        [Description("The Folder Is Log Files Path!")]
        public String LOG_FILE_PATH { get; set; }
        [Description("Input the website link from which to download files")]
        public String GENERATE_DOWNLOAD_REPORTS_LIST_PATH { get; set; }
    }

    public class HKWarrantAndCBBCFilesDownload : GeneratorBase
    {
        private static readonly String CONFIGFILE_NAME = ".\\Config\\HK\\HK_WarrantAndCBBCFilesDownload.config";
        private static readonly string HOLIDAY_INFO_FILE_PATH = ".\\Config\\HK\\Holiday.xml";
        private HKWarrantAndCBBCFilesDownloadConfig configObj = null;
        private Logger logger = null;
        private static DateTime startDate = DateTime.Now.AddDays(-1);
        private static DateTime endDate = DateTime.Now;
        private static List<DateTime> holidayList = new List<DateTime>();
        private List<FID> CBBCList = new List<FID>();
        private List<FID> WarrantList = new List<FID>();
        private static string downloadTime = DateTime.Today.ToString("yyyyMMdd");
        private static string generateFilesTime = DateTime.Now.ToString().Replace(":", "-");

        protected override void Initialize()
        {

            base.Initialize();
            try
            {
                configObj = ConfigUtil.ReadConfig(CONFIGFILE_NAME, typeof(HKWarrantAndCBBCFilesDownloadConfig)) as HKWarrantAndCBBCFilesDownloadConfig;
            }
            catch (System.Exception ex)
            {
                Logger.Log("Error happens when initializing task... Ex: " + ex.Message);
            }

            //logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.New);
            holidayList = ConfigUtil.ReadConfig(HOLIDAY_INFO_FILE_PATH, typeof(List<DateTime>)) as List<DateTime>;
            if (!string.IsNullOrEmpty(configObj.END_DATE))
            {
                try
                {
                    DateTime dt = Convert.ToDateTime(configObj.END_DATE.Trim().ToString());
                    double seconds = dt.Subtract(endDate).TotalSeconds;
                    if (seconds > 0)
                    {
                        Logger.Log("EndDate is over Today,Defult is Now!");
                        throw new System.Exception("ENDDATE Setting is error!");
                    }
                    else
                    {
                        endDate = DateTime.ParseExact(configObj.END_DATE, "yyyy-MM-dd", new CultureInfo("en-US"));
                    }
                }
                catch (System.FormatException fEx)
                {
                    Logger.Log("ENDDATE format is error!" + fEx.Message);
                }
                catch (System.Exception ex)
                {
                    Logger.Log(ex.Message);
                }

            }
            if (!string.IsNullOrEmpty(configObj.START_DATE))
            {
                startDate = DateTime.ParseExact(configObj.START_DATE, "yyyy-MM-dd", new CultureInfo("en-US"));
            }
            else
            {
                startDate = MiscUtil.GetLastTradingDay(endDate, holidayList, 1);
            }

            if (string.IsNullOrEmpty(configObj.CBBC_DOWNLOAD_FOLDER))
            {
                configObj.CBBC_DOWNLOAD_FOLDER = @"D:\";
            }
            if (string.IsNullOrEmpty(configObj.WARRANT_DOWNLOAD_FOLDER))
            {
                configObj.WARRANT_DOWNLOAD_FOLDER = @"D:\";
            }
            if (string.IsNullOrEmpty(configObj.GENERATE_DOWNLOAD_REPORTS_LIST_PATH))
            {
                configObj.GENERATE_DOWNLOAD_REPORTS_LIST_PATH = @"D:\";
            }
            #region
            //if (string.IsNullOrEmpty(configObj.START_DATE))
            //{
            //    string[] date = DateTime.Now.AddDays(-1).ToString().Split(' ')[0].Split('-');
            //    startDate.Year = date[0];
            //    startDate.Month = date[1];
            //    startDate.Day = date[2];
            //}
            //else 
            //{
            //    string[] temp = configObj.START_DATE.Split('_');
            //    startDate.Year = temp[0].ToString();
            //    startDate.Month = temp[1].ToString();
            //    startDate.Day = temp[2].ToString();
            //}
            //if (string.IsNullOrEmpty(configObj.END_DATE))
            //{
            //    startDate.Year = DateTime.Now.Year.ToString();
            //    startDate.Month = DateTime.Now.Month.ToString();
            //    startDate.Day = DateTime.Now.Day.ToString();
            //}
            //else 
            //{
            //    string[] temp = configObj.END_DATE.Split('_');
            //    endDate.Year = temp[0].ToString();
            //    endDate.Month = temp[1].ToString();
            //    endDate.Day = temp[2].ToString();
            //}
            #endregion

        }

        protected override void Start()
        {
            try
            {
                StartHKWarrantAndCBBCFilesDownloadJob();
            }
            catch (System.Exception ex)
            {
                Logger.Log("Ex: " + ex.Message + "Stack Trace: " + ex.StackTrace);
            }
        }

        public void StartHKWarrantAndCBBCFilesDownloadJob()
        {
            string url = "http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.aspx";
            CBBCList = getFidList(url, startDate, endDate, 19, 166, "CBBC");
            WarrantList = getFidList(url, startDate, endDate, 17, 162, "Warrant");
            GenerateFiles();
        }

        private void GenerateFiles()
        {
            try
            {
                string CBBC_filename = "CBBC_TradingSummaries_" + generateFilesTime + ".xls";
                string Warrant_filename = "Warrant_TradingSummaries_" + generateFilesTime + ".xls";
                string path = configObj.GENERATE_DOWNLOAD_REPORTS_LIST_PATH;
                GenerateOneFile(CBBC_filename, path, CBBCList);
                GenerateOneFile(Warrant_filename, path, WarrantList);
            }
            catch (System.Exception ex)
            {
                Logger.Log("Can't Generate Reports Files! The Error is : " + ex.Message);
            }
        }

        private void GenerateOneFile(string filename, string path, List<FID> list)
        {
            using (ExcelApp app = new ExcelApp(true, true))
            {
                string filepath = null;
                if (path.EndsWith("\\"))
                    filepath = path + filename;
                else
                    filepath = path + "\\" + filename;
                if (app.ExcelAppInstance == null)
                {
                    String msg = "Excel could not be started. Check that your office installation and project reference correct!!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, filepath);
                Worksheet worksheet = ExcelUtil.GetWorksheet("Sheet1", workbook);

                if (worksheet == null)
                {
                    String msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                ((Range)worksheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                worksheet.Cells[1, 1] = "FileName";
                worksheet.Cells[1, 2] = "LocalPath";
                worksheet.Cells[1, 3] = "Weblink";

                int startline = 2;
                foreach (var item in list)
                {
                    worksheet.Cells[startline, 1] = item.FileName;
                    worksheet.Cells[startline, 2] = item.LocalPath;
                    worksheet.Hyperlinks.Add(ExcelUtil.GetRange(startline, 3, worksheet), item.WebLink, Missing.Value, "Web Link for the file", Missing.Value);
                    startline++;
                }
                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbook.Save();
                if (filename.Contains("CBBC"))
                    AddResult(filename, filepath, "Generate CBBC Reports File");
                if (filename.Contains("Warrant"))
                    AddResult(filename, filepath, "Generate Warrant Reports File");
            }
        }

        private List<FID> getFidList(string url, DateTime startDate, DateTime endDate, int tier2GroupNum, int tier2Num, string type)
        {
            List<FID> FIDList = new List<FID>();
            AdvancedWebClient wc = new AdvancedWebClient();
            string postData = string.Format("__VIEWSTATEENCRYPTED=&ctl00%24txt_today=20130221&ctl00%24hfStatus=AEM&ctl00%24hfAlert=&ctl00%24txt_stock_code=&ctl00%24txt_stock_name=&ctl00%24rdo_SelectDocType=rbAfter2006&ctl00%24sel_tier_1=5&ctl00%24sel_DocTypePrior2006=-1&ctl00%24sel_tier_2_group={0}&ctl00%24sel_tier_2={1}&ctl00%24ddlTierTwo=23%2C1%2C3&ctl00%24ddlTierTwoGroup=10%2C2&ctl00%24txtKeyWord=&ctl00%24rdo_SelectDateOfRelease=rbManualRange&ctl00%24sel_DateOfReleaseFrom_d={2}&ctl00%24sel_DateOfReleaseFrom_m={3}&ctl00%24sel_DateOfReleaseFrom_y={4}&ctl00%24sel_DateOfReleaseTo_d={5}&ctl00%24sel_DateOfReleaseTo_m={6}&ctl00%24sel_DateOfReleaseTo_y={7}&ctl00%24sel_defaultDateRange=SevenDays&ctl00%24rdo_SelectSortBy=rbDateTime", tier2GroupNum, tier2Num, startDate.ToString("dd"), startDate.ToString("MM"), startDate.ToString("yyyy"), endDate.ToString("dd"), endDate.ToString("MM"), endDate.ToString("yyyy"));

            string viewState = getViewStateValue(WebClientUtil.GetPageSource(wc, url, 1800000, postData));
            postData = string.Format("__VIEWSTATE={0}&{1}", viewState, postData);

            string pageSource = WebClientUtil.GetPageSource(wc, url, 1800000, postData);

            FIDList.AddRange(getPerPageFIDList(pageSource, type));

            while (true)
            {
                if (pageSource.Contains("ctl00_gvMain_ctl24_btnNext"))
                {
                    viewState = getViewStateValue(pageSource);
                    postData = "__VIEWSTATEENCRYPTED=&ctl00%24gvMain%24ctl24%24btnNext.x=15&ctl00%24gvMain%24ctl24%24btnNext.y=12";
                    postData = string.Format("__VIEWSTATE={0}&{1}", viewState, postData);
                    pageSource = WebClientUtil.GetPageSource(wc, url, 1800000, postData);
                    FIDList.AddRange(getPerPageFIDList(pageSource, type));
                }
                break;
            }

            if (type == "CBBC")
                AddResult("CBBC_" + downloadTime, Path.Combine(configObj.CBBC_DOWNLOAD_FOLDER, ("CBBC_" + downloadTime)), "Have Finished Download CBBC Files");
            if (type == "Warrant")
                AddResult("Warrant_" + downloadTime, Path.Combine(configObj.WARRANT_DOWNLOAD_FOLDER, ("Warrant_" + downloadTime)), "Have Finished Download Warrant Files");

            return FIDList;
        }


        //To Do:
        private List<FID> getPerPageFIDList(string pageSource, string type)
        {
            List<FID> FIDList = new List<FID>();
            if (pageSource == null)
            {
                return FIDList;
            }
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(pageSource);
            HtmlNodeCollection node = doc.DocumentNode.SelectNodes("//div")[3].SelectNodes(".//table")[0].SelectNodes(".//tr");
            if (node.Count > 5)
            {
                for (int i = 3; i < node.Count - 3; i++)
                {
                    FID fid = new FID();
                    HtmlNode trNode = node[i];
                    var tdNodeList = trNode.SelectNodes("td");
                    string[] date = GetCleanTextFromHtml(tdNodeList[0].InnerText.ToString()).Split('/');
                    string releaseYear = date[2].Substring(0, 4);
                    string time = date[2].Substring(4).Replace(":", "");
                    fid.WebLink = "http://www.hkexnews.hk" + GetCleanTextFromHtml(tdNodeList[3].SelectSingleNode("a").Attributes["href"].Value);
                    fid.FileName = GetCleanTextFromHtml(tdNodeList[3].InnerText.ToString().Split(']')[1].Split('(')[0]) + "_" + releaseYear + date[1] + date[0] + "_" + time + ".xls";
                    while (fid.FileName.Contains("/"))
                        fid.FileName = fid.FileName.Replace("/", "-");
                    if (type == "CBBC")
                    {
                        fid.LocalPath = Path.Combine(string.Format("{0}\\{1}", configObj.CBBC_DOWNLOAD_FOLDER, ("CBBC_" + downloadTime)), fid.FileName);
                    }
                    if (type == "Warrant")
                    {
                        fid.LocalPath = Path.Combine(string.Format("{0}\\{1}", configObj.WARRANT_DOWNLOAD_FOLDER, ("Warrant_" + downloadTime)), fid.FileName);
                    }
                    WebClientUtil.DownloadFile(fid.WebLink, 180000, fid.LocalPath);
                    FIDList.Add(fid);
                }
            }
            return FIDList;
        }

        private string getViewStateValue(string pageSource)
        {
            Regex r = new Regex("id=\"__VIEWSTATE\" value=\"(?<ViewState>.*?)\"\\s{0,}/>");
            Match m = r.Match(pageSource);
            return HttpUtility.UrlEncode(m.Groups["ViewState"].Value);

        }

        private string GetCleanTextFromHtml(string html)
        {
            return HttpUtility.HtmlDecode(html).Replace("\t", "").Replace("\r", "").Replace("\n", "").Trim();
        }
    }
}
