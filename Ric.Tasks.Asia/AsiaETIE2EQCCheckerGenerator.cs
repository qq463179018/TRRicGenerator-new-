using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Office.Interop.Excel;
using MSAD.Common.MiscUtility;
using MSAD.Common.OfficeUtility;
using Ric.Util;
using ExcelApp = Ric.Util.ExcelApp;
using Ric.Core;

namespace Ric.Tasks.Asia
{
    public class ReportContent
    {
        public List<string> HeaderList;
        public List<List<string>> ContentList;
        public string ColToLookup;
        public string ReportFileName;
        public ReportContent()
        {
            HeaderList = new List<string>();
            ContentList = new List<List<string>>();
            ColToLookup = "RIC";
        }
    }

    public class ReportsList
    {
        public string ReportFileName { get; set; }
        public string ColumnName { get; set; }
        public string Column { get; set; }
        public string InformationProvided { get; set; }
        public string ReportType { get; set; }
    }

    [ConfigStoredInDB]
    public class ETIE2EQCCheckerConfig
    {
        [StoreInDB]
        public string ReportsListFilePath { get; set; }

        [StoreInDB]
        public string RccReportsFilePath { get; set; }

        [StoreInDB]
        public string DcrsReportsFilePath { get; set; }

        [StoreInDB]
        public string OptionReportBaseUrl { get; set; }

        [StoreInDB]
        public string FutureReportBaseUrl { get; set; }

        [StoreInDB]
        public string RccShareDrivePath { get; set; }

        [StoreInDB]
        [Description("The format should be like\"2012_06_29_14_10\", if this field not set, default is the last bussiness day.")]
        public string StartTime { get; set; }

        [StoreInDB]
        public string TagListFilePath { get; set; }

        [StoreInDB]
        public string TagListFileSheetName { get; set; }

        [StoreInDB]
        public string OptionRicListFilePath { get; set; }

        [StoreInDB]
        public string OptionRicListSheetName { get; set; }

        [StoreInDB]
        public string ExCodeListFilePath { get; set; }

        [StoreInDB]
        public string ExCodeListSheetName { get; set; }

        [StoreInDB]
        public string HongkongResultFilePath { get; set; }

        [StoreInDB]
        public string JapanResultFilePath { get; set; }

        [StoreInDB]
        public string TaiwanResultFilePath { get; set; }

        [StoreInDB]
        public string KoreaResultFilePath { get; set; }

        [StoreInDB]
        public string ChinaResultFilePath { get; set; }

        [StoreInDB]
        public string ThailandResultFilePath { get; set; }

        [StoreInDB]
        public string IndonesiaResultFilePath { get; set; }

        [StoreInDB]
        public string ResultFileDir { get; set; }

        [StoreInDB]
        public string OseFuturesCheckSheetPath { get; set; }

        [StoreInDB]
        public string OseFuturesCheckSheetName { get; set; }

        [StoreInDB]
        public string TseFuturesCheckSheetPath { get; set; }

        [StoreInDB]
        public string TseFuturesCheckSheetName { get; set; }

        //[Category("MAIL_ACCOUNT")]
        //[Description("Config for getting the email account")]
        //public MailAccount MAIL_ACCOUNT { get; set; }

        [StoreInDB]
        [Category("MAIL_SEARCH_QUERY")]
        [Description("Account name which used to search the target mail, like \"UC159450\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("MAIL_SEARCH_QUERY")]
        [Description("Password")]
        public string Password { get; set; }

        [StoreInDB]
        [Category("MAIL_SEARCH_QUERY")]
        [Description("Domain of the mail account, like \"ten\", \"apac\"")]
        public string Domain { get; set; }

        [StoreInDB]
        [Category("MAIL_SEARCH_QUERY")]
        [Description("Mail address, like \"eti.asia@thomsonreuters.com\"")]
        public string MailAddress { get; set; }


        //[Category("FUTURE_RIC_EMAIL_QUERY")]
        //[Description("Config for getting future ric email")]
        //public RCCReportMailQuery FutureRicEmailQuery { get; set; }

        [StoreInDB]
        [Category("FUTURE_RIC_EMAIL_QUERY")]
        [Description("Target mail sender")]
        public string FutureRicSender { get; set; }

        [StoreInDB]
        [Category("FUTURE_RIC_EMAIL_QUERY")]
        [Description("Mail folder path")]
        public string FutureRicMailFolderPath { get; set; }

        [StoreInDB]
        [Category("FUTURE_RIC_EMAIL_QUERY")]
        [Description("Keyword of mail subject which can specify the target mail")]
        public string FutureRicSubjectKeyword { get; set; }


        //[Category("RCC_REPORT_EMAIL_QUERY")]
        //[Description("Config for getting rcc report email")]
        //public RCCReportMailQuery RccReportEmailQuery { get; set; }

        [StoreInDB]
        [Category("RCC_REPORT_EMAIL_QUERY")]
        [Description("Target mail sender")]
        public string RccReportSender { get; set; }

        [StoreInDB]
        [Category("RCC_REPORT_EMAIL_QUERY")]
        [Description("Mail folder path")]
        public string RccReportMailFolderPath { get; set; }

        [StoreInDB]
        [Category("RCC_REPORT_EMAIL_QUERY")]
        [Description("Keyword of mail subject which can specify the target mail")]
        public string RccReportSubjectKeyword { get; set; }


        [StoreInDB]
        [Description("The mail format should contain the full name, such as \"xxx.xxx@thomsonreuters.com\".")]
        public List<string> AlertMailToList { get; set; }

        [StoreInDB]
        [Description("The mail format should contain the full name, such as \"xxx.xxx@thomsonreuters.com\".")]
        public List<string> AlertMailCcList { get; set; }
    }

    public class HolidayList
    {
        public List<MarketHoliday> MarketHolidayList { get; set; }
    }


    public class MarketHoliday
    {
        public string Name { get; set; }
        public List<DateTime> Holidays { get; set; }
    }

    public class AsiaETIE2EQCCheckerGenerator : GeneratorBase
    {
        #region Properties

        private readonly string holidayFilePath = ".\\Config\\Asia\\Holiday.xml";
        private ETIE2EQCCheckerConfig configObj = null;
        private const string rccType = "RCC";
        private const string dcrsType = "DCRS";
        private Dictionary<String, DateTime> startDate = new Dictionary<String, DateTime>();
        private Dictionary<String, DateTime> endDate = new Dictionary<String, DateTime>();
        private Dictionary<String, List<string>> exCodeList;
        private Dictionary<String, List<string>> ricSufix;
        private Dictionary<String, List<string>> tagList;
        private Dictionary<String, List<string>> newRicList;
        private Dictionary<String, List<string>> sendNewRicReport = new Dictionary<string, List<string>>();
        private String newRicReportContent = null;
        private Dictionary<String, String> marketCondition = new Dictionary<string, string>();
        private DateTime startTime = DateTime.Now.AddDays(-1);
        private DateTime endTime = DateTime.Now.Date;
        private HolidayList holidayList = new HolidayList();
        private String dcrsDownloadTime = DateTime.Now.ToString();
        private ExchangeService service = null;

        //Specify if need to send some alert messages by emails.
        private int alertTag = 0;

        #endregion

        #region Interface implementation

        public override bool IsPrerequisiteMeet()
        {
            //Get to know if there are RCC report files in RCC Share Drive
            bool tag = HasRCCReports(configObj.RccShareDrivePath);

            return tag;
        }

        public override void OnPrerequisiteNotMeet()
        {
            
        }

        public override void ExecuteUnderPrerequisiteNotMeet()
        {
            alertTag = 1;
            Start();
            SaveTaskResult();
        }

        protected override void Start()
        {
            try
            {
                StartETIE2EQCCheckerJob();
            }
            catch (Exception ex)
            {
                LogMessage(ex.Message, Logger.LogType.Error);
                LogMessage(ex.StackTrace, Logger.LogType.Error);
            }
        }

        protected override void Initialize()
        {
            base.Initialize();
            try
            {
                configObj = Config as ETIE2EQCCheckerConfig;
                service = EWSUtility.CreateService(new System.Net.NetworkCredential(configObj.AccountName, configObj.Password, configObj.Domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
            }
            catch (Exception ex)
            {
                LogMessage("Error happens when initializing task... Ex: " + ex.Message, Logger.LogType.Error);
            }
            holidayList = Ric.Util.ConfigUtil.ReadConfig(holidayFilePath, typeof(HolidayList)) as HolidayList;
            startDate = GetLastTradingDay(DateTime.Today, holidayList, 1);
            startTime = Misc.GetLastWorkingDay(DateTime.Today, 1);

            if (!string.IsNullOrEmpty(configObj.StartTime))
            {
                try
                {
                    string[] keys = startDate.Keys.ToArray();
                    foreach (string key in keys)
                    {
                        startTime = DateTime.ParseExact(configObj.StartTime, "yyyy_MM_dd_HH_mm", new CultureInfo("en-US"));
                        startDate[key] = startTime;
                    }
                }
                catch (Exception ex)
                {
                    LogMessage("Error happens when parse date time for start date. Exception: " + ex.Message, Logger.LogType.Error);
                }
            }
            else
            {
                try
                {
                    string[] keys = startDate.Keys.ToArray();
                    foreach (string key in keys)
                    {
                        startDate[key] = startTime;
                    }
                }
                catch (Exception ex)
                {
                    LogMessage("Error happens when parse date time for start date. Exception: " + ex.Message, Logger.LogType.Error);
                }
            }
        }

        private void StartETIE2EQCCheckerJob()
        {
            tagList = GetTagList();
            exCodeList = GetExchangeCodeList();
            ricSufix = GetRicSufix();

            newRicList = GetNewRicList();

            AddReports(newRicList, ricSufix);
            LogMessage("Start to get DCRS Reports **********");
            List<string> dcrsReportList = GetDCRSReports(configObj.OptionReportBaseUrl, configObj.FutureReportBaseUrl);
            LogMessage("Got DCRS Reports **********");

            //Get the path of all suited excel files.
            LogMessage("Start to get RCC Reports **********");
            List<string> rccReportList = GetRCCReports(configObj.RccShareDrivePath);
            LogMessage("Got RCC Reports **********");

            // Get reports

            var reports = new Dictionary<String, List<ReportContent>>();

            foreach (Dictionary<string, List<ReportContent>> dcrsReports in dcrsReportList.Select(file => GetReportAfterComparison("RIC", GetDCRsReportConntent(file), ricSufix)))
            {
                AddReports(dcrsReports, reports);
            }

            if (rccReportList.Count > 0)
            {
                using (var app = new ExcelApp(false, false))
                {
                    foreach (string file in rccReportList)
                    {
                        var rccReports = new Dictionary<String, List<ReportContent>>();
                        var rccReports_New = new Dictionary<String, List<ReportContent>>();
                        var rccReports_Old = new Dictionary<String, List<ReportContent>>();
                        var workbook = ExcelUtil.CreateOrOpenExcelFile(app, file);
                        var worksheet = workbook.Worksheets[1] as Worksheet;
                        if (worksheet == null)
                        {
                            LogMessage("No worksheet for file " + file, Logger.LogType.Warning);
                        }
                        else
                        {
                            ReportContent reportContent = GetRCCReportContent(file, worksheet);
                            
                            string fileName = Path.GetFileNameWithoutExtension(file);
                            //Find the file whose name likes "UsysATMIV RICs".
                            if (fileName.Contains("ATMIV"))
                            {
                                rccReports_Old = GetReportAfterComparison("RIC", reportContent, ricSufix);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysChain RICs Query".
                            if (fileName.Contains("Chain RIC"))
                            {
                                rccReports_Old = GetReportAfterComparison("PSEUDO EXCHANGE", reportContent, exCodeList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysFandO RICs missing underlying".
                            if (fileName.Contains("missing underlying"))
                            {
                                rccReports_Old = GetReportAfterComparison("EXCHANGE", reportContent, exCodeList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysFandO RICs missng segment".
                            if (fileName.Contains("missing segment"))
                            {
                                rccReports_Old = GetReportAfterComparison("TAG", reportContent, tagList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysFuture RICs missing newfacts".
                            if (fileName.Contains("RICs missing newfacts"))
                            {
                                rccReports_Old = GetReportAfterComparison("TAG", reportContent, tagList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysFutures RICs missing IDN facts".
                            if (fileName.Contains("IDN facts") && !fileName.Contains("Option RICs missing IDN facts"))
                            {
                                rccReports_Old = GetReportAfterComparison("TAG", reportContent, tagList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysNewRootMissingFacts".
                            if (fileName.Contains("NewRoot"))
                            {
                                rccReports_Old = GetReportAfterComparison("CLA TAG  VALUE D", reportContent, tagList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysNewStubMissingFacts".
                            if (fileName.Contains("NewStub"))
                            {
                                rccReports_Old = GetReportAfterComparison("RIC STUB", reportContent, ricSufix);
                                rccReports_New = GetReportAfterComparison("RIC STUB", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysOption RICs missing IDN facts".
                            if (fileName.Contains("Option RICs missing IDN facts"))
                            {
                                rccReports_Old = GetReportAfterComparison("TAG", reportContent, tagList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "Usysoptions rics missing newfacts".
                            if (fileName.Contains("rics missing newfacts"))
                            {
                                rccReports_Old = GetReportAfterComparison("TAG", reportContent, tagList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysRICs not tagged futures".
                            if (fileName.Contains("not tagged futures"))
                            {
                                rccReports_Old = GetReportAfterComparison("EXCHANGE", reportContent, exCodeList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            //Find the file whose name likes "UsysRICs not tagged options".
                            if (fileName.Contains("not tagged options"))
                            {
                                rccReports_Old = GetReportAfterComparison("EXCHANGE", reportContent, exCodeList);
                                rccReports_New = GetReportAfterComparison("RIC", reportContent, newRicList);
                            }
                            rccReports = MergeDcrReports(rccReports_New, rccReports_Old);
                        }
                        workbook.Close(false, workbook.FullName, false);
                        AddReports(rccReports, reports);
                    }
                }
            }
            else
            {
                alertTag = 1;
            }

            // Generate the result file of all markets respectively

            foreach (string marketKey in reports.Keys)
            {
                string resultFilePath = GetResultFilePath(configObj, marketKey);
                //Save the reports to different files according to the different markets.
                LogMessage("Start to generate Result files **********");
                GenerateResultFile(resultFilePath, reports[marketKey]);
                LogMessage("Finished generating Result files **********");
                string content = GetResultContent(reports[marketKey]);
                marketCondition[marketKey] = content;
            }

            newRicReportContent = GetNewRicReportContent();
            LogMessage("Start to send emails **********");
            SendAlertMail(service, marketCondition);
            LogMessage("Finished sending emails **********");
        }

        #endregion

        private Dictionary<String, List<ReportContent>> MergeDcrReports(Dictionary<String, List<ReportContent>> reportDic_1, Dictionary<String, List<ReportContent>> reportDic_2)
        {
            int keyCount_1 = reportDic_1.Keys.Count;
            int keyCount_2 = reportDic_2.Keys.Count;
            var mergeDic = (keyCount_1 <= keyCount_2) ? GetMergeResult(reportDic_1, reportDic_2) : GetMergeResult(reportDic_2, reportDic_1);
            return mergeDic;
        }

        //merge the lesser information Dic into the more
        private Dictionary<String, List<ReportContent>> GetMergeResult(Dictionary<String, List<ReportContent>> reportDicLess, Dictionary<String, List<ReportContent>> reportDicMore)
        {
            foreach (string key in reportDicLess.Keys)
            {
                if (reportDicMore.Keys.Contains(key))
                {
                    foreach (ReportContent reportLess in reportDicLess[key])
                    {
                        foreach (ReportContent reportMore in reportDicMore[key].Where(reportMore => reportMore.ReportFileName.Equals(reportLess.ReportFileName) && reportLess.ContentList.Count != 0))
                        {
                            reportMore.ContentList.AddRange(reportLess.ContentList);
                        }
                    }
                }
                else
                {
                    reportDicMore.Add(key, reportDicLess[key]);
                }
            }
            return reportDicMore;
        }

        /// <summary>
        /// Read the TAG List worksheet and Get Set of Several Markets Tags Respectively.
        /// </summary>
        /// <returns>The Dictionary structure whose keys are markets' name and value are the corresponding tag lists. </returns>
        private Dictionary<String, List<string>> GetTagList()
        {
            var tagList = new Dictionary<String, List<string>>();

            using (var app = new ExcelApp(false, false))
            {
                var hkMarket = new List<string>();
                var krMarket = new List<string>();
                var chMarket = new List<string>();
                var twMarket = new List<string>();
                var thMarket = new List<string>();
                var inMarket = new List<string>();
                var jpMarket = new List<string>();

                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.TagListFilePath);
                Worksheet worksheet = ExcelUtil.GetWorksheet(configObj.TagListFileSheetName, workbook);
                if (worksheet == null)
                {
                    LogMessage(string.Format("Cannot find the worksheet {0}in file {1}. ", configObj.TagListFileSheetName, configObj.TagListFilePath), Logger.LogType.Warning);
                }
                else
                {
                    int lastUsedCol = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                    int colToLookup = 4;
                    int marketCol = 1;
                    bool findTag = false;
                    bool findMarket = false;

                    for (int i = 1; i <= lastUsedCol; i++)
                    {
                        Range range = ExcelRangeUtil.GetRange(1, i, worksheet);
                        string colName = range.Text.ToString().ToLower();

                        if (colName == "tag")
                        {
                            colToLookup = i;
                            findTag = true;
                        }
                        if (colName == "market")
                        {
                            marketCol = i;
                            findMarket = true;
                        }
                        if ((findTag) && (findMarket))
                        {
                            break;
                        }
                    }
                    using (var reader = new ExcelLineWriter(worksheet, 2, colToLookup, ExcelLineWriter.Direction.Down))
                    {
                        int j = 2;
                        string marketName = "";
                        while (reader.Row <= lastUsedRow)
                        {
                            string text = reader.ReadLineCellText();
                            if (!string.IsNullOrEmpty(text))
                            {
                                Range range = ExcelRangeUtil.GetRange(j, marketCol, worksheet);
                                marketName = range.Text.ToString().Replace(" ", "").ToLower();
                                switch (marketName)
                                {
                                    case "hongkong":
                                        hkMarket.Add(text);
                                        break;
                                    case "japan":
                                        jpMarket.Add(text);
                                        break;
                                    case "taiwan":
                                        twMarket.Add(text);
                                        break;
                                    case "korea":
                                        krMarket.Add(text);
                                        break;
                                    case "thailand":
                                        thMarket.Add(text);
                                        break;
                                    case "china":
                                        chMarket.Add(text);
                                        break;
                                    case "indonesia":
                                        inMarket.Add(text);
                                        break;
                                }
                            }
                            j++;
                        }
                        tagList.Add("hongkong", hkMarket);
                        tagList.Add("korea", krMarket);
                        tagList.Add("china", chMarket);
                        tagList.Add("taiwan", twMarket);
                        tagList.Add("thailand", thMarket);
                        tagList.Add("japan", jpMarket);
                        tagList.Add("indonesia", inMarket);
                    }
                }
                workbook.Close(false, workbook.FullName, false);
            }
            return tagList;
        }

        /// <summary>
        /// Read the Exchange Code worksheet and Get Set of several Markets' Exchange Codes respectively.
        /// </summary>
        /// <returns>The Dictionary structure whose keys are markets' name and value are the corresponding Exchange Code lists.</returns>
        private Dictionary<String, List<string>> GetExchangeCodeList()
        {
            var exchangeCodeList = new Dictionary<String, List<string>>();

            using (var app = new ExcelApp(false, false))
            {
                var hkMarket = new List<string>();
                var krMarket = new List<string>();
                var chMarket = new List<string>();
                var twMarket = new List<string>();
                var thMarket = new List<string>();
                var inMarket = new List<string>();
                var jpMarket = new List<string>();

                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.ExCodeListFilePath);
                Worksheet worksheet = ExcelUtil.GetWorksheet(configObj.ExCodeListSheetName, workbook);
                if (worksheet == null)
                {
                    LogMessage(string.Format("Cannot find the worksheet {0}in file {1}. ", configObj.ExCodeListSheetName, configObj.ExCodeListFilePath), Logger.LogType.Warning);
                }
                else
                {
                    int lastUsedCol = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                    for (int i = 1; i <= lastUsedCol; i++)
                    {
                        using (var reader = new ExcelLineWriter(worksheet, 1, i, ExcelLineWriter.Direction.Down))
                        {
                            Range range = ExcelUtil.GetRange(1, i, worksheet);
                            string marketName = range.Text.ToString().Replace(" ", "").ToLower();
                            reader.MoveNext();
                            while (reader.Row <= lastUsedRow)
                            {
                                string text = reader.ReadLineCellText();
                                if (!string.IsNullOrEmpty(text))
                                {
                                    switch (marketName)
                                    {
                                        case "hongkong":
                                            hkMarket.Add(text);
                                            break;
                                        case "japan":
                                            jpMarket.Add(text);
                                            break;
                                        case "taiwan":
                                            twMarket.Add(text);
                                            break;
                                        case "korea":
                                            krMarket.Add(text);
                                            break;
                                        case "thailand":
                                            thMarket.Add(text);
                                            break;
                                        case "china":
                                            chMarket.Add(text);
                                            break;
                                        case "indonesia":
                                            inMarket.Add(text);
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    exchangeCodeList.Add("hongkong", hkMarket);
                    exchangeCodeList.Add("korea", krMarket);
                    exchangeCodeList.Add("china", chMarket);
                    exchangeCodeList.Add("taiwan", twMarket);
                    exchangeCodeList.Add("thailand", thMarket);
                    exchangeCodeList.Add("japan", jpMarket);
                    exchangeCodeList.Add("indonesia", inMarket);
                }
                workbook.Close(false, workbook.FullName, false);
            }
            return exchangeCodeList;
        }

        /// <summary>
        /// Get the new ric lists from a special email folder and two excle file whose name likes  TSEFuturesCheckSheetR1 and OSEFuturesCheckSheetR1.
        /// </summary>
        /// <returns>A dictionary data structure which contains all markets' new ric lists.</returns>
        private Dictionary<String, List<string>> GetNewRicList()
        {
            var ricList = new Dictionary<string, List<string>>();
            var tempRicList = new List<string>();
            //Get ricList from TSEFuturesCheckSheetR1.xls excel file
            GetRicList(tempRicList, "TSE");

            //Get ricList from OSEFuturesCheckSheetR1.xls excel file
            GetRicList(tempRicList, "OSE");

            ricList.Add("japan", tempRicList);
            //Get ricList from mail
            Dictionary<String, List<string>> ricListFromMail = GetRicListFromMail(configObj.MailAddress, configObj.FutureRicSender, configObj.FutureRicMailFolderPath, configObj.FutureRicSubjectKeyword);
            AddReports(ricListFromMail, ricList);
            return ricList;
        }

        private void GetRicList(List<string> list, string str)
        {
            using (var app = new ExcelApp(false, false))
            {
                Workbook workbook = null;
                Worksheet worksheet = null;
                int[] colIndex;
                string state = null;
                if (str.Equals("TSE"))
                {
                    workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.TseFuturesCheckSheetPath);
                    worksheet = ExcelUtil.GetWorksheet(configObj.TseFuturesCheckSheetName, workbook);
                    if (worksheet == null)
                    {
                        LogMessage(string.Format("Cannot find the worksheet {0} in file {1}. ", configObj.TseFuturesCheckSheetName, configObj.TseFuturesCheckSheetPath), Logger.LogType.Warning);
                    }
                    colIndex = new[] { 2, 3, 4, 10, 19, 20, 21, 23, 24, 25, 27, 28, 29, 31, 32, 33 };
                    state = "TSE";
                }
                else    //"OSE"
                {
                    workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.OseFuturesCheckSheetPath);
                    worksheet = ExcelUtil.GetWorksheet(configObj.OseFuturesCheckSheetName, workbook);
                    if (worksheet == null)
                    {
                        LogMessage(string.Format("Cannot find the worksheet {0} in file {1}. ", configObj.OseFuturesCheckSheetName, configObj.OseFuturesCheckSheetPath), Logger.LogType.Warning);
                    }
                    colIndex = new[] { 2, 3, 4, 14, 19, 20, 21 };
                    state = "OSE";
                }
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                int startLine = 3;
                int colCount = colIndex.Length;

                while (startLine <= lastUsedRow)
                {
                    int i = 0;
                    while (i < colCount)
                    {
                        Range range = ExcelUtil.GetRange(startLine, colIndex[i], worksheet);
                        string temp = range.Text.ToString().Replace(" ", "");
                        if (!string.IsNullOrEmpty(temp)) list.Add(temp);
                        i++;
                    }
                    if (state.Equals("TSE")) startLine += 6;     //Go to next new line to get record
                    if (state.Equals("OSE")) startLine += 7;
                }
                workbook.Close(false, workbook.FullName, false);
            }
        }

        /// <summary>
        /// Get the different mail Subject's real market name and the corresponding  NEW RIC list
        /// </summary>
        /// <param name="account">The email account which receives the emails.</param>
        /// <param name="newRicQuery">The query key words of the future ric emails.</param>
        /// <returns>A dictionary data structure which keys are market name and values are the corresponding new ric lists</returns>
        private Dictionary<String, List<string>> GetRicListFromMail(string mailAddress, string sender, string folderPath, string subjectKeyword)
        {
            var list = new Dictionary<string, List<string>>();
            try
            {
                var query = new EWSMailSearchQuery(sender, mailAddress, folderPath, subjectKeyword, "", startTime, DateTime.Now);

                List<EmailMessage> mailList = EWSMailSearchQuery.SearchMail(service, query);
                foreach (EmailMessage mail in mailList)
                {
                    mail.Load();
                    List<string> ricListFromMail = GetMailContent(mail);
                    sendNewRicReport.Add(mail.Subject, ricListFromMail);
                    if (mail.Subject.Contains("Exchanges-TM"))
                    {
                        List<string> ricListFromMail_TW = ricListFromMail.ToList();
                        if (list.Keys.Contains("taiwan"))
                            list["taiwan"].AddRange(ricListFromMail_TW);
                        else
                            list.Add("taiwan", ricListFromMail_TW);
                    }
                    if (mail.Subject.Contains("Exchanges-HF"))
                    {
                        List<string> ricListFromMail_HK = ricListFromMail.ToList();
                        List<string> ricListFromMail_CN = ricListFromMail.ToList();
                        if (list.Keys.Contains("hongkong"))
                            list["hongkong"].AddRange(ricListFromMail_HK);
                        else
                            list.Add("hongkong", ricListFromMail_HK);

                        if (list.Keys.Contains("china"))
                            list["china"].AddRange(ricListFromMail_CN);
                        else
                            list.Add("china", ricListFromMail_CN);
                    }
                    if (mail.Subject.Contains("Exchanges-SB"))
                    {
                        List<string> ricListFromMail_IN = ricListFromMail.ToList();
                        if (list.Keys.Contains("indonesia"))
                            list["indonesia"].AddRange(ricListFromMail_IN);
                        else
                            list.Add("indonesia", ricListFromMail_IN);
                    }
                    if (mail.Subject.Contains("Region-KOR"))
                    {
                        List<string> ricListFromMail_KR = ricListFromMail.ToList();
                        if (list.Keys.Contains("korea"))
                            list["korea"].AddRange(ricListFromMail_KR);
                        else
                            list.Add("korea", ricListFromMail_KR);
                    }
                    if (mail.Subject.Contains("Region-TOK"))
                    {
                        List<string> ricListFromMail_JP = ricListFromMail.ToList();
                        if (list.Keys.Contains("japan"))
                            list["japan"].AddRange(ricListFromMail_JP);
                        else
                            list.Add("japan", ricListFromMail_JP);
                    }
                    if (mail.Subject.Contains("Exchanges-BA|AFET") || mail.Subject.Contains("Exchanges-BK|SET"))
                    {
                        List<string> ricListFromMail_TH = ricListFromMail.ToList();
                        if (list.Keys.Contains("thailand"))
                            list["thailand"].AddRange(ricListFromMail_TH);
                        else
                            list.Add("thailand", ricListFromMail_TH);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage("Get RCC reports from mail failed. Ex: " + ex.Message, Logger.LogType.Error);
            }
            return list;
        }

        /// <summary>
        /// Get NEW RIC from the special email.
        /// </summary>
        /// <param name="item">The special email which contains NEW RIC.</param>
        /// <returns>The NEW RIC list.</returns>
        private List<string> GetMailContent(EmailMessage item)
        {
            var temp = new List<string>();
            string originalBody = item.Body;
            string formatBody_1 = originalBody.Replace("\r\n", "");
            if (formatBody_1.Contains("***********F U T U R E    R E P O R T I N G**************"))
            {
                var regex = new Regex("E X P I R E D   R I C s   R E P O R T(.|\n)*F U T U R E    R E P O R T I N G");
                Match result = regex.Match(formatBody_1);
                if (result.Success) formatBody_1 = result.Value;
            }
            else
            {
                var regex = new Regex("E X P I R E D   R I C s   R E P O R T(.|\n)*LISTS SUBMITTED FOR GENERATION:");
                Match result = regex.Match(formatBody_1);
                if (result.Success) formatBody_1 = result.Value;
            }
            string[] formatBody_2 = formatBody_1.Split('-');
            for (int i = 0; i < formatBody_2.Length; i++)
            {
                //Except some irrelevant lines
                if (formatBody_2[i].Contains("*") || formatBody_2[i].Contains("Exchange") || string.IsNullOrEmpty(formatBody_2[i]) || formatBody_2[i].Contains("LISTS SUBMITTED FOR GENERATION:")) continue;

                //Find the relevant line
                int j = i + 2;

                //If not be out of the texts
                if (j < formatBody_2.Length)
                {
                    //If the line doesn't contain blank that indicates there is not the New Ric
                    if (!formatBody_2[j].Contains(" "))
                    {
                        i += 2;     //Skip to the next
                        continue;
                    }

                    //There is the New Ric 
                    //Divide the line into the groups by blanks
                    string[] group = formatBody_2[j].Split(' ');
                    int count = 0;
                    foreach (string grp in @group.Where(grp => !string.IsNullOrEmpty(grp)))
                    {
                        count++;

                        //Find the second element in the group, which is the New Ric column
                        if (count == 2)
                        {
                            temp.Add(grp);
                            i += 4;
                            break;
                        }
                    }
                }
                else
                {
                    break;
                }
            }
            return temp;
        }

        /// <summary>
        /// Read the RIC Sufix worksheet and Get Set of several Markets' Exchange Codes respectively.
        /// </summary>
        /// <returns>The Dictionary structure whose keys are markets' name and value are the corresponding Ric Sufix lists.</returns>
        private Dictionary<String, List<string>> GetRicSufix()
        {
            var ricSufixList = new Dictionary<String, List<string>>();

            using (var app = new ExcelApp(false, false))
            {
                var hkMarket = new List<string>();
                var krMarket = new List<string>();
                var chMarket = new List<string>();
                var twMarket = new List<string>();
                var thMarket = new List<string>();
                var inMarket = new List<string>();
                var jpMarket = new List<string>();

                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.OptionRicListFilePath);
                Worksheet worksheet = ExcelUtil.GetWorksheet(configObj.OptionRicListSheetName, workbook);
                if (worksheet == null)
                {
                    LogMessage(string.Format("Cannot find the worksheet {0}in file {1}. ", configObj.OptionRicListSheetName, configObj.OptionRicListFilePath), Logger.LogType.Warning);
                }
                else
                {
                    int lastUsedCol = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                    for (int i = 1; i <= lastUsedCol; i++)
                    {
                        using (var reader = new ExcelLineWriter(worksheet, 1, i, ExcelLineWriter.Direction.Down))
                        {
                            Range range = ExcelUtil.GetRange(1, i, worksheet);
                            string marketName = range.Text.ToString().Replace(" ", "").ToLower();
                            reader.MoveNext();
                            while (reader.Row <= lastUsedRow)
                            {
                                string text = reader.ReadLineCellText();
                                if (!string.IsNullOrEmpty(text))
                                {
                                    switch (marketName)
                                    {
                                        case "hongkong":
                                            hkMarket.Add(text);
                                            break;
                                        case "japan":
                                            jpMarket.Add(text);
                                            break;
                                        case "taiwan":
                                            twMarket.Add(text);
                                            break;
                                        case "korea":
                                            krMarket.Add(text);
                                            break;
                                        case "thailand":
                                            thMarket.Add(text);
                                            break;
                                        case "china":
                                            chMarket.Add(text);
                                            break;
                                        case "indonesia":
                                            inMarket.Add(text);
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    ricSufixList.Add("hongkong", hkMarket);
                    ricSufixList.Add("korea", krMarket);
                    ricSufixList.Add("china", chMarket);
                    ricSufixList.Add("taiwan", twMarket);
                    ricSufixList.Add("thailand", thMarket);
                    ricSufixList.Add("japan", jpMarket);
                    ricSufixList.Add("indonesia", inMarket);
                }
                workbook.Close(false, workbook.FullName, false);
            }
            return ricSufixList;
        }

        private int GetIndex(string str, Worksheet worksheet)
        {
            int colNumber = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            int index = 0;
            for (int i = 1; i <= colNumber; i++)
            {
                Range range = ExcelRangeUtil.GetRange(1, i, worksheet);
                string colName = range.Text.ToString().Replace(" ", "").ToLower();
                if (colName == str)
                {
                    index = i;
                }
            }
            if (index != 0)
            {
                return index;
            }
            return -1;
        }

        #region GET DCRS Reports from website
        private List<string> GetDCRSReports(string optionUrl, string futureUrl)
        {
            var dcrsReportFiles = new List<string>();
            dcrsDownloadTime = DateTime.Now.ToString("DCRS_ddMMMHHmm");
            dcrsReportFiles.AddRange(GetDCRSReports(optionUrl));
            dcrsReportFiles.AddRange(GetDCRSReports(futureUrl));
            return dcrsReportFiles;
        }

        /// <summary>
        /// Download all the files from the special url and Get all the files path which are been downloaded.
        /// </summary>
        /// <param name="url">A web url.</param>
        /// <returns>A list data structure which consists of the path of all the downloaded files.</returns>
        private List<string> GetDCRSReports(string url)
        {
            var reportFiles = new List<string>();
            if (url == null)
            {
                return reportFiles;
            }
            HtmlDocument htmlDoc = WebClientUtil.GetHtmlDocument(url, 180000);
            var nodeList = htmlDoc.DocumentNode.SelectNodes("//tr");
            for (int i = 1; i < nodeList.Count; i++)
            {
                HtmlNode node = nodeList[i];
                var tdNodeList = node.SelectNodes("td");
                DateTime releaseTime = DateTime.ParseExact(Misc.GetCleanTextFromHtml(tdNodeList[1].InnerText), "dd/MM/yyyy HH:mm:ss", new CultureInfo("en-US"));
                if (releaseTime >= startTime && releaseTime <= endTime)
                {
                    string fileUrl = Misc.GetCleanTextFromHtml(tdNodeList[0].SelectSingleNode("a").Attributes["href"].Value);
                    string fileLocalPath = Path.Combine(string.Format("{0}\\{1}", configObj.DcrsReportsFilePath, dcrsDownloadTime), Misc.GetCleanTextFromHtml(tdNodeList[0].SelectSingleNode("a").InnerText));
                    WebClientUtil.DownloadFile(fileUrl, 18000, fileLocalPath);
                    reportFiles.Add(fileLocalPath);
                }
            }
            return reportFiles;
        }
        #endregion

        #region GET RCCReports from Share Driver or Mail
        private bool HasRCCReports(string rccShareDrivePath)
        {
            string month = DateTime.Now.ToString("MMM");
            string day = DateTime.Now.Day.ToString("D2");
            List<string> rccReportsFullPath = new List<string>();

            string path1 = month;
            path1 = Path.Combine(rccShareDrivePath, path1);
            string path2 = string.Format("RCC {0} {1}", month, day);
            path2 = Path.Combine(path1, path2);
            return Directory.Exists(path2);
        }

        //Get the full path of the all files which are RCC Reports from RCC share drive.
        private List<string> GetRCCReports(string rccShareDrivePath)
        {
            string month = DateTime.Now.ToString("MMM");
            string day = DateTime.Now.Day.ToString("D2");
            var rccReportsFullPath = new List<string>();

            string path1 = month;
            path1 = Path.Combine(rccShareDrivePath, path1);
            string path2=string.Format("RCC {0} {1}", month, day);
            path2 = Path.Combine(path1, path2);
            if (!Directory.Exists(path2))
            {
                return rccReportsFullPath;
            }
            string folderPath = string.Format("{0}\\{1}",configObj.RccReportsFilePath,DateTime.Now.ToString("RCC_ddMMM_HHmm"));
            foreach (var file in Directory.GetFiles(path2, "*.xlsx",SearchOption.AllDirectories))
            {
                CopyFile(folderPath, file);
                rccReportsFullPath.Add(Path.Combine(folderPath, Path.GetFileName(file)));
            }
            return rccReportsFullPath;
        }

        private void CopyFile(string destPath, string fileFullPath)
        {
            var file = new FileInfo(fileFullPath);
            if (!Directory.Exists(destPath))
            {
                Directory.CreateDirectory(destPath);
            }
            string destFilePath = Path.Combine(destPath, Path.GetFileName(fileFullPath));
                
            file.CopyTo(destFilePath);
        }

        /// <summary>
        /// Download the attachment files from the special mail Inbox folder and Get the path of all the attachment files.
        /// </summary>
        /// <param name="account">The email account which receives the emails about RRC.</param>
        /// <param name="mailQuery">The email query.</param>
        /// <returns>A list data structure which consists of the path of all the RCC report files.</returns>
        private List<string> DownloadRCCReports()
        {
            List<string> rccReports = new List<string>();

            try
            {
                EWSMailSearchQuery query = new EWSMailSearchQuery(configObj.RccReportSender, configObj.MailAddress, configObj.RccReportMailFolderPath, configObj.RccReportSubjectKeyword, "", startTime, DateTime.Now);

                List<EmailMessage> mailList = EWSMailSearchQuery.SearchMail(service, query);

                foreach (EmailMessage mail in mailList)
                {
                    mail.Load();
                    string subject = mail.Subject;
                    List<string> attachments = EWSMailSearchQuery.DownloadAttachments(service, mail, "", "", string.Format("{0}\\{1}", configObj.RccReportsFilePath, DateTime.Now.ToString("RCC_ddMMM_HHmm")));
                    rccReports.AddRange(attachments);
                }
            }
            catch (Exception ex)
            {
                LogMessage("Get RCC reports from mail failed. Ex: " + ex.Message, Logger.LogType.Error);
            }
            return rccReports;
        }

        #endregion

        //After Comparing according the arguments "lookupStr" and waysToLookup , get all the markets' Report Content.
        private Dictionary<String, List<ReportContent>> GetReportAfterComparison(string lookupStr, ReportContent originalReport, Dictionary<String, List<string>> waysToLookup)
        {
            var reportContent = new Dictionary<string, List<ReportContent>>();
            foreach (string key in waysToLookup.Keys)
            {
                var report = new ReportContent
                {
                    ColToLookup = lookupStr,
                    HeaderList = originalReport.HeaderList,
                    ReportFileName = originalReport.ReportFileName
                };
                var list = new List<ReportContent>();
                int i = 0;
                int keyCount = waysToLookup[key].Count;
                int ricIdx = report.HeaderList.IndexOf(lookupStr);

                if (ricIdx < 0) return reportContent;

                if ((lookupStr == "RIC") || (lookupStr == "RIC STUB"))
                {
                    while (i < originalReport.ContentList.Count)
                    {
                        for (int j = 0; j < keyCount; j++)
                        {
                            if (originalReport.ContentList[i][ricIdx].Contains('.'))
                            {
                                int dotIdx = originalReport.ContentList[i][ricIdx].LastIndexOf('.');
                                string subString = originalReport.ContentList[i][ricIdx].Substring(dotIdx);
                                if (subString == waysToLookup[key][j])
                                {
                                    report.ContentList.Add(originalReport.ContentList[i]);
                                }
                            }
                            else if (originalReport.ContentList[i][ricIdx].Length >= 3)
                            {
                                if ((originalReport.ContentList[i][ricIdx].Substring(0, 3) == waysToLookup[key][j]))
                                {
                                    report.ContentList.Add(originalReport.ContentList[i]);
                                }
                                //Match the New Ric List
                                if (originalReport.ContentList[i][ricIdx].Equals(waysToLookup[key][j]))
                                {
                                    report.ContentList.Add(originalReport.ContentList[i]);
                                }
                            }
                        }
                        i++;
                    }
                    list.Add(report);
                    reportContent.Add(key, list);
                }
                else
                {
                    while (i < originalReport.ContentList.Count)
                    {
                        for (int j = 0; j < keyCount; j++)
                        {
                            if (originalReport.ContentList[i][ricIdx] == waysToLookup[key][j])
                            {
                                report.ContentList.Add(originalReport.ContentList[i]);
                            }
                        }
                        i++;
                    }
                    list.Add(report);
                    reportContent.Add(key, list);
                }
            }
            return reportContent;
        }

        #region Collect the header and the content from the report

        private ReportContent GetDCRsReportConntent(string reportFile)
        {
            var report = new ReportContent {ReportFileName = Path.GetFileNameWithoutExtension(reportFile)};
            string[] contents = File.ReadAllLines(reportFile);
            report.HeaderList = contents[0].Split(',').ToList();
            for (int i = 1; i < contents.Length; i++)
            {
                report.ContentList.Add(contents[i].Split(',').ToList());
            }
            return report;
        }

        private ReportContent GetRCCReportContent(string reportFile, Worksheet worksheet)
        {
            var report = new ReportContent();
            int lastUsedRow = worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row - 1;
            int lastUsedCol = worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column - 1;

            report.ReportFileName = Path.GetFileNameWithoutExtension(reportFile);
            using (var reader = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
            {
                while (reader.Col <= lastUsedCol)
                {
                    report.HeaderList.Add(reader.ReadLineCellText());
                }
                reader.Reset(worksheet, 2, 1, ExcelLineWriter.Direction.Right);

                for (int i = 2; i <= lastUsedRow; i++)
                {
                    var contents = new List<string>();
                    while (reader.Col <= lastUsedCol)
                    {
                        contents.Add(reader.ReadLineCellText());
                    }
                    report.ContentList.Add(contents);
                    reader.PlaceNext(reader.Row + 1, 1);
                }
            }
            return report;
        }
        #endregion

        #region Make every file information together

        private void AddReports(Dictionary<String, List<ReportContent>> originalReports, Dictionary<String, List<ReportContent>> resultReports)
        {
            foreach (string key in originalReports.Keys)
            {
                if (resultReports.Keys.Contains(key))
                {
                    resultReports[key].AddRange(originalReports[key]);
                }
                else
                {
                    resultReports.Add(key, originalReports[key]);
                }
            }
        }

        /// <summary>
        /// Add a special dictionary which keys are market name and values are the corresponding new ric lists to the result dictionary which already
        /// contains some markets' new ric records.
        /// </summary>
        /// <param name="originalReports">A special dictionary data structure which contains some markets's ric records.</param>
        /// <param name="resultReports">The result dictionary which already contains some markets' new ric records.</param>
        private void AddReports(Dictionary<String, List<string>> originalReports, Dictionary<String, List<string>> resultReports)
        {
            foreach (string key in originalReports.Keys)
            {
                if (resultReports.Keys.Contains(key))
                {
                    resultReports[key].AddRange(originalReports[key]);
                }
                else
                {
                    resultReports.Add(key, originalReports[key]);
                }
            }
        }

        #endregion

        #region generate result file
        private string GetResultFilePath(ETIE2EQCCheckerConfig configObj, string marketKey)
        {
            string resultFile = "";
            if (marketKey == "hongkong")
            {
                resultFile = Path.Combine(configObj.HongkongResultFilePath, string.Format("{0}_ETIE2EQCCheck_HONGKONG.xls", DateTime.Today.ToString("ddMMyyHHmm")));
            }
            if (marketKey == "japan")
            {
                resultFile = Path.Combine(configObj.JapanResultFilePath, string.Format("{0}_ETIE2EQCCheck_JAPAN.xls", DateTime.Today.ToString("ddMMyyHHmm")));
            }
            if (marketKey == "taiwan")
            {
                resultFile = Path.Combine(configObj.TaiwanResultFilePath, string.Format("{0}_ETIE2EQCCheck_TAIWAN.xls", DateTime.Today.ToString("ddMMyyHHmm")));
            }
            if (marketKey == "china")
            {
                resultFile = Path.Combine(configObj.ChinaResultFilePath, string.Format("{0}_ETIE2EQCCheck_CHINA.xls", DateTime.Today.ToString("ddMMyyHHmm")));
            }
            if (marketKey == "thailand")
            {
                resultFile = Path.Combine(configObj.ThailandResultFilePath, string.Format("{0}_ETIE2EQCCheck_THAILAND.xls", DateTime.Today.ToString("ddMMyyHHmm")));
            }
            if (marketKey == "korea")
            {
                resultFile = Path.Combine(configObj.KoreaResultFilePath, string.Format("{0}_ETIE2EQCCheck_KOREA.xls", DateTime.Today.ToString("ddMMyyHHmm")));
            }
            if (marketKey == "indonesia")
            {
                resultFile = Path.Combine(configObj.IndonesiaResultFilePath, string.Format("{0}_ETIE2EQCCheck_INDONESIA.xls", DateTime.Today.ToString("ddMMyyHHmm")));
            }
            return resultFile;
        }

        //Generate result file
        private void GenerateResultFile(string resultFilePath, List<ReportContent> reports)
        {
            using (var app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, resultFilePath);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                using (var writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    foreach (ReportContent report in reports.Where(report => report.ContentList != null && report.ContentList.Count != 0))
                    {
                        writer.WriteInPlace(report.ReportFileName);
                        writer.PlaceNext(writer.Row + 1, 1);
                        foreach (string header in report.HeaderList)
                        {
                            writer.WriteLine(header);
                        }
                        writer.PlaceNext(writer.Row + 1, 1);
                        foreach (List<string> contents in report.ContentList)
                        {
                            foreach (string contentField in contents)
                            {
                                writer.WriteLine(contentField);
                            }
                            writer.PlaceNext(writer.Row + 1, 1);
                        }
                        writer.PlaceNext(writer.Row + 1, 1);
                    }
                }
                workbook.Save();
                workbook.Close(true, workbook.FullName, false);
            }
        }
        #endregion

        #region send a mail
        //Get result content.
        private string GetResultContent(List<ReportContent> reports)
        {
            var sb = new StringBuilder();
            foreach (ReportContent report in reports.Where(report => report.ContentList != null && report.ContentList.Count != 0))
            {
                sb.AppendLine(report.ReportFileName);
                string headLine = report.HeaderList.Aggregate(string.Empty, (current, header) => current + string.Format(",{0}", header));
                sb.AppendLine(headLine.Substring(1));

                foreach (string contentLine in report.ContentList.Select(contents => contents.Aggregate(string.Empty, (current, contentField) => current + string.Format(",{0}", contentField))))
                {
                    sb.AppendLine(contentLine.Substring(1));
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private string GetNewRicReportContent()
        {
            var sb = new StringBuilder();
            foreach (string key in sendNewRicReport.Keys)
            {
                sb.AppendLine(key);
                foreach (string content in sendNewRicReport[key])
                {
                    sb.AppendLine(content);
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private void SendAlertMail(ExchangeService service, Dictionary<String, String> condition)
        {
            var attacheFileList = new List<string>();
            string content = "";
            if (alertTag == 1)
            {
                content += String.Format("\t\t<p>{0}</p>\r\n", "Alert: can't find RCC Reports files in the RCC Share drive path!");
            }
            foreach (string key in condition.Keys)
            {
                string attachedFile = Path.Combine(configObj.ResultFileDir, String.Format("{0}_Result_{1}.csv", key.ToUpper(), DateTime.Now.ToString("ddMMMHHmm")));
                File.WriteAllText(attachedFile, condition[key]);
                AddResult(String.Format("{0} result file", key), attachedFile, "file");
                //TaskResultList.Add(new TaskResultEntry(string.Format("{0} result file", key), "", attachedFile));
                attacheFileList.Add(attachedFile);
                if (!string.IsNullOrEmpty(condition[key]))
                {
                    content += string.Format("\t\t<p>{0}</p>\r\n", key + " market exists error!");
                }
            }
            if (newRicReportContent != null)
            {
                string newRicListFile = Path.Combine(configObj.ResultFileDir, string.Format("{0}_Result_{1}.csv", "NewRicList", DateTime.Now.ToString("ddMMMHHmm")));
                string contents = newRicReportContent.Replace(",", ";");
                File.WriteAllText(newRicListFile, contents);
                AddResult("New Ric list file", newRicListFile, "file");
                //TaskResultList.Add(new TaskResultEntry("New Ric List File", "Download All New Ric From Mails", newRicListFile));
                attacheFileList.Add(newRicListFile);
            }
            else
            {
                content += string.Format("\t\t<p>{0}</p>\r\n", "No New Ric in Mails!");
            }
            string subject = string.Format("E2E QC Check Result for {0}", DateTime.Now.ToString("ddMMM_HH:mm"));
            if (content == "")
            {
                content += string.Format("\t\t<p>{0}</p>", "All Market is Right!");
            }
            content = string.Format("<html>\r\n\t<body>\r\n{0}\r\n</body>\r\n</html>", content);
            try
            {
                EWSUtility.CreateAndSendMail(service, configObj.AlertMailToList, configObj.AlertMailCcList, new List<string>(), subject, content, attacheFileList);
            }
            catch (Exception ex)
            {
                LogMessage("send alert mail failed. ex: " + ex.Message, Logger.LogType.Warning);
            }
        }
        #endregion

        public static Dictionary<String, DateTime> GetLastTradingDay(DateTime originDate, HolidayList holidayList, int deltaDay)
        {
            var dictionary = new Dictionary<String, DateTime>();

            for (int i = 0; i < holidayList.MarketHolidayList.Count; i++)
            {
                string name = holidayList.MarketHolidayList[i].Name;
                var strHolidayList = from holiday in holidayList.MarketHolidayList[i].Holidays select holiday.ToString("yyyy-MM-dd");
                var holidaySet = new HashSet<string>(strHolidayList);

                DateTime curDay = originDate;
                int dayLeft = deltaDay;

                while (true)
                {
                    if (curDay.DayOfWeek == DayOfWeek.Saturday || curDay.DayOfWeek == DayOfWeek.Sunday ||
                        holidaySet.Contains(curDay.ToString("yyyy-MM-dd")))
                    {
                        curDay = curDay.AddDays(-1);
                    }
                    else
                    {
                        curDay = curDay.AddDays(-1);
                        dayLeft--;
                    }

                    if (dayLeft == 0)
                    {
                        while (true)
                        {
                            if (curDay.DayOfWeek == DayOfWeek.Saturday || curDay.DayOfWeek == DayOfWeek.Sunday ||
                                holidaySet.Contains(curDay.ToString("yyyy-MM-dd")))
                            {
                                curDay = curDay.AddDays(-1);
                            }
                            else
                            {
                                break;
                            }
                        }
                        break;
                    }
                }
                dictionary.Add(name, curDay);
            }
            return dictionary;
        }
    }
}
