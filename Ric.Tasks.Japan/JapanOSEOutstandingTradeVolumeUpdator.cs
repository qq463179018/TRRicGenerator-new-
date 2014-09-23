using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.ComponentModel;
using System.Drawing.Design;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using HtmlAgilityPack;
using Ric.Core;
using Ric.Util;
using Ric.Db.Manager;

namespace Ric.Tasks.Japan
{
    public class OSETradeCompanyInfo
    {
        //public string TradeType { get; set; } //Can be "BUY" or "SELL"
        public string OriginalName { get; set; }
        public string EnglishName { get; set; }
        public string JapaneseName { get; set; }
        public string OPINT { get; set; }
    }

    [ConfigStoredInDB]
    public class JapanOSETradeVolumeUpdatorConfig
    {
        [Category("Announcement Date")]
        [Description("Date format: MMdd. E.g. 1206")]
        public string Date { get; set; }

        [StoreInDB]
        [Category("Down File Path")]
        [DefaultValue("C:\\Japan\\source\\")]
        [Description("File folder where the files will be downloaded")]
        public string DownloadFilePath { get; set; }

        [StoreInDB]
        [Category("Template File Path")]
        public string OS225FUUTemplateFile { get; set; }

        [StoreInDB]
        [Category("Template File Path")]
        public string OS225OPZTemplateFile { get; set; }

        [StoreInDB]
        [Category("Template File Path")]
        public string OSOPT35_46TemplateFile { get; set; }
        //[Description("User name of the OSE account")]
        //public string USERNAME { get; set; }

        //[Description("The password of the OSE account ")]
        //public string PASSWORD { get; set; }
        [StoreInDB]
        [Description("Seq ID when getting files from website")]
        public string SEQ_ID { get; set; }

        [StoreInDB]
        [Description("The printer with which to print the files, if set the value to be empty, files wouldn't be printed.")]
        public string PRINTER_NAME { get; set; }

        [StoreInDB]
        [Description("File path of baknote file")]
        public string BAKNOTE_FILE_PATH { get; set; }

        [StoreInDB]
        [Description("Excel Version, by now, it can be \"03\", \"07\" ")]
        public string TARGET_FILE_VERSION { get; set; }

        [StoreInDB]
        [DefaultValue("Sheet1")]
        [Description("Worksheet name of the baknote file")]
        public string WORKSHEETNAME_BAKNOTE { get; set; }

        [StoreInDB]
        [Description("The path of the file which contains the holiday information of Japan market")]
        public string HOLIDAY_INFO_FILE_PATH { get; set; }



        [StoreInDB]
        [Description("Target file path")]
        public string TARGET_FILE_DIR { get; set; }

        private Dictionary<string, int> businessDay = new Dictionary<string, int>();

        private void InitialBusinessDay()
        {
            businessDay.Add("Monday",-3);
            businessDay.Add("Tuesday", -4);
            businessDay.Add("Wednesday", -5);
            businessDay.Add("Thursday", -6);
            businessDay.Add("Friday", -7);
            businessDay.Add("Saturday", -8);
            businessDay.Add("Sunday", -9);
        }

        private bool IsHoliday(List<DateTime> holidays,DateTime date)
        {
            if (holidays.Contains(date))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public JapanOSETradeVolumeUpdatorConfig()
        {
            InitialBusinessDay();
            List<DateTime> holidays = HolidayManager.SelectHoliday(5);
            DateTime date = DateTime.Today;
            date = date.AddDays(businessDay[date.DayOfWeek.ToString()]);
            while (IsHoliday(holidays,date))
            {
                date.AddDays(-1);
            }
            Date = date.ToString("yyyyMMdd");
        }
                       
    }

    public class tradingInfo
    {
        public DateTime tradingDate1 { get; set; }
        public DateTime tradingDate2 { get; set; }
        public List<OSETradeCompanyInfo> info1 = new List<OSETradeCompanyInfo>();
        public List<OSETradeCompanyInfo> info2 = new List<OSETradeCompanyInfo>();

        public tradingInfo()
        {
            tradingDate1 = new DateTime();
            tradingDate2 = new DateTime();
            //List<OSETradeCompanyInfo> info1 = new List<OSETradeCompanyInfo>();
            //List<OSETradeCompanyInfo> info2 = new List<OSETradeCompanyInfo>();
        }
    }

    public class OS225FUUOriginData
    {
        public DateTime updateDate { get; set; }
        public tradingInfo OS225FUU { get; set; }
        public tradingInfo OS225FUX { get; set; }
        public tradingInfo OSTRADF { get; set; }

        public OS225FUUOriginData()
        {
            updateDate = new DateTime();
            OS225FUU = new tradingInfo();
            OS225FUX = new tradingInfo();
            OSTRADF = new tradingInfo();
        }
    }

    public class JapanOSEOutstandingTradeVolumeUpdator :GeneratorBase
    {
        private static JapanOSETradeVolumeUpdatorConfig configObj = null;
        Dictionary<string, NameMap> nameDic = null;
        private ExcelApp app;
        protected override void Start()
        {
            StartOSETradeVolumeUpdatorJob();
            app.Dispose();
        }

        protected override void Initialize()
        {
            base.Initialize();
            app = new ExcelApp(false, false);
            if (app.ExcelAppInstance == null)
            {
                Logger.Log("Excel cannot be started", Logger.LogType.Error);
            }
            configObj = Config as JapanOSETradeVolumeUpdatorConfig;
            nameDic = JapanShared.GetNameMap(configObj.BAKNOTE_FILE_PATH, configObj.WORKSHEETNAME_BAKNOTE);
        }

        //Start OSE Trade Volume Updator jobs according to the tasks configured in the configuration file
        private void StartOSETradeVolumeUpdatorJob()
        {
            downloadSourceFiles();
            OS225FUUJob();
            OS225OPZJob();
            OSOPT35_46Job();
        }

        private void downloadSourceFiles()
        {
            string fileUrl = null;
            try
            {
                HtmlDocument htc = new HtmlDocument();
                string uri = String.Format("http://www.ose.or.jp/market/trading_data/open_interest_by_participant");
                htc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                HtmlNodeCollection tr = htc.DocumentNode.SelectNodes("//tr/td/a");
                foreach (HtmlNode node in tr)
                {
                    if (node.Attributes["href"].Value.Contains(configObj.Date + "_index_futures_OI_by_participant.xls"))
                    {
                        fileUrl = node.Attributes["href"].Value;
                        WebClientUtil.DownloadFile(@"http://www.ose.or.jp/" + fileUrl, 18000, configObj.DownloadFilePath + "\\" + configObj.Date + "_index_futures_OI_by_participant.xls");
                    }
                    else if (node.Attributes["href"].Value.Contains(configObj.Date + "_index_options_OI_by_participant.xls"))
                    {
                        fileUrl = node.Attributes["href"].Value;
                        WebClientUtil.DownloadFile(@"http://www.ose.or.jp/" + fileUrl, 18000, configObj.DownloadFilePath + "\\" + configObj.Date + "_index_options_OI_by_participant.xls");
                    }
                    else if (node.Attributes["href"].Value.Contains(configObj.Date + "_individual_options_OI_by_participant.xls"))
                    {
                        fileUrl = node.Attributes["href"].Value;
                        WebClientUtil.DownloadFile(@"http://www.ose.or.jp/" + fileUrl, 18000, configObj.DownloadFilePath + "\\" + configObj.Date + "_individual_options_OI_by_participant.xls");
                    }

                    
                }
                
            }

            catch (Exception ex)
            {
                Logger.Log("Error when downloading files from website" + ex.Message);
                LogMessage("Error when downloading files from website");
            }

            finally
            {
              
            }
        }

        private Workbook InitialExcelFile(string filePath)
        {
            Workbook book = null;
            try
            {
                book = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
            }
            catch (Exception ex)
            {
                Logger.Log("Error when initial excel file" + ex.Message);
                LogMessage("Error when initial excel file");
                return book;
            }
            return book;
            
        }

        private void updateCompanyInfo(OSETradeCompanyInfo companyInfo, Dictionary<string, NameMap> nameDic)
        {
            foreach (var item in nameDic)
            {
                if (item.Key == companyInfo.OriginalName.Trim())
                {
                    companyInfo.JapaneseName = item.Value.JapaneseName;
                    companyInfo.EnglishName = item.Value.EnglistName;
                    break;
                }
            }
        }

        private void GetOS225FUUData(Workbook book, ref OS225FUUOriginData data)
        {
            if (book == null || data == null)
            {
                return;
            }
            try
            {
                Worksheet sheet = book.Worksheets[1] as Worksheet;
                if (sheet == null)
                {
                    return;
                }
                string updateDate = ExcelUtil.GetRange("A2", sheet).Text.ToString();
                data.updateDate = JapanShared.TransferJpDate(updateDate);

                int lastUsedRow = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1;
                int currentRow = 1;
                while (currentRow <= lastUsedRow)
                {
                    if (ExcelUtil.GetRange(currentRow, 1, sheet).Value2 != null && ExcelUtil.GetRange(currentRow, 1, sheet).Value2.ToString() != string.Empty)
                    {
                        string text = ExcelUtil.GetRange(currentRow, 1, sheet).Value2.ToString();
                        if (text.Contains("日経225先物"))
                        {
                            string date = ExcelUtil.GetRange(currentRow + 1, 2, sheet).Text.ToString();
                            data.OS225FUU.tradingDate1 = JapanShared.TransferJpDate(date);

                            date = ExcelUtil.GetRange(currentRow + 1, 9, sheet).Text.ToString();
                            data.OS225FUU.tradingDate2 = JapanShared.TransferJpDate(date);

                            for (int i = 0; i < 15; i++)
                            {
                                OSETradeCompanyInfo companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 2);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OS225FUU.info1.Add(companyInfo);

                                companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 4);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OS225FUU.info1.Add(companyInfo);

                                companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 9);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OS225FUU.info2.Add(companyInfo);

                                companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 11);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OS225FUU.info2.Add(companyInfo);
                            }
                            currentRow += 20;
                        }
                        else if (text.Contains("日経225mini"))
                        {
                            string date = ExcelUtil.GetRange(currentRow + 1, 2, sheet).Text.ToString();
                            data.OS225FUX.tradingDate1 = JapanShared.TransferJpDate(date);

                            date = ExcelUtil.GetRange(currentRow + 1, 9, sheet).Text.ToString();
                            data.OS225FUX.tradingDate2 = JapanShared.TransferJpDate(date);

                            for (int i = 0; i < 15; i++)
                            {
                                OSETradeCompanyInfo companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 2);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OS225FUX.info1.Add(companyInfo);

                                companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 4);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OS225FUX.info1.Add(companyInfo);

                                companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 9);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OS225FUX.info2.Add(companyInfo);

                                companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 11);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OS225FUX.info2.Add(companyInfo);
                            }
                            currentRow += 20;
                        }

                        else if (text.Contains("TOPIX先物"))
                        {
                            string date = ExcelUtil.GetRange(currentRow + 1, 2, sheet).Text.ToString();
                            data.OSTRADF.tradingDate1 = JapanShared.TransferJpDate(date);

                            date = ExcelUtil.GetRange(currentRow + 1, 9, sheet).Text.ToString();
                            data.OSTRADF.tradingDate2 = JapanShared.TransferJpDate(date);

                            for (int i = 0; i < 15; i++)
                            {
                                OSETradeCompanyInfo companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 2);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OSTRADF.info1.Add(companyInfo);

                                companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 4);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OSTRADF.info1.Add(companyInfo);

                                companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 9);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OSTRADF.info2.Add(companyInfo);

                                companyInfo = OS225FUU.GetTradingInfo(sheet, currentRow + 4 + i, 11);
                                updateCompanyInfo(companyInfo, nameDic);
                                data.OSTRADF.info2.Add(companyInfo);
                            }
                            currentRow += 20;
                        }
                        else
                        {
                            currentRow++;
                        }
                    }
                    else
                    {
                        currentRow++;
                    }
                }

            }
            catch (Exception ex)
            {
                Logger.Log("GetOS225FUUData failed" + ex.Message);
                LogMessage("GetOS225FUUData failed");
            }
        }

        private void GenerateOS225FUU(Workbook book, OS225FUUOriginData data)
        {
            OS225FUU.WriteOS225FUU(book,data);
            OS225FUU.WriteOS225FUX(book, data);
            OS225FUU.WriteOSTRADF(book, data);
        }

        private void OS225FUUJob()
        {
            string sourceFile = configObj.DownloadFilePath + "\\" + configObj.Date + "_index_futures_OI_by_participant.xls";
            string targetFile = configObj.OS225FUUTemplateFile;
            OS225FUUOriginData  data = new OS225FUUOriginData();
            Workbook source = InitialExcelFile(sourceFile);
            Workbook target = InitialExcelFile(targetFile);

            if (source == null || target == null)
            {
                return;
            }

            try
            {
                GetOS225FUUData(source, ref data);
                GenerateOS225FUU(target,data);
                target.Save();
            }
            catch (Exception ex)
            {
                Logger.Log("OS225FUUJob failed" + ex.Message);
                LogMessage("OS225FUUJob failed");
            }

        }

        private void OS225OPZJob()
        {
            string sourceFile = configObj.DownloadFilePath + "\\" + configObj.Date + "_index_options_OI_by_participant.xls";
            string targetFile = configObj.OS225OPZTemplateFile;
            OS225OPZOriginData data = new OS225OPZOriginData();
            Workbook source = InitialExcelFile(sourceFile);
            Workbook target = InitialExcelFile(targetFile);

            if (source == null || target == null)
            {
                return;
            }

            try
            {
                OS225OPZ.GetOS225OPZData(source, ref data);

                foreach (OSETradeCompanyInfo item in data.putInfo)
                {
                    updateCompanyInfo(item, nameDic);
                }

                foreach (OSETradeCompanyInfo item in data.callInfo)
                {
                    updateCompanyInfo(item, nameDic);
                }

                OS225OPZ.GenerateOS225OPZ(target, data);
                target.Save();
            }
            catch (Exception ex)
            {
                Logger.Log("OS225OPZJob failed" + ex.Message);
                LogMessage("OS225OPZJob failed");
            }
        }

        private void OSOPT35_46Job()
        {
            string sourceFile = configObj.DownloadFilePath + "\\" + configObj.Date + "_individual_options_OI_by_participant.xls";
            string targetFile = configObj.OSOPT35_46TemplateFile;
            OSOPT35_46OriginData data = new OSOPT35_46OriginData();
            Workbook source = InitialExcelFile(sourceFile);
            Workbook target = InitialExcelFile(targetFile);

            if (source == null || target == null)
            {
                return;
            }

            try
            {
                OSOPT35_46.GetOSOPT35_46Data(source,ref data,nameDic);
                OSOPT35_46.GenerateOSOPT35_46(target, data);
                target.Save();
            }
            catch (Exception ex)
            {
                Logger.Log("OSOPT35_46Job failed" + ex.Message);
                LogMessage("OSOPT35_46Job failed");
            }
        }
    }
}
