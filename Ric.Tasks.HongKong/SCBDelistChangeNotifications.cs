using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Microsoft.Office.Interop.Excel;
using System.IO;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    public class SCBDelistChangeNotificationConfig
    {
        public string GENERATED_SCB_FILE_PATH { get; set; }
        public string SOURCE_FM_FILE_PATH { get; set; }
        public string LOG_FILE_PATH { get; set; }
    }

    public class SCBChangeInfo
    {
        public string EventType { get; set; }
        public string Function { get; set; }
        public string FileCode { get; set; }
        public string Isin { get; set; }
        public string Sedol { get; set; }
        public string OldTicker { get; set; }
        public string NewTicker { get; set; }
        public string OldRic { get; set; }
        public string NewRic { get; set; }
        public string Mic { get; set; }
        public string ExchangeCode { get; set; }
        public string ExchangeName { get; set; }
        public string OldCompName { get; set; }
        public string NewCompName { get; set; }
        public string AnnouncementDate { get; set; }
        public string EffectiveDate { get; set; }
        public SCBChangeInfo()
        {
            EventType = "CHAN";
            Function = "NEW";
            FileCode = "1";
            NewTicker = "N.A.";
            NewRic = "N.A.";
            Mic = "XHKG";
            ExchangeCode = "HKG";
            ExchangeName = "The Stock Exchange of Hong Kong";
            Isin = string.Empty;
            Sedol = string.Empty;
        }

    }
    public class SCBDelistChangeNotification : GeneratorBase
    {
        private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_SCBDelistChangeNotification.config";
        private static SCBDelistChangeNotificationConfig configObj = null;
        //private static Logger logger = null;


        protected override void Start()
        {
            StartSCBDelistChangeNotificationJob();
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(SCBDelistChangeNotificationConfig)) as SCBDelistChangeNotificationConfig;
            //logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.New);
        }

        public void StartSCBDelistChangeNotificationJob()
        {
            SCBChangeInfo changeInfo = GetSCBChangeInfo();
            GenerateSCBFile(changeInfo);
        }

        public SCBChangeInfo GetSCBChangeInfo()
        {
            SCBChangeInfo changeInfo = new SCBChangeInfo();
            using (ExcelApp app = new ExcelApp(false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.SOURCE_FM_FILE_PATH);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                int i = lastUsedRow;
                while ((i <= lastUsedRow) && (i > 0))
                {
                    if (ExcelUtil.GetRange(i, 1, worksheet).Text != null)
                    {
                        string sA1 = ExcelUtil.GetRange(i, 1, worksheet).Text.ToString();
                        if (ExcelUtil.GetRange(i, 2, worksheet).Text == null)
                        {
                            LogMessage("The second column is null.");
                            //return changeInfo;
                        }
                        string sA2 = ExcelUtil.GetRange(i, 2, worksheet).Text.ToString();
                        switch (sA1)
                        {
                            case "Exchange Symbol:":
                                changeInfo.OldTicker = sA2;
                                if (sA2 == string.Empty)
                                    Logger.Log("Can't find the value for Old Ticker.");
                                break;
                            case "Underlying RIC:":
                                changeInfo.OldRic = sA2;
                                if (sA2 == string.Empty)
                                    Logger.Log("Can't find the value for Old Ric.");
                                break;
                            case "Organisation Name (DIRNAME) (OLD):":
                                changeInfo.OldCompName = sA2;
                                if (sA2 == string.Empty)
                                    Logger.Log("Can't find the value for Old Company Name.");
                                break;
                            case "Organisation Name (DIRNAME) (NEW):":
                                changeInfo.NewCompName = sA2;
                                if (sA2 == string.Empty)
                                    Logger.Log("Can't find the value for New Company Name.");
                                break;
                            case "Effective Date:":
                                changeInfo.EffectiveDate = sA2;
                                if (sA2 == string.Empty)
                                    Logger.Log("Can't find the value for Effective Date.");
                                break;
                            default:
                                break;
                        }
                    }
                    i--;
                }
                changeInfo.AnnouncementDate = GetAnnouncementDateFromFile(configObj.SOURCE_FM_FILE_PATH);
                if (changeInfo.AnnouncementDate == string.Empty)
                {
                    Logger.Log("Announcement date got from file name is invalid. Now it's empty.");
                }
                workbook.Close(false, workbook.FullName, false);
            }
            return changeInfo;
        }

        public void GenerateSCBFile(SCBChangeInfo changeInfo)
        {
            using (ExcelApp app = new ExcelApp(false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.GENERATED_SCB_FILE_PATH);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                int lastUsedCol = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
                ExcelUtil.GetRange(1, 1, lastUsedRow, lastUsedCol, worksheet).ClearContents();
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    writer.WriteLine("EventType");
                    writer.WriteLine("Function");
                    writer.WriteLine("FILE_CODE");
                    writer.WriteLine("ISIN");
                    writer.WriteLine("SEDOL");
                    writer.WriteLine("OLD TICKER");
                    writer.WriteLine("NEW TICKER");
                    writer.WriteLine("Old RIC");
                    writer.WriteLine("New RIC");
                    writer.WriteLine("MIC");
                    writer.WriteLine("EXCHANGE_CODE");
                    writer.WriteLine("EXCHANGE_NAME");
                    writer.WriteLine("OLD COMPANY_NAME");
                    writer.WriteLine("NEW COMPANY_NAME");
                    writer.WriteLine("ANNOUNCEMENT DATE");
                    writer.WriteLine("EFFECTIVE DATE");
                    writer.PlaceNext(2, 1);
                    writer.WriteLine(changeInfo.EventType);
                    writer.WriteLine(changeInfo.Function);
                    writer.WriteLine(changeInfo.FileCode);
                    writer.WriteLine(changeInfo.Isin);
                    writer.WriteLine(changeInfo.Sedol);
                    writer.WriteLine(changeInfo.OldTicker);
                    writer.WriteLine(changeInfo.NewTicker);
                    writer.WriteLine(changeInfo.OldRic);
                    writer.WriteLine(changeInfo.NewRic);
                    writer.WriteLine(changeInfo.Mic);
                    writer.WriteLine(changeInfo.ExchangeCode);
                    writer.WriteLine(changeInfo.ExchangeName);
                    writer.WriteLine(changeInfo.OldCompName);
                    writer.WriteLine(changeInfo.NewCompName);
                    writer.WriteLine(changeInfo.AnnouncementDate);
                    writer.WriteLine(changeInfo.EffectiveDate);
                }
                workbook.SaveCopyAs(GetSCBGenerateFilePath(configObj.GENERATED_SCB_FILE_PATH));
                workbook.Close(false, workbook.FullName, false);
            }
        }

        public string GetAnnouncementDateFromFile(string filePath)
        {
            //02NOV2011
            //2-Nov-11
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string temp = fileName.Split('_')[3];
            string date = int.Parse(temp.Substring(0, 2)).ToString();
            date += "-";
            date += temp[2].ToString().ToUpper();
            date += temp.Substring(3, 2).ToLower();
            date += "-";
            date += temp.Substring(7, 2);
            return date;
        }

        public string GetSCBGenerateFilePath(string originalFilePath)
        {
            string dir = Path.GetDirectoryName(originalFilePath);
            string fileName = "Reuters_CHAN_DLST_";
            string currentDate = DateTime.Now.ToString("yyyyMMdd");
            fileName += currentDate;
            fileName += "_100000";
            fileName += ".CSV";
            string filePath = Path.Combine(dir, fileName);
            return filePath;
        }
    }
}
