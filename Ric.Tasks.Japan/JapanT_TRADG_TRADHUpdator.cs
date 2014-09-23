using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks.Japan
{
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class JapanTRADG_TRADHUpdatorConfig
    {
        [StoreInDB]
        [DisplayName("Source file directory")]
        [Description("Path of the folder under which the source file will be saved")]
        public string SourceFileDir { get; set; }

        [StoreInDB]
        [DisplayName("Target file directory")]
        [Description("Path of the folder under which the target file will be ")]
        public string TargetFileDir { get; set; }

        [StoreInDB]
        [DisplayName("Target file version")]
        [Description("Excel Version, by now, it can be \"03\", \"07\" ")]
        public string TargetFileVersion { get; set; }

        [StoreInDB]
        [DisplayName("Bak note File path")]
        [Description("Path of the bakNode excel file")]
        public string BaknoteFilePath { get; set; }

        [StoreInDB]
        [DisplayName("Printer name")]
        [Description("Printer name: \"\\BJZITRENDOS1.apac.ime.reuters.com\\BJS01PTR012\"")]
        public string PrinterName { get; set; }

        [StoreInDB]
        [DisplayName("Worksheet name")]
        [Description("Worksheet name of the bakNote file")]
        public string WorkSheetName { get; set; }

        [StoreInDB]
        [DisplayName("Special case")]
        [Description("Special case, like: yes ")]
        public string SpecialCase { get; set; }
    }

    public class Transaction
    {
        public CompanyInfo CompanyInfo { get; set; }
        public string SellSum { get; set; }
        public string BuySum { get; set; }
        public string TotalSum { get; set; }
        public Transaction()
        {
            CompanyInfo = new CompanyInfo();
        }
    }

    public class JpTRADGInfo
    {
        public string Date { get; set; }
        public string GSellSum { get; set; }
        public string GBuySum { get; set; }
        public string GTotalSum { get; set; }
        public string Top15SellSum { get; set; }
        public string Top15BuySum { get; set; }
        public string Top15TotalSum { get; set; }
        public string CurrentSellSum { get; set; }
        public string CurrentBuySum { get; set; }
        public string CurrentNChg1 { get; set; }
        public string CurrentNChg2 { get; set; }
        public string NextSellSum { get; set; }
        public string NextBuySum { get; set; }
        public string NextNChg1 { get; set; }
        public string NextNChg2 { get; set; }
        public string TotalSellSum { get; set; }
        public string TotalBuySum { get; set; }
        public string TotalNchg1 { get; set; }
        public string TotalNchg2 { get; set; }
        public List<Transaction> TransactionList { get; set; }
        public JpTRADGInfo()
        {
            this.TransactionList = new List<Transaction>();
        }
    }

    public class JpTRADHInfo
    {
        public JpTRADHInfoPart1 Part1 { get; set; }
        public JpTRADHInfoPart2 Part2 { get; set; }
        public JpTRADHInfoPart3 Part3 { get; set; }
        public JpTRADHInfo()
        {
            Part1 = new JpTRADHInfoPart1();
            Part2 = new JpTRADHInfoPart2();
            Part3 = new JpTRADHInfoPart3();
        }
    }

    public class JpTRADHInfoPart1
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Top15SellSum { get; set; }
        public string Top15BuySum { get; set; }
        public string Top15TotalSum { get; set; }
        public string GTotalSellSum { get; set; }
        public string GTotalBuySum { get; set; }
        public string GTotalTotalSum { get; set; }
        public List<Transaction> TransactionList { get; set; }
        public JpTRADHInfoPart1()
        {
            TransactionList = new List<Transaction>();
        }
    }

    public class JpTRADHInfoPart2
    {
        public string Title { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public List<string> ShareSumList { get; set; }
        public List<string> TradeVolList { get; set; }
        public JpTRADHInfoPart2()
        {
            ShareSumList = new List<string>();
            TradeVolList = new List<string>();
        }
    }

    public class JpTRADHInfoPart3
    {
        public string Title { get; set; }
        public DateTime Date { get; set; }
        public List<string> ShareSumList { get; set; }
        public List<string> ShareIncrementList { get; set; }
        public List<string> TradeVolSumList { get; set; }
        public List<string> TradeVolIncrementList { get; set; }
        public JpTRADHInfoPart3()
        {
            ShareSumList = new List<string>();
            ShareIncrementList = new List<string>();
            TradeVolIncrementList = new List<string>();
            TradeVolSumList = new List<string>();
        }
    }

    public class JapanT_TRADG_TRADHUpdator : GeneratorBase
    {
        private string sourceFileDir;
        private string targretFileDir;
        private string targetFileVersion;
        private string baknoteFilePath;
        private string printerName;
        private string workSheetName;
        private string specialCase;
        private Dictionary<string, NameMap> nameDic;
        private JapanTRADG_TRADHUpdatorConfig configObj;
        private List<DateTime> holidayList = new List<DateTime>();
        private string dailySourceFilePath;
        private string weeklySourceFilePath;
        private string tse_url = @"http://www.tse.or.jp/market/data/program/index.html";
        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as JapanTRADG_TRADHUpdatorConfig;
            sourceFileDir = configObj.SourceFileDir.Trim();
            targretFileDir = configObj.TargetFileDir.Trim();
            targetFileVersion = configObj.TargetFileVersion.Trim();
            baknoteFilePath = configObj.BaknoteFilePath.Trim();
            printerName = configObj.PrinterName.Trim();
            workSheetName = configObj.WorkSheetName.Trim();
            specialCase = configObj.SpecialCase.Trim();
            InitializeDir(sourceFileDir);
        }
        protected override void Cleanup()
        {
        }
        // Make sure no dupl
        private void InitializeDir(string dir)
        {
            string[] files = Directory.GetFiles(dir, "*.xls", SearchOption.TopDirectoryOnly);
            if (files.Length != 0)
            {
                foreach (string file in files)
                {
                    File.Delete(file);
                }
            }
        }
        protected override void Start()
        {
            startTradgTradhUpdating();
        }
        private void startTradgTradhUpdating()
        {
            List<string> linkList = GetSourceFileLinks(tse_url);
            nameDic = JapanShared.GetNameMap(baknoteFilePath, workSheetName);
            // For tradgUpdate
            try
            {
                dailySourceFilePath = DownloadFile(linkList[0]);
                JpTRADGInfo tradgTradingInfo = printAndGetJpDailyTradingInfo(dailySourceFilePath);
                GenerateTTradgTargetFile(targretFileDir, tradgTradingInfo);
            }
            catch (Exception ex)
            {
                Logger.Log("There's error during task: TRADG. Exception: " + ex.Message);
            }
            // For TRADH
            try
            {
                if (linkList.Count == 2)
                {
                    weeklySourceFilePath = DownloadFile(linkList[1]);
                    JpTRADHInfo tradhTradeInfo = PrintAndGetJpWeeklyTradingInfo(weeklySourceFilePath);
                    GenerateTradhTargetFile(targretFileDir, tradhTradeInfo);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("There's error during task: TRADH. Exception: " + ex.Message);
            }
        }

        #region TRADH Related
        private void GenerateTradhTargetFile(string targetFileDir, JpTRADHInfo tradeInfo)
        {
            string fileName = "Weekly_TTRADH";
            fileName += MiscUtil.getFileExtension(targetFileVersion);
            if (!Directory.Exists(targetFileDir))
            {
                Directory.CreateDirectory(targetFileDir);
            }
            fileName = Path.Combine(targetFileDir, fileName);
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, fileName);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                ((Range)worksheet.Columns["A:I", Missing.Value]).ColumnWidth = 15;
                ((Range)worksheet.Columns["A:I", Type.Missing]).Font.Name = "明朝";
                worksheet.Cells[1, 1] = "ARBITRAGE TRANSACTION (T/TRADH)";
                //Part 1:
                worksheet.Cells[2, 6] = string.Format("{0}-{1}", tradeInfo.Part1.StartDate.ToString("MMMdd", new CultureInfo("en-US")), tradeInfo.Part1.EndDate.ToString("MMMdd", new CultureInfo("en-US"))).ToUpper();
                ExcelUtil.GetRange(2, 6, worksheet).Interior.Color = ColorTranslator.ToOle(Color.Violet);
                worksheet.Cells[3, 2] = "SEUCURITIES";
                worksheet.Cells[3, 3] = "ENGLISH";
                worksheet.Cells[3, 4] = "JAPANESE";
                worksheet.Cells[3, 5] = "SELL";
                worksheet.Cells[3, 6] = "BUY";
                worksheet.Cells[3, 7] = "TOTAL";
                ExcelUtil.GetRange(3, 2, 3, 7, worksheet).Font.Bold = true;
                WriteTransactionInfo(tradeInfo.Part1.TransactionList, worksheet, 4, 2);
                worksheet.Cells[19, 2] = "TOP15";
                worksheet.Cells[20, 2] = "G.TOTAL";
                ExcelUtil.GetRange(19, 2, 20, 2, worksheet).Font.Bold = true;
                worksheet.Cells[19, 5] = tradeInfo.Part1.Top15SellSum;
                worksheet.Cells[19, 6] = tradeInfo.Part1.Top15BuySum;
                worksheet.Cells[19, 7] = tradeInfo.Part1.Top15TotalSum;
                worksheet.Cells[20, 5] = tradeInfo.Part1.GTotalSellSum;
                worksheet.Cells[20, 6] = tradeInfo.Part1.GTotalBuySum;
                worksheet.Cells[20, 7] = tradeInfo.Part1.GTotalTotalSum;
                ExcelUtil.GetRange(4, 5, 20, 7, worksheet).HorizontalAlignment = XlHAlign.xlHAlignRight;
                ExcelUtil.GetRange(3, 1, 20, 7, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                //Part 2:
                worksheet.Cells[22, 1] = tradeInfo.Part2.Title;
                worksheet.Cells[22, 6] = string.Format("{0}-{1}", tradeInfo.Part2.StartDate.ToString("MMMdd", new CultureInfo("en-US")), tradeInfo.Part2.EndDate.ToString("MMMdd", new CultureInfo("en-US"))).ToUpper();
                ExcelUtil.GetRange(22, 6, worksheet).Interior.Color = ColorTranslator.ToOle(Color.Violet);
                worksheet.Cells[23, 2] = "ﾌﾟﾛｸﾞﾗﾑ売買";
                worksheet.Cells[24, 6] = "裁定取引に係る売買";
                worksheet.Cells[24, 8] = "裁定取引以外の売買";
                worksheet.Cells[25, 2] = "売り付け";
                worksheet.Cells[25, 5] = "買い付け";
                worksheet.Cells[25, 6] = "売り付け";
                worksheet.Cells[25, 7] = "買い付け";
                worksheet.Cells[25, 8] = "売り付け";
                worksheet.Cells[25, 9] = "買い付け";
                worksheet.Cells[26, 1] = "株数";
                worksheet.Cells[27, 1] = "金額";
                WriteShareOrIncreaseSumList(tradeInfo.Part2.ShareSumList, worksheet, 26, 2);
                WriteShareOrIncreaseSumList(tradeInfo.Part2.TradeVolList, worksheet, 27, 2);
                ExcelUtil.GetRange(26, 2, 27, 9, worksheet).HorizontalAlignment = XlHAlign.xlHAlignRight;
                ExcelUtil.GetRange(23, 1, 27, 9, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightSlateGray);
                //Part 3:
                worksheet.Cells[29, 1] = tradeInfo.Part3.Title;
                ExcelUtil.GetRange(29, 6, worksheet).NumberFormat = "@";
                worksheet.Cells[29, 6] = tradeInfo.Part3.Date.ToString("MMMdd", new CultureInfo("en-US")).ToUpper();
                ExcelUtil.GetRange(29, 6, worksheet).Interior.Color = ColorTranslator.ToOle(Color.Violet);
                worksheet.Cells[30, 2] = "'         売りﾎﾟｼｼｮﾝ";
                worksheet.Cells[30, 7] = "'       買ﾎﾟｼﾞｼｮﾝ";
                worksheet.Cells[31, 2] = "   当限";
                worksheet.Cells[31, 5] = "翌限以降";
                worksheet.Cells[31, 6] = "  合計  ";
                worksheet.Cells[31, 7] = "   当限";
                worksheet.Cells[31, 8] = "翌限以降";
                worksheet.Cells[31, 9] = "  合計  ";
                worksheet.Cells[32, 1] = "株数";
                worksheet.Cells[33, 1] = "前週末比";
                worksheet.Cells[34, 1] = "金額";
                worksheet.Cells[35, 1] = "前週末比";
                WriteShareOrIncreaseSumList(tradeInfo.Part3.ShareSumList, worksheet, 32, 2);
                WriteShareOrIncreaseSumList(tradeInfo.Part3.ShareIncrementList, worksheet, 33, 2);
                WriteShareOrIncreaseSumList(tradeInfo.Part3.TradeVolSumList, worksheet, 34, 2);
                WriteShareOrIncreaseSumList(tradeInfo.Part3.TradeVolIncrementList, worksheet, 35, 2);
                ExcelUtil.GetRange(32, 2, 35, 9, worksheet).HorizontalAlignment = XlHAlign.xlHAlignRight;
                ExcelUtil.GetRange(30, 1, 35, 9, worksheet).Interior.Color = ColorTranslator.ToOle(Color.Olive);
                worksheet.UsedRange.NumberFormat = "@";
                workbook.Save();
                workbook.Close(false, workbook.FullName, false);
                AddResult("Weekly_TRADH", fileName, "file");
                //TaskResultList.Add(new TaskResultEntry("Weekly_TRADH", "Weekly Task Result File", fileName));
            }
        }

        private void WriteShareOrIncreaseSumList(List<string> sumList, _Worksheet worksheet, int curRow, int curCol)
        {
            for (int i = 0; i < sumList.Count; i++)
            {
                if (i == 0)
                {
                    worksheet.Cells[curRow, curCol] = sumList[i];
                }
                else
                {
                    worksheet.Cells[curRow, curCol + i + 2] = sumList[i];
                }
            }
        }

        // Print Source File and Get Weekly TradingInfo
        private JpTRADHInfo PrintAndGetJpWeeklyTradingInfo(string filePath)
        {
            JpTRADHInfo tradhInfo = new JpTRADHInfo();
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                JapanShared.PrintWorksheet(worksheet, printerName, XlPageOrientation.xlPortrait);
                worksheet.UsedRange.UnMerge();
                JpTRADHInfoPart1 part1 = new JpTRADHInfoPart1();
                DateTime[] part1DateArr = ParseToGetTimeDuration(ExcelUtil.GetRange(33, 3, worksheet).Value2.ToString());
                part1.StartDate = part1DateArr[0];
                part1.EndDate = part1DateArr[1];
                part1.TransactionList = GetTransactionList(37, 3, worksheet);
                part1.Top15SellSum = ExcelUtil.GetRange(52, 7, worksheet).Value2.ToString();
                part1.Top15BuySum = ExcelUtil.GetRange(52, 9, worksheet).Value2.ToString();
                part1.Top15TotalSum = ExcelUtil.GetRange(52, 13, worksheet).Value2.ToString();
                part1.GTotalSellSum = ExcelUtil.GetRange(52, 7, worksheet).Value2.ToString();
                part1.GTotalBuySum = ExcelUtil.GetRange(52, 9, worksheet).Value2.ToString();
                part1.GTotalTotalSum = ExcelUtil.GetRange(52, 13, worksheet).Value2.ToString();
                JpTRADHInfoPart2 part2 = new JpTRADHInfoPart2();
                string part2SourceTitle = ExcelUtil.GetRange(5, 3, worksheet).Value2.ToString().Trim();
                part2.Title = part2SourceTitle.Split('（')[0].Trim();
                DateTime[] part2DateArr = ParseToGetTimeDuration(part2SourceTitle.Split('（')[1].Trim());
                part2.StartDate = part2DateArr[0];
                part2.EndDate = part2DateArr[1];
                for (int i = 5; i < 16; i++)
                {
                    if (ExcelUtil.GetRange(9, i, worksheet).Value2 != null && ExcelUtil.GetRange(9, i, worksheet).Value2.ToString().Trim() != string.Empty)
                    {
                        part2.ShareSumList.Add(ExcelUtil.GetRange(9, i, worksheet).Value2.ToString().Trim());
                    }
                }
                for (int i = 5; i < 16; i++)
                {
                    if (ExcelUtil.GetRange(10, i, worksheet).Value2 != null && ExcelUtil.GetRange(10, i, worksheet).Value2.ToString().Trim() != string.Empty)
                    {
                        part2.TradeVolList.Add(ExcelUtil.GetRange(10, i, worksheet).Value2.ToString().Trim());
                    }
                }
                JpTRADHInfoPart3 part3 = new JpTRADHInfoPart3();
                string part3SourceTitle = ExcelUtil.GetRange(12, 3, worksheet).Value2.ToString();
                part3.Title = part3SourceTitle.Split('（')[0].Trim();
                part3.Date = JapanShared.TransferJpDate(part3SourceTitle.Split('（')[1].Trim());
                for (int i = 5; i < 16; i++)
                {
                    if (ExcelUtil.GetRange(15, i, worksheet).Value2 != null && ExcelUtil.GetRange(15, i, worksheet).Value2.ToString().Trim() != string.Empty)
                    {
                        part3.ShareSumList.Add(ExcelUtil.GetRange(15, i, worksheet).Value2.ToString().Trim());
                    }
                }
                for (int i = 5; i < 16; i++)
                {
                    if (ExcelUtil.GetRange(16, i, worksheet).Value2 != null && ExcelUtil.GetRange(16, i, worksheet).Value2.ToString().Trim() != string.Empty)
                    {
                        part3.ShareIncrementList.Add(ExcelUtil.GetRange(16, i, worksheet).Value2.ToString().Trim());
                    }
                }
                for (int i = 5; i < 16; i++)
                {
                    if (ExcelUtil.GetRange(17, i, worksheet).Value2 != null && ExcelUtil.GetRange(17, i, worksheet).Value2.ToString().Trim() != string.Empty)
                    {
                        part3.TradeVolSumList.Add(ExcelUtil.GetRange(17, i, worksheet).Value2.ToString().Trim());
                    }
                }
                for (int i = 5; i < 16; i++)
                {
                    if (ExcelUtil.GetRange(18, i, worksheet).Value2 != null && ExcelUtil.GetRange(18, i, worksheet).Value2.ToString().Trim() != string.Empty)
                    {
                        part3.TradeVolIncrementList.Add(ExcelUtil.GetRange(18, i, worksheet).Value2.ToString().Trim());
                    }
                }
                tradhInfo.Part1 = part1;
                tradhInfo.Part2 = part2;
                tradhInfo.Part3 = part3;
            }
            return tradhInfo;
        }

        private DateTime[] ParseToGetTimeDuration(string sourceTimeDuration)
        {
            //平成24年3月5日～3月9日
            string[] timeDuration = sourceTimeDuration.Split('～');
            DateTime[] dateArr = { DateTime.Today, DateTime.Today };
            dateArr[0] = JapanShared.TransferJpDate(timeDuration[0]);
            dateArr[1] = JapanShared.TransferJpDate(timeDuration[1]);
            return dateArr;
        }

        #endregion
        //Generate TTRadgTarget File
        private void GenerateTTradgTargetFile(string targetFileDir, JpTRADGInfo tradingInfo)
        {
            string fileName = "TTRADG";
            fileName += MiscUtil.getFileExtension(targetFileVersion);
            if (!Directory.Exists(targetFileDir))
            {
                Directory.CreateDirectory(targetFileDir);
            }
            fileName = Path.Combine(targetFileDir, fileName);
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, fileName);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                ((Range)worksheet.Columns["A", Type.Missing]).ColumnWidth = 8.43;
                ((Range)worksheet.Columns["B", Type.Missing]).ColumnWidth = 11.57;
                ((Range)worksheet.Columns["C", Type.Missing]).ColumnWidth = 14.57;
                ((Range)worksheet.Columns["D", Type.Missing]).ColumnWidth = 23.43;
                ((Range)worksheet.Columns["E", Type.Missing]).ColumnWidth = 16.57;
                ((Range)worksheet.Columns["F", Type.Missing]).ColumnWidth = 11.29;
                ((Range)worksheet.Columns["G", Type.Missing]).ColumnWidth = 15.29;
                ((Range)worksheet.Columns["A:G", Type.Missing]).Font.Name = "明朝";
                worksheet.Cells[1, 1] = "ARBITRAGE TRANSACTION (T/TRADG)";
                Range a1Range = ExcelUtil.GetRange(1, 1, worksheet);
                a1Range.Font.Size = 16;
                worksheet.Cells[1, 6] = tradingInfo.Date;
                worksheet.Cells[1, 7] = "AS OF " + tradingInfo.Date.ToUpper();
                Range f1Range = ExcelUtil.GetRange(1, 7, worksheet);
                f1Range.Interior.Color = ColorTranslator.ToOle(Color.Violet);
                worksheet.Cells[2, 7] = "↑飛ばす前に２営業日前になっているかどうか確認";
                Range g2Range = ExcelUtil.GetRange(2, 7, worksheet);
                g2Range.Font.Color = ColorTranslator.ToOle(Color.Red);
                g2Range.Font.Size = 9;
                int currentRow = 4;
                for (int i = 0; i < 15; i++)
                {
                    worksheet.Cells[currentRow + i, 1] = (i + 1).ToString();
                }
                worksheet.Cells[3, 2] = "SEUCURITIES";
                worksheet.Cells[3, 3] = "ENGLISH";
                worksheet.Cells[3, 4] = "JAPANESE";
                worksheet.Cells[3, 5] = "SELL";
                worksheet.Cells[3, 6] = "BUY";
                worksheet.Cells[3, 7] = "TOTAL";
                Range row3Range = ExcelUtil.GetRange(3, 2, 3, 7, worksheet);
                row3Range.Font.Size = 10;
                row3Range.Font.Bold = true;
                row3Range.Borders.LineStyle = 1;
                row3Range.Borders.Weight = XlBorderWeight.xlMedium;
                currentRow = 4;
                foreach (Transaction tran in tradingInfo.TransactionList)
                {
                    worksheet.Cells[currentRow, 3] = tran.CompanyInfo.EnglishName;
                    worksheet.Cells[currentRow, 4] = tran.CompanyInfo.JapaneseName;
                    worksheet.Cells[currentRow, 5] = tran.SellSum;
                    worksheet.Cells[currentRow, 6] = tran.BuySum;
                    worksheet.Cells[currentRow, 7] = tran.TotalSum;
                    currentRow++;
                }
                Range F4g18Range = ExcelUtil.GetRange(4, 5, 18, 7, worksheet);
                F4g18Range.Font.Bold = true;
                F4g18Range.Font.Size = 12;
                ExcelUtil.GetRange(4, 2, 18, 2, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                ExcelUtil.GetRange(4, 3, 18, 3, worksheet).Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                ExcelUtil.GetRange(4, 4, 18, 6, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                ExcelUtil.GetRange(19, 2, 20, 4, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                ExcelUtil.GetRange(23, 2, 24, 7, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                Range b4g18Range = ExcelUtil.GetRange(4, 2, 18, 7, worksheet);
                b4g18Range.Borders.LineStyle = 1;
                b4g18Range.Borders.Weight = XlBorderWeight.xlMedium;
                worksheet.Cells[19, 2] = "TOP15";
                worksheet.Cells[19, 5] = tradingInfo.Top15SellSum;
                worksheet.Cells[19, 6] = tradingInfo.Top15BuySum;
                worksheet.Cells[19, 7] = tradingInfo.Top15TotalSum;
                worksheet.Cells[20, 2] = "G.TOTAL";
                worksheet.Cells[20, 5] = tradingInfo.GSellSum;
                worksheet.Cells[20, 6] = tradingInfo.GBuySum;
                worksheet.Cells[20, 7] = tradingInfo.GTotalSum;
                Range b19g20Range = ExcelUtil.GetRange(19, 2, 20, 7, worksheet);
                b19g20Range.Font.Size = 12;
                b19g20Range.Font.Bold = true;
                b19g20Range.Borders.LineStyle = 1;
                b19g20Range.Borders.Weight = XlBorderWeight.xlMedium;
                worksheet.Cells[22, 2] = "SELL";
                worksheet.Cells[22, 5] = "N.CHG";
                worksheet.Cells[22, 6] = "BUY";
                worksheet.Cells[22, 7] = "N.CHG";
                Range row22Range = ExcelUtil.GetRange(22, 1, 22, 7, worksheet);
                row22Range.Font.Bold = true;
                row22Range.Font.Size = 11;
                row22Range.Borders.LineStyle = 1;
                row22Range.Borders.Weight = XlBorderWeight.xlMedium;
                worksheet.Cells[21, 4] = "OK";
                worksheet.Cells[23, 1] = "CURRENT MTH";
                worksheet.Cells[23, 2] = tradingInfo.CurrentSellSum;
                worksheet.Cells[23, 5] = tradingInfo.CurrentNChg1;
                worksheet.Cells[23, 6] = tradingInfo.CurrentBuySum;
                worksheet.Cells[23, 7] = tradingInfo.CurrentNChg2;
                worksheet.Cells[24, 1] = "NEXT-M ONWRD";
                worksheet.Cells[24, 2] = tradingInfo.NextSellSum;
                worksheet.Cells[24, 5] = tradingInfo.NextNChg1;
                worksheet.Cells[24, 6] = tradingInfo.NextBuySum;
                worksheet.Cells[24, 7] = tradingInfo.NextNChg2;
                worksheet.Cells[25, 1] = "TOTAL";
                worksheet.Cells[25, 2] = tradingInfo.TotalSellSum;
                worksheet.Cells[25, 5] = tradingInfo.TotalNchg1;
                worksheet.Cells[25, 6] = tradingInfo.TotalBuySum;
                worksheet.Cells[25, 7] = tradingInfo.TotalNchg2;
                Range a23a25Range = ExcelUtil.GetRange(23, 1, 25, 1, worksheet);
                a23a25Range.Borders.LineStyle = 1;
                a23a25Range.Borders.Weight = XlBorderWeight.xlMedium;
                Range b23g25Range = ExcelUtil.GetRange(23, 2, 25, 7, worksheet);
                b23g25Range.Borders.LineStyle = 1;
                b23g25Range.Borders.Weight = XlBorderWeight.xlMedium;
                Range b22d25Range = ExcelUtil.GetRange(22, 2, 25, 4, worksheet);
                b22d25Range.Merge(true);
                workbook.Close(true, workbook.FullName, false);
                AddResult("Target file", fileName, "file");
                //TaskResultList.Add(new TaskResultEntry("Target File", "", fileName));
            }
        }

        #region TRADG Related

        private DateTime GetDateTime(string sourceDateStr)
        {
            string year = DateTime.Today.ToString("yy");
            string month = sourceDateStr.Substring(sourceDateStr.IndexOf("年") + 1, sourceDateStr.IndexOf("月") - sourceDateStr.IndexOf("年") - 1);
            string day = sourceDateStr.Substring(sourceDateStr.IndexOf("月") + 1, sourceDateStr.IndexOf("日") - sourceDateStr.IndexOf("月") - 1);
            string dateTime = month + "/" + day + "/" + year;
            DateTime date = DateTime.Parse(dateTime);
            return date;
        }
        private JpTRADGInfo printAndGetJpDailyTradingInfo(string filePath)
        {
            JpTRADGInfo tradgInfo = new JpTRADGInfo();
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                JapanShared.PrintWorksheet(worksheet, printerName, XlPageOrientation.xlPortrait);
                tradgInfo.GSellSum = ExcelUtil.GetRange(7, 5, worksheet).Value2.ToString().Trim();
                tradgInfo.GBuySum = ExcelUtil.GetRange(7, 10, worksheet).Value2.ToString().Trim();
                tradgInfo.CurrentSellSum = ExcelUtil.GetRange(12, 5, worksheet).Value2.ToString().Trim();
                tradgInfo.NextSellSum = ExcelUtil.GetRange(12, 6, worksheet).Value2.ToString().Trim();
                tradgInfo.TotalSellSum = ExcelUtil.GetRange(12, 8, worksheet).Value2.ToString().Trim();
                tradgInfo.CurrentNChg1 = ExcelUtil.GetRange(13, 5, worksheet).Value2.ToString();
                tradgInfo.NextNChg1 = ExcelUtil.GetRange(13, 6, worksheet).Value2.ToString();
                tradgInfo.TotalNchg1 = ExcelUtil.GetRange(13, 8, worksheet).Value2.ToString();
                tradgInfo.CurrentBuySum = ExcelUtil.GetRange(12, 10, worksheet).Value2.ToString();
                tradgInfo.NextBuySum = ExcelUtil.GetRange(12, 12, worksheet).Value2.ToString();
                tradgInfo.TotalBuySum = ExcelUtil.GetRange(12, 14, worksheet).Value2.ToString();
                tradgInfo.CurrentNChg2 = ExcelUtil.GetRange(13, 10, worksheet).Value2.ToString();
                tradgInfo.NextNChg2 = ExcelUtil.GetRange(13, 12, worksheet).Value2.ToString();
                tradgInfo.TotalNchg2 = ExcelUtil.GetRange(13, 14, worksheet).Value2.ToString().Trim();
                tradgInfo.Date = GetDateTime(ExcelUtil.GetRange(33, 3, worksheet).Value2.ToString().Trim()).ToString("ddMMMyy");
                int currentRow = 37;
                while (ExcelUtil.GetRange(currentRow, 3, worksheet).Value2.ToString().Trim() != "-")
                {
                    Transaction tran = new Transaction
                    {
                        CompanyInfo =
                        {
                            OriginalName =
                                ExcelUtil.GetRange(currentRow, 3, worksheet).Value2.ToString().Replace("　", "")
                        }
                    };
                    if (nameDic.ContainsKey(tran.CompanyInfo.OriginalName))
                    {
                        tran.CompanyInfo.EnglishName = nameDic[tran.CompanyInfo.OriginalName].EnglistName;
                        tran.CompanyInfo.ShortEnglishName = nameDic[tran.CompanyInfo.OriginalName].ShortName;
                        tran.CompanyInfo.JapaneseName = nameDic[tran.CompanyInfo.OriginalName].JapaneseName;
                    }
                    else
                    {
                        Logger.Log(string.Format("There's no such name for {0}, please check the baknote file.", tran.CompanyInfo.OriginalName), Logger.LogType.Warning);
                    }
                    tran.SellSum = ExcelUtil.GetRange(currentRow, 7, worksheet).Value2.ToString();
                    tran.BuySum = ExcelUtil.GetRange(currentRow, 9, worksheet).Value2.ToString();
                    tran.TotalSum = ExcelUtil.GetRange(currentRow, 13, worksheet).Value2.ToString();
                    tradgInfo.TransactionList.Add(tran);
                    currentRow++;
                }
                tradgInfo.Top15SellSum = ExcelUtil.GetRange(52, 7, worksheet).Value2.ToString().Trim();
                tradgInfo.Top15BuySum = ExcelUtil.GetRange(52, 9, worksheet).Value2.ToString().Trim();
                tradgInfo.Top15TotalSum = ExcelUtil.GetRange(52, 13, worksheet).Value2.ToString().Trim();
                tradgInfo.GTotalSum = ExcelUtil.GetRange(53, 13, worksheet).Value2.ToString().Trim();
                workbook.Close(false, workbook.FullName, Missing.Value);
            }
            return tradgInfo;
        }

        private List<string> GetSourceFileLinks(string baseUrl)
        {
            holidayList = HolidayManager.SelectHoliday(5);
            List<string> linkList = new List<string>();
            string pageSource = WebClientUtil.GetPageSource(baseUrl, 1800000);
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(pageSource);
            var tableList = htmlDoc.DocumentNode.SelectNodes("//table[@class = 'styleShiryo']");
            var nodeDailyList = tableList[1].SelectNodes("tr/td");
            string dailyExcelLink = "http://www.tse.or.jp" + MiscUtil.GetCleanTextFromHtml(nodeDailyList[2].SelectSingleNode("a").Attributes["href"].Value);
            linkList.Add(dailyExcelLink);
            if (DateTime.Today == MiscUtil.GetNextWeeklyTradingDay(DateTime.Today, holidayList, 2) || specialCase.Equals("yes"))
            {
                var nodeWeeklyList = tableList[2].SelectNodes("tr/td");
                string titleWeekly = MiscUtil.GetCleanTextFromHtml(nodeWeeklyList[0].InnerText);
                if (IsLastWeekInfo(titleWeekly) || specialCase.Equals("yes"))
                {
                    string weeklyExcelLink = "http://www.tse.or.jp" + MiscUtil.GetCleanTextFromHtml(nodeWeeklyList[2].SelectSingleNode("a").Attributes["href"].Value);
                    linkList.Add(weeklyExcelLink);
                }
            }
            return linkList;
        }

        private bool IsLastWeekInfo(string title)
        {
            string[] arr = title.Split('～');
            if (arr.Length == 2)
            {
                DateTime startDate = GetDateFromTitle(arr[0].Trim());
                if (!arr[1].Contains(" 年"))
                {
                    arr[1] = startDate.Year + "年" + arr[1];
                }
                DateTime endDate = GetDateFromTitle(arr[1].Trim());
                return DateTime.Today.AddDays(-7) >= startDate && DateTime.Today.AddDays(-7) <= endDate;
            }
            Logger.Log("The weekly information is not available yet, please wait.......");
            return false;
        }

        private DateTime GetDateFromTitle(string title)
        {
            int indexYear = title.IndexOf("年");
            int indexMonth = title.IndexOf("月");
            int indexDay = title.IndexOf("日");
            string year = title.Substring(0, indexYear);
            string month = title.Substring(indexYear + 1, indexMonth - indexYear - 1);
            string day = title.Substring(indexMonth + 1, indexDay - indexMonth - 1);
            string dateTime = month + "/" + day + "/" + year;
            DateTime date = DateTime.Parse(dateTime);
            return date;
        }

        #endregion

        private string DownloadFile(string linkUrl)
        {
            string[] part = linkUrl.Split('/');
            string filePath = Path.Combine(sourceFileDir, part[part.Length - 1]);
            WebClientUtil.DownloadFile(linkUrl, 180000, filePath);
            return filePath;
        }

        //get all the transactions
        private List<Transaction> GetTransactionList(int startRow, int startCol, Worksheet worksheet)
        {
            List<Transaction> transactionList = new List<Transaction>();
            int currentRow = startRow;
            int currentCol = startCol;
            while (currentRow < startRow + 15)
            {
                Range r = ExcelUtil.GetRange(currentRow, currentCol, worksheet);
                if (r.Value2 != null && r.Value2.ToString().Trim() != string.Empty && r.Value2.ToString().Trim() != "-")
                {
                    string warningMsg = string.Empty;
                    Transaction transaction = new Transaction
                    {
                        CompanyInfo =
                        {
                            OriginalName = ExcelUtil.GetRange(currentRow, 3, worksheet).Value2.ToString().Trim()
                        }
                    };
                    JapanShared.UpdateCompanyInfo(nameDic, transaction.CompanyInfo, out warningMsg);
                    Logger.Log(warningMsg);
                    transaction.SellSum = ExcelUtil.GetRange(currentRow, 7, worksheet).Value2.ToString().Trim();
                    transaction.BuySum = ExcelUtil.GetRange(currentRow, 9, worksheet).Value2.ToString().Trim();
                    transaction.TotalSum = ExcelUtil.GetRange(currentRow, 13, worksheet).Value2.ToString().Trim();
                    transactionList.Add(transaction);
                }
                currentRow++;
            }
            return transactionList;
        }
        private void WriteTransactionInfo(IEnumerable<Transaction> transactionList, Worksheet worksheet, int currentRow, int currentCol)
        {
            for (int i = currentRow; i < currentRow + 15; i++)
            {
                worksheet.Cells[i, 1] = (i - currentRow + 1).ToString();
            }
            using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, currentRow, currentCol + 1, ExcelLineWriter.Direction.Right))
            {
                foreach (Transaction transaction in transactionList)
                {
                    writer.WriteLine(transaction.CompanyInfo.EnglishName);
                    writer.WriteLine(transaction.CompanyInfo.JapaneseName);
                    writer.WriteLine(transaction.SellSum);
                    writer.WriteLine(transaction.BuySum);
                    writer.WriteLine(transaction.TotalSum);
                    writer.PlaceNext(writer.Row + 1, currentCol + 1);
                }
            }
        }
    }
}
