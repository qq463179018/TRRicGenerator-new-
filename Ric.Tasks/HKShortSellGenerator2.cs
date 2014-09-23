using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.ComponentModel;
using System.Drawing.Design;
using Reuters.ProcessQuality.ContentAuto.Lib;
using Microsoft.Office.Interop.Excel;

namespace Ric.Generator.Lib.HK
{
    public class TradingInfo
    {
        public List<StockInfo> StockList { get; set; }
        public string DateStr;
        public string DesignatedSecuritiesRecordingSum { get; set; }
        public string DesignatedSharesShortSoldSum { get; set; }
        public string DesignatedShortSellTurnoverShares { get; set; }
        public string DesignatedShortSellTurnoverValue { get; set; }
        public string HKDTurnoverValue { get; set; }

        public string NonDesignatedSecuritiesRecordingSum { get; set; }
        public string NonDesignatedSharesShortSoldSum { get; set; }
        public string NonDesignatedShortSellTransactionSum { get; set; }
    }

    public class StockInfo
    {
        public string Ric { get; set; }
        public string StockName { get; set; }
        public string Shares { get; set; }
        public string Turnover { get; set; }

        //Format the information of a stock to be in one string line
        public string ToSingleLine()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(this.Ric.PadRight(12));
            sb.Append(this.StockName.PadRight(27));
            sb.Append(this.Shares.PadLeft(15));
            sb.Append(this.Turnover.PadLeft(20));
            return sb.ToString();
        }
    }

    //Items can be configured by user
    public class HKShortSellConfig
    {    

        [Category("URI")]
        [Description("MainBoard URI:  http://www.hkex.com.hk/eng/stat/smstat/ssturnover/ncms/ASHTMAIN.HTM")]
        public string MainBoard_URI { get; set; }

        [Category("URI")]
        [Description("GEM URI:  http://www.hkex.com.hk/eng/stat/smstat/ssturnover/ncms/ASHTGEM.HTM")]
        public string GEM_URI { get; set; }

        [Description("Short Sell Tasks: task name should be one of \"HK-GEMSS\", \"HK-MAINSS\", \"HKMAIN01-10\" and the tasks will run in order")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<String> SHORTSELL_TASKLIST { get; set; }

        [Category("HKMAIN01_10_CONFIG")]
        [Description("Config for generating HKMAIN01_10.xls")]
        public HKMAIN01_10_Config HKMAIN01_10_CONFIG { get; set; }

        [Category("HKGEMSS_CONFIG")]
        [Description("Config for generating HKGEMSS.xls")]
        public HKGEMSS_Config HKGEMSS_CONFIG {get;set;}

        [Category("HKMAINSS_CONFIG")]
        [Description("Config for generating HKMAINSS.xls")]
        public HKMAINSS_Config HKMAINSS_CONFIG { get; set; }
    }


    //Config file for updating gemss file
    [TypeConverter(typeof(ExpandableObjectConverter))] 
    public class HKGEMSS_Config
    {
        public string WORKBOOK_PATH{get;set;}
        public string WORKSHEET_NAME{get;set;}
        public string XML_FILE_PATH { get; set; }
    }

    //Config file for generating Main01_10 file
    [TypeConverter(typeof(ExpandableObjectConverter))] 
    public class HKMAIN01_10_Config
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
        public string XML_FILE_PATH { get; set; }
    }

    //Config for generating MAINSS file
    [TypeConverter(typeof(ExpandableObjectConverter))] 
    public class HKMAINSS_Config
    {
        public string XML_FILE_PATH { get; set; }
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
        public int PAGE_NUM_TO_COPY { get; set; }
        public int LINE_EACH_PAGE_TO_COPY { get; set; }
        public int START_PAGE { get; set; }
        public int LINE_NUM { get; set; }
        public int TOTAL_PAGE_NUM { get; set; }
    }

    /// <summary>
    /// Short Sell Information Gatherer
    /// </summary>
    public class HKShortSellGenerator : GeneratorBase
    {
        private readonly string configFilePath = ".\\Config\\HK\\HKShortSell.config";
        private readonly int recordAlignrightLength = 74;
        private readonly int recordTotalLength = 80;
        private HKShortSellConfig configObj = null;

        private readonly string newsTitlePrefix = "Short Sell Turnover (Main Board) up to day close today".Trim().ToLower();

        protected override void Start()
        {
            try
            {
                StartShortSellJob();
            }
            catch(Exception ex)
            {
                logger.Log(ex.Message);
                logger.Log(ex.StackTrace);
            }
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(configFilePath, typeof(HKShortSellConfig)) as HKShortSellConfig;
            isEikonExcelDisable = true;          

        }
       

        /// <summary>
        /// Start short sell jobs according with the tasks configured in the HKShortSell.config file.
        /// </summary>
        public void StartShortSellJob()
        {
            if (configObj.SHORTSELL_TASKLIST == null || configObj.SHORTSELL_TASKLIST.Count < 1)
            {
                logger.LogErrorAndRaiseException("Please select at least one job! And the job name should be one of HK-GEMSS, HK-MAINSS, HKMAIN01-10");
            }

            foreach (string shortsellTask in configObj.SHORTSELL_TASKLIST)
            {
                if (shortsellTask.ToLower().Contains("hk-gemss"))
                {
                    logger.Log("Start running task: HK-GEMSS***********************");
                    Start_HK_GEMSS();
                    logger.Log("Task: HK-GEMSS Completed***************************");
                }

                else if (shortsellTask.ToLower().Contains("hk-mainss"))
                {
                    logger.Log("Start running task: HK-MAINSS**********************");
                    Start_HK_MAINSS();
                    logger.Log("Task: HK-MAINSS Completed**************************");
                }

                else if (shortsellTask.ToLower().Contains("hkmain01-10"))
                {
                    logger.Log("Start running task: HKMAINSS01-10******************");
                    Start_HKMAIN01_10();
                    logger.Log("Task: HKMAINSS01-10 Completed**********************");
                }
            }
        }

        /// <summary>
        /// Check if the dispaly date is the current day.
        /// </summary>
        /// <param name="displayDate">the display date.</param>
        /// <returns>a bool value which is true or false.</returns>
        public bool IsCurrentDay(string displayDate)
        {
            return displayDate == DateTime.Today.ToString("dd/MM/yyyy");
        }

        #region ShortSell Task1: Updating GEMSS File

        /// <summary>
        /// Update GEMSS file.
        /// </summary>
        public void Start_HK_GEMSS()
        {
            string uri = configObj.GEM_URI;
                //"/eng/stat/smstat/ssturnover/ncms/ASHTGEM.HTM";
            //uri = MiscUtil.UrlCombine(configObj.BASE_URI, uri);;
            TradingInfo gemTradingInfo = GetGemTradingInfo(uri);
            UpdateGemssFileAndGenerateXMLFile(gemTradingInfo); 
        }

        /// <summary>
        /// Get GEM trading News Info.
        /// </summary>
        /// <param name="uri">the uri of the source data which is a html page.</param>
        /// <returns>a instance of the Class TradingInfo which is about the GEM trading News Info.</returns>
        public TradingInfo GetGemTradingInfo(string uri)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = null;
            htmlDoc = WebClientUtil.GetHtmlDocument(uri, 180000);

            TradingInfo gemTradingInfo = new TradingInfo();
            gemTradingInfo.StockList = new List<StockInfo>();
            List<string> valueList = new List<string>();
            string dateStr = DateTime.Now.ToString("dddd dd/MM/yyyy");
            gemTradingInfo.DateStr = string.Format("Recorded as of {0} 04:00 pm :-", dateStr);

            //Get the trading news information
            string gemTradingNewsInfo = htmlDoc.DocumentNode.SelectSingleNode("//body/pre/font").InnerText;

            //Parse and get the required information
            if (!IsNewsExist(gemTradingNewsInfo))
            {
                gemTradingInfo.StockList = null ;
            }

            using (StringReader sr = new StringReader(gemTradingNewsInfo))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    string[] frags = line.Split(new string[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
                    int ric = 0;
                    if (frags.Length == 4 && int.TryParse(frags[0].Trim(), out ric))
                    {
                        StockInfo stockInfo = new StockInfo();
                        stockInfo.Ric = string.Format("<{0}.HK>", ric.ToString("D4"));
                        stockInfo.StockName = frags[1].Trim();
                        stockInfo.Shares = frags[2].Trim();
                        stockInfo.Turnover = frags[3].Trim();
                        gemTradingInfo.StockList.Add(stockInfo);
                        continue;
                    }
                    
                    frags = line.Split(':');
                    if (frags.Length == 2 && frags[1].Trim() != "")
                    {
                        valueList.Add(line.Trim());
                    }
                }
            }

            //Parse to get the summary info of Main Board
            if (valueList != null && valueList.Count > 0)
            {
                gemTradingInfo.DesignatedSecuritiesRecordingSum = valueList[1];
                gemTradingInfo.DesignatedSharesShortSoldSum = valueList[2];
                gemTradingInfo.DesignatedShortSellTurnoverShares = valueList[3];
                gemTradingInfo.DesignatedShortSellTurnoverValue = valueList[4];
                if (valueList[5].Contains("Short Selling Turnover Total Value ($)          : HKD  "))
                {
                    gemTradingInfo.HKDTurnoverValue = valueList[5];
                }
            }
            return gemTradingInfo;
        }

        /// <summary>
        /// Update HK-GEMSS.xls file and create HK-GEMSS.xml
        /// </summary>
        /// <param name="gemTradingInfo">the GEM trading News Info.</param>
        public void UpdateGemssFileAndGenerateXMLFile(TradingInfo gemTradingInfo)
        {
           // string gemssFilePath = MiscUtil.BackupFileWithNewName(configObj.HKGEMSS_CONFIG.WORKBOOK_PATH);
            string gemssFilePath = MiscUtil.BackUpFileWithDateFolder(configObj.HKGEMSS_CONFIG.WORKBOOK_PATH, false);
            using (ExcelApp app = new ExcelApp(false,false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, gemssFilePath);
                var worksheet = ExcelUtil.GetWorksheet(configObj.HKGEMSS_CONFIG.WORKSHEET_NAME, workbook);
                if (worksheet == null)
                {
                    logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.HKGEMSS_CONFIG.WORKSHEET_NAME, workbook.Name));
                }

                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Down))
                {
                    // Allocate spaces
                    ExcelUtil.InsertBlankRows(ExcelUtil.GetRange("A1", worksheet), 24);
                    writer.PlaceNext(1, 1);

                    // Form 1st line and write
                    string line = string.Empty;
                    writer.WriteLine(gemTradingInfo.DateStr);

                    // Copy 2nd and 3rd line from line 26 and line 27
                    writer.WriteLine(worksheet.Cells[writer.Row + 24, writer.Col]);
                    writer.WriteLine(worksheet.Cells[writer.Row + 24, writer.Col]);

                    if (gemTradingInfo.StockList == null)
                    {
                        writer.WriteLine("NIL");
                        while (writer.Row < 19)
                        {
                            writer.WriteLine(" ");
                        }
                    }
                    else
                    {
                        // Fill stock info
                        foreach (StockInfo stockInfo in gemTradingInfo.StockList)
                        {
                            StringBuilder sb = new StringBuilder();
                            sb.Append(stockInfo.Ric.PadRight(12));
                            sb.Append(stockInfo.StockName.PadRight(27));
                            sb.Append(stockInfo.Shares.PadLeft(15));
                            sb.Append(stockInfo.Turnover.PadLeft(20));
                            line = sb.ToString();
                            writer.WriteLine(line);
                        }
                        while (writer.Row < 19)
                        {
                            writer.WriteLine(" ");
                        }

                        //Fill non-designated securities summary if non-designated securities exist
                        writer.PlaceNext(19, 1);
                        if (!string.IsNullOrEmpty(gemTradingInfo.NonDesignatedSecuritiesRecordingSum) && gemTradingInfo.NonDesignatedSecuritiesRecordingSum != "0")
                        {

                            writer.WriteLine(gemTradingInfo.NonDesignatedSecuritiesRecordingSum.PadLeft(recordAlignrightLength));
                            writer.WriteLine(gemTradingInfo.NonDesignatedSharesShortSoldSum.PadLeft(recordAlignrightLength));
                            writer.WriteLine(gemTradingInfo.NonDesignatedShortSellTransactionSum.PadLeft(recordAlignrightLength));
                        }
                        // Fill designated securities summary
                        writer.WriteLine(gemTradingInfo.DesignatedSecuritiesRecordingSum.PadLeft(recordAlignrightLength));
                        writer.WriteLine(gemTradingInfo.DesignatedSharesShortSoldSum.PadLeft(recordAlignrightLength));
                        writer.WriteLine(gemTradingInfo.DesignatedShortSellTurnoverShares.PadLeft(recordAlignrightLength));
                        writer.WriteLine(gemTradingInfo.DesignatedShortSellTurnoverValue.PadLeft(recordAlignrightLength));
                        if (!string.IsNullOrEmpty(gemTradingInfo.HKDTurnoverValue))
                        {
                            writer.WriteLine(gemTradingInfo.HKDTurnoverValue.PadLeft(recordAlignrightLength));
                        }
                    }

                    for (int page = 1; page <= 25; page++)
                    {
                        string id = "HK/GEMSS" + page.ToString("D2");
                        int upRow = 24 * (page - 1) + 1;
                        int downRow = upRow + 23;
                        writer.PlaceNextAndWriteLine(upRow, 3, id);
                        writer.PlaceNextAndWriteLine(downRow, 3, id);
                    }

                    //Fill color for C columns
                    Range range = ExcelUtil.GetRange(1, 3, 24, 3, worksheet);
                    range.Interior.Color = ExcelUtil.GetRange(49, 3, worksheet).Interior.Color;

                    ProductionXmlFileTemplate xmlFileTem = GetGemssXmlFileContent(worksheet);
                    ConfigUtil.WriteXml(configObj.HKGEMSS_CONFIG.XML_FILE_PATH, xmlFileTem);
                    TaskResultList.Add(new TaskResultEntry("XML file for HK-GEMSS", "", configObj.HKGEMSS_CONFIG.XML_FILE_PATH));


                    if (File.Exists(configObj.HKGEMSS_CONFIG.WORKBOOK_PATH))
                    {
                        File.Delete(configObj.HKGEMSS_CONFIG.WORKBOOK_PATH);
                    }
                    //Save files as a copy
                    workbook.SaveCopyAs(configObj.HKGEMSS_CONFIG.WORKBOOK_PATH);
                    TaskResultList.Add(new TaskResultEntry("HK-GEMSS","",configObj.HKGEMSS_CONFIG.WORKBOOK_PATH));
                    //Close current opend excel file
                    workbook.Close(false, gemssFilePath, false);
                }
            }
        }

        private ProductionXmlFileTemplate GetGemssXmlFileContent(Worksheet gemssWorksheet)
        {
            ProductionXmlFileTemplate template = new ProductionXmlFileTemplate();
            int lastAddedRicNum = 0;
            for (int i = 1; i <= 240; i++)
            {
                if (i % 24 == 1)
                {
                    lastAddedRicNum = template.rics.rics.Count - 1;
                    if (lastAddedRicNum > 0)
                    {
                        Fid lastFid = new Fid();
                        lastFid.Id = 339;
                        lastFid.Value = string.Format("\"Continued on <{0}>\"", template.rics.rics[lastAddedRicNum].Name);
                    }
                    Reuters.ProcessQuality.ContentAuto.Lib.Ric ric = new Reuters.ProcessQuality.ContentAuto.Lib.Ric();
                    ric.Name = string.Format("HK/GEMSS{0}", (i / 24+1).ToString("D2"));
                    template.rics.rics.Add(ric);
                }

                if (i % 24 == 0)
                {
                    continue;
                }

                Fid fid = new Fid();
                fid.Id = 316+(i-1)%24;
                if (ExcelUtil.GetRange(i, 1, gemssWorksheet).Text != null)
                {
                    fid.Value = string.Format("\"{0}\"", ExcelUtil.GetRange(i, 1, gemssWorksheet).Text.ToString());
                     if (fid.Value == "\" \"")
                    {
                        fid.Value = " ".PadLeft(recordTotalLength);
                        fid.Value = "\"" + fid.Value + "\"";
                    }
                }
                else
                {
                    fid.Value = string.Format("\"{0}\"", " ".PadLeft(recordAlignrightLength));
                }

                lastAddedRicNum = template.rics.rics.Count - 1;
                template.rics.rics[lastAddedRicNum].fids.Add(fid);
            }
            return template;
        }
        
        /// <summary>
        /// Check if there's content of trading news.
        /// </summary>
        /// <param name="tradingNewsInfo">the trading News Information which is getted from a special web  html.</param>
        /// <returns>true if exists, false otherwise</returns>
        public bool IsNewsExist(string tradingNewsInfo)
        {
            Regex Regex = new Regex("\\bCODE\\b");
            Match m = Regex.Match(tradingNewsInfo);
            if (m.Length > 0)
            {
                return true;
            }
            return false;
        }

        #endregion


        #region ShortSell Task2: Generation Main01-10 File

        /// <summary>
        /// GeneratingHKMAIN01_10 File.
        /// </summary>
        public void Start_HKMAIN01_10()
        {
            TradingInfo main01_10TradingInfo = GetMain01_10TradingInfo();
            GenerateMain01_10ExcelAndXMLFile(main01_10TradingInfo);
        }

        /// <summary>
        /// Get trading information about "(Main Board) up to day" .
        /// </summary>
        /// <returns>a instance of the Class TradingInfo which is about the Main01_10 trading News Info.</returns>
        public TradingInfo GetMain01_10TradingInfo()
        {

            string url = configObj.MainBoard_URI;
            //  "/eng/stat/smstat/ssturnover/ncms/ASHTMAIN.HTM";
            //  url = MiscUtil.UrlCombine(configObj.BASE_URI, url);

            TradingInfo main01_10TradingInfo = new TradingInfo();
            main01_10TradingInfo.StockList = new List<StockInfo>();
            List<string> valueList = new List<string>();
            string tradingNewsInfo = string.Empty;
            string dateStr = DateTime.Now.ToString("dddd dd/MM/yyyy");
            main01_10TradingInfo.DateStr = string.Format("Recorded as of {0} 04:00 pm :-", dateStr);
            
            var htmlDoc = WebClientUtil.GetHtmlDocument(url, 180000);

            tradingNewsInfo = htmlDoc.DocumentNode.SelectSingleNode("//body/pre/font").InnerText;
            using (StringReader sr = new StringReader(tradingNewsInfo))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    //Parse the stock info

                    string[] frags = line.Split(new string[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
                    int ric = 0;
                    if (frags.Length == 4 && int.TryParse(frags[0].Trim(), out ric))
                    {
                        StockInfo stockInfo = new StockInfo();
                        stockInfo.Ric = string.Format("<{0}.HK>", ric.ToString("D4"));
                        stockInfo.StockName = frags[1].Trim();
                        stockInfo.Shares = frags[2].Trim();
                        stockInfo.Turnover = frags[3].Trim();
                        main01_10TradingInfo.StockList.Add(stockInfo);
                        continue;
                    }
                    else if (frags.Length == 4 && !int.TryParse(frags[0].Trim(), out ric))
                    {
                        Regex regExpression = new Regex("[0-9]+");
                        Match match = regExpression.Match(frags[0].Trim());
                        if (match.Success)
                        {
                            StockInfo stockInfo = new StockInfo();
                            stockInfo.Ric = string.Format("<{0}.HK>", match.Value);
                            stockInfo.StockName = frags[1].Trim();
                            stockInfo.Shares = frags[2].Trim();
                            stockInfo.Turnover = frags[3].Trim();
                            main01_10TradingInfo.StockList.Add(stockInfo);
                            continue;
                        }
                    }

                    frags = line.Split(':');
                    if (frags.Length == 2 && frags[1].Trim() != "")
                    {
                        valueList.Add(line.Trim());
                    }
                }

                //Parse to get the summary info of Main Board
                if (valueList != null && valueList.Count > 0)
                {
                    main01_10TradingInfo.DesignatedSecuritiesRecordingSum = valueList[1];
                    main01_10TradingInfo.DesignatedSharesShortSoldSum = valueList[2];
                    main01_10TradingInfo.DesignatedShortSellTurnoverShares = valueList[3];
                    main01_10TradingInfo.DesignatedShortSellTurnoverValue = valueList[4];
                    main01_10TradingInfo.HKDTurnoverValue = valueList[5];
                }
            }
            return main01_10TradingInfo;
        }

        /// <summary>
        /// Generate the HKMAIN01_10Excel file and the HKMAIN01_10Xml file.
        /// </summary>
        /// <param name="main01_10TradingInfo">the Main01_10 trading News Info.</param>
        public void GenerateMain01_10ExcelAndXMLFile(TradingInfo main01_10TradingInfo)
        {
            string main01_10FilePath = MiscUtil.BackUpFileWithDateFolder(configObj.HKMAIN01_10_CONFIG.WORKBOOK_PATH, false); //MiscUtil.BackupFileWithNewName(configObj.HKMAIN01_10_CONFIG.WORKBOOK_PATH);
            using (ExcelApp app = new ExcelApp(true,false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, main01_10FilePath);
                var worksheet = ExcelUtil.GetWorksheet(configObj.HKMAIN01_10_CONFIG.WORKSHEET_NAME, workbook);
                if (worksheet == null)
                {
                    logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.HKMAIN01_10_CONFIG.WORKSHEET_NAME, workbook.Name));
                }

                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    // Update the date 
                    string line = string.Empty;
                    writer.WriteLine(main01_10TradingInfo.DateStr);

                    //Copy from line 4
                    writer.PlaceNext(4,1);
                    foreach (StockInfo stock in main01_10TradingInfo.StockList)
                    {
                        writer.WriteLine(stock.Ric);
                        writer.WriteLine(stock.StockName);
                        writer.WriteLine(stock.Shares);
                        writer.WriteLine(stock.Turnover);
                        writer.PlaceNext(writer.Row + 1, 1);
                    }
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                    ExcelUtil.GetRange(writer.Row,writer.Col,lastUsedRow,writer.Col+4,worksheet).Clear();

                    writer.PlaceNext(writer.Row+2,1);
                    writer.WriteLine(main01_10TradingInfo.DesignatedSecuritiesRecordingSum);

                    writer.PlaceNext(writer.Row+1,1);
                    writer.WriteLine(main01_10TradingInfo.DesignatedSharesShortSoldSum);

                    writer.PlaceNext(writer.Row+1,1);
                    writer.WriteLine(main01_10TradingInfo.DesignatedShortSellTurnoverShares);

                    writer.PlaceNext(writer.Row + 1, 1);
                    writer.WriteLine(main01_10TradingInfo.DesignatedShortSellTurnoverValue);
                    
                    writer.PlaceNext(writer.Row + 1, 1);
                    writer.WriteLine(main01_10TradingInfo.HKDTurnoverValue);
                    
                    ExcelUtil.GetRange(1, 1, writer.Row-1, 4, worksheet).Borders.LineStyle = 0;
                }

                ProductionXmlFileTemplate xmlFileTem = GetHKMAINXmlFileContent(worksheet);
                ConfigUtil.WriteXml(configObj.HKMAIN01_10_CONFIG.XML_FILE_PATH, xmlFileTem);
                TaskResultList.Add(new TaskResultEntry("XML file for HK-MAIN01-10", "", configObj.HKMAIN01_10_CONFIG.XML_FILE_PATH));
               
                //Save the HK-MAIN01-10 file as a copy

                if (File.Exists(configObj.HKMAIN01_10_CONFIG.WORKBOOK_PATH))
                {
                    File.Delete(configObj.HKMAIN01_10_CONFIG.WORKBOOK_PATH);
                }
                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbook.SaveCopyAs(configObj.HKMAIN01_10_CONFIG.WORKBOOK_PATH);

                TaskResultList.Add(new TaskResultEntry("HKMAINSS01-10", "", configObj.HKMAIN01_10_CONFIG.WORKBOOK_PATH));
                workbook.Close(false, main01_10FilePath, true);
            }
        }

        private ProductionXmlFileTemplate GetHKMAINXmlFileContent(Worksheet worksheet)
        {
            ProductionXmlFileTemplate template = new ProductionXmlFileTemplate();
            int lastAddedRicNum = 0;
            string dateStr = ExcelUtil.GetRange(1, 1, worksheet).Text.ToString();
            string titleLine = "RIC         STOCK NAME                          SHARES        TURNOVER ($)";
            string partingLine = "=========   ====================       ===============     ===============";
            for (int i= 4; i <= 503; i++)
            {
                if (i % 20 == 4)
                {
                    Reuters.ProcessQuality.ContentAuto.Lib.Ric ric = new Reuters.ProcessQuality.ContentAuto.Lib.Ric();
                    ric.Name = string.Format("HK/MAINSS{0}", (i / 20 + 1).ToString("D2"));
                    template.rics.rics.Add(ric);
                    Fid fidDate = new Fid();
                    fidDate.Id = 316;
                    fidDate.Value = string.Format("\"{0}\"", dateStr);
                    lastAddedRicNum = template.rics.rics.Count - 1;
                    template.rics.rics[lastAddedRicNum].fids.Add(fidDate);

                    Fid fidTitleLine = new Fid();
                    fidTitleLine.Id = 317;
                    fidTitleLine.Value = string.Format("\"{0}\"", titleLine);
                    template.rics.rics[lastAddedRicNum].fids.Add(fidTitleLine);

                    Fid fidPartingLine = new Fid();
                    fidPartingLine.Id = 318;
                    fidPartingLine.Value = string.Format("\"{0}\"", partingLine);
                    template.rics.rics[lastAddedRicNum].fids.Add(fidPartingLine);
                    
                }
                Fid fid2 = new Fid();
                fid2.Id = (i - 4) % 20 + 319;
                fid2.Value = GetFidValue(worksheet, i);

                lastAddedRicNum = template.rics.rics.Count - 1;
                template.rics.rics[lastAddedRicNum].fids.Add(fid2);
            }
            return template;
        }

        private string GetFidValue(Worksheet worksheet, int i)
        {
            string value = "";
            string colValue = "";
            try
            {
                value += "\"";
                if (ExcelUtil.GetRange(i, 1, worksheet).Text == null)
                {
                    value += string.Format("{0}", " ".PadLeft(recordAlignrightLength));
                }
                else
                {
                    string firstColText = ExcelUtil.GetRange(i, 1, worksheet).Text.ToString().Trim();
                    if (firstColText.StartsWith("<"))
                    {
                        colValue = ExcelUtil.GetRange(i, 1, worksheet).Text.ToString().Trim();
                        //value += string.Format("{0}   ",colValue.PadRight(11));
                        value += string.Format("{0}", colValue.PadRight(12));

                        colValue = ExcelUtil.GetRange(i, 2, worksheet).Text.ToString().Trim();
                        //value += string.Format("{0}       ", colValue.PadRight(20));
                        value += string.Format("{0}", colValue.PadRight(27));

                        colValue = ExcelUtil.GetRange(i, 3, worksheet).Text.ToString().Trim();
                        //value += string.Format("{0}     ", colValue.PadLeft(15));
                        value += string.Format("{0}", colValue.PadLeft(15));

                        colValue = ExcelUtil.GetRange(i, 4, worksheet).Text.ToString().Trim();
                        //value += string.Format("{0}        ", colValue.PadLeft(15));
                        value += string.Format("{0}", colValue.PadLeft(20));
                    }
                    else
                    {
                        colValue = ExcelUtil.GetRange(i, 1, worksheet).Text.ToString().Equals(string.Empty) ? " " : ExcelUtil.GetRange(i, 1, worksheet).Text.ToString().Trim();
                        value += string.Format("{0}", colValue.PadLeft(recordAlignrightLength));
                    }
                }
                
                value += "\"";

            }
            catch (Exception ex)
            { string errInfo = ex.ToString(); }

            return value;
        }

        /// <summary>
        /// Get "(Main Board) up to day "links from uri http://www.hkex.com.hk/eng/market/sec_tradinfo/tradnews/today/news.htm  .
        /// </summary>
        /// <param name="uri">the uri of the data source which is a web html.</param>
        /// <returns>a list which consists of the uris of the links</returns>
        public List<string> GetUrlLinksFromMainPage(string uri)
        {
            List<string> linkUrlList = new List<string>();
            //uri = MiscUtil.UrlCombine(configObj.BASE_URI, uri);
            var htmlDoc = WebClientUtil.GetHtmlDocument(uri,180000);
            var linkNodeList = htmlDoc.DocumentNode.SelectNodes("//span[@id='Content']/table/tbody/tr//td//a");
            foreach (var linkNode in linkNodeList)
            {
                if (linkNode.Attributes["href"] != null)
                {
                    string linkText = linkNode.InnerText;
                    string linkUrl = linkNode.Attributes["href"].Value;
                    if (!MiscUtil.IsAbsUrl(linkUrl))
                    {
                        linkUrl = MiscUtil.UrlCombine(uri, linkUrl);
                    }
                    if (linkText.Trim().ToLower().StartsWith(newsTitlePrefix)&&!(linkUrlList.Contains(linkUrl)))
                    {
                        linkUrlList.Add(linkUrl);
                    }
                }
            }
            return linkUrlList;
        }

        #endregion


        #region ShortSell Task3: Generating HK_Mainss File

        /// <summary>
        /// Generate HK-MAINSS file by copying contents from HK-MAIN01-10.xls .
        /// </summary>
        public void Start_HK_MAINSS()
        {
            TradingInfo tradingInfo = GetTradingInfoFromMain01_10File(configObj.HKMAINSS_CONFIG.PAGE_NUM_TO_COPY, configObj.HKMAINSS_CONFIG.LINE_EACH_PAGE_TO_COPY);
            GenerateMAINSSExcelAndXmlFile(configObj.HKMAINSS_CONFIG.WORKBOOK_PATH, tradingInfo);
        }

        //
        /// <summary>
        /// Get trading information from the HK-MAIN01-10.xls file.
        /// </summary>
        /// <param name="pageNum">the page number to copy.</param>
        /// <param name="lineNumEachPage">the line number of each page to copy.</param>
        /// <returns>the Main01_10 trading News Information.</returns>
        public TradingInfo GetTradingInfoFromMain01_10File(int pageNum, int lineNumEachPage)
        {
            TradingInfo tradingInfo = new TradingInfo();
            tradingInfo.StockList = new List<StockInfo>();

            //Open HK-MAIN01-10.xls
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.HKMAIN01_10_CONFIG.WORKBOOK_PATH);
                var worksheet = ExcelUtil.GetWorksheet(configObj.HKMAIN01_10_CONFIG.WORKSHEET_NAME, workbook);
                if (worksheet == null)
                {
                    logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.HKMAIN01_10_CONFIG.WORKSHEET_NAME, workbook.Name));
                }

                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    tradingInfo.DateStr = ExcelUtil.GetRange(1, 1, worksheet).Text.ToString();
                    reader.PlaceNext(4, 1);
                    //Range range = ExcelUtil.GetRange(reader.Row, 1, worksheet);
                    while ((reader.Row < (pageNum * lineNumEachPage) + 11))
                    {
                        if (ExcelUtil.GetRange(reader.Row, 1, worksheet).Text != null)
                        {
                            string firstColText = ExcelUtil.GetRange(reader.Row, 1, worksheet).Text.ToString().Trim();
                            if (firstColText.StartsWith("<"))
                            {
                                StockInfo stockInfo = new StockInfo();
                                stockInfo.Ric = reader.ReadLineCellText();
                                stockInfo.StockName = reader.ReadLineCellText();
                                stockInfo.Shares = reader.ReadLineCellText();
                                stockInfo.Turnover = reader.ReadLineCellText();
                                tradingInfo.StockList.Add(stockInfo);
                                reader.PlaceNext(reader.Row + 1, 1);
                            }

                            else
                            {
                                if (firstColText != string.Empty)
                                {
                                    tradingInfo.DesignatedSecuritiesRecordingSum = ExcelUtil.GetRange(reader.Row, 1, worksheet).Text.ToString().Trim();
                                    tradingInfo.DesignatedSharesShortSoldSum = ExcelUtil.GetRange(reader.Row + 1, 1, worksheet).Text.ToString().Trim();
                                    tradingInfo.DesignatedShortSellTurnoverShares = ExcelUtil.GetRange(reader.Row + 2, 1, worksheet).Text.ToString().Trim();
                                    tradingInfo.DesignatedShortSellTurnoverValue = ExcelUtil.GetRange(reader.Row + 3, 1, worksheet).Text.ToString().Trim();
                                    if (ExcelUtil.GetRange(reader.Row + 4, 1, worksheet).Text != null)
                                    {
                                        tradingInfo.HKDTurnoverValue = ExcelUtil.GetRange(reader.Row + 4, 1, worksheet).Text.ToString().Trim();
                                    }
                                    reader.PlaceNext(reader.Row + 5, 1);
                                }
                                else
                                {
                                    reader.PlaceNext(reader.Row + 1, 1);
                                }
                            }
                        }
                        else
                        {
                            reader.PlaceNext(reader.Row + 1, 1);
                        }
                        
                    }
                    workbook.Close(false, configObj.HKMAIN01_10_CONFIG.WORKBOOK_PATH, false);
                }
            }
            return tradingInfo;
        }

        private void GenerateMAINSSExcelAndXmlFile(string xlsFilePath, TradingInfo tradingInfo)
        {
            xlsFilePath = MiscUtil.BackUpFileWithDateFolder(xlsFilePath, false); //MiscUtil.BackupFileWithNewName(xlsFilePath);
            List<string> linesToWrite = new List<string>();
            foreach (StockInfo stockInfo in tradingInfo.StockList)
            {
                linesToWrite.Add(stockInfo.ToSingleLine());
            }
            linesToWrite.Add("");
            if (!string.IsNullOrEmpty(tradingInfo.DesignatedSecuritiesRecordingSum))
            {
                linesToWrite.Add(tradingInfo.DesignatedSecuritiesRecordingSum.PadLeft(recordAlignrightLength));
            }
            if (!string.IsNullOrEmpty(tradingInfo.DesignatedSharesShortSoldSum))
            {
                linesToWrite.Add(tradingInfo.DesignatedSharesShortSoldSum.PadLeft(recordAlignrightLength));
            }
            if (!string.IsNullOrEmpty(tradingInfo.DesignatedShortSellTurnoverShares))
            {
                linesToWrite.Add(tradingInfo.DesignatedShortSellTurnoverShares.PadLeft(recordAlignrightLength));
            }
            if (!string.IsNullOrEmpty(tradingInfo.DesignatedShortSellTurnoverValue))
            {
                linesToWrite.Add(tradingInfo.DesignatedShortSellTurnoverValue.PadLeft(recordAlignrightLength));
            }
            if (!string.IsNullOrEmpty(tradingInfo.HKDTurnoverValue))
            {
                linesToWrite.Add(tradingInfo.HKDTurnoverValue.PadLeft(recordAlignrightLength));
            }

            using (ExcelApp app = new ExcelApp(false,false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, xlsFilePath);
                var worksheet = ExcelUtil.GetWorksheet(configObj.HKMAINSS_CONFIG.WORKSHEET_NAME, workbook);
                if (worksheet == null)
                {
                    logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.HKGEMSS_CONFIG.WORKSHEET_NAME, workbook.Name));
                }

                //Generate MAINSSExcelFile
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Down))
                {
                    // Allocate spaces
                    int startOffset = (configObj.HKMAINSS_CONFIG.START_PAGE-1) * configObj.HKMAINSS_CONFIG.LINE_NUM;
                    int startRow = startOffset + 1;
                    int curLineNum = 0;

                    int insertBlankPageNum = (linesToWrite.Count - 1) / configObj.HKMAINSS_CONFIG.LINE_EACH_PAGE_TO_COPY + 1;
                    ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(startRow, 1, worksheet), insertBlankPageNum * 24);
                    writer.PlaceNext(startRow, 1);

                    // Form 1st line and write
                    string line = string.Empty;
                    List<string> headerLineList = new List<string>();
                    headerLineList.Add(tradingInfo.DateStr);
                    line = "RIC         STOCK NAME                          SHARES        TURNOVER ($)";
                    headerLineList.Add(line);
                    line = "=========   ====================       ===============     ===============";
                    headerLineList.Add(line);

                    for (int i = 0; i < insertBlankPageNum; i++)
                    {
                        int curStartRow = startOffset + i * configObj.HKMAINSS_CONFIG.LINE_NUM + 1;
                        int nextStartRow = curStartRow + 24;
                        writer.WriteLine(headerLineList[0]);
                        writer.WriteLine(headerLineList[1]);
                        writer.WriteLine(headerLineList[2]);
                        //keep a line without text.
                        writer.MoveNext();

                        while (writer.Row < nextStartRow && curLineNum < linesToWrite.Count)
                        {
                            writer.WriteLine(linesToWrite[curLineNum]);
                            curLineNum++;
                        }

                        writer.PlaceNext(nextStartRow, 1);
                    }

                    // Fill designated securities summary with page no. and background color.
                    for (int page = configObj.HKMAINSS_CONFIG.START_PAGE; page <= configObj.HKMAINSS_CONFIG.TOTAL_PAGE_NUM; page++)
                    {
                        string id = "HK/MAINSS" + page.ToString("D2");
                        int upRow = configObj.HKMAINSS_CONFIG.LINE_NUM * (page - 1) + 1;
                        int downRow = upRow + (configObj.HKMAINSS_CONFIG.LINE_NUM-1);
                        writer.PlaceNextAndWriteLine(upRow, 2, id);
                        writer.PlaceNextAndWriteLine(downRow, 2, id);
                        Range pageRange = ExcelUtil.GetRange(upRow, 2, downRow, 2, worksheet);
                        if (page % 2 == 0)
                        {
                            pageRange.Interior.Color = 16777164.0;
                        }
                        else
                        {
                            pageRange.Interior.Color = 10092543.0;
                        }
                    }

                    Range rangToRemove = ExcelUtil.GetRange(writer.Row, 1, writer.Row + 24 * insertBlankPageNum, 2, worksheet);
                    rangToRemove.Clear();
                }

                //Generate MAINSSXmlFile
                ProductionXmlFileTemplate xmlFileTem = GetHKMAINSSXmlFileContent(worksheet);
                ConfigUtil.WriteXml(configObj.HKMAINSS_CONFIG.XML_FILE_PATH, xmlFileTem);
                TaskResultList.Add(new TaskResultEntry("XML file for HKMAINSS", "", configObj.HKMAINSS_CONFIG.XML_FILE_PATH));


                if (File.Exists(configObj.HKMAINSS_CONFIG.WORKBOOK_PATH))
                {
                    File.Delete(configObj.HKMAINSS_CONFIG.WORKBOOK_PATH);
                }
                workbook.SaveCopyAs(configObj.HKMAINSS_CONFIG.WORKBOOK_PATH);
                TaskResultList.Add(new TaskResultEntry("HKMAINSS","",configObj.HKMAINSS_CONFIG.WORKBOOK_PATH));
                workbook.Close(false, xlsFilePath, true);
            }
        }

        private ProductionXmlFileTemplate GetHKMAINSSXmlFileContent(Worksheet worksheet)
        {
            ProductionXmlFileTemplate template = new ProductionXmlFileTemplate();
            int lastAddedRicNum = 0;
            // the default id value of the first fid node in each ric node
            int lastStartId = 316;
            // the start row of the range which will be copied from the HKMAINSSExcel file to the HKMAINSSXml file
            int startRow = (configObj.HKMAINSS_CONFIG.START_PAGE - 1) * configObj.HKMAINSS_CONFIG.LINE_NUM + 1;
            // the end row of the range which will be copied from the HKMAINSSExcel file to the HKMAINSSXml file
            int endRow = configObj.HKMAINSS_CONFIG.TOTAL_PAGE_NUM * configObj.HKMAINSS_CONFIG.LINE_NUM;
            //copy the records whose Range is from startRow to endRow in the HKMAINSSExcel file
            for (int i = startRow; i <= endRow; i++)
            {
                if ((i - startRow) % configObj.HKMAINSS_CONFIG.LINE_NUM == 0)
                {
                    lastAddedRicNum = template.rics.rics.Count - 1;
                    
                    Reuters.ProcessQuality.ContentAuto.Lib.Ric ric = new Reuters.ProcessQuality.ContentAuto.Lib.Ric();
                    ric.Name = string.Format("HK/MAINSS{0}", (((i - startRow) / configObj.HKMAINSS_CONFIG.LINE_NUM) + configObj.HKMAINSS_CONFIG.START_PAGE).ToString("D2"));
                    template.rics.rics.Add(ric);

                    if (lastAddedRicNum >= 0)
                    {
                        Fid lastFid = new Fid();
                        lastFid.Id = 339;
                        string continuedPageDescription = "Continued on <" + template.rics.rics[lastAddedRicNum + 1].Name + ">";
                        continuedPageDescription = continuedPageDescription.PadLeft(recordTotalLength);
                        lastFid.Value = string.Format("\"{0}\"", continuedPageDescription);
                        template.rics.rics[lastAddedRicNum].fids.Add(lastFid);
                    }
                }

                if ((i - startRow) % configObj.HKMAINSS_CONFIG.LINE_NUM == 3)
                {
                    lastStartId = 315;
                    continue;
                }

                Fid fid = new Fid();
                fid.Id = lastStartId + (i - startRow) % configObj.HKMAINSS_CONFIG.LINE_NUM;
                string colValue = GetFormatedValue(i, worksheet);
                fid.Value = string.Format("\"{0}\"", colValue);

                lastAddedRicNum = template.rics.rics.Count - 1;
                template.rics.rics[lastAddedRicNum].fids.Add(fid);

                if ((i - startRow) % configObj.HKMAINSS_CONFIG.LINE_NUM == (configObj.HKMAINSS_CONFIG.LINE_NUM-1))
                {
                    lastStartId = 316;
                }
            }
            return template;
        }

        private string GetFormatedValue(int row, Worksheet worksheet)
        {
            string value="";
            if (ExcelUtil.GetRange(row, 1, worksheet).Text == null)
            {
                value = " ".PadLeft(recordAlignrightLength);
            }
            else
            {
                string firstCloumText = ExcelUtil.GetRange(row, 1, worksheet).Text.ToString();
                if (firstCloumText.Equals(string.Empty))
                {
                    value = " ".PadLeft(recordAlignrightLength);
                }
                else
                {
                    value = firstCloumText;
                }
            }
            return value;
        }

        #endregion
    }
}
