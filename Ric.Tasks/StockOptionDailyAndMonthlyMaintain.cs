using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Util;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Collections;
using System.Diagnostics;
using Ric.Db.Manager;
using System.Text.RegularExpressions;
using System.Drawing;
using Ric.Db.Info;
using System.Windows.Forms;
using System.ComponentModel;
using Ric.Core;

namespace Ric.Tasks
{

    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_StockOptionDailyAndMonthlyMaintainConfig
    {
        [StoreInDB]
        [Category("File")]
        [Description("The full path of 'KSO STRIKE_MACRO.xls'. E.g.C:\\Korea_Auto\\Option\\KSO STRIKE_MACRO.xls ")]
        [DefaultValue("C:\\Korea_Auto\\Option\\KSO STRIKE_MACRO.xls")]
        public string RequirementFile { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Option\\GEDA\\")]
        [Description("Path for saving generated GEDA files \nE.g. C:\\Korea_Auto\\Option\\GEDA\\ ")]
        public string GEDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Option\\NDA\\")]
        [Description("Path for saving generated NDA files \nE.g. C:\\Korea_Auto\\Option\\NDA\\ ")]
        public string NDA { get; set; }

        [Category("USD OPT")]
        [Description("Value of USD last price today.")]
        public string PriceUSD { get; set; }

    }

    public class StockOptionDailyAndMonthlyMaintain : GeneratorBase
    {
        private KOREA_StockOptionDailyAndMonthlyMaintainConfig configObj = null;
        private ExcelApp excelApp = null;
        private Workbook workbook = null;
        private Hashtable callPutOption = new Hashtable();
        private Hashtable callOption = new Hashtable();
        private Hashtable putOption = new Hashtable();
        private Hashtable stockOptions = new Hashtable();

        private string gatsServer = string.Empty;
        private List<ContractMonth> contractMonths = new List<ContractMonth>();

        //For GATS
        private string ricsFromExcel = string.Empty;
        int retryTime = 2;
        //For USD
        private List<ContractMonth> contractMonthsUSD = new List<ContractMonth>();
        private StockOptionData optionUSD = new StockOptionData();
        private GATSUtil gats = new GATSUtil();
        private string gedaPath = string.Empty;
        private string ndaPath = string.Empty;

        protected override void Initialize()
        {
            base.Initialize();
            isEikonExcelDisable = true;
            configObj = Config as KOREA_StockOptionDailyAndMonthlyMaintainConfig;

            string today = DateTime.Today.ToString("yyyy-MM-dd");
            gedaPath = Path.Combine(configObj.GEDA, today);
            ndaPath = Path.Combine(configObj.NDA, today);

            if (!CheckTaskRequirement())
            {
                string msg = "Error found when initialize stock option.";
                throw new Exception(msg);
            }
            Logger.Log("Initialize - OK!");
            TaskResultList.Add(new TaskResultEntry("Log File", "Log File", Logger.FilePath));
        }


        private bool CheckTaskRequirement()
        {
            //price 
            if (string.IsNullOrEmpty(configObj.PriceUSD))
            {
                MessageBox.Show("Please fill the USD price!", "Error");
                return false;
            }

            Double priceDouble;
            if (!Double.TryParse(configObj.PriceUSD, out priceDouble))
            {
                MessageBox.Show("Bad format for USD price!", "Error");
                return false;
            }

            if (string.IsNullOrEmpty(configObj.RequirementFile))
            {
                string msg = "RequirementPath in configuration can not be blank!";
                MessageBox.Show(msg, "Error");
                return false;
            }

            if (!File.Exists(configObj.RequirementFile))
            {
                string msg = string.Format("Can not find file {0}!", configObj.RequirementFile);
                MessageBox.Show(msg, "Error");
                return false;
            }

            gatsServer = gats.serverIP;

            return true;
        }


        protected override void Start()
        {           
            GenerateBackup();
           // COMAddIn addIn = null;
            InitializeExcel();
            try
            {
                PrepareData();
                GetDataFromGats();
                FormatData();
                WriteBackToRequirement();
                GenerateFiles();
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in running Stock Option task. " + ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception(msg);
            }
            finally
            {
                DisposeExcel();
            }
        }        

        /// <summary>
        /// Backup requirement file into folder named by today.
        /// </summary>
        private void GenerateBackup()
        {
            string sourceFile = configObj.RequirementFile;
            string today = DateTime.Today.ToString("yyyy-MM-dd");
            string destFilePath = Path.Combine(Path.GetDirectoryName(sourceFile), "Backup");
            destFilePath = Path.Combine(destFilePath, today);

            if (!Directory.Exists(destFilePath))
            {
                Directory.CreateDirectory(destFilePath);
            }

            try
            {
                string destFile = Path.Combine(destFilePath, Path.GetFileName(sourceFile));
                if (!File.Exists(destFile))
                {
                    File.Copy(sourceFile, destFile, true);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in generate back up file for KSO STRIKE_MACRO.xls.\r\n" + ex.Message, Logger.LogType.Error);
            }
            Logger.Log(string.Format("Backup - OK!  Backup file: {0} to folder:{1}.", Path.GetFileName(sourceFile), destFilePath));
            TaskResultList.Add(new TaskResultEntry("Backup Folder", "Backup Folder", destFilePath));
            TaskResultList.Add(new TaskResultEntry(Path.GetFileName(configObj.RequirementFile), Path.GetFileNameWithoutExtension(configObj.RequirementFile), configObj.RequirementFile));
        }

        /// <summary>
        /// Initialize excel app.
        /// </summary>
        private void InitializeExcel()
        {
            excelApp = new ExcelApp(false, false);
            if (excelApp == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            
            workbook = ExcelUtil.CreateOrOpenExcelFile(excelApp, configObj.RequirementFile);
            if (workbook == null)
            {
                string msg = string.Format("Error found when openning file {0}!", configObj.RequirementFile);
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        /// <summary>
        /// Get basic infomation from requirement file.
        /// RIC, company code, price and code exsited from KSO sheets.
        /// Last trading day , call put options from mapping table sheet.
        /// </summary>
        private void PrepareData()
        {
            GetMapDataFromDB();
            GetDataFromKSO();
            GetUSDDataFromKSO();
            Logger.Log("Get RICs, Prices and Codes from Excel - OK!");
        }

        private void GetMapDataFromDB()
        {
            GetLastTradingDays();
            GetLastTradingDaysForUSD();
            GetCallPutOptions();
        }

        private void GetLastTradingDays()
        {
            DateTime currentDate = DateTime.Today;
            int currentMonth = currentDate.Month;
            if (KoreaOptionMapManager.CheckLaterThanLastTradingDay(currentDate))
            {
                currentMonth++;
            }
            int currentYear = currentDate.Year;
            for (int i = 0; i < 4; i++)
            {
                if (currentMonth > 12)
                {
                    currentYear++;
                    currentMonth = currentMonth % 12;
                }
                if (i == 3)
                {
                    while (currentMonth % 3 != 0)
                    {
                        currentMonth++;
                    }
                }
                string monthUpper = ((StockOptionMonth)currentMonth).ToString();
                KoreaOptionLastTradingDayInfo lastTradingDayItem = KoreaOptionMapManager.SelectLastTradingDayByYearMonth(currentYear.ToString(), monthUpper);
                if (lastTradingDayItem == null)
                {
                    string msg = string.Format("Can not get last trading day from database. For year; {0}, month: {1}. Please check the Korea Code Map table", currentYear, monthUpper);
                    throw new Exception(msg);
                }
                string lastDate = lastTradingDayItem.LastTradingDay.ToString("dd-MMM-yy");
                int highest = 0;
                int lowest = 0;
                ContractMonth contractMonthItem = new ContractMonth(monthUpper, lastDate, highest, lowest);
                contractMonths.Add(contractMonthItem);
                currentMonth++;
            }
            Logger.Log("Get last trading days from DB - OK!");
        }

        private void GetLastTradingDaysForUSD()
        {
            DateTime currentDate = DateTime.Today;
            int currentMonth = currentDate.Month;
            if (KoreaOptionMapManager.CheckLaterThanLastTradingDayUSD(currentDate))
            {
                currentMonth++;
            }
            int currentYear = currentDate.Year;
            for (int i = 0; i < 4; i++)
            {
                if (currentMonth > 12)
                {
                    currentYear++;
                    currentMonth = currentMonth % 12;
                }
                if (i == 3)
                {
                    while (currentMonth % 3 != 0)
                    {
                        currentMonth++;
                    }
                }
                string monthUpper = ((StockOptionMonth)currentMonth).ToString();
                KoreaOptionLastTradingDayInfo lastTradingDayItem = KoreaOptionMapManager.SelectLastTradingDayByYearMonth(currentYear.ToString(), monthUpper);
                if (lastTradingDayItem == null)
                {
                    string msg = string.Format("Can not get last trading day from database. For year; {0}, month: {1}. Please check the Korea Code Map table", currentYear, monthUpper);
                    throw new Exception(msg);
                }              

                //For USD
                string lastDateUSD = lastTradingDayItem.LastTradingDayForUSDOPT.ToString("dd-MMM-yy");
                ContractMonth contractMonthItemUSD = new ContractMonth(monthUpper, lastDateUSD, 0, 0);
                contractMonthsUSD.Add(contractMonthItemUSD);
                currentMonth++;

            }
            Logger.Log("Get last trading days for USD from DB - OK!");
        }

        private void GetCallPutOptions()
        {
            List<KoreaCodeMapInfo> codeMap = KoreaOptionMapManager.SelectCodeMapByType(KoreaCodeMapType.Call);
            foreach (KoreaCodeMapInfo codeItem in codeMap)
            {
                callPutOption.Add(codeItem.Code, codeItem.Name);
                callOption.Add(codeItem.Name, codeItem.Code);
            }
            codeMap = KoreaOptionMapManager.SelectCodeMapByType(KoreaCodeMapType.Put);
            foreach (KoreaCodeMapInfo codeItem in codeMap)
            {
                callPutOption.Add(codeItem.Code, codeItem.Name);
                putOption.Add(codeItem.Name, codeItem.Code);
            }
            Logger.Log("Get Call/Put Option from DB - OK!");
        }

        private void GetDataFromKSO()
        {
            for (int i = 1; i <= 5; i++)
            {
                Worksheet worksheet = ExcelUtil.GetWorksheet("KSO" + i, workbook);
                if (worksheet == null)
                {
                    throw new Exception(string.Format("Cannot find worksheet {0} in file {1}.", "KSO" + i, configObj.RequirementFile));
                }
                GetRICList(worksheet, ref ricsFromExcel);

            }
            ricsFromExcel = ricsFromExcel.Substring(0, ricsFromExcel.Length - 1);
        }


        private void GetRICList(Worksheet worksheet, ref string rics)
        {
            int col = 0;
            while (true)
            {
                col++;
                string ric = ((Range)worksheet.Cells[1, 3 * col]).Text.ToString().Trim();

                if (string.IsNullOrEmpty(ric) || (!ric.Contains(".")))
                {
                    break;
                }
                rics += ric + ",";

                string companyCode = ((Range)worksheet.Cells[2, 3 * col - 1]).Text.ToString().Trim();
                if (companyCode.Length > 3)
                {
                    companyCode = companyCode.Substring(0, 3);
                }
                int startLine = 10;
                List<StockCodePrice> codePrices = new List<StockCodePrice>();
                while (true)
                {
                    string code = ((Range)worksheet.Cells[startLine, 3 * col - 1]).Text.ToString().Trim();
                    if (string.IsNullOrEmpty(code))
                    {
                        break;
                    }
                    string price = ((Range)worksheet.Cells[startLine, 3 * col]).Text.ToString().Trim();

                    StockCodePrice codePrice = new StockCodePrice(code, price);
                    codePrices.Add(codePrice);
                    startLine++;

                }
                StockOptionData option = new StockOptionData();
                option.RIC = ric;
                option.Column = col;
                option.KSOsheet = worksheet.Name;
                option.CompanyCode = companyCode;
                option.PriceExisted = codePrices;

                List<ContractMonth> contractMonthForOption = new List<ContractMonth>();
                foreach (ContractMonth month in contractMonths)
                {
                    ContractMonth monthItem = new ContractMonth();
                    monthItem.Month = month.Month;
                    monthItem.LastTradeDay = month.LastTradeDay;
                    contractMonthForOption.Add(monthItem);
                }
                option.ContractMonths = contractMonthForOption;
                stockOptions.Add(ric, option);
            }
        }


        /// <summary>
        /// Get last price for each RIC.
        /// highest and lowest values for four months of each RIC. 
        /// </summary>
        private void GetDataFromGats()
        {
            string file = "Tools\\Data2XML\\Data2XML.exe";
            string path = "Tools\\Data2XML\\";
            Process gatsProcess = new Process();
            gatsProcess.StartInfo.FileName = file;
            gatsProcess.StartInfo.WorkingDirectory = path;
            gatsProcess.StartInfo.UseShellExecute = false;
            gatsProcess.StartInfo.CreateNoWindow = true;

            GetLastPriceFromGats(gatsProcess, ricsFromExcel);
            GetAllPricesFromGats(gatsProcess);
            GetAllPricesFromGatsForUSD();

        }

        private void GetLastPriceFromGats(Process gatsProcess, string rics)
        {
            string fids = "TRDPRC_1";
            string arg = string.Format("/c-quiet -dbout -raw_enum_vals  -ph {0} -pn IDN_SELECTFEED -rics \"{1}\" -fids \"{2}\"", gatsServer, rics, fids);
            string result = GetGatsResponse(gatsProcess, arg);
            if (string.IsNullOrEmpty(result))
            {
                throw new Exception("GATS returns no response.");
            }

            string pattern = string.Format(@"\r\n(?<RIC>[0-9]+\.KS) +{0} +(?<LastPrice>.*?)\r\n", fids);
            Regex regex = new Regex(pattern);
            MatchCollection match = regex.Matches(result);
            int rightCount = match.Count;
            foreach (Match m in match)
            {
                string ric = m.Groups["RIC"].Value;
                if (stockOptions.Contains(ric))
                {
                    try
                    {
                        StockOptionData option = (StockOptionData)stockOptions[ric];
                        ((StockOptionData)stockOptions[ric]).LastPrice = Convert.ToInt32(m.Groups["LastPrice"].Value);
                    }
                    catch (Exception ex)
                    {
                        string msg = string.Format("Error found when getting last price from GATS. For RIC:{0}. Error message:", ric, ex.Message);
                        Logger.Log(msg, Logger.LogType.Error);
                        rightCount--;
                    }
                }
            }
            if ((rightCount != rics.Split(',').Length))
            {
                VerifyAllLastPrice(gatsProcess);
            }
            Logger.Log("Get all last price from GATS - OK!");
        }

        /// <summary>
        /// Check all RICs contains last price.
        /// If not, get last price from GATS again.
        /// </summary>
        /// <param name="gatsProcess">gats process</param>       
        private void VerifyAllLastPrice(Process gatsProcess)
        {
            List<string> ricsLeft = new List<string>();
            foreach (StockOptionData option in stockOptions)
            {
                if (option.LastPrice == 0)
                {
                    ricsLeft.Add(option.RIC);
                }
            }
            if (ricsLeft.Count > 0 && retryTime-- > 0)
            {
                string rics = string.Join(",", ricsLeft.ToArray());
                GetLastPriceFromGats(gatsProcess, rics);
                return;
            }

            foreach (string ric in ricsLeft)
            {
                string msg = string.Format("After tried 2 times. We still can't get last price for RIC:{0}. Remove it from today's option list.", ric);
                Logger.Log(msg, Logger.LogType.Error);
                stockOptions.Remove(ric);
            }
        }


        private void GetAllPricesFromGats(Process gatsProcess)
        {
            int count = 15;
            string[] ricArray = ricsFromExcel.Split(',');
            string ricPara = string.Empty;
            for (int i = 0; i < ricArray.Length; i++)
            {
                string ric = ricArray[i].Replace(".KS", "");
                for (int j = 0; j <= count; j++)
                {
                    if (j > 9)
                    {
                        //OptionPageNumber pageNum = (OptionPageNumber)j;
                        string pageNum = Convert.ToString(j, 16).ToUpper();
                        ricPara += string.Format("{0}#{1}*.KS,", pageNum, ric);
                        continue;
                    }
                    ricPara += string.Format("{0}#{1}*.KS,", j, ric);
                }
            }
            ricPara = ricPara.Substring(0, ricPara.Length - 1);

            string fidPara = "LONGLINK2,LONGLINK4,LONGLINK6,LONGLINK8,LONGLINK10,LONGLINK12,LONGLINK14,LONGNEXTLR";
            string arg = string.Format("/c-quiet -dbout -raw_enum_vals  -ph {2} -pn IDN_SELECTFEED -rics \"{0}\" -fids \"{1}\"", ricPara, fidPara, gatsServer);
            string result = GetGatsResponse(gatsProcess, arg);
            if (string.IsNullOrEmpty(result))
            {
                throw new Exception("Can not get response from GATS!");
            }

            result += GetExtraPrices(count, result, gatsProcess);

            FormatOptionPriceRange(result);

            Logger.Log("Get all prices range from GATS - OK!");
        }

        private string GetExtraPrices(int count, string result, Process gatsProcess)
        {
            int startNum = count + 1;
            string pageNum1 = Convert.ToString(startNum, 16).ToUpper();
            string checkSign = string.Format("LONGNEXTLR +{0}#(?<RIC>[0-9]+)\\*\\.KS", pageNum1);
            Regex regex = new Regex(checkSign);
            MatchCollection matches = regex.Matches(result);
            if (matches.Count == 0)
            {
                return "";
            }
            int endNum = count + 10;
            string ricPara = string.Empty;
            for (int i = 0; i < matches.Count; i++)
            {
                Match m = matches[i];
                string ric = m.Groups["RIC"].Value;
                for (int j = startNum; j <= endNum; j++)
                {
                    //OptionPageNumber pageNum = (OptionPageNumber)j;
                    string pageNum = Convert.ToString(j, 16).ToUpper();
                    ricPara += string.Format("{0}#{1}*.KS,", pageNum, ric);
                }
            }
            ricPara = ricPara.Substring(0, ricPara.Length - 1);

            string fidPara = "LONGLINK2,LONGLINK4,LONGLINK6,LONGLINK8,LONGLINK10,LONGLINK12,LONGLINK14,LONGNEXTLR";
            string arg = string.Format("/c-quiet -dbout -raw_enum_vals  -ph {2} -pn IDN_SELECTFEED -rics \"{0}\" -fids \"{1}\"", ricPara, fidPara, gatsServer);
            string gatsresult = GetGatsResponse(gatsProcess, arg);
            if (string.IsNullOrEmpty(result))
            {
                return "";
            }

            gatsresult += GetExtraPrices(endNum, gatsresult, gatsProcess);
            return gatsresult;
        }

        private void FormatOptionPriceRange(string result)
        {
            string[] ricArray = ricsFromExcel.Split(',');
            for (int i = 0; i < ricArray.Length; i++)
            {
                string ricKey = ricArray[i];
                string ric = ricKey.Replace(".KS", "");
                string pattern = string.Format("#{0}.* +LONGLINK\\d+ +(?<Code>[A-Z]+)(?<Price>\\d+)(?<Month>\\w+).KS", ric);
                Regex regex = new Regex(pattern);
                MatchCollection matches = regex.Matches(result);
                StockOptionData option = stockOptions[ricKey] as StockOptionData;
                foreach (Match match in matches)
                {
                    string codeItem = match.Groups["Code"].Value;
                    string priceItem = match.Groups["Price"].Value;
                    string monthItem = match.Groups["Month"].Value;

                    string month = callPutOption[monthItem.Substring(0, 1)].ToString();
                    int value = Convert.ToInt32(priceItem);

                    ModifyContractMonth(month, value, option);
                }
            }
        }

        private void ModifyContractMonth(string month, int value, StockOptionData option)
        {
            for (int i = 0; i < option.ContractMonths.Count; i++)
            {
                if (option.ContractMonths[i].Month == month)
                {
                    if (option.ContractMonths[i].Highest <= value || option.ContractMonths[i].Highest == 0)
                    {
                        option.ContractMonths[i].Highest = value;
                    }
                    if (option.ContractMonths[i].Lowest >= value || option.ContractMonths[i].Lowest == 0)
                    {
                        option.ContractMonths[i].Lowest = value;
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// Give GATS a command line. Get the response.
        /// </summary>
        /// <param name="gatsProcess">GATS process</param>
        /// <param name="command">command</param>
        /// <returns>response</returns>
        private string GetGatsResponse(Process gatsProcess, string command)
        {
            if (gatsProcess == null)
            {
                return null;
            }
            try
            {
                int timeout = 20;
                string filePath = GetOutputFilePath();
                string fileName = "GATS_Response.txt";
                filePath = Path.Combine(filePath, fileName);

                command += string.Format(" -tee \"{0}\"", filePath);

                int retry = 3;
                bool success = false;
                while (!success && retry-- > 0)
                {
                    gatsProcess.StartInfo.Arguments = command;
                    gatsProcess.Start();
                    success = gatsProcess.WaitForExit(timeout * 1000);
                    if (!success)
                    {
                        gatsProcess.Kill();
                    }
                }
                string response = null;
                if (success)
                {
                    response = File.ReadAllText(filePath);
                }
                return response;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Find the closest value in price range exsited.
        /// Calculate the 9 highlight prices.
        /// Find their code.
        /// If code not existed, generate new code for it. Refresh the price and code exsited.
        /// </summary>
        private void FormatData()
        {
            foreach (DictionaryEntry de in stockOptions)
            {
                StockOptionData option = de.Value as StockOptionData;
                option.FindClosestValue();
                option.FormatPriceExisted();
                option.FormatPricePredictions();
            }

            //For USD
            optionUSD.FindClosestValueUSD();
            optionUSD.FormatPriceExistedUSD();
            optionUSD.FormatPricePredictionsUSD();

            Logger.Log("Format prices and predictions - OK!");
        }

        /// <summary>
        /// Write the last price, 4 contract months, 9 highlight price, price and code range back to requirement.
        /// </summary>
        private void WriteBackToRequirement()
        {
            foreach (DictionaryEntry de in stockOptions)
            {
                StockOptionData option = de.Value as StockOptionData;
                Worksheet worksheet = ExcelUtil.GetWorksheet(option.KSOsheet, workbook);
                if (worksheet == null)
                {
                    Logger.Log(string.Format("Can not find worksheet {0} in workbook {1}", option.KSOsheet, workbook.Name));
                    continue;
                }

                WriteLastPrice(worksheet, option);
                WritePriceExsited(worksheet, option);

            }

            //For USD
            Worksheet worksheetUSD = ExcelUtil.GetWorksheet("KSO6", workbook);
            WriteLastPrice(worksheetUSD, optionUSD);
            WritePriceExsited(worksheetUSD, optionUSD);

            Logger.Log("Write data back to Excel - OK!");

        }

        private void WriteLastPrice(Worksheet worksheet, StockOptionData option)
        {
            if (option.Column < 0)
            {
                return;
            }
            worksheet.Cells[5, 3 * option.Column - 1] = "Auto";
            worksheet.Cells[5, 3 * option.Column] = option.LastPrice.ToString();
        }

        private void WriteContractMonth(Worksheet worksheet, StockOptionData option)
        {
            worksheet.Cells[5, 3 * option.Column] = "Contract Months";
            for (int i = 0; i < option.ContractMonths.Count; i++)
            {
                worksheet.Cells[6 + i, 3 * option.Column] = option.ContractMonths[i].LastTradeDay;
            }
        }

        private void WritePricePredition(Worksheet worksheet, StockOptionData option)
        {

            for (int i = 0; i < option.PricePredictions.Count; i++)
            {
                StockCodePrice codePrice = option.PricePredictions[i];

                worksheet.Cells[11 + i, 3 * option.Column - 1] = codePrice.Code;
                worksheet.Cells[11 + i, 3 * option.Column] = codePrice.Price;
                ((Range)worksheet.Cells[11 + i, 3 * option.Column]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.DodgerBlue);
            }
            ((Range)worksheet.Cells[15, 3 * option.Column]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
        }

        private void WritePriceExsited(Worksheet worksheet, StockOptionData option)
        {
            for (int i = 0; i < option.PriceExisted.Count; i++)
            {
                StockCodePrice codePrice = option.PriceExisted[i];

                if (codePrice.Status.Equals(StockPriceStatus.New))
                {
                    worksheet.Cells[10 + i, 3 * option.Column - 2] = "New";
                    ((Range)worksheet.Cells[10 + i, 3 * option.Column - 2]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
                }
                else
                {
                    worksheet.Cells[10 + i, 3 * option.Column - 2] = "";
                }

                ((Range)worksheet.Cells[10 + i, 3 * option.Column - 1]).NumberFormat = "@";
                ((Range)worksheet.Cells[10 + i, 3 * option.Column]).ClearFormats();
                ((Range)worksheet.Cells[10 + i, 3 * option.Column]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                worksheet.Cells[10 + i, 3 * option.Column - 1] = codePrice.Code;
                worksheet.Cells[10 + i, 3 * option.Column] = codePrice.Price;

                if (i == option.ClosestIndex)
                {
                    ((Range)worksheet.Cells[10 + i, 3 * option.Column]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                }
                else if (Math.Abs(option.ClosestIndex - i) < 5)
                {
                    ((Range)worksheet.Cells[10 + i, 3 * option.Column]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.DodgerBlue);
                }
            }

            for (int i = option.PriceExisted.Count; i < option.PriceExisted.Count + 20; i++)
            {
                worksheet.Cells[10 + i, 3 * option.Column - 2] = "";
                worksheet.Cells[10 + i, 3 * option.Column - 1] = "";
                worksheet.Cells[10 + i, 3 * option.Column] = "";
            }
        }


        /// <summary>
        /// Generate GEDA, NDA files, etc.
        /// </summary>
        private void GenerateFiles()
        {
            bool normalGEDA = false;
            bool newGEDA = false;

            string fileNameGEDA = string.Format("KR_OPT_STOCK_{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
            string fileNameNew = string.Format("KR_OPT_NEW_{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
            string fileNameNDA = string.Format("KR_OPT_LTD&LOTSIZE_{0}.csv", DateTime.Today.ToString("ddMMMyyyy").ToUpper());

            string filePathGEDA = Path.Combine(gedaPath, fileNameGEDA);
            string filePathNew = Path.Combine(gedaPath, fileNameNew);
            string filePathNDA = Path.Combine(ndaPath, fileNameNDA);

            string titleGEDA = "EXL_NAME\tEXPIR_DATE\tSTRIKE_PRICE\tINP_EXCERCISE_CODE\tINP_INCL_CHAIN";
            string titleNDA = "RIC,DERIVATIVES LAST TRADING DAY,DERIVATIVES LOT SIZE";


            foreach (DictionaryEntry de in stockOptions)
            {
                StockOptionData option = de.Value as StockOptionData;
                GeneratePriceOutput(option, ref normalGEDA, ref newGEDA, filePathGEDA, titleGEDA, filePathNDA, titleNDA, filePathNew);
            }

            if (newGEDA)
            {
                TaskResultList.Add(new TaskResultEntry(fileNameNew, "New Added", filePathNew));
            }
            if (normalGEDA)
            {
                TaskResultList.Add(new TaskResultEntry(fileNameGEDA, "GEDA File", filePathGEDA));
                TaskResultList.Add(new TaskResultEntry(fileNameNDA, "NDA File", filePathNDA));
            }

            GenerateUSDFiles();           
            Logger.Log("Generate output files - OK!");
        }

        private void GenerateUSDFiles()
        {
            bool normalGEDA = false;
            bool newGEDA = false;

            string fileNameGEDA = string.Format("KR_OPT_KRW_{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
            string fileNameNew = string.Format("KR_OPT_KRW_NEW_{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
            string fileNameNDA = string.Format("KR_OPT_KRW_LTD&LOTSIZE_{0}.csv", DateTime.Today.ToString("ddMMMyyyy").ToUpper());

            string filePathGEDA = Path.Combine(gedaPath, fileNameGEDA);
            string filePathNew = Path.Combine(gedaPath, fileNameNew);
            string filePathNDA = Path.Combine(ndaPath, fileNameNDA);

            string titleGEDA = "EXL_NAME\tEXPIR_DATE\tSTRIKE_PRICE\tINP_EXCERCISE_CODE\tINP_INCL_CHAIN";
            string titleNDA = "RIC,DERIVATIVES LAST TRADING DAY,DERIVATIVES LOT SIZE";

            GeneratePriceOutput(optionUSD, ref normalGEDA, ref newGEDA, filePathGEDA, titleGEDA, filePathNDA, titleNDA, filePathNew);


            if (newGEDA)
            {
                TaskResultList.Add(new TaskResultEntry(fileNameNew, "KRW New Added", filePathNew));
            }
            if (normalGEDA)
            {
                TaskResultList.Add(new TaskResultEntry(fileNameGEDA, "KRW GEDA File", filePathGEDA));
                TaskResultList.Add(new TaskResultEntry(fileNameNDA, "KRW NDA File", filePathNDA));
            }
        }

        private void GeneratePriceOutput(StockOptionData option, ref bool normalGEDA, ref bool newGEDA, string filePathGEDA, string titleGEDA, string filePathNDA, string titleNDA, string filePathNew)
        {
            foreach (ContractMonth month in option.ContractMonths)
            {
                foreach (StockCodePrice codePrice in option.PricePredictions)
                {
                    if (codePrice.Price <= month.Highest && codePrice.Price >= month.Lowest)
                    {
                        continue;
                    }

                    string monthPutCode = putOption[month.Month].ToString();
                    string monthCallCode = callOption[month.Month].ToString();

                    StockOptionOutput outputItemCall = new StockOptionOutput(codePrice.Price, codePrice.Code, month.LastTradeDay, monthCallCode, option.CompanyCode);
                    StockOptionOutput outputItemPut = new StockOptionOutput(codePrice.Price, codePrice.Code, month.LastTradeDay, monthPutCode, option.CompanyCode);
                    string formatedDate = DateTime.Parse(month.LastTradeDay).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));

                    string recordCall = string.Format("{0}\t{1}\t{2}\t{3}\t{4}", outputItemCall.Name, outputItemCall.ExpireDate, outputItemCall.StrikePrice, outputItemCall.Code, outputItemCall.Chain);
                    string recordPut = string.Empty;
                    //string.Format("{0}\t{1}\t{2}\t{3}\t{4}", outputItemPut.Name, outputItemPut.ExpireDate, outputItemPut.StrikePrice, outputItemPut.Code, outputItemPut.Chain);

                    FileUtil.WriteSingleLine(filePathGEDA, recordCall, titleGEDA);
                    normalGEDA = true;

                    if (codePrice.Status == StockPriceStatus.New)
                    {
                        FileUtil.WriteSingleLine(filePathNew, recordCall, titleGEDA);
                        newGEDA = true;
                    }
                    recordCall = string.Format("{0},{1},{2}", outputItemCall.RIC, formatedDate, outputItemCall.LotSize);
                    recordPut = string.Format("{0},{1},{2}", outputItemPut.RIC, formatedDate, outputItemPut.LotSize);

                    FileUtil.WriteSingleLine(filePathNDA, recordCall + "\r\n" + recordPut, titleNDA);
                }
            }
        }

        /// <summary>
        /// Dispose excel app.
        /// </summary>
        private void DisposeExcel()
        {
            if (excelApp == null)
            {
                return;
            }
            excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
            excelApp.ExcelAppInstance.DisplayAlerts = false;
            if (workbook != null)
            {
                workbook.Save();
                workbook.Close(true, Type.Missing, Type.Missing);
            }           
            excelApp.Dispose();
        }


        #region  For USD

        private void GetUSDDataFromKSO()
        {
            Worksheet worksheet = ExcelUtil.GetWorksheet("KSO6", workbook);
            if (worksheet == null)
            {
                throw new Exception(string.Format("Cannot find worksheet {0} in file {1}.", "KSO6", configObj.RequirementFile));
            }

            int startLine = 10;
            List<StockCodePrice> codePrices = new List<StockCodePrice>();
            while (true)
            {
                string code = ((Range)worksheet.Cells[startLine, 2]).Text.ToString().Trim();
                if (string.IsNullOrEmpty(code))
                {
                    break;
                }
                string price = ((Range)worksheet.Cells[startLine, 3]).Text.ToString().Trim();

                StockCodePrice codePrice = new StockCodePrice(code, price);
                codePrices.Add(codePrice);
                startLine++;

            }
            optionUSD.RIC = "KRW";
            optionUSD.Column = 1;
            optionUSD.KSOsheet = worksheet.Name;
            optionUSD.CompanyCode = "KRW";
            optionUSD.PriceExisted = codePrices;
            optionUSD.ContractMonths = contractMonthsUSD;
            optionUSD.LastPrice = Convert.ToInt32(Math.Floor(Double.Parse(configObj.PriceUSD)));
        }

        private void GenerateGedaFileForUSD(string filePathUSD, string titleGEDA)
        {
            List<string> data = new List<string>();

            foreach (ContractMonth month in contractMonthsUSD)
            {
                foreach (StockCodePrice codePrice in optionUSD.PricePredictions)
                {
                    StockOptionOutput outputItem = new StockOptionOutput(codePrice.Price, codePrice.Code, month.LastTradeDay, "", optionUSD.CompanyCode);
                    string record = string.Format("{0}\t{1}\t{2}\t{3}\t{4}", outputItem.Name, outputItem.ExpireDate, outputItem.StrikePrice, outputItem.Code, outputItem.Chain);
                    data.Add(record);
                }
            }
            FileUtil.WriteOutputFile(filePathUSD, data, titleGEDA, WriteMode.Overwrite);
        }

        private void GetAllPricesFromGatsForUSD()
        {
            string rics = "0#KRW*:,1#KRW*:,2#KRW*:,3#KRW*:,4#KRW*:,5#KRW*:,6#KRW*:,7#KRW*:,8#KRW*:,9#KRW*:";
            string fids = "LINK_2,LINK_4,LINK_6,LINK_8,LINK_10,LINK_12,LINK_14,NEXT_LR";
            string response = gats.GetGatsResponse(rics, fids);
            response += GetExtraPricesUSD(9, response);

            FormatOptionPriceRangeUSD(response);

        }

        private string GetExtraPricesUSD(int count, string result)
        {
            int startNum = count + 1;
            string reponse = string.Empty;
            reponse += GetExtraPricesUsdWithTwoType(startNum, true, result);
            reponse += GetExtraPricesUsdWithTwoType(startNum, false, result);
            return reponse;
        }

        /// <summary>
        /// Use Demical or Hex NEXT_LR to find if exsits next page.
        /// </summary>
        /// <param name="numCode">next page number</param>
        /// <param name="aryFlag">true represents Demical, false represents Hex</param>
        /// <param name="result">last response of GATS</param>
        /// <returns>output result</returns>
        private string GetExtraPricesUsdWithTwoType(int numCode, bool aryFlag, string result)
        {
            string startNum = aryFlag ? numCode.ToString() : Convert.ToString(numCode, 16).ToUpper();
            string checkSign = string.Format("NEXT_LR +{0}#KRW\\*:", startNum);
            Regex regex = new Regex(checkSign);
            Match match = regex.Match(result);
            if (!match.Success)
            {
                return "";
            }

            List<string> ricList = new List<string>();
            for (int i = numCode; i < numCode + 10; i++)
            {
                string ricPre = aryFlag ? i.ToString() : Convert.ToString(i, 16);
                string ric = string.Format("{0}#KRW*:", ricPre);
                ricList.Add(ric);
            }

            string rics = string.Join(",", ricList.ToArray());
            string fids = "LINK_2,LINK_4,LINK_6,LINK_8,LINK_10,LINK_12,LINK_14,NEXT_LR";
            string output = gats.GetGatsResponse(rics, fids);
            output += GetExtraPricesUSD(numCode + 10, output);
            return output;
        }

        private void FormatOptionPriceRangeUSD(string result)
        {
            string pattern = "#KRW\\*: +LINK_\\d+ +KRW(?<Price>\\d+)(?<Month>\\w+)\r\n";
            Regex regex = new Regex(pattern);
            MatchCollection matches = regex.Matches(result);

            foreach (Match match in matches)
            {
                string priceItem = match.Groups["Price"].Value;
                string monthItem = match.Groups["Month"].Value;

                string month = callPutOption[monthItem.Substring(0, 1)].ToString();
                int value = Convert.ToInt32(priceItem);

                ModifyContractMonth(month, value, optionUSD);
            }
        }


        #endregion


    }


    public class StockOptionData
    {
        public int LastPrice { get; set; }
        public string RIC { get; set; }
        public string CompanyCode { get; set; }
        public int Column { get; set; }
        public int ClosestValue { get; set; }
        public int ClosestIndex { get; set; }
        public string KSOsheet { set; get; }
        public List<ContractMonth> ContractMonths { get; set; }
        public List<StockCodePrice> PricePredictions { get; set; }
        public List<StockCodePrice> PriceExisted { get; set; }
        public List<StockOptionOutput> PriceOutPut { get; set; }
        public bool CodeChanged { get; set; }

        public StockOptionData()
        {
            this.LastPrice = 0;
            this.RIC = string.Empty;
            this.CompanyCode = string.Empty;
            this.ClosestValue = -1;
            this.Column = -1;
            this.KSOsheet = string.Empty;
            this.ContractMonths = new List<ContractMonth>();
            this.PricePredictions = new List<StockCodePrice>();
            this.PriceExisted = new List<StockCodePrice>();
            this.CodeChanged = false;

        }

        public void FindClosestValue()
        {
            if (this.LastPrice <= 0)
            {
                return;
            }
            int lastPrice = this.LastPrice;
            //Requirement contains no prices and codes for this option
            if (this.PriceExisted == null || this.PriceExisted.Count == 0)
            {
                this.ClosestValue = GenerateClosestValue(lastPrice);
                this.ClosestIndex = -1;
                return;
            }

            int index = -1;
            int absValue = 9999999;

            for (int i = 0; i < this.PriceExisted.Count; i++)
            {
                if (this.PriceExisted[i].Price <= 0)
                {
                    continue;
                }
                int tempValue = this.PriceExisted[i].Price;

                if (Math.Abs(tempValue - lastPrice) < absValue)
                {
                    index = i;
                    absValue = Math.Abs(tempValue - lastPrice);
                }
            }
            this.ClosestValue = this.PriceExisted[index].Price;
            this.ClosestIndex = index;
        }

        public int GenerateClosestValue(int lastPrice)
        {
            if (lastPrice <= 0)
            {
                return 0;
            }
            int increment = 0;
            int level = 0;
            if (lastPrice >= 0 && lastPrice < 5000)
            {
                level = 2000;
                increment = 100;
            }
            else if (lastPrice >= 5000 && lastPrice < 10000)
            {
                level = 5000;
                increment = 200;
            }
            else if (lastPrice >= 10000 && lastPrice < 20000)
            {
                level = 10000;
                increment = 500;
            }
            else if (lastPrice >= 20000 && lastPrice < 50000)
            {
                level = 20000;
                increment = 1000;
            }
            else if (lastPrice >= 50000 && lastPrice < 100000)
            {
                level = 50000;
                increment = 2000;
            }
            else if (lastPrice >= 100000 && lastPrice < 200000)
            {
                level = 100000;
                increment = 5000;
            }
            else if (lastPrice >= 200000 && lastPrice < 500000)
            {
                level = 200000;
                increment = 10000;
            }
            else if (lastPrice >= 500000 && lastPrice < 1000000)
            {
                level = 500000;
                increment = 20000;
            }
            else if (lastPrice >= 1000000 && lastPrice < 2000000)
            {
                level = 1000000;
                increment = 50000;
            }

            int closestValue = level + ((lastPrice - level) / increment + (((lastPrice - level) % increment) / (increment / 2))) * increment;
            return closestValue;
        }

        public void FormatPriceExisted()
        {
            if (this.PriceExisted.Count == 0)
            {
                GeneratePriceExsited();
                return;
            }

            //Need price up
            if (this.ClosestIndex < 4)
            {
                int startCode = FindMaxCode();
                int lastPriceValue = PriceExisted[0].Price;
                for (int i = 0; i < 4 - ClosestIndex; i++)
                {
                    startCode++;
                    string code = string.Format("{0,3:000}", startCode);
                    lastPriceValue = GenerateOnePrice(lastPriceValue, true, OptionType.StockOption);
                    StockCodePrice codePriceNew = new StockCodePrice(code, lastPriceValue.ToString());
                    codePriceNew.Status = StockPriceStatus.New;
                    PriceExisted.Insert(0, codePriceNew);
                }
                ClosestIndex = 4;
                CodeChanged = true;
            }

            //Need price down
            else if (this.ClosestIndex > this.PriceExisted.Count - 5)
            {
                int codeDistance = 5 - (this.PriceExisted.Count - this.ClosestIndex);
                int startCode = FindMaxCode();
                int lastPriceValue = PriceExisted[PriceExisted.Count - 1].Price;                
                for (int i = 0; i < codeDistance; i++)
                {
                    startCode++;
                    string code = string.Format("{0,3:000}", startCode);
                    lastPriceValue = GenerateOnePrice(lastPriceValue, false, OptionType.StockOption);
                    StockCodePrice codePriceNew = new StockCodePrice(code, lastPriceValue.ToString());
                    codePriceNew.Status = StockPriceStatus.New;
                    PriceExisted.Add(codePriceNew);
                }
                ClosestIndex = PriceExisted.Count - 5;
                CodeChanged = true;
            }

        }

        public void FormatPricePredictions()
        {
            for (int i = -4; i <= 4; i++)
            {
                int pos = ClosestIndex + i;
                StockCodePrice codePriceExist = PriceExisted[pos];
                PricePredictions.Add(codePriceExist);
            }
        }

        private int FindMaxCode()
        {
            int maxCode = 0;
            foreach (StockCodePrice codePrice in PriceExisted)
            {
                int tempCode = Convert.ToInt32(codePrice.Code.TrimStart('0'));
                maxCode = tempCode > maxCode ? tempCode : maxCode;
            }
            return maxCode;
        }

        public void GeneratePriceExsited()
        {
            if (this.ClosestValue <= 0)
            {
                return;
            }
            int lastPriceValue = this.ClosestValue;

            int startCode = 5;
            this.ClosestIndex = startCode - 1;
            string code = string.Format("{0,3:000}", startCode);
            StockCodePrice codePrice = new StockCodePrice("005", lastPriceValue.ToString());
            codePrice.Status = StockPriceStatus.New;
            this.PriceExisted.Add(codePrice);

            for (int i = 1; i < 5; i++)
            {
                lastPriceValue = GenerateOnePrice(lastPriceValue, true, OptionType.StockOption);
                code = string.Format("{0,3:000}", startCode + i);
                StockCodePrice codePriceAdd = new StockCodePrice(code, lastPriceValue.ToString());
                codePriceAdd.Status = StockPriceStatus.New;
                this.PriceExisted.Insert(0, codePriceAdd);
            }

            lastPriceValue = this.ClosestValue;
            for (int i = 1; i < 5; i++)
            {
                lastPriceValue = GenerateOnePrice(lastPriceValue, false, OptionType.StockOption);
                code = string.Format("{0,3:000}", startCode - i);
                StockCodePrice codePriceAdd = new StockCodePrice(code, lastPriceValue.ToString());
                codePriceAdd.Status = StockPriceStatus.New;
                this.PriceExisted.Add(codePriceAdd);
            }

            CodeChanged = true;
        }

        private int GenerateOnePrice(int lastPriceValue, bool direction, OptionType optionType)
        {
            int increment = 0;
            if (direction)
            {
                if (optionType.Equals(OptionType.StockOption))
                {
                    increment = getIncrementAdd(lastPriceValue);
                }
                else if (optionType.Equals(OptionType.USDOpiton))
                {
                    increment = 10;
                }
                return lastPriceValue + increment;
            }
            else
            {
                if (optionType.Equals(OptionType.StockOption))
                {
                    increment = getIncrementMinus(lastPriceValue);
                }
                else if (optionType.Equals(OptionType.USDOpiton))
                {
                    increment = -10;
                }
                return lastPriceValue + increment;
            }
        }

        private int getIncrementAdd(int nowData)
        {
            int increment = 0;
            if (nowData >= 0 && nowData < 5000)
            {
                increment = 100;
            }
            if (nowData >= 5000 && nowData < 10000)
            {
                increment = 200;
            }
            if (nowData >= 10000 && nowData < 20000)
            {
                increment = 500;
            }
            if (nowData >= 20000 && nowData < 50000)
            {
                increment = 1000;
            }
            if (nowData >= 50000 && nowData < 100000)
            {
                increment = 2000;
            }
            if (nowData >= 100000 && nowData < 200000)
            {
                increment = 5000;
            }
            if (nowData >= 200000 && nowData < 500000)
            {
                increment = 10000;
            }
            if (nowData >= 500000 && nowData < 1000000)
            {
                increment = 20000;
            }
            if (nowData >= 1000000 && nowData < 2000000)
            {
                increment = 50000;
            }
            return increment;
        }

        private int getIncrementMinus(int nowData)
        {
            int increment = 0;
            if (nowData > 0 && nowData <= 5000)
            {
                increment = -100;
            }
            if (nowData > 5000 && nowData <= 10000)
            {
                increment = -200;
            }
            if (nowData > 10000 && nowData <= 20000)
            {
                increment = -500;
            }
            if (nowData > 20000 && nowData <= 50000)
            {
                increment = -1000;
            }
            if (nowData > 50000 && nowData <= 100000)
            {
                increment = -2000;
            }
            if (nowData > 100000 && nowData <= 200000)
            {
                increment = -5000;
            }
            if (nowData > 200000 && nowData <= 500000)
            {
                increment = -10000;
            }
            if (nowData > 500000 && nowData <= 1000000)
            {
                increment = -20000;
            }
            if (nowData > 1000000 && nowData <= 2000000)
            {
                increment = -50000;
            }
            return increment;
        }


        //For USD Option
        public void FindClosestValueUSD()
        {
            if (this.LastPrice <= 0)
            {
                return;
            }
            int lastPrice = this.LastPrice;

            int index = -1;
            int absValue = 999999;

            for (int i = 0; i < this.PriceExisted.Count; i++)
            {
                if (this.PriceExisted[i].Price <= 0)
                {
                    continue;
                }
                int tempValue = this.PriceExisted[i].Price;

                if (Math.Abs(tempValue - lastPrice) <= absValue)
                {
                    index = i;
                    absValue = Math.Abs(tempValue - lastPrice);
                }
            }
            this.ClosestValue = this.PriceExisted[index].Price;
            this.ClosestIndex = index;
        }

        public void FormatPriceExistedUSD()
        {
            if (this.PriceExisted.Count == 0)
            {
                return;
            }

            //Need price up
            if (this.ClosestIndex < 3)
            {
                int lastPriceValue = PriceExisted[0].Price;
                for (int i = 0; i < 3 - ClosestIndex; i++)
                {
                    lastPriceValue = GenerateOnePrice(lastPriceValue, false, OptionType.USDOpiton);
                    int startCode = lastPriceValue / 10;
                    string code = string.Format("{0,3:000}", startCode);
                    StockCodePrice codePriceNew = new StockCodePrice(code, lastPriceValue.ToString());
                    codePriceNew.Status = StockPriceStatus.New;
                    PriceExisted.Insert(0, codePriceNew);
                }
                ClosestIndex = 3;
                CodeChanged = true;
            }

            //Need price down
            else if (this.ClosestIndex > this.PriceExisted.Count - 4)
            {
                int lastPriceValue = PriceExisted[PriceExisted.Count - 1].Price;
                for (int i = 0; i < 3 - ClosestIndex; i++)
                {
                    lastPriceValue = GenerateOnePrice(lastPriceValue, true, OptionType.USDOpiton);
                    int startCode = lastPriceValue / 10;
                    string code = string.Format("{0,3:000}", startCode);
                    StockCodePrice codePriceNew = new StockCodePrice(code, lastPriceValue.ToString());
                    codePriceNew.Status = StockPriceStatus.New;
                    PriceExisted.Add(codePriceNew);
                }
                ClosestIndex = PriceExisted.Count - 4;
                CodeChanged = true;
            }

        }

        public void FormatPricePredictionsUSD()
        {
            for (int i = -3; i <= 3; i++)
            {
                int pos = ClosestIndex + i;
                StockCodePrice codePriceExist = PriceExisted[pos];
                PricePredictions.Add(codePriceExist);
            }
        }
    }

    public class StockOptionOutput
    {
        //For GEDA
        public string Name { get; set; }
        public string ExpireDate { get; set; }
        public string StrikePrice { get; set; }
        public string Code { get; set; }
        public string Chain { get; set; }

        //For NDA
        public string RIC { get; set; }
        public string LotSize { get; set; }

        public StockPriceStatus Status { get; set; }

        public StockOptionOutput(int price, string code, string expireDate, string optionCode, string companyCode)
        {   //KS200262U3.KS
            expireDate = DateTime.Parse(expireDate).ToString("dd-MMM-yy", new CultureInfo("en-US"));
            if (companyCode == "KRW")
            {
                string yearSuffix = expireDate.Substring(expireDate.Length - 1);
                this.RIC = string.Format("{0}{1}{2}{3}", companyCode, price, optionCode, yearSuffix);
                this.Name = "KOFEX_EQO_KRW";
                this.LotSize = "10000";
            }
            else
            {
                string yearSuffix = expireDate.Substring(expireDate.Length - 1);
                this.RIC = string.Format("{0}{1}{2}{3}.KS", companyCode, price, optionCode, yearSuffix);
                this.Name = string.Format("KSO_EQO_{0}", companyCode);
                this.LotSize = "10";
            }
            this.ExpireDate = expireDate;
            this.StrikePrice = price.ToString();
            this.Code = code;
            this.Chain = "Y";
            
        }
    }

    public class ContractMonth
    {
        public string Month { get; set; }
        public string LastTradeDay { get; set; }
        public int Highest { get; set; }
        public int Lowest { get; set; }

        public ContractMonth()
        {
            this.Month = string.Empty;
            this.LastTradeDay = string.Empty;
            this.Highest = 0;
            this.Lowest = 0;
        }

        public ContractMonth(string month, string lastTradeDay, int highest, int lowest)
        {
            this.LastTradeDay = lastTradeDay;
            this.Highest = highest;
            this.Lowest = lowest;
            this.Month = month;
        }
    }

    public class StockCodePrice
    {
        public string Code { get; set; }
        public int Price { get; set; }
        public StockPriceStatus Status { get; set; }

        public StockCodePrice()
        {
            this.Code = string.Empty;
            this.Price = 0;
            this.Status = 0;
        }

        public StockCodePrice(string code, int price)
        {
            this.Code = code;
            this.Price = price;
        }
        public StockCodePrice(string code, string price)
        {
            this.Code = code;
            this.Price = Convert.ToInt32(price);
        }
    }

    public enum StockPriceStatus : int
    {
        InMonthRange = 0,
        OutMonthRange = 1,
        New = 2
    }

    public enum OptionType
    {
        StockOption,
        USDOpiton,
        IndexOption
    }

    public enum OptionPageNumber
    {
        A = 10, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z
    }
}



