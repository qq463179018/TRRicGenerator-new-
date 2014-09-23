using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;
using HtmlAgilityPack;
using Ric.Util;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace Ric.Tasks.Thailand
{
    [ConfigStoredInDB]
    class TFEXNdaOptionCreationConfig
    {
        [StoreInDB]
        [Category("WorkFolder")]
        [DisplayName("Folder path")]
        [Description("Generated Files Folder Path,like: G:\\China")]
        public string FolderPath { get; set; }

        [Category("DateTime")]
        [DisplayName("Source date")]
        [Description("download the source from website by date time")]
        public string DateOfSource { get; set; }

        public TFEXNdaOptionCreationConfig()
        {
            DateOfSource = DateTime.Now.ToString("yyyy-MM-dd");
        }
    }

    class TFEXNdaOptionCreation : GeneratorBase
    {
        private TFEXNdaOptionCreationConfig configObj = null;
        private string sourceUrl = @"http://www.tfex.co.th/tfex/downloadSeriesProfile.html;jsessionid=C45FA8D23303CD199D84DD094C425602?locale=en_US";
        private List<string> title = null;
        List<TFEXBulkFileTemplate> listTFEX = new List<TFEXBulkFileTemplate>();

        protected override void Initialize()
        {
            configObj = Config as TFEXNdaOptionCreationConfig;
            title = new List<string>() {
                "RIC",
                "TAG",
                "TYPE",
                "CATEGORY",
                "ASSET COMMON NAME",
                "ASSET SHORT NAME",
                "CALL PUT OPTION",
                "DERIVATIVES QUOTE UNDERLYING ASSET",
                "EXPIRY DATE",
                "DERIVATIVES LAST TRADING DAY",
                "RETIRE DATE",
                "DERIVATIVES LOT SIZE",
                "STRIKE PRICE",
                "DERIVATIVES LOT UNIT",
                "EXCHANGE",
                "CURRENCY",
                "DERIVATIVES METHOD OF DELIVERY",
                "OPTIONS EXERCISE STYLE",
                "TICKER SYMBOL",
                "DERIVATIVES TICK VALUE",
                "RCS ASSET CLASS",
                "DERIVATIVES TRADING STYLE",
                "DERIVATIVES CONTRACT TYPE",
                "DERIVATIVES PERIODICITY",
                "DERIVATIVES SERIES DESCRIPTION",
                "RADING SYMBOL",
                "DERIVATIVES FIRSTRADING DAY",
                "OPTION STUB"
            };
        }

        protected override void Start()
        {
            LogMessage("start to download file from source site...");
            List<string> downloadFilePath = GetDownloadFilePath();
            LogMessage(string.Format("download file count:{0}", downloadFilePath == null ? "0" : downloadFilePath.Count.ToString()));

            if (downloadFilePath == null || downloadFilePath.Count == 0)
                return;

            LogMessage("start to extract data from download file...");
            ExtractDataFromFiles(downloadFilePath);
            LogMessage(string.Format("extract colunm count:{0}", listTFEX == null ? "0" : listTFEX.Count.ToString()));

            if (listTFEX == null || listTFEX.Count == 0)
                return;

            LogMessage("start to generate bulk file...");
            GenerateOutPutFile(listTFEX);

        }

        private void GenerateOutPutFile(List<TFEXBulkFileTemplate> listTFEX)
        {
            List<List<string>> listListOutput = new List<List<string>>();
            string path = Path.Combine(configObj.FolderPath, string.Format("TFEX_QA_Add_{0}.csv", configObj.DateOfSource));

            try
            {
                listListOutput.Add(title);
                FormateTemplate(listTFEX, listListOutput);
                XlsOrCsvUtil.GenerateStringCsv(path, listListOutput);
                AddResult("TFEX_yyy-MM-dd.csv", path, "Bulk file");
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void FormateTemplate(List<TFEXBulkFileTemplate> listTFEX, List<List<string>> listList)
        {
            foreach (var item in listTFEX)
            {
                List<string> line = new List<string>();

                line.Add(item.Ric);
                line.Add(item.Tag);
                line.Add(item.Type);
                line.Add(item.Category);
                line.Add(item.AssetCommonName);
                line.Add(item.AssetShortName);
                line.Add(item.CallPutOption);
                line.Add(item.DerivativesQuoteUnderlyingAsset);
                line.Add(item.ExpiryDate);
                line.Add(item.DerivativesLastTradingDay);
                line.Add(item.RetireDate);
                line.Add(item.DerivativesLotSize);
                line.Add(item.StrikePrice);
                line.Add(item.DerivativesLotUnit);
                line.Add(item.Exchange);
                line.Add(item.Currency);
                line.Add(item.DerivativesMethodOfDelivery);
                line.Add(item.OptionsExerciseStyle);
                line.Add(item.TickerSymbol);
                line.Add(item.DerivativesTickValue);
                line.Add(item.RcsAssetClass);
                line.Add(item.DerivativesTradingStyle);
                line.Add(item.DerivativesContractType);
                line.Add(item.DerivativesPeriodicity);
                line.Add(item.DerivativesSeriesDescription);
                line.Add(item.TradingSymbol);
                line.Add(item.DerivativesFirstTradingDay);
                line.Add(item.OptionStub);

                listList.Add(line);
            }
        }

        private void ExtractDataFromFiles(List<string> downloadFilePath)
        {
            try
            {
                foreach (var item in downloadFilePath)
                {
                    ExtractDataFromFile(item);
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
        }

        private void ExtractDataFromFile(string item)
        {
            string line = null;
            int count = 0;
            StreamReader sr = new StreamReader(item, Encoding.Default);
            while ((line = sr.ReadLine()) != null)
            {
                count++;
                if (count <= 2)
                    continue;

                if (!line.Contains(','))
                    continue;

                string[] lineColumn = line.Split(',');
                if (lineColumn == null || lineColumn.Length < 18)
                    continue;

                if (!lineColumn[3].Contains("OPT"))
                    continue;

                TFEXBulkFileTemplate tfex = new TFEXBulkFileTemplate(new DEXSRPTemplate(lineColumn));
                listTFEX.Add(tfex);
            }
        }

        private List<string> GetDownloadFilePath()
        {
            List<string> result = new List<string>();
            HtmlDocument htc = null;
            string baseUrl = @"http://www.tfex.co.th";
            string downloadFolder = Path.Combine(configObj.FolderPath, "Download");

            try
            {
                RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
                {
                    htc = WebClientUtil.GetHtmlDocument(sourceUrl, 3000);
                });

                if (htc == null)
                    throw new Exception(string.Format("open website: {0} error.", sourceUrl));

                HtmlNodeCollection trs = htc.DocumentNode.SelectNodes(".//table")[7].SelectNodes(".//tr");
                List<string> urlFoundAll = GetUrlFoundAll(trs);

                if (!Directory.Exists(downloadFolder))
                    Directory.CreateDirectory(downloadFolder);

                string tdNewSeries = GetUrlFoundOne(urlFoundAll);
                if ((tdNewSeries + "").Trim().Length == 0)
                    return null;

                string dexsrp = Path.Combine(downloadFolder, string.Format("dexsrp{0}.txt", DateTimeConvert(configObj.DateOfSource, "yyyy-MM-dd", "yyyyMMdd")));

                if (File.Exists(dexsrp))
                    File.Delete(dexsrp);

                RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
                {
                    WebClientUtil.DownloadFile(string.Format("{0}{1}", baseUrl, tdNewSeries), 30000, dexsrp);
                });

                if (File.Exists(dexsrp))
                    result.Add(dexsrp);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return result;
        }

        private string GetUrlFoundOne(List<string> urlFoundAll)
        {
            if (urlFoundAll == null || urlFoundAll.Count == 0)
                return string.Empty;

            foreach (var item in urlFoundAll)
            {
                if (IsSameWithDate(item))
                    return item;
            }

            return string.Empty;
        }

        private List<string> GetUrlFoundAll(HtmlNodeCollection trs)
        {
            List<string> result = new List<string>();

            if (trs == null || trs.Count <= 1)
                return null;

            try
            {
                string url = string.Empty;
                for (int i = 1; i < trs.Count; i++)
                {
                    try
                    {
                        url = trs[i].SelectNodes(".//td")[2].SelectSingleNode(".//a").Attributes["href"].Value.Trim();
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(string.Format("get download source url error.trs.count:{0}.msg:{1}", trs.Count.ToString(), ex.Message), Logger.LogType.Warning);
                        continue;
                    }

                    if ((url + "").Trim().Length == 0)
                        continue;

                    result.Add(url);
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

            return result;
        }

        private bool IsSameWithDate(string str)
        {
            bool result = false;
            string sourceDate = string.Empty;

            try
            {
                sourceDate = DateTimeConvert(configObj.DateOfSource, "yyyy-MM-dd", "yyyyMMdd");

                if (str.Contains(string.Format("dexsrp{0}.txt", sourceDate)))
                    result = true;

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return false;
            }
        }

        private string DateTimeConvert(string dateTimeStr, string formatAfter, string formatBefor)
        {
            string result = string.Empty;
            DateTime dt;

            if (DateTime.TryParseExact(configObj.DateOfSource, formatAfter, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
                result = dt.ToString(formatBefor);

            if (string.IsNullOrEmpty(result))
                throw new Exception(string.Format("convert configObj.DateTime: {0} error.", configObj.DateOfSource));

            return result;
        }
    }

    class DEXSRPTemplate
    {
        public DEXSRPTemplate(string[] lineColumn)
        {
            MarketID = lineColumn[0];
            ListID = lineColumn[1];
            SegmentID = lineColumn[2];
            InstrumentID = lineColumn[3];
            OrderBookID = lineColumn[4];
            Event = lineColumn[5];
            PreviousTIShortName = lineColumn[6];
            NewTIShortName = lineColumn[7];
            FirstTradeDate = lineColumn[8];
            LastTradeDate = lineColumn[9];
            ExpiredDate = lineColumn[10];
            ReferencePrice = lineColumn[11];
            PriceQuotationFactor = lineColumn[12];
            ContractSize = lineColumn[13];
            StrikePrice = lineColumn[14];
            OptionsType = lineColumn[15];
            OptionsStyle = lineColumn[16];
            PhysicalDelivery = lineColumn[17];
        }

        public string MarketID { get; set; }                //MarketID
        public string ListID { get; set; }                  //ListID	
        public string SegmentID { get; set; }               //SegmentID	
        public string InstrumentID { get; set; }            //InstrumentID	
        public string OrderBookID { get; set; }             //OrderBookID	
        public string Event { get; set; }                   //Event	
        public string PreviousTIShortName { get; set; }     //PreviousTIShortName	
        public string NewTIShortName { get; set; }          //NewTIShortName	
        public string FirstTradeDate { get; set; }          //FirstTradeDate	
        public string LastTradeDate { get; set; }           //LastTradeDate	
        public string ExpiredDate { get; set; }             //Expired Date	
        public string ReferencePrice { get; set; }          //ReferencePrice	
        public string PriceQuotationFactor { get; set; }    //PriceQuotationFactor	
        public string ContractSize { get; set; }            //ContractSize	
        public string StrikePrice { get; set; }             //StrikePrice	
        public string OptionsType { get; set; }             //OptionsType	
        public string OptionsStyle { get; set; }            //OptionsStyle	
        public string PhysicalDelivery { get; set; }        //PhysicalDelivery
    }

    class TFEXBulkFileTemplate
    {
        public string Ric { get; set; }
        public string Tag { get; set; }
        public string Type { get; set; }
        public string Category { get; set; }
        public string AssetCommonName { get; set; }
        public string AssetShortName { get; set; }
        public string CallPutOption { get; set; }
        public string DerivativesQuoteUnderlyingAsset { get; set; }
        public string ExpiryDate { get; set; }
        public string DerivativesLastTradingDay { get; set; }
        public string RetireDate { get; set; }
        public string DerivativesLotSize { get; set; }
        public string StrikePrice { get; set; }
        public string DerivativesLotUnit { get; set; }
        public string Exchange { get; set; }
        public string Currency { get; set; }
        public string DerivativesMethodOfDelivery { get; set; }
        public string OptionsExerciseStyle { get; set; }
        public string TickerSymbol { get; set; }
        public string DerivativesTickValue { get; set; }
        public string RcsAssetClass { get; set; }
        public string DerivativesTradingStyle { get; set; }
        public string DerivativesContractType { get; set; }
        public string DerivativesPeriodicity { get; set; }
        public string DerivativesSeriesDescription { get; set; }
        public string TradingSymbol { get; set; }
        public string DerivativesFirstTradingDay { get; set; }
        public string OptionStub { get; set; }

        public string callOrPut { get; set; }
        public string expireYearYY { get; set; }
        public string expireMonthMMM { get; set; }

        private Dictionary<string, string> dicPut = new Dictionary<string, string>();
        private Dictionary<string, string> dicCall = new Dictionary<string, string>();

        public TFEXBulkFileTemplate(DEXSRPTemplate dex)
        {
            callOrPut = GetCallOrPut(dex.NewTIShortName);
            AddDic();
            Tag = "47512";
            Type = "DERIVATIVE";
            Category = "EIO";
            CallPutOption = SetCallPutOption(callOrPut);
            DerivativesQuoteUnderlyingAsset = ".SET50";
            ExpiryDate = SetExpiryDate(dex.ExpiredDate);
            DerivativesLastTradingDay = DateTimeConvert(dex.LastTradeDate.Replace("/", "-"), "yyyy-MM-dd", "dd-MMM-yyyy");
            RetireDate = SetRetireDate(DerivativesLastTradingDay);
            DerivativesLotSize = dex.ContractSize;
            StrikePrice = dex.StrikePrice.Replace(".000000", "");
            DerivativesLotUnit = "INDEX";
            Exchange = "TFX";
            Currency = "THB";
            DerivativesMethodOfDelivery = "CASH";
            OptionsExerciseStyle = "E";
            TickerSymbol = "S50";
            DerivativesTickValue = "20";
            RcsAssetClass = "OPT";
            DerivativesTradingStyle = "E";
            DerivativesContractType = "7";
            DerivativesPeriodicity = "Q";
            DerivativesSeriesDescription = "Thailand Futures Exchange SET 50 Index Option";
            TradingSymbol = dex.NewTIShortName;
            DerivativesFirstTradingDay = DateTimeConvert(dex.FirstTradeDate.Replace("/", "-"), "yyyy-MM-dd", "dd-MMM-yyyy");
            OptionStub = "S50.FX";

            Ric = SetRic(dex);
            AssetCommonName = string.Format("S50 {0}{1} {2}{3}", expireMonthMMM, expireYearYY, StrikePrice, callOrPut).ToUpper();
            AssetShortName = AssetCommonName;
        }

        private string SetRic(DEXSRPTemplate dex)
        {
            string result = string.Empty;
            string monthCode = string.Empty;
            string yearCode = string.Empty;
            if (callOrPut.Equals("C"))
            {
                if (dicCall.ContainsKey(expireMonthMMM))
                    monthCode = dicCall[expireMonthMMM];
            }
            else if (callOrPut.Equals("P"))
            {
                if (dicPut.ContainsKey(expireMonthMMM))
                    monthCode = dicPut[expireMonthMMM];
            }

            yearCode = expireYearYY;
            result = string.Format("S50{0}{1}{2}.FX", StrikePrice, monthCode, yearCode);
            return result;
        }

        private void AddDic()
        {
            dicCall.Add("Jan", "A");
            dicCall.Add("Feb", "B");
            dicCall.Add("Mar", "C");
            dicCall.Add("Apr", "D");
            dicCall.Add("May", "E");
            dicCall.Add("Jun", "F");
            dicCall.Add("Jul", "G");
            dicCall.Add("Aug", "H");
            dicCall.Add("Sep", "I");
            dicCall.Add("Oct", "J");
            dicCall.Add("Nov", "K");
            dicCall.Add("Dec", "L");

            dicPut.Add("Jan", "M");
            dicPut.Add("Feb", "N");
            dicPut.Add("Mar", "O");
            dicPut.Add("Apr", "P");
            dicPut.Add("May", "Q");
            dicPut.Add("Jun", "R");
            dicPut.Add("Jul", "S");
            dicPut.Add("Aug", "T");
            dicPut.Add("Sep", "U");
            dicPut.Add("Oct", "V");
            dicPut.Add("Nov", "W");
            dicPut.Add("Dec", "X");
        }

        private string SetRetireDate(string p)
        {
            string result = string.Empty;
            DateTime dt;
            if (DateTime.TryParseExact(p.Replace("/", "-"), "dd-MMM-yyyy", new CultureInfo("en-US"), DateTimeStyles.None, out dt))
                result = dt.AddDays(+4).ToString("dd-MMM-yyyy");

            return result;
        }

        private string DateTimeConvert(string dateTimeStr, string formatAfter, string formatBefor)
        {
            string result = string.Empty;
            DateTime dt;

            if (DateTime.TryParseExact(dateTimeStr, formatAfter, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
                result = dt.ToString(formatBefor);

            return result;
        }

        private string SetExpiryDate(string p)
        {
            string result = string.Empty;
            DateTime dt;
            if (DateTime.TryParseExact(p.Replace("/", "-"), "yyyy-MM-dd", new CultureInfo("en-US"), DateTimeStyles.None, out dt))
            {
                result = dt.ToString("dd-MMM-yyyy");
                expireYearYY = result.Substring(result.Length - 1, 1);
                expireMonthMMM = result.Substring(3, 3);
            }
            return result;
        }

        private string SetCallPutOption(string callOrPut)
        {
            if (string.IsNullOrEmpty(callOrPut))
                return string.Empty;

            if (callOrPut.Equals("C"))
                return "CALL";
            else if (callOrPut.Equals("P"))
                return "PUT";
            else
                return string.Empty;
        }

        private string GetCallOrPut(string p)
        {
            Match ma = (new Regex(@"\S{6}(?<Value>(P|C){1})")).Match(p);
            if (!ma.Success)
                return string.Empty;

            return ma.Groups["Value"].Value;
        }
    }
}
