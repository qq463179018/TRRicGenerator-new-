using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Ric.Util;
using Ric.Core;

namespace Ric.Tasks.HongKong
{

    public class StockOption
    {
        public string ClassName { get; set; }
        public string ExpiryMonth { get; set; }
        public string ExpiryYear { get; set; }
        public string Price { get; set; }


        public List<string> CallRicList = new List<string>();
        public List<string> PutRicList = new List<string>();

        public void GetRicList(List<UnderlyingCodeConversion> underlyCodeMap, List<MonthConversion> contractCodeMap)
        {
            Price = Price.Replace("/", "");
            if (Price.StartsWith("$"))
            {
                Price = Price.Remove(0, 1);
            }
            string[] priceArr = Price.Split('$');
            foreach (string price in priceArr)
            {
                string Ric = string.Empty;
                string callRic = string.Empty;
                string putRic = string.Empty;

                Ric += getUnderlyingCode(underlyCodeMap);
                Ric += getPriceCode(price);
                //Version Code : 0
                Ric += "0";

                string[] callPutCode = getContractMonthCode(contractCodeMap);
                string yearCode = getYearCode();

                callRic = Ric + callPutCode[0];
                putRic = Ric + callPutCode[1];
                callRic += yearCode;
                callRic += ".HK";

                putRic += yearCode;
                putRic += ".HK";
                CallRicList.Add(callRic);
                PutRicList.Add(putRic);
            }
        }

        private string getUnderlyingCode(List<UnderlyingCodeConversion> underlyCodeMap)
        {
            string underlyingCode = string.Empty;
            foreach (UnderlyingCodeConversion underlyingConversion in underlyCodeMap)
            {
                if (ClassName == underlyingConversion.ClassName)
                {
                    underlyingCode = underlyingConversion.UnderlyCode;
                    break;
                }
            }
            return underlyingCode;
        }

        private string[] getContractMonthCode(List<MonthConversion> contractCodeMap)
        {
            string[] callPutCode = { "", "" };
            foreach (MonthConversion contractConversion in contractCodeMap)
            {
                if (contractConversion.Month == ExpiryMonth)
                {
                    callPutCode[0] = contractConversion.CallContractCode;
                    callPutCode[1] = contractConversion.PutContactCode;
                    break;
                }
            }
            return callPutCode;
        }

        private string getPriceCode(string price)
        {
            price = price.Replace("$", "");
            string priceCode = string.Empty;
            string[] arr = price.Split('.');
            price = price.Replace(".", "");
            if (arr.Length == 1)
            {
                priceCode = price + "00";
            }
            else if (arr.Length == 2)
            {
                if (arr[1].Length == 0)
                {
                    priceCode = price + "00";
                }
                else if (arr[1].Length == 1)
                {
                    priceCode = price + "0";
                }
                else
                {
                    priceCode = price;
                }
            }
            else
            {
                throw (new Exception());
            }
            return priceCode;
        }
        private string getYearCode()
        {
            return ExpiryYear[ExpiryYear.Length - 1].ToString();
        }
    }

    public class UnderlyingCodeConversion
    {
        public string ClassName { get; set; }
        public string UnderlyCode { get; set; }
    }
    public class ClassUnderlyCodeMap
    {
        public List<UnderlyingCodeConversion> UnderlyCodeMap { get; set; }
    }


    public class MonthConversion
    {
        public string Month { get; set; }
        public string CallContractCode { get; set; }
        public string PutContactCode { get; set; }

    }
    public class MonthContractCodeMap
    {
        public List<MonthConversion> ContractCodeMap { get; set; }
    }

    public class OptionRicGeneratorConfig
    {
        public string RIC_GENERATE_FILE_DIR { get; set; }
        public string LOG_FILE_PATH { get; set; }
        public string MAIN_PAGE_URI { get; set; }
    }

    public class StockOptionRicGenerator : GeneratorBase
    {
        private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_OptionRicGenerator.config";
        private static readonly string ClASS_UNDERLYCODE_MAP_PATH = ".\\Config\\HK\\HK_ClassUnderlyCodeMap.xml";
        private static readonly string MONTH_CONTRACTCODE_MAP_PATH = ".\\Config\\HK\\HK_MonthContractCodeMap.xml";
        //private static readonly string MAIN_PAGE_URI = "http://www.hkex.com.hk/eng/market/sec_tradinfo/tradnews/prvtrad_day/news.htm";//"http://www.hkex.com.hk/eng/market/sec_tradinfo/tradnews/today/news.htm";
        private static readonly string NEWS_TITLE_PREFIX = "OPTN NEWS- NEW OPTION SERIES";
        private static OptionRicGeneratorConfig configObj = null;
        private static ClassUnderlyCodeMap underlyingCodeMap = null;
        private static MonthContractCodeMap monthCodeMap = null;
        //private static Logger logger = null;

        //private static readonly string NEWS_TITLE_PREFIX = "Short Sell Turnover (Main Board) up to day close today".Trim().ToLower();

        protected override void Start()
        {
            StartOptionRicGeneratorJob();
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(OptionRicGeneratorConfig)) as OptionRicGeneratorConfig;
            underlyingCodeMap = ConfigUtil.ReadConfig(ClASS_UNDERLYCODE_MAP_PATH, typeof(ClassUnderlyCodeMap)) as ClassUnderlyCodeMap;
            monthCodeMap = ConfigUtil.ReadConfig(MONTH_CONTRACTCODE_MAP_PATH, typeof(MonthContractCodeMap)) as MonthContractCodeMap;
            //logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.New);
        }

        public void StartOptionRicGeneratorJob()
        {

            List<StockOption> stockOptionList = new List<StockOption>();
            stockOptionList = GetStockOptionList();
            using (ExcelApp app = new ExcelApp(false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, Path.Combine(configObj.RIC_GENERATE_FILE_DIR, NewFileName()));
                var worksheet = workbook.Worksheets[1] as Worksheet;
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    foreach (StockOption stockOption in stockOptionList)
                    {
                        for (int i = 0; i < stockOption.PutRicList.Count; i++)
                        {
                            writer.WriteLine(stockOption.CallRicList[i]);
                            writer.WriteLine("D" + stockOption.CallRicList[i]);
                            writer.WriteLine("/" + stockOption.CallRicList[i]);
                            writer.PlaceNext(writer.Row + 1, 1);
                            writer.WriteLine(stockOption.PutRicList[i]);
                            writer.WriteLine("D" + stockOption.PutRicList[i]);
                            writer.WriteLine("/" + stockOption.PutRicList[i]);
                            writer.PlaceNext(writer.Row + 1, 1);
                        }
                    }
                }
                workbook.Save();
            }
        }

        public List<StockOption> GetStockOptionList()
        {
            List<StockOption> stockOptionList = new List<StockOption>();
            List<string> linkUrlList = GetUrlLinksFromMainPage();
            foreach (string uri in linkUrlList)
            {
                //string uri = "http://www.hkex.com.hk/eng/market/sec_tradinfo/tradnews/today/enew.htm";
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc = WebClientUtil.GetHtmlDocument(uri, 270000);
                string stockOptionInfo = htmlDoc.DocumentNode.SelectSingleNode("//tbody/tr[@valign='top']/td[@valign='top']/pre").InnerText;
                using (StringReader sr = new StringReader(stockOptionInfo))
                {
                    string line1;
                    while (((line1 = sr.ReadLine()) != null))
                    {
                        //Parse the stock info 
                        Regex stockInfoRegex = new Regex("(?s)(?<className>\\w{3})\\s{2,}(?<month>\\w{3})\\s{1,}(?<year>\\d{2})\\s{2,}(?<price>.*\\d)");
                        Match stockInfoMatch = stockInfoRegex.Match(line1);
                        if (stockInfoMatch.Success)
                        {
                            StockOption stockInfo = new StockOption();
                            stockInfo.ClassName = stockInfoMatch.Groups["className"].Value;
                            stockInfo.ExpiryMonth = stockInfoMatch.Groups["month"].Value;
                            stockInfo.ExpiryYear = stockInfoMatch.Groups["year"].Value;
                            stockInfo.Price = stockInfoMatch.Groups["price"].Value;
                            if (line1.EndsWith("/"))
                            {
                                stockInfo.Price += sr.ReadLine();
                            }
                            stockInfo.GetRicList(underlyingCodeMap.UnderlyCodeMap, monthCodeMap.ContractCodeMap);

                            stockOptionList.Add(stockInfo);
                            continue;
                        }
                    }
                }
            }
            return stockOptionList;
        }

        public List<string> GetUrlLinksFromMainPage()
        {
            List<string> linkUrlList = new List<string>();
            var htmlDoc = WebClientUtil.GetHtmlDocument(configObj.MAIN_PAGE_URI, 270000);
            var linkNodeList = htmlDoc.DocumentNode.SelectNodes("//span[@id='Content']/table/tbody/tr//td//a");
            foreach (var linkNode in linkNodeList)
            {
                if (linkNode.Attributes["href"] != null)
                {
                    string linkText = linkNode.InnerText;
                    string linkUrl = linkNode.Attributes["href"].Value;
                    if (!MiscUtil.IsAbsUrl(linkUrl))
                    {
                        linkUrl = MiscUtil.UrlCombine(configObj.MAIN_PAGE_URI, linkUrl);
                    }
                    if (linkText.Trim().StartsWith(NEWS_TITLE_PREFIX) && !(linkUrlList.Contains(linkUrl)))
                    {
                        linkUrlList.Add(linkUrl);
                    }
                }
            }
            return linkUrlList;
        }

        public string NewFileName()
        {
            string[] month = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            string fileName = "StockOptionRic";
            fileName += "_";
            string currentDay = DateTime.Now.ToString("dd_MM_yyyy");
            string[] dateTime = currentDay.Split('_');
            fileName += dateTime[0];
            fileName += month[int.Parse(dateTime[1])];
            fileName += dateTime[2];
            fileName += "_";
            fileName += Guid.NewGuid().ToString();
            fileName += ".xls";
            return fileName;
        }
    }
}
