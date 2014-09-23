using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.DataStream
{
    public class DataStreamRicCreationWithDseHongKong : GeneratorBase
    {
        #region Fields

        private DataStreamRicCreationWithDseConfig configObj;
        private Dictionary<string, string> namesAbbs = new Dictionary<string, string>();
        private CookieContainer cookies = new CookieContainer();
        private bool ifDownNameRule;
        private bool loginSuccess;
        private bool noshSessioned;

        private string normalOutputLine = string.Empty;
        private string hShareOutputLine = string.Empty;

        private string mFileFolder = string.Empty;

        #endregion

        #region Initialize and Start

        protected override void Initialize()
        {
            configObj = Config as DataStreamRicCreationWithDseConfig;
            AddResult("LOG File", Logger.FilePath, "LOG File");

            InitializeFileDirectory();

            string msg = "Initialize...OK!";
            Logger.Log(msg);
        }

        private void InitializeFileDirectory()
        {
            string outputFolder = Path.Combine(configObj.OutputPath, DateTime.Today.ToString("yyyy-MM-dd"));

            mFileFolder = Path.Combine(outputFolder, "DSE_FILES");

            if (!Directory.Exists(mFileFolder))
            {
                Directory.CreateDirectory(mFileFolder);
            }

            AddResult("DSE_FILES", mFileFolder, "DSE_FILES");
        }

        protected override void Start()
        {
            StartJob();
        }

        private void StartJob()
        {
            DownloadFtpFiles();
            List<string> records = GetRecords();
            if (records == null || records.Count == 0)
            {
                Logger.Log("No RIC creation today.");
                return;
            }
            List<DataStreamRicCreationInfo> ricCreations = FormatRecords(records);

            if (ricCreations == null || ricCreations.Count == 0)
            {
                Logger.Log("No RIC creation today.");
                return;
            }

            if (ifDownNameRule)
            {
                DownloadNameRules();
            }

            FormatOutputLine(ricCreations);
            GenerateFile();

        }

        #endregion

        private string GetCurrrentFileDate(string fileStart)
        {
            string inputDate = configObj.Date;
            DateTime dateToUse = DateTime.Now;
            if (fileStart == "EM01")
            {
                int daysToAdd = -1;
                dateToUse = DateTime.ParseExact(inputDate, "MMdd", CultureInfo.InvariantCulture);
                if (dateToUse.DayOfWeek == DayOfWeek.Monday)
                {
                    daysToAdd = -1;
                }
                return dateToUse.AddDays(daysToAdd).ToString("MMdd");
            }
            return inputDate;
        }

        #region Get FTP records

        private void DownloadFtpFiles()
        {
            string currentDate = String.Empty;
            string[] fileStartArr = { "0001", "EM01" };

            foreach (string fileStart in fileStartArr)
            {

                currentDate = GetCurrrentFileDate(fileStart);

                string fileName = string.Format("{0}{1}.M", fileStart, currentDate);

                string ftpfullpath = @"ftp://ASIA2:ASIA2@ds1.rds.reuters.com/" + fileName;

                try
                {
                    WebClient request = new WebClient();

                    if (!string.IsNullOrEmpty(configObj.IP) && !string.IsNullOrEmpty(configObj.Port))
                    {
                        WebProxy proxy = new WebProxy(configObj.IP, Convert.ToInt32(configObj.Port));
                        request.Proxy = proxy;
                    }

                    request.Credentials = new NetworkCredential("ASIA2", "ASIA2");

                    string mfilePath = Path.Combine(mFileFolder, fileName);

                    request.DownloadFile(ftpfullpath, mfilePath);

                    Logger.Log(string.Format("Download FTP File {0}... OK!", fileName));
                }
                catch (Exception ex)
                {
                    string msg = string.Format("Can not download file: {0} from FTP. Response:{1}", fileName, ex.Message);
                    Logger.Log(msg, Logger.LogType.Error);
                    LogMessage(string.Format("Can not download file: {0}", fileName));
                }
            }

        }

        private List<string> GetRecords()
        {
            List<string> xeRecord = new List<string>();
            string currentDate = String.Empty;

            string[] fileStartArr = { "0001", "EM01" };

            foreach (string fileStart in fileStartArr)
            {
                currentDate = GetCurrrentFileDate(fileStart);
                string fileName = string.Format("{0}{1}.M", fileStart, currentDate);

                string mfilePath = Path.Combine(mFileFolder, fileName);

                if (!File.Exists(mfilePath))
                {
                    continue;
                }


                using (StreamReader sr = new StreamReader(mfilePath))
                {
                    string tmp;
                    while ((tmp = sr.ReadLine()) != null)
                    {
                        if (!tmp.StartsWith("XE"))
                        {
                            continue;
                        }
                        if (tmp.StartsWith("XEIPO"))
                        {
                            continue;
                        }
                        string ric = tmp.Substring(0, tmp.IndexOf(' ')).Replace("XE", "");
                        if (!ric.EndsWith("HK"))
                        {
                            continue;
                        }

                        string ticker = ric.Replace(".HK", "");

                        int tickerNum;

                        if (!int.TryParse(ticker, out tickerNum))
                        {
                            continue;
                        }

                        xeRecord.Add(tmp);

                        string msg = string.Format("Get 1 record from file: {0}. RIC:{1}", fileName, ric);
                        Logger.Log(msg);

                    }
                }

            }

            return xeRecord;
        }

        #endregion

        #region Format DSE records

        private List<DataStreamRicCreationInfo> FormatRecords(IEnumerable<string> records)
        {
            Dictionary<string, DseFieldPosition> dseFields = InitializeDseFields();
            List<DataStreamRicCreationInfo> ricCreations = new List<DataStreamRicCreationInfo>();
            foreach (string record in records)
            {
                DataStreamRicCreationInfo item = new DataStreamRicCreationInfo
                {
                    Ticker = FormatField(record, "Ticker", dseFields),
                    Sedol = FormatField(record, "Sedol", dseFields),
                    CompanyName = FormatField(record, "CompanyName", dseFields)
                };
                if (!string.IsNullOrEmpty(item.CompanyName))
                {
                    ifDownNameRule = true;
                }
                item.FirstTradingDate = FormatField(record, "FirstTradingDate", dseFields);
                item.ExchangeCode = FormatField(record, "ExchangeCode", dseFields);
                item.Isin = FormatField(record, "Isin", dseFields);
                item.SecurityDescription = FormatField(record, "Security Description", dseFields);

                if (item.SecurityDescription.EndsWith("RTS"))
                {
                    string ric = FormatField(record, "RIC", dseFields);
                    string msg = string.Format("RIC:{0} is a RTS. Ignore it.", ric);
                    Logger.Log(msg);
                    continue;
                }
                if (item.SecurityDescription.EndsWith(" (TEMP)"))
                {
                    string ric = FormatField(record, "RIC", dseFields);
                    string msg = string.Format("RIC:{0} is parallel trading temp. Ignore it.", ric);
                    Logger.Log(msg);
                    continue;
                }

                string assetCategory = FormatField(record, "AssetCategory", dseFields);
                if (!assetCategory.Equals("ORD"))
                {
                    string ric = FormatField(record, "RIC", dseFields);
                    string msg = string.Format("RIC:{0} is a {1}. Ignore it.", ric, assetCategory);
                    Logger.Log(msg);
                    continue;
                }

                ricCreations.Add(item);
            }
            return ricCreations;
        }

        private Dictionary<string, DseFieldPosition> InitializeDseFields()
        {
            Dictionary<string, DseFieldPosition> dseFields = new Dictionary<string, DseFieldPosition>
            {
                {"Security Description", new DseFieldPosition(23, 58)},
                {"Ticker", new DseFieldPosition(412, 436)},
                {"Sedol", new DseFieldPosition(68, 74)},
                {"CompanyName", new DseFieldPosition(105, 184)},
                {"FirstTradingDate", new DseFieldPosition(1022, 1029)},
                {"ExchangeCode", new DseFieldPosition(98, 100)},
                {"Isin", new DseFieldPosition(84, 95)},
                {"RIC", new DseFieldPosition(3, 22)},
                {"AssetCategory", new DseFieldPosition(407, 410)}
            };
            return dseFields;
        }

        private string FormatField(string record, string field, Dictionary<string, DseFieldPosition> dseFields)
        {
            if (!dseFields.ContainsKey(field))
            {
                return "";
            }
            DseFieldPosition dseField = dseFields[field];
            if (record.Length < dseField.EndPosition + 1)
            {
                string msg = string.Format("At FormatField(). Can not get field:{0}. Out of length.", field);
                Logger.Log(msg, Logger.LogType.Error);
                return "";
            }

            return record.Substring(dseField.StartPosition - 1, dseField.EndPosition - dseField.StartPosition + 1).Trim();
        }

        #endregion

        #region Format Output records

        private void FormatOutputLine(List<DataStreamRicCreationInfo> ricCreations)
        {
            List<string> normalLine = new List<string>();
            List<string> hShareLine = new List<string>();
            foreach (DataStreamRicCreationInfo item in ricCreations)
            {
                string temp = "";
                DataStreamRicCreationBulkTemplate bulkTemp = new DataStreamRicCreationBulkTemplate();
                bulkTemp.Seqn = temp.PadRight(7, ' ');
                string initials = GetInitialsFromName(item.CompanyName);
                bulkTemp.Mnem = "K:" + initials;
                if (string.IsNullOrEmpty(item.Sedol))
                {
                    string sedol = GetSedol(item.Isin);
                    if (string.IsNullOrEmpty(sedol))
                    {
                        sedol = temp.PadRight(9, ' ');
                    }
                    else
                    {
                        sedol = "UK" + sedol;
                    }
                    bulkTemp.Sedol = sedol;
                }
                else
                {
                    bulkTemp.Sedol = "UK" + item.Sedol;
                }

                bulkTemp.IsHShare = CheckIfHShare(item.SecurityDescription);

                FormatBulkNames(item.CompanyName, bulkTemp);

                FormatCName(bulkTemp);

                if (bulkTemp.IsHShare)
                {
                    FormatBulkNamesForHShare(bulkTemp);
                }

                if (item.FirstTradingDate.Equals("-9999999"))
                {
                    bulkTemp.Bdate = temp.PadRight(8, ' ');
                }
                else
                {
                    DateTime dt;
                    string bdate = string.Empty;
                    bdate = DateTime.TryParseExact(item.FirstTradingDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal, out dt) ? dt.ToString("dd/MM/yy", DateTimeFormatInfo.InvariantInfo) : temp.PadRight(8, ' ');
                    bulkTemp.Bdate = bdate;
                }


                bulkTemp.DefMki = "HK";

                bulkTemp.Nosh = "1";
                bulkTemp.Mnem2 = "HK" + item.Ticker.TrimStart('0');
                bulkTemp.Mnem2 = bulkTemp.Mnem2.PadRight(12, ' ');
                bulkTemp.Isin = string.IsNullOrEmpty(item.Isin) ? temp.PadRight(12, ' ') : item.Isin;

                if (bulkTemp.IsHShare)
                {
                    bulkTemp.Geog = "051";
                    bulkTemp.Dcur = "079";
                }
                else
                {
                    bulkTemp.Geog = temp.PadRight(3, ' ');
                    bulkTemp.Dcur = temp.PadRight(3, ' ');
                }

                string line = string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",\"{12}\"",
                                            bulkTemp.Seqn, bulkTemp.Mnem, bulkTemp.Sedol, bulkTemp.Name1, bulkTemp.Name2,
                                            bulkTemp.Bdate, bulkTemp.DefMki, bulkTemp.Nosh, bulkTemp.Mnem2, bulkTemp.Isin, bulkTemp.CName, bulkTemp.Geog, bulkTemp.Dcur);
                if (bulkTemp.IsHShare)
                {
                    hShareLine.Add(line);
                }
                else
                {
                    normalLine.Add(line);
                }

                if (ricCreations.Count > 1)
                {
                    System.Threading.Thread.Sleep(5000);
                }

            }
            normalOutputLine = string.Join("\r\n", normalLine.ToArray());
            hShareOutputLine = string.Join("\r\n", hShareLine.ToArray());
        }

        private string GetInitialsFromName(string companyName)
        {
            if (string.IsNullOrEmpty(companyName))
            {
                return "";
            }
            string[] names = companyName.Split(' ');
            string initials = string.Empty;

            if (names.Length < 4)
            {
                companyName = Regex.Replace(companyName, " +", "");
                return companyName.Substring(0, 4).ToUpper();
            }

            foreach (string name in names)
            {
                if (name.Trim() != "" && initials.Length < 4)
                {
                    initials += name.Substring(0, 1).ToUpper();
                }
            }
            return initials;
        }

        #region For Field NAME1,NAME2 and CNAME

        private void FormatCName(DataStreamRicCreationBulkTemplate bulkTemp)
        {
            string name1 = bulkTemp.Name1.Trim();
            string name2 = bulkTemp.Name2.Trim();

            string names = string.Empty;
            if (!name1.EndsWith("."))
            {
                names = name1 + " " + name2;
            }
            else
            {
                names = name1 + name2;
            }

            if (names.EndsWith(".") && (names.Length + bulkTemp.NameSuffix.Length) <= 24)
            {
                names += bulkTemp.NameSuffix;
            }
            else if ((!names.EndsWith(".")) && (names.Length + 1 + bulkTemp.NameSuffix.Length) <= 24)
            {
                names += " " + bulkTemp.NameSuffix;
            }

            else if (names.Length > 24)
            {
                names = names.Substring(0, 24);
            }

            bulkTemp.CName = names.PadRight(24, ' ');
        }

        private void FormatBulkNames(string companyName, DataStreamRicCreationBulkTemplate bulkTemp)
        {
            //Rules:
            //1. 提取第一个单词后trim
            //2. 去掉5个后缀 if 有
            //3. 去掉company后缀， if 有
            //4. 如果Joint。。。Company出现在中间， 用JSC.替换
            //5. 开始找缩写
            //6. 开始format两个name
            //7. if > 48, user input
            string temp = string.Empty;
            if (string.IsNullOrEmpty(companyName))
            {
                bulkTemp.Name1 = temp.PadRight(24, ' ');
                bulkTemp.Name2 = temp.PadRight(24, ' ');
                return;
            }

            string inputConpanyName = companyName;
            companyName = companyName.ToUpper();

            string[] nameWords = companyName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            //Company Name only contains one word.
            if (nameWords.Length == 1)
            {
                string abbr = companyName;
                if (namesAbbs.ContainsKey(companyName))
                {
                    abbr = namesAbbs[companyName];
                }
                bulkTemp.Name1 = abbr.PadRight(24, ' ');
                bulkTemp.Name2 = temp.PadRight(24, ' ');
                return;
            }

            companyName = FormatCompanyName(companyName);

            string nameSuffix = string.Empty;
            string nameLeft = TrimCompanyEndings(companyName, ref nameSuffix);
            bulkTemp.NameSuffix = nameSuffix;

            List<string> abbreviations = GetNameAbbreviations(nameLeft);

            string namePart1 = string.Empty;
            string namePart2 = string.Empty;

            bool reFormat = false;
            do
            {
                if (reFormat)
                {
                    if (abbreviations[abbreviations.Count - 1] == "CORP.")
                    {
                        bulkTemp.NameSuffix = "CORP.";
                    }

                    if (abbreviations.Contains("CORP."))
                    {
                        abbreviations.Remove("CORP.");
                    }
                    if (abbreviations.Contains("COMPANY"))
                    {
                        abbreviations.Remove("COMPANY");
                    }
                }
                string formatedName = string.Join("", abbreviations.ToArray()).Trim();
                if (formatedName.Length > 48)
                {
                    if (reFormat)
                    {
                        string[] namesInput = DataStreamRicCreationNamesInput.Prompt(inputConpanyName, formatedName);
                        if (namesInput != null)
                        {
                            namePart1 = namesInput[0];
                            namePart2 = namesInput[1];
                        }
                    }
                    else
                    {
                        reFormat = true;
                    }
                }
                else if (formatedName.Length <= 24)
                {
                    namePart1 = formatedName;
                }

                else
                {
                    int index = 24;
                    int index2 = 24;
                    char[] charArr = formatedName.ToCharArray();
                    if (charArr[index] == ' ')
                    {
                        index2 = index;
                    }
                    else
                    {
                        while (true)
                        {
                            char middleChar = charArr[index];
                            if (middleChar == ' ' || middleChar == '.')
                            {
                                break;
                            }
                            index--;
                        }
                        index2 = index + 1;
                    }
                    namePart1 = formatedName.Substring(0, index2);
                    namePart2 = formatedName.Substring(index2).Trim();
                    if (namePart2.Length > 24)
                    {
                        if (reFormat)
                        {
                            string[] namesInput = DataStreamRicCreationNamesInput.Prompt(inputConpanyName, formatedName);
                            if (namesInput != null)
                            {
                                namePart1 = namesInput[0];
                                namePart2 = namesInput[1];
                            }
                        }
                        else
                        {
                            reFormat = true;
                        }
                    }
                }
            }
            while (reFormat);

            bulkTemp.Name1 = namePart1.PadRight(24, ' ');
            bulkTemp.Name2 = namePart2.PadRight(24, ' ');
        }

        private List<string> GetNameAbbreviations(string nameLeft)
        {
            List<string> abbrevations = new List<string>();

            List<string> nameLeftArr = nameLeft.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            //Rule 14 in requirement v3.0. Abbreviate from right to left.            

            if (nameLeft.Length <= 24)
            {
                abbrevations.AddRange(nameLeftArr.Select(subName => subName + " "));
                return abbrevations;
            }

            string firstWord = nameLeftArr[0] + " ";
            nameLeftArr.RemoveAt(0);

            while (nameLeftArr.Count > 0)
            {
                string nameToFind = string.Join(" ", nameLeftArr.ToArray());
                for (int j = 0; j <= nameLeftArr.Count - 1; j++)
                {
                    if (!namesAbbs.ContainsKey(nameToFind) && j != nameLeftArr.Count - 1)
                    {
                        nameToFind = nameToFind.Replace(nameLeftArr[j], "").Trim();
                        continue;
                    }
                    string abbr = string.Empty;
                    if (j == nameLeftArr.Count - 1)
                    {
                        if (namesAbbs.ContainsKey(nameToFind))
                        {
                            abbr = namesAbbs[nameToFind] + ".";
                        }
                        else
                        {
                            abbr = nameLeftArr[j];
                            if (!abbr.Contains("."))
                            {
                                abbr = abbr + " ";
                            }
                        }
                        if (abbr.Equals("&."))
                        {
                            abbr = "&";
                        }

                    }
                    else
                    {
                        abbr = namesAbbs[nameToFind] + ".";
                    }

                    int k = nameLeftArr.Count - 1 - j;
                    while (k-- >= 0)
                    {
                        nameLeftArr.RemoveAt(nameLeftArr.Count - 1);
                    }

                    if (!string.IsNullOrEmpty(abbr))
                    {
                        abbrevations.Add(abbr);
                        string nameFormatTemp = firstWord + string.Join(" ", nameLeftArr.ToArray()) + " " + string.Join("", abbrevations.ToArray());
                        if (nameFormatTemp.Length <= 24)
                        {
                            for (int i = nameLeftArr.Count - 1; i >= 0; i--)
                            {
                                abbrevations.Add(nameLeftArr[i] + " ");
                            }

                            abbrevations.Add(firstWord);
                            abbrevations.Reverse();
                            return abbrevations;
                        }
                    }
                }
            }
            abbrevations.Add(firstWord);
            abbrevations.Reverse();
            return abbrevations;
        }

        private string FormatCompanyName(string companyName)
        {
            if (companyName.Contains("CO LTD"))
            {
                companyName = companyName.Replace("CO LTD", "COMPANY LIMITED");
            }
            return companyName;
        }


        private string TrimCompanyEndings(string nameLeft, ref string nameSuffix)
        {

            if (nameLeft.Contains("COMPANY LIMITED"))
            {
                return nameLeft;
            }

            string[] endings = { "JSC", "JOINT STOCK COMPANY", "JOINT STOCK CORPORATION", "JOINT-STOCK CORPORATION", "PUBLIC LIMITED COMPANY", "INCORPORATED", "COMPANY", "LIMITED", "LTD" };
            foreach (string ending in endings.Where(ending => (nameLeft.Length > ending.Length) && (nameLeft.Substring(nameLeft.Length - ending.Length - 1).Equals(" " + ending))))
            {
                nameLeft = nameLeft.Replace(ending, "");
                if (ending.Equals("JSC") || ending.Equals("JOINT STOCK COMPANY"))
                {
                    nameSuffix = "JSC.";
                }
                if (ending.Equals("LTD"))
                {
                    nameSuffix = "LTD";
                }
                if (ending.Equals("LIMITED"))
                {
                    nameSuffix = "LTD.";
                }
            }

            if (nameLeft.Contains(" JOINT STOCK COMPANY "))
            {
                nameLeft = nameLeft.Replace("JOINT STOCK COMPANY", "JSC.");
            }

            return nameLeft.Trim();
        }

        private void DownloadNameRules()
        {
            string url = @"http://dataops.datastream.com/cgi-bin/readfile.pl?filename=H:/Production/Loaders/Global/DataStream/Tools/Abbreviation/Mload/abbreviations.rep&warnold=1";
            HtmlDocument doc = null;
            int retry = 3;
            while (doc == null && retry-- > 0)
            {
                try
                {
                    string pageSource = WebClientUtil.GetPageSource(null, url, 180000, null, Encoding.GetEncoding("ISO-8859-1"));
                    if (pageSource != null)
                    {
                        doc = new HtmlDocument();
                        doc.LoadHtml(pageSource);
                    }
                }
                catch (Exception ex)
                {
                    string msg = "At DownloadNameRules(). Error found when downloading the name abbreviations file. " + ex.Message;
                    Logger.Log(msg);
                }
            }
            HtmlNode xmpNode = doc.DocumentNode.SelectSingleNode("//xmp");
            if (xmpNode == null)
            {
                string msg = string.Format("At DownloadNameRules(). Can not get the name abbreviation in file.");
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }
            string content = xmpNode.InnerText;

            string[] lines = content.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            int startLine = 0;
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("MEANING") && lines[i].Contains("ABBREVATION"))
                {
                    startLine = i + 1;
                    break;
                }
            }
            for (int i = startLine; i < lines.Length; i++)
            {
                string singleLine = lines[i];
                if (!singleLine.Contains("   "))
                {
                    continue;
                }
                string[] names = singleLine.Split(new[] { "   " }, StringSplitOptions.RemoveEmptyEntries);
                if (names.Length != 2)
                {
                    string msg = string.Format("At DownloadNameRules(). Irregular name and abbrevation at line: {0} in 'Abbrevation file'. Ignore it.\r\n\t\t\t\t Line content:{1}.", i, singleLine);
                    Logger.Log(msg);
                    continue;
                }
                if (!namesAbbs.ContainsKey(names[0].Trim()))
                {
                    namesAbbs.Add(names[0].Trim(), names[1].Trim());
                }
                else
                {
                    string msg = string.Format("At DownloadNameRules(). Repeated name at line: {0} in Abbrevation file. Line content:{1}.", i, singleLine);
                    Logger.Log(msg);
                }
            }
        }

        private bool CheckIfHShare(string securityDescription)
        {
            if (string.IsNullOrEmpty(securityDescription))
            {
                return false;
            }

            if (securityDescription.EndsWith(" H"))
            {
                return true;
            }

            return false;
        }

        private void FormatBulkNamesForHShare(DataStreamRicCreationBulkTemplate bulkTemp)
        {
            string hShareEnding = "'H'";
            if (bulkTemp.Name2.Trim() != "")
            {
                bulkTemp.Name2 = bulkTemp.Name2.Substring(0, 21) + hShareEnding;
            }
            else if (bulkTemp.Name1.Trim().Length > 20)
            {
                bulkTemp.Name2 = hShareEnding.PadRight(24, ' ');
            }
            else
            {
                bulkTemp.Name1 = bulkTemp.Name1.Substring(0, 21) + hShareEnding;
            }
        }

        #endregion

        #region For Filed SEDOL

        private string jSessionId = string.Empty;

        /// <summary>
        /// To validate everytime HTTPS Certificate
        /// </summary>
        /// <param name="senter"></param>
        /// <param name="certificate"></param>
        /// <param name="chain"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private static bool CheckValidationResult(object senter, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }

        private string GetSedol(string isin)
        {
            if (string.IsNullOrEmpty(isin))
            {
                return "";
            }

            try
            {
                if (!loginSuccess)
                {
                    LoginToWebsite();
                    loginSuccess = true;
                }
                string response = QuerySedol(isin);
                return ParseSedol(response, isin);
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GetSedol(). Error found in getting SEDOL for ISIN:{0}. Error Message: {1}. ", isin, ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return "";
            }
        }

        /// <summary>
        /// Connect to the website and get the Cookie
        /// </summary>
        private void LoginToWebsite()
        {
            string username = Encode(configObj.Username);
            string uri = string.Format("https://www.unavista.londonstockexchange.com/datasolutions/dwr/call/plaincall/LoginHelper.doLogin.dwr");
            string postData = string.Format("callCount=1\r\nwindowName=unavista_datasolutions\r\nc0-scriptName=LoginHelper\r\nc0-methodName=doLogin\r\nc0-id=0\r\nc0-param0=string:{0}\r\nc0-param1=null:null\r\nc0-param2=string:{1}\r\nc0-param3=boolean:false\r\nc0-param4=string:Google%20Chrome\r\nc0-param5=null:null\r\nc0-param6=boolean:false\r\nc0-param7=null:null\r\nc0-param8=null:null\r\nc0-param9=string:11.9.0.23449\r\nbatchId=2\r\npage=%2Fdatasolutions%2Flogin.html\r\nhttpSessionId=\r\nscriptSessionId=", username, configObj.Password);
            ServicePointManager.ServerCertificateValidationCallback = CheckValidationResult;
            try
            {
                HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.ContentType = "text/plain";
                request.UserAgent = "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
                request.Method = "POST";
                request.KeepAlive = true;
                request.AllowAutoRedirect = true;
                request.Referer = "https://www.unavista.londonstockexchange.com/datasolutions/login.html";
                request.Accept = "*/*";
                request.ContentLength = 2080;

                request.CookieContainer = cookies;
                request.Credentials = CredentialCache.DefaultCredentials;


                byte[] buf = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);

                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();
                string jsessionid = httpResponse.GetResponseHeader("Set-Cookie");

                Regex regex = new Regex("JSESSIONID=(?<JSessionId>.*?); Path");
                MatchCollection matches = regex.Matches(jsessionid);
                foreach (Match match in matches)
                {
                    jSessionId = match.Groups["JSessionId"].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error in LoginToWebsite():" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private string QuerySedol(string isin)
        {
            string uri = "https://www.unavista.londonstockexchange.com/datasolutions/dwr/call/plaincall/ClusterListHelper.loadClusterList.dwr";
            string postData;

            string[] postContent =
            {             
                "callCount=1", 
                "windowName=unavista_datasolutions", 
                "c0-scriptName=ClusterListHelper", 
                "c0-methodName=loadClusterList", 
                "c0-id=0", 
                "c0-param0=number:2199194", 
                "c0-param1=number:2199194", 
                "c0-param2=number:1085", 
                "c0-param3=number:10", 
                "c0-param4=boolean:false", 
                "c0-param5=null:null", 
                "c0-param6=null:null", 
                "c0-param7=array:[]", 
                "c0-param8=array:[]", 
                "c0-param9=null:null", 
                "c0-e2=string:(", 
                "c0-e3=string:ts.%5BISIN%5D", 
                "c0-e4=string:OR", 
                "c0-e5=string:false", 
                "c0-e6=string:" + isin, 
                "c0-e7=string:string", 
                "c0-e8=null:null", 
                "c0-e9=string:%3D", 
                "c0-e1=Object_Object:{openparen:reference:c0-e2, field:reference:c0-e3, logicaloperator:reference:c0-e4, unicode:reference:c0-e5, value:reference:c0-e6, datatype:reference:c0-e7, closeparen:reference:c0-e8, operator:reference:c0-e9}", 
                "c0-e11=null:null", 
                "c0-e12=string:ts.%5BSEDOL%5D", 
                "c0-e13=string:OR", 
                "c0-e14=string:false", 
                "c0-e15=string:" + isin, 
                "c0-e16=string:string", 
                "c0-e17=null:null", 
                "c0-e18=string:%3D", 
                "c0-e10=Object_Object:{openparen:reference:c0-e11, field:reference:c0-e12, logicaloperator:reference:c0-e13, unicode:reference:c0-e14, value:reference:c0-e15, datatype:reference:c0-e16, closeparen:reference:c0-e17, operator:reference:c0-e18}", 
                "c0-e20=null:null", 
                "c0-e21=string:ts.%5BprevSEDOL%5D", 
                "c0-e22=null:null", 
                "c0-e23=string:false", 
                "c0-e24=string:" + isin, 
                "c0-e25=string:string", 
                "c0-e26=string:)", 
                "c0-e27=string:%3D", 
                "c0-e19=Object_Object:{openparen:reference:c0-e20, field:reference:c0-e21, logicaloperator:reference:c0-e22, unicode:reference:c0-e23, value:reference:c0-e24, datatype:reference:c0-e25, closeparen:reference:c0-e26, operator:reference:c0-e27}", 
                "c0-param10=array:[reference:c0-e1,reference:c0-e10,reference:c0-e19]", 
                "c0-param11=boolean:false", 
                "c0-param12=boolean:false", 
                "c0-param13=boolean:false", 
                "c0-param14=null:null", 
                "c0-param15=null:null", 
                "c0-param16=null:null", 
                "c0-param17=array:[]", 
                "c0-param18=boolean:false", 
                "c0-param19=boolean:false", 
                "c0-param20=null:null", 
                "c0-param21=null:null", 
                "c0-param22=null:null", 
                "c0-param23=null:null", 
                "batchId=15", 
                "page=%2Fdatasolutions%2Funavistalanding.html", 
                "httpSessionId=" + jSessionId, 
                "scriptSessionId=" 
            };
            postData = string.Join("\r\n", postContent);
            try
            {
                HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.ContentType = "text/plain";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
                request.Method = "POST";
                request.KeepAlive = true;
                request.AllowAutoRedirect = true;
                request.Referer = "https://www.unavista.londonstockexchange.com/datasolutions/unavistalanding.html";
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                request.CookieContainer = cookies;
                request.Credentials = CredentialCache.DefaultCredentials;

                byte[] buf = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);

                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();

                StreamReader sr = new StreamReader(httpResponse.GetResponseStream());

                string content = sr.ReadToEnd();

                return content;

            }
            catch (Exception ex)
            {
                string msg = "Error found in QuerySedol():" + ex;
                Logger.Log(msg, Logger.LogType.Error);
                return "";
            }

        }

        private string ParseSedol(string response, string isin)
        {
            string pattern = string.Format(@"data:\[\[.*?{0}.*?\]\]", isin);
            Regex regex = new Regex(pattern);
            Match match = regex.Match(response);
            if (match.Success)
            {
                string[] dataList = match.Value.Split(new[] { "\",\"" }, StringSplitOptions.RemoveEmptyEntries);
                if (dataList.Length > 8)
                {
                    return dataList[7];
                }
            }
            Logger.Log("At ParseSedol(). Can not get SEDOL from website.", Logger.LogType.Error);
            return "";
        }

        /// <summary>
        /// Encode text to bits
        /// </summary>
        /// <param name="encodeText">text to encode</param>
        /// <returns>encoded result</returns>
        public static string Encode(string encodeText)
        {
            return Regex.Replace(encodeText, "[^a-zA-Z0-9]",
                match =>
                    "%" + BitConverter.ToString(Encoding.GetEncoding("euc-kr").GetBytes(match.Value)).Replace("-", "%"));
        }

        #endregion

        #endregion

        #region Generate File

        private void GenerateFile()
        {
            try
            {
                string outputFolder = Path.Combine(configObj.OutputPath, DateTime.Today.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }
                AddResult("Output Folder", outputFolder, "Output Folder");

                string template = string.Empty;
                string dataLine = string.Empty;
                string fileName = string.Empty;
                string filePath = string.Empty;

                if (!string.IsNullOrEmpty(normalOutputLine))
                {
                    template = InitializeMacTemplate();
                    dataLine = "[DataStreamRicCreationInfoLine]";
                    template = template.Replace(dataLine, normalOutputLine);
                    fileName = string.Format("HongKong_{0}.MAC", configObj.Date);
                    filePath = Path.Combine(outputFolder, fileName);
                    File.WriteAllText(filePath, template, Encoding.UTF8);
                    AddResult("MAC File", filePath, "MAC File");

                    string msg = string.Format("Generate normal RIC creation output file...OK!");
                    Logger.Log(msg);
                }
                if (!string.IsNullOrEmpty(hShareOutputLine))
                {
                    template = InitializeMacTemplateHShare();
                    dataLine = "[DataStreamRicCreationInfoLine]";
                    template = template.Replace(dataLine, hShareOutputLine);
                    fileName = string.Format("HongKongHshareRicCreation_{0}.MAC", configObj.Date);
                    filePath = Path.Combine(outputFolder, fileName);
                    File.WriteAllText(filePath, template, Encoding.UTF8);
                    AddResult("H share MAC File", filePath, "H share MAC File");

                    string msg = string.Format("Generate H-share RIC creation output file...OK!");
                    Logger.Log(msg);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("At GenerateFile(). Error found in generating MAC file. Error message:{0}.", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private string InitializeMacTemplate()
        {

            string template = ";Hongkong Basic Add\r\n"
                            + ";Start on Primary Menu\r\n\r\n"
                            + "LOOP:\r\n" + "IF &ENDOFDATA=FALSE THEN\t\r\n"
                            + "INPUT SEQN,MNEM,SEDOL,NAME1,NAME2,BDATE,DEF_MKT,NOSH,MNEM2,ISIN,CNAME,GEOG,DCUR\r\n\r\n"
                            + "Send (\"14\") ;select EQUITY\r\n" + "Send (\"S\")  ;go to NEWST\r\n"
                            + "Send (\"[DOWN][DOWN][DOWN]048YR\")  ;Select Hongkong, Domestic Quote, Research \r\n"
                            + "Send (\"[ENTER]\")\r\n"
                            + "SET DSCD TO mid$(&screen,182,6)\r\n"
                            + "Send (SEQN+MNEM+MNEM2+\"[TAB]\"+SEDOL+NAME1+NAME2+\"[DOWN][DOWN][DOWN]\"+DEF_MKT+\"[DOWN]D\"+BDATE)\r\n"
                            + "Send (\"[ENTER]\")\r\n" + "Send (\"C\")\r\n" + "Send (\"[ENTER]\")\r\n"
                            + "Send (NOSH+\"[ENTER]\")\r\n"
                            + "Send (\"C\")\r\n"
                            + "Send (CNAME+\"[DOWN]116[ENTER]\")\r\n"
                            + "Send (\"Y\")\r\n\r\n"
                            + ";Set QFLAG\r\n"
                            + "Send (\"[HOME]QFLAG\")\r\n"
                            + "Send ( MNEM )\r\n"
                            + "Send (\"YY\")\r\n\r\n"
                            + ";Set MAJOR\r\n"
                            + "Send (\"[HOME]MAJOR\")\r\n"
                            + "Send ( MNEM )\r\n"
                            + "Send (\"Y\")\r\n"
                            + "Send (\"S\")\r\n"
                            + "Send (\"Y\")\r\n\r\n"
                            + ";Set PRIME flag\r\n"
                            + "Send (\"[PF3]\")\r\n"
                            + "Send (\"[TAB][TAB]\"+MNEM)\r\n"
                            + "Send (\"Y\")\r\n"
                            + "Send (\"S\")\r\n"
                            + "Send (\"S\")\r\n"
                            + "Send (\"Y\")\r\n\r\n"
                            + ";CTREE\r\n"
                            + "Send (\"[PF3]\")\r\n"
                            + "Send (\"[PF3]\")\r\n"
                            + "Send (\"[HOME]CTREE\")\r\n"
                            + "Send ( MNEM+\"[DOWN]Y\" )\r\n"
                            + "Send (\"Y\")\r\n\r\n"
                            + ";BASIC\r\n"
                            + "Send (\"[HOME]BASIC\") \r\n"
                            + "Send ( MNEM )\r\n\r\n"
                            + ";INDEX\r\n" + "Send (\"[HOME]INDEX\")\r\n"
                            + "Send ( DSCD )\r\n"
                            + "Send (\"[DOWN][DOWN][DOWN]I\"+MNEM2+\"E[TAB][TAB][TAB][TAB][TAB][TAB][TAB][TAB][TAB]I\"+ISIN+\"V E\")\r\n"
                            + "Send ( \"Y\" )\r\n\r\n"
                            + "GOTO LOOP\r\n"
                            + "ENDIF\r\n"
                            + "END\r\n\r\n"
                            + "LIST:\r\n"
                            + "DATA\r\n\r\n"
                            + "[DataStreamRicCreationInfoLine]\r\n"
                            + "ENDDATA";
            return template;
        }

        private string InitializeMacTemplateHShare()
        {

            string template = ";Hongkong Basic Add\r\n"
                            + ";Start on Primary Menu\r\n\r\n"
                            + "LOOP:\r\n" + "IF &ENDOFDATA=FALSE THEN\t\r\n"
                            + "INPUT SEQN,MNEM,SEDOL,NAME1,NAME2,BDATE,DEF_MKT,NOSH,MNEM2,ISIN,CNAME,GEOG,DCUR\r\n\r\n"
                            + "Send (\"14\") ;select EQUITY\r\n" + "Send (\"S\")  ;go to NEWST\r\n"
                            + "Send (\"[DOWN][DOWN][DOWN]048YR\")  ;Select Hongkong, Domestic Quote, Research \r\n"
                            + "Send (\"[ENTER]\")\r\n"
                            + "SET DSCD TO mid$(&screen,182,6)\r\n"
                            + "Send (SEQN+MNEM+MNEM2+\"[TAB]\"+SEDOL+NAME1+NAME2+\"[DOWN][DOWN][DOWN]\"+DEF_MKT+\"[DOWN]D\"+BDATE)\r\n"
                            + "Send (\"[ENTER]\")\r\n" + "Send (\"C\")\r\n" + "Send (\"[ENTER]\")\r\n"
                            + "Send (NOSH+\"[ENTER]\")\r\n"
                            + "Send (\"C\")\r\n"
                            + "Send (CNAME+\"[DOWN]116[ENTER]\")\r\n"
                            + "Send (\"Y\")\r\n\r\n"
                            + ";Set QFLAG\r\n"
                            + "Send (\"[HOME]QFLAG\")\r\n"
                            + "Send ( MNEM )\r\n"
                            + "Send (\"YY\")\r\n\r\n"
                            + ";Set MAJOR\r\n"
                            + "Send (\"[HOME]MAJOR\")\r\n"
                            + "Send ( MNEM )\r\n"
                            + "Send (\"Y\")\r\n"
                            + "Send (\"S\")\r\n"
                            + "Send (\"Y\")\r\n\r\n"
                            + ";Set PRIME flag\r\n"
                            + "Send (\"[PF3]\")\r\n"
                            + "Send (\"[TAB][TAB]\"+MNEM)\r\n"
                            + "Send (\"Y\")\r\n"
                            + "Send (\"S\")\r\n"
                            + "Send (\"S\")\r\n"
                            + "Send (\"Y\")\r\n\r\n"
                            + ";CTREE\r\n"
                            + "Send (\"[PF3]\")\r\n"
                            + "Send (\"[PF3]\")\r\n"
                            + "Send (\"[HOME]CTREE\")\r\n"
                            + "Send ( MNEM+\"[DOWN]Y\" )\r\n"
                            + "Send (\"Y\")\r\n\r\n"
                            + ";BASIC\r\n"
                            + "Send (\"[HOME]BASIC\") \r\n"
                            + "Send ( MNEM+\"[DOWN][DOWN]\"+GEOG+\"[DOWN][DOWN][DOWN][DOWN][DOWN][DOWN]\"+DCUR )\r\n\r"
                            + "Send (\"[Enter]\")\r\n"
                            + "Send (\"Y\")\r\n"
                            + "Send (\"[Enter]\")\r\n\r\n"
                            + ";INDEX\r\n" + "Send (\"[HOME]INDEX\")\r\n"
                            + "Send ( DSCD )\r\n"
                            + "Send (\"[DOWN][DOWN][DOWN]I\"+MNEM2+\"E[TAB][TAB][TAB][TAB][TAB][TAB][TAB][TAB][TAB]I\"+ISIN+\"V E\")\r\n"
                            + "Send ( \"Y\" )\r\n\r\n"
                            + "GOTO LOOP\r\n"
                            + "ENDIF\r\n"
                            + "END\r\n\r\n"
                            + "LIST:\r\n"
                            + "DATA\r\n\r\n"
                            + "[DataStreamRicCreationInfoLine]\r\n"
                            + "ENDDATA";
            return template;
        }

        #endregion

    }
}
