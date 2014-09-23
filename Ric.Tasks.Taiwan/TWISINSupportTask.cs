using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Taiwan
{
    #region Model & Config

    public class DbConfig
    {
        public string ConnectionString { get; set; }
    }

    [ConfigStoredInDB]
    public class TWISINSupportTaskConfig
    {
        [StoreInDB]
        [Description("Path of the result excel file")]
        [DisplayName("Result file directory")]
        public string ResultFileDir { get; set; }

        [StoreInDB]
        [Description("Name of the equity option table file")]
        [DisplayName("Equity option table file")]
        public string EquityOptionTableFile { get; set; }
    }

    public class RicStructureSummary
    {
        public List<FutureRicStructure> FutureRicStructureList { get; set; }
        public List<OptionRicStructure> OptionRicStructureList { get; set; }
    }

    public class FutureRicStructure
    {
        public string FutureType { get; set; }
        public string OffclCodeIndexes { get; set; }//Split by ,
        public string PreMonthCodeStr { get; set; }//For 每周到期期货， this string can be configured as MTX[1]W, in which <1> should be replaced by the first digit number
    }

    public class OptionRicStructure
    {
        public string OptionType { get; set; }
        public string SourcePrefix { get; set; } //For 臺指選擇權買權/臺指選擇權賣權, this field value is "TXO"
        public string ReplacedPrefix { get; set; }//For 臺指選擇權買權/臺指選擇權賣權, this field value is "TX"
        public string Suffix { get; set; }
    }

    public class MonthConversion
    {
        public string Month { get; set; }
        public string OffclCodeMonthCode { get; set; }
        public string RICContractMonthCode { get; set; }
    }

    public class TWISINOffclCode : IComparable<TWISINOffclCode>
     {
         public string Ric { get; set; }
         public string OffclCode { get; set; }
         public string ISIN { get; set; }
         public string FutureOptionType { get; set; }

         #region IComparable<TWISINOffclCode> Members

         public int CompareTo(TWISINOffclCode other)
         {
             return this.OffclCode.CompareTo(other.OffclCode);
         }

         #endregion
     }

    #endregion

    #region

    class TWISINSupportTask:GeneratorBase
    {
        #region Properties

        private const string DbConnectionStringConfigPath = ".\\DbConfig.xml";
        private const string MonthCodeFilePath = ".\\Config\\TW\\TW_MonthContractCodeMap.xml";
        private string _equityOptionTablePath;
        private const string RicStructureSummaryFilePath = ".\\Config\\TW\\TWISINSupportTaskRicStructure.xml";

        private Dictionary<string, string> _monthContractCodeMap;
        private Dictionary<string, string> _codeRicRootMap;
        private RicStructureSummary _ricStructureSummaryObj;
        private TWISINSupportTaskConfig _configObj;
        private static string _connectionString = string.Empty;

        #endregion

        #region GeneratorBase

        protected override void Initialize()
        {
            base.Initialize();
            try
            {
                _configObj = Config as TWISINSupportTaskConfig;
                _connectionString = (ConfigUtil.ReadConfig(DbConnectionStringConfigPath, typeof(DbConfig)) as DbConfig).ConnectionString; 

                var monthCodeConversionList = ConfigUtil.ReadConfig(MonthCodeFilePath, typeof(List<MonthConversion>)) as List<MonthConversion>;
                _monthContractCodeMap = GetMonthContractCodeMap(monthCodeConversionList);

                _equityOptionTablePath = _configObj.ResultFileDir + _configObj.EquityOptionTableFile;
                _codeRicRootMap = GetEquityOptionCodeRicRootMap(_equityOptionTablePath);

                _ricStructureSummaryObj = ConfigUtil.ReadConfig(RicStructureSummaryFilePath, typeof(RicStructureSummary)) as RicStructureSummary;
            }
            catch (Exception ex)
            {
                LogMessage("Error happens when initializing task... Ex: " + ex.Message, Logger.LogType.Error);
            }
            AddResult("Log File", Logger.FilePath, "log");
        }

        protected override void Start()
        {
            try
            {
                List<TWISINOffclCode> listFromDB = GetAll(_connectionString);
                HtmlNode rootNode = GetHtmlSource();
                List<TWISINOffclCode> listFromSource = GetISINList(rootNode);
                List<TWISINOffclCode> newAddedList = GetNewAddISINOffclCode(listFromSource, listFromDB);

                foreach (TWISINOffclCode isin in newAddedList)
                {
                    insert(isin, _connectionString);
                }

                foreach (TWISINOffclCode isinOffclCode in newAddedList)
                {
                    isinOffclCode.Ric = ComposeRic(isinOffclCode);
                }

                string gedaFilePath = Path.Combine(_configObj.ResultFileDir, string.Format("TW_ISIN_{0}.txt", DateTime.Now.ToString("ddMMMhhmmss", new CultureInfo("en-US"))));
                GenerateGEDAFile(gedaFilePath, newAddedList);
                AddResult("Geda file", gedaFilePath, "geda");
            }
            catch (Exception ex)
            {
                LogMessage(ex.Message, Logger.LogType.Error);
                LogMessage(ex.StackTrace, Logger.LogType.Error);
                throw (ex);
            }
        }

        #endregion

        #region Download and parse TW website to get the ISIN 

        private HtmlNode GetHtmlSource()
        {
            string source = WebClientUtil.GetPageSource(null, "http://isin.twse.com.tw/isin/C_public.jsp?strMode=6",
                180000, "strMode=6", Encoding.GetEncoding("big5"));
            var document = new HtmlDocument();
            document.LoadHtml(source);
            return document.DocumentNode;
        }

        private List<TWISINOffclCode> GetISINList(HtmlNode rootNode)
        {
            var offclCodeISINList = new List<TWISINOffclCode>();
            HtmlNodeCollection collection = rootNode.SelectNodes("//tr");
            string kindOfOption="";
            foreach (var node in collection.Skip(1))
            {
                if (node.SelectSingleNode("./td").Attributes["colspan"] != null)
                {
                    kindOfOption=MiscUtil.GetCleanTextFromHtml(node.InnerText);
                    continue;
                }
                var ISINOffclCode = new TWISINOffclCode();
                HtmlNodeCollection tempCollection = node.SelectNodes("./td");
                int index = tempCollection[0].InnerText.IndexOf(" ");
                if (index < 0)
                {
                    index = tempCollection[0].InnerText.IndexOf("　");
                }
                ISINOffclCode.OffclCode = tempCollection[0].InnerText.Substring(0, index);
                ISINOffclCode.ISIN = tempCollection[1].InnerText;

                ISINOffclCode.FutureOptionType = kindOfOption;
                offclCodeISINList.Add(ISINOffclCode);
            }
            return offclCodeISINList;
        }

        #endregion

        private void GenerateGEDAFile(string filePath, List<TWISINOffclCode> isinRicList)
        {
            var sb = new StringBuilder();
            sb.Append("RIC\t#INSTMOD_OFFC_CODE2\t#INSTMOD_#DDS_ISIN_CODE\r\n");
            foreach (TWISINOffclCode item in isinRicList)
            {
                sb.Append(item.Ric);
                sb.Append("\t");
                sb.Append(item.ISIN);
                sb.Append("\t");
                sb.Append(item.ISIN);
                sb.Append("\r\n");
            }

            String dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        //Compare the list got from TW website and the list stored in the DB to get the newly added ISIN, using Linq
        private List<TWISINOffclCode> GetNewAddISINOffclCode(List<TWISINOffclCode> listFromSourcePage, List<TWISINOffclCode> listFromDB)
        {
            List<TWISINOffclCode> newAddISINOffclCodeList = 
                (from sourceISINOffclCode in listFromSourcePage
                where (!(new HashSet<string>(listFromDB.Select(x => x.OffclCode.Trim())).Contains(sourceISINOffclCode.OffclCode.Trim())))
                select sourceISINOffclCode
                ).ToList();

            return newAddISINOffclCodeList;
        }

        #region Compose Ric

        //compose Ric based on the offcl code and the future/ option type
        private string ComposeRic(TWISINOffclCode isinOffclCode)
        {
            string ric = string.Empty;
            if (isinOffclCode.FutureOptionType.Contains("期貨"))
            {
                FutureRicStructure stru = FindFutureRicStructure(isinOffclCode.FutureOptionType, _ricStructureSummaryObj.FutureRicStructureList);
                ric = ComposeFutureRic(isinOffclCode.OffclCode, stru.OffclCodeIndexes, stru.PreMonthCodeStr);
            }
            else if ((isinOffclCode.FutureOptionType == "股票選擇權買權") || (isinOffclCode.FutureOptionType == "股票選擇權賣權"))
            {
                ric = ComposeRicForSpecialCase(isinOffclCode.OffclCode, _codeRicRootMap);
            }
            else
            {
                OptionRicStructure stru =FindOptionRicStructure(isinOffclCode.FutureOptionType,_ricStructureSummaryObj.OptionRicStructureList);
                ric = ComposeOptionRic(isinOffclCode.OffclCode, stru.SourcePrefix, stru.ReplacedPrefix, stru.Suffix);
            }

            return ric;
        }

        private FutureRicStructure FindFutureRicStructure(string futureOptionType, List<FutureRicStructure> futureRicList)
        {
            foreach (FutureRicStructure stru in futureRicList.Where(stru => stru.FutureType.Trim() == futureOptionType.Trim()))
            {
                return stru;
            }
            LogMessage(string.Format("Cannot find this type {0} in FutureRicStructureList, please check the ric structure config file",futureOptionType), Logger.LogType.Error);
            return null;
        }

        private OptionRicStructure FindOptionRicStructure(string futureOptionType, IEnumerable<OptionRicStructure> optionRicList)
        {
            foreach (OptionRicStructure stru in optionRicList.Where(stru => stru.OptionType.Trim() == futureOptionType.Trim()))
            {
                return stru;
            }
            LogMessage(string.Format("Cannot find this type {0} in OptionRicStructureList, please check the ric structure config file",futureOptionType), Logger.LogType.Error);
            return null;
        }

        private string ComposeFutureRic(string offclCode, string offclCodeIndexes, string preMonthCodeStr)
        {
            string ric = string.Empty;
            ric += GetSubString(offclCode, offclCodeIndexes);
            ric += GetPreMonthCodeStr(offclCode, preMonthCodeStr);
            ric += GetContractMonthCode(offclCode, _monthContractCodeMap);
            ric += offclCode[offclCode.Length - 1];
            return ric;
        }

        private string GetPreMonthCodeStr(string offclCode, string preMonthCodeExp)
        {
            Regex r = new Regex("[(?<index>\\d{1,})]");
            Match m = r.Match(preMonthCodeExp);
            if (!string.IsNullOrEmpty(m.Value))
            {
                try
                {
                    string indexStr = m.Value;
                    int index = int.Parse(indexStr);
                    preMonthCodeExp = preMonthCodeExp.Replace(string.Format("[{0}]",m.Value), string.Format("{0}", GetDigitNum(index, offclCode)));
                }
                catch (Exception ex)
                {
                    LogMessage("Error happens when getting preMonthCodeStr, offclCode: "+offclCode+"preMonthCodeExp: "+preMonthCodeExp+"Ex: "+ex.Message, Logger.LogType.Error);
                }
            }
            return preMonthCodeExp;
        }

        //Find the index-th digital number
        private string GetDigitNum(int index, string offclCode)
        {
            Regex r = new Regex("(?<digit>\\d{"+index+"}.*?)");
            Match m = r.Match(offclCode);
            return m.Groups["digit"].Value.Substring(index - 1);
        }

        private string GetContractMonthCode(string offclCode, Dictionary<string,string>monthContractCodeMap)
        {
            string offclMonthCode = offclCode[offclCode.Length - 2].ToString().Trim();
            if (!monthContractCodeMap.Keys.Contains(offclMonthCode))
            {
                LogMessage("Cannot find the month contract code for " + offclCode, Logger.LogType.Warning);
                return string.Empty;
            }
            return monthContractCodeMap[offclMonthCode];
        }

        private string GetSubString(string offclCode, string offclCodeIndexes)
        {
            string subStr = string.Empty;
            if (!string.IsNullOrEmpty(offclCodeIndexes))
            {
                string[] indexes = offclCodeIndexes.Split(',');
                if (indexes.Length > 0)
                {
                    foreach (string indexStr in indexes)
                    {
                        try
                        {
                            int index = int.Parse(indexStr.Trim());
                            subStr += offclCode[index];
                        }
                        catch (Exception ex)
                        {
                            LogMessage("Error happens when getting substring for offclCode, please check OffclCodeIndexes filed value in RicStructureSummary config file. Ex: " + ex.Message, Logger.LogType.Error);
                        }
                    }
                }
            }
            return subStr;
        }

        private string ComposeOptionRic(string offclCode, string sourcePrefix, string replacedPrefix, string suffix)
        {
            Regex r = new Regex("(?<prefix>[a-zA-Z]{1,})");
            Match m = r.Match(offclCode);
            string ric = m.Groups["prefix"].Value;
            if (!string.IsNullOrEmpty(sourcePrefix))
            {
                ric = ric.Replace(sourcePrefix, replacedPrefix);
            }
            ric += RemoveZeros(offclCode);
            ric += offclCode.Substring(offclCode.Length - 2, 2);
            ric += suffix;

            return ric;
        }

        //股票選擇權買權/股票選擇權賣權
        private string ComposeRicForSpecialCase(string offclCode, Dictionary<string,string> codeRicRootMap)
        {
            string ric = string.Empty;

            //Convert the first three characters to numbers by look up the Equity Option table.
            //Remove the zeros before the first non-zero number in the orginal OFFCL_CODE, ignore the last digit if it is zero
            if (!offclCode.StartsWith("TX"))
            {
                string code = offclCode.Substring(0, 3);
                if (!codeRicRootMap.Keys.Contains(code))
                {
                    LogMessage(string.Format("Error happens for getting root code for {0}, cannot find the code {1} in equity option table.", offclCode, code), Logger.LogType.Warning);
                }

                else
                {
                    ric += codeRicRootMap[code];
                    string nonZeroNum = RemoveZeros(offclCode);
                    if (nonZeroNum.EndsWith("0"))
                    {
                        nonZeroNum = nonZeroNum.Remove(nonZeroNum.Length - 1);
                    }
                    ric += nonZeroNum;
                    ric += offclCode.Substring(offclCode.Length - 2, 2);
                }
            }
                //If the OFFCL_CODE not starts with "TX", Replace the left fourth position 0 with "W",
            else
            {
                ric = offclCode;
                if (ric[3] == '0')
                {
                    ric = ric.Remove(3, 1);
                    ric = ric.Insert(3, "W");
                }
            }
            ric = ric + ".TM";

            return ric;
        }

        //Remove the zeros before the first non-zero number in the original OFFCL_CODE
        private string RemoveZeros(string offclCode)
        {
            Regex r = new Regex("0{0,}(?<nonZeroNum>\\d{1,})");
            Match m = r.Match(offclCode);

            return m.Groups["nonZeroNum"].Value;
        }

        #endregion

        #region For DB Operations
        //Get all the items in DB
        private List<TWISINOffclCode> GetAll(string connectionString)
        {
            var infoList = new List<TWISINOffclCode>();
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    using (var comm = new SqlCommand("select * from ETI_TW_OFFCLCODE_ISIN", conn))
                    {
                        using (SqlDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    var info = new TWISINOffclCode
                                    {
                                        OffclCode = Convert.ToString(dr["offclCode"]),
                                        ISIN = Convert.ToString(dr["ISIN"])
                                    };
                                    infoList.Add(info);
                                }
                            }
                        }
                    }
                }
                return infoList;
            }
            catch (Exception ex)
            {
                LogMessage("Exception happens when getting TWISINOffclCode list from the DB. ex: " + ex.Message, Logger.LogType.Error);
                return null;
            }
        }

        //Insert an item to db
        private bool insert(TWISINOffclCode info, string connectionString)
        {
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    using (var comm = new SqlCommand())
                    {
                        comm.Connection = conn;
                        comm.CommandText = "insert into ETI_TW_OFFCLCODE_ISIN(ISIN,OFFCLCODE) values(@ISIN,@OFFCLCODE)";
                        comm.Parameters.Add(new SqlParameter("@ISIN", info.ISIN));
                        comm.Parameters.Add(new SqlParameter("@OFFCLCODE", info.OffclCode));

                        int rowAffected = comm.ExecuteNonQuery();

                        if (rowAffected == 0)
                        {
                            return false;
                        }
                        return true;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region For Initialize
        //Get the month contract code map
        private Dictionary<string,string> GetMonthContractCodeMap(List<MonthConversion> monthConversionList)
        {
            var monthContractCodeMap = new Dictionary<string, string>();
            foreach (MonthConversion codeConversion in monthConversionList)
            {
                if (string.IsNullOrEmpty(codeConversion.OffclCodeMonthCode))
                {
                    LogMessage("Please check the month code contract file.", Logger.LogType.Warning);
                    continue;
                }
                if (!monthContractCodeMap.Keys.Contains(codeConversion.OffclCodeMonthCode))
                {
                    monthContractCodeMap.Add(codeConversion.OffclCodeMonthCode, codeConversion.RICContractMonthCode);
                    if (string.IsNullOrEmpty(codeConversion.RICContractMonthCode))
                    {
                        LogMessage(string.Format("The value for {0} is empty. Please check the month code contract file.",codeConversion.RICContractMonthCode), Logger.LogType.Warning);
                    }
                }
            }
            return monthContractCodeMap;
        }

        //Get the equity option table
        private Dictionary<string, string> GetEquityOptionCodeRicRootMap(string equityOptionTablePath)
        {
            var equityOptionCodeRicRootMap = new Dictionary<string, string>();
            if(!File.Exists(equityOptionTablePath))
            {
                throw (new Exception(string.Format("Cannot find the Equity Option Table. {0}",equityOptionTablePath)));
            }
            string[] contents = File.ReadAllLines(equityOptionTablePath);
            foreach (string[] parts in contents.Where(line => !string.IsNullOrEmpty(line)).Select(line => line.Split(','))
                                               .Where(parts => !equityOptionCodeRicRootMap.Keys.Contains(parts[0])))
            {
                equityOptionCodeRicRootMap.Add(parts[0], parts[1]);
            }
            return equityOptionCodeRicRootMap;
        }

        #endregion

    }

    #endregion

}
