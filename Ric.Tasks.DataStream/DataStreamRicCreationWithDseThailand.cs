using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;
using System.Net;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Ric.Util;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Globalization;

namespace Ric.Tasks.DataStream
{
    #region [Config]
    [ConfigStoredInDB]
    class DataStreamRicCreationWithDseThailandConfig
    {
        [StoreInDB]
        [Category("Path")]
        [DefaultValue("D:\\DataStream\\RIC_Creation\\")]
        [Description("Path to save generated output file. E.g.D:\\DataStream\\RIC_Creation\\")]
        public string OutputPath { get; set; }

        [Category("Announcement Date")]
        [Description("Date format: MMdd. E.g. 1206")]
        public string Date { get; set; }

        [StoreInDB]
        [Category("SEDOL Login Info")]
        [DefaultValue("xxx.xxx@thomsonreuters.com")]
        [Description("Username for login to the SEDOL source website.")]
        public string Username { get; set; }

        [StoreInDB]
        [Category("SEDOL Login Info")]
        [Description("Password for login to the SEDOL source website.")]
        public string Password { get; set; }

        [StoreInDB]
        [Category("Proxy")]
        [DefaultValue("10.40.14.56")]
        [Description("Proxy IP address for assess to DSE.")]
        public string IP { get; set; }

        [StoreInDB]
        [Category("Proxy")]
        [DefaultValue("80")]
        [Description("Proxy port for assess to DSE.")]
        public string Port { get; set; }

        public DataStreamRicCreationWithDseThailandConfig()
        {
            Date = DateTime.Today.AddDays(-1).ToString("MMdd");
        }
    }
    #endregion

    class DataStreamRicCreationWithDseThailand : GeneratorBase
    {
        private static DataStreamRicCreationWithDseThailandConfig configObj = null;
        private List<string> listDownloadFtpCode = null;
        private List<string> listDownloadFtpFilePath = null;//list of download file name
        private List<string> listExtractedRecords = null;//get records from ftp files
        private bool loginSuccess = false;
        private string jSessionId = string.Empty;
        private CookieContainer cookies = new CookieContainer();
        private Dictionary<string, string> namesAbbs = new Dictionary<string, string>();

        protected override void Initialize()
        {
            configObj = Config as DataStreamRicCreationWithDseThailandConfig;
            listDownloadFtpCode = new List<string>() { 
                "0760",
                "9077",
                "0049",
                "EM01",
                "EM11"
            };
        }

        protected override void Start()
        {
            if (!PrepearConfig())
                return;

            Logger.Log("prepear config succeed.");

            listDownloadFtpFilePath = DownloadFilesFromFtp(listDownloadFtpCode);
            Logger.Log(string.Format("ftp file count:{0}",
                listDownloadFtpFilePath == null ? "0" : listDownloadFtpFilePath.Count.ToString()));

            listExtractedRecords = ExtractRecordsFromFiles(listDownloadFtpFilePath);
            Logger.Log(string.Format("extracted records count:{0}",
                listExtractedRecords == null ? "0" : listExtractedRecords.Count.ToString()));

            List<DataStreamRicCreationEntity> listDSRicCreation = FormatRecords(listExtractedRecords);
            Logger.Log(string.Format("format records count:{0}",
                listDSRicCreation == null ? "0" : listDSRicCreation.Count.ToString()));

            DownloadNameRules();

            List<ThailandTemplate> listThailandTemplate = FormatTemplates(listDSRicCreation);
            Logger.Log(string.Format("format records count:{0}",
                listThailandTemplate == null ? "0" : listThailandTemplate.Count.ToString()));

            GenetateFile(listThailandTemplate);
        }

        #region [generate bulk file]
        private void GenetateFile(List<ThailandTemplate> listThailandTemplate)
        {
            if (listThailandTemplate == null || listThailandTemplate.Count == 0)
            {
                Logger.Log("listThailandTemplate.count==0, no data need to ouput. ", Logger.LogType.Warning);
                return;
            }

            var sb = new StringBuilder();
            try
            {
                string fileFolder = Path.Combine(configObj.OutputPath, DateTime.Today.ToString("yyyy-MM-dd"));
                string filePath = Path.Combine(fileFolder, string.Format("Thailand_{0}.MAC", configObj.Date));

                foreach (var item in listThailandTemplate)
                {
                    sb.AppendFormat("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",\"{12}\",\"{13}\",\"{14}\",\"{15}\",\"{16}\",\"{17}\",\"{18}\",\"{19}\"",
                                     item.Mnem, item.Sedol, item.Name1, item.Name2,
                                     item.Bdate, item.DefMkt, item.Mnem2, item.Isin, item.Cname,
                                     item.Geog, item.Dcur, item.Remk, item.Secty, item.Grpi, item.Fname1, item.Mic,
                                     item.Fname2, item.Cusip, item.Qci, item.ParentRic);
                    sb.AppendLine();
                }

                if (!Directory.Exists(fileFolder))
                    Directory.CreateDirectory(fileFolder);

                string templateString = File.ReadAllText(@"Config\DataStream\Thailand.MAC", Encoding.ASCII);
                string resultString = string.Format(templateString, sb.ToString());
                File.WriteAllText(filePath, resultString, Encoding.ASCII);
                AddResult("Output Folder", fileFolder, "Output Folder");
                AddResult("MAC File", filePath, "MAC File");
                Logger.Log("Generate normal RIC creation output file...OK!");
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
        #endregion

        #region download rules
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
            string[] lines = content.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
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

                string[] names = singleLine.Split(new string[] { "   " }, StringSplitOptions.RemoveEmptyEntries);

                if (names.Length != 2)
                {
                    string msg = string.Format("At DownloadNameRules(). Irregular name and abbrevation at line: {0} in 'Abbrevation file'. Ignore it.\r\n\t\t\t\t Line content:{1}.", i, singleLine);
                    //Logger.Log(msg);
                    continue;
                }

                if (!namesAbbs.ContainsKey(names[0].Trim()))
                {
                    namesAbbs.Add(names[0].Trim(), names[1].Trim());
                }
                else
                {
                    string msg = string.Format("At DownloadNameRules(). Repeated name at line: {0} in Abbrevation file. Line content:{1}.", i, singleLine);
                    //Logger.Log(msg);
                }
            }

            Logger.Log(string.Format("download the rules count:{0}", namesAbbs == null ? "0" : namesAbbs.Keys.Count.ToString()));
        }
        #endregion

        #region format template
        private List<ThailandTemplate> FormatTemplates(List<DataStreamRicCreationEntity> listDSRicCreation)
        {
            List<ThailandTemplate> list = new List<ThailandTemplate>();
            string sedol = string.Empty;
            string name1 = string.Empty;
            string name2 = string.Empty;
            string suffix = string.Empty;
            try
            {
                if (listExtractedRecords == null || listExtractedRecords.Count == 0)
                    return null;
                foreach (var item in listDSRicCreation)
                {
                    ThailandTemplate template = new ThailandTemplate(item);
                    //Sedol 
                    if ((item.Sedol + "").Trim().Length == 0)
                    {
                        sedol = GetSedol(item.Isin);

                        if (string.IsNullOrEmpty(sedol))
                            sedol = "         ";//9 space
                        else
                            sedol = "UK" + sedol;
                    }
                    else
                    {
                        sedol = "UK" + item.Sedol;
                    }

                    template.Sedol = sedol;
                    //Name1 
                    //Name2 
                    //Cname 
                    //Fname1 
                    //Fname2
                    if ((item.AssetCategory + "").Trim().Equals("PRF") || (item.AssetCategory + "").Trim().Equals("CPR"))
                    {
                        FormatBulkNames(item.SecurityLongDescription.Trim(), 24, 24, nameInputType.Name, ref name1, ref name2, ref suffix);
                        template.NameSuffix = suffix;
                        template.Name1 = name1;
                        template.Name2 = name2;

                        FormatBulkNames(item.SecurityLongDescription.Trim(), 19, 24, nameInputType.Fname, ref name1, ref name2, ref suffix);
                        template.Fname1 = name1;
                        template.Fname2 = name2;
                    }
                    else
                    {
                        FormatBulkNames(item.CompanyName.Trim(), 24, 24, nameInputType.Name, ref name1, ref name2, ref suffix);
                        template.NameSuffix = suffix;
                        template.Name1 = name1;
                        template.Name2 = name2;

                        FormatBulkNames(item.CompanyName.Trim(), 19, 24, nameInputType.Fname, ref name1, ref name2, ref suffix);
                        template.Fname1 = name1;
                        template.Fname2 = name2;
                    }

                    if (item.AssetCategory.CompareTo("PRF") == 0 || item.AssetCategory.CompareTo("CPR") == 0)
                    {
                        FormatCName(template, item.CompanyName);
                    }
                    else
                    {
                        name1 = template.Name1.Trim();
                        name2 = template.Name2.Trim();

                        string names = string.Empty;
                        if (!name1.EndsWith("."))
                        {
                            names = name1 + " " + name2;
                        }
                        else
                        {
                            names = name1 + name2;
                        }
                        FormatCName(template, names);
                    }

                    if (IsValidThailandTemplate(item))
                        list.Add(template);
                }
                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private void FormatCName(ThailandTemplate bulkTemp, string names)
        {

            if (names.EndsWith(".") && (names.Length + bulkTemp.NameSuffix.Length) <= 24)
            {
                names += bulkTemp.NameSuffix;
            }
            else if ((!names.EndsWith(".")) && (names.Length + 1 + bulkTemp.NameSuffix.Length) <= 24)
            {
                if (names.EndsWith(" "))
                {
                    names += bulkTemp.NameSuffix;
                }
                else
                {
                    names += " " + bulkTemp.NameSuffix;
                }

            }
            else if (names.Length > 24)
            {
                names = names.Substring(0, 24);
            }

            bulkTemp.Cname = names.PadRight(24, ' ');
        }

        private void FormatBulkNames(string companyName, int nameLength1, int nameLength2, nameInputType type, ref string name1, ref string name2, ref string suffix)
        {
            string temp = string.Empty;
            if (string.IsNullOrEmpty(companyName))
            {
                name1 = temp.PadRight(nameLength1, ' ');
                name2 = temp.PadRight(nameLength2, ' ');
                return;
            }

            string inputConpanyName = companyName;
            companyName = companyName.ToUpper();

            string[] nameWords = companyName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            //Company Name only contains one word.
            if (nameWords.Length == 1)
            {
                string abbr = companyName;
                if (namesAbbs.ContainsKey(companyName))
                {
                    abbr = namesAbbs[companyName];
                }
                name1 = abbr.PadRight(nameLength1, ' ');
                name2 = temp.PadRight(nameLength2, ' ');
                return;
            }

            string nameSuffix = string.Empty;
            string nameLeft = TrimCompanyEndings(companyName, ref nameSuffix);
            suffix = nameSuffix;

            List<string> abbreviations = GetNameAbbreviations(nameLeft, nameLength1, nameLength2);

            string namePart1 = string.Empty;
            string namePart2 = string.Empty;

            bool reFormat = false;
            do
            {
                if (reFormat)
                {
                    if (abbreviations[abbreviations.Count - 1] == "CORP.")
                    {
                        suffix = "CORP.";
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
                if (formatedName.Length > (nameLength1 + nameLength2))
                {
                    if (reFormat)
                    {
                        string[] namesInput = DataStreamRicCreationNamesInput.Prompt(inputConpanyName, formatedName, type);
                        if (namesInput != null)
                        {
                            namePart1 = namesInput[0];
                            namePart2 = namesInput[1];
                        }
                        reFormat = false;
                    }
                    else
                    {
                        reFormat = true;
                    }
                }
                else if (formatedName.Length <= nameLength1)
                {
                    namePart1 = formatedName;
                }

                else
                {
                    int index = nameLength1;
                    int index2 = nameLength2;
                    char[] charArr = formatedName.ToCharArray();
                    if (charArr[index] == ' ')
                    {
                        index2 = index;
                    }
                    else
                    {
                        char middleChar = charArr[index];
                        if (middleChar == ' ' || middleChar == '.')
                        {
                            index--;
                        }
                        while (true)
                        {
                            middleChar = charArr[index];
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
                            string[] namesInput = DataStreamRicCreationNamesInput.Prompt(inputConpanyName, formatedName, type);
                            if (namesInput != null)
                            {
                                namePart1 = namesInput[0];
                                namePart2 = namesInput[1];
                            }
                            reFormat = false;
                        }
                        else
                        {
                            reFormat = true;
                        }
                    }
                }
            }
            while (reFormat);

            name1 = namePart1.PadRight(nameLength1, ' ');
            name2 = namePart2.PadRight(nameLength2, ' ');
        }

        private List<string> GetNameAbbreviations(string nameLeft, int nameLength1, int nameLength2)
        {

            List<string> abbrevations = new List<string>();

            List<string> nameLeftArr = nameLeft.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            //Rule 14 in requirement v3.0. Abbreviate from right to left.            

            if (nameLeft.Length <= nameLength1)
            {
                foreach (string subName in nameLeftArr)
                {
                    abbrevations.Add(subName + " ");
                }
                return abbrevations;
            }

            string firstWord = nameLeftArr[0] + " ";
            nameLeftArr.RemoveAt(0);

            while (nameLeftArr.Count > 0)
            {
                //string namePart = nameLeftArr[i];
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
                    //for (int k = 0; k <= nameLeftArr.Count - 1 - j; )
                    //{
                    //    nameLeftArr.RemoveAt(nameLeftArr.Count - 1);
                    //}

                    int k = nameLeftArr.Count - 1 - j;
                    while (k-- >= 0)
                    {
                        nameLeftArr.RemoveAt(nameLeftArr.Count - 1);
                    }

                    if (!string.IsNullOrEmpty(abbr))
                    {
                        abbrevations.Add(abbr);
                        string nameFormatTemp = firstWord + string.Join(" ", nameLeftArr.ToArray()) + " " + string.Join("", abbrevations.ToArray());
                        if (nameFormatTemp.Length <= nameLength1)
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

        private string TrimCompanyEndings(string nameLeft, ref string nameSuffix)
        {
            string[] endings = new string[] { "JSC", "JOINT STOCK COMPANY", "JOINT STOCK CO", "JOINT STOCK CORPORATION", "JOINT-STOCK CORPORATION", "PUBLIC LIMITED COMPANY", "INCORPORATED", "COMPANY", "LIMITED", "LTD", "CORPORATION", "CO.", "CORP", "INC", "LLC", "PLC.", "PLC", "Public Limited Company", "Public Limited Co", "Public Limited Co." };
            foreach (string ending in endings)
            {
                if ((nameLeft.Length > ending.Length) && ((nameLeft.Substring(nameLeft.Length - ending.Length - 1).Equals(" " + ending)) || (nameLeft.Substring(nameLeft.Length - ending.Length - 1).Equals("." + ending))))
                {
                    nameLeft = nameLeft.Replace(ending, "");
                    if (ending.Equals("JSC") || ending.Equals("JOINT STOCK COMPANY") || ending.Equals("JOINT STOCK CO") || ending.Equals("JOINT STOCK CORPORATION"))
                    {
                        nameSuffix = "JSC.";
                    }
                    else if (ending.Equals("CORPORATION"))
                    {
                        nameSuffix = "CORP.";
                    }
                    else if (ending.Equals("LIMITED") || ending.Equals("LTD"))
                    {
                        nameSuffix = "LTD.";
                    }
                    else if (ending.Equals("INCORPORATED"))
                    {
                        nameSuffix = "INCO.";
                    }
                    else if (ending.Equals("LLC"))
                    {
                        nameSuffix = "LLC.";
                    }
                    else if (ending.Equals("PLC.") || ending.Equals("PLC") || ending.Equals("Public Limited Company") || ending.Equals("Public Limited Co") || ending.Equals("Public Limited Co."))
                    {
                        nameSuffix = "PLC.";
                    }
                }
            }

            if (nameLeft.Contains(" JOINT STOCK COMPANY "))
            {
                nameLeft = nameLeft.Replace("JOINT STOCK COMPANY", "JSC.");
            }

            return nameLeft.Trim();
        }

        private string GetSedol(string isin)
        {
            string result = string.Empty;
            try
            {
                if ((isin + "").Trim().Length == 0)
                    return result;

                if (!loginSuccess)
                {
                    LoginToWebsite();
                    loginSuccess = true;
                }

                string response = QuerySedol(isin.Trim());
                result = ParseSedol(response, isin.Trim());

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return result;
            }
        }

        private string ParseSedol(string response, string isin)
        {
            string result = string.Empty;
            string pattern = string.Format(@"data:\[\[.*?{0}.*?\]\]", isin);
            Regex regex = new Regex(pattern);
            Match match = regex.Match(response);

            if (match.Success)
            {
                string[] dataList = match.Value.Split(new string[] { "\",\"" }, StringSplitOptions.RemoveEmptyEntries);
                if (dataList.Length > 8)
                {
                    return dataList[7];
                }
            }

            Logger.Log("At ParseSedol(). Can not get SEDOL from website.", Logger.LogType.Error);

            return result;
        }

        private string QuerySedol(string isin)
        {
            string uri = "https://www.unavista.londonstockexchange.com/datasolutions/dwr/call/plaincall/ClusterListHelper.loadClusterList.dwr";
            string postData = "";

            string[] postContent = new string[] {             
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
                string msg = "Error found in QuerySedol():" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return "";
            }
        }

        private void LoginToWebsite()
        {
            string username = Encode(configObj.Username);
            string uri = string.Format("https://www.unavista.londonstockexchange.com/datasolutions/login.html");
            string postData = string.Format("callCount=1\r\nwindowName=unavista_datasolutions\r\nc0-scriptName=LoginHelper\r\nc0-methodName=doLogin\r\nc0-id=0\r\nc0-param0=string:{0}\r\nc0-param1=null:null\r\nc0-param2=string:{1}\r\nc0-param3=boolean:false\r\nc0-param4=string:Google%20Chrome\r\nc0-param5=null:null\r\nc0-param6=boolean:false\r\nc0-param7=null:null\r\nc0-param8=null:null\r\nc0-param9=string:11.9.0.23449\r\nbatchId=2\r\npage=%2Fdatasolutions%2Flogin.html\r\nhttpSessionId=\r\nscriptSessionId=", username, configObj.Password);
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);

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

        public static string Encode(string encodeText)
        {
            return Regex.Replace(encodeText, "[^a-zA-Z0-9]", delegate(Match match) { return "%" + BitConverter.ToString(Encoding.GetEncoding("euc-kr").GetBytes(match.Value)).Replace("-", "%"); });
        }

        private static bool CheckValidationResult(object senter, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }

        private bool IsValidThailandTemplate(DataStreamRicCreationEntity item)
        {
            bool result = true;
            try
            {
                if ((item.ExchangeCode + "").Trim().Equals("RTS"))
                    result = false;

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
        #endregion

        #region [format records]
        private List<DataStreamRicCreationEntity> FormatRecords(List<string> listExtractedRecords)
        {
            List<DataStreamRicCreationEntity> list = new List<DataStreamRicCreationEntity>();
            try
            {
                if (listExtractedRecords == null || listExtractedRecords.Count == 0)
                    return null;

                foreach (var record in listExtractedRecords)
                {
                    DataStreamRicCreationEntity dsInfo = new DataStreamRicCreationEntity(record);

                    if (IsValidDataStreamRicCreationInfo(dsInfo))
                        list.Add(dsInfo);
                }

                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private bool IsValidDataStreamRicCreationInfo(DataStreamRicCreationEntity dsInfo)
        {
            try
            {
                bool result = true;

                if (dsInfo == null)
                    result = false;

                if (dsInfo.fileName.StartsWith("EM") && !dsInfo.ExchangeCode.Equals("SET"))
                    result = false;

                if (dsInfo.ThomsonReutersClassificationScheme.Trim().Equals("RTS"))
                    result = false;

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
        #endregion

        #region [extract records]
        private List<string> ExtractRecordsFromFiles(List<string> listDownloadFtpFilePath)
        {
            List<string> list = new List<string>();
            try
            {
                if (listDownloadFtpFilePath == null || listDownloadFtpFilePath.Count == 0)
                {
                    string msg = string.Format("no ftp file need to extract .");
                    Logger.Log(msg, Logger.LogType.Warning);
                    return null;
                }

                foreach (var item in listDownloadFtpFilePath)
                {
                    ExtractRecord(item, list);
                }
                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private void ExtractRecord(string item, List<string> list)
        {
            try
            {
                if (!File.Exists(item))
                    return;

                int start = item.LastIndexOf("\\");
                string fileName = item.Substring(start + 1, item.Length - start - 1);

                using (StreamReader sr = new StreamReader(item))
                {
                    string line = null;
                    string exchangeCode = string.Empty;

                    while ((line = sr.ReadLine().ToUpper()) != null)
                    {
                        if (!line.StartsWith("XE"))
                            continue;

                        if (line.Length < 1031)
                            continue;

                        list.Add(line + fileName);
                    }
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
        #endregion

        #region [download files]
        private List<string> DownloadFilesFromFtp(List<string> listDownloadFtpCode)
        {
            List<string> list = new List<string>();

            try
            {
                string fileName = string.Empty;
                string fileUrl = string.Empty;
                string targetFolder = Path.Combine(
                    Path.Combine(configObj.OutputPath, DateTime.Now.ToString("yyyy-MM-dd")),
                    "DSE_FILES");
                string targetPath = string.Empty;

                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                AddResult("DSE_FILES", targetFolder, "DSE_FILES");

                foreach (var item in listDownloadFtpCode)
                {
                    fileName = string.Format("{0}{1}.M", item, configObj.Date);
                    fileUrl = @"ftp://ASIA2:ASIA2@ds1.rds.reuters.com/" + fileName;
                    targetPath = Path.Combine(targetFolder, fileName);

                    if (!File.Exists(targetPath))
                        DownloadFileFromFtp(targetPath, fileUrl);

                    if (File.Exists(targetPath))
                        list.Add(targetPath);
                }

                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private void DownloadFileFromFtp(string targetPath, string fileUrl)
        {
            try
            {
                WebClient request = new WebClient();
                WebProxy proxy = new WebProxy(configObj.IP, Convert.ToInt32(configObj.Port));
                request.Proxy = proxy;
                request.Credentials = new NetworkCredential("ASIA2", "ASIA2");
                request.DownloadFile(fileUrl, targetPath);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                LogMessage(string.Format("Can not download file: {0}", targetPath));
            }
        }
        #endregion

        #region [check fonfig]
        private bool PrepearConfig()
        {
            try
            {
                bool result = true;

                if ((configObj.OutputPath + "").Trim().Length == 0)
                    return false;

                if (!Directory.Exists(configObj.OutputPath))
                    Directory.CreateDirectory(configObj.OutputPath);

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
        #endregion
    }

    public class ThailandTemplate
    {
        public string Mnem { get; set; }
        public string Sedol { get; set; }
        public string Name1 { get; set; }
        public string Name2 { get; set; }
        public string Bdate { get; set; }
        public string DefMkt { get; set; }
        public string Mnem2 { get; set; }
        public string Isin { get; set; }
        public string Cname { get; set; }
        public string Geog { get; set; }
        public string Dcur { get; set; }
        public string Remk { get; set; }
        public string Secty { get; set; }
        public string Grpi { get; set; }
        public string Fname1 { get; set; }
        public string Mic { get; set; }
        public string Fname2 { get; set; }
        public string Cusip { get; set; }
        public string Qci { get; set; }
        public string ParentRic { get; set; }
        public string NameSuffix { get; set; }

        public ThailandTemplate(DataStreamRicCreationEntity dsInfo)
        {
            //Mnem    
            this.Mnem = GetMnem(dsInfo);
            //Sedol 

            //Name1 

            //Name2 

            //Bdate 
            this.Bdate = GetBdate(dsInfo);

            //DefMkt 
            this.DefMkt = GetDefMkt(dsInfo);

            //Mnem2 
            this.Mnem2 = "            ";

            //Isin 
            if ((dsInfo.Isin + "").Trim().Length > 0)
                this.Isin = dsInfo.Isin;
            else
                this.Isin = "            ";

            //Cname 

            //Geog
            this.Geog = "016";

            //Dcur 
            this.Dcur = "031";

            //Remk 
            this.Remk = GetRemk(dsInfo);

            //Secty 
            //Grpi 
            if ((dsInfo.ThomsonReutersClassificationScheme + "").Trim().Equals("ETF"))
            {
                this.Secty = "ET";
                this.Grpi = "159";
            }
            else
            {
                this.Secty = "EQ";
                this.Grpi = "116";
            }

            //Fname1 

            //Mic
            this.Mic = "FB ";

            //Fname2

            //Cusip
            this.Cusip = ("TH:" + dsInfo.Ticker.Replace("-", "")).PadRight(12, ' ');

            //Qci 
            if (dsInfo.CompanyName.Contains("ADR"))
                this.Qci = "22";
            else if (dsInfo.CompanyName.Contains("GDR"))
                this.Qci = "D5";
            else
                this.Qci = "  ";

            //ParentRic
            if (dsInfo.RIC.Trim().Equals(dsInfo.ReutersEditorialRIC.Trim()))
                this.ParentRic = "";
            else
                this.ParentRic = dsInfo.ReutersEditorialRIC;

        }

        private string GetRemk(DataStreamRicCreationEntity dsInfo)
        {
            string result = "FBBF";
            string remk = (dsInfo.marketSegmentName + "").Trim();//field 53

            //*** If Field 53 = "SET" or "MAI", REMK ="MBBK"
            //***All others:  REMK = "FBBF"
            if (remk.Equals("SET") || remk.Equals("MAI"))
                result = "MBBK";

            return result;
        }

        private string GetDefMkt(DataStreamRicCreationEntity dsInfo)
        {
            string result = "BK";
            string defMkt = (dsInfo.marketSegmentName + "").Trim();//field 53

            //*** If Field 53 = "SET Foreign Board" or "MAI Foreign Board", DEF_MKT="BF"
            //***Others, REMK="BK"
            if (defMkt.Equals("SET Foreign Board") || defMkt.Equals("MAI Foreign Board"))
                result = "BF";

            return result;
        }

        private string GetBdate(DataStreamRicCreationEntity dsInfo)
        {
            string result = "        ";
            string firstTradeDate = (dsInfo.FirstTradingDate + "").Trim();//field 89

            try
            {
                if (firstTradeDate.Equals("-9999999"))
                    return result;

                DateTime dt;

                if (DateTime.TryParseExact(
                    firstTradeDate,
                    "yyyyMMdd",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AdjustToUniversal,
                    out dt))
                {
                    return dt.ToString("dd/MM/yy", DateTimeFormatInfo.InvariantInfo);
                }

                return result;
            }
            catch (Exception)
            {
                return result;
            }

        }

        private string GetMnem(DataStreamRicCreationEntity dsInfo)
        {
            string result = "      ";
            try
            {
                string marketSegmentName = (dsInfo.marketSegmentName + "").Trim();//field 53
                string thomsonReutersClassificationScheme = (dsInfo.ThomsonReutersClassificationScheme + "").Trim();//field 74
                string ticker = (dsInfo.Ticker + "").Trim();//field 45
                string ric = (dsInfo.RIC + "").Replace(" ", "").Replace("_", "").Replace(".", "").Replace("-", "");//field 2

                if (marketSegmentName.Equals("SET") || marketSegmentName.Equals("MAL"))
                {
                    //*** If Field 53 = "SET" or "MAI" & Field 74 ="ORD",  MNEM = "Q:"+First 3 characters of Field 45 + "P"  (i.e. Q:PCSP) 
                    //*** If Field 53 = "SET" or "MAI" & Field 74 ="NVDR", MNEM = "Q:"+First 3 characters of Field 45 + "R" (i.e. Q:PCSR)
                    if (thomsonReutersClassificationScheme.Equals("ORD"))
                        return "Q:" + ticker.Substring(0, 3) + "P";
                    else if (thomsonReutersClassificationScheme.Equals("NVDR"))
                        return "Q:" + ticker.Substring(0, 3) + "R";
                }

                //*** If Field 53 = "SET Foreign Board" or "MAI Foreign Board", MNEM = "Q:"+First 3 characters of Field 45 + "F" (i.e. Q:PCSF)
                if (marketSegmentName.Equals("SET Foreign Board") || marketSegmentName.Equals("MAI Foreign Board"))
                    return "Q:" + ticker.Substring(0, 3) + "F";

                //Field 74 ="PRF" or "CPR" or "PREFERRED"
                //dot of Field 2. Remove any "_" or space or  "." or  "-"  (i.e. Field 2 = BH_pn.BK -> Q:BHPN)
                //*Always 6 characters, add trailing spaces if value is too short
                if (thomsonReutersClassificationScheme.Equals("PRF") ||
                    thomsonReutersClassificationScheme.Equals("CPR") ||
                    thomsonReutersClassificationScheme.Equals("PREFERRED"))
                {
                    return "Q:" + ticker.Substring(0, 3) + ticker.Substring(ticker.Length - 1, 1);
                }

                return result;
            }
            catch (Exception)
            {
                return result;
            }
        }
    }
}
