using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Ric.Core;
using Ric.Util;
using System.Threading;

namespace Ric.Tasks.HongKong
{
    #region Model & Config

    [ConfigStoredInDB]
    public class HKTradingNewsConfig
    {
        [StoreInDB]
        [DisplayName("Link Name")]
        [Description("Search the Link name, like OPTN NEWS- NEW OPTION SERIES.")]
        public string LinkName { get; set; }

        [StoreInDB]
        [DisplayName("Save Path")]
        [Description("The path of the generated file to save.")]
        public string SavePath { get; set; }

        [StoreInDB]
        [DisplayName("Xml Config Path")]
        [Description("Path of the XML file")]
        public string XmlConfigPath { get; set; }
    }

    public class HKTradingNewsExpirDateAndExlNameConfig
    {
        public List<Expir> ExpirDateMapping { get; set; }
        public List<ExlName> ExlNameMapping { get; set; }
    }

    public class Expir
    {
        public string SourceDate { get; set; }
        public string MappingDate { get; set; }
    }

    public class ExlName
    {
        public string ClassName { get; set; }
        public string Code { get; set; }
    }

    #endregion

    #region Task

    public class HKTradingNews : GeneratorBase
    {
        #region Properties

        private const string XmlConfigPath = ".\\Config\\HK\\HK_TradingNews.xml";
        private static HKTradingNewsConfig _configObj;
        private static HKTradingNewsExpirDateAndExlNameConfig _xmlConfigObj;

        private const string beforeUrl = @"http://www.hkex.com.hk/eng/market/sec_tradinfo/tradnews/today/";
        //private const string beforeUrl = @"http://www.hkex.com.hk/eng/market/sec_tradinfo/tradnews/today/news.htm";
        //Below link used for test yesterday's news.
        //private string beforeUrl = @"http://www.hkex.com.hk/eng/market/sec_tradinfo/tradnews/prvtrad_day/";

        List<string> optLinks = new List<string>();

        private int _linkLength;
        private string _fullpath = "";

        #endregion

        #region GeneratorBase implementation

        protected override void Start()
        {
            StartTradingNewsJob();
        }

        protected override void Initialize()
        {
            base.Initialize();

            _configObj = Config as HKTradingNewsConfig;
            _xmlConfigObj = ConfigUtil.ReadConfig(_configObj.XmlConfigPath, typeof(HKTradingNewsExpirDateAndExlNameConfig)) as HKTradingNewsExpirDateAndExlNameConfig;
        }

        private void StartTradingNewsJob()
        {
            GetListInfo("news.htm");
        }

        #endregion

        #region Getting infos

        public void GetListInfo(string afterUrl)
        {
            var allItems = new List<List<string>>();
            optLinks.Clear();

            DateTime now = DateTime.Now;
            _fullpath = _configObj.SavePath + "\\opt" + now.Year + now.Month + now.Day + ".txt";
            _linkLength = _configObj.LinkName.Length;
            try
            {
                LogMessage("Fetching website");
                string url = beforeUrl + afterUrl;
                WebRequest wr = WebRequest.Create(url);
                Thread.Sleep(3000);
                WebResponse rs = wr.GetResponse();
                var sr = new StreamReader(rs.GetResponseStream());
                string htm = sr.ReadToEnd();

                LogMessage("Get all options position");
                GetAllOptPos(htm);

                if (optLinks.Count <= 0)
                {
                    MessageBox.Show("No new option published, please run this task later!", "Notes");
                    return;
                }

                LogMessage("Iterate through options pages");
                foreach (string optLink in optLinks)
                {
                    if (optLink.Contains(".htm"))
                    {
                        allItems.AddRange(GetOptionSeries(optLink));
                    }
                    else
                    {
                        LogMessage(optLink);
                    }
                }

                LogMessage("Starting write file");
                WriteOptText(allItems);
            }
            catch (Exception e)
            {
                LogMessage(e.Message);
            }
        }

        private void GetAllOptPos(string htm)
        {
            int pos = htm.IndexOf(_configObj.LinkName);
            if (pos > 0)
            {
                string beforeHtm = htm.Substring(0, pos);
                AddOptLinks(beforeHtm);
                string afterHtm = htm.Substring(pos + _linkLength, htm.Length - pos - _linkLength);
                GetAllOptPos(afterHtm);
            }

        }

        private void AddOptLinks(string htm)
        {
            int end = htm.LastIndexOf('"');
            int start = htm.Substring(0, end).LastIndexOf('"');
            string summaryUrl = htm.Substring(start + 1, end - start - 1);
            optLinks.Add(summaryUrl);
        }

        private IEnumerable<List<string>> GetOptionSeries(string summaryUrl)
        {
            var allItems = new List<List<string>>();
            string optUrl = beforeUrl + summaryUrl;
            WebRequest wr = WebRequest.Create(optUrl);
            WebResponse rs = wr.GetResponse();
            var sr = new StreamReader(rs.GetResponseStream());
            string htm = sr.ReadToEnd();

            int firstPos = htm.IndexOf("<PRE>");
            int lastPos = htm.IndexOf("</PRE>");

            string text = htm.Substring(firstPos + 5, lastPos - firstPos - 5);

            allItems = ExtractText(text);
            return allItems;

        }

        public List<List<string>> ExtractText(string txt)
        {
            var allItems = new List<List<string>>();
            if (txt != "")
            {
                txt = txt.Replace("\r", "").TrimStart('\n').TrimEnd('\n');
                List<string> optionSeries = txt.Split('\n').ToList();

                foreach (var splittedLine in optionSeries.Skip(7).Select(option => option.Replace("   ", ",").Replace(", ", ",").Split(',').ToList()))
                {
                    splittedLine.RemoveAll(String.IsNullOrEmpty);

                    if (splittedLine.Count > 1)
                    {
                        if (splittedLine[0].Contains("("))
                        {
                            splittedLine[0] = splittedLine[0].Remove(splittedLine[0].IndexOf("("));
                        }
                        allItems.Add(splittedLine);
                    }
                    else
                    {
                        if (allItems == null || allItems.Count == 0)
                            continue;

                        if (!allItems[allItems.Count - 1][2].EndsWith("/") && !splittedLine.First().Trim().StartsWith("/"))
                        {
                            allItems[allItems.Count - 1][2] += "/";
                        }
                        allItems[allItems.Count - 1][2] += splittedLine.First().Trim();
                    }
                }
            }
            return allItems;
        }

        private string GetMappingName(string source)
        {
            string code = "";
            foreach (ExlName item in _xmlConfigObj.ExlNameMapping.Where(item => source.Equals(item.ClassName)))
            {
                code = item.Code;
            }
            return code;
        }

        private string GetExpirDateMapping(string source)
        {
            string mappingDate = "";
            foreach (Expir item in _xmlConfigObj.ExpirDateMapping.Where(item => source.Equals(item.SourceDate)))
            {
                mappingDate = item.MappingDate;
            }
            return mappingDate;
        }

        #endregion

        #region Writing files

        private void WriteOptText(List<List<string>> allItems)
        {
            string exlName = "";
            string expirDate = "";
            var contents = new List<string>();
            var sbOpt = new StringBuilder();

            sbOpt.Append("EXL_NAME");
            sbOpt.Append("\t");
            sbOpt.Append("EXPIR_DATE");
            sbOpt.Append("\t");
            sbOpt.Append("STRIKE_PRICE");
            sbOpt.Append("\t");
            sbOpt.Append("INP_VER_NBR");
            sbOpt.Append("\t");
            sbOpt.Append("INP_COM_CODE");

            contents.Add(sbOpt.ToString());
            sbOpt.Remove(0, sbOpt.Length);

            int itemCount = 0;
            try
            {
                while (itemCount < allItems.Count)
                {
                    List<string> rowData = allItems[itemCount];
                    rowData[0] = rowData[0].Replace(" ($0.01)", "").Replace(" ($0.001)", "");
                    if (rowData[0].TrimStart(' ').IndexOf('$') == 0)
                    {
                        List<string> continuePrice = rowData[0].TrimStart(' ').TrimEnd(' ').Split('/').ToList();
                        foreach (string price in continuePrice)
                        {
                            sbOpt.Append(exlName);
                            sbOpt.Append("\t");
                            sbOpt.Append(expirDate);
                            sbOpt.Append("\t");
                            sbOpt.Append(price.Replace('$', ' ').TrimStart(' ').TrimEnd(' '));

                            contents.Add(sbOpt.ToString());
                            sbOpt.Remove(0, sbOpt.Length);
                        }
                    }
                    else
                    {
                        exlName = "SEHK_EQO_" + GetMappingName(rowData[0].TrimStart(' ').TrimEnd(' ').ToUpper());
                        expirDate = GetExpirDateMapping(rowData[1].TrimStart(' ').TrimEnd(' ').ToUpper());
                        sbOpt.Append(exlName);
                        sbOpt.Append("\t");
                        sbOpt.Append(expirDate);
                        sbOpt.Append("\t");
                        List<string> strikePrice = rowData.Count == 4 ? rowData[3].TrimStart(' ').TrimEnd(' ').Split('/').ToList() : rowData[2].TrimStart(' ').TrimEnd(' ').Split('/').ToList();
                        sbOpt.Append(strikePrice[0].Replace('$', ' ').TrimStart(' ').TrimEnd(' '));

                        contents.Add(sbOpt.ToString());
                        sbOpt.Remove(0, sbOpt.Length);

                        for (int colIndex = 1; colIndex < strikePrice.Count; colIndex++)
                        {
                            if (
                                !strikePrice[colIndex].Replace('$', ' ')
                                    .TrimStart(' ')
                                    .TrimEnd(' ')
                                    .Equals(string.Empty))
                            {
                                sbOpt.Append(exlName);
                                sbOpt.Append("\t");
                                sbOpt.Append(expirDate);
                                sbOpt.Append("\t");
                                sbOpt.Append(strikePrice[colIndex].Replace('$', ' ').TrimStart(' ').TrimEnd(' '));

                                contents.Add(sbOpt.ToString());
                                sbOpt.Remove(0, sbOpt.Length);
                            }
                        }
                    }
                    itemCount++;
                }
            }
            catch (Exception ex)
            {
                string errInfo = ex.ToString();
            }

            string[] newContent = CopyContent(contents.ToArray());

            if (!Directory.Exists(_configObj.SavePath))
            {
                CoreObj.CreateDir(_configObj.SavePath);
            }
            WriteTxtFile(_fullpath, newContent);
            AddResult("TradingNewsResult", _fullpath, "file");
        }

        private string[] CopyContent(IEnumerable<string> content)
        {
            var item = new List<string>();
            foreach (string str in content)
            {
                string temp = str.Replace("\r", "").Replace("\n", "").Trim();
                temp = temp.Substring(temp.IndexOf('\t'));
                if (!temp.Equals(string.Empty))
                {
                    item.Add(str);
                }
                else
                {
                    break;
                }
            }
            string[] newContent = item.ToArray();
            return newContent;
        }

        private void WriteTxtFile(string fullpath, string[] content)
        {
            try
            {
                File.WriteAllLines(fullpath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LogMessage("write to file error:" + ex.Message);
            }
        }

        #endregion
    }

    #endregion
}
