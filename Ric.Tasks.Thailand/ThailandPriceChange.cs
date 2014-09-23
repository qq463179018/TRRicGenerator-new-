using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using HtmlAgilityPack;
using Ric.Core;
using Ric.FileLib;
using Ric.FormatLib;
using Ric.Util;
using FileMode = Ric.FileLib.Enum.FileMode;

namespace Ric.Tasks.Thailand
{
    #region Configuration

    [ConfigStoredInDB]
    public class ThailandPriceChangeConfig
    {
        [StoreInDB]
        [DisplayName("Result folder")]
        [Description("The path where the result will be written.\nEg: C:/Mydrive/")]
        public string ResultFolderPath { get; set; }
    }

    #endregion

    #region Task

    class ThailandPriceChange : GeneratorBase
    {
        #region Initialization

        private static ThailandPriceChangeConfig configObj;

        #endregion

        #region GeneratorBase Implementation

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as ThailandPriceChangeConfig;
        }

        protected override void Start()
        {
            try
            {
                int i = 1;
                var propsPc = new List<Dictionary<string, string>>();
                foreach (Dictionary<string, string> infos in GetPages().Select(page => CleanInfos(GetInfos(page)))
                                                                       .Where(infos => infos != null)
                                                                       .Where(infos => infos["subject"] == "Adjustment"))
                {
                    propsPc.Add(infos);
                    i++;
                }
                if (i > 1)
                {
                    string path = string.Format("{0}Price_Change_{1}.xlsx", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM"));
                    List<string> title = new List<string>() { "RIC", "Symbol", "ISIN", "OLD_RATIO", "#INSTMOD_WNT_RATIO", "OLD_PRC", "#INSTMOD_STRIKE_PRC" };
                    List<List<string>> bulkFile = new List<List<string>>();
                    bulkFile.Add(title);
                    FillInBulkFile(bulkFile, propsPc);
                    XlsOrCsvUtil.GenerateXls0rCsv(path, bulkFile);
                    AddResult(Path.GetFileNameWithoutExtension(path), path, "");

                    //var ndaDw1 = new Nda(FileMode.WriteOnly);
                    //ndaDw1.LoadFromTemplate(TemplateFm.ThPriceChange, propsPc);
                    //ndaDw1.Save(String.Format("{0}Price_Change_{1}.xlsx", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                    //AddResult("Price change Nda bulk file", ndaDw1.Path, "nda");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Task failed, error: " + ex.Message, ex);
            }
        }

        private void FillInBulkFile(List<List<string>> bulkFile, List<Dictionary<string, string>> propsPc)
        {
            try
            {
                List<string> line = null;
                foreach (var dic in propsPc)
                {
                    line = new List<string>();

                    line.Add(dic["ric"]);
                    line.Add(dic["symbol"]);
                    line.Add(string.Empty);
                    line.Add(dic["oldratio"]);
                    line.Add(dic["newratio"]);
                    line.Add(dic["oldprice"]);
                    line.Add(dic["newprice"]);

                    bulkFile.Add(line);
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

        private Dictionary<string, string> CleanInfos(Dictionary<string, string> oldInfos)
        {
            try
            {
                return (new Dictionary<string, string>
                {
                    {"ric", GetRic(oldInfos["Symbol"])},
                    {"subject", oldInfos["Subject"]},
                    {"symbol", oldInfos["Symbol"]},
                    {"newprice", oldInfos["After Exercise Price"]},
                    {"oldprice", oldInfos["Before Exercise Price"]},
                    {"oldratio", oldInfos["Before Adjustment Exercise Ratio"]},
                    {"newratio", oldInfos["After Adjustment Exercise Ratio"]}
                });
            }
            catch
            {
                return null;
            }
        }

        private string GetRic(string symbol)
        {
            return (symbol.Length > 8) ? symbol.Remove(symbol.Length - 5, 1) + ".BK" : (symbol + ".BK");
        }

        private Dictionary<string, string> GetInfos(string page)
        {
            try
            {
                var request = WebRequest.Create(page) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
                request.Method = "GET";
                request.KeepAlive = true;
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
                request.Credentials = CredentialCache.DefaultCredentials;
                request.Headers["Cache-Control"] = "max-age=0";
                request.Headers["Accept-Encoding"] = "gzip,deflate,sdch";
                request.Headers["Accept-Language"] = "en-US,en;q=0.8";
                var httpResponse = (HttpWebResponse)request.GetResponse();

                Stream pageStream = httpResponse.GetResponseStream();
                var html = new HtmlDocument();
                html.Load(pageStream);

                HtmlNode pre = html.DocumentNode.SelectSingleNode(".//pre");
                string[] lineSeparators = { "\n" };
                string[] pointSeparators = { ":" };

                string[] lines = pre.InnerText.Split(lineSeparators, StringSplitOptions.None);
                return lines.Select(line => line.Replace(" (DW :", "")
                                                .Replace(" (THB/share)", "")
                                                .Replace(" (Update)", "")
                                                .Split(pointSeparators, StringSplitOptions.None))
                                                .Where(lineParts => lineParts.Length >= 2)
                                                .ToDictionary(lineParts => lineParts[0].Trim(), lineParts => lineParts[1].Trim());
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
            return null;
        }

        private IEnumerable<string> GetPages()
        {
            ServicePointManager.Expect100Continue = false;
            var results = new List<string>();

            int toAdd = DateTime.Now.DayOfWeek == DayOfWeek.Monday ? -3 : -1;
            //if (DateTime.Now.DayOfWeek == DayOfWeek.Monday)
            //{
            //    toAdd = -3;
            //}
            //else
            //{
            //    toAdd = -1;
            //}

            foreach (string url in Enumerable.Range(0, 10).Select(i => string.Format("http://www.set.or.th/set/newslist.do?to={0}%2F{1}%2F{2}&headline=exercise+price&submit=Search&symbol=&currentpage={3}&from={0}%2F{1}%2F{2}&newsType=&country=US&exchangeSymbols=&company=true&exchangeNews=on&language=en&exchange=true"
                , DateTime.Now.AddDays(toAdd).ToString("dd"), DateTime.Now.AddDays(toAdd).ToString("MM"), DateTime.Now.Year, i)))
            {
                try
                {
                    var htc = WebClientUtil.GetHtmlDocument(url, 300000);

                    HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                    HtmlNodeCollection links = tables[15].SelectNodes(".//a");
                    results.AddRange(from link in links
                                     select link.Attributes["href"].Value.Trim() into rawlink
                                     select rawlink.Substring(rawlink.IndexOf("?")) into rawlink
                                     select rawlink.Replace("&amp;", "&") into rawlink
                                     select "http://www.set.or.th/set/newsdetails.do" + rawlink);
                }
                catch
                {
                    break;
                }
            }
            return results;
        }
    }

    #endregion
}
