using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.IO;

namespace Ric.Tasks.China
{
    [ConfigStoredInDB]
    class CnIpoAutomationConfig
    {
        [StoreInDB]
        [Category("InputFilePath")]
        [Description("path of source file,like")]
        public string OutputPath { get; set; }

        [StoreInDB]
        [Category("OutputFolder")]
        [DefaultValue("yaqiong.wang")]
        [Description("UserName")]
        public string AssignedTo { get; set; }
    }

    class CnIpoAutomation : GeneratorBase
    {
        private CnIpoAutomationConfig configObj = null;
        protected override void Initialize()
        {
            configObj = Config as CnIpoAutomationConfig;
        }

        protected override void Start()
        {
            List<string> urlBase = new List<string>() { 
                @"http://www.sge.sh/publish/sge/xqzx/myxq/index.htm",
                @"http://www.sge.sh/publish/sge/xqzx/myxq/index1.htm",
                @"http://www.sge.sh/publish/sge/xqzx/myxq/index2.htm",
                @"http://www.sge.sh/publish/sge/xqzx/myxq/index3.htm",
                @"http://www.sge.sh/publish/sge/xqzx/myxq/index4.htm",
                @"http://www.sge.sh/publish/sge/xqzx/myxq/index5.htm",
                @"http://www.sge.sh/publish/sge/xqzx/myxq/index6.htm",
            };
            List<string> urlAll = GetAllUrl(urlBase);
            List<ExcelEntity> bulkFile = ExtractEntityValue(urlAll);
            List<List<string>> bulkFile2 = ConvertToList(bulkFile);
            Ric.Util.XlsOrCsvUtil.GenerateStringCsv(Path.Combine(configObj.OutputPath, "jinlan.csv"), bulkFile2);
        }

        private List<List<string>> ConvertToList(List<ExcelEntity> bulkFile)
        {
            List<List<string>> result = new List<List<string>>();
            List<string> title = new List<string>() {
            "year",
            "day",
            "month",
            "key",
            "value"
            };
            result.Add(title);
            foreach (var item in bulkFile)
            {
                List<string> line = new List<string>();
                line.Add(item.Year);
                line.Add(item.Month);
                line.Add(item.Day.Replace("\r", "").Replace("\n", "").Trim());
                line.Add(item.Key);
                line.Add(item.Value.Replace("\r", "").Replace("\n", "").Trim());

                result.Add(line);
            }

            return result;
        }

        private List<ExcelEntity> ExtractEntityValue(List<string> urlAll)
        {
            List<ExcelEntity> result = new List<ExcelEntity>();

            try
            {
                HtmlDocument htc = null;
                string year = string.Empty;
                string month = string.Empty;
                string day = string.Empty;
                string key = "Ag(T+D)";
                string value = string.Empty;
                foreach (var item in urlAll)
                {
                    Ric.Util.RetryUtil.Retry(5, TimeSpan.FromSeconds(5), true, delegate
                    {
                        htc = Ric.Util.WebClientUtil.GetHtmlDocument(item, 3000);
                    });

                    try
                    {
                        string title = htc.DocumentNode.SelectNodes(".//h1")[0].InnerText;
                        Match ma = (new Regex(@"(?<year>\d{2,4})年(?<month>\d{1,2})月")).Match(title);
                        if (ma.Success)
                        {
                            year = ma.Groups["year"].Value;
                            month = ma.Groups["month"].Value;
                        }

                        HtmlNodeCollection trs = htc.DocumentNode.SelectNodes(".//tbody")[0].SelectNodes(".//tr");
                        for (int i = 1; i < trs.Count; i++)
                        {
                            int countTitle = trs[0].SelectNodes(".//td").Count;
                            HtmlNodeCollection tds = trs[i].SelectNodes(".//td");
                            if (tds.Count == countTitle)
                                day = tds[0].InnerText;

                            if (tds[0].InnerText.Contains(key))
                            {
                                value = tds[9].InnerText;

                                ExcelEntity entity = new ExcelEntity();
                                entity.Year = year;
                                entity.Month = month;
                                entity.Day = day;
                                entity.Key = key;
                                entity.Value = value;

                                result.Add(entity);
                            }

                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            return null;
        }

        private List<string> GetAllUrl(List<string> baseUrl)
        {
            List<string> result = new List<string>();

            try
            {
                HtmlDocument htc = null;
                foreach (var item in baseUrl)
                {
                    Ric.Util.RetryUtil.Retry(5, TimeSpan.FromSeconds(5), true, delegate
                    {
                        htc = Ric.Util.WebClientUtil.GetHtmlDocument(item, 3000);
                    });

                    try
                    {
                        HtmlNodeCollection dls = htc.DocumentNode.SelectNodes(".//dl");
                        string innerText = string.Empty;
                        for (int i = 0; i < dls.Count; i++)
                        {
                            innerText = dls[i].SelectSingleNode(".//a").Attributes["href"].Value.Trim();
                            if ((innerText + "").Trim().Length == 0)
                            {
                                System.Windows.Forms.MessageBox.Show("error extract html");
                                continue;
                            }

                            result.Add(string.Format(@"http://www.sge.sh/publish/sge/xqzx/myxq/{0}", innerText));
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

            return null;
        }
    }

    struct ExcelEntity
    {
        public string Year { get; set; }
        public string Month { get; set; }
        public string Day { get; set; }
        public string Key { get; set; }
        public string Value { get; set; }
    }
}
