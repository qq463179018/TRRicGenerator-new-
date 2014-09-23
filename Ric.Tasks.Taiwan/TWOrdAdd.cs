using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using Ric.Core;
using Ric.FileLib;
using Ric.FormatLib;
using Ric.Util;
using FileMode = Ric.FileLib.Enum.FileMode;

namespace Ric.Tasks.Taiwan
{
    #region Config

    [ConfigStoredInDB]
    public class TwOrdAddConfig
    {
        [StoreInDB]
        [DisplayName("Working directory")]
        [Description("Directory where files will be written")]
        public string WorkingDir { get; set; }

        [StoreInDB]
        [Description("List of Rics")]
        public List<string> Rics { get; set; }

        [StoreInDB]
        [Description("Market : Choose between EMG,  TWSE or GTSM")]
        public string Market { get; set; }
    }

    #endregion

    public class TwOrdAdd : GeneratorBase
    {
        #region Initialization

        private TwOrdAddConfig _configObj;
        private HFile _ordIdn;
        private HFile _ordNda;
        private HFile _ordQuoteFuture;
        private HFile _ordIssueFuture;
        private VFile _ordFm;

        #endregion

        #region GeneratorBase Implementation

        protected override void Initialize()
        {
            base.Initialize();
            _configObj = Config as TwOrdAddConfig;
        }

        protected override void Start()
        {
            var datas = new Dictionary<string, string>();
            var cleanData = new List<Dictionary<string, string>>();
            var ordNda = new Nda(FileMode.WriteOnly);
            var ordQuoteFuture = new Nda(FileMode.WriteOnly);
            var ordIssueFuture = new Nda(FileMode.WriteOnly);
            var ordFm = new Fm();
            var ordIdn = new Idn(FileMode.WriteOnly);

            try
            {
                SetTemplates();
                foreach (var ric in _configObj.Rics)
                {
                    datas = GetInfos(ric.Trim());
                    cleanData.Add(CleanData(datas));
                }
                ordIssueFuture.LoadFromTemplate(_ordIssueFuture, cleanData);
                ordIssueFuture.Save(Path.Combine(_configObj.WorkingDir, String.Format("IssueFutureAdd{0}.csv", _configObj.Market)));
                AddResult("Issue Future Add file", ordIssueFuture.Path, "nda");


                ordIdn.LoadFromTemplate(_ordIdn, cleanData);
                ordIdn.Save(Path.Combine(_configObj.WorkingDir, String.Format("IdnAdd{0}.txt", _configObj.Market)));
                AddResult("Idn Add file", ordIdn.Path, "idn");

                ordNda.LoadFromTemplate(_ordNda, cleanData);
                ordNda.Save(Path.Combine(_configObj.WorkingDir, String.Format("NdaAdd{0}.csv", _configObj.Market)));
                AddResult("Nda Add file", ordNda.Path, "nda");

                ordQuoteFuture.LoadFromTemplate(_ordQuoteFuture, cleanData);
                ordQuoteFuture.Save(Path.Combine(_configObj.WorkingDir, String.Format("QuoteFutureAdd{0}.csv", _configObj.Market)));
                AddResult("Quote Future Add file", ordQuoteFuture.Path, "nda");

                ordFm.LoadFromTemplate(_ordFm, cleanData);
                ordFm.Save(Path.Combine(_configObj.WorkingDir, String.Format("Fm{0}.txt", _configObj.Market)));
                AddResult("Fm file", ordFm.Path, "fm");
            }
            catch (Exception ex)
            {
                LogMessage("Task failed, error: " + ex.Message, Logger.LogType.Error);
                throw new Exception("Task failed, error: " + ex.Message, ex);
            }
            finally
            {
                
            }
        }

        #endregion

        private Dictionary<string, string> CleanData(Dictionary<string, string> datas)
        {
            var cleanData = new Dictionary<string, string>
            {
                {"code", datas["股票代號"]},
                {"displayname", datas["英文全名"].Replace(" CO.,", "").Replace("LTD", "").Trim().ToUpper()},
                {"shortname", datas["英文簡稱"].ToUpper()},
                {"lotsize", "1000"},
                {"fullname", datas["英文全名"].Replace(" CO.,", "").Replace("LTD", "").Trim().ToUpper()},
                {"isin", GetIsin(datas["股票代號"])},
                {"chinesename", datas["chinesename"]},
                {"fullchinesename", datas["公司名稱"]}
            };
            if (_configObj.Market.Contains("TWSE"))
            {
                cleanData.Add("effectivedate", ConvertDate(datas["上市日期"]).ToString("dd-MMM-yy"));
                cleanData.Add("effectiveDateLong", ConvertDate(datas["上市日期"]).ToString("dd-MMM-yyyy"));
                cleanData.Add("firstTradingDay", ConvertDate(datas["上市日期"]).ToString("dd-MMM-yyyy"));
            }
            else if (_configObj.Market.Contains("GTSM"))
            {
                cleanData.Add("effectivedate", ConvertDate(datas["上櫃日期"]).ToString("dd-MMM-yy"));
                cleanData.Add("effectiveDateLong", ConvertDate(datas["上櫃日期"]).ToString("dd-MMM-yyyy"));
                cleanData.Add("firstTradingDay", ConvertDate(datas["上櫃日期"]).ToString("dd-MMM-yyyy"));
            }
            else if (_configObj.Market.Contains("EMG"))
            {
                cleanData.Add("effectivedate", ConvertDate(datas["興櫃日期"]).ToString("dd-MMM-yy"));
                cleanData.Add("effectiveDateLong", ConvertDate(datas["興櫃日期"]).ToString("dd-MMM-yyyy"));
                cleanData.Add("firstTradingDay", ConvertDate(datas["興櫃日期"]).ToString("dd-MMM-yyyy"));
            }
            return cleanData;
        }

        private string GetIsin(string code)
        {
            try
            {
                string url = String.Format("http://isin.twse.com.tw/isin/single_main.jsp?owncode={0}", code);
                var htc = WebClientUtil.GetHtmlDocument(url, 300000, null);

                return htc.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[2]/tr[2]/td[2]").InnerText.Trim();
            }
            catch (Exception)
            {
                throw new Exception("Isin did not came out on the website yet, try again later");
            }
        }

        private void SetTemplates()
        {
            _ordIssueFuture = Template.TwOrdIssueFuture;
            if (_configObj.Market.Contains("TWSE"))
            {
                _ordIdn = TemplateIdn.TwOrdIdnTwse;
                _ordNda = Template.TwOrdAddTwse;
                _ordQuoteFuture = Template.TwOrdQuoteFutureTwse;
                _ordFm = TemplateFm.TwOrdTwse;
            }
            else if (_configObj.Market.Contains("GTSM"))
            {
                _ordIdn = TemplateIdn.TwOrdIdnGtsm;
                _ordNda = Template.TwOrdAddGtsm;
                _ordQuoteFuture = Template.TwOrdQuoteFutureGtsm;
                _ordFm = TemplateFm.TwOrdGtsm;
            }
            else if (_configObj.Market.Contains("EMG"))
            {
                _ordIdn = TemplateIdn.TwOrdIdnEmg;
                _ordNda = Template.TwOrdAddEmg;
                _ordQuoteFuture = Template.TwOrdQuoteFutureEmg;
                _ordFm = TemplateFm.TwOrdEmg;
            }
            else
            {
                throw new Exception("The Market you choose does not exist, please change the configuration and choose an existing one.");
            }
        }


        #region Get Informations

        /// <summary>
        /// Converts Taiwan date to 'normal' date
        /// eg: 102/11/13   --> 13NOV13
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        private DateTime ConvertDate(string date)
        {
            if (date == "0")
                return DateTime.Now;
            string[] datePart = date.Split('/');
            return DateTime.ParseExact(String.Format("{0}/{1}/{2}", datePart[1], datePart[2], (Convert.ToInt32(datePart[0]) + 1911)), "MM/dd/yyyy", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// From a given ric, find in TWSE website informations about it
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        private Dictionary<string, string> GetInfos(string code)
        {
            const string url = "http://mops.twse.com.tw/mops/web/ajax_t05st03";
            bool skipped = false;
            bool skipped2 = false;
            var datas = new Dictionary<string, string>();
            string postData = String.Format("encodeURIComponent=1&step=1&firstin=1&off=1&keyword4={0}&code1=&TYPEK2=&checkbtn=1&queryName=co_id&TYPEK=all&co_id={0}"
                , code);
            var htc = WebClientUtil.GetHtmlDocument(url, 300000, postData);
            HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
            foreach (var tr in tables[1].SelectNodes(".//tr"))
            {
                var ths = tr.SelectNodes(".//th");
                var tds = tr.SelectNodes(".//td");
                int i = 0;
                foreach (var th in ths)
                {
                    if (i == 31 && skipped == false)
                    {
                        skipped = true;
                        continue;
                    }

                    if (i == 32 && skipped2 == false)
                    {
                        skipped2 = true;
                        continue;
                    }

                    if (!datas.ContainsKey(th.InnerText.Replace("&nbsp;", "").Replace("&nbsp", "").Trim()) && i < tds.Count && i != 29)
                    {
                        datas.Add(th.InnerText.Replace("&nbsp;", "").Replace("&nbsp", "").Trim(),
                            tds[i].InnerText.Replace("&nbsp;", "").Replace("&nbsp", "").Trim());
                    }
                    i++;
                }
            }
            datas.Add("chinesename", GetChineseName(htc));

            return datas;
        }

        private string GetChineseName(HtmlDocument htc)
        {
            var pNodes = htc.DocumentNode.SelectNodes("//span").Where(node => node.InnerText.Contains("公司)"));
            string chineseName = String.Empty;


            foreach (var htmlNode in pNodes)
            {
                Trace.TraceInformation(htmlNode.InnerText + "\r\n");
                chineseName = htmlNode.InnerText.Split('\n')[1];
            }
            return chineseName;
        }


        #endregion

    }
}
