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
    public class TwOrdDropConfig
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

        [StoreInDB]
        [DisplayName("Retire date")]
        [Description("Retire date of selected Rics")]
        public DateTime RetireDate { get; set; }

    }

    #endregion

    public class TwOrdDrop : GeneratorBase
    {
        #region Initialization

        private TwOrdDropConfig _configObj;
        private HFile _ordIdn;
        private HFile _ordNda;

        #endregion

        #region GeneratorBase Implementation

        protected override void Initialize()
        {
            base.Initialize();
            _configObj = Config as TwOrdDropConfig;
        }

        protected override void Start()
        {
            var datas = new List<Dictionary<string, string>>();
            var ordNda = new Nda(FileMode.WriteOnly);
            var ordIdn = new Idn(FileMode.WriteOnly);

            try
            {
                SetTemplates();
                datas.AddRange(_configObj.Rics.Select(GetData));

                ordIdn.LoadFromTemplate(_ordIdn, datas);
                ordIdn.Save(Path.Combine(_configObj.WorkingDir, String.Format("IdnAdd{0}.txt", _configObj.Market)));
                AddResult("Idn Add file", ordIdn.Path, "idn");

                ordNda.LoadFromTemplate(_ordNda, datas);
                ordNda.Save(Path.Combine(_configObj.WorkingDir, String.Format("NdaAdd{0}.csv", _configObj.Market)));
                AddResult("Nda Add file", ordNda.Path, "nda");
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

        private Dictionary<string, string> GetData(string ric)
        {
            return new Dictionary<string, string>
            {
                {"code", ric},
                {"market", _configObj.Market},
                {"retireDate", _configObj.RetireDate.ToString("dd-MMM-yy")},
            };
        }

        private void SetTemplates()
        {
            _ordIdn = TemplateIdn.TwOrdDrop;
            if (_configObj.Market.ToLower().Contains("twse"))
            {
                _ordNda = Template.TwOrdDropTwse;
            }
            else if (_configObj.Market.ToLower().Contains("gtsm"))
            {
                _ordNda = Template.TwOrdDropGtsm;
            }
            else
            {
                _ordNda = Template.TwOrdDropEmg;
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
