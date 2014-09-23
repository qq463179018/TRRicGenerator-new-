using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using HtmlAgilityPack;
using System.IO;
using System.ComponentModel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
{
    [ConfigStoredInDB]   
    public class KoreaKtbFutureRolloverConfig
    {
        [StoreInDB]
        [DefaultValue("C:\\Korea_Auto\\KTB Future\\GEDA\\")]
        [Description("Path for saving generated GEDA files (KTB Future)\nE.g. C:\\Korea_Auto\\KTB Future\\GEDA\\ ")]
        public string GEDA { get; set; }
    }

    public class KoreaKtbFutureRollover : GeneratorBase
    {

        private List<KoreaKtbFutureInfo> futures = new List<KoreaKtbFutureInfo>();
        
        private KoreaKtbFutureRolloverConfig configObj = null;

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KoreaKtbFutureRolloverConfig;
            if (string.IsNullOrEmpty(configObj.GEDA))
            {
                configObj.GEDA = GetOutputFilePath();
            }
            else
            {
                configObj.GEDA = Path.Combine(configObj.GEDA, DateTime.Today.ToString("yyyy-MM-dd"));
            }
        }

        protected override void Start()
        {
            GetFutures();            
            GenerateFile();
        }

        private void GetFutures()
        {
            string urlThreeYears = "http://eng.krx.co.kr/m3/m3_3/m3_3_3/m3_3_3_1/UHPENG03003_03_01.html";
            string urlFiveYears = "http://eng.krx.co.kr/m3/m3_3/m3_3_3/m3_3_3_2/UHPENG03003_03_02.html";
            string urlTenYears = "http://eng.krx.co.kr/m3/m3_3/m3_3_3/m3_3_3_3/UHPENG03003_03_03.html";

            Dictionary<string, string[]> pageChain = new Dictionary<string, string[]>();
            pageChain.Add(urlThreeYears, new string[2] { "KOSBND_OTC_KTB1U_CHAIN", "KOSBND_OTC_KTB2U_CHAIN" });
            pageChain.Add(urlFiveYears, new string[2] { "KOSBND_OTC_5TB1U_CHAIN", "KOSBND_OTC_5TB2U_CHAIN" });
            pageChain.Add(urlTenYears, new string[2] { "KOSBND_OTC_10TB1U_CHAIN", "KOSBND_OTC_10TB2U_CHAIN" });

            foreach (var de in pageChain)
            {
                GetChainCode(de.Key, de.Value);
            }
        }

        private void GetChainCode(string pageUrl, string[] chains)
        {
            HtmlDocument doc = WebClientUtil.GetHtmlDocument(pageUrl, 180000);
            List<HtmlNode> codeTables = new List<HtmlNode>();
            HtmlNodeCollection tables = doc.DocumentNode.SelectNodes("//table");

            foreach (HtmlNode table in tables)
            {
                if (table.Attributes["summary"] == null)
                {
                    break;
                }
                
                string startTitle = table.SelectSingleNode(".//tr[1]/th[1]").InnerText.Trim();
               
                if (startTitle.Equals("Issue Name"))
                {
                    codeTables.Add(table);
                }
            }

            if (codeTables.Count != 2)
            {
                string msg = string.Format("Error! Can not get 2 tables for Underlying Basket Bonds from web page: {0}", pageUrl);
                throw new Exception(msg);
            }

            for (int i = 0; i < 2; i++)
            {
                GetBasketBondsDetail(codeTables[i], chains[i]);
            }
        }

        private void GetBasketBondsDetail(HtmlNode table, string chain)
        {
            HtmlNodeCollection trs = table.SelectNodes("tbody/tr");
            KoreaKtbFutureInfo chainCode = new KoreaKtbFutureInfo();
            chainCode.Chain = chain;
            for (int i = 0; i < trs.Count; i++)
            {                
                string codeNo = trs[i].SelectSingleNode("td[2]").InnerText.Trim();
                if (!string.IsNullOrEmpty(codeNo))
                {
                    codeNo = codeNo.Substring(0, codeNo.Length - 1) + "=KQ";

                    chainCode.CodeList.Add(codeNo);
                }
            }
            futures.Add(chainCode);
        }
      
        private void GenerateFile()
        {
            AddResult("GEDA Folder",configObj.GEDA,"GEDA Folder");
            foreach (KoreaKtbFutureInfo item in futures)
            { 
                string fileName = item.Chain.Replace("KOSBND_OTC_", "") + ".txt";
                string filePath = Path.Combine(configObj.GEDA, fileName);
                string content = item.Chain + "\r\n";
                content += string.Join("\r\n", item.CodeList.ToArray());
                FileUtil.WriteOutputFile(filePath, content);
                AddResult(fileName,filePath,fileName);
            }
        }
    }

    public class KoreaKtbFutureInfo
    {
        public string Chain { get; set; }
        public List<string> CodeList { get; set; }

        public KoreaKtbFutureInfo()
        {
            Chain = string.Empty;
            CodeList = new List<string>();
        }
    }
}
