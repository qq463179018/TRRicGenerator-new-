using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks
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

        private  List<KoreaKtbFutureInfo> futures = new List<KoreaKtbFutureInfo>();
        
        private KoreaKtbFutureRolloverConfig configObj;

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KoreaKtbFutureRolloverConfig;
            configObj.GEDA = string.IsNullOrEmpty(configObj.GEDA) ? GetOutputFilePath() : Path.Combine(configObj.GEDA, DateTime.Today.ToString("yyyy-MM-dd"));
        }

        protected override void Start()
        {
            GetFutures();            
            GenerateFile();
        }

        private void GetFutures()
        {
            const string urlThreeYears = "http://eng.krx.co.kr/m3/m3_3/m3_3_3/m3_3_3_1/UHPENG03003_03_01.html";
            const string urlFiveYears = "http://eng.krx.co.kr/m3/m3_3/m3_3_3/m3_3_3_2/UHPENG03003_03_02.html";
            const string urlTenYears = "http://eng.krx.co.kr/m3/m3_3/m3_3_3/m3_3_3_3/UHPENG03003_03_03.html";

            Dictionary<string, string[]> pageChain = new Dictionary<string, string[]>
            {
                {urlThreeYears, new[] {"KOSBND_OTC_KTB1U_CHAIN", "KOSBND_OTC_KTB2U_CHAIN"}},
                {urlFiveYears, new[] {"KOSBND_OTC_5TB1U_CHAIN", "KOSBND_OTC_5TB2U_CHAIN"}},
                {urlTenYears, new[] {"KOSBND_OTC_10TB1U_CHAIN", "KOSBND_OTC_10TB2U_CHAIN"}}
            };

            foreach (var de in pageChain)
            {
                GetChainCode(de.Key, de.Value);
            }
        }

        private void GetChainCode(string pageUrl, string[] chains)
        {
            HtmlDocument doc = WebClientUtil.GetHtmlDocument(pageUrl, 180000);
            HtmlNodeCollection tables = doc.DocumentNode.SelectNodes("//table");

            List<HtmlNode> codeTables = (from table in tables.TakeWhile(table => table.Attributes["summary"] != null) 
                                         let startTitle = table.SelectSingleNode(".//tr[1]/th[1]").InnerText.Trim() 
                                         where startTitle.Equals("Issue Name") 
                                         select table).ToList();

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
            KoreaKtbFutureInfo chainCode = new KoreaKtbFutureInfo {Chain = chain};
            foreach (string codeNo in 
                     from tr in trs 
                     select tr.SelectSingleNode("td[2]").InnerText.Trim() 
                     into codeNo 
                        where !string.IsNullOrEmpty(codeNo) 
                        select codeNo.Substring(0, codeNo.Length - 1) + "=KQ")
            {
                chainCode.CodeList.Add(codeNo);
            }
            futures.Add(chainCode);
        }
      
        private void GenerateFile()
        {
            TaskResultList.Add(new TaskResultEntry("GEDA Folder", "GEDA Folder", configObj.GEDA));
            foreach (KoreaKtbFutureInfo item in futures)
            { 
                string fileName = item.Chain.Replace("KOSBND_OTC_", "") + ".txt";
                string filePath = Path.Combine(configObj.GEDA, fileName);
                string content = item.Chain + "\r\n";
                content += string.Join("\r\n", item.CodeList.ToArray());
                FileUtil.WriteOutputFile(filePath, content);
                TaskResultList.Add(new TaskResultEntry(fileName, fileName, filePath, FileProcessType.GEDA_BULK_CHAIN_RIC_CREATION));
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
