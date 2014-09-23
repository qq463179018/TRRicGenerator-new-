using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class ETFinsertTaskConfig
    {
        [StoreInDB]
        [DisplayName("Result file directory")]
        public string ResultFileDir { get; set; }

        [StoreInDB]
        [DisplayName("Value in file five")]
        public string ValueInFileFive { get; set; }

        [StoreInDB]
        [DisplayName("Email 1 recipients")]
        [Description("The mail format should contain the full name, such as \"xxx.xxx@thomsonreuters.com\".")]
        public List<string> TwEmailOne_MailToSend { get; set; }

        [StoreInDB]
        [DisplayName("Email 1 recipients (CC)")]
        [Description("The mail format should contain the full name, such as \"xxx.xxx@thomsonreuters.com\".")]
        public List<string> TwEmailOne_MailToSend_CC { get; set; }

        [StoreInDB]
        [DisplayName("Email 1 signature")]
        [Description("The mail format should contain the full name, such as \"xxx.xxx@thomsonreuters.com\".")]
        public List<string> TwEmailOne_MailToSend_Signature { get; set; }

        [StoreInDB]
        [DisplayName("Email 2 recipients (CC)")]
        [Description("The mail format should contain the full name, such as \"xxx.xxx@thomsonreuters.com\".")]
        public List<string> TwEmailTwo_MailToSend_CC { get; set; }

        [StoreInDB]
        [DisplayName("Email 2 signature")]
        [Description("The mail format should contain the full name, such as \"xxx.xxx@thomsonreuters.com\".")]
        public List<string> TwEmailTwo_MailToSend_Signature { get; set; }

        [StoreInDB]
        [DisplayName("Email 2 recipients")]
        [Description("The mail format should contain the full name, such as \"xxx.xxx@thomsonreuters.com\".")]
        public List<string> TwEmailTwo_MailToSend { get; set; }
    }

    class ETFinsertTask : GeneratorBase
    {
        private class ETFinsertTaskTemplate
        {
            public ETFinsertTaskTemplate()
            {
                GrabTime = DateTime.Now;
            }
            public string Code { set; get; }
            public string ReferenceValue { set; get; }
            public string EstimatedValue { set; get; }
            public DateTime GrabTime { set; get; }
        }

        private const string HolidayListFilePath = ".\\Config\\TW\\TW_ Holiday.xml";
        private const string TableName = "ETI_TW_ETFinsert";
        private ETFinsertTaskConfig _configObj;
        private List<DateTime> _holidayList;

        #region initialize and start 
        protected override void Initialize()
        {
            base.Initialize();
            _configObj = Config as ETFinsertTaskConfig;
            _holidayList = ConfigUtil.ReadConfig(HolidayListFilePath, typeof(List<DateTime>)) as List<DateTime>;
        }
        protected override void Start()
        {
            try
            {
                List<ETFinsertTaskTemplate> res = GrabDataFromWeb();
                storeInDb(res);
                GenerateFiles(res);
                GenerateFileFive();
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
                Logger.Log(ex.StackTrace);
                throw ex;
            }
        }
        #endregion

        #region grab data
        private IEnumerable<ETFinsertTaskTemplate> GrabDataFromSiteOne()
        {
            HtmlDocument root = WebClientUtil.GetHtmlDocument("http://www.p-shares.com/fundindex_i_now.asp", 6000);
            HtmlNodeCollection trs = root.DocumentNode.SelectNodes("//tr");
            return (from tr in trs
                    select tr.SelectNodes("./td")
                    into tds
                    let regex = new Regex("\\d{1,}")
                    let match = regex.Match(tds[0].InnerText.TrimStart().TrimEnd())
                    where match.Success
                    select new ETFinsertTaskTemplate
                    {
                        Code = tds[0].InnerText.TrimStart().TrimEnd(), ReferenceValue = tds[2].InnerText.TrimStart().TrimEnd(), EstimatedValue = tds[3].InnerText.TrimStart().TrimEnd()
                    }).ToList();
        }

        private IEnumerable<ETFinsertTaskTemplate> GrabDataFromSiteTwo()
        {
            HtmlDocument root = WebClientUtil.GetHtmlDocument("http://www.assetmanagement.hsbc.com.tw/HSBC_WCTS_FE/feEvents/ETFSite/etf_message.asp?ETFOK=Y", 6000);
            HtmlNodeCollection trs = root.DocumentNode.SelectNodes("//table[@class=\"data\"]/tr");
            return (from tr in trs
                    select tr.SelectNodes("./td")
                    into tds
                    where tds != null && tds.Count >= 4
                    let regex = new Regex("\\d{1,}.\\d{1,}")
                    let match = regex.Match(tds[2].InnerText.TrimStart().TrimEnd())
                    where match.Success
                    select new ETFinsertTaskTemplate
                    {
                        Code = tds[0].InnerText.TrimStart().TrimEnd(), ReferenceValue = match.Value.TrimStart().TrimEnd(), EstimatedValue = tds[3].InnerText.TrimStart().TrimEnd()
                    }).ToList();
        }

        private IEnumerable<ETFinsertTaskTemplate> GrabDataFromSiteThree()
        {
            HtmlDocument root = WebClientUtil.GetHtmlDocument("http://etrade.fsit.com.tw/ETF/etf/realtime_fund_nav_content.aspx", 6000);
            HtmlNodeCollection trs = root.DocumentNode.SelectNodes("//table/tr");
            return (from tr in trs
                    select tr.SelectNodes("./td")
                    into tds
                    where tds != null && tds.Count >= 4
                    let regex = new Regex("\\d{1,}.\\d{1,}")
                    let match = regex.Match(tds[2].InnerText.TrimStart().TrimEnd())
                    where match.Success
                    select new ETFinsertTaskTemplate
                    {
                        Code = tds[0].InnerText.TrimStart().TrimEnd(), ReferenceValue = match.Value.TrimStart().TrimEnd(), EstimatedValue = tds[3].InnerText.TrimStart().TrimEnd()
                    }).ToList();
        }

        private IEnumerable<ETFinsertTaskTemplate> GrabDataFromSiteFour()
        {
            List<ETFinsertTaskTemplate> partOne = new List<ETFinsertTaskTemplate>();
            HtmlDocument root = WebClientUtil.GetHtmlDocument("http://sitc.sinopac.com/web/etf/Ajax/GetTradeinfo_nav.aspx", 6000, "fundcode=37");
            HtmlNodeCollection trs = root.DocumentNode.SelectNodes("//table/tbody/tr");
            foreach (var tr in trs)
            {
                HtmlNodeCollection tds = tr.SelectNodes("./td");
                if (tds == null || tds.Count < 4)
                    continue;
                Regex regex = new Regex("\\d{1,}.\\d{1,}");
                Match match = regex.Match(tds[2].InnerText.TrimStart().TrimEnd());
                if (!match.Success)
                {
                    continue;
                }
                ETFinsertTaskTemplate oneRes = new ETFinsertTaskTemplate
                {
                    ReferenceValue = match.Value.TrimStart().TrimEnd()
                };
                match = regex.Match(tds[0].InnerText.TrimStart().TrimEnd());
                oneRes.Code = match.Value.TrimStart().TrimEnd();
                match = regex.Match(tds[3].InnerText.TrimStart().TrimEnd());
                oneRes.EstimatedValue = match.Value.TrimStart().TrimEnd();
                partOne.Add(oneRes);
            }
            return partOne;
        }

        private IEnumerable<ETFinsertTaskTemplate> GrabDataFromSiteFive()
        {
            List<ETFinsertTaskTemplate> partOne = new List<ETFinsertTaskTemplate>();
            HtmlDocument root = WebClientUtil.GetHtmlDocument("https://www.kgifund.com.tw/50_05.asp", 6000);
            HtmlNodeCollection trs = root.DocumentNode.SelectNodes("/html[1]/body[1]/table[1]/tr[1]/td[2]/table[1]/tr[3]/td[2]/table[1]/tr[2]/td[1]/table[1]/tr[2]/td[2]/div[1]/table[1]/tr[3]/td[1]/table[1]/tr[2]/td[1]/table[1]/tr");
            HtmlNode codeNode = root.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[1]/tr[1]/td[2]/table[1]/tr[3]/td[2]/table[1]/tr[2]/td[1]/table[1]/tr[2]/td[2]/div[1]/table[1]/tr[3]/td[1]/table[1]/tr[1]/td[1]/table[1]/tr");
            ETFinsertTaskTemplate oneRes = new ETFinsertTaskTemplate();
            Regex regex = new Regex("\\d{1,}");
            Match match = regex.Match(codeNode.InnerText.TrimStart().TrimEnd());
            oneRes.Code = match.Value;
            foreach (var tr in trs)
            {
                if (tr.InnerText.Contains("昨日收盤淨值(台幣)"))
                {
                    regex = new Regex("\\d{1,}.\\d{1,}");
                    match = regex.Match(tr.InnerText.TrimStart().TrimEnd());
                    oneRes.ReferenceValue = match.Value;
                }
                if (tr.InnerText.Contains("盤中預估淨值(台幣)"))
                {
                    regex = new Regex("\\d{1,}.\\d{1,}");
                    match = regex.Match(tr.InnerText.TrimStart().TrimEnd());
                    oneRes.EstimatedValue = match.Value;
                }

            }

            partOne.Add(oneRes);
            return partOne;
        }

        private IEnumerable<ETFinsertTaskTemplate> GrabDataFromSiteSix()
        {
            List<ETFinsertTaskTemplate> partOne = new List<ETFinsertTaskTemplate>();
            HtmlDocument root = WebClientUtil.GetHtmlDocument("http://www.fhtrust.com.tw/funds/fund_ETF_RTnav.asp", 6000, "QueryFund=ETF01&AgreeFlag=Y");
            HtmlNodeCollection trs = root.DocumentNode.SelectNodes("//table[@class=\"tb_2\"][2]/tr");

            HtmlNodeCollection tds = trs[2].SelectNodes("./td");
            Regex regex = new Regex("\\d{1,}.\\d{1,}");
            Match match = regex.Match(tds[1].InnerText.TrimStart().TrimEnd());
            ETFinsertTaskTemplate oneRes = new ETFinsertTaskTemplate
            {
                ReferenceValue = match.Value.TrimStart().TrimEnd()
            };

            match = regex.Match(tds[0].InnerText.TrimStart().TrimEnd());
            oneRes.Code = match.Value.TrimStart().TrimEnd();

            match = regex.Match(tds[2].InnerText.TrimStart().TrimEnd());
            oneRes.EstimatedValue = match.Value.TrimStart().TrimEnd();
            partOne.Add(oneRes);

            return partOne;
        }

        private List<ETFinsertTaskTemplate> GrabDataFromWeb()
        {
            List<ETFinsertTaskTemplate> res = new List<ETFinsertTaskTemplate>();
            res.AddRange(GrabDataFromSiteOne());
            res.AddRange(GrabDataFromSiteTwo());
            res.AddRange(GrabDataFromSiteThree());
            res.AddRange(GrabDataFromSiteFour());
            res.AddRange(GrabDataFromSiteFive());
            res.AddRange(GrabDataFromSiteSix());
            return res;
        }
        #endregion

        #region Db operate

        private void storeInDb(IEnumerable<ETFinsertTaskTemplate> res)
        {
            DataTable table = ManagerBase.Select("ETI_TW_ETFinsert", new[] { "*" }, string.Format("where GrabTime='{0}'", DateTime.Now.ToString("yyyy-MM-dd", new CultureInfo("en-US"))));
            if (table.Rows.Count != 0)
                return;
            foreach (var term in res)
            {
                
                DataRow row = table.Rows.Add();
                row[0] = term.Code;
                row[1] = term.ReferenceValue;
                row[2] = term.EstimatedValue;
                row[3] = term.GrabTime.ToString("yyyy-MM-dd",new CultureInfo("en-US"));
            }
            ManagerBase.UpdateDbTable(table, "ETI_TW_ETFinsert");
        }

        #endregion

        #region compare and generate result

        private void GenerateFiles(List<ETFinsertTaskTemplate> res)
        {
            DateTime lastTradingDay = MiscUtil.GetLastTradingDay(DateTime.Now, _holidayList, 1);
            StringBuilder sb = new StringBuilder();
            string fileContents="";
            sb.Append("TQSK.DAT;21-MAR-2012 15:00:00;TPS;\r\n");
            sb.Append("GEN_VAL6;GV6_TEXT;\r\n");
            foreach (var term in res)
            {
                sb.AppendFormat("{0}.TW;{1};144700;\r\n", term.Code, term.EstimatedValue);
            }
            sb.Append("\r\n\r\n\r\n\r\n\r\n");
            fileContents+=sb.ToString();
            sb.Remove(0,sb.Length);
            //file one

            sb.Append("TAM.DAT;21-MAR-2012 15:00:00;TPS;\r\n");
            sb.Append("GEN_VAL7;\r\n");
            foreach (var term in from term in res 
                                 let @where = string.Format("where Code='{0}' and GrabTime='{1}' and EstimatedValue='{2}'", term.Code, lastTradingDay.ToString("yyyy/MM/dd", new CultureInfo("en-US")), term.EstimatedValue) 
                                 let table = ManagerBase.Select(TableName, new[] { "*" }, @where) 
                                 where table.Rows.Count == 0 
                                 select term)
            {
                sb.AppendFormat("{0}.TW;{1};\r\n", term.Code, term.ReferenceValue);
            }
            sb.Append("\r\n\r\n\r\n\r\n\r\n");
            fileContents+=sb.ToString();
            string file2 = sb.ToString();
            sb.Remove(0,sb.Length);
            //file two

            sb.Append("TAM.DAT;21-MAR-2012 15:00:00;TPS;\r\n");
            sb.Append("SPARE_DATE5;\r\n");
            foreach (var term in res)
            {
                sb.AppendFormat("{0}.TW;{1};\r\n", term.Code, DateTime.Now.ToString("dd/MM/yyyy"));
            }
            sb.Append("\r\n\r\n\r\n\r\n\r\n");
            fileContents+=sb.ToString();
            sb.Remove(0,sb.Length);
            //file three

            sb.Append("TAM.DAT;21-MAR-2012 15:00:00;TPS;\r\n");
            sb.Append("SPARE_DATE5;\r\n");

            sb.AppendFormat("006201.TWO;{0};\r\n", lastTradingDay.ToString("dd/MM/yyyy"));
            sb.AppendFormat("006202.TWO;{0};\r\n", lastTradingDay.ToString("dd/MM/yyyy"));
            sb.Append("\r\n\r\n\r\n\r\n\r\n");
            fileContents+=sb.ToString();
            sb.Remove(0,sb.Length);
            //file four

            string path = string.Format("{0}\\TQSK RICs.txt", _configObj.ResultFileDir);
            File.WriteAllText(path, fileContents, Encoding.UTF8);
            TaskResultList.Add(new TaskResultEntry(Path.GetFileNameWithoutExtension(path), path, path, CreateMailOne(file2)));
        }
        private void GenerateFileFive()
        {
            StringBuilder sb = new StringBuilder();
            string fileContents = "";
            sb.Append("TQSK.DAT;21-MAR-2012 15:00:00;TPS;\r\n");
            sb.Append("HIGH_1;LOW_1;\r\n");
            sb.AppendFormat(".TWOIRR;{0};{0};\r\n",_configObj.ValueInFileFive);
            fileContents += sb.ToString();
            sb.Remove(0, sb.Length);
            //file five
            string path = string.Format("{0}\\TWOIRR.txt", _configObj.ResultFileDir);
            File.WriteAllText(path, fileContents, Encoding.UTF8);
            TaskResultList.Add(new TaskResultEntry(Path.GetFileNameWithoutExtension(path), path, path, CreateMailTwo()));
        }
        #endregion

        #region create email

        private MailToSend CreateMailOne(string file2)
        {
            MailToSend mail = new MailToSend();
            StringBuilder mailbodyBuilder = new StringBuilder();
            string path = string.Format("{0}\\TQSK RICs.txt", _configObj.ResultFileDir);
            string subject = string.Format("TQRIC insert file for Taiwan real-time data ETF NAV wef {0}",DateTime.Now.ToString("dd-MMM-yyyy",new CultureInfo("en-US")));
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("Hi Global Real-Time Service Desk  & IDN DR support , \r\n");
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("Please help to input attached 4  TQRIC insert files for Taiwan ETF NAV data in TQSK . Thanks in advance.   \r\n");
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("cc to Timeseries,  \r\n");
            mailbodyBuilder.Append("Please help amend previous day’s ETF NAV in database.\r\n");
            mailbodyBuilder.Append("ETF NAV  20 MAY 2013\r\n");
            mailbodyBuilder.Append(file2);
            mailbodyBuilder.Append("\r\n");
            foreach (var sig in _configObj.TwEmailOne_MailToSend_Signature)
            {
                mailbodyBuilder.Append(sig + "\r\n");
            }
            mail.ToReceiverList.AddRange(_configObj.TwEmailOne_MailToSend);
            mail.CCReceiverList.AddRange(_configObj.TwEmailOne_MailToSend_CC);
            mail.AttachFileList.Add(path);
            mail.MailSubject = subject;
            mail.MailBody = mailbodyBuilder.ToString();
            return mail;
        }

        private MailToSend CreateMailTwo()
        {
            MailToSend mail = new MailToSend();
            string path = string.Format("{0}\\TWOIRR.txt", _configObj.ResultFileDir);
            string subject = string.Format("TQRIC insert file for Taiwan total return index <.TWOIRR> wef {0}", DateTime.Now.ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
            StringBuilder mailbodyBuilder = new StringBuilder();
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("Hi Global Real-Time Service Desk  & IDN DR support , \r\n");
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("Please help to insert attached TQRIC insert files for Taiwan total return index <.TWOIRR> . Thanks in advance.\r\n");
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("\r\n");
            foreach (var sig in _configObj.TwEmailTwo_MailToSend_Signature)
            {
                mailbodyBuilder.Append(sig + "\r\n");
            }
            mail.ToReceiverList.AddRange(_configObj.TwEmailTwo_MailToSend);
            mail.CCReceiverList.AddRange(_configObj.TwEmailTwo_MailToSend_CC);
            mail.AttachFileList.Add(path);
            mail.MailSubject = subject;
            mail.MailBody = mailbodyBuilder.ToString();
            return mail;
        }
        #endregion
    }
}
