using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Ric.Core;
using Ric.Util;
using Finance.Lib;
using FileSys = System.IO;
using System.IO;

namespace Ric.Tasks.Taiwan
{
    #region Config

    [ConfigStoredInDB]
    public class TWCBAnnouncementConfig
    {
        [Description("Directory where FM will be written")]
        [StoreInDB]
        public string WorkingDir { get; set; }

        [Description("List of Rics")]
        [StoreInDB]
        public List<string> Rics { get; set; }

        [Category("Mail")]
        [Description("Recipients to send email to.")]
        [StoreInDB]
        public List<string> TO_TYPE_RECIPIENTS { get; set; }

        [Category("Mail")]
        [Description("Recipients to send email to in CC.")]
        [StoreInDB]
        public List<string> CC_TYPE_RECIPIENTS { get; set; }
    }

    public class FmInfo
    {
        public string Ric { get; set; }
        public string Isin { get; set; }
        public DateTime EffectiveDate { get; set; }
        public DateTime MatureDate { get; set; }
        public string ChineseName { get; set; }
        public string EnglishName { get; set; }
        public string Strikes { get; set; }
        public string Shares { get; set; }
        public string Type { get; set; }
    }

    #endregion

    public class TWCBAnnouncement : GeneratorBase
    {
        #region Initialization

        private static TWCBAnnouncementConfig configObj = null;
        private Dictionary<string, string> dplyNames = new Dictionary<string, string>();
        private List<string> cbAnnouncementFileList = new List<string>();
        List<string> toGrab = new List<string>() { "發行日期", "到期日期", "國際編碼", "債券代碼", "債券簡稱", "債券英文名稱", "最新轉(交)換價格", "實際發行總額" };

        #endregion

        #region GeneratorBase Implementation

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as TWCBAnnouncementConfig;
        }

        protected override void Start()
        {
            List<string> rics = configObj.Rics;
            try
            {
                dplyNames = QueryGats(PrepareRicsForGats(rics));
                var bulkFile = new Nda();
                foreach (string ric in rics)
                {
                    var infosRic = GenerateFm(ric);
                    if (infosRic != null)
                    {
                        bulkFile.AddProp(infosRic);
                    }
                }
                if (bulkFile.format.Prop.Length > 0)
                {
                    //remove the "\r\n" in the last line in the file end
                    bulkFile.GenerateAndSave("TwIdnCb", String.Format("{0}TwIdnCb_{1}.txt", configObj.WorkingDir, DateTime.Now.ToString("ddMM")));
                    //bulkFile.GenerateAndSave("TwCbBulk", String.Format("{0}TwCbBulk_{1}.csv", configObj.WorkingDir, DateTime.Now.ToString("ddMM")));
                    RemoveEndLineStr(String.Format("{0}TwIdnCb_{1}.txt", configObj.WorkingDir, DateTime.Now.ToString("ddMM")));
                    //TaskResultList.Add(new TaskResultEntry("bulk file", "", String.Format("{0}TwIdnCb_{1}.txt", configObj.WorkingDir, DateTime.Now.ToString("ddMM"))));
                    //TaskResultList.Add(new TaskResultEntry("bulk file", "", String.Format("{0}TwCbBulk_{1}.csv", configObj.WorkingDir, DateTime.Now.ToString("ddMM"))));
                    AddResult("bulk file1", String.Format("{0}TwIdnCb_{1}.txt", configObj.WorkingDir, DateTime.Now.ToString("ddMM")), "");
                    //AddResult("bulk file2", String.Format("{0}TwCbBulk_{1}.csv", configObj.WorkingDir, DateTime.Now.ToString("ddMM")), "");
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Task failed, error: " + ex.Message, Logger.LogType.Error);
                throw new Exception("Task failed, error: " + ex.Message, ex);
            }
        }

        private void RemoveEndLineStr(string fileName)
        {
            StringBuilder fileStr = new StringBuilder();

            if ((fileName + "").Trim().Length == 0)
                return;

            try
            {
                if (!FileSys.File.Exists(fileName))
                    return;

                fileStr.Append(FileSys.File.ReadAllText(fileName));
                if (fileStr.ToString().EndsWith("\t\r\n"))
                    fileStr.Length = fileStr.Length - 3;

                FileSys.File.WriteAllText(fileName, fileStr.ToString());
            }
            catch (Exception ex)
            {
                LogMessage(string.Format("remove \t\r\n error ,in the file {0},msg:{1}", fileName, ex.Message));
            }
        }

        #endregion

        #region Get Informations

        /// <summary>
        /// From informations create a FmInfo object
        /// </summary>
        /// <param name="results"></param>
        /// <returns></returns>
        private FmInfo InfosToFm(List<string> results)
        {
            FmInfo fm = new FmInfo();
            Dictionary<string, string> infos = new Dictionary<string, string>();
            char[] stringSeparators = new char[] { '：' };

            foreach (string entry in results)
            {
                string[] parts = entry.Split(stringSeparators, StringSplitOptions.None);
                if (parts.Length == 2)
                {
                    infos.Add(parts[0].Trim(), parts[1].Trim());
                }
            }

            fm.Isin = infos["國際編碼"];
            fm.MatureDate = ConvertDate(infos["到期日期"]);
            fm.EffectiveDate = ConvertDate(infos["發行日期"]);
            fm.ChineseName = infos["債券簡稱"];
            fm.Strikes = infos["最新轉(交)換價格"];
            fm.Shares = infos["實際發行總額"];
            fm.Ric = infos["債券代碼"];
            string[] nameParts = infos["債券英文名稱"].Split(fm.Ric.Substring(4).ToCharArray());
            fm.EnglishName = nameParts[0];
            string url = "http://mops.twse.com.tw/mops/web/ajax_t05st03";
            string postData = String.Format("encodeURIComponent=1&step=1&firstin=1&off=1&keyword4={0}&code1=&TYPEK2=&checkbtn=1&queryName=co_id&TYPEK=all&co_id={1}", fm.Ric.Substring(0, 4), fm.Ric.Substring(0, 4));
            var htc = WebClientUtil.GetHtmlDocument(url, 300000, postData);

            string fullHtml = htc.DocumentNode.InnerText;
            if (fullHtml.Contains("(上市公司)"))
            {
                fm.Type = ".TW";
            }
            else
            {
                fm.Type = ".TWO";
            }
            return fm;
        }

        /// <summary>
        /// Converts Taiwan date to 'normal' date
        /// eg: 102/11/13   --> 13NOV13
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        private DateTime ConvertDate(string date)
        {
            string[] datePart = date.Split('/');
            return DateTime.ParseExact(String.Format("{0}/{1}/{2}", datePart[1], datePart[2], (Convert.ToInt32(datePart[0]) + 1911).ToString()), "MM/dd/yyyy", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// From a given ric, find in TWSE website informations about it
        /// </summary>
        /// <param name="ric"></param>
        /// <returns></returns>
        private FmInfo GetInfos(string ric)
        {
            string url = "http://mops.twse.com.tw/mops/web/t120sg01";
            Dictionary<string, string> chineseNumbers = new Dictionary<string, string>
            {
                {"0", "零"},
                {"1", "一"},
                {"2", "二"},
                {"3", "三"},
                {"4", "四"},
                {"5", "五"},
                {"6", "六"},
                {"7", "七"},
                {"8", "八"}, 
                {"9", "九"}
            };
            for (DateTime dateToCheck = DateTime.Now.AddMonths(1); dateToCheck.Year > 2005; dateToCheck = dateToCheck.AddMonths(-1))
            {
                try
                {
                    string getData = String.Format("?encodeURIComponent=1&firstin=ture&tg=&pg=&step=0&TYPEK=&monyr_reg={3}{4}&issuer_stock_code={0}&bond_kind=5&bond_yrn={1}&bond_subn=$M00000001&bond_id={2}&come=2"
                        , ric.Substring(0, 4), ric.Substring(4), ric, dateToCheck.Year, dateToCheck.ToString("MM"));
                    var htc = WebClientUtil.GetHtmlDocument(url + getData, 300000, null);
                    List<string> results = new List<string>();

                    HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                    foreach (HtmlNode td in tables[3].SelectNodes(".//td"))
                    {
                        foreach (string keys in toGrab)
                        {
                            if (td.InnerText.Contains(keys))
                            {
                                results.Add(td.InnerText.Trim());
                            }
                        }
                    }
                    return InfosToFm(results);
                }
                catch
                { }
            }
            for (DateTime dateToCheck = DateTime.Now.AddMonths(1); dateToCheck.Year > 2005; dateToCheck = dateToCheck.AddMonths(-1))
            {
                try
                {
                    string getData = String.Format("?encodeURIComponent=1&firstin=ture&tg=&pg=&step=0&TYPEK=&monyr_reg={3}{4}&issuer_stock_code={0}&bond_kind=5&bond_yrn={1}&bond_subn=1&bond_id={2}&come=2"
                        , ric.Substring(0, 4), ric.Substring(4), ric, dateToCheck.Year, dateToCheck.ToString("MM"));
                    var htc = WebClientUtil.GetHtmlDocument(url + getData, 300000, null);
                    List<string> results = new List<string>();

                    HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                    foreach (HtmlNode td in tables[3].SelectNodes(".//td"))
                    {
                        foreach (string keys in toGrab)
                        {
                            if (td.InnerText.Contains(keys))
                            {
                                results.Add(td.InnerText.Trim());
                            }
                        }
                    }

                    if (results == null || results.Count == 0)
                        LogMessage("There is no valid infomation in http://mops.twse.com.tw/mops/web/t120sg01");

                    return InfosToFm(results);
                }
                catch
                { }
            }
            throw new Exception("Cannot find FM");
        }

        #endregion

        #region Gats

        /// <summary>
        /// Query GATS to find DSPLY_NAME fids from rics
        /// </summary>
        /// <param name="rics"></param>
        /// <returns></returns>
        private Dictionary<string, string> QueryGats(string rics)
        {
            GatsUtil gats = new GatsUtil(GatsUtil.Server.Idn);
            Dictionary<string, string> gatsValues = new Dictionary<string, string>();
            Regex rgxSpace = new Regex(@"\s+");
            Regex rgxPre = new Regex(@"^DSPLY_NAME", RegexOptions.IgnoreCase);
            try
            {
                string[] stringSeparators = new string[] { "\r\n" };
                char[] stringSeparators2 = new char[] { ' ' };

                string test = gats.GetGatsResponse(rics, "DSPLY_NAME");

                string[] lines = test.Split(stringSeparators, StringSplitOptions.None);
                foreach (string line in lines)
                {
                    string formattedLine = rgxSpace.Replace(line, " ");
                    string[] lineTab = formattedLine.Split(stringSeparators2);
                    if (lineTab.Length > 2
                        && rgxPre.IsMatch(lineTab[1])
                        && lineTab[2] != "")
                    {
                        if (!gatsValues.ContainsKey(lineTab[0].Trim()))
                        {
                            string tmpName = "";
                            for (int i = 0; i < lineTab.Count(); i++)
                            {
                                if (i >= 2)
                                {
                                    tmpName += lineTab[i] + " ";
                                }
                            }
                            gatsValues.Add(lineTab[0].Trim(), tmpName.Trim());
                        }
                    }
                }
                return gatsValues;
            }
            catch (Exception ex)
            {
                Logger.Log("Error in QueryGats", Logger.LogType.Error);
                throw new Exception("Error While using Gats: " + ex.Message);
            }
        }

        /// <summary>
        /// From the list string of rics generates a string with
        /// full rics to give to GATS
        /// </summary>
        /// <param name="ricsToSearch"></param>
        /// <returns></returns>
        private string PrepareRicsForGats(List<string> ricsToSearch)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string ric in ricsToSearch)
            {
                sb.AppendFormat("{0}.TW,{0}.TWO,", ric.Substring(0, 4));
            }
            return sb.ToString();
        }

        /// <summary>
        /// Find in gats results the displayname from given ric
        /// </summary>
        /// <param name="ric"></param>
        /// <param name="type">.TW or .TWO</param>
        /// <returns></returns>
        private string GetGatsDisplayName(string ric, string type)
        {
            string displayName = dplyNames[ric.Substring(0, 4) + type];
            if (displayName.Length >= 12)
            {
                return displayName.Substring(0, 12);
            }
            return displayName;
        }

        #endregion

        #region Generate FM

        /// <summary>
        /// Generates Fm from website informations
        /// </summary>
        /// <param name="ric"></param>
        private Dictionary<string, string> GenerateFm(string ric)
        {
            try
            {
                FmInfo res = GetInfos(ric);
                var fm = new Fm();
                var infos = new Dictionary<string, string>
                    {
                        {"ric", res.Ric},
                        {"name", res.EnglishName},
                        {"displayname", GetGatsDisplayName(res.Ric, res.Type)},
                        {"chinesename", res.ChineseName},
                        {"units", String.Format("{0:n0}", Convert.ToInt64(res.Shares.Replace("元", "").Replace(",", "")) / 100)},
                        {"maturedate", res.MatureDate.ToString("ddMMMyy").ToUpper()},
                        {"effectivedate", res.EffectiveDate.ToString("ddMMMyy").ToUpper()},
                        {"effectivedateidn", res.EffectiveDate.ToString("dd/MM/yyyy")},
                        {"isin", res.Isin},
                        {"abbrev", res.EnglishName.ToUpper()},
                        {"strike", res.Strikes.Replace("元", "")},
                        {"type", res.Type}
                    };
                fm.AddProp(infos);
                string filename = String.Format("{0}{1}_{2}.txt", configObj.WorkingDir, res.Ric, DateTime.Now.ToString("ddMM"));
                fm.GenerateAndSave("TwTemplate", filename);
                //TaskResultList.Add(new TaskResultEntry("result file" + res.Ric + " FM", "", filename));
                AddResult("result file", filename, "");
                //add "=== End of Proforma ===" in the file end
                AddWordInTheEnd(filename);

                //BCU.txt
                string fileNameBCU = Path.Combine(configObj.WorkingDir, "BCU.txt");
                GeneratorBCUFile(res, fileNameBCU);
                return infos;
            }
            catch (Exception ex)
            {
                Logger.Log("Fm generation failed for this ric, error: " + ex.Message, Logger.LogType.Warning);
                return null;
            }
        }

        private void GeneratorBCUFile(FmInfo res, string fileName)
        {
            string content = string.Empty;
            string bcuType = string.Empty;
            if (res.Type.Equals(".TWO"))
                bcuType = string.Format("OTCTWS_EQ_{0}_REL", res.Ric);
            else if (res.Type.Equals(".TW"))
                bcuType = string.Format("TAIW_EQ_{0}_REL", res.Ric);
            else
                LogMessage(string.Format("res.Type:{0} is not .TWO or .TW, so bcuType==empty.", res.Type));

            content = "Ric\tBCU\r\n" + res.Ric + "\t" + bcuType;
            System.IO.File.WriteAllText(fileName, content);
            AddResult("result file", fileName, "");
        }

        private void AddWordInTheEnd(string filename)
        {
            string fileStr = string.Empty;

            if ((filename + "").Trim().Length == 0)
                return;

            try
            {
                if (!FileSys.File.Exists(filename))
                    return;

                fileStr = FileSys.File.ReadAllText(filename);
                if (fileStr.EndsWith("\r\n"))
                    fileStr = fileStr + "=== End of Proforma ===";

                FileSys.File.WriteAllText(filename, fileStr);
            }
            catch (Exception ex)
            {
                LogMessage(string.Format("add the word in the end of the file{0} error.msg{1}", filename, ex.Message));
            }
        }

        #endregion

        #region Send Email

        //Send mail with CB Announcement text file attached.
        private void sendMail()
        {
            string errMsg = string.Empty;
            List<string> fileList = new List<string>();
            string mailSuj = "TW CB PRICE " + DateTime.Now.ToString("ddMMMyy").ToUpper();
            string mailBody = @"Hi, 
 
Grateful your help to insert attached  file for real-time data correction asap.  
Thank you in advance. 
 
Regards";
            if (fileList.Count != 0)
            {
                using (OutlookApp app = new OutlookApp())
                {
                    OutlookUtil.CreateAndSendMail(app, cbAnnouncementFileList, mailSuj, configObj.TO_TYPE_RECIPIENTS, configObj.CC_TYPE_RECIPIENTS, mailBody, out errMsg);
                }
            }
        }

        #endregion
    }
}
