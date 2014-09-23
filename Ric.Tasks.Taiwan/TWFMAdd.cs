using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks.Taiwan
{
    [ConfigStoredInDB]
    public class TWFMAddConfig
    {
        [StoreInDB]
        [DisplayName("Working path")]
        public string WorkingPath { get; set; }

        public string OneTimeRun { get; set; }

        public TWFMAddConfig()
        {
            OneTimeRun = "0";
        }
    }

    public class TWFMAdd : GeneratorBase
    {
        private TWFMAddConfig configObj = null;
        string twUrl = "http://mops.twse.com.tw/mops/web/ajax_quickpgm?GOOD_ID=&TYPEK=all&TYPEK2=sii&encodeURIComponent=1&firstin=ture&keyword4=&off=1&queryName=GOOD_ID&step=2";
        string twoUrl = "http://mops.twse.com.tw/mops/web/ajax_quickpgm?GOOD_ID=&TYPEK=all&TYPEK2=otc&encodeURIComponent=1&firstin=ture&keyword4=&off=1&queryName=GOOD_ID&step=2";
        string kgiUrl = "https://derivatives.kgi.com.tw/EDWebSite/EDWeb/Warrant/WarrantIssue2.aspx?PageID=1042";
        string isinUrl = "http://brk.twse.com.tw:8000/isin/single_main.jsp";
        string warrantDetailUrl = "http://mops.twse.com.tw/mops/web/ajax_t05st48";

        string kgiFileFullName = null;
        string fmFileFullName = null;
        string lastWarrantListFileFullName = null;
        string missedWarrantFileFullName = null;

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as TWFMAddConfig;

            kgiFileFullName = Path.Combine(configObj.WorkingPath, "temp.xls");
            fmFileFullName = Path.Combine(configObj.WorkingPath, "TW_Warrant_ADD_FM.xls");//ric lose
            lastWarrantListFileFullName = Path.Combine(configObj.WorkingPath, "last warrant list.xml");
            missedWarrantFileFullName = Path.Combine(configObj.WorkingPath, "MISSED_WARRANTS.xls");

            LogMessage("Initialization done");
        }

        protected override void Start()
        {
            StartFMBulkFileGenertor();
        }

        /// <summary>
        /// Step 1: Load yesterday's missed data and download data from two websites
        /// Step 2: Get the detail warrant infomation
        /// Step 3: Update the issue date and price from KGI 
        /// Step 4: Generating FM files
        /// </summary>
        private void StartFMBulkFileGenertor()
        {
            List<TWWarrant> newWarrantList = null;

            if (configObj.OneTimeRun.Trim().Equals("1"))
            {
                newWarrantList = new List<TWWarrant>();
            }

            else
            {
                // 1. Get today's list
                List<TWWarrantBaseInfo> newList = GetTwoWarrantList();

                List<TWWarrantBaseInfo> existingWarrants = new List<TWWarrantBaseInfo>();

                if (File.Exists(lastWarrantListFileFullName))
                {
                    existingWarrants =
                    ConfigUtil.ReadConfig(lastWarrantListFileFullName, typeof(List<TWWarrantBaseInfo>)) as List<TWWarrantBaseInfo>;
                }

                BackUpWarrantsToFile(newList, lastWarrantListFileFullName);

                // 2. Get added list
                newWarrantList = GetNewAddedsWarrants(newList, existingWarrants);

            }
            // 3. Get last missing list
            List<TWWarrant> missedWarrants = ReadMissedData();

            // 4. The whole added list
            if (missedWarrants != null && missedWarrants.Count != 0)
            {
                newWarrantList.AddRange(missedWarrants);
            }

            if (newWarrantList == null || newWarrantList.Count == 0)
            {
                Logger.Log("There is no new-added code today!");
                return;
            }

            // 5. Get Warrant info from Exchange Website but IssueDate, Price, ISIN
            newWarrantList = UpdateWarrants(newWarrantList);

            // 6. Get Warrant IssueDate, Price info from KGI Website
            UpdateIssueDateAndRationBasedKGI(newWarrantList);

            // 7. Generate FM and get ISIN from Website
            GenerateFMFile(newWarrantList);
        }

        #region Step1 Load yesterday's missed data and download data from two websites

        /// <summary>
        /// Read yesterday's missed data. If exists, add them into the data need to run today.
        /// </summary>
        /// <returns></returns>
        private List<TWWarrant> ReadMissedData()
        {
            string filePath = Path.GetFullPath(missedWarrantFileFullName);
            if (!File.Exists(filePath))
            {
                return null;
            }

            List<TWWarrant> missedWarrants = new List<TWWarrant>();

            try
            {
                ExcelApp app = new ExcelApp(false, false);

                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                if (worksheet != null)
                {
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                    for (int i = 2; i <= lastUsedRow; i++)
                    {
                        if (ExcelUtil.GetRange(i, 1, worksheet).Text.ToString().Trim() != "")
                        {
                            TWWarrant warrant = new TWWarrant();
                            warrant.WarrantCode = ExcelUtil.GetRange(i, 1, worksheet).Text.ToString().Trim();
                            warrant.WarrantNameAbb = ExcelUtil.GetRange(i, 2, worksheet).Text.ToString().Trim();
                            warrant.Type = ExcelUtil.GetRange(i, 3, worksheet).Text.ToString().Trim();
                            if (warrant.WarrantCode.Contains("."))
                            {
                                string[] ric = warrant.WarrantCode.Split('.');
                                warrant.WarrantCode = ric[0];
                                warrant.Type = ric[1];
                            }
                            missedWarrants.Add(warrant);
                        }
                    }

                    if (lastUsedRow > 1)
                    {
                        try
                        {
                            File.Copy(missedWarrantFileFullName,
                                Path.Combine(GetPreviousFolder(), GetBackupFileName(missedWarrantFileFullName)),
                                true);

                            Range delRow = ExcelUtil.GetRange("A2:D" + lastUsedRow, worksheet);
                            delRow.Clear();
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("Failed to backup missed warrant file, skip: " + ex.Message, Logger.LogType.Warning);
                        }
                    }
                    workbook.Save();
                    workbook.Close();
                }
            }
            catch (Exception ex1)
            {
                Logger.Log("Failed to read missed warrant file, skip: " + ex1.Message, Logger.LogType.Warning);
            }
            return missedWarrants;
        }

        string GetBackupFileName(string filename)
        {
            return Path.GetFileNameWithoutExtension(filename) + DateTime.Now.ToString("yyyyMMdd") + Path.GetExtension(filename);
        }

        string GetPreviousFolder()
        {
            string path = Path.Combine(configObj.WorkingPath, "Previous");

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            return path;
        }

        private List<TWWarrantBaseInfo> GetTwoWarrantList()
        {
            List<TWWarrantBaseInfo> warrantList1 = null;
            List<TWWarrantBaseInfo> warrantList2 = null;

            int retriesLeft = 2;
            while ((warrantList1 == null || warrantList1.Count == 0) && retriesLeft-- > 0)
            {
                try
                {
                    warrantList1 = GetWarrantList(twUrl, "TW");
                }
                catch (Exception)
                {
                    Logger.Log("Error happens when download TW data from, try again");
                }
            }
            if (warrantList1 == null)
            {
                Logger.Log(string.Format("Cannot download page {0}", twUrl));
                throw new Exception();
            }
            int retriesLeftTwo = 2;
            while ((warrantList2 == null || warrantList2.Count == 0) && retriesLeftTwo-- > 0)
            {
                try
                {
                    warrantList2 = GetWarrantList(twoUrl, "TWO");
                }
                catch (Exception)
                {
                    Logger.Log("Error happens when download TWO data from, try again");
                }
            }
            if (warrantList2 == null)
            {
                Logger.Log(string.Format("Cannot download page {0}", twoUrl));
                throw new Exception();
            }
            warrantList1.AddRange(warrantList2);
            return warrantList1;
        }

        /// <summary>
        /// Get the table's content into TWWarrantBaseInfo, screen the valid nodes(code.length=6, all digits/end with P/B/C) 
        /// </summary>
        /// <param name="url">original data's url</param>
        /// <param name="warrantType">TW or TWO</param>
        /// <returns></returns>
        private List<TWWarrantBaseInfo> GetWarrantList(string url, string warrantType)
        {
            List<TWWarrantBaseInfo> warrantList = new List<TWWarrantBaseInfo>();
            //HtmlAgilityPack.HtmlDocument htmlDoc = WebClientUtil.GetHtmlDocument(url, 180000);
            HtmlDocument htmlDoc = null;
            RetryUtil.Retry(5, TimeSpan.FromSeconds(2), true, delegate
            {
                htmlDoc = WebClientUtil.GetHtmlDocument(url, 180000);
            });

            var nodeList = htmlDoc.DocumentNode.SelectNodes("//table/tr/td/a");
            for (int i = 0; i < nodeList.Count; i += 2)
            {
                string warrantCode = MiscUtil.GetCleanTextFromHtml(nodeList[i].InnerText).Trim();
                string warrantNameAbb = MiscUtil.GetCleanTextFromHtml(nodeList[i + 1].InnerText).Trim();
                string codeEnd = warrantCode.Substring(warrantCode.Length - 1);
                bool isValid = false;
                if (Regex.IsMatch(codeEnd, @"^\d+$")
                    || codeEnd.ToUpper().Equals("P")
                    || codeEnd.ToUpper().Equals("B")
                    || codeEnd.ToUpper().Equals("C")
                    || codeEnd.ToUpper().Equals("X")
                    || codeEnd.ToUpper().Equals("Y"))
                {
                    isValid = true;
                }
                if (warrantCode.Length == 6 && isValid && warrantNameAbb.Length != 0)
                {
                    TWWarrantBaseInfo warrant = new TWWarrantBaseInfo();
                    warrant.WarrantCode = warrantCode;
                    warrant.WarrantNameAbb = warrantNameAbb;
                    warrant.Type = warrantType;
                    warrantList.Add(warrant);
                }
            }
            return warrantList;
        }

        /// <summary>
        /// backup the download data to .\Previous\TW+WARRANT+ADD+FM_Previous_MMM_dd_HH_mm.xml
        /// </summary>
        /// <param name="warrantInfoList">list to backup</param>
        /// <param name="backUpFilePath">download data's full path</param>
        private void BackUpWarrantsToFile(List<TWWarrantBaseInfo> warrantInfoList, string backUpFilePath)
        {
            string previousFilePath = Path.GetDirectoryName(backUpFilePath) + "\\Previous";
            if (!Directory.Exists(previousFilePath))
            {
                Directory.CreateDirectory(previousFilePath);
            }
            if (File.Exists(backUpFilePath))
            {
                File.Copy(backUpFilePath, Path.Combine(previousFilePath,
                    string.Format("{0}_Previous_{1}.xml", Path.GetFileNameWithoutExtension(backUpFilePath), DateTime.Now.ToString("MMM_dd_HH_mm"))));
            }
            ConfigUtil.WriteXml(backUpFilePath, warrantInfoList);
        }

        /// <summary>
        /// Get today's new added warrants and code.
        /// </summary>
        /// <param name="allNewWarrants">today's all data</param>
        /// <param name="existingWarrants">yesterday's all data</param>
        /// <returns></returns>
        private List<TWWarrant> GetNewAddedsWarrants(List<TWWarrantBaseInfo> allNewWarrants, List<TWWarrantBaseInfo> existingWarrants)
        {
            List<TWWarrant> newWarrantList = new List<TWWarrant>();
            Hashtable existingWarrantsTab = new Hashtable();
            foreach (TWWarrantBaseInfo warrant in existingWarrants)
            {
                if (!existingWarrantsTab.Contains(warrant.WarrantCode))
                {
                    existingWarrantsTab.Add(warrant.WarrantCode, warrant);
                }
            }
            foreach (TWWarrantBaseInfo warrant in allNewWarrants)
            {
                if (!existingWarrantsTab.ContainsKey(warrant.WarrantCode))
                {
                    TWWarrant twWarrant = new TWWarrant();
                    twWarrant.WarrantCode = warrant.WarrantCode;
                    twWarrant.WarrantNameAbb = warrant.WarrantNameAbb;
                    twWarrant.Type = warrant.Type;
                    newWarrantList.Add(twWarrant);
                }
            }
            return newWarrantList;
        }

        #endregion

        #region Step2 get detail information of each warrant

        private List<TWWarrant> UpdateWarrants(List<TWWarrant> warrantListNeedDetail)
        {
            List<TWWarrant> reUpdateWarrant = new List<TWWarrant>();
            List<List<string>> deleteWarrant = new List<List<string>>();

            int maxRetry = 5;
            int retryTimes = 0;

            for (int i = 0; i < warrantListNeedDetail.Count; i++)
            {
                retryTimes = 0;

                try
                {
                    TWWarrant warrant = warrantListNeedDetail[i];
                    string postData = string.Empty;
                    Regex r = new Regex("");
                    if (warrant.Type == "TW")
                    {
                        postData = string.Format("encodeURIComponent=1&run=Y&step=1&TYPEK=sii&GOOD_ID={0}&firstin=true&off=1", warrant.WarrantCode);
                        r = new Regex("上市日期.*?<td class='odd'.*?(?<ListingDate>民國\\d{1,}年\\d{2}月\\d{1,}日)", RegexOptions.IgnoreCase);
                    }
                    if (warrant.Type == "TWO")
                    {
                        postData = string.Format("encodeURIComponent=1&run=Y&step=1&TYPEK=otc&GOOD_ID={0}&firstin=true&off=1", warrant.WarrantCode);
                        r = new Regex("上櫃日期.*?<td class='odd'.*?(?<ListingDate>民國\\d{1,}年\\d{2}月\\d{1,}日)", RegexOptions.IgnoreCase);
                    }
                    string pageSource = WebClientUtil.GetPageSource(warrantDetailUrl, 180000, postData);
                    pageSource = pageSource.Replace("\n", string.Empty).Replace("\r", string.Empty);
                    Match m = r.Match(pageSource);

                    if (!pageSource.Contains("查無此筆資料"))
                    {
                        string listingDate = MiscUtil.GetCleanTextFromHtml(m.Groups["ListingDate"].Value);

                        // If the server returns that "query too frequently". Sleep 21s until get the detail data.
                        while (listingDate == "" && retryTimes < maxRetry)
                        {
                            ++retryTimes;
                            Logger.Log(string.Format("Retry to get info of {0}: {1}", warrant.WarrantCode, retryTimes));

                            try
                            {
                                System.Threading.Thread.Sleep(10500);
                                pageSource = WebClientUtil.GetPageSource(warrantDetailUrl, 180000, postData);
                                pageSource = pageSource.Replace("\n", string.Empty).Replace("\r", string.Empty);
                                m = r.Match(pageSource);
                                listingDate = MiscUtil.GetCleanTextFromHtml(m.Groups["ListingDate"].Value);
                            }
                            catch (Exception)
                            {
                                Logger.Log(string.Format("Failed when get detail of warrant {0}", warrant.WarrantCode));
                            }
                        }

                        if (listingDate != "")
                        {
                            //Only get the warrant that its effective day is greater than yesterday.                          
                            DateTime effectiveDate = ConvertTWYear(listingDate);
                            //DateTime currentDate = System.DateTime.Now;
                            //TimeSpan ts = effectiveDate.Subtract(currentDate);
                            //if (ts.Days >= -3)
                            //{
                            warrant.ListingDate = effectiveDate.ToString("dd-MMM-yy", new System.Globalization.CultureInfo("en-US"));

                            r = new Regex("權證類型.*?<td class='odd'.*?>(?<WarrantType>.*?)</TD>", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.WarrantType = MiscUtil.GetCleanTextFromHtml(m.Groups["WarrantType"].Value);

                            r = new Regex("履約方式.*?<td class='odd'.*?>(?<PaymentType>.*?)</TD>", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.PaymentType = m.Groups["PaymentType"].Value;

                            if (string.IsNullOrEmpty(warrant.WarrantNameAbb))
                            {
                                r = new Regex("權證簡稱.*?<td class='odd'.*?>(?<WarrantName>.*?)</TD>", RegexOptions.IgnoreCase);
                                m = r.Match(pageSource);
                                warrant.WarrantNameAbb = m.Groups["WarrantName"].Value;
                            }

                            r = new Regex("發行單位數量.*?<td class='odd'.*?>(?<IssuerSum>.*?)</TD>", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.IssueSum = MiscUtil.GetCleanTextFromHtml(m.Groups["IssuerSum"].Value).Replace(",", "");


                            r = new Regex("權證英文簡稱.*?<td class='odd'.*?>(?<EnglishNameAbb>.*?)</TD>", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.WarrantEnglishNameAbb = MiscUtil.GetCleanTextFromHtml(m.Groups["EnglishNameAbb"].Value);

                            r = new Regex("原始上限價格.*?<td class='odd'.*?>(?<OrigCeilingPrice>.*?)</TD>", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.OrigCeilingPrice = MiscUtil.GetCleanTextFromHtml(m.Groups["OrigCeilingPrice"].Value);

                            r = new Regex("原始下限價格.*?<td class='odd'.*?>(?<OrigLowerPrice>.*?)</TD>", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.OrigLowerPrice = MiscUtil.GetCleanTextFromHtml(m.Groups["OrigLowerPrice"].Value);

                            r = new Regex("中/英文名稱.*?<tr class='odd'>.*?\\d{4}.*?<td.*?>(?<ChiEngNameAbb>.*?)</td>");
                            m = r.Match(pageSource);
                            warrant.ChiEngNameAbb = MiscUtil.GetCleanTextFromHtml(m.Groups["ChiEngNameAbb"].Value);

                            r = new Regex("最新標的履約配發數量.*?<td.*?<td.*?<td.*?<td.*?>(?<newTargetSum>.*?)</td>", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.NewTargetSum = MiscUtil.GetCleanTextFromHtml(m.Groups["newTargetSum"].Value);

                            r = new Regex("申請機構名稱.*?<td class='odd'.*?>(?<IssuerOrgName>.*?)</TD", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.IssuerOrgName = MiscUtil.GetCleanTextFromHtml(m.Groups["IssuerOrgName"].Value);

                            r = new Regex("履約開始日.*?<td class='odd'.*?>(?<ContractStartDay>.*?)</TD", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.ContractStartDay = MiscUtil.GetCleanTextFromHtml(m.Groups["ContractStartDay"].Value);

                            r = new Regex("履約截止日.*?<td class='odd'.*?>(?<ContractExpireDay>.*?)</TD", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.ContractExpireDay = MiscUtil.GetCleanTextFromHtml(m.Groups["ContractExpireDay"].Value);

                            r = new Regex("標的證券/指數代號.*?<tr class='odd'>.*?>(?<TargetCode>.*?\\d{2,}\\s{0,})</td>");
                            m = r.Match(pageSource);
                            warrant.TargetCode = MiscUtil.GetCleanTextFromHtml(m.Groups["TargetCode"].Value);

                            r = new Regex("原始履約價格.*?<td class='odd'.*?>(?<OrigContracePrice>.*?)</TD", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.OrigContracePrice = MiscUtil.GetCleanTextFromHtml(m.Groups["OrigContracePrice"].Value);

                            r = new Regex("最新履約價格.*?</td>.*?>(?<NewContractPrice>.*?)</TD", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.NewContactPrice = MiscUtil.GetCleanTextFromHtml(m.Groups["NewContractPrice"].Value);

                            r = new Regex("到期日.*?<td class='odd'.*?(?<ExpireDay>民國\\d{1,}年\\d{2}月\\d{1,}日)", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.ExpireDay = MiscUtil.GetCleanTextFromHtml(m.Groups["ExpireDay"].Value);

                            r = new Regex("結算方式說明.*?</td>.*?>(?<SettlementIndicator>.*?)</TD", RegexOptions.IgnoreCase);
                            m = r.Match(pageSource);
                            warrant.SettlementIndicator = MiscUtil.GetCleanTextFromHtml(m.Groups["SettlementIndicator"].Value);

                            //}
                            //else
                            //{
                            //    Logger.Log(string.Format("Effective day of Warrant {0} is {1}, So delete it from the List.", warrant.WarrantCode, effectiveDate.ToString("yyyy-MM-dd")));
                            //    warrantListNeedDetail.RemoveAt(i);
                            //    i--;
                            //}
                        }
                        else
                        {
                            //i--;
                            Logger.Log(string.Format("Failed to get info of {0}, skipped", warrant.WarrantCode));
                        }
                    }
                    else
                    {
                        //Put the 查無此筆資料's warrants into deleteItem and output. Then run it tomorrow.
                        List<string> deleteItem = new List<string>();
                        deleteItem.Add(string.Format("'{0}", warrant.WarrantCode));
                        deleteItem.Add(warrant.WarrantNameAbb);
                        deleteItem.Add(warrant.Type);
                        deleteItem.Add("查無此筆資料");
                        deleteWarrant.Add(deleteItem);
                        Logger.Log(string.Format("Cannot get detail of Warrant {0}(查無此筆資料), So delete it from the List.", warrant.WarrantCode));
                        warrantListNeedDetail.RemoveAt(i);
                        i--;
                    }
                }

                catch (Exception)
                {
                    Logger.Log(string.Format("Cannot update warrant {0}", warrantListNeedDetail[i].WarrantCode));
                }
            }

            if (deleteWarrant.Count != 0)
            {
                WriteDeleteWarrantToXls(deleteWarrant);
            }

            return warrantListNeedDetail;
        }

        /// <summary>
        /// Write the delete data into MISSED_WARRANTS.xls
        /// </summary>
        /// <param name="deleteWarrant">data to write</param>
        private void WriteDeleteWarrantToXls(List<List<string>> deleteWarrant)
        {
            if (File.Exists(missedWarrantFileFullName))
            {
                MiscUtil.BackupFileWithNewName(missedWarrantFileFullName);
            }

            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, missedWarrantFileFullName);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 2, 1, ExcelLineWriter.Direction.Right))
                {
                    foreach (List<string> items in deleteWarrant)
                    {
                        for (int i = 0; i < items.Count; i++)
                        {
                            writer.WriteLine(items[i]);
                        }
                        writer.PlaceNext(writer.Row + 1, 1);
                    }
                }
                worksheet.UsedRange.NumberFormat = "@";
                try
                {
                    workbook.Save();
                }
                catch (Exception)
                {
                    Logger.Log("Failed to write delete warrants to MISSED_WARRANTS.xls. The backup is in folder BACKUP");
                    WriteWarrantToTxt(deleteWarrant);
                }
                AddResult("Missed warrants", missedWarrantFileFullName, "warrant");
                //TaskResultList.Add(new TaskResultEntry("MISSED_WARRANTS", "", missedWarrantFileFullName));
            }
        }

        /// <summary>
        /// If failed to write the delete data to xls. then create a txt backup.
        /// </summary>
        /// <param name="deleteWarrant"></param>
        private void WriteWarrantToTxt(List<List<string>> deleteWarrant)
        {
            string filePath = Path.Combine(configObj.WorkingPath + "\\BACKUP", "MissedWarrantsBackup_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt");
            string[] content = new string[deleteWarrant.Count + 2];
            content[0] = DateTime.Now + "\tDelete warrants count:" + deleteWarrant.Count;
            content[1] = "Code\tName\tType\tMissReason";
            for (int i = 0; i < deleteWarrant.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(deleteWarrant[i][0] + "\t" + deleteWarrant[i][1] + "\t" + deleteWarrant[i][2] + "\t" + deleteWarrant[i][3]);
                content[i + 2] = sb.ToString();
                sb.Remove(0, sb.Length);
            }
            WriteTxtFile(filePath, content);
            AddResult("Missed warrant backup", filePath, "warrant");
            //TaskResultList.Add(new TaskResultEntry("MissedWarrantsBackup.txt", filePath, filePath));

        }

        private void WriteTxtFile(string fullpath, string[] content)
        {
            if (!Directory.Exists(Path.GetDirectoryName(fullpath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(fullpath));
            }
            try
            {
                File.WriteAllLines(fullpath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                string errInfo = ex.ToString();
            }
        }
        #endregion

        #region Step3 update KGI's infomation to DB, get issue date and price from DB
        private void UpdateIssueDateAndRationBasedKGI(List<TWWarrant> newWarrantList)
        {
            List<TWWarrant> warrantsFromKGIList = GetWarrantsFromKGI();

            if (warrantsFromKGIList == null)
            {
                LogMessage("Failed to get warrant info from KGI website, end the KGI process.", Logger.LogType.Warning);
                return;
            }

            UpdateDatabaseUseKGI(warrantsFromKGIList);

            //find the infomation from database

            foreach (TWWarrant warrant in newWarrantList)
            {
                TWIssueDatePriceInfo infoExist = TWIssueDatePriceManager.GetByWarrantName(warrant.WarrantNameAbb);
                if (infoExist != null)
                {
                    DateTime issueDate;
                    string dtStr = infoExist.IssueDate;

                    if (DateTime.TryParse(infoExist.IssueDate, out issueDate))
                    {
                        dtStr = issueDate.ToString("yy-MMM-dd", new CultureInfo("en-US"));
                    }

                    warrant.IssueDate = dtStr;
                    warrant.IssuePrice = infoExist.IssuePrice;
                }
                else
                {
                    warrant.IssueDate = "";
                    warrant.IssuePrice = "";
                    LogMessage(string.Format("Can't find warrant {0} in KGI.", warrant.WarrantCode));
                }
            }
        }

        private List<TWWarrant> GetWarrantsFromKGI()
        {
            string postData = "{0}&ctl00%24DropDownList2=%E5%85%A8%E9%83%A8&ctl00%24DropDownList1=%E5%85%A8%E9%83%A8&ctl00%24ContentPlaceHolder1%24DropDownList1=%E6%89%80%E6%9C%89%E7%99%BC%E8%A1%8C%E5%88%B8%E5%95%86&ctl00%24ContentPlaceHolder1%24DropDownList2=ALL&ctl00%24ContentPlaceHolder1%24ImageButton8.x=10&ctl00%24ContentPlaceHolder1%24ImageButton8.y=6";

            try
            {
                string page = WebClientUtil.GetPageSource(kgiUrl, 180000);

                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(page);

                List<string> aspxFormData = new List<string>();

                aspxFormData.Add("__EVENTTARGET=" + HttpUtility.UrlEncode(doc.GetElementbyId("__EVENTTARGET").GetAttributeValue("value", "")));
                aspxFormData.Add("__EVENTARGUMENT=" + HttpUtility.UrlEncode(doc.GetElementbyId("__EVENTARGUMENT").GetAttributeValue("value", "")));
                aspxFormData.Add("__VIEWSTATE=" + HttpUtility.UrlEncode(doc.GetElementbyId("__VIEWSTATE").GetAttributeValue("value", "")));
                aspxFormData.Add("__EVENTVALIDATION=" + HttpUtility.UrlEncode(doc.GetElementbyId("__EVENTVALIDATION").GetAttributeValue("value", "")));

                postData = string.Format(postData, string.Join("&", aspxFormData.ToArray()));

                if (File.Exists(kgiFileFullName))
                {
                    File.Delete(kgiFileFullName);
                }

                WebClientUtil.DownloadFile(kgiUrl, 180000, kgiFileFullName, postData);
            }
            catch (Exception ex)
            {
                LogMessage("Download excel file from KGI website failed: " + ex.Message, Logger.LogType.Error);
                return null;
            }

            System.Threading.Thread.Sleep(2000);

            return GetWarrantsFromExcel(kgiFileFullName);
        }

        /// <summary>
        /// Download today's KGI file and update the table ETI_TW_ISSUE_DATE_PRICE. 
        /// </summary>
        private void UpdateDatabaseUseKGI(List<TWWarrant> warrantsFromKGIList)
        {
            foreach (TWWarrant warrant in warrantsFromKGIList)
            {
                DateTime dt;
                string issueDate = warrant.IssueDate;

                if (DateTime.TryParseExact(warrant.IssueDate.Replace("/", "-"), new string[] { "MM-dd-yyyy", "M-d-yyyy", "dd-MM-yyyy", "d-M-yyyy" }, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
                {
                    issueDate = dt.ToString("dd-MMM-yyyy");
                }
                else if (DateTime.TryParse(warrant.IssueDate, out dt))
                {
                    issueDate = dt.ToString("dd-MMM-yyyy");
                }

                try
                {
                    TWIssueDatePriceInfo infoExist = TWIssueDatePriceManager.GetByWarrantName(warrant.WarrantNameAbb);
                    if (infoExist != null)
                    {
                        TWIssueDatePriceInfo info = new TWIssueDatePriceInfo();
                        info.ShortName = warrant.KGIShortrName;
                        info.WarrantName = warrant.WarrantNameAbb;
                        info.IssueDate = issueDate;
                        info.IssuePrice = warrant.IssuePrice;
                        TWIssueDatePriceManager.Update(info);
                    }
                    else
                    {
                        TWIssueDatePriceInfo info = new TWIssueDatePriceInfo();
                        info.ShortName = warrant.KGIShortrName;
                        info.WarrantName = warrant.WarrantNameAbb;
                        info.IssueDate = issueDate;
                        info.IssuePrice = warrant.IssuePrice;
                        TWIssueDatePriceManager.Insert(info);
                    }
                }
                catch (Exception ex)
                {
                    LogMessage("Failed update database using data from KGI website: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Open the target excel file. Get the data and fill into warrantList 
        /// </summary>
        /// <param name="filePath">the file path needed to load data</param>
        /// <returns></returns>
        private List<TWWarrant> GetWarrantsFromExcel(string filePath)
        {
            try
            {
                List<TWWarrant> warrantList = new List<TWWarrant>();

                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                    var worksheet = workbook.Worksheets[1] as Worksheet;
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                    //in the file, the data range is from the 2rd row to (lastUsedRow-2) row.                
                    for (int i = 2; i <= lastUsedRow - 2; i++)
                    {
                        TWWarrant warrant = new TWWarrant();
                        warrant.WarrantNameAbb = ExcelUtil.GetRange(i, 2, worksheet).Text.ToString().Trim();
                        warrant.TargetCode = ExcelUtil.GetRange(i, 3, worksheet).Text.ToString().Trim();
                        warrant.KGIShortrName = ExcelUtil.GetRange(i, 4, worksheet).Text.ToString().Trim();
                        warrant.IssueDate = ExcelUtil.GetRange(i, 5, worksheet).Text.ToString().Trim();
                        warrant.IssuePrice = ExcelUtil.GetRange(i, 6, worksheet).Text.ToString().Trim();
                        warrant.IssueSum = ExcelUtil.GetRange(i, 7, worksheet).Text.ToString().Trim();
                        warrant.OrigContracePrice = ExcelUtil.GetRange(i, 8, worksheet).Text.ToString().Trim();
                        warrantList.Add(warrant);
                    }

                    workbook.Save();
                    workbook.Close();
                }

                return warrantList;
            }
            catch (Exception ex)
            {
                LogMessage("Failed to parse excel file downloaded from KGI website: " + ex.Message, Logger.LogType.Error);
                return null;
            }
        }

        #endregion

        #region functions to form the data of FMTemplate.

        private DateTime ConvertTWYear(string TWDate)
        {
            Regex r = new Regex("民國(?<year>\\d{1,})年(?<month>\\d{1,})月(?<day>\\d{1,})日");
            Match m = r.Match(TWDate);
            int year = (int.Parse(m.Groups["year"].Value)) + 1911;
            int month = int.Parse(m.Groups["month"].Value);
            int day = int.Parse(m.Groups["day"].Value);
            return new DateTime(year, month, day);
        }

        private string GetDisplayName(TWWarrant warrant, string underlyingRic)
        {
            string displayName = string.Empty;
            string chineseShortName = warrant.ChineseShortName;
            try
            {
                displayName += TWIssueManager.GetByChineseShortName(chineseShortName).EnglishShortName;
                displayName += "-";
                if (warrant.isIndex)
                {
                    displayName += underlyingRic.Replace(".", "").Trim();
                }
                else
                {
                    displayName += warrant.TargetCode;
                }
                displayName += " ";
                displayName += warrant.callPut;

                string expireDay = ConvertTWYear(warrant.ExpireDay).ToString("yyMM", new System.Globalization.CultureInfo("en-US"));
                displayName += expireDay;
            }
            catch (Exception)
            {
                Logger.Log(string.Format("Cannot get displyName for warrant {0}. Issue: {1} is not in Database(TABLE:ETI_TW_ISSUE_INFO/COLUMN:ChineseShortName) ", warrant.WarrantCode, chineseShortName));
            }
            return displayName;
        }

        private string GetLongLink4(TWWarrant warrant)
        {
            if (warrant.isCBBC)
            {
                return string.Empty;
            }
            return string.Format("{0}va.{1}", warrant.WarrantCode, warrant.Type);
        }

        private string GetLongLink5(TWWarrant warrant)
        {
            if (warrant.isIndex)
            {
                return string.Empty;
            }
            return string.Format("0#{0}rel.{1}", warrant.TargetCode, warrant.Type);
        }

        private string GetLonglink1MenuPage(TWWarrant warrant)
        {
            if (warrant.Type == "TW")
            {
                return warrant.isCBBC ? "TW/CBBC01" : "TW/WTS1";
            }
            return warrant.isCBBC ? "TWO/CBBC01" : "TWO/WTS01";
        }

        private string GetLonkLink2_WT_Chain(TWWarrant warrant)
        {
            string lonkLink2 = string.Empty;

            if (warrant.isCBBC)
            {
                if (warrant.isIndex)
                {
                    lonkLink2 = string.Format("0#.TWIIC.{0}", warrant.Type);
                }
                else
                {
                    lonkLink2 = string.Format("0#{0}cbbc.{1}", warrant.TargetCode, warrant.Type);
                }
            }

            else
            {
                if (warrant.isIndex)
                {
                    lonkLink2 = string.Format("0#.TWIIW.{0}", warrant.Type);
                }
                else
                {
                    lonkLink2 = string.Format("0#{0}wt.{1}", warrant.TargetCode, warrant.Type);
                }
            }

            return lonkLink2;
        }

        private string GetISIN(TWWarrant warrant)
        {
            string url = string.Format("{0}?owncode={1}&stockname=", isinUrl, warrant.WarrantCode);
            try
            {
                string pageSource = WebClientUtil.GetPageSource(url, 18000);
                Regex r = new Regex("(?<ISIN>[-a-zA-Z_0-9]{12,14})", RegexOptions.IgnoreCase);
                Match m = r.Match(pageSource);
                return m.Groups["ISIN"].Value;
            }
            catch (Exception)
            {
                Logger.Log(string.Format("Error happans when getting ISIN for {0}", warrant.WarrantCode));
                return string.Empty;
            }

        }

        private string GetChainRic(TWWarrant warrant)
        {
            char indexNum = '0';
            for (int i = 0; i < warrant.WarrantCode.Length - 1; i++)
            {
                if (warrant.WarrantCode[i] != '0')
                {
                    indexNum = warrant.WarrantCode[i];
                    break;
                }
            }
            string chainRic = string.Empty;
            if (!warrant.isCBBC)
            {
                if (warrant.isETF)
                {
                    chainRic = string.Format("0#EWNT.{0}, 0#CWNT.{0}", warrant.Type);
                    chainRic += string.Format(", 0#{0}rel.TW, 0#{0}wt.TW, 0#CWNT{1}.TW", warrant.TargetCode, indexNum);
                }
                else if (warrant.isIndex)
                {
                    chainRic = string.Format("0#CWNT.{0}, 0#IWNT.{0}, 0#.TWIIW.{0}", warrant.Type);
                    chainRic += string.Format(", 0#CWNT{0}.{1}", indexNum, warrant.Type);
                }
                else
                {
                    chainRic = string.Format("0#SWNT.{0}, 0#CWNT.{0}", warrant.Type);
                    chainRic += string.Format(", 0#{0}rel.TW, 0#{0}wt.TW, 0#CWNT{1}.TW", warrant.TargetCode, indexNum);
                }
            }

            else
            {
                if (warrant.isIndex)
                {
                    chainRic = string.Format("0#CBBC.{0}, 0#.TWIIC.TW, 0#CWNT{1}.TW", warrant.Type, indexNum);
                }
                else
                {
                    chainRic = string.Format("0#CBBC.{0}, 0#{1}rel.{0}, 0#{1}cbbc.{0}, 0#CBBC3.{0}", warrant.Type, warrant.TargetCode);
                }
            }

            return chainRic;
        }

        private string GetCoiDisplyNmll(TWWarrant warrant)
        {
            string coiDisplyNmll = string.Empty;
            coiDisplyNmll += warrant.WarrantNameAbb;
            coiDisplyNmll += "-";
            coiDisplyNmll += warrant.ChiEngNameAbb;
            coiDisplyNmll += " " + warrant.callPut;

            string expireDay = ConvertTWYear(warrant.ExpireDay).ToString("yyMM", new System.Globalization.CultureInfo("en-US"));
            coiDisplyNmll += expireDay;
            return coiDisplyNmll;
        }

        private string GetCoiSectorChain(TWWarrant warrant)
        {
            string coiSectorChain = string.Empty;
            char indexNum = '0';
            for (int i = 0; i < warrant.WarrantCode.Length - 1; i++)
            {
                if (warrant.WarrantCode[i] != '0')
                {
                    indexNum = warrant.WarrantCode[i];
                    break;
                }
            }

            if (warrant.Type == "TWO")
            {
                if (!warrant.isCBBC)
                {
                    coiSectorChain += "股票備兌認股權證,備兌認股權證,";
                    coiSectorChain += warrant.ChiEngNameAbb;
                    coiSectorChain += ",";
                    coiSectorChain += warrant.ChiEngNameAbb;
                    coiSectorChain += ",";
                    coiSectorChain += "認股權證 (0)";
                }
                else
                {
                    coiSectorChain += "上櫃牛熊證,";
                    coiSectorChain += warrant.ChiEngNameAbb;
                    coiSectorChain += ",";
                    coiSectorChain += warrant.ChiEngNameAbb;
                    coiSectorChain += ",";
                    coiSectorChain += "上櫃牛熊證 (0)";
                }

                return coiSectorChain;
            }

            if (!warrant.isCBBC)
            {
                if (warrant.isIndex)
                {
                    coiSectorChain += "備兌認股權證,單一指數權證,台灣加權指數,認股權證 ";
                }
                else
                {
                    coiSectorChain += "股票備兌認股權證,備兌認股權證,";
                    coiSectorChain += warrant.ChiEngNameAbb;
                    coiSectorChain += ",";
                    coiSectorChain += warrant.ChiEngNameAbb;
                    coiSectorChain += ",";
                    coiSectorChain += "認股權證 ";
                }
            }

            else
            {
                if (warrant.isIndex)
                {
                    coiSectorChain += "上市牛熊證,單一指數牛熊證,台灣加權指數,上市牛熊證 ";
                }
                else
                {
                    coiSectorChain += "上市牛熊證,";
                    coiSectorChain += warrant.ChiEngNameAbb;
                    coiSectorChain += ",";
                    coiSectorChain += warrant.ChiEngNameAbb;
                    coiSectorChain += ",";
                    coiSectorChain += "上市牛熊證 ";
                }
            }
            coiSectorChain += "(0";
            coiSectorChain += indexNum.ToString();
            coiSectorChain += ")";

            return coiSectorChain;
        }

        private string GetBcastRef(TWWarrant warrant)
        {
            if (warrant.Type == "TWO")
            {
                return string.Format("{0}.TWO", warrant.TargetCode);
            }
            if (warrant.isIndex)
            {
                return ".TWII";//string.Format("{0}.TWII", warrant.TargetCode);
            }
            return string.Format("{0}.TW", warrant.TargetCode);
        }

        private string GetIDNLongName(TWWarrant warrant, TWUnderlyingNameInfo underlying)
        {
            string IDNLongName = string.Empty;

            if (warrant.isIndex)
            {
                IDNLongName += "TAIWAN WEIGHTED";
            }

            else
            {
                try
                {
                    IDNLongName += underlying.EnglishDisplay;
                }
                catch (Exception)
                {
                    IDNLongName += "*****";
                    Logger.Log(string.Format("Cannot get IDN LongName for warrant {0}.The ChiEngName(中/英文名稱) {1} is not in the Database(TABLE:ETI_TW_UNDERLYING_NAME/COLUMN:ChineseDisplay) .", warrant.WarrantCode + warrant.WarrantNameAbb, warrant.ChiEngNameAbb));
                }
            }

            IDNLongName += "@";

            string issuePart = string.Empty;
            string chineseNameShort = warrant.ChineseShortName;
            TWIssueInfo issueInfo = TWIssueManager.GetByChineseShortName(chineseNameShort);
            if (issueInfo == null)
            {
                return "";
            }

            IDNLongName += issueInfo.EnglishName;
            IDNLongName += " ";
            IDNLongName += ConvertTWYear(warrant.ExpireDay).ToString("MMMyy", new System.Globalization.CultureInfo("en-US")).ToUpper();
            IDNLongName += " ";
            IDNLongName += warrant.NewContactPrice;
            IDNLongName += warrant.callPut;
            IDNLongName += "WNT";
            return IDNLongName;
        }

        private string GetLocalSectorClassification(TWWarrant warrant)
        {
            if (warrant.isIndex)
            {
                return "Index warrants";
            }
            if (warrant.isETF)
            {
                return "ETF warrants";
            }
            return "covered warrants";
        }
        #endregion

        #region Step4 Generate the FM file
        /// <summary>
        /// Generate FM File for warrants. If the issue price/organization name/display name is null, then yellow it.
        /// </summary>
        /// <param name="fmFileDir">the directory to write FM file</param>
        /// <param name="newWarrantList">warrant list need to output</param>
        private void GenerateFMFile(List<TWWarrant> newWarrantList)
        {
            if (File.Exists(fmFileFullName))
            {
                MiscUtil.BackupFileWithNewName(fmFileFullName);
            }

            List<string> fmTitle = new List<string>
            { "RIC", "Issue Date", "Issue Price", "Cap. Price", "Effective Date", "Displayname", "Official Code", 
                                                        "Exchange Symbol", "OFFC_CODE2 (ISIN)", "Currency", "Recordtype", "Chain Ric", 
                                                        "Position in chain", "Lot Size", "COIDISPLY_NMLL", "COI SECTOR CHAIN", "BCAST_REF",
                                                        "WNT_RATIO", "STRIKE_PRC (WT)", "MATUR_DATE", "CONV_FAC", "ISIN", "IDN Longname",
                                                        "Issue Classification", "Primary Listing", "Organisation Name", "Underlying RIC",
                                                        "Issued Company Name", "RIC", "Local Sector Classification", "Index RIC(s)", 
                                                        "Total Shares Outstanding", "Composite Chain RIC", "Longlink 1 (Full Quote RIC)", 
                                                        "Longlink 2 (TA RIC)", "Longlink 3 (stat RIC)", "Longlink 4 (WNT Value Added RIC)", 
                                                        "Longlink 5 (stock Relative RICs)", "Longlink 6 (TAS RIC)", "Longlink 7 (Underlying RIC)", 
                                                        "Longlink 8 (DOM RIC)", "Longlink 9 (DIVCF)", "BOND_TYPE", "PUTCALLIND", 
                                                        "ISS_TP_FLG (Warrant Type)", "GEN_TEXT16 (Exercise Indicator)", 
                                                        "GN_TXT16_2 (Settlement Indicator)", "Longlink 1 (Issuer)", "Longlink 2 (Menu Page)", 
                                                        "Gearing", "Premium", "LONGLINK1 (#800) (TAS RIC)", "LONKLINK2 (#801) (wt Chain)",
                                                        "LONKLINK3 (#802) (Technology Analysis RIC)", "LONKLINK4 (#803) (Value Added RIC)" };

            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbookFM = ExcelUtil.CreateOrOpenExcelFile(app, fmFileFullName);
                Worksheet worksheet = workbookFM.Worksheets[1] as Worksheet;
                //title
                for (int i = 0; i < fmTitle.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = fmTitle[i];
                }
                for (int j = 0; j < newWarrantList.Count; j++)
                {
                    TWWarrant warrant = newWarrantList[j];
                    TWFMTemplate item = GetSingleTemplate(warrant);
                    worksheet.Cells[j + 2, 1] = item.Ric;
                    ((Range)(worksheet.Cells[j + 2, 2])).NumberFormat = "@";
                    worksheet.Cells[j + 2, 2] = item.IssueDate;
                    worksheet.Cells[j + 2, 3] = item.IssuePrice;
                    worksheet.Cells[j + 2, 4] = item.CapPrice;
                    ((Range)(worksheet.Cells[j + 2, 5])).NumberFormat = "@";
                    worksheet.Cells[j + 2, 5] = item.EffectiveDate;
                    worksheet.Cells[j + 2, 6] = item.DisplayName;
                    worksheet.Cells[j + 2, 7] = string.Format("'{0}", item.OfficialCode);
                    worksheet.Cells[j + 2, 8] = string.Format("'{0}", item.ExchangeSymbol);
                    worksheet.Cells[j + 2, 9] = item.OffcCode2;
                    worksheet.Cells[j + 2, 10] = item.Currency;
                    worksheet.Cells[j + 2, 11] = item.RecordType;
                    worksheet.Cells[j + 2, 12] = item.ChainRic;
                    worksheet.Cells[j + 2, 13] = item.PositionInChain;
                    worksheet.Cells[j + 2, 14] = item.LotSize;
                    worksheet.Cells[j + 2, 15] = item.CoiDisplyNmll;
                    worksheet.Cells[j + 2, 16] = item.CoiSectorChain;
                    worksheet.Cells[j + 2, 17] = item.BcastRef;
                    worksheet.Cells[j + 2, 18] = item.WntRatio;
                    worksheet.Cells[j + 2, 19] = string.Format("'{0}", item.StrikePrc);//保留2位小数
                    ((Range)(worksheet.Cells[j + 2, 20])).NumberFormat = "@";
                    worksheet.Cells[j + 2, 20] = item.MaturDate;
                    worksheet.Cells[j + 2, 21] = item.ConvFac;
                    worksheet.Cells[j + 2, 22] = item.Isin;
                    worksheet.Cells[j + 2, 23] = item.IDNLongName;
                    worksheet.Cells[j + 2, 24] = item.IssueClassification;
                    worksheet.Cells[j + 2, 25] = item.PrimaryListing;
                    worksheet.Cells[j + 2, 26] = item.OrganisationName;
                    worksheet.Cells[j + 2, 27] = item.UnderlyingRIC;
                    worksheet.Cells[j + 2, 28] = item.IssuedCompanyName;
                    worksheet.Cells[j + 2, 29] = item.Ric;
                    worksheet.Cells[j + 2, 30] = item.LocalSectorClassification;
                    worksheet.Cells[j + 2, 31] = item.IndexRic;
                    worksheet.Cells[j + 2, 32] = item.TotalSharesOutstanding;
                    worksheet.Cells[j + 2, 33] = item.CompositeChainRic;
                    worksheet.Cells[j + 2, 34] = item.LongLink1;
                    worksheet.Cells[j + 2, 35] = item.LongLink2;
                    worksheet.Cells[j + 2, 36] = item.LongLink3;
                    worksheet.Cells[j + 2, 37] = item.LongLink4;
                    worksheet.Cells[j + 2, 38] = item.LongLink5;
                    worksheet.Cells[j + 2, 39] = item.LongLink6;
                    worksheet.Cells[j + 2, 40] = item.LongLink7;
                    worksheet.Cells[j + 2, 41] = item.LongLink8;
                    worksheet.Cells[j + 2, 42] = item.LongLink9;
                    worksheet.Cells[j + 2, 43] = item.BondType;
                    worksheet.Cells[j + 2, 44] = item.PutCallInd;
                    worksheet.Cells[j + 2, 45] = item.ISS_TP_FLG;
                    worksheet.Cells[j + 2, 46] = item.GEN_TEXT16;
                    worksheet.Cells[j + 2, 47] = item.GN_TXT16_2;
                    worksheet.Cells[j + 2, 48] = item.Longlink1_Issuer;
                    worksheet.Cells[j + 2, 49] = item.Longlink2_MenuPage;
                    worksheet.Cells[j + 2, 50] = "";//Gearing
                    worksheet.Cells[j + 2, 51] = ""; //Premium
                    worksheet.Cells[j + 2, 52] = item.LONGLINK1_TAS_RIC;
                    worksheet.Cells[j + 2, 53] = item.LONKLINK2_WT_Chain;
                    worksheet.Cells[j + 2, 54] = item.LONKLINK3_Tech_Ric;
                    worksheet.Cells[j + 2, 55] = item.LONKLINK4_ValueAdded_Ric;

                    //if issue price is "", set the issuePrice and issueDate Cells Yellow
                    if (item.IssuePrice == "")
                    {
                        ((Range)worksheet.Cells[j + 2, 1]).Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                        ((Range)worksheet.Cells[j + 2, 2]).Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                        ((Range)worksheet.Cells[j + 2, 3]).Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                    }
                }

                worksheet.UsedRange.NumberFormat = "@";
                workbookFM.Save();
            }
            AddResult("Fm file", fmFileFullName, "fm");
            //TaskResultList.Add(new TaskResultEntry("FM File", "FM File", fmFileFullName));

        }

        /// <summary>
        /// Get FM templater for each warrant
        /// </summary>
        /// <param name="warrant">TWWarrant warrant</param>
        /// <returns>TWFMTemplate tem</returns>
        private TWFMTemplate GetSingleTemplate(TWWarrant warrant)
        {
            TWFMTemplate tem = new TWFMTemplate();
            try
            {
                tem.Ric = string.Format("{0}.{1}", warrant.WarrantCode, warrant.Type);

                if (warrant.WarrantCode.EndsWith("B") || warrant.WarrantCode.EndsWith("C"))
                {
                    warrant.isCBBC = true;
                }
                if (warrant.TargetCode.StartsWith("IX"))
                {
                    warrant.isIndex = true;
                }
                if (warrant.TargetCode.StartsWith("00"))
                {
                    warrant.isETF = true;
                }
                if (warrant.WarrantType.Contains("認售權證"))
                {
                    warrant.callPut = "P";
                }

                DateTime dt;
                string dtStr = warrant.IssueDate;

                if (!string.IsNullOrEmpty(warrant.IssueDate) && DateTime.TryParseExact(warrant.IssueDate, "yy-MMM-dd", new CultureInfo("en-US"), DateTimeStyles.None, out dt))
                {
                    dtStr = dt.ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
                }

                warrant.ChineseShortName = SubStringChinese(warrant.WarrantNameAbb);
                if (!TWIssueManager.ExistChineseName(warrant.ChineseShortName))
                {
                    warrant.ChineseShortName = warrant.IssuerOrgName.Substring(0, 2);
                    if (!TWIssueManager.ExistChineseName(warrant.ChineseShortName))
                    {
                        TWIssueInfo issuer = IssuerAdd.Prompt(tem.Ric, warrant.WarrantNameAbb, warrant.IssuerOrgName);
                        int issueRow = TWIssueManager.InsertNewIssuer(issuer);
                        Logger.Log(string.Format("Insert {0} issue record to database.", issueRow));
                    }
                }

                TWUnderlyingNameInfo underlying = TWUnderlyingNameManager.GetByChiEngName(warrant.ChiEngNameAbb);
                if (underlying == null)
                {
                    underlying = TWUnderlyingNameManager.GetByCode(warrant.TargetCode);
                }
                if (underlying == null)
                {
                    underlying = UnderlyingAdd.Prompt(tem.Ric, warrant.ChiEngNameAbb, warrant.TargetCode);
                    int row = TWUnderlyingNameManager.InsertNewUnderlying(underlying);
                    Logger.Log(string.Format("Insert {0} underlying record to database.", row));
                }

                tem.UnderlyingRIC = underlying.UnderlyingRIC;
                tem.BcastRef = underlying.UnderlyingRIC;

                tem.IssueDate = string.IsNullOrEmpty(dtStr) ? "" : dtStr;
                tem.LONKLINK3_Tech_Ric = string.Format("{0}ta.{1}", warrant.WarrantCode, warrant.Type);
                tem.LONKLINK4_ValueAdded_Ric = warrant.isCBBC ? string.Empty : string.Format("{0}va.{1}", warrant.WarrantCode, warrant.Type);
                tem.IssuePrice = string.IsNullOrEmpty(warrant.IssuePrice) ? "" : warrant.IssuePrice;
                tem.CapPrice = FormatCapPrice(warrant);//warrant.isCBBC ? (string.IsNullOrEmpty(warrant.OrigCeilingPrice) ? warrant.OrigLowerPrice : warrant.OrigCeilingPrice) : "N/A";
                tem.EffectiveDate = warrant.ListingDate;
                tem.DisplayName = GetDisplayName(warrant, underlying.UnderlyingRIC);
                tem.OfficialCode = warrant.WarrantCode;
                tem.ExchangeSymbol = (warrant.Type == "TWO" ? string.Format("O{0}", warrant.WarrantCode) : warrant.WarrantCode);
                tem.LongLink1 = tem.Ric;
                tem.LongLink2 = string.Format("{0}ta.{1}", tem.OfficialCode, warrant.Type);
                tem.LongLink3 = string.Format("{0}stat.{1}", tem.OfficialCode, warrant.Type);
                tem.LongLink4 = GetLongLink4(warrant);
                tem.LongLink5 = GetLongLink5(warrant);
                tem.LongLink6 = string.Format("t{0}", tem.Ric);
                //tem.LongLink7 = string.Format("{0}{1}", warrant.TargetCode, (warrant.isIndex ? ".TW" : ".TWII"));
                tem.LongLink8 = string.Format("D{0}", tem.Ric);
                // tem.LongLink9 = warrant.isIndex ? ("NA") : string.Format("{0}DIVCF.TW", tem.LongLink7);
                tem.CompositeChainRic = string.Format("0#{0}", tem.Ric);
                tem.LONGLINK1_TAS_RIC = string.Format("t{0}", tem.Ric);
                tem.BondType = warrant.isIndex ? "Index WARRANTS" : "WARRANTS";
                //tem.UnderlyingRIC = warrant.isIndex ? (".TWII") : (warrant.TargetCode + "." + warrant.Type);
                tem.Longlink2_MenuPage = GetLonglink1MenuPage(warrant); //warrant.isCBBC ? "TW/CBBC01" : "TW/WTS1";
                tem.LONKLINK2_WT_Chain = GetLonkLink2_WT_Chain(warrant);
                tem.PutCallInd = string.Format("{0}_{1}", warrant.callPut, warrant.PaymentType == "1" ? "AM" : "EU");
                //tem.GN_TXT16_2 = warrant.callPut;
                tem.GEN_TEXT16 = warrant.PaymentType == "1" ? "A" : "E";
                tem.ConvFac = ((double.Parse(warrant.NewTargetSum, System.Globalization.NumberStyles.Float)) / 1000).ToString();
                tem.WntRatio = string.Format("'1:{0}", tem.ConvFac);
                tem.OffcCode2 = GetISIN(warrant);
                tem.ChainRic = GetChainRic(warrant);
                tem.CoiDisplyNmll = GetCoiDisplyNmll(warrant);
                tem.CoiSectorChain = GetCoiSectorChain(warrant);
                //tem.BcastRef = GetBcastRef(warrant);
                tem.LongLink7 = tem.BcastRef;
                string underlyingCode = tem.LongLink7.Split('.')[0];
                tem.LongLink9 = warrant.isIndex ? ("NA") : string.Format("{0}DIVCF.{1}", underlyingCode, warrant.Type);

                tem.StrikePrc = warrant.NewContactPrice.Replace(",", "");
                tem.MaturDate = ConvertTWYear(warrant.ExpireDay).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                tem.Isin = tem.OffcCode2;
                tem.IDNLongName = GetIDNLongName(warrant, underlying);//To Do
                tem.LocalSectorClassification = GetLocalSectorClassification(warrant);
                tem.TotalSharesOutstanding = ((int.Parse(warrant.IssueSum, NumberStyles.Integer)) * 1000).ToString("###,### units");
                string chineseShortName = warrant.ChineseShortName;

                tem.IssuedCompanyName = underlying.EnglishDisplay;
                TWIssueInfo issuerTemp = TWIssueManager.GetByChineseShortName(chineseShortName);
                tem.OrganisationName = issuerTemp.EnglishFullName;
                tem.PrimaryListing = issuerTemp.IssueCode;
                tem.Longlink1_Issuer = tem.PrimaryListing;
                if (warrant.isIndex)
                {
                    tem.ISS_TP_FLG = "I";
                }

                tem.GN_TXT16_2 = FormatSettlementIndicator(warrant);

            }
            catch (Exception ex)
            {
                Logger.Log(ex.ToString());
            }
            return tem;
        }

        private string FormatSettlementIndicator(TWWarrant warrant)
        {
            /*
             *  現金結算	C
                證券給付	S
                證券給付，惟 	B
                證券給付，如 	B
              
             * */

            if (warrant.SettlementIndicator.Contains("現金結算"))
            {
                return "C";
            }

            if (warrant.SettlementIndicator.Contains("證券給付") && (warrant.SettlementIndicator.Contains("惟") || warrant.SettlementIndicator.Contains("如")))
            {
                return "B";
            }

            if (warrant.SettlementIndicator.Contains("證券給付"))
            {
                return "S";
            }

            return "";
        }

        private string FormatCapPrice(TWWarrant warrant)
        {
            if (warrant.isCBBC)
            {
                return string.IsNullOrEmpty(warrant.OrigCeilingPrice) ? warrant.OrigLowerPrice : warrant.OrigCeilingPrice;
            }
            else
            {
                string contractPrice = warrant.NewTargetSum;
                contractPrice = contractPrice.Replace(",", "").Replace(" ", "").Trim();
                if (contractPrice.Contains("."))
                {
                    contractPrice = contractPrice.Split('.')[0];
                }
                return contractPrice;
            }
        }

        #endregion

        /// <summary>
        /// warrantNameAbb contains Chinese Characters and numbers or English Characters, 
        /// this method can get the Chinese short name of warrantNameAbb
        /// </summary>
        /// <param name="stringToSub">the string to cut</param>
        /// <returns></returns>
        public string SubStringChinese(string stringToSub)
        {
            string newString = "";
            for (int i = 0; i <= stringToSub.Length - 1; i++)
            {
                char c = Convert.ToChar(stringToSub.Substring(i, 1));
                if (((int)c > 255) || ((int)c < 0)) // is Chinese Characters
                {
                    newString += c;
                }
                else
                {
                    if (newString != "")
                    {
                        return newString;
                    }
                }
            }
            return newString;
        }

    }
}
