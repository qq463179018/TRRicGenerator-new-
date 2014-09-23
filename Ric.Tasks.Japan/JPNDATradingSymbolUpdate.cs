using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using Microsoft.Exchange.WebServices.Data;
using Ric.Db.Info;
using Ric.Db.Manager;
using System.Drawing.Design;
using System.Windows.Forms;
using MSAD.Common.OfficeUtility;
using System.Threading;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Japan
{
    [ConfigStoredInDB]
    class JPNDATradingSymbolUpdateConfig
    {
        [StoreInDB]
        [Category("OutputPath")]
        [Description("path of generate result file")]
        public string OutputPath { get; set; }

        [Category("EmailReceivedDate")]
        [Description("Email which day email you want to get, like: \"2013-05-06\" Default:Today's Email")]
        public string EmailDate { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [DefaultValue("UC159450")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [Description("Email Folder.like:Inbox\\")]
        [DefaultValue("Inbox")]
        public string MailFoder { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }

        DateTime dt = DateTime.Now.ToUniversalTime().AddHours(+8);

        public JPNDATradingSymbolUpdateConfig()
        {
            if (Convert.ToInt32(dt.ToString("yyyy-H").Remove(0, 5)) > 12)
                EmailDate = dt.ToString("yyyy-MM-dd");
            else
                EmailDate = dt.AddDays(-1).ToString("yyyy-MM-dd");
        }
    }

    class JPNDATradingSymbolUpdate : GeneratorBase
    {
        public static JPNDATradingSymbolUpdateConfig configObj = null;
        private ExchangeService service;
        EWSMailSearchQuery query;
        private EmailAccountInfo emailAccount = null;
        private string emailFolder = string.Empty;
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private List<string> attacheFileList = new List<string>();
        private List<string> listErrorEamil = new List<string>();
        private DateTime startDate;
        private DateTime endDate;
        private List<string> listRic = new List<string>();//ric got from iffm and stk
        private List<string> listRicIFFM = null;//iffm ric got from email
        private string keyWordIFFM = string.Empty;
        private string patternIFFM = string.Empty;
        private string iffmPath = string.Empty;
        private List<string> listRicOthers = null;//stk ric got from email and attachment
        private List<string> listKeyWordOthersAttachment = new List<string>();
        private List<string> listKeyWordOthersBody = new List<string>();
        private string patternOthersBody = string.Empty;
        private string othersPath = string.Empty;
        private string outPutPath = string.Empty;
        //bulk iffm
        Dictionary<string, List<string>> dicListIFFM = new Dictionary<string, List<string>>();
        //bulk file index
        Dictionary<string, List<string>> dicListOthers = new Dictionary<string, List<string>>();

        protected override void Initialize()
        {
            configObj = Config as JPNDATradingSymbolUpdateConfig;

            if (!string.IsNullOrEmpty(configObj.AccountName))
                emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());

            if (emailAccount == null)
            {
                MessageBox.Show(configObj.AccountName + "is not exist in DB.");
                return;
            }

            emailFolder = configObj.MailFoder.Trim();
            service = MSAD.Common.OfficeUtility.EWSUtility.CreateService(new System.Net.NetworkCredential(emailAccount.AccountName, emailAccount.Password, emailAccount.Domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            startDate = Convert.ToDateTime(configObj.EmailDate);
            endDate = startDate.AddDays(+1);
            //keyWordIFFM = "OSEOMX IFFM - EDA Automation Report";
            keyWordIFFM = "OSD_OTFC Report for SO_6 of OSD feed";
            //patternIFFM = @"(?<RIC>[0-9A-Za-z]+\.OS)\s{5}\S+";
            patternIFFM = @"\s+(?<RIC>[0-9A-Z]+\.OS)\s+";
            listKeyWordOthersBody.Add("JNI Index Option additional strike price added");
            listKeyWordOthersBody.Add("JTI Index Option additional strike price added");
            patternOthersBody = @"(?<RIC>[^\\n][0-9A-Za-z]+\.OS)";
            //listKeyWordOthersAttachment.Add("OSE JNI Index Options Rollover Ric List");
            //listKeyWordOthersAttachment.Add("OSE JTI Index Options Rollover Ric List");
            listKeyWordOthersAttachment.Add("OSE JNI/JTI Index Options Rollover Ric List");
            outPutPath = configObj.OutputPath.Trim();
            iffmPath = Path.Combine(outPutPath, string.Format("NDA_OSE_Stk_Opt_Ref_{0}.csv", DateTime.Now.ToString("ddMMMyyy")));
            othersPath = Path.Combine(outPutPath, string.Format("NDA_OSE_Index_Opt_Ref_{0}.csv", DateTime.Now.ToString("ddMMMyyy")));
            List<string> listIFFMTitle = new List<string>() { "RIC", "DERIVATIVES LAST TRADING DAY", "DERIVATIVES LOT SIZE", "TRADING SYMBOL" };
            dicListIFFM.Add("title", listIFFMTitle);
            List<string> listIndeTitle = new List<string>() { "RIC", "TRADING SYMBOL" };
            dicListOthers.Add("title", listIndeTitle);
        }

        protected override void Start()
        {
            //get ric from iffm email
            listRicIFFM = GetRicFromEmailBody(keyWordIFFM, patternIFFM);
            //get lot from idn use ric list
            Dictionary<string, string> dicIFFMLot = GetIFFMLot(listRicIFFM);
            //add ric.insert(4,"L");
            List<string> listRicIFFMAll = AddLOptionRic(listRicIFFM);
            //get csv body from idn
            FillInDicIFFMBody(listRicIFFMAll, dicIFFMLot, dicListIFFM);
            //generate iffm csv
            GenerateBulkFile(iffmPath, dicListIFFM);

            //get ric from others email
            listRicOthers = GetRicFromOthers();
            //add ric.insert(3,"L")   "1"+ric   "2"+ric 
            List<string> listOptionRicOthersALL = AddL12RicOthers(listRicOthers);
            //get csv body from idn
            FillInDicOthersBody(listOptionRicOthersALL, dicListOthers);
            //generate iffm csv
            GenerateBulkFile(othersPath, dicListOthers);

            SendEmial();
        }

        private void SendEmial()
        {
            string subject = string.Empty;
            string content = string.Empty;

            if (listErrorEamil.Count == 0)
            {
                if (attacheFileList.Count != 0)
                {
                    subject = "Japan NDA Trading Symbol Update" + DateTime.Now.ToString("dd-MMM-yyyy");
                    content = " ";
                    SendMail(subject, content, attacheFileList);
                }
                else//ok
                {
                    subject = "Japan NDA Trading Symbol Update" + DateTime.Now.ToString("dd-MMM-yyyy");
                    content = "no ric today";
                    SendMail(subject, content, new List<string>());
                }
            }
            else
            {
                string failedEmail = GetFailedList(listErrorEamil);

                if (attacheFileList.Count != 0)
                {
                    subject = "Japan NDA Trading Symbol Update" + DateTime.Now.ToString("dd-MMM-yyyy");
                    content = "GetFollowingEmailError:\r\n" + failedEmail;
                    SendMail(subject, content, attacheFileList);
                }
                else
                {
                    subject = "Japan NDA Trading Symbol Update" + DateTime.Now.ToString("dd-MMM-yyyy");
                    content = "GetFollowingEmailError:\r\n" + failedEmail;
                    SendMail(subject, content, new List<string>());
                }
            }
        }

        private string GetFailedList(List<string> listErrorEamil)
        {
            string result = string.Empty;

            foreach (var item in listErrorEamil)
            {
                result += item + "\r\n";
            }

            return result;
        }

        private void FillInDicOthersBody(List<string> listOptionRicOthersALL, Dictionary<string, List<string>> dicListOthers)
        {
            List<string> listQueryString = GetQueryString(listOptionRicOthersALL);

            if (listQueryString == null || listQueryString.Count == 0)
                return;

            foreach (var item in listQueryString)
            {
                GetIDNOthers(item, dicListOthers);
            }
        }

        private void GetIDNOthers(string item, Dictionary<string, List<string>> dicList)
        {
            try
            {
                string pattern = @"(?<RIC>[0-9A-Z]+\.OS)\s+PROV_SYMB\s+(?<PROV_SYMB>[0-9\-]+)";
                string fids = "PROV_SYMB";
                GatsUtil gats = new GatsUtil();
                //string response = gats.GetGatsResponse(item, "PROD_PERM");
                string response = gats.GetGatsResponse(item, fids);
                Regex regex = new Regex(pattern);
                MatchCollection matches = regex.Matches(response);
                string ric = string.Empty;

                foreach (Match match in matches)
                {
                    ric = match.Groups["RIC"].Value.ToString().Trim();
                    List<string> list = new List<string>();
                    list.Add(ric);
                    list.Add(match.Groups["PROV_SYMB"].Value.ToString().Trim());

                    dicList.Add(ric, list);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("get exist ric in the gats error.:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        private List<string> AddL12RicOthers(List<string> listRicOthers)
        {
            List<string> list = new List<string>();
            string ricL = string.Empty;
            string ric1 = string.Empty;
            string ric2 = string.Empty;

            if (listRicOthers == null)
                return null;

            if (listRicOthers.Count == 0)
                return list;

            foreach (var item in listRicOthers)
            {
                if (!list.Contains(item))
                    list.Add(item);

                ricL = item.Insert(3, "L");
                if (!list.Contains(ricL))
                    list.Add(ricL);

                ric1 = "1" + item;
                if (!list.Contains(ric1))
                    list.Add(ric1);

                ric2 = "2" + item;
                if (!list.Contains(ric2))
                    list.Add(ric2);
            }

            return list;
        }

        private void GenerateBulkFile(string file, Dictionary<string, List<string>> dicList)
        {
            if (dicList == null || dicList.Count <= 1)
            {
                string msg = string.Format("no column need to generate bulk file.");
                Logger.Log(msg, Logger.LogType.Info);
                return;
            }

            try
            {
                if (!Directory.Exists(outPutPath))
                {
                    Directory.CreateDirectory(outPutPath);
                }

                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("delete old the file :{0} error. msg:{1}", file, ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                XlsOrCsvUtil.GenerateStringCsv(file, dicList);
                attacheFileList.Add(file);
                AddResult("bulk file", file, "CSV Bulk File");
                //TaskResultList.Add(new TaskResultEntry(MethodBase.GetCurrentMethod().DeclaringType.FullName.Replace("Ric.Generator.Lib.", ""), "ResultFile", file));
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate the file :{0} error. msg:{1}", file, ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private Dictionary<string, string> GetIFFMLot(List<string> listOptionRicIFFM)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            List<string> listQueryString = GetQueryString(listOptionRicIFFM);

            if (listQueryString == null || listQueryString.Count == 0)
                return null;

            foreach (var item in listQueryString)
            {
                GetIFFMLotFromIDN(item, dic);
            }

            return dic;
        }

        private void GetIFFMLotFromIDN(string item, Dictionary<string, string> dic)
        {
            try
            {
                string pattern = @"(?<RIC>[0-9A-Z]+\.OS)\s+LOT_SIZE_A\s+(?<LOT>[0-9\.]+)";
                string fids = "LOT_SIZE_A";
                GatsUtil gats = new GatsUtil();
                //string response = gats.GetGatsResponse(item, "PROD_PERM");
                string response = gats.GetGatsResponse(item, fids);
                Regex regex = new Regex(pattern);
                MatchCollection matches = regex.Matches(response);
                string ric = string.Empty;
                string lot = string.Empty;
                string ricL = string.Empty;

                foreach (Match match in matches)
                {
                    ric = match.Groups["RIC"].Value.ToString().Trim();
                    lot = match.Groups["LOT"].Value.ToString().Trim();

                    if (!dic.ContainsKey(ric))
                        dic.Add(ric, lot);

                    ricL = ric.Insert(4, "L");

                    if (!dic.ContainsKey(ricL))
                        dic.Add(ricL, lot);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("get exist ric in the gats error.:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        private void FillInDicIFFMBody(List<string> listOptionRicIFFMAll, Dictionary<string, string> dicLot, Dictionary<string, List<string>> dicListIFFM)
        {
            List<string> listQueryString = GetQueryString(listOptionRicIFFMAll);

            if (listQueryString == null || listQueryString.Count == 0)
                return;

            foreach (var item in listQueryString)
            {
                GetIDNIFFM(item, dicLot, dicListIFFM);
            }
        }

        private void GetIDNIFFM(string item, Dictionary<string, string> dicLot, Dictionary<string, List<string>> dicList)
        {
            try
            {
                string pattern = @"(?<RIC>[0-9A-Z]+\.OS)\s+EXPIR_DATE\s+(?<EXPIR_DATE>[0-9]{1,2}\s*[A-Z]+\s*[0-9]{2,4})\r\n[0-9A-Z]+\.OS\s+PROV_SYMB\s+(?<PROV_SYMB>[0-9\-]+)";
                string fids = "EXPIR_DATE,PROV_SYMB";
                GatsUtil gats = new GatsUtil();
                //string response = gats.GetGatsResponse(item, "PROD_PERM");
                string response = gats.GetGatsResponse(item, fids);
                Regex regex = new Regex(pattern);
                MatchCollection matches = regex.Matches(response);
                string ric = string.Empty;

                foreach (Match match in matches)
                {
                    ric = match.Groups["RIC"].Value.ToString().Trim();
                    List<string> list = new List<string>();
                    list.Add(ric);
                    list.Add(FormatExpirDate(match.Groups["EXPIR_DATE"].Value.ToString().Trim()));
                    list.Add(dicLot.ContainsKey(ric) ? dicLot[ric].Replace(".00", "").Replace(".0", "") : "  ");
                    list.Add(match.Groups["PROV_SYMB"].Value.ToString().Trim());

                    dicList.Add(ric, list);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("get exist ric in the gats error.:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        private string FormatExpirDate(string p)
        {
            string expirDate = string.Empty;

            try
            {
                if ((p + "").Length < 11)
                    return expirDate;

                string day = p.Substring(0, 2);
                string monthFirst = p.Substring(3, 1);
                string monthSecond = p.Substring(4, 2);
                string year = p.Substring(7, 4);
                expirDate = string.Format("{0}-{1}{2}-{3}", day, monthFirst, monthSecond.ToLower(), year);
                return expirDate;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return expirDate;
            }
        }

        private List<string> AddLOptionRic(List<string> listOptionRicIFFM)
        {
            List<string> list = new List<string>();
            string ricL = string.Empty;

            if (listOptionRicIFFM == null)
                return null;

            if (listOptionRicIFFM.Count == 0)
                return list;

            foreach (var item in listOptionRicIFFM)
            {
                if (!list.Contains(item))
                    list.Add(item);

                ricL = item.Insert(4, "L");
                if (!list.Contains(ricL))
                    list.Add(ricL);
            }

            return list;
        }

        private List<string> GetQueryString(List<string> listRic)
        {
            List<string> list = new List<string>();

            if (listRic == null || listRic.Count == 0)
                return null;

            string strQuery = string.Empty;
            int count = listRic.Count;
            int fenMu = 2000;
            int qiuYu = count % fenMu;
            int qiuShang = count / fenMu;

            if (qiuShang > 0)
            {
                for (int i = 0; i < qiuShang; i++)
                {
                    for (int j = 0; j < fenMu; j++)
                    {
                        string strTmp = listRic[i * fenMu + j].ToString().Trim();
                        if (!string.IsNullOrEmpty(strTmp))
                        {
                            strQuery += string.Format(",{0}", strTmp);
                        }
                    }

                    strQuery = strQuery.Remove(0, 1);
                    list.Add(strQuery);
                    strQuery = string.Empty;
                }
            }

            for (int i = qiuShang * fenMu; i < count; i++)
            {
                string strTmp = listRic[i].ToString().Trim();

                if (!string.IsNullOrEmpty(strTmp))
                {
                    strQuery += string.Format(",{0}", strTmp);
                }
            }

            strQuery = strQuery.Remove(0, 1);
            list.Add(strQuery);

            return list;
        }

        private List<string> GetRicFromOthers()
        {
            List<string> list = new List<string>();
            List<string> listRicBody = null;
            List<string> listRicAttachment = null;
            listRicBody = GetBodyRic();
            listRicAttachment = GetAttachmentRic();

            if (listRicBody != null && listRicBody.Count > 0)
            {
                list.AddRange(listRicBody);
                string msg = string.Format("get {0} ric from two of others email body .", listRicBody.Count);
                Logger.Log(msg, Logger.LogType.Info);
            }

            if (listRicAttachment != null && listRicAttachment.Count > 0)
            {
                list.AddRange(listRicAttachment);
                string msg = string.Format("get {0} ric from two of others email attachment file .", listRicAttachment.Count);
                Logger.Log(msg, Logger.LogType.Info);
            }

            return list;

        }

        private List<string> GetAttachmentRic()
        {
            List<string> list = new List<string>();
            List<string> listItem = null;
            string downloadPath = Path.Combine(outPutPath, "download");
            CreateFolder(downloadPath);
            DeleteFiles(downloadPath);

            foreach (var item in listKeyWordOthersAttachment)
            {
                listItem = GetAttachmentRic(item, downloadPath);

                if (listItem != null && listItem.Count > 0)
                    list.AddRange(listItem);
            }

            return list;
        }

        private List<string> GetAttachmentRic(string keyWord, string downloadPath)
        {
            List<string> list = new List<string>();

            try
            {
                query = new EWSMailSearchQuery("", emailAccount.MailAddress, emailFolder, keyWord, "", startDate, endDate);
                List<EmailMessage> mailList = GetEmailList(service, query, keyWord);

                if (mailList == null)
                {
                    string msg = string.Format("email account error.");
                    Logger.Log(msg, Logger.LogType.Error);
                    return null;//ric list ==null
                }

                if (mailList.Count == 0)
                {
                    string msg = string.Format("no email in this account within this email folder");
                    Logger.Log(msg, Logger.LogType.Warning);
                    return list;//ric list.count==0
                }

                EmailMessage email = mailList[0];
                email.Load();
                List<string> attachments = EWSMailSearchQuery.DownloadAttachments(service, email, "", "", downloadPath);
                return GetRicFromFile(attachments);
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetRicFromIFFM() failed. msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return list;
            }
        }

        private List<string> GetRicFromFile(List<string> attachments)
        {
            List<string> list = new List<string>();
            List<string> listItem = null;

            foreach (var item in attachments)
            {
                if (item.Contains(".xls") || item.Contains(".csv"))
                    listItem = GetRicFromFile(item);

                if (listItem != null && listItem.Count > 0)
                    list.AddRange(listItem);
            }

            return list;
        }

        private List<string> GetRicFromFile(string item)
        {
            List<string> list = new List<string>();
            try
            {
                using (Ric.Util.ExcelApp app = new Ric.Util.ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, item);
                    var worksheet = workbook.Worksheets[1] as Worksheet;

                    if (worksheet != null)
                    {
                        int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                        for (int i = 2; i <= lastUsedRow; i++)
                        {
                            object value = ExcelUtil.GetRange(i, 1, worksheet).Value2;
                            string ric = string.Empty;

                            if (value != null && value.ToString().Trim() != string.Empty)
                            {
                                ric = value.ToString().Trim();

                                if (ric.Contains("RIC") && list.Contains(ric))
                                    continue;

                                list.Add(ric);
                            }
                        }
                    }
                    workbook.Close(false, workbook.FullName, Missing.Value);
                }

                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetRicFromFile(string item) failed. item:{0},msg:{1}", item, ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }
        }

        private void CreateFolder(string path)
        {
            if (!Directory.Exists(path))
            {
                try
                {
                    Directory.CreateDirectory(path);
                }
                catch (Exception ex)
                {
                    string msg = string.Format("can't cerate directory {0} \r\n.ex:{1}", path, ex.ToString());
                    Logger.Log(msg, Logger.LogType.Error);
                }
            }
        }

        private void DeleteFiles(string path)
        {
            if (!Directory.Exists(path))
                return;

            DirectoryInfo fatherFolder = new DirectoryInfo(path);
            FileInfo[] files = fatherFolder.GetFiles();

            foreach (FileInfo file in files)
            {
                string fileName = file.Name;

                try
                {
                    File.Delete(file.FullName);
                }
                catch (Exception ex)
                {
                    string msg = string.Format("the file: {0} delete failed. please close pdf file.\r\nmsg: {1}", file.Name, ex.ToString());
                    Logger.Log(msg, Logger.LogType.Error);
                }
            }

            foreach (DirectoryInfo childFolder in fatherFolder.GetDirectories())
            {
                DeleteFiles(childFolder.FullName);
            }
        }

        private List<string> GetBodyRic()
        {
            List<string> list = new List<string>();
            List<string> listItem = null;

            foreach (var item in listKeyWordOthersBody)
            {
                listItem = GetRicFromEmailBody(item, patternOthersBody);

                if (listItem != null && listItem.Count > 0)
                    list.AddRange(listItem);
            }

            return list;
        }

        private List<string> GetRicFromEmailBody(string keyWord, string pattern)
        {
            List<string> list = new List<string>();

            try
            {
                query = new EWSMailSearchQuery("", emailAccount.MailAddress, emailFolder, keyWord, "", startDate, endDate);
                List<EmailMessage> mailList = GetEmailList(service, query, keyWord);

                if (mailList == null)
                {
                    string msg = string.Format("email account error.");
                    Logger.Log(msg, Logger.LogType.Error);
                    return null;//ric list ==null
                }

                if (mailList.Count == 0)
                {
                    string msg = string.Format("no email in this account within this email folder");
                    Logger.Log(msg, Logger.LogType.Warning);
                    return list;//ric list.count==0
                }

                EmailMessage email = mailList[0];
                email.Load();
                string body = email.Body.ToString();

                return GetRicIFFM(body, pattern);
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetRicFromIFFM() failed. msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return list;
            }
        }

        private List<string> GetRicIFFM(string body, string pattern)
        {
            List<string> list = new List<string>();
            string ric = string.Empty;

            if ((body + "").Trim() == "")
                return list;//ric count ==0;

            Regex regex = new Regex(pattern);
            MatchCollection matches = regex.Matches(ClearHtmlTags(body));

            foreach (Match match in matches)
            {
                ric = match.Groups["RIC"].Value.Trim();

                if (!list.Contains(ric))
                    list.Add(ric);
            }

            return list;
        }

        private string ClearHtmlTags(string strTags)
        {
            strTags = Regex.Replace(strTags, @"<script[^>]*?>.*?</script>", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
            //strTags = Regex.Replace(strTags, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"-->", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"<!--.*", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(quot|#34);", "\"", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(nbsp|#160);", " ", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&#(\d+);", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"<img[^>]*>;", "", RegexOptions.IgnoreCase);
            strTags.Replace("<", "");
            strTags.Replace(">", "");
            //strTags.Replace("\r\n", "");

            return strTags;
        }

        private List<EmailMessage> GetEmailList(ExchangeService service, EWSMailSearchQuery query, string keyword)
        {
            List<EmailMessage> list = null;

            for (int i = 0; ; i++)
            {
                try
                {
                    list = EWSMailSearchQuery.SearchMail(service, query);
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(2000);
                    if (i == 4)
                    {
                        string msg = string.Format("execute GetEmailList(ExchangeService service, EWSMailSearchQuery query) failed. msg:{0}", ex.Message);
                        Logger.Log(msg, Logger.LogType.Error);
                        listErrorEamil.Add(keyword);
                        return list;
                    }
                }
            }

            return list;
        }

        private void SendMail(string subject, string content, List<string> attacheFileList)
        {
            StringBuilder bodyBuilder = new StringBuilder();
            bodyBuilder.Append(content);
            bodyBuilder.Append("<p>");

            foreach (string signatureLine in configObj.MailSignature)
            {
                bodyBuilder.AppendFormat("{0}<br />", signatureLine);
            }

            bodyBuilder.Append("</p>");
            content = bodyBuilder.ToString();

            if (configObj.MailCC.Count > 1 || (configObj.MailCC.Count == 1 && configObj.MailCC[0] != ""))
            {
                listMailCC = configObj.MailCC;
            }

            MSAD.Common.OfficeUtility.EWSUtility.CreateAndSendMail(service, listMailTo, listMailCC, new List<string>(), subject, content, attacheFileList);
        }
    }
}
