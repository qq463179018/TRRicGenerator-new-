using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;
using MSAD.Common.OfficeUtility;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    #region Configuration
    [ConfigStoredInDB]
    class HKIPOIDNConfig
    {
        [StoreInDB]
        [Category("EmailAccount")]
        [DisplayName("Mail folder path")]
        [Description("Mail folder path,like: Inbox/XXXXX")]
        public string MailFolderPath { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [DisplayName("Name")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("File")]
        [DisplayName("Txt file path")]
        [Description("GeneratedFilePath")]
        public string TxtFilePath { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Recipients")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Recipients (Cc)")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Signature")]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }
    }
    #endregion

    class HKIPOIDN : GeneratorBase
    {
        #region Description
        private static HKIPOIDNConfig configObj = null;
        private string accountName = string.Empty;//UC169XXX
        private string password = string.Empty;//********
        private string domain = string.Empty;//TEN
        private string mailAdress = string.Empty;//eti.XXXXXX@thomsonreuters.com
        private string mailFolder = string.Empty;//Inbox/XXXXX
        private List<string> listMailTo = new List<string>();
        private List<string> listMailCC = new List<string>();
        private List<string> listMailSignature = new List<string>();
        private ExchangeService service;
        private List<string> listHKIPO = new List<string>();// first get ipo from email
        private bool isExistEmail = false;
        private bool isEmptyEmail = false;
        private string strEmailKeyWord = string.Empty;
        private DateTime startDate = DateTime.Now.ToUniversalTime().AddHours(+8).AddHours(-DateTime.Now.ToUniversalTime().AddHours(+8).Hour);
        private DateTime endDate = DateTime.Now.ToUniversalTime().AddHours(+8);
        private string strPatternEmail = string.Empty;
        private List<string> listAttachementFile = new List<string>();
        private Dictionary<string, string> dicHKIPOFromGATS = new Dictionary<string, string>();//get ipo from GATS
        private string strPatternGATS = string.Empty;
        private List<string> attacheFileList = new List<string>();
        private string txtFileNameAllRIC = string.Empty;
        private string txtFileNameRemovedRIC = string.Empty;
        private string txtFilePath = string.Empty;
        private bool isExistIPO = false;//IsExist ipo file in the attachement


        protected override void Initialize()
        {
            configObj = Config as HKIPOIDNConfig;
            EmailAccountInfo emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());
            accountName = emailAccount.AccountName;
            password = emailAccount.Password;
            domain = emailAccount.Domain;
            mailAdress = emailAccount.MailAddress;
            mailFolder = configObj.MailFolderPath.Trim();
            listMailTo = configObj.MailTo;
            listMailSignature = configObj.MailSignature;
            txtFilePath = configObj.TxtFilePath.Trim();
            strEmailKeyWord = "HKFM_";
            strPatternGATS = @"\r\n(?<Ric>[A-Za-z0-9#]+\.HK)\b\s+\bPROD_PERM\b\s+\b(?<Value>[0-9]+)\r\n"; ;//find ipo from GATS
            txtFileNameAllRIC = DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + "All_HK_IPO_IDN.txt"; ;//all ric file name
            txtFileNameRemovedRIC = DateTime.Now.ToUniversalTime().AddHours(+8).ToString("MMdd") + "Missing_HK_IPO_IDN.txt";//removed ric file name 
        }
        #endregion

        protected override void Start()
        {
            GetAttachementFromEmail(listAttachementFile, strEmailKeyWord, ref isExistEmail, ref isEmptyEmail);
            GetHKIPOFromAttachement(listHKIPO, listAttachementFile);
            AddNewRicList(listHKIPO);
            GenerateFile(listHKIPO, txtFilePath, txtFileNameAllRIC);//all ipo
            GetHKIPOFromGatsToDic(listHKIPO, strPatternGATS, dicHKIPOFromGATS);
            RemoveExistHKIPO(listHKIPO, dicHKIPOFromGATS);
            GenerateFile(listHKIPO, txtFilePath, txtFileNameRemovedRIC);//ipo removed
            SendEmail(listHKIPO);
        }

        private void AddNewRicList(List<string> listHKIPO)
        {
            try
            {
                if (listHKIPO == null || listHKIPO.Count == 0)
                    return;

                string newRic = string.Empty;
                int length = listHKIPO.Count;
                for (int i = 0; i < length; i++)
                {
                    newRic = listHKIPO[i].Replace(".HK", ".HS");
                    if ((newRic + "").Trim().Length == 0 || listHKIPO.Contains(newRic))
                        continue;

                    listHKIPO.Add(newRic);
                }
            }
            catch (Exception ex)
            {
                LogMessage("Add new ric list error. from *.HK TO *.HS.\r\n msg:" + ex.Message);
            }
        }


        #region DownloadAttachementFileFromEmail
        /// <summary>
        /// DownloadAttachementFileFromEmail
        /// </summary>
        /// <param name="listAttachementFile">store file name</param>
        /// <param name="strEmailKeyWord">subject</param>
        /// <param name="isExistEmail">bool</param>
        /// <param name="isEmptyEmail">bool</param>
        private void GetAttachementFromEmail(List<string> listAttachementFile, string strEmailKeyWord, ref bool isExistEmail, ref bool isEmptyEmail)
        {
            try
            {
                service = MSAD.Common.OfficeUtility.EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
                EWSMailSearchQuery query = new EWSMailSearchQuery("", mailAdress, mailFolder, strEmailKeyWord, "", startDate, endDate);
                List<EmailMessage> mailList = null;
                string strEmailContent = string.Empty;
                for (int i = 0; i < 5; i++)
                {
                    try
                    {
                        mailList = EWSMailSearchQuery.SearchMail(service, query);
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(5000);
                        if (i == 4)
                        {
                            throw;
                        }
                    }
                }
                if (mailList.Count > 0)
                {
                    isExistEmail = true;
                    string attachmentPath = Path.Combine(txtFilePath, "attachment");
                    EmailMessage mail = mailList[0];
                    if (!Directory.Exists(attachmentPath))
                    {
                        Directory.CreateDirectory(attachmentPath);
                    }
                    else
                    {
                        string[] rootFiles = Directory.GetFiles(attachmentPath);
                        foreach (string file in rootFiles)
                        {
                            File.Delete(file);
                        }
                    }
                    mail.Load();
                    List<string> attachments = EWSMailSearchQuery.DownloadAttachments(service, mail, "", "", attachmentPath);
                    if (attachments != null && attachments.Count != 0)
                    {
                        string dir = Path.GetDirectoryName(attachments[0]);
                        foreach (string zipFile in attachments)
                        {
                            string err = null;
                            if (!Ric.Util.ZipUtil.UnZipFile(zipFile, dir, out err))
                            {
                                Logger.Log(string.Format("Error happens when unzipping the file {0}. Exception message: {1}", zipFile, err));
                            }
                        }
                        foreach (var file in Directory.GetFiles(Path.GetDirectoryName(attachments[0]), "*.xls"))
                        {
                            listAttachementFile.Add(file);
                        }
                    }
                    else
                    {
                        isEmptyEmail = true;
                        Logger.Log("Found email but no attachement");
                    }
                }
                else
                {
                    Logger.Log("there is no email in mail box!");
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Get Data from mail failed. Ex: " + ex.Message);
            }
        }
        #endregion

        #region ReadXlsFile(step:one)
        /// <summary>
        /// ReadXlsFile
        /// </summary>
        /// <param name="listHKIPO">ric</param>
        /// <param name="listAttachementFile">file path</param>
        private void GetHKIPOFromAttachement(List<string> listHKIPO, List<string> listAttachementFile)
        {
            if (listAttachementFile != null && listAttachementFile.Count != 0)
            {
                string strFileName = string.Empty;
                foreach (string strFilePath in listAttachementFile)
                {
                    int start = strFilePath.LastIndexOf("\\");
                    strFileName = strFilePath.Substring(start + 1, strFilePath.Length - start - 1);
                    if (strFileName.Contains(".xls") && strFileName.Contains("_IPO_"))
                    {
                        isExistIPO = true;
                        GetHKIPOFromExcelFile(listHKIPO, strFilePath);
                    }
                }
            }
        }
        #endregion

        #region ReadXlsFile(step:two)
        /// <summary>
        /// ReadXlsFile
        /// </summary>
        /// <param name="listHKIPO">ric</param>
        /// <param name="strFilePath">file path</param>
        private void GetHKIPOFromExcelFile(List<string> listHKIPO, string strFilePath)
        {
            try
            {
                string worksheetName = "FM";
                using (Ric.Util.ExcelApp app = new Ric.Util.ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, strFilePath);
                    var worksheet = ExcelUtil.GetWorksheet(worksheetName, workbook);
                    if (worksheet != null)
                    {
                        int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                        for (int i = 1; i <= lastUsedRow; i++)
                        {
                            object key = ExcelUtil.GetRange(i, 1, worksheet).Value2;
                            object value = ExcelUtil.GetRange(i, 2, worksheet).Value2;
                            if (key != null && key.ToString().Trim() != string.Empty && value != null && value.ToString().Trim() != string.Empty)
                            {
                                if (key.ToString().Trim().Contains("Underlying RIC:") || key.ToString().Trim().Contains("Composite chain RIC:") || key.ToString().Trim().Contains("Broker page RIC:") || key.ToString().Trim().Contains("Misc.Info page RIC:"))
                                {
                                    if (!listHKIPO.Contains(value.ToString().Trim()))
                                    {
                                        listHKIPO.Add(value.ToString().Trim());
                                    }
                                }
                            }
                        }
                    }
                    workbook.Close(false, workbook.FullName, Missing.Value);
                }
            }
            catch (Exception e)
            {
                Logger.Log(string.Format("Error happens when get data. Ex: {0} .", e.Message));
            }
        }
        #endregion

        #region GenerateTxtFile
        /// <summary>
        /// Generate Txt File
        /// </summary>
        /// <param name="listHKIPO">ric</param>
        /// <param name="txtFilePath">file path</param>
        /// <param name="txtFileName">file name</param>
        private void GenerateFile(List<string> listHKIPO, string txtFilePath, string txtFileName)
        {
            if (isExistEmail && listHKIPO != null && listHKIPO.Count > 0)
            {
                string filePath = Path.Combine(txtFilePath, txtFileName);
                string content = string.Empty;

                foreach (var str in listHKIPO)
                    content += string.Format(",{0}", str);

                content = content.Remove(0, 1);
                try
                {
                    File.WriteAllText(filePath, content);
                    attacheFileList.Add(filePath);
                    AddResult("HKIPOQC", filePath, "file");
                    //TaskResultList.Add(new TaskResultEntry("HKIPOQC", "HKIPO", filePath));
                }
                catch (Exception ex)
                {
                    Logger.Log(string.Format("Error happens when generating file. Ex: {0} .", ex.Message));
                }
            }
        }
        #endregion

        #region GetDataToDicFromGATS(step:one)
        /// <summary>
        /// Get Data From GATS
        /// </summary>
        /// <param name="listHKIPO">Ric</param>
        /// <param name="strPatternGATS">pattern</param>
        /// <param name="dicHKIPOFromGATS">ric from gats</param>
        private void GetHKIPOFromGatsToDic(List<string> listHKIPO, string strPatternGATS, Dictionary<string, string> dicHKIPOFromGATS)
        {
            if (listHKIPO != null && listHKIPO.Count > 0)
            {
                string strQuery = string.Empty;
                foreach (string str in listHKIPO)
                {
                    strQuery += str + ",";
                }
                strQuery = strQuery.Remove(strQuery.Length - 1, 1);
                GetDataFromGATSToExistDic(strQuery, dicHKIPOFromGATS, strPatternGATS);
            }
        }
        #endregion

        #region GetDataToDicFromGATS(step:two)
        /// <summary>
        /// GetDataToDicFromGATS
        /// </summary>
        /// <param name="strQuery">input ric chain to gats</param>
        /// <param name="dicFromGATS">dic</param>
        /// <param name="strPatternGATS">pattern</param>
        private void GetDataFromGATSToExistDic(string strQuery, Dictionary<string, string> dicFromGATS, string strPatternGATS)
        {
            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(strQuery, "PROD_PERM");
            Regex regex = new Regex(strPatternGATS);
            MatchCollection matches = regex.Matches(response);
            string tmp = string.Empty;
            foreach (Match match in matches)
            {
                tmp = match.Groups["Ric"].Value;
                if (!dicFromGATS.ContainsKey(tmp))
                {
                    dicFromGATS.Add(match.Groups["Ric"].Value, match.Groups["Value"].Value);
                }
            }
        }
        #endregion

        #region CleanRicExistInDic
        /// <summary>
        /// CleanRicExistInDic
        /// </summary>
        /// <param name="listHKIPO">ric</param>
        /// <param name="dicHKIPOFromGATS">dic</param>
        private void RemoveExistHKIPO(List<string> listHKIPO, Dictionary<string, string> dicHKIPOFromGATS)
        {
            if (isExistEmail)
            {
                int index = listHKIPO.Count;
                for (int i = index - 1; i >= 0; i--)
                {
                    if (dicHKIPOFromGATS.ContainsKey(listHKIPO[i]))
                    {
                        listHKIPO.Remove(listHKIPO[i]);
                    }
                }
            }
        }
        #endregion

        #region SendEmail(step:one)
        /// <summary>
        /// send email
        /// </summary>
        /// <param name="listHKIPO">ric</param>
        private void SendEmail(List<string> listHKIPO)
        {
            string subject = string.Empty;
            string content = string.Empty;
            if (isExistEmail)
            {
                if (!isEmptyEmail)
                {
                    if (attacheFileList.Count == 2)//ok
                    {
                        subject = "HK IPO - IDN Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Exist Email]";
                        content = "<center>********IPO RIC missing********</center><br /><br />";
                    }
                    else if (attacheFileList.Count == 1)//ok
                    {
                        subject = "HK IPO - IDN Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Exist Email]";
                        content = "<center>********No Missing IPO ********</center><br /><br />";
                    }
                    else//ok
                    {
                        if (isExistIPO)//ok
                        {
                            subject = "HK IPO - IDN Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Exist Email]";
                            content = "<center>********Failed To Extract IPO RIC From FM********</center><br /><br />";//2.Failed to extract IPO RIC from FM
                        }
                        else//ok
                        {
                            subject = "HK IPO - IDN Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Exist Email]";
                            content = "<center>********No IPO Today********</center><br /><br />";//1.No IPO File Today
                        }
                    }
                }
                else//ok
                {
                    subject = "HK IPO - IDN Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[Empty Email]";
                    content = "<center>********FM attachment not found********</center><br />";
                }
            }
            else//ok
            {
                subject = "HK IPO - IDN Automation Report for" + DateTime.Now.ToString("dd-MMM-yyyy") + "[No  Email]";
                content = "<center>********FM Email Not Found********</center><br /><br />";
            }
            SendMail(service, subject, content, attacheFileList);
        }
        #endregion

        #region SendEmail(step:two)
        /// <summary>
        /// send email
        /// </summary>
        /// <param name="service">login</param>
        /// <param name="subject">subject</param>
        /// <param name="content">content</param>
        /// <param name="attacheFileList"></param>
        private void SendMail(ExchangeService service, string subject, string content, List<string> attacheFileList)
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
        #endregion
    }
}
