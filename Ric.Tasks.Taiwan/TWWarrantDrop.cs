using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using MSAD.Common.OfficeUtility;
using Ric.Core;
using Microsoft.Exchange.WebServices.Data;

namespace Ric.Tasks.Taiwan
{
    #region Configuration
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class TWWarrantDropConfig
    {
        [StoreInDB]
        [Category("EmailAccount")]
        [Description("Mail folder path,like: Inbox/XXXXX")]
        public string MailFolderPath { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [Description("Password")]
        public string Password { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [Description("Domain of the mail account, like: \"ten\"")]
        public string Domain { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [Description("Mail address, like: \"eti.XXXXXX@thomsonreuters.com\"")]
        public string MailAddress { get; set; }

        [StoreInDB]
        [Category("FilePath")]
        [Description("GeneratedFilePath")]
        public string FilePath { get; set; }

        [StoreInDB]
        [Category("EmailReceivedDate")]
        [Description("Email which day email you want to get, like: \"2013-05-06\" Default:Today's Email")]
        public string EmailDate { get; set; }
    }
    #endregion

    #region DefinitionEntity
    public class ForcedropTW
    {
        public string RIC { get; set; }
        public string REGION_CODE { get; set; }
        public string XQS_HEADEND { get; set; }
    }

    public class IAChg
    {
        public string PILC { get; set; }
        public string ISIN { get; set; }
        public string TAIWAN_CODE { get; set; }
    }

    #endregion

    class TWWarrantDrop : GeneratorBase
    {
        #region Declaration

        private DateTime startTime;
        private DateTime endTime;
        private ExchangeService service;
        private List<List<string>> listEXL = new List<List<string>>();
        private List<string> ricTAIW_EQLB_WNTandTAIW_INX_WNT = new List<string> { "stat", "ta", "va", "D", "0#" };
        private List<string> ricTAIW_CBBC = new List<string> { "stat", "ta", "D", "0#" };
        private List<string> ricOTCTWS_WNTandOTCTWS_INX_WNT = new List<string> { "stat", "ta", "va", "f", "D", "0#" };
        private List<ForcedropTW> listForcedropTW = new List<ForcedropTW>();
        public TWWarrantDropConfig configObj = null;
        private string accountName;//UC169XXX
        private string password;//********
        private string domain;//TEN
        private string mailAdress;//eti.XXXXXX@thomsonreuters.com
        private string mailFolder;//Inbox/XXXXX
        private string FilePath;
        private string emailDate;//2013-05-06ED
        private string pattern;
        DateTime dateTime = DateTime.Now.ToUniversalTime().AddHours(+8);

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as TWWarrantDropConfig;
            accountName = configObj.AccountName.Trim();
            password = configObj.Password;
            domain = configObj.Domain.Trim();
            mailAdress = configObj.MailAddress.Trim();
            mailFolder = configObj.MailFolderPath.Trim();
            FilePath = configObj.FilePath.Trim();
            emailDate = configObj.EmailDate.Trim();
        }
    #endregion

        protected override void Start()
        {
            FilllistEXLFromEmail();
            GenerateForcedropTWEntity();
            GenerateTxtFile();
        }

        #region FilllistEXLFromEmail
        private void FilllistEXLFromEmail()
        {
            try
            {
                service = EWSUtility.CreateService(new System.Net.NetworkCredential(accountName, password, domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
                EWSMailSearchQuery query;
                if (string.IsNullOrEmpty(emailDate))
                {
                    query = new EWSMailSearchQuery("", mailAdress, mailFolder, "TWWNT_DROP_IFFM Expired RICs Housekeeping Report for TWWNT_DROP", "", dateTime.AddHours(-dateTime.Hour).AddMinutes(-dateTime.Minute), dateTime);
                }
                else
                {
                    startTime = Convert.ToDateTime(emailDate);
                    endTime = startTime.AddDays(+1);
                    query = new EWSMailSearchQuery("", mailAdress, mailFolder, "TWWNT_DROP_IFFM Expired RICs Housekeeping Report for TWWNT_DROP", "", startTime, endTime);
                }
                List<EmailMessage> mailList = null;
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
                    string emailString = string.Empty;
                    EmailMessage mail = mailList[0];
                    mail.Load();
                    emailString = TWHelper.ClearHtmlTags(mail.Body.ToString());
                    string[] emailEXL = emailString.Replace("EXL", "~").Split('~');
                    foreach (string email in emailEXL.Where(email => email.Trim().StartsWith("(")))
                    {
                        List<string> listTmp;
                        string tmp = email.Substring(email.IndexOf("(") + 1, email.IndexOf(")") - 2).Trim();
                        switch (tmp)
                        {
                            case "TAIW_INX_WNT":
                                listTmp = new List<string> {tmp};
                                RegexPattern(email, listTmp);
                                break;
                            case "TAIW_EQLB_WNT":
                                listTmp = new List<string> {tmp};
                                RegexPattern(email, listTmp);
                                break;
                            case "TAIW_CBBC":
                                listTmp = new List<string> {tmp};
                                RegexPattern(email, listTmp);
                                break;
                            case "OTCTWS_WNT":
                                listTmp = new List<string> {tmp};
                                RegexPattern(email, listTmp);
                                break;
                            case "OTCTWS_INX_WNT":
                                listTmp = new List<string> {tmp};
                                RegexPattern(email, listTmp);
                                break;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No email of TWWNT_DROP_IFFM in Outlook!");
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Get Data from mail failed. Ex: " + ex.Message + "Try To Execute {EWSMailSearchQuery.SearchMail(service, query)} 5 times");
            }
        }
        #endregion

        #region FunctionRegexPattern
        /// <summary>
        /// Get Email string from analyst's computer exclude html tags but run on my computer include html tags!
        /// </summary>
        /// <param name="emailEXL"></param>
        /// <param name="listTmp"></param>
        private void RegexPattern(string emailEXL, List<string> listTmp)
        {
            pattern = @"(?<RIC>[A-Z0-9]{6}\.{1}[A-Z]{2,3})";
            Regex regex = new Regex(pattern);
            MatchCollection matches = regex.Matches(emailEXL);
            listTmp.AddRange(from Match match in matches 
                             select match.Groups["RIC"].Value);
            listEXL.Add(listTmp);
        }
        #endregion

        #region GenerateForcedropTWEntity
        private void GenerateForcedropTWEntity()
        {
            try
            {
                foreach (var listStr in listEXL)
                {
                    if (listStr.Contains("TAIW_CBBC"))
                    {
                        for (int i = 1; i < listStr.Count; i++)
                        {
                            string tmp = listStr[i];
                            foreach (var str in ricTAIW_CBBC)
                            {
                                var tw = new ForcedropTW();
                                if (str == "D" || str == "0#")
                                {
                                    tw.RIC = str + tmp;
                                }
                                else
                                {
                                    tw.RIC = tmp.Substring(0, tmp.IndexOf(".")) + str + tmp.Substring(tmp.IndexOf("."), tmp.Length - tmp.IndexOf("."));
                                }
                                tw.REGION_CODE = "HKG";
                                tw.XQS_HEADEND = "STQSK";
                                listForcedropTW.Add(tw);
                            }
                        }
                    }
                    else if (listStr.Contains("TAIW_EQLB_WNT") || listStr.Contains("TAIW_INX_WNT"))
                    {
                        for (int i = 1; i < listStr.Count; i++)
                        {
                            string tmp = listStr[i];
                            foreach (var str in ricTAIW_EQLB_WNTandTAIW_INX_WNT)
                            {
                                var tw = new ForcedropTW();
                                if (str == "D" || str == "0#")
                                {
                                    tw.RIC = str + tmp;
                                }
                                else
                                {
                                    tw.RIC = tmp.Substring(0, tmp.IndexOf(".")) + str + tmp.Substring(tmp.IndexOf("."), tmp.Length - tmp.IndexOf("."));
                                }
                                tw.REGION_CODE = "HKG";
                                tw.XQS_HEADEND = "STQSK";
                                listForcedropTW.Add(tw);
                            }
                        }
                    }
                    else
                    {
                        for (int i = 1; i < listStr.Count; i++)
                        {
                            string tmp = listStr[i];
                            foreach (var str in ricOTCTWS_WNTandOTCTWS_INX_WNT)
                            {
                                var tw = new ForcedropTW();
                                if (str == "D" || str == "0#")
                                {
                                    tw.RIC = str + tmp;
                                }
                                else
                                {
                                    tw.RIC = tmp.Substring(0, tmp.IndexOf(".")) + str + tmp.Substring(tmp.IndexOf("."), tmp.Length - tmp.IndexOf("."));
                                }
                                tw.REGION_CODE = "HKG";
                                tw.XQS_HEADEND = "STQSK";
                                listForcedropTW.Add(tw);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Fill ForcedropTWEntity failed. Ex: " + ex.Message);
            }
        }
        #endregion

        #region GenerateTxtFile
        private void GenerateTxtFile()
        {
            string fileName;
            if (string.IsNullOrEmpty(emailDate))
            {
                fileName = dateTime.Year + dateTime.Month + dateTime.Day + "Force_drop-TW.txt";
            }
            else
            {
                DateTime date = Convert.ToDateTime(emailDate);
                fileName = date.Year.ToString() + date.Month + date.Day + "Force_drop-TW.txt";
            }
            string forcedropTwTxtFilePath = Path.Combine(FilePath, fileName);
            string content = "RIC\tREGION_CODE\tXQS_HEADEND\r\n";
            foreach (ForcedropTW tw in listForcedropTW)
            {
                content += string.Format("{0}\t", tw.RIC);
                content += string.Format("{0}\t", tw.REGION_CODE);
                content += string.Format("{0}", tw.XQS_HEADEND);
                content += "\r\n";
            }
            try
            {
                File.WriteAllText(forcedropTwTxtFilePath, content);
                TaskResultList.Add(new TaskResultEntry("TWWarranrDrop", "txtFile", forcedropTwTxtFilePath));
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating ForcedropTW txt file. Ex: {0} .", ex.Message));
            }
        }
        #endregion
    }
}
