using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Office.Interop.Excel;
using MSAD.Common.OfficeUtility;
using Ric.Core;
using Ric.Util;
using ExcelApp = Ric.Util.ExcelApp;

namespace Ric.Tasks.Asia
{
    /// <summary>
    /// Configuration Class For the PreIPO Task
    /// </summary>
    [ConfigStoredInDB]
    public class ASIAPreIPOCheckerConfig
    {
        [StoreInDB]
        [Category("Credentials")]
        [Description("Domain\nE.g, ten")]
        public string Domain { get; set; }

        [StoreInDB]
        [Category("Credentials")]
        [Description("Username")]
        public string Username { get; set; }

        [StoreInDB]
        [Category("Credentials")]
        [Description("Password\nWrite your password on the first line")]
        public List<string> Password { get; set; }

        [StoreInDB]
        [Category("CSV")]
        [DisplayName("Result path")]
        [Description("Path of the generated CSV")]
        public string ResultsWorkbookPath { get; set; }

        [StoreInDB]
        [Category("Mail Account")]
        [DisplayName("Mail username")]
        [Description("Config for getting the email account")]
        public string MailUsername { get; set; }

        [StoreInDB]
        [Category("Mail Account")]
        [DisplayName("Mail password")]
        [Description("Config for getting the email account")]
        public string MailPassword { get; set; }

        [StoreInDB]
        [Category("Mail Account")]
        [DisplayName("Mail domain")]
        [Description("Config for getting the email account")]
        public string MailDomain { get; set; }

        [StoreInDB]
        [Category("Mail Account")]
        [Description("Config for getting the email account")]
        public string Email { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Recipients")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Recipients (CC)")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Mail signature")]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }
    }

    /// <summary>
    /// The class task is to parse a given CSV file with IPO informations, find new entries for
    ///  IPOs in a predefined list of asian countries, then create a new CSV file with those.
    ///  The File is then saved and send by email with the name of countries with new IPOs in the title.
    /// </summary>
    public class PreIPO : GeneratorBase
    {
        private ASIAPreIPOCheckerConfig _configObj;

        private string _csvUrl = String.Empty;
        private string _targetFilePath = String.Empty;
        private string _resultFilename = String.Empty;
        private List<string> updatedCountries = new List<string>();
        private List<List<string>> ipoList = new List<List<string>>();
        private readonly List<string> _xlsTitle = new List<string>{ "TRANSACTION_ID", "COMPANY_NAME", "CUSIP", "TICKER", "CREATE_STAMP"
            , "UPDATE_STAMP", "TRADE_DATE", "EXCHANGE", "OA_PERMID", "VALUE_IN_USD_MILLIONS", "VALUE_IN_HOST_MILLIONS"
            , "OFFER_CURRENCY", "NATION_OF_EXCHANGE", "SECURITY_TYPE", "HEAD_QUARTERS", "LONGNAME", "EXPECTED_ISSUE_DATE"
            , "IPO_DATE", "OFFER_PRICE", "TOTAL_SHRS_ALL_MKTS", "TRANSACTION_STATUS", "SHARES_OUT_BEFORE_OFFER"
            , "SHARES_OUT_AFTER_OFFER", "COMMENTS" };
        private readonly List<string> _countries = new List<string>{ "HongKong", "China", "Taiwan", "SouthKorea", "Indonesia"
            , "Japan", "Thailand", "Vietnam"};

        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        private ExchangeService _service;

        protected override void Start()
        {
            FindLastFile();
            DownloadCsv();
            ReadCSV();
            GenerateCSV();
            SendEmail();
        }

        protected override void Initialize()
        {
            base.Initialize();
            _configObj = Config as ASIAPreIPOCheckerConfig;
            _service = EWSUtility.CreateService(new NetworkCredential(_configObj.MailUsername, _configObj.MailPassword, _configObj.MailDomain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
        }

        /// <summary>
        /// Find newest CSV file
        /// (updated automatically every 3 months)
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        private void FindLastFile()
        {
            try
            {
                int nb = 1;
                DateTime compareDate = DateTime.Parse("01-Apr-2013");
                DateTime now = DateTime.Now;

                while (now.CompareTo(compareDate) > 0)
                {
                    compareDate = compareDate.AddMonths(3);
                    nb++;
                    if (nb == 5)
                    {
                        nb = 1;
                    }
                }
                _csvUrl = String.Format("https://theshare.thomsonreuters.com/sites/Business_Events_Team/Shared%20Documents/IPO%20Report/ipo_report_Generate_Q{0}_{1}.csv"
                    , nb, now.ToString("MMM_yyyy"));
            }
            catch (Exception ex)
            {
                string msg = "Cannot find last CSV file on sharepoint :" + ex;
                LogMessage(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Download CSV from Thomson Reuters website
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        private void DownloadCsv()
        {
            try
            {
                _targetFilePath = Path.GetDirectoryName(_configObj.ResultsWorkbookPath) + @"\" + DateTime.Now.ToOADate() + "_tmp.csv";

                var request = WebRequest.Create(_csvUrl) as HttpWebRequest;
                request.Timeout = 100000;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
                request.Method = "GET";
                request.Referer = "https://theshare.thomsonreuters.com/sites/Business_Events_Team/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FBusiness%5FEvents%5FTeam%2FShared%20Documents%2FIPO%20Report&FolderCTID=0x012000A8A0CECADD9AD44CB1EF5591A9D30E6F&View=%7b4E40A23B-5EFC-4CB0-8F04-87C242ED25CD%7d";
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";

                var cache = new CredentialCache
                {
                    {
                        new Uri(_csvUrl), "NTLM",
                        new NetworkCredential(_configObj.Username, _configObj.Password[0], _configObj.Domain)
                    }
                };
                request.Credentials = cache; 

                var httpResponse = (HttpWebResponse)request.GetResponse();
                Stream httpResponseStream = httpResponse.GetResponseStream();
                const int bufferSize = 1024;
                byte[] buffer = new byte[bufferSize];
                int bytesRead = 0;

                FileStream fileStream = File.Create(_targetFilePath);
                while ((bytesRead = httpResponseStream.Read(buffer, 0, bufferSize)) != 0)
                {
                    fileStream.Write(buffer, 0, bytesRead);
                }
                fileStream.Close();
                httpResponseStream.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot download CSV :" + ex;
                LogMessage(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Parsing CSV Finding new IPOs
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        private void ReadCSV()
        {
            var app = new ExcelApp(false, false);
            if (app.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            try
            {
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, _targetFilePath);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;

                double yesterday = DateTime.Now.AddDays(-2).ToOADate();
                double firstDate = DateTime.Parse("15-Apr-13").ToOADate();
                string country = String.Empty;

                for (int line = 2; worksheet.get_Range("C" + line, Type.Missing).Value2 != null; line++)
                {
                    var newIpo = new List<string>();
                    double createStamp = Convert.ToDouble(worksheet.Range["E" + line, Type.Missing].Value2.ToString());
                    double updateStamp = Convert.ToDouble(worksheet.Range["F" + line, Type.Missing].Value2.ToString());
                    country = worksheet.Range["M" + line, Type.Missing].Value2.ToString();

                    if (((updateStamp >= yesterday && createStamp >= firstDate) || createStamp >= yesterday)
                        && _countries.Contains(country))
                    {
                        if (!updatedCountries.Contains(country))
                        {
                            updatedCountries.Add(country);
                        }
                        for (int column = 0; column < _xlsTitle.Count; column++)
                        {
                            newIpo.Add(worksheet.Range[Alphabet.Substring(column, 1) + line, Type.Missing].Value2.ToString());
                        }
                        ipoList.Add(newIpo);
                    }
                }
                workbook.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot read CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
            }
            finally
            {
                app.Dispose();
            }
        }

        /// <summary>
        /// Creating new CSV with requested companies (new IPOs)
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        private void GenerateCSV()
        {
            var app = new ExcelApp(false, false);

            if (app.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !";
                LogMessage(msg, Logger.LogType.Error);
                return;
            }
            try
            {
                _resultFilename = _configObj.ResultsWorkbookPath.Replace(".csv", "_" + DateTime.Now.ToString("ddMMM_HH_mm_ss") + ".csv");
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, _resultFilename);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;

                for (int column = 0; column < _xlsTitle.Count; column++)
                {
                    worksheet.Cells[1, column + 1] = _xlsTitle[column];
                }
                for (int line = 2; line <= ipoList.Count + 1; line++)
                {
                    for (int column = 1; column <= _xlsTitle.Count; column++)
                    {
                        if (column == 5 || column == 6)
                        {
                            DateTime formatDate = DateTime.FromOADate(Convert.ToDouble(ipoList[line - 2][column - 1]));
                            worksheet.Cells[line, column] = formatDate.ToString("d");
                        }
                        else
                        {
                            worksheet.Cells[line, column] = ipoList[line - 2][column - 1];
                        }
                    }
                }
                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                AddResult("Result file", _resultFilename, "csv");
                workbook.SaveAs(workbook.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                LogMessage("Generated CSV file Successfully. Filepath is " + _resultFilename);
                workbook.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
            }
            finally
            {
                app.Dispose();
                File.Delete(_targetFilePath);
            }
        }

        /// <summary>
        /// Sending Email with the newly created CSV attached
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        private void SendEmail()
        {
            try
            {
                var mailCC= new List<string>();
                var attachement = new List<string>();
                string body = String.Empty;
                string title = String.Empty;
                if (updatedCountries.Count == 0)
                {
                    title = "No New Pre-IPO Records Today";
                }
                else
                {
                    var bodyBuilder = new StringBuilder();
                    title = "New Pre-IPO Records";
                    bodyBuilder.Append("Please find the new Pre-IPO records for ");
                    for (int count = 0; count < updatedCountries.Count; count++)
                    {
                        bodyBuilder.Append(updatedCountries[count]);
                        if (updatedCountries.Count >= 2 && count == updatedCountries.Count - 2)
                        {
                            bodyBuilder.Append(" & ");
                        }
                        else if (count != updatedCountries.Count - 1)
                        {
                            bodyBuilder.Append(", ");
                        }
                        else
                        {
                            bodyBuilder.Append(".");
                        }
                    }
                    bodyBuilder.Append("<p>");
                    foreach (string signatureLine in _configObj.MailSignature)
                    {
                        bodyBuilder.AppendFormat("{0}<br />", signatureLine);
                    }
                    bodyBuilder.Append("</p>");
                    body = bodyBuilder.ToString();
                    attachement.Add(_resultFilename);
                }
                if (_configObj.MailCC.Count > 1 || (_configObj.MailCC.Count == 1 && _configObj.MailCC[0] != ""))
                {
                    mailCC = _configObj.MailCC;
                }
                EWSUtility.CreateAndSendMail(_service, _configObj.MailTo, mailCC, new List<string>(), title, body, attachement);
            }
            catch (Exception ex)
            {
                string msg = "Cannot send mail :" + ex;
                LogMessage(msg, Logger.LogType.Error);
            }
        }
    }
}
