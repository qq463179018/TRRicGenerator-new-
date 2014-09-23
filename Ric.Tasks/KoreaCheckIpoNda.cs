using Microsoft.Exchange.WebServices.Data;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class KoreaCheckIpoNdaConfig
    {
        [StoreInDB]
        [Category("Proxy")]
        [DefaultValue("10.40.14.56")]
        [Description("Proxy IP address for assess to DSE.")]
        public string IP { get; set; }

        [StoreInDB]
        [Category("Proxy")]
        [DefaultValue("80")]
        [Description("Proxy port for assess to DSE.")]
        public string Port { get; set; }

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

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Account name")]
        [Description("Account name which used to login the outlook account, E.g.: \"UC165188\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [DisplayName("Download file path")]
        [Description("file path download from ftp, E.g.: \"\\10.35.16.40\\build\\private\\QC source file\"")]
        public string DownloadFilePath { get; set; }
    }

    public enum ReutersProductionType
    {
        IDN,
        NDA
    }

    public class KoreaCheckIpoData
    {
        public ReutersProductionType ProductionType { get; set; }

        public string TickerFm { get; set; }

        public string TickerProduct { get; set; }

        public bool IsTickerSame { get; set; }

        public string IsinFm { get; set; }

        public string IsinProduct { get; set; }

        public bool IsIsinSame { get; set; }

        public string IdnDisplayNameFm { get; set; }

        public string IdnDisplayNameProduct { get; set; }

        public bool IsIdnDisplayNameSame { get; set; }

        public string BcastRefFm { get; set; }

        public string BcastRefProduct { get; set; }

        public bool IsBcastRefSame { get; set; }

        public KoreaCheckIpoData()
        {
            IsBcastRefSame = IsIdnDisplayNameSame = IsTickerSame = IsIsinSame = true;

            TickerFm = TickerProduct = IdnDisplayNameFm = IdnDisplayNameProduct = BcastRefFm = BcastRefProduct
                = IsinFm = IsinProduct = string.Empty;
        }
    }

    /// <summary>
    /// Check FM with FTP.
    /// FM data from database(Tactical Automation).
    /// FTP files: 0184MMdd.M, 0673MMdd.M, 3073MMdd.M
    /// </summary>
    public class KoreaCheckIpoNda : GeneratorBase
    {
        private KoreaCheckIpoNdaConfig configObj;

        private string mFileFolder = string.Empty;

        private ExchangeService service;

        private Dictionary<string, DseFileRule> ruleDic;

        private List<KoreaEquityInfo> ipos;

        private Dictionary<string, DseFileInfo> dseDic;

        private Dictionary<string, ReutersIdnInfo> idnDic;

        private List<KoreaCheckIpoData> changedIpo;

        private List<KoreaCheckIpoData> missedIpo;

        private EmailAccountInfo emailAccount;

        private const string KoreaIpoQcFileName = "Korea IPO Check QC_{0}.xls";

        private bool isAfternoonTask;

        protected override void Initialize()
        {
            configObj = Config as KoreaCheckIpoNdaConfig;

            DateTime dtChina = TimeUtil.ConvertToChina(DateTime.UtcNow);
            emailAccount = EmailAccountManager.SelectEmailAccountByAccountName(configObj.AccountName.Trim());
            //AM
            if (dtChina.Hour >= 12)
            {
                isAfternoonTask = true;
            }

            string currentDate = DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US"));

            if (!isAfternoonTask)
            {
                currentDate = DateTime.Today.AddDays(-1).ToString("yyyyMMdd", new CultureInfo("en-US"));
            }

            mFileFolder = configObj.DownloadFilePath.Trim();

            InitializeMailAccount();

            InitializeDseRule();
        }

        private void Initialize1(KoreaCheckIpoNdaConfig obj, bool _isAfternoonTask)
        {
            configObj = obj;//Config as KoreaCheckIpoNdaConfig;

            isAfternoonTask = _isAfternoonTask;

            string currentDate = DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US"));

            if (!isAfternoonTask)
            {
                currentDate = DateTime.Today.AddDays(-1).ToString("yyyyMMdd", new CultureInfo("en-US"));
            }

            InitializeMailAccount();

            InitializeDseRule();
        }

        private void InitializeMailAccount()
        {
            service = MSAD.Common.OfficeUtility.EWSUtility.CreateService(new NetworkCredential(emailAccount.AccountName, emailAccount.Password, emailAccount.Domain), new Uri(@"https://apac.mail.erf.thomson.com/EWS/Exchange.asmx"));
        }

        protected override void Start()
        {

            //Sourcing
            GetIpo();

            if (ipos == null || ipos.Count == 0)
            {
                LogMessage("No IPOs.");
                return;
            }

            GetDseData();

            if (isAfternoonTask)
            {
                GetGatsData();
            }

            CompareIpo();

            if (changedIpo != null && changedIpo.Count > 0)
            {
                GenerateFile();
            }

            SendEmail();


        }

        public void StartJob(KoreaCheckIpoNdaConfig obj, bool _isAfternoonTask)
        {
            //Sourcing
            Initialize1(obj, _isAfternoonTask);

            GetIpo();

            GetDseData();

            CompareIpo();

            if (changedIpo != null && changedIpo.Count > 0)
            {
                GenerateFile();
            }

            SendEmail();
        }

        private void GetIpo()
        {
            ipos = new List<KoreaEquityInfo>();

            string currentDate = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));

            if (!isAfternoonTask)
            {
                currentDate = DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            }

            ipos = KoreaEquityManager.SelectEquityByDate(currentDate);

        }

        private void GetDseData()
        {
            DownloadFtpFiles();

            GetRecords();
        }

        private void DownloadFtpFiles()
        {
            string currentDate = String.Empty;
            string[] fileStartArr;
            if (isAfternoonTask)
            {
                fileStartArr = new[] { "0184", "0673", "3073" };
            }
            else
            {
                //if we should download EM01 file. use this part.
                fileStartArr = new[] { "EM01" };

                //After the EM01 file stored in a common path, use this part.
                //return;
            }

            WebClient request = new WebClient();

            if (!string.IsNullOrEmpty(configObj.IP) && !string.IsNullOrEmpty(configObj.Port))
            {
                WebProxy proxy = new WebProxy(configObj.IP, Convert.ToInt32(configObj.Port));
                request.Proxy = proxy;
            }

            request.Credentials = new NetworkCredential("ASIA2", "ASIA2");

            foreach (string fileStart in fileStartArr)
            {
                currentDate = GetCurrrentFileDate(fileStart);

                string fileName = string.Format("{0}{1}.M", fileStart, currentDate);

                string mfilePath = Path.Combine(mFileFolder, fileName);

                if (File.Exists(mfilePath))
                {
                    continue;
                }

                string ftpfullpath = @"ftp://ASIA2:ASIA2@ds1.rds.reuters.com/" + fileName;

                try
                {
                    request.DownloadFile(ftpfullpath, mfilePath);

                    LogMessage(string.Format("Download FTP File {0}... OK!", fileName));
                }
                catch (Exception ex)
                {
                    string msg = string.Format("Can not download file: {0} from FTP. Response:{1}", fileName, ex.Message);
                    LogMessage(msg, Logger.LogType.Error);
                }
            }

        }

        private string GetCurrrentFileDate(string fileStart)
        {
            string inputDate = DateTime.Today.ToString("MMdd");//configObj.Date;
            DateTime dateToUse = DateTime.Now;
            if (fileStart == "EM01")
            {
                int daysToAdd = -1;
                dateToUse = DateTime.ParseExact(inputDate, "MMdd", CultureInfo.InvariantCulture);
                //if (dateToUse.DayOfWeek == DayOfWeek.Monday)
                //{
                //    daysToAdd = -1;
                //}
                return dateToUse.AddDays(daysToAdd).ToString("MMdd");
            }

            return inputDate;

        }

        private void GetRecords()
        {
            //List<DseFileInfo> xeRecord = new List<DseFileInfo>();
            dseDic = new Dictionary<string, DseFileInfo>();

            string currentDate = string.Empty;

            string[] fileStartArr = { "0184", "0673", "3073", "EM01" };


            foreach (string fileStart in fileStartArr)
            {
                currentDate = GetCurrrentFileDate(fileStart);

                string fileName = string.Format("{0}{1}.M", fileStart, currentDate);

                string mfilePath = Path.Combine(mFileFolder, fileName);

                if (!File.Exists(mfilePath))
                {
                    continue;
                }

                using (StreamReader sr = new StreamReader(mfilePath))
                {
                    string tmp = null;
                    while ((tmp = sr.ReadLine()) != null)
                    {
                        if (!tmp.StartsWith("XE"))
                        {
                            continue;
                        }
                        string ric = tmp.Substring(0, tmp.IndexOf(' ')).Replace("XE", "");
                        if (!(ric.EndsWith("KS") || ric.EndsWith("KQ") || ric.EndsWith("KN")))
                        {
                            continue;
                        }

                        DseFileInfo info = ParseDseRecord(tmp);

                        if (info != null && !dseDic.ContainsKey(info.Ticker))
                        {
                            dseDic.Add(info.Ticker, info);
                        }

                        string msg = string.Format("Get 1 record from file: {0}. RIC:{1}", fileName, ric);
                        LogMessage(msg);

                    }
                }

            }

        }

        private DseFileInfo ParseDseRecord(string record)
        {
            DseFileInfo parsedRecord = new DseFileInfo();

            PropertyInfo[] properties = parsedRecord.GetType().GetProperties();

            foreach (PropertyInfo p in properties)
            {
                if (ruleDic.ContainsKey(p.Name))
                {
                    DseFileRule dseRule = ruleDic[p.Name] as DseFileRule;
                    string value = dseRule.ParseField(record);
                    SetPropertyValue(parsedRecord.GetType(), p.Name, value, parsedRecord);
                }
            }

            return parsedRecord;
        }

        private void CompareIpo()
        {
            changedIpo = new List<KoreaCheckIpoData>();
            missedIpo = new List<KoreaCheckIpoData>();

            foreach (var ipo in ipos)
            {
                CompareIpoNda(ipo);
                CompareIpoIdn(ipo);
            }
        }

        private void CompareIpoNda(KoreaEquityInfo ipo)
        {
            if (dseDic == null)
            {
                return;
            }

            if (dseDic.ContainsKey(ipo.Ticker))
            {
                DseFileInfo dseInfo = dseDic[ipo.Ticker];

                string[] securityNames = dseInfo.SecurityDescription.Split(' ');

                string type = string.Empty;

                if (securityNames.Length > 1)
                {
                    type = securityNames[securityNames.Length - 1];
                    dseInfo.SecurityDescription = dseInfo.SecurityDescription.Replace(type, "").Trim();
                }

                if ((dseInfo.ISIN == ipo.ISIN) && (dseInfo.SecurityDescription == ipo.IDNDisplayName))
                {
                    return;
                }
                KoreaCheckIpoData changeData = new KoreaCheckIpoData
                {
                    ProductionType = ReutersProductionType.NDA,
                    TickerFm = ipo.Ticker,
                    TickerProduct = dseInfo.Ticker,
                    IsTickerSame = true,
                    IsinFm = ipo.ISIN,
                    IsinProduct = dseInfo.ISIN,
                    IdnDisplayNameFm = ipo.IDNDisplayName,
                    IdnDisplayNameProduct = dseInfo.SecurityDescription
                };

                if (dseInfo.SecurityDescription != ipo.IDNDisplayName)
                {
                    //Mark change
                    changeData.IsIdnDisplayNameSame = false;
                }
                if (dseInfo.ISIN != ipo.ISIN)
                {
                    changeData.IsIsinSame = false;
                }

                changedIpo.Add(changeData);
            }
            else
            {
                // Mark Ticker missed.
                KoreaCheckIpoData missData = new KoreaCheckIpoData
                {
                    ProductionType = ReutersProductionType.NDA,
                    TickerFm = ipo.Ticker,
                    TickerProduct = string.Empty,
                    IsinFm = ipo.ISIN,
                    IsinProduct = string.Empty,
                    IdnDisplayNameFm = ipo.IDNDisplayName,
                    IdnDisplayNameProduct = string.Empty,
                    IsTickerSame = false,
                    IsIdnDisplayNameSame = false,
                    IsIsinSame = false
                };

                missedIpo.Add(missData);
            }
        }

        private void CompareIpoIdn(KoreaEquityInfo ipo)
        {
            if (idnDic == null)
            {
                return;
            }

            if (idnDic.ContainsKey(ipo.Ticker))
            {
                ReutersIdnInfo idnInfo = idnDic[ipo.Ticker];

                if ((idnInfo.BcastRef == ipo.BcastRef) && (idnInfo.DsplyName == ipo.IDNDisplayName) && (idnInfo.OffclCode == ipo.Ticker))
                {
                    return;
                }

                else
                {

                    KoreaCheckIpoData changeData = new KoreaCheckIpoData
                    {
                        ProductionType = ReutersProductionType.IDN,
                        TickerFm = ipo.Ticker,
                        TickerProduct = idnInfo.OffclCode,
                        BcastRefFm = ipo.BcastRef,
                        BcastRefProduct = idnInfo.BcastRef,
                        IdnDisplayNameFm = ipo.IDNDisplayName,
                        IdnDisplayNameProduct = idnInfo.DsplyName
                    };

                    if (idnInfo.OffclCode != ipo.Ticker)
                    {
                        changeData.IsTickerSame = false;
                    }

                    if (idnInfo.DsplyName != ipo.IDNDisplayName)
                    {
                        //Mark change
                        changeData.IsIdnDisplayNameSame = false;
                    }
                    if (idnInfo.BcastRef != ipo.BcastRef)
                    {
                        changeData.IsIsinSame = false;
                    }

                    changedIpo.Add(changeData);
                }
            }
            else
            {
                // Mark Ticker missed.
                KoreaCheckIpoData missData = new KoreaCheckIpoData
                {
                    ProductionType = ReutersProductionType.IDN,
                    TickerFm = ipo.Ticker,
                    TickerProduct = string.Empty,
                    IdnDisplayNameFm = ipo.IDNDisplayName,
                    IdnDisplayNameProduct = string.Empty,
                    BcastRefFm = ipo.BcastRef,
                    BcastRefProduct = string.Empty,
                    IsTickerSame = false,
                    IsIdnDisplayNameSame = false,
                    IsBcastRefSame = false
                };

                missedIpo.Add(missData);
            }
        }

        private void GenerateFile()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                excelApp.Dispose();
                throw new Exception(msg);
            }

            try
            {
                string fileName = string.Format(KoreaIpoQcFileName, DateTime.Today.ToString("yyyy-MM-dd"));

                string filePath = Path.Combine(GetOutputFilePath(), fileName);  //"C:\\Korea_Auto\\Equity_Warrant\\Name_Change\\" + filename;

                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                    throw new Exception(msg);
                }

                GenerateExcelFileTitle(wSheet);
                int row = 2;
                foreach (var item in changedIpo)
                {
                    string productionType = item.ProductionType.ToString();
                    wSheet.Cells[row, 1] = productionType;
                    ((Range)wSheet.Cells[row, 2]).NumberFormat = "@";
                    wSheet.Cells[row, 2] = item.TickerFm;
                    ((Range)wSheet.Cells[row, 3]).NumberFormat = "@";
                    wSheet.Cells[row, 3] = item.TickerProduct;

                    wSheet.Cells[row, 4] = item.IsTickerSame.ToString();
                    if (!item.IsTickerSame)
                    {
                        ((Range)wSheet.Cells[row, 4]).Interior.Color = Color.Yellow;
                    }

                    wSheet.Cells[row, 5] = item.IsinFm;
                    wSheet.Cells[row, 6] = item.IsinProduct;
                    if (string.IsNullOrEmpty(item.IsinFm))
                    {
                        wSheet.Cells[row, 7] = "";
                    }
                    else
                    {
                        wSheet.Cells[row, 7] = item.IsIsinSame.ToString();
                    }
                    if (!item.IsIsinSame)
                    {
                        ((Range)wSheet.Cells[row, 7]).Interior.Color = Color.Yellow;
                    }

                    wSheet.Cells[row, 8] = item.IdnDisplayNameFm;

                    wSheet.Cells[row, 9] = item.IdnDisplayNameProduct;

                    wSheet.Cells[row, 10] = item.IsIdnDisplayNameSame.ToString();
                    if (!item.IsIdnDisplayNameSame)
                    {
                        ((Range)wSheet.Cells[row, 10]).Interior.Color = Color.Yellow;
                    }

                    ((Range)wSheet.Cells[row, 11]).NumberFormat = "@";
                    wSheet.Cells[row, 11] = item.BcastRefFm;
                    ((Range)wSheet.Cells[row, 12]).NumberFormat = "@";
                    wSheet.Cells[row, 12] = item.BcastRefProduct;
                    if (string.IsNullOrEmpty(item.BcastRefFm))
                    {
                        wSheet.Cells[row, 13] = "";
                    }
                    else
                    {
                        wSheet.Cells[row, 13] = item.IsBcastRefSame.ToString();
                    }
                    if (!item.IsBcastRefSame)
                    {
                        ((Range)wSheet.Cells[row, 13]).Interior.Color = Color.Yellow;
                    }

                    row++;
                }

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();

                AddResult("Compared file", filePath, "file");
                Logger.Log("Generate FM file. Filepath is " + filePath);
            }
            catch (Exception ex)
            {
                string msg = "Error found in GenerateNameChangeFMFile()   : \r\n" + ex.ToString();
                LogMessage(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }

        }

        private void GenerateExcelFileTitle(Worksheet wSheet)
        {
            wSheet.Cells[1, 1] = "Production Type";
            wSheet.Cells[1, 2] = "Ticker in FM";
            wSheet.Cells[1, 3] = "Ticker in Production";
            wSheet.Cells[1, 4] = "Compare Result";
            wSheet.Cells[1, 5] = "ISIN in FM";
            wSheet.Cells[1, 6] = "ISIN in Production";
            wSheet.Cells[1, 7] = "Compare Result";
            wSheet.Cells[1, 8] = "IDN Display Name in FM";
            wSheet.Cells[1, 9] = "IDN Display Name in Production";
            wSheet.Cells[1, 10] = "Compare Result";
            wSheet.Cells[1, 11] = "BCAST_REF in FM";
            wSheet.Cells[1, 12] = "BCAST_REF in Production";
            wSheet.Cells[1, 13] = "Compare Result";

        }

        private void SendEmail()
        {

            string subject = "Korea IPO QC Check Automation Report for " + DateTime.Now.ToString("dd-MMM-yyyy");
            string mailBody = string.Empty;

            List<string> attachFileList = new List<string>();
            if (changedIpo.Count == 0 && missedIpo.Count == 0)
            {
                subject += " [No Missed and Changed Ticker]";
                mailBody = "******** All the IPO Tickers existed in production(NDA and IDN) and had no change ********";
            }
            else if (changedIpo.Count > 0 && missedIpo.Count > 0)
            {
                string mTickersInNda = string.Join(",", (from p in missedIpo where p.ProductionType.Equals(ReutersProductionType.NDA) select p.TickerFm).ToArray());

                string mTickersInIdn = string.Join(",", (from p in missedIpo where p.ProductionType.Equals(ReutersProductionType.IDN) select p.TickerFm).ToArray());

                string cTickersInNda = string.Join(",", (from p in changedIpo where p.ProductionType.Equals(ReutersProductionType.NDA) select p.TickerFm).ToArray());

                string cTickersInIdn = string.Join(",", (from p in changedIpo where p.ProductionType.Equals(ReutersProductionType.IDN) select p.TickerFm).ToArray());

                subject += " [Missed Ticker and Changed Field]";
                mailBody = "Missed Ticker In NDA: " + (string.IsNullOrEmpty(mTickersInNda) ? "None" : mTickersInNda) + "<br />"
                         + "Changed Ticker In NDA:" + (string.IsNullOrEmpty(cTickersInNda) ? "None" : cTickersInNda) + "<br />"
                         + "Missed Ticker In IDN: " + (string.IsNullOrEmpty(mTickersInIdn) ? "None" : mTickersInIdn) + "<br />"
                         + "Changed Ticker In IDN:" + (string.IsNullOrEmpty(cTickersInIdn) ? "None" : cTickersInIdn) + "<br />"
                         + "For the changed fields, please see in the attachment.";

                string fileName = string.Format(KoreaIpoQcFileName, DateTime.Today.ToString("yyyy-MM-dd"));
                string filePath = Path.Combine(GetOutputFilePath(), fileName);
                attachFileList.Add(filePath);
            }
            else if (changedIpo.Count > 0)
            {
                subject += " [Changed Field]";

                string cTickersInNda = string.Join(",", (from p in changedIpo where p.ProductionType.Equals(ReutersProductionType.NDA) select p.TickerFm).ToArray());

                string cTickersInIdn = string.Join(",", (from p in changedIpo where p.ProductionType.Equals(ReutersProductionType.IDN) select p.TickerFm).ToArray());

                mailBody = "Changed Ticker In NDA:" + (string.IsNullOrEmpty(cTickersInNda) ? "None" : cTickersInNda) + "<br />"
                         + "Changed Ticker In IDN:" + (string.IsNullOrEmpty(cTickersInNda) ? "None" : cTickersInIdn) + "<br />";

                string fileName = string.Format(KoreaIpoQcFileName, DateTime.Today.ToString("yyyy-MM-dd"));
                string filePath = Path.Combine(GetOutputFilePath(), fileName);
                attachFileList.Add(filePath);
            }
            else if (missedIpo.Count > 0)
            {
                subject += " [Missed Ticker]";
                string mTickersInNda = string.Join(",", (from p in missedIpo where p.ProductionType.Equals(ReutersProductionType.NDA) select p.TickerFm).ToArray());

                string mTickersInIdn = string.Join(",", (from p in missedIpo where p.ProductionType.Equals(ReutersProductionType.IDN) select p.TickerFm).ToArray());

                mailBody = "Missed Ticker In NDA: " + (string.IsNullOrEmpty(mTickersInNda) ? "None" : mTickersInNda) + "<br />"
                         + "Missed Ticker In IDN: " + (string.IsNullOrEmpty(mTickersInNda) ? "None" : mTickersInIdn) + "<br />";
            }

            SendMail(service, subject, mailBody, attachFileList);
        }

        #region SendMail
        /// <summary>
        /// SendMail
        /// </summary>
        /// <param name="service">Login Email</param>
        /// <param name="subject">subject</param>
        /// <param name="content">Body</param>
        /// <param name="attacheFileList">Attachements</param>
        private void SendMail(ExchangeService service, string subject, string content, List<string> attacheFileList)
        {
            StringBuilder bodyBuilder = new StringBuilder();
            bodyBuilder.Append(content);
            bodyBuilder.Append("<p></p>");
            bodyBuilder.Append("<p>");
            foreach (string signatureLine in configObj.MailSignature)
            {
                bodyBuilder.AppendFormat("{0}<br />", signatureLine);
            }
            bodyBuilder.Append("</p>");
            content = bodyBuilder.ToString();
            List<string> mailCc = null;
            if (configObj.MailCC.Count > 1 || (configObj.MailCC.Count == 1 && configObj.MailCC[0] != ""))
            {
                mailCc = configObj.MailCC;
            }
            MSAD.Common.OfficeUtility.EWSUtility.CreateAndSendMail(service, configObj.MailTo, mailCc, new List<string>(), subject, content, attacheFileList);
        }
        #endregion

        private void InitializeDseRule()
        {
            string dseFileRulePath = @".\Config\Asia\DseFileRules.xml";

            List<DseFileRule> dseRules = ConfigUtil.ReadConfig(dseFileRulePath, typeof(List<DseFileRule>)) as List<DseFileRule>;

            ruleDic = new Dictionary<string, DseFileRule>();

            ruleDic = dseRules.ToDictionary(e => e.PropertyName, StringComparer.Ordinal);

        }

        private void SetPropertyValue(Type type, string property, string value, object obj)
        {
            string propertyTypeName = type.GetProperty(property).PropertyType.Name;
            object valueToSet = null;
            if (propertyTypeName.ToLower().Contains("string"))
            {
                valueToSet = value;
            }
            type.GetProperty(property).SetValue(obj, valueToSet, null);
        }


        #region IDN Check

        private void GetGatsData()
        {
            idnDic = new Dictionary<string, ReutersIdnInfo>();

            string rics = string.Join(",", (from p in ipos select p.RIC).ToArray());
            string fids = "OFFCL_CODE,BCAST_REF,DSPLY_NAME";

            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(rics, fids);

            if (string.IsNullOrEmpty(response))
            {
                //No result from gats.

                return;
            }


            foreach (var ipo in ipos)
            {

                if (!response.Contains(ipo.RIC))
                {
                    continue;
                }

                ReutersIdnInfo idn = new ReutersIdnInfo();

                string offclPattern = string.Format("{0} +OFFCL_CODE +(?<OfficalCode>.*?)\r\n", ipo.RIC);
                string bcastPattern = string.Format("{0} +BCAST_REF +(?<BcastRef>.*?)\r\n", ipo.RIC);
                string diaplayPattern = string.Format("{0} +DSPLY_NAME +(?<DisplayName>.*?)\r\n", ipo.RIC);

                Regex r = new Regex(offclPattern);
                Match m = r.Match(response);
                if (m.Success)
                {
                    idn.OffclCode = m.Groups["OfficalCode"].Value.Trim();
                }

                r = new Regex(bcastPattern);
                m = r.Match(response);
                if (m.Success)
                {
                    idn.BcastRef = m.Groups["BcastRef"].Value.Trim();
                }

                r = new Regex(diaplayPattern);
                m = r.Match(response);
                if (m.Success)
                {
                    idn.DsplyName = m.Groups["DisplayName"].Value.Trim();
                }

                idnDic.Add(ipo.Ticker, idn);
            }
        }


        #endregion
    }

    public class DseFileRule
    {
        public string Field { get; set; }
        public string Description { get; set; }
        public string MaxLength { get; set; }
        public string Decimal { get; set; }
        public string PositionFrom { get; set; }
        public string PositionTo { get; set; }
        public string PropertyName { get; set; }

        public string ParseField(string record)
        {
            int startPosition = 0;
            int endPosition = 0;
            int maxLength = 0;

            startPosition = Convert.ToInt32(this.PositionFrom);
            endPosition = Convert.ToInt32(this.PositionTo);
            maxLength = Convert.ToInt32(this.MaxLength);

            return record.Substring(startPosition - 1, maxLength).Trim();//record.Substring(startPosition - 1, endPosition - startPosition + 1).Trim();
        }
    }

    public class DseFileInfo
    {
        public string RecordType { get; set; }
        public string RIC { get; set; }
        public string SecurityDescription { get; set; }
        public string CUSIP { get; set; }
        public string SEDOL { get; set; }
        public string CommonCode { get; set; }
        public string ISIN { get; set; }
        public string IssueClassification { get; set; }
        public string ExchangeCode { get; set; }
        public string CurrencyCode { get; set; }
        public string TradingStatus { get; set; }
        public string CompanyName { get; set; }
        public string MSCIIndustrialClassificationCode { get; set; }
        public string IssuerORGID { get; set; }
        public string CompanyShortName { get; set; }
        public string CompanyLegalDomicile { get; set; }
        public string AustraliaCode { get; set; }
        public string AustriaCode { get; set; }
        public string BelgiumCode { get; set; }
        public string FranceCode { get; set; }
        public string Wertpapier { get; set; }
        public string SICC { get; set; }
        public string NetherlandsCode { get; set; }
        public string Nolongerinused { get; set; }
        public string SaoPauloCode { get; set; }
        public string Valoren { get; set; }
        public string TaiwanCode { get; set; }
        public string HongKongCode { get; set; }
        public string MalaysiaCode { get; set; }
        public string SingaporeCode { get; set; }
        public string PILC { get; set; }
        public string FinsburyCompanyCode { get; set; }
        public string ISOCFICode { get; set; }
        public string ReutersEditorialRIC { get; set; }
        public string GICSIndustryCode { get; set; }
        public string NoLongerused1 { get; set; }
        public string NoLongerused2 { get; set; }
        public string PEcode { get; set; }
        public string QuotronSymbol { get; set; }
        public string BelgianCode { get; set; }
        public string MarketIdentifierCode { get; set; }
        public string OPOL { get; set; }
        public string AssetCategory { get; set; }
        public string PrimaryListedRIC { get; set; }
        public string Ticker { get; set; }
        public string SecurityLongDescription { get; set; }
        public string TRBCIndustryCode { get; set; }
        public string MiFIDIndicator { get; set; }
        public string CFICode { get; set; }
        public string PlaceofListingFlag { get; set; }
        public string PrimaryReferenceMarketQuote { get; set; }
        public string PrimaryExecutionVenue { get; set; }
        public string MarketSegmentName { get; set; }
        public string MarketMIC { get; set; }
        public string CRAORGID { get; set; }
        public string CESREEARegulated { get; set; }
        public string CESRMostRelevantMarket { get; set; }
        public string CESRAveargeDailyTurnover { get; set; }
        public string CESRAverageDailyTurnoverCurrencyCode { get; set; }
        public string CRAAverageDailyTurnover { get; set; }
        public string CRAAverageDailyTurnoverCurrencyCode { get; set; }
        public string CESRAverageValueofOrdersExecuted { get; set; }
        public string CESRAverageValueofOrdersExecutedCurrencyCode { get; set; }
        public string CRAAverageValueofOrdersExecuted { get; set; }
        public string CRAAverageValueofOrdersExecutedCurrencyCode { get; set; }
        public string CESRFreeFloat { get; set; }
        public string CESRFreeFloatcurrencyCode { get; set; }
        public string CRAFreeFloat { get; set; }
        public string CRAFreeFloatcurrencyCode { get; set; }
        public string CESRStandardMarketSize { get; set; }
        public string CESRStandardMarketSizeCurrencyCode { get; set; }
        public string NetherlandCode { get; set; }
        public string RoundLotSize { get; set; }
        public string ThomsonReutersClassificationScheme { get; set; }
        public string SuspendFlag { get; set; }
        public string DepositoryAssetUnderlying { get; set; }
        public string ILXCode { get; set; }
        public string TradingSymbol { get; set; }
        public string WhenIssuedFlag { get; set; }
        public string RegisteredFlag144A { get; set; }
        public string AssetRatioFor { get; set; }
        public string AssetRatioAgainst { get; set; }
        public string EuronextTradingGroup { get; set; }
        public string KazakhstanCode { get; set; }
        public string CountryofIncorporation { get; set; }
        public string QuoteID { get; set; }
        public string AssetID { get; set; }
        public string StampDutyFlag { get; set; }
        public string FirstTradeDate { get; set; }
        public string INAVRIC { get; set; }
        public string CINCode { get; set; }
        public string CountryofTaxation { get; set; }
    }

    public class ReutersIdnInfo
    {
        public string BcastRef { get; set; }

        public string OffclCode { get; set; }

        public string DsplyName { get; set; }
    }
}
