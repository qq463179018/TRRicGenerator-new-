using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Info;
using MySql.Data.MySqlClient;
using Ric.Db.Manager;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using Ric.Db.Config;
using System.Text.RegularExpressions;
using pdftron;
using pdftron.PDF;
using Ric.Core;
using Ric.Util;
using System.Globalization;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace Ric.Tasks.Corax
{
    [ConfigStoredInDB]
    public class CoraxProfitDistributionPDFConfig
    {
        [StoreInDB]
        [Category("OutputPath")]
        [Description("path of generate result file")]
        public string OutputPath { get; set; }

        [StoreInDB]
        [Category("QueryType")]
        [DefaultValue("yaqiong.wang")]
        [Description("UserName")]
        public string AssignedTo { get; set; }

        [Category("QueryType")]
        [Description("Status of record")]
        public Status Status { get; set; }

        [Category("GenerateBulkFile")]
        [Description("BOD_notice_In_Scope.csv ?")]
        public BODNoticeInScope BODNoticeInScopeFile { get; set; }

        [Category("GenerateBulkFile")]
        [Description("AGM_EGM_notice_for_NDA_and_CAM(.csv and .xls)")]
        public AGMEGMNoticeForNDAAndCAM AGMEGMNoticeForNDAAndCAMFile { get; set; }

        [Category("GenerateBulkFile")]
        [Description("AGM_EGM_notice_for_CAM.xls")]
        public AGMEGMNoticeForCAM AGMEGMNoticeForCAMFile { get; set; }

        [Category("GenerateBulkFile")]
        [Description("AGM_EGM_result_In_Scope.csv")]
        public AGMEGMResultInScope AGMEGMResultInScopeFile { get; set; }

        [Category("GenerateBulkFile")]
        [Description("Implement_notice.csv")]
        public ImplementNotice ImplementNoticeFile { get; set; }

        [Category("SetOOSToAce")]
        [Description("if set ScopeType OOS ACE DB")]
        public UpdateOOS UpdateOOS { get; set; }
    }

    public enum BODNoticeInScope : int { Yes, No };

    public enum AGMEGMNoticeForNDAAndCAM : int { Yes, No };

    public enum AGMEGMNoticeForCAM : int { Yes, No };

    public enum AGMEGMResultInScope : int { Yes, No };

    public enum ImplementNotice : int { Yes, No };

    public enum UpdateOOS : int { Yes, No };

    public enum Status : int { Assigned, Completed, Pending, OutOfScope }//default from 0,1 ...

    public class CoraxProfitDistributionPDF : GeneratorBase
    {
        public static CoraxProfitDistributionPDFConfig configObj = null;
        private string downloadPath = string.Empty;
        private string typePath = string.Empty;
        private string resultPath = string.Empty;
        private ExtendDataContext extendContext = null;
        private DealsContext dealsContext = null;
        private List<CHNProcessItem> historicalAnnouncements = new List<CHNProcessItem>();
        private string sourcingEngine = string.Empty;
        private string scopeType = string.Empty;
        private List<string> listNewDownloadPdf = new List<string>();//save download pdf in this time
        private List<string> listDownloadPdfError = new List<string>();//save download pdf link when generate download error 
        private Dictionary<string, string> dicTitlePattern = new Dictionary<string, string>();
        private string status = string.Empty;
        private string assignedTo = string.Empty;
        private string classAGMNotice = string.Empty;
        private List<string> listAGMNotice = new List<string>();
        private string classAGMResult = string.Empty;
        private List<string> listAGMResult = new List<string>();
        private string classBODNotice = string.Empty;
        private List<string> listBODNotice = new List<string>();
        private string classImplementNotice = string.Empty;
        private List<string> listImplementNotice = new List<string>();
        private string classOthers = string.Empty;
        private List<string> listOthers = new List<string>();
        private List<string> listOOSFolder = null;
        private List<int> listOOSId = null;
        private Dictionary<string, string> dicBShare = new Dictionary<string, string>();//table tickers on ace 
        string eventFile1 = string.Empty;

        protected override void Initialize()
        {
            configObj = Config as CoraxProfitDistributionPDFConfig;
            downloadPath = Path.Combine(configObj.OutputPath.Trim(), @"Download");
            CreateFolder(downloadPath);
            typePath = Path.Combine(configObj.OutputPath.Trim(), @"Classification");
            resultPath = Path.Combine(configObj.OutputPath.Trim(), @"Result");
            sourcingEngine = "CORAX CHN";
            scopeType = "InScope";
            assignedTo = configObj.AssignedTo.Trim();

            if (configObj.Status.Equals(Status.Assigned))
                status = "Assigned";
            else if (configObj.Status.Equals(Status.Completed))
                status = "Completed";
            else if (configObj.Status.Equals(Status.Pending))
                status = "Pending";
            else if (configObj.Status.Equals(Status.OutOfScope))
                status = "OutOfScope";
            else
                MessageBox.Show("please check configuration value.");

            extendContext = new ExtendDataContext(new MySqlConnection(AceDbConfig.DataSourceStringDeals));
            dealsContext = new DealsContext(new MySqlConnection(AceDbConfig.DataSourceStringDeals));
            dicTitlePattern.Add("董事会.*决议的?公告", "BOD_notice");
            dicTitlePattern.Add("股东大会.*通知.*Meeting", "AGM_EGM_notice");
            dicTitlePattern.Add("股东大会.*决议的?公告", "AGM_EGM_result");
            dicTitlePattern.Add("实施的?公告", "Implement_notice");
            classAGMNotice = Path.Combine(typePath, "AGM_EGM_notice");
            classAGMResult = Path.Combine(typePath, "AGM_EGM_result");
            classBODNotice = Path.Combine(typePath, "BOD_notice");
            classImplementNotice = Path.Combine(typePath, "Implement_notice");
            classOthers = Path.Combine(typePath, "Others");

            //list folder need to set outof scope to ace 
            listOOSFolder = new List<string>() { 
                Path.Combine(classAGMResult,"AGM_EGM_result_OO_Scope"),
                Path.Combine(classBODNotice,"BOD_notice_OO_Scope")
            };
        }

        protected override void Start()
        {
            //string path = @"F:\Task\ETI-431 Corax Profit Distribution PDF automation\Classification\BOD_notice\BOD_notice_In_Scope";
            //List<string> listPdf = Directory.GetFiles(path, "*.PDF").ToList();
            //string annDate1 = "30-04-2014";
            //string meetingDate1 = "30-04-2014";
            //list.Add(GetCDIRate(item));                    //CDI rate                   pdf
            //list.Add(GetSDIRate(item));                    //SDI rate                   pdf
            //list.Add(GetSISRatio(item));                   //SIS ratio                  pdf
            //List<string> listTest = new List<string>();

            //foreach (var item in listPdf)
            //{
            //    listTest.Add(GetCDIRate(item));
            //}

            LogMessage("start to download pdf ... ");
            FirstJopDownloadAndClassify();
            LogMessage("start to update to ace ... ");
            SecondJopUpdateToDB();
            LogMessage("start to generate bulk files for classification pdf... ");
            ThirdJopGenerateBulkFile();

            LogMessage("start to generate bulk file for result file ... ");
            FourJopGenerateBulkFile();
            LogMessage("all finished ... ");
        }

        private void FourJopGenerateBulkFile()
        {
            if ((eventFile1 + "").Trim().Length == 0 || !File.Exists(eventFile1))
            {
                Logger.Log("Meeting_yyyymmdd.xls is not exist.");
                return;
            }

        }

        #region [Start]
        private void ThirdJopGenerateBulkFile()
        {
            string path = Path.Combine(resultPath, DateTime.Now.ToString("dd-MM-yyyy"));
            string filePathNDAAndCAMXls = string.Empty;
            string filePathCAMXls = string.Empty;
            CreateFolder(path);

            LogMessage("start to get table of BShare in ace.");
            #region [Get Dictionary B share]
            dicBShare = GetBShareTable();
            #endregion

            LogMessage("start to generate BOD_notice_In_Scope.csv");
            #region [BOD_notice_In_Scope.csv]
            try
            {
                if (configObj.BODNoticeInScopeFile.Equals(BODNoticeInScope.Yes))
                {
                    Dictionary<string, List<string>> dicListBOD = new Dictionary<string, List<string>>();
                    List<string> listBODTitle = new List<string>() { "RIC", "Announcement date", "BOD date", "CDI rate", "SDI rate", "SIS ratio", "Annual/Semi-annual Report", "Source Date Local", "URL" };
                    dicListBOD.Add("title", listBODTitle);
                    List<string> listBODFoldet = new List<string>() { Path.Combine(classBODNotice, "BOD_notice_In_Scope") };
                    List<string> listBODFiles = null;
                    listBODFiles = GetAllFiles(listBODFoldet);
                    FillCSVBody(dicListBOD, listBODFiles);
                    string filePath = Path.Combine(path, "BOD_notice_In_Scope.csv");
                    XlsOrCsvUtil.GenerateXls0rCsv(filePath, dicListBOD);
                    AddResult("BOD_notice_In_Scope", filePath, "CSV Bulk File");
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            LogMessage("start to generate AGM_EGM_notice_for_NDA_and_CAM.xls/AGM_EGM_notice_for_NDA_and_CAM.csv");
            #region [AGM_EGM_notice_for_NDA_and_CAM.xls/AGM_EGM_notice_for_NDA_and_CAM.csv]
            try
            {
                if (configObj.AGMEGMNoticeForNDAAndCAMFile.Equals(AGMEGMNoticeForNDAAndCAM.Yes))
                {
                    List<string> listNDAandCAMFolder = new List<string>() { Path.Combine(classAGMNotice, "AGM_EGM_notice_for_NDA_and_CAM") };
                    List<string> listNDAandCAMFiles = null;
                    listNDAandCAMFiles = GetAllFiles(listNDAandCAMFolder);

                    //generate AGM_EGM_notice_for_NDA_and_CAM csv file
                    Dictionary<string, List<string>> dicListNDAandCAMCSV = new Dictionary<string, List<string>>();
                    List<string> listNDAandCAMTitleCSV = new List<string>() { "RIC", "Meeting date", "Meeting type", "Source Date Local", "Source Date GMT", "URL" };
                    dicListNDAandCAMCSV.Add("title", listNDAandCAMTitleCSV);
                    FillNDAandCAMCSVBody(dicListNDAandCAMCSV, listNDAandCAMFiles);
                    string filePathCsv = Path.Combine(path, "AGM_EGM_notice_for_NDA_and_CAM.csv");
                    XlsOrCsvUtil.GenerateXls0rCsv(filePathCsv, dicListNDAandCAMCSV);
                    AddResult("AGM_EGM_notice_for_NDA_and_CAM", filePathCsv, "CSV Bulk File");

                    //generate AGM_EGM_notice_for_NDA_and_CAM xls file
                    Dictionary<string, List<string>> dicListNDAandCAMXLS = new Dictionary<string, List<string>>();
                    List<string> listNDAandCAMTitleXLS = new List<string>() { "RIC", "Announcement date", "Meeting type", "Meeting date", "Record date", "Location", "Source Date Local", "URL" };
                    dicListNDAandCAMXLS.Add("title", listNDAandCAMTitleXLS);
                    FillNDAandCAMXLSBody(dicListNDAandCAMXLS, listNDAandCAMFiles);
                    filePathNDAAndCAMXls = Path.Combine(path, "AGM_EGM_notice_for_NDA_and_CAM.xls");
                    XlsOrCsvUtil.GenerateXls0rCsv(filePathNDAAndCAMXls, dicListNDAandCAMXLS);
                    AddResult("AGM_EGM_notice_for_NDA_and_CAM", filePathNDAAndCAMXls, "XLS Bulk File");
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            LogMessage("start to generate AGM_EGM_notice_for_CAM.xls");
            #region [AGM_EGM_notice_for_CAM.xls]
            try
            {
                if (configObj.AGMEGMNoticeForCAMFile.Equals(AGMEGMNoticeForCAM.Yes))
                {
                    List<string> listCAMFolder = new List<string>() { Path.Combine(classAGMNotice, "AGM_EGM_notice_for_CAM") };
                    List<string> listCAMFiles = null;
                    listCAMFiles = GetAllFiles(listCAMFolder);

                    //generate AGM_EGM_notice_for_CAM xls file
                    Dictionary<string, List<string>> dicListCAMXLS = new Dictionary<string, List<string>>();
                    List<string> listCAMTitleXLS = new List<string>() { "RIC", "Announcement date", "Meeting type", "Meeting date", "Record date", "Location", "Source Date Local", "URL" };
                    dicListCAMXLS.Add("title", listCAMTitleXLS);
                    FillNDAandCAMXLSBody(dicListCAMXLS, listCAMFiles);
                    filePathCAMXls = Path.Combine(path, "AGM_EGM_notice_for_CAM.xls");
                    XlsOrCsvUtil.GenerateXls0rCsv(filePathCAMXls, dicListCAMXLS);
                    AddResult("AGM_EGM_notice_for_CAM", filePathCAMXls, "XLS Bulk File");
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            LogMessage("start to generate AGM_EGM_result_In_Scope.csv");
            #region [AGM_EGM_result_In_Scope.csv]
            try
            {
                if (configObj.AGMEGMResultInScopeFile.Equals(AGMEGMResultInScope.Yes))
                {
                    List<string> listAGMResultFolder = new List<string>() { Path.Combine(classAGMResult, "AGM_EGM_result_In_Scope") };
                    List<string> listAGMResultFiles = null;
                    listAGMResultFiles = GetAllFiles(listAGMResultFolder);

                    //generate AGM_EGM_result_In_Scope.csv file
                    Dictionary<string, List<string>> dicListAGMResult = new Dictionary<string, List<string>>();
                    List<string> listAGMResultTitle = new List<string>() { "RIC", "Status", "Source Date Local", "Source Date GMT", "URL" };
                    dicListAGMResult.Add("title", listAGMResultTitle);

                    FillAGMResultBody(dicListAGMResult, listAGMResultFiles);
                    string filePath = Path.Combine(path, "AGM_EGM_result_In_Scope.csv");
                    XlsOrCsvUtil.GenerateXls0rCsv(filePath, dicListAGMResult);
                    AddResult("AGM_EGM_result_In_Scope", filePath, "CSV Bulk File");
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            LogMessage("start to generate Implement_notice.csv");
            #region [Implement_notice.csv]
            try
            {
                if (configObj.ImplementNoticeFile.Equals(ImplementNotice.Yes))
                {
                    List<string> listImplementNoticeFolder = new List<string>() { classImplementNotice };
                    List<string> listImplementNoticeFiles = null;
                    listImplementNoticeFiles = GetAllFiles(listImplementNoticeFolder);

                    //generate AGM_EGM_result_In_Scope.csv file
                    Dictionary<string, List<string>> dicListImplementNotice = new Dictionary<string, List<string>>();
                    List<string> listImplementNoticeTitle = new List<string>() { "RIC", "CDI rate", "SDI ratio", "SIS ratio", "Record date", "Ex date", "Pay date", "Exchange Rate", "Exchange Currency", "Listing date", "Source Date Local", "Source Date GMT", "URL" };
                    dicListImplementNotice.Add("title", listImplementNoticeTitle);

                    FillImplementNoticeBody(dicListImplementNotice, listImplementNoticeFiles);
                    string filePath = Path.Combine(path, "Implement_notice.csv");
                    XlsOrCsvUtil.GenerateXls0rCsv(filePath, dicListImplementNotice);
                    AddResult("Implement_notice", filePath, "CSV Bulk File");
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
            #endregion

            LogMessage("start to generate Meeting_yyyymmdd.xls");
            #region [Meeting_yyyymmdd.xls]
            if (!File.Exists(filePathNDAAndCAMXls) || !File.Exists(filePathCAMXls))
            {
                LogMessage(string.Format("{0} or {1} is not exist.", filePathCAMXls, filePathNDAAndCAMXls));
                return;
            }

            string pathMeeting = Path.Combine(path, string.Format("Meeting_{0}.xls", DateTime.Now.ToString("yyyyMMdd")));
            Dictionary<string, List<string>> dicListMeeting = new Dictionary<string, List<string>>();
            List<string> listMeetingTitle = new List<string>() { "RIC", "MEET Announcement Date", "Event Type", "Meeting Date", "MEET Record Date", "Meeting Location", "EVENT_SOURCE_LOCAL_TIME" };
            dicListMeeting.Add("title", listMeetingTitle);
            AddToDicList(ReadFromXLS(filePathNDAAndCAMXls), dicListMeeting);
            AddToDicList(ReadFromXLS(filePathCAMXls), dicListMeeting);
            XlsOrCsvUtil.GenerateXls0rCsv(pathMeeting, dicListMeeting);
            AddResult("Meeting_yyyymmdd", pathMeeting, "XLS Bulk File");
            #endregion

            LogMessage("start to generate BulkInsertMeeting_yyyymmdd.xls");
            #region [BulkInsertMeeting_yyyymmdd.xls]
            string pathBulkInsert = Path.Combine(path, string.Format("BulkInsertMeeting_{0}.xls", DateTime.Now.ToString("yyyyMMdd")));
            Dictionary<string, List<string>> dicListBulkInsert = new Dictionary<string, List<string>>();
            List<string> listBulkInsertTitle = new List<string>() 
            { 
                "RIC", "ISIN", "MEET Underlying OrgId", "Event Type", "Event Status", "MandVollnd", "Processing Status",
                "MEET Announcement Date", "MEET Rescindeddate", "MEET Record Date", "Meeting Status", "Meeting Date",
                "Meet End Date1", "Meetingdate Timezonecode", "Meeting Date2", "Meet End Date2", "Meetingdate2 Timezonecode",
                "Meeting Date3", "Meet End Date3", "Meetingdate3 Timezonecode", "Meeting Location", "Location Url",
                "Event Url", "Webcast Type", "Webcast Link", "Webcast Exp Date", "PriCountrycode", "Live Pri Dial",
                "SecCountrycode", "Live Sec Dial", "Live Pass Code", "Dial in Url", "Dial in Notes", "Replay Start Date",
                "Replay End Date", "Replay Pri Dial Countrycode", "Replay Pri Dial", "Replay Sec Dial Countrycode", 
                "Replay Sec Dial", "Replay Pass Code", "Replay in Notes", "RSVP by Date", "RSVP Phone Countrycode", 
                "RSVP Phone", "RSVP Fax Countrycode", "RSVP Fax", "RSVP Email", "RSVP Url", "Description Type", "Description",
                "Analyst Notes", "Data Entry Status", "EVENT_SOURCE_EST_CODE", "EVENT_SOURCE_LOCAL_TIME", "EVENT_SOURCE_TZ_CODE",
                "EVENT_SOURCE_SP_CODE", "EVENT_SOURCE_DESCRIPTION", "EVENT_SOURCE_LINK" 
            };
            dicListBulkInsert.Add("title", listBulkInsertTitle);
            FillInBulkInsert(dicListMeeting, dicListBulkInsert);
            XlsOrCsvUtil.GenerateXls0rCsv(pathBulkInsert, dicListBulkInsert);
            #endregion
        }

        private void FillInBulkInsert(Dictionary<string, List<string>> dicListMeeting, Dictionary<string, List<string>> dicListBulkInsert)
        {
            try
            {
                if (dicListMeeting == null || dicListMeeting.Count == 0)
                {
                    LogMessage("meeting list is empty,so neednot genetate bulk insert meeting file.");
                    return;
                }

                List<MeetingEventTemplate> meetingEvert = FillWithMeeting(dicListMeeting);
                FillInOthersValue(meetingEvert);
                ConvertTemplateToDicList(meetingEvert, dicListBulkInsert);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void ConvertTemplateToDicList(List<MeetingEventTemplate> meetingEvert, Dictionary<string, List<string>> dicListBulkInsert)
        {
            throw new NotImplementedException();
        }

        private void FillInOthersValue(List<MeetingEventTemplate> meetingEvert)
        {
            /*The following value has been added in template, so this function need to add the valude not been added.*/
            //0"RIC", 
            //1"MEET Announcement Date",
            //2"Event Type",
            //3"Meeting Date", 
            //4"MEET Record Date",
            //5"Meeting Location", 
            //6"EVENT_SOURCE_LOCAL_TIME" 

            foreach (var tmp in meetingEvert)
            {
                //RIC 
                //ISIN 
                //MEETUnderlyingOrgId 
                //EventType 
                //EventStatus 
                //MandVollnd 
                //ProcessingStatus 
                //MEETAnnouncementDate 
                //MEETRescindeddate 
                //MEETRecordDate 
                //MeetingStatus 
                //MeetingDate 
                //MeetEndDate1 
                //MeetingdateTimezonecode 
                //MeetingDate2 
                //MeetEndDate2 
                //Meetingdate2Timezonecode 
                //MeetingDate3 
                //MeetEndDate3 
                //Meetingdate3Timezonecode 
                //MeetingLocation 
                //LocationUrl 
                //EventUrl 
                //WebcastType 
                //WebcastLink 
                //WebcastExpDate 
                //PriCountrycode 
                //LivePriDial 
                //SecCountrycode 
                //LiveSecDial 
                //LivePassCode 
                //DialInUrl 
                //DialInNotes 
                //ReplayStartDate 
                //ReplayEndDate 
                //ReplayPriDialCountrycode 
                //ReplayPriDial 
                //ReplaySecDialCountrycode 
                //ReplaySecDial 
                //ReplayPassCode 
                //ReplayInNotes 
                //RSVPByDate 
                //RSVPPhoneCountrycode 
                //RSVPPhone 
                //RSVPFaxCountrycode 
                //RSVPFax 
                //RSVPEmail 
                //RSVPUrl 
                //DescriptionType 
                //Description 
                //AnalystNotes 
                //DataEntryStatus 
                //EventSourceESTCode 
                //EventSourceLocalTime 
                //EventSourceTZCode 
                //EventSourceSPCode 
                //EventSourceDescription 
                //EventSourceLink 
            }
        }

        private List<MeetingEventTemplate> FillWithMeeting(Dictionary<string, List<string>> dicListMeeting)
        {
            List<MeetingEventTemplate> result = new List<MeetingEventTemplate>();
            MeetingEventTemplate tmp;

            List<List<string>> meetingValue = dicListMeeting.Values.ToList();
            foreach (var item in meetingValue)
            {
                tmp = new MeetingEventTemplate();
                tmp.RIC = item[0];                  //0"RIC", 
                tmp.MEETAnnouncementDate = item[1]; //1"MEET Announcement Date",
                tmp.EventType = item[2];            //2"Event Type",
                tmp.MeetingDate = item[3];          //3"Meeting Date", 
                tmp.MEETRecordDate = item[4];       //4"MEET Record Date",
                tmp.MeetingLocation = item[5];      //5"Meeting Location", 
                tmp.EventSourceLocalTime = item[6]; //6"EVENT_SOURCE_LOCAL_TIME" 
                result.Add(tmp);
            }

            return result;
        }

        private void AddToDicList(List<List<string>> list, Dictionary<string, List<string>> dicList)
        {
            if (list == null || list.Count <= 1)
                return;

            for (int i = 1; i < list.Count; i++)
                dicList.Add(dicList.Count.ToString(), list[i]);
        }

        private List<List<string>> ReadFromXLS(string path)
        {
            try
            {
                ExcelApp app = new ExcelApp(false, false);
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, path);
                return WorkbookExtension.ToList(workbook, 1);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return null;
        }

        private void SecondJopUpdateToDB()
        {
            if (!configObj.UpdateOOS.Equals(UpdateOOS.Yes))
                return;

            try
            {
                listOOSId = GetOOSId(listOOSFolder);
                UpdateToAceDB(listOOSId);
            }
            catch (Exception ex)
            {
                string msg = string.Format("SecondJopUpdateToDB() failed. msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void FirstJopDownloadAndClassify()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(configObj.OutputPath) || string.IsNullOrWhiteSpace(configObj.AssignedTo))
                {
                    MessageBox.Show("Please check configuration value.");
                    return;
                }

                GetDownloadListFromDB(historicalAnnouncements);

                if (historicalAnnouncements == null || historicalAnnouncements.Count == 0)
                {
                    MessageBox.Show("No new records in DB .");
                    return;
                }

                DeleteFiles(downloadPath);
                DeleteFiles(typePath);
                listNewDownloadPdf = DownloadFile(historicalAnnouncements);
                OutPutDownloadErrorInfo(listDownloadPdfError);
                FindByTitle(listNewDownloadPdf);
                ClassifySecondTime();
            }
            catch (Exception ex)
            {
                string msg = string.Format("FirstJopDownloadAndClassify() failed. msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }
        #endregion

        #region [ThirdJopGenerateBulkFile]

        private Dictionary<string, string> GetBShareTable()
        {
            List<Tickers> listBShare = new List<Tickers>();

            if (extendContext == null)
            {
                Logger.Log("can not connect ace db to get BShare table");
                return null;
            }

            try
            {
                var query = string.Format("SELECT * FROM tickers");
                listBShare.AddRange(extendContext.ExecuteQuery<Tickers>(query, 60).ToList());

                return GetBShareDictionary(listBShare);

            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private Dictionary<string, string> GetBShareDictionary(List<Tickers> listBShare)
        {
            Dictionary<string, string> dicBShare = new Dictionary<string, string>();

            try
            {
                if (listBShare == null || listBShare.Count == 0)
                {
                    Logger.Log("table of tickers in sce is empty.", Logger.LogType.Warning);
                    return dicBShare;
                }

                foreach (var item in listBShare)
                {
                    if (dicBShare.ContainsKey(item.TickerA))
                        continue;

                    dicBShare.Add(item.TickerA, item.TickerB);
                }

                return dicBShare;

            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private void FillImplementNoticeBody(Dictionary<string, List<string>> dicListBody, List<string> listFiles)
        {
            int annId = 0;
            string sourceDateLocal = string.Empty;
            string sourceDateGMT = string.Empty;
            string ric = string.Empty;

            if (listFiles == null || listFiles.Count == 0)
            {
                string msg = string.Format("no valid *.pdf in AGM_EGM_result_In_Scope folder. ");
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            foreach (var item in listFiles)
            {
                List<string> list = new List<string>();
                annId = GetIdFromPath(item);
                string tickerA = string.Empty;
                bool isExistBShare = false;
                string recordDate = string.Empty;
                string exDate = string.Empty;

                if (annId == 0)
                {
                    string msg = string.Format("can't get id of the path {0}", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    continue;
                }

                sourceDateLocal = GetSourceDateLocal(annId);
                sourceDateGMT = GetSourceDateGMT(annId);
                ric = GetRic(annId);
                tickerA = GetTicker(annId);

                if (!string.IsNullOrEmpty(tickerA) && dicBShare.ContainsKey(tickerA))
                    isExistBShare = true;

                recordDate = GetRecordDateFromImplementNotice(item);//dd-MM-yyyy
                exDate = GetExDate(item);//dd-MM-yyyy

                string exDateAShare = CalculateAShareExDate(recordDate, exDate);

                list.Add(ric);                                    //RIC 
                list.Add(GetCDIRate(item));                       //CDI rate                    
                list.Add(GetSDIRate(item));                       //SDI ratio                    
                list.Add(GetSISRatio(item));                      //SIS ratio                    
                list.Add(recordDate);                             //Record date                    
                list.Add(exDateAShare);                           //Ex date
                list.Add(exDateAShare);                           //Pay date
                list.Add(string.Empty);                           //Exchange Rate
                list.Add(GetExchangeCurrency(ric));               //Exchange Currency
                list.Add(GetListingDate(item));                   //Listing date
                list.Add(sourceDateLocal);                        //Source Date Local
                list.Add(sourceDateGMT);                          //Source Date GMT
                list.Add(GetUrl(annId));                          //URL

                dicListBody.Add(annId.ToString(), list);

                if (!isExistBShare)
                    continue;

                //Add Bshare Records
                List<string> listBShare = new List<string>();

                listBShare.Add(ConvertTickerToRic(dicBShare[tickerA]));           //RIC 
                listBShare.Add(GetCDIRate(item));                                 //CDI rate                    
                listBShare.Add(GetSDIRate(item));                                 //SDI ratio                    
                listBShare.Add(GetSISRatio(item));                                //SIS ratio                    
                listBShare.Add(AddBuinessDay(exDate, 2, "dd-MM-yyyy"));           //Record date    
                listBShare.Add(exDate);                                           //Ex date
                listBShare.Add(string.Empty);                                     //Pay date
                listBShare.Add(GetExchangeRate(item));                            //Exchange Rate
                listBShare.Add(GetExchangeCurrency(ric));                         //Exchange Currency
                listBShare.Add(GetListingDate(item));                             //Listing date
                listBShare.Add(sourceDateLocal);                                  //Source Date Local
                list.Add(sourceDateGMT);                                          //Source Date GMT
                listBShare.Add(GetUrl(annId));                                    //URL

                dicListBody.Add(annId.ToString() + "(BShare)", listBShare);
            }
        }

        private string CalculateAShareExDate(string recordDate, string exDate)
        {
            //Calculation Record date or Ex date:
            //Ex date is always 1 business day after Record Date.
            //If only Record date can be found, calculate the Ex date:
            //Ex date=Record date + 1 business day
            //If only Ex date can be found, calculate the Record date:
            //Record date=Ex date – 1 business day

            if (string.IsNullOrEmpty(recordDate) && string.IsNullOrEmpty(exDate))
                return string.Empty;

            if (string.IsNullOrEmpty(recordDate))
                return exDate;

            if (string.IsNullOrEmpty(exDate))
                return AddBuinessDay(recordDate, 1, "dd-MM-yyyy");

            //string ss = AddBuinessDay(exDate, -1, "dd-MM-yyyy");
            //if (!recordDate.Equals(ss))
            //    MessageBox.Show(recordDate + "\r\n" + ss);

            if (recordDate.Equals(AddBuinessDay(exDate, -1, "dd-MM-yyyy")))
                return exDate;

            return string.Empty;
        }

        private string AddBuinessDay(string dateTime, int businessDayRange, string dateFormatString)
        {
            //If there is B share, calculate the Record date:
            //Record date=Ex date+2 business day
            //Record date cannot be grabbed from the announcement for B share.
            string resultPath = string.Empty;

            if (string.IsNullOrEmpty(dateTime))
                return string.Empty;

            DateTime dt;
            if (!DateTime.TryParseExact(dateTime, dateFormatString, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
                return string.Empty;

            int dayRange = 1;
            if (businessDayRange < 0)
            {
                businessDayRange = -businessDayRange;
                dayRange = -1;
            }

            int count = 0;
            while (count < businessDayRange)
            {
                dt = dt.AddDays(dayRange);
                if (dt.DayOfWeek != DayOfWeek.Saturday && dt.DayOfWeek != DayOfWeek.Sunday)
                    count++;
            }

            return dt.ToString(dateFormatString);
        }

        private string GetExDate(string item)
        {
            try
            {
                //除权除息日：2013 年 5 月 13 日
                //除权除息日为 2013 年 5 月 3 日
                string patternPdfExDate = @".{0,10}(?<RecordDate>\d\s*\d\s*\d?\s*\d?\s*年\s*\d\s*\d?\s*月\s*\d\s*\d?\s*日).{0,30}";
                List<PdfString> listExDate = FindListPdfString(patternPdfExDate, item);
                if (listExDate == null || listExDate.Count == 0)
                    return string.Empty;

                string patternExDate = @"除权除息日\D*(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日";
                List<string> groupName = new List<string>() { "year", "month", "day" };
                List<string> exDate = FindValueFromListPdfString(patternExDate, listExDate, groupName);
                if (exDate == null || exDate.Count == 0)
                    return string.Empty;

                return string.Format("{0}-{1}-{2}",
                                    exDate[2].Trim().Length == 1 ? "0" + exDate[2].Trim() : exDate[2].Trim(),//year
                                    exDate[1].Trim().Length == 1 ? "0" + exDate[1].Trim() : exDate[1].Trim(),//month
                                    exDate[0].Trim().Length == 2 ? "20" + exDate[0].Trim() : exDate[0].Trim()//day
                                    );//dd-MM-yyyy
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return string.Empty;
        }

        private string GetRecordDateFromImplementNotice(string item)
        {
            try
            {
                //股权登记日：2013 年 5 月 13 日
                //股权登记日为 2013 年 5 月 3 日
                string patternPdfRecordDate = @".{0,10}(?<RecordDate>\d\s*\d\s*\d?\s*\d?\s*年\s*\d\s*\d?\s*月\s*\d\s*\d?\s*日).{0,30}";
                List<PdfString> listRecordDate = FindListPdfString(patternPdfRecordDate, item);
                if (listRecordDate == null || listRecordDate.Count == 0)
                    return string.Empty;

                string patternRecordDate = @"股权登记日\D*(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日";
                List<string> groupName = new List<string>() { "year", "month", "day" };
                List<string> recordDate = FindValueFromListPdfString(patternRecordDate, listRecordDate, groupName);
                if (recordDate == null || recordDate.Count == 0)
                    return string.Empty;

                return string.Format("{0}-{1}-{2}",
                                    recordDate[2].Trim().Length == 1 ? "0" + recordDate[2].Trim() : recordDate[2].Trim(),//year
                                    recordDate[1].Trim().Length == 1 ? "0" + recordDate[1].Trim() : recordDate[1].Trim(),//month
                                    recordDate[0].Trim().Length == 2 ? "20" + recordDate[0].Trim() : recordDate[0].Trim()//day
                                    );//dd-MM-yyyy
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return string.Empty;
        }

        private string GetExchangeRate(string item)
        {
            //人民币中间价（1:6.1537）
            //中间价1：6.1710
            //中间价（1 港币＝0.7955 人民币）
            //中间价（美元：人民币=1：6.7958）
            //Exchange rate= 1/the grabbed number
            //Output Format: xx.xxxxxx(round-off to 6 decimal places)
            //Eg. 美元兑换人民币中间价（1:6.1537）
            //Exchange rate=1/6.1537=0.162504
            try
            {
                string patternPdfExchangeRate = @"中\s*间\s*价.{0,30}";
                List<PdfString> listExchangeRate = FindListPdfString(patternPdfExchangeRate, item);
                if (listExchangeRate == null || listExchangeRate.Count == 0)
                    return string.Empty;

                string patternExchangeRate = @"中间价.*(?<Value>(＝|=|:|：|比|等于)[0-9\.]{4,8})";
                List<string> listGroupName = new List<string>() { "Value" };
                List<string> exchangeRate = FindValueFromListPdfString(patternExchangeRate, listExchangeRate, listGroupName);
                if (exchangeRate == null || exchangeRate.Count == 0)
                    return string.Empty;

                decimal exchangeRateDecimal;
                if (decimal.TryParse(exchangeRate[0], out exchangeRateDecimal))
                    return Math.Round(1 / exchangeRateDecimal, 6).ToString("f6");
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return string.Empty;
        }

        private List<string> FindValueFromListPdfString(string pattern, List<PdfString> listPdfString, List<string> groupName)
        {
            List<string> result = new List<string>();

            try
            {
                foreach (var item in listPdfString)
                {
                    Match match = (new Regex(pattern)).Match(item.ToString().Replace(" ", ""));

                    if (match.Success)
                    {
                        foreach (var name in groupName)
                            result.Add(match.Groups[name].Value);

                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return null;
        }

        private List<PdfString> FindListPdfString(string pattern, string path)
        {
            PDFDoc doc = null;

            try
            {
                if (!File.Exists(path))
                {
                    string msg = string.Format("the file {0} is not exist.", path);
                    Logger.Log(msg, Logger.LogType.Error);
                    return null;
                }

                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(path);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", path);
                    Logger.Log(msg, Logger.LogType.Error);
                    return null;
                }

                PdfAnalyzer pa = new PdfAnalyzer();
                return pa.RegexSearchAllPages(doc, pattern);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
            finally
            {
                doc.Dispose();
            }
        }

        private string GetListingDate(string item)
        {
            string listingDate = "     ";//dd-MM-yyyy
            PDFDoc doc = null;

            try
            {
                string patternListingDate = @".{0,20}(?<ListingDate>\d{2,4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)";

                if (!File.Exists(item))
                {
                    string msg = string.Format("the file {0} is not exist.", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return listingDate;
                }

                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(item);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return listingDate;
                }

                PdfAnalyzer pa = new PdfAnalyzer();
                List<PdfString> listListingDate = pa.RegexSearchAllPages(doc, patternListingDate);

                if (listListingDate == null || listListingDate.Count == 0)
                {
                    string msg = string.Format("can't found BODDate in {0} by keyWord {1}).", item, patternListingDate);
                    Logger.Log(msg, Logger.LogType.Error);
                    return listingDate;
                }

                return GetListingDateSecond(listListingDate);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return listingDate;
            }
            finally
            {
                doc.Dispose();
            }
        }

        private string GetListingDateSecond(List<PdfString> listListingDate)
        {
            string listingDate = "     ";//dd-MM-yyyy

            try
            {
                if (listListingDate == null || listListingDate.Count == 0)
                    return listingDate;

                string result = string.Empty;
                string patternListingDate = @"股份上市.{0,10}(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日?";

                foreach (var item in listListingDate)
                {
                    result = item.ToString().Replace(" ", "");

                    Match ma = (new Regex(patternListingDate)).Match(result);

                    if (ma.Success)
                    {
                        listingDate = FormatedRecordDateFound(ma);//record date and listing date are the same format "dd-MM-yyyy"

                        if (!string.IsNullOrEmpty(listingDate))
                            return listingDate;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return listingDate;
        }

        private string GetExchangeCurrency(string ric)
        {
            string exchangeCurrency = string.Empty;

            if (string.IsNullOrEmpty(ric))
                return exchangeCurrency;

            try
            {
                if (ric.Trim().StartsWith("9"))
                    exchangeCurrency = "USD";
                else if (ric.Trim().StartsWith("2"))
                    exchangeCurrency = "HKD";
                else
                    exchangeCurrency = "     ";
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return exchangeCurrency;
        }

        private void FillAGMResultBody(Dictionary<string, List<string>> dicListBody, List<string> listFiles)
        {
            int annId = 0;
            string sourceDateLocal = string.Empty;
            string sourceDateGMT = string.Empty;

            if (listFiles == null || listFiles.Count == 0)
            {
                string msg = string.Format("no valid *.pdf in AGM_EGM_result_In_Scope folder. ");
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            foreach (var item in listFiles)
            {
                List<string> list = new List<string>();
                annId = GetIdFromPath(item);

                if (annId == 0)
                {
                    string msg = string.Format("can't get id of the path {0}", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    continue;
                }

                sourceDateLocal = GetSourceDateLocal(annId);
                sourceDateGMT = GetSourceDateGMT(annId);

                //"RIC", "Status", "Source Date Local", "Source Date GMT", "URL" };
                list.Add(GetRic(annId));                            //RIC
                list.Add(GetStatus(item));                          //Status
                list.Add(sourceDateLocal);                          //Source Date Local
                list.Add(sourceDateGMT);                            //Source Date GMT
                list.Add(GetUrl(annId));                            //URL

                dicListBody.Add(annId.ToString(), list);
            }
        }

        private string GetStatus(string pdfPath)
        {
            string status = "NULL";//Output format: “APPD” or “NULL”
            PDFDoc doc = null;

            try
            {
                List<string> listPatternStatus = new List<string>()
                {
                    @"(不存在|无|未有|没有|未)(出现|涉及|有)?\d*(变更|否决|新增|增加|或|及|提案|议案|新提案|修改提案|、)+",
                    @"是否有否决(提案|议案)的情况.{0,4}(否|无|没有|不存在|未有)"
                };

                if (!File.Exists(pdfPath))
                {
                    string msg = string.Format("the file {0} is not exist.", pdfPath);
                    Logger.Log(msg, Logger.LogType.Error);
                    return status;
                }

                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(pdfPath);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", pdfPath);
                    Logger.Log(msg, Logger.LogType.Error);
                    return status;
                }

                if (IsValidSataus(doc, listPatternStatus))
                    status = "APPD";

                return status;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return status;
            }
            finally
            {
                doc.Dispose();
            }
        }

        private bool IsValidSataus(PDFDoc doc, List<string> listPatternStatus)
        {
            bool boolSataus = false;

            try
            {
                PdfAnalyzer pa = new PdfAnalyzer();
                List<PdfString> listStatus = null;

                foreach (var item in listPatternStatus)
                {
                    listStatus = pa.RegexSearchAllPages(doc, item);

                    if (listStatus != null && listStatus.Count > 0)
                    {
                        boolSataus = true;
                        break;
                    }
                }

                return boolSataus;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return boolSataus;
            }
        }

        private void FillNDAandCAMXLSBody(Dictionary<string, List<string>> dicListBody, List<string> listFiles)
        {
            int annId = 0;
            string annDate = string.Empty;
            string meetingDate = string.Empty;
            string meetingType = string.Empty;
            string sourceDateLocal = string.Empty;
            string tickerA = string.Empty;
            string tickerB = string.Empty;

            if (listFiles == null || listFiles.Count == 0)
            {
                string msg = string.Format("no valid *.pdf in BOD_notice folder. ");
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            foreach (var item in listFiles)
            {
                List<string> list = new List<string>();

                annId = GetIdFromPath(item);

                if (annId == 0)
                {
                    string msg = string.Format("can't get id of the path {0}", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    continue;
                }

                annDate = GetAnnouncementDate(annId);
                meetingDate = GetMeetingDate(item, annDate);
                meetingType = GetMeetingTypeMEET(annId);
                sourceDateLocal = GetSourceDateLocal(annId);

                list.Add(GetRic(annId));                            //RIC
                list.Add(annDate);                                  //Announcement date
                list.Add(meetingType);                              //Meetingtype
                list.Add(meetingDate);                              //Meeting date
                list.Add(GetRecordDate(item, annDate, meetingDate)); //Record date 
                list.Add(GetLocation(item));                        //Location
                list.Add(sourceDateLocal);                          //Source Date Local
                list.Add(GetUrl(annId));                            //URL

                dicListBody.Add(annId.ToString(), list);

                tickerA = GetTicker(annId);

                if (string.IsNullOrEmpty(tickerA) || !dicBShare.ContainsKey(tickerA))
                    continue;

                tickerB = dicBShare[tickerA];
                dicListBody.Add(tickerB + dicListBody.Count.ToString(), ReplaceTickerAByTickerB(tickerB, list));
            }
        }

        private string GetMeetingTypeMEET(int annId)
        {
            string meetingType = "    ";

            try
            {
                var chn = historicalAnnouncements.FirstOrDefault(p => p.ID == annId);

                if (chn == null)
                {
                    string msg = string.Format("can't found the record in ace db use id={0}", annId.ToString());
                    Logger.Log(msg, Logger.LogType.Error);
                    return meetingType;
                }

                if (chn.Title.Contains("年度"))
                    meetingType = "MEET";
                else
                    meetingType = "XMET";
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return meetingType;
        }

        private string GetLocation(string item)
        {
            string location = "     ";        //dd-MMM-yyyy
            PDFDoc doc = null;
            try
            {
                //会议召开地点：…
                //会议地点：… 
                string patternLocation = @"(?<Location>会\s*议\s*(召\s*开)?\s*地\s*点\s*\:?.{50})";

                if (!File.Exists(item))
                {
                    string msg = string.Format("the file {0} is not exist.", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return location;
                }

                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(item);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return location;
                }

                PdfAnalyzer pa = new PdfAnalyzer();
                List<PdfString> listLocation = pa.RegexSearchAllPages(doc, patternLocation);

                if (listLocation == null || listLocation.Count == 0)
                {
                    string msg = string.Format("can't found meetingDate in {0} by keyWord {1}).", item, patternLocation);
                    Logger.Log(msg, Logger.LogType.Error);
                    return location;
                }

                location = listLocation[0].ToString().Replace(" ", "");

                return location;

            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetLocation(string item) error.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return location;
            }
            finally
            {
                doc.Dispose();
            }

        }

        private string GetRecordDate(string item, string annDate, string meetingDate)
        {
            string recordDate = "     ";        //dd-MMM-yyyy
            PDFDoc doc = null;
            try
            {
                //股权登记日：2013 年 5 月 13 日
                //股权登记日为 2013 年 5 月 3 日
                //截至2013年5月23日
                //截止 2013 年 5 月 3 日
                //（1）截止2013年5月16日下午3时交易结束后在中国证券登记结算有限
                //1.凡于2014年5月23日（星期五）下午收市后在中国证券登记结算有限责
                //1、截止 2014 年 5月1 5日下午收市时在中国证券登记结算有限责任公司上海
                //（1）凡是2014年5月12日下午收市后，在中国证券登记结算有限责任公司深
                //（1）于2014年5月15日下午深圳证券交易所收市后在中国证券登记结算有限责
                //（1）截至2014年4月28日下午15:00收市后在中国证券登记结算有限公司深圳分
                //1、2014年5月16日下午15:00收市后在中国证券登记结算有限公司深圳分公司登记

                //string patternRecordDate = @"(?<RecordDate>.{0,10}\d{2,4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日).{0,30}";
                string patternRecordDate = @".{0,10}(?<RecordDate>\d\s*\d\s*\d?\s*\d?\s*年\s*\d\s*\d?\s*月\s*\d\s*\d?\s*日).{0,30}";
                List<string> listPatternRecordDate = new List<string>();
                listPatternRecordDate.Add(@"(股权登记日|截止|截至)\D*(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日");
                listPatternRecordDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日.*(收市).*");

                if (!File.Exists(item))
                {
                    string msg = string.Format("the file {0} is not exist.", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return recordDate;
                }

                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(item);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return recordDate;
                }

                PdfAnalyzer pa = new PdfAnalyzer();
                List<PdfString> listRecordMDate = pa.RegexSearchAllPages(doc, patternRecordDate);

                if (listRecordMDate == null || listRecordMDate.Count == 0)
                {
                    string msg = string.Format("can't found meetingDate in {0} by keyWord {1}).", item, patternRecordDate);
                    Logger.Log(msg, Logger.LogType.Error);
                    return recordDate;
                }

                return GetRecordDateSecond(listRecordMDate, listPatternRecordDate, annDate, meetingDate);
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetRcordDate(string item, string annDate, string meetingDate) error.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return recordDate;
            }
            finally
            {
                doc.Dispose();
            }
        }

        private string GetRecordDateSecond(List<PdfString> listRecordMDate, List<string> listPatternRecordDate, string annDate, string meetingDate)
        {
            string recordDate = "     ";

            foreach (PdfString pdfString in listRecordMDate)
            {
                string result = pdfString.ToString().Replace(" ", "");

                foreach (var pattern in listPatternRecordDate)
                {
                    Match ma = (new Regex(pattern)).Match(result);

                    if (ma.Success)
                    {
                        recordDate = FormatedRecordDateFound(ma);

                        if (IsValidRecordDate(recordDate, annDate, meetingDate))
                        {
                            DateTime dtRecord;

                            if (DateTime.TryParse(FormatDateString(recordDate), out dtRecord))
                                return dtRecord.ToString("dd-MMM-yyyy");

                            //DateTime dt = DateTime.Parse(recordDate);
                            //return dt.ToString("dd-MMM-yyyy");
                        }
                        else
                            recordDate = "     ";
                    }
                }
            }

            return recordDate;
        }

        private string FormatDateString(string recordDate)
        {
            string result = "     ";

            if ((recordDate + "").Trim().Length == 0)
                return result;

            try
            {
                //23-01-2014
                string year = recordDate.Substring(6, 4);
                string month = recordDate.Substring(3, 2);
                string day = recordDate.Substring(0, 2);
                result = string.Format("{0}-{1}-{2}", month, day, year);

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                     System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                     System.Reflection.MethodBase.GetCurrentMethod().Name,
                     ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return result;
            }

        }

        private bool IsValidRecordDate(string recordDate, string annDate, string meetingDate)// annDate <= recordDate <= meetingDate
        {
            bool result = false;

            if (string.IsNullOrWhiteSpace(recordDate))
                return result;

            if (string.IsNullOrWhiteSpace(annDate))
            {
                if (IsFirstSmallSecond(recordDate, meetingDate))
                    result = true;
            }
            else if (string.IsNullOrWhiteSpace(meetingDate))
            {
                if (IsFirstSmallSecond(annDate, recordDate))
                    result = true;
            }
            else
            {
                if (IsFirstSmallSecond(annDate, recordDate) && IsFirstSmallSecond(recordDate, meetingDate))
                    result = true;
            }

            return result;
        }

        private bool IsFirstSmallSecond(string first, string second)//first must small second
        {
            bool result = false;

            if (string.IsNullOrWhiteSpace(first) || string.IsNullOrWhiteSpace(second))
                return true;

            int firstYear = System.Convert.ToInt32(first.Substring(6, 4));  //dd-mm-yyyy
            int firstMonth = System.Convert.ToInt32(first.Substring(3, 2));
            int firstDay = System.Convert.ToInt32(first.Substring(0, 2));
            int secondYear = System.Convert.ToInt32(second.Substring(6, 4));  //dd-mm-yyyy
            int secondMonth = System.Convert.ToInt32(second.Substring(3, 2));
            int secondDay = System.Convert.ToInt32(second.Substring(0, 2));

            if (firstYear < secondYear)
            {
                result = true;
            }
            else if (firstYear == secondYear)
            {
                if (firstMonth < secondMonth)
                {
                    result = true;
                }
                else if (firstMonth == secondMonth)
                {
                    if (firstDay <= secondDay)
                        result = true;
                }
            }

            return result;
        }

        private string FormatedRecordDateFound(Match ma)
        {
            string result = "     ";
            string year = string.Empty;
            string month = string.Empty;
            string day = string.Empty;
            year = ma.Groups["year"].Value;
            month = ma.Groups["month"].Value;
            day = ma.Groups["day"].Value;
            month = month.Length == 1 ? "0" + month : month;
            day = day.Length == 1 ? "0" + day : day;
            result = string.Format("{0}-{1}-{2}", day, month, year);

            return result;
        }

        private void FillNDAandCAMCSVBody(Dictionary<string, List<string>> dicListBody, List<string> listFiles)
        {
            int annId = 0;
            string annDate = string.Empty;
            string sourceDateLocal = string.Empty;
            string sourceDateGMT = string.Empty;
            string tickerA = string.Empty;
            string tickerB = string.Empty;

            if (listFiles == null || listFiles.Count == 0)
            {
                string msg = string.Format("no valid *.pdf in BOD_notice folder. ");
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            foreach (var item in listFiles)
            {
                List<string> list = new List<string>();
                annId = GetIdFromPath(item);

                if (annId == 0)
                {
                    string msg = string.Format("can't get id of the path {0}", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    continue;
                }

                annDate = GetAnnouncementDate(annId);
                sourceDateLocal = GetSourceDateLocal(annId);
                sourceDateGMT = GetSourceDateGMT(annId);

                list.Add(GetRic(annId));                                    //RIC
                list.Add(GetMeetingDate(item, annDate));                    //Meeting date
                list.Add(GetMeetingTypeAGM((GetFileName(item))));          //Meeting type
                list.Add(sourceDateLocal);                                  //Sourced Date Local
                list.Add(sourceDateGMT);                                    //Source Date GMT
                list.Add(GetUrl(annId));                                    //URL

                dicListBody.Add(annId.ToString(), list);

                tickerA = GetTicker(annId);

                if (string.IsNullOrEmpty(tickerA) || !dicBShare.ContainsKey(tickerA))
                    continue;

                tickerB = dicBShare[tickerA];
                dicListBody.Add(tickerB + dicListBody.Count.ToString(), ReplaceTickerAByTickerB(tickerB, list));
            }
        }

        private string GetMeetingTypeAGM(string title)
        {
            if ((title + "").Trim().Length == 0)
                return string.Empty;

            if (title.Contains("年度"))
                return "AGM";
            else
                return "EGM";
        }

        private string GetMeetingDate(string item, string annDate)
        {
            string meetingDate = "     ";//dd-mm-yyyy HH24:mm:ss
            PDFDoc doc = null;
            try
            {
                string patternNDAandCAMDate = @"(?<NDAandCAMDate>会\s*议\s*.{2,15}\d{2,4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日.{0,15})";
                List<string> listPatternNDAandCAMDate = new List<string>();
                //会议时间：2013 年 5 月 9 日（星期四）上午 9:30
                //会议召开时间：2013 年 5 月 8 日上午 9 时 30 分
                //会议时间：2013 年 5 月 10 日（星期五）上午 10：00 时
                //会议召开时间：2013 年 5 月 17 日（星期五） 上午 10:00
                //会议召开时间：2013 年 5 月 31 日（周五）上午 9：00 时
                //会议召开日期、时间：2013 年 5 月 10 日（星期五）9:30
                listPatternNDAandCAMDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})(\:|点|时)(?<minute>\d{1,2})");//0
                listPatternNDAandCAMDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})(\:|点|时)(?<minute>\d{1,2})");   //1 
                listPatternNDAandCAMDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})点半");//2
                listPatternNDAandCAMDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})点半");   //3 
                listPatternNDAandCAMDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})(点|时)");//4
                listPatternNDAandCAMDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})(点|时)");   //5 
                listPatternNDAandCAMDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日");                                             //default dd-mm-yyyy 00:00:00                      //6 

                if (!File.Exists(item))
                {
                    string msg = string.Format("the file {0} is not exist.", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return meetingDate;
                }

                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(item);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return meetingDate;
                }

                PdfAnalyzer pa = new PdfAnalyzer();
                List<PdfString> listNDAandCAMDate = pa.RegexSearchAllPages(doc, patternNDAandCAMDate);

                if (listNDAandCAMDate == null || listNDAandCAMDate.Count == 0)
                {
                    string msg = string.Format("can't found meetingDate in {0} by keyWord {1}).", item, patternNDAandCAMDate);
                    Logger.Log(msg, Logger.LogType.Error);
                    return meetingDate;
                }

                return GetMeetingDateSecond(listNDAandCAMDate, listPatternNDAandCAMDate, annDate);
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetMeetingDate(string item) error.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return meetingDate;
            }
            finally
            {
                doc.Dispose();
            }
        }

        private string GetMeetingDateSecond(List<PdfString> listNDAandCAMDate, List<string> listPatternNDAandCAMDate, string annDate)
        {
            string meetingDate = "     ";

            foreach (PdfString pdfString in listNDAandCAMDate)
            {
                string result = pdfString.ToString().Replace(" ", "");

                for (int i = 0; i < listPatternNDAandCAMDate.Count; i++)
                {
                    Match ma = (new Regex(listPatternNDAandCAMDate[i])).Match(result);

                    if (ma.Success)
                    {
                        meetingDate = FormatedResultFound(ma, i);

                        if (IsBiggerThanAnnDate(meetingDate, annDate))
                            return meetingDate;
                        else
                            meetingDate = "     ";
                    }
                }
            }

            return meetingDate;
        }

        private bool IsBiggerThanAnnDate(string meetingDate, string annDate)//meetingDate must bigger than annDate
        {
            bool result = false;

            if (string.IsNullOrWhiteSpace(meetingDate) || string.IsNullOrWhiteSpace(annDate))
                return result;

            int firstYear = System.Convert.ToInt32(meetingDate.Substring(6, 4));  //dd-mm-yyyy hh:mm:ss
            int firstMonth = System.Convert.ToInt32(meetingDate.Substring(3, 2));
            int firstDay = System.Convert.ToInt32(meetingDate.Substring(0, 2));
            int secondYear = System.Convert.ToInt32(annDate.Substring(6, 4));  //dd-mm-yyyy
            int secondMonth = System.Convert.ToInt32(annDate.Substring(3, 2));
            int secondDay = System.Convert.ToInt32(annDate.Substring(0, 2));

            if (firstYear > secondYear)
            {
                result = true;
            }
            else if (firstYear == secondYear)
            {
                if (firstMonth > secondMonth)
                {
                    result = true;
                }
                else if (firstMonth == secondMonth)
                {
                    if (firstDay >= secondDay)
                        result = true;
                }
            }

            return result;
        }

        private string FormatedResultFound(Match ma, int i)
        {
            string result = "     ";
            string year = string.Empty;
            string month = string.Empty;
            string day = string.Empty;
            string hour = string.Empty;
            string minute = string.Empty;
            string second = "00";
            year = ma.Groups["year"].Value;
            month = ma.Groups["month"].Value;
            day = ma.Groups["day"].Value;

            if (i == 0)         //"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})(\:|点|时)(?<minute>\d{1,2})");//0
            {
                hour = ma.Groups["hour"].Value;
                minute = ma.Groups["minute"].Value;
            }
            else if (i == 1)    //"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})(\:|点|时)(?<minute>\d{1,2})");   //1
            {
                hour = (System.Convert.ToInt32(ma.Groups["hour"].Value) + 12).ToString();
                minute = ma.Groups["minute"].Value;
            }
            else if (i == 2)    //"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})点半");//2
            {
                hour = ma.Groups["hour"].Value;
                minute = "30";
            }
            else if (i == 3)    //"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})点半");   //3
            {
                hour = (System.Convert.ToInt32(ma.Groups["hour"].Value) + 12).ToString();
                minute = "30";
            }
            else if (i == 4)    //"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})(点|时)");//4
            {
                hour = ma.Groups["hour"].Value;
                minute = "00";
            }
            else if (i == 5)    //"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)(星期|周)(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})(点|时)");   //5
            {
                hour = (System.Convert.ToInt32(ma.Groups["hour"].Value) + 12).ToString();
                minute = "00";
            }
            else              //"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日");                                             //default dd-mm-yyyy 00:00:00                      //6 
            {
                hour = "00";
                minute = "00";
            }

            month = month.Length == 1 ? "0" + month : month;
            day = day.Length == 1 ? "0" + day : day;
            hour = hour.Length == 1 ? "0" + hour : hour;
            minute = minute.Length == 1 ? "0" + minute : minute;
            result = string.Format("{0}-{1}-{2} {3}:{4}:{5}", day, month, year, hour, minute, second);

            return result;
        }

        private void FillCSVBody(Dictionary<string, List<string>> dicListBOD, List<string> listBODFiles)
        {
            int annId = 0;
            string annDate = string.Empty;
            string sourceDateLocal = string.Empty;
            string tickerA = string.Empty;
            string tickerB = string.Empty;

            if (listBODFiles == null || listBODFiles.Count == 0)
            {
                string msg = string.Format("no valid *.pdf in BOD_notice folder. ");
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            Dictionary<string, string> dicOthersFiles = GetOtherFilesInfo();

            foreach (var item in listBODFiles)
            {
                List<string> list = new List<string>();

                annId = GetIdFromPath(item);

                if (annId == 0)
                {
                    string msg = string.Format("can't get id of the path {0}", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    continue;
                }

                annDate = GetAnnouncementDate(annId);
                sourceDateLocal = GetSourceDateLocal(annId);

                list.Add(GetRic(annId));                       //RIC                        ace by id        
                list.Add(annDate);                             //Announcement date          ace by id
                list.Add(GetBODDate(item, annDate));           //BOD date                   pdf
                list.Add(GetCDIRate(item));                    //CDI rate                   pdf
                list.Add(GetSDIRate(item));                    //SDI rate                   pdf
                list.Add(GetSISRatio(item));                   //SIS ratio                  pdf
                list.Add(GetAnnual(item, dicOthersFiles));     //Annual/Semi-annual Report  Other folder
                list.Add(sourceDateLocal);                     //Source Date Local          ace by id
                list.Add(GetUrl(annId));                       //URL                        ace by id

                dicListBOD.Add(annId.ToString(), list);

                tickerA = GetTicker(annId);

                if (string.IsNullOrEmpty(tickerA) || !dicBShare.ContainsKey(tickerA))
                    continue;

                tickerB = dicBShare[tickerA];
                dicListBOD.Add(tickerB + dicListBOD.Count.ToString(), ReplaceTickerAByTickerB(tickerB, list));
            }
        }

        private List<string> ReplaceTickerAByTickerB(string tickerB, List<string> list)
        {
            List<string> result = new List<string>();

            try
            {
                result.Add(ConvertTickerToRic(tickerB));

                for (int i = 1; i < list.Count; i++)
                    result.Add(list[i]);

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string ConvertTickerToRic(string ticker)
        {
            string ric = string.Empty;

            if (string.IsNullOrEmpty(ticker))
                return ric;

            if (ticker.StartsWith("5") || ticker.StartsWith("6") || ticker.StartsWith("9"))
                ric = ticker + ".SS";
            else if (ticker.StartsWith("1") || ticker.StartsWith("00") || ticker.StartsWith("3"))
                ric = ticker + ".SZ";
            else
                ric = ticker + ".NotFound";

            return ric;
        }

        private string GetSourceDateLocal(int annId)
        {
            string sourceDateLocal = "     ";

            try
            {
                var chn = historicalAnnouncements.FirstOrDefault(p => p.ID == annId);

                if (chn == null)
                {
                    string msg = string.Format("can't found the record in ace db use id={0}", annId.ToString());
                    Logger.Log(msg, Logger.LogType.Error);
                    return sourceDateLocal;
                }

                sourceDateLocal = chn.CreateDate.ToString("dd-MMM-yyyy HH:mm:ss");
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return sourceDateLocal;
        }

        private string GetSourceDateGMT(int annId)
        {
            string sourceDateGMT = "     ";

            try
            {
                var chn = historicalAnnouncements.FirstOrDefault(p => p.ID == annId);

                if (chn == null)
                {
                    string msg = string.Format("can't found the record in ace db use id={0}", annId.ToString());
                    Logger.Log(msg, Logger.LogType.Error);
                    return sourceDateGMT;
                }

                sourceDateGMT = chn.CreateDate.AddHours(-8).ToString("dd-MMM-yyyy HH:mm:ss");
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return sourceDateGMT;
        }

        private Dictionary<string, string> GetOtherFilesInfo()
        {
            Dictionary<string, string> dicOtherFilesInfo = new Dictionary<string, string>();

            try
            {
                List<string> listOthersFolder = new List<string>() { classOthers };
                List<string> listOthersFiles = null;
                listOthersFiles = GetAllFiles(listOthersFolder);
                //string fileName = string.Empty;

                foreach (var item in listOthersFiles)
                {
                    //fileName = GetFileName(item);
                    if (string.IsNullOrEmpty(item) || dicOtherFilesInfo.ContainsKey(item))
                        continue;

                    dicOtherFilesInfo.Add(item, GetTicker(GetFileName(item)));
                }

                return dicOtherFilesInfo;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return dicOtherFilesInfo;
            }
        }

        private string GetAnnual(string item, Dictionary<string, string> dicOtherFiles)
        {
            string result = "No";

            try
            {
                if (dicOtherFiles == null || dicOtherFiles.Count == 0)
                    return result;

                string ticker = GetTicker(GetFileName(item));
                var otherCollection = dicOtherFiles.Where(p => p.Value == ticker).Select(p => p.Key).ToList();

                foreach (var title in otherCollection)
                {
                    if (IsAnnualReport(title))
                    {
                        result = "Yes";
                        break;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return result;
            }
        }

        private bool IsAnnualReport(string path)
        {
            bool result = false;
            PdfAnalyzer pa = null;
            PDFDoc doc = null;
            string pdfPattern = @"审\s*议.{0,30}年\s*(报|度\s*报\s*告)";

            try
            {
                if (!File.Exists(path))
                    return result;

                pa = new PdfAnalyzer();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(path);
                doc.InitSecurityHandler();
                List<PdfString> found = pa.RegexSearchAllPages(doc, pdfPattern);
                if (found != null && found.Count > 0)
                    result = true;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return result;
        }

        private string GetTicker(string name)
        {
            string ticker = string.Empty;
            string patternGetTicker = @"\d{2,4}-\d{1,2}-\d{1,2}_(?<Ticker>\d*)_";
            Match ma = (new Regex(patternGetTicker)).Match(name);

            if (ma.Success)
                ticker = ma.Groups["Ticker"].Value;

            return ticker;
        }

        private string GetUrl(int annId)
        {
            string announcementDate = "     ";
            try
            {
                var chn = historicalAnnouncements.FirstOrDefault(p => p.ID == annId);

                if (chn == null)
                {
                    string msg = string.Format("can't found the record in ace db use id={0}", annId.ToString());
                    Logger.Log(msg, Logger.LogType.Error);
                    return announcementDate;
                }

                announcementDate = chn.SourceLink;
                return announcementDate;

            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetAnnouncementDate(int annId) error. msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return announcementDate;
            }
        }

        private string GetSISRatio(string item)
        {
            string sisRatio = "00.000000";
            PDFDoc doc = null;
            try
            {
                PdfAnalyzer pa = new PdfAnalyzer();

                if (!File.Exists(item))
                {
                    string msg = string.Format("the file {0} is not exist.", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return sisRatio;
                }

                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(item);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return sisRatio;
                }

                List<PdfString> listPdfFirst = null;
                string patternCDIFirst = @"每\s*(10|十)\s*股\s*转\s*增\D*(?<money>[0-9\.]*)\s*股";    // 每 十股转增 x 股  //每 10 股转增 x 股
                listPdfFirst = pa.RegexSearchAllPages(doc, patternCDIFirst);
                decimal moneyTen;
                string patternMoney = @"转增(?<money>[0-9\.]+)股";

                if (listPdfFirst != null && listPdfFirst.Count >= 1)
                {
                    List<string> listMoneyTen = GetAllMoney(listPdfFirst, patternMoney, "money");

                    if (listMoneyTen.Count >= 2)
                        return string.Empty;

                    if (listMoneyTen.Count == 1)
                    {
                        if (decimal.TryParse(listMoneyTen[0], out moneyTen))
                            sisRatio = Math.Round(moneyTen, 6).ToString("f6");

                        return sisRatio;
                    }
                }

                List<PdfString> listPdfSecond = null;
                string patternCDISecond = @"每\s*(一|1)?\s*股\s*转\s*增\D*(?<money>[0-9\.]*)\s*股";    //每 股转增 x 股  //每 1 股转增 x 股 //每 一 股转增 x 股
                listPdfSecond = pa.RegexSearchAllPages(doc, patternCDISecond);
                decimal moneyOne;

                if (listPdfSecond != null && listPdfSecond.Count >= 1)
                {
                    List<string> listMoneyOne = GetAllMoney(listPdfSecond, patternMoney, "money");

                    if (listMoneyOne.Count >= 2)
                        return string.Empty;

                    if (listMoneyOne.Count == 1)
                    {
                        if (decimal.TryParse(listMoneyOne[0], out moneyOne))
                            sisRatio = Math.Round(moneyOne * 100, 6).ToString("f6");

                        return sisRatio;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetSISRatio(string item) error.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                doc.Dispose();
            }

            return sisRatio;
        }

        private string GetSDIRate(string item)
        {
            string sdiRate = "00.000000";
            PDFDoc doc = null;
            try
            {
                PdfAnalyzer pa = new PdfAnalyzer();

                if (!File.Exists(item))
                {
                    string msg = string.Format("the file {0} is not exist.", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return sdiRate;
                }


                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(item);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return sdiRate;
                }

                /*how do in this case.
                 * 现拟定本公司2013年度利润分配预案：每10股派发现金红利0.20元(含税)，不送股，也不实施资本公积金转增股本。该分配预案尚须提交2013年度股东大会审议通过。
                 * 同意公司以2013年12月31日的股本总数774,018,313股为基数，按每10股派送1.00元派发现金红利（含税），共分配现金股利77,401,831.30元，公司剩余未分配利润1,119,311,455.30元结转至下一年度。资本公积金不转增股本。
                 * 公司拟以2013年末总股本510,260,000股为基数，每10股派现金红利0.15元（含税），不转增，不送红股，共派现金红利7,653,900.00元, 公司（母公司）剩余未分配利润262,907,343.66元转入下一年度未分配利润。公司董事会认为该利润分配预案合法、合规。符合公司实际情况，
                 * 不以未分配利润派送红股，不以资本公积转增股本。剩余可供分配利
                 * 公司拟以2013年末总股本1,722,495,752股为基数，向全体股东按每10股派送现金红利2.30元（含税）
                 * 2013年不实施送股或资本公积转增股本。
                 * ，向全体股东每10股派送现金股利1.8元人民币（含税），计人民币15.84亿
                 */

                List<PdfString> listPdfFirst = null;
                //string patternCDIFirst = @"每\s*(10|十)\s*股\s*派?\s*送\D*(?<money>[0-9\.]*)\s*股";    //每 十股送 x 股  //每 10 股送 x 股
                string patternCDIFirst = @"每\s*(10|十)\s*股\s*送\D*(?<money>[0-9\.]*)\s*股";    //每 十股送 x 股  //每 10 股送 x 股
                listPdfFirst = pa.RegexSearchAllPages(doc, patternCDIFirst);
                decimal moneyTen;
                string patternMoney = @"送(?<money>[0-9\.]+)股";

                if (listPdfFirst != null && listPdfFirst.Count >= 1)
                {
                    List<string> listMoneyTen = GetAllMoney(listPdfFirst, patternMoney, "money");

                    if (listMoneyTen.Count >= 2)
                        return string.Empty;

                    if (listMoneyTen.Count == 1)
                    {
                        if (decimal.TryParse(listMoneyTen[0], out moneyTen))
                            sdiRate = Math.Round(moneyTen * 10, 6).ToString("f6");

                        return sdiRate;
                    }
                }

                List<PdfString> listPdfSecond = null;
                string patternCDISecond = @"每\s*(一|1)?\s*股\s*送\D*(?<money>[0-9\.]*)\s*股";    //每股送…x股  //每1股送…x股 //每一股送…x股
                listPdfSecond = pa.RegexSearchAllPages(doc, patternCDISecond);
                decimal moneyOne;

                if (listPdfSecond != null && listPdfSecond.Count >= 1)
                {
                    List<string> listMoneyOne = GetAllMoney(listPdfSecond, patternMoney, "money");

                    if (listMoneyOne.Count >= 2)
                        return string.Empty;

                    if (listMoneyOne.Count == 1)
                    {
                        if (decimal.TryParse(listMoneyOne[0], out moneyOne))
                            sdiRate = Math.Round(moneyOne * 100, 6).ToString("f6");

                        return sdiRate;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetSDIRate(string item) error.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                doc.Dispose();
            }

            return sdiRate;
        }

        private string GetCDIRate(string item)
        {
            string cdiRate = "     ";
            PDFDoc doc = null;
            try
            {
                PdfAnalyzer pa = new PdfAnalyzer();

                if (!File.Exists(item))
                {
                    string msg = string.Format("the file {0} is not exist.", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return cdiRate;
                }


                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(item);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return cdiRate;
                }

                #region [cdi rate =0]
                List<PdfString> listPdfFirst = null;
                List<string> listPatternCDIFirst = new List<string>() { 
                    "不分红", 
                    "不进行利润分配",
                    "不派发现金",
                    "不分配",
                    "不做利润分配",
                    "不进行现金分红",
                    "不进行现金股利分配"
                };

                foreach (var pattern in listPatternCDIFirst)
                {
                    listPdfFirst = pa.RegexSearchAllPages(doc, pattern);

                    if (listPdfFirst != null && listPdfFirst.Count >= 1)
                    {
                        cdiRate = "00.000000";
                        return cdiRate;
                    }
                }
                #endregion

                #region [cdi rate of ten]
                List<PdfString> listPdfSecond = null;
                string patternCDISecond = @"每\s*(10|十)\s*股\D*(?<money>[0-9\.]*)\s*元";
                listPdfSecond = pa.RegexSearchAllPages(doc, patternCDISecond);
                decimal moneyTen;
                string patternMoney = @"(?<money>[0-9\.]*)元";

                if (listPdfSecond != null && listPdfSecond.Count >= 1)
                {
                    List<string> listMoneyTen = GetAllMoney(listPdfSecond, patternMoney, "money");

                    if (listMoneyTen.Count >= 2)
                        return string.Empty;

                    if (listMoneyTen.Count == 1)
                    {
                        if (decimal.TryParse(listMoneyTen[0], out moneyTen))
                            cdiRate = Math.Round(moneyTen / 10, 6).ToString("f6");

                        return cdiRate;
                    }
                }
                #endregion

                #region [cdi rate of one]
                List<PdfString> listPdfThird = null;
                string patternCDIThird = @"每\s*(一|1)?\s*股\D*(?<money>[0-9\.]*)\s*元";
                listPdfThird = pa.RegexSearchAllPages(doc, patternCDIThird);
                decimal moneyOne;

                if (listPdfThird != null && listPdfThird.Count >= 1)
                {
                    List<string> listMoneyOne = GetAllMoney(listPdfThird, patternMoney, "money");

                    if (listMoneyOne.Count >= 2)
                        return string.Empty;

                    if (listMoneyOne.Count == 1)
                    {
                        if (decimal.TryParse(listMoneyOne[0], out moneyOne))
                            cdiRate = Math.Round(moneyOne, 6).ToString("f6");

                        return cdiRate;
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetCDIRate(string item) error.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                doc.Dispose();
            }

            return cdiRate;
        }

        private List<string> GetAllMoney(List<PdfString> listPdfSecond, string patternMoney, string groupName)
        {
            string value = string.Empty;
            List<string> listMoney = new List<string>();

            foreach (var item in listPdfSecond)
            {
                Match ma = (new Regex(patternMoney)).Match(item.ToString().Replace(" ", ""));
                if (!ma.Success)
                    continue;

                value = ma.Groups[groupName].Value;
                if (listMoney.Contains(value))
                    continue;

                listMoney.Add(value);
            }
            return listMoney;

        }

        private string GetBODDate(string item, string annDate)
        {
            string bodDate = "     ";//dd-mm-yyyy hh:mm:ss
            PDFDoc doc = null;
            try
            {
                //string patternBODDate = @"[^截止]\s*(\:|\：)?\s*(?<BODDate>\d{2,4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日.{15})";
                string patternBODDate = @"(?<BODDate>\d{2,4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日.{0,15})";
                List<string> listPatternBODDate = new List<string>();
                listPatternBODDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)星期(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})(\:|点|时)(?<minute>\d{1,2})");//0
                listPatternBODDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)星期(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})(\:|点|时)(?<minute>\d{1,2})");   //1

                listPatternBODDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)星期(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})点半");                       //2         //4
                listPatternBODDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)星期(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})点半");                          //3        //5

                listPatternBODDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)星期(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})(点|时)");                    //4       //2
                listPatternBODDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)星期(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})(点|时)");                       //5       //3

                //listPatternBODDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)星期(一|二|三|四|五|六|七)(\)|\）))?(上午)?(?<hour>\d{1,2})(点半|\:30)");                        //4
                //listPatternBODDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日((\(|\（)星期(一|二|三|四|五|六|七)(\)|\）))?下午(?<hour>\d{1,2})(点半|\:30)");                           //5

                listPatternBODDate.Add(@"(?<year>\d{2,4})年(?<month>\d{1,2})月(?<day>\d{1,2})日");                                             //default dd-mm-yyyy 00:00:00                      //6                                    

                if (!File.Exists(item))
                {
                    string msg = string.Format("the file {0} is not exist.", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return bodDate;
                }

                //PDFNet.Initialize();
                pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
                doc = new PDFDoc(item);
                doc.InitSecurityHandler();

                if (doc == null)
                {
                    string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", item);
                    Logger.Log(msg, Logger.LogType.Error);
                    return bodDate;
                }

                PdfAnalyzer pa = new PdfAnalyzer();
                List<PdfString> listBODDate = pa.RegexSearchAllPages(doc, patternBODDate);

                if (listBODDate == null || listBODDate.Count == 0)
                {
                    string msg = string.Format("can't found BODDate in {0} by keyWord {1}).", item, patternBODDate);
                    Logger.Log(msg, Logger.LogType.Error);
                    return bodDate;
                }

                string middleBodDate = GetBODDateSecond(listBODDate, listPatternBODDate, annDate);

                if (IsValidRangeDay(GetRangeOfDay(middleBodDate, annDate)))
                    bodDate = middleBodDate;

                return bodDate;
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetBODDate(string item, string annDate) error [pdfPath:{0}]. msg:{1}", item, ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return bodDate;
            }
            finally
            {
                doc.Dispose();
            }
        }

        private bool IsValidRangeDay(int dayRange)
        {
            bool InFourRange = false;

            if (dayRange >= 0 && dayRange <= 4)
                InFourRange = true;

            return InFourRange;
        }

        private int GetRangeOfDay(string smallDate, string bigDate)
        {
            int dayRange = -1;
            DateTime dtSmall;
            DateTime dtBig;

            if (DateTime.TryParse(FormatDateString(smallDate), out dtSmall) && DateTime.TryParse(FormatDateString(bigDate), out dtBig))
                dayRange = (dtBig - dtSmall).Days;

            return dayRange;
        }

        /*
         *[new requirement:6.24]confirm the "businese" from yaqiong...
         *Announcement Date - BOD Date < 4 business day (of stock exchange). 
         *If the grabbed date does not meet this condition, we leave the BOD date as blank.
         */
        private string GetBODDateSecond(List<PdfString> listBODDate, List<string> listPatternBODDate, string annDate)
        {
            string bodDate = "     ";
            string firstFound = string.Empty;
            string secondFound = string.Empty;
            string bodDateOne = listBODDate[0].ToString().Replace(" ", "");
            string bodDateTwo = string.Empty;

            firstFound = GetResultFound(listPatternBODDate, bodDateOne);

            if (listBODDate.Count > 1)//if exist two bod date extract second boddate
            {
                bodDateTwo = listBODDate[1].ToString().Replace(" ", "");

                if (bodDateTwo.Equals(bodDateOne) && listBODDate.Count > 2)
                    bodDateTwo = listBODDate[2].ToString().Replace(" ", "");//some times the listBODDate[0]===listBODDate[1]

                secondFound = GetResultFound(listPatternBODDate, bodDateTwo);
            }

            if (string.IsNullOrEmpty(secondFound))//only find one boddate in pdf
            {
                bodDate = FindSmallInTwoDate(firstFound, annDate);
                return bodDate;
            }

            string bigDate = FindBigInTwoDate(firstFound, secondFound);//find bigger date in the first two value in pdf
            bodDate = FindSmallInTwoDate(bigDate, annDate);//use the bigger value in pdf to compare with annDate in db

            return bodDate;
        }

        private string FindBigInTwoDate(string first, string second)
        {
            string result = first;

            int firstYear = System.Convert.ToInt32(first.Substring(6, 4));  //dd-mm-yyyy hh:mm:ss
            int firstMonth = System.Convert.ToInt32(first.Substring(3, 2));
            int firstDay = System.Convert.ToInt32(first.Substring(0, 2));

            int secondYear = System.Convert.ToInt32(second.Substring(6, 4));  //dd-mm-yyyy
            int secondMonth = System.Convert.ToInt32(second.Substring(3, 2));
            int secondDay = System.Convert.ToInt32(second.Substring(0, 2));

            if (firstYear < secondYear)
                result = second;
            else if (firstYear == secondYear)
            {
                if (firstMonth < secondMonth)
                    result = second;
                else if (firstMonth == secondMonth)
                {
                    if (firstDay <= secondDay)
                        result = second;
                }
            }
            return result;
        }

        private string FindSmallInTwoDate(string first, string second)
        {
            string result = "     ";

            int firstYear = System.Convert.ToInt32(first.Substring(6, 4));  //dd-mm-yyyy hh:mm:ss
            int firstMonth = System.Convert.ToInt32(first.Substring(3, 2));
            int firstDay = System.Convert.ToInt32(first.Substring(0, 2));

            int secondYear = System.Convert.ToInt32(second.Substring(6, 4));  //dd-mm-yyyy
            int secondMonth = System.Convert.ToInt32(second.Substring(3, 2));
            int secondDay = System.Convert.ToInt32(second.Substring(0, 2));

            if (firstYear == secondYear)
            {
                if (firstMonth < secondMonth)
                    result = first;
                else if (firstMonth == secondMonth)
                {
                    if (firstDay <= secondDay)
                        result = first;
                }
            }

            return result;
        }

        private string GetResultFound(List<string> listPatternBODDate, string bodDate)
        {
            string result = string.Empty;
            string year = string.Empty;
            string month = string.Empty;
            string day = string.Empty;
            string hour = string.Empty;
            string minute = string.Empty;
            string second = "00";

            for (int i = 0; i < listPatternBODDate.Count; i++)
            {
                Match ma = (new Regex(listPatternBODDate[i])).Match(bodDate);
                if (ma.Success)
                {
                    year = ma.Groups["year"].Value;
                    month = ma.Groups["month"].Value;
                    day = ma.Groups["day"].Value;

                    if (i == 0)         //?(上午)?(?<hour>\d{1,2})(\:|点|时)(?<minute>\d{1,2})");//0
                    {
                        hour = ma.Groups["hour"].Value;
                        minute = ma.Groups["minute"].Value;
                    }
                    else if (i == 1)    //?下午(?<hour>\d{1,2})(\:|点|时)(?<minute>\d{1,2})");   //1
                    {
                        hour = (System.Convert.ToInt32(ma.Groups["hour"].Value) + 12).ToString();
                        minute = ma.Groups["minute"].Value;
                    }
                    else if (i == 2)    //?(上午)?(?<hour>\d{1,2})点半");                        //2
                    {
                        hour = ma.Groups["hour"].Value;
                        minute = "30";
                    }
                    else if (i == 3)    //?下午(?<hour>\d{1,2})点半");                           //3
                    {
                        hour = (System.Convert.ToInt32(ma.Groups["hour"].Value) + 12).ToString();
                        minute = "30";
                    }
                    else if (i == 4)    //?(上午)?(?<hour>\d{1,2})(点|时)");                     //4
                    {
                        hour = ma.Groups["hour"].Value;
                        minute = "00";
                    }
                    else if (i == 5)    //?下午(?<hour>\d{1,2})(点|时)");                        //5
                    {
                        hour = (System.Convert.ToInt32(ma.Groups["hour"].Value) + 12).ToString();
                        minute = "00";
                    }
                    else              //default dd-mm-yyyy 00:00:00                              //6
                    {
                        hour = "00";
                        minute = "00";
                    }

                    year = year.Length == 2 ? "20" + year : year;
                    month = month.Length == 1 ? "0" + month : month;
                    day = day.Length == 1 ? "0" + day : day;
                    hour = hour.Length == 1 ? "0" + hour : hour;
                    minute = minute.Length == 1 ? "0" + minute : minute;
                    result = string.Format("{0}-{1}-{2} {3}:{4}:{5}", day, month, year, hour, minute, second);
                    break;
                }
            }

            return result;
        }

        private string GetAnnouncementDate(int annId)
        {
            string announcementDate = "     ";
            try
            {
                var chn = historicalAnnouncements.FirstOrDefault(p => p.ID == annId);

                if (chn == null)
                {
                    string msg = string.Format("can't found the record in ace db use id={0}", annId.ToString());
                    Logger.Log(msg, Logger.LogType.Error);
                    return announcementDate;
                }

                announcementDate = chn.AnnouceDate.ToString("dd-MM-yyyy");
                return announcementDate;

            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetAnnouncementDate(int annId) error. msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return announcementDate;
            }
        }

        private string GetTicker(int annId)
        {
            try
            {
                var chn = historicalAnnouncements.FirstOrDefault(p => p.ID == annId);

                if (chn == null)
                {
                    string msg = string.Format("can't found the record in ace db use id={0}", annId.ToString());
                    Logger.Log(msg, Logger.LogType.Error);

                    return null;
                }

                return chn.Ticker;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string GetRic(int annId)
        {
            string ric = "     ";
            try
            {
                var chn = historicalAnnouncements.FirstOrDefault(p => p.ID == annId);

                if (chn == null)
                {
                    string msg = string.Format("can't found the record in ace db use id={0}", annId.ToString());
                    Logger.Log(msg, Logger.LogType.Error);
                    return ric;
                }

                ric = chn.Ticker;

                if (ric.StartsWith("5") || ric.StartsWith("6") || ric.StartsWith("9"))
                    ric = ric + ".SS";
                else if (ric.StartsWith("1") || ric.StartsWith("00") || ric.StartsWith("3"))
                    ric = ric + ".SZ";
                else
                    ric = ric + ".NotFound";

                return ric;
            }
            catch (Exception ex)
            {
                string msg = string.Format("execute GetRic(int annId) error.msg:{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return ric;
            }
        }

        private List<string> GetAllFiles(List<string> listBODFoldet)
        {
            List<string> list = new List<string>();

            foreach (var item in listBODFoldet)
            {
                if (Directory.Exists(item))
                    list.AddRange(Directory.GetFiles(item, "*.PDF", SearchOption.TopDirectoryOnly).ToList());
            }

            return list;
        }

        private int GetIdFromPath(string path)
        {
            int id;
            string patternGetId = @".*_(?<ID>\d+).PDF";
            Match ma = (new Regex(patternGetId)).Match(path);

            if (ma.Success)
                if (int.TryParse(ma.Groups["ID"].Value, out id))
                    return id;

            return 0;
        }
        #endregion

        #region SecondJopUpdateToDB
        private void UpdateToAceDB(List<int> listOOSId)
        {
            List<CHNProcessItem> listCHNUpdate = new List<CHNProcessItem>();

            if (listOOSId == null || listOOSId.Count == 0)
            {
                string msg = string.Format("no records from loacl oos folder.");
                Logger.Log(msg, Logger.LogType.Info);
                return;
            }

            foreach (var item in listOOSId)
            {
                var query = string.Format("SELECT * FROM source WHERE Id = {0}", item);
                listCHNUpdate.AddRange(dealsContext.ExecuteQuery<CHNProcessItem>(query, 60).ToList());
            }

            foreach (var item in listCHNUpdate)
            {
                item.Comments = "OO-Scope by RG";
                //item.Status = "Completed";        //Note:  For testing period, do not change status to “Completed”. Add comments on ACE as “OO-Scope by RG”.
            }

            try
            {
                dealsContext.SubmitChanges();
                Logger.Log(string.Format("update {0} records to ace db succeed.", listCHNUpdate.Count), Logger.LogType.Info);
            }
            catch (Exception ex)
            {
                string msg = string.Format("update {0} records to ace db failed. msg:{1}", listCHNUpdate.Count, ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private List<int> GetOOSId(List<string> listOOSFolder)
        {
            List<int> list = new List<int>();
            List<string> listOOSPath = new List<string>();
            string patternGetId = @".*_(?<ID>\d+).PDF";
            Regex regex = new Regex(patternGetId);

            foreach (var item in listOOSFolder)
            {
                if (Directory.Exists(item))
                    listOOSPath.AddRange(Directory.GetFiles(item, "*.PDF", SearchOption.TopDirectoryOnly).ToList());
            }

            foreach (var item in listOOSPath)
            {
                MatchCollection matches = regex.Matches(item);
                string tmp = string.Empty;
                foreach (Match match in matches)
                {
                    list.Add(System.Convert.ToInt32(match.Groups["ID"].Value));
                }
            }

            return list;
        }
        #endregion

        #region FirstJopDownloadAndClassify
        private void OutPutDownloadErrorInfo(List<string> list)
        {
            if (list == null || list.Count == 0)
            {
                string msg = string.Format("all links download succeed!");
                Logger.Log(msg, Logger.LogType.Info);
                return;
            }

            GenerateInfo(list);
        }

        private void GenerateInfo(List<string> list)//downloadPath
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Links Of Download Error:\r\n");

            foreach (var item in list)
            {
                sb.AppendFormat("{0}\r\n", item);
            }

            if (!Directory.Exists(downloadPath))
                CreateFolder(downloadPath);

            File.WriteAllText(Path.Combine(downloadPath, "DownloadError.txt"), sb.ToString());
            AddResult(TaskName, Path.Combine(downloadPath, "DownloadError.txt"), "DownloadErrorLog");
        }

        #region [Classify Second Time]
        private void ClassifySecondTime()
        {
            if (Directory.Exists(classAGMNotice))
            {
                listAGMNotice = Directory.GetFiles(classAGMNotice, "*.PDF", SearchOption.TopDirectoryOnly).ToList();
                ClassifyAGMNotice(listAGMNotice);
            }

            if (Directory.Exists(classAGMResult))
            {
                listAGMResult = Directory.GetFiles(classAGMResult, "*.PDF", SearchOption.TopDirectoryOnly).ToList();
                ClassifyAGMResult(listAGMResult);
            }

            if (Directory.Exists(classBODNotice))
            {
                listBODNotice = Directory.GetFiles(classBODNotice, "*.PDF", SearchOption.TopDirectoryOnly).ToList();
                ClassifyBODNotice(listBODNotice);
            }

        }

        #region [BOD notice]
        private void ClassifyBODNotice(List<string> listBODNotice)
        {
            string path = string.Empty;
            string file = string.Empty;
            PDFDoc doc = null;
            List<string> listKeyWord = new List<string>();//key word for AGM Result In Scope folder
            listKeyWord.Add("利润分配");
            listKeyWord.Add("配股");
            listKeyWord.Add("股权分置改革");
            listKeyWord.Add("股改");
            listKeyWord.Add("缩股");

            foreach (var item in listBODNotice)
            {
                file = GetFileName(item);
                if (!File.Exists(item)) continue;

                if (IsExistKeyWord(item, listKeyWord, doc))
                {
                    path = Path.Combine(classBODNotice, "BOD_notice_In_Scope");
                    CreateFolder(path);
                    File.Copy(item, Path.Combine(path, file), true);
                }
                else
                {
                    path = Path.Combine(classBODNotice, "BOD_notice_OO_Scope");
                    CreateFolder(path);
                    File.Copy(item, Path.Combine(path, file), true);
                }

                if (File.Exists(item))
                    File.Delete(item);
            }
        }
        #endregion

        #region [AGM/EGM result]
        private void ClassifyAGMResult(List<string> listAGMResult)
        {
            string path = string.Empty;
            string file = string.Empty;
            PDFDoc doc = null;
            List<string> listKeyWord = new List<string>();//key word for AGM Result In Scope folder
            listKeyWord.Add("利润分配");
            listKeyWord.Add("配股");
            listKeyWord.Add("股权分置改革");
            listKeyWord.Add("股改");
            listKeyWord.Add("缩股");

            foreach (var item in listAGMResult)
            {
                file = GetFileName(item);
                if (!File.Exists(item)) continue;

                if (file.Contains("年度股东大会决议") || IsExistKeyWord(item, listKeyWord, doc))
                {
                    path = Path.Combine(classAGMResult, "AGM_EGM_result_In_Scope");
                    CreateFolder(path);
                    File.Copy(item, Path.Combine(path, file), true);
                }
                else
                {
                    path = Path.Combine(classAGMResult, "AGM_EGM_result_OO_Scope");
                    CreateFolder(path);
                    File.Copy(item, Path.Combine(path, file), true);
                }

                if (File.Exists(item))
                    File.Delete(item);
            }
        }

        private bool IsExistKeyWord(string item, List<string> listKeyWord, PDFDoc doc)
        {
            bool exist = false;

            if (!File.Exists(item))
            {
                string msg = string.Format("the file {0} is not exist.", item);
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception(msg);
            }

            //PDFNet.Initialize();
            pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
            doc = new PDFDoc(item);
            doc.InitSecurityHandler();

            if (doc == null)
            {
                string msg = string.Format("can't load pdf to doc = new PDFDoc({0}); ", item);
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception(msg);
            }

            PdfAnalyzer pa = new PdfAnalyzer();

            foreach (var word in listKeyWord)
            {
                if (pa.RegexSearchAllPages(doc, word).Count > 0)
                {
                    exist = true;
                    break;
                }
            }

            doc.Dispose();

            return exist;
        }

        #endregion

        #region [AGM/EGM notice]
        private void ClassifyAGMNotice(List<string> listAGMNotice)
        {
            string path = string.Empty;
            string file = string.Empty;
            PDFDoc doc = null;
            List<string> listKeyWord = new List<string>();//key word for AGM Result In Scope folder
            listKeyWord.Add("利润分配");
            listKeyWord.Add("配股");
            listKeyWord.Add("股权分置改革");
            listKeyWord.Add("股改");
            listKeyWord.Add("缩股");

            foreach (var item in listAGMNotice)
            {
                file = GetFileName(item);
                if (!File.Exists(item)) continue;

                if (IsExistKeyWord(item, listKeyWord, doc))
                {
                    path = Path.Combine(classAGMNotice, "AGM_EGM_notice_for_NDA_and_CAM");
                    CreateFolder(path);
                    File.Copy(item, Path.Combine(path, file), true);
                }
                else
                {
                    path = Path.Combine(classAGMNotice, "AGM_EGM_notice_for_CAM");
                    CreateFolder(path);
                    File.Copy(item, Path.Combine(path, file), true);
                }

                if (File.Exists(item))
                    File.Delete(item);
            }
        }
        #endregion
        #endregion

        #region [Classify First Time]
        private void FindByTitle(List<string> list)
        {
            if (list == null || list.Count == 0)
            {
                string msg = string.Format("no download file in local folder ,please check new record in ACE DB .");
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }

            foreach (var item in list)
            {
                File.Copy(item, FileTypeFolder(item), true);
            }
        }

        private List<string> DownloadFile(List<CHNProcessItem> historicalAnnouncements)
        {
            List<string> list = new List<string>();

            if (historicalAnnouncements == null || historicalAnnouncements.Count == 0)
            {
                string msg = string.Format("no historicalAnnouncements in Ace DB from dateTime.");
                Logger.Log(msg, Logger.LogType.Warning);
                return null;
            }

            foreach (var item in historicalAnnouncements)
            {
                DownloadPdf(item, list);
            }

            return list;
        }

        private string FileTypeFolder(string path)
        {
            string typeFolder = "Others";

            foreach (var item in dicTitlePattern.Keys)
            {
                if (Regex.IsMatch(GetFileName(path), item))
                {
                    typeFolder = dicTitlePattern[item];
                    break;
                }
            }

            typeFolder = Path.Combine(typePath, typeFolder);
            CreateFolder(typeFolder);

            return Path.Combine(typeFolder, GetFileName(path));
        }

        private string GetFileName(string path)
        {
            int start = path.LastIndexOf("\\");
            return path.Substring(start + 1, path.Length - start - 1);
        }

        private void DownloadPdf(CHNProcessItem item, List<string> list)
        {
            try
            {
                string strFileName = Path.Combine(downloadPath, string.Format("{0}_{1}_{2}_{3}_{4}.PDF", item.AnnouceDate.ToString("yyyy-MM-dd"), item.Ticker, FilterSpecial(item.Title), item.DealType, item.ID));
                RetryUtil.Retry(5, TimeSpan.FromSeconds(1), true, delegate
                {
                    if (!File.Exists(strFileName))
                        WebClientUtil.DownloadFile(item.SourceLink, 30000, strFileName);
                });

                if (list.Contains(strFileName))
                    return;

                list.Add(strFileName);
            }
            catch (Exception ex)
            {
                listDownloadPdfError.Add(item.SourceLink);
                string msg = string.Format("can't download file :{0}.msg :{1}", item.SourceLink, ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        public string FilterSpecial(string str)
        {
            string[] aryReg = { "#", "%", "&", "*", "|", "\\", ":", "\"", "<", ">", "?", "/" };

            for (int i = 0; i < aryReg.Length; i++)
            {
                str = str.Replace(aryReg[i], "");
            }

            return str;
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

        #endregion

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

        private void GetDownloadListFromDB(List<CHNProcessItem> historicalAnnouncements)
        {
            if (extendContext == null)
            {
                string msg = string.Format("DataSourceString is null .");
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                var query = string.Format("SELECT * FROM source WHERE SourcingEngine = '{0}' AND ScopeType = '{1}' AND AssignedTo = '{2}' AND Status = '{3}'", sourcingEngine, scopeType, assignedTo, status);
                //var query = string.Format("SELECT * FROM source WHERE SourcingEngine = '{0}' AND ScopeType = '{1}' AND AnnouceDate >= '{2}' Limit 0,1000", sourcingEngine, scopeType, startDateTime);
                historicalAnnouncements.AddRange(extendContext.ExecuteQuery<CHNProcessItem>(query, 60).ToList());
                Logger.Log(string.Format("{0} historical items loaded.", historicalAnnouncements.Count), Logger.LogType.Info);
            }
            catch (Exception ex)
            {
                string msg = string.Format("load historicalAnnouncements from Ace DB error. ex:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }
        #endregion
    }
}
