using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class HKSEAConfig
    {
        [Category("1.Today File")]
        [StoreInDB]
        [DisplayName("Source file path")]
        [Description("the formal xls file from an email.")]
        public string FormalSourceFilePath { get; set; }

        [Category("1.Today File")]
        [StoreInDB]
        [DisplayName("Source worksheet name")]
        public string FormalSourceFileWorksheetName { get; set; }

        [Category("2. Common File")]
        [StoreInDB]
        [DisplayName("HK SEA Z2 file path")]
        public string HkSeaZ2FilePath { get; set; }

        [Category("Holiday")]
        [DisplayName("Count")]
        public int HolidayCount { get; set; }

        [Category("5. Add-Drop File")]
        [StoreInDB]
        [DisplayName("Add Drop file path")]
        public string AddDropFilePath { get; set; }

        [Category("5. Add-Drop File")]
        [StoreInDB]
        [DisplayName("Add Drop raw worksheet name")]
        public string AddDropRawWorksheetName { get; set; }

        [Category("6. Template File")]
        [StoreInDB]
        [DisplayName("RIC Cows Template V2 file path")]
        public string RicCowsTemplateV2FilePath { get; set; }

        [Category("6. Template File")]
        [StoreInDB]
        [DisplayName("Template V2 worksheet name")]
        public string TemplateV2WorksheetName { get; set; }       

        [Category("3. Revised File")]
        [StoreInDB]
        [Description("the special xls file which comes from an email and consists of revised records.")]
        [DisplayName("Revised source file path")]
        public string RevisedSourceFilePath { get; set; }

        [Category("3. Revised File")]
        [StoreInDB]
        [DisplayName("Revised source worksheet name")]
        public string RevsiedSourceFileWorksheetName { get; set; }

        [Category("4. Canceled File")]
        [StoreInDB]
        [Description("the special xls file which comes from an email and consists of cancelled records.")]
        [DisplayName("Cancelled source file path")]
        public string CancelledSourceFilePath { get; set; }

        [Category("4. Canceled File")]
        [StoreInDB]
        [DisplayName("Cancelled source worksheet name")]
        public string CancelledSourceFileWorksheetName { get; set; } 

        public HKSEAConfig()
        {
            HolidayCount = 0;
        }
    }

    public class HKDelistingAndNameChangeConfig
    {
        public List<SectorConfig> SectorConfig { get; set; }

        public List<WarrantDelistingConfig> WarrantDelistingConfig { get; set; }

        public List<CBBCDelistingConfig> CBBCDelistingConfig { get; set; }

        public List<WarrantStartWith> WarrantStartWith { get; set; }

        public string OthersWarrantChainRic { get; set; }

        public string OthersCBBCChainRic { get; set; }

        public List<string> Hyphen { get; set; }

    }

    public class SectorConfig
    {
        public string Sector { get; set; }
        public string SectorChain { get; set; }
    }

    public class WarrantDelistingConfig
    {
        public string UnderlyingChain { get; set; }
        public string ChainRic { get; set; }
    }

    public class CBBCDelistingConfig
    {
        public string UnderlyingChain { get; set; }
        public string ChainRic { get; set; }
    }

    public class WarrantStartWith
    {
        public string StartWith { get; set; }
        public string ChainRic { get; set; }
    }

    public class BordLotSizeChangeInfo
    {
        public string Code;
        public string Name;
        public string EventType;
        public string BoardLotSize;
        public string Date;
    }

    public class BrokerChangeInfo : IComparable<BrokerChangeInfo>
    {
        public string Name { get; set; }
        public string Ric { get; set; }
        public string Type { get; set; }
        public string Date { get; set; }

        #region IComparable<BrokerChangeInfo> Members

        public int CompareTo(BrokerChangeInfo other)
        {
            return Date.CompareTo(other.Date);
        }

        #endregion
    }

    public class SpecialTradingInfo : IComparable<SpecialTradingInfo>
    {
        public string Ric { get; set; }
        public string Name { get; set; }
        public string BoardLotSize { get; set; }
        public string Date { get; set; }

        public int CompareTo(SpecialTradingInfo other)
        {
            return Date.CompareTo(other.Date);
        }
    }


    public class ExtraChangingInfo
    {
        public string ActionType { get; set; }
        public string Ric { get; set; }
        public string CurrentName { get; set; }
        public string OldName { get; set; }
        public string Event { get; set; }
        public string EffectiveDate { get; set; }
        public string NotificationDate { get; set; }
        public bool IsUsed { get; set; }
    }

    public class ChangeInfo
    {
        public List<BrokerChangeInfo> AllChangedBrokerList { get; set; }
        public List<SameDayChangeInfo> DifDayChangeInfo { get; set; }

        public List<SameDayChangeInfo> GetDifDayChangeInfo()
        {
            List<SameDayChangeInfo> sameDayChangeList = new List<SameDayChangeInfo>();
            SameDayChangeInfo sameDayChangeInfo = null;
            string date = string.Empty;
            for (int i = 0; i < AllChangedBrokerList.Count; i++)
            {
                BrokerChangeInfo broker = AllChangedBrokerList[i];
                if (broker.Date == date)
                {
                    if (broker.Type.ToLower() == "new")
                    {
                        sameDayChangeInfo.NewBrokerList.Add(broker);
                    }

                    else if (broker.Type.ToLower() == "delete")
                    {
                        sameDayChangeInfo.DeleteBrokerList.Add(broker);
                    }
                    else if (broker.Type.ToLower() == "change")
                    {
                        sameDayChangeInfo.BordLotSizeChangeList.Add(broker);
                    }

                    else if (broker.Type.ToLower() == "name chg to")
                    {
                        sameDayChangeInfo.NameChangeToList.Add(broker);
                    }
                }

                else
                {
                    if (i != 0)
                    {
                        sameDayChangeList.Add(sameDayChangeInfo);
                        sameDayChangeInfo.RicChangeToList = sameDayChangeInfo.GetChangeToList();
                    }
                    sameDayChangeInfo = new SameDayChangeInfo
                    {
                        Date = broker.Date,
                        NewBrokerList = new List<BrokerChangeInfo>(),
                        DeleteBrokerList = new List<BrokerChangeInfo>(),
                        NameChangeToList = new List<BrokerChangeInfo>(),
                        RicChangeToList = new List<StockNameRicChange>(),
                        BordLotSizeChangeList = new List<BrokerChangeInfo>()
                    };

                    date = broker.Date;

                    if (broker.Type.ToLower() == "new")
                    {
                        sameDayChangeInfo.NewBrokerList.Add(broker);
                    }
                    else if (broker.Type.ToLower() == "delete")
                    {
                        sameDayChangeInfo.DeleteBrokerList.Add(broker);
                    }
                    else if (broker.Type.ToLower() == "change")
                    {
                        sameDayChangeInfo.BordLotSizeChangeList.Add(broker);
                    }
                    else if (broker.Type.ToLower() == "name chg to")
                    {
                        sameDayChangeInfo.NameChangeToList.Add(broker);
                    }
                }
            }
            sameDayChangeList.Add(sameDayChangeInfo);
            return sameDayChangeList;
        }
    }

    public class SameDayChangeInfo : IComparable<SameDayChangeInfo>
    {
        public string Date { get; set; }
        public List<BrokerChangeInfo> NewBrokerList { get; set; }
        public List<BrokerChangeInfo> DeleteBrokerList { get; set; }
        public List<BrokerChangeInfo> NameChangeToList { get; set; }
        public List<StockNameRicChange> RicChangeToList { get; set; }
        public List<BrokerChangeInfo> BordLotSizeChangeList { get; set; }

        public List<StockNameRicChange> GetChangeToList()
        {
            foreach (BrokerChangeInfo broker in NewBrokerList)
            {
                for (int j = 0; j < DeleteBrokerList.Count; j++)
                {
                    if (broker.Name == DeleteBrokerList[j].Name)
                    {
                        StockNameRicChange stockRicChange = new StockNameRicChange
                        {
                            NewValue = broker.Ric,
                            Name = broker.Name,
                            OldValue = DeleteBrokerList[j].Ric,
                            Date = NewBrokerList[j].Date,
                            Ric = NewBrokerList[j].Ric,
                            Type = "RIC CHG TO"
                        };
                        RicChangeToList.Add(stockRicChange);
                    }
                }
            }
            return RicChangeToList;
        }

        public int CompareTo(SameDayChangeInfo y)
        {
            if (Date == null)
            {
                if (y.Date == null)
                    return 0;
                return -1;
            }
            if (y.Date == null)
                return 1;

            DateTime d1 = ParseDateTime(Date);
            DateTime d2 = ParseDateTime(y.Date);

            return d1.CompareTo(d2);
        }

        public DateTime ParseDateTime(string time)
        {
            DateTime dt;

            if (time.Contains("/"))
            {
                string[] timeItem = time.Split('/');
                time = timeItem[2] + "-" + timeItem[1] + "-" + timeItem[0];
            }

            if (!DateTime.TryParse(time, out dt))
            {
                throw new InvalidDataException(time);
            }
            return dt;
        }

    }

    public class StockNameRicChange
    {
        public string OldValue;
        public string NewValue;
        public string Ric;
        public string Name;
        public string Date;
        public string Type;
    }

    public class HKSEAUpdatorAndDelistingAndDrop : GeneratorBase
    {
        private static readonly string sectorConfigPath = ".\\Config\\HK\\HK_DelistingAndNameChange.xml";
        private HKSEAConfig configObj;
        private HKDelistingAndNameChangeConfig xmlConfigObj;

        private List<List<string>> cbbcList = new List<List<string>>();
        private List<List<string>> warrantList = new List<List<string>>();
        private List<List<string>> nameChangeList = new List<List<string>>();
        private List<List<List<string>>> apartCbbcList = new List<List<List<string>>>();
        private List<List<List<string>>> apartWarrantList = new List<List<List<string>>>();

        private List<DateTime> holidayList = new List<DateTime>();

        //{ '-', '(', ' ' };
        private char[] hyphen;
        private string underlyingStockCode = "";
        private string chainRic = "";
        private string oldShortName = "";
        private string oldLongName = "";
        private string sector = "";
        private string outputPath = "";

        //Backup formal file for StartDelistingAndNameChangeJob().
        private string formalFileBackupPath = string.Empty;

        protected override void Start()
        {
            try
            {
                StartSEAUpdaterJob();
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
                Logger.Log(ex.StackTrace);
            }
            try
            {
                StartHKEquityWarrantCBBCDropJob();
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
                Logger.Log(ex.StackTrace);
            }
            try
            {
                StartDelistingAndNameChangeJob();
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
                Logger.Log(ex.StackTrace);
            }
            AddResult("logger", Logger.FilePath, "log");
            //new TaskResultEntry("Log", "Log", Logger.FilePath));
        }

        protected override void Initialize()
        {
            configObj = Config as HKSEAConfig;
            xmlConfigObj = ConfigUtil.ReadConfig(sectorConfigPath, typeof(HKDelistingAndNameChangeConfig)) as HKDelistingAndNameChangeConfig;
            BackUpFormalFile();
            InitilizeHoliday();
        }

        private void InitilizeHoliday()
        {
            const string holidayTable = "ETI_Holiday";
            string where = string.Format("where DATEPART(yyyy,HolidayDate) = '{0}' and MarketId = 1", DateTime.Today.Year);

            System.Data.DataTable dt = ManagerBase.Select(holidayTable, new[] { "*" }, where);
            if (dt == null || dt.Rows.Count == 0)
            {
                return;
            }
            foreach (DataRow dr in dt.Rows)
            {
                holidayList.Add(Convert.ToDateTime(dr["HolidayDate"]));   
            }
        }

        private void BackUpFormalFile()
        {
            string dir = Path.GetDirectoryName(configObj.FormalSourceFilePath);
            dir = Path.Combine(dir, DateTime.Now.ToString("yyyy-MM-dd"));
            outputPath = dir;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }        

            string newFileName = Path.GetFileNameWithoutExtension(configObj.FormalSourceFilePath) + "_Updated.xls";
            formalFileBackupPath = Path.Combine(dir, newFileName);
            if (File.Exists(formalFileBackupPath))
            {
                File.Delete(formalFileBackupPath);
            }
            File.Copy(configObj.FormalSourceFilePath, formalFileBackupPath);

            TaskResultList.Add(new TaskResultEntry("Output Folder", "Output Folder", outputPath));
        }

        private void StartSEAUpdaterJob()
        {
            Logger.Log("Start the HKSEAUpdater Job......");
            UpdateFiles();
            Logger.Log("Finished the HKSEAUpdater Job......");
        }

        private void UpdateFiles()
        {
            int lastUsedRowNum = -1;
            List<BordLotSizeChangeInfo> bordLotSizeChangeList = new List<BordLotSizeChangeInfo>();

            //Preprocess the source file and get the parallel trading records.

            //Three Phase Parallel Trading Records
            List<SpecialTradingInfo> smallRicParallelTradingListOne = new List<SpecialTradingInfo>();
            List<SpecialTradingInfo> bigRicParallelTradingListOne = new List<SpecialTradingInfo>();

            //Two Phase Parallel Trading Records
            List<SpecialTradingInfo> smallRicParallelTradingListTwo = new List<SpecialTradingInfo>();
            List<SpecialTradingInfo> bigRicParallelTradingListTwo = new List<SpecialTradingInfo>();

            //The first type single special record
            List<SpecialTradingInfo> firstTypeSingleTradingList = new List<SpecialTradingInfo>();

            //Revised Trading Records
            List<ExtraChangingInfo> revisedRecordList = new List<ExtraChangingInfo>();

            //Cancelled Trading Records
            List<ExtraChangingInfo> cancelledRecordList = new List<ExtraChangingInfo>();

            GetCancelledRecord(configObj.CancelledSourceFilePath, out cancelledRecordList);
            GetRevisedRecord(configObj.RevisedSourceFilePath, out revisedRecordList);
            GetSpecialTradingInfo(formalFileBackupPath, out smallRicParallelTradingListOne, out bigRicParallelTradingListOne,
                out smallRicParallelTradingListTwo, out bigRicParallelTradingListTwo, out firstTypeSingleTradingList);

            using (ExcelApp app = new ExcelApp(false, false))
            {
                COMAddIns addIns = app.ExcelAppInstance.COMAddIns;
                COMAddIn addIn = null;
                foreach (COMAddIn item in addIns.Cast<COMAddIn>().Where(item => item.ProgId.Trim().Equals("PowerlinkCOMAddIn.COMAddIn") && item.Connect))
                {
                    Logger.Log("Disabled Eikon Excel.");
                    item.Connect = false;
                    addIn = item;
                }

                Logger.Log("Start to generate the Add-Drop file.***");
                string addDropFilePath = BackupFiles(configObj.AddDropFilePath);
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, addDropFilePath);
                Worksheet worksheetRaw = ExcelUtil.GetWorksheet(configObj.AddDropRawWorksheetName, workbook);
                using (ExcelApp appSource = new ExcelApp(false, false))
                {
                    var workbookSource = ExcelUtil.CreateOrOpenExcelFile(appSource, formalFileBackupPath);
                    Range AKRange = GetAKRange(workbookSource, out lastUsedRowNum, out bordLotSizeChangeList);
                    Range r = ExcelUtil.GetRange(1, 1, lastUsedRowNum, 11, worksheetRaw);
                    AKRange.Copy(Missing.Value);
                    r.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);
                    ExcelUtil.GetRange(1, 1, worksheetRaw).Copy(Missing.Value);
                    workbookSource.Close(false, workbookSource.FullName, false);
                }

                //Run Macros
                worksheetRaw.Activate();


                app.ExcelAppInstance.GetType().InvokeMember("Run",
                    BindingFlags.Default | BindingFlags.InvokeMethod,
                    null,
                    app.ExcelAppInstance,
                    new object[] { "Macro1" });

                var worksheet5 = ExcelUtil.GetWorksheet("Sheet5", workbook);
                if (worksheet5 == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet5 from workbook {0}", workbook.FullName));
                }

                UpdateSheet5(worksheet5, bordLotSizeChangeList);

                var worksheetAddDrop = ExcelUtil.GetWorksheet("ADD DROP", workbook);
                if (worksheetAddDrop == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet ADD DROP from workbook {0}", workbook.FullName));
                }
                UpdateAddDropSheet(worksheetAddDrop, bordLotSizeChangeList);
                int lastRowContainInfo = 2;
                while (ExcelUtil.GetRange(lastRowContainInfo, 1, worksheetAddDrop).Value2 != null)
                {
                    string value = (ExcelUtil.GetRange(lastRowContainInfo, 1, worksheetAddDrop).Value2.ToString().Trim().ToLower());
                    if (value.Contains("add") || (value.Contains("delete") || (value.Contains("change"))))
                    {
                        lastRowContainInfo++;
                    }
                    else
                        break;
                }

                ChangeInfo changedBrokerInfo = GetAllBrokerChangeInfo(worksheet5, lastRowContainInfo - 1);
                changedBrokerInfo.DifDayChangeInfo.Sort();
                UpdateHKSEAFile(changedBrokerInfo, bordLotSizeChangeList);
                UpdateHKSEAFileWithParallelTradeRecord(smallRicParallelTradingListOne, bigRicParallelTradingListOne, smallRicParallelTradingListTwo,
                    bigRicParallelTradingListTwo, firstTypeSingleTradingList);
                UpdateHKSEAFileWithCancelledRecord(cancelledRecordList);
                UpdateHKSEAFileWithRevisedRecord(revisedRecordList);

                string hkSEAFilePath;
                string date = DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US"));
                string hkSEAFileName = string.Format("HKSEA-Z2{0}.xls", date.ToUpper());
                string hkSEAFileDir = Path.GetDirectoryName(configObj.HkSeaZ2FilePath) + "\\" + DateTime.Today.ToString("yyyy-MM-dd");
                hkSEAFilePath = Path.Combine(hkSEAFileDir, hkSEAFileName);

                File.Copy(configObj.HkSeaZ2FilePath, hkSEAFilePath, true);
                TaskResultList.Add(new TaskResultEntry("HKSEA File", "Generate the HKSEA file", hkSEAFilePath));

                Range A2Range = ExcelUtil.GetRange(2, 1, lastRowContainInfo - 1, 16, worksheetAddDrop);
                UpdateTemplateV2File(A2Range, lastRowContainInfo);

                string newAddDropFilePath = Path.Combine(Path.GetDirectoryName(workbook.FullName), Path.GetFileName(configObj.AddDropFilePath));
                workbook.SaveCopyAs(newAddDropFilePath);
                workbook.Close(false, workbook.FullName, false);
                TaskResultList.Add(new TaskResultEntry("Add-Drop File", "Generate the new version of add-drop.xls file", newAddDropFilePath));
                File.Delete(addDropFilePath);
                Logger.Log("Finished generating the Add-Drop file.***");
                if (addIn != null)
                {
                    addIn.Connect = true;
                }
            }
        }

        private void GetRevisedRecord(string revisedXlsFilePath, out List<ExtraChangingInfo> revisedRecordList)
        {
            List<ExtraChangingInfo> recordList = new List<ExtraChangingInfo>();
            if (File.Exists(revisedXlsFilePath))
            {
                Logger.Log("Exists a revised record xls file, starting to get the revised records.");
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, revisedXlsFilePath);
                    var worksheet = ExcelUtil.GetWorksheet(configObj.RevsiedSourceFileWorksheetName, workbook);
                    if (worksheet == null)
                    {
                        Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.RevsiedSourceFileWorksheetName, configObj.RevisedSourceFilePath));
                    }
                    using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 16, 3, ExcelLineWriter.Direction.Right))
                    {
                        int endRowPos = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                        while (reader.Row <= endRowPos)
                        {
                            if (ExcelUtil.GetRange(reader.Row, 3, worksheet).Text != null && !string.IsNullOrEmpty(ExcelUtil.GetRange(reader.Row, 3, worksheet).Text.ToString()))
                            {
                                string ric = reader.ReadLineCellText();
                                ExtraChangingInfo extraTradingEntry = new ExtraChangingInfo {ActionType = "REVISED"};
                                int ricLength = ric.Length;

                                extraTradingEntry.Ric = string.Format("{0}.HK", ricLength <= 3 ? ric.PadLeft(4, '0') : ric);
                                extraTradingEntry.CurrentName = ExcelUtil.GetRange(reader.Row, 4, worksheet).Text.ToString();

                                string effectiveDate = ExcelUtil.GetRange(reader.Row, 9, worksheet).Text.ToString();
                                extraTradingEntry.EffectiveDate = effectiveDate;

                                string notificationDate = ExcelUtil.GetRange(reader.Row, 11, worksheet).Text.ToString();
                                extraTradingEntry.NotificationDate = notificationDate;

                                if (ExcelUtil.GetRange(reader.Row, 6, worksheet).Text != null && "Change of Name".Equals(ExcelUtil.GetRange(reader.Row, 6, worksheet).Text.ToString()))
                                {
                                    extraTradingEntry.Event = "NAME CHG TO";
                                }
                                else if (ExcelUtil.GetRange(reader.Row, 6, worksheet).Text != null && "New Listing".Equals(ExcelUtil.GetRange(reader.Row, 6, worksheet).Text.ToString()))
                                {
                                    extraTradingEntry.Event = "NEW";
                                }
                                else if (ExcelUtil.GetRange(reader.Row, 6, worksheet).Text != null && "Delisting".Equals(ExcelUtil.GetRange(reader.Row, 6, worksheet).Text.ToString()))
                                {
                                    extraTradingEntry.Event = "DELETE";
                                }
                                extraTradingEntry.IsUsed = false;
                                recordList.Add(extraTradingEntry);
                                reader.PlaceNext(reader.Row + 1, 3);
                            }
                            else
                            {
                                reader.PlaceNext(reader.Row + 1, 3);
                            }
                        }
                        revisedRecordList = recordList;
                    }
                }
                Logger.Log("Finished getting the revised records.");
            }
            else
            {
                revisedRecordList = recordList;
            }
        }
        //Exchange the date which like "2013-02-20" to the format looking like "20/02/2013".
        private string ExchangeDate(string initialDate)
        {
            string[] dateArray = initialDate.Split(new char[] { '-' });
            if (dateArray[1].Length == 1)
            {
                dateArray[1] = dateArray[1].PadLeft(2, '0');
            }
            if (dateArray[2].Length == 1)
            {
                dateArray[2] = dateArray[2].PadLeft(2, '0');
            }

            return string.Format("{0}/{1}/{2}", dateArray[2], dateArray[1], dateArray[0]);
        }

        private void GetCancelledRecord(string cancelledXlsFilePath, out List<ExtraChangingInfo> cancelledRecordList)
        {
            List<ExtraChangingInfo> recordList = new List<ExtraChangingInfo>();
            if (File.Exists(cancelledXlsFilePath))
            {
                Logger.Log("Exists a cancelled record xls file, starting to get the cancelled records.");
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, cancelledXlsFilePath);
                    var worksheet = ExcelUtil.GetWorksheet(configObj.CancelledSourceFileWorksheetName, workbook);
                    if (worksheet == null)
                    {
                        Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.CancelledSourceFileWorksheetName, configObj.CancelledSourceFilePath));
                    }
                    using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 16, 3, ExcelLineWriter.Direction.Right))
                    {
                        int endRowPos = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                        while (reader.Row <= endRowPos)
                        {
                            if (ExcelUtil.GetRange(reader.Row, 3, worksheet).Text != null && !string.IsNullOrEmpty(ExcelUtil.GetRange(reader.Row, 3, worksheet).Text.ToString()))
                            {
                                string ric = reader.ReadLineCellText();
                                ExtraChangingInfo extraTradingEntry = new ExtraChangingInfo();
                                extraTradingEntry.ActionType = "CANCELLED";
                                int ricLength = ric.Length;
                                extraTradingEntry.Ric = string.Format("{0}.HK", ricLength <= 3 ? ric.PadLeft(4, '0') : ric);
                                extraTradingEntry.CurrentName = ExcelUtil.GetRange(reader.Row, 4, worksheet).Text.ToString();
                                extraTradingEntry.EffectiveDate = ExcelUtil.GetRange(reader.Row, 9, worksheet).Text.ToString();

                                if (ExcelUtil.GetRange(reader.Row, 6, worksheet).Text != null && "Change of Name".Equals(ExcelUtil.GetRange(reader.Row, 6, worksheet).Text.ToString()))
                                {
                                    extraTradingEntry.Event = "NAME CHG TO";
                                }
                                else if (ExcelUtil.GetRange(reader.Row, 6, worksheet).Text != null && "New Listing".Equals(ExcelUtil.GetRange(reader.Row, 6, worksheet).Text.ToString()))
                                {
                                    extraTradingEntry.Event = "NEW";
                                }
                                else if (ExcelUtil.GetRange(reader.Row, 6, worksheet).Text != null && "Delisting".Equals(ExcelUtil.GetRange(reader.Row, 6, worksheet).Text.ToString()))
                                {
                                    extraTradingEntry.Event = "DELETE";
                                }
                                extraTradingEntry.IsUsed = false;
                                recordList.Add(extraTradingEntry);
                                reader.PlaceNext(reader.Row + 1, 3);
                            }
                            else
                            {
                                reader.PlaceNext(reader.Row + 1, 3);
                            }
                        }
                        cancelledRecordList = recordList;
                    }
                }
                Logger.Log("Finished getting the cancelled records.");
            }
            else
            {
                cancelledRecordList = recordList;
            }
        }

        private void GetSpecialTradingInfo(string xlsFilePath, out List<SpecialTradingInfo> smallRicParallelTradingListOne,
            out List<SpecialTradingInfo> bigRicParallelTradingListOne, out List<SpecialTradingInfo> smallRicParallelTradingListTwo,
            out List<SpecialTradingInfo> bigRicParallelTradingListTwo, out List<SpecialTradingInfo> firstTypeSingleTradingList)
        {
            List<SpecialTradingInfo> firstParallelTradingList = new List<SpecialTradingInfo>();
            List<SpecialTradingInfo> secondParallelTradingList = new List<SpecialTradingInfo>();
            List<SpecialTradingInfo> thirdParallelTradingList = new List<SpecialTradingInfo>();
            List<SpecialTradingInfo> forthParallelTradingList = new List<SpecialTradingInfo>();
            List<SpecialTradingInfo> firstSingleTradingList = new List<SpecialTradingInfo>();

            Dictionary<int, int> ricCountDictionary = new Dictionary<int, int>();

            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, xlsFilePath);
                var worksheet = ExcelUtil.GetWorksheet(configObj.FormalSourceFileWorksheetName, workbook);
                if (worksheet == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", xlsFilePath, configObj.FormalSourceFilePath));
                }

                //Check whether exists the special Trading Records.
                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 16, 3, ExcelLineWriter.Direction.Right))
                {
                    int endRowPos = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                    while (reader.Row <= endRowPos)
                    {
                        if (ExcelUtil.GetRange(reader.Row, 3, worksheet).Text != null && !string.IsNullOrEmpty(ExcelUtil.GetRange(reader.Row, 3, worksheet).Text.ToString()))
                        {
                            int ric = int.Parse(reader.ReadLineCellText());
                            if (!ricCountDictionary.ContainsKey(ric))
                            {
                                ricCountDictionary.Add(ric, 1);
                            }
                            else
                            {
                                ricCountDictionary[ric] += 1;
                            }
                            //Check whether exists the special record whose event is "Change of Board Lot Size"
                            if (ExcelUtil.GetRange(reader.Row, 6, worksheet).Text != null && "Change of Board Lot Size".Equals(ExcelUtil.GetRange(reader.Row, 6, worksheet).Text.ToString()))
                            {
                                SpecialTradingInfo specialTradingEntry = new SpecialTradingInfo();
                                int ricLength = ric.ToString().Length;

                                specialTradingEntry.Ric = ricLength <= 3 ? string.Format("{0}.HK", ric.ToString().PadLeft(4, '0')) : string.Format("{0}.HK", ric);
                                specialTradingEntry.Name = ExcelUtil.GetRange(reader.Row, 4, worksheet).Text.ToString();
                                specialTradingEntry.BoardLotSize = ExcelUtil.GetRange(reader.Row, 8, worksheet).Text.ToString();
                                specialTradingEntry.Date = ExcelUtil.GetRange(reader.Row, 9, worksheet).Text.ToString();
                                firstSingleTradingList.Add(specialTradingEntry);
                                RemoveRange(worksheet, reader.Row, 3, reader.Row, 10);
                            }

                            reader.PlaceNext(reader.Row + 1, 3);
                        }
                        else
                        {
                            reader.PlaceNext(reader.Row + 1, 3);
                        }
                    }

                    //Exists the situation in which there are Codes appearing 3 times.
                    if (ricCountDictionary.ContainsValue(3))
                    {
                        Logger.Log("Exists parallel trading records which change in three phase in the formal source xls file, starting to get the three phase parallel trading records.");
                        var query = from d in ricCountDictionary
                                    where d.Value == 3
                                    select d.Key;
                        List<int> threetimeRicList = query.ToList();
                        //Maybe exist many three phase Parallel Trading records.
                        int pairsRic = threetimeRicList.Count / 2;
                        List<int>[] specialRicListArray = new List<int>[pairsRic];
                        for (int i = 0; i < pairsRic; i++)
                        {
                            specialRicListArray[i] = new List<int>();
                            specialRicListArray[i].Add(threetimeRicList[i * 2]);
                            specialRicListArray[i].Add(threetimeRicList[i * 2 + 1]);
                        }
                        foreach (List<int> list in specialRicListArray)
                        {
                            list.Sort();
                        }

                        List<int> specialRicRow = new List<int>();
                        foreach (List<int> ricList in specialRicListArray)
                        {
                            reader.PlaceNext(16, 3);
                            int smallRicSum = 0, bigRicSum = 0;
                            while (reader.Row <= endRowPos && (smallRicSum < 3 || bigRicSum < 3))
                            {
                                if (ExcelUtil.GetRange(reader.Row, 3, worksheet).Text != null && !string.IsNullOrEmpty(ExcelUtil.GetRange(reader.Row, 3, worksheet).Text.ToString()))
                                {
                                    string ric = ExcelUtil.GetRange(reader.Row, 3, worksheet).Text.ToString();
                                    if (ric.Equals(ricList[0].ToString()))
                                    {
                                        SpecialTradingInfo parallelTradingEntry = new SpecialTradingInfo();
                                        int ricLength = ric.Length;
                                        parallelTradingEntry.Ric = string.Format("{0}.HK", ricLength <= 3 ? ric.PadLeft(4, '0') : ric);
                                        parallelTradingEntry.Name = ((Range)worksheet.Cells[reader.Row, 4]).Text.ToString();
                                        parallelTradingEntry.Date = ((Range)worksheet.Cells[reader.Row, 9]).Text.ToString();
                                        firstParallelTradingList.Add(parallelTradingEntry);
                                        smallRicSum++;
                                        specialRicRow.Add(reader.Row);

                                        reader.PlaceNext(reader.Row + 1, 3);
                                    }
                                    else if (ric.Equals(ricList[1].ToString()))
                                    {
                                        SpecialTradingInfo parallelTradingEntry = new SpecialTradingInfo();
                                        int ricLength = ric.Length;
                                        parallelTradingEntry.Ric = string.Format("{0}.HK", ricLength <= 3 ? ric.PadLeft(4, '0') : ric);
                                        parallelTradingEntry.Name = ((Range)worksheet.Cells[reader.Row, 4]).Text.ToString();
                                        parallelTradingEntry.Date = ((Range)worksheet.Cells[reader.Row, 9]).Text.ToString();
                                        secondParallelTradingList.Add(parallelTradingEntry);
                                        bigRicSum++;
                                        specialRicRow.Add(reader.Row);

                                        reader.PlaceNext(reader.Row + 1, 3);
                                    }
                                    else
                                    {
                                        reader.PlaceNext(reader.Row + 1, 3);
                                    }
                                }
                                else
                                {
                                    reader.PlaceNext(reader.Row + 1, 3);
                                }
                            }
                        }

                        //Remove the ThreePhase Parallel Trading Records from the xls file.
                        foreach (int ricRow in specialRicRow)
                        {
                            RemoveRange(worksheet, ricRow, 3, ricRow, 10);
                        }
                        Logger.Log("Finished getting the three phase parallel trading records.");
                    }

                    //Exists the situation in which there are Codes appearing 2 times.
                    if (ricCountDictionary.ContainsValue(2))
                    {
                        var query = from d in ricCountDictionary
                                    where d.Value == 2
                                    select d.Key;
                        List<int> twotimeRicList = query.ToList();

                        //Maybe exist many two phase Parallel Trading records.
                        int pairsRic = twotimeRicList.Count / 2;
                        if (pairsRic > 0)
                        {
                            Logger.Log("Exists parallel trading records which change in two phases in the formal source xls file, starting to get the two phase parallel trading records.");
                            List<int>[] specialRicListArray = new List<int>[pairsRic];
                            for (int i = 0; i < pairsRic; i++)
                            {
                                specialRicListArray[i] = new List<int> {twotimeRicList[i*2], twotimeRicList[i*2 + 1]};
                            }
                            foreach (List<int> ric in specialRicListArray)
                            {
                                ric.Sort();
                            }

                            List<int> specialRicRow = new List<int>();
                            foreach (List<int> specialRicList in specialRicListArray)
                            {
                                reader.PlaceNext(16, 3);
                                int smallRicSum = 0, bigRicSum = 0;
                                while (reader.Row <= endRowPos && (smallRicSum < 2 || bigRicSum < 2))
                                {
                                    if (ExcelUtil.GetRange(reader.Row, 3, worksheet).Text != null && !string.IsNullOrEmpty(ExcelUtil.GetRange(reader.Row, 3, worksheet).Text.ToString()))
                                    {
                                        string ric = ExcelUtil.GetRange(reader.Row, 3, worksheet).Text.ToString();
                                        if (ric.Equals(specialRicList[0].ToString()))
                                        {
                                            SpecialTradingInfo parallelTradingEntry = new SpecialTradingInfo();
                                            int ricLength = ric.Length;
                                            parallelTradingEntry.Ric = string.Format("{0}.HK", ricLength <= 3 ? ric.PadLeft(4, '0') : ric);
                                            parallelTradingEntry.Name = ((Range)worksheet.Cells[reader.Row, 4]).Text.ToString();
                                            parallelTradingEntry.BoardLotSize = ((Range)worksheet.Cells[reader.Row, 8]).Text == null ? "" : ((Range)worksheet.Cells[reader.Row, 8]).Text.ToString();
                                            parallelTradingEntry.Date = ((Range)worksheet.Cells[reader.Row, 9]).Text.ToString();
                                            thirdParallelTradingList.Add(parallelTradingEntry);
                                            smallRicSum++;
                                            specialRicRow.Add(reader.Row);

                                            reader.PlaceNext(reader.Row + 1, 3);
                                        }
                                        else if (ric.Equals(specialRicList[1].ToString()))
                                        {
                                            SpecialTradingInfo parallelTradingEntry = new SpecialTradingInfo();
                                            int ricLength = ric.Length;
                                            parallelTradingEntry.Ric = string.Format("{0}.HK", ricLength <= 3 ? ric.PadLeft(4, '0') : ric);
                                            parallelTradingEntry.Name = ((Range)worksheet.Cells[reader.Row, 4]).Text.ToString();
                                            parallelTradingEntry.Date = ((Range)worksheet.Cells[reader.Row, 9]).Text.ToString();
                                            forthParallelTradingList.Add(parallelTradingEntry);
                                            bigRicSum++;
                                            specialRicRow.Add(reader.Row);

                                            reader.PlaceNext(reader.Row + 1, 3);
                                        }
                                        else
                                        {
                                            reader.PlaceNext(reader.Row + 1, 3);
                                        }
                                    }
                                    else
                                    {
                                        reader.PlaceNext(reader.Row + 1, 3);
                                    }
                                }
                            }

                            //Remove the TwoPhase Parallel Trading Records from the xls file.
                            foreach (int ricRow in specialRicRow)
                            {
                                RemoveRange(worksheet, ricRow, 3, ricRow, 10);
                            }
                            Logger.Log("Finished getting the two phase parallel trading records.");
                        }
                    }
                }

                smallRicParallelTradingListOne = firstParallelTradingList;
                bigRicParallelTradingListOne = secondParallelTradingList;
                smallRicParallelTradingListTwo = thirdParallelTradingList;
                bigRicParallelTradingListTwo = forthParallelTradingList;
                firstTypeSingleTradingList = firstSingleTradingList;
                if (firstTypeSingleTradingList.Count > 0)
                {
                    Logger.Log("Exists the special single line record whose event is \"Change of Board Lot Size\" ,and finished getting such records.");
                }
                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbook.Save();
                TaskResultList.Add(new TaskResultEntry("Updated Source File", "Remove three phase parallel trading records from source file", xlsFilePath));
                workbook.Close();
            }
        }

        private void RemoveRange(Worksheet worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            Range rangeToRemove = ExcelUtil.GetRange(startRow, startCol, endRow, endCol, worksheet);
            rangeToRemove.Clear();
        }

        public void UpdateHKSEAFile(ChangeInfo ChangedBrokerInfo, List<BordLotSizeChangeInfo> bordLotSizeChangeList)
        {
            using (ExcelApp appSEA = new ExcelApp(false, false))
            {
                string seaFilePath = BackupFiles(configObj.HkSeaZ2FilePath);
                var workbookSEA = ExcelUtil.CreateOrOpenExcelFile(appSEA, seaFilePath);
                var worksheetSEA = ExcelUtil.GetWorksheet("Sheet2", workbookSEA);
                int rowPos;
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheetSEA, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    int curRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                    for (int i = ChangedBrokerInfo.DifDayChangeInfo.Count - 1; i > -1; i--)
                    {
                        SameDayChangeInfo sameDayChangedInfo = ChangedBrokerInfo.DifDayChangeInfo[i];
                        rowPos = curRowPos;
                        while (rowPos > 0)
                        {
                            //底部空行需要删掉
                            Range range1 = ExcelUtil.GetRange(rowPos, 4, worksheetSEA);

                            object nameCell = ExcelUtil.GetRange(rowPos, 1, worksheetSEA).Value2;
                            object ricCell = ExcelUtil.GetRange(rowPos, 2, worksheetSEA).Value2;
                            object actionCell = ExcelUtil.GetRange(rowPos, 3, worksheetSEA).Value2;
                            object dateCell = ExcelUtil.GetRange(rowPos, 4, worksheetSEA).Value2;

                            string name = string.Empty;
                            string ric = string.Empty;
                            string action = string.Empty;
                            string date = string.Empty;

                            if (nameCell != null)
                            {
                                name = nameCell.ToString();
                            }
                            if (ricCell != null)
                            {
                                ric = ricCell.ToString();
                            }
                            if (actionCell != null)
                            {
                                action = actionCell.ToString();
                            }
                            if (dateCell != null)
                            {
                                date = dateCell.ToString();
                            }

                            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(ric) && string.IsNullOrEmpty(action) && string.IsNullOrEmpty(date))
                            {
                                rowPos--;
                                curRowPos = rowPos;
                                continue;
                            }

                            if (ExcelUtil.GetRange(rowPos, 4, worksheetSEA).Value2 == null || string.IsNullOrEmpty(ExcelUtil.GetRange(rowPos, 4, worksheetSEA).Value2.ToString()))
                            {
                                rowPos--;
                                continue;
                            }
                            if (DateTimeCompare(sameDayChangedInfo.Date, ExcelUtil.GetRange(rowPos, 4, worksheetSEA).Text.ToString()))
                            {
                                if (rowPos == curRowPos)
                                {
                                    curRowPos = rowPos;
                                    InsertChangeInfo(rowPos + 1, writer, sameDayChangedInfo, worksheetSEA, bordLotSizeChangeList);
                                    break;
                                }

                                if (range1 == null || (string.IsNullOrEmpty(range1.Text.ToString().Trim())))
                                {
                                    rowPos--;
                                }
                                else if (!DateTimeCompare(sameDayChangedInfo.Date, range1.Text.ToString().Trim()))
                                {
                                    curRowPos = rowPos;
                                    InsertChangeInfo(rowPos + 1, writer, sameDayChangedInfo, worksheetSEA, bordLotSizeChangeList);
                                    break;
                                }

                                else
                                {
                                    string s = range1.Text.ToString();
                                    rowPos--;
                                }
                            }

                            else
                            {
                                if (rowPos == 1)
                                {
                                    curRowPos = rowPos;
                                    InsertChangeInfo(rowPos + 1, writer, sameDayChangedInfo, worksheetSEA, bordLotSizeChangeList);
                                    break;
                                }

                                if (ExcelUtil.GetRange(rowPos - 1, 4, worksheetSEA) == null || (ExcelUtil.GetRange(rowPos - 1, 4, worksheetSEA).Text == null))
                                {
                                    rowPos--;
                                }
                                else if (DateTimeCompare(sameDayChangedInfo.Date, ExcelUtil.GetRange(rowPos - 1, 4, worksheetSEA).Text.ToString()))
                                {
                                    curRowPos = rowPos;
                                    InsertChangeInfo(rowPos, writer, sameDayChangedInfo, worksheetSEA, bordLotSizeChangeList);
                                    break;
                                }
                                else
                                {
                                    rowPos--;
                                }
                            }
                        }
                    }
                }
                workbookSEA.SaveCopyAs(configObj.HkSeaZ2FilePath);
                workbookSEA.Close(false, workbookSEA.FullName, false);
                File.Delete(seaFilePath);
            }
        }

        private void UpdateHKSEAFileWithRevisedRecord(List<ExtraChangingInfo> revisedRecordList)
        {
            if (revisedRecordList.Count > 0)
            {
                Logger.Log("Start to update HKSEA File with the revised records.");
                using (ExcelApp appSEA = new ExcelApp(false, false))
                {
                    string seaFilePath = BackupFiles(configObj.HkSeaZ2FilePath);
                    var workbookSEA = ExcelUtil.CreateOrOpenExcelFile(appSEA, seaFilePath);
                    var worksheetSEA = ExcelUtil.GetWorksheet("Sheet2", workbookSEA);
                    int lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                    int curRowPos = lastRowPos;
                    using (ExcelLineWriter writer = new ExcelLineWriter(worksheetSEA, 1, 1, ExcelLineWriter.Direction.Right))
                    {
                        //Process the records whose Event Type is "NEW" or "DELETE" or "NAME CHG TO". 
                        foreach (ExtraChangingInfo revisedRecord in revisedRecordList)
                        {
                            if ("NAME CHG TO".Equals(revisedRecord.Event))
                            {
                                while (curRowPos >= 1)
                                {

                                    if (revisedRecord.CurrentName.Equals(ExcelUtil.GetRange(curRowPos + 1, 1, worksheetSEA).Text.ToString()) &&
                                        revisedRecord.Ric.Equals(ExcelUtil.GetRange(curRowPos, 2, worksheetSEA).Text.ToString()) &&
                                        "NAME CHG TO".Equals(ExcelUtil.GetRange(curRowPos, 3, worksheetSEA).Text.ToString()))
                                    {
                                        string date = ExcelUtil.GetRange(curRowPos, 4, worksheetSEA).Text.ToString();
                                        if (!string.IsNullOrEmpty(date))
                                        {
                                            if (DateTimeLessThan(revisedRecord.NotificationDate, date))
                                            {
                                                List<string> insertList = new List<string>(8);
                                                for (int row = curRowPos; row <= curRowPos + 1; row++)
                                                {
                                                    for (int col = 1; col <= 3; col++)
                                                    {
                                                        string value = ExcelUtil.GetRange(row, col, worksheetSEA).Text.ToString();
                                                        value = string.IsNullOrEmpty(value) ? "" : value;
                                                        insertList.Add(value);
                                                    }
                                                    insertList.Add(row == curRowPos ? revisedRecord.EffectiveDate : "");
                                                }
                                                object color = ExcelUtil.GetRange(curRowPos, 1, curRowPos + 1, 4, worksheetSEA).Interior.Color;
                                                ExcelUtil.GetRange(curRowPos, 1, curRowPos + 1, 4, worksheetSEA).Delete();
                                                lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                                                revisedRecord.IsUsed = true;

                                                InsertTwoLineRecord(lastRowPos, writer, worksheetSEA, insertList, 3, color);
                                                break;
                                            }
                                            curRowPos--;
                                        }
                                        else
                                        {
                                            date = ExcelUtil.GetRange(curRowPos + 1, 4, worksheetSEA).Text.ToString();
                                            if (DateTimeLessThan(revisedRecord.NotificationDate, date))
                                            {
                                                List<string> insertList = new List<string>(8);
                                                for (int row = curRowPos; row <= curRowPos + 1; row++)
                                                {
                                                    for (int col = 1; col <= 3; col++)
                                                    {
                                                        string value = ExcelUtil.GetRange(row, col, worksheetSEA).Text.ToString();
                                                        value = string.IsNullOrEmpty(value) ? "" : value;
                                                        insertList.Add(value);
                                                    }
                                                    insertList.Add(row == curRowPos ? "" : revisedRecord.EffectiveDate);
                                                }
                                                object color = ExcelUtil.GetRange(curRowPos, 1, curRowPos + 1, 4, worksheetSEA).Interior.Color;
                                                ExcelUtil.GetRange(curRowPos, 1, curRowPos + 1, 4, worksheetSEA).Delete();
                                                lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                                                revisedRecord.IsUsed = true;

                                                InsertTwoLineRecord(lastRowPos, writer, worksheetSEA, insertList, 7, color);
                                                break;
                                            }
                                            curRowPos--;
                                        }
                                    }
                                    else
                                    {
                                        curRowPos--;
                                    }
                                }
                                lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                                curRowPos = lastRowPos;
                            }
                            else
                            {
                                //Process the records whose event is "NEW" or "DELETE", but not "RIC CHG TO".
                                while (curRowPos >= 1)
                                {
                                    if (revisedRecord.CurrentName.Equals(ExcelUtil.GetRange(curRowPos, 1, worksheetSEA).Text.ToString()) &&
                                        revisedRecord.Ric.Equals(ExcelUtil.GetRange(curRowPos, 2, worksheetSEA).Text.ToString()) &&
                                        revisedRecord.Event.Equals(ExcelUtil.GetRange(curRowPos, 3, worksheetSEA).Text.ToString()) &&
                                        DateTimeLessThan(revisedRecord.NotificationDate, ExcelUtil.GetRange(curRowPos, 4, worksheetSEA).Text.ToString()))
                                    {
                                        List<string> insertList = new List<string>(4);
                                        for (int col = 1; col <= 3; col++)
                                        {
                                            string value = ExcelUtil.GetRange(curRowPos, col, worksheetSEA).Text.ToString();
                                            insertList.Add(value);
                                        }
                                        insertList.Add(revisedRecord.EffectiveDate);

                                        object color = ExcelUtil.GetRange(curRowPos, 1, curRowPos, 4, worksheetSEA).Interior.Color;
                                        ExcelUtil.GetRange(curRowPos, 1, curRowPos, 4, worksheetSEA).Delete();
                                        lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                                        revisedRecord.IsUsed = true;

                                        InsertOneLineRecord(lastRowPos, writer, worksheetSEA, insertList, 3, color);
                                        break;
                                    }
                                    string date = ExcelUtil.GetRange(curRowPos, 4, worksheetSEA).Text.ToString();
                                    if (!string.IsNullOrEmpty(date) && !DateTimeLessThan(revisedRecord.NotificationDate, date))
                                    {
                                        break;
                                    }
                                    curRowPos--;
                                }
                                lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                                curRowPos = lastRowPos;
                            }
                        }

                        List<ExtraChangingInfo> ricCHGToList = revisedRecordList.Where(revisedRecord => !revisedRecord.IsUsed).ToList();

                        //Exists the records which means a "RIC CHG TO" event.
                        if (ricCHGToList.Count > 0)
                        {
                            List<ExtraChangingInfo>[] ricCHGToListArray = new List<ExtraChangingInfo>[ricCHGToList.Count / 2];
                            for (int i = 0; i < ricCHGToList.Count / 2; i++)
                            {
                                ricCHGToListArray[i] = new List<ExtraChangingInfo>
                                {
                                    ricCHGToList[i*2],
                                    ricCHGToList[i*2 + 1]
                                };
                            }

                            lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                            curRowPos = lastRowPos;
                            //Process the records whose event is "RIC CHG TO".
                            foreach (List<ExtraChangingInfo> infoList in ricCHGToListArray)
                            {
                                while (curRowPos >= 1)
                                {
                                    if (infoList[0].CurrentName.Equals(infoList[1].CurrentName) &&
                                        infoList[0].CurrentName.Equals(ExcelUtil.GetRange(curRowPos, 1, worksheetSEA).Text.ToString()) &&
                                        "RIC CHG TO".Equals(ExcelUtil.GetRange(curRowPos, 3, worksheetSEA).Text.ToString()) &&
                                        DateTimeLessThan(infoList[0].NotificationDate, ExcelUtil.GetRange(curRowPos, 4, worksheetSEA).Text.ToString()))
                                    {
                                        List<string> insertList = new List<string>(8);
                                        for (int row = curRowPos; row <= curRowPos + 1; row++)
                                        {
                                            for (int col = 1; col <= 3; col++)
                                            {
                                                string value = ExcelUtil.GetRange(row, col, worksheetSEA).Text.ToString();
                                                value = string.IsNullOrEmpty(value) ? "" : value;
                                                insertList.Add(value);
                                            }
                                            insertList.Add(row == curRowPos ? infoList[0].EffectiveDate : "");
                                        }
                                        object color = ExcelUtil.GetRange(curRowPos, 1, curRowPos + 1, 4, worksheetSEA).Interior.Color;
                                        ExcelUtil.GetRange(curRowPos, 1, curRowPos + 1, 4, worksheetSEA).Delete();
                                        lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;

                                        InsertTwoLineRecord(lastRowPos, writer, worksheetSEA, insertList, 3, color);
                                        break;
                                    }
                                    curRowPos--;
                                }

                                lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                                curRowPos = lastRowPos;
                            }
                        }
                    }
                    workbookSEA.SaveCopyAs(configObj.HkSeaZ2FilePath);
                    workbookSEA.Close(false, workbookSEA.FullName, false);
                    File.Delete(seaFilePath);
                }
                Logger.Log("Finished updating HKSEA File with the revised records.");
            }
        }

        private void UpdateHKSEAFileWithCancelledRecord(List<ExtraChangingInfo> cancelledRecordList)
        {
            if (cancelledRecordList.Count > 0)
            {
                Logger.Log("Start to update HKSEA File with the cancelled records.");
                using (ExcelApp appSEA = new ExcelApp(false, false))
                {
                    string seaFilePath = BackupFiles(configObj.HkSeaZ2FilePath);
                    var workbookSEA = ExcelUtil.CreateOrOpenExcelFile(appSEA, seaFilePath);
                    var worksheetSEA = ExcelUtil.GetWorksheet("Sheet2", workbookSEA);
                    int lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                    int curRowPos = lastRowPos;
                    using (ExcelLineWriter reader = new ExcelLineWriter(worksheetSEA, 1, 1, ExcelLineWriter.Direction.Right))
                    {
                        foreach (ExtraChangingInfo cancelledRecord in cancelledRecordList)
                        {
                            if ("NAME CHG TO".Equals(cancelledRecord.Event))
                            {
                                while (curRowPos >= 1)
                                {
                                    if (cancelledRecord.CurrentName.Equals(ExcelUtil.GetRange(curRowPos + 1, 1, worksheetSEA).Text.ToString()) &&
                                        cancelledRecord.Ric.Equals(ExcelUtil.GetRange(curRowPos, 2, worksheetSEA).Text.ToString()) &&
                                        "NAME CHG TO".Equals(ExcelUtil.GetRange(curRowPos, 3, worksheetSEA).Text.ToString()) &&
                                        (cancelledRecord.EffectiveDate.Equals(ExcelUtil.GetRange(curRowPos, 4, worksheetSEA).Text.ToString()) ||
                                        cancelledRecord.EffectiveDate.Equals(ExcelUtil.GetRange(curRowPos + 1, 4, worksheetSEA).Text.ToString())))
                                    {
                                        ExcelUtil.GetRange(curRowPos, 1, curRowPos + 1, 4, worksheetSEA).Delete();

                                        cancelledRecord.IsUsed = true;
                                        break;
                                    }
                                    curRowPos--;
                                }
                                lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                                curRowPos = lastRowPos;
                            }
                            else
                            {
                                while (curRowPos >= 1)
                                {
                                    if (cancelledRecord.CurrentName.Equals(ExcelUtil.GetRange(curRowPos, 1, worksheetSEA).Text.ToString()) &&
                                        cancelledRecord.Ric.Equals(ExcelUtil.GetRange(curRowPos, 2, worksheetSEA).Text.ToString()) &&
                                        cancelledRecord.Event.Equals(ExcelUtil.GetRange(curRowPos, 3, worksheetSEA).Text.ToString()) &&
                                        cancelledRecord.EffectiveDate.Equals(ExcelUtil.GetRange(curRowPos, 4, worksheetSEA).Text.ToString()))
                                    {
                                        ExcelUtil.GetRange(curRowPos, 1, curRowPos, 4, worksheetSEA).Delete();

                                        cancelledRecord.IsUsed = true;
                                        break;
                                    }
                                    string date = ExcelUtil.GetRange(curRowPos, 4, worksheetSEA).Text.ToString();
                                    if (!string.IsNullOrEmpty(date) && !DateTimeLessThan(cancelledRecord.EffectiveDate, date))
                                    {
                                        break;
                                    }
                                    curRowPos--;
                                }
                                lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                                curRowPos = lastRowPos;
                            }
                        }

                        List<ExtraChangingInfo> ricCHGToList = cancelledRecordList.Where(cancelledRecord => !cancelledRecord.IsUsed).ToList();

                        //Exists the records which means a "RIC CHG TO" event.
                        if (ricCHGToList.Count > 0)
                        {
                            List<ExtraChangingInfo>[] ricCHGToListArray = new List<ExtraChangingInfo>[ricCHGToList.Count / 2];
                            for (int i = 0; i < ricCHGToList.Count / 2; i++)
                            {
                                ricCHGToListArray[i] = new List<ExtraChangingInfo>
                                {
                                    ricCHGToList[i*2],
                                    ricCHGToList[i*2 + 1]
                                };
                            }

                            lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                            curRowPos = lastRowPos;
                            foreach (List<ExtraChangingInfo> infoList in ricCHGToListArray)
                            {
                                while (curRowPos >= 1)
                                {
                                    if (infoList[0].CurrentName.Equals(infoList[1].CurrentName) &&
                                        infoList[0].CurrentName.Equals(ExcelUtil.GetRange(curRowPos, 1, worksheetSEA).Text.ToString()) &&
                                        infoList[0].EffectiveDate.Equals(ExcelUtil.GetRange(curRowPos, 4, worksheetSEA).Text.ToString()) &&
                                        "RIC CHG TO".Equals(ExcelUtil.GetRange(curRowPos, 3, worksheetSEA).Text.ToString()))
                                    {
                                        ExcelUtil.GetRange(curRowPos, 1, curRowPos + 1, 4, worksheetSEA).Delete();

                                        curRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                                        break;
                                    }
                                    curRowPos--;
                                }
                            }
                        }
                    }
                    workbookSEA.SaveCopyAs(configObj.HkSeaZ2FilePath);
                    workbookSEA.Close(false, workbookSEA.FullName, false);
                    File.Delete(seaFilePath);
                }
                Logger.Log("Finished updating HKSEA File with the cancelled records.");
            }
        }

        private void UpdateHKSEAFileWithParallelTradeRecord(List<SpecialTradingInfo> smallRicListOne, List<SpecialTradingInfo> bigRicListOne,
            List<SpecialTradingInfo> smallRicListTwo, List<SpecialTradingInfo> bigRicListTwo, List<SpecialTradingInfo> firstTypeSingleTradingList)
        {
            Logger.Log("Start to update HKSEA File with the parallel trading records.");
            using (ExcelApp appSEA = new ExcelApp(false, false))
            {
                string seaFilePath = BackupFiles(configObj.HkSeaZ2FilePath);
                var workbookSEA = ExcelUtil.CreateOrOpenExcelFile(appSEA, seaFilePath);
                var worksheetSEA = ExcelUtil.GetWorksheet("Sheet2", workbookSEA);
                int lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheetSEA, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    for (int i = 0; i < smallRicListOne.Count / 3; i++)
                    {
                        InsertFirstPhase3ParallelTradingInfo(lastRowPos, writer, worksheetSEA, smallRicListOne[i * 3], bigRicListOne[i * 3]);
                        lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;

                        InsertSecondPhase3ParallelTradingInfo(lastRowPos, writer, worksheetSEA, smallRicListOne[i * 3 + 1], bigRicListOne[i * 3 + 1], bigRicListOne[i * 3].Name);
                        lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;

                        InsertThirdPhase3ParallelTradingInfo(lastRowPos, writer, worksheetSEA, smallRicListOne[i * 3 + 2], smallRicListOne[i * 3 + 1].Name, bigRicListOne[i * 3 + 2]);
                        lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                    }
                    for (int i = 0; i < smallRicListTwo.Count / 2; i++)
                    {
                        InsertFirshPhase2ParallelTradingInfo(lastRowPos, writer, worksheetSEA, smallRicListTwo[i * 2], smallRicListTwo[i * 2 + 1].Name, bigRicListTwo[i * 2]);
                        lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;

                        InsertSecondPhase2ParallelTradingInfo(lastRowPos, writer, worksheetSEA, smallRicListTwo[i * 2 + 1], smallRicListTwo[i * 2].Name, bigRicListTwo[i * 2 + 1]);
                        lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                    }
                    foreach (SpecialTradingInfo info in firstTypeSingleTradingList)
                    {
                        InsertFirstTypeSingleTradingInfo(lastRowPos, writer, worksheetSEA, info);
                        lastRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                    }
                }
                workbookSEA.SaveCopyAs(configObj.HkSeaZ2FilePath);
                workbookSEA.Close(false, workbookSEA.FullName, false);
                File.Delete(seaFilePath);
            }
            Logger.Log("Finished updating HKSEA File with the parallel trading records.");
        }

        private void InsertOneLineRecord(int lastRowPos, ExcelLineWriter writer, Worksheet worksheet, List<string> insertList, int effectiveDateIndex, object color)
        {
            int curRowPos = lastRowPos;
            while (curRowPos >= 1)
            {
                if (ExcelUtil.GetRange(curRowPos, 4, worksheet).Text == null || string.IsNullOrEmpty(ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    curRowPos--;
                    continue;
                }
                if (DateTimeLessThan(insertList[effectiveDateIndex], ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    if (curRowPos == 1)
                    {
                        ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(1, 1, worksheet), 1);
                        writer.PlaceNext(1, 1);
                        writer.WriteLine(insertList[0]);
                        writer.WriteLine(insertList[1]);
                        writer.WriteLine(insertList[2]);
                        writer.WriteLine(insertList[3]);
                        ExcelUtil.GetRange(1, 1, 1, 4, worksheet).Interior.Color = color;
                        break;
                    }
                    curRowPos--;
                    continue;
                }
                ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos + 1, 1, worksheet), 1);
                writer.PlaceNext(curRowPos + 1, 1);
                writer.WriteLine(insertList[0]);
                writer.WriteLine(insertList[1]);
                writer.WriteLine(insertList[2]);
                writer.WriteLine(insertList[3]);
                ExcelUtil.GetRange(curRowPos + 1, 1, curRowPos + 1, 4, worksheet).Interior.Color = color;
                break;
            }
        }

        private void InsertTwoLineRecord(int lastRowPos, ExcelLineWriter writer, Worksheet worksheet, List<string> insertList, int effectiveDateIndex, object color)
        {
            int curRowPos = lastRowPos;
            while (curRowPos >= 1)
            {
                if (ExcelUtil.GetRange(curRowPos, 4, worksheet).Text == null || string.IsNullOrEmpty(ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    curRowPos--;
                    continue;
                }
                if (DateTimeLessThan(insertList[effectiveDateIndex], ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    if (curRowPos == 1)
                    {
                        ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(1, 1, worksheet), 2);
                        writer.PlaceNext(1, 1);
                        writer.WriteLine(insertList[0]);
                        writer.WriteLine(insertList[1]);
                        writer.WriteLine(insertList[2]);
                        writer.WriteLine(insertList[3]);

                        writer.PlaceNext(2, 1);
                        writer.WriteLine(insertList[4]);
                        writer.WriteLine(insertList[5]);
                        writer.WriteLine(insertList[6]);
                        writer.WriteLine(insertList[7]);
                        ExcelUtil.GetRange(1, 1, 2, 4, worksheet).Interior.Color = color;
                        break;
                    }
                    curRowPos--;
                    continue;
                }
                ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos + 1, 1, worksheet), 2);
                writer.PlaceNext(curRowPos + 1, 1);
                writer.WriteLine(insertList[0]);
                writer.WriteLine(insertList[1]);
                writer.WriteLine(insertList[2]);
                writer.WriteLine(insertList[3]);

                writer.PlaceNext(curRowPos + 2, 1);
                writer.WriteLine(insertList[4]);
                writer.WriteLine(insertList[5]);
                writer.WriteLine(insertList[6]);
                writer.WriteLine(insertList[7]);
                ExcelUtil.GetRange(curRowPos + 1, 1, curRowPos + 2, 4, worksheet).Interior.Color = color;
                break;
            }
        }

        private void InsertFirstPhase3ParallelTradingInfo(int lastRowPos, ExcelLineWriter writer, Worksheet worksheet, SpecialTradingInfo smallRicRecord, SpecialTradingInfo bigRicRecord)
        {
            int curRowPos = lastRowPos;

            while (curRowPos >= 1)
            {
                if (ExcelUtil.GetRange(curRowPos, 4, worksheet).Text == null || string.IsNullOrEmpty(ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    curRowPos--;
                    continue;
                }
                if (DateTimeLessThan(smallRicRecord.Date, ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    if (curRowPos == 1)
                    {
                        ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos, 1, worksheet), 2);
                        writer.PlaceNext(1, 1);
                        writer.WriteLine(smallRicRecord.Name);
                        writer.WriteLine(smallRicRecord.Ric);
                        writer.WriteLine("RIC CHG TO");
                        writer.WriteLine(smallRicRecord.Date);

                        writer.PlaceNext(2, 2);
                        writer.WriteLine(bigRicRecord.Ric);
                        writer.PlaceNext(2, 4);
                        writer.WriteLine(bigRicRecord.Date);
                        ExcelUtil.GetRange(1, 1, 2, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(146, 208, 80).ToArgb();
                        break;
                    }
                    curRowPos--;
                    continue;
                }
                ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos + 1, 1, worksheet), 2);
                writer.PlaceNext(curRowPos + 1, 1);
                writer.WriteLine(smallRicRecord.Name);
                writer.WriteLine(smallRicRecord.Ric);
                writer.WriteLine("RIC CHG TO");
                writer.WriteLine(smallRicRecord.Date);

                writer.PlaceNext(curRowPos + 2, 2);
                writer.WriteLine(bigRicRecord.Ric);
                writer.PlaceNext(curRowPos + 2, 4);
                writer.WriteLine(bigRicRecord.Date);
                ExcelUtil.GetRange(curRowPos + 1, 1, curRowPos + 2, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(146, 208, 80).ToArgb();
                break;
            }
        }

        private void InsertSecondPhase3ParallelTradingInfo(int lastRowPos, ExcelLineWriter writer, Worksheet worksheet, SpecialTradingInfo smallRicRecord, SpecialTradingInfo bigRicRecord, string bigRicRecordOldName)
        {
            int curRowPos = lastRowPos;

            while (curRowPos >= 1)
            {
                if (ExcelUtil.GetRange(curRowPos, 4, worksheet).Text == null || string.IsNullOrEmpty(ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    curRowPos--;
                    continue;
                }
                if (DateTimeLessThan(smallRicRecord.Date, ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    if (curRowPos == 1)
                    {
                        ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(1, 1, worksheet), 3);
                        writer.PlaceNext(1, 1);
                        writer.WriteLine(smallRicRecord.Name);
                        writer.WriteLine(smallRicRecord.Ric);
                        writer.WriteLine("NEW");
                        writer.WriteLine(smallRicRecord.Date);

                        writer.PlaceNext(2, 1);
                        writer.WriteLine(bigRicRecordOldName);
                        writer.WriteLine(bigRicRecord.Ric);
                        writer.WriteLine("NAME CHG TO");

                        writer.PlaceNext(3, 1);
                        writer.WriteLine(bigRicRecord.Name);
                        writer.PlaceNext(3, 4);
                        writer.WriteLine(bigRicRecord.Date);
                        ExcelUtil.GetRange(1, 1, 3, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(146, 208, 80).ToArgb();
                        break;
                    }
                    curRowPos--;
                    continue;
                }
                ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos + 1, 1, worksheet), 3);
                writer.PlaceNext(curRowPos + 1, 1);
                writer.WriteLine(smallRicRecord.Name);
                writer.WriteLine(smallRicRecord.Ric);
                writer.WriteLine("NEW");
                writer.WriteLine(smallRicRecord.Date);

                writer.PlaceNext(curRowPos + 2, 1);
                writer.WriteLine(bigRicRecordOldName);
                writer.WriteLine(bigRicRecord.Ric);
                writer.WriteLine("NAME CHG TO");

                writer.PlaceNext(curRowPos + 3, 1);
                writer.WriteLine(bigRicRecord.Name);
                writer.PlaceNext(curRowPos + 3, 4);
                writer.WriteLine(bigRicRecord.Date);
                ExcelUtil.GetRange(curRowPos + 1, 1, curRowPos + 3, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(146, 208, 80).ToArgb();
                break;
            }
        }

        private void InsertThirdPhase3ParallelTradingInfo(int lastRowPos, ExcelLineWriter writer, Worksheet worksheet, SpecialTradingInfo smallRicRecord, string smallRicRecordOldName, SpecialTradingInfo bigRicRecord)
        {
            int curRowPos = lastRowPos;

            while (curRowPos >= 1)
            {
                if (ExcelUtil.GetRange(curRowPos, 4, worksheet).Text == null || string.IsNullOrEmpty(ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    curRowPos--;
                    continue;
                }
                if (DateTimeLessThan(smallRicRecord.Date, ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    if (curRowPos == 1)
                    {
                        ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(1, 1, worksheet), 3);
                        writer.PlaceNext(1, 1);
                        writer.WriteLine(smallRicRecordOldName);
                        writer.WriteLine(smallRicRecord.Ric);
                        writer.WriteLine("NAME CHG TO");
                        writer.WriteLine(smallRicRecord.Date);

                        writer.PlaceNext(2, 1);
                        writer.WriteLine(smallRicRecord.Name);

                        writer.PlaceNext(3, 1);
                        writer.WriteLine(bigRicRecord.Name);
                        writer.WriteLine(bigRicRecord.Ric);
                        writer.WriteLine("DELETE");
                        writer.WriteLine(bigRicRecord.Date);
                        ExcelUtil.GetRange(1, 1, 3, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(146, 208, 80).ToArgb();
                        break;
                    }
                    curRowPos--;
                    continue;
                }
                ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos + 1, 1, worksheet), 3);
                writer.PlaceNext(curRowPos + 1, 1);
                writer.WriteLine(smallRicRecordOldName);
                writer.WriteLine(smallRicRecord.Ric);
                writer.WriteLine("NAME CHG TO");
                writer.WriteLine(smallRicRecord.Date);

                writer.PlaceNext(curRowPos + 2, 1);
                writer.WriteLine(smallRicRecord.Name); ;

                writer.PlaceNext(curRowPos + 3, 1);
                writer.WriteLine(bigRicRecord.Name);
                writer.WriteLine(bigRicRecord.Ric);
                writer.WriteLine("DELETE");
                writer.WriteLine(bigRicRecord.Date);
                ExcelUtil.GetRange(curRowPos + 1, 1, curRowPos + 3, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(146, 208, 80).ToArgb();
                break;
            }
        }

        private void InsertFirshPhase2ParallelTradingInfo(int lastRowPos, ExcelLineWriter writer, Worksheet worksheet, SpecialTradingInfo smallRicRecord, string smallRicOldName, SpecialTradingInfo bigRicRecord)
        {
            int curRowPos = lastRowPos;

            while (curRowPos >= 1)
            {
                if (ExcelUtil.GetRange(curRowPos, 4, worksheet).Text == null || string.IsNullOrEmpty(ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    curRowPos--;
                    continue;
                }
                if (DateTimeLessThan(smallRicRecord.Date, ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    if (curRowPos == 1)
                    {
                        ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos, 1, worksheet), 3);
                        writer.PlaceNext(1, 1);
                        writer.WriteLine(smallRicOldName);
                        writer.WriteLine(smallRicRecord.Ric);
                        writer.WriteLine("NAME CHG TO");
                        writer.WriteLine(smallRicRecord.Date);

                        writer.PlaceNext(2, 1);
                        writer.WriteLine(smallRicRecord.Name);
                        writer.PlaceNext(2, 3);
                        writer.WriteLine(string.Format("LOT CHG {0}", smallRicRecord.BoardLotSize));

                        writer.PlaceNext(3, 1);
                        writer.WriteLine(bigRicRecord.Name);
                        writer.WriteLine(bigRicRecord.Ric);
                        writer.WriteLine("NEW");
                        writer.WriteLine(bigRicRecord.Date);
                        ExcelUtil.GetRange(1, 1, 3, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(149, 179, 215).ToArgb();
                        break;
                    }
                    curRowPos--;
                    continue;
                }
                ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos + 1, 1, worksheet), 3);
                writer.PlaceNext(curRowPos + 1, 1);
                writer.WriteLine(smallRicOldName);
                writer.WriteLine(smallRicRecord.Ric);
                writer.WriteLine("NAME CHG TO");
                writer.WriteLine(smallRicRecord.Date);

                writer.PlaceNext(curRowPos + 2, 1);
                writer.WriteLine(smallRicRecord.Name);
                writer.PlaceNext(curRowPos + 2, 3);
                writer.WriteLine(string.Format("LOT CHG {0}", smallRicRecord.BoardLotSize));

                writer.PlaceNext(curRowPos + 3, 1);
                writer.WriteLine(bigRicRecord.Name);
                writer.WriteLine(bigRicRecord.Ric);
                writer.WriteLine("NEW");
                writer.WriteLine(bigRicRecord.Date);
                ExcelUtil.GetRange(curRowPos + 1, 1, curRowPos + 3, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(149, 179, 215).ToArgb();
                break;
            }
        }

        private void InsertSecondPhase2ParallelTradingInfo(int lastRowPos, ExcelLineWriter writer, Worksheet worksheet, SpecialTradingInfo smallRicRecord, string smallRicOldName, SpecialTradingInfo bigRicRecord)
        {
            int curRowPos = lastRowPos;

            while (curRowPos >= 1)
            {
                if (ExcelUtil.GetRange(curRowPos, 4, worksheet).Text == null || string.IsNullOrEmpty(ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    curRowPos--;
                    continue;
                }
                if (DateTimeLessThan(smallRicRecord.Date, ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    if (curRowPos == 1)
                    {
                        ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos, 1, worksheet), 3);
                        writer.PlaceNext(1, 1);
                        writer.WriteLine(smallRicOldName);
                        writer.WriteLine(smallRicRecord.Ric);
                        writer.WriteLine("NAME CHG TO");
                        writer.WriteLine(smallRicRecord.Date);

                        writer.PlaceNext(2, 1);
                        writer.WriteLine(smallRicRecord.Name);
                        writer.PlaceNext(2, 4);
                        writer.WriteLine(smallRicRecord.Date);

                        writer.PlaceNext(3, 1);
                        writer.WriteLine(bigRicRecord.Name);
                        writer.WriteLine(bigRicRecord.Ric);
                        writer.WriteLine("DELETE");
                        writer.WriteLine(bigRicRecord.Date);
                        ExcelUtil.GetRange(1, 1, 3, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(149, 179, 215).ToArgb();
                        break;
                    }
                    curRowPos--;
                    continue;
                }
                ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos + 1, 1, worksheet), 3);
                writer.PlaceNext(curRowPos + 1, 1);
                writer.WriteLine(smallRicOldName);
                writer.WriteLine(smallRicRecord.Ric);
                writer.WriteLine("NAME CHG TO");
                writer.WriteLine(smallRicRecord.Date);

                writer.PlaceNext(curRowPos + 2, 1);
                writer.WriteLine(smallRicRecord.Name);
                writer.PlaceNext(curRowPos + 2, 4);
                writer.WriteLine(smallRicRecord.Date);

                writer.PlaceNext(curRowPos + 3, 1);
                writer.WriteLine(bigRicRecord.Name);
                writer.WriteLine(bigRicRecord.Ric);
                writer.WriteLine("DELETE");
                writer.WriteLine(bigRicRecord.Date);
                ExcelUtil.GetRange(curRowPos + 1, 1, curRowPos + 3, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(149, 179, 215).ToArgb();
                break;
            }
        }

        private void InsertFirstTypeSingleTradingInfo(int lastRowPos, ExcelLineWriter writer, Worksheet worksheet, SpecialTradingInfo firstTypeSingleRecord)
        {
            int curRowPos = lastRowPos;

            while (curRowPos >= 1)
            {
                if (ExcelUtil.GetRange(curRowPos, 4, worksheet).Text == null || string.IsNullOrEmpty(ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    curRowPos--;
                    continue;
                }
                if (DateTimeLessThan(firstTypeSingleRecord.Date, ExcelUtil.GetRange(curRowPos, 4, worksheet).Text.ToString()))
                {
                    if (curRowPos == 1)
                    {
                        ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos, 1, worksheet), 1);
                        writer.PlaceNext(1, 1);
                        writer.WriteLine(firstTypeSingleRecord.Name);
                        writer.WriteLine(firstTypeSingleRecord.Ric);
                        writer.WriteLine(string.Format("Board Lot Size Change to {0}", firstTypeSingleRecord.BoardLotSize));
                        writer.WriteLine(firstTypeSingleRecord.Date);
                        ExcelUtil.GetRange(1, 1, 1, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(250, 200, 144).ToArgb();
                        break;
                    }
                    curRowPos--;
                    continue;
                }
                ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(curRowPos + 1, 1, worksheet), 1);
                writer.PlaceNext(curRowPos + 1, 1);
                writer.WriteLine(firstTypeSingleRecord.Name);
                writer.WriteLine(firstTypeSingleRecord.Ric);
                writer.WriteLine(string.Format("Board Lot Size Change to {0}", firstTypeSingleRecord.BoardLotSize));
                writer.WriteLine(firstTypeSingleRecord.Date);
                ExcelUtil.GetRange(curRowPos + 1, 1, curRowPos + 1, 4, worksheet).Interior.Color = System.Drawing.Color.FromArgb(250, 200, 144).ToArgb();
                break;
            }
        }

        private void StartHKEquityWarrantCBBCDropJob()
        {
            Logger.Log("Start the HKEquityWarrantCBBCDrop Job......");
            Logger.Log("Start to generate the GEDA Drop file.");
            List<string> mainRicList = GetMainRic(configObj.HkSeaZ2FilePath);
            List<string> miRicList = new List<string>();
            string miRic;
            foreach (string[] mainRicParts in mainRicList.Select(mainRic => mainRic.Split(".".ToCharArray())))
            {
                miRic = string.Format("{0}MI.HK", mainRicParts[0]);
                miRicList.Add(miRic);
            }
            GenerateGEDAFiles(mainRicList, miRicList);
            Logger.Log("Finished generating the GEDA Drop file.");
            Logger.Log("Finished the HKEquityWarrantCBBCDrop Job......");
        }

        private List<string> GetMainRic(string sourceFilePath)
        {
            List<string> mainRicList = new List<string>();
            using (ExcelApp app = new ExcelApp(false, false))
            {
                if (!File.Exists(sourceFilePath))
                {
                    Logger.Log(string.Format("Can't find the HKSEA file: {0} .", sourceFilePath));
                    return mainRicList;
                }
                var workbookSEA = ExcelUtil.CreateOrOpenExcelFile(app, sourceFilePath);
                Worksheet worksheet = ExcelUtil.GetWorksheet("Sheet2", workbookSEA);
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                Range columnC, columnD;
                string columnCText, columnDText, columnBText;
                
                DateTime tomorrow = MiscUtil.GetNextTradingDay(DateTime.Today, holidayList, 1);
              
                for (int i = lastUsedRow; i >= worksheet.UsedRange.Row; i--)
                {
                    columnC = ExcelUtil.GetRange(string.Format("C{0}", i), worksheet);
                    columnD = ExcelUtil.GetRange(string.Format("D{0}", i), worksheet);
                    columnCText = columnC.Text.ToString();
                    columnDText = columnD.Text.ToString();
                    if (!string.IsNullOrEmpty(columnCText) && !string.IsNullOrEmpty(columnDText))
                    {
                        DateTime effectiveDate = ParseDateTime(columnDText);

                        if ("DELETE".Equals(columnCText.ToUpper()) && effectiveDate.CompareTo(tomorrow) == 0)
                        {
                            columnBText = ExcelUtil.GetRange(string.Format("B{0}", i), worksheet).Text.ToString();
                            if (!string.IsNullOrEmpty(columnBText))
                            {
                                mainRicList.Add(columnBText);
                            }
                        }
                        if (effectiveDate.CompareTo(tomorrow) < 0)
                        {
                            break;
                        }
                    }
                }
                workbookSEA.Close(false, workbookSEA.FullName, false);
            }
            return mainRicList;
        }

        private void GenerateGEDAFiles(IEnumerable<string> mainRicList, IEnumerable<string> miRicList)
        {
            string date = DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US"));
            string fileName = string.Format("CBBC_MI_Drop_{0}.txt", date.ToUpper());
            string fileDir = outputPath;
            if (!Directory.Exists(fileDir))
            {
                Directory.CreateDirectory(fileDir);
            }
            string gedaDropFilePath = Path.Combine(fileDir, fileName);
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("RIC");
            foreach (string mainRic in mainRicList)
            {
                sb.AppendLine(mainRic);
            }
            foreach (string miRic in miRicList)
            {
                sb.AppendLine(miRic);
            }
            string content = sb.ToString();
            try
            {
                File.WriteAllText(gedaDropFilePath, content, Encoding.UTF8);
                TaskResultList.Add(new TaskResultEntry("GEDA File", "Generate HK CBBC MI DROP GEDA file", gedaDropFilePath, FileProcessType.GEDA_BULK_RIC_DELETE));
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating GEDA files. EX: {0} .", ex.Message));
            }
        }

        public void InsertChangeInfo(int rowPos, ExcelLineWriter writer, SameDayChangeInfo sameDayChangeInfo, Worksheet worksheet, List<BordLotSizeChangeInfo> bordLotSizeChangeInfoList)
        {
            InsertNewDeleteChangeInfo(rowPos--, writer, sameDayChangeInfo.DeleteBrokerList, worksheet);
            int curPos = rowPos;
            while (curPos > 0)
            {
                Range r = ExcelUtil.GetRange(curPos - 1, 4, worksheet);
                if (r.Text == null || r.Text.ToString() == "")
                {
                    curPos--;
                    continue;
                }
                if (DateTimeEqual(sameDayChangeInfo.Date, r.Text.ToString()))
                {
                    curPos--;
                    continue;
                }
                InsertNameChange(curPos, writer, sameDayChangeInfo.NameChangeToList, worksheet);
                InsertRicChange(curPos, writer, sameDayChangeInfo.RicChangeToList, worksheet);
                InsertBordSizeChangeInfo(curPos, writer, sameDayChangeInfo.BordLotSizeChangeList, bordLotSizeChangeInfoList, worksheet);
                InsertNewDeleteChangeInfo(curPos, writer, sameDayChangeInfo.NewBrokerList, worksheet);
                break;
            }
            int lastUsedRowNum = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            ExcelUtil.GetRange(1, 4, lastUsedRowNum, 4, worksheet).NumberFormat = "dd/MM/yyyy";
        }

        public void InsertBordSizeChangeInfo(int rowPos, ExcelLineWriter writer, List<BrokerChangeInfo> brokerChangeList, List<BordLotSizeChangeInfo> bordLotSizeChangeInfoList, Worksheet worksheet)
        {
            ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(rowPos, 1, worksheet), brokerChangeList.Count);
            if (brokerChangeList.Count == 0)
            {
                return;
            }
            int lastUsedCol = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            foreach (BrokerChangeInfo broker in brokerChangeList)
            {
                string type = string.Empty;
                writer.PlaceNext(rowPos++, 1);
                writer.WriteLine(broker.Name);
                writer.WriteLine(broker.Ric);
                foreach (BordLotSizeChangeInfo bordLotSizeChangeInfo in bordLotSizeChangeInfoList.Where(bordLotSizeChangeInfo => bordLotSizeChangeInfo.Name == broker.Name))
                {
                    type = string.Format("Bord Lot Size Change to {0}", bordLotSizeChangeInfo.BoardLotSize);
                    break;
                }
                writer.WriteLine(type);
                writer.WriteLine(broker.Date);
                ExcelUtil.GetRange(rowPos - 1, 1, rowPos - 1, lastUsedCol, worksheet).Interior.Color = IsIPos(broker.Ric) ? 10092543.0 : 16777215.0;
            }
        }
        public void InsertNewDeleteChangeInfo(int rowPos, ExcelLineWriter writer, List<BrokerChangeInfo> brokerChangeList, Worksheet worksheet)
        {
            ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(rowPos, 1, worksheet), brokerChangeList.Count);
            if (brokerChangeList.Count == 0)
            {
                return;
            }

            int lastUsedCol = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            foreach (BrokerChangeInfo broker in brokerChangeList)
            {
                writer.PlaceNext(rowPos++, 1);
                writer.WriteLine(broker.Name);
                writer.WriteLine(broker.Ric);
                writer.WriteLine(broker.Type);
                writer.WriteLine(broker.Date);
                ExcelUtil.GetRange(rowPos - 1, 1, rowPos - 1, lastUsedCol, worksheet).Interior.Color = IsIPos(broker.Ric) ? 10092543.0 : 16777215.0;
            }
            ExcelUtil.GetRange(rowPos - brokerChangeList.Count, 4, rowPos - 1, 4, worksheet).NumberFormat = "dd/MM/yyyy";
            if (brokerChangeList[0].Type.ToLower() == "delete")
            {
                ExcelUtil.GetRange(rowPos - brokerChangeList.Count, 1, rowPos - 1, lastUsedCol, worksheet).Font.Color = ParseDateTime(brokerChangeList[0].Date) == MiscUtil.GetNextTradingDay(DateTime.Today, holidayList, 1) ? 255.0 : 0.0;
            }
            else
            {
                ExcelUtil.GetRange(rowPos - brokerChangeList.Count, 1, rowPos - 1, lastUsedCol, worksheet).Font.Color = 0.0;
            }
        }

        public void InsertNameChange(int rowPos, ExcelLineWriter writer, List<BrokerChangeInfo> nameChangeList, Worksheet worksheet)
        {
            if (nameChangeList == null || nameChangeList.Count == 0)
            {
                return;
            }
            ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(rowPos, 1, worksheet), 2 * nameChangeList.Count);
            int lastUsedCol = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            foreach (BrokerChangeInfo nameChange in nameChangeList)
            {
                writer.PlaceNext(rowPos++, 2);
                writer.WriteLine(nameChange.Ric);
                writer.WriteLine(nameChange.Type);
                writer.WriteLine(nameChange.Date);
                writer.PlaceNext(rowPos++, 1);
                writer.WriteLine(nameChange.Name);
                writer.PlaceNextAndWriteLine(rowPos, 4, nameChange.Date);
                ExcelUtil.GetRange(rowPos - 2, lastUsedCol, rowPos - 1, lastUsedCol, worksheet).Interior.Color = IsIPos(nameChange.Ric) ? 10092543.0 : 16777164.0;
            }
        }

        public void InsertRicChange(int rowPos, ExcelLineWriter writer, List<StockNameRicChange> ricChangeList, Worksheet worksheet)
        {
            if (ricChangeList == null || ricChangeList.Count == 0)
            {
                return;
            }
            ExcelUtil.InsertBlankRows(ExcelUtil.GetRange(rowPos, 1, worksheet), 2 * ricChangeList.Count);
            int lastUsedCol = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            foreach (StockNameRicChange ricChange in ricChangeList)
            {
                writer.PlaceNext(rowPos++, 1);
                writer.WriteLine(ricChange.Name);
                writer.WriteLine(ricChange.OldValue);
                writer.WriteLine(ricChange.Type);
                writer.WriteLine(ricChange.Date);
                writer.PlaceNext(rowPos++, 2);
                writer.WriteLine(ricChange.NewValue);
                if (IsUpgrade(ricChange))
                {
                    writer.WriteLine("UPGRADE");
                }

                writer.PlaceNextAndWriteLine(rowPos, 4, ricChange.Date);
                ExcelUtil.GetRange(rowPos - 2, lastUsedCol, rowPos - 1, lastUsedCol, worksheet).Interior.Color = IsIPos(ricChange.Ric) ? 10092543.0 : 16777164.0;
            }
        }

        public bool IsUpgrade(StockNameRicChange ricChange)
        {
            return ricChange.OldValue.Trim().StartsWith("8") && (!ricChange.NewValue.Trim().StartsWith("8"));
        }

        public bool IsIPos(string ricStr)
        {
            int ricInt = 0;
            ricStr = ricStr.Replace(".HK", "").Replace("<", "").Replace(">", "");
            bool result = int.TryParse(ricStr, out ricInt);
            if (!result)
            {
                Logger.LogErrorAndRaiseException(string.Format("Can't parse {0}", ricStr));
            }
            else
            {
                int temp = ricInt / 10000;
                if (temp == 1 || temp == 2 || temp == 6)
                {
                    result = false;
                }
                else
                    return  true;
            }
            return result;
        }

        public ChangeInfo GetAllBrokerChangeInfo(Worksheet worksheet, int lastRowContainInfo)
        {
            ChangeInfo changedBrokerInfo = new ChangeInfo {AllChangedBrokerList = new List<BrokerChangeInfo>()};
            for (int rowNum = 2; rowNum <= lastRowContainInfo; rowNum++)
            {
                BrokerChangeInfo brokerChangeInfo = new BrokerChangeInfo
                {
                    Name = ExcelUtil.GetRange(rowNum, 1, worksheet).Value2.ToString(),
                    Ric = ExcelUtil.GetRange(rowNum, 2, worksheet).Value2.ToString(),
                    Type = ExcelUtil.GetRange(rowNum, 3, worksheet).Value2.ToString(),
                    Date = ExcelUtil.GetRange(rowNum, 4, worksheet).Value2.ToString()
                };
                changedBrokerInfo.AllChangedBrokerList.Add(brokerChangeInfo);
            }
            changedBrokerInfo.DifDayChangeInfo = changedBrokerInfo.GetDifDayChangeInfo();
            return changedBrokerInfo;
        }

        public void UpdateSheet5(Worksheet worksheet5, List<BordLotSizeChangeInfo> bordLotSizeChangeList)
        {
            using (ExcelLineWriter writer = new ExcelLineWriter(worksheet5, 2, 2, ExcelLineWriter.Direction.Down))
            {
                while (true)
                {
                    Range nameCell = ExcelUtil.GetRange(writer.Row, 1, worksheet5);
                    if (nameCell == null || nameCell.Text.ToString().Trim() == "0")
                    {
                        break;
                    }
                    if (bordLotSizeChangeList != null && bordLotSizeChangeList.Count != 0)
                    {
                        if (ExcelUtil.GetRange(writer.Row, 3, worksheet5).Text.ToString().Trim() == "#N/A")
                        {
                            writer.PlaceNextAndWriteLine(writer.Row, 3, "CHANGE");
                            writer.PlaceNext(writer.Row - 1, writer.Col);
                        }
                    }
                    Range r = ExcelUtil.GetRange(writer.Row, 2, worksheet5);
                    string line = r.Value2.ToString().Trim(new[] { '>', '<' });
                    writer.PlaceNextAndWriteLine(writer.Row, 2, line);
                }
            }
        }

        public void UpdateAddDropSheet(Worksheet worksheetAddDrop, List<BordLotSizeChangeInfo> bordLotSizeChangeList)
        {
            using (ExcelLineWriter writer = new ExcelLineWriter(worksheetAddDrop, 2, 1, ExcelLineWriter.Direction.Down))
            {
                while (true)
                {
                    Range changeTypeRange = ExcelUtil.GetRange(writer.Row, 1, worksheetAddDrop);
                    if (changeTypeRange == null || changeTypeRange.Text.ToString().Trim() == "#N/A")
                    {
                        break;
                    }
                    if (changeTypeRange != null && changeTypeRange.Text.ToString().Trim() == "FALSE")
                    {
                        writer.WriteLine("CHANGE");
                        writer.PlaceNextAndWriteLine(writer.Row, 3, ExcelUtil.GetRange(writer.Row, 4, worksheetAddDrop).Text.ToString());
                    }
                    else if (changeTypeRange != null && changeTypeRange.Text.ToString().Trim() == "Change")
                    {
                        writer.PlaceNextAndWriteLine(writer.Row, 3, ExcelUtil.GetRange(writer.Row, 4, worksheetAddDrop).Text.ToString());
                    }
                    else
                    {
                        writer.MoveNext();
                    }
                }
            }
        }

        public Range GetAKRange(Workbook workbookSource, out int lastUsedRowNum, out List<BordLotSizeChangeInfo> bordLotSizeChangeList)
        {
            var worksheetSource = ExcelUtil.GetWorksheet(configObj.FormalSourceFileWorksheetName, workbookSource);
            if (worksheetSource == null)
            {
                Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.FormalSourceFileWorksheetName, workbookSource.Name));
            }

            bordLotSizeChangeList = GetBordChangeList(worksheetSource);

            //Update the worksheet format to fit some requiements: 
            //
            //1. Delete A column;
            //2. Delete D column;
            //3. Insert blank column after F column;
            //4. Insert black row before 11 row.
            HKShared.UpdateWorksheetFormat(worksheetSource);

            //Insert a column at D column
            ExcelUtil.InsertBlankCols(ExcelUtil.GetRange(1, 4, worksheetSource), 1);
            lastUsedRowNum = worksheetSource.UsedRange.Row + worksheetSource.UsedRange.Rows.Count - 1;
            Range akRange = ExcelUtil.GetRange(1, 1, lastUsedRowNum, 11, worksheetSource);
            return akRange;
        }

        public void UpdateTemplateV2File(Range A2Range, int lastRowContainInfo)
        {
            Logger.Log("Start to generate the RIC_Convs_Template_V2 file.");
            A2Range.Copy(Missing.Value);
            string TemplateV2FilePath = BackupFiles(configObj.RicCowsTemplateV2FilePath);
            using (ExcelApp appTemplateV2 = new ExcelApp(false, false))
            {
                var workbookTemplateV2 = ExcelUtil.CreateOrOpenExcelFile(appTemplateV2, TemplateV2FilePath);
                var worksheetTemplateV2 = ExcelUtil.GetWorksheet(configObj.TemplateV2WorksheetName, workbookTemplateV2);
                if (worksheetTemplateV2 == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.TemplateV2WorksheetName, workbookTemplateV2.Name));
                }

                Range C3Range = ExcelUtil.GetRange(3, 3, lastRowContainInfo, 18, worksheetTemplateV2);
                C3Range.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);
                ExcelUtil.GetRange(1, 1, worksheetTemplateV2).Copy(Missing.Value);
                SetCreateValue(worksheetTemplateV2, lastRowContainInfo);
                //Run Macros
                appTemplateV2.ExcelAppInstance.GetType().InvokeMember("Run",
                    BindingFlags.Default | BindingFlags.InvokeMethod,
                    null,
                    appTemplateV2.ExcelAppInstance,
                    new object[] { "FormatData" });

                string newTemplateV2FilePath = Path.Combine(Path.GetDirectoryName(workbookTemplateV2.FullName), Path.GetFileName(configObj.RicCowsTemplateV2FilePath));
                workbookTemplateV2.SaveCopyAs(newTemplateV2FilePath);
                workbookTemplateV2.Close(false, workbookTemplateV2.FullName, false);
                TaskResultList.Add(new TaskResultEntry("RIC_Convs_Template_V2  File", "Generate the RIC_Convs_Template_V2.xls file", newTemplateV2FilePath));
                File.Delete(TemplateV2FilePath);
            }
            Logger.Log("Finished generating the RIC_Convs_Template_V2 file.");
        }

        public List<BordLotSizeChangeInfo> GetBordChangeList(Worksheet worksheet)
        {
            List<BordLotSizeChangeInfo> bordLotSizeChangeList = new List<BordLotSizeChangeInfo>();
            int lastUsedRowNum = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            int rowNum = 16;
            while (++rowNum <= lastUsedRowNum)
            {
                if (ExcelUtil.GetRange(rowNum, 4, worksheet).Value2 != null)
                {
                    if (ExcelUtil.GetRange(rowNum, 4, worksheet).Value2.ToString().ToLower() == "Change of Board Lot Size Size".ToLower())
                    {
                        BordLotSizeChangeInfo bordLotSizeChangeInfo = new BordLotSizeChangeInfo();
                        bordLotSizeChangeInfo.Code = ExcelUtil.GetRange(rowNum, 2, worksheet).Value2.ToString();
                        bordLotSizeChangeInfo.Name = ExcelUtil.GetRange(rowNum, 3, worksheet).Value2.ToString();
                        bordLotSizeChangeInfo.EventType = ExcelUtil.GetRange(rowNum, 4, worksheet).Value2.ToString();
                        bordLotSizeChangeInfo.BoardLotSize = ExcelUtil.GetRange(rowNum, 6, worksheet).Value2.ToString();
                        bordLotSizeChangeInfo.Date = ExcelUtil.GetRange(rowNum, 8, worksheet).Value2.ToString();
                        bordLotSizeChangeList.Add(bordLotSizeChangeInfo);
                    }
                }
            }
            return bordLotSizeChangeList;
        }

        //Set A col value as "Create" for TemplateV2
        public void SetCreateValue(Worksheet worksheet, int lastRowNum)
        {
            using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 3, 1, ExcelLineWriter.Direction.Down))
            {
                while (ExcelUtil.GetRange(writer.Row, 1, worksheet) != null)
                {
                    string value = ExcelUtil.GetRange(writer.Row, 3, worksheet).Value2.ToString().ToLower();
                    if (value.Contains("add") || value.Contains("delete") || value.Contains("change"))
                    {
                        writer.WriteLine("Create");
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }

        //Create sub-folder to store the new generated files return new file path
        public string BackupFiles(string filePath)
        {
            string dir = Path.GetDirectoryName(filePath);
            dir = Path.Combine(dir, DateTime.Now.ToString("yyyy-MM-dd"));
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            string newFileName = NewFileName(filePath);
            string newFilePath = Path.Combine(dir, newFileName);
            if (File.Exists(newFilePath))
            {
                File.Delete(newFilePath);
            }
            File.Copy(filePath, newFilePath);
            return newFilePath;
        }

        //Create a new file name 
        public string NewFileName(string currentFileName)
        {
            string newFileName = Path.GetFileNameWithoutExtension(currentFileName);
            newFileName += "_";
            newFileName += DateTime.Now.ToString("HH-mm-ss");
            newFileName += "_";
            newFileName += Guid.NewGuid().ToString();
            newFileName += Path.GetExtension(currentFileName);
            return newFileName;
        }

        public DateTime GetNextBusinessDay(DateTime currentDate, int holidayCount)
        {
            DateTime nextBusinessDay;
            if (holidayCount == 0)
            {
                if (currentDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    nextBusinessDay = currentDate.AddDays(-2);
                }
                else if (currentDate.DayOfWeek == DayOfWeek.Monday)
                {
                    nextBusinessDay = currentDate.AddDays(-3);
                }
                else
                {
                    nextBusinessDay = currentDate.AddDays(-1);
                }
            }
            else
            {
                nextBusinessDay = currentDate.AddDays(holidayCount - 1);
            }

            return nextBusinessDay;
        }

        public bool DateTimeCompare(string time1, string time2)
        {
            if (string.IsNullOrEmpty(time1) || string.IsNullOrEmpty(time2))
            {
                return false;
            }

            //Some Col D value in HKSEA-Z2 file can't be specified. It's value showed as "to be announced" 
            if (time2.ToLower() == "to be announced")
            {
                return false;
            }
            try
            {

                DateTime dateTime1 = ParseDateTime(time1);
                DateTime dateTime2 = ParseDateTime(time2);
                return dateTime1 >= dateTime2;
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
                throw ex;
            }
        }

        private bool DateTimeLessThan(string time1, string time2)
        {
            if (string.IsNullOrEmpty(time1) || string.IsNullOrEmpty(time2))
            {
                return false;
            }

            //Some Col D value in HKSEA-Z2 file can't be specified. It's value showed as "to be announced" 
            if (time2.ToLower() == "to be announced")
            {
                return false;
            }
            DateTime dateTime1 = ParseDateTime(time1);
            DateTime dateTime2 = ParseDateTime(time2);
            return dateTime1 <= dateTime2;
        }

        public bool DateTimeEqual(string time1, string time2)
        {
            if (string.IsNullOrEmpty(time1) || string.IsNullOrEmpty(time2))
            {
                return false;
            }
            DateTime dateTime1 = ParseDateTime(time1);
            DateTime dateTime2 = ParseDateTime(time2);
            return dateTime1 == dateTime2;
        }

        public DateTime ParseDateTime(string time)
        {
            DateTime dt;

            if (time.Contains("/"))
            {
                string[] timeItem = time.Split('/');
                time = timeItem[2] + "-" + timeItem[1] + "-" + timeItem[0];
            }

            if (!DateTime.TryParse(time, out dt))
            {
                throw new InvalidDataException(time);
            }
            return dt;
        }

        #region HK DelistingAndNameChange Job
        private void StartDelistingAndNameChangeJob()
        {
            cbbcList.Clear();
            warrantList.Clear();
            nameChangeList.Clear();
            apartCbbcList.Clear();
            apartWarrantList.Clear();
            InitialHyphen();
            int startPosition = 16;

            if (File.Exists(formalFileBackupPath))
            {

                Logger.Log("Start DelistingAndNameChange Job......");
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, formalFileBackupPath);
                    var worksheet = ExcelUtil.GetWorksheet(configObj.FormalSourceFileWorksheetName, workbook);
                    if (worksheet == null)
                    {
                        Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.FormalSourceFileWorksheetName, formalFileBackupPath));
                    }

                    Range rangeCol6 = ExcelUtil.GetRange(startPosition, 6, worksheet);
                    Range rangeCol3 = ExcelUtil.GetRange(startPosition, 3, worksheet);
                    if (rangeCol6.Value[Missing.Value] != null && rangeCol3.Value[Missing.Value] != null)
                    {
                        int usedRange = worksheet.UsedRange.Rows.Count;
                        GetDelistingAndNameChangeData(usedRange, worksheet, startPosition);
                        ApartList(apartCbbcList, cbbcList);
                        ApartList(apartWarrantList, warrantList);
                    }
                    else
                    {
                        Logger.LogErrorAndRaiseException(" Check today's file in the folder Delisting!");

                    }
                    app.ExcelAppInstance.AlertBeforeOverwriting = false;
                    workbook.Close(false, formalFileBackupPath, true);
                }

                GenerateDelistingFMFiles();

                Logger.Log("Finished DelistingAndNameChange Job......");
            }
            else
            {
                Logger.LogErrorAndRaiseException(string.Format(" Can't find today's delisting file {0}!", configObj.FormalSourceFilePath));
            }
        }

        private void GenerateDelistingFMFiles()
        {
            Application xlApp = new Application();
            try
            {
                WriteDataIntoFile(xlApp);
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in generate today's delisting FM files. {0}", ex.Message);
                Logger.Log(msg);
                throw new Exception(msg);
            }
            finally
            {
                CoreObj.KillExcelProcess(xlApp);
            }
        }

        private void InitialHyphen()
        {
            hyphen = new char[xmlConfigObj.Hyphen.Count];
            int count = 0;
            foreach (string item in xmlConfigObj.Hyphen)
            {
                hyphen[count] = (item.ToCharArray())[0];
                count++;
            }
        }

        private void GetDelistingAndNameChangeData(int usedRange, Worksheet delSheet, int StartPosition)
        {
            // Currently not check the records whose event is "Change of Name"
            for (int rowIndex = StartPosition; rowIndex < usedRange; rowIndex++)
            {
                Range rangeCol6 = ExcelUtil.GetRange(rowIndex, 6, delSheet);
                Range rangeCol3 = ExcelUtil.GetRange(rowIndex, 3, delSheet);

                if (rangeCol6.Value[Missing.Value] == null || rangeCol3.Value[Missing.Value] == null)
                    continue;

                string cel6Value = rangeCol6.Value[Missing.Value].ToString();
                string cel3Value = rangeCol3.Value[Missing.Value].ToString();

                List<string> itemList;
                if (cel6Value.Equals("Delisting") && cel3Value.Length == 5 && cel3Value.StartsWith("6"))
                {
                    itemList = ReadDelistingRowData(delSheet, rowIndex);
                    cbbcList.Add(itemList);
                }
                if (cel6Value.Equals("Delisting") && cel3Value.Length == 5 && (cel3Value.StartsWith("1") || cel3Value.StartsWith("2")))
                {
                    itemList = ReadDelistingRowData(delSheet, rowIndex);
                    warrantList.Add(itemList);
                }
            }
        }

        private List<string> ReadDelistingRowData(Worksheet wSheet, int rowIndex)
        {
            List<string> itemList = new List<string>();
            string col3Value = ExcelUtil.GetRange(rowIndex, 3, wSheet).Value[Missing.Value].ToString();
            itemList.Add(col3Value);
            Range range = ExcelUtil.GetRange(rowIndex, 4, wSheet);
            if (range.Value[Missing.Value] != null)
            {
                itemList.Add(range.Value[Missing.Value].ToString());
            }
            else
            {
                Logger.LogErrorAndRaiseException(string.Format("Check the source, no stock name in the row {0}", rowIndex));
            }
            range = ExcelUtil.GetRange(rowIndex, 9, wSheet);
            if (range.Value[Missing.Value] != null)
            {
                itemList.Add(range.Value[Missing.Value].ToString());
            }
            else
            {
                Logger.LogErrorAndRaiseException(string.Format("Check the source, no date in the row {0}", rowIndex));
            }
            return itemList;
        }

        private void ApartList(List<List<List<string>>> apartList, List<List<string>> cbbcList)
        {
            while (cbbcList.Count > 0)
            {
                ApartListAccordingDate(apartList, cbbcList);
            }
        }

        //group the cbbcList with the Date value.
        private void ApartListAccordingDate(List<List<List<string>>> apartList, List<List<string>> cbbcList)
        {
            List<List<string>> newCbbcList = new List<List<string>> {cbbcList[0]};
            string date = cbbcList[0][2];
            cbbcList.Remove(cbbcList[0]);

            int count = 0;
            while (count < cbbcList.Count)
            {
                if (date.Equals(cbbcList[count][2]))
                {
                    newCbbcList.Add(cbbcList[count]);
                    cbbcList.Remove(cbbcList[count]);
                }
                else
                {
                    count++;
                }
            }
            apartList.Add(newCbbcList);
        }

        private void WriteDataIntoFile(Application xlApp)
        {
            // Currently exclude the records whose event is "Change of Name"
            if (apartCbbcList.Count > 0)
            {
                foreach (List<List<string>> cbbcItem in apartCbbcList)
                {
                    WriteCBBCDelistingFMFile(xlApp, cbbcItem);
                }
                Logger.Log("Exists CBBCDelisting records and generated the CBBCDelisting FM file.");
            }
            if (apartWarrantList.Count > 0)
            {
                foreach (List<List<string>> warrantItem in apartWarrantList)
                {
                    WriteWarrantDelistingFMFile(xlApp, warrantItem);
                }
                Logger.Log("Exists WARRANTDelisting records and generated the WARRANTDelisting FM file.");
            }
        }

        private void WriteCBBCDelistingFMFile(_Application xlApp, List<List<string>> dataList)
        {
            string delistingSavePath = outputPath;
            DateTime date = ParseDateTime(dataList[0][2]);

            string formatDate = TimeUtil.GetFormatDate(date);
            string effectiveDate = TimeUtil.GetEffectiveDate(date);

            int filesCount;
            if ((dataList.Count % 20) > 0)
            {
                filesCount = dataList.Count / 20 + 1;
            }
            else
            {
                filesCount = dataList.Count / 20;
            }

            for (int i = 0; i < filesCount; i++)
            {
                string fileName = "HK" + TimeUtil.shortYear + "-_DELETE";
                Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                wSheet.Application.Cells.HorizontalAlignment = Microsoft.Office.Core.XlHAlign.xlHAlignLeft;
                wSheet.Cells[1, 1] = "Please action the delete of the following HK stock on TQS.";
                wSheet.Cells[3, 1] = "FM Serial Number:";
                wSheet.Cells[3, 2] = "HK06-411";
                wSheet.Cells[4, 1] = "Effective Date:";
                wSheet.Cells[4, 2] = effectiveDate;
                int rowIndex = 6;
                for (int j = i * 20; j < i * 20 + 20; j++)
                {
                    string code = "0";
                    if (j < dataList.Count)
                    {
                        code = dataList[j][0];
                        string name = dataList[j][1];

                        wSheet.Cells[rowIndex++, 1] = "+DELETION+";
                        wSheet.Cells[rowIndex++, 1] = "'---------------------------------------------------------------------------------------------------";
                        Range range = ExcelUtil.GetRange(rowIndex, 1, wSheet);
                        range.NumberFormatLocal = "(0)";
                        wSheet.Cells[rowIndex++, 1] = (j % 20 + 1).ToString();
                        wSheet.Cells[rowIndex, 1] = "Underlying RIC:";
                        wSheet.Cells[rowIndex++, 2] = code + ".HK";
                        wSheet.Cells[rowIndex, 1] = "Composite chain RIC:";
                        wSheet.Cells[rowIndex++, 2] = "0#" + code + ".HK";
                        wSheet.Cells[rowIndex, 1] = "Broker page RIC:";
                        wSheet.Cells[rowIndex++, 2] = code + "bk.HK";
                        wSheet.Cells[rowIndex, 1] = "Official Code:";
                        wSheet.Cells[rowIndex++, 2] = code;
                        wSheet.Cells[rowIndex, 1] = "Exchange Symbol:";
                        wSheet.Cells[rowIndex++, 2] = code;
                        wSheet.Cells[rowIndex, 1] = "Misc.Info page RIC:";
                        wSheet.Cells[rowIndex++, 2] = code + "MI.HK";
                        wSheet.Cells[rowIndex, 1] = "Displayname:";
                        wSheet.Cells[rowIndex++, 2] = name;
                        wSheet.Cells[rowIndex, 1] = "Chain RIC:";
                        wSheet.Cells[rowIndex++, 2] = "0#CBBC.HK";
                        wSheet.Cells[rowIndex, 1] = "Warrant Type:";
                        string bearOrBull = "";
                        if (name.Substring(9, 1).Equals("P"))
                        {
                            bearOrBull = "Bear";
                        }
                        else if (name.Substring(9, 1).Equals("C"))
                        {
                            bearOrBull = "Bull";
                        }
                        wSheet.Cells[rowIndex++, 2] = "Callable " + bearOrBull + " Contracts";
                        wSheet.Cells[rowIndex, 1] = "Misc Info page Chain RIC:";
                        wSheet.Cells[rowIndex++, 2] = "0#MI.HK";
                        wSheet.Cells[rowIndex, 1] = "Index RIC:";
                        wSheet.Cells[rowIndex++, 2] = "N/A";
                        wSheet.Cells[rowIndex, 1] = "Recordtype:";
                        wSheet.Cells[rowIndex++, 2] = "097";
                        wSheet.Cells[rowIndex++, 1] = "---------------------------------------------------------------------------------------------------";
                    }
                    if (!code.Equals("0"))
                    {
                        fileName = fileName + "_" + code;
                    }
                }
                fileName += "_" + formatDate + ".xls";
                wBook.Application.DisplayAlerts = false;
                if (!Directory.Exists(delistingSavePath))
                {
                    Directory.CreateDirectory(delistingSavePath);
                }
                string cbbcDelistingFilePath = Path.Combine(delistingSavePath, fileName);
                if (File.Exists(cbbcDelistingFilePath))
                {
                    File.Delete(cbbcDelistingFilePath);
                }
                wBook.SaveAs(cbbcDelistingFilePath, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                TaskResultList.Add(new TaskResultEntry("CBBCDelisting FM  File", "Generate the CBBCDelisting FM file", cbbcDelistingFilePath));
                wBook.Close(Missing.Value, Missing.Value, Missing.Value);
            }
        }

        private void WriteWarrantDelistingFMFile(_Application xlApp, List<List<string>> dataList)
        {
            string delistingSavePath = outputPath;
            DateTime date = ParseDateTime(dataList[0][2]);

            string formatDate = TimeUtil.GetFormatDate(date);
            string effectiveDate = TimeUtil.GetEffectiveDate(date);

            int filesCount = 0;
            if ((dataList.Count % 20) > 0)
            {
                filesCount = dataList.Count / 20 + 1;
            }
            else
            {
                filesCount = dataList.Count / 20;
            }

            for (int i = 0; i < filesCount; i++)
            {
                string fileName = "HK" + TimeUtil.shortYear + "-_DELETE";
                Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                wSheet.Application.Cells.HorizontalAlignment = Microsoft.Office.Core.XlHAlign.xlHAlignLeft;
                wSheet.Cells[1, 1] = "Please action the delete of the following HK stock on TQS.";
                wSheet.Cells[3, 1] = "FM Serial Number:";
                wSheet.Cells[3, 2] = "HK06-411";
                wSheet.Cells[4, 1] = "Effective Date:";
                wSheet.Cells[4, 2] = effectiveDate;
                int rowIndex = 6;
                for (int j = i * 20; j < i * 20 + 20; j++)
                {
                    string code = "0";
                    if (j < dataList.Count)
                    {
                        code = dataList[j][0];
                        string name = dataList[j][1];

                        try
                        {
                            GetInfoFromInternet(code, "Warrant");
                        }
                        catch
                        {
                            underlyingStockCode = "UNDERLYINGCODE";
                        }


                        wSheet.Cells[rowIndex++, 1] = "+DELETION+";
                        wSheet.Cells[rowIndex++, 1] = "'---------------------------------------------------------------------------------------------------";
                        Range range = ExcelUtil.GetRange(rowIndex, 1, wSheet);
                        range.NumberFormatLocal = "(0)";
                        wSheet.Cells[rowIndex++, 1] = (j % 20 + 1).ToString();
                        wSheet.Cells[rowIndex, 1] = "Underlying RIC:";
                        wSheet.Cells[rowIndex++, 2] = code + ".HK";
                        wSheet.Cells[rowIndex, 1] = "Composite chain RIC:";
                        wSheet.Cells[rowIndex++, 2] = "0#" + code + ".HK";
                        wSheet.Cells[rowIndex, 1] = "Broker page RIC:";
                        wSheet.Cells[rowIndex++, 2] = code + "bk.HK";
                        wSheet.Cells[rowIndex, 1] = "Official Code:";
                        wSheet.Cells[rowIndex++, 2] = code;
                        wSheet.Cells[rowIndex, 1] = "Exchange Symbol:";
                        wSheet.Cells[rowIndex++, 2] = code;
                        wSheet.Cells[rowIndex, 1] = "Misc.Info page RIC:";
                        wSheet.Cells[rowIndex++, 2] = code + "MI.HK";
                        wSheet.Cells[rowIndex, 1] = "Displayname:";
                        wSheet.Cells[rowIndex++, 2] = name;
                        wSheet.Cells[rowIndex, 1] = "Chain RIC:";
                        wSheet.Cells[rowIndex++, 2] = chainRic;
                        wSheet.Cells[rowIndex, 1] = "Chain RIC:";
                        wSheet.Cells[rowIndex++, 2] = underlyingStockCode.Contains("HSI") ? "N/A" : "0#CWRTS.HK";
                        wSheet.Cells[rowIndex, 1] = "Chain RIC:";
                        wSheet.Cells[rowIndex++, 2] = "0#WARRANTS.HK";
                        wSheet.Cells[rowIndex, 1] = "Warrant Type:";
                        wSheet.Cells[rowIndex++, 2] = underlyingStockCode.Contains("HSI") ? "Hang Seng Index Warrant" : "Equity Warrant";
                        wSheet.Cells[rowIndex, 1] = "Misc Info page Chain RIC:";
                        wSheet.Cells[rowIndex++, 2] = "0#MI.HK";
                        wSheet.Cells[rowIndex, 1] = "Index RIC:";
                        wSheet.Cells[rowIndex++, 2] = "N/A";
                        wSheet.Cells[rowIndex, 1] = "Recordtype:";
                        wSheet.Cells[rowIndex++, 2] = "097";
                        wSheet.Cells[rowIndex++, 1] = "---------------------------------------------------------------------------------------------------";
                    }
                    if (!code.Equals("0"))
                    {
                        fileName = fileName + "_" + code;
                    }
                }
                fileName += "_" + formatDate + ".xls";
                wBook.Application.DisplayAlerts = false;
                if (!Directory.Exists(delistingSavePath))
                {
                    Directory.CreateDirectory(delistingSavePath);
                }
                string warrantDelistingFilePath = Path.Combine(delistingSavePath, fileName);
                if (File.Exists(warrantDelistingFilePath))
                {
                    File.Delete(warrantDelistingFilePath);
                }
                wBook.SaveAs(warrantDelistingFilePath, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                TaskResultList.Add(new TaskResultEntry("WarrantDelisting FM  File", "Generate the WarrantDelisting FM file", warrantDelistingFilePath));
                wBook.Close(Missing.Value, Missing.Value, Missing.Value);
            }
        }

        private void GetInfoFromInternet(string code, string condition)
        {
            string url = "http://www.hkex.com.hk/eng/invest/company/profile_page_e.asp?WidCoID=" + code +
                         "&WidCoAbbName=&Month=&langcode=e";

            foreach (var tryCount in Enumerable.Range(0, 10))
            {
                try
                {
                    WebRequest wr = WebRequest.Create(url);
                    WebResponse rs = wr.GetResponse();
                    StreamReader sr = new StreamReader(rs.GetResponseStream());
                    string htm = sr.ReadToEnd();

                    int pos = htm.IndexOf("Company/Securities Name:", StringComparison.Ordinal);
                    string beforeHtm = htm.Substring(0, pos);
                    string afterHtm = htm.Substring(pos, htm.Length - pos);

                    int beforeTablePos = beforeHtm.LastIndexOf("<table", StringComparison.Ordinal);
                    int afterTablePos = afterHtm.IndexOf("</table>", StringComparison.Ordinal);
                    string beforeTable = beforeHtm.Substring(beforeTablePos, beforeHtm.Length - beforeTablePos);
                    string afterTable = afterHtm.Substring(0, afterTablePos) + "</table>";
                    string simpleHtml = "<html><head></head><body>" +
                                        (beforeTable + afterTable).Replace("\r", "").Replace("\n", "").Replace("\t", "") +
                                        "</body></html>";

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(simpleHtml);

                    if (condition.Equals("Warrant"))
                    {
                        GetUnderlyingStockCode(htmlDoc);
                    }
                    else
                    {
                        GetNameChangeInfo(htmlDoc);
                    }
                    return;
                }
                catch (Exception e)
                {

                }
            }
            Logger.LogErrorAndRaiseException("Get HTML raw data error!");
        }

        private void GetUnderlyingStockCode(HtmlAgilityPack.HtmlDocument htmlDoc)
        {
            try
            {
                underlyingStockCode = htmlDoc.DocumentNode.SelectSingleNode("//table/tr[10]/td[2]").InnerText.Replace("&nbsp;", "").Trim();
            }
            catch (Exception e)
            {
                Logger.LogErrorAndRaiseException("Can't find underlying Stock code", e);
            }
            GetWarrantChainRic(underlyingStockCode);
        }

        private void GetWarrantChainRic(string underlyingStockCode)
        {
            chainRic = "";
            foreach (WarrantStartWith item in xmlConfigObj.WarrantStartWith.Where(item => underlyingStockCode.StartsWith(item.StartWith)))
            {
                chainRic = item.ChainRic;
            }

            if (chainRic.Equals(string.Empty))
            {
                Regex ricRegex = new Regex("[1-9]");
                if ((ricRegex.Match(underlyingStockCode)).Success)
                {
                    for (int length = underlyingStockCode.Length; length < 4; length++)
                    {
                        underlyingStockCode = "0" + underlyingStockCode;
                    }
                    chainRic = "0#" + underlyingStockCode + "W.HK";
                }
            }

            if (chainRic.Equals(string.Empty))
            {
                foreach (WarrantDelistingConfig warrantItem in xmlConfigObj.WarrantDelistingConfig.Where(warrantItem => underlyingStockCode.Equals(warrantItem.UnderlyingChain)))
                {
                    chainRic = warrantItem.ChainRic;
                }
            }

            if (chainRic.Equals(string.Empty))
            {
                chainRic = xmlConfigObj.OthersWarrantChainRic;
            }
        }

        private void GetNameChangeInfo(HtmlAgilityPack.HtmlDocument htmlDoc)
        {
            try
            {
                oldShortName = htmlDoc.DocumentNode.SelectSingleNode("//table/tr/td").InnerText.Replace("&nbsp;", "").Trim();
                int codePos = oldShortName.IndexOf("(");
                oldShortName = oldShortName.Substring(0, codePos).Trim();

                oldLongName = htmlDoc.DocumentNode.SelectSingleNode("//table/tr[2]/td[2]").InnerText.Replace("&nbsp;", "").Trim().TrimEnd('.');
                oldLongName = ChangeNameIntoCamel(oldLongName);

                sector = htmlDoc.DocumentNode.SelectSingleNode("//table/tr[8]/td[2]").InnerText.Replace("&nbsp;", "").Trim();
                sector = sector.Substring(0, sector.IndexOf(' '));
            }
            catch (Exception e)
            {
                Logger.LogErrorAndRaiseException("Grab information from web sit error!", e);
            }
        }

        private string ChangeNameIntoCamel(string name)
        {
            string[] nameParts = name.Split(' ');
            string camelName = nameParts.Select(t => t.Substring(0, 1) + t.Substring(1, t.Length - 1).ToLower()).Aggregate("", (current, formatName) => current + " " + formatName);
            return camelName.Trim();
        }
        #endregion
    }
}
