using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Threading;
using Ric.Util;

namespace Ric.Tasks.HongKong
{

    //public class HKSEAConfig
    //{
    //    public string SOURCE_FILE_PATH { get; set; }
    //    public string ADD_DROP_FILE_PATH { get; set; }
    //    public string RIC_COWS_TEMPLATE_V2_FILE_PATH { get; set; }
    //    public string HKSEA_Z2_FILE_PATH { get; set; }
    //    public string LOG_FILE_PATH { get; set; }
    //    public string SOURCE_FILE_WORKSHEET_NAME { get; set; }
    //    public string ADD_DROP_RAW_WORKSHEET_NAME { get; set; }
    //    public string TEMPLATE_V2_WORKSHEET_NAME { get; set; }
    //    public int HOLIDAY_COUNT { get; set; }
    //}

    //public class BordLotSizeChangeInfo
    //{
    //    public string Code;
    //    public string Name;
    //    public string EventType;
    //    public string BoardLotSize;
    //    public string Date;
    //}

    public class BrokerChangeInfo : IComparable<BrokerChangeInfo>
    {
        public string Name { get; set; }
        public string Ric { get; set; }
        public string Type { get; set; }
        public string Date { get; set; }

        #region IComparable<BrokerChangeInfo> Members

        public int CompareTo(BrokerChangeInfo other)
        {
            return this.Date.CompareTo(other.Date);
        }

        #endregion
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
                    sameDayChangeInfo = new SameDayChangeInfo();
                    sameDayChangeInfo.Date = broker.Date;
                    date = broker.Date;
                    sameDayChangeInfo.NewBrokerList = new List<BrokerChangeInfo>();
                    sameDayChangeInfo.DeleteBrokerList = new List<BrokerChangeInfo>();
                    sameDayChangeInfo.NameChangeToList = new List<BrokerChangeInfo>();
                    sameDayChangeInfo.RicChangeToList = new List<StockNameRicChange>();
                    sameDayChangeInfo.BordLotSizeChangeList = new List<BrokerChangeInfo>();
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

    public class SameDayChangeInfo
    {
        public string Date { get; set; }
        public List<BrokerChangeInfo> NewBrokerList { get; set; }
        public List<BrokerChangeInfo> DeleteBrokerList { get; set; }
        //public List<StockNameRicChange> NameChangeToList { get; set; }
        public List<BrokerChangeInfo> NameChangeToList { get; set; }
        public List<StockNameRicChange> RicChangeToList { get; set; }
        public List<BrokerChangeInfo> BordLotSizeChangeList { get; set; }

        public List<StockNameRicChange> GetChangeToList()
        {
            for (int i = 0; i < NewBrokerList.Count; i++)
            {
                for (int j = 0; j < DeleteBrokerList.Count; j++)
                {

                    if (NewBrokerList[i].Name == DeleteBrokerList[j].Name)
                    {
                        StockNameRicChange stockRicChange = new StockNameRicChange();
                        stockRicChange.NewValue = NewBrokerList[i].Ric;
                        stockRicChange.Name = NewBrokerList[i].Name;
                        stockRicChange.OldValue = DeleteBrokerList[j].Ric;
                        stockRicChange.Date = NewBrokerList[j].Date;
                        stockRicChange.Ric = NewBrokerList[j].Ric;
                        stockRicChange.Type = "RIC CHG TO";
                        RicChangeToList.Add(stockRicChange);
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            return RicChangeToList;
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

    public class HKSEAUpdator : GeneratorBase
    {
        private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_SEAUpdate.config";
        private static Logger logger = null;
        private static HKSEAConfig configObj = null;

        protected override void Start()
        {
            StartSEAUpdaterJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(HKSEAConfig)) as HKSEAConfig;
            logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.LogMode_New);
        }

        public void StartSEAUpdaterJob()
        {
            UpdateFiles();
        }

        public void UpdateFiles()
        {
            int lastUsedRowNum = -1;
            List<BordLotSizeChangeInfo> bordLotSizeChangeList = new List<BordLotSizeChangeInfo>();
            //Range AKRange = GetAKRange(out lastUsedRowNum);
            //Range AKRange = null;


            using (ExcelApp app = new ExcelApp(false, false))
            {
                string addDropFilePath = BackupFiles(configObj.ADD_DROP_FILE_PATH);
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, addDropFilePath);
                Worksheet worksheetRaw = ExcelUtil.GetWorksheet(configObj.ADD_DROP_RAW_WORKSHEET_NAME, workbook);
                using (ExcelApp appSource = new ExcelApp(false, false))
                {
                    var workbookSource = ExcelUtil.CreateOrOpenExcelFile(appSource, configObj.SOURCE_FILE_PATH);
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
                    logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet5 from workbook {0}", workbook.FullName));
                }

                UpdateSheet5(worksheet5, bordLotSizeChangeList);

                var worksheetAddDrop = ExcelUtil.GetWorksheet("ADD DROP", workbook);
                if (worksheetAddDrop == null)
                {
                    logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet ADD DROP from workbook {0}", workbook.FullName));
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

                ChangeInfo ChangedBrokerInfo = GetAllBrokerChangeInfo(worksheet5, lastRowContainInfo - 1);
                UpdateHKSEAFile(ChangedBrokerInfo, bordLotSizeChangeList);
                Range A2Range = ExcelUtil.GetRange(2, 1, lastRowContainInfo - 1, 16, worksheetAddDrop);
                UpdateTemplateV2File(A2Range, lastRowContainInfo);
                workbook.SaveCopyAs(Path.Combine(Path.GetDirectoryName(workbook.FullName), Path.GetFileName(configObj.ADD_DROP_FILE_PATH)));
                workbook.Close(false, workbook.FullName, false);
                File.Delete(addDropFilePath);
            }
        }

        public void UpdateHKSEAFile(ChangeInfo ChangedBrokerInfo, List<BordLotSizeChangeInfo> bordLotSizeChangeList)
        {
            using (ExcelApp appSEA = new ExcelApp(false, false))
            {
                string SEAFilePath = BackupFiles(configObj.HKSEA_Z2_FILE_PATH);
                var workbookSEA = ExcelUtil.CreateOrOpenExcelFile(appSEA, SEAFilePath);
                var worksheetSEA = ExcelUtil.GetWorksheet("Sheet2", workbookSEA);
                int rowPos = 1;
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheetSEA, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    int curRowPos = worksheetSEA.UsedRange.Row + worksheetSEA.UsedRange.Rows.Count - 1;
                    for (int i = ChangedBrokerInfo.DifDayChangeInfo.Count - 1; i > -1; i--)
                    {
                        SameDayChangeInfo sameDayChangedInfo = ChangedBrokerInfo.DifDayChangeInfo[i];
                        rowPos = curRowPos;
                        while (rowPos > 0)
                        {
                            if (ExcelUtil.GetRange(rowPos, 4, worksheetSEA).Value2 == null || string.IsNullOrEmpty(ExcelUtil.GetRange(rowPos, 4, worksheetSEA).Value2.ToString()))
                            {
                                rowPos--;
                                continue;
                            }
                            else if (DateTimeCompare(sameDayChangedInfo.Date, ExcelUtil.GetRange(rowPos, 4, worksheetSEA).Text.ToString()))
                            {
                                if (rowPos == curRowPos)
                                {
                                    curRowPos = rowPos;
                                    InsertChangeInfo(rowPos + 1, writer, sameDayChangedInfo, worksheetSEA, bordLotSizeChangeList);
                                    break;
                                }

                                else if (ExcelUtil.GetRange(rowPos + 1, 4, worksheetSEA) == null || (string.IsNullOrEmpty(ExcelUtil.GetRange(rowPos + 1, 4, worksheetSEA).Text.ToString())))
                                {
                                    rowPos--;
                                }
                                else if (!DateTimeCompare(sameDayChangedInfo.Date, ExcelUtil.GetRange(rowPos + 1, 4, worksheetSEA).Text.ToString()))
                                {
                                    curRowPos = rowPos;
                                    InsertChangeInfo(rowPos + 1, writer, sameDayChangedInfo, worksheetSEA, bordLotSizeChangeList);
                                    break;
                                }

                                else
                                {
                                    rowPos--;
                                    continue;
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

                                else if (ExcelUtil.GetRange(rowPos - 1, 4, worksheetSEA) == null || (ExcelUtil.GetRange(rowPos - 1, 4, worksheetSEA).Text == null))
                                {
                                    rowPos--;
                                }
                                else if (DateTimeCompare(sameDayChangedInfo.Date, ExcelUtil.GetRange(rowPos - 1, 4, worksheetSEA).Text.ToString()))
                                {
                                    curRowPos = rowPos;
                                    InsertChangeInfo(rowPos + 1, writer, sameDayChangedInfo, worksheetSEA, bordLotSizeChangeList);
                                    break;
                                }
                                else
                                {
                                    rowPos--;
                                    continue;
                                }
                            }
                        }
                    }
                }
                //workbookSEA.SaveCopyAs(Path.Combine(Path.GetDirectoryName(workbookSEA.FullName), Path.GetFileName(configObj.HKSEA_Z2_FILE_PATH)));
                workbookSEA.SaveCopyAs(configObj.HKSEA_Z2_FILE_PATH);
                workbookSEA.Close(false, workbookSEA.FullName, false);
                //File.Delete(SEAFilePath);
            }
        }

        public void InsertChangeInfo(int rowPos, ExcelLineWriter writer, SameDayChangeInfo sameDayChangeInfo, Worksheet worksheet, List<BordLotSizeChangeInfo> bordLotSizeChangeInfoList)
        {
            InsertNewDeleteChangeInfo(--rowPos, writer, sameDayChangeInfo.DeleteBrokerList, worksheet);
            int curPos = rowPos;
            while (curPos > 0)
            {
                Range r = ExcelUtil.GetRange(curPos - 1, 4, worksheet);
                if (r.Text == null || r.Text.ToString() == "")
                {
                    curPos--;
                    continue;
                }
                else if (DateTimeEqual(sameDayChangeInfo.Date, r.Text.ToString()))
                {
                    curPos--;
                    continue;
                }

                else
                {
                    InsertNameChange(curPos, writer, sameDayChangeInfo.NameChangeToList, worksheet);
                    InsertRicChange(curPos, writer, sameDayChangeInfo.RicChangeToList, worksheet);
                    InsertBordSizeChangeInfo(curPos, writer, sameDayChangeInfo.BordLotSizeChangeList, bordLotSizeChangeInfoList, worksheet);
                    InsertNewDeleteChangeInfo(curPos, writer, sameDayChangeInfo.NewBrokerList, worksheet);
                    break;
                }
            }
            int lastUsedRowNum = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            ExcelUtil.GetRange(1, 4, lastUsedRowNum, 4, worksheet).NumberFormat = "MM/dd/yyyy";
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
                foreach (BordLotSizeChangeInfo bordLotSizeChangeInfo in bordLotSizeChangeInfoList)
                {
                    if (bordLotSizeChangeInfo.Name == broker.Name)
                    {
                        type = string.Format("Bord Lot Size Change to {0}", bordLotSizeChangeInfo.BoardLotSize);
                        break;
                    }
                    else continue;
                }
                writer.WriteLine(type);
                writer.WriteLine(broker.Date);
                if (IsIPos(broker.Ric))
                {
                    ExcelUtil.GetRange(rowPos - 1, 1, rowPos - 1, lastUsedCol, worksheet).Interior.Color = 10092543.0;
                }

                else
                {
                    ExcelUtil.GetRange(rowPos - 1, 1, rowPos - 1, lastUsedCol, worksheet).Interior.Color = 16777215.0;
                }
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
                if (IsIPos(broker.Ric))
                {
                    //Mark as yellow
                    ExcelUtil.GetRange(rowPos - 1, 1, rowPos - 1, lastUsedCol, worksheet).Interior.Color = 10092543.0;
                }

                else
                {
                    ExcelUtil.GetRange(rowPos - 1, 1, rowPos - 1, lastUsedCol, worksheet).Interior.Color = 16777215.0;
                }
            }
            ExcelUtil.GetRange(rowPos - brokerChangeList.Count, 4, rowPos - 1, 4, worksheet).NumberFormat = "MM/dd/yyyy";
            if (brokerChangeList[0].Type == "delete")
            {
                if (ParseDateTime(brokerChangeList[0].Date) == GetNextBusinessDay(DateTime.Now, configObj.HOLIDAY_COUNT))
                {
                    //Mark as red
                    ExcelUtil.GetRange(rowPos - brokerChangeList.Count, 1, rowPos - 1, lastUsedCol, worksheet).Font.Color = 255.0;
                }
                else
                    ExcelUtil.GetRange(rowPos - brokerChangeList.Count, 1, rowPos - 1, lastUsedCol, worksheet).Font.Color = 0.0;
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
            else
            {
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
                    if (IsIPos(nameChange.Ric))
                    {
                        ExcelUtil.GetRange(rowPos - 2, lastUsedCol, rowPos - 1, lastUsedCol, worksheet).Interior.Color = 10092543.0;
                    }
                    else
                        ExcelUtil.GetRange(rowPos - 2, lastUsedCol, rowPos - 1, lastUsedCol, worksheet).Interior.Color = 16777164.0;
                }
            }
        }

        public void InsertRicChange(int rowPos, ExcelLineWriter writer, List<StockNameRicChange> ricChangeList, Worksheet worksheet)
        {
            if (ricChangeList == null || ricChangeList.Count == 0)
            {
                return;
            }

            else
            {
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
                    if (IsIPos(ricChange.Ric))
                    {
                        ExcelUtil.GetRange(rowPos - 2, lastUsedCol, rowPos - 1, lastUsedCol, worksheet).Interior.Color = 10092543.0;
                    }
                    else
                        ExcelUtil.GetRange(rowPos - 2, lastUsedCol, rowPos - 1, lastUsedCol, worksheet).Interior.Color = 16777164.0;
                }
            }
        }

        public bool IsUpgrade(StockNameRicChange ricChange)
        {
            if (ricChange.OldValue.Trim().StartsWith("8") && (!ricChange.NewValue.Trim().StartsWith("8")))
            {
                return true;
            }
            else
                return false;
        }
        public bool IsIPos(string ricStr)
        {
            int ricInt = 0;
            ricStr = ricStr.Replace(".HK", "").Replace("<", "").Replace(">", "");
            bool result = int.TryParse(ricStr, out ricInt);
            if (!result)
            {
                logger.LogErrorAndRaiseException(string.Format("Can't parse {0}", ricStr));
            }
            else
            {
                int temp = ricInt / 10000;
                if (temp == 1 || temp == 2 || temp == 6)
                {
                    result = false; ;
                }
                else
                    result = true;
            }
            return result;
        }

        public ChangeInfo GetAllBrokerChangeInfo(Worksheet worksheet, int lastRowContainInfo)
        {
            ChangeInfo ChangedBrokerInfo = new ChangeInfo();
            ChangedBrokerInfo.AllChangedBrokerList = new List<BrokerChangeInfo>();
            for (int rowNum = 2; rowNum <= lastRowContainInfo; rowNum++)
            {
                BrokerChangeInfo brokerChangeInfo = new BrokerChangeInfo();
                brokerChangeInfo.Name = ExcelUtil.GetRange(rowNum, 1, worksheet).Value2.ToString();
                brokerChangeInfo.Ric = ExcelUtil.GetRange(rowNum, 2, worksheet).Value2.ToString();
                brokerChangeInfo.Type = ExcelUtil.GetRange(rowNum, 3, worksheet).Value2.ToString();
                brokerChangeInfo.Date = ExcelUtil.GetRange(rowNum, 4, worksheet).Value2.ToString();
                ChangedBrokerInfo.AllChangedBrokerList.Add(brokerChangeInfo);
            }
            ChangedBrokerInfo.DifDayChangeInfo = ChangedBrokerInfo.GetDifDayChangeInfo();
            //ChangedBrokerInfo.AllChangedBrokerList.Sort();
            return ChangedBrokerInfo;
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

                    else
                    {
                        if (bordLotSizeChangeList != null && bordLotSizeChangeList.Count != 0)
                        {
                            if (ExcelUtil.GetRange(writer.Row, 3, worksheet5).Text.ToString().Trim() == "#N/A")
                            {
                                writer.PlaceNextAndWriteLine(writer.Row, 3, "CHANGE");
                                writer.PlaceNext(writer.Row - 1, writer.Col);
                            }
                        }
                        Range r = ExcelUtil.GetRange(writer.Row, 2, worksheet5);
                        string line = r.Value2.ToString().Trim(new char[] { '>', '<' });
                        writer.PlaceNextAndWriteLine(writer.Row, 2, line);
                    }
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

                    else
                    {
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
        }

        public Range GetAKRange(Workbook workbookSource, out int lastUsedRowNum, out List<BordLotSizeChangeInfo> bordLotSizeChangeList)
        {
            Range AKRange = null;
            var worksheetSource = ExcelUtil.GetWorksheet(configObj.SOURCE_FILE_WORKSHEET_NAME, workbookSource);
            if (worksheetSource == null)
            {
                logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.SOURCE_FILE_WORKSHEET_NAME, workbookSource.Name));
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
            AKRange = ExcelUtil.GetRange(1, 1, lastUsedRowNum, 11, worksheetSource);
            return AKRange;
        }

        public void UpdateTemplateV2File(Range A2Range, int lastRowContainInfo)
        {
            A2Range.Copy(Missing.Value);
            string TemplateV2FilePath = BackupFiles(configObj.RIC_COWS_TEMPLATE_V2_FILE_PATH);
            using (ExcelApp appTemplateV2 = new ExcelApp(false, false))
            {
                var workbookTemplateV2 = ExcelUtil.CreateOrOpenExcelFile(appTemplateV2, TemplateV2FilePath);
                var worksheetTemplateV2 = ExcelUtil.GetWorksheet(configObj.TEMPLATE_V2_WORKSHEET_NAME, workbookTemplateV2);
                if (worksheetTemplateV2 == null)
                {
                    logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.TEMPLATE_V2_WORKSHEET_NAME, workbookTemplateV2.Name));
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

                workbookTemplateV2.SaveCopyAs(Path.Combine(Path.GetDirectoryName(workbookTemplateV2.FullName), Path.GetFileName(configObj.RIC_COWS_TEMPLATE_V2_FILE_PATH)));
                workbookTemplateV2.Close(false, workbookTemplateV2.FullName, false);
                File.Delete(TemplateV2FilePath);
            }
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
                    else
                        continue;
                }
                else
                {
                    continue;
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
            dir += "//";
            dir += DateTime.Now.ToString("yyyy-MM-dd");
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
            DateTime dateTime1 = ParseDateTime(time1);
            DateTime dateTime2 = ParseDateTime(time2);
            return dateTime1 >= dateTime2;
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
            List<string> formatList = new List<string>();
            formatList.Add("dd/MM/yyyy");
            formatList.Add("d/MM/yyyy");
            formatList.Add("d/M/yyyy");
            formatList.Add("dd/M/yyyy");
            DateTime dateTime = DateTime.MinValue;
            bool result = false;
            int count = formatList.Count;
            while (result == false && count >= 0)
            {
                foreach (string format in formatList)
                {
                    result = DateTime.TryParseExact(time, format, null, System.Globalization.DateTimeStyles.None, out dateTime);
                    if (result == true)
                    {
                        break;
                    }
                    count--;
                }
            }
            if (result == false)
            {
                logger.LogErrorAndRaiseException(string.Format("Parse DateTime Error, {0}", time));
            }
            return dateTime;
        }
    }
}
