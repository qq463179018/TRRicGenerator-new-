using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;
using System.Threading;

namespace Ric.Tasks.HongKong
{
    //Items user can be configured
    [ConfigStoredInDB]
    public class HKBrokerPageFMConfig
    {
        [StoreInDB]
        public string CHINESE_VERSION_XML_FILE_PATH { get; set; }

        [StoreInDB]
        public string ENGLISH_VERSION_XML_FILE_PATH { get; set; }

        [StoreInDB]
        public string PREVIOUS_SEHKFILE_PATH { get; set; }

        [StoreInDB]
        public string CURRENT_SEHKFILE_PATH { get; set; }

        [StoreInDB]
        public string BROKER3H_FILE_PATH { get; set; }

        [StoreInDB]
        public string LOG_FILE_PATH { get; set; }

        [StoreInDB]
        public string MASTER_WORKSHEET_NAME { get; set; }

        [StoreInDB]
        public string INSERT_WORKSHEET_NAME { get; set; }

        [StoreInDB]
        public string CHINESE_VERSION_FILE_PATH { get; set; }
    }

    public class HKBrokerPageNameMap
    {
        public string longName { get; set; }
        public string shortName { get; set; }
    }

    public class BrokerInfo : IComparable<BrokerInfo>
    {
        public string originalName { get; set; }
        public string ChineseName { get; set; }
        public string ChineseShortName { get; set; }
        public string ChineseLongName { get; set; }
        public string longName { get; set; }
        public string shortName { get; set; }
        public string brokerNumbers { get; set; }

        public string err = string.Empty;

        public string hall { get; set; }
        public string terminal { get; set; }

        public int GroupId { get; set; }
        public List<int> IdList { get; set; }

        private static readonly string NAME_MAP_FILE_PATH = ".\\Config\\HK\\HK_NameMap.xls";
        private static Dictionary<string, HKBrokerPageNameMap> chineseNameMap = new Dictionary<string, HKBrokerPageNameMap>();
        private static Dictionary<string, HKBrokerPageNameMap> englishNameMap = new Dictionary<string, HKBrokerPageNameMap>();


        public BrokerInfo()
        {
            if (chineseNameMap.Keys.Count == 0 || englishNameMap.Keys.Count == 0)
            {
                getNameMap();
            }
        }
        public string GetHallInfo()
        {
            List<int> unitDigitList = new List<int>();
            string hallInfo = string.Empty;
            if (IdList[0] > 3)
            {
                hallInfo = ".";
            }
            else
            {
                hallInfo = ((GroupId * 10) + IdList[0]).ToString("D4");
                for (int i = 1; i < IdList.Count; i++)
                {
                    if (IdList[i] > 3)
                    {
                        break;
                    }
                    else
                    {
                        if (IdList[i] == IdList[i - 1] + 1)
                        {
                            if (hallInfo[hallInfo.Length - 2] != '/')
                            {
                                hallInfo = hallInfo + '/';
                                hallInfo = hallInfo + IdList[i].ToString();
                                continue;
                            }
                            else
                            {
                                hallInfo = hallInfo.Remove(hallInfo.Length - 1);
                                hallInfo = hallInfo + IdList[i].ToString();
                            }
                        }

                        else
                        {
                            hallInfo = hallInfo + ",";
                            hallInfo = hallInfo + IdList[i].ToString();
                        }
                    }
                }
            }
            return hallInfo;
        }

        public string GetTerminalInfo()
        {
            string terminalInfo = string.Empty;
            if (IdList[IdList.Count - 1] < 4)
            {
                terminalInfo = ".";
            }

            else
            {
                int startPos = 0;
                for (int i = 0; i < IdList.Count; i++)
                {
                    if (IdList[i] < 4)
                    {
                        continue;
                    }
                    else
                    {
                        startPos = i;
                        terminalInfo = (GroupId * 10 + IdList[i]).ToString("D4");
                        break;
                    }
                }

                for (int i = startPos + 1; i < IdList.Count; i++)
                {
                    if (IdList[i] == IdList[i - 1] + 1)
                    {
                        if (terminalInfo[terminalInfo.Length - 2] != '/')
                        {
                            terminalInfo = terminalInfo + '/';
                            terminalInfo = terminalInfo + IdList[i].ToString();
                            continue;
                        }
                        else
                        {
                            terminalInfo = terminalInfo.Remove(terminalInfo.Length - 1);
                            terminalInfo = terminalInfo + IdList[i].ToString();
                        }
                    }

                    else
                    {
                        terminalInfo = terminalInfo + ",";
                        terminalInfo = terminalInfo + IdList[i].ToString();
                    }
                }
            }
            return terminalInfo;
        }

        public void getNameMap()
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, NAME_MAP_FILE_PATH);
                var worksheet = (Worksheet)workbook.Worksheets[1];
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    while (reader.Row <= lastUsedRow)
                    {
                        string englishOrigName = reader.ReadLineCellText().ToUpper();
                        if (!englishNameMap.ContainsKey(englishOrigName))
                        {
                            englishNameMap.Add(englishOrigName, new HKBrokerPageNameMap());
                            englishNameMap[englishOrigName].longName = reader.ReadLineCellText().ToUpper();
                            englishNameMap[englishOrigName].shortName = reader.ReadLineCellText().ToUpper();
                        }

                        else
                        {
                            reader.PlaceNext(reader.Row, reader.Col + 2);
                        }

                        reader.MoveNext();
                        string chineseOrigName = reader.ReadLineCellText().ToUpper();
                        if (!chineseNameMap.ContainsKey(chineseOrigName))
                        {
                            chineseNameMap.Add(chineseOrigName, new HKBrokerPageNameMap());
                            chineseNameMap[chineseOrigName].longName = reader.ReadLineCellText().ToUpper();
                            chineseNameMap[chineseOrigName].shortName = reader.ReadLineCellText().ToUpper();
                        }
                        reader.PlaceNext(reader.Row + 1, 1);
                    }
                }
                workbook.Close(false, workbook.FullName, false);
            }
        }

        public override string ToString()
        {
            return originalName + ": " + brokerNumbers;
        }

        //Get broker English long name and short name
        public void GetBrokerLongAndShortName()
        {
            originalName = originalName.ToUpper().Trim();
            if (englishNameMap.ContainsKey(originalName))
            {
                this.longName = englishNameMap[originalName].longName;
                this.shortName = englishNameMap[originalName].shortName;
            }
            else
            {
                err = string.Format("There's no such name \"{0}\", please check the HK_NAMEMAP.xls", originalName);
            }

        }

        //Get short Chinese name for broker
        public void GetBrokerChineseLongAndShortName()
        {
            if (chineseNameMap.ContainsKey(ChineseName))
            {
                this.ChineseLongName = chineseNameMap[ChineseName].longName;
                this.ChineseShortName = chineseNameMap[ChineseName].shortName;
            }
            else
            {
                err = string.Format("There's no such name \"{0}\", please check the HK_NAMEMAP.xls", ChineseName);
            }
        }

        #region IComparable<BrokerInfo> Members

        public int CompareTo(BrokerInfo other)
        {
            return this.GroupId.CompareTo(other.GroupId);
        }

        #endregion
    }

    public class PageInfo
    {
        public string minNumberStr;
        public string maxNumberStr;
        public string pageNumberStr;
        public string pageNumberChineseStr;
    }

    //Generator for broker.3h file
    public class HKBrokerPageFM : GeneratorBase
    {
        //private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HKBrokerPageFM.config";
        //private static Logger logger = null;
        private static HKBrokerPageFMConfig configObj = null;

        private static readonly string BASE_URI = "http://www.hkex.com.hk/eng/plw/plw_search.asp";

        protected override void Start()
        {
            StartBrokerPageFMJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as HKBrokerPageFMConfig;



            //configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(HKBrokerPageFMConfig)) as HKBrokerPageFMConfig;
        }

        public void StartBrokerPageFMJob()
        {
            //For testing
            //string previousCSVFilePath = Path.Combine(Path.Combine(Path.GetDirectoryName(configObj.CURRENT_SEHKFILE_PATH), "Bak"), Path.GetFileName(configObj.CURRENT_SEHKFILE_PATH));
            //List<BrokerInfo> brokerList = GetBrokerList(previousCSVFilePath);

            DownLoadSCVFile();
            Thread.Sleep(2000);
            List<BrokerInfo> brokerList = GetBrokerList(configObj.CURRENT_SEHKFILE_PATH);

            Logger.Log("Begin to Generate XML Format Files");

            Thread.Sleep(2000);
            GenerateChineseVersionXmlFile(brokerList, configObj.CHINESE_VERSION_XML_FILE_PATH);
            Thread.Sleep(2000);
            GenerateEnglishVersionXmlFile(brokerList, configObj.ENGLISH_VERSION_XML_FILE_PATH);
            Thread.Sleep(2000);

            Logger.Log("Begin to Compare the Current Day SEHK File and the Previous Day SEHK File");
            Logger.Log("**************************************************************************");
            Thread.Sleep(2000);
            string previousCSVFilePath = Path.Combine(Path.Combine(Path.GetDirectoryName(configObj.CURRENT_SEHKFILE_PATH), "Bak"), Path.GetFileName(configObj.CURRENT_SEHKFILE_PATH));
            List<BrokerInfo> previousBrokerList = GetBrokerList(previousCSVFilePath);
            Thread.Sleep(2000);
            CompareSEHKFile(previousBrokerList, brokerList);
            Logger.Log("**************************************************************************");
            Thread.Sleep(2000);
            GenerateBroker3hFile(brokerList);
        }

        private string formatString(string originalStr, int length)
        {
            string str = originalStr;
            int del = length;
            foreach (char c in originalStr)
            {
                if (Regex.IsMatch(c.ToString(), @"[\u4e00-\u9fbb]+$"))
                {
                    del -= 2;
                }

                else
                {
                    del -= 1;
                }
            }

            if (del > 0)
            {
                for (int i = 0; i < del; i++)
                {
                    str += " ";
                }
            }
            return str;
        }

        private string formatStringForChinese(string originalStr, int width, int allBytes)
        {
            string str = formatString(originalStr, width);

            while (Encoding.UTF8.GetBytes(str).Length > allBytes)
            {
                str = str.Remove(str.Length - 1);
            }
            return str;
        }

        private string getFidValueForChineseVersionFile(BrokerInfo brokerInfo)
        {
            string value = string.Empty;
            value += "\"";
            value += brokerInfo.hall.PadRight(9).Substring(0, 8);
            value += brokerInfo.terminal.PadRight(10).Substring(0, 10);
            if (!string.IsNullOrEmpty(brokerInfo.ChineseName))
            {
                if (!string.IsNullOrEmpty(brokerInfo.ChineseLongName))
                {
                    value += formatStringForChinese(brokerInfo.ChineseLongName, 42, 58);

                    int bytesNumForChineseShortName = 96 - Encoding.UTF8.GetBytes(value).Length + 1;
                    value += formatStringForChinese(brokerInfo.ChineseShortName, 24, bytesNumForChineseShortName);
                }
                if (string.IsNullOrEmpty(brokerInfo.ChineseLongName))
                {
                    value += formatStringForChinese(brokerInfo.ChineseName, 42, 58);
                    int bytesNumForChineseShortName = 96 - Encoding.UTF8.GetBytes(value).Length + 1;
                    value += formatStringForChinese(brokerInfo.ChineseName, 24, bytesNumForChineseShortName);
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(brokerInfo.longName))
                {
                    value += formatStringForChinese(brokerInfo.longName, 42, 58);
                    int bytesNumForChineseShortName = 96 - Encoding.UTF8.GetBytes(value).Length + 1;
                    value += formatStringForChinese(brokerInfo.shortName, 24, bytesNumForChineseShortName);
                }
                else
                {
                    value += formatStringForChinese(brokerInfo.originalName, 42, 58);
                    int bytesNumForChineseShortName = 96 - Encoding.UTF8.GetBytes(value).Length + 1;
                    value += formatStringForChinese(brokerInfo.originalName, 24, bytesNumForChineseShortName);
                }
            }
            value += "\"";
            return value;
        }

        private string getFidValueForEnglishVersionFile(BrokerInfo brokerInfo)
        {
            string value = string.Empty;
            value += "\"";
            value += brokerInfo.hall.PadRight(9).Substring(0, 9);
            value += brokerInfo.terminal.PadRight(10).Substring(0, 10);
            if (!string.IsNullOrEmpty(brokerInfo.longName))
            {
                value += brokerInfo.longName.PadRight(45).Substring(0, 45);
                value += brokerInfo.shortName.PadRight(16).Substring(0, 16);
            }
            else
            {
                value += brokerInfo.originalName.PadRight(45).Substring(0, 45);
                value += brokerInfo.originalName.PadRight(16).Substring(0, 16);
            }
            value += "\"";
            return value;
        }

        //17 lines each page
        private void GenerateChineseVersionXmlFile(List<BrokerInfo> brokerList, string xmlFilePath)
        {
            int lineNumEachPage = 17;
            int lastAddedRicNum = 0;
            ProductionXmlFileTemplate template = new ProductionXmlFileTemplate();

            for (int i = 0; i < lineNumEachPage * 44; i++)
            {
                if (i % lineNumEachPage == 0)
                {
                    Core.Ric ric = new Core.Ric();
                    ric.Name = string.Format("HKBKC{0}", ((i / lineNumEachPage) + 2).ToString("D2"));
                    template.rics.rics.Add(ric);
                }
                Fid fid = new Fid();
                fid.Id = 1361 + i % lineNumEachPage;
                fid.Locale = true;
                if (i < brokerList.Count)
                {
                    fid.Value = getFidValueForChineseVersionFile(brokerList[i]);
                }

                else
                {
                    fid.Value = "\"\"";
                }
                lastAddedRicNum = template.rics.rics.Count - 1;
                template.rics.rics[lastAddedRicNum].fids.Add(fid);
            }


            ConfigUtil.WriteXml(xmlFilePath, template);
            TaskResultList.Add(new TaskResultEntry("Chinese Version XML Format File", "", xmlFilePath));

        }

        //17 lines each page
        private void GenerateEnglishVersionXmlFile(List<BrokerInfo> brokerList, string xmlFilePath)
        {
            int lineNumEachPage = 22;
            int lastAddedRicNum = 0;
            ProductionXmlFileTemplate template = new ProductionXmlFileTemplate();
            for (int i = 0; i < lineNumEachPage * 34; i++)
            {
                if (i % lineNumEachPage == 0)
                {
                    Core.Ric ric = new Core.Ric();
                    ric.Name = string.Format("HKBK{0}", ((i / lineNumEachPage) + 2).ToString("D2"));
                    template.rics.rics.Add(ric);
                }
                Fid fid = new Fid();
                fid.Id = 317 + i % lineNumEachPage;
                if (i < brokerList.Count)
                {
                    fid.Value = getFidValueForEnglishVersionFile(brokerList[i]);
                }
                else
                {
                    fid.Value = "\"\"";
                }
                lastAddedRicNum = template.rics.rics.Count - 1;
                template.rics.rics[lastAddedRicNum].fids.Add(fid);
            }

            ConfigUtil.WriteXml(xmlFilePath, template);
            TaskResultList.Add(new TaskResultEntry("English Version XML Format File", "", xmlFilePath));

        }

        //Compare two SEHK file
        public void CompareSEHKFile(List<BrokerInfo> previousDayBrokerList, List<BrokerInfo> currentDayBrokerList)
        {
            Dictionary<string, BrokerInfo> previousMap = new Dictionary<string, BrokerInfo>();
            Dictionary<string, BrokerInfo> currentMap = new Dictionary<string, BrokerInfo>();
            foreach (BrokerInfo broker in previousDayBrokerList)
            {
                if (!previousMap.ContainsKey(broker.originalName))
                {
                    previousMap.Add(broker.originalName, broker);
                }
            }
            foreach (BrokerInfo broker in currentDayBrokerList)
            {
                if (!currentMap.ContainsKey(broker.originalName))
                {
                    currentMap.Add(broker.originalName, broker);
                }
            }

            List<string> allKeys = new List<string>(previousMap.Keys.Union(currentMap.Keys));
            foreach (string key in allKeys)
            {
                bool previousFound = previousMap.ContainsKey(key);
                bool currentFound = currentMap.ContainsKey(key);

                if (previousFound && currentFound)
                {
                    if (previousMap[key].brokerNumbers.CompareTo(currentMap[key].brokerNumbers) != 0)
                    {
                        Logger.Log("PreviousDay: " + previousMap[key].ToString() + " || " + "CurrentDay: " + currentMap[key].ToString());
                    }
                }
                else if (previousFound)
                {
                    Logger.Log("PreviousDay: " + previousMap[key].ToString() + " || " + "CurrentDay: Null");
                }
                else
                {
                    Logger.Log("PreviousDay: Null" + " || " + "CurrentDay: " + currentMap[key].ToString());
                }
            }
        }

        //Generate broker3h file
        public void GenerateBroker3hFile(List<BrokerInfo> brokerList)
        {
            string broker3hFilePath = MiscUtil.BackupFileWithNewName(configObj.BROKER3H_FILE_PATH);
            if (File.Exists(configObj.CHINESE_VERSION_FILE_PATH))
            {
                string ChineseVersionFileBakFilePath = MiscUtil.BackupFileWithNewName(configObj.CHINESE_VERSION_FILE_PATH);
            }
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, broker3hFilePath);
                if (workbook.ReadOnly == true)
                {
                    workbook.ReadOnlyRecommended = false;
                }
                int brokerRowsSum = 0;
                Logger.Log("**************************************************************************");

                //Copy the content in Master sheet in Insert sheet
                Logger.Log("Begin to Generate the Insert Sheet");
                GenerateInsertSheet(workbook, brokerRowsSum, brokerList);

                Logger.Log("**************************************************************************");

                //Generate Chinese version File
                workbook.SaveCopyAs(configObj.CHINESE_VERSION_FILE_PATH);
                TaskResultList.Add(new TaskResultEntry("Broker3h_Chinese.xls", "Chinese Version Broker3h file", configObj.CHINESE_VERSION_FILE_PATH));
                UpdateCColAsEnglish(brokerList, ExcelUtil.GetWorksheet(configObj.MASTER_WORKSHEET_NAME, workbook));
                UpdateDColAsEnglish(brokerList, ExcelUtil.GetWorksheet(configObj.MASTER_WORKSHEET_NAME, workbook));
                UpdateCColAsEnglish(brokerList, ExcelUtil.GetWorksheet(configObj.INSERT_WORKSHEET_NAME, workbook));
                UpdateDColAsEnglish(brokerList, ExcelUtil.GetWorksheet(configObj.INSERT_WORKSHEET_NAME, workbook));
                UpdatePageId(ExcelUtil.GetWorksheet(configObj.INSERT_WORKSHEET_NAME, workbook));

                workbook.SaveCopyAs(configObj.BROKER3H_FILE_PATH);
                TaskResultList.Add(new TaskResultEntry("Broker3h.xls", "English Version Broker3h file", configObj.BROKER3H_FILE_PATH));
                workbook.Close(false, workbook.FullName, false);
            }
        }

        //Update C column value as the English version name
        public void UpdateCColAsEnglish(List<BrokerInfo> brokerList, Worksheet worksheet)
        {
            int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            ExcelUtil.GetRange(4, 2, lastUsedRow, 4, worksheet).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.White);
            using (ExcelLineWriter EnglishWriter = new ExcelLineWriter(worksheet, 4, 3, ExcelLineWriter.Direction.Down))
            {
                foreach (BrokerInfo broker in brokerList)
                {
                    if (!string.IsNullOrEmpty(broker.longName))
                    {
                        EnglishWriter.WriteLine(broker.longName);
                    }

                    else
                    {
                        EnglishWriter.WriteLine(broker.originalName);
                        ExcelUtil.GetRange(EnglishWriter.Row, EnglishWriter.Col, worksheet).Interior.Color = 255.0;
                    }
                }
            }
        }

        //Update page Id from HKBKC to HKBK for English version broker3h file.
        public void UpdatePageId(Worksheet worksheet)
        {
            int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

            using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 4, 5, ExcelLineWriter.Direction.Down))
            {
                while (writer.Row <= lastUsedRow)
                {
                    Range r = ExcelUtil.GetRange(writer.Row, 5, worksheet);
                    while (r.Text != null && r.Text.ToString().Contains("HKBKC"))
                    {
                        writer.WriteLine(r.Text.ToString().Replace("HKBKC", "HKBK"));
                        writer.PlaceNext(writer.Row + 15, 5);
                        writer.WriteLine(ExcelUtil.GetRange(writer.Row, 5, worksheet).Text.ToString().Replace("HKBKC", "HKBK"));
                    }
                    break;
                }

                for (int i = 0; i < 3; i++)
                {
                    writer.PlaceNext(2, 9 + i * 3);
                    while (ExcelUtil.GetRange(writer.Row, 7, worksheet).Text != null && !ExcelUtil.GetRange(writer.Row, 7, worksheet).Text.ToString().Contains("REFER"))
                    {
                        Range referRange = ExcelUtil.GetRange(writer.Row, writer.Col, worksheet);
                        if (referRange.Text != null)
                        {
                            writer.WriteLine(referRange.Text.ToString().Replace("HKBKC", "HKBK"));
                        }
                    }
                }
            }
        }

        //Update D column value as the English version name
        public void UpdateDColAsEnglish(List<BrokerInfo> brokerList, Worksheet worksheet)
        {
            int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            ExcelUtil.GetRange(4, 2, lastUsedRow, 4, worksheet).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.White);

            using (ExcelLineWriter EnglishWriter = new ExcelLineWriter(worksheet, 4, 4, ExcelLineWriter.Direction.Down))
            {
                foreach (BrokerInfo broker in brokerList)
                {
                    if (!string.IsNullOrEmpty(broker.shortName))
                    {
                        EnglishWriter.WriteLine(broker.shortName);
                    }

                    else
                    {
                        EnglishWriter.WriteLine("");
                        ExcelUtil.GetRange(EnglishWriter.Row, EnglishWriter.Col, worksheet).Interior.Color = 255.0;
                    }
                }
            }
        }

        //Generate master sheet
        public Range GenerateMasterSheet(Workbook workbook, List<BrokerInfo> brokerList, out int brokerRowsSum)
        {
            var worksheetMaster = ExcelUtil.GetWorksheet(configObj.MASTER_WORKSHEET_NAME, workbook);
            if (worksheetMaster == null)
            {
                LogMessage(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.MASTER_WORKSHEET_NAME, workbook.Name));
            }

            int lastUsedRow = worksheetMaster.UsedRange.Row + worksheetMaster.UsedRange.Rows.Count - 1;

            ExcelUtil.GetRange(1, 2, lastUsedRow, 4, worksheetMaster).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.White);

            Range copyRange = null;
            using (ExcelLineWriter writerMaster = new ExcelLineWriter(worksheetMaster, 4, 1, ExcelLineWriter.Direction.Right))
            {
                foreach (BrokerInfo brokerInfo in brokerList)
                {
                    writerMaster.WriteLine(brokerInfo.hall);
                    writerMaster.WriteLine(brokerInfo.terminal);

                    if (!string.IsNullOrEmpty(brokerInfo.ChineseName))
                    {
                        if (!string.IsNullOrEmpty(brokerInfo.ChineseLongName))
                        {
                            writerMaster.WriteLine(brokerInfo.ChineseLongName);
                            writerMaster.WriteLine(brokerInfo.ChineseShortName);
                        }
                        if (string.IsNullOrEmpty(brokerInfo.ChineseLongName))
                        {
                            writerMaster.WriteLine(brokerInfo.ChineseName);
                            writerMaster.WriteLine(brokerInfo.ChineseName);
                            ExcelUtil.GetRange(writerMaster.Row, writerMaster.Col - 1, writerMaster.Row, writerMaster.Col, worksheetMaster).Interior.Color = 255;//Highlight as red
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(brokerInfo.longName))
                        {
                            writerMaster.WriteLine(brokerInfo.longName);
                            writerMaster.WriteLine(brokerInfo.shortName);
                        }
                        else
                        {
                            writerMaster.WriteLine(brokerInfo.originalName);
                            writerMaster.WriteLine(brokerInfo.originalName);
                            ExcelUtil.GetRange(writerMaster.Row, writerMaster.Col - 1, writerMaster.Row, writerMaster.Col, worksheetMaster).Interior.Color = 255;
                        }
                    }
                    writerMaster.PlaceNext(writerMaster.Row + 1, 1);
                }
                Range r = ExcelUtil.GetRange(4, 1, writerMaster.Row - 1, 1, worksheetMaster);
                r.Interior.Color = 13409484.0;
                brokerRowsSum = writerMaster.Row - 1;
                if (lastUsedRow > writerMaster.Row)
                {
                    Range range = ExcelUtil.GetRange(writerMaster.Row, 1, lastUsedRow, 4, worksheetMaster);
                    range.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                }
                writerMaster.WriteLine("STOP");
                writerMaster.WriteLine("STOP");
                writerMaster.WriteLine("STOP");
                writerMaster.WriteLine("STOP");
                writerMaster.PlaceNext(writerMaster.Row + 1, 2);
                writerMaster.WriteLine("STANDARD SECURITIES ASIA LTD");
                copyRange = ExcelUtil.GetRange(4, 1, brokerRowsSum, 4, worksheetMaster);
            }
            return copyRange;
        }

        //Generate insert sheet 
        public void GenerateInsertSheet(Workbook workbook, int brokerRowsSum, List<BrokerInfo> brokerList)
        {
            var worksheetInsert = ExcelUtil.GetWorksheet(configObj.INSERT_WORKSHEET_NAME, workbook);
            if (worksheetInsert == null)
            {
                LogMessage(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.INSERT_WORKSHEET_NAME, workbook.Name));
            }
            Range copyRange = GenerateMasterSheet(workbook, brokerList, out brokerRowsSum);
            using (ExcelLineWriter writerInsert = new ExcelLineWriter(worksheetInsert, 4, 1, ExcelLineWriter.Direction.Right))
            {
                Range r = ExcelUtil.GetRange(4, 1, brokerRowsSum, 4, worksheetInsert);
                copyRange.Copy(r);

                int lastUsedRowInsert = worksheetInsert.UsedRange.Row + worksheetInsert.UsedRange.Rows.Count - 1;
                if (lastUsedRowInsert > (brokerRowsSum))
                {
                    ExcelUtil.GetRange(brokerRowsSum + 1, 1, lastUsedRowInsert, 4, worksheetInsert).Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                }

                ExcelUtil.GetRange(4, 5, lastUsedRowInsert, 5, worksheetInsert).ClearContents();
                int pageSum = ((brokerRowsSum - 4)) / 17 + 1;
                List<PageInfo> pageList = new List<PageInfo>();
                for (int i = 0; i < pageSum; i++)
                {
                    PageInfo pageInfo = new PageInfo();
                    string id = "HKBKC" + (i + 2).ToString("D2");
                    int upRow = 17 * i + 4;
                    int downRow = upRow + 16;
                    writerInsert.PlaceNextAndWriteLine(upRow, 5, id);
                    writerInsert.PlaceNextAndWriteLine(downRow, 5, id);
                    pageInfo.pageNumberChineseStr = "<" + id + ">";
                    pageInfo.pageNumberStr = pageInfo.pageNumberChineseStr.Replace("HKBKC", "HKBK");
                    pageInfo.minNumberStr = GetMinMaxValue(ExcelUtil.GetRange(upRow, 1, worksheetInsert).Text.ToString() + "," + ExcelUtil.GetRange(upRow, 2, worksheetInsert).Text.ToString())[0];
                    if (downRow < (brokerRowsSum))
                    {
                        pageInfo.maxNumberStr = GetMinMaxValue(ExcelUtil.GetRange(downRow, 1, worksheetInsert).Text.ToString() + "," + ExcelUtil.GetRange(downRow, 2, worksheetInsert).Text.ToString())[1];
                    }
                    else
                    {
                        pageInfo.maxNumberStr = GetMinMaxValue(ExcelUtil.GetRange(brokerRowsSum, 1, worksheetInsert).Text.ToString() + ExcelUtil.GetRange(brokerRowsSum, 2, worksheetInsert).Text.ToString())[1];
                    }
                    pageList.Add(pageInfo);
                }

                writerInsert.PlaceNext(2, 7);
                int referStartRow = 0;
                int referLastRow = 0;
                while ((ExcelUtil.GetRange(writerInsert.Row + 1, 7, worksheetInsert).Text.ToString() != "") && (ExcelUtil.GetRange(writerInsert.Row + 2, 7, worksheetInsert).Text.ToString() != ""))
                {
                    if (ExcelUtil.GetRange(writerInsert.Row, 7, worksheetInsert).Text.ToString() == "REFER <HKBK01>")
                    {
                        referStartRow = writerInsert.Row;
                    }
                    writerInsert.PlaceNext(writerInsert.Row + 1, 7);
                }
                referLastRow = writerInsert.Row + 2;

                Range referRange = ExcelUtil.GetRange(referStartRow, 7, referLastRow, 15, worksheetInsert);
                Range bakRange = ExcelUtil.GetRange(referLastRow + 1, 7, referLastRow + 1 + (referLastRow - referStartRow), 15, worksheetInsert);
                referRange.Copy(bakRange);

                ExcelUtil.GetRange(2, 7, referStartRow - 1, 15, worksheetInsert).Clear();

                writerInsert.PlaceNext(1, 7);
                for (int j = 0; j < pageList.Count; j++)
                {
                    if (j % 3 == 0)
                    {
                        writerInsert.PlaceNext(writerInsert.Row + 1, 7);
                    }
                    else
                    {
                    }
                    writerInsert.WriteLine(pageList[j].minNumberStr);
                    writerInsert.WriteLine(pageList[j].maxNumberStr);
                    writerInsert.WriteLine(pageList[j].pageNumberChineseStr);
                }
                ExcelUtil.GetRange(2, 7, writerInsert.Row, 15, worksheetInsert).Font.FontStyle = ExcelUtil.GetRange(4, 1, worksheetInsert).Font.FontStyle;
                ExcelUtil.GetRange(2, 7, writerInsert.Row, 15, worksheetInsert).Font.Size = ExcelUtil.GetRange(4, 1, worksheetInsert).Font.Size;
                ExcelUtil.GetRange(2, 7, writerInsert.Row, 9, worksheetInsert).Interior.Color = 16777164.0;
                ExcelUtil.GetRange(2, 10, writerInsert.Row, 12, worksheetInsert).Interior.Color = 13409484.0;
                ExcelUtil.GetRange(2, 13, writerInsert.Row, 15, worksheetInsert).Interior.Color = 10092543.0;
                Range referR = ExcelUtil.GetRange(writerInsert.Row + 1, 7, writerInsert.Row + (referLastRow - referStartRow) + 1, 15, worksheetInsert);
                bakRange.Copy(referR);
                if ((referLastRow - referStartRow) > (writerInsert.Row - referStartRow))
                {
                    ExcelUtil.GetRange(writerInsert.Row + (referLastRow - referStartRow) + 3, 7, referLastRow + 1 + (referLastRow - referStartRow), 15, worksheetInsert).Clear();
                }
            }
        }

        //Get all the broker information from the current csv file
        public List<BrokerInfo> GetBrokerList(string currentCSVFilePath)
        {
            List<BrokerInfo> brokerList = new List<BrokerInfo>();
            using (ExcelApp app = new ExcelApp(false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, currentCSVFilePath);
                Worksheet worksheet = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                if (worksheet == null)
                {
                    LogMessage(string.Format("Cannot get worksheet {0} from workbook {1}", Path.GetFileNameWithoutExtension(currentCSVFilePath), workbook.Name));
                }
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 2, 1, ExcelLineWriter.Direction.Down))
                {
                    while (reader.Row < lastUsedRow)
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(ExcelUtil.GetRange(reader.Row, 1, worksheet).Text.ToString()))
                            {
                                string brokerNumbers = ExcelUtil.GetRange(reader.Row, 2, worksheet).Text.ToString();
                                if ((brokerNumbers != null) && (brokerNumbers != ""))
                                {
                                    string originalName = ExcelUtil.GetRange(reader.Row, 3, worksheet).Text.ToString().Trim();
                                    string ChineseName = ExcelUtil.GetRange(reader.Row, 4, worksheet).Text.ToString().Trim();
                                    if (brokerList.Count == 0 || originalName != brokerList[brokerList.Count - 1].originalName)
                                    {
                                        if (brokerList.Count != 0 && brokerNumbers == brokerList[brokerList.Count - 1].brokerNumbers)
                                        {
                                            reader.PlaceNext(reader.Row + 1, 1);
                                            continue;
                                        }
                                        Dictionary<int, BrokerInfo> brokerGroupMap = new Dictionary<int, BrokerInfo>();
                                        foreach (string id in brokerNumbers.Split(','))
                                        {
                                            int curId = int.Parse(id.Trim());
                                            int part1 = curId / 10;
                                            int part2 = curId % 10;

                                            if (brokerGroupMap.ContainsKey(part1))
                                            {
                                                brokerGroupMap[part1].IdList.Add(part2);
                                            }
                                            else
                                            {
                                                BrokerInfo brokerInfo = new BrokerInfo();
                                                brokerInfo.ChineseName = ChineseName.Trim();
                                                brokerInfo.originalName = originalName;
                                                brokerInfo.brokerNumbers = brokerNumbers;
                                                brokerInfo.GetBrokerLongAndShortName();
                                                if (brokerInfo.err != string.Empty)
                                                {
                                                    Logger.Log(brokerInfo.err);
                                                    brokerInfo.err = string.Empty;
                                                }
                                                brokerInfo.GetBrokerChineseLongAndShortName();
                                                if (brokerInfo.err != string.Empty)
                                                {
                                                    Logger.Log(brokerInfo.err);
                                                    brokerInfo.err = string.Empty;
                                                }
                                                brokerInfo.GroupId = part1;
                                                brokerInfo.IdList = new List<int>();
                                                brokerInfo.IdList.Add(part2);
                                                brokerGroupMap.Add(part1, brokerInfo);
                                            }
                                        }

                                        foreach (BrokerInfo broker in brokerGroupMap.Values)
                                        {
                                            broker.hall = broker.GetHallInfo();
                                            broker.terminal = broker.GetTerminalInfo();
                                            if (!brokerList.Contains(broker))
                                            {
                                                brokerList.Add(broker);
                                            }
                                        }
                                    }
                                }

                                reader.PlaceNext(reader.Row + 1, 1);
                            }
                            else
                            {
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMessage(string.Format("msg:{0}", ex.Message));
                        }
                    }
                }
                workbook.Close(false, workbook.FullName, false);
            }
            brokerList.Sort();
            return brokerList;
        }

        //Download csv file from page http://www.hkex.com.hk/eng/plw/plw_search.asp
        public void DownLoadSCVFile()
        {
            string uri = GetCSVFileSrc();
            Thread.Sleep(2000);

            if (File.Exists(configObj.CURRENT_SEHKFILE_PATH))
            {
                MiscUtil.BackUpFile(configObj.CURRENT_SEHKFILE_PATH);
                File.Delete(configObj.CURRENT_SEHKFILE_PATH);
            }
            WebClientUtil.DownloadFile(uri, 180000, configObj.CURRENT_SEHKFILE_PATH);
        }

        //Get the target csv file source address
        public string GetCSVFileSrc()
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = null;
            htmlDoc = WebClientUtil.GetHtmlDocument(BASE_URI, 180000);
            string csvFileSrc = htmlDoc.DocumentNode.SelectSingleNode("//table/tbody/tr/td/a/img[@name='btn_download']").ParentNode.Attributes["href"].Value;
            csvFileSrc = csvFileSrc.Remove(0, csvFileSrc.IndexOf("eng", 0) + 3);
            string baseUri = BASE_URI.Remove(BASE_URI.IndexOf("/plw", 0));
            csvFileSrc = baseUri + csvFileSrc;
            return csvFileSrc;
        }

        //Get the min and max value, minValue: minMaxValue[0]; maxValue: minMaxValue[1]
        public string[] GetMinMaxValue(string value)
        {
            string[] minMaxValue = { "", "" };
            Regex regex = new Regex("(?<brokerNum>\\d{4})");
            MatchCollection mc = regex.Matches(value);
            minMaxValue[0] = mc[0].Groups["brokerNum"].Value;
            if (Char.IsDigit(value[value.Length - 1]))
            {
                minMaxValue[1] = minMaxValue[0].Remove(3, 1) + value[value.Length - 1];
            }
            else
            {
                if ((Char.IsDigit(value[value.Length - 2])) && (value[value.Length - 1] == '.'))
                {
                    minMaxValue[1] = minMaxValue[0].Remove(3, 1) + value[value.Length - 2];
                }
                else
                {
                    minMaxValue[1] = minMaxValue[0].Remove(3, 1) + value[value.Length - 3];
                }
            }
            return minMaxValue;
        }
    }
}
