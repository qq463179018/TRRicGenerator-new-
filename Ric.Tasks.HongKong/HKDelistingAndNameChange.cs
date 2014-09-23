using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.ComponentModel;
using System.Drawing.Design;
using System.Text.RegularExpressions;
using Ric.Util;
using Ric.Core;

namespace Ric.Tasks.HongKong
{
    [ConfigStoredInDB]
    public class HKDelistingAndNameChangConfig
    {
        [StoreInDB]
        [Category("Source File")]
        [Description("Search from the start  position to find the delisting and name change rows.")]
        public string DELISTING_START_POSITION { get; set; }

        [StoreInDB]
        [Category("Source File")]
        [Description("The path of the source file.")]
        public string SOURCE_FILE_PATH { get; set; }

        [StoreInDB]
        [Category("Source File")]
        public string SOURCE_FILE_WORKSHEET_NAME { get; set; }

        [StoreInDB]
        [Category("Save Path")]
        [Description("The save path of the generated cbbc delisting file.")]
        public string CBBC_DELISTING_SAVE_PATH { get; set; }

        [StoreInDB]
        [Category("Save Path")]
        [Description("The save path of the generated warrant delisting file.")]
        public string WARRANT_DELISTING_SAVE_PATH { get; set; }

        [StoreInDB]
        [Category("Save Path")]
        [Description("The save path of the generated name change file.")]
        public string NAMECHANGE_SAVE_PATH { get; set; }
    }

    public class HKDelistingAndNameChangSectorConfig
    {
        public List<SectorConfig> SECTORCONFIG { get; set; }

        public List<WarrantDelistingConfig> WARRANTDELISTINGCONFIG { get; set; }

        public List<CBBCDelistingConfig> CBBCDELISTINGCONFIG { get; set; }

        public List<WarrantStartWith> WARRANTCHAINRICSTARTWITH { get; set; }

        public string OthersWarrantChainRic { get; set; }

        public string OthersCBBCChainRic { get; set; }

        public List<string> HYPHEN { get; set; }

    }

    //public class SectorConfig
    //{
    //    public string SECTOR { get; set; }
    //    public string SECTORCHAIN { get; set; }
    //}

    //public class WarrantDelistingConfig
    //{
    //    public string UnderlyingChain { get; set; }
    //    public string ChainRic { get; set; }
    //}

    //public class CBBCDelistingConfig
    //{
    //    public string UnderlyingChain { get; set; }
    //    public string ChainRic { get; set; }
    //}

    //public class WarrantStartWith
    //{
    //    public string StartWith { get; set; }
    //    public string ChainRic { get; set; }
    //}

    public class HKDelistingAndNameChange : GeneratorBase
    {
        private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_DelistingAndNameChange.config";
        private static readonly string SECTOR_CONFIG_PATH = ".\\Config\\HK\\HK_DelistingAndNameChange.xml";
        private static HKDelistingAndNameChangConfig configObj = null;
        private static HKDelistingAndNameChangSectorConfig xmlConfigObj = null;

        protected override void Start()
        {
            StartDelistingAndNameChangeJob();
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as HKDelistingAndNameChangConfig;
            xmlConfigObj = ConfigUtil.ReadConfig(SECTOR_CONFIG_PATH, typeof(HKDelistingAndNameChangSectorConfig)) as HKDelistingAndNameChangSectorConfig;
        }

        private string fileName = "ggg";

        private char[] hyphen = null; //{ '-', '(', ' ' };
        private string nameChangeCode = "";
        public string NameChangeCode
        {
            get
            {
                if (nameChangeCode.Length < 4)
                {
                    for (int i = 0; i < (4 - nameChangeCode.Length); i++)
                    {
                        nameChangeCode = "0" + nameChangeCode;
                    }
                }
                return nameChangeCode;
            }
            set
            {
                this.nameChangeCode = value;
            }
        }

        private string oldShortName = "";
        private string oldLongName = "";
        private string lotSize = "";
        private string sector = "";
        private string underlyingStockCode = "";
        private string chainRic = "";
        private List<List<string>> cbbcList = new List<List<string>>();
        private List<List<string>> warrantList = new List<List<string>>();
        private List<List<string>> nameChangeList = new List<List<string>>();
        private List<List<List<string>>> apartCbbcList = new List<List<List<string>>>();
        private List<List<List<string>>> apartWarrantList = new List<List<List<string>>>();

        public void FindSourceFile(string sourceFolderPath)
        {
            DateTime now = DateTime.Now;
            string formatDate = now.ToString("yy-MM-dd", new CultureInfo("en-US"));
            List<string> fileNames = new List<string>();

            fileNames = FileUtil.GetTodayFileNameFromFolder(sourceFolderPath);
            foreach (string name in fileNames)
            {
                if (name.Contains(formatDate))
                {
                    fileName = name;
                    break;
                }
            }
        }

        private void InitialHyphen()
        {
            hyphen = new char[xmlConfigObj.HYPHEN.Count];
            int count = 0;
            foreach (string item in xmlConfigObj.HYPHEN)
            {
                hyphen[count] = (item.ToCharArray())[0];
                count++;
            }
        }

        public void StartDelistingAndNameChangeJob()
        {
            cbbcList.Clear();
            warrantList.Clear();
            nameChangeList.Clear();
            apartCbbcList.Clear();
            apartWarrantList.Clear();
            InitialHyphen();
            int StartPosition = int.Parse(configObj.DELISTING_START_POSITION);
            string sourceFolderPath = configObj.SOURCE_FILE_PATH; //Z:\Hong Kong\HK Circular
            //Core coreObj = new Core();
            //FindSourceFile(sourceFolderPath);

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (File.Exists(configObj.SOURCE_FILE_PATH))
            {
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.SOURCE_FILE_PATH);
                    var worksheet = ExcelUtil.GetWorksheet(configObj.SOURCE_FILE_WORKSHEET_NAME, workbook);
                    if (worksheet == null)
                    {
                        LogMessage(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.SOURCE_FILE_WORKSHEET_NAME, configObj.SOURCE_FILE_PATH));
                    }

                    Range rangeCol6 = ExcelUtil.GetRange(StartPosition, 6, worksheet);
                    Range rangeCol3 = ExcelUtil.GetRange(StartPosition, 3, worksheet);
                    if (rangeCol6.get_Value(Missing.Value) != null && rangeCol3.get_Value(Missing.Value) != null)
                    {
                        int usedRange = worksheet.UsedRange.Count;
                        GetDelistingAndNameChangeData(usedRange, worksheet, StartPosition);
                        ApartList(apartCbbcList, cbbcList);
                        ApartList(apartWarrantList, warrantList);
                    }
                    else
                    {
                        LogMessage(" Check today's file in the folder Delisting!");
                        //coreObj.KillExcelProcess(xlApp);
                    }
                    workbook.Close(false, configObj.SOURCE_FILE_PATH, true);
                    WriteDataIntoFile(xlApp);
                }
                //coreObj.KillExcelProcess(xlApp);
            }
            else
            {
                LogMessage(string.Format(" Can't find today's delisting file {0}", configObj.SOURCE_FILE_PATH));
            }

        }

        public void GetDelistingAndNameChangeData(int usedRange, Worksheet delSheet, int StartPosition)
        {
            //for (int rowIndex = StartPosition; rowIndex < usedRange; rowIndex = rowIndex + 3)
            for (int rowIndex = StartPosition; rowIndex < usedRange; rowIndex++)
            {
                Range rangeCol6 = ExcelUtil.GetRange(rowIndex, 6, delSheet);
                Range rangeCol3 = ExcelUtil.GetRange(rowIndex, 3, delSheet);

                if (rangeCol6.get_Value(Missing.Value) == null || rangeCol3.get_Value(Missing.Value) == null)
                    continue;

                string cel6Value = rangeCol6.get_Value(Missing.Value).ToString();
                string cel3Value = rangeCol3.get_Value(Missing.Value).ToString();

                List<string> itemList = null;
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

                //name change
                if (cel6Value.Equals("Change of Name"))
                {
                    itemList = ReadNameChangeData(delSheet, rowIndex);
                    nameChangeList.Add(itemList);
                }
            }
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
            List<List<string>> newCbbcList = new List<List<string>>();
            newCbbcList.Add(cbbcList[0]);
            string date = cbbcList[0][2].ToString();
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

        public void WriteDataIntoFile(Microsoft.Office.Interop.Excel.Application xlApp)
        {
            if (apartCbbcList.Count > 0)
            {
                foreach (List<List<string>> cbbcItem in apartCbbcList)
                {
                    WriteCBBCDelistingFMFile(xlApp, cbbcItem);
                }
            }
            if (apartWarrantList.Count > 0)
            {
                foreach (List<List<string>> warrantItem in apartWarrantList)
                {
                    WriteWarrantDelistingFMFile(xlApp, warrantItem);
                }
            }
            if (nameChangeList.Count > 0)
            {
                WriteNameChangeFile(xlApp, nameChangeList);
            }
        }

        public List<string> ReadDelistingRowData(Worksheet wSheet, int rowIndex)
        {
            List<string> itemList = new List<string>();
            string col3Value = ExcelUtil.GetRange(rowIndex, 3, wSheet).get_Value(Missing.Value).ToString();
            itemList.Add(col3Value);
            Range range = ExcelUtil.GetRange(rowIndex, 4, wSheet);
            if (range.get_Value(Missing.Value) != null)
            {
                itemList.Add(range.get_Value(Missing.Value).ToString());
            }
            else
            {
                LogMessage(string.Format("Check the source, no stock name in the row {0}", rowIndex));
            }
            range = ExcelUtil.GetRange(rowIndex, 9, wSheet);
            if (range.get_Value(Missing.Value) != null)
            {
                itemList.Add(range.get_Value(Missing.Value).ToString());
            }
            else
            {
                LogMessage(string.Format("Check the source, no date in the row {0}", rowIndex));
            }
            return itemList;
        }

        private void GetUnderlyingStockCode(HtmlAgilityPack.HtmlDocument htmlDoc)
        {
            try
            {
                underlyingStockCode = htmlDoc.DocumentNode.SelectSingleNode("//table/tr[10]/td[2]").InnerText.Replace("&nbsp;", "").Trim();
            }
            catch (Exception e)
            {
                LogMessage("Can't find underlying Stock code:" + e);
            }
            GetWarrantChainRic(underlyingStockCode);
        }

        private void GetWarrantChainRic(string underlyingStockCode)
        {
            chainRic = "";
            foreach (WarrantStartWith item in xmlConfigObj.WARRANTCHAINRICSTARTWITH)
            {
                if (underlyingStockCode.StartsWith(item.StartWith))
                {
                    chainRic = item.ChainRic;
                }
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
                foreach (WarrantDelistingConfig warrantItem in xmlConfigObj.WARRANTDELISTINGCONFIG)
                {
                    if (underlyingStockCode.Equals(warrantItem.UnderlyingChain))
                    {
                        chainRic = warrantItem.ChainRic;
                    }
                }
            }

            if (chainRic.Equals(string.Empty))
            {
                chainRic = xmlConfigObj.OthersWarrantChainRic;
            }

        }

        private void WriteWarrantDelistingFMFile(Microsoft.Office.Interop.Excel.Application xlApp, List<List<string>> dataList)
        {
            //string delistingSavePath = @"D:\zhang fan\Delisting";
            //folderPath = @"Z:\Hong Kong\FM\Today";
            string delistingSavePath = configObj.WARRANT_DELISTING_SAVE_PATH;
            DateTime date = TimeUtil.ParseTime(dataList[0][2].ToString());
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
                fileName = "HK" + TimeUtil.shortYear + "-_DELETE";
                Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                wSheet.Application.Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
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
                        code = dataList[j][0].ToString();
                        string name = dataList[j][1].ToString();
                        GetInfoFromInternet(code, "Warrant");

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
                if (System.IO.File.Exists(delistingSavePath + "\\" + fileName))
                {
                    System.IO.File.Delete(delistingSavePath + "\\" + fileName);
                }
                wBook.SaveAs(delistingSavePath + "\\" + fileName, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                wBook.Close(Missing.Value, Missing.Value, Missing.Value);
            }
        }

        public void WriteCBBCDelistingFMFile(Microsoft.Office.Interop.Excel.Application xlApp, List<List<string>> dataList)
        {
            //string delistingSavePath = @"D:\zhang fan\Delisting";
            //folderPath = @"Z:\Hong Kong\FM\Today";
            string delistingSavePath = configObj.CBBC_DELISTING_SAVE_PATH;
            DateTime date = TimeUtil.ParseTime(dataList[0][2].ToString());
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
                fileName = "HK" + TimeUtil.shortYear + "-_DELETE";
                Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                wSheet.Application.Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
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
                        code = dataList[j][0].ToString();
                        string name = dataList[j][1].ToString();

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
                if (System.IO.File.Exists(delistingSavePath + "\\" + fileName))
                {
                    System.IO.File.Delete(delistingSavePath + "\\" + fileName);
                }
                wBook.SaveAs(delistingSavePath + "\\" + fileName, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                wBook.Close(Missing.Value, Missing.Value, Missing.Value);
            }

        }


        private string ChangeNameIntoCamel(string name)
        {
            string[] nameParts = name.Split(' ');
            string camelName = "";
            for (int i = 0; i < nameParts.Length; i++)
            {
                string formatName = nameParts[i].Substring(0, 1) + nameParts[i].Substring(1, nameParts[i].Length - 1).ToLower();
                camelName = camelName + " " + formatName;

            }
            return camelName.Trim();
        }

        private string GetChainRicFromSector(string sector)
        {
            string chainRic = "";
            if (!sector.Equals(string.Empty))
            {
                chainRic = "0#.HK" + sector.Substring(0, 2).ToUpper() + "f";
            }
            return chainRic;

        }

        private char GetClassifiedCodeFromName(string name)
        {
            char classifiedChar = ' ';
            char[] chars = name.Substring(0, 1).ToCharArray();
            char firstChar = chars[0];
            if (firstChar >= 'A' && firstChar < 'H')
            {
                classifiedChar = 'A';
            }
            else if (firstChar >= 'H' && firstChar < 'P')
            {
                classifiedChar = 'H';
            }
            else
            {
                classifiedChar = 'P';
            }
            return classifiedChar;
        }

        private string GetSectorChainFromSector(string sector, string shortName)
        {
            string sectorChain = "";
            foreach (SectorConfig item in xmlConfigObj.SECTORCONFIG)
            {
                if (item.Sector.Equals(sector))
                {
                    sectorChain = item.SectorChain;
                }
                if (sector.Equals("Industrial") || sector.Equals("Consolidated"))
                {
                    sectorChain = sectorChain + GetClassifiedCodeFromName(shortName).ToString() + ".HK";
                }
            }
            if (sectorChain.Equals(string.Empty))
            {
                LogMessage("Check the SECTORCONFIG in xml file!");
            }
            return sectorChain;
        }

        public List<string> ReadNameChangeData(Worksheet wSheet, int rowIndex)
        {

            List<string> itemList = new List<string>();
            string col1Value = ExcelUtil.GetRange(rowIndex, 2, wSheet).get_Value(Missing.Value).ToString();

            itemList.Add(col1Value);
            Range range = ExcelUtil.GetRange(rowIndex, 3, wSheet);
            if (range.get_Value(Missing.Value) != null)
            {
                string shortName = range.get_Value(Missing.Value).ToString();
                itemList.Add(shortName);
            }
            else
            {
                LogMessage(string.Format("Check the source, no short name in the row {0}", rowIndex));
            }
            range = ExcelUtil.GetRange(rowIndex + 1, 3, wSheet);
            if (range.get_Value(Missing.Value) != null)
            {
                string longName = range.get_Value(Missing.Value).ToString();
                string name = ChangeNameIntoCamel(longName);
                itemList.Add(name);
            }
            else
            {
                LogMessage(string.Format("Check the source, no long name in the row {0}", rowIndex));
            }
            range = ExcelUtil.GetRange(rowIndex, 8, wSheet);
            if (range.get_Value(Missing.Value) != null)
            {
                itemList.Add(range.get_Value(Missing.Value).ToString());
            }
            else
            {
                LogMessage(string.Format("Check the source, no date in the row {0}", rowIndex));
            }
            range = ExcelUtil.GetRange(rowIndex + 1, 8, wSheet);
            if (range.get_Value(Missing.Value) != null)
            {
                string chineseName = range.get_Value(Missing.Value).ToString();
                itemList.Add(chineseName);
                itemList.Add(FormatName(chineseName));
            }
            else
            {
                LogMessage(string.Format("Check the source, no chinese name in the row {0}", rowIndex));
            }
            return itemList;
        }

        #region Remove the fake name change from list
        private void RemoveInvalidNameChangeRow(List<List<string>> nameList, int start)
        {
            if (nameList.Count == 0 || nameList[0].Count == 0 || (start >= nameList.Count))
                return;

            int lastDataIndex = nameList[start].Count - 1;
            string firstName = nameList[start][lastDataIndex];
            List<int> deleteRow = new List<int>();
            for (int i = start + 1; i < nameList.Count; i++)
            {
                string currentName = nameList[i][lastDataIndex];
                if (firstName.Equals(currentName))
                {
                    deleteRow.Add(i);
                }
            }
            if (deleteRow.Count > 0)
            {
                deleteRow.Add(start);
                DeleteNameListRows(nameList, deleteRow, start);
            }
            else
            {
                start++;
                RemoveInvalidNameChangeRow(nameList, start);
            }
        }

        private void DeleteNameListRows(List<List<string>> nameList, List<int> deleteRow, int start)
        {
            foreach (int rowIndex in deleteRow)
            {
                nameList.Remove(nameList[rowIndex]);
            }
            RemoveInvalidNameChangeRow(nameList, start);
        }
        #endregion

        private string FormatName(string name)
        {
            name = name.TrimStart();
            int hyphenIndex = name.IndexOfAny(hyphen);
            if (hyphenIndex > 0)
                name = name.Substring(0, hyphenIndex);
            return name.TrimEnd();
        }

        private void GetNameChangeInfo(HtmlAgilityPack.HtmlDocument htmlDoc)
        {
            try
            {
                oldShortName = htmlDoc.DocumentNode.SelectSingleNode("//table/tr/td").InnerText.Replace("&nbsp;", "").Trim();
                int codePos = oldShortName.IndexOf("(");
                oldShortName = oldShortName.Substring(0, codePos).Trim();
                HtmlAgilityPack.HtmlNode node = htmlDoc.DocumentNode.SelectSingleNode("//table");

                oldLongName = htmlDoc.DocumentNode.SelectSingleNode("//table/tr[2]/td[2]").InnerText.Replace("&nbsp;", "").Trim().TrimEnd('.');
                oldLongName = ChangeNameIntoCamel(oldLongName);
                lotSize = htmlDoc.DocumentNode.SelectSingleNode("//table/tr[15]/td[2]").InnerText.Replace("&nbsp;", "").Trim();

                sector = htmlDoc.DocumentNode.SelectSingleNode("//table/tr[8]/td[2]").InnerText.Replace("&nbsp;", "").Trim();
                sector = sector.Substring(0, sector.IndexOf(' '));
            }
            catch (Exception e)
            {
                LogMessage("Grab information from web sit error!:" + e);
            }
        }

        public void GetInfoFromInternet(string code, string condition)
        {
            string url = "http://www.hkex.com.hk/eng/invest/company/profile_page_e.asp?WidCoID=" + code + "&WidCoAbbName=&Month=&langcode=e";
            try
            {
                WebRequest wr = WebRequest.Create(url);
                WebResponse rs = wr.GetResponse();
                StreamReader sr = new StreamReader(rs.GetResponseStream());
                string htm = sr.ReadToEnd();

                int pos = htm.IndexOf("Company/Securities Name:");
                string beforeHtm = htm.Substring(0, pos);
                string afterHtm = htm.Substring(pos, htm.Length - pos);

                int beforeTablePos = beforeHtm.LastIndexOf("<table");
                int afterTablePos = afterHtm.IndexOf("</table>");
                string beforeTable = beforeHtm.Substring(beforeTablePos, beforeHtm.Length - beforeTablePos);
                string afterTable = afterHtm.Substring(0, afterTablePos) + "</table>";
                string simpleHtml = "<html><head></head><body>" + (beforeTable + afterTable).Replace("\r", "").Replace("\n", "").Replace("\t", "") + "</body></html>";

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
            }
            catch (Exception e)
            {
                LogMessage("Get HTML raw data error!:" + e);
            }

        }

        public void WriteNameChangeFile(Microsoft.Office.Interop.Excel.Application xlApp, List<List<string>> nameList)
        {
            string nameChangSavePath = configObj.NAMECHANGE_SAVE_PATH;

            RemoveInvalidNameChangeRow(nameList, 0);

            for (int i = 0; i < nameList.Count; i++)
            {
                GetInfoFromInternet(nameList[i][0], "NameChange");

                fileName = "HK" + TimeUtil.shortYear + "-2996_CHANGE";
                DateTime date = TimeUtil.ParseTime(nameList[i][3].ToString());
                string formatDate = TimeUtil.GetFormatDate(date);
                string effectiveDate = TimeUtil.GetEffectiveDate(date);

                Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                wSheet.Application.Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                wSheet.Cells[1, 1] = "Please action the change of the following HK stock on TQS.";
                wSheet.Cells[3, 1] = "FM Serial Number:";
                wSheet.Cells[3, 2] = "HK" + TimeUtil.shortYear;
                wSheet.Cells[4, 1] = "Effective Date:";
                wSheet.Cells[4, 2] = effectiveDate;

                nameChangeCode = nameList[i][0];
                string newShortName = nameList[i][1].ToString();
                string newLongName = ChangeNameIntoCamel(nameList[i][2].ToString());
                int rowIndex = 6;
                wSheet.Cells[rowIndex++, 1] = "+AMENDMENT+";
                wSheet.Cells[rowIndex++, 1] = "'---------------------------------------------------------------------------------------------------";
                Range range = ExcelUtil.GetRange(rowIndex, 1, wSheet);
                range.NumberFormatLocal = "(0)";
                wSheet.Cells[rowIndex++, 1] = "1";
                wSheet.Cells[rowIndex++, 1] = "**For TQS**";
                rowIndex++;
                wSheet.Cells[rowIndex, 1] = "Underlying RIC:";
                wSheet.Cells[rowIndex++, 2] = NameChangeCode + ".HK";
                wSheet.Cells[rowIndex, 1] = "Composite chain RIC:";
                wSheet.Cells[rowIndex++, 2] = "0#" + NameChangeCode + ".HK";
                wSheet.Cells[rowIndex, 1] = "Broker page RIC:";
                wSheet.Cells[rowIndex++, 2] = NameChangeCode + "bk.HK";
                wSheet.Cells[rowIndex, 1] = "Misc.Info page RIC:";
                wSheet.Cells[rowIndex++, 2] = NameChangeCode + "MI.HK";
                wSheet.Cells[rowIndex, 1] = "Displayname (OLD):";
                wSheet.Cells[rowIndex++, 2] = oldShortName;
                wSheet.Cells[rowIndex, 1] = "Displayname (NEW):";
                wSheet.Cells[rowIndex++, 2] = newShortName;
                wSheet.Cells[rowIndex, 1] = "Official Code:";
                range = ExcelUtil.GetRange(rowIndex, 2, wSheet);
                range.NumberFormatLocal = "@";
                wSheet.Cells[rowIndex++, 2] = NameChangeCode;
                wSheet.Cells[rowIndex, 1] = "Exchange Symbol:";
                range = ExcelUtil.GetRange(rowIndex, 2, wSheet);
                range.NumberFormatLocal = "@";
                wSheet.Cells[rowIndex++, 2] = NameChangeCode;
                wSheet.Cells[rowIndex, 1] = "Currency:";
                wSheet.Cells[rowIndex++, 2] = "HKD";
                wSheet.Cells[rowIndex, 1] = "Recordtype:";
                wSheet.Cells[rowIndex++, 2] = "113";
                wSheet.Cells[rowIndex, 1] = "Spare_Ubytes8:";
                wSheet.Cells[rowIndex++, 2] = "EQTY";
                wSheet.Cells[rowIndex, 1] = "IDN Longname (OLD):";
                wSheet.Cells[rowIndex++, 2] = oldShortName + "@STK";
                wSheet.Cells[rowIndex, 1] = "IDN Longname (NEW):";
                wSheet.Cells[rowIndex++, 2] = newShortName + "@STK";
                wSheet.Cells[rowIndex, 1] = "Chain RIC (OLD):";
                wSheet.Cells[rowIndex++, 2] = "0#" + oldShortName.Substring(0, 1).ToString() + ".HK";
                wSheet.Cells[rowIndex, 1] = "Chain RIC (NEW):";
                wSheet.Cells[rowIndex++, 2] = "0#" + newShortName.Substring(0, 1).ToString() + ".HK";
                wSheet.Cells[rowIndex, 1] = "Chain RIC (OLD):";
                wSheet.Cells[rowIndex++, 2] = GetSectorChainFromSector(sector, oldShortName);
                wSheet.Cells[rowIndex, 1] = "Chain RIC (NEW):";
                wSheet.Cells[rowIndex++, 2] = GetSectorChainFromSector(sector, oldShortName);
                wSheet.Cells[rowIndex, 1] = "Misc.Info page Chain RIC:";
                wSheet.Cells[rowIndex++, 2] = "0#MI.HK";


                wSheet.Cells[rowIndex, 1] = "Lot Size:";
                range = ExcelUtil.GetRange(rowIndex, 2, wSheet);
                range.NumberFormatLocal = "#,##0";
                wSheet.Cells[rowIndex++, 2] = lotSize;
                wSheet.Cells[rowIndex, 1] = "COI DSPLY_NMLL (OLD):";
                wSheet.Cells[rowIndex++, 2] = "";
                wSheet.Cells[rowIndex, 1] = "COI DSPLY_NMLL (NEW):";
                wSheet.Cells[rowIndex++, 2] = nameList[i][4].ToString();

                wSheet.Cells[rowIndex++, 1] = "---------------------------------------------------------------------------------------------------";
                wSheet.Cells[rowIndex++, 1] = "LONGLINK1 (Warrant chain RIC):";
                wSheet.Cells[rowIndex, 1] = "BCAST_REF:";
                wSheet.Cells[rowIndex++, 2] = NameChangeCode + ".HK";
                wSheet.Cells[rowIndex++, 1] = "---------------------------------------------------------------------------------------------------";


                wSheet.Cells[rowIndex, 1] = "Organisation Name (DIRNAME) (OLD):";
                wSheet.Cells[rowIndex++, 2] = oldLongName;
                wSheet.Cells[rowIndex, 1] = "Organisation Name (DIRNAME) (NEW):";
                wSheet.Cells[rowIndex++, 2] = newLongName;
                wSheet.Cells[rowIndex, 1] = "Legal Registered Name (OLD):";
                wSheet.Cells[rowIndex++, 2] = oldLongName.Replace("Ltd", "Limited");
                wSheet.Cells[rowIndex, 1] = "Legal Registered Name (NEW):";
                wSheet.Cells[rowIndex++, 2] = newLongName.Replace("Ltd", "Limited");


                wSheet.Cells[rowIndex, 1] = "Old Organisation Name as Alias - Previous Name:";
                wSheet.Cells[rowIndex++, 2] = "Yes";
                wSheet.Cells[rowIndex, 1] = "Issue Classification:";
                wSheet.Cells[rowIndex++, 2] = "CS";
                wSheet.Cells[rowIndex, 1] = "Local Sector Classification Name:";
                wSheet.Cells[rowIndex++, 2] = sector;
                wSheet.Cells[rowIndex++, 1] = "---------------------------------------------------------------------------------------------------";
                wSheet.Cells[rowIndex++, 1] = "GN_TX20_14";
                wSheet.Cells[rowIndex, 1] = "-Index chain RIC, pos 1:";
                wSheet.Cells[rowIndex++, 2] = "********";
                wSheet.Cells[rowIndex, 1] = "-Sector chain RIC, pos 9 (OLD):";
                wSheet.Cells[rowIndex++, 2] = GetSectorChainFromSector(sector, oldShortName);
                wSheet.Cells[rowIndex, 1] = "-Sector chain RIC, pos 9 (NEW):";
                wSheet.Cells[rowIndex++, 2] = GetSectorChainFromSector(sector, oldShortName);
                wSheet.Cells[rowIndex++, 1] = "---------------------------------------------------------------------------------------------------";
                wSheet.Cells[rowIndex++, 1] = "**FOR AFE**";
                wSheet.Cells[rowIndex++, 1] = "GV1_FLAG:";
                wSheet.Cells[rowIndex++, 1] = "ISS_TP_FLG:";
                wSheet.Cells[rowIndex++, 1] = "RDM_CUR:";
                wSheet.Cells[rowIndex++, 1] = "LONGLINK14:";
                wSheet.Cells[rowIndex, 1] = "BOND_TYPE:";
                wSheet.Cells[rowIndex++, 2] = "BOND_TYPE_186";
                wSheet.Cells[rowIndex++, 1] = "LEG1_STR:";
                wSheet.Cells[rowIndex++, 1] = "LEG2_STR:";
                wSheet.Cells[rowIndex++, 1] = "GN_TXT24_1:";
                wSheet.Cells[rowIndex++, 1] = "GN_TXT24_2:";
                wSheet.Cells[rowIndex, 1] = "Chain RIC:";
                wSheet.Cells[rowIndex++, 2] = GetChainRicFromSector(sector);
                wSheet.Cells[rowIndex, 1] = "Position in chain:";
                wSheet.Cells[rowIndex++, 2] = "(in numerical order)";
                wSheet.Cells[rowIndex++, 1] = "---------------------------------------------------------------------------------------------------";

                fileName += "_" + NameChangeCode + "_" + formatDate + ".xls";
                wBook.Application.DisplayAlerts = false;
                if (System.IO.File.Exists(nameChangSavePath + "\\" + fileName))
                    System.IO.File.Delete(nameChangSavePath + "\\" + fileName);
                wBook.SaveAs(nameChangSavePath + "\\" + fileName, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                wBook.Close(Missing.Value, Missing.Value, Missing.Value);

            }

        }

    }
}
