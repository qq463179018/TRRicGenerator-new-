using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    public class ExcelOperation
    {
        public static Workbook OpenWorkBook(Application xlApp, string folderName, string fileName)
        {
            Workbook wBook = null;
            if (FileUtil.IsFolderExist(folderName))
            {
                if (!fileName.Equals(string.Empty))
                {
                    wBook = xlApp.Workbooks.Open(folderName + "\\" + fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
            }
            return wBook;
        }

        public void InsertBlankRows(Worksheet wSheet, int rowIndex, int count)
        {
            Range range = ((Range)wSheet.Cells[rowIndex, 1]).EntireRow;
            for (int i = 0; i < count; i++)
            {
                range.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                range.ApplyOutlineStyles();
            }
        }
    }

    class WarrantDataItem
    {
        #region Warrant Items
        private List<string> warrantRic = new List<string>();
        public List<string> WarrantRic
        {
            get
            {
                return warrantRic;
            }
        }
        private List<string> issuer = new List<string>();
        public List<string> Issuer
        {
            get
            {
                return issuer;
            }
        }
        private List<string> issueSize = new List<string>();
        public List<string> IssueSize
        {
            get
            {
                return issueSize;
            }
        }
        private List<string> issuePrice = new List<string>();
        public List<string> IssuePrice
        {
            get
            {
                return issuePrice;
            }
        }
        private List<string> underlyingRic = new List<string>();
        public List<string> UnderlyingRic
        {
            get
            {
                return underlyingRic;
            }
        }
        private List<string> convRatio = new List<string>();
        public List<string> ConvRatio
        {
            get
            {
                return convRatio;
            }
        }
        private List<string> strike = new List<string>();
        public List<string> Strike
        {
            get
            {
                return strike;
            }
        }
        #endregion

        public void GetWarrantDataFromWorkBook(Workbook wBook, HKSpeedGuideConfig configObj, HKSpeedGuideXML xmlObj, Logger logger)
        {
            if (wBook != null)
            {
                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                GetWarrantData(wSheet, xmlObj, logger);
            }

        }

        public string GetOrgNameMappingResult(string source, HKSpeedGuideXML xmlObj, Logger logger)
        {
            string mappingResult = "";
            foreach (NameMapping item in xmlObj.NameConfig.Where(item => item.OrganisationName.Equals(source)))
            {
                mappingResult = item.mappingName;
            }
            if (mappingResult.Equals(string.Empty))
            {
                logger.Log(string.Format("Check the xml file, can't find the mapping name according to the source name {0}", source));
            }
            return mappingResult;
        }

        private void GetWarrantData(_Worksheet wSheet, HKSpeedGuideXML xmlObj, Logger logger)
        {
            int rowIndex = 6;
            string cellValue = wSheet.Range[wSheet.Cells[rowIndex, 1], wSheet.Cells[rowIndex, 1]].Value2.ToString();
            while (cellValue.Equals("+ADDITION+"))
            {
                rowIndex++;
                cellValue = wSheet.Range[wSheet.Cells[rowIndex, 1], wSheet.Cells[rowIndex, 1]].Value2.ToString();
                while (!cellValue.Equals("+ADDITION+"))
                {
                    if (cellValue == "" && wSheet.Range[wSheet.Cells[rowIndex + 1, 1], wSheet.Cells[rowIndex + 1, 1]].Value2 == null && wSheet.Range[wSheet.Cells[rowIndex + 2, 1], wSheet.Cells[rowIndex + 2, 1]].Value2 == null)
                    {
                        break;
                    }
                    string col2Value = "";
                    switch (cellValue)
                    {
                        case "Underlying RIC:":
                            col2Value = wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2 == null ? "" : wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2.ToString();
                            warrantRic.Add(col2Value);
                            break;
                        case "Organisation Name (DIRNAME):":
                            string formatValue = "";
                            if (wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2 == null)
                            {
                                col2Value = "";
                            }
                            else
                            {
                                col2Value = wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2.ToString();
                                formatValue = GetOrgNameMappingResult(col2Value, xmlObj, logger);
                            }
                            if (!formatValue.Equals(("")))
                            {
                                issuer.Add(formatValue);
                            }
                            break;
                        case "GN_TX20_6:":
                            col2Value = wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2 == null ? "" : wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2.ToString();
                            underlyingRic.Add(col2Value);
                            break;
                        case "ISSUE PRICE:":
                            col2Value = wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2 == null ? "" : wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2.ToString();
                            issuePrice.Add(col2Value);
                            break;
                        case "WNT_RATIO:":
                            col2Value = wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2 == null ? "" : wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2.ToString();
                            convRatio.Add(col2Value);
                            break;
                        case "STRIKE_PRC:":
                            col2Value = wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2 == null ? "" : wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2.ToString();
                            strike.Add(col2Value);
                            break;
                        case "GN_TX20_11:":
                            if (wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2 == null)
                            {
                                col2Value = "";
                            }
                            else
                            {
                                col2Value = wSheet.Range[wSheet.Cells[rowIndex, 2], wSheet.Cells[rowIndex, 2]].Value2.ToString();
                                col2Value = col2Value.Replace("M", "");
                            }
                            issueSize.Add(col2Value);
                            break;
                    }
                    rowIndex++;
                    cellValue = wSheet.Range[wSheet.Cells[rowIndex, 1], wSheet.Cells[rowIndex, 1]].Value2 == null ? "" : wSheet.Range[wSheet.Cells[rowIndex, 1], wSheet.Cells[rowIndex, 1]].Value2.ToString();

                }
                int t = rowIndex;
            }
        }
    }

    class WarrantDataWriter
    {
        private WarrantDataItem _warrantDataItem;
        ExcelOperation excelOperation = new ExcelOperation();
        private Worksheet masterSheet;
        public Worksheet MasterSheet
        {
            get
            {
                return masterSheet;
            }
        }

        public void WriteIntoMasterSheet(Workbook warGuideBook, WarrantDataItem warrantDataItem, Logger logger)
        {
            _warrantDataItem = warrantDataItem;
            if (_warrantDataItem.WarrantRic != null && _warrantDataItem.WarrantRic.Count > 0)
            {
                WriteWarrantDataIntoWorkBook(warGuideBook);
            }
            else
            {
                logger.Log("No warrant data today");
            }
        }

        private void WriteWarrantDataIntoWorkBook(Workbook warGuideBook)
        {

            if (warGuideBook != null)
            {
                masterSheet = ExcelUtil.GetWorksheet("MASTER", warGuideBook);
                if (masterSheet != null)
                {
                    excelOperation.InsertBlankRows(masterSheet, 4, _warrantDataItem.WarrantRic.Count);
                    WriteWarrantDataIntoSheet(masterSheet, 4);
                }
            }
        }

        private void WriteWarrantDataIntoSheet(Worksheet wSheet, int rowIndex)
        {
            string formatToday = DateTime.Now.ToString("dd-MMM-yy", new CultureInfo("en-US"));
            for (int i = 0; i < _warrantDataItem.WarrantRic.Count; i++)
            {
                Range rangeDate = ExcelUtil.GetRange(rowIndex + i, 1, wSheet);
                rangeDate.NumberFormatLocal = "@";
                wSheet.Cells[rowIndex + i, 1] = formatToday;
                wSheet.Cells[rowIndex + i, 2] = "<" + _warrantDataItem.WarrantRic[i] + ">";
                wSheet.Cells[rowIndex + i, 3] = _warrantDataItem.Issuer[i];
                Range rangeIssueSize = ExcelUtil.GetRange(rowIndex + i, 4, wSheet);
                rangeIssueSize.NumberFormatLocal = "@";
                wSheet.Cells[rowIndex + i, 4] = _warrantDataItem.IssueSize[i];
                wSheet.Cells[rowIndex + i, 5] = _warrantDataItem.IssuePrice[i];
                wSheet.Cells[rowIndex + i, 6] = _warrantDataItem.UnderlyingRic[i];

                Range convRationRange = ExcelUtil.GetRange(rowIndex + i, 7, wSheet);
                convRationRange.NumberFormatLocal = "@";

                string convRatio = decimal.Parse(_warrantDataItem.ConvRatio[i].Trim(), NumberStyles.Float).ToString();
                wSheet.Cells[rowIndex + i, 7] = convRatio;
                wSheet.Cells[rowIndex + i, 8] = _warrantDataItem.Strike[i];

            }
        }

        public void CopyData(Worksheet fromSheet, Worksheet toSheet)
        {
            int masterRowsCount = fromSheet.UsedRange.Rows.Count;
            Range fromRange = ExcelUtil.GetRange(4, 2, masterRowsCount, 8, fromSheet);
            Range toRange = ExcelUtil.GetRange(86, 1, toSheet);

            fromRange.Copy(Missing.Value);

            toRange.PasteSpecial(XlPasteType.xlPasteFormulas, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);

        }
    }

    public class DeleteItems
    {
        private List<string> deletedItem = new List<string>();
        public List<string> DeletedItem
        {
            get
            {
                return deletedItem;
            }
        }

        public void GetDeleteItemsFromWorkBook(Workbook hkSEABook)
        {

            if (hkSEABook != null)
            {
                Worksheet hkSEASheet = ExcelUtil.GetWorksheet("Sheet2", hkSEABook);

                if (hkSEASheet != null)
                {
                    GetDeletedItems(hkSEASheet);
                }

            }
        }

        private void GetDeletedItems(Worksheet hkSEASheet)
        {
            int rowCount = hkSEASheet.UsedRange.Rows.Count;
            string col3Value;
            string colFourValue;
            DateTime now = DateTime.Now;
            string formatDate = TimeUtil.GetDDMMYYWithSplit(now);
            for (int i = 1; i <= rowCount; i++)
            {
                Range rangeColFour = ExcelUtil.GetRange(i, 4, hkSEASheet);
                if (rangeColFour.Value[Missing.Value] == null)
                {
                    continue;
                }
                colFourValue = rangeColFour.Value[Missing.Value].ToString();
                if (colFourValue.Equals(formatDate))
                {
                    Range rangeColThr = ExcelUtil.GetRange(i, 3, hkSEASheet);
                    if (rangeColThr.Value[Missing.Value] != null)
                    {
                        col3Value = rangeColThr.Value[Missing.Value].ToString();
                        if (col3Value.Equals("DELETE"))
                        {
                            Range rangeColTwo = ExcelUtil.GetRange(i, 2, hkSEASheet);
                            string colTwoValue = rangeColTwo.Value[Missing.Value].ToString();
                            deletedItem.Add(colTwoValue);
                        }
                    }
                }
            }
        }
    }

    class DeleteWarrantData
    {
        public void DeleteDataByRows(Worksheet masterSheet, List<string> deletedItem)
        {
            int rowCount = masterSheet.UsedRange.Rows.Count;
            for (int rowIndex = 4; rowIndex < rowCount + 1; rowIndex++)
            {
                Range range = ExcelUtil.GetRange(rowIndex, 2, masterSheet);
                if (range.Value[Missing.Value] == null)
                {
                    break;
                }
                string cellValue = range.Value[Missing.Value].ToString().TrimStart('<').TrimEnd('>').ToString();
                foreach (string item in deletedItem.Where(item => item.Equals(cellValue)))
                {
                    range.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
        }
    }

    public class HKSpeedGuide : GeneratorBase
    {
        private const string XmlFilePath = ".\\Config\\HK\\HK_SpeedGuide.xml";
        private static HKSpeedGuideConfig configObj;
        private static HKSpeedGuideXML xmlObj;

        protected override void Start()
        {
            StartSpeedGuideJob();
        }

        protected override void Initialize()
        {
            base.Initialize();

            IsEikonExcelDisable = true;
            configObj = Config as HKSpeedGuideConfig;
            xmlObj = ConfigUtil.ReadConfig(XmlFilePath, typeof(HKSpeedGuideXML)) as HKSpeedGuideXML;

        }

        WarrantDataItem warrantDataItem = new WarrantDataItem();
        WarrantDataWriter warrantDataWriter = new WarrantDataWriter();
        DeleteItems deleteItems = new DeleteItems();
        DeleteWarrantData deleteWarrantDataLines = new DeleteWarrantData();

        public void StartSpeedGuideJob()
        {

            var xlApp = new Application();
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            LogMessage("Getting data from folder warrant");
            GetDataFromFolderWarrent(xlApp, xmlObj, Logger);

            LogMessage("Write into Master");
            Workbook warGuideBook = WriteIntoMaster(xlApp);

            LogMessage("Get deleted items");
            GetDeleteItems(xlApp);

            LogMessage("Delete rows");
            DeleteRows();

            LogMessage("Copy workbook");
            Copy(warGuideBook);

            AddResult("war_GuideB.xls", warGuideBook.FullName, "file");

            LogMessage("Generating result files");
            generateTXTFile(warGuideBook);
            generateXMLFile(warGuideBook);

            LogMessage("war_guideB.xls was updated.");
        }


        private void generateXMLFile(Workbook warGuideBook)
        {
            ProductionXmlFileTemplate xmlObj = getSpeedGuideXmlFileContent(warGuideBook);
            ConfigUtil.WriteXml(Path.Combine(configObj.WarGuideBFolderPath, "War_guideB.xml"), xmlObj);
            AddResult("war_guideB.xml (for FTP)", Path.Combine(configObj.WarrantGuideBTxtFileDir, "war_guideB.xml"), "file");
        }

        private ProductionXmlFileTemplate getSpeedGuideXmlFileContent(Workbook warGuideBWorkbook)
        {
            var template = new ProductionXmlFileTemplate();
            Worksheet insertV3Sheet = ExcelUtil.GetWorksheet(configObj.WorksheetWarrantGuideBInsertV3, warGuideBWorkbook);
            if (insertV3Sheet == null)
            {
                LogMessage(string.Format("There's no worksheet {0} in workbook {1}", configObj.WorksheetWarrantGuideBInsertV3, warGuideBWorkbook.FullName));
            }

            int lastAddedRicNum;
            for (int i = 86; i < 1966; i++)
            {
                if (i % 20 == 6)
                {
                    Core.Ric ric = new Core.Ric { Name = string.Format("HK/WTS{0}", (i / 20 + 2).ToString("D2")) };
                    template.rics.rics.Add(ric);
                }

                Fid fid = new Fid { Id = ((i - 86) % 20) + 319, Value = getFidValue(insertV3Sheet, i) };

                lastAddedRicNum = template.rics.rics.Count - 1;
                template.rics.rics[lastAddedRicNum].fids.Add(fid);
            }
            return template;
        }

        private string getFidValue(Worksheet worksheet, int rowNum)
        {
            string value = string.Empty;
            string colVal = string.Empty;
            try
            {
                value += "\"";

                colVal = ExcelUtil.GetRange(rowNum, 1, worksheet).Value2 == null ? string.Empty.PadRight(11) : ExcelUtil.GetRange(rowNum, 1, worksheet).Value2.ToString().Trim().PadRight(11);
                value += colVal;

                string issuer = ExcelUtil.GetRange(rowNum, 2, worksheet).Value2 == null ? string.Empty.PadRight(18) : ExcelUtil.GetRange(rowNum, 2, worksheet).Value2.ToString().Trim();
                value += (issuer.Length > 16 ? issuer.Substring(0, 16) : issuer).PadRight(18);

                string issueSize = ExcelUtil.GetRange(rowNum, 3, worksheet).Value2 == null ? string.Empty.PadLeft(9) : ExcelUtil.GetRange(rowNum, 3, worksheet).Value2.ToString().Trim();
                value += (issueSize.Length > 9 ? issueSize.Substring(0, 9) : issueSize).PadLeft(9);

                colVal = ExcelUtil.GetRange(rowNum, 4, worksheet).Value2 == null ? string.Empty.PadLeft(11) : ExcelUtil.GetRange(rowNum, 4, worksheet).Value2.ToString().Trim().PadLeft(11);
                value += colVal;

                colVal = ExcelUtil.GetRange(rowNum, 5, worksheet).Value2 == null ? string.Empty.PadLeft(11) : ExcelUtil.GetRange(rowNum, 5, worksheet).Value2.ToString().Trim().PadLeft(11);
                value += colVal;


                if (ExcelUtil.GetRange(rowNum, 6, worksheet).Value2 != null)
                {
                    string convRatio = decimal.Parse(ExcelUtil.GetRange(rowNum, 6, worksheet).Value2.ToString().Trim(), NumberStyles.Float).ToString();
                    value += (convRatio.Length > 9 ? convRatio.Substring(0, 9) : convRatio).PadLeft(10);
                }

                else
                {
                    value += string.Empty.PadLeft(10);
                }

                if (ExcelUtil.GetRange(rowNum, 7, worksheet).Value2 != null)
                {
                    string strike = decimal.Parse(ExcelUtil.GetRange(rowNum, 7, worksheet).Value2.ToString().Trim(), NumberStyles.Float).ToString();
                    value += (strike.Length > 8 ? strike.Substring(0, 8) : strike).PadLeft(10);
                }
                else
                {
                    value += string.Empty.PadLeft(10);
                }
                value += "\"";
            }
            catch (Exception ex)
            {
                LogMessage(string.Format("Get fid value failed. Row number: {0}.Exception message: {1}", rowNum, ex.Message));
            }
            return value;
        }

        private void generateTXTFile(Workbook warGuideBWorkbook)
        {
            Worksheet insertV3Sheet = ExcelUtil.GetWorksheet(configObj.WorksheetWarrantGuideBInsertV3, warGuideBWorkbook);
            if (insertV3Sheet == null)
            {
                LogMessage(string.Format("There's no worksheet {0} in workbook {1}", configObj.WorksheetWarrantGuideBInsertV3, warGuideBWorkbook.FullName));
            }

            int lastUsedRow = insertV3Sheet.UsedRange.Row + insertV3Sheet.UsedRange.Rows.Count - 1;
            var sb = new StringBuilder();
            sb.AppendLine("HKSE;29-JUN-29;7:1;TPS;");
            sb.AppendLine("ROW80_5;ROW80_6;ROW80_7;ROW80_8;ROW80_9;ROW80_10;ROW80_11;ROW80_12;ROW80_13;ROW80_14;ROW80_15;ROW80_16;ROW80_17;ROW80_18;ROW80_19;ROW80_20;ROW80_21;ROW80_22;ROW80_23;ROW80_24;ROW80_25;");
            sb.Append("HK/WTS06;");
            using (var reader = new ExcelLineWriter(insertV3Sheet, 86, 1, ExcelLineWriter.Direction.Right))
            {
                string line = string.Empty;
                int pageNum = -1;
                while (reader.Row <= lastUsedRow)
                {
                    line += reader.ReadLineCellText().PadRight(9);
                    line += reader.ReadLineCellText().PadRight(16);
                    line += reader.ReadLineCellText().PadLeft(9);
                    line += reader.ReadLineCellText().PadLeft(8);
                    line += reader.ReadLineCellText().PadRight(10);
                    line += reader.ReadLineCellText().PadLeft(9);
                    line += reader.ReadLineCellText().PadLeft(8);
                    line += ";";
                    sb.Append(line);
                    line = string.Empty;
                    if (reader.Row % 20 == 5)
                    {
                        pageNum = reader.Row / 20 + 1;
                        sb.AppendLine(string.Format("                                               PREV <HK/WTS{0}> / NEXT <HK/WTS{1}>;", (pageNum - 1).ToString("D2"), (pageNum + 1).ToString("D2")));
                        line = string.Format("HK/WTS{0};", (pageNum + 1).ToString("D2"));
                    }
                    reader.PlaceNext(reader.Row + 1, 1);
                }
            }

            using (var sw = new StreamWriter(Path.Combine(configObj.WarrantGuideBTxtFileDir, "war_guideB.txt"), false, Encoding.UTF8))
            {
                sw.WriteLine(sb.ToString());
                sw.AutoFlush = true;
                sw.Close();
                AddResult("war_guideB.txt", Path.Combine(configObj.WarrantGuideBTxtFileDir, "war_guideB.txt"), "file");
            }
        }

        private void GetDataFromFolderWarrent(Application xlApp, HKSpeedGuideXML xmlObj, Logger logger)
        {
            string folderPath = configObj.WarrantFolderPath;    //@"D:\HKRicTemplate\Warrant";
            List<string> fileNames = FileUtil.GetTodayFileNameFromFolder(folderPath);
            if (fileNames != null && fileNames.Count > 0)
            {
                foreach (Workbook wBook in fileNames.Select(fileName => ExcelOperation.OpenWorkBook(xlApp, folderPath, fileName)))
                {
                    warrantDataItem.GetWarrantDataFromWorkBook(wBook, configObj, xmlObj, logger);
                    //wBook.Close(Missing.Value, Missing.Value, Missing.Value);
                }
            }
        }

        private Workbook WriteIntoMaster(Application xlApp)
        {
            string folderPath = configObj.WarGuideBFolderPath;
            string fileName = "war_guideB.xls";
            Workbook warGuideBook = ExcelOperation.OpenWorkBook(xlApp, folderPath, fileName);

            CopyWarGuideBToBak(warGuideBook);

            warGuideBook = ExcelOperation.OpenWorkBook(xlApp, folderPath, fileName);
            warrantDataWriter.WriteIntoMasterSheet(warGuideBook, warrantDataItem, Logger);
            return warGuideBook;
        }

        private void CopyWarGuideBToBak(_Workbook warGuideBook)
        {
            DateTime yesterday = DateTime.Now.AddDays(-1);
            string formatYesterday = "_" + yesterday.Day + yesterday.Month + yesterday.Year;
            string savePath = configObj.BackupWarGuideBFolderPath + "war_guideB" + formatYesterday + ".xls";
            if (!Directory.Exists(Path.GetDirectoryName(savePath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(savePath));
            }
            warGuideBook.Application.DisplayAlerts = false;
            warGuideBook.SaveAs(savePath, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            warGuideBook.Close(Missing.Value, Missing.Value, Missing.Value);
        }

        private void GetDeleteItems(Application xlApp)
        {
            string folderPath = configObj.HKSeaFolderPath;
            string fileName = "HKSEA-Z2.xls";
            Workbook hkSEABook = ExcelOperation.OpenWorkBook(xlApp, folderPath, fileName);
            deleteItems.GetDeleteItemsFromWorkBook(hkSEABook);
            hkSEABook.Close(Missing.Value, Missing.Value, Missing.Value);
        }

        private void DeleteRows()
        {
            List<string> deleteItem = deleteItems.DeletedItem;
            if (deleteItem != null && deleteItem.Count > 0)
            {
                deleteWarrantDataLines.DeleteDataByRows(warrantDataWriter.MasterSheet, deleteItem);
            }
        }

        // Copy from MASTER (B4) to INSERT V3
        private void Copy(Workbook warGuideBook)
        {
            Worksheet masterSheet = ExcelUtil.GetWorksheet("MASTER", warGuideBook); //warrantDataWriter.MasterSheet;
            Worksheet insertSheet = ExcelUtil.GetWorksheet("INSERT V3", warGuideBook);
            if (masterSheet != null && insertSheet != null)
            {
                warrantDataWriter.CopyData(masterSheet, insertSheet);
                warGuideBook.Application.DisplayAlerts = false;
                warGuideBook.SaveAs(warGuideBook.FullName);
                //warGuideBook.Save();
            }
        }
    }

    [ConfigStoredInDB]
    public class HKSpeedGuideConfig
    {
        [StoreInDB]
        [Category("Folder")]
        [DisplayName("Warrant")]
        public string WarrantFolderPath { get; set; }

        [StoreInDB]
        [Category("Folder")]
        [DisplayName("Wrt Guide B")]
        public string WarGuideBFolderPath { get; set; }

        [StoreInDB]
        [Category("Folder")]
        [DisplayName("Backup Wrt Guide B")]
        public string BackupWarGuideBFolderPath { get; set; }

        [StoreInDB]
        [Category("Folder")]
        [DisplayName("HK SEA")]
        public string HKSeaFolderPath { get; set; }

        [StoreInDB]
        [Category("Folder")]
        [DisplayName("Wrt Guide B Txt")]
        public string WarrantGuideBTxtFileDir { get; set; }

        [StoreInDB]
        [Category("Worksheet")]
        [DisplayName("Warrant Guide B Insert V3")]
        public string WorksheetWarrantGuideBInsertV3 { get; set; }
    }

    public class NameMapping
    {
        public string OrganisationName { get; set; }
        public string mappingName { get; set; }
    }

    public class HKSpeedGuideXML
    {
        public List<NameMapping> NameConfig;
    }
}
