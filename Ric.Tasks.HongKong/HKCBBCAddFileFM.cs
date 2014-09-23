using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    [ConfigStoredInDB]
    public class HKCBBCAddFileFMConfig
    {
        [StoreInDB]
        public string CBBC_SEC_FILES_DIR { get; set; }

        [StoreInDB]
        public string CBBC_PRODUCER_MACRO_FILE_PATH { get; set; }

        [StoreInDB]
        public string WRT_ADD_FILE_PATH { get; set; }

        [StoreInDB]
        public string CBBC_SEC_FILE_WORKSHEET { get; set; }

        [StoreInDB]
        public string MACRO_FM_WORKSHEET { get; set; }

        [StoreInDB]
        public string MACRO_INPUT_WORKSHEET { get; set; }

        [StoreInDB]
        public string MACRO_OUTPUT_WORKSHEET { get; set; }

        [StoreInDB]
        public int HOLIDAY_COUNT { get; set; }
    }

    public class BullBearInfo
    {
        public string Code;
        public string Ratio;
        public string Date;
    }

    public class HKCBBCAddFileFM : GeneratorBase
    {
        private static readonly string HOLIDAY_FILE_PATH = ".\\Config\\HK\\Holiday.xml";
        private static HKCBBCAddFileFMConfig configObj = null;
        private static List<DateTime> holidayList = null;

        protected override void Start()
        {
            StartCBBCAddFileFMJobs();
        }

        protected override void Initialize()
        {
            base.Initialize();
            IsEikonExcelDisable = true;
            configObj = Config as HKCBBCAddFileFMConfig;
            holidayList = ConfigUtil.ReadConfig(HOLIDAY_FILE_PATH, typeof(List<DateTime>)) as List<DateTime>;
        }


        public void StartCBBCAddFileFMJobs()
        {
            using (ExcelApp appMacro = new ExcelApp(false, false))
            {
                LogMessage(" - 1 - ");
                var workbookMacro = ExcelUtil.CreateOrOpenExcelFile(appMacro, configObj.CBBC_PRODUCER_MACRO_FILE_PATH);
                Worksheet worksheetFM = ExcelUtil.GetWorksheet(configObj.MACRO_FM_WORKSHEET, workbookMacro);

                LogMessage(" - 2 - ");
                //Copy content from SEC files into new CBBC Producer MACRO
                using (ExcelApp appSec = new ExcelApp(false, false))
                {
                    LogMessage(" - 2 - 1 - ");
                    Range toCopyRange = MergeCBBCSECFiles(configObj.CBBC_SEC_FILES_DIR, appSec);
                    toCopyRange.Copy(Missing.Value);
                    Range r = ExcelUtil.GetRange(1, 1, worksheetFM);
                    LogMessage(" - 3 - ");
                    r.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);
                    ExcelUtil.GetRange(1, 1, worksheetFM).Copy(Missing.Value);
                }

                LogMessage(" - 4 - ");
                //Run Macro: IDN_CBBC
                worksheetFM.Activate();
                appMacro.ExcelAppInstance.GetType().InvokeMember("Run",
                    BindingFlags.Default | BindingFlags.InvokeMethod,
                    null,
                    appMacro.ExcelAppInstance,
                    new object[] { "IDN_CBBC" });

                LogMessage(" - 5 - ");
                var worksheetInput = ExcelUtil.GetWorksheet(configObj.MACRO_INPUT_WORKSHEET, workbookMacro);
                UpdateRation(worksheetInput);

                LogMessage(" - 6 - ");
                var worksheetOutput = ExcelUtil.GetWorksheet(configObj.MACRO_OUTPUT_WORKSHEET, workbookMacro);
                //int lastUsedColNum = worksheetOutput.UsedRange.Column + worksheetOutput.UsedRange.Count - 1;
                int rowNum = 2;
                int lastUsedColNum = 1;
                while (ExcelUtil.GetRange(rowNum, 2, worksheetOutput).Text != null && ExcelUtil.GetRange(rowNum, 2, worksheetOutput).Text.ToString().Trim() != "0")
                {
                    rowNum++;
                }

                LogMessage(" - 7 - ");
                while (ExcelUtil.GetRange(1, lastUsedColNum, worksheetOutput).Text != null && ExcelUtil.GetRange(1, lastUsedColNum, worksheetOutput).Text.ToString() != "")
                {
                    lastUsedColNum++;
                }
                LogMessage(" - 8 - ");
                Range r1 = ExcelUtil.GetRange(1, 1, rowNum - 1, lastUsedColNum, worksheetOutput);
                using (ExcelApp appTarget = new ExcelApp(false, false))
                {
                    var workbookTarget = ExcelUtil.CreateOrOpenExcelFile(appTarget, configObj.WRT_ADD_FILE_PATH);
                    Worksheet worksheetTarget = workbookTarget.Worksheets[1] as Worksheet;
                    Range r2 = ExcelUtil.GetRange(1, 1, worksheetTarget);
                    r1.Copy(Missing.Value);

                    LogMessage(" - 9 - ");

                    r2.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);
                    ExcelUtil.GetRange(1, 1, worksheetTarget).Copy(Missing.Value);
                    workbookMacro.Close(false, configObj.CBBC_PRODUCER_MACRO_FILE_PATH, false);
                    UpdateBNColumn(worksheetTarget);
                    worksheetTarget.Activate();

                    LogMessage(" - 10 - ");

                    UpdateCGColumn(worksheetTarget, holidayList);
                    string targetFilePath = GetTargetFilePath(configObj.WRT_ADD_FILE_PATH);
                    workbookTarget.SaveCopyAs(targetFilePath);
                    //TaskResultList.Add(new TaskResultEntry(Path.GetFileName(targetFilePath), "", targetFilePath));
                    AddResult(Path.GetFileName(targetFilePath), targetFilePath, "");
                    workbookTarget.Close(false, configObj.WRT_ADD_FILE_PATH, false);
                }
            }
        }

        public void UpdateCGColumn(Worksheet worksheet, List<DateTime> holidayList)
        {
            int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            ExcelUtil.GetRange(1, 32, lastUsedRow, 32, worksheet).NumberFormat = "@";
            ExcelUtil.GetRange(1, 1, lastUsedRow, 150, worksheet).NumberFormat = "@";
            ExcelUtil.GetRange(1, 1, lastUsedRow, 151, worksheet).ClearFormats();
            for (int i = 2; i <= lastUsedRow; i++)
            {
                Range expirationDateRange = ExcelUtil.GetRange(i, 32, worksheet);
                if (expirationDateRange.Value2 != null && expirationDateRange.Value2.ToString().Trim() != string.Empty)
                {
                    DateTime expirationDate = DateTime.ParseExact(expirationDateRange.Value2.ToString().Trim(), "dd/MM/yyyy", null);
                    worksheet.Cells[i, 85] = "Last Trading Day is " + MiscUtil.GetLastTradingDay(expirationDate, holidayList, 1).ToString("dd-MMM-yyyy");
                }
            }
        }

        //Update BN underlying column: (Removing the first number redundant "0", 00005.HK --> 0005.HK, 00038.HK --> 0038.HK, and changes "HSC" to "HSCE"
        //There's a bug in the macro, this method is a work around
        public void UpdateBNColumn(Worksheet worksheet)
        {
            int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 2, 66, ExcelLineWriter.Direction.Down))
            {
                while (reader.Row <= lastUsedRow)
                {
                    var cellValue = ExcelUtil.GetRange(reader.Row, reader.Col, worksheet).Text;
                    if (cellValue != null)
                    {
                        string text = cellValue.ToString().Trim().ToUpper();
                        if (text.Contains(".HSC"))
                        {
                            reader.WriteInPlace(text.Replace(".HSC", ".HSCE"));
                        }
                        //00005.HK
                        else if (text.Contains(".HK"))
                        {
                            string underlyingCode = text.Substring(0, text.IndexOf('.'));
                            if (underlyingCode.Length == 5 && underlyingCode.StartsWith("0"))
                            {
                                reader.WriteInPlace(text.Remove(0, 1));
                            }
                        }
                    }
                    reader.PlaceNext(reader.Row + 1, reader.Col);
                }
            }
        }
        //Get target file path
        public string GetTargetFilePath(string oldFilePath)
        {
            string dir = Path.GetDirectoryName(oldFilePath);

            string newFileName = NewFileName();
            string newFilePath = Path.Combine(dir, newFileName);
            if (File.Exists(newFilePath))
            {
                File.Delete(newFilePath);
            }
            return newFilePath;
        }

        //Generate new file name according to the current datetime

        public string NewFileName()
        {
            string[] month = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            string fileName = "WRT_ADD";
            fileName += "_";
            string currentDay = DateTime.Now.ToString("dd_MM_yyyy");
            string[] dateTime = currentDay.Split('_');
            fileName += dateTime[0];
            fileName += month[int.Parse(dateTime[1])];
            fileName += dateTime[2];
            fileName += "_hongkongturbo";
            fileName += ".csv";
            return fileName;
        }

        //Update Ration column
        public void UpdateRation(Worksheet worksheetInput)
        {
            List<BullBearInfo> bullBearInfoList = GetBullBearInfo();
            using (ExcelLineWriter writer = new ExcelLineWriter(worksheetInput, 4, 11, ExcelLineWriter.Direction.Down))
            {
                while ((ExcelUtil.GetRange(writer.Row, 1, worksheetInput).Text != null) && (ExcelUtil.GetRange(writer.Row, 1, worksheetInput).Text.ToString().Trim() != "0"))
                {
                    foreach (BullBearInfo bullBear in bullBearInfoList.Where(bullBear => ExcelUtil.GetRange(writer.Row, 1, worksheetInput).Text.ToString() == bullBear.Code))
                    {
                        writer.WriteLine(bullBear.Ratio);
                        break;
                    }
                }
            }
        }

        //Get Ration info from http://www.hkex.com.hk/eng/cbbc/newissue/newlaunch.htm
        public List<BullBearInfo> GetBullBearInfo()
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = null;
            string uri = "http://www.hkex.com.hk/eng/cbbc/newissue/newlaunch.htm";
            htmlDoc = WebClientUtil.GetHtmlDocument(uri, 180000);
            HtmlAgilityPack.HtmlNodeCollection normalNodeList = htmlDoc.DocumentNode.SelectNodes("//table/tbody/tr[@class='tr_normal']");
            return (from normalNode in normalNodeList
                    select normalNode.ChildNodes
                        into tableblankNodeList
                        where DateEqualLastBusinessDay(tableblankNodeList[2 * 11 + 1].InnerText)
                        select new BullBearInfo
                        {
                            Code = tableblankNodeList[2 * 1 + 1].ChildNodes[0].InnerText,
                            Date = tableblankNodeList[2 * 11 + 1].InnerText,
                            Ratio = tableblankNodeList[2 * 9 + 1].InnerText
                        }).ToList();
        }

        public bool DateEqualLastBusinessDay(string time)
        {
            string lastBusinessDay = MiscUtil.GetLastBusinessDay(configObj.HOLIDAY_COUNT, DateTime.Now).ToString("dd-MM-yyyy");
            return time.Trim() == lastBusinessDay;
        }


        //Copy all the SEC files content into the first Sec file and return the first sec file path
        public Range MergeCBBCSECFiles(string CBBCSecFileDir, ExcelApp appSec)
        {
            string[] secFilePathArr = GetAllSecFiles(CBBCSecFileDir);
            Range toCopyRange = null;
            LogMessage(" - 2 - 2- ");
            if (secFilePathArr == null || secFilePathArr.Length == 0)
            {
                LogMessage(string.Format("There's no CBBC SEC file, please check in directory {0}", CBBCSecFileDir));
            }
            else
            {
                LogMessage(" - 2 - 3 - ");
                var workbookSEC1 = ExcelUtil.CreateOrOpenExcelFile(appSec, secFilePathArr[0]);
                var worksheetSEC1 = ExcelUtil.GetWorksheet(configObj.CBBC_SEC_FILE_WORKSHEET, workbookSEC1);
                int curLastRowNum = worksheetSEC1.UsedRange.Row + worksheetSEC1.UsedRange.Count - 1;
                if (secFilePathArr.Length > 1)
                {
                    LogMessage(" - 2 - 4 - ");
                    for (int i = 1; i < secFilePathArr.Length; i++)
                    {
                        string secFilePath = secFilePathArr[i];
                        curLastRowNum = worksheetSEC1.UsedRange.Row + worksheetSEC1.UsedRange.Rows.Count - 1;
                        using (ExcelApp app = new ExcelApp(false))
                        {
                            LogMessage(" - 2 - 5 - ");
                            var workbookSEC = ExcelUtil.CreateOrOpenExcelFile(appSec, secFilePath);
                            var worksheetSEC = ExcelUtil.GetWorksheet(configObj.CBBC_SEC_FILE_WORKSHEET, workbookSEC);
                            int lastUsedRowNum = worksheetSEC.UsedRange.Row + worksheetSEC.UsedRange.Rows.Count - 1;
                            LogMessage(" - 2 - 6 - ");
                            Range r = ExcelUtil.GetRange(8, 1, lastUsedRowNum, 3, worksheetSEC);
                            r.Copy(Missing.Value);
                            LogMessage(" - 2 - 7 - ");
                            ExcelUtil.GetRange(curLastRowNum + 1, 1, worksheetSEC1).PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);
                            ExcelUtil.GetRange(1, 1, worksheetSEC1).Copy(Missing.Value);
                            LogMessage(" - 2 - 8 - ");
                            workbookSEC.Close(false, workbookSEC.FullName, false);
                        }
                    }
                }
                LogMessage(" - 2 - 9 - ");
                int lastUsedRow = worksheetSEC1.UsedRange.Row + worksheetSEC1.UsedRange.Rows.Count - 1;
                toCopyRange = ExcelUtil.GetRange(1, 1, lastUsedRow, 3, worksheetSEC1);
            }
            LogMessage(" - 2 - 10 - ");
            return toCopyRange;
        }


        //Get all the CBBC SEC file path in CBBC SEC file directory
        public string[] GetAllSecFiles(string CBBCSecFileDir)
        {
            string[] fileArr = Directory.GetFiles(CBBCSecFileDir, "*.xls", SearchOption.TopDirectoryOnly);
            return fileArr;
        }
    }
}
