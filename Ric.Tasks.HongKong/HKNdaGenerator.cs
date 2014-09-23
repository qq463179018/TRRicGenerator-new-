using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    #region Configuration

    [ConfigStoredInDB]
    public class HKNDAGeneratorFMConfig
    {
        [StoreInDB]
        [Category("Files")]
        [Description("The path of the folder where the NDA files will be created.")]
        public string ResultsFolder { get; set; }

        [StoreInDB]
        [Category("Files")]
        [Description("The path of the folder where the FM files are.")]
        public string SourceFolder { get; set; }

        [StoreInDB]
        [Category("FM Type")]
        [Description("The Type of FM file\nPossible Values are : GEM / MBD / Rights / ETF / REIT / HShare / LotSize / Name Change / MBD Parallel / GEM Parallel")]
        public string Type { get; set; }
    }

    #endregion

    #region Task

    public class HKNDAGeneratorFM : GeneratorBase
    {
        #region Declaration

        private static HKNDAGeneratorFMConfig configObj;

        delegate void FillCSV(Dictionary<string, string> content, int index);

        private List<string> ndaTitle = new List<string> { "RIC", "ASSET COMMON NAME", "ASSET SHORT NAME", "CURRENCY", "EXCHANGE", "TAG", "TYPE", "CATEGORY", "TICKER SYMBOL", "ROUND LOT SIZE", "TRADING SEGMENT", "BASE ASSET", "EQUITY FIRST TRADING DAY", "SETTLEMENT PERIOD" };
        private List<string> ndaTitleChange = new List<string> { "RIC", "ASSET COMMON NAME", "ASSET SHORT NAME" };
        private List<string> ndaIaTitle = new List<string> { "TYPE", "CATEGORY", "EQUITY ISSUER", "RCS ASSET CLASS", "ASSET COMMON NAME" };
        private List<string> ndaIaTitleChange = new List<string> { "BASE ASSET", "ASSET COMMON NAME" };
        private List<string> fieldList = new List<string> { "Chain RIC:", "Currency:", "Lot Size:", "Displayname:", "Effective Date:", "Official Code:", "Legal Registered Name:" };
        private List<string> typeList = new List<string> { "LotSize", "HShare", "GEM Parallel", "GEM", "RTS", "RIGHTS", "REIT", "ETF", "MBD Parallel", "Name Change", "MBD" };
        private int[] typeNb = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        private ExcelApp app;

        #endregion

        #region Interface implementation

        /// <summary>
        /// Getting files in given folder
        /// Then take informations file by file and create new NDA files
        /// </summary>
        protected override void Start()
        {
            string[] fileEntries = Directory.GetFiles(configObj.SourceFolder);
            FillCSV[] filling = { FillLotSize, FillHShare, FillGEMParallel, FillGEM, FillRights, FillRights, FillREIT, FillETF, FillMBDParallel, FillNameChange, FillMBD };
            var result = new Dictionary<string, string>();
            int index;

            foreach (string fileName in fileEntries)
            {
                LogMessage("Reading file");
                result = configObj.Type.Contains("Parallel") ? ReadExcelPara(fileName) : ReadExcel(fileName);
                LogMessage("Cleaning results");
                CleanResult(ref result);
                index = typeList.FindIndex(s => s == configObj.Type);
                LogMessage("Filling new file");
                filling[index](result, typeNb[index]);
                typeNb[index]++;
            }
            CleanExit();
        }

        /// <summary>
        /// Basic initialization
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as HKNDAGeneratorFMConfig;

            app = new ExcelApp(false, false);
            if (app.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !";
                LogMessage(msg, Logger.LogType.Error);
            }
            if (!typeList.Contains(configObj.Type))
            {
                throw (new Exception("FM Type not supported : Choose between those values : GEM / MBD / Rights / ETF / REIT / HShare / LotSize / Name Change / MBD Parallel / GEM Parallel"));
            }
        }

        /// <summary>
        /// Close the Excel App in case of error raised
        /// If not this function is last called to be sure everything was closed correctly for the next call to this task
        /// </summary>
        /// <returns>true if success</returns>
        private bool CleanExit()
        {
            app.Dispose();

            return true;
        }

        #endregion

        #region Getting information from FM

        /// <summary>
        /// Read the CSV from path in parameters, parse it and return dictionary with wanted informations
        /// </summary>
        /// <param name="filename"></param>
        /// <returns>A dictionary with the wanted 
        /// field name as key and their value as value</returns>
        private Dictionary<string, string> ReadExcel(string filename)
        {
            var entry = new Dictionary<string, string>();
            var excelValues = new List<List<string>>();
            try
            {
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, filename);
                excelValues = workbook.ToList();

                foreach (List<string> row in excelValues.Where(row => fieldList.Contains(row[0].Replace(" (NEW):", ":"))))
                {
                    if (entry.ContainsKey(row[0].Replace(" (NEW):", ":")))
                    {
                        entry[row[0].Replace(" (NEW):", ":")] += row[1];
                    }
                    else
                    {
                        entry.Add(row[0], row[1]);
                    }
                }
                workbook.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot read CSV file :" + ex;
                Logger.Log(msg, Logger.LogType.Error);
                CleanExit();
            }
            return entry;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private Dictionary<string, string> ReadExcelPara(string filename)
        {
            var entry = new Dictionary<string, string>();
            var excelValues = new List<List<string>>();
            int section = 0;
            try
            {
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, filename);
                excelValues = workbook.ToList();

                foreach (List<string> row in excelValues)
                {
                    if (row[0] == "-1" || row[0] == "-2" || row[0] == "-3")
                    {
                        section++;
                    }
                    if (fieldList.Contains(row[0].Replace(" (NEW):", ":")))
                    {
                        if (entry.ContainsKey(String.Format("[{0}] {1}", section, row[0].Replace(" (NEW):", ":"))))
                        {
                            entry[String.Format("[{0}] {1}", section, row[0].Replace(" (NEW):", ":"))] += row[1];
                        }
                        else
                        {
                            entry.Add(String.Format("[{0}] {1}", section, row[0]), row[1]);
                        }
                    }
                }
                workbook.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot read CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
            return entry;
        }

        /// <summary>
        /// Go through the results and clean them to make further processing easier
        /// </summary>
        /// <param name="content"></param>
        private void CleanResult(ref Dictionary<string, string> content)
        {
            try
            {
                if (content.ContainsKey("Official Code:"))
                {
                    if (content["Official Code:"].Length < 4)
                    {
                        content["Official Code:"] = "0" + content["Official Code:"];
                    }
                }
                if (content.ContainsKey("Displayname:"))
                {
                    int index = content["Displayname:"].IndexOf("<--");
                    if (index > 0)
                    {
                        content["Displayname:"] = content["Displayname:"].Remove(index);
                    }
                    content["Displayname:"] = content["Displayname:"].Replace("-NEW", "");
                    content["Displayname:"] = content["Displayname:"].TrimEnd();
                }

                if (content.ContainsKey("[1] Official Code:"))
                {
                    if (content["[1] Official Code:"].Length < 4)
                    {
                        content["[1] Official Code:"] = "0" + content["[1] Official Code:"];
                    }
                }
                if (content.ContainsKey("[1] Displayname:"))
                {
                    int index = content["[1] Displayname:"].IndexOf("<--");
                    if (index > 0)
                    {
                        content["[1] Displayname:"] = content["[1] Displayname:"].Remove(index);
                    }
                    content["[1] Displayname:"] = content["[1] Displayname:"].Replace("-NEW", "");
                    content["[1] Displayname:"] = content["[1] Displayname:"].TrimEnd();
                }
                if (content.ContainsKey("[2] Official Code:"))
                {
                    if (content["[2] Official Code:"].Length < 4)
                    {
                        content["[2] Official Code:"] = "0" + content["[2] Official Code:"];
                    }
                }
                if (content.ContainsKey("[2] Displayname:"))
                {
                    int index = content["[2] Displayname:"].IndexOf("<--");
                    if (index > 0)
                    {
                        content["[2] Displayname:"] = content["[2] Displayname:"].Remove(index);
                    }
                    content["[2] Displayname:"] = content["[2] Displayname:"].Replace("-NEW", "");
                    content["[2] Displayname:"] = content["[2] Displayname:"].TrimEnd();
                }
                if (content.ContainsKey("Effective Date:"))
                {
                    content["Effective Date:"] = DateTime.FromOADate(Convert.ToDouble(content["Effective Date:"])).ToString("d-MMM-yyyy");
                }
                else if (content.ContainsKey("[0] Effective Date:"))
                {
                    content["[0] Effective Date:"] = DateTime.FromOADate(Convert.ToDouble(content["[0] Effective Date:"])).ToString("d-MMM-yyyy");
                }
            }
            catch (Exception ex)
            {
                string msg = "Error: " + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        #endregion

        #region writing NDA

        /// <summary>
        /// Initialize the Workbook and worksheet depending the type
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="worksheet"></param>
        /// <param name="type"></param>
        private void InitWorkbookAndWorksheet(out Workbook workbook, out Worksheet worksheet, string type)
        {
            string titleWork = configObj.ResultsFolder + type + "_" + DateTime.Now.ToString("dMMMyyyy") + ".csv";
            workbook = ExcelUtil.CreateOrOpenExcelFile(app, titleWork);
            worksheet = workbook.Worksheets[1] as Worksheet;
            worksheet.Name = type;
        }

        /// <summary>
        /// Filling the newly created Lot Size file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillLotSize(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Worksheet worksheetNda;
            string titleWork = configObj.ResultsFolder + "LotSize_" + DateTime.Now.ToString("dMMMyyyy") + ".csv";

            try
            {
                workbookNda = ExcelUtil.CreateOrOpenExcelFile(app, titleWork);
                worksheetNda = workbookNda.Worksheets[1] as Worksheet;
                worksheetNda.Name = "LotSize";

                int firstLine = 2 + index * 3;

                int ricNb = Int32.Parse(content["Official Code:"]);
                // NDA Titles
                if (index == 0)
                {
                    worksheetNda.Cells[1, 1] = "RIC";
                    worksheetNda.Cells[1, 2] = "ROUND LOT SIZE";
                }

                // RIC
                worksheetNda.Cells[firstLine, 1] = content["Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["Official Code:"] + "ta.HK";
                worksheetNda.Cells[firstLine + 2, 1] = content["Official Code:"] + "stat.HK";

                if (content.ContainsKey("Lot Size:"))
                {
                    worksheetNda.Cells[firstLine, 2] = content["Lot Size:"];
                }
                else
                {
                    worksheetNda.Cells[firstLine, 2] = content["Lot Size (NEW):"];
                }

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                workbookNda.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling the newly created Name Change file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillNameChange(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Workbook workbookNdaIa;
            Worksheet worksheetNda;
            Worksheet worksheetNdaIa;
            try
            {
                InitWorkbookAndWorksheet(out workbookNda, out worksheetNda, "Name_Change");
                InitWorkbookAndWorksheet(out workbookNdaIa, out worksheetNdaIa, "Name_Change_IA");
                int firstLine = 2 + index * 3;
                int firstLineIa = 2 + index;
                if (index == 0)
                {
                    FillTitlesNameChange(ref worksheetNda, ref worksheetNdaIa);
                }

                worksheetNda.Cells[firstLine, 1] = content["Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["Official Code:"] + "ta.HK";
                worksheetNda.Cells[firstLine + 2, 1] = content["Official Code:"] + "stat.HK";

                for (int count = 0; count < 3; count++)
                {
                    worksheetNda.Cells[firstLine + count, 2] = content["Displayname (NEW):"] + " ORD";
                    worksheetNda.Cells[firstLine + count, 3] = content["Displayname (NEW):"];
                }

                worksheetNdaIa.Cells[firstLineIa, 2] = content["Legal Registered Name (NEW):"].Replace(" Limited", "") + " Ord Shs";

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNdaIa.SaveAs(workbookNdaIa.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                AddResult(Path.GetFileNameWithoutExtension(workbookNdaIa.FullName), workbookNdaIa.FullName, "nda ia");
                workbookNda.Close();
                workbookNdaIa.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling the newly created H Share IPO file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillHShare(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Workbook workbookNdaIa;
            Worksheet worksheetNda;
            Worksheet worksheetNdaIa;
            try
            {
                InitWorkbookAndWorksheet(out workbookNda, out worksheetNda, "H_Share_IPO");
                InitWorkbookAndWorksheet(out workbookNdaIa, out worksheetNdaIa, "H_Share_IPO_IA");
                int firstLine = 2 + index * 3;
                int firstLineIa = 2 + index;

                if (index == 0)
                {
                    FillTitles(ref worksheetNda, ref worksheetNdaIa);
                    worksheetNda.Cells[1, 15] = "QUOTE SHARE CLASS";
                }

                worksheetNda.Cells[firstLine, 1] = content["Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["Official Code:"] + "ta.HK";
                worksheetNda.Cells[firstLine + 2, 1] = content["Official Code:"] + "stat.HK";

                worksheetNda.Cells[firstLine, 6] = "1";
                worksheetNda.Cells[firstLine + 1, 6] = "64377";
                worksheetNda.Cells[firstLine + 2, 6] = "60019";

                worksheetNda.Cells[firstLine, 9] = content["Official Code:"];
                worksheetNda.Cells[firstLine + 2, 9] = content["Official Code:"];

                worksheetNda.Cells[firstLine, 10] = content["Lot Size:"];

                worksheetNda.Cells[firstLine, 11] = "HKG:XHKG";
                worksheetNda.Cells[firstLine + 2, 11] = "HKG:XHKG";

                for (int count = 0; count < 3; count++)
                {
                    worksheetNda.Cells[firstLine + count, 2] = content["Displayname:"] + " ORD H";
                    worksheetNda.Cells[firstLine + count, 3] = content["Displayname:"];
                    worksheetNda.Cells[firstLine + count, 4] = content["Currency:"];
                    worksheetNda.Cells[firstLine + count, 5] = "HKG";
                    worksheetNda.Cells[firstLine + count, 7] = "EQUITY";
                    worksheetNda.Cells[firstLine + count, 8] = "ORD";
                    worksheetNda.Cells[firstLine + count, 13] = content["Effective Date:"];
                    worksheetNda.Cells[firstLine + count, 15] = "Class H";
                }

                worksheetNdaIa.Cells[firstLineIa, 1] = "EQUITY";
                worksheetNdaIa.Cells[firstLineIa, 2] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 3] = "";
                worksheetNdaIa.Cells[firstLineIa, 4] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 5] = content["Legal Registered Name:"].Replace(" Limited", "") + " Ord Shs H";

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNdaIa.SaveAs(workbookNdaIa.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                AddResult(Path.GetFileNameWithoutExtension(workbookNdaIa.FullName), workbookNdaIa.FullName, "nda ia");
                workbookNda.Close();
                workbookNdaIa.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling the newly created GEM Parallel file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillGEMParallel(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Workbook workbookNdaIa;
            Worksheet worksheetNda;
            Worksheet worksheetNdaIa;
            try
            {
                InitWorkbookAndWorksheet(out workbookNda, out worksheetNda, "GEM_Parallel");
                InitWorkbookAndWorksheet(out workbookNdaIa, out worksheetNdaIa, "GEM_Parallel_IA");
                int firstLine = 2 + index * 3;
                int firstLineIa = 2 + index;

                if (index == 0)
                {
                    FillTitles(ref worksheetNda, ref worksheetNdaIa);
                }

                worksheetNda.Cells[firstLine, 1] = content["Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["Official Code:"] + "ta.HK";
                worksheetNda.Cells[firstLine + 2, 1] = content["Official Code:"] + "stat.HK";

                worksheetNda.Cells[firstLine, 6] = "1";
                worksheetNda.Cells[firstLine + 1, 6] = "64377";
                worksheetNda.Cells[firstLine + 2, 6] = "60019";

                worksheetNda.Cells[firstLine, 9] = content["Official Code:"];
                worksheetNda.Cells[firstLine + 2, 9] = content["Official Code:"];

                worksheetNda.Cells[firstLine, 10] = content["Lot Size (NEW):"];

                worksheetNda.Cells[firstLine, 11] = "HKG:XGEM";
                worksheetNda.Cells[firstLine + 2, 11] = "HKG:XGEM";

                for (int count = 0; count < 3; count++)
                {
                    worksheetNda.Cells[firstLine + count, 2] = content["Displayname:"] + " ORD (TEMP)";
                    worksheetNda.Cells[firstLine + count, 3] = content["Displayname:"];
                    worksheetNda.Cells[firstLine + count, 4] = content["Currency:"];
                    worksheetNda.Cells[firstLine + count, 5] = "HKG";
                    worksheetNda.Cells[firstLine + count, 7] = "EQUITY";
                    worksheetNda.Cells[firstLine + count, 8] = "ORD";
                    worksheetNda.Cells[firstLine + count, 13] = content["Effective Date:"];
                }

                worksheetNdaIa.Cells[firstLineIa, 1] = "EQUITY";
                worksheetNdaIa.Cells[firstLineIa, 2] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 3] = "45707";
                worksheetNdaIa.Cells[firstLineIa, 4] = content["Displayname:"].Replace(" Limited", "") + " Ord Shs";

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNdaIa.SaveAs(workbookNdaIa.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(workbookNdaIa.FullName), workbookNdaIa.FullName, "nda ia");
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                workbookNda.Close();
                workbookNdaIa.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling the newly created MBD Parallel file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillMBDParallel(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Workbook workbookNdaIa;
            Worksheet worksheetNda;
            Worksheet worksheetNdaIa;
            try
            {
                InitWorkbookAndWorksheet(out workbookNda, out worksheetNda, "MBD_Parallel");
                InitWorkbookAndWorksheet(out workbookNdaIa, out worksheetNdaIa, "MBD_Parallel_IA");
                int firstLine = 2 + index * 3;
                int firstLineIa = 2 + index;

                if (index == 0)
                {
                    FillTitles(ref worksheetNda, ref worksheetNdaIa);
                }

                worksheetNda.Cells[firstLine, 1] = content["[2] Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["[2] Official Code:"] + "ta.HK";
                worksheetNda.Cells[firstLine + 2, 1] = content["[2] Official Code:"] + "stat.HK";

                worksheetNda.Cells[firstLine, 6] = "1";
                worksheetNda.Cells[firstLine + 1, 6] = "64377";
                worksheetNda.Cells[firstLine + 2, 6] = "60019";

                worksheetNda.Cells[firstLine, 9] = content["[2] Official Code:"];
                worksheetNda.Cells[firstLine + 2, 9] = content["[2] Official Code:"];

                if (content.ContainsKey("[2] Lot Size (NEW):"))
                {
                    worksheetNda.Cells[firstLine, 10] = content["[2] Lot Size (NEW):"];
                }
                else
                {
                    worksheetNda.Cells[firstLine, 10] = content["[2] Lot Size:"];
                }

                worksheetNda.Cells[firstLine, 11] = "HKG:XHKG";
                worksheetNda.Cells[firstLine + 2, 11] = "HKG:XHKG";

                for (int count = 0; count < 3; count++)
                {
                    worksheetNda.Cells[firstLine + count, 2] = content["Displayname:"] + " ORD (TEMP)";
                    worksheetNda.Cells[firstLine + count, 3] = content["Displayname:"];
                    worksheetNda.Cells[firstLine + count, 4] = content["Currency:"];
                    worksheetNda.Cells[firstLine + count, 5] = "HKG";
                    worksheetNda.Cells[firstLine + count, 7] = "EQUITY";
                    worksheetNda.Cells[firstLine + count, 8] = "ORD";
                    worksheetNda.Cells[firstLine + count, 13] = content["Effective Date:"];
                }

                worksheetNdaIa.Cells[firstLineIa, 1] = "EQUITY";
                worksheetNdaIa.Cells[firstLineIa, 2] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 3] = "45707";
                worksheetNdaIa.Cells[firstLineIa, 4] = content["Displayname:"].Replace(" Limited", "") + " Ord Shs";

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNdaIa.SaveAs(workbookNdaIa.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(workbookNdaIa.FullName), workbookNdaIa.FullName, "nda ia");
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                workbookNda.Close();
                workbookNdaIa.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling the newly created MBD IPO file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillMBD(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Workbook workbookNdaIa;
            Worksheet worksheetNda;
            Worksheet worksheetNdaIa;
            try
            {
                InitWorkbookAndWorksheet(out workbookNda, out worksheetNda, "MBD_IPO");
                InitWorkbookAndWorksheet(out workbookNdaIa, out worksheetNdaIa, "MBD_IPO_IA");
                int firstLine = 2 + index * 3;
                int firstLineIa = 2 + index;

                if (index == 0)
                {
                    FillTitles(ref worksheetNda, ref worksheetNdaIa);
                }

                worksheetNda.Cells[firstLine, 1] = content["Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["Official Code:"] + "ta.HK";
                worksheetNda.Cells[firstLine + 2, 1] = content["Official Code:"] + "stat.HK";

                worksheetNda.Cells[firstLine, 6] = "1";
                worksheetNda.Cells[firstLine + 1, 6] = "64377";
                worksheetNda.Cells[firstLine + 2, 6] = "60019";

                worksheetNda.Cells[firstLine, 9] = content["Official Code:"];
                worksheetNda.Cells[firstLine + 2, 9] = content["Official Code:"];

                worksheetNda.Cells[firstLine, 10] = content["Lot Size:"];

                worksheetNda.Cells[firstLine, 11] = "HKG:XHKG";
                worksheetNda.Cells[firstLine + 2, 11] = "HKG:XHKG";

                for (int count = 0; count < 3; count++)
                {
                    worksheetNda.Cells[firstLine + count, 2] = content["Displayname:"] + " ORD";
                    worksheetNda.Cells[firstLine + count, 3] = content["Displayname:"];
                    worksheetNda.Cells[firstLine + count, 4] = content["Currency:"];
                    worksheetNda.Cells[firstLine + count, 5] = "HKG";
                    worksheetNda.Cells[firstLine + count, 7] = "EQUITY";
                    worksheetNda.Cells[firstLine + count, 8] = "ORD";
                    worksheetNda.Cells[firstLine + count, 13] = content["Effective Date:"];
                    worksheetNda.Cells[firstLine + count, 14] = "T+2";
                }

                worksheetNdaIa.Cells[firstLineIa, 1] = "EQUITY";
                worksheetNdaIa.Cells[firstLineIa, 2] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 3] = "";
                worksheetNdaIa.Cells[firstLineIa, 4] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 5] = content["Legal Registered Name:"].Replace(" Limited", "") + " Ord Shs";

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNdaIa.SaveAs(workbookNdaIa.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(workbookNdaIa.FullName), workbookNdaIa.FullName, "nda ia");
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                workbookNda.Close();
                workbookNdaIa.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling the newly created GEM IPO file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillGEM(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Workbook workbookNdaIa;
            Worksheet worksheetNda;
            Worksheet worksheetNdaIa;
            try
            {
                InitWorkbookAndWorksheet(out workbookNda, out worksheetNda, "GEM_IPO");
                InitWorkbookAndWorksheet(out workbookNdaIa, out worksheetNdaIa, "GEM_IPO_IA");
                int firstLine = 2 + index * 3;
                int firstLineIa = 2 + index;
                int ricNb = Int32.Parse(content["Official Code:"]);
                // NDA Titles
                if (index == 0)
                {
                    FillTitles(ref worksheetNda, ref worksheetNdaIa);
                }

                // RIC
                worksheetNda.Cells[firstLine, 1] = content["Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["Official Code:"] + "ta.HK";
                worksheetNda.Cells[firstLine + 2, 1] = content["Official Code:"] + "stat.HK";

                worksheetNda.Cells[firstLine, 6] = "1";
                worksheetNda.Cells[firstLine + 1, 6] = "64377";
                worksheetNda.Cells[firstLine + 2, 6] = "60019";

                worksheetNda.Cells[firstLine, 9] = ricNb.ToString();
                worksheetNda.Cells[firstLine + 2, 9] = ricNb.ToString();

                worksheetNda.Cells[firstLine, 10] = content["Lot Size:"];

                worksheetNda.Cells[firstLine, 11] = "HKG:XGEM";
                worksheetNda.Cells[firstLine + 2, 11] = "HKG:XGEM";

                for (int count = 0; count < 3; count++)
                {
                    worksheetNda.Cells[firstLine + count, 2] = content["Displayname:"] + " ORD";
                    worksheetNda.Cells[firstLine + count, 3] = content["Displayname:"];
                    worksheetNda.Cells[firstLine + count, 4] = content["Currency:"];
                    worksheetNda.Cells[firstLine + count, 5] = "HKG";
                    worksheetNda.Cells[firstLine + count, 7] = "EQUITY";
                    worksheetNda.Cells[firstLine + count, 8] = "ORD";
                    worksheetNda.Cells[firstLine + count, 13] = content["Effective Date:"];
                    worksheetNda.Cells[firstLine + count, 14] = "T+2";
                }

                worksheetNdaIa.Cells[firstLineIa, 1] = "EQUITY";
                worksheetNdaIa.Cells[firstLineIa, 2] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 3] = "";
                worksheetNdaIa.Cells[firstLineIa, 4] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 5] = content["Legal Registered Name:"].Replace(" Limited", "") + " Ord Shs";

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNdaIa.SaveAs(workbookNdaIa.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                AddResult(Path.GetFileNameWithoutExtension(workbookNdaIa.FullName), workbookNdaIa.FullName, "nda ia");
                workbookNda.Close();
                workbookNdaIa.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling the newly created Right file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillRights(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Workbook workbookNdaIa;
            Worksheet worksheetNda;
            Worksheet worksheetNdaIa;
            try
            {
                InitWorkbookAndWorksheet(out workbookNda, out worksheetNda, "RTS");
                InitWorkbookAndWorksheet(out workbookNdaIa, out worksheetNdaIa, "RTS_IA");

                int firstLine = 2 + index * 2;
                int firstLineIa = 2 + index;
                int ricNb = Int32.Parse(content["Official Code:"]);

                // NDA Titles
                if (index == 0)
                {
                    FillTitles(ref worksheetNda, ref worksheetNdaIa);
                }

                // RIC
                worksheetNda.Cells[firstLine, 1] = content["Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["Official Code:"] + "ta.HK";

                worksheetNda.Cells[firstLine, 6] = "1";
                worksheetNda.Cells[firstLine + 1, 6] = "64377";


                worksheetNda.Cells[firstLine, 9] = ricNb.ToString();

                worksheetNda.Cells[firstLine, 10] = content["Lot Size:"];

                worksheetNda.Cells[firstLine, 11] = "HKG:XHKG";

                for (int count = 0; count < 2; count++)
                {
                    worksheetNda.Cells[firstLine + count, 2] = content["Displayname:"];
                    worksheetNda.Cells[firstLine + count, 3] = content["Displayname:"].Replace(" RTS", "");
                    worksheetNda.Cells[firstLine + count, 4] = content["Currency:"];
                    worksheetNda.Cells[firstLine + count, 5] = "HKG";
                    worksheetNda.Cells[firstLine + count, 7] = "EQUITY";
                    worksheetNda.Cells[firstLine + count, 8] = "RTS";
                    worksheetNda.Cells[firstLine + count, 13] = content["Effective Date:"];
                }

                worksheetNdaIa.Cells[firstLineIa, 1] = "EQUITY";
                worksheetNdaIa.Cells[firstLineIa, 2] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 3] = "";
                worksheetNdaIa.Cells[firstLineIa, 4] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 5] = content["Legal Registered Name:"].Replace(" Limited", "") + " Ord Shs";

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbookNdaIa.SaveAs(workbookNdaIa.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                AddResult(Path.GetFileNameWithoutExtension(workbookNdaIa.FullName), workbookNdaIa.FullName, "nda ia");
                workbookNda.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling the newly created ETF file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillETF(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Workbook workbookNdaIa;
            Worksheet worksheetNda;
            Worksheet worksheetNdaIa;
            try
            {
                InitWorkbookAndWorksheet(out workbookNda, out worksheetNda, "ETF");
                InitWorkbookAndWorksheet(out workbookNdaIa, out worksheetNdaIa, "ETF_IA");

                int firstLine = 2 + index * 3;
                int firstLineIa = 2 + index;
                int ricNb = Int32.Parse(content["Official Code:"]);

                // NDA Titles
                if (index == 0)
                {
                    FillTitles(ref worksheetNda, ref worksheetNdaIa);
                }

                worksheetNda.Cells[firstLine, 1] = content["Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["Official Code:"] + "ta.HK";
                worksheetNda.Cells[firstLine + 2, 1] = content["Official Code:"] + "stat.HK";

                worksheetNda.Cells[firstLine, 6] = "1";
                worksheetNda.Cells[firstLine + 1, 6] = "64377";
                worksheetNda.Cells[firstLine + 2, 6] = "60019";

                worksheetNda.Cells[firstLine, 9] = ricNb.ToString();
                worksheetNda.Cells[firstLine + 2, 9] = ricNb.ToString();

                worksheetNda.Cells[firstLine, 10] = content["Lot Size:"];

                worksheetNda.Cells[firstLine, 11] = "HKG:XHKG";
                worksheetNda.Cells[firstLine + 2, 11] = "HKG:XHKG";

                for (int count = 0; count < 3; count++)
                {
                    worksheetNda.Cells[firstLine + count, 2] = content["Displayname:"] + " ETF";
                    worksheetNda.Cells[firstLine + count, 3] = content["Displayname:"];
                    worksheetNda.Cells[firstLine + count, 4] = content["Currency:"];
                    worksheetNda.Cells[firstLine + count, 5] = "HKG";
                    worksheetNda.Cells[firstLine + count, 7] = "EQUITY";
                    worksheetNda.Cells[firstLine + count, 8] = "ECL";
                    worksheetNda.Cells[firstLine + count, 13] = content["Effective Date:"];
                }

                worksheetNdaIa.Cells[firstLineIa, 1] = "EQUITY";
                worksheetNdaIa.Cells[firstLineIa, 2] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 3] = "";
                worksheetNdaIa.Cells[firstLineIa, 4] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 5] = content["Legal Registered Name:"].Replace(" Limited", "") + " Ord Shs";

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbookNdaIa.SaveAs(workbookNdaIa.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                AddResult(Path.GetFileNameWithoutExtension(workbookNdaIa.FullName), workbookNdaIa.FullName, "nda ia");
                workbookNda.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling the newly created REIT file with new entry
        /// Create the file if doesn't exist yet
        /// </summary>
        /// <param name="content"></param>
        /// <param name="index"></param>
        private void FillREIT(Dictionary<string, string> content, int index)
        {
            Workbook workbookNda;
            Workbook workbookNdaIa;
            Worksheet worksheetNda;
            Worksheet worksheetNdaIa;
            try
            {
                InitWorkbookAndWorksheet(out workbookNda, out worksheetNda, "REIT");
                InitWorkbookAndWorksheet(out workbookNdaIa, out worksheetNdaIa, "REIT_IA");

                int firstLine = 2 + index * 3;
                int firstLineIa = 2 + index;
                int ricNb = Int32.Parse(content["Official Code:"]);

                // NDA Titles
                if (index == 0)
                {
                    FillTitles(ref worksheetNda, ref worksheetNdaIa);
                }

                worksheetNda.Cells[firstLine, 1] = content["Official Code:"] + ".HK";
                worksheetNda.Cells[firstLine + 1, 1] = content["Official Code:"] + "ta.HK";
                worksheetNda.Cells[firstLine + 2, 1] = content["Official Code:"] + "stat.HK";

                worksheetNda.Cells[firstLine, 6] = "1";
                worksheetNda.Cells[firstLine + 1, 6] = "64377";
                worksheetNda.Cells[firstLine + 2, 6] = "60019";

                worksheetNda.Cells[firstLine, 9] = ricNb.ToString();
                worksheetNda.Cells[firstLine + 2, 9] = ricNb.ToString();

                worksheetNda.Cells[firstLine, 10] = content["Lot Size:"];

                worksheetNda.Cells[firstLine, 11] = "HKG:XHKG";
                worksheetNda.Cells[firstLine + 2, 11] = "HKG:XHKG";

                for (int count = 0; count < 3; count++)
                {
                    worksheetNda.Cells[firstLine + count, 2] = content["Displayname:"] + " REIT";
                    worksheetNda.Cells[firstLine + count, 3] = content["Displayname:"];
                    worksheetNda.Cells[firstLine + count, 4] = content["Currency:"];
                    worksheetNda.Cells[firstLine + count, 5] = "HKG";
                    worksheetNda.Cells[firstLine + count, 7] = "EQUITY";
                    worksheetNda.Cells[firstLine + count, 8] = "ECL";
                    worksheetNda.Cells[firstLine + count, 13] = content["Effective Date:"];
                }

                worksheetNdaIa.Cells[firstLineIa, 1] = "EQUITY";
                worksheetNdaIa.Cells[firstLineIa, 2] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 3] = "";
                worksheetNdaIa.Cells[firstLineIa, 4] = "ORD";
                worksheetNdaIa.Cells[firstLineIa, 5] = content["Legal Registered Name:"].Replace(" Limited", "") + " Ord Shs";

                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbookNda.SaveAs(workbookNda.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbookNdaIa.SaveAs(workbookNdaIa.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(workbookNda.FullName), workbookNda.FullName, "nda");
                AddResult(Path.GetFileNameWithoutExtension(workbookNdaIa.FullName), workbookNdaIa.FullName, "nda ia");
                workbookNda.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot generate CSV file :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                CleanExit();
            }
        }

        /// <summary>
        /// Filling titles for newly created NDAs
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="worksheetIa"></param>
        private void FillTitles(ref Worksheet worksheet, ref Worksheet worksheetIa)
        {
            for (int column = 0; column < ndaTitle.Count; column++)
            {
                worksheet.Cells[1, column + 1] = ndaTitle[column];
            }
            for (int column = 0; column < ndaIaTitle.Count; column++)
            {
                worksheetIa.Cells[1, column + 1] = ndaIaTitle[column];
            }
        }

        /// <summary>
        /// Filling titles for newly created name change NDAs
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="worksheetIa"></param>
        private void FillTitlesNameChange(ref Worksheet worksheet, ref Worksheet worksheetIa)
        {
            for (int column = 0; column < ndaTitleChange.Count; column++)
            {
                worksheet.Cells[1, column + 1] = ndaTitleChange[column];
            }
            for (int column = 0; column < ndaIaTitleChange.Count; column++)
            {
                worksheetIa.Cells[1, column + 1] = ndaIaTitleChange[column];
            }

        }

        #endregion
    }

    #endregion
}
