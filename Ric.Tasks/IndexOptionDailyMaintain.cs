using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_IndexOptionDailyMaintainConfig
    {
        [StoreInDB]
        [Category("File Path")]
        [DisplayName("Requirement file")]
        [Description("The full path of 'KSO STRIKE_MACRO_requirements.xls'. E.g.C:\\Korea_Auto\\Option\\KSO STRIKE_MACRO_requirements.xls ")]
        [DefaultValue("C:\\Korea_Auto\\Option\\KSO STRIKE_MACRO_requirements.xls")]
        public string RequirementFile { get; set; }

        [StoreInDB]
        [Category("File Path")]
        [DisplayName("Nda")]
        [DefaultValue("C:\\Korea_Auto\\Option\\NDA\\")]
        [Description("Path for saving generated NDA file \nE.g. C:\\Korea_Auto\\Option\\NDA\\ ")]
        public string NDA { get; set; }
    }

    public class OperateExcel//get data from source
    {
        /* ******************************************************************
         * this class generate datas from excel and convert it to List<List<string>>
         * 
         * 
         ********************************************************************/

        public static List<List<string>> ReadDateFromExcelSheet(int startRow, int startCol, int endRow, int endCol,
            Worksheet worksheet, bool orderByRow)
        {
            if (worksheet == null)
                throw new Exception("can't write to a worksheet which is not existing");
            List<List<string>> dataToStore = new List<List<string>>();
            if (orderByRow)
                for (int i = startRow; i <= endRow; i++)
                {
                    List<string> tempRow = new List<string>();
                    for (int j = startCol; j <= endCol; j++)
                    {
                        tempRow.Add(ExcelUtil.GetRange(i, j, worksheet).Text.ToString());
                    }
                    dataToStore.Add(tempRow);
                }
            else
            {
                for (int j = startCol; j <= endCol; j++)
                {
                    List<string> temprow = new List<string>();
                    {
                        for (int i = startRow; i <= endRow; i++)
                            temprow.Add(ExcelUtil.GetRange(i, j, worksheet).Text.ToString());
                    }
                    dataToStore.Add(temprow);
                }
            }
            return dataToStore;
        }

        public static List<List<string>> ReadDateFromExcelSheet(string RangStr, Worksheet worksheet, bool OrderByRow)
        {
            if (worksheet == null)
                throw new Exception("can't write to a worksheet which is not existing");
            List<List<string>> DataToStore = new List<List<string>>();
            Range temprang = ExcelUtil.GetRange(RangStr, worksheet);
            if (OrderByRow)
                for (int i = 1; i <= temprang.Rows.Count; i++)
                {
                    List<string> temprow = new List<string>();
                    for (int j = 1; j <= temprang.Columns.Count; j++)
                    {
                        temprow.Add(((Range)temprang.Cells[i, j]).Text.ToString());
                    }
                    DataToStore.Add(temprow);
                }
            else
            {
                for (int j = 1; j <= temprang.Columns.Count; j++)
                {
                    List<string> temprow = new List<string>();

                    for (int i = 1; i <= temprang.Rows.Count; i++)
                    {
                        temprow.Add(((Range)temprang.Cells[i, j]).Text.ToString());
                    }
                    DataToStore.Add(temprow);
                }
            }
            return DataToStore;
        }

        public static void WriteToExcelSheet(int row, int col, Worksheet worksheet, List<List<string>> data)
        {
            if (worksheet == null)
                throw new Exception("can't write to a worksheet which is not existing");
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    worksheet.Cells[row + i, col + j] = data[i][j];
                }
            }
        }
        public static void WriteToCSV(string filepath, List<List<string>> data)
        {
            FileStream fs = new FileStream(filepath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            foreach (List<string> list in data)
            {
                for (int j = 0; j < list.Count; j++)
                {
                    if (j != list.Count - 1)
                        sw.Write(list[j] + ",");
                    else
                        sw.Write(list[j]);

                }
                sw.Write("\r\n");
            }
            sw.Close();
            fs.Close();

        }
        public static void WriteToCSV(string filepath, List<List<string>> data, FileMode mode)
        {
            FileStream fs = new FileStream(filepath, mode);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            foreach (List<string> list in data)
            {
                for (int j = 0; j < list.Count; j++)
                {
                    if (j != list.Count - 1)
                        sw.Write(list[j] + ",");
                    else
                        sw.Write(list[j]);

                }
                sw.Write("\r\n");
            }
            sw.Close();
            fs.Close();

        }
        public static void WriteToTXT(string filepath, List<List<string>> data)
        {
            FileStream fs = new FileStream(filepath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    sw.Write(data[i][j]);
                }
                sw.Write("\r\n");
            }
            sw.Close();
            fs.Close();
        }
    }

    public class IndexOptionDailyMaintain : GeneratorBase
    {
        public class UsefulDateFormat : IComparable<UsefulDateFormat>//Data Format can be used in every class
        {
            public string Year { set; get; }
            public string Name { set; get; }
            public double Low { set; get; }
            public double High { set; get; }

            public int CompareTo(UsefulDateFormat y)
            {
                if (Name == null)
                {
                    if (y.Name == null)
                        return 0;
                    return -1;
                }
                if (y.Name == null)
                    return 1;
                int yearCompare = Year.CompareTo(y.Year);
                if (yearCompare == 0)
                {
                    int monthx = (int)(Enum.Parse(typeof(StockOptionMonth), Name.Substring(0, 3)));
                    int monthy = (int)(Enum.Parse(typeof(StockOptionMonth), y.Name.Substring(0, 3)));
                    return monthx.CompareTo(monthy);
                }
                return yearCompare;
            }

        }

        public class GenerateUsefulDateFromList
        {
            /* ******************************************************************
             * this class generate datas which can be usd to write to another excel
             *  data source from List<List<string>> 
             *  convet data to the List<UsefulDateFormat> 
             *  UsefulDateFormat is a class which define DateFormat we need
             * 
            **********************************************************************/
            public string Lastvalue { set; get; }
            public List<UsefulDateFormat> Data { set; get; }


            public GenerateUsefulDateFromList()
            {
                Lastvalue = null;
                Data = null;

            }
            public GenerateUsefulDateFromList(List<List<string>> DataToParse)
            {
                Lastvalue = null;
                Data = null;
                ParseData(DataToParse);
            }
            private void AddValueHighOrLow(List<List<string>> dataToParse, int index)//updata Data
            {
                //to judge value which may hava same month
                bool findFlag = false;
                double value = Convert.ToDouble(dataToParse[0][index]);
                foreach (UsefulDateFormat format in Data.Where(format => format.Name == dataToParse[2][index]))
                {
                    if (format.High <= value)
                    {
                        format.High = value;
                    }
                    if (format.Low >= value)
                    {
                        format.Low = value;
                    }
                    findFlag = true;
                    break;
                }
                if (!findFlag)
                {
                    UsefulDateFormat temp = new UsefulDateFormat
                    {
                        Name = dataToParse[2][index],
                        High = value,
                        Low = value
                    };
                    string year = temp.Name.Substring(3);
                    FormatYear(ref year);
                    temp.Year = year;
                    Data.Add(temp);
                }
            }

            private void FormatYear(ref string year)
            {
                int currentYear = DateTime.Today.Year;
                year = year == currentYear.ToString().Substring(3) ? currentYear.ToString() : (currentYear + 1).ToString();
            }

            /// <summary> 
            /// Parse data from dataToParse and convert it to List<UsefulDateFormat> 
            /// result stored in Data 
            /// </summary>
            /// <param name="dataToParse"></param>
            public void ParseData(List<List<string>> dataToParse)
            {
                Data = new List<UsefulDateFormat>();
                for (int i = 0; i < dataToParse[0].Count; i++)
                {
                    if (string.IsNullOrEmpty(dataToParse[2][i]))
                    {
                        break;
                    }
                    AddValueHighOrLow(dataToParse, i);
                }
            }
        }

        private Hashtable lastTradingDayInDB = new Hashtable();

        private KOREA_IndexOptionDailyMaintainConfig configObj;
        private bool monthQuarter;
        private bool monthOther;
        private Hashtable callPutOption = new Hashtable();
        private Hashtable callOption = new Hashtable();
        private Hashtable yearCode = new Hashtable();
        private Hashtable putOption = new Hashtable();


        private GenerateUsefulDateFromList usefulData;
        private double[] valueWriteToLineH;
        private double[] valueWriteToLineI;
        private List<List<string>> ndaDataWriteToRequirement = new List<List<string>>();
        private ExcelApp excelApp;
        private Workbook workbook;
        public string gatsServer = string.Empty;


        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREA_IndexOptionDailyMaintainConfig;

            if (string.IsNullOrEmpty(configObj.RequirementFile))
            {
                string msg = "RequirementPath in configuration can not be blank!";
                MessageBox.Show(msg, "Error");
                throw new Exception(msg);
            }

            if (!File.Exists(configObj.RequirementFile))
            {
                string msg = string.Format("Can not find file {0}!", configObj.RequirementFile);
                MessageBox.Show(msg, "Error");
                throw new Exception(msg);
            }

            GatsUtil gats = new GatsUtil();
            gatsServer = gats.ServerIp;

            Logger.Log("Initialize - OK!");
            TaskResultList.Add(new TaskResultEntry("LOG", "Log File", Logger.FilePath));
        }

        protected override void Start()
        {
            try
            {
                GenerateBackup();
                CheckIfNeedMonthlyMaintain();//Update monthOther and month369
                GetMapDataFromDB();
                InitializeExcel();
                if (monthOther || monthQuarter)
                {
                    ExecuteMonthlyMaintain();
                }
                else
                {
                    ExecuteDailyMaintain();
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in running Index Option task. " + ex.Message;
                Logger.Log(msg);
                throw new Exception(msg); ;
            }
            finally
            {
                DisposeExcel();
            }

        }

        private void GetMapDataFromDB()
        {
            GetCallPutOptions();
            GetYearCodes();
        }

        private void GetYearCodes()
        {
            int year = DateTime.Today.Year;
            KoreaCodeMapInfo yearCodeInDB = KoreaOptionMapManager.SelectOneYearCode(year);
            if (yearCodeInDB == null)
            {
                string msg = string.Format("Can not find Year Code for {0} in DB!", year);
                throw new Exception(msg);
            }
            yearCode.Add(yearCodeInDB.Name, yearCodeInDB.Code);
            if (DateTime.Today.Month > 6)
            {
                year++;
                yearCodeInDB = KoreaOptionMapManager.SelectOneYearCode(year);
                if (yearCodeInDB == null)
                {
                    string msg = string.Format("Can not find Year Code for {0} in DB!", year);
                    throw new Exception(msg);
                }
                yearCode.Add(yearCodeInDB.Name, yearCodeInDB.Code);
            }
        }

        private void GetCallPutOptions()
        {
            List<KoreaCodeMapInfo> codeMap = KoreaOptionMapManager.SelectCodeMapByType(KoreaCodeMapType.Call);
            foreach (KoreaCodeMapInfo codeItem in codeMap)
            {
                callPutOption.Add(codeItem.Code, codeItem.Name);
                callOption.Add(codeItem.Name, codeItem.Code);
            }
            codeMap = KoreaOptionMapManager.SelectCodeMapByType(KoreaCodeMapType.Put);
            foreach (KoreaCodeMapInfo codeItem in codeMap)
            {
                callPutOption.Add(codeItem.Code, codeItem.Name);
                putOption.Add(codeItem.Name, codeItem.Code);
            }
            Logger.Log("Get Call/Put Option from DB - OK!");
        }

        /// <summary>
        /// Initialize excel app.
        /// </summary>
        private void InitializeExcel()
        {
            excelApp = new ExcelApp(false, false);
            if (excelApp == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            workbook = ExcelUtil.CreateOrOpenExcelFile(excelApp, configObj.RequirementFile);
            if (workbook == null)
            {
                string msg = string.Format("Error found when openning file {0}!", configObj.RequirementFile);
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        /// <summary>
        /// Dispose excel app.
        /// </summary>
        private void DisposeExcel()
        {
            if (excelApp == null)
            {
                return;
            }
            excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
            excelApp.ExcelAppInstance.DisplayAlerts = false;
            if (workbook != null)
            {
                workbook.Save();
                workbook.Close(true, Type.Missing, Type.Missing);
            }
            excelApp.Dispose();
        }

        /// <summary>
        /// Backup requirement file into folder named by today.
        /// </summary>
        private void GenerateBackup()
        {
            string sourceFile = configObj.RequirementFile;
            string today = DateTime.Today.ToString("yyyy-MM-dd");
            string destFilePath = Path.Combine(Path.GetDirectoryName(sourceFile), "Backup");
            destFilePath = Path.Combine(destFilePath, today);

            if (!Directory.Exists(destFilePath))
            {
                Directory.CreateDirectory(destFilePath);
            }

            try
            {
                string destFile = Path.Combine(destFilePath, Path.GetFileName(sourceFile));
                if (!File.Exists(destFile))
                {
                    File.Copy(sourceFile, destFile, true);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in generate backup file.\r\n" + ex.Message, Logger.LogType.Error);
            }
            Logger.Log(string.Format("Backup - OK!  Backup file: {0} to folder:{1}.", Path.GetFileName(sourceFile), destFilePath));
            AddResult("Backup folder", destFilePath, "folder");
            AddResult(Path.GetFileNameWithoutExtension(configObj.RequirementFile), configObj.RequirementFile, "file");
            //TaskResultList.Add(new TaskResultEntry("Backup Folder", "Backup Folder", destFilePath));
            //TaskResultList.Add(new TaskResultEntry(Path.GetFileName(configObj.RequirementFile), Path.GetFileNameWithoutExtension(configObj.RequirementFile), configObj.RequirementFile));
        }

        //Update monthOther and month369
        private void CheckIfNeedMonthlyMaintain()
        {
            string nowDateTime = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            int nowMonthNumber = DateTime.Today.Month;
            if (KoreaOptionMapManager.CheckLastTradingDay(nowDateTime))
            {
                if (nowMonthNumber == 3 || nowMonthNumber == 6 || nowMonthNumber == 9 || nowMonthNumber == 12)
                    monthQuarter = true;
                if (!monthQuarter)
                    monthOther = true;
            }
        }

        private int FindClosestNumIndex(string dataScope, int startIndex, double lastValue, Worksheet worksheet)
        {
            //to find closest num betwen lastvalue and DataScope in sheet KOS5  
            int index = -1;
            double absValue = 99999;
            List<List<string>> highlightValueSource = OperateExcel.ReadDateFromExcelSheet(dataScope, worksheet, false);
            for (int i = 0; i < highlightValueSource[0].Count; i++)
            {
                if (string.IsNullOrEmpty(highlightValueSource[0][i]))
                {
                    continue;
                }
                double tempValue = 0;
                try
                {
                    tempValue = Convert.ToDouble(highlightValueSource[0][i]);
                }
                catch (Exception ex)
                {
                    string err = ex.ToString();
                    break;
                }
                if (Math.Abs(tempValue - lastValue) < absValue)
                {
                    index = i + startIndex;
                    absValue = Math.Abs(tempValue - lastValue);
                }
            }
            return index;
        }
        /// <summary>
        /// this function find the highest or lowest num in a line 
        /// </summary>
        /// <param name="startNum">value true means find the highest</param>
        /// <param name="row">line start row</param>
        /// <param name="col">line start col</param>
        /// <param name="worksheet">worksheet contains line</param>
        /// <returns></returns>
        public int GetNumIndex(bool startNum, int row, int col, Worksheet worksheet)
        {
            /*****************************************************************************
             * to find highest num or lowest num in a line in a sheet if startNum == true
             * find highest num
             * else
             * find lowest num
             *****************************************************************************/
            double num = -1;
            string temp;
            int index = -1;
            if (startNum)
            {
                for (int i = row; ; i++)
                {
                    temp = ((Range)worksheet.Cells[i, col]).Text.ToString();
                    if (string.IsNullOrEmpty(temp))
                        continue;
                    index = i;
                    break;
                }
            }
            else
            {
                for (int i = row; ; i++)
                {
                    temp = ((Range)worksheet.Cells[i, col]).Text.ToString();
                    if (string.IsNullOrEmpty(temp))
                    {
                        if (num == -1)
                            continue;
                        index = i - 1;
                        break;
                    }
                    try
                    {
                        num = Convert.ToDouble(temp);
                    }
                    catch
                    {
                        index = i - 1;
                        break;
                    }
                }
            }
            return index;
        }
        private void GenerateArray(double centernum, int lenth, double step, ref double[] arry)
        {
            arry = new double[lenth * 2 + 1];
            for (int i = lenth * 2; i >= 0; i--)
            {
                arry[lenth * 2 - i] = centernum + (i - lenth) * step;
            }
        }

        /// <summary>
        /// this function is write to a certain sheet in this task 
        /// can't be used in any other place
        /// it can write highlight palce in a certain line and add flag "New" to added number
        /// </summary>
        /// <param name="centernum">highlight num which in center area</param>
        /// <param name="startnum">start num of the line</param>
        /// <param name="endnum">end num of the line</param>
        /// <param name="step">step  in two number </param>
        /// <param name="row">start row</param>
        /// <param name="col">start cow</param>
        /// <param name="highlightnum">the highlight area ' lenth is highlightnum*2+1</param>
        /// <param name="worksheet">worksheet need to be written</param>
        /// <param name="flag">if flag is true  means line H else means line I </param>
        private void WriteHighlightDataByConfirmedStep(double centernum, double startnum, double endnum, double step, int row, int col, int highlightnum, Worksheet worksheet, bool flag)
        {
            double tempStartnum = startnum;
            double tempsEndnum = endnum;
            if (startnum - step * highlightnum < centernum)
            {
                startnum = centernum + step * highlightnum;
            }
            if (endnum + step * highlightnum > centernum)
            {
                endnum = centernum - step * highlightnum;
            }
            for (int i = 0; ; i++)
            {
                worksheet.Cells[row + i, col] = startnum.ToString();
                if (flag)
                {
                    if (((Range)worksheet.Cells[row + i, col - 1]).Text.ToString() == "New")
                    {
                        worksheet.Cells[row + i, col - 1] = "";
                    }
                }
                else
                {
                    if (((Range)worksheet.Cells[row + i, col + 1]).Text.ToString() == "New")
                    {
                        worksheet.Cells[row + i, col + 1] = "";
                    }
                }

                ((Range)worksheet.Cells[row + i, col]).ClearFormats();
                ((Range)worksheet.Cells[row + i, col]).HorizontalAlignment = XlHAlign.xlHAlignCenter;

                if (startnum == centernum)
                {
                    ((Range)worksheet.Cells[row + i, col]).Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                }
                else if (Math.Abs(startnum - centernum) <= step * highlightnum)
                {
                    ((Range)worksheet.Cells[row + i, col]).Interior.Color = ColorTranslator.ToOle(Color.DodgerBlue);
                }
                startnum -= step;
                if (startnum < endnum)
                {
                    for (int j = 1; j < 20; j++)
                    {
                        ((Range)worksheet.Cells[row + i + j, col]).ClearFormats();
                        worksheet.Cells[row + i + j, col] = "";
                    }
                    break;
                }
            }
        }

        private void WriteFourMonthDataToRequirement(_Worksheet worksheet)
        {
            /*to write G6----I9*/
            worksheet.Cells[3, 8] = usefulData.Lastvalue;
            worksheet.Cells[5, 8] = "Lowest";
            worksheet.Cells[5, 9] = "Highest";
            for (int i = 0; i < usefulData.Data.Count; i++)
            {
                worksheet.Cells[6 + i, 7] = usefulData.Data[i].Name.Substring(0, usefulData.Data[i].Name.Length - 1);
                worksheet.Cells[6 + i, 8] = usefulData.Data[i].Low;
                worksheet.Cells[6 + i, 9] = usefulData.Data[i].High;
            }
        }

        /// <summary>
        /// Change month number to upper English words. e.g. 08 -> AUG
        /// </summary>
        /// <param name="month">month number</param>
        /// <returns>month three chars</returns>
        private string FormatMonthToWords(int month)
        {
            if (monthQuarter)
            {
                month += 6;
            }
            else if (monthOther)
            {
                month += 3;
            }
            month = (month - 1) % 12 + 1;
            string monthtoadd = DateTime.Parse("2013-" + month + "-22").ToString("MMM", new CultureInfo("en-US")).ToUpper();
            return monthtoadd;
        }

        private UsefulDateFormat GenerateNewMonthUsefulDateTerm()
        {
            UsefulDateFormat term = new UsefulDateFormat();
            string yearSuffix = string.Empty;
            int year = Convert.ToInt16(usefulData.Data[0].Name.Substring(usefulData.Data[0].Name.Length - 1, 1));
            int nowMonthNumber = DateTime.Now.Month;

            if (nowMonthNumber > 9 && nowMonthNumber < 12)
            {
                yearSuffix = (year + 1).ToString();
            }
            else
            {
                yearSuffix = year.ToString();
            }
            if (nowMonthNumber == 9 && monthQuarter)
            {
                yearSuffix = (year + 1).ToString();
            }
            if (yearSuffix.Length > 1)
            {
                yearSuffix = yearSuffix.Substring(yearSuffix.Length - 1);
            }
            term.Year = FormatYear(yearSuffix);
            term.Name = FormatMonthToWords(nowMonthNumber) + yearSuffix;
            term.High = 0;
            term.Low = 0;
            return term;
        }

        /// <summary>
        /// Write data to KSO requirement file.
        /// </summary>
        private void WriteDataToRequirement()
        {
            Worksheet worksheet = ExcelUtil.GetWorksheet("KSO5", workbook);
            if (worksheet == null)
            {
                throw new Exception(string.Format("Cannot find worksheet {0} in file {1}.", "KSO5", configObj.RequirementFile));
            }
            try
            {
                // write data to KSO_STIKE_MACRO requirement.xls which needs highlight
                //rewrite the H24-end in sheet KOS5
                double highestnum;
                double centernum;
                double lowestnum;
                int index;
                string dataScope = "H24:H" + worksheet.UsedRange.Rows.Count;
                index = FindClosestNumIndex(dataScope, 24, Convert.ToDouble(usefulData.Lastvalue), worksheet);
                centernum = Convert.ToDouble(ExcelUtil.GetRange("H" + index, worksheet).Text.ToString());
                index = GetNumIndex(true, 24, 8, worksheet);
                highestnum = Convert.ToDouble(ExcelUtil.GetRange("H" + index, worksheet).Text.ToString());
                index = GetNumIndex(false, 24, 8, worksheet);
                lowestnum = Convert.ToDouble(ExcelUtil.GetRange("H" + index, worksheet).Text.ToString());
                GenerateArray(centernum, 6, 2.5, ref valueWriteToLineH);//generate arry contains Next Day’s Price?
                WriteHighlightDataByConfirmedStep(centernum, highestnum, lowestnum, 2.5, 25, 8, 6, worksheet, true);
                //rewrite the I24-end in sheet KOS5
                dataScope = "I24:I" + worksheet.UsedRange.Rows.Count;
                index = FindClosestNumIndex(dataScope, 24, Convert.ToDouble(usefulData.Lastvalue), worksheet);
                centernum = Convert.ToDouble(ExcelUtil.GetRange("I" + index, worksheet).Text.ToString());
                index = GetNumIndex(true, 24, 9, worksheet);
                highestnum = Convert.ToDouble(ExcelUtil.GetRange("I" + index, worksheet).Text.ToString());
                index = GetNumIndex(false, 24, 9, worksheet);
                lowestnum = Convert.ToDouble(ExcelUtil.GetRange("I" + index, worksheet).Text.ToString());
                GenerateArray(centernum, 3, 5, ref valueWriteToLineI);//generate arry contains Next Day’s Price?
                WriteHighlightDataByConfirmedStep(centernum, highestnum, lowestnum, 5, 25, 9, 3, worksheet, false);

                WriteFourMonthDataToRequirement(worksheet);

            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }
        }

        /// <summary>
        /// Generate NDA file.
        /// </summary>
        private void GenerateNDAFile()
        {
            configObj.NDA = Path.Combine(configObj.NDA, DateTime.Today.ToString("yyyy-MM-dd"));
            if (!Directory.Exists(configObj.NDA))
            {
                Directory.CreateDirectory(configObj.NDA);
            }
            string filePath = Path.Combine(configObj.NDA, "KR" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + "_chg_iss.csv");
            try
            {
                ndaDataWriteToRequirement = FillData();
                if (ndaDataWriteToRequirement.Count > 0)
                {
                    List<string> header = new List<string>
                    { "RIC", "TAG", "CATEGORY", "TYPE", "CURRENCY","EXCHANGE", 
                                                        "TRADING SEGMENT ", "DERIVATIVES LAST TRADING DAY", "EXPIRY DATE", "RETIRE DATE", "CALL PUT OPTION", 
                                                        "DERIVATIVES LOT SIZE", "DERIVATIVES TICK VALUE", "DERIVATIVES LOT UNIT", "OPTION STUB", "DERIVATIVES QUOTE UNDERLYING ASSET", 
                                                        "DERIVATIVES METHOD OF DELIVERY", "OPTIONS EXERCISE STYLE", "RCS ASSET CLASS", "DERIVATIVES TRADING STYLE", 
                                                        "DERIVATIVES CONTRACT TYPE", "DERIVATIVES PERIODICITY", "DERIVATIVES SERIES DESCRIPTION", "TICKER SYMBOL", 
                                                        "TRADING SYMBOL", "STRIKE PRICE", "ASSET SHORT NAME", "ASSET COMMON NAME", };
                    List<List<string>> dataWriteToCsv = GenerateUltimateRes(ndaDataWriteToRequirement);
                    FileUtil.WriteOutputFile(filePath, dataWriteToCsv, header, WriteMode.Overwrite);
                    Logger.Log("Generate NDA file - OK!");
                    AddResult("Nda bulk file", filePath, "nda");
                    //TaskResultList.Add(new TaskResultEntry(Path.GetFileNameWithoutExtension(filePath), "NDA File", filePath, FileProcessType.NDA));
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found when generating NDA file.\n" + ex.Message);
            }
        }


        /// <summary>   
        /// daily: add the out of range values in col H and col I
        /// monthly : case 2 , actually it is the same logic as daily
        /// monthly : case 1 , the past quarterly month is different. Add all the .5 values in the range of lowest to highest for the past quarterly month
        /// the output values need to be in order desc.
        /// </summary>           
        /// <returns>the values need to be written to CSV. values in H+I</returns>
        private List<List<string>> FillData()
        {
            List<List<string>> filtres = new List<List<string>>();
            List<UsefulDateFormat> fourMonthData = usefulData.Data;
            bool rangeDone = false;  // flag that if all 2.5 values in the Range have been added(case1)
            bool isLower = false;
            List<string> temp;
            double[] valueToFiltColH = valueWriteToLineH;
            double[] valueToFiltColI = valueWriteToLineI;
            int i, j = 0;
            for (j = 0; j < 3; j++)
            {
                if (j == 2 && monthQuarter && !rangeDone)  // case 1: all the 13 values are lower than the Lowest of past quarterly month
                {
                    if (valueToFiltColH[0] <= fourMonthData[j].Low)
                    {
                        double addNum = 0.0;   //the number need to add, from Highest to Lowest, the step is 5

                        if (fourMonthData[j].High > (int)fourMonthData[j].High)
                            addNum = fourMonthData[j].High;
                        else
                            addNum = fourMonthData[j].High - 2.5;
                        while (!isLower)   //when add_num is lower than Lowest, the while can be over
                        {
                            if (addNum >= fourMonthData[j].Low)
                            {
                                temp = new List<string>();
                                temp.Add(fourMonthData[j].Name);
                                temp.Add(fourMonthData[j].High.ToString());
                                temp.Add(fourMonthData[j].Low.ToString());
                                temp.Add(((int)addNum).ToString());
                                temp.Add(".5");
                                filtres.Add(temp);
                                addNum = addNum - 5;
                            }
                            else
                            {
                                isLower = true;
                            }
                        }
                        rangeDone = true;
                    }
                }

                for (i = 0; i < valueToFiltColH.Length; i++)  //all situation should check if the 13 values are in Range
                {
                    if (valueToFiltColH[i] < fourMonthData[j].Low || valueToFiltColH[i] > fourMonthData[j].High)
                    {
                        temp = new List<string>
                        {
                            fourMonthData[j].Name,
                            fourMonthData[j].High.ToString(),
                            fourMonthData[j].Low.ToString(),
                            ((int) valueToFiltColH[i]).ToString(),
                            valueToFiltColH[i] > (int) valueToFiltColH[i] ? ".5" : ""
                        };
                        filtres.Add(temp);
                    }
                    else // the value is in the Range, for daily and case2, no steps. for case1, all the 2.5 values in the H_L range need to be added
                    {
                        if (j == 2 && monthQuarter && !rangeDone)
                        {
                            double addNum = 0.0; //the number need to add, from Highest to Lowest, the step is 5
                            if (fourMonthData[j].High > (int)fourMonthData[j].High)
                                addNum = fourMonthData[j].High;
                            else
                                addNum = fourMonthData[j].High - 2.5;
                            while (!isLower)
                            {
                                if (addNum >= fourMonthData[j].Low)
                                {
                                    temp = new List<string>
                                    {
                                        fourMonthData[j].Name,
                                        fourMonthData[j].High.ToString(),
                                        fourMonthData[j].Low.ToString(),
                                        ((int) addNum).ToString(),
                                        ".5"
                                    };
                                    filtres.Add(temp);
                                    addNum = addNum - 5;
                                }
                                else
                                {
                                    isLower = true;
                                }
                            }
                            rangeDone = true;
                        }
                    }
                }
                if (j == 2 && monthQuarter && !rangeDone)  // case 1: all the 13 values are higher than the Lowest of past quarterly month
                {
                    if (valueToFiltColH[valueToFiltColH.Length - 1] >= fourMonthData[j].High)
                    {
                        double addNum = 0.0; //the number need to add, from Highest to Lowest, the step is 5
                        if (fourMonthData[j].High > (int)fourMonthData[j].High)
                            addNum = fourMonthData[j].High;
                        else
                            addNum = fourMonthData[j].High - 2.5;
                        while (!isLower)
                        {
                            if (addNum >= fourMonthData[j].Low)
                            {
                                temp = new List<string>
                                {
                                    fourMonthData[j].Name,
                                    fourMonthData[j].High.ToString(),
                                    fourMonthData[j].Low.ToString(),
                                    ((int) addNum).ToString(),
                                    ".5"
                                };
                                filtres.Add(temp);
                                addNum = addNum - 5;
                            }
                            else
                            {
                                isLower = true;
                            }
                        }
                        rangeDone = true;
                    }
                }
            }

            foreach (double val in valueToFiltColI)
            {
                j = 3;
                if (val < fourMonthData[j].Low || val > fourMonthData[j].High) // values out of Range need to be added
                {
                    temp = new List<string>
                    {
                        fourMonthData[j].Name,
                        fourMonthData[j].High.ToString(),
                        fourMonthData[j].Low.ToString(),
                        ((int) val).ToString(),
                        val > (int) val ? ".5" : ""
                    };
                    filtres.Add(temp);
                }
            }
            return filtres;
        }

        private List<List<string>> GenerateUltimateRes(List<List<string>> filtRes)
        {
            List<List<string>> DateToWrite = new List<List<string>>();
            foreach (List<string> res in filtRes)
            {
                List<string> temp = GenerateOneUltimateRes(res, 1);
                DateToWrite.Add(temp);
                temp = GenerateOneUltimateRes(res, 2);
                DateToWrite.Add(temp);
            }
            return DateToWrite;
        }

        private List<string> GenerateOneUltimateRes(List<string> oneFiltRes, int callOrPut)
        {

            string yearSuffix = oneFiltRes[0].Substring(oneFiltRes[0].Length - 1, 1);
            string yearReal = FormatYear(yearSuffix);

            List<string> OneUltimateRes = new List<string>();
            string C_P = "";
            C_P = callOrPut == 1 ? "C" : "P";

            #region A-G
            string month = oneFiltRes[0].Substring(0, oneFiltRes[0].Length - 1);
            string LastnumberOfYear = oneFiltRes[0].Substring(oneFiltRes[0].Length - 1, 1);
            string Ric = "KS200" + oneFiltRes[3] + FindRicCallOrPut(month, callOrPut) + LastnumberOfYear + ".KS";
            OneUltimateRes.Add(Ric);//A
            OneUltimateRes.Add("7130");//B
            OneUltimateRes.Add("EIO");//C
            OneUltimateRes.Add("DERIVATIVE");//D fixed
            OneUltimateRes.Add("KRW");//E
            OneUltimateRes.Add("KSC");//F
            OneUltimateRes.Add("KSC:XKRX");//G
            #endregion
            #region H-J need to confirm
            //string LastTrading = FindTradingDataOrRetireDate(month, 1);
            string LastTrading = ((KoreaOptionLastTradingDayInfo)lastTradingDayInDB[month]).LastTradingDay.ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            string RetireDay = Convert.ToDateTime(LastTrading).AddDays(4).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            OneUltimateRes.Add(LastTrading);//H
            OneUltimateRes.Add(LastTrading);//I
            OneUltimateRes.Add(RetireDay);//J

            #endregion
            #region K

            OneUltimateRes.Add(callOrPut == 1 ? "CALL" : "PUT");

            #endregion
            #region L-W
            OneUltimateRes.Add("500000");//L     fixed
            OneUltimateRes.Add("5000");//M       fixed
            OneUltimateRes.Add("INDEX");//N
            OneUltimateRes.Add("KS200.KS");//O
            OneUltimateRes.Add(".KS200");//P
            OneUltimateRes.Add("CASH");//Q
            OneUltimateRes.Add("E");//R
            OneUltimateRes.Add("OPT");//S
            OneUltimateRes.Add("E");//T
            OneUltimateRes.Add("12");//U
            OneUltimateRes.Add("M");//V
            OneUltimateRes.Add("KSE Korea Exchange Korea Se KOSPI 200 Electronic Day Session Index option");//W
            #endregion
            #region X need to confirm 数据不一致
            if (callOrPut == 1)
                OneUltimateRes.Add("201");//X
            else
                OneUltimateRes.Add("301");//X
            #endregion
            #region Y
            string TRADING_SYMBOL = null;
            string contractMonthNumber = ((KoreaOptionLastTradingDayInfo)lastTradingDayInDB[month]).ContractMonthNumber;

            if (callOrPut == 1)
                TRADING_SYMBOL = "201" + FindYearCode(yearReal) + contractMonthNumber + oneFiltRes[3];
            else
                TRADING_SYMBOL = "301" + FindYearCode(yearReal) + contractMonthNumber + oneFiltRes[3];
            OneUltimateRes.Add(TRADING_SYMBOL);//Y
            #endregion
            #region Z
            OneUltimateRes.Add(oneFiltRes[3] + oneFiltRes[4]);//Z 
            #endregion
            string expiryDate = DateTime.Parse(LastTrading).ToString("MMMy", new CultureInfo("en-US"));
            expiryDate = expiryDate.Substring(0, 3).ToUpper() + expiryDate.Substring(expiryDate.Length - 1, 1);

            #region AA

            OneUltimateRes.Add("KS200" + " " + expiryDate + " " + oneFiltRes[3] + " " + C_P);//AA 
            #endregion
            #region AB
            OneUltimateRes.Add("KS200" + " " + expiryDate + " " + oneFiltRes[3] + " " + C_P);//AB
            #endregion
            return OneUltimateRes;
        }

        private string FindRicCallOrPut(string month, int flag)
        {
            string monthCode = string.Empty;
            if (flag == 1)
            {
                if (callOption.Contains(month))
                {
                    monthCode = callOption[month].ToString().Trim();
                }
            }
            else
            {
                if (putOption.Contains(month))
                {
                    monthCode = putOption[month].ToString().Trim();
                }
            }
            return monthCode;

        }

        private string FindYearCode(string year)
        {
            return yearCode.Contains(year) ? yearCode[year].ToString().Trim() : null;
        }

        private string FormatYear(string yearSuffix)
        {
            int currentYear = DateTime.Today.Year;
            string currentYearStr = currentYear.ToString().Trim();
            string currentYearSuffix = currentYearStr.Substring(currentYearStr.Length - 1);
            if (!currentYearSuffix.Equals(yearSuffix))
            {
                currentYear++;
            }
            return currentYear.ToString();
        }

        private List<List<string>> FormatCsvAddedData(List<List<string>> csvAddedData)
        {
            List<List<string>> tempRes = new List<List<string>>();
            for (int i = 0; i < 4; i++)
            {
                tempRes.Add(new List<string>());
            }

            foreach (List<string> data in csvAddedData)
            {
                for (int j = 0; j < 4; j++)
                {
                    if (data[0].Equals(usefulData.Data[j].Name))
                    {
                        tempRes[j].Add(data[3] + data[4]);
                    }
                }
            }
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 13; j++)
                    tempRes[i].Add("");
            }
            return tempRes;
        }

        /// <summary>
        /// Write the added values of four month to KSO requirement file.
        /// </summary>
        /// <param name="csvAddedData">added data</param>
        private void WriteAddedDataToRequirement(List<List<string>> csvAddedData)
        {
            string filePath = configObj.RequirementFile;
            List<List<string>> tempRes = FormatCsvAddedData(csvAddedData);
            var worksheet = ExcelUtil.GetWorksheet("KSO5", workbook);
            if (worksheet == null)
            {
                throw new Exception(string.Format("Cannot find worksheet {0} in file {1}.", "KSO5", filePath));
            }
            try
            {
                OperateExcel.WriteToExcelSheet(6, 10, worksheet, tempRes);
            }
            catch (Exception ex)
            {
                Logger.Log("Error found when writing relative data to requirement file." + ex.Message, Logger.LogType.Error);
            }
            Logger.Log("Write data back to Excel - OK!");
        }


        #region the DailyMaintain task will go on the following step

        private void ExecuteDailyMaintain()
        {
            string msg = string.Format("\r\nOn {0}, we did a daily Index Option maintance.", DateTime.Now.ToString("yyyy-MMM-dd HH:mm:ss"));
            Logger.Log(msg);
            PrepareData();
            GetLastTradingDay();
            WriteDataToRequirement();
            GenerateNDAFile();
            WriteAddedDataToRequirement(ndaDataWriteToRequirement);
        }

        #endregion

        #region the MonthlyMaintain task will go on the following step

        private void ExecuteMonthlyMaintain()
        {

            PrepareData();

            usefulData.Data.RemoveAt(0);

            string msg = string.Empty;
            if (monthQuarter)
            {
                msg = "On {0}, we did a quarterly Index Option maintain.";
                usefulData.Data.Add(GenerateNewMonthUsefulDateTerm());
            }
            else
            {
                msg = "On {0}, we did a monthly Index Option maintain.";
                usefulData.Data.Insert(2, GenerateNewMonthUsefulDateTerm());
            }
            msg = string.Format(msg, DateTime.Now.ToString("yyyy-MMM-dd HH:mm:ss"));
            Logger.Log(msg);

            GetLastTradingDay();

            WriteDataToRequirement();
            GenerateNDAFile();
            WriteAddedDataToRequirement(ndaDataWriteToRequirement);
        }

        private void GetLastTradingDay()
        {
            foreach (var dataItem in usefulData.Data)
            {
                string month = dataItem.Name.Substring(0, 3);
                KoreaOptionLastTradingDayInfo lastTradingDayItem = KoreaOptionMapManager.SelectLastTradingDayByYearMonth(dataItem.Year, month);
                if (lastTradingDayItem == null)
                {
                    string msg = string.Format("Can not get last trading day from database. For year; {0}, month: {1}. Please check the Korea Code Map table", dataItem.Year, dataItem.Name.Substring(0, 3));
                    throw new Exception(msg);
                }
                lastTradingDayInDB.Add(month, lastTradingDayItem);
            }
        }
        #endregion


        /// <summary>
        /// Get data from GATS.
        /// </summary>
        private void PrepareData()
        {
            string file = "Tools\\Data2XML\\Data2XML.exe";
            string path = "Tools\\Data2XML\\";
            Process gatsProcess = new Process
            {
                StartInfo =
                {
                    FileName = file,
                    WorkingDirectory = path,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };
            string lastPrice = GetLastPrice(gatsProcess);
            List<List<string>> dataToParse = GetLongLinks(gatsProcess);
            if (!string.IsNullOrEmpty(lastPrice) && dataToParse != null)
            {
                usefulData = new GenerateUsefulDateFromList();
                usefulData.ParseData(dataToParse);
                usefulData.Lastvalue = lastPrice;
                usefulData.Data.Sort();
            }
            Logger.Log("Get data from GATS - OK!");
        }

        private List<List<string>> GetLongLinks(Process gatsProcess)
        {
            List<List<string>> data = new List<List<string>>();
            List<string> longlinks = new List<string>();
            List<string> prices = new List<string>();
            List<string> months = new List<string>();

            int count = 20;
            string ricPara = string.Empty;
            string ric = "KS200";
            for (int j = 0; j <= count; j++)
            {
                ricPara += string.Format("{0}#{1}*.KS,", j, ric);
            }
            ricPara = ricPara.Substring(0, ricPara.Length - 1);

            string fidPara = "LONGLINK2,LONGLINK4,LONGLINK6,LONGLINK8,LONGLINK10,LONGLINK12,LONGLINK14,LONGNEXTLR";
            string arg = string.Format("/c-quiet -dbout -raw_enum_vals  -ph {2} -pn IDN_SELECTFEED -rics \"{0}\" -fids \"{1}\"", ricPara, fidPara, gatsServer);
            //string result = GetGatsResponse(gatsProcess, arg);

            GatsUtil gats = new GatsUtil();
            string result = gats.GetGatsResponse(ricPara, fidPara);
            if (string.IsNullOrEmpty(result))
            {
                throw new Exception("Can not get response from GATS!");
            }

            result += GetExtraPrices(count, result, gatsProcess);

            string pattern = @"LONGLINK\d+ +(?<LongLink>.*?)\r\n";
            Regex regex = new Regex(pattern);
            MatchCollection matches = regex.Matches(result);
            if (matches.Count > 0)
            {
                for (int i = 0; i < matches.Count; i++)
                {
                    Match m = matches[i];
                    string longlink = m.Groups["LongLink"].Value.Trim();
                    if (string.IsNullOrEmpty(longlink))
                    {
                        continue;
                    }

                    string s = longlink;
                    s = s.Replace("KS200", "").Replace(".KS", "");
                    string month = s.Substring(s.Length - 2);
                    string price = s.Substring(0, s.Length - 2);
                    //string year = s.Substring(s.Length - 1);

                    FormatPrice(ref price);
                    FormatMonth(ref month);

                    if (!string.IsNullOrEmpty(month) && !string.IsNullOrEmpty(price))
                    {
                        longlinks.Add(longlink);
                        prices.Add(price);
                        months.Add(month);

                    }
                }
            }
            data.Add(prices);
            data.Add(longlinks);
            data.Add(months);

            return data;
        }

        private string GetExtraPrices(int count, string result, Process gatsProcess)
        {
            int startNum = count + 1;
            string checkSign = string.Format("LONGNEXTLR +{0}#(?<RIC>[0-9]+)\\*\\.KS", startNum);
            Regex regex = new Regex(checkSign);
            Match match = regex.Match(result);
            if (!match.Success)
            {
                return "";
            }
            int endNum = count + 10;
            string ricPara = string.Empty;
            string ric = "KS200";
            for (int j = startNum; j <= endNum; j++)
            {
                ricPara += string.Format("{0}#{1}*.KS,", j, ric);
            }

            ricPara = ricPara.Substring(0, ricPara.Length - 1);

            string fidPara = "LONGLINK2,LONGLINK4,LONGLINK6,LONGLINK8,LONGLINK10,LONGLINK12,LONGLINK14,LONGNEXTLR";
            string arg = string.Format("/c-quiet -dbout -raw_enum_vals  -ph {2} -pn IDN_SELECTFEED -rics \"{0}\" -fids \"{1}\"", ricPara, fidPara, gatsServer);
            //string gatsresult = GetGatsResponse(gatsProcess, arg);

            GatsUtil gats = new GatsUtil();
            string gatsresult = gats.GetGatsResponse(ricPara, fidPara);
            if (string.IsNullOrEmpty(result))
            {
                return "";
            }

            gatsresult += GetExtraPrices(endNum, gatsresult, gatsProcess);
            return gatsresult;
        }

        private string GetLastPrice(Process gatsProcess)
        {
            string fids = "TRDPRC_1";
            string indexRic = ".KS200";
            string arg = string.Format("/c-quiet -dbout -raw_enum_vals  -ph {0} -pn IDN_SELECTFEED -rics \"{1}\" -fids \"{2}\"", gatsServer, indexRic, fids);
            //string result = GetGatsResponse(gatsProcess, arg);

            GatsUtil gats = new GatsUtil();
            string result = gats.GetGatsResponse(indexRic, fids);
            if (string.IsNullOrEmpty(result))
            {
                throw new Exception("GATS returns no response.");
            }
            string pattern = string.Format(@"{0} +(?<LastPrice>.*?)\r\n", fids);
            Regex regex = new Regex(pattern);
            Match match = regex.Match(result);
            if (match.Success)
            {
                return match.Groups["LastPrice"].Value.Trim();
            }
            string msg = "Can not get last price from GATS.";
            throw new Exception(msg);
        }

        /// <summary>
        /// Give GATS a command line. Get the response.
        /// </summary>
        /// <param name="gatsProcess">GATS process</param>
        /// <param name="command">command</param>
        /// <returns>response</returns>
        private string GetGatsResponse(Process gatsProcess, string command)
        {
            if (gatsProcess == null)
            {
                return null;
            }
            try
            {
                int timeout = 20;
                string filePath = GetOutputFilePath();
                string fileName = "GATS_Response.txt";
                filePath = Path.Combine(filePath, fileName);

                command += string.Format(" -tee \"{0}\"", filePath);

                int retry = 3;
                bool success = false;
                while (!success && retry-- > 0)
                {
                    gatsProcess.StartInfo.Arguments = command;
                    gatsProcess.Start();
                    success = gatsProcess.WaitForExit(timeout * 1000);
                    if (!success)
                    {
                        gatsProcess.Kill();
                    }
                }
                string response = null;
                if (success)
                {
                    response = File.ReadAllText(filePath);
                }
                return response;
            }
            catch
            {
                return null;
            }
        }

        private void FormatMonth(ref string month)
        {
            if (string.IsNullOrEmpty(month) || month.Length != 2)
            {
                month = null;
                return;
            }

            string monthCode = month.Substring(0, 1);

            if (!callPutOption.Contains(monthCode))
            {
                month = null;
                return;
            }
            month = callPutOption[monthCode] + month.Substring(1);
        }

        private void FormatPrice(ref string price)
        {
            if (string.IsNullOrEmpty(price))
            {
                price = null;
                return;
            }
            string last = price.Substring(price.Length - 1);
            if (last.Equals("2") || last.Equals("7"))
            {
                price += ".5";
            }
        }
    }
}
