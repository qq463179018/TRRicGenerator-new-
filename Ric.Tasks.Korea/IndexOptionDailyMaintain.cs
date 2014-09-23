using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Globalization;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Collections;
using Ric.Db.Manager;
using Ric.Db.Info;
using System.Diagnostics;
using System.Reflection;
using Ric.Core;
using Ric.Util;


namespace Ric.Tasks.Korea
{
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_IndexOptionDailyMaintainConfig
    {
        [StoreInDB]
        [Category("File Path")]
        [Description("The full path of 'KSO STRIKE_MACRO_requirements.xls'. E.g.C:\\Korea_Auto\\Option\\KSO STRIKE_MACRO_requirements.xls ")]
        [DefaultValue("C:\\Korea_Auto\\Option\\KSO STRIKE_MACRO_requirements.xls")]
        public string RequirementFile { get; set; }

        [StoreInDB]
        [Category("File Path")]
        [DefaultValue("C:\\Korea_Auto\\Option\\NDA\\")]
        [Description("Path for saving generated NDA file \nE.g. C:\\Korea_Auto\\Option\\NDA\\ ")]
        public string NDA { get; set; }

        [StoreInDB]
        [Category("File Path")]
        [DefaultValue("C:\\Korea_Auto\\Option\\IDN\\")]
        [Description("Path for saving generated IDN file \nE.g. C:\\Korea_Auto\\Option\\IDN\\ ")]
        public string IDN { get; set; }
    }

    public class OperateExcel//get data from source
    {
        /* ******************************************************************
         * this class generate datas from excel and convert it to List<List<string>>
         * 
         * 
         ********************************************************************/
        public static List<List<string>> ReadDateFromExcelSheet(int startRow, int startCol, int endRow, int endCol, Worksheet worksheet, bool orderByRow)
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
                    worksheet.Cells[row + i, col + j] = data[i][j].ToString();
                }
            }
        }
        public static void WriteToCSV(string filepath, List<List<string>> data)
        {
            FileStream fs = new FileStream(filepath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    if (j != data[i].Count - 1)
                        sw.Write(data[i][j] + ",");
                    else
                        sw.Write(data[i][j]);

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
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    if (j != data[i].Count - 1)
                        sw.Write(data[i][j] + ",");
                    else
                        sw.Write(data[i][j]);

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
                if (this.Name == null)
                {
                    if (y.Name == null)
                        return 0;
                    else
                        return -1;
                }
                else
                {
                    if (y.Name == null)
                        return 1;
                    int yearCompare = this.Year.CompareTo(y.Year);
                    if (yearCompare == 0)
                    {
                        int monthx = (int)(Enum.Parse(typeof(StockOptionMonth), this.Name.Substring(0, 3)));
                        int monthy = (int)(Enum.Parse(typeof(StockOptionMonth), y.Name.Substring(0, 3)));
                        return monthx.CompareTo(monthy);
                    }
                    else
                    {
                        return yearCompare;
                    }
                }
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
                for (int i = 0; i < Data.Count; i++)
                {
                    if (Data[i].Name == dataToParse[2][index])
                    {
                        if (Data[i].High <= value)
                        {
                            Data[i].High = value;
                        }
                        if (Data[i].Low >= value)
                        {
                            Data[i].Low = value;
                        }
                        findFlag = true;
                        break;
                    }
                }
                if (!findFlag)
                {
                    UsefulDateFormat temp = new UsefulDateFormat();
                    temp.Name = dataToParse[2][index];
                    temp.High = value;
                    temp.Low = value;
                    string year = temp.Name.Substring(3);
                    FormatYear(ref year);
                    temp.Year = year;
                    Data.Add(temp);
                }
            }

            private void FormatYear(ref string year)
            {
                year = string.Format("{0}{1}", DateTime.Now.Year.ToString().Substring(0, 3), year);
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
                    else
                    {
                        AddValueHighOrLow(dataToParse, i);
                    }
                }
            }
        }

        private Hashtable lastTradingDayInDB = new Hashtable();
        private Hashtable lastTradingDayInDBNew = new Hashtable();

        private KOREA_IndexOptionDailyMaintainConfig configObj = null;
        private bool monthQuarter = false;
        private bool monthOther = false;

        private bool quarterly = false;
        private bool monthly = false;
        private bool semiAnnualSix = false;
        private bool semiAnnualTwelve = false;

        private Hashtable callPutOption = new Hashtable();
        private Hashtable callOption = new Hashtable();
        private Hashtable yearCode = new Hashtable();
        private Hashtable putOption = new Hashtable();


        private GenerateUsefulDateFromList usefulData = null;
        private double[] valueWriteToLineG = null;
        private double[] valueWriteToLineH = null;
        private double[] valueWriteToLineI = null;
        private List<List<string>> ndaDataWriteToRequirement = new List<List<string>>();
        private List<List<string>> idnDataWriteToRequirement = new List<List<string>>();
        private ExcelApp excelApp = null;
        private Workbook workbook = null;
        public string gatsServer = string.Empty;

        private Dictionary<string, List<double>> monthPriceIDN = new Dictionary<string, List<double>>();

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREA_IndexOptionDailyMaintainConfig;
            //gatsServer = ConfigureOperator.GetGatsServer();

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

            //GatsUtil gats = new GatsUtil();
            //gatsServer = gats.ServerIp;

            Logger.Log("Initialize - OK!");
            AddResult("LOG", Logger.FilePath, "Log File");
        }

        protected override void Start()
        {
            //  COMAddIn addIn = null;
            try
            {
                GenerateBackup();
                CheckIfNeedMonthlyMaintain();//Update monthOther and month369
                GetMapDataFromDB();
                InitializeExcel();

                //if (monthOther || monthQuarter)
                //{
                //    ExecuteMonthlyMaintain();
                //}
                //else
                //{
                //    ExecuteDailyMaintain();
                //}

                if (quarterly || semiAnnualSix || semiAnnualTwelve || monthly)
                    ExecuteMonthlyMaintain();
                else
                    ExecuteDailyMaintain();
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
            //GetLastTradingDays();
            GetCallPutOptions();
            //GetYearCodes();
            GetYearCodesNew();
        }

        private void GetYearCodes()
        {
            int year = DateTime.Today.Year;
            KoreaCodeMapInfo yearCodeInDB = KoreaOptionMapNewManager.SelectOneYearCode(year);
            if (yearCodeInDB == null)
            {
                string msg = string.Format("Can not find Year Code for {0} in DB!", year);
                throw new Exception(msg);
            }
            yearCode.Add(yearCodeInDB.Name, yearCodeInDB.Code);
            if (DateTime.Today.Month > 6)
            {
                year++;
                yearCodeInDB = KoreaOptionMapNewManager.SelectOneYearCode(year);
                if (yearCodeInDB == null)
                {
                    string msg = string.Format("Can not find Year Code for {0} in DB!", year);
                    throw new Exception(msg);
                }
                yearCode.Add(yearCodeInDB.Name, yearCodeInDB.Code);
            }
        }

        private void GetYearCodesNew()
        {
            int year = DateTime.Today.Year;
            KoreaCodeMapInfo yearCodeInDB = null;
            while ((yearCodeInDB = KoreaOptionMapNewManager.SelectOneYearCode(year)) != null)
            {
                yearCode.Add(yearCodeInDB.Name, yearCodeInDB.Code);
                year++;
            }
        }

        private void GetCallPutOptions()
        {
            List<KoreaCodeMapInfo> codeMap = KoreaOptionMapNewManager.SelectCodeMapByType(KoreaCodeMapType.Call);
            foreach (KoreaCodeMapInfo codeItem in codeMap)
            {
                callPutOption.Add(codeItem.Code, codeItem.Name);
                callOption.Add(codeItem.Name, codeItem.Code);
            }
            codeMap = KoreaOptionMapNewManager.SelectCodeMapByType(KoreaCodeMapType.Put);
            foreach (KoreaCodeMapInfo codeItem in codeMap)
            {
                callPutOption.Add(codeItem.Code, codeItem.Name);
                putOption.Add(codeItem.Name, codeItem.Code);
            }
            Logger.Log("Get Call/Put Option from DB - OK!");
        }
        /*
        private void GetLastTradingDays()
        {
            int currentYear = DateTime.Today.Year;
            int currentMonth = DateTime.Today.Month;
            int startIndex = 0;
            List<KoreaOptionLastTradingDayInfo> lastTradingDayDB = KoreaOptionMapNewManager.SelectLastTradingDayByYear(currentYear.ToString());
            if (lastTradingDayDB == null)
            {
                string msg = string.Format("Can not get last trading days for year:{0} information from DB.", currentYear);
                throw new Exception(msg);
            }
            if (lastTradingDayDB.Count != 12)
            {
                string msg = string.Format("The last trading days information for year:{0} in DB are incomplete!", currentYear);
                throw new Exception(msg);
            }
            if (currentMonth > 6)
            {
                startIndex = 6;
                List<KoreaOptionLastTradingDayInfo> lastTradingDayOfNextYear = KoreaOptionMapNewManager.SelectLastTradingDayByYear((currentYear + 1).ToString());

                if (lastTradingDayOfNextYear == null)
                {
                    string msg = string.Format("Can not get last trading days for year:{0} information from DB.", currentYear + 1);
                    throw new Exception(msg);
                }
                if (lastTradingDayOfNextYear.Count != 12)
                {
                    string msg = string.Format("The last trading days information for year:{0} in DB are incomplete!", currentYear + 1);
                    throw new Exception(msg);
                }

                lastTradingDayDB.AddRange(lastTradingDayOfNextYear);
            }

            for (int i = startIndex; i < startIndex + 12; i++)
            {
                KoreaOptionLastTradingDayInfo item = lastTradingDayDB[i];
                lastTradingDayInDB.Add(item.Month, item);
            }
            Logger.Log("Get last trading days from DB - OK!");
        }
        */
        private void InitializeExcelApp()
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

        private void DisposeExcelApp()
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
            AddResult("Backup Folder", destFilePath, "Backup Folder");
            AddResult(Path.GetFileName(configObj.RequirementFile), configObj.RequirementFile, Path.GetFileNameWithoutExtension(configObj.RequirementFile));
        }

        //Update monthOther and month369
        private void CheckIfNeedMonthlyMaintain()
        {
            string nowDateTime = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            int nowMonthNumber = DateTime.Today.Month;
            if (KoreaOptionMapNewManager.CheckLastTradingDay(nowDateTime))
            {
                //if (nowMonthNumber == 3 || nowMonthNumber == 6 || nowMonthNumber == 9 || nowMonthNumber == 12)
                //    monthQuarter = true;
                //if (!monthQuarter)
                //    monthOther = true;

                if (nowMonthNumber == 3 || nowMonthNumber == 9) //MAR, SEP
                    quarterly = true;
                else if (nowMonthNumber == 6)                   //JUN
                    semiAnnualSix = true;
                else if (nowMonthNumber == 12)                  //DEC
                    semiAnnualTwelve = true;
                else                                            //JAN, FEB, APR, MAY, JUL, AUG, OCT, NOV
                    monthly = true;
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
                    else
                    {
                        index = i;
                        break;
                    }
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
                        else
                        {
                            index = i - 1;
                            break;
                        }
                    }
                    else
                    {
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
                ((Range)worksheet.Cells[row + i, col]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                if (startnum == centernum)
                {
                    ((Range)worksheet.Cells[row + i, col]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                }
                else if (Math.Abs(startnum - centernum) <= step * highlightnum)
                {
                    ((Range)worksheet.Cells[row + i, col]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.DodgerBlue);
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

        private void WriteHighlightDataByConfirmedStep(double centernum, double startnum, double endnum, double step, int row, int col, int highlightnum, Worksheet worksheet)
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
                ((Range)worksheet.Cells[row + i, col]).ClearFormats();
                ((Range)worksheet.Cells[row + i, col]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                if (startnum == centernum)
                {
                    ((Range)worksheet.Cells[row + i, col]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                }
                else if (Math.Abs(startnum - centernum) <= step * highlightnum)
                {
                    ((Range)worksheet.Cells[row + i, col]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.DodgerBlue);
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
        //private void WriteFourMonthDataToRequirement(Worksheet worksheet)
        //{
        //    /*to write G6----I9*/
        //    worksheet.Cells[3, 8] = usefulData.Lastvalue;
        //    worksheet.Cells[5, 8] = "Lowest";
        //    worksheet.Cells[5, 9] = "Highest";
        //    for (int i = 0; i < usefulData.Data.Count; i++)
        //    {
        //        worksheet.Cells[6 + i, 7] = usefulData.Data[i].Name.Substring(0, usefulData.Data[i].Name.Length - 1);
        //        worksheet.Cells[6 + i, 8] = usefulData.Data[i].Low;
        //        worksheet.Cells[6 + i, 9] = usefulData.Data[i].High;
        //    }
        //}

        private void WriteElevenMonthDataToRequirement(Worksheet worksheet)
        {
            /*to write G6----I16*/
            worksheet.Cells[3, 8] = usefulData.Lastvalue;
            worksheet.Cells[5, 8] = "Lowest";
            worksheet.Cells[5, 9] = "Highest";
            for (int i = 0; i < usefulData.Data.Count; i++)
            {
                ((Range)worksheet.Cells[6 + i, 7]).NumberFormatLocal = "@";
                ((Range)worksheet.Cells[6 + i, 7]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[6 + i, 7] = usefulData.Data[i].Name;

                ((Range)worksheet.Cells[6 + i, 8]).NumberFormatLocal = "@";
                ((Range)worksheet.Cells[6 + i, 8]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[6 + i, 8] = usefulData.Data[i].Low;

                ((Range)worksheet.Cells[6 + i, 9]).NumberFormatLocal = "@";
                ((Range)worksheet.Cells[6 + i, 9]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
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
        /// return current month +6 month
        /// </summary>
        /// <returns></returns>
        private UsefulDateFormat GenerateNewMonthUsefulDateTermNew(DateTime dt)
        {
            UsefulDateFormat term = new UsefulDateFormat();
            term.Year = dt.Year.ToString();
            term.Name = FormateNameValue(dt.ToString("dd-MMM-yyyy"));
            term.High = 0.0;
            term.Low = 0.0;
            //term.High = 275.0;
            //term.Low = 220.0;
            return term;
        }

        private string FormateNameValue(string p)
        {
            return string.Format("{0}{1}", p.Substring(3, 3).ToUpper(), p.Substring(p.Length - 1, 1));
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
                string dataScope;

                //rewrite the G24:G-end in sheet KOS5
                dataScope = "G24:G" + worksheet.UsedRange.Rows.Count;
                index = FindClosestNumIndex(dataScope, 24, Convert.ToDouble(usefulData.Lastvalue), worksheet);
                centernum = Convert.ToDouble(ExcelUtil.GetRange("G" + index, worksheet).Text.ToString());
                index = GetNumIndex(true, 24, 7, worksheet);
                highestnum = Convert.ToDouble(ExcelUtil.GetRange("G" + index, worksheet).Text.ToString());
                index = GetNumIndex(false, 24, 7, worksheet);
                lowestnum = Convert.ToDouble(ExcelUtil.GetRange("G" + index, worksheet).Text.ToString());
                GenerateArray(centernum, 12, 2.5, ref valueWriteToLineG);//generate arry contains Next Day’s Price?
                WriteHighlightDataByConfirmedStep(centernum, highestnum, lowestnum, 2.5, 25, 7, 12, worksheet);

                //rewrite the H24:H-end in sheet KOS5
                dataScope = "H24:H" + worksheet.UsedRange.Rows.Count;
                index = FindClosestNumIndex(dataScope, 24, Convert.ToDouble(usefulData.Lastvalue), worksheet);
                centernum = Convert.ToDouble(ExcelUtil.GetRange("H" + index, worksheet).Text.ToString());
                index = GetNumIndex(true, 24, 8, worksheet);
                highestnum = Convert.ToDouble(ExcelUtil.GetRange("H" + index, worksheet).Text.ToString());
                index = GetNumIndex(false, 24, 8, worksheet);
                lowestnum = Convert.ToDouble(ExcelUtil.GetRange("H" + index, worksheet).Text.ToString());
                GenerateArray(centernum, 6, 5, ref valueWriteToLineH);//generate arry contains Next Day’s Price?
                WriteHighlightDataByConfirmedStep(centernum, highestnum, lowestnum, 5, 25, 8, 6, worksheet);

                //rewrite the I24:I-end in sheet KOS5
                dataScope = "I24:I" + worksheet.UsedRange.Rows.Count;
                index = FindClosestNumIndex(dataScope, 24, Convert.ToDouble(usefulData.Lastvalue), worksheet);
                centernum = Convert.ToDouble(ExcelUtil.GetRange("I" + index, worksheet).Text.ToString());
                index = GetNumIndex(true, 24, 9, worksheet);
                highestnum = Convert.ToDouble(ExcelUtil.GetRange("I" + index, worksheet).Text.ToString());
                index = GetNumIndex(false, 24, 9, worksheet);
                lowestnum = Convert.ToDouble(ExcelUtil.GetRange("I" + index, worksheet).Text.ToString());
                GenerateArray(centernum, 3, 10, ref valueWriteToLineI);//generate arry contains Next Day’s Price?
                WriteHighlightDataByConfirmedStep(centernum, highestnum, lowestnum, 10, 25, 9, 3, worksheet);

                WriteElevenMonthDataToRequirement(worksheet);

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
            string path = Path.Combine(configObj.NDA, DateTime.Today.ToString("yyyy-MM-dd"));
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filePath = Path.Combine(path, "KR" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + "_chg_iss.csv");
            try
            {
                ndaDataWriteToRequirement = FillDataNew();
                if (ndaDataWriteToRequirement.Count > 0)
                {
                    List<string> header = new List<string>() { "RIC", "TAG", "CATEGORY", "TYPE", "CURRENCY","EXCHANGE", 
                                                        "TRADING SEGMENT ", "DERIVATIVES LAST TRADING DAY", "EXPIRY DATE", "RETIRE DATE", "CALL PUT OPTION", 
                                                        "DERIVATIVES LOT SIZE", "DERIVATIVES TICK VALUE", "DERIVATIVES LOT UNIT", "OPTION STUB", "DERIVATIVES QUOTE UNDERLYING ASSET", 
                                                        "DERIVATIVES METHOD OF DELIVERY", "OPTIONS EXERCISE STYLE", "RCS ASSET CLASS", "DERIVATIVES TRADING STYLE", 
                                                        "DERIVATIVES CONTRACT TYPE", "DERIVATIVES PERIODICITY", "DERIVATIVES SERIES DESCRIPTION", "TICKER SYMBOL", 
                                                        "TRADING SYMBOL", "STRIKE PRICE", "ASSET SHORT NAME", "ASSET COMMON NAME", };
                    List<List<string>> dataWriteToCsv = GenerateUltimateRes(ndaDataWriteToRequirement);
                    FileUtil.WriteOutputFile(filePath, dataWriteToCsv, header, WriteMode.Overwrite);
                    Logger.Log("Generate NDA file - OK!");
                    AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "NDA File");
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found when generating NDA file.\n" + ex.Message);
            }
        }

        private List<List<string>> FillDataNew()
        {
            List<List<string>> filtres = new List<List<string>>();
            List<UsefulDateFormat> elevenMonthData = usefulData.Data;

            for (int i = 0; i < elevenMonthData.Count; i++)
            {
                if (i < 6)
                {
                    AddFiltres(i, valueWriteToLineG, elevenMonthData, filtres);
                }
                else if (i < 8 && i >= 6)
                {
                    AddFiltres(i, valueWriteToLineH, elevenMonthData, filtres);
                }
                else if (i < 11 && i >= 8)
                {
                    AddFiltres(i, valueWriteToLineI, elevenMonthData, filtres);
                }
            }

            return filtres;
        }

        private void AddFiltres(int monthIndex, double[] valueWriteToLine, List<UsefulDateFormat> elevenMonthData, List<List<string>> filtres)
        {
            double highIDN, lowIDN;
            double highXLS, lowXLS;
            highIDN = elevenMonthData[monthIndex].High;
            lowIDN = elevenMonthData[monthIndex].Low;
            highXLS = valueWriteToLine[0];
            lowXLS = valueWriteToLine[valueWriteToLine.Length - 1];

            if (Convert.ToInt32(elevenMonthData[monthIndex].High + elevenMonthData[monthIndex].Low) <= 0)
            {
                for (int j = 0; j < valueWriteToLine.Length; j++)
                {
                    filtres.Add(AddRange(monthIndex, elevenMonthData, j, valueWriteToLine));
                }
            }
            else
            {
                for (int j = 0; j < valueWriteToLine.Length; j++)
                {
                    if (valueWriteToLine[j] < lowIDN || valueWriteToLine[j] > highIDN)
                    {
                        filtres.Add(AddRange(monthIndex, elevenMonthData, j, valueWriteToLine));
                    }
                    else
                    {
                        if (monthPriceIDN.ContainsKey(elevenMonthData[monthIndex].Name) &&
                            !monthPriceIDN[elevenMonthData[monthIndex].Name].Contains(valueWriteToLine[j]))
                            filtres.Add(AddRange(monthIndex, elevenMonthData, j, valueWriteToLine));
                    }
                }
            }
        }

        private List<string> AddRange(int indexMonthDate, List<UsefulDateFormat> monthData, int indexFiltCol, double[] valueToFiltCol)
        {
            List<string> temp = new List<string>();
            temp.Add(monthData[indexMonthDate].Name);
            temp.Add(monthData[indexMonthDate].High.ToString());
            temp.Add(monthData[indexMonthDate].Low.ToString());
            temp.Add(((int)valueToFiltCol[indexFiltCol]).ToString());

            if (valueToFiltCol[indexFiltCol] > (int)valueToFiltCol[indexFiltCol])
                temp.Add(".5");
            else
                temp.Add("");

            return temp;
        }

        private IEnumerable<string> DoubleToString(double[] range)
        {
            List<string> result = new List<string>();
            if (range == null || range.Length <= 0)
                return result;

            foreach (var item in range)
            {
                result.Add(item.ToString());
            }

            return result;
        }

        private List<List<string>> GenerateUltimateRes(List<List<string>> filtRes)
        {
            List<List<string>> DateToWrite = new List<List<string>>();
            for (int i = 0; i < filtRes.Count; i++)
            {
                List<string> temp = GenerateOneUltimateRes(filtRes[i], 1);
                DateToWrite.Add(temp);
                temp = GenerateOneUltimateRes(filtRes[i], 2);
                DateToWrite.Add(temp);
            }
            return DateToWrite;
        }

        private List<string> GenerateOneUltimateRes(List<string> oneFiltRes, int callOrPut)
        {

            string yearSuffix = oneFiltRes[0].Substring(oneFiltRes[0].Length - 1, 1);
            //string Year = "201" + oneFiltRes[0].Substring(oneFiltRes[0].Length - 1, 1);
            string yearReal = FormatYear(yearSuffix);

            List<string> OneUltimateRes = new List<string>();
            string C_P = "";
            if (callOrPut == 1)
                C_P = "C";
            else
            {
                C_P = "P";
            }

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
            //string LastTrading = ((KoreaOptionLastTradingDayInfo)lastTradingDayInDB[month]).LastTradingDay.ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            if (!lastTradingDayInDBNew.ContainsKey(oneFiltRes[0]))
            {
                string msg = string.Format("please add MonthYear:{0} to DB table ETI_Korea_OptionLTDNew", oneFiltRes[0].ToString());
                LogMessage(msg);
            }
            string LastTrading = ((KoreaOptionLastTradingDayInfo)lastTradingDayInDBNew[oneFiltRes[0]]).LastTradingDay.ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            string RetireDay = Convert.ToDateTime(LastTrading).AddDays(4).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            OneUltimateRes.Add(LastTrading);//H
            OneUltimateRes.Add(LastTrading);//I
            OneUltimateRes.Add(RetireDay);//J

            #endregion
            #region K
            if (callOrPut == 1)
                OneUltimateRes.Add("CALL");//K
            else
            {
                OneUltimateRes.Add("PUT");//K
            }
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
            //string contractMonthNumber = ((KoreaOptionLastTradingDayInfo)lastTradingDayInDB[month]).ContractMonthNumber;
            if (!lastTradingDayInDBNew.ContainsKey(oneFiltRes[0]))
            {
                string msg = string.Format("please add MonthYear:{0} to DB table ETI_Korea_OptionLTDNew", oneFiltRes[0].ToString());
                LogMessage(msg);
            }
            string contractMonthNumber = ((KoreaOptionLastTradingDayInfo)lastTradingDayInDBNew[oneFiltRes[0]]).ContractMonthNumber;

            //if (callOrPut == 1)
            //    TRADING_SYMBOL = "K201" + FindYearCode(yearReal) + contractMonthNumber + oneFiltRes[3];
            //else
            //    TRADING_SYMBOL = "K301" + FindYearCode(yearReal) + contractMonthNumber + oneFiltRes[3];

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
            if (yearCode.Contains(year))
            {
                return yearCode[year].ToString().Trim();
            }
            else
            {
                return null;
            }
        }

        //private string FormatYear(string yearSuffix)
        //{
        //    int currentYear = DateTime.Today.Year;
        //    string currentYearStr = currentYear.ToString().Trim();
        //    string currentYearSuffix = currentYearStr.Substring(currentYearStr.Length - 1);
        //    if (!currentYearSuffix.Equals(yearSuffix))
        //    {
        //        currentYear++;
        //    }

        //    return currentYear.ToString();
        //}

        private string FormatYear(string yearSuffix)
        {
            return string.Format("{0}{1}", DateTime.Now.Year.ToString().Substring(0, 3), yearSuffix);
        }

        private List<List<string>> FormatCsvAddedData(List<List<string>> csvAddedData)
        {
            List<List<string>> tempRes = new List<List<string>>();
            for (int i = 0; i < 4; i++)
            {
                tempRes.Add(new List<string>());
            }

            for (int i = 0; i < csvAddedData.Count; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    if (csvAddedData[i][0].Equals(usefulData.Data[j].Name))
                    {
                        tempRes[j].Add(csvAddedData[i][3] + csvAddedData[i][4]);
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
        /*
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
                //AddResult(Path.GetFileNameWithoutExtension(filePath),filePath,"Requirement File");
            }
            catch (Exception ex)
            {
                Logger.Log("Error found when writing relative data to requirement file." + ex.Message, Logger.LogType.Error);
            }
            Logger.Log("Write data back to Excel - OK!");
        }
        */

        #region the DailyMaintain task will go on the following step

        private void ExecuteDailyMaintain()
        {
            string msg = string.Format("\r\nOn {0}, we did a daily Index Option maintance.", DateTime.Now.ToString("yyyy-MMM-dd HH:mm:ss"));
            Logger.Log(msg);
            PrepareData();
            GetLastTradingDay();
            WriteDataToRequirement();
            GenerateNDAFile();
            GenerateIDNFile();
            //WriteAddedDataToRequirement(ndaDataWriteToRequirement);
        }

        #endregion

        #region generate idn file by read nda file
        private void GenerateIDNFile()
        {
            string idnFolderPath = Path.Combine(configObj.IDN, DateTime.Today.ToString("yyyy-MM-dd"));
            string idnFilePath = Path.Combine(idnFolderPath, string.Format("KOSIO_IDN_BULK_FILE_{0}.txt", DateTime.Now.ToString("dd-MMM-yyyy").Replace("-", "_")));
            string ndaFolderPath = Path.Combine(configObj.NDA, DateTime.Today.ToString("yyyy-MM-dd"));
            string ndaFilePath = Path.Combine(ndaFolderPath, "KR" + DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")) + "_chg_iss.csv");

            if (!Directory.Exists(idnFolderPath))
                Directory.CreateDirectory(idnFolderPath);

            if (!File.Exists(ndaFilePath))
            {
                string msg = string.Format("idn rely on nda file. but the {0} is exist .", ndaFilePath);
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }

            List<string> title = new List<string>() { "SYMBOL", "DSPLY_NAME", "RIC", "OFFCL_CODE", "EX_SYMBOL", "EXPIR_DATE", "CONTR_MNTH", "CONTR_SIZE", "STRIKE_PRC", "PUTCALLIND", "BCKGRNDPAG", "#INSTMOD_MNEMONIC", "#INSTMOD_PROV_SYMB", "EXL_NAME" };
            idnDataWriteToRequirement.Add(title);//add idn file title
            FillIDNBody(idnDataWriteToRequirement, ndaFilePath);
            Logger.Log("get data from nda");

            if (idnDataWriteToRequirement == null || idnDataWriteToRequirement.Count <= 1)
            {
                string msg = string.Format("no valid data in nda file :{0} .", ndaFilePath);
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }

            GenerateIDNBulkFile(idnDataWriteToRequirement, idnFilePath);
            Logger.Log("Generate IDN file - OK!");
            AddResult(Path.GetFileNameWithoutExtension(idnFilePath), idnFilePath, "IDN File");
        }

        private void GenerateIDNBulkFile(List<List<string>> listList, string path)
        {
            StringBuilder sb = new StringBuilder();

            foreach (var list in listList)
            {
                foreach (var item in list)
                {
                    sb.AppendFormat("{0}\t", item);
                }

                sb.Length = sb.Length - 1;
                sb.Append("\r\n");
            }

            sb.Length = sb.Length - 2;
            File.WriteAllText(path, sb.ToString());
        }

        private void FillIDNBody(List<List<string>> idn, string path)
        {
            try
            {
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, path);
                    Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                    if (worksheet != null)
                    {
                        for (int i = 2; i <= lastUsedRow; i++)
                        {
                            List<string> column = new List<string>();
                            //symbol	
                            object ric = ExcelUtil.GetRange(i, 1, worksheet).Value2;
                            string symbol = (ric == null ? string.Empty : ric.ToString());
                            column.Add(symbol);
                            //dsply_name
                            object assetShortName = ExcelUtil.GetRange(i, 27, worksheet).Value2;
                            column.Add(assetShortName == null ? string.Empty : assetShortName.ToString());
                            //ric
                            column.Add(symbol);
                            //offcl_code
                            object tradingSymbol = ExcelUtil.GetRange(i, 25, worksheet).Value2;
                            string offclCode = (tradingSymbol == null ? string.Empty : tradingSymbol.ToString());
                            column.Add(offclCode);
                            //ex_symbol
                            column.Add(offclCode);
                            //expir_date  DD-MMM-YYYY
                            object expiryDateNDA = ExcelUtil.GetRange(i, 9, worksheet).Value;
                            string expiryDate = string.Empty;

                            if (expiryDateNDA != null)
                                expiryDate = Convert.ToDateTime(expiryDateNDA.ToString()).ToString("dd-MMM-yyyy"); //dt.ToString("dd-MMM-yyyy");

                            column.Add(expiryDate);
                            //contr_mnth  MMMY
                            string contrMnth = string.Empty;

                            if ((expiryDate + "").Trim().Length != 0)
                                contrMnth = expiryDate.Substring(3, 3).ToUpper() + expiryDate.Substring(10, 1);

                            column.Add(contrMnth);
                            //contr_size
                            column.Add("1");
                            //strike_prc
                            object strikePrice = ExcelUtil.GetRange(i, 26, worksheet).Value2;
                            column.Add(strikePrice == null ? string.Empty : strikePrice.ToString());
                            //putcallind
                            object callPut = ExcelUtil.GetRange(i, 11, worksheet).Value2;
                            string putcallind = string.Empty;

                            if (callPut != null)
                                putcallind = callPut.ToString().Trim().Equals("CALL") ? "C_EU" : "P_EU";

                            column.Add(putcallind);
                            //bckgrndpag
                            column.Add("****");
                            //#instmod_mnemonic
                            object tickerSymbol = ExcelUtil.GetRange(i, 24, worksheet).Value2;
                            column.Add(tickerSymbol == null ? string.Empty : tickerSymbol.ToString());
                            //#instmod_prov_symb
                            column.Add(offclCode);
                            //exl_name												
                            column.Add("KOSIO_EQO_KS200");

                            idn.Add(column);
                        }
                    }
                    workbook.Close(false, workbook.FullName, Missing.Value);
                }
            }
            catch (Exception e)
            {
                Logger.Log(string.Format("Error happens when get data. Ex: {0} .", e.Message));
            }
        }
        #endregion

        #region the MonthlyMaintain task will go on the following step

        private void ExecuteMonthlyMaintain()
        {

            PrepareData();

            usefulData.Data.RemoveAt(0);

            //string msg = string.Empty;
            //if (monthQuarter)
            //{
            //    msg = "On {0}, we did a quarterly Index Option maintain.";
            //    usefulData.Data.Add(GenerateNewMonthUsefulDateTerm());
            //}
            //else
            //{
            //    msg = "On {0}, we did a monthly Index Option maintain.";
            //    usefulData.Data.Insert(2, GenerateNewMonthUsefulDateTerm());
            //}
            //msg = string.Format(msg, DateTime.Now.ToString("yyyy-MMM-dd HH:mm:ss"));
            //Logger.Log(msg);


            //string nowDateTime = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            //int nowMonthNumber = DateTime.Today.Month;

            if (monthly)
                usefulData.Data.Insert(5, GenerateNewMonthUsefulDateTermNew(DateTime.Now.AddMonths(6)));
            else if (quarterly)
                usefulData.Data.Insert(7, GenerateNewMonthUsefulDateTermNew(DateTime.Now.AddYears(1)));
            else if (semiAnnualSix)
                usefulData.Data.Insert(9, GenerateNewMonthUsefulDateTermNew(DateTime.Now.AddYears(2)));
            else
                usefulData.Data.Add(GenerateNewMonthUsefulDateTermNew(DateTime.Now.AddYears(3)));

            GetLastTradingDay();

            WriteDataToRequirement();
            GenerateNDAFile();
            GenerateIDNFile();
            //WriteAddedDataToRequirement(ndaDataWriteToRequirement);
        }

        private void GetLastTradingDay()
        {
            string test = string.Empty;
            foreach (var dataItem in usefulData.Data)
            {
                try
                {
                    string month = dataItem.Name.Substring(0, 3);
                    KoreaOptionLastTradingDayInfo lastTradingDayItem = KoreaOptionMapNewManager.SelectLastTradingDayByYearMonth(dataItem.Year, month);
                    if (lastTradingDayItem == null)
                    {
                        string msg = string.Format("Can not get last trading day from database. For year; {0}, month: {1}. Please check the Korea Code Map table", dataItem.Year, dataItem.Name.Substring(0, 3));
                        throw new Exception(msg);
                    }
                    //exist month when range from 4 to 11 in different year
                    if (!lastTradingDayInDBNew.ContainsKey(dataItem.Name))
                        lastTradingDayInDBNew.Add(dataItem.Name, lastTradingDayItem);

                    //if (!lastTradingDayInDB.ContainsKey(month))
                    //    lastTradingDayInDB.Add(month, lastTradingDayItem);
                }
                catch (Exception ex)
                {
                    LogMessage("error");
                }
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
            Process gatsProcess = new Process();
            gatsProcess.StartInfo.FileName = file;
            gatsProcess.StartInfo.WorkingDirectory = path;
            gatsProcess.StartInfo.UseShellExecute = false;
            gatsProcess.StartInfo.CreateNoWindow = true;
            string lastPrice = GetLastPrice(gatsProcess);
            List<List<string>> dataToParse = GetLongLinks(gatsProcess);
            if (!string.IsNullOrEmpty(lastPrice) && dataToParse != null)
            {
                usefulData = new GenerateUsefulDateFromList();
                usefulData.ParseData(dataToParse);
                usefulData.Lastvalue = lastPrice;
                //usefulData.Lastvalue = "300.25";
                usefulData.Data.Sort();
            }
            //handle the first time run. There is only 4 month data in DB, then will get 11 month data.
            if (usefulData.Data.Count == 4)
            {
                AddDefaultMonthData(usefulData);
            }

            Logger.Log("Get data from GATS - OK!");
        }

        /// <summary>
        /// 2014.09.01 the first time run this task,there 
        /// is only 4 month data in IDN,so generate 
        /// default value for other 7 month
        /// </summary>
        /// <param name="usefulData"></param>
        private void AddDefaultMonthData(GenerateUsefulDateFromList usefulData)
        {
            try
            {
                double value = 0.0;
                string year = string.Empty;
                List<string> monthQueen = new List<string>() { "SEP4", "OCT4", "NOV4", "DEC4", "JAN5", "FEB5", "MAR5", "JUN5", "DEC5", "JUN6", "DEC6" };//when this month is 01-Sep-2014
                //List<string> monthQueen = new List<string>() { "AUG4", "SEP4", "OCT4", "NOV4", "DEC4", "JAN5", "MAR5", "JUN5", "DEC5", "JUN6", "DEC6" };//when this month is 01-Aug-2014
                foreach (var item in usefulData.Data)
                {
                    if (monthQueen.Contains(item.Name))
                        monthQueen.Remove(item.Name);
                }

                foreach (var item in monthQueen)
                {
                    UsefulDateFormat temp = new UsefulDateFormat();
                    temp.Name = item;
                    temp.High = value;
                    temp.Low = value;
                    //temp.High = 265.23;
                    //temp.Low = 200.36;
                    temp.Year = GetFullYearCode(item.Substring(3));
                    usefulData.Data.Add(temp);
                }
                usefulData.Data.Sort();
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

        private string GetFullYearCode(string item)
        {
            string year = DateTime.Now.Year.ToString();
            if (year.EndsWith(item))
                return year;

            return string.Format("{0}{1}", year.Substring(0, 3), item);
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

                        GroupMonthPriceIDN(price, month);
                    }
                }
            }

            data.Add(prices);
            data.Add(longlinks);
            data.Add(months);
            SortMonthPriceIDN();

            return data;
        }

        private void SortMonthPriceIDN()
        {
            List<string> keys = monthPriceIDN.Keys.ToList();
            foreach (var key in keys)
            {
                monthPriceIDN[key].Sort();
            }
        }

        private void GroupMonthPriceIDN(string price, string month)
        {
            double priceIDN;
            if (!double.TryParse(price, out priceIDN))
            {
                LogMessage(string.Format("convert {0} to dauble type error.", price));
                return;
            }

            if (monthPriceIDN.ContainsKey(month))
            {
                if (!monthPriceIDN[month].Contains(priceIDN))
                    monthPriceIDN[month].Add(priceIDN);
            }
            else
            {
                List<double> temp = new List<double>();
                temp.Add(priceIDN);
                monthPriceIDN.Add(month, temp);
            }
        }

        private string GetExtraPrices(int count, string result, Process gatsProcess)
        {
            int startNum = count + 1;
            //string checkSign = string.Format("LONGNEXTLR +{0}#(?<RIC>[0-9]+)\\*\\.KS", startNum);
            string checkSign = string.Format(@"LONGNEXTLR\s*{0}#KS(?<RIC>[0-9]+)\*\.KS", startNum);
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
            else
            {
                string msg = "Can not get last price from GATS.";
                throw new Exception(msg);
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
            month = callPutOption[monthCode].ToString() + month.Substring(1);
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
