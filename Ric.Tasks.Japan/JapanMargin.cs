using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks.Japan
{
    #region Configuration

    [ConfigStoredInDB]
    public class JapanMarginConfig
    {
        [StoreInDB]
        [DisplayName("Result file")]
        [Description("The path where the result will be written. Must include the file name as well\n\nEg: C:/Mydrive/fileName.xls")]
        public string ResultFilePath { get; set; }

        [StoreInDB]
        [DefaultValue("1")]
        [Description("Choose what to do : 1 for creation, 2 for update")]
        public string type { get; set; }
    }

    #endregion

    class JapanMargin : GeneratorBase
    {
        #region Declaration

        private static JapanMarginConfig configObj;
        private const string baseUrl = "http://www.tse.or.jp/";
        private ExcelApp app;
        public List<List<string>> results1 = new List<List<string>>();
        public List<List<string>> results2 = new List<List<string>>();
        public string alphabet = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        private List<string> titles1 = new List<string> { "", "合", "計", "", "売残高", "", "買残高", "" };
        private List<string> titles2 = new List<string> { "売残高", "前週比", "買残高", "前週比", "一般信用取引", "制度信用取引", "一般信用取引", "制度信用取引" };
        private List<DateTime> holidays = new List<DateTime>();
        private Workbook workbook;

        private enum LastWeek { FirstDay, LastDay };

        #endregion

        #region Interface implementation

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as JapanMarginConfig;
            app = new ExcelApp(false, false);
            if (app.ExcelAppInstance == null)
            {
                Logger.Log("Excel cannot be started", Logger.LogType.Error);
            }
        }

        protected override void Start()
        {
            int type = GetType();
            GetHolidays();
            if (type == 1)
            {
                if (File.Exists(configObj.ResultFilePath))
                {
                    File.Delete(configObj.ResultFilePath);
                }
                workbook = InitializeExcel(configObj.ResultFilePath);
                workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[1]);
                DownloadFirstFile();
                FillTemplate();
                FillTemplate2();
                ReadExcel();
                FillExcel();
            }
            else
            {
                workbook = InitializeExcel(configObj.ResultFilePath);
                DownloadSecondFile();
                ReadExcel2();
                FillExcel2();
            }
            Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
            worksheet.Activate();
            workbook.SaveAs(workbook.FullName);
            AddResult("Result file", workbook.FullName, "file");
            workbook.Close();
            app.Dispose();
        }

        #endregion

        #region Config initialization

        /// <summary>
        /// Find from Config if it's the creation or the update
        /// </summary>
        /// <returns>1 for creation and 2 for update</returns>
        private int GetType()
        {
            if (configObj.type == "1" || configObj.type == "2")
            {
                return Convert.ToInt16(configObj.type);
            }
            return 1;
        }

        #endregion

        #region Download files

        /// <summary>
        /// Download Excel for creation
        /// </summary>
        private void DownloadFirstFile()
        {
            try
            {
                HtmlDocument htc = new HtmlDocument();
                string uri = String.Format("http://www.tse.or.jp/market/data/margin/index.html");
                htc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                int trNb = tables[3].SelectNodes(".//tr").Count;
                string source1Url = tables[3].SelectSingleNode(String.Format(".//tr[{0}]/td[2]/a", trNb)).Attributes["href"].Value.Trim();
                source1Url = baseUrl + source1Url;
                WebClientUtil.DownloadFile(source1Url, 100000, Path.GetDirectoryName(configObj.ResultFilePath) + @"\test1.xls");
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        /// <summary>
        /// Download Excel for update
        /// </summary>
        private void DownloadSecondFile()
        {
            try
            {
                HtmlDocument htc = new HtmlDocument();
                string uri = String.Format("http://www.tse.or.jp/market/data/margin/index.html");
                htc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                int trNb = tables[4].SelectNodes(".//tr").Count;
                string source2Url = tables[4].SelectSingleNode(String.Format(".//tr[{0}]/td[2]/a", trNb)).Attributes["href"].Value.Trim();
                source2Url = baseUrl + source2Url;
                WebClientUtil.DownloadFile(source2Url, 100000, Path.GetDirectoryName(configObj.ResultFilePath) + @"\test2.xls");
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        #endregion

        #region Getting holidays

        /// <summary>
        /// Getting list of holidays date from workbook path given in parameters
        /// </summary>
        private void GetHolidays()
        {
            try
            {
                HolidayManager.SelectHoliday(5);
            }
            catch (Exception ex)
            {
                Logger.Log("Cannot get holidays", Logger.LogType.Error);
                throw new Exception("Cannot get holidays: " + ex.Message);
            }
        }

        /// <summary>
        /// Get last week first and last business days
        /// </summary>
        /// <param name="lastweek">Enum for first or last</param>
        /// <returns>The requested date</returns>
        private string GetSpecificDate(LastWeek lastweek)
        {
            int dayWeek = (int)DateTime.Now.DayOfWeek;
            DateTime dateToFind;
            int daysToAdd = 0;
            try
            {
                if (lastweek == LastWeek.FirstDay)
                {
                    dateToFind = DateTime.Now.AddDays((dayWeek + 6) * -1);
                    if (holidays.Contains(dateToFind))
                    {
                        daysToAdd = 1;
                    }
                }
                else
                {
                    dateToFind = DateTime.Now.AddDays((dayWeek + 2) * -1);
                    if (holidays.Contains(dateToFind))
                    {
                        daysToAdd = -1;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Cannot get Specific date", Logger.LogType.Error);
                throw new Exception("Cannot get specific date"  + ex.Message);
            }
            return dateToFind.AddDays(daysToAdd).ToString("ddMMM").ToUpper();
        }

        #endregion

        #region Reading files

        /// <summary>
        /// Get informations from Excel for creation
        /// </summary>
        private void ReadExcel()
        {
            Workbook workbookSource = InitializeExcel(Path.GetDirectoryName(configObj.ResultFilePath) + @"\test1.xls");
            try
            {
                Worksheet worksheet = workbookSource.Worksheets[1] as Worksheet;
                for (int line = 7; line < 25; line++)
                {
                    List<string> newLine = new List<string>();
                    for (int column = 5; column < 16; column++)
                    {
                        if (column != 10 && column != 12 && column != 14)
                        {
                            newLine.Add(worksheet.get_Range(alphabet.Substring(column, 1) + line).Value2.ToString());
                        }
                    }
                    results1.Add(newLine);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Cannot get informations from Excel", Logger.LogType.Error);
                throw new Exception("Cannot get informations from Excel: " + ex.Message);
            }
            finally
            {
                workbookSource.SaveAs(workbookSource.FullName);
                workbookSource.Close();
            }
        }

        /// <summary>
        /// Get informations from Excel for update
        /// </summary>
        private void ReadExcel2()
        {
            Workbook workbookSource = InitializeExcel(Path.GetDirectoryName(configObj.ResultFilePath) + @"\test2.xls");
            try
            {
                Worksheet worksheet = workbookSource.Worksheets[1] as Worksheet;
                for (int line = 7; line < 13; line++)
                {
                    List<string> newLine = new List<string>();
                    for (int column = 4; column < 12; column++)
                    {
                        newLine.Add(worksheet.get_Range(alphabet.Substring(column, 1) + line).Value2.ToString());
                    }
                    results2.Add(newLine);
                }
                for (int line = 17; line < 23; line++)
                {
                    List<string> newLine = new List<string>();
                    int startColumn = line % 2 == 0 ? 6 : 4;
                    for (int column = startColumn; column < 14; column++)
                    {
                        newLine.Add(worksheet.get_Range(alphabet.Substring(column, 1) + line).Value2.ToString());
                    }
                    results2.Add(newLine);
                }
                List<string> newLine2 = new List<string>();
                for (int line = 28; line < 31; line++)
                {
                    newLine2.Add(worksheet.Range["D34"].Value2.ToString());
                    newLine2.Add(worksheet.Range["D36"].Value2.ToString());
                    newLine2.Add(worksheet.Range["D38"].Value2.ToString());
                }
                results2.Add(newLine2);
                for (int line = 27; line < 30; line++)
                {
                    if (line == 28)
                    {
                        continue;
                    }
                    List<string> newLine = new List<string>();
                    for (int column = 4; column < 8; column++)
                    {
                        newLine.Add(worksheet.get_Range(alphabet.Substring(column, 1) + line).Value2.ToString());
                    }
                    results2.Add(newLine);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Cannot get informations from 2nd Excel", Logger.LogType.Error);
                throw new Exception("Cannot get informations from 2nd Excel: " + ex.Message);
            }
            finally
            {
                workbookSource.SaveAs(workbookSource.FullName);
                workbookSource.Close();
            }
        }

        #endregion

        #region Writing Template

        /// <summary>
        /// Create template Excel Worksheet for creation
        /// </summary>
        private void FillTemplate()
        {
            Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
            worksheet.Name = "第2営業日";

            worksheet.Cells[3, 2] = "WEEK OF";
            worksheet.Cells[3, 3] = GetSpecificDate(LastWeek.FirstDay);
            Range rg1 = (Range)worksheet.Cells[3, 3];
            rg1.NumberFormat = "ddmmm";
            worksheet.Cells[3, 4] = GetSpecificDate(LastWeek.LastDay);
            Range rg2 = (Range)worksheet.Cells[3, 4];
            rg2.NumberFormat = "ddmmm";
            worksheet.Cells[3, 5] = DateTime.Now.ToString("ddMMMyy").ToUpper();
            Range rg3 = (Range)worksheet.Cells[3, 5];
            rg3.NumberFormat = "ddmmmyy";

            worksheet.Range["B6", "B13"].Merge();
            worksheet.Range["B6", "B13"].BorderAround();
            worksheet.Cells[6, 2] = "合計";

            worksheet.Range["B14", "B21"].Merge();
            worksheet.Range["B14", "B21"].BorderAround();
            worksheet.Cells[14, 2] = "市場第一部1st";

            worksheet.Range["B22", "B29"].Merge();
            worksheet.Range["B22", "B29"].BorderAround();
            worksheet.Cells[22, 2] = "市場第二部2nd";

            worksheet.Range["C6", "C7"].Merge();
            worksheet.Range["C6", "C7"].BorderAround();
            worksheet.Cells[6, 3] = "二市場";
            worksheet.Cells[6, 4] = "株数";
            worksheet.Cells[7, 4] = "金額";

            worksheet.Range["C9", "C10"].Merge();
            worksheet.Range["C9", "C10"].BorderAround();
            worksheet.Cells[9, 3] = "東京";
            worksheet.Cells[9, 4] = "株数";
            worksheet.Cells[10, 4] = "金額";

            worksheet.Range["C12", "C13"].Merge();
            worksheet.Range["C12", "C13"].BorderAround();
            worksheet.Cells[12, 3] = "名古屋";
            worksheet.Cells[12, 4] = "株数";
            worksheet.Cells[13, 4] = "金額";

            worksheet.Range["C14", "C15"].Merge();
            worksheet.Range["C14", "C15"].BorderAround();
            worksheet.Cells[14, 3] = "二市場";
            worksheet.Cells[14, 4] = "株数";
            worksheet.Cells[15, 4] = "金額";

            worksheet.Range["C17", "C18"].Merge();
            worksheet.Range["C17", "C18"].BorderAround();
            worksheet.Cells[17, 3] = "東京";
            worksheet.Cells[17, 4] = "株数";
            worksheet.Cells[18, 4] = "金額";

            worksheet.Range["C20", "C21"].Merge();
            worksheet.Range["C20", "C21"].BorderAround();
            worksheet.Cells[20, 3] = "名古屋";
            worksheet.Cells[20, 4] = "株数";
            worksheet.Cells[21, 4] = "金額";

            worksheet.Range["C22", "C23"].Merge();
            worksheet.Range["C22", "C23"].BorderAround();
            worksheet.Cells[22, 3] = "二市場";
            worksheet.Cells[22, 4] = "株数";
            worksheet.Cells[23, 4] = "金額";

            worksheet.Range["C25", "C26"].Merge();
            worksheet.Range["C25", "C26"].BorderAround();
            worksheet.Cells[25, 3] = "東京";
            worksheet.Cells[25, 4] = "株数";
            worksheet.Cells[26, 4] = "金額";

            worksheet.Range["C28", "C29"].Merge();
            worksheet.Range["C28", "C29"].BorderAround();
            worksheet.Cells[28, 3] = "名古屋";
            worksheet.Cells[28, 4] = "株数";
            worksheet.Cells[29, 4] = "金額";

            ((Range)worksheet.Columns["A"]).ColumnWidth = 2;
            ((Range)worksheet.Columns["B"]).ColumnWidth = 15;
            ((Range)worksheet.Columns["C:E"]).ColumnWidth = 10;
            ((Range)worksheet.Columns["I:L"]).ColumnWidth = 13;

            for (int line = 4; line < 30; line++)
            {
                for (int column = 4; column < 13; column++)
                {
                    worksheet.Range[alphabet.Substring(column, 1) + line].BorderAround();
                }
            }
            for (int column = 5; column < (titles1.Count + 5); column++)
            {
                worksheet.Cells[4, column] = titles1[column - 5];
                worksheet.Cells[5, column] = titles2[column - 5];
            }
            worksheet.Range["B4", "L29"].BorderAround(Type.Missing, XlBorderWeight.xlMedium);
        }

        /// <summary>
        /// Create template Excel Worksheet for update
        /// </summary>
        private void FillTemplate2()
        {
            Worksheet worksheet = workbook.Worksheets[2] as Worksheet;
            try
            {
                worksheet.Name = "第3営業日";

                worksheet.Cells[3, 2] = "WEEK OF";

                worksheet.Cells[4, 5] = "       委            託";
                worksheet.Cells[4, 9] = "       自          己";

                worksheet.Cells[5, 5] = "売残高";
                worksheet.Range["E5"].BorderAround();
                worksheet.Cells[5, 6] = "前週比";
                worksheet.Range["F5"].BorderAround();
                worksheet.Cells[5, 7] = "買残高";
                worksheet.Range["G5"].BorderAround();
                worksheet.Cells[5, 8] = "前週比";
                worksheet.Range["H5"].BorderAround();
                worksheet.Cells[5, 9] = "売残高";
                worksheet.Range["I5"].BorderAround();
                worksheet.Cells[5, 10] = "前週比";
                worksheet.Range["J5"].BorderAround();
                worksheet.Cells[5, 11] = "買残高";
                worksheet.Range["K5"].BorderAround();
                worksheet.Cells[5, 12] = "前週比";
                worksheet.Range["L5"].BorderAround();

                //////////////////////////////////////

                worksheet.Cells[6, 2] = "二";
                worksheet.Cells[7, 2] = "市";
                worksheet.Cells[9, 2] = "場";
                worksheet.Cells[10, 2] = "合";

                worksheet.Cells[18, 2] = "二";
                worksheet.Cells[19, 2] = "市";
                worksheet.Cells[21, 2] = "場";
                worksheet.Cells[22, 2] = "合";

                ///////////////////////////////////////

                worksheet.Cells[9, 3] = "東京";
                worksheet.Cells[12, 3] = "名古屋";
                worksheet.Cells[21, 3] = "東京";
                worksheet.Cells[24, 3] = "名古屋";

                /////////////////////////////////////////

                worksheet.Range["B6", "D13"].BorderAround();
                worksheet.Range["B18", "D25"].BorderAround();

                worksheet.Range["C9", "D10"].BorderAround();
                worksheet.Range["C12", "D13"].BorderAround();
                worksheet.Range["D6", "D13"].BorderAround();
                worksheet.Range["C9", "C13"].BorderAround();

                worksheet.Range["C21", "D22"].BorderAround();
                worksheet.Range["C24", "D25"].BorderAround();
                worksheet.Range["D18", "D25"].BorderAround();
                worksheet.Range["C21", "C25"].BorderAround();


                ////////////////////////////////////////

                worksheet.Cells[6, 4] = "株数";
                worksheet.Cells[7, 4] = "金額";

                worksheet.Cells[9, 4] = "株数";
                worksheet.Cells[10, 4] = "金額";

                worksheet.Cells[12, 4] = "株数";
                worksheet.Cells[13, 4] = "金額";

                worksheet.Cells[18, 4] = "株数";
                worksheet.Cells[19, 4] = "金額";

                worksheet.Cells[21, 4] = "株数";
                worksheet.Cells[22, 4] = "金額";

                worksheet.Cells[24, 4] = "株数";
                worksheet.Cells[25, 4] = "金額";

                /////////////////////////////////////

                worksheet.Range["E16", "E17"].Merge();
                worksheet.Range["F16", "F17"].Merge();
                worksheet.Range["E16", "E17"].BorderAround();
                worksheet.Range["F16", "F17"].BorderAround();
                worksheet.Cells[16, 5] = "社内対当";
                worksheet.Cells[16, 6] = "前週比";

                /////////////////////////////////////////

                worksheet.Cells[16, 7] = "貸借取引残高";
                worksheet.Cells[17, 7] = "貸株";
                worksheet.Cells[17, 8] = "前週比";
                worksheet.Cells[17, 9] = "融資";
                worksheet.Cells[17, 10] = "前週比";
                worksheet.Range["G17"].BorderAround();
                worksheet.Range["H17"].BorderAround();
                worksheet.Range["I17"].BorderAround();
                worksheet.Range["J17"].BorderAround();

                /////////////////////////////////////////

                worksheet.Range["K16", "L16"].Merge();
                worksheet.Range["M16", "N16"].Merge();
                worksheet.Cells[16, 11] = "自己貸株";
                worksheet.Cells[16, 13] = "自己融資";

                worksheet.Cells[17, 12] = "前週比";
                worksheet.Cells[17, 14] = "前週比";

                worksheet.Range["K16", "L17"].BorderAround();
                worksheet.Range["M16", "N17"].BorderAround();
                worksheet.Range["L17"].BorderAround();
                worksheet.Range["N17"].BorderAround();

                /////////////////////////////////////////

                worksheet.Cells[27, 4] = "東京市場分・最高最低記録";
                worksheet.Cells[29, 4] = "売残高";
                worksheet.Cells[30, 4] = "買残高";
                worksheet.Cells[31, 4] = "自己融資";
                worksheet.Cells[32, 4] = "信用倍率";

                worksheet.Cells[28, 5] = "最高";
                worksheet.Cells[28, 6] = "最低";

                worksheet.Cells[27, 8] = "売残高に対する買残高の倍率（株数）";
                worksheet.Cells[28, 8] = "二市場合計";
                worksheet.Cells[29, 8] = "東京";
                worksheet.Cells[30, 8] = "名古屋";

                //////////////////////////////////////////////////////

                for (int line = 28; line < 33; line++)
                {
                    for (int column = 4; column < 7; column++)
                    {
                        worksheet.Range[alphabet.Substring(column, 1) + line].BorderAround();
                    }
                }
                for (int line = 28; line < 31; line++)
                {
                    for (int column = 8; column < 10; column++)
                    {
                        worksheet.Range[alphabet.Substring(column, 1) + line].BorderAround();
                    }
                }

                for (int line = 6; line < 14; line++)
                {
                    if (line == 8 || line == 11)
                    {
                        continue;
                    }
                    for (int column = 5; column < 13; column++)
                    {
                        worksheet.Range[alphabet.Substring(column, 1) + line].BorderAround();
                    }
                }

                for (int line = 18; line < 26; line++)
                {
                    if (line == 20 || line == 23)
                    {
                        continue;
                    }
                    for (int column = 5; column < 15; column++)
                    {
                        worksheet.Range[alphabet.Substring(column, 1) + line].BorderAround();
                    }
                }
                ///////////////////////////////////////////////////

                ((Range)worksheet.Columns["A"]).ColumnWidth = 2;
                ((Range)worksheet.Columns["B:E"]).ColumnWidth = 10;
                ((Range)worksheet.Columns["H"]).ColumnWidth = 12;

                /////////////////////////////////////////////////////////////////

                worksheet.Range["B4", "L13"].BorderAround(Type.Missing, XlBorderWeight.xlMedium);
                worksheet.Range["B16", "N25"].BorderAround(Type.Missing, XlBorderWeight.xlMedium);
            }
            catch (Exception ex)
            {
                Logger.Log("Cannot write template", Logger.LogType.Error);
                throw new Exception("Cannot write template: " + ex.Message);
            }
        }

        #endregion

        #region Writing Files

        /// <summary>
        /// Fill 1st Worksheet with information needed for creation
        /// </summary>
        private void FillExcel()
        {
            Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
            try
            {
                for (int line = 6, i = 0; line < 30; line++)
                {
                    if (line == 8 || line == 11 || line == 16 || line == 19 || line == 24 || line == 27)
                    {
                        continue;
                    }
                    for (int column = 5; column < 13; column++)
                    {
                        worksheet.Cells[line, column] = results1[i][column - 5];
                    }
                    i++;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Cannot fill 1st excel with infos", Logger.LogType.Error);
                throw new Exception("Cannot fill 1st excel with infos: " + ex.Message);
            }
            finally
            {
                File.Delete(Path.GetDirectoryName(configObj.ResultFilePath) + @"\test1.xls");
            }
        }

        /// <summary>
        /// Fill 2nd Worksheet with information needed for update
        /// </summary>
        private void FillExcel2()
        {
            Worksheet worksheet = workbook.Worksheets[2] as Worksheet;
            Worksheet worksheetFirst = workbook.Worksheets[1] as Worksheet;
            try
            {
                worksheet.Cells[3, 3] = GetSpecificDate(LastWeek.FirstDay);
                Range rg1 = (Range)worksheet.Cells[3, 3];
                rg1.NumberFormat = "ddmmm";
                worksheet.Cells[3, 4] = GetSpecificDate(LastWeek.LastDay);
                Range rg2 = (Range)worksheet.Cells[3, 4];
                rg2.NumberFormat = "ddmmm";
                worksheet.Cells[3, 5] = DateTime.Now.ToString("ddMMMyy").ToUpper();
                Range rg3 = (Range)worksheet.Cells[3, 5];
                rg3.NumberFormat = "ddmmmyy";

                worksheetFirst.Cells[3, 5] = DateTime.Now.ToString("ddMMMyy").ToUpper();
                Range rg4 = (Range)worksheet.Cells[3, 5];
                rg4.NumberFormat = "ddmmmyy";

                int i = 0;

                for (int line = 6; line < 14; line++)
                {
                    if (line == 8 || line == 11)
                    {
                        continue;
                    }
                    for (int column = 5; column < 13; column++)
                    {
                        worksheet.Cells[line, column] = results2[i][column - 5];
                    }
                    i++;
                }
                for (int line = 18; line < 26; line++)
                {
                    int startColumn;
                    if (line == 20 || line == 23)
                    {
                        continue;
                    }
                    if (line == 19 || line == 22 || line == 25)
                    {
                        startColumn = 7;
                    }
                    else
                    {
                        startColumn = 5;
                    }
                    for (int column = startColumn; column < 15; column++)
                    {
                        worksheet.Cells[line, column] = results2[i][column - startColumn];
                    }
                    i++;
                }

                for (int line = 29; line < 32; line++)
                {
                    worksheet.Cells[line, 5] = Math.Round(Convert.ToDouble(results2[13][line - 29]) / 10, 1, MidpointRounding.AwayFromZero);
                    worksheet.Cells[line, 6] = Math.Round(Convert.ToDouble(results2[14][line - 29]) / 10, 1, MidpointRounding.AwayFromZero);
                    worksheet.Range["E29", "F31"].NumberFormat = ".0";
                }

                worksheet.Cells[32, 5] = Convert.ToDouble(results2[13][3]);
                worksheet.Cells[32, 6] = Convert.ToDouble(results2[14][3]);
                worksheet.Range["E32", "F32"].NumberFormat = ".00";

                worksheet.Cells[28, 9] = Math.Round(Convert.ToDouble(results2[12][0]), 2, MidpointRounding.AwayFromZero);
                worksheet.Cells[29, 9] = Math.Round(Convert.ToDouble(results2[12][1]), 2, MidpointRounding.AwayFromZero);
                worksheet.Cells[30, 9] = Math.Round(Convert.ToDouble(results2[12][2]), 2, MidpointRounding.AwayFromZero);
                worksheet.Range["I28", "I30"].NumberFormat = ".00";
            }
            catch (Exception ex)
            {
                Logger.Log("Cannot fill 2nd excel with infos", Logger.LogType.Error);
                throw new Exception("Cannot fill 2nd excel with infos: " + ex.Message);
            }
            finally
            {
                File.Delete(Path.GetDirectoryName(configObj.ResultFilePath) + @"\test2.xls");
            }
        }

        /// <summary>
        /// Workbook Initialization from path
        /// </summary>
        /// <param name="path">Path of the workbook</param>
        /// <returns>The Created or opened workbook</returns>
        private Workbook InitializeExcel(string path)
        {
            Workbook workbook;
            try
            {
                workbook = ExcelUtil.CreateOrOpenExcelFile(app, path);
                return workbook;
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex;
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception("Cannot load workbook");
            }
        }

        #endregion
    }
}
