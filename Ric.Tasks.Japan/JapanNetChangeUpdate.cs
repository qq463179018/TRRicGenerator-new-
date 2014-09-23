using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Japan
{
    #region Config

    [ConfigStoredInDB]
    public class JapanNetChangeUpdateConfig
    {
        [StoreInDB]
        [Category("Source File")]
        [DisplayName("Template file")]
        [Description("The path where the template is.")]
        public string TemplateFilePath { get; set; }
    }

    #endregion

    #region Task

    class JapanNetChangeUpdate : GeneratorBase
    {
        #region Declaration

        private static JapanNetChangeUpdateConfig configObj;
        private string baseUrl = "http://www.tse.or.jp/";
        private ExcelApp app;
        private int fileNb;
        private Workbook template;

        private const int ENTRY_OFFSET = 39;
        private List<string> months = new List<string> { "", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
        private List<string> TypeName = new List<string> { "", "株数", "金額" };
        
        #endregion

        #region Interface implementation

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as JapanNetChangeUpdateConfig;
            app = new ExcelApp(false, false);
            if (app.ExcelAppInstance == null)
            {
                Logger.Log("Excel cannot be started", Logger.LogType.Error);
            }
            template = ExcelUtil.CreateOrOpenExcelFile(app, configObj.TemplateFilePath);
        }

        protected override void Start()
        {
            DownloadFiles();
            for (fileNb = 1; fileNb < 3; fileNb++)
            {
                FillTemplate();
            }
            app.Dispose();
        }

        #endregion

        #region Download Excel

        /// <summary>
        /// 
        /// </summary>
        private void DownloadFiles()
        {
            try
            {
                HtmlDocument htc = new HtmlDocument();
                string uri = String.Format("http://www.tse.or.jp/market/data/sector/index.html");
                htc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                string source1Url = tables[2].SelectSingleNode(".//tr[2]/td[3]/a").Attributes["href"].Value.Trim();
                source1Url = baseUrl + source1Url;
                string source2Url = tables[2].SelectSingleNode(".//tr[2]/td[5]/a").Attributes["href"].Value.Trim();
                source2Url = baseUrl + source2Url;
                WebClientUtil.DownloadFile(source1Url, 100000, Path.GetDirectoryName(configObj.TemplateFilePath) + @"\test1.xls");
                WebClientUtil.DownloadFile(source2Url, 100000, Path.GetDirectoryName(configObj.TemplateFilePath) + @"\test2.xls");
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        #endregion

        #region Fill Excel

        /// <summary>
        /// Get starting and finishing date in List from Worksheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private List<string> GetDateFromWorkbook(ref Workbook workbook)
        {
            Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
            string rawDate = worksheet.get_Range("A4").Value2.ToString();

            rawDate = rawDate.Substring(rawDate.IndexOf("("));
            rawDate = rawDate.Replace("(", "").Replace(")", "");
            string[] dateTab = rawDate.Split(new[] { '-' });
            return (from newdate in dateTab 
                    //select newdate.Replace("日", "") into tmp 
                    select newdate.Split(new[] { '/' }) into tmpdate 
                    select String.Format("{0} {1}", tmpdate[1], months[Convert.ToInt32(tmpdate[0])])).ToList();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="date"></param>
        private void WriteDateInTemplate(List<string> date)
        {
            Worksheet templateWs = template.Worksheets[fileNb] as Worksheet;
            templateWs.Cells[4, 1] = date[0];
            templateWs.Cells[4, 2] = date[1];
        }

        /// <summary>
        /// Write the full worksheet in template
        /// </summary>
        private void FillTemplate()
        {
            Workbook workbook = InitializeExcel();
            List<string> date = GetDateFromWorkbook(ref workbook);

            for (int sheetNb = 0; sheetNb < 5; sheetNb++)
            {
                WriteWorksheetInTemplate(ref workbook, sheetNb);
            }
            WriteDateInTemplate(date);

            template.Save();
        }

        /// <summary>
        /// Convert values to the right range and decimals
        /// </summary>
        /// <param name="value">value to change</param>
        /// <param name="minus">yes if it's a difference</param>
        /// <returns>the corrected value</returns>
        private string ConvertValue(string value, bool minus = false)
        {
            if (value == "")
            {
                return "";
            }

            value = value.Replace(",", "");
            int length = 0;
            if (value.StartsWith("-"))
            {
                length = value.Length - 1;
            }
            else
            {
                length = value.Length;
            }
            int divider;
            int mainPrecision;
            if (fileNb == 1)
            {
                divider = 1000;
                mainPrecision = 4;
            }
            else
            {
                divider = 100000;
                mainPrecision = 6;
            }

            if (length >= mainPrecision)
            {
                double tmp = Convert.ToDouble(value);
                tmp /= divider;
                tmp = Math.Round(tmp, MidpointRounding.AwayFromZero);
                //if (minus)
                //{
                //    tmp *= -1;
                //}
                return tmp.ToString();
            }
            else
            {
                int precision = mainPrecision - length;
                double tmp = Convert.ToDouble(value);
                tmp /= divider;
                tmp = Math.Round(tmp, precision, MidpointRounding.AwayFromZero);
                //if (minus)
                //{
                //    tmp *= -1;
                //}
                return tmp.ToString();
            }
        }

        /// <summary>
        /// Get NetChange Type J=>P
        /// </summary>
        /// <param name="name"></param>
        /// <returns>the type</returns>
        private string GetType(string name)
        {
            if (fileNb == 1)
            {
                if (name.Contains("TSE 1st"))
                {
                    return "<NETCHJ>";
                }
                if (name.Contains("TSE 2nd"))
                {
                    return "<NETCHL>";
                }
                if (name.Contains("TSE Mothers"))
                {
                    return "<NETCHN>";
                }
                if (name.Contains("TSE JASDAQ"))
                {
                    return "<NETCHP>";
                }
                if (name.Contains("Tokyo & Nagoya"))
                {
                    return "<NETCHH>";
                }
            }
            else
            {
                if (name.Contains("TSE 1st"))
                {
                    return "<NETCHI>";
                }
                if (name.Contains("TSE 2nd"))
                {
                    return "<NETCHK>";
                }
                if (name.Contains("TSE Mothers"))
                {
                    return "<NETCHM>";
                }
                if (name.Contains("TSE JASDAQ"))
                {
                    return "<NETCHO>";
                }
                if (name.Contains("Tokyo & Nagoya"))
                {
                    return "<NETCHG>";
                }
            }
            return "";
        }

        /// <summary>
        /// Write some data in Template Worksheet
        /// </summary>
        /// <param name="workbook">Reference of the template worbook</param>
        /// <param name="offset">Number of the entry in the worksheet</param>
        private void WriteWorksheetInTemplate(ref Workbook workbook, int offset)
        {
            Worksheet worksheet = workbook.Worksheets[offset + 1] as Worksheet;
            Worksheet templateWs = template.Worksheets[fileNb] as Worksheet;

            offset *= ENTRY_OFFSET;

            templateWs.Cells[3 + offset, 1] = GetType(worksheet.Name);
            templateWs.Cells[3 + offset, 4] = String.Format("{0} {1}", worksheet.Name, TypeName[fileNb]);

            if (fileNb == 1)
            {
                templateWs.Cells[3 + offset, 7] = "(単位：100万株）";
            }
            else
            {
                templateWs.Cells[3 + offset, 7] = "(単位：億円）";
            }

            templateWs.Cells[7 + offset, 1] = ConvertValue(worksheet.Range["K19"].Value2.ToString(), true);  // X14
            templateWs.Cells[7 + offset, 2] = ConvertValue(worksheet.Range["K20"].Value2.ToString()); //AN14
            templateWs.Cells[7 + offset, 3] = ConvertValue(worksheet.Range["K13"].Value2.ToString(), true); //BD14
            templateWs.Cells[7 + offset, 4] = ConvertValue(worksheet.Range["K14"].Value2.ToString()); //BT14
            templateWs.Cells[7 + offset, 5] = ConvertValue(worksheet.Range["K16"].Value2.ToString(), true); //CJ14
            templateWs.Cells[7 + offset, 6] = ConvertValue(worksheet.Range["K17"].Value2.ToString()); //CZ14

            templateWs.Cells[8 + offset, 1] = ConvertValue(worksheet.Range["G19"].Value2.ToString(), true);  // X10
            templateWs.Cells[8 + offset, 2] = ConvertValue(worksheet.Range["G20"].Value2.ToString()); //AN10
            templateWs.Cells[8 + offset, 3] = ConvertValue(worksheet.Range["G13"].Value2.ToString(), true); //BD10
            templateWs.Cells[8 + offset, 4] = ConvertValue(worksheet.Range["G14"].Value2.ToString()); //BT10
            templateWs.Cells[8 + offset, 5] = ConvertValue(worksheet.Range["G16"].Value2.ToString(), true); //CJ10
            templateWs.Cells[8 + offset, 6] = ConvertValue(worksheet.Range["G17"].Value2.ToString()); //CZ10

            templateWs.Cells[13 + offset, 1] = ConvertValue(worksheet.Range["K24"].Value2.ToString(), true);  // T26
            templateWs.Cells[13 + offset, 2] = ConvertValue(worksheet.Range["K25"].Value2.ToString()); //AF26
            templateWs.Cells[13 + offset, 3] = ConvertValue(worksheet.Range["K27"].Value2.ToString(), true); //AR26
            templateWs.Cells[13 + offset, 4] = ConvertValue(worksheet.Range["K28"].Value2.ToString()); //BD26
            templateWs.Cells[13 + offset, 5] = ConvertValue(worksheet.Range["K30"].Value2.ToString(), true); //BP26
            templateWs.Cells[13 + offset, 6] = ConvertValue(worksheet.Range["K31"].Value2.ToString()); //CB26
            templateWs.Cells[13 + offset, 7] = ConvertValue(worksheet.Range["K33"].Value2.ToString(), true); //CN26
            templateWs.Cells[13 + offset, 8] = ConvertValue(worksheet.Range["K34"].Value2.ToString()); //CZ26

            templateWs.Cells[14 + offset, 1] = ConvertValue(worksheet.Range["G24"].Value2.ToString(), true);  // T22
            templateWs.Cells[14 + offset, 2] = ConvertValue(worksheet.Range["G25"].Value2.ToString()); //AF22
            templateWs.Cells[14 + offset, 3] = ConvertValue(worksheet.Range["G27"].Value2.ToString(), true); //AR22
            templateWs.Cells[14 + offset, 4] = ConvertValue(worksheet.Range["G28"].Value2.ToString()); //BD22
            templateWs.Cells[14 + offset, 5] = ConvertValue(worksheet.Range["G30"].Value2.ToString(), true); //BP22
            templateWs.Cells[14 + offset, 6] = ConvertValue(worksheet.Range["G31"].Value2.ToString()); //CB22
            templateWs.Cells[14 + offset, 7] = ConvertValue(worksheet.Range["G33"].Value2.ToString(), true); //CN22
            templateWs.Cells[14 + offset, 8] = ConvertValue(worksheet.Range["G34"].Value2.ToString()); //CZ22

            templateWs.Cells[19 + offset, 5] = ConvertValue(worksheet.Range["K38"].Value2.ToString(), true); //AZ38
            templateWs.Cells[19 + offset, 6] = ConvertValue(worksheet.Range["K39"].Value2.ToString()); //
            templateWs.Cells[19 + offset, 7] = ConvertValue(worksheet.Range["K41"].Value2.ToString(), true); //
            templateWs.Cells[19 + offset, 8] = ConvertValue(worksheet.Range["K42"].Value2.ToString()); //
            templateWs.Cells[19 + offset, 9] = ConvertValue(worksheet.Range["K44"].Value2.ToString(), true); //
            templateWs.Cells[19 + offset, 10] = ConvertValue(worksheet.Range["K45"].Value2.ToString()); //
            templateWs.Cells[19 + offset, 11] = ConvertValue(worksheet.Range["K47"].Value2.ToString(), true); //
            templateWs.Cells[19 + offset, 12] = ConvertValue(worksheet.Range["K48"].Value2.ToString()); //

            templateWs.Cells[20 + offset, 5] = ConvertValue(worksheet.Range["G38"].Value2.ToString(), true); //
            templateWs.Cells[20 + offset, 6] = ConvertValue(worksheet.Range["G39"].Value2.ToString()); //
            templateWs.Cells[20 + offset, 7] = ConvertValue(worksheet.Range["G41"].Value2.ToString(), true); //
            templateWs.Cells[20 + offset, 8] = ConvertValue(worksheet.Range["G42"].Value2.ToString()); //
            templateWs.Cells[20 + offset, 9] = ConvertValue(worksheet.Range["G44"].Value2.ToString(), true); //
            templateWs.Cells[20 + offset, 10] = ConvertValue(worksheet.Range["G45"].Value2.ToString()); //
            templateWs.Cells[20 + offset, 11] = ConvertValue(worksheet.Range["G47"].Value2.ToString(), true); //
            templateWs.Cells[20 + offset, 12] = ConvertValue(worksheet.Range["G48"].Value2.ToString()); //

            templateWs.Cells[24 + offset, 5] = ConvertValue(worksheet.Range["K52"].Value2.ToString(), true); //
            templateWs.Cells[24 + offset, 6] = ConvertValue(worksheet.Range["K53"].Value2.ToString()); //
            templateWs.Cells[24 + offset, 7] = ConvertValue(worksheet.Range["K55"].Value2.ToString(), true); //
            templateWs.Cells[24 + offset, 8] = ConvertValue(worksheet.Range["K56"].Value2.ToString()); //
            templateWs.Cells[24 + offset, 9] = ConvertValue(worksheet.Range["K58"].Value2.ToString(), true); //
            templateWs.Cells[24 + offset, 10] = ConvertValue(worksheet.Range["K59"].Value2.ToString()); //
            templateWs.Cells[24 + offset, 11] = ConvertValue(worksheet.Range["K61"].Value2.ToString(), true); //
            templateWs.Cells[24 + offset, 12] = ConvertValue(worksheet.Range["K62"].Value2.ToString()); //

            templateWs.Cells[25 + offset, 5] = ConvertValue(worksheet.Range["G52"].Value2.ToString(), true);
            templateWs.Cells[25 + offset, 6] = ConvertValue(worksheet.Range["G53"].Value2.ToString()); //
            templateWs.Cells[25 + offset, 7] = ConvertValue(worksheet.Range["G55"].Value2.ToString(), true); //
            templateWs.Cells[25 + offset, 8] = ConvertValue(worksheet.Range["G56"].Value2.ToString()); //
            templateWs.Cells[25 + offset, 9] = ConvertValue(worksheet.Range["G58"].Value2.ToString(), true); //
            templateWs.Cells[25 + offset, 10] = ConvertValue(worksheet.Range["G59"].Value2.ToString()); //
            templateWs.Cells[25 + offset, 11] = ConvertValue(worksheet.Range["G61"].Value2.ToString(), true); //
            templateWs.Cells[25 + offset, 12] = ConvertValue(worksheet.Range["G62"].Value2.ToString()); //

            templateWs.Cells[17 + offset, 1] = ConvertValue(worksheet.Range["B9"].Value2.ToString()); //

            templateWs.Cells[22 + offset, 1] = Math.Round(Convert.ToDouble(worksheet.Range["E9"].Value2.ToString().Replace("%", "")), 1, MidpointRounding.AwayFromZero); //
            templateWs.Cells[22 + offset, 2] = Math.Round(Convert.ToDouble(worksheet.Range["G9"].Value2.ToString().Replace("%", "")), 1, MidpointRounding.AwayFromZero); //
            templateWs.Cells[22 + offset, 3] = Math.Round(Convert.ToDouble(worksheet.Range["I9"].Value2.ToString().Replace("%", "")) * 1.0, 1, MidpointRounding.AwayFromZero); //
            templateWs.Range["A" + (22 + offset), "C" + (22 + offset)].NumberFormat = ".0";


            templateWs.Cells[27 + offset, 2] = ConvertValue(worksheet.Range["C68"].Value2.ToString()); //
            templateWs.Cells[28 + offset, 2] = Convert.ToDouble(worksheet.Range["D68"].Value2.ToString().Replace("%", "")); //
            templateWs.Cells[29 + offset, 2] = ConvertValue(worksheet.Range["C69"].Value2.ToString()); //
            templateWs.Cells[30 + offset, 2] = Convert.ToDouble(worksheet.Range["D69"].Value2.ToString().Replace("%", "")); //
            templateWs.Cells[31 + offset, 2] = ConvertValue(worksheet.Range["C70"].Value2.ToString()); //
            templateWs.Cells[32 + offset, 2] = ConvertValue(worksheet.Range["C71"].Value2.ToString()); //
            templateWs.Range["B" + (28 + offset), "C" + (28 + offset)].NumberFormat = ".0";

            templateWs.Cells[27 + offset, 3] = ConvertValue(worksheet.Range["E68"].Value2.ToString()); //
            templateWs.Cells[28 + offset, 3] = Convert.ToDouble(worksheet.Range["F68"].Value2.ToString().Replace("%", "")); //
            templateWs.Cells[29 + offset, 3] = ConvertValue(worksheet.Range["E69"].Value2.ToString()); //
            templateWs.Cells[30 + offset, 3] = Convert.ToDouble(worksheet.Range["F69"].Value2.ToString().Replace("%", "")); //
            templateWs.Cells[31 + offset, 3] = ConvertValue(worksheet.Range["E70"].Value2.ToString()); //
            templateWs.Cells[32 + offset, 3] = ConvertValue(worksheet.Range["E71"].Value2.ToString()); //
            worksheet.Range["B" + (30 + offset), "C" + (30 + offset)].NumberFormat = ".0";

            templateWs.Cells[36 + offset, 2] = ConvertValue(worksheet.Range["D75"].Value2.ToString()); //
            templateWs.Cells[36 + offset, 3] = ConvertValue(worksheet.Range["F75"].Value2.ToString()); //
            templateWs.Cells[37 + offset, 2] = ConvertValue(worksheet.Range["D76"].Value2.ToString()); //
            templateWs.Cells[37 + offset, 3] = ConvertValue(worksheet.Range["F76"].Value2.ToString()); //
        }

        /// <summary>
        /// Initialize workbook
        /// </summary>
        /// <returns>the created or opened workbook</returns>
        private Workbook InitializeExcel()
        {
            Workbook workbook;
            try
            {
                workbook = ExcelUtil.CreateOrOpenExcelFile(app, Path.GetDirectoryName(configObj.TemplateFilePath) + @"\test" + fileNb + ".xls");
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

    #endregion
}
