using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using Ric.Util;

namespace Ric.Tasks.Japan
{
    public class JapanShared
    {
        /// <summary>
        /// Get nameMap from baknote.xls file
        /// </summary>
        /// <param name="bakNoteFilePath">baknote file path</param>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        public static Dictionary<string, NameMap> GetNameMap(string bakNoteFilePath, string worksheetName)
        {
            Dictionary<string, NameMap> nameDic = new Dictionary<string, NameMap>();
            try
            {
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, bakNoteFilePath);
                    var worksheet = ExcelUtil.GetWorksheet(worksheetName, workbook);
                    if (worksheet == null)
                    {
                        throw new System.Exception(string.Format("There's no such worksheet {0}", worksheetName));
                    }

                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                    for (int i = 4; i <= lastUsedRow; i++)
                    {
                        string key = string.Empty;
                        if (ExcelUtil.GetRange(i, 4, worksheet).Value2 != null && ExcelUtil.GetRange(i, 4, worksheet).Value2.ToString() != string.Empty)
                        {
                            key = ExcelUtil.GetRange(i, 2, worksheet).Value2.ToString().Trim();
                        }
                        if (!nameDic.ContainsKey(key))
                        {
                            NameMap map = new NameMap();
                            if (ExcelUtil.GetRange(i, 4, worksheet).Value2 != null)
                            {
                                map.JapaneseName = ExcelUtil.GetRange(i, 4, worksheet).Value2.ToString().Trim();
                            }
                            if (ExcelUtil.GetRange(i, 3, worksheet).Value2 != null)
                            {
                                map.EnglistName = ExcelUtil.GetRange(i, 3, worksheet).Value2.ToString().Trim();
                            }

                            if (ExcelUtil.GetRange(i, 5, worksheet).Value2 != null)
                            {
                                map.Ric = ExcelUtil.GetRange(i, 5, worksheet).Value2.ToString().Trim();
                            }

                            if (ExcelUtil.GetRange(i, 6, worksheet).Value2 != null)
                            {
                                map.ShortName = ExcelUtil.GetRange(i, 6, worksheet).Value2.ToString().Trim();
                            }
                            nameDic.Add(key, map);
                        }
                    }

                    workbook.Close(false, workbook.FullName, Missing.Value);
                }
            }
            catch (System.Exception)
            {
                throw new System.Exception("There's an error when getting name map from file baknote");
            }

            return nameDic;
        }


        /// <summary>
        /// Update company information with Japanese name, English name and short name according to the name map
        /// </summary>
        /// <param name="nameDic"></param>
        /// <param name="companyInfo"></param>
        public static void UpdateCompanyInfo(Dictionary<string, NameMap> nameDic, CompanyInfo companyInfo, out string warningMsg)
        {
            warningMsg = string.Empty;
            foreach (var item in nameDic.Where(item => item.Key.Trim() == companyInfo.OriginalName.Trim()))
            {
                companyInfo.JapaneseName = item.Value.JapaneseName;
                companyInfo.EnglishName = item.Value.EnglistName;
                companyInfo.ShortEnglishName = item.Value.ShortName;
                companyInfo.Ric = item.Value.Ric.Replace("0#","");
                break;
            }

            if (string.IsNullOrEmpty(companyInfo.EnglishName))
            {
                warningMsg = string.Format("There's no such name for {0}, please check the baknote file.", companyInfo.OriginalName);
            }
        }

        /// <summary>
        /// Print excel worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet which will be print</param>
        /// <param name="printerName">The printer which will be used to print the worksheet</param>
        /// <param name="pageOrientation">The page orientation</param>
        public static void PrintWorksheet(Worksheet worksheet, string printerName, XlPageOrientation pageOrientation)
        {
            if (!string.IsNullOrEmpty(printerName))
            {
                worksheet.PageSetup.Orientation = pageOrientation;
                worksheet.PrintOut(Missing.Value,
                                     Missing.Value,
                                     1,
                                     Missing.Value,
                                     printerName,
                                     Missing.Value,
                                     Missing.Value,
                                     Missing.Value);
            }
        }

        public static DateTime TransferJpDate(string sourceDateStr)
        {
            sourceDateStr = sourceDateStr.Trim().Replace("（", "").Replace("）", "").Replace(" ","");
            Regex regex = new Regex("(?s)((?<yearName>.*?)(?<year>\\d{1,})年){0,}((?<month>\\d{1,})月){0,}((?<day>\\d{1,})日{0,}){0,}");
            Match m = regex.Match(sourceDateStr);
            string year = m.Groups["year"].Value;
            string month = m.Groups["month"].Value;
            string day = m.Groups["day"].Value;
            if (year == string.Empty)
            {
                year = DateTime.Now.ToString("yyyy");
            }
            else
            {
                int yearNum = int.Parse(year);
                year = (yearNum + 1988).ToString(); //平成
            }
            string dateTime = month;
            if (day.Trim() != string.Empty)
            {
                dateTime += "/";
                dateTime += day;
            }
            dateTime += "/";
            dateTime += year;
            DateTime date = DateTime.Parse(dateTime);
            return date;
        }

        //P404
        public static string GetWesternYear(string source)
        {
            Regex r = new Regex("\\d{1,}");
            Match m = r.Match(source);
            string year = m.Value[0].ToString();
            string currentYear = (int.Parse(DateTime.Today.ToString("yyyy"))-1988).ToString(); //平成年
            year = currentYear.Replace(currentYear[currentYear.Length - 1], year[0]); // potential bug: TO BE IMPROVED
            int yearInt = int.Parse(year);
            yearInt += 1988;//平成年
            return yearInt.ToString();
        }

        public static string GetChineseVersionNum(int number)
        {
            string[] numberChineseVersionArr = { "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二" };
            string ChineseVersionNum = string.Empty;
            if (number > 0 && number < 13)
            {
                ChineseVersionNum = numberChineseVersionArr[number - 1];
            }
            return ChineseVersionNum;
        }


        public static void GenerateSecurityOptionTargetFile(string filePath, List<SecurityOptionSector> sectorList, string contentTitle, int rowNumEachSector, int startPageNum, int maxPageNum, string pageNamePre, DateTime lastBusinessDay)
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                GenerateSecurityOptionTargetFile(worksheet, sectorList, contentTitle, rowNumEachSector, startPageNum, maxPageNum, pageNamePre, lastBusinessDay);
                workbook.Save();
            }
        }

        public static void GenerateSecurityOptionTargetFile(Worksheet worksheet, List<SecurityOptionSector> sectorList,
            string contentTitle, int rowNumEachSector, int startPageNum, int maxPageNum, string pageNamePre,
            DateTime lastBusinessDay)
        {
            worksheet.Cells[1, 1] = contentTitle;
            int pageNum = startPageNum;
            string pageName = "";
            bool isFirstPage = true;
            int curCol = 1;
            for (int i = 0; i < sectorList.Count; i++)
            {
                // No page T/OPT20, T/OPT30
                if (pageNum%10 == 0)
                {
                    pageNum++;
                }
                pageName = string.Format("{0}{1}", pageNamePre, pageNum);

                curCol = i % 2 == 0 ? 1 : 9;

                WriteSecurityOptionSector(worksheet, sectorList[i], pageName, (i/2) * (rowNumEachSector + 5) + 1, curCol,
                    lastBusinessDay);
                if (i % 2 != 0 && i < sectorList.Count - 1)
                {
                    pageNum++;
                }
            }
            if (sectorList.Count%2 == 1)
            {
                WriteEmptySecurityOptionSector(worksheet, pageName, (sectorList.Count/2)*(rowNumEachSector + 5) + 1, 9,
                    lastBusinessDay, isFirstPage);
                isFirstPage = false;
            }
            int skipNum = 0;
            pageNum++;
            while (pageNum <= maxPageNum)
            {
                if (pageNum%10 == 0)
                {
                    skipNum++;
                    pageNum++;
                }
                int curRow = (rowNumEachSector + 5)*(pageNum - startPageNum - skipNum) + 1;
                pageName = string.Format("{0}{1}", pageNamePre, pageNum);
                curCol = (curCol == 1 ? 9 : 1);
                WriteEmptySecurityOptionSector(worksheet, pageName, curRow, curCol, lastBusinessDay, isFirstPage);
                isFirstPage = false;
                curCol = (curCol == 1 ? 9 : 1);
                WriteEmptySecurityOptionSector(worksheet, pageName, curRow, curCol, lastBusinessDay, isFirstPage);
                pageNum++;
            }
            FormatSecurityOptionTargetFile(worksheet);
        }

        public static void WriteEmptySecurityOptionSector(Worksheet worksheet, string pageName, int curRow, int curCol, DateTime lastBusinessDay, bool isFirstPage)
        {
            ExcelUtil.GetRange(curRow, curCol, curRow + 25, curCol + 7, worksheet).NumberFormat = "@";
            worksheet.Cells[curRow, curCol] = pageName;
            ExcelUtil.GetRange(curRow, curCol, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightSkyBlue);
            ExcelUtil.GetRange(curRow, curCol, worksheet).Font.Bold = true;
            worksheet.Cells[curRow + 1, curCol + 2] = lastBusinessDay.ToString("ddMMMyy").ToUpper();

            worksheet.Cells[curRow + 3, curCol + 1] = "SELL";
            ExcelUtil.GetRange(curRow + 3, curCol + 1, curRow + 3, curCol + 3, worksheet).MergeCells = true;
            ExcelUtil.GetRange(curRow + 3, curCol + 1, curRow + 3, curCol + 3, worksheet).HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Cells[curRow + 3, curCol + 5] = "BUY";
            ExcelUtil.GetRange(curRow + 3, curCol + 4, curRow + 3, curCol + 6, worksheet).MergeCells = true;
            ExcelUtil.GetRange(curRow + 3, curCol + 4, curRow + 3, curCol + 6, worksheet).HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ExcelUtil.GetRange(curRow + 3, curCol + 1, curRow + 3, curCol + 6, worksheet).Font.Size = 13;
            ExcelUtil.GetRange(curRow + 3, curCol + 1, curRow + 3, curCol + 6, worksheet).Font.Bold = true;

            worksheet.Cells[curRow + 4, curCol + 1] = "JAPANESE";
            worksheet.Cells[curRow + 4, curCol + 2] = "ENGLISH";
            worksheet.Cells[curRow + 4, curCol + 3] = "OP/INT";

            worksheet.Cells[curRow + 4, curCol + 4] = "JAPANESE";
            worksheet.Cells[curRow + 4, curCol + 5] = "ENGLISH";
            worksheet.Cells[curRow + 4, curCol + 6] = "OP/INT";

            if (isFirstPage)
            {
                worksheet.Cells[curRow + 5, curCol + 1] = "出来ず";
                worksheet.Cells[curRow + 5, curCol + 2] = "UNQ";
                worksheet.Cells[curRow + 5, curCol + 4] = "出来ず";
                worksheet.Cells[curRow + 5, curCol + 5] = "UNQ";
            }

            ExcelUtil.GetRange(curRow + 4, curCol + 1, curRow + 4, curCol + 6, worksheet).Font.Size = 13;
            ExcelUtil.GetRange(curRow + 4, curCol + 1, curRow + 4, curCol + 6, worksheet).Font.Bold = true;
            ExcelUtil.GetRange(curRow + 4, curCol + 1, curRow + 4, curCol + 6, worksheet).Font.Size = 13;
            ExcelUtil.GetRange(curRow + 4, curCol + 1, curRow + 4, curCol + 6, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            
        }


        public static void WriteSecurityOptionSector(Worksheet worksheet, SecurityOptionSector securityOptionSector, string pageName, int curRow, int curCol, DateTime lastBusinessDay)
        {
            ExcelUtil.GetRange(curRow, curCol, curRow + 25, curCol + 7, worksheet).NumberFormat = "@";
            worksheet.Cells[curRow, curCol] = pageName;
            ExcelUtil.GetRange(curRow, curCol, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightSkyBlue);
            ExcelUtil.GetRange(curRow, curCol, worksheet).Font.Bold = true;
            worksheet.Cells[curRow + 1, curCol + 1] = securityOptionSector.TradeDate.ToString("MMMyy");
            worksheet.Cells[curRow + 1, curCol + 2] = lastBusinessDay.ToString("ddMMMyy").ToUpper();

            if (!string.IsNullOrEmpty(securityOptionSector.TradedCompany.CompanyInfo.EnglishName))
            {
                worksheet.Cells[curRow + 2, curCol + 1] = securityOptionSector.TradedCompany.CompanyInfo.EnglishName.ToUpper();
            }

            if (!string.IsNullOrEmpty(securityOptionSector.TradedCompany.CompanyInfo.Ric))
            {
                worksheet.Cells[curRow + 2, curCol + 2] = securityOptionSector.TradedCompany.CompanyInfo.Ric;
            }
            worksheet.Cells[curRow + 2, curCol + 3] = securityOptionSector.TradedCompany.CompanyInfo.OriginalName;
            worksheet.Cells[curRow + 2, curCol + 4] = securityOptionSector.Type;
            worksheet.Cells[curRow + 2, curCol + 5] = securityOptionSector.TradedCompany.TransSum;

            worksheet.Cells[curRow + 3, curCol + 1] = "SELL";
            ExcelUtil.GetRange(curRow + 3, curCol + 1, curRow + 3, curCol + 3, worksheet).MergeCells = true;
            ExcelUtil.GetRange(curRow + 3, curCol + 1, curRow + 3, curCol + 3, worksheet).HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Cells[curRow + 3, curCol + 5] = "BUY";
            ExcelUtil.GetRange(curRow + 3, curCol + 4, curRow + 3, curCol + 6, worksheet).MergeCells = true;
            ExcelUtil.GetRange(curRow + 3, curCol + 4, curRow + 3, curCol + 6, worksheet).HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ExcelUtil.GetRange(curRow + 3, curCol + 1, curRow + 3, curCol + 6, worksheet).Font.Size = 13;
            ExcelUtil.GetRange(curRow + 3, curCol + 1, curRow + 3, curCol + 6, worksheet).Font.Bold = true;

            worksheet.Cells[curRow + 4, curCol + 1] = "JAPANESE";
            worksheet.Cells[curRow + 4, curCol + 2] = "ENGLISH";
            worksheet.Cells[curRow + 4, curCol + 3] = "OP/INT";

            worksheet.Cells[curRow + 4, curCol + 4] = "JAPANESE";
            worksheet.Cells[curRow + 4, curCol + 5] = "ENGLISH";
            worksheet.Cells[curRow + 4, curCol + 6] = "OP/INT";

            ExcelUtil.GetRange(curRow + 4, curCol + 1, curRow + 4, curCol + 6, worksheet).Font.Size = 13;
            ExcelUtil.GetRange(curRow + 4, curCol + 1, curRow + 4, curCol + 6, worksheet).Font.Bold = true;
            ExcelUtil.GetRange(curRow + 4, curCol + 1, curRow + 4, curCol + 6, worksheet).Font.Size = 13;
            ExcelUtil.GetRange(curRow + 4, curCol + 1, curRow + 4, curCol + 6, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            WriteTransCompanyInfosForSecurityOption(worksheet, curRow+5, curCol+1, securityOptionSector);     
        }

        public static void FormatSecurityOptionTargetFile(Worksheet worksheet)
        {
            ((Range)worksheet.Columns["A:O", Type.Missing]).ColumnWidth = 14;
            ((Range)worksheet.Columns["A:O", Type.Missing]).Font.Name = "明朝";
            worksheet.UsedRange.NumberFormat = "@";
        }

        /// <summary>
        /// Write a sector for Security Option transaction
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="startRow"></param>
        /// <param name="oseTransSector"></param>
        public static void WriteTransCompanyInfosForSecurityOption(Worksheet worksheet, int startRow, int curCol, SecurityOptionSector securityOptionTransSector)
        {
            int currentRow = startRow;
            if (securityOptionTransSector.SellCompanyInfoList.Count == 0)
            {
                worksheet.Cells[currentRow, curCol+1] = "UNQ";
                worksheet.Cells[currentRow, curCol] = "出来ず";
            }
            if (securityOptionTransSector.BuyCompanyInfoList.Count == 0)
            {
                worksheet.Cells[currentRow, curCol+4] = "UNQ";
                worksheet.Cells[currentRow, curCol+3] = "出来ず";
            }
            using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, currentRow, curCol, ExcelLineWriter.Direction.Right))
            {
                for (int i = 0; i < securityOptionTransSector.SellCompanyInfoList.Count; i++)
                {
                    OSETransSingleCompanyInfo singleCompanyInfo = securityOptionTransSector.SellCompanyInfoList[i];
                    writeCompanyInfoForSecurityOption(worksheet, writer, currentRow, singleCompanyInfo);
                    writer.PlaceNext(startRow+i+1 , curCol);
                }

                writer.PlaceNext(startRow, curCol+3);

                for (int i = 0; i < securityOptionTransSector.BuyCompanyInfoList.Count; i++)
                {
                    OSETransSingleCompanyInfo singleCompanyInfo = securityOptionTransSector.BuyCompanyInfoList[i];
                    writeCompanyInfoForSecurityOption(worksheet, writer, startRow, singleCompanyInfo);
                    writer.PlaceNext(startRow+i+1, curCol + 3);
                }
            }
        }

        public static void writeCompanyInfoForSecurityOption(Worksheet worksheet, ExcelLineWriter writer, int startRow, OSETransSingleCompanyInfo singleCompanyInfo)
        {
            if (singleCompanyInfo != null && singleCompanyInfo.CompanyInfo != null)
            {
                if (!string.IsNullOrEmpty(singleCompanyInfo.CompanyInfo.JapaneseName))
                {
                    writer.WriteLine(singleCompanyInfo.CompanyInfo.JapaneseName);
                    ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }
                else
                {
                    writer.WriteLine(singleCompanyInfo.CompanyInfo.OriginalName);
                    ExcelUtil.GetRange(writer.Row, writer.Col - 1, writer.Row, writer.Col, worksheet).Font.Color = ColorTranslator.ToOle(Color.Red);
                    ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }

                if (!string.IsNullOrEmpty(singleCompanyInfo.CompanyInfo.EnglishName))
                {
                    writer.WriteLine(singleCompanyInfo.CompanyInfo.EnglishName);
                    ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }
                else
                {
                    writer.WriteLine(singleCompanyInfo.CompanyInfo.OriginalName);
                    ExcelUtil.GetRange(writer.Row, writer.Col - 1, writer.Row, writer.Col, worksheet).Font.Color = ColorTranslator.ToOle(Color.Red);
                    ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }

                writer.WriteLine(singleCompanyInfo.Volume_OP_INT.Replace(",", ""));
                ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignRight;

                writer.PlaceNext(writer.Row + 1, writer.Col);
            }
        }

        /// <summary>
        /// Write a sector for OSE transaction
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="startRow"></param>
        /// <param name="oseTransSector"></param>
        public static void WriteTransCompanyInfoSector(Worksheet worksheet, int startRow, OSETransSector oseTransSector, bool isFuture)
        {
            int currentRow = startRow;
            if (!isFuture)
            {
                if (oseTransSector.SellCompanyInfoList.Count == 0)
                {
                    worksheet.Cells[currentRow, 2] = "出来ず";
                    worksheet.Cells[currentRow, 3] = "UNQ";
                }
                if (oseTransSector.BuyCompanyInfoList.Count == 0)
                {
                    worksheet.Cells[currentRow, 5] = "出来ず";
                    worksheet.Cells[currentRow, 6] = "UNQ";
                }
            }
            else
            {
                if (oseTransSector.SellCompanyInfoList.Count == 0)
                {
                    worksheet.Cells[currentRow, 3] = "UNQ";
                    worksheet.Cells[currentRow, 4] = "出来ず";
                }
                if (oseTransSector.BuyCompanyInfoList.Count == 0)
                {
                    worksheet.Cells[currentRow, 7] = "UNQ";
                    worksheet.Cells[currentRow, 8] = "出来ず";
                }
            }

            using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, currentRow, 2, ExcelLineWriter.Direction.Right))
            {
                foreach (OSETransSingleCompanyInfo singleCompanyInfo in oseTransSector.SellCompanyInfoList)
                {
                    if (singleCompanyInfo != null && singleCompanyInfo.CompanyInfo != null)
                    {
                        if (!string.IsNullOrEmpty(singleCompanyInfo.TransSum))
                        {
                            writer.WriteLine(singleCompanyInfo.TransSum.Replace(",",""));
                            ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignRight;
                        }
                        else
                        {
                            writer.WriteLine(singleCompanyInfo.CompanyInfo.OriginalName);
                            ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        }
                        if (string.IsNullOrEmpty(singleCompanyInfo.CompanyInfo.EnglishName))
                        {
                            writer.WriteLine(singleCompanyInfo.CompanyInfo.OriginalName);
                            ExcelUtil.GetRange(writer.Row, writer.Col - 1, writer.Row, writer.Col, worksheet).Font.Color = ColorTranslator.ToOle(Color.Red);
                            ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        }
                        else
                        {
                            writer.WriteLine(singleCompanyInfo.CompanyInfo.EnglishName);
                            ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        }
                        writer.WriteLine(singleCompanyInfo.CompanyInfo.JapaneseName);
                        ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;

                        writer.WriteLine(singleCompanyInfo.Volume_OP_INT.Replace(",",""));
                        ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignRight;

                        writer.PlaceNext(writer.Row + 1, 2);
                    }
                    ExcelUtil.GetRange(startRow, 5, startRow + oseTransSector.SellCompanyInfoList.Count, 9,worksheet).HorizontalAlignment = XlHAlign.xlHAlignRight;
                }
                writer.PlaceNext(currentRow, 6);
                foreach (OSETransSingleCompanyInfo singleCompanyInfo in oseTransSector.BuyCompanyInfoList.Where(singleCompanyInfo => singleCompanyInfo != null && singleCompanyInfo.CompanyInfo != null))
                {
                    if (!string.IsNullOrEmpty(singleCompanyInfo.TransSum))
                    {
                        writer.WriteLine(singleCompanyInfo.TransSum.Replace(",", ""));
                        ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignRight;
                    }
                    else
                    {
                        writer.WriteLine(singleCompanyInfo.CompanyInfo.OriginalName);
                        ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }
                    if (string.IsNullOrEmpty(singleCompanyInfo.CompanyInfo.EnglishName))
                    {
                        writer.WriteLine(singleCompanyInfo.CompanyInfo.OriginalName);
                        ExcelUtil.GetRange(writer.Row, writer.Col - 1, writer.Row, writer.Col, worksheet).Font.Color = ColorTranslator.ToOle(Color.Red);
                        ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }
                    else
                    {
                        writer.WriteLine(singleCompanyInfo.CompanyInfo.EnglishName);
                        ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }
                    writer.WriteLine(singleCompanyInfo.CompanyInfo.JapaneseName);
                    ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    writer.WriteLine(singleCompanyInfo.Volume_OP_INT.Replace(",",""));
                    ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet).HorizontalAlignment = XlHAlign.xlHAlignRight;

                    writer.PlaceNext(writer.Row + 1, 6);
                }
            }
        }

        public static void FormatCompanyInfoSector(Worksheet worksheet, int startRow, int maxRow, int startCol)
        {
            ExcelUtil.GetRange(startRow, startCol, startRow + maxRow, startCol, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ExcelUtil.GetRange(startRow, startCol + 1, startRow + maxRow, startCol + 1, worksheet).Interior.Color = ColorTranslator.ToOle(Color.YellowGreen);
            ExcelUtil.GetRange(startRow, startCol + 2, startRow + maxRow, startCol + 2, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
            ExcelUtil.GetRange(startRow, startCol + 3, startRow + maxRow, startCol + 3, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ExcelUtil.GetRange(startRow, startCol + 4, startRow + maxRow, startCol + 4, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ExcelUtil.GetRange(startRow, startCol + 5, startRow + maxRow, startCol + 5, worksheet).Interior.Color = ColorTranslator.ToOle(Color.YellowGreen);
            ExcelUtil.GetRange(startRow, startCol + 6, startRow + maxRow, startCol + 6, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
            ExcelUtil.GetRange(startRow, startCol + 7, startRow + maxRow, startCol + 7, worksheet).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Range)worksheet.Columns["A:I", Type.Missing]).ColumnWidth = 15;
            ((Range)worksheet.Columns["A:I", Type.Missing]).Font.Name = "明朝";
            ExcelUtil.GetRange(startRow, startCol, startRow + maxRow, startCol + 7, worksheet).NumberFormat = "@";
            for (int i = 1; i <= maxRow; i++)
            {
                worksheet.Cells[i+startRow-1, 1] = i.ToString();
            }
        }

        #region Email Related

        /// <summary>
        ///  Get source files according to the file name keyword
        /// </summary>
        /// <param name="item"></param>
        /// <param name="keyword"></param>
        /// <param name="fileSavedDir"></param>
        /// <returns></returns>
        public static string GetTargetAttachedFile(MailItem item, string keyword, string fileSavedDir)
        {
            string sourceFile = string.Empty;
            try
            {
                List<string> attachedFileList = OutlookUtil.DownloadAttachments(item, keyword, null, fileSavedDir);
                if (attachedFileList == null || attachedFileList.Count == 0)
                {
                    throw new System.Exception(string.Format("Can't find the file name of which contains {0}", keyword));
                }
                if (attachedFileList.Count > 1)
                {
                    throw new System.Exception(string.Format("There're more than 1 files found name of which contails {0}", keyword));
                }
                sourceFile = attachedFileList[0];
            }
            catch (System.Exception ex)
            {
                throw new System.Exception(string.Format("Error found: {0}", ex.Message));
            }
            return sourceFile;
        }
        #endregion
    }
}
