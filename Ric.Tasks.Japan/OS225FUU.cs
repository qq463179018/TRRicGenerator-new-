using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Ric.Util;

namespace Ric.Tasks.Japan
{
    class OS225FUU
    {
        public static OSETradeCompanyInfo GetTradingInfo(Worksheet sheet, int row,int col)
        {
            OSETradeCompanyInfo companyInfo = new OSETradeCompanyInfo();
            companyInfo.OriginalName = ExcelUtil.GetRange(row, col, sheet).Text.ToString();
            companyInfo.OPINT = ExcelUtil.GetRange(row, col+1, sheet).Text.ToString().Replace(",","");
            return companyInfo;
        }

        

        

        public static void WriteOS225FUU(Workbook book, OS225FUUOriginData data)
        {

            Worksheet sheet = book.Worksheets[1] as Worksheet;
            sheet.Cells[2, 1] = data.updateDate.ToString("yyyy");
            
            sheet.Cells[2, 2] = data.updateDate.ToString("MM");
            sheet.Cells[2, 3] = data.updateDate.ToString("dd");

            sheet.Cells[2, 8] = data.updateDate.ToString("dd");
            sheet.Cells[2, 9] = data.updateDate.ToString("MMM").ToUpper();
            sheet.Cells[2, 10] = data.updateDate.ToString("yy");

            sheet.Cells[3, 2] = data.OS225FUU.tradingDate1.ToString("MMM").ToUpper() + data.OS225FUU.tradingDate1.ToString("yy");
            sheet.Cells[3, 3] = data.OS225FUU.tradingDate1.ToString("yy");
            sheet.Cells[3, 5] = data.OS225FUU.tradingDate1.ToString("MM");

            int currentRow = 5;
            int currentInfo = 0;
            foreach (OSETradeCompanyInfo item in data.OS225FUU.info1)
            {
                currentInfo++;
                if (currentInfo % 2 == 1)
                {
                    sheet.Cells[currentRow, 2] = item.OriginalName;
                    sheet.Cells[currentRow, 3] = item.EnglishName;
                    sheet.Cells[currentRow, 4] = item.JapaneseName;
                    sheet.Cells[currentRow, 5] = item.OPINT;
                }
                else
                {
                    sheet.Cells[currentRow, 6] = item.OriginalName;
                    sheet.Cells[currentRow, 7] = item.EnglishName;
                    sheet.Cells[currentRow, 8] = item.JapaneseName;
                    sheet.Cells[currentRow, 9] = item.OPINT;
                    currentRow++;
                }

            }

            sheet.Cells[21, 2] = data.OS225FUU.tradingDate2.ToString("MMM").ToUpper() + data.OS225FUU.tradingDate2.ToString("yy");
            sheet.Cells[21, 3] = data.OS225FUU.tradingDate2.ToString("yy");
            sheet.Cells[21, 5] = data.OS225FUU.tradingDate2.ToString("MM");

            currentRow = 23;
            currentInfo = 0;
            foreach (OSETradeCompanyInfo item in data.OS225FUU.info2)
            {
                currentInfo++;
                if (currentInfo % 2 == 1)
                {
                    sheet.Cells[currentRow, 2] = item.OriginalName;
                    sheet.Cells[currentRow, 3] = item.EnglishName;
                    sheet.Cells[currentRow, 4] = item.JapaneseName;
                    sheet.Cells[currentRow, 5] = item.OPINT;
                }
                else
                {
                    sheet.Cells[currentRow, 6] = item.OriginalName;
                    sheet.Cells[currentRow, 7] = item.EnglishName;
                    sheet.Cells[currentRow, 8] = item.JapaneseName;
                    sheet.Cells[currentRow, 9] = item.OPINT;
                    currentRow++;
                }

            }

        }

        public static void WriteOS225FUX(Workbook book, OS225FUUOriginData data)
        {
            Worksheet sheet = book.Worksheets[2] as Worksheet;
            sheet.Cells[2, 1] = data.updateDate.ToString("yyyy");
           
            sheet.Cells[2, 2] = data.updateDate.ToString("MM");
            sheet.Cells[2, 3] = data.updateDate.ToString("dd");

            sheet.Cells[2, 8] = data.updateDate.ToString("dd");
            sheet.Cells[2, 9] = data.updateDate.ToString("MMM").ToUpper();
            sheet.Cells[2, 10] = data.updateDate.ToString("yy");

            sheet.Cells[3, 2] = data.OS225FUX.tradingDate1.ToString("MMM").ToUpper() + data.OS225FUX.tradingDate1.ToString("yy");
            sheet.Cells[3, 3] = data.OS225FUX.tradingDate1.ToString("yy");
            sheet.Cells[3, 5] = data.OS225FUX.tradingDate1.ToString("MM");

            int currentRow = 5;
            int currentInfo = 0;
            foreach (OSETradeCompanyInfo item in data.OS225FUX.info1)
            {
                currentInfo++;
                if (currentInfo % 2 == 1)
                {
                    sheet.Cells[currentRow, 2] = item.OriginalName;
                    sheet.Cells[currentRow, 3] = item.EnglishName;
                    sheet.Cells[currentRow, 4] = item.JapaneseName;
                    sheet.Cells[currentRow, 5] = item.OPINT;
                }
                else
                {
                    sheet.Cells[currentRow, 6] = item.OriginalName;
                    sheet.Cells[currentRow, 7] = item.EnglishName;
                    sheet.Cells[currentRow, 8] = item.JapaneseName;
                    sheet.Cells[currentRow, 9] = item.OPINT;
                    currentRow++;
                }

            }

            sheet.Cells[21, 2] = data.OS225FUX.tradingDate2.ToString("MMM").ToUpper() + data.OS225FUX.tradingDate2.ToString("yy");
            sheet.Cells[21, 3] = data.OS225FUX.tradingDate2.ToString("yy");
            sheet.Cells[21, 5] = data.OS225FUX.tradingDate2.ToString("MM");

            currentRow = 23;
            currentInfo = 0;
            foreach (OSETradeCompanyInfo item in data.OS225FUX.info2)
            {
                currentInfo++;
                if (currentInfo % 2 == 1)
                {
                    sheet.Cells[currentRow, 2] = item.OriginalName;
                    sheet.Cells[currentRow, 3] = item.EnglishName;
                    sheet.Cells[currentRow, 4] = item.JapaneseName;
                    sheet.Cells[currentRow, 5] = item.OPINT;
                }
                else
                {
                    sheet.Cells[currentRow, 6] = item.OriginalName;
                    sheet.Cells[currentRow, 7] = item.EnglishName;
                    sheet.Cells[currentRow, 8] = item.JapaneseName;
                    sheet.Cells[currentRow, 9] = item.OPINT;
                    currentRow++;
                }

            }
        }

        public static void WriteOSTRADF(Workbook book, OS225FUUOriginData data)
        {
            Worksheet sheet = book.Worksheets[3] as Worksheet;
            sheet.Cells[2, 1] = data.updateDate.ToString("yyyy");
           
            sheet.Cells[2, 2] = data.updateDate.ToString("MM");
            sheet.Cells[2, 3] = data.updateDate.ToString("dd");

            sheet.Cells[2, 8] = data.updateDate.ToString("dd");
            sheet.Cells[2, 9] = data.updateDate.ToString("MMM").ToUpper();
            sheet.Cells[2, 10] = data.updateDate.ToString("yy");

            sheet.Cells[3, 2] = data.OSTRADF.tradingDate1.ToString("MMM").ToUpper() + data.OSTRADF.tradingDate1.ToString("yy");
            sheet.Cells[3, 3] = data.OSTRADF.tradingDate1.ToString("yy");
            sheet.Cells[3, 5] = data.OSTRADF.tradingDate1.ToString("MM");

            int currentRow = 5;
            int currentInfo = 0;
            foreach (OSETradeCompanyInfo item in data.OSTRADF.info1)
            {
                currentInfo++;
                if (currentInfo % 2 == 1)
                {
                    sheet.Cells[currentRow, 2] = item.OriginalName;
                    sheet.Cells[currentRow, 3] = item.EnglishName;
                    sheet.Cells[currentRow, 4] = item.JapaneseName;
                    sheet.Cells[currentRow, 5] = item.OPINT;
                }
                else
                {
                    sheet.Cells[currentRow, 6] = item.OriginalName;
                    sheet.Cells[currentRow, 7] = item.EnglishName;
                    sheet.Cells[currentRow, 8] = item.JapaneseName;
                    sheet.Cells[currentRow, 9] = item.OPINT;
                    currentRow++;
                }

            }

            sheet.Cells[21, 2] = data.OSTRADF.tradingDate2.ToString("MMM").ToUpper() + data.OSTRADF.tradingDate2.ToString("yy");
            sheet.Cells[21, 3] = data.OSTRADF.tradingDate2.ToString("yy");
            sheet.Cells[21, 5] = data.OSTRADF.tradingDate2.ToString("MM");

            currentRow = 23;
            currentInfo = 0;
            foreach (OSETradeCompanyInfo item in data.OSTRADF.info2)
            {
                currentInfo++;
                if (currentInfo % 2 == 1)
                {
                    sheet.Cells[currentRow, 2] = item.OriginalName;
                    sheet.Cells[currentRow, 3] = item.EnglishName;
                    sheet.Cells[currentRow, 4] = item.JapaneseName;
                    sheet.Cells[currentRow, 5] = item.OPINT;
                }
                else
                {
                    sheet.Cells[currentRow, 6] = item.OriginalName;
                    sheet.Cells[currentRow, 7] = item.EnglishName;
                    sheet.Cells[currentRow, 8] = item.JapaneseName;
                    sheet.Cells[currentRow, 9] = item.OPINT;
                    currentRow++;
                }

            }
        }
    }
}
