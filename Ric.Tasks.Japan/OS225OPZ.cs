using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Ric.Util;

namespace Ric.Tasks.Japan
{
    class OS225OPZ
    {
        public static void GetOS225OPZData(Workbook book, ref OS225OPZOriginData data)
        {
            if (book == null || data == null)
            {
                return;
            }
            Worksheet sheet = book.Worksheets[1] as Worksheet;
            if (sheet == null)
            {
                return;
            }
            string updateDate = ExcelUtil.GetRange("A2", sheet).Text.ToString();
            data.updateDate = JapanShared.TransferJpDate(updateDate);

            string putTradingDate = ExcelUtil.GetRange("B4", sheet).Text;
            putTradingDate = putTradingDate.Replace("プット", "");
            data.putTradingDate = JapanShared.TransferJpDate(putTradingDate);

            string callTradingDate = ExcelUtil.GetRange("I4", sheet).Text;
            callTradingDate = callTradingDate.Replace("コール", "");
            data.callTradingDate = JapanShared.TransferJpDate(callTradingDate);

            int offset = 17;
            for (int i = 0; i < 5; i++ )
            {
                string exchangeCount = ExcelUtil.GetRange(6 + i * offset, 1, sheet).Text.ToString();
                exchangeCount = exchangeCount.Replace("(", "").Replace(")", "").Replace("円", "").Replace(",","").Trim();
                data.putExchangeCount.Add(exchangeCount);

                exchangeCount = ExcelUtil.GetRange(6 + i * offset, 8, sheet).Text.ToString();
                exchangeCount = exchangeCount.Replace("(", "").Replace(")", "").Replace("円", "").Replace(",", "").Trim();
                data.callExchangeCount.Add(exchangeCount);

                for (int j = 0; j < 15; j++ )
                {
                    OSETradeCompanyInfo Info = OS225FUU.GetTradingInfo(sheet, 8 + i * offset + j,2);
                    data.putInfo.Add(Info);

                    Info = OS225FUU.GetTradingInfo(sheet, 8 + i * offset + j, 4);
                    data.putInfo.Add(Info);

                    Info = OS225FUU.GetTradingInfo(sheet, 8 + i * offset + j, 9);
                    data.callInfo.Add(Info);

                    Info = OS225FUU.GetTradingInfo(sheet, 8 + i * offset + j, 11);
                    data.callInfo.Add(Info);
                }
                

                
            }
        }

        private static void WriteOS225OPZPUT(Workbook book, OS225OPZOriginData data)
        {
            if (book == null || data == null)
            {
                return;
            }
            Worksheet sheet = book.Worksheets[1] as Worksheet;
            sheet.Cells[3, 2] = data.putTradingDate.ToString("MMM").ToUpper() + data.putTradingDate.ToString("yy");
            sheet.Cells[3, 3] = data.putTradingDate.ToString("yy");
            sheet.Cells[3, 4] = data.putTradingDate.ToString("MM");

            sheet.Cells[3, 6] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
            sheet.Cells[3, 7] = data.updateDate.ToString("yyyy") + "年" + data.updateDate.ToString("MM") + "月" + data.updateDate.ToString("dd") + "日";

            int offset = 17;
            int index = 0;
            for (int i = 0; i < 5; i++ )
            {
                for (int j = 0; j < 15;j++ )
                {
                    sheet.Cells[5 + i * offset + j, 2] = data.putInfo[index].OriginalName;
                    sheet.Cells[5 + i * offset + j, 3] = data.putInfo[index].EnglishName;
                    sheet.Cells[5 + i * offset + j, 4] = data.putInfo[index].JapaneseName;
                    sheet.Cells[5 + i * offset + j, 5] = data.putInfo[index].OPINT;
                    index++;

                    sheet.Cells[5 + i * offset + j, 6] = data.putInfo[index].OriginalName;
                    sheet.Cells[5 + i * offset + j, 7] = data.putInfo[index].EnglishName;
                    sheet.Cells[5 + i * offset + j, 8] = data.putInfo[index].JapaneseName;
                    sheet.Cells[5 + i * offset + j, 9] = data.putInfo[index].OPINT;
                    index++;
                }
                

                if (i == 0)
                {
                    sheet.Cells[3, 8] = data.putExchangeCount[i];
                }
                else
                {
                    sheet.Cells[4 + i * offset, 8] = data.putExchangeCount[i];
                }
            }
        }

        private static void WriteOS225OPZCALL(Workbook book, OS225OPZOriginData data)
        {
            if (book == null || data == null)
            {
                return;
            }
            Worksheet sheet = book.Worksheets[2] as Worksheet;
            sheet.Cells[3, 2] = data.callTradingDate.ToString("MMM").ToUpper() + data.callTradingDate.ToString("yy");
            sheet.Cells[3, 3] = data.callTradingDate.ToString("yy");
            sheet.Cells[3, 4] = data.callTradingDate.ToString("MM");

            sheet.Cells[3, 6] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
            sheet.Cells[3, 7] = data.updateDate.ToString("yyyy") + "年" + data.updateDate.ToString("MM") + "月" + data.updateDate.ToString("dd") + "日";

            int offset = 17;
            int index = 0;
            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 15; j++ )
                {
                    sheet.Cells[5 + i * offset + j, 2] = data.callInfo[index].OriginalName;
                    sheet.Cells[5 + i * offset + j, 3] = data.callInfo[index].EnglishName;
                    sheet.Cells[5 + i * offset + j, 4] = data.callInfo[index].JapaneseName;
                    sheet.Cells[5 + i * offset + j, 5] = data.callInfo[index].OPINT;
                    index++;

                    sheet.Cells[5 + i * offset + j, 6] = data.callInfo[index].OriginalName;
                    sheet.Cells[5 + i * offset + j, 7] = data.callInfo[index].EnglishName;
                    sheet.Cells[5 + i * offset + j, 8] = data.callInfo[index].JapaneseName;
                    sheet.Cells[5 + i * offset + j, 9] = data.callInfo[index].OPINT;
                    index++;
                }
               

                if (i == 0)
                {
                    sheet.Cells[3, 8] = data.callExchangeCount[i];
                }
                else
                {
                    sheet.Cells[4 + i * offset, 8] = data.callExchangeCount[i];
                }
            }
        }

        public static void GenerateOS225OPZ(Workbook book, OS225OPZOriginData data)
        {
            WriteOS225OPZPUT(book,data);
            WriteOS225OPZCALL(book,data);
        }
    }

    public class OS225OPZOriginData
    {
        public DateTime updateDate { get; set; }
        public DateTime putTradingDate { get; set; }
        public DateTime callTradingDate { get; set; }
        public List<string> putExchangeCount = new List<string>();
        public List<string> callExchangeCount = new List<string>();
        public List<OSETradeCompanyInfo> putInfo = new List<OSETradeCompanyInfo>();
        public List<OSETradeCompanyInfo> callInfo = new List<OSETradeCompanyInfo>();
    }
}
