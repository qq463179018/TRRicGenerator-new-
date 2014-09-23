using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Ric.Util;

namespace Ric.Tasks.Japan
{
    class OSOPT35_46
    {
        private static string GetExchangeCount2(Worksheet sheet, int row, int col)
        {
            string count = ExcelUtil.GetRange(row, col, sheet).Text.ToString();
            count = count.Replace(",", "").Replace(")", "").Replace("円", "");
            if (count.Length > 3)
            {
                count = count.Substring(0, 3);
            }
            return count;
        }

        private static void updateCompanyInfo(OSETradeCompanyInfo companyInfo, Dictionary<string, NameMap> nameDic)
        {
            foreach (var item in nameDic)
            {
                if (item.Key == companyInfo.OriginalName.Trim())
                {
                    companyInfo.JapaneseName = item.Value.JapaneseName;
                    companyInfo.EnglishName = item.Value.EnglistName;
                    break;
                }
            }
        }

        private static string GetPutOrCall(Worksheet sheet, int row, int col)
        {
            string tmp = ExcelUtil.GetRange(row, col, sheet).Text.ToString();
            if (tmp.Equals("プット"))
            {
                tmp = "PUT";
            }
            else if (tmp.Equals("コール"))
            {
                tmp = "CALL";
            }
            return tmp;
        }

        public static OSETradeCompanyInfo GetExchangeName(Worksheet sheet, int row, int col)
        {
            OSETradeCompanyInfo tmp = new OSETradeCompanyInfo();
            tmp.OriginalName = ExcelUtil.GetRange(row, col, sheet).Text.ToString();
            return tmp;
        }

        private static OSETradeCompanyInfo GetTradingInfo(Worksheet sheet, int row, int col,int offset)
        {
            OSETradeCompanyInfo companyInfo = new OSETradeCompanyInfo();
            companyInfo.OriginalName = ExcelUtil.GetRange(row, col, sheet).Text.ToString();
            companyInfo.OPINT = ExcelUtil.GetRange(row, col + offset, sheet).Text.ToString().Replace(",", "");
            return companyInfo;
        }

        public static void GetOSOPT35_46Data(Workbook book, ref OSOPT35_46OriginData data, Dictionary<string, NameMap> nameDic)
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

            int offset = 19;
            for (int i = 0; i < 10; i++ )
            {
                string count = ExcelUtil.GetRange(6 + i * offset, 3, sheet).Text.ToString();
                data.code.Add(count);
                count = ExcelUtil.GetRange(6 + i * offset, 12, sheet).Text.ToString();
                data.code.Add(count);

                count = GetExchangeCount2(sheet, 7 + i * offset, 7);                
                data.exchangeCount2.Add(count);
                count = GetExchangeCount2(sheet, 7 + i * offset, 16);
                data.exchangeCount2.Add(count);

                string tmp = GetPutOrCall(sheet,6 + i * offset,7);
                data.putOrCall.Add(tmp);
                tmp = GetPutOrCall(sheet, 6 + i * offset, 16);
                data.putOrCall.Add(tmp);

                OSETradeCompanyInfo exchangeName = GetExchangeName(sheet, 6 + i * offset, 3);
                updateCompanyInfo(exchangeName, nameDic);
                data.exchangeName.Add(exchangeName);
                exchangeName = GetExchangeName(sheet, 6 + i * offset, 12);
                updateCompanyInfo(exchangeName, nameDic);
                data.exchangeName.Add(exchangeName);

                string exchangDate = ExcelUtil.GetRange(6 + i * offset, 6, sheet).Text.ToString();
                data.exchangeDate.Add(JapanShared.TransferJpDate(exchangDate));
                exchangDate = ExcelUtil.GetRange(6 + i * offset, 15, sheet).Text.ToString();
                data.exchangeDate.Add(JapanShared.TransferJpDate(exchangDate));

                for (int j = 9; j <= 23; j++ )
                {
                    OSETradeCompanyInfo info = GetTradingInfo(sheet,j+i*offset,3,2);
                    updateCompanyInfo(info, nameDic);
                    data.info.Add(info);

                    info = GetTradingInfo(sheet, j + i * offset, 6, 1);
                    updateCompanyInfo(info, nameDic);
                    data.info.Add(info);

                    info = GetTradingInfo(sheet, j + i * offset, 12, 2);
                    updateCompanyInfo(info, nameDic);
                    data.info.Add(info);

                    info = GetTradingInfo(sheet, j + i * offset, 15, 1);
                    updateCompanyInfo(info, nameDic);
                    data.info.Add(info);
                }
            }

        }

        public static void GenerateOSOPT35_46(Workbook book, OSOPT35_46OriginData data)
        {
            WriteOSOPT35_37(book, data);
            WriteOSOPT38_41(book, data);
            WriteOSOPT42_46(book, data);
        }

        private static void WriteOSOPT35_37(Workbook book, OSOPT35_46OriginData data)
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

            int index = 0;
            int index2 = 0;
            int offset = 22;
            for (int i = 0; i < 3;i++ )
            {
                sheet.Cells[5 + i * offset, 2] = data.exchangeDate[index].ToString("MMM") + data.exchangeDate[index].ToString("yy");
                sheet.Cells[5 + i * offset, 3] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
                sheet.Cells[6 + i * offset, 2] = data.exchangeName[index].EnglishName;
                sheet.Cells[6 + i * offset, 4] = data.exchangeName[index].JapaneseName;
                sheet.Cells[6 + i * offset, 3] = data.code[index];
                sheet.Cells[6 + i * offset, 5] = data.putOrCall[index];
                sheet.Cells[6 + i * offset, 6] = data.exchangeCount2[index];
                index++;

                sheet.Cells[5 + i * offset, 10] = data.exchangeDate[index].ToString("MMM") + data.exchangeDate[index].ToString("yy");
                sheet.Cells[5 + i * offset, 11] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
                sheet.Cells[6 + i * offset, 10] = data.exchangeName[index].EnglishName;
                sheet.Cells[6 + i * offset, 12] = data.exchangeName[index].JapaneseName;
                sheet.Cells[6 + i * offset, 11] = data.code[index];
                sheet.Cells[6 + i * offset, 13] = data.putOrCall[index];
                sheet.Cells[6 + i * offset, 14] = data.exchangeCount2[index];
                index++;

                for (int j = 9; j <= 23; j++ )
                {
                    sheet.Cells[j + i * offset, 2] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 3] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 4] = data.info[index2].OPINT;
                    index2++;

                    sheet.Cells[j + i * offset, 5] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 6] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 7] = data.info[index2].OPINT;
                    index2++;

                    sheet.Cells[j + i * offset, 10] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 11] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 12] = data.info[index2].OPINT;
                    index2++;

                    sheet.Cells[j + i * offset, 13] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 14] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 15] = data.info[index2].OPINT;
                    index2++;
                }

            }
            
        }

        private static void WriteOSOPT38_41(Workbook book, OSOPT35_46OriginData data)
        {
            if (book == null || data == null)
            {
                return;
            }
            Worksheet sheet = book.Worksheets[2] as Worksheet;
            if (sheet == null)
            {
                return;
            }

            int index = 6;
            int index2 = 180;
            int offset = 22;
            for (int i = 0; i < 3; i++)
            {
                sheet.Cells[5 + i * offset, 2] = data.exchangeDate[index].ToString("MMM") + data.exchangeDate[index].ToString("yy");
                sheet.Cells[5 + i * offset, 3] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
                sheet.Cells[6 + i * offset, 2] = data.exchangeName[index].EnglishName;
                sheet.Cells[6 + i * offset, 4] = data.exchangeName[index].JapaneseName;
                sheet.Cells[6 + i * offset, 3] = data.code[index];
                sheet.Cells[6 + i * offset, 5] = data.putOrCall[index];
                sheet.Cells[6 + i * offset, 6] = data.exchangeCount2[index];
                index++;

                sheet.Cells[5 + i * offset, 10] = data.exchangeDate[index].ToString("MMM") + data.exchangeDate[index].ToString("yy");
                sheet.Cells[5 + i * offset, 11] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
                sheet.Cells[6 + i * offset, 10] = data.exchangeName[index].EnglishName;
                sheet.Cells[6 + i * offset, 12] = data.exchangeName[index].JapaneseName;
                sheet.Cells[6 + i * offset, 11] = data.code[index];
                sheet.Cells[6 + i * offset, 13] = data.putOrCall[index];
                sheet.Cells[6 + i * offset, 14] = data.exchangeCount2[index];
                index++;

                for (int j = 9; j <= 23; j++)
                {
                    sheet.Cells[j + i * offset, 2] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 3] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 4] = data.info[index2].OPINT;
                    index2++;

                    sheet.Cells[j + i * offset, 5] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 6] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 7] = data.info[index2].OPINT;
                    index2++;

                    sheet.Cells[j + i * offset, 10] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 11] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 12] = data.info[index2].OPINT;
                    index2++;

                    sheet.Cells[j + i * offset, 13] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 14] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 15] = data.info[index2].OPINT;
                    index2++;
                }

            }
        }
        private static void WriteOSOPT42_46(Workbook book, OSOPT35_46OriginData data)
        {
            if (book == null || data == null)
            {
                return;
            }
            Worksheet sheet = book.Worksheets[3] as Worksheet;
            if (sheet == null)
            {
                return;
            }

            int index = 12;
            int index2 = 360;
            int offset = 22;
            for (int i = 0; i < 4; i++)
            {
                sheet.Cells[5 + i * offset, 2] = data.exchangeDate[index].ToString("MMM") + data.exchangeDate[index].ToString("yy");
                sheet.Cells[5 + i * offset, 3] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
                sheet.Cells[6 + i * offset, 2] = data.exchangeName[index].EnglishName;
                sheet.Cells[6 + i * offset, 4] = data.exchangeName[index].JapaneseName;
                sheet.Cells[6 + i * offset, 3] = data.code[index];
                sheet.Cells[6 + i * offset, 5] = data.putOrCall[index];
                sheet.Cells[6 + i * offset, 6] = data.exchangeCount2[index];
                index++;

                sheet.Cells[5 + i * offset, 10] = data.exchangeDate[index].ToString("MMM") + data.exchangeDate[index].ToString("yy");
                sheet.Cells[5 + i * offset, 11] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
                sheet.Cells[6 + i * offset, 10] = data.exchangeName[index].EnglishName;
                sheet.Cells[6 + i * offset, 12] = data.exchangeName[index].JapaneseName;
                sheet.Cells[6 + i * offset, 11] = data.code[index];
                sheet.Cells[6 + i * offset, 13] = data.putOrCall[index];
                sheet.Cells[6 + i * offset, 14] = data.exchangeCount2[index];
                index++;

                for (int j = 9; j <= 23; j++)
                {
                    sheet.Cells[j + i * offset, 2] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 3] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 4] = data.info[index2].OPINT;
                    index2++;

                    sheet.Cells[j + i * offset, 5] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 6] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 7] = data.info[index2].OPINT;
                    index2++;

                    sheet.Cells[j + i * offset, 10] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 11] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 12] = data.info[index2].OPINT;
                    index2++;

                    sheet.Cells[j + i * offset, 13] = data.info[index2].JapaneseName;
                    sheet.Cells[j + i * offset, 14] = data.info[index2].EnglishName;
                    sheet.Cells[j + i * offset, 15] = data.info[index2].OPINT;
                    index2++;
                }

            }

            sheet.Cells[71, 3] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
            sheet.Cells[71, 11] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
            sheet.Cells[93, 3] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
            sheet.Cells[93, 11] = data.updateDate.ToString("dd") + data.updateDate.ToString("MMM").ToUpper() + data.updateDate.ToString("yy");
        }
    }

    public class OSOPT35_46OriginData
    {
        public DateTime updateDate { get; set; }
        public List<string> code = new List<string>();
        public List<string> exchangeCount2 = new List<string>();
        public List<OSETradeCompanyInfo> exchangeName = new List<OSETradeCompanyInfo>();
        public List<DateTime> exchangeDate = new List<DateTime>();
        public List<string> putOrCall = new List<string>();
        public List<OSETradeCompanyInfo> info = new List<OSETradeCompanyInfo>();
    }
}
