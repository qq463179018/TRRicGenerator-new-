using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;

namespace Ric.Util
{
    public class MiscUtil
    {
        public static string UrlCombine(string p1, string p2)
        {
            return (new Uri(new Uri(p1), p2).AbsoluteUri);
        }

        public static bool IsAbsUrl(string url)
        {
            return Uri.IsWellFormedUriString(url, UriKind.Absolute);
        }

        //Back-up file with a new file name
        public static string BackupFileWithNewName(string filePath)
        {
            string dir = Path.GetDirectoryName(filePath);
            dir = Path.Combine(dir, "Bak");
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            string bakFilePath = Path.Combine(dir, NewFileName(filePath));
            File.Move(filePath, bakFilePath);
            return bakFilePath;
        }

        //Back-up file with the original file name
        public static string BackUpFile(string filePath)
        {
            string dir = Path.GetDirectoryName(filePath);
            dir = Path.Combine(dir, "Bak");
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            string bakFilePath = Path.Combine(dir, Path.GetFileName(filePath));
            if (File.Exists(bakFilePath))
            {
                File.Delete(bakFilePath);
            }
            File.Copy(filePath, bakFilePath);
            return bakFilePath;
        }

        //Back-up file with the original file name
        public static string BackUpFileWithDateFolder(string filePath, bool overWrite)
        {
            string dir = Path.GetDirectoryName(filePath);
            dir = Path.Combine(dir, "Bak");
            dir = Path.Combine(dir, DateTime.Today.ToString("yyyy-MM-dd"));
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            string bakFilePath = Path.Combine(dir, Path.GetFileName(filePath));
           
            
            if (File.Exists(bakFilePath))
            {
                if (overWrite)
                {
                    File.Delete(bakFilePath);
                }
                else
                {
                    return bakFilePath;
                }
            }

            File.Copy(filePath, bakFilePath);

            return bakFilePath;
        }

        //Generate a new file name based on the current date.
        public static string NewFileName(string fileName)
        {
            string newFileName = Path.GetFileNameWithoutExtension(fileName);
            newFileName += "_";
            newFileName += "Previous";
            newFileName += "_";
            newFileName += DateTime.Now.ToString("yyyy-MM-dd");
            newFileName += "_";
            newFileName += Guid.NewGuid();
            newFileName += Path.GetExtension(fileName);
            return newFileName;
        }


        //TO DO: 
        //public static string ParseDateTime(DateTime dateTime, string format)

        // Parse DateTime as the format 01Jan2011.
        public static string ParseDateTime(DateTime dateTime)
        {
            string dateTimeStr = string.Empty;
            string[] month = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            string temp = dateTime.ToString("dd_MM_yyyy");
            string[] tempArr = temp.Split('_');
            dateTimeStr = tempArr[0];
            dateTimeStr += month[int.Parse(tempArr[1])];
            dateTimeStr += tempArr[2];
            return dateTimeStr;
        }

        // Parse DateTime as the format 20 SEP-2011
        public static string ParseDateTimeWithBlank(DateTime dateTime)
        {
            string dateTimeStr = string.Empty;
            string[] month = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            string temp = dateTime.ToString("dd_MM_yyyy");
            string[] tempArr = temp.Split('_');
            dateTimeStr = tempArr[0];
            dateTimeStr += " ";
            dateTimeStr += month[int.Parse(tempArr[1])].ToUpper();
            dateTimeStr += "-";
            dateTimeStr += tempArr[2];
            return dateTimeStr;
        }

        //Get last business day
        public static DateTime GetLastBusinessDay(int holidayCount, DateTime time)
        {
            DateTime lastBusinessDay = DateTime.MinValue;
            if (time.DayOfWeek == DayOfWeek.Monday)
            {
                lastBusinessDay = time.AddDays(-3-holidayCount);
            }

            else if (time.DayOfWeek == DayOfWeek.Sunday)
            {
                lastBusinessDay = time.AddDays(-2 - holidayCount);
            }
            else 
            {
                lastBusinessDay = time.AddDays(-1 - holidayCount);
            }
            return lastBusinessDay;
        }

        public static DateTime GetLastTradingDay(DateTime originDate, List<DateTime> holidayList, int deltaDay)
        {
            var strHolidayList = from holiday in holidayList select holiday.ToString("yyyy-MM-dd");
            HashSet<string> holidaySet = new HashSet<string>(strHolidayList);

            DateTime curDay = originDate;
            int dayLeft = deltaDay;

            while (true)
            {
                if (curDay.DayOfWeek == DayOfWeek.Saturday || curDay.DayOfWeek == DayOfWeek.Sunday ||
                    holidaySet.Contains(curDay.ToString("yyyy-MM-dd")))
                {
                    curDay = curDay.AddDays(-1);
                }
                else
                {
                    curDay = curDay.AddDays(-1);
                    dayLeft--;
                }

                if (dayLeft == 0)
                {
                    while (true)
                    {
                        if (curDay.DayOfWeek == DayOfWeek.Saturday || curDay.DayOfWeek == DayOfWeek.Sunday ||
                            holidaySet.Contains(curDay.ToString("yyyy-MM-dd")))
                        {
                            curDay = curDay.AddDays(-1);
                        }
                        else
                        {
                            break;
                        }
                    }
                    break;
                }
            }

            return curDay;
        }

        public static DateTime GetNextTradingDay(DateTime originDate, List<DateTime> holidayList, int deltaDay)
        {
            var strHolidayList = from holiday in holidayList select holiday.ToString("yyyy-MM-dd");
            HashSet<string> holidaySet = new HashSet<string>(strHolidayList);

            DateTime curDay = originDate;
            int dayLeft = deltaDay;

            while (true)
            {
                if (curDay.DayOfWeek == DayOfWeek.Saturday || curDay.DayOfWeek == DayOfWeek.Sunday ||
                    holidaySet.Contains(curDay.ToString("yyyy-MM-dd")))
                {
                    curDay = curDay.AddDays(1);
                }
                else
                {
                    curDay = curDay.AddDays(1);
                    dayLeft--;
                }

                if (dayLeft == 0)
                {
                    while (true)
                    {
                        if (curDay.DayOfWeek == DayOfWeek.Saturday || curDay.DayOfWeek == DayOfWeek.Sunday ||
                            holidaySet.Contains(curDay.ToString("yyyy-MM-dd")))
                        {
                            curDay = curDay.AddDays(1);
                        }
                        else
                        {
                            break;
                        }
                    }
                    break;
                }
            }

            return curDay;
        }
    
        //Get the delta days th trading day
        public static DateTime GetNextWeeklyTradingDay(DateTime originDate, List<DateTime> holidayList, int deltaDay)
        {
            DateTime curDay = originDate;
            if (curDay.DayOfWeek == DayOfWeek.Monday)
            {
                curDay = originDate.AddDays(-1);
            }
            else if (curDay.DayOfWeek == DayOfWeek.Tuesday)
            {
                curDay = originDate.AddDays(-2);
            }
            else if (curDay.DayOfWeek == DayOfWeek.Wednesday)
            {
                curDay = originDate.AddDays(-3);
            }
            else if (curDay.DayOfWeek == DayOfWeek.Thursday)
            {
                curDay = originDate.AddDays(-4);
            }
            else if (curDay.DayOfWeek == DayOfWeek.Friday)
            {
                curDay = originDate.AddDays(-5);
            }
            curDay = GetNextTradingDay(curDay, holidayList, deltaDay);
            return curDay;
        }

        //Get next business day
        public static DateTime GetNextBusinessDay(int holidayCount, DateTime time)
        {
            DateTime nextBusinessDay = DateTime.MinValue;
            if (time.DayOfWeek == DayOfWeek.Friday)
            {
                nextBusinessDay = time.AddDays(3 + holidayCount);
            }
            if (time.DayOfWeek == DayOfWeek.Saturday)
            {
                nextBusinessDay = DateTime.Now.AddDays(2 + holidayCount);
            }
            else
            {
                nextBusinessDay = DateTime.Now.AddDays(1 + holidayCount);
            }
            return nextBusinessDay;
        }

        public static string GetCleanTextFromHtml(string html)
        {
            return HttpUtility.HtmlDecode(html).Replace("\t", "").Replace("\r", "").Replace("\n", "").Trim();
        }

        public static string NewFileNameWithDate(string origFilePath)
        {
            string newFilePath = Path.GetFileNameWithoutExtension(origFilePath);
            newFilePath += "_";
            newFilePath += DateTime.Now.ToString("mm_hh_dd_MMM_yyyy");
            newFilePath += ".";
            newFilePath += Path.GetExtension(origFilePath);
            return newFilePath;
        }

        // Transfer full code and half code
        public static string CharConverter(string source)
        {
            System.Text.StringBuilder result = new System.Text.StringBuilder(source.Length, source.Length);
            for (int i = 0; i < source.Length; i++)
            {
                if (source[i] >= 65281 && source[i] <= 65373)
                {
                    result.Append((char)(source[i] - 65248));
                }
                else if (source[i] == 12288)
                {
                    result.Append(' ');
                }
                else
                {
                    result.Append(source[i]);
                }
            }
            return result.ToString();
        }

        // Get excel file extension 
        public static string getFileExtension(string source)
        {
            string excelFileExtension = string.Empty;
            if (source.Contains("03"))
            {
                excelFileExtension = ".xls";
            }

            else if (source.Contains("07"))
            {
                excelFileExtension = ".xlsx";
            }
            else
            {
                excelFileExtension = ".xlsx";
            }
            return excelFileExtension;
        }
    }
}
