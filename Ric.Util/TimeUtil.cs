using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace Ric.Util
{
    public static class TimeUtil
    {
        public static string shortYear = DateTime.Now.Year.ToString().Substring(2, 2);
        public static string GetEffectiveDate(DateTime date)
        {
            return date.ToString("dd") + "-" + date.ToString("MMM", new CultureInfo("en-US")) + "-" + date.ToString("yy");
        }

        public static string GetFormatDate(DateTime date)
        {
            return date.ToString("dd") + date.ToString("MMM", new CultureInfo("en-US")).ToUpper() + date.ToString("yyyy");
        }
        public static string GetYYYYMMDD(DateTime date)
        {
            return date.ToString("yyyyMMdd", new CultureInfo("en-US"));
        }
        // return dd/mm/yyyy
        public static string GetDDMMYYWithSplit(DateTime date)
        {
            return date.ToString("dd") + "/" + date.ToString("MM") + "/" + date.ToString("yyyy");
        }

        public static DateTime ParseTime(string value)
        {
            string[] dateParts = value.Split('/');
            DateTime date = DateTime.Parse(dateParts[1] + "/" + dateParts[0] + "/" + dateParts[2], new CultureInfo("en-US"));
            return date;
        }

        /// <summary>
        /// 转化UTC时间为PST时间
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static DateTime ConvertToPST(this DateTime dateTime)
        {
            return System.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime, "Pacific Standard Time");
        }

        /// <summary>
        /// 转化UTC时间为PST时间
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static DateTime ConvertToChina(this DateTime dateTime)
        {
            return System.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime, "China Standard Time");
        }
    }
}
