using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Ric.Core;

namespace Ric.Tasks.Taiwan
{
    public class TWHelper
    {
        private static Logger logger = null;
        // Get TW year (2012 maps to 101)
        public static string GetTWYear(string year)
        {
            if (year.Length != 4)
            {
                throw new Exception(string.Format("The dateTime year {0} should be in the format yyyy", year));
            }
            int temp = int.Parse(year);
            int twYear = temp - 1911;
            return twYear.ToString();
        }

        public static string GetYearFromTWYear(string year)
        {
            int temp = int.Parse(year) + 1911;
            return temp.ToString();
        }

        public static string DateStringForm(string dateStr, string stringForm)
        {
            DateTime dateValue;
            string dateStrFormed = null;
            try
            {
                string[] formats = {
                                       "d-MMM-yyyy",
                                    "M-d-yy", "MM-d-yy", "M-dd-yy", "MM-dd-yy", "MMM-dd-yy",
                                    "M-d-yyyy", "MM-d-yyyy", "M-dd-yyyy", "MM-dd-yyyy", "MMM-dd-yyyy",
                                    "M-yy-d", "MM-yy-d", "M-yy-dd", "MM-yy-dd", "MMM-yy-dd",
                                    "M-yyyy-d", "MM-yyyy-d", "M-yyyy-dd", "MM-yyyy-dd", "MMM-yyyy-dd",
                                    "d-M-yy", "d-MM-yy", "dd-M-yy", "dd-MM-yy","dd-MMM-yy",
                                    "d-M-yyyy", "d-MM-yyyy", "dd-M-yyyy", "dd-MM-yyyy", "dd-MMM-yyyy",
                                    "d-yy-M", "d-yy-MM", "dd-yy-M", "dd-yy-MM", "dd-yy-MMM",
                                    "d-yyyy-M", "d-yyyy-MM", "dd-yyyy-M", "dd-yyyy-MM", "dd-yyyy-MMM",
                                    "yy-M-d", "yy-MM-d", "yy-M-dd", "yy-MM-dd", "yy-MMM-dd",
                                    "yyyy-M-d", "yyyy-MM-d", "yyyy-M-dd", "yyyy-MM-dd", "yyyy-MMM-dd",
                                    "yy-d-M", "yy-d-MM", "yy-dd-M", "yy-dd-MM", "yy-dd-MMM",
                                    "yyyy-d-M", "yyyy-d-MM", "yyyy-dd-M", "yyyy-dd-MM", "yyyy-dd-MMM",
                                    "M/d/yy", "MM/d/yy", "M/dd/yy", "MM/dd/yy", "MMM/dd/yy",
                                    "M/d/yyyy", "MM/d/yyyy", "M/dd/yyyy", "MM/dd/yyyy", "MMM/dd/yyyy",
                                    "M/yy/d", "MM/yy/d", "M/yy/dd", "MM/yy/dd", "MMM/yy/dd",
                                    "M/yyyy/d", "MM/yyyy/d", "M/yyyy/dd", "MM/yyyy/dd", "MMM/yyyy/dd",
                                    "d/M/yy", "d/MM/yy", "dd/M/yy", "dd/MM/yy","dd/MMM/yy",
                                    "d/M/yyyy", "d/MM/yyyy", "dd/M/yyyy", "dd/MM/yyyy", "dd/MMM/yyyy",
                                    "d/yy/M", "d/yy/MM", "dd/yy/M", "dd/yy/MM", "dd/yy/MMM",
                                    "d/yyyy/M", "d/yyyy/MM", "dd/yyyy/M", "dd/yyyy/MM", "dd/yyyy/MMM",
                                    "yy/M/d", "yy/MM/d", "yy/M/dd", "yy/MM/dd", "yy/MMM/dd",
                                    "yyyy/M/d", "yyyy/MM/d", "yyyy/M/dd", "yyyy/MM/dd", "yyyy/MMM/dd",
                                    "yy/d/M", "yy/d/MM", "yy/dd/M", "yy/dd/MM", "yy/dd/MMM",
                                    "yyyy/d/M", "yyyy/d/MM", "yyyy/dd/M", "yyyy/dd/MM", "yyyy/dd/MMM",
                                    "M.d.yy", "MM.d.yy", "M.dd.yy", "MM.dd.yy", "MMM.dd.yy",
                                    "M.d.yyyy", "MM.d.yyyy", "M.dd.yyyy", "MM.dd.yyyy", "MMM.dd.yyyy",
                                    "M.yy.d", "MM.yy.d", "M.yy.dd", "MM.yy.dd", "MMM.yy.dd",
                                    "M.yyyy.d", "MM.yyyy.d", "M.yyyy.dd", "MM.yyyy.dd", "MMM.yyyy.dd",
                                    "d.M.yy", "d.MM.yy", "dd.M.yy", "dd.MM.yy","dd.MMM.yy",
                                    "d.M.yyyy", "d.MM.yyyy", "dd.M.yyyy", "dd.MM.yyyy", "dd.MMM.yyyy",
                                    "d.yy.M", "d.yy.MM", "dd.yy.M", "dd.yy.MM", "dd.yy.MMM",
                                    "d.yyyy.M", "d.yyyy.MM", "dd.yyyy.M", "dd.yyyy.MM", "dd.yyyy.MMM",
                                    "yy.M.d", "yy.MM.d", "yy.M.dd", "yy.MM.dd", "yy.MMM.dd",
                                    "yyyy.M.d", "yyyy.MM.d", "yyyy.M.dd", "yyyy.MM.dd", "yyyy.MMM.dd",
                                    "yy.d.M", "yy.d.MM", "yy.dd.M", "yy.dd.MM", "yy.dd.MMM",
                                    "yyyy.d.M", "yyyy.d.MM", "yyyy.dd.M", "yyyy.dd.MM", "yyyy.dd.MMM",
                                    };
                dateValue = DateTime.ParseExact(dateStr, formats, new CultureInfo("en-US"), DateTimeStyles.None);
                dateStrFormed = dateValue.ToString(stringForm, new CultureInfo("en-US"));
                return dateStrFormed;
            }
            catch (Exception ex)
            {
                logger.Log(string.Format("Error happens when Format DateTime String {0} .", ex.Message));
                return dateStrFormed;
            }
        }
        
        public static string ClearHtmlTags(string strTags)
        {
            strTags = Regex.Replace(strTags, @"<script[^>]*?>.*?</script>", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"-->", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"<!--.*", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(quot|#34);", "\"", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(nbsp|#160);", " ", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"&#(\d+);", "", RegexOptions.IgnoreCase);
            strTags = Regex.Replace(strTags, @"<img[^>]*>;", "", RegexOptions.IgnoreCase);
            strTags.Replace("<", "");
            strTags.Replace(">", "");
            strTags.Replace("\r\n", "");
            return strTags;
        }
    }
}