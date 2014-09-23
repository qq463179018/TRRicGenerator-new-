using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace PdfTronWrapper.Utility
{
    /// <summary>
    /// This file contains all the basic string operations
    /// </summary>
    internal static class StringExtension
    {
        /// <summary>
        /// Remove the blank space of the string
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string RemoveBlankSpace(this string str)
        {
            return str.Replace(" ", "").Replace(" ", "").Replace("　", "");
        }


        /// <summary>
        /// Get the first letter of the string
        /// </summary>
        /// <param name="str">A string</param>
        /// <returns>The first letter of the string</returns>
        public static string GetFirstLetter(this string str, bool isChinese)
        {
            for (int i = 0; i < str.Length; i++)
            {
                Regex rx = new Regex(isChinese ? "^[\u4e00-\u9fa5|0-9]$" : "^[a-zA-Z]$");
                string letter = str[i].ToString();
                if (rx.IsMatch(letter))
                    return letter;
            }
            return str;
        }

    }
}
