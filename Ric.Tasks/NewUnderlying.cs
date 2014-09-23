using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks
{
    public class NewUnderlying : GeneratorBase
    {

        /// <summary>
        /// Encode text to bits
        /// </summary>
        /// <param name="encodeText">text to encode</param>
        /// <returns>encoded result</returns>
        public static string Encode(string encodeText)
        {
            return Regex.Replace(encodeText, "[^a-zA-Z0-9]",
                match =>
                    "%" + BitConverter.ToString(Encoding.GetEncoding("euc-kr").GetBytes(match.Value)).Replace("-", "%"));
        }

        /// <summary>
        /// Search the ISIN website for given korean name and key word: 보통주 first time: search with koreanName + 보통주. second time: search with koreanName
        /// If mutiple records or no record found. User need to input ISIN. 
        /// </summary>
        /// <param name="koreanName">korean name</param>
        /// <param name="times">searched times</param>
        /// <returns>isin</returns>
        public static string GetIsinByName(string koreanName,int times)
        {
            string postData = "";
            string encodeName = Encode(koreanName);
            if (koreanName.Contains("-"))
            {
                encodeName = Encode(koreanName.Split('-')[1]);
            }

            if (times == 0)
            {
                postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun=01&lst_yn1=Y&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word={0}&ef_iss_inst_cd=&ef_isu_nm={0}%BA%B8%C5%EB%C1%D6&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=", encodeName);
            }
            else if (times == 1)
            {
                postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun=01&lst_yn1=Y&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word={0}&ef_iss_inst_cd=&ef_isu_nm=%BA%B8%C5%EB%C1%D6&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=", encodeName);
            }           
            string uri = "http://isin.krx.co.kr/jsp/realBoard01.jsp";
            HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
            request.Timeout = 100000;
            request.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:6.0.2) Gecko/20100101 Firefox/6.0.2";
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Referer = "http://isin.krx.co.kr/jsp/BA_LT113.jsp";
            request.Headers.Add("Accept-Encoding: gzip,deflate,sdch");

            string pageSource = null;
            int retries = 3;
            while (pageSource == null && retries-- > 0)
            {
                try
                {
                    pageSource = WebClientUtil.GetDynamicPageSource(request, postData, Encoding.GetEncoding("EUC-KR"));
                }
                catch
                {
                    System.Threading.Thread.Sleep(5000);
                }
            }
            string isin = "";
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(pageSource);

            HtmlNodeCollection trs = null;
            try
            {
                trs = doc.DocumentNode.SelectNodes("/html[1]/body[1]/table[1]/tr[2]/td[1]/table[1]/tr");
            }
            catch
            {
                string msg = "Error found in searching new underlying record found for " + koreanName;
                //Logger.Log(msg, Logger.LogType.Error);
                isin = InputISIN.Prompt(koreanName, "Underlying Name");
            }
            //not found
            if (trs == null || trs.Count > 2)
            {
                isin = InputISIN.Prompt(koreanName, "Underlying Name");
            }
            else if (trs.Count == 1)
            {
                isin = times == 0 ? GetIsinByName(koreanName, 1) : InputISIN.Prompt(koreanName, "Underlying Name");
            }
            //find only one record
            else if (trs.Count == 2)
            {
                HtmlNode tr = trs[1];
                HtmlNodeCollection tds = tr.SelectNodes("./td");
                isin = tds[1].InnerText.Trim();
            }
            return isin;
        }

        /// <summary>
        /// Grab new underlying info with korean name and isin.
        /// </summary>
        /// <param name="koreaName">koreaName</param>
        /// <returns>new underlying info</returns>
        public static KoreaUnderlyingInfo GrabNewUnderlyingInfo(string koreaName)
        {
            return GrabNewUnderlyingInfo(koreaName, GetIsinByName(koreaName, 0));
        }

        /// <summary>
        /// Grab new underlying info with korean name.
        /// </summary>
        /// <param name="koreaName">korean name</param>
        /// <returns>new underlying info</returns>
        public static KoreaUnderlyingInfo GrabNewUnderlyingInfo(string koreaName, string isin)
        {           
            if (string.IsNullOrEmpty(isin))
            {                      
                return null;
            }
            KoreaUnderlyingInfo newUnderlying = new KoreaUnderlyingInfo();
            string uri = string.Format("http://isin.krx.co.kr/jsp/BA_VW010.jsp?isu_cd={0}&modi=f&req_no=", isin);

            string pageSource = null;
            int retry = 3;
            while (pageSource == null && retry-- > 0)
            {
                try
                {
                    pageSource = WebClientUtil.GetPageSource(null, uri, 6000, "", Encoding.GetEncoding("EUC-KR"));
                }
                catch
                {
                    continue;
                }
            }
            if (pageSource == null)
            {
                string msg = "Can not get the New Underlying infos in ISIN webpage. For ISIN:" + isin + ". please check if the webpage can be accessed!";
                return null;
            }
            HtmlDocument isinRoot = new HtmlDocument();
            isinRoot.LoadHtml(pageSource);
            HtmlNode isinTable = isinRoot.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[1]/tr[3]/td[1]/table[1]");
            HtmlNodeCollection isinTrs = isinTable.SelectNodes("./tr");

            string ric = isinTrs[2].SelectNodes("./td")[3].InnerText.TrimStart().TrimEnd();
            string sixDigit = ric.Substring(ric.Length - 6);
            string underEngName = isinTrs[10].SelectNodes("./td")[1].InnerText.TrimStart().TrimEnd();
            string suffix = string.IsNullOrEmpty(isinTrs[11].SelectNodes("./td")[2].InnerText.TrimStart().TrimEnd()) ? "KQ" : "KS";
            string usName = isinTrs[10].SelectNodes("./td")[3].InnerText.Trim();
            string symbol = isinTrs[3].SelectNodes("./td")[1].InnerText.Trim();
            string companyName = isinTrs[0].SelectNodes("./td")[1].InnerText.Trim();
            Regex regex = new Regex(@"\[.+?]");
            Match m = regex.Match(companyName);
            companyName = m.Value.Trim(new[] { ' ', '[', ']' });
            companyName = Regex.Replace(companyName, "&nbsp;", "");

            newUnderlying.UnderlyingRIC = sixDigit + "." + suffix;
            if (sixDigit.Substring(5) == "0")
            {
                sixDigit = sixDigit.Substring(0, 5);
            }
            sixDigit = "kr;" + sixDigit.TrimStart('0');
            if (suffix == "KQ")
            {
                sixDigit += "K";
            }
            string ndaTc = ClearCoLtdForName(underEngName.ToUpper());
            newUnderlying.QACommonNamePart = ndaTc;
            newUnderlying.NDATCUnderlyingTitle = ndaTc;
            newUnderlying.BNDUnderlying = sixDigit;
            newUnderlying.KoreaName = koreaName;
            newUnderlying.KoreaNameFM2 = koreaName;
            newUnderlying.KoreaNameDrop = koreaName;           
            newUnderlying.IDNDisplayNamePart = GetIDNDisplayName(symbol, usName, ndaTc);
            newUnderlying.ISIN = isin;
            newUnderlying.CompanyName = companyName;

            return newUnderlying;
        }

        /// <summary>
        /// Get IDN display name by symbol or usName or company name. 
        /// If the database contains a same display name, then change it.
        /// </summary>
        /// <param name="symbol"></param>
        /// <param name="usName"></param>
        /// <param name="ndaTc"></param>
        /// <returns></returns>
        public static string GetIDNDisplayName(string symbol, string usName, string ndaTc)
        {
            //symbol usName<=7digits 取length长的
            //usName>7 symbol!= "" 取symbol  
            // symbol=null   1. 1 word 前四后三  2. >2 words  1四2三
            //判断是否重复 ，若重复，用ndatc生成
            usName = usName.ToUpper();
            string result = string.Empty;
            string nameNoBlank = Regex.Replace(usName, "([ ]+)", "");
            if (nameNoBlank.Length <= 7)
            {
                result = nameNoBlank.Length >= symbol.Length ? nameNoBlank : symbol;
            }
            else
            {
                if (!string.IsNullOrEmpty(symbol))
                {
                    result = symbol;
                }
                else
                {
                    nameNoBlank = Regex.Replace(nameNoBlank, "([^0-9A-Z]+)", "");
                    Regex regex = new Regex("([0-9]+)");
                    MatchCollection match = regex.Matches(nameNoBlank);
                    if (match.Count > 0)
                    {
                        string subNum = "";
                        Dictionary<int, string> numDe = new Dictionary<int, string>();
                        string nameNoDigit = Regex.Replace(nameNoBlank, "([^A-Z]+)", "");
                        int totalLength = 0;
                        for (int i = match.Count - 1; i >= 0; i--)
                        {
                            totalLength += match[i].Length;
                            if ((match[i].Index + match[i].Length) > 7)
                            {
                                subNum = match[i].Value.Trim() + subNum;
                            }
                            else
                            {
                                numDe.Add(match[i].Index, match[i].Value);
                            }
                        }
                        nameNoDigit = nameNoDigit.Substring(0, 7 - totalLength);
                        nameNoDigit = numDe.Aggregate(nameNoDigit, (current, item) => current.Insert(item.Key, item.Value));
                        result = nameNoDigit + subNum;
                    }
                    else
                    {
                        result = nameNoBlank.Substring(0, 7);
                    }
                }
            }
            int retry = 1;
            while (KoreaUnderlyingManager.ExsitDisplayName(result))
            {
                result = GetIDNDisplayNamePortion(ndaTc, retry++);
            }
            return result;
        }



        /// <summary>
        /// Use some rules to generate IDN display name for new underlying. And check the DB if it is unique.
        /// </summary>
        /// <param name="usName"></param>
        /// <returns></returns>
        public static string GetUniqueIdnName(string usName, string companyName)
        {
            usName = usName.ToUpper();
            string result = "";
            string nameNoBlank = Regex.Replace(usName, "([ ]+)", "");
            if (nameNoBlank.Length <= 7)
            {
                result = nameNoBlank;
            }
            else
            {
                nameNoBlank = Regex.Replace(nameNoBlank, "([^0-9A-Z]+)", "");
                Regex regex = new Regex("([0-9]+)");
                MatchCollection match = regex.Matches(nameNoBlank);
                if (match.Count > 0)
                {
                    string subNum = "";
                    Dictionary<int, string> numDe = new Dictionary<int, string>();
                    string nameNoDigit = Regex.Replace(nameNoBlank, "([^A-Z]+)", "");
                    int totalLength = 0;
                    for (int i = match.Count - 1; i >= 0; i--)
                    {
                        totalLength += match[i].Length;
                        if ((match[i].Index + match[i].Length) > 7)
                        {
                            subNum = match[i].Value.Trim() + subNum;
                        }
                        else
                        {
                            numDe.Add(match[i].Index, match[i].Value);
                        }
                    }
                    nameNoDigit = nameNoDigit.Substring(0, 7 - totalLength);
                    nameNoDigit = numDe.Aggregate(nameNoDigit, (current, item) => current.Insert(item.Key, item.Value));
                    result = nameNoDigit + subNum;
                }
                else
                {
                    usName = ClearCoLtdForName(usName);
                    result = GetIDNDisplayNamePortion(usName, 0);
                }
            }
            int retry = 1;
            while (KoreaUnderlyingManager.ExsitDisplayName(result))
            {
                result = GetIDNDisplayNamePortion(companyName, retry++);
            }
            return result;
        }

        /// <summary>
        /// Get IDNDisplay name(new underlying) (less than 7 characters) with underlying English campany name.
        /// </summary>
        /// <param name="companyName">underlying English campany name</param>
        /// <returns>IDN Display name</returns>
        public static string GetIDNDisplayNamePortion(string companyName, int retry)
        {
            companyName = companyName.ToUpper();
            string cond = @"[A-Z0-9]+";
            Regex r = new Regex(cond);
            MatchCollection m = r.Matches(companyName);
            int n = m.Count;
            if (n == 1)
            {
                return (m[0].Value.Substring(0, 7 - retry) + m[0].Value.Substring(m[0].Length - retry, retry));
            }
            string idnName = m[0].Value.Length > 4 ? m[0].Value.Substring(0, 4) : m[0].Value;
            int subLength = (7 / n) == 0 ? 1 : (7 / n);
            for (int i = 1; i < n; i++)
            {
                Match item = m[i];
                if (item.Length >= (subLength + retry - 1))
                {
                    idnName += item.ToString().Substring(retry - 1, subLength);
                }
                else
                {
                    idnName += item.Value;
                }
            }
            idnName = idnName.Length > 7 ? idnName.Substring(0, 7) : idnName;
            return idnName;
        }

        /// <summary>
        /// Remove the infos of company like CO LTD CORP INC CORPARATION
        /// </summary>
        /// <param name="underEngName">full name</param>
        /// <returns>name without company infos</returns>
        public static string ClearCoLtdForName(string underEngName)
        {
            underEngName = underEngName.ToUpper();
            underEngName = underEngName.Replace("CORPORATION", "");
            List<string> names = underEngName.Split(new[] { ' ', ',', '.' }).ToList();
            names.Remove("CO");
            names.Remove("LTD");
            names.Remove("INC");
            names.Remove("CORP");
            string result = names
                            .Where(name => name != "" && name != " ")
                            .Aggregate("", (current, name) => current + (name + " "));
            return result.TrimEnd();
        }
    }
}
