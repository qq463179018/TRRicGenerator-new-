using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using HtmlAgilityPack;
using Ric.Core;
using Ric.FileLib;
using Ric.FormatLib;
using Ric.Util;
using FileMode = Ric.FileLib.Enum.FileMode;

namespace Ric.Tasks.Thailand
{
    #region Configuration

    [ConfigStoredInDB]
    public class ThailandDwAddConfig
    {
        [StoreInDB]
        [DisplayName("Result folder")]
        [Description("The path where the result will be written.\nEg: C:/Mydrive/")]
        public string ResultFolderPath { get; set; }
    }

    #endregion

    #region Task

    class ThailandDwAdd : GeneratorBase
    {
        #region Initialization

        private static ThailandDwAddConfig configObj;
        private const string BaseUrl = "http://www.set.or.th";
        private CookieContainer cookies;
        private CookieContainer cookiesIsin;
        private const string Alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        #endregion

        #region GeneratorBase Implementation

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as ThailandDwAddConfig;
        }

        protected override void Start()
        {
            try
            {
                //Login1();
                //Login2();
                int i = 1;
                var propsDw = new List<Dictionary<string, string>>();
                foreach (Dictionary<string, string> infos in GetPages().Select(page => CleanInfos(GetInfos(page))).Where(infos => infos["type"] == "Derivative warrant"))
                {
                    infos.Add("counter", i.ToString());
                    propsDw.Add(infos);
                    i++;
                }
                if (i > 1)
                {
                    //GetFSharpTitleValue(propsDw);


                    var ndaDw1 = new Fm(Format.Horizontal, FileMode.WriteOnly);
                    ndaDw1.LoadFromTemplate(TemplateFm.ThFm, propsDw);
                    ndaDw1.Save(String.Format("{0}Thailand_DW_ADD_{1}.xls", configObj.ResultFolderPath,
                        DateTime.Now.ToString("ddMM")));
                    AddResult("Dw Fm bulk file", ndaDw1.Path, "fm");

                    var ndaDw2 = new NdaQa(FileMode.WriteOnly);
                    ndaDw2.LoadFromTemplate(Template.QaAdd, propsDw);
                    ndaDw2.Save(String.Format("{0}NDA_QA_ADD_{1}.csv", configObj.ResultFolderPath,
                        DateTime.Now.ToString("ddMM")));
                    AddResult("Dw Nda bulk file", ndaDw2.Path, "nda");

                    var ndaDw3 = new Idn(FileMode.WriteOnly);
                    ndaDw3.LoadFromTemplate(TemplateIdn.DomChain, propsDw);
                    ndaDw3.Save(String.Format("{0}IDN_Domchain_{1}.txt", configObj.ResultFolderPath,
                        DateTime.Now.ToString("ddMM")));
                    AddResult("Dw Idn bulk file", ndaDw3.Path, "idn");

                    var ndaDw4 = new Idn(FileMode.WriteOnly);
                    ndaDw4.LoadFromTemplate(TemplateIdn.ForIdn, propsDw);
                    ndaDw4.Save(String.Format("{0}IDN_ADD_{1}.txt", configObj.ResultFolderPath,
                        DateTime.Now.ToString("ddMM")));
                    AddResult("Dw Idn bulk file", ndaDw4.Path, "idn");

                    var ndaDw5 = new NdaIa(FileMode.WriteOnly);
                    ndaDw5.LoadFromTemplate(Template.IaAdd, propsDw);
                    ndaDw5.Save(String.Format("{0}NDA_IA_ADD_{1}.csv", configObj.ResultFolderPath,
                        DateTime.Now.ToString("ddMM")));
                    AddResult("Dw Nda bulk file", ndaDw5.Path, "nda");

                    var ndaDw6 = new Tc(FileMode.WriteOnly);
                    ndaDw6.LoadFromTemplate(TemplateTc.WrtAdd, propsDw);
                    ndaDw6.Save(String.Format("{0}WRT_ADD_{1}_thailand.csv", configObj.ResultFolderPath,
                        DateTime.Now.ToString("ddMMMyyyy")));
                    AddResult("Dw t&c bulk file", ndaDw6.Path, "t&c");
                }
                else
                {
                    LogMessage("No Dw announcement today");
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        //private void GetFSharpTitleValue(List<Dictionary<string, string>> propsDw)
        //{
        //    string result = string.Empty;
        //    foreach (var item in propsDw[0])
        //    {
        //        result = result + item.Key + " : " + item.Value + "\r\n";
        //    }
        //    Console.WriteLine(result);
        //}

        #endregion

        /*
        private void Login1()
        {
            ServicePointManager.Expect100Continue = false;
            string pageSource = string.Empty;
            Encoding encoding = Encoding.UTF8;
            string postData = String.Format("txtLogin=11794278&txtPassword=Thailand3&x=19&y=2");
            string uri = String.Format("http://www.setsmart.com/ism/ism_verifyuser.jsp");
            byte[] buf = encoding.GetBytes(postData);

            var request = WebRequest.Create(uri) as HttpWebRequest;
            request.Timeout = 300000;
            request.AllowAutoRedirect = true;
            request.CookieContainer = cookiesIsin;
            request.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:6.0.2) Gecko/20100101 Firefox/6.0.2";
            request.Method = "POST";
            request.Referer = "http://www.setsmart.com/ism/login.jsp";
            request.ContentType = "text/html;charset=tis-620";
            request.KeepAlive = true;
            if (!string.IsNullOrEmpty(postData))
            {
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);
            }
            using (WebResponse response = request.GetResponse())
            {
                var sr = new StreamReader(response.GetResponseStream());
                pageSource = sr.ReadToEnd();
            }
        }

        private void Login2()
        {
            ServicePointManager.Expect100Continue = false;
            string pageSource = string.Empty;
            Encoding encoding = Encoding.UTF8;
            byte[] buf = null;
            string postData = String.Format("txtLogin=11794278&txtPassword=Thailand3&hidBrowser=null&hidLang=English");
            string uri = String.Format("http://www.setsmart.com/ism/LoginRep.jsp");
            buf = encoding.GetBytes(postData);

            var request = WebRequest.Create(uri) as HttpWebRequest;
            request.Timeout = 300000;
            request.AllowAutoRedirect = true;
            request.CookieContainer = cookiesIsin;
            request.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:6.0.2) Gecko/20100101 Firefox/6.0.2";
            request.Method = "POST";
            request.Referer = "http://www.setsmart.com/ism/ism_verifyuser.jsp";
            request.ContentType = "text/html;charset=tis-620";
            request.KeepAlive = true;
            if (!string.IsNullOrEmpty(postData))
            {
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);
            }
            using (WebResponse response = request.GetResponse())
            {
                var sr = new StreamReader(response.GetResponseStream());
                pageSource = sr.ReadToEnd();
            }
        } 
        */

        private Dictionary<string, string> CleanInfos(Dictionary<string, string> oldInfos)
        {
            if (oldInfos["Listing"] == "Derivative warrant")
            {
                return GetDwInfos(oldInfos);
            }
            var newInfos = new Dictionary<string, string>
            {
                {"type", "other"}
            };
            return newInfos;
        }

        private Dictionary<string, string> GetDwInfos(Dictionary<string, string> oldInfos)
        {
            var newInfos = new Dictionary<string, string>
            { 
                {"type", oldInfos["Listing"]},
                {"code", oldInfos["DW name"]},
                {"ric", oldInfos["DW name"].Remove(oldInfos["DW name"].Length - 5, 1)},
                {"market", GetMarket(oldInfos["DW name"])},
                {"tradingdate", oldInfos["Trading date"]},
                {"tradingdatewrt", DateTime.ParseExact(oldInfos["Trading date"], "dd-MMM-yyyy", null).ToString("dd/MM/yyyy")},
                {"lasttradingdate", oldInfos["Last trading date"]},
                {"lasttradingdatewrt", DateTime.ParseExact(oldInfos["Last trading date"], "dd-MMM-yyyy", null).ToString("dd/MM/yyyy")},
                {"maturitydate", oldInfos["Maturity date"]},
                {"maturitydatewrt", DateTime.ParseExact(oldInfos["Maturity date"], "dd-MMM-yyyy", null).ToString("dd/MM/yyyy")},
                {"codestart", GetStart(oldInfos["DW name"])},
                {"cp", GetCp(oldInfos["DW name"])},
                {"asset", oldInfos["Underlying Asset"]},
                {"warrantissuer", "45828"},
                {"assetclass", "TRAD"},
                {"lastexercisedate", oldInfos["Last exercise date"]},
                {"lastexercisedatewrt", DateTime.ParseExact(oldInfos["Last exercise date"], "dd-MMM-yyyy", null).ToString("dd/MM/yyyy")},
                {"name", oldInfos["DW issuer"].Substring(0, oldInfos["DW issuer"].LastIndexOf("(")).TrimEnd()},
                {
                    "expiredate",
                    DateTime.ParseExact(oldInfos["Maturity date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .AddDays(1)
                        .ToString("dd-MMM-yyyy")
                },
                {
                    "expiredatewrt",
                    DateTime.ParseExact(oldInfos["Maturity date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .AddDays(1)
                        .ToString("dd/MM/yyyy")
                },
                {
                    "number",
                    oldInfos["Number of derivative warrants ("].Remove(
                        oldInfos["Number of derivative warrants ("].IndexOf(" "))
                },
                {
                    "abbr",
                    oldInfos["DW issuer"].Substring(oldInfos["DW issuer"].LastIndexOf("("))
                        .Replace("(", "")
                        .Replace(")", "")
                }
            };

            newInfos.Add("extension", GetExtension(newInfos["market"]));
            if (newInfos["extension"].Equals(".SET50"))
                newInfos["asset"] = string.Empty;

            newInfos.Add("rawnumber", newInfos["number"].Replace(",", ""));
            newInfos.Add("bigexpiredate", newInfos["expiredate"].ToUpper());
            newInfos.Add("display", GetDisplayName(newInfos["codestart"], newInfos["cp"], newInfos["abbr"], newInfos["market"]));

            if (oldInfos.ContainsKey("Exercise price"))
            {
                newInfos.Add("price", oldInfos["Exercise price"]);
            }
            else if (oldInfos.ContainsKey("Strike level (Index Point)"))
            {
                newInfos.Add("price", oldInfos["Strike level (Index Point)"]);
            }
            else
            {
                newInfos.Add("price", "");
            }

            newInfos.Add("ratio", oldInfos.ContainsKey("Exercise ratio") ? oldInfos["Exercise ratio"] : "1");
            newInfos.Add("multiplier", oldInfos.ContainsKey("Multiplier (THB : Index Point)") ? oldInfos["Multiplier (THB : Index Point)"] : "");
            newInfos.Add("bigdate", GetBigDate(newInfos["maturitydate"], newInfos["price"]));
            return newInfos;
        }

        private string GetMarket(string dwName)
        {
            if (dwName.StartsWith("S50"))
            {
                return "S50";
            }
            return dwName.StartsWith("SET") ? "SET" : "BK";
        }

        private string GetExtension(string market)
        {
            if (market == "S50")
            {
                return ".SET50";
            }
            return market == "SET" ? ".SETI" : ".BK";
        }

        private string GetIsin(string code)
        {
            ServicePointManager.Expect100Continue = false;
            string pageSource = string.Empty;
            Encoding encoding = Encoding.UTF8;
            byte[] buf = null;
            string postData = String.Format("symbol=ADVA13C1406B&submit.x=13&submit.y=3");
            string uri = String.Format("http://www.setsmart.com/ism/companyprofile.html");
            buf = encoding.GetBytes(postData);

            var request = WebRequest.Create(uri) as HttpWebRequest;
            request.Timeout = 300000;
            request.AllowAutoRedirect = true;
            request.CookieContainer = cookiesIsin;
            request.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:6.0.2) Gecko/20100101 Firefox/6.0.2";
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            if (!string.IsNullOrEmpty(postData))
            {
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);
            }
            using (WebResponse response = request.GetResponse())
            {
                var sr = new StreamReader(response.GetResponseStream());
                pageSource = sr.ReadToEnd();
            }
            var doc = new HtmlDocument();
            doc.LoadHtml(pageSource);

            return "";
        }

        private string GetDisplayName(string start, string cp, string abbr, string market)
        {
            var displayNameBuilder = new StringBuilder();
            var startText = market == "S50" ? market : start;

            displayNameBuilder.Append(startText);
            displayNameBuilder.Append(" BY ");
            displayNameBuilder.Append(abbr);
            if (startText.Length + abbr.Length == 9)
            {
                displayNameBuilder.Append(cp == "C" ? "C" : "P");
            }
            else if (start.Length + abbr.Length == 8)
            {
                displayNameBuilder.Append(cp == "C" ? " C" : " P");
            }
            else
            {
                displayNameBuilder.Append(cp == "C" ? " CA" : " PU");
            }
            displayNameBuilder.Append("#A");
            return displayNameBuilder.ToString();
        }

        private string GetBigDate(string date, string price)
        {
            if (date.Contains(" "))
            {
                date = date.TrimStart().Remove(date.IndexOf(' '));
            }
            DateTime bigDate = DateTime.ParseExact(date, "dd-MMM-yyyy", CultureInfo.InvariantCulture);
            if (price != "")
            {
                price = Convert.ToDouble(price).ToString();
            }
            return price + bigDate.ToString("MMMyy").ToUpper();
        }

        private string GetCp(string name)
        {
            for (int i = (name.Length - 2); i >= 0; i--)
            {
                if (Alpha.Contains(name[i].ToString()))
                {
                    return (name[i].ToString());
                }
            }
            return null;
        }

        private string GetStart(string dwname)
        {
            string res = "";
            foreach (char t in dwname)
            {
                if (!Alpha.Contains(t.ToString()))
                {
                    return res;
                }
                res += t;
            }
            return res;
        }

        private Dictionary<string, string> GetInfos(string page)
        {
            var infos = new Dictionary<string, string>();
            var request = WebRequest.Create(page) as HttpWebRequest;
            request.ProtocolVersion = HttpVersion.Version11;
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0";
            request.Method = "GET";
            request.KeepAlive = true;
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.CookieContainer = cookies;
            request.Credentials = CredentialCache.DefaultCredentials;
            request.Headers["Cache-Control"] = "max-age=0";
            request.Headers["Accept-Encoding"] = "gzip,deflate,sdch";
            request.Headers["Accept-Language"] = "en-US,en;q=0.8";
            var httpResponse = (HttpWebResponse)request.GetResponse();

            Stream pageStream = httpResponse.GetResponseStream();
            var html = new HtmlDocument();
            html.Load(pageStream);

            HtmlNode pre = html.DocumentNode.SelectSingleNode(".//pre");
            string[] lineSeparators = { "\n" };
            string[] pointSeparators = { ":" };

            string[] lines = pre.InnerText.Split(lineSeparators, StringSplitOptions.None);
            string tmpKey = "";
            foreach (string[] lineParts in lines.Select(line => line.Replace(" (DW : Underlying asset)", "")
                                                                    .Replace(" (DW : Underlying Asset)", "")
                                                                    .Replace("unit:", "")
                                                                    .Replace("Warrant : ", "")
                                                                    .Replace(" (Baht)", "")
                                                                    .Replace(" (baht)", "")
                                                                    .Split(pointSeparators, StringSplitOptions.None)))
            {
                if (lineParts.Length >= 2)
                {
                    infos.Add(lineParts[0].Trim(), lineParts[1].Trim());
                    tmpKey = lineParts[0].Trim();
                }
                if (lineParts.Length == 1 && tmpKey != "")
                {
                    infos[tmpKey] += (" " + lineParts[0]);
                }
            }
            return infos;
        }

        private IEnumerable<string> GetPages()
        {
            ServicePointManager.Expect100Continue = false;
            var results = new List<string>();


            //foreach (string url in Enumerable.Range(0, 10).Select(i => string.Format("http://www.set.or.th/set/newslist.do?to={0}%2F{1}%2F{2}&headline=adds+new+listed&submit=Search&symbol=&currentpage={3}&from={0}%2F{1}%2F{2}&newsType=&country=US&exchangeSymbols=&company=false&exchangeNews=on&language=en&exchange=true"
            //      , DateTime.Now.AddDays(-13).ToString("dd"), DateTime.Now.AddDays(-13).ToString("MM"), DateTime.Now.Year, i)))
            foreach (string url in Enumerable.Range(0, 10).Select(i => string.Format("http://www.set.or.th/set/newslist.do?to={0}%2F{1}%2F{2}&headline=adds+new+listed&submit=Search&symbol=&currentpage={3}&from={0}%2F{1}%2F{2}&newsType=&country=US&exchangeSymbols=&company=false&exchangeNews=on&language=en&exchange=true"
                , DateTime.Now.AddDays(0).ToString("dd"), DateTime.Now.AddDays(0).ToString("MM"), DateTime.Now.Year, i)))
            {
                try
                {
                    var htc = WebClientUtil.GetHtmlDocument(url, 300000);

                    HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                    HtmlNodeCollection links = tables[15].SelectNodes(".//a");
                    results.AddRange(from link in links
                                     select link.Attributes["href"].Value.Trim() into rawlink
                                     select rawlink.Substring(rawlink.IndexOf("?")) into rawlink
                                     select rawlink.Replace("&amp;", "&") into rawlink
                                     select "http://www.set.or.th/set/newsdetails.do" + rawlink);
                }
                catch
                {
                    break;
                }
            }
            return results;
        }
    }

    #endregion
}
