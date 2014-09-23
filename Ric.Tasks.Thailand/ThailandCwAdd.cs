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
    public class ThailandCwAddConfig
    {
        [StoreInDB]
        [DisplayName("Result folder")]
        [Description("The path where the result will be written.\nEg: C:/Mydrive/")]
        public string ResultFolderPath { get; set; }
    }

    #endregion

    #region Task

    class ThailandCwAdd : GeneratorBase
    {
        #region Initialization

        private static ThailandCwAddConfig configObj;
        private const string BaseUrl = "http://www.set.or.th";
        private CookieContainer cookies;
        private CookieContainer cookiesIsin;
        private const string Alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        #endregion

        #region GeneratorBase Implementation

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as ThailandCwAddConfig;
        }

        protected override void Start()
        {
            try
            {
                int j = 0;
                var propsCw = new List<Dictionary<string, string>>();
                foreach (Dictionary<string, string> infos in GetPages()
                                .Select(page => CleanInfos(GetInfos(page)))
                                .Where(infos => infos["type"] == "Warrant"))
                {
                    infos.Add("counter", j.ToString());
                    propsCw.Add(infos);
                    j++;
                }
                if (j > 0)
                {
                    var ndaIaCw = new NdaIa(FileMode.WriteOnly);
                    ndaIaCw.LoadFromTemplate(Template.IaAddCw, propsCw);
                    ndaIaCw.Save(String.Format("{0}Nda_Ia_Cw_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                    AddResult("Cw Nda bulk file", ndaIaCw.Path, "nda");

                    var ndaQaCw = new NdaQa(FileMode.WriteOnly);
                    ndaQaCw.LoadFromTemplate(Template.QaAddCw, propsCw);
                    ndaQaCw.Save(String.Format("{0}Nda_Qa_Cw_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                    AddResult("Cw Nda bulk file", ndaQaCw.Path, "nda");

                    var idnNvdrCw = new Idn(FileMode.WriteOnly);
                    idnNvdrCw.LoadFromTemplate(TemplateIdn.CwNvdr, propsCw);
                    idnNvdrCw.Save(String.Format("{0}IDN_Cw_Nvdr_{1}.txt", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                    AddResult("Cw Idn bulk file", idnNvdrCw.Path, "idn");

                    var idnMain = new Idn(FileMode.WriteOnly);
                    idnMain.LoadFromTemplate(TemplateIdn.CwMain, propsCw);
                    idnMain.Save(String.Format("{0}IDN_Cw_Main_{1}.txt", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                    AddResult("Cw Idn bulk file", idnMain.Path, "idn");

                    #region [error in F# dll in this "TemplateFm.ThFmCw"]
                    //var fmCw = new Fm(Format.Horizontal, FileMode.WriteOnly);
                    //fmCw.LoadFromTemplate(TemplateFm.ThFmCw, propsCw);
                    //fmCw.Save(String.Format("{0}Thailand_Fm_Cw_{1}.xls", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                    //AddResult("Cw Fm bulk file", fmCw.Path, "fm");
                    #endregion

                    string path = String.Format("{0}Thailand_Fm_Cw_{1}.xls", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM"));
                    generateXls(propsCw, path);
                    AddResult("Cw Fm bulk file", path, "fm");

                    var tcCw = new Tc(FileMode.WriteOnly);
                    tcCw.LoadFromTemplate(TemplateTc.WrtAddCw, propsCw);
                    tcCw.Save(String.Format("{0}Wrt_Add_Cw_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                    AddResult("Cw t&c bulk file", tcCw.Path, "t&c");
                }
                else
                {
                    LogMessage("No Cw announcement today");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Task failed, error: " + ex.Message, ex);
            }
        }

        private void generateXls(List<Dictionary<string, string>> propsCw, string path)
        {
            var dicList = new Dictionary<string, List<string>>();
            var listTile = new List<string>()
            {
                "Effective Date",
                "SYMBOL",
                "DSPLY_NAME",
                "RIC",
                "OFFCL_CODE",
                "EX_SYMBOL",
                "BCKGRNDPAG",
                "BCAST_REF",
                "#INSTMOD_EXPIR_DATE",
                "#INSTMOD_LONGLINK1",
                "#INSTMOD_LONGLINK2",
                "#INSTMOD_MATUR_DATE",
                "#INSTMOD_OFFC_CODE2",
                "#INSTMOD_STRIKE_PRC",
                "#INSTMOD_WNT_RATIO",
                "#INSTMOD_MNEMONIC",
                "#INSTMOD_TDN_SYMBOL",
                "#INSTMOD_LONGLINK3",
                "EXL_NAME",
                "Old Chain",
                "New BCU",
                "NDA Common Name",
                "Primary Listing",
                "Organisation Name DIRNAME",
                "Issue Date",
                "First Exercise Date",
                "Last Exercise Date",
                "Outstanding Warrant Quantity",
                "Exercise Period"
            };
            dicList.Add("title", listTile);

            //foreach (var dic in propsCw)
            for (int index = 0; index < propsCw.Count; index++)
            {
                List<string> listBody = propsCw[index].Values.ToList();
                dicList.Add(index.ToString(), listBody);
            }

            if (dicList == null || dicList.Count <= 1)
                return;

            XlsOrCsvUtil.GenerateXls0rCsv(path, dicList);
        }

        #endregion

        private Dictionary<string, string> CleanInfos(Dictionary<string, string> oldInfos)
        {
            if (oldInfos["Listing"] == "Warrant")
            {
                return GetCwInfos(oldInfos);
            }
            var newInfos = new Dictionary<string, string>
            {
                {"type", "other"}
            };
            return newInfos;
        }

        private Dictionary<string, string> GetCwInfos(Dictionary<string, string> oldInfos)
        {
            var newInfos = new Dictionary<string, string>
            {
                {"type", oldInfos["Listing"]},
                {"market", oldInfos["Secondary market"]},
                {"price", oldInfos["Exercise price (baht per share)"]},
                {"ratio", oldInfos["Exercise ratio (Common stock)"]},
                {"symbol", oldInfos["Warrant trading symbol"]},
                {"warrantnumber", oldInfos["Warrant trading symbol"].Substring(oldInfos["Warrant trading symbol"].Length - 1, 1)},
                {"tradingdate", oldInfos["Trading date"]},
                {"lastexercisedateLong", oldInfos["Last exercise date"]},
                {"lastexercisedateLongWrt", oldInfos["Last exercise date"].Replace("/", "-")},
                {"firstexercisedate", oldInfos["First exercise date"]}, 
                {"warrantissuer", "100616405"},
                {"assetclass", "COWNT"},
                {"cp", "C"},
                {
                    "lastexercisedate",
                    DateTime.ParseExact(oldInfos["Last exercise date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .ToString("dd/MM/yyyy")
                },
                {
                    "lastexercisedateShortYear",
                    DateTime.ParseExact(oldInfos["Last exercise date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .ToString("dd-MMM-yy")
                },
                {
                    "firstExerciseDateShortYear",
                    DateTime.ParseExact(oldInfos["First exercise date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .ToString("dd-MMM-yy")
                },
                {
                    "tradingDateShortYear",
                    DateTime.ParseExact(oldInfos["Trading date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .ToString("dd-MMM-yy")
                },
                {
                    "lastexercisedateShort",
                    DateTime.ParseExact(oldInfos["Last exercise date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .ToOADate().ToString()
                },
                {
                    "issuedateWrt",
                    DateTime.ParseExact(oldInfos["Trading date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .ToString("dd/MM/yyyy")
                },
                {
                    "issuedate",
                    DateTime.ParseExact(oldInfos["Last exercise date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .ToString("dd/MM/yyyy")
                },
                {
                    "issueDateShortYear",
                    DateTime.ParseExact(oldInfos["First exercise date"], "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                        .AddDays(1)
                        .ToString("dd-MMM-yy")
                },
                {
                    "number",
                    oldInfos["Number of listed warrants ("]
                    .Remove(
                        oldInfos["Number of listed warrants ("].IndexOf(" "))
                    .Replace(
                        ",", "")
                },
                {"name", oldInfos["Company name"].Substring(0, oldInfos["Company name"].LastIndexOf("("))},
                {
                    "abbr",
                    oldInfos["Company name"].Substring(oldInfos["Company name"].LastIndexOf("("))
                        .Replace("(", "")
                        .Replace(")", "")
                }
            };
            newInfos.Add("bigexpiredate", newInfos["lastexercisedate"].ToUpper());
            newInfos.Add("rawnumber", newInfos["number"].Replace(",", ""));
            newInfos.Add("symbolDot", newInfos["symbol"].Replace("-", "."));
            newInfos.Add("bigdate", GetBigDate(newInfos["lastexercisedate"], newInfos["price"]));
            return newInfos;
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

        private string GetDisplayName(string start, string cp, string abbr)
        {
            string res = start + " BY " + abbr;
            if (start.Length + abbr.Length == 9)
            {
                if (cp == "C")
                {
                    res += "C";
                }
                else
                {
                    res += "P";
                }
            }
            else if (start.Length + abbr.Length == 8)
            {
                if (cp == "C")
                {
                    res += " C";
                }
                else
                {
                    res += " P";
                }
            }
            else
            {
                if (cp == "C")
                {
                    res += " CA";
                }
                else
                {
                    res += " PU";
                }
            }
            res += "#A";
            return res;
        }

        private string GetBigDate(string date, string price)
        {
            if (date.Contains(" "))
            {
                date = date.TrimStart().Remove(date.IndexOf(' '));
            }
            DateTime bigDate = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture);
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
                                                                    .Replace(" (Update)", "")
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

            foreach (string url in Enumerable.Range(0, 10).Select(i => string.Format("http://www.set.or.th/set/newslist.do?to={0}%2F{1}%2F{2}&headline=adds+new+listed&submit=Search&symbol=&currentpage={3}&from={0}%2F{1}%2F{2}&newsType=&country=US&exchangeSymbols=&company=false&exchangeNews=on&language=en&exchange=true"
                , "11", "06", DateTime.Now.Year, i)))
            //, DateTime.Now.AddDays(0).ToString("dd"), DateTime.Now.AddDays(0).ToString("MM"), DateTime.Now.Year, i)))
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
