using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;
using ETI.Core;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks
{
    public class ELWDrop : GeneratorBase
    {       
        private List<ELWFMDropModel> dropList = new List<ELWFMDropModel>();    
        private Hashtable tableOfPilc = null;
        bool updatedTagPilc = true;
     
        private KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig configObj = null;
        public ELWDrop(List<TaskResultEntry> resultList, Logger Logger)
        {
            this.TaskResultList = resultList;
            this.Logger = Logger;
        }
        private void Initialize(KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig obj)
        {
            configObj = obj;          
        }

        public int StartELWDropJob(KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig obj)
        {
            Initialize(obj);           
            GrabOrgSourceDataFromWebpage();
            if (dropList.Count > 0)
            {
                tableOfPilc = ReadReferenceTableOfPlic();
                FormatELWFMDropModel();
                GetISINFromWebpage();
                GenerateELWDropFmFile();
               
                GenerateDropGeda();
                GenerateNDA();

                if (updatedTagPilc)
                {
                    UpdateELWDropDb();
                }
            }
            else
            {
                Logger.Log("No drop announcement.");
            }
            int count = dropList.Count;
            return count;
        }

        private void UpdateELWDropDb()
        {
            var todayric = from drop in dropList
                           select ("'" + drop.RIC + "'");

            string rics = string.Join(",", todayric.ToArray());

            int row = KoreaELWManager.InsertELWDrop(dropList, rics);
            string msge = string.Format("Updated {0} ELW Drop records in database.", row);
            Logger.Log(msge);

           
            row = KoreaELWManager.DeleteELWFM2(rics);
            msge = string.Format("Deleted {0} ELW FM2 record(s) in database. RIC list: {1}.", row, rics);
            Logger.Log(msge);
        }
       
        private Hashtable ReadReferenceTableOfPlic()
        {
            Hashtable tableOfPlic = new Hashtable();
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel applcation could not be created ,please check your office installation is corrected !!";
                Logger.Log(msg, Logger.LogType.Error);
                throw (new Exception(msg));
            }

            try
            {
                string fpath = configObj.TagPilcFile;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, fpath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                //Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet 1", wBook);
                if (wSheet == null)
                {
                    string msg = string.Format("worksheet{0} couldn't be found !!", "Issuer");
                    Logger.Log(msg, Logger.LogType.Error);
                    throw (new Exception(msg));
                }

                string expiryDate = null;
                string newExpiryDateLong = DateTime.Now.ToString("ddMMMyy", new CultureInfo("en-US"));
                string newExpiryDateShort = newExpiryDateLong.Substring(2).ToUpper();

                for (int i = 2; ; i++)
                {
                    string ric = ((Range)wSheet.Cells[i, 2]).Text.ToString();
                    if (!string.IsNullOrEmpty(ric))
                    {
                        PilcTemplate pilc = new PilcTemplate();
                        pilc.ExpiryDate = ((Range)wSheet.Cells[i, 12]).Text.ToString();
                        pilc.QACommonName = ((Range)wSheet.Cells[i, 5]).Text.ToString().ToUpper();
                        pilc.IACommonName = ((Range)wSheet.Cells[i, 11]).Text.ToString();

                        expiryDate = GetExpiryDateFromCommonName(pilc.IACommonName);
                        if (!string.IsNullOrEmpty(expiryDate))
                        {
                            pilc.QACommonName = pilc.QACommonName.Replace(expiryDate.Substring(2).ToUpper(), newExpiryDateShort);
                            pilc.IACommonName = pilc.IACommonName.Replace(expiryDate, newExpiryDateLong);
                        }

                        tableOfPlic[ric] = pilc;
                    }
                    else
                    {
                        break;
                    }
                }
                wBook.Save();
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
                Logger.Log(ex.StackTrace);
                throw ex;
            }
            finally
            {
                excelApp.Dispose();
            }
            return tableOfPlic;
        }      

		private string GetExpiryDateFromCommonName(string name)
		{
			name = name.Trim();
			int idx = name.LastIndexOf(' ');
			if (idx < 0)
			{
				return null;
			}

			name = name.Substring(idx + 1);

			if (!Regex.IsMatch(name, @"\d{2}[a-zA-Z]{3}\d{2}"))
			{
				return null;
			}

			return name;
		}

        private void GrabOrgSourceDataFromWebpage()
        {
            try
            {
                string startDate = configObj.StartDate;
                string endDate = configObj.EndDate;
                string uri = "http://kind.krx.co.kr/disclosure/disclosurebystocktype.do";
                //string postData = string.Format("method=searchDisclosureByStockTypeSub&currentPageSize=15&pageIndex=1&menuIndex=3&orderIndex=1&forward=disclosurebystocktype_sub&elwIsuCd=&elwUly=&lpMbr=&corpNameList=&marketType=&fromData={0}&toData={1}&reportNm=&elwRsnClss=0914&elwRsnClss=0912&elwRsnClss=0915&elwRsnClss=0910", startDate, endDate);

                string postData = String.Format("method=searchDisclosureByStockTypeElwSub&currentPageSize=15&pageIndex={0}&orderMode=1&orderStat=D&forward=disclosurebystocktype_elw_sub&elwIsuCd=&elwUly=&lpMbr=&corpNameList=&fromDate={1}&toDate={2}&reportNm=%EC%83%81%EC%9E%A5%ED%8F%90%EC%A7%80", 0, startDate, endDate);
                
                string pageSource = WebClientUtil.GetDynamicPageSource(uri, 180000, postData);
                HtmlAgilityPack.HtmlDocument htc = new HtmlAgilityPack.HtmlDocument();
                if (!string.IsNullOrEmpty(pageSource))
                    htc.LoadHtml(pageSource);
                if (htc != null)
                {
                    HtmlNode tbodyNode = htc.DocumentNode.SelectNodes("//table")[0].SelectSingleNode(".//tbody");
                    HtmlNodeCollection nodeCollections = tbodyNode.SelectNodes(".//tr");
                    int count = nodeCollections.Count;
                    for (var i = 0; i < count; i++)
                    {
                        var item = nodeCollections[i] as HtmlNode;
                        HtmlNode titleNode = item.SelectSingleNode(".//td[4]/a");
                        if (titleNode == null)
                        {
                            continue;
                        }
                        string title = titleNode.InnerText.Trim().ToString();
                        if (!string.IsNullOrEmpty(title) && title.Equals("주식워런트증권 상장폐지조치"))
                        {
                            string companyname = item.SelectSingleNode(".//td[3]").InnerText;
                            string attribute = string.Empty;
                            attribute = item.SelectSingleNode(".//td[4]/a").Attributes["onclick"].Value.Trim().ToString();
                            attribute = attribute.Split('(')[1].Split(',')[0].Trim(new char[] { ' ', '\'', ',' }).ToString();
                            string url = string.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno=&viewerhost=&viewerport=", attribute);
                            System.Threading.Thread.Sleep(2000);
                            string source = WebClientUtil.GetDynamicPageSource(url, 300000, null);
                            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                            if (!string.IsNullOrEmpty(source))
                                doc.LoadHtml(source);
                            if (doc != null)
                            {
                                //string parameter = doc.DocumentNode.SelectSingleNode(".//div/select[@id='mainDocId']/option[2]").Attributes["value"].Value.Trim().ToString();
                                String parameter = doc.DocumentNode.SelectSingleNode(".//select[@id='mainDoc']/option[2]").Attributes["value"].Value.Trim().ToString();
                                parameter = parameter.Trim().ToString().Replace("|Y", "");  
                               
                                attribute = attribute.Insert(4, "/").Insert(7, "/").Insert(10, "/").Trim().ToString();
                                url = String.Format("http://kind.krx.co.kr/external/{0}/{1}/68955.htm", attribute, parameter);
                                doc = WebClientUtil.GetHtmlDocument(url, 300000, null);
                                if (doc != null)
                                {
                                    string str_pre = doc.DocumentNode.SelectSingleNode(".//pre").InnerText.Trim().ToString();
                                    str_pre = str_pre.Trim().ToString();

                                    int str_judge_start_pos = str_pre.IndexOf("상장폐지사유") + "상장폐지사유".Length;

                                    string str_judge = FormatDataWithPos(str_judge_start_pos, str_pre);

                                    if (str_judge != "최종거래일 도래")
                                    {
                                        int str_effective_date_start_pos = str_pre.IndexOf("상장폐지일") + "상장폐지일".Length;
                                        string str_effective_date = FormatDataWithPos(str_effective_date_start_pos, str_pre);

                                        int str_company_start_pos = str_pre.IndexOf("상장폐지 주식워런트증권 종목명") + "상장폐지 주식워런트증권 종목명".Length; //33
                                        string preLeft = str_pre.Substring(str_company_start_pos);

                                        int indexNum = 1;
                                        string indexSuffix = ".";
                                        string pattern = @"\n.*?(?<IndexNum>\d).*?상장폐지 주식워런트증권 종목명";
                                        Regex regex = new Regex(pattern);
                                        Match match = regex.Match(str_pre);
                                        if (match.Success)
                                        {
                                            indexNum = Convert.ToInt16(match.Groups["IndexNum"].Value);
                                            indexSuffix = match.Value.Replace("상장폐지 주식워런트증권 종목명", "").Replace(indexNum.ToString(), "").Trim(' ', '\n');
                                        }
                                        string nextIndex = (indexNum + 1).ToString() + indexSuffix;
                                        int strCompanyEndPos = preLeft.IndexOf(nextIndex);
                                        string str_company_arr = str_pre.Substring(str_company_start_pos, strCompanyEndPos).Trim('\n', '\r', '\t', ' ');

                                        string[] company_arr = str_company_arr.TrimStart(new char[] { '-', ' ', '▶' }).Split('\n');

                                        for (var x = 0; x < company_arr.Length; x++)
                                        {
                                            ELWFMDropModel ELWDrop = new ELWFMDropModel();
                                            ELWDrop.OrgSource = company_arr[x].Trim(new Char[] { '-', ' ', '▶' }).ToString();
                                            ELWDrop.EffectiveDate = str_effective_date.Trim().ToString();
                                            ELWDrop.UpdateDate = ELWDrop.EffectiveDate;
                                            ELWDrop.Publisher = companyname;
                                            dropList.Add(ELWDrop);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabOrgSourceDataFromWebpage()     : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        public static string FormatDataWithPos(int pos, string strPre)
        {
            string temp = strPre.Trim(new char[] { '\r', '\n', ' ', ',' }).ToString();

            char[] tempArr = temp.ToCharArray();
            string result = "";
            while (tempArr[pos] != '\n')    //||tempArr[pos] != '\r'
            {
                result += tempArr[pos].ToString();
                if ((pos + 1) < tempArr.Length)
                    pos++;
                else
                    break;
            }
            result = result.Trim(new char[] { ' ', ':' });
            return result;
        }

        private void FormatELWFMDropModel()
        {
            try
            {
                for (var i = 0; i < dropList.Count; i++)
                {
                    ELWFMDropModel drop = (ELWFMDropModel)dropList[i];     //미래1507OCI머티콜
                    char[] temp_array = drop.OrgSource.ToCharArray();
                    string str_issuername = "";
                    foreach (var item in temp_array)
                    {
                        int asciiCode = (int)item;
                        if (asciiCode > 47 && asciiCode < 58)
                            break;
                        str_issuername += item;
                    }

                    int no = 0;
                    for (var j = 0; j < temp_array.Length; j++)
                    {
                        int asciiCode = (int)temp_array[j];
                        if (asciiCode > 47 && asciiCode < 58)
                        {
                            no = j;
                            break;
                        }
                    }
                    string CharLen = drop.OrgSource.Substring(no, 4).Trim().ToString();

                    drop.Type = "WNT";

                    string shortname = "";
                    KoreaIssuerInfo issuer = KoreaIssuerManager.SelectIssuer(str_issuername);
                    if (issuer!=null)
                    {
                        drop.Ticker = issuer.IssuerCode2 + CharLen;
                        drop.RIC = issuer.IssuerCode2 + CharLen + ".KS";
                        shortname = issuer.IssuerCode4;
                    }
                    string tempCompany = string.Empty;
                    //현대1882삼성엔지콜
                    if (!drop.OrgSource.Contains("조기종료"))
                    {
                        tempCompany = drop.OrgSource;
                        drop.Comment = "Premature";
                    }
                    else
                    {
                        tempCompany = drop.OrgSource.Replace("조기종료", "");
                        drop.Comment = "KOBA Drop";
                    }                        
                    string last = tempCompany.Substring((no + 4)).Trim().ToString();
                    string str_underlying_Dsply_Nmll = last.Substring(0, (last.Length - 1)).Trim().ToString();
                    if (str_underlying_Dsply_Nmll == "KOSPI200")
                        str_underlying_Dsply_Nmll = "코스피";
                    else if (str_underlying_Dsply_Nmll == "스탠차")
                        str_underlying_Dsply_Nmll = "스탠다드차타드";
                    else if (str_underlying_Dsply_Nmll == "IBK")
                        str_underlying_Dsply_Nmll = "아이비케이";
                    else if (str_underlying_Dsply_Nmll == "HMC")
                        str_underlying_Dsply_Nmll = "에이치엠씨";
                    else if (str_underlying_Dsply_Nmll == "KB")
                        str_underlying_Dsply_Nmll = "케이비";                 
                    string str_call_or_put = last.Substring((last.Length - 1)).Trim().ToString();

                    string idn_name = "***";
                    KoreaUnderlyingInfo underlying = KoreaUnderlyingManager.SelectUnderlying(str_underlying_Dsply_Nmll, KoreaNameType.KoreaNameForDrop);
                    if (underlying == null)
                    {
                        Logger.Log("Can not find underlying info with Korea Name for Drop:" + str_underlying_Dsply_Nmll +". Please input the ISIN.", Logger.LogType.Warning);
                        string isin = InputISIN.Prompt(str_underlying_Dsply_Nmll, "Korea Name For Drop");
                        if (!string.IsNullOrEmpty(isin))
                        {
                            underlying = KoreaUnderlyingManager.SelectUnderlyingByISIN(isin);
                            KoreaUnderlyingManager.UpdateKoreaNameDrop(str_underlying_Dsply_Nmll, isin);
                        }                                              
                    }
                    if (underlying != null)
                    {
                        idn_name = underlying.IDNDisplayNamePart;
                    }
                    if (drop.Comment == "KOBA Drop")
                    {
                        idn_name += "KO";
                    }
                    if (str_call_or_put == "콜")
                        str_call_or_put = "C";
                    else if (str_call_or_put == "풋")
                        str_call_or_put = "P";

                    string _idn_display_name = (shortname + CharLen + idn_name + str_call_or_put).ToString();
                    drop.IDNDisplayName = _idn_display_name;  

                    if (str_issuername == "스탠차")
                        str_issuername = "스탠다드차타드";
                    if (str_issuername == "IBK")
                        str_issuername = "아이비케이";
                    if (str_issuername == "HMC")
                        str_issuername = "에이치엠씨";
                    if (str_issuername == "KB")
                        str_issuername = "케이비";

                    //우리1C83삼성테크콜
                    drop.Issuername = str_issuername;
                    drop.Num = "제" + CharLen + "호";
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in _ELWFMDropModelFormat()    : " + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GetISINFromWebpage()
        {
            try
            {
               // string uri = "http://isin.krx.co.kr/jsp/BA_LT113.jsp";
                string uri = "http://isin.krx.co.kr/jsp/realBoard07.jsp";
                foreach (var item in dropList)
                {
                    string postData = string.Empty;
                    string issuername = HttpUtility.UrlEncode(item.Issuername, Encoding.GetEncoding("euc-kr"));
                    string num = HttpUtility.UrlEncode(item.Num, Encoding.GetEncoding("euc-kr"));
                    //                                                  %c7%f6%b4%eb           %c1%a61730%c8%a3       
                    //postData = string.Format("kind=W&pg_no=1&cb_search_column=co_nm&ef_key_word={0}&ef_isu_nm={1}&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst_dt_from=&ef_lst_dt_to=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&chk_bs410=W", issuername, num);
                    postData = string.Format("kind=&ef_std_cd_grnt_dt_from=&ef_std_cd_grnt_dt_to=&secuGubun=07&lst_yn_all=on&lst_yn1=Y&lst_yn2=N&lst_yn3=D&els_dls_all=on&els_dls1=els&els_dls2=dls&so_gb_all=on&so_gb1=s&so_gb2=o&jp_gb_all=on&jp_gb1=c&jp_gb2=t&jp_gb3=r&jp_gb4=i&hg_gb_all=on&hg_gb1=h&hg_gb2=g&tg_gb_all=on&tg_gb1=x&tg_gb2=z&df_gb_all=on&df_gb1=df1&df_gb2=df2&df_gb3=df3&df_gb4=df4&df_gb5=df5&df_gb6=df6&df_gb7=df7&cb_search_column=co_nm&ef_key_word={0}&ef_iss_inst_cd=&ef_isu_nm={1}&ef_iss_dt_from=&ef_iss_dt_to=&ef_lst",issuername,num);
                    AdvancedWebClient wc = new AdvancedWebClient();
                    string pageSource = WebClientUtil.GetPageSource(wc, uri, 300000, postData);
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    if (!string.IsNullOrEmpty(pageSource))
                        doc.LoadHtml(pageSource);
                    if (doc != null)
                    {
                        HtmlNode node = doc.DocumentNode.SelectNodes("//table/tr/td/table/tr")[1].SelectNodes("td")[1];
                        string isin = string.Empty;
                        if (node != null)
                            isin = node.InnerText.Trim().ToString();
                        if (!string.IsNullOrEmpty(isin))
                            item.ISIN = isin;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetISINFromWebpage()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GenerateELWDropFmFile()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference correct!!!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            try
            {
                if (string.IsNullOrEmpty(ELWDropELWFM1ELWFileBulkGenerate.filename))
                {
                    ELWDropELWFM1ELWFileBulkGenerate.filename = "Korea FM for " + DateTime.Today.ToString("dd-MMM-yyyy", new CultureInfo("en-US")).Replace("-", " ") + " (Morning).xls";
                }
                string ipath = Path.Combine(configObj.FM, ELWDropELWFM1ELWFileBulkGenerate.filename);           // "C:\\Korea_Auto\\ELW_FM\\ELW_Drop\\" + filename;
                               
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet =  wSheet = (Worksheet)wBook.Worksheets[1];                
                if (wSheet == null)
                {
                    string msg = "Error found in PrintFurtherIssueToExcel :(WorkSheet could not be created. Check that your office installation and project reference correct!!!)";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                int startLine = 5;
                while (wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim() != string.Empty) startLine++;

                GenerateExcelFileTitle(wSheet, startLine, "common");
                startLine = startLine + 7;
                AppendDataToFile(wSheet, startLine, "common");

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
                //TaskResultList.Add(new TaskResultEntry(Path.GetFileNameWithoutExtension(ipath), "", ipath, creatFm1Mail()));

            }
            catch (Exception ex)
            {
                string msg = "Error found in _print_ELWFMDroppTemplate : " + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void GenerateExcelFileTitle(Worksheet wSheet, int startLine, string type)
        {
            try
            {
                string columns = type.Equals("master") ? "C1" : "C" + (startLine + 6);
                startLine = type.Equals("master") ? 1 : (startLine + 6);
                if (wSheet.get_Range(columns, Type.Missing).Value2 == null)
                {
                    if (type.Equals("master"))
                    {
                        ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 18;
                        ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 18;
                        ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 18;
                        ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 15;
                        ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 30;
                        ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 20;
                        ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 15;
                        ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 15;
                        ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 15;
                        ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 25;
                        ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 15;
                        ((Range)wSheet.Columns["A:K", System.Type.Missing]).Font.Name = "Arial";
                    }
                    else
                    {
                        wSheet.Cells[(startLine - 1), 1] = "DROP";
                        ((Range)wSheet.Cells[(startLine - 1), 1]).Font.Underline = true;
                        wSheet.Cells[(startLine - 4), 1] = "CHANGE";
                        ((Range)wSheet.Cells[(startLine - 4), 1]).Font.Underline = true;
                    }

                    ((Range)wSheet.Rows[startLine, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)wSheet.Rows[startLine, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

                    wSheet.Cells[startLine, 1] = "Updated Date";
                    wSheet.Cells[startLine, 2] = "Effective Date";
                    wSheet.Cells[startLine, 3] = "RIC";
                    wSheet.Cells[startLine, 4] = "Type";
                    wSheet.Cells[startLine, 5] = "IDN Display Name";
                    wSheet.Cells[startLine, 6] = "ISIN";
                    wSheet.Cells[startLine, 7] = "Ticker";
                    wSheet.Cells[startLine, 8] = "Maturity Date";
                    wSheet.Cells[startLine, 9] = "Comment";
                    if (type.Equals("master"))
                    {
                        wSheet.Cells[startLine, 10] = "CompanyName";
                        wSheet.Cells[startLine, 11] = "Publisher";
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GenerateExcelFileTitle()    : \r\n" + ex.ToString() + "  innerException  : \r\n" + ex.InnerException;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void AppendDataToFile(Worksheet wSheet, int startLine, string type)
        {
            try
            {
                for (var i = 0; i < dropList.Count; i++)
                {
                    ELWFMDropModel dropTemp = (ELWFMDropModel)dropList[i];
                    ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                    wSheet.Cells[startLine, 1] = Convert.ToDateTime(dropTemp.UpdateDate).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine, 2] = Convert.ToDateTime(dropTemp.EffectiveDate).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    wSheet.Cells[startLine, 3] = dropTemp.RIC;
                    wSheet.Cells[startLine, 4] = dropTemp.Type;
                    wSheet.Cells[startLine, 5] = dropTemp.IDNDisplayName;
                    wSheet.Cells[startLine, 6] = dropTemp.ISIN;
                    ((Range)wSheet.Cells[startLine, 7]).NumberFormat = "@";
                    wSheet.Cells[startLine, 7] = dropTemp.Ticker;
                    wSheet.Cells[startLine, 8] = dropTemp.MaturityDate;
                    wSheet.Cells[startLine, 9] = dropTemp.Comment;
                    if (type.Equals("master"))
                    {
                        wSheet.Cells[startLine, 10] = dropTemp.OrgSource;
                        wSheet.Cells[startLine, 11] = dropTemp.Publisher;
                    }
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in AppendDataToFile()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }
        }

        /*=======================================================================================================*/
        private string CreateFolder(string foldername)
        {
            string ipath = string.Empty;
            try
            {
                ipath = Path.Combine(configObj.BulkFile, foldername);
                Common.CreateDirectory(ipath);
            }
            catch (Exception ex)
            {
                string msg = "Connot create the folder for bulk file  : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return ipath;
        }

        private void GenerateDropGeda()
        {
            string foldername = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            string dir = CreateFolder(foldername);
            string filePath = Path.Combine(dir, "KR_ELW_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) +"_DROP_BCU_RIC_REMOVE.txt");
            List<List<string>> res = new List<List<string>>();           
            List<string> title = new List<string>() { "RIC" };

            foreach (var item in dropList)
            {
                List<string> oneRes = new List<string>();
                oneRes.Add(item.RIC);
                res.Add(oneRes);
            }
            FileUtil.WriteOutputFile(filePath, res, title, WriteMode.Append);         
            TaskResultList.Add(new TaskResultEntry(Path.GetFileNameWithoutExtension(filePath), "DROP GEDA File", filePath, FileProcessType.GEDA_BULK_RIC_DELETE));
        }
        #region Drop Nda
        private void GenerateNDA()
        {
            GenerateIA();
            GenerateQA();
        }

        private void GenerateIA()
        {
            string foldername = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            string dir = CreateFolder(foldername);
            string filePath = Path.Combine(dir, "KR_ELW_DROP_jj" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "IAChg.csv");
            List<List<string>> res = new List<List<string>>();        
            List<string> title = new List<string>() { "ISIN", "ASSET COMMON NAME" };
			foreach (var item in dropList)
            {				
				List<string> oneRes = new List<string>();
				oneRes.Add(item.ISIN);
                if (tableOfPilc.Contains(item.RIC))
                {
                    oneRes.Add((tableOfPilc[item.RIC] as PilcTemplate).IACommonName);
                }
                else
                {
                    string msg = string.Format("TAG and PILC file doesn't contains record for RIC:{0}.", item.RIC);
                    Logger.Log(msg, Logger.LogType.Error);
                    oneRes.Add("***");
                }
				res.Add(oneRes);
            }
            FileUtil.WriteOutputFile(filePath, res, title, WriteMode.Append); 
            TaskResultList.Add(new TaskResultEntry(Path.GetFileNameWithoutExtension(filePath), "DROP NDA IA File", filePath, FileProcessType.NDA));
       
        }

        private void GenerateQA()
        {
            string foldername = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            string dir = CreateFolder(foldername);
            string filePath = Path.Combine(dir, "KR_ELW_DROP_jj" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "QAChg.csv");
            List<List<string>> res = new List<List<string>>();           
            List<string> title = new List<string>() { "RIC", "ASSET COMMON NAME", "EXPIRY DATE" };            

			foreach (var item in dropList)
            {				
				List<string> oneRes = new List<string>();
				oneRes.Add(item.RIC);
                if (tableOfPilc.Contains(item.RIC))
                {
                    oneRes.Add((tableOfPilc[item.RIC] as PilcTemplate).QACommonName);
                    updatedTagPilc = false;
                }
                else
                {
                    string msg = string.Format("TAG and PILC file doesn't contains record for RIC:{0}.", item.RIC);
                    Logger.Log(msg, Logger.LogType.Error);
                    oneRes.Add("***");
                }
				oneRes.Add(DateTime.Now.ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
				res.Add(oneRes);

				oneRes = new List<string>();
				oneRes.Add(item.RIC.Replace(".KS", "F.KS"));				
                if (tableOfPilc.Contains(item.RIC))
                {
                    oneRes.Add((tableOfPilc[item.RIC] as PilcTemplate).QACommonName);
                }
                else
                {
                    oneRes.Add("***");
                }
				oneRes.Add(DateTime.Now.ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
				res.Add(oneRes);
            }

            if (!updatedTagPilc)
            {
                MessageBox.Show("Please update the TAG and PILC file and run this task again.");
            }

            FileUtil.WriteOutputFile(filePath, res, title, WriteMode.Append);          
            TaskResultList.Add(new TaskResultEntry(Path.GetFileNameWithoutExtension(filePath), "DROP NDA QA File", filePath, FileProcessType.NDA));
        }
        #endregion

    }
}
