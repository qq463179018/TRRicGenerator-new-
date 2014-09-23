using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Collections;

using System.Globalization;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.Threading;
using HtmlAgilityPack;
using System.Drawing;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Util;


namespace Ric.Tasks.Korea
{
    public class ELWSingleSearch : GeneratorBase
    {
        const string CONFIGFILE_NAME = ".\\Config\\Korea\\KOREA_ELWSearchByISINGenerator.config";
        const string LOGFILE_NAME = "SingleSearch-Log.txt";

        public WarrantTemplate SingleTemp = new WarrantTemplate();
        private List<WarrantTemplate> koreaList = new List<WarrantTemplate>();
        private Hashtable KoreaIssuerNameHash = null;
        private Hashtable koreaUnderlyingHash = null;
        private Logger logger = null;
        private KOREA_ELWSearchByISINGeneratorConfig configObj = null;

        protected override void Start()
        {
            StartSingleSearchFMJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(CONFIGFILE_NAME, typeof(KOREA_ELWSearchByISINGeneratorConfig)) as KOREA_ELWSearchByISINGeneratorConfig;
            logger = new Logger(configObj.LOG_FILE_PATH + LOGFILE_NAME, Logger.LogMode.New);
        }

        public void StartSingleSearchFMJob()
        {
            LoadIssuerAndUnderlyingData();
            GrabDataByISIN();
            GenerateELWFMSingleSearchTemplate_xls();
        }

        private void LoadIssuerAndUnderlyingData()
        {
            KoreaIssuerNameHash = new Hashtable();
            koreaUnderlyingHash = new Hashtable();
            string codeMapPath = ".\\Config\\Korea\\KOREA_ELWFMCodeMap.xml";
            ClassCodeMap codeMapObj = ConfigUtil.ReadConfig(codeMapPath, typeof(ClassCodeMap)) as ClassCodeMap;


            List<IssuerCodeConversion> iList = codeMapObj.IssuerCodeMap.ToList();
            List<UnderlyingCodeConversion> uList = codeMapObj.UnderlyingCodeMap.ToList();

            foreach (var item in iList)
            {
                if (!KoreaIssuerNameHash.Contains(item.DSPLY_NMLL))
                {
                    item.fullname = item.fullname.Contains("@") ? item.fullname.Replace("@", "&") : item.fullname;
                    item.orgname = item.orgname.Contains("@") ? item.orgname.Replace("@", "&") : item.orgname;
                    item.shortname = item.shortname.Contains("@") ? item.shortname.Replace("@", "&") : item.shortname;
                    item.NDA_TC = item.NDA_TC.Contains("@") ? item.NDA_TC.Replace("@", "&") : item.NDA_TC;
                    KoreaIssuerNameHash.Add(item.DSPLY_NMLL, item);
                }
            }

            foreach (var item in uList)
            {
                if (!koreaUnderlyingHash.Contains(item.KoreaName))
                {
                    item.QACommonName = item.QACommonName.Contains("@") ? item.QACommonName.Replace("@", "&") : item.QACommonName;
                    item.IDNDisplayName = item.IDNDisplayName.Contains("@") ? item.IDNDisplayName.Replace("@", "&") : item.IDNDisplayName;
                    item.NDA_TC = item.NDA_TC.Contains("@") ? item.NDA_TC.Replace("@", "&") : item.NDA_TC;
                    koreaUnderlyingHash.Add(item.KoreaName, item);
                }
            }
        }

        private void GrabDataByISIN()
        {
            List<String> lists = configObj.ISINList.ToList();
            if (lists.Count > 0)
            {
                foreach (var item in lists)
                {
                    if (String.IsNullOrEmpty(item)) continue;
                    String uri = String.Format("http://isin.krx.co.kr/jsp/BA_VW021.jsp?isu_cd={0}&modi=f&req_no=", item);
                    HtmlDocument htc = WebClientUtil.GetHtmlDocument(uri, 300000);
                    if (htc != null)
                    {
                        HtmlNode table = htc.DocumentNode.SelectNodes("//table")[2];
                        if (table != null)
                        {
                            String tr5_td2 = table.SelectSingleNode(".//tr[5]/td[2]").InnerText.Trim().ToString();
                            //String tr3_td2 = table.SelectSingleNode(".//tr[3]/td[2]").InnerText.Trim().ToString();
                            if (tr5_td2.Equals("사모")) continue;
                            WarrantTemplate wt = new WarrantTemplate();
                            wt.ISIN = item.Trim().ToString();
                            DataAnalysisForSingleSearch(table, wt);
                            koreaList.Add(wt);
                        }
                    }
                }
            }
        }

        private void DataAnalysisForSingleSearch(HtmlNode table, WarrantTemplate wt)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            try
            {
                String tr2_td4 = table.SelectSingleNode(".//tr[2]/td[4]").InnerText.Trim();
                String tr3_td4 = table.SelectSingleNode(".//tr[3]/td[4]").InnerText.Trim();
                String tr4_td2 = table.SelectSingleNode(".//tr[4]/td[2]").InnerText.Trim();
                String tr4_td4 = table.SelectSingleNode(".//tr[4]/td[4]").InnerText.Trim();
                String tr6_td2 = table.SelectSingleNode(".//tr[6]/td[2]").InnerText.Trim();
                String tr6_td4 = table.SelectSingleNode(".//tr[6]/td[4]").InnerText.Trim();
                String tr7_td2 = table.SelectSingleNode(".//tr[7]/td[2]").InnerText.Trim();
                String tr7_td4 = table.SelectSingleNode(".//tr[7]/td[4]").InnerText.Trim();
                String tr8_td2 = table.SelectSingleNode(".//tr[8]/td[2]").InnerText.Trim();
                String tr11_td4 = table.SelectSingleNode(".//tr[11]/td[4]").InnerText;
                String tr13_td2 = table.SelectSingleNode(".//tr[13]/td[2]").InnerText;
                String tr20_td4 = table.SelectSingleNode(".//tr[20]/td[4]").InnerText.Trim();
                String tr21_td2 = table.SelectSingleNode(".//tr[21]/td[2]").InnerText.Trim();

                if (String.IsNullOrEmpty(tr13_td2))
                    tr13_td2 = "KOSPI200";
                else
                    tr13_td2 = tr13_td2.Contains("(주)") ? tr13_td2.Replace("(주)", "") : tr13_td2;
                wt.Underlying_Korea_Name = tr13_td2;

                string issuer_koreaname = tr3_td4;
                Char[] array_issuer = issuer_koreaname.ToCharArray();
                string ikoreaname = String.Empty;
                foreach (var item in array_issuer)
                {
                    if (item > 48 && item < 58) break;
                    ikoreaname += item.ToString();
                }
                wt.Issuer_Korea_Name = ikoreaname;

                wt.Updated_Date = DateTime.Today.ToString("dd-MMM-yy");
                wt.Effective_Date = DateTime.Today.Year.ToString();
                wt.FM = "1";
                wt.Ticker = tr2_td4.Substring(1);
                wt.RIC = wt.Ticker + ".KS";
                wt.Org_Mat_date = tr6_td4.Trim().ToString();
                wt.Mat_Date = Convert.ToDateTime(tr6_td4.Trim()).ToString("dd-MMM-yy");
                wt.Strike_Price = tr20_td4.Contains(",") ? tr20_td4.Replace(",", "") : tr20_td4;
                wt.Quanity_of_Warrants = tr7_td4.Contains(",") ? tr7_td4.Replace(",", "") : tr7_td4;
                wt.Issue_Price = tr7_td2.Contains(",") ? tr7_td2.Replace(",", "") : tr7_td2;
                wt.Org_Issue_date = tr6_td2;
                wt.Issue_Date = Convert.ToDateTime(tr6_td2).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                wt.Conversion_Ratio = tr8_td2;
                wt.Issuer = tr4_td2.ToUpper();
                wt.Korea_Warrant_Name = tr3_td4;
                if (!string.IsNullOrEmpty(tr21_td2))
                    wt.KnockOut_price = tr21_td2.Contains(",") ? tr21_td2.Replace(",", "") : tr21_td2;


                int count = tr4_td4.Split('-').Length;
                String PorC = tr4_td4.Split('-')[(count - 1)];
                wt.CallOrPut = PorC.Equals("C") ? "CALL" : "PUT";

                CreateNewIDNDisplayName(wt, PorC);

                CreateNewBcastRef(wt);

                CreateNewQACommonName(wt, PorC);

                CreateTheNewChain(wt);

            }
            catch (Exception ex)
            {
                String msg = "Error found in DataAnalysisForSingleSearch    :InnerException  \r\n" + ex.InnerException + "  :  \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void CreateNewIDNDisplayName(WarrantTemplate wt, String PorC)
        {
            try
            {
                //IDN Display Name
                if (KoreaIssuerNameHash.Contains(wt.Issuer_Korea_Name))
                {
                    IssuerCodeConversion iCode = ((IssuerCodeConversion)KoreaIssuerNameHash[wt.Issuer_Korea_Name]);
                    String shortname = iCode.shortname;
                    if (!string.IsNullOrEmpty(shortname))
                    {
                        String NumLen = wt.Ticker.Substring(2, 4);
                        String underly = String.Empty;
                        if (!String.IsNullOrEmpty(wt.Underlying_Korea_Name))
                        {
                            if (koreaUnderlyingHash.Contains(wt.Underlying_Korea_Name))
                                underly = ((UnderlyingCodeConversion)koreaUnderlyingHash[(wt.Underlying_Korea_Name)]).IDNDisplayName;
                            else
                                underly = "***";
                        }
                        String IDN = String.IsNullOrEmpty(wt.KnockOut_price) ? (shortname + NumLen + underly + PorC) : (shortname + NumLen + underly + "KO" + PorC);
                        wt.IDN_Display_Name = IDN;
                    }
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in create new IDN Display Name   : \r\n" + ex.ToString() + "\r\n  InnerException   :   " + ex.InnerException;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void CreateNewBcastRef(WarrantTemplate wt)
        {
            try
            {
                //BCAST_REF
                String underlying_ric = String.Empty;
                if (!String.IsNullOrEmpty(wt.Underlying_Korea_Name))
                {
                    if (koreaUnderlyingHash.Contains(wt.Underlying_Korea_Name))
                        underlying_ric = ((UnderlyingCodeConversion)koreaUnderlyingHash[(wt.Underlying_Korea_Name)]).UnderlyingRic;
                    else
                        underlying_ric = "***";
                }
                wt.BCAST_REF = underlying_ric.ToUpper();
            }
            catch (Exception ex)
            {
                String msg = "Error found in create new Bcast Ref   : \r\n" + ex.ToString() + "\r\n  InnerException   :   " + ex.InnerException;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void CreateNewQACommonName(WarrantTemplate wt, String PorC)
        {
            try
            {
                //QA Common Name
                String cname = koreaUnderlyingHash.Contains(wt.Underlying_Korea_Name) ? ((UnderlyingCodeConversion)koreaUnderlyingHash[(wt.Underlying_Korea_Name)]).QACommonName : "***";
                String sname = KoreaIssuerNameHash.Contains(wt.Issuer_Korea_Name) ? ((IssuerCodeConversion)KoreaIssuerNameHash[(wt.Issuer_Korea_Name)]).shortname : "***";
                String mtime = Convert.ToDateTime(wt.Org_Mat_date).ToString("MMM-yy", new CultureInfo("en-US")).Replace("-", "").ToUpper();
                String price = wt.Strike_Price.Contains(".") ? wt.Strike_Price.Split('.')[0] : wt.Strike_Price;
                price = price.Length >= 4 ? price.Substring(0, 4) : price;
                String last = wt.Underlying_Korea_Name.Equals("KOSPI200") ? "IW" : "WNT";
                String qacommonname = String.Empty;
                if (String.IsNullOrEmpty(wt.KnockOut_price))
                    qacommonname = cname + " " + sname + " " + mtime + " " + price + " " + PorC + last;
                else
                    qacommonname = cname + " " + sname + " " + mtime + " " + price + " KO " + PorC + last;
                wt.QA_Common_Name = qacommonname.ToUpper();
            }
            catch (Exception ex)
            {
                String msg = "Error found in create new QA Common Name   : \r\n" + ex.ToString() + "\r\n  InnerException   :   " + ex.InnerException;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void CreateTheNewChain(WarrantTemplate wt)
        {
            //Chain
            try
            {
                String chain = String.Empty;
                if (String.IsNullOrEmpty(wt.KnockOut_price))
                    chain = wt.Underlying_Korea_Name.Equals("KOSPI200") ? ("0#WARRANTS.KS, 0#ELW.KS, 0#.KS200W.KS") : "0#WARRANTS.KS, 0#ELW.KS, 0#CELW.KS, 0#" + wt.BCAST_REF.Split('.')[0] + "W." + wt.BCAST_REF.Split('.')[1];
                else
                    chain = "0#KOBA.KS";
                wt.Chain = chain.ToUpper();
            }
            catch (Exception ex)
            {
                String msg = "Error found in create new Chain   : \r\n" + ex.ToString() + "\r\n  InnerException   :   " + ex.InnerException;
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GenerateELWFMSingleSearchTemplate_xls()
        {
            if (koreaList.Count > 0)
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                ExcelApp excelApp = new ExcelApp(false, false);
                if (excelApp.ExcelAppInstance == null)
                {
                    Logger.Log("Excel could not be started. Check that your office installation and project reference are correct!", Logger.LogType.Error);
                    return;
                }

                try
                {
                    String filename = "Search by ISIN " + DateTime.Today.ToString("yyyy-MM-dd") + ".xls";
                    String ipath = configObj.Korea_SearchByISIN_GenerateFileConfig.WORKBOOK_PATH + filename;
                    Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                    Worksheet wSheet = ExcelUtil.GetWorksheet(configObj.Korea_SearchByISIN_GenerateFileConfig.WORKSHEET_NAME, wBook);
                    if (wSheet == null)
                    {
                        Logger.Log("Worksheet could not be created. Check that your office installation and project reference are correct!", Logger.LogType.Error);
                        return;
                    }

                    CreateExcelTitle(wSheet);

                    int startLine = 2;
                    while (wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString() != String.Empty) startLine++;

                    LoopPrintKoreaListData(wSheet, startLine);

                    excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                    wBook.Save();
                }
                catch (Exception ex)
                {
                    Logger.Log("Error found in GenerateELWFMSingleSearchTemplate_xls : " + ex.ToString(), Logger.LogType.Warning);
                    return;
                }
                finally
                {
                    excelApp.Dispose();
                }
            }
        }

        private void CreateExcelTitle(Worksheet wSheet)
        {
            if (wSheet.get_Range("C1", Type.Missing).Value2 == null)
            {
                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 5;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 35;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 60;
                ((Range)wSheet.Columns["Q", System.Type.Missing]).ColumnWidth = 35;
                ((Range)wSheet.Columns["R", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["S", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["A:S", System.Type.Missing]).Font.Name = "Arial";

                ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

                wSheet.Cells[1, 1] = "Updated Date";
                wSheet.Cells[1, 2] = "Effective Date";
                wSheet.Cells[1, 3] = "RIC";
                wSheet.Cells[1, 4] = "FM";
                wSheet.Cells[1, 5] = "IDN Display Name";
                wSheet.Cells[1, 6] = "ISIN";
                wSheet.Cells[1, 7] = "Ticker";
                wSheet.Cells[1, 8] = "BCAST_REF";
                wSheet.Cells[1, 9] = "QA Common Name";
                wSheet.Cells[1, 10] = "Mat Date";
                wSheet.Cells[1, 11] = "Strike Price";
                wSheet.Cells[1, 12] = "Quanity of Warrants";
                wSheet.Cells[1, 13] = "Issue Price";
                wSheet.Cells[1, 14] = "Issue Date";
                wSheet.Cells[1, 15] = "Conversion Ratio";
                wSheet.Cells[1, 16] = "Issuer";
                wSheet.Cells[1, 17] = "Exchange Warrant Name";
                wSheet.Cells[1, 18] = "Chain";
                wSheet.Cells[1, 19] = "Knock-out Price";
            }
        }

        private void LoopPrintKoreaListData(Worksheet wSheet, int startLine)
        {
            foreach (var item in koreaList)
            {
                ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                wSheet.Cells[startLine, 1] = item.Updated_Date;
                wSheet.Cells[startLine, 2] = item.Effective_Date;
                wSheet.Cells[startLine, 3] = item.RIC;
                wSheet.Cells[startLine, 4] = item.FM;
                wSheet.Cells[startLine, 5] = item.IDN_Display_Name;
                wSheet.Cells[startLine, 6] = item.ISIN;
                ((Range)wSheet.Cells[startLine, 7]).NumberFormat = "@";
                wSheet.Cells[startLine, 7] = item.Ticker;
                wSheet.Cells[startLine, 8] = item.BCAST_REF;
                wSheet.Cells[startLine, 9] = item.QA_Common_Name;
                ((Range)wSheet.Cells[startLine, 10]).NumberFormat = "@";
                wSheet.Cells[startLine, 10] = item.Mat_Date;
                wSheet.Cells[startLine, 11] = item.Strike_Price;
                wSheet.Cells[startLine, 12] = item.Quanity_of_Warrants;
                wSheet.Cells[startLine, 13] = item.Issue_Price;
                ((Range)wSheet.Cells[startLine, 14]).NumberFormat = "@";
                wSheet.Cells[startLine, 14] = item.Issue_Date;
                wSheet.Cells[startLine, 15] = item.Conversion_Ratio;
                wSheet.Cells[startLine, 16] = item.Issuer;
                wSheet.Cells[startLine, 17] = item.Korea_Warrant_Name;
                wSheet.Cells[startLine, 18] = item.Chain;
                wSheet.Cells[startLine, 19] = item.KnockOut_price;
                startLine++;
            }
        }





        /*--------------------------------------Useless Code------------------------------------*/
        /*[Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]*/
        /*private void GrabDataByISIN()
        {
            selenium.SetTimeout("600000");
            selenium.OpenWindow(defaultURI, "kwarrant");
            selenium.WaitForPopUp("kwarrant", "600000");
            selenium.SelectWindow("kwarrant");
            Thread.Sleep(3000);
            selenium.Click("I1");
            selenium.WaitForPageToLoad("600000");
            selenium.SelectFrame("main");
            selenium.Type("ef_isu_nm", configObj.Korea_ELWFMSingleSearch_GenerateFile_CONFIG.InputISIN);//SingleTemp.ISIN
            selenium.Click("chk_bs410");
            selenium.Click("button1");
            selenium.WaitForPageToLoad("600000");

            try
            {
                bool flag = selenium.IsElementPresent("//tr[2]/td[1]");
                if (flag)
                {
                    String attribute = selenium.GetAttribute("//tr[8]/td/table/tbody/tr[2]/td[1]/a@href");
                    String varTime = attribute.Split('(')[1].Split(',')[0].Split('\'')[1];
                    String varIsin = attribute.Split('(')[1].Split(',')[1].Split('\'')[1];
                    String uri = String.Format("http://isin.krx.co.kr/jsp/BA_VW021.jsp?isu_cd={0}&modi=t&req_no={1}", varIsin, varTime);
                    HtmlDocument doc = WebClientUtil.GetHtmlDocument(uri, 30000);
                    String judge = doc.DocumentNode.SelectSingleNode("//tr[5]/td[2]").InnerText;
                    if (judge == "사모")
                    {
                        String msg = "There is a '사모' in this page";
                        Logger.Log(msg, Logger.LogType.Error);
                        return;
                    }
                    SingleTemp.ISIN = varIsin;
                    String tr3_td2 = doc.DocumentNode.SelectSingleNode("//tr[3]/td[2]").InnerText;
                    if (tr3_td2.Split(' ')[1].Equals("조기종료"))
                    {
                        String knockout_price = doc.DocumentNode.SelectSingleNode("//tr[21]/td[2]").InnerText;
                        SingleTemp.KnockOut_price = knockout_price.IndexOf('.') > 0 ? knockout_price.Replace(".", "").ToString() : knockout_price;
                    }

                    DataAnalysisForSingleSearch(doc);

                    selenium.SelectWindow("kwarrant");
                    selenium.WaitForPopUp("kwarrant", "6000");
                    selenium.Close();
                    selenium.Stop();
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GrabDataByISIN :" + ex.ToString(), Logger.LogType.Error);
                return;
            }
        }*/

        /*private void DataAnalysisForSingleSearch(HtmlDocument doc)
        {
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                String tr2_td4 = doc.DocumentNode.SelectSingleNode("//tr[2]/td[4]").InnerText;
                String tr3_td4 = doc.DocumentNode.SelectSingleNode("//tr[3]/td[4]").InnerText;
                String tr4_td2 = doc.DocumentNode.SelectSingleNode("//tr[4]/td[2]").InnerText;
                String tr4_td4 = doc.DocumentNode.SelectSingleNode("//tr[4]/td[4]").InnerText;
                String tr6_td2 = doc.DocumentNode.SelectSingleNode("//tr[6]/td[2]").InnerText;
                String tr6_td4 = doc.DocumentNode.SelectSingleNode("//tr[6]/td[4]").InnerText;
                String tr7_td2 = doc.DocumentNode.SelectSingleNode("//tr[7]/td[2]").InnerText;
                String tr7_td4 = doc.DocumentNode.SelectSingleNode("//tr[7]/td[4]").InnerText;
                String tr8_td2 = doc.DocumentNode.SelectSingleNode("//tr[8]/td[2]").InnerText;
                String tr11_td4 = doc.DocumentNode.SelectSingleNode("//tr[11]/td[4]").InnerText;
                String tr13_td2 = doc.DocumentNode.SelectSingleNode("//tr[13]/td[2]").InnerText;
                String tr20_td4 = doc.DocumentNode.SelectSingleNode("//tr[20]/td[4]").InnerText;

                if (tr13_td2 == string.Empty)
                    tr13_td2 = "KOSPI200";
                if (tr13_td2.Contains("(주)"))
                    tr13_td2 = tr13_td2.Replace("(주)", "");

                SingleTemp.Underlying_Korea_Name = tr13_td2;

                String issuer_koreaname = tr3_td4;
                Char[] str_issuer = issuer_koreaname.ToCharArray();
                String ikoreaname = "";
                foreach (var item in str_issuer)
                {
                    int x = (int)item;
                    if (item > 48 && item < 58) break;
                    ikoreaname += item.ToString();
                }
                SingleTemp.Issuer_Korea_Name = ikoreaname;

                SingleTemp.Updated_Date = DateTime.Today.ToString("dd-MMM-yy");
                SingleTemp.Effective_Date = DateTime.Today.Year.ToString();
                SingleTemp.FM = "1";

                SingleTemp.RIC = tr2_td4.Substring(1) + ".KS";
                //SingleTemp.ISIN = tr2_td2;
                SingleTemp.Ticker = tr2_td4.Substring(1).ToString();
                SingleTemp.Org_Mat_date = tr6_td4.Trim().ToString();
                String mat = tr6_td4.Trim().ToString();//tr6_td4.Split('-')[2] + "-" + (Convert.ToDateTime(tr6_td4.Split('-')[0] + "-" + tr6_td4.Split('-')[1]).ToString("MMM-yy"));
                SingleTemp.Mat_Date = Convert.ToDateTime(mat).ToString("dd-MMM-yy"); //DateTime.ParseExact(tr22_td4, "dd-MMM-yy", null).ToString();
                String strike_Price = tr20_td4.Replace(",", ""); //(tr20_td4.IndexOf('.') > 0) == true ? tr20_td4.Replace(",", "").Split('.')[0] : tr20_td4.Replace(",", "");
                SingleTemp.Strike_Price = strike_Price;
                SingleTemp.Quanity_of_Warrants = tr7_td4.Replace(",", "");
                String issue_price = (tr7_td2.IndexOf('.') > 0) == true ? tr7_td2.Replace(",", "").Split('.')[0] : tr7_td2.Replace(",", "");
                SingleTemp.Issue_Price = issue_price;
                SingleTemp.Org_Issue_date = tr6_td2.Trim().ToString();
                SingleTemp.Issue_Date = Convert.ToDateTime(tr6_td2).ToString("dd-MMM-yy");  //DateTime.ParseExact(tr6_td2, "dd-MMM-yy", null).ToString();
                SingleTemp.Conversion_Ratio = tr8_td2;
                SingleTemp.Issuer = tr4_td2.ToUpper();
                SingleTemp.Korea_Warrant_Name = tr3_td4;

                //混合的
                String org = tr4_td2.Substring(0, 5);
                int accout = tr4_td4.Split('-').Count();
                String last = tr4_td4.Split('-')[(accout - 1)];
                SingleTemp.CallOrPut = last == "C" ? "CALL" : "PUT";

                //IDN Display Name
                if (KoreaIssuerNameHash.Contains(SingleTemp.Issuer_Korea_Name))
                {
                    IssuerCodeConversion iCode = ((IssuerCodeConversion)KoreaIssuerNameHash[SingleTemp.Issuer_Korea_Name]);
                    String shortname = iCode.shortname;
                    if (shortname != String.Empty)
                    {
                        String NumLen = tr2_td4.Substring(3, 4);
                        String underly = "";
                        if (tr11_td4 == String.Empty && tr13_td2 != String.Empty)
                        {
                            if (koreaUnderlyingHash.Contains(tr13_td2))
                                underly = ((UnderlyingCodeConversion)koreaUnderlyingHash[tr13_td2]).IDNDisplayName;
                            else
                                underly = "*****";
                        }
                        //String underly = ((UnderlyingCode)koreaUnderlyingHash[tr13_td2]).IDN_Name == "" ? "*************" : ((UnderlyingCode)koreaUnderlyingHash[tr13_td2]).IDN_Name;
                        String CharLen = tr11_td4 == "KOSPI200" ? "KOSPI" : underly;
                        SingleTemp.IDN_Display_Name = SingleTemp.KnockOut_price != null ? (shortname + NumLen + CharLen + "KO" + last) : (shortname + NumLen + CharLen + last);
                        if (SingleTemp.IDN_Display_Name.Contains("*****"))
                            SingleTemp.IDN_Display_Name = "**********";
                    }
                }



                //BCAST_REF
                String underlying_ric = "";
                if (tr11_td4 == String.Empty && tr13_td2 != String.Empty)
                {
                    SingleTemp.BCAST_REF = tr11_td4 == "KOSPI200" ? ".KS200" : null;
                    if (koreaUnderlyingHash.Contains(tr13_td2))
                        underlying_ric = ((UnderlyingCodeConversion)koreaUnderlyingHash[tr13_td2]).UnderlyingRic;
                    else
                        underlying_ric = "*****";
                }
                else
                {
                    if (tr11_td4.Equals("KOSPI200"))
                        underlying_ric = ".KS200";
                }
                SingleTemp.BCAST_REF = underlying_ric.ToUpper();

                //QA Common Name
                String sname = tr11_td4 == "KOSPI200" ? "KSPI" : (koreaUnderlyingHash.Contains(tr13_td2) == true ? ((UnderlyingCodeConversion)koreaUnderlyingHash[tr13_td2]).QACommonName : "*****");
                String cname = KoreaIssuerNameHash.Contains(SingleTemp.Issuer_Korea_Name) == true ? (((IssuerCodeConversion)KoreaIssuerNameHash[SingleTemp.Issuer_Korea_Name]).shortname) : "****";
                String mtime = SingleTemp.Mat_Date.Split('-')[1] + SingleTemp.Mat_Date.Split('-')[2];
                String price = SingleTemp.Strike_Price.IndexOf('.') > 0 == true ? SingleTemp.Strike_Price.Split('.')[0] : SingleTemp.Strike_Price;
                price = (price.Length > 4) == true ? price.Substring(0, 4) : price;
                //String price = warrantTemp.Strike_Price.Length > 4 == true ? warrantTemp.Strike_Price.Substring(0, 4) : warrantTemp.Strike_Price;
                String slast = tr11_td4 == "KOSPI200" ? "IW" : "WNT";
                SingleTemp.QA_Common_Name = SingleTemp.KnockOut_price != null ? (sname + " " + cname + " " + mtime + " " + price + " " + "KO " + last + slast) : (sname + " " + cname + " " + mtime + " " + price + " " + last + slast);
                SingleTemp.QA_Common_Name = SingleTemp.QA_Common_Name.ToUpper().Trim();

                //Chain
                String chain = tr11_td4 == "KOSPI200" ? ("0#WARRANTS.KS, 0#ELW.KS, 0#.KS200W.KS") : "0#WARRANTS.KS, 0#ELW.KS, " + "0#CELW.KS, 0#" + SingleTemp.BCAST_REF.Split('.')[0] + "W." + SingleTemp.BCAST_REF.Split('.')[1]; ;
                SingleTemp.Chain = chain;//"0#WARRANTS.KS, 0#ELW.KS," + "0#." + warrantTemp.BCAST_REF.Split('.')[0] + "W." + warrantTemp.BCAST_REF.Split('.')[1];
                SingleTemp.Last_Trading_Date = " ";
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in DataAnalysisForSingleSearch : " + ex.ToString(), Logger.LogType.Warning);
                return;
            }
        }*/

        /*private void GenerateELWFMSingleSearchTemplate_xls(WarrantTemplate SingleTemp)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                Logger.Log("Excel could not be started. Check that your office installation and project reference are correct!", Logger.LogType.Error);
                return;
            }

            try
            {
                String filename = SingleTemp.ISIN + ".xls";
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, configObj.Korea_SearchByISIN_GenerateFileConfig.WORKBOOK_PATH + filename);
                Worksheet wSheet = ExcelUtil.GetWorksheet(configObj.Korea_SearchByISIN_GenerateFileConfig.WORKSHEET_NAME, wBook);
                if (wSheet == null)
                {
                    Logger.Log("Worksheet could not be created. Check that your office installation and project reference are correct!", Logger.LogType.Error);
                    return;
                }

                if (wSheet.get_Range("C1", Type.Missing).Value2 == null)
                {
                    ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 15;
                    ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                    ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
                    ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 5;
                    ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 35;
                    ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 20;
                    ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 15;
                    ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 20;
                    ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 65;
                    ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 15;
                    ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 10;
                    ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 15;
                    ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 10;
                    ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 15;
                    ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 15;
                    ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 60;
                    ((Range)wSheet.Columns["Q", System.Type.Missing]).ColumnWidth = 35;
                    ((Range)wSheet.Columns["R", System.Type.Missing]).ColumnWidth = 65;
                    ((Range)wSheet.Columns["S", System.Type.Missing]).ColumnWidth = 20;
                    ((Range)wSheet.Columns["A:S", System.Type.Missing]).Font.Name = "Arial";

                    ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

                    wSheet.Cells[1, 1] = "Updated Date";
                    wSheet.Cells[1, 2] = "Effective Date";
                    wSheet.Cells[1, 3] = "RIC";
                    wSheet.Cells[1, 4] = "FM";
                    wSheet.Cells[1, 5] = "IDN Display Name";
                    wSheet.Cells[1, 6] = "ISIN";
                    wSheet.Cells[1, 7] = "Ticker";
                    wSheet.Cells[1, 8] = "BCAST_REF";
                    wSheet.Cells[1, 9] = "QA Common Name";
                    wSheet.Cells[1, 10] = "Mat Date";
                    wSheet.Cells[1, 11] = "Strike Price";
                    wSheet.Cells[1, 12] = "Quanity of Warrants";
                    wSheet.Cells[1, 13] = "Issue Price";
                    wSheet.Cells[1, 14] = "Issue Date";
                    wSheet.Cells[1, 15] = "Conversion Ratio";
                    wSheet.Cells[1, 16] = "Issuer";
                    wSheet.Cells[1, 17] = "Exchange Warrant Name";
                    wSheet.Cells[1, 18] = "Chain";
                    wSheet.Cells[1, 19] = "Knock-out Price";
                }

                int startLine = 2;
                while (wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString() != String.Empty) startLine++;

                ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                wSheet.Cells[startLine, 1] = SingleTemp.Updated_Date;
                ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                wSheet.Cells[startLine, 2] = SingleTemp.Effective_Date;
                wSheet.Cells[startLine, 3] = SingleTemp.RIC;
                wSheet.Cells[startLine, 4] = SingleTemp.FM;
                wSheet.Cells[startLine, 5] = SingleTemp.IDN_Display_Name;
                ((Range)wSheet.Cells[startLine, 6]).NumberFormat = "@";
                wSheet.Cells[startLine, 6] = SingleTemp.ISIN;
                ((Range)wSheet.Cells[startLine, 7]).NumberFormat = "@";
                wSheet.Cells[startLine, 7] = SingleTemp.Ticker;
                wSheet.Cells[startLine, 8] = SingleTemp.BCAST_REF;
                wSheet.Cells[startLine, 9] = SingleTemp.QA_Common_Name;
                ((Range)wSheet.Cells[startLine, 10]).NumberFormat = "@";
                wSheet.Cells[startLine, 10] = SingleTemp.Mat_Date;
                wSheet.Cells[startLine, 11] = SingleTemp.Strike_Price;
                wSheet.Cells[startLine, 12] = SingleTemp.Quanity_of_Warrants;
                wSheet.Cells[startLine, 13] = SingleTemp.Issue_Price;
                ((Range)wSheet.Cells[startLine, 14]).NumberFormat = "@";
                wSheet.Cells[startLine, 14] = SingleTemp.Issue_Date;
                wSheet.Cells[startLine, 15] = SingleTemp.Conversion_Ratio;
                wSheet.Cells[startLine, 16] = SingleTemp.Issuer;
                wSheet.Cells[startLine, 17] = SingleTemp.Korea_Warrant_Name;
                wSheet.Cells[startLine, 18] = SingleTemp.Chain;
                wSheet.Cells[startLine, 19] = SingleTemp.KnockOut_price;

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateELWFMSingleSearchTemplate_xls : " + ex.ToString(), Logger.LogType.Warning);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }*/
    }
}
