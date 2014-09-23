using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections;
using System.ComponentModel;
using Selenium;
using Ric.Core;
using Ric.Util;
using System.IO;
using System.Net;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Globalization;
using System.Drawing.Design;

namespace Ric.Tasks.Thailand
{
    public struct TailandRicTemplate
    {
        public String dwIssuerStr;
        public String dwIssueShortNameStr;
        public String totalRicsStr;
        public String exchangeStr;
        public String effectiveDateStr;
        public String ricStr;
        public String displayNameStr;
        public String officialCodeStr;
        public String offc_code2Str;
        public String currencyStr;
        public String recordtypeStr;
        public String chainRicStr;
        public String chainRicOnListingDateStr;
        public String bcastRefStr;
        public String underlyingRicStr;
        public String longlink2Str;
        public String wntRatioStr;
        public String strikePrcStr;
        public DateTime maturDateDT;
        public String exerciseDateStr;
        public String lastTradingDateStr;
        public String numberOfWarrants;
    }

    public class ThailandGeneratorConfig
    {
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> TailandRicFields { get; set; }
        public string SavePath { get; set; }
        public string username { get; set; }
        public string password { get; set; }
        public int start_position { get; set; }
        public int end_position { get; set; }
        public string html_url { get; set; }
        public List<string> HeadLineContains { get; set; }
        public string Symbol { get; set; }
        public string NewsTodayUrl { get; set; }

    }

    public class ThailandRicGenerator : GeneratorBase
    {

        private ISelenium selenium;
        private int count = 0;

        private static readonly string CONFIG_FILE_PATH = "Config\\Thailand\\TailandRic.Config";
        private static ThailandGeneratorConfig configObj = null;
        private static Logger logger = null;

        protected override void Initialize()
        {
            base.Initialize();

            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(ThailandGeneratorConfig)) as ThailandGeneratorConfig;

            logger = new Logger(configObj.SavePath, Logger.LogMode.New);
        }


        [DllImport("user32.dll")]
        private static extern void GetWindowThreadProcessId(IntPtr hWnd, out int k);

        protected override void Start()
        {
            StartGrabber();
        }

        /**
         * Public method called by outside to start grab data and generate template for Tailand Market
         * 
         */
        public void StartGrabber()
        {
            //DataCaptureFromSetSmart("PTTG01CA"); //just for test

            HtmlAgilityPack.HtmlDocument htmlDoc = null;
            string url = configObj.NewsTodayUrl;  // "http://www.set.or.th/set/todaynews.do?language=en&country=US";
            SetupTest();
            if (configObj.html_url == "NA")
            {
                string htm = GrabberHTMLFromNewsToday(url);
                htmlDoc = ParseNewsTodayHTML(htm);

                if (htmlDoc != null)
                {
                    SearchSET(htmlDoc);
                }
                //DataCaptureFromSET();
            }
            else
            {
                url = configObj.html_url;
                GrabDataFromInnerHtm(url);
                //DataCaptureByUrl();
            }
            ClearTest();

            if (count == 0)
            {
                logger.LogErrorAndRaiseException("Not found new RIC information");
            }
            else
            {
                logger.LogErrorAndRaiseException(count.ToString() + " Ric template file generated under " + configObj.SavePath);
            }
        }

        private bool IsSymbolContain(string symbol)
        {
            bool isContain = false;
            string countBackSecond = symbol.Substring(symbol.Length - 2, 1);
            foreach (string item in configObj.HeadLineContains)
            {
                if (countBackSecond.Equals(item))
                {
                    isContain = true;
                }
            }
            return isContain;
        }

        private void SearchSET(HtmlAgilityPack.HtmlDocument htmlDoc)
        {
            int startPosition = 5;
            string symbol = htmlDoc.DocumentNode.SelectSingleNode("//tr[" + startPosition + "]/td[3]").InnerText;
            while (!string.IsNullOrEmpty(symbol))
            {
                if (symbol.Equals(configObj.Symbol))
                {
                    string headLine = htmlDoc.DocumentNode.SelectSingleNode("//tr[" + startPosition + "]/td[4]").InnerText;
                    if (headLine.Contains("adds new"))
                    {
                        startPosition += 2;
                        symbol = htmlDoc.DocumentNode.SelectSingleNode("//tr[" + startPosition + "]/td[3]").InnerText;
                        if (IsSymbolContain(symbol))
                        {
                            string url = htmlDoc.DocumentNode.SelectSingleNode("//tr[" + startPosition + "]/td[5]/a[text()='HTML']@href").InnerText;

                            GrabDataFromInnerHtm(url);
                        }
                    }

                }

                startPosition += 2;
                if (htmlDoc.DocumentNode.SelectSingleNode("//tr[" + startPosition + "]/td[3]") == null)
                    break;
                symbol = htmlDoc.DocumentNode.SelectSingleNode("//tr[" + startPosition + "]/td[3]").InnerText;

            }

        }

        private void GrabDataFromInnerHtm(string url)
        {
            count++;
            HtmlAgilityPack.HtmlDocument htmlDoc = WebClientUtil.GetHtmlDocument(url, 18000);
            string txtFile = htmlDoc.DocumentNode.SelectSingleNode("//pre").InnerText;
            DataAnalysis(txtFile);
        }

        public string GrabberHTMLFromNewsToday(string url)
        {
            string htm = "";
            try
            {
                WebRequest wr = WebRequest.Create(url);
                WebResponse rs = wr.GetResponse();
                StreamReader sr = new StreamReader(rs.GetResponseStream());
                htm = sr.ReadToEnd();
            }
            catch (Exception e)
            {
                logger.LogErrorAndRaiseException("Grab HTML error!", e);
            }
            return htm;
        }

        private HtmlAgilityPack.HtmlDocument ParseNewsTodayHTML(string htm)
        {

            HtmlAgilityPack.HtmlDocument htmlDoc = null;

            DateTime today = DateTime.Now;
            string formatToday = today.ToString("dd", new CultureInfo("en-US")) + " " + today.ToString("MMM", new CultureInfo("en-US")) + " " + today.ToString("yyyy");
            try
            {
                int formatTodayPos = htm.LastIndexOf(formatToday);

                string afterHtm = htm.Substring(formatTodayPos, htm.Length - formatTodayPos);

                int firstTableInAfterHtmPos = afterHtm.IndexOf("<table");
                int remarkPos = afterHtm.IndexOf("Remark");
                string beforeRemarkHtm = afterHtm.Substring(0, remarkPos);
                int tabaleBeforeRemarkPos = beforeRemarkHtm.IndexOf("</table>");

                string tableHtm = afterHtm.Substring(firstTableInAfterHtmPos, tabaleBeforeRemarkPos - firstTableInAfterHtmPos + 8);

                string simpleHtm = "<html><head></head><body>" + tableHtm + "</body></html>";

                htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(simpleHtm);
            }
            catch (Exception e)
            {
                logger.LogErrorAndRaiseException("Parse the html of News Today failed", e);
            }
            return htmlDoc;
        }



        /**
         * SetupTest used to create a Selenium object and start it
         * 
         */
        private void SetupTest()
        {
            string filename = "java";
            string parameters = @" -jar selenium-server.jar";
            System.Diagnostics.Process[] allProcess = System.Diagnostics.Process.GetProcessesByName("java");
            try
            {

                if (allProcess.Length == 0)
                {
                    System.Diagnostics.Process.Start(filename, parameters);
                    System.Threading.Thread.Sleep(3000);

                    selenium = new DefaultSelenium("localhost", 4444, "*iexplore", "http://www.set.or.th");
                    selenium.Start();
                    selenium.SetTimeout("100000");
                    selenium.UseXpathLibrary("javascript-xpath");
                }
            }
            catch (System.Exception e)
            {
                logger.LogErrorAndRaiseException("Selenium console not started." + e.ToString());
                System.Windows.Forms.Application.Exit();
            }
        }

        /**
         * ClearTest used to stop selenium running
         * 
         */
        private void ClearTest()
        {
            try
            {
                selenium.Stop();
                System.Diagnostics.Process[] allProcess = System.Diagnostics.Process.GetProcessesByName("java");
                if (allProcess.Length != 0)
                {
                    for (int i = 0; i < allProcess.Length; i++)
                    {
                        allProcess[i].Kill();
                    }
                }
            }
            catch (SeleniumException e)
            {
                logger.LogErrorAndRaiseException("Error found in ClearTest:" + e.ToString());
            }
        }


        /**
         * Quick method to get data from a dedicated URL
         * 
         */
        //private void DataCaptureByUrl()
        //{
        //    try
        //    {
        //        selenium.Open(configObj.html_url);

        //        string txt = selenium.GetText("//pre");
        //        DataAnalysis(txt);
        //        count++;
        //    }
        //    catch (SeleniumException ex)
        //    {
        //        logger.LogErrorAndRaiseException(ex.ToString());
        //    }

        //}

        /**
         * DataAnalysis method used to split pre txt to get required data
         * Return   : void
         * Parameter: String sourceStr
         */
        private void DataAnalysis(String sourceStr)
        {

            TailandRicTemplate objAQSTQS = new TailandRicTemplate();
            Hashtable ht = new Hashtable();

            //create a backup string to keep source string
            String bksourceStr = sourceStr;

            for (int i = 0; i < configObj.TailandRicFields.Count; i++)
            {
                sourceStr = sourceStr.Substring(sourceStr.IndexOf(configObj.TailandRicFields[i]));


                if (configObj.TailandRicFields[i] == "Exercise ratio (DW : Underlying asset)")
                {
                    sourceStr = sourceStr.Substring(sourceStr.IndexOf(")") + 6);
                }
                else if (configObj.TailandRicFields[i] == "Number of derivative warrants (unit:")
                {
                    sourceStr = sourceStr.Substring(sourceStr.IndexOf(":") + 8);
                }
                else
                {
                    sourceStr = sourceStr.Substring(sourceStr.IndexOf(":") + 2);
                }

                ht.Add(configObj.TailandRicFields[i], sourceStr.Substring(0, sourceStr.IndexOf("\n")));
            }

            //Get short name of issuer, need dealing if no short name contained in ()
            String tempStr = bksourceStr.Substring(bksourceStr.IndexOf(configObj.TailandRicFields[0]), bksourceStr.IndexOf(configObj.TailandRicFields[1]) - bksourceStr.IndexOf(configObj.TailandRicFields[0]));

            if (tempStr.Contains("(THAILAND)"))
            {
                tempStr = tempStr.Substring(tempStr.IndexOf("(THAILAND)") + "(THAILAND)".Length);

            }
            objAQSTQS.dwIssueShortNameStr = tempStr.Substring(tempStr.IndexOf("(") + 1, tempStr.IndexOf(")") - tempStr.IndexOf("(") - 1);





            objAQSTQS.dwIssuerStr = DWIssuerFormat((string)ht["DW issuer"], objAQSTQS.dwIssueShortNameStr);

            objAQSTQS.totalRicsStr = "1";
            objAQSTQS.exchangeStr = "SET";

            objAQSTQS.effectiveDateStr = (string)ht["Trading date"];

            objAQSTQS.officialCodeStr = (string)ht["DW name"];
            objAQSTQS.ricStr = objAQSTQS.officialCodeStr + "tc.BK";

            objAQSTQS.displayNameStr = (string)ht["Underlying asset(s)"] + " BY " + objAQSTQS.dwIssueShortNameStr + " CA#" + objAQSTQS.officialCodeStr.Substring(objAQSTQS.officialCodeStr.Length - 1, 1);

            //Get ISIN number
            objAQSTQS.offc_code2Str = DataCaptureFromSetSmart(objAQSTQS.officialCodeStr);


            objAQSTQS.currencyStr = "THB";
            objAQSTQS.recordtypeStr = "97";
            objAQSTQS.chainRicStr = "0#IPO.BK";
            objAQSTQS.chainRicOnListingDateStr = "0#DW.BK, 0#" + ((string)ht["Underlying asset(s)"]).Substring(0, 1) + ".BK";
            objAQSTQS.bcastRefStr = objAQSTQS.dwIssueShortNameStr + ".BK";
            objAQSTQS.underlyingRicStr = (string)ht["Underlying asset(s)"] + ".BK";
            objAQSTQS.longlink2Str = (string)ht["DW name"] + "tcta.BK";
            objAQSTQS.wntRatioStr = (string)ht["Exercise ratio (DW : Underlying asset)"];
            objAQSTQS.strikePrcStr = (string)ht["Exercise price (baht per share)"];
            objAQSTQS.lastTradingDateStr = (string)ht["Last trading date"];

            objAQSTQS.exerciseDateStr = (string)ht["Last exercise date"];
            DateTime mDT = new DateTime();
            if (DateTime.TryParse(objAQSTQS.exerciseDateStr, out mDT))
            {
                objAQSTQS.maturDateDT = mDT.AddDays(1);
            }
            else
            {
                logger.LogErrorAndRaiseException("Fail to parse maturity date to date");
                objAQSTQS.maturDateDT = DateTime.Now;
            }

            objAQSTQS.numberOfWarrants = (string)ht["Number of derivative warrants (unit:"];

            GenerateFMTemplate(objAQSTQS);
            GenerateIDNBulkFile(objAQSTQS);
        }

        /**
         * Generate IDN Bulk file
         * Return   : void
         * Parameter: TailandRicTemplate objAQSTQS
         */
        private void GenerateIDNBulkFile(TailandRicTemplate objAQSTQS)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                logger.LogErrorAndRaiseException("EXCEL could not be started. Check that your office installation and project references are correct");
                return;
            }
            xlApp.Visible = false;


            try
            {
                //create a new excel file
                if (!Directory.Exists(configObj.SavePath))
                {
                    DirectoryInfo dir = new DirectoryInfo(configObj.SavePath);
                    dir.Create();
                }

                Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                wSheet.Name = "SET_EQLB_W";

                ((Range)wSheet.Columns["A:Q", System.Type.Missing]).ColumnWidth = 18;
                xlApp.Cells.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;


                if (wSheet == null)
                {
                    logger.LogErrorAndRaiseException("Worksheet could not be created. Check that your office installation and project references are correct.");
                }

                ((Range)wSheet.Rows["1", System.Type.Missing]).Font.Bold = true;
                ((Range)wSheet.Rows["1", System.Type.Missing]).Font.Color = ColorTranslator.ToOle(Color.Red);
                ((Range)wSheet.Rows["1", System.Type.Missing]).Interior.Color = ColorTranslator.ToOle(Color.Yellow);


                wSheet.Cells[1, 1] = "SYMBOL";
                wSheet.Cells[2, 1] = objAQSTQS.ricStr;

                wSheet.Cells[1, 2] = "DSPLY_NAME";
                wSheet.Cells[2, 2] = objAQSTQS.displayNameStr;

                wSheet.Cells[1, 3] = "RIC";
                wSheet.Cells[2, 3] = objAQSTQS.ricStr;

                wSheet.Cells[1, 4] = "OFFCL_CODE";
                wSheet.Cells[2, 4] = objAQSTQS.officialCodeStr;


                wSheet.Cells[1, 5] = "EX_SYMBOL";
                wSheet.Cells[2, 5] = "DUMMY-" + objAQSTQS.ricStr.Substring(0, objAQSTQS.ricStr.IndexOf(".")).ToUpper();

                wSheet.Cells[1, 6] = "BCKGRNDPAG";
                wSheet.Cells[2, 6] = "****";

                wSheet.Cells[1, 7] = "BCAST_REF";
                wSheet.Cells[2, 7] = objAQSTQS.bcastRefStr;

                ((Range)wSheet.Cells[2, 8]).NumberFormat = "@";
                wSheet.Cells[1, 8] = "#INSTMOD_EXPIR_DATE";
                wSheet.Cells[2, 8] = objAQSTQS.maturDateDT.ToString("dd/MM/yyyy");

                wSheet.Cells[1, 9] = "#INSTMOD_LONGLINK1";
                wSheet.Cells[2, 9] = objAQSTQS.underlyingRicStr;

                wSheet.Cells[1, 10] = "#INSTMOD_LONGLINK2";
                wSheet.Cells[2, 10] = objAQSTQS.longlink2Str;

                ((Range)wSheet.Cells[2, 11]).NumberFormat = "@";
                wSheet.Cells[1, 11] = "#INSTMOD_MATUR_DATE";
                wSheet.Cells[2, 11] = objAQSTQS.maturDateDT.ToString("dd/MM/yyyy");

                wSheet.Cells[1, 12] = "#INSTMOD_OFFC_CODE2";
                wSheet.Cells[2, 12] = objAQSTQS.offc_code2Str;

                wSheet.Cells[1, 13] = "#INSTMOD_STRIKE_PRC";
                wSheet.Cells[2, 13] = objAQSTQS.strikePrcStr;

                ((Range)wSheet.Cells[2, 14]).NumberFormat = "@";
                wSheet.Cells[1, 14] = "#INSTMOD_WNT_RATIO";
                wSheet.Cells[2, 14] = objAQSTQS.wntRatioStr.Substring(0, objAQSTQS.wntRatioStr.IndexOf(":") - 1);

                wSheet.Cells[1, 15] = "#INSTMOD_MNEMONIC";
                wSheet.Cells[2, 15] = objAQSTQS.officialCodeStr;

                wSheet.Cells[1, 16] = "#INSTMOD_TDN_SYMBOL";
                wSheet.Cells[2, 16] = objAQSTQS.officialCodeStr;

                wSheet.Cells[1, 17] = "EXL_NAME";
                wSheet.Cells[2, 17] = "SET_EQLB_W";



                xlApp.DisplayAlerts = false;
                xlApp.AlertBeforeOverwriting = false;

                string fullpath = configObj.SavePath + "\\" + objAQSTQS.officialCodeStr + "_" + DateTime.Now.ToString("d").Replace('/', '-') + ".xls";
                wBook.SaveCopyAs(fullpath);

            }
            catch (SystemException ex)
            {
                logger.LogErrorAndRaiseException(ex.ToString());
            }
            finally
            {
                xlApp.Quit();
                KillExcelProcess(xlApp);
            }
        }

        /**
         * Generate FM template file
         * Return   : void
         * Parameter: TailandRicTemplate objAQSTQS
         */
        private void GenerateFMTemplate(TailandRicTemplate objAQSTQS)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                logger.LogErrorAndRaiseException("EXCEL could not be started. Check that your office installation and project references are correct");
                return;
            }
            xlApp.Visible = false;


            try
            {
                //create a new excel file
                if (!Directory.Exists(configObj.SavePath))
                {
                    DirectoryInfo dir = new DirectoryInfo(configObj.SavePath);
                    dir.Create();
                }

                Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                wSheet.Name = "TailandRicInfo";

                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 30;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 4;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 50;
                xlApp.Cells.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;


                if (wSheet == null)
                {
                    logger.LogErrorAndRaiseException("Worksheet could not be created. Check that your office installation and project references are correct.");
                }

                //Generate data for AQS/TQS template

                wSheet.Cells[1, 1] = "Official Code";
                wSheet.Cells[1, 2] = ":";
                wSheet.Cells[1, 3] = objAQSTQS.officialCodeStr;

                ((Range)wSheet.Cells[3, 1]).NumberFormat = "@";
                wSheet.Cells[3, 1] = "========================================================";

                wSheet.Cells[4, 1] = "Total RICs";
                wSheet.Cells[4, 2] = ":";
                wSheet.Cells[4, 3] = objAQSTQS.totalRicsStr;

                wSheet.Cells[5, 1] = "Exchange";
                wSheet.Cells[5, 2] = ":";
                wSheet.Cells[5, 3] = objAQSTQS.exchangeStr;


                ((Range)wSheet.Cells[7, 1]).Font.Bold = true;
                ((Range)wSheet.Cells[8, 1]).Font.Bold = true;
                wSheet.Cells[7, 1] = "1";
                wSheet.Cells[8, 1] = "For AQS/TQS";
                wSheet.Cells[9, 1] = "--------------------";

                ((Range)wSheet.Cells[10, 3]).NumberFormat = "@";
                wSheet.Cells[10, 1] = "Effective Date";
                wSheet.Cells[10, 2] = ":";
                wSheet.Cells[10, 3] = objAQSTQS.effectiveDateStr;


                wSheet.Cells[11, 1] = "RIC";
                wSheet.Cells[11, 2] = ":";
                wSheet.Cells[11, 3] = objAQSTQS.ricStr;

                wSheet.Cells[12, 1] = "Displayname";
                wSheet.Cells[12, 2] = ":";
                wSheet.Cells[12, 3] = objAQSTQS.displayNameStr;

                wSheet.Cells[13, 1] = "Official Code";
                wSheet.Cells[13, 2] = ":";
                wSheet.Cells[13, 3] = objAQSTQS.officialCodeStr;

                ((Range)wSheet.Cells[14, 1]).Font.Color = ColorTranslator.ToOle(Color.Red);
                ((Range)wSheet.Cells[14, 2]).Font.Color = ColorTranslator.ToOle(Color.Red);
                ((Range)wSheet.Cells[14, 1]).Font.Bold = true;
                wSheet.Cells[14, 1] = "Exchange Symbol";
                wSheet.Cells[14, 2] = ":";

                wSheet.Cells[15, 1] = "OFFC_CODE2";
                wSheet.Cells[15, 2] = ":";
                wSheet.Cells[15, 3] = objAQSTQS.offc_code2Str;

                wSheet.Cells[16, 1] = "Currency";
                wSheet.Cells[16, 2] = ":";
                wSheet.Cells[16, 3] = objAQSTQS.currencyStr;

                wSheet.Cells[17, 1] = "Recordtype";
                wSheet.Cells[17, 2] = ":";
                wSheet.Cells[17, 3] = objAQSTQS.recordtypeStr;


                wSheet.Cells[18, 1] = "Chain RIC (Old)";
                wSheet.Cells[18, 2] = ":";
                wSheet.Cells[18, 3] = objAQSTQS.chainRicStr;

                wSheet.Cells[19, 1] = "Chain RIC (New)";
                wSheet.Cells[19, 2] = ":";
                wSheet.Cells[19, 3] = objAQSTQS.chainRicOnListingDateStr;

                wSheet.Cells[21, 1] = "BCAST-REF";
                wSheet.Cells[21, 2] = ":";
                wSheet.Cells[21, 3] = objAQSTQS.bcastRefStr;

                wSheet.Cells[22, 1] = "UNDERLYING RIC (LONGLINK1)";
                wSheet.Cells[22, 2] = ":";
                wSheet.Cells[22, 3] = objAQSTQS.underlyingRicStr;

                wSheet.Cells[23, 1] = "LONGLINK2";
                wSheet.Cells[23, 2] = ":";
                wSheet.Cells[23, 3] = objAQSTQS.longlink2Str;

                ((Range)wSheet.Cells[24, 3]).NumberFormat = "@";
                wSheet.Cells[24, 1] = "WNT_RATIO (DW per Underlying)";
                wSheet.Cells[24, 2] = ":";
                wSheet.Cells[24, 3] = objAQSTQS.wntRatioStr.Substring(0, objAQSTQS.wntRatioStr.IndexOf(":") - 1);

                ((Range)wSheet.Cells[25, 3]).NumberFormat = "0.00";
                wSheet.Cells[25, 1] = "STRIKE_PRC (WT)";
                wSheet.Cells[25, 2] = ":";
                wSheet.Cells[25, 3] = objAQSTQS.strikePrcStr;

                ((Range)wSheet.Cells[26, 3]).NumberFormat = "@";
                wSheet.Cells[26, 1] = "MATUR_DATE / EXPIR_DATE";
                wSheet.Cells[26, 2] = ":";
                wSheet.Cells[26, 3] = objAQSTQS.maturDateDT.Day + "-" + MonthTrans(objAQSTQS.maturDateDT) + "-" + objAQSTQS.maturDateDT.Year;


                //Generate data for NDA Template
                ((Range)wSheet.Cells[28, 1]).Font.Bold = true;
                wSheet.Cells[28, 1] = "For NDA";
                wSheet.Cells[29, 1] = "--------------------";

                ((Range)wSheet.Cells[30, 1]).Font.Bold = true;
                wSheet.Cells[30, 1] = "B. Existing Organisation Listing";

                wSheet.Cells[31, 1] = "RIC";
                wSheet.Cells[31, 2] = ":";
                wSheet.Cells[31, 3] = objAQSTQS.ricStr;

                wSheet.Cells[32, 1] = "IDN Longname";
                wSheet.Cells[32, 2] = ":";
                wSheet.Cells[32, 3] = objAQSTQS.underlyingRicStr.Substring(0, objAQSTQS.underlyingRicStr.IndexOf(".BK")) + "_" + objAQSTQS.dwIssueShortNameStr + "@" + IDNLongNameFormat(objAQSTQS.strikePrcStr) + MonthTrans(objAQSTQS.maturDateDT).ToUpper() + (objAQSTQS.maturDateDT.Year).ToString().Substring(2, 2) + "CWNT";

                wSheet.Cells[33, 1] = "Issue Classification";
                wSheet.Cells[33, 2] = ":";
                wSheet.Cells[33, 3] = "WT";

                wSheet.Cells[34, 1] = "Primary Listing (RIC) /EDCOID";
                wSheet.Cells[34, 2] = ":";
                wSheet.Cells[34, 3] = objAQSTQS.bcastRefStr;

                wSheet.Cells[35, 1] = "Organisation Name (DIRNAME)";
                wSheet.Cells[35, 2] = ":";
                wSheet.Cells[35, 3] = objAQSTQS.dwIssuerStr;

                ((Range)wSheet.Cells[37, 1]).Font.Bold = true;
                wSheet.Cells[37, 1] = "Note (For Warrant Term & Conditions)";

                wSheet.Cells[38, 1] = "RIC";
                wSheet.Cells[38, 2] = ":";
                wSheet.Cells[38, 3] = objAQSTQS.ricStr;

                ((Range)wSheet.Cells[39, 3]).NumberFormat = "@";
                wSheet.Cells[39, 1] = "Issue Date";
                wSheet.Cells[39, 2] = ":";
                wSheet.Cells[39, 3] = objAQSTQS.effectiveDateStr;

                ((Range)wSheet.Cells[40, 3]).NumberFormat = "@";
                wSheet.Cells[40, 1] = "Listed Date";
                wSheet.Cells[40, 2] = ":";
                wSheet.Cells[40, 3] = objAQSTQS.effectiveDateStr;

                ((Range)wSheet.Cells[41, 3]).NumberFormat = "@";
                wSheet.Cells[41, 1] = "First Exercise Date";
                wSheet.Cells[41, 2] = ":";
                wSheet.Cells[41, 3] = objAQSTQS.exerciseDateStr + "(Automatic Exercise)";

                ((Range)wSheet.Cells[42, 3]).NumberFormat = "@";
                wSheet.Cells[42, 1] = "Last Exercise Date";
                wSheet.Cells[42, 2] = ":";
                wSheet.Cells[42, 3] = objAQSTQS.exerciseDateStr + "(Automatic Exercise)";

                ((Range)wSheet.Cells[43, 3]).NumberFormat = "@";
                wSheet.Cells[43, 1] = "Last Actual Trading Date";
                wSheet.Cells[43, 2] = ":";
                wSheet.Cells[43, 3] = objAQSTQS.lastTradingDateStr;

                ((Range)wSheet.Cells[44, 3]).NumberFormat = "@";
                wSheet.Cells[44, 1] = "Expiry Date";
                wSheet.Cells[44, 2] = ":";
                wSheet.Cells[44, 3] = objAQSTQS.maturDateDT.Day + "-" + MonthTrans(objAQSTQS.maturDateDT) + "-" + objAQSTQS.maturDateDT.Year;

                ((Range)wSheet.Cells[45, 3]).NumberFormat = "0.00";
                wSheet.Cells[45, 1] = "Strike Price";
                wSheet.Cells[45, 2] = ":";
                wSheet.Cells[45, 3] = objAQSTQS.strikePrcStr;

                ((Range)wSheet.Cells[46, 3]).NumberFormat = "@";
                wSheet.Cells[46, 1] = "Ratio warrant : underlying";
                wSheet.Cells[46, 2] = ":";
                wSheet.Cells[46, 3] = objAQSTQS.wntRatioStr + " ( " + objAQSTQS.wntRatioStr.Substring(0, objAQSTQS.wntRatioStr.IndexOf(':') - 1) + " Derivative Warrants " + " : " + objAQSTQS.wntRatioStr.Substring(objAQSTQS.wntRatioStr.IndexOf(':') + 2) + " Stock )";

                ((Range)wSheet.Cells[47, 3]).NumberFormat = "@";
                wSheet.Cells[47, 1] = "Outstanding Warrant Quantity";
                wSheet.Cells[47, 2] = ":";
                wSheet.Cells[47, 3] = objAQSTQS.numberOfWarrants;

                wSheet.Cells[48, 1] = "ISIN";
                wSheet.Cells[48, 2] = ":";
                wSheet.Cells[48, 3] = objAQSTQS.offc_code2Str;

                //((Range)wSheet.Rows["49", System.Type.Missing]).Font.Color = ColorTranslator.ToOle(Color.Blue);
                wSheet.Cells[49, 1] = "Exercise Period";
                wSheet.Cells[49, 2] = ":";
                wSheet.Cells[49, 3] = "European Style; DW can be exercised only on Automatic Exercise Date.";

                ((Range)wSheet.Cells[51, 1]).Font.Bold = true;
                ((Range)wSheet.Cells[51, 3]).Font.Bold = true;
                wSheet.Cells[51, 1] = "RICs(ADD) in the Search";
                wSheet.Cells[51, 2] = ":";
                wSheet.Cells[51, 3] = objAQSTQS.ricStr;

                ((Range)wSheet.Cells[53, 1]).NumberFormat = "@";
                wSheet.Cells[53, 1] = "===========================END=============================";

                xlApp.DisplayAlerts = false;
                xlApp.AlertBeforeOverwriting = false;

                string fullpath = configObj.SavePath + "\\" + "FM Temp-" + objAQSTQS.officialCodeStr + "_" + DateTime.Now.ToString("d").Replace('/', '-') + ".xls";
                wBook.SaveCopyAs(fullpath);


            }
            catch (SystemException ex)
            {
                logger.LogErrorAndRaiseException(ex.ToString());
            }
            finally
            {
                xlApp.Quit();
                KillExcelProcess(xlApp);
            }
        }

        //public static string GetDynamicPageSource(string uri, int timeout, string postData, string cookie)
        //{
        //    string pageSource = string.Empty;
        //    Encoding encoding = Encoding.UTF8;

        //    byte[] buf = null;
        //    int retryiesLeft = 5;
        //    buf = encoding.GetBytes(postData);

        //    while (string.IsNullOrEmpty(pageSource) && retryiesLeft-- > 0)
        //    {
        //        HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
        //        request.Timeout = timeout;
        //        request.Headers["Cookie"] = cookie;
        //        request.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:6.0.2) Gecko/20100101 Firefox/6.0.2";
        //        request.Method = "POST";
        //        request.ContentType = "application/x-www-form-urlencoded";
        //        request.ContentLength = buf.Length;
        //        request.GetRequestStream().Write(buf, 0, buf.Length);
        //        using (WebResponse response = request.GetResponse())
        //        {
        //            StreamReader sr = new StreamReader(response.GetResponseStream());
        //            pageSource = sr.ReadToEnd();
        //        }
        //    }
        //    if (pageSource == string.Empty)
        //    {
        //        throw new Exception(string.Format("Cannot download page {0}b with post data {1}", uri, postData));
        //    }
        //    return pageSource;
        //}

        //public static HttpWebResponse GetResponse(string uri, int timeout, string postData)
        //{
        //    string pageSource = string.Empty;
        //    Encoding encoding = Encoding.UTF8;

        //    byte[] buf = null;
        //    buf = encoding.GetBytes(postData);


        //    HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
        //    request.Timeout = timeout;
        //    //request.Headers["Cookie"] = "80922B3FB48D76DD2F6A268906563EC5.simpn9; Path=/ism";
        //    request.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:6.0.2) Gecko/20100101 Firefox/6.0.2";
        //    request.Method = "POST";
        //    request.ContentType = "application/x-www-form-urlencoded";
        //    request.ContentLength = buf.Length;
        //    request.GetRequestStream().Write(buf, 0, buf.Length);
        //    return request.GetResponse() as HttpWebResponse;
        //}

        //private void DataCaptureFromSetSmart(string symbol)
        //{
        //    string loginUrl = "http://www.setsmart.com/ism/login.jsp";
        //    string loginPostData = "txtLogin=25932633&txtPassword=reuters1&hidBrowser=null&hidLang=English";
        //    string searchUrl = "http://www.setsmart.com/ism/companyprofile.html";
        //    string searchData = "symbol=" + symbol;
        //    var req = GetResponse(loginUrl, 10000, loginPostData);
        //    var cookie = req.Headers["Set-Cookie"];
        //    cookie = cookie.Substring(0, cookie.IndexOf(";")).Trim();
        //    var content = GetDynamicPageSource(searchUrl, 10000, searchData, cookie);

        //}

        /**
         * Grab Data from http://isin.krx.co.kr/
         *  
         */
        //private void DataCaptureFromSET()
        //{
        //    String symbolType = "SET";
        //    int startPosition = configObj.start_position;
        //    int endPosition = configObj.end_position;

        //    try
        //    {

        //        selenium.SetTimeout("50000");
        //        selenium.OpenWindow("/set/todaynews.do?language=en&country=US", "parentWindow");
        //        selenium.WaitForPopUp("parentWindow", "50000");
        //        selenium.SelectWindow("parentWindow");

        //        //No defination for searching end of position
        //        if (configObj.end_position == 0)
        //        {
        //            while (selenium.IsElementPresent("//tr[2]/td/table/tbody/tr[" + startPosition + "]/td[3]"))
        //            {
        //                String symbolStr = selenium.GetText("//tr[2]/td/table/tbody/tr[" + startPosition + "]/td[3]");
        //                if (symbolStr == symbolType)
        //                {
        //                    String headlineStr = selenium.GetText("//tr[2]/td/table/tbody/tr[" + startPosition + "]/td[4]");
        //                    if (headlineStr.Contains("adds new") && (headlineStr.Contains("CC") || headlineStr.Contains("CB") || headlineStr.Contains("CA") || headlineStr.Contains("CD") || headlineStr.Contains("CE") || headlineStr.Contains("CF")))
        //                    {

        //                        String htmlUrl = selenium.GetAttribute("//tr[2]/td/table/tbody/tr[" + startPosition + "]/td[5]/a[text()='HTML']@href");
        //                        selenium.OpenWindow(htmlUrl, "detailWindow" + startPosition);
        //                        selenium.WaitForPopUp("detailWindow" + startPosition, "10000");
        //                        selenium.SelectWindow("detailWindow" + startPosition);

        //                        String txtFile = selenium.GetText("//pre");

        //                        //here call txt analysis method
        //                        DataAnalysis(txtFile);

        //                        count++;
        //                        selenium.SelectWindow("parentWindow");

        //                    }
        //                }

        //                startPosition = startPosition + 2;
        //            }//end while

        //        }
        //        else if (configObj.end_position >= configObj.start_position)
        //        {
        //            for (int i = startPosition; i <= endPosition; i = i + 2)
        //            {
        //                String symbolStr = selenium.GetText("//tr[2]/td/table/tbody/tr[" + startPosition + "]/td[3]");
        //                if (symbolStr == symbolType)
        //                {
        //                    String headlineStr = selenium.GetText("//tr[2]/td/table/tbody/tr[" + startPosition + "]/td[4]");
        //                    if (headlineStr.Contains("adds new") && (headlineStr.Contains("CC") || headlineStr.Contains("CB") || headlineStr.Contains("CA") || headlineStr.Contains("CD") || headlineStr.Contains("CE") || headlineStr.Contains("CF")))
        //                    {

        //                        String htmlUrl = selenium.GetAttribute("//tr[2]/td/table/tbody/tr[" + startPosition + "]/td[5]/a[text()='HTML']@href");
        //                        selenium.OpenWindow(htmlUrl, "detailWindow" + startPosition);
        //                        selenium.WaitForPopUp("detailWindow" + startPosition, "10000");
        //                        selenium.SelectWindow("detailWindow" + startPosition);

        //                        String txtFile = selenium.GetText("//pre");

        //                        //here call txt analysis method
        //                        DataAnalysis(txtFile);

        //                        count++;
        //                        selenium.SelectWindow("parentWindow");

        //                    }
        //                }
        //            }//end for

        //        }//end else if
        //        else
        //        {
        //            logger.LogErrorAndRaiseException("end_position must bigger than start_position");
        //        }

        //    }//end try
        //    catch (SeleniumException e)
        //    {
        //        logger.LogErrorAndRaiseException(e.ToString());
        //        ClearTest();

        //    }

        //}//end DataCaptureFromSET()

        /**
         * Get ISIN number from http://isin.krx.co.kr by specific official code
         * 
         */
        private String DataCaptureFromSetSmart(String officialCode)
        {
            ISelenium setsmart = new DefaultSelenium("localhost", 4444, "*iexplore", "http://www.setsmart.com");
            try
            {

                setsmart.Start();
                setsmart.UseXpathLibrary("javascript-xpath");
                setsmart.Open("/ism/login.jsp");
                setsmart.WaitForPageToLoad("60000");
                setsmart.Type("txtLogin", configObj.username);
                setsmart.Type("txtPassword", configObj.password);
                setsmart.Click("//input[@type='image']");
                //setsmart.Click("//input[@type='image' and @onclick='return validate();']");
                setsmart.WaitForPageToLoad("60000");
                setsmart.Click("link=Company Profile");
                setsmart.WaitForPageToLoad("60000");
                setsmart.Type("symbol", officialCode);
                setsmart.Click("//input[@type='image']");
                //setsmart.Click("//input[@name='submit' and @value='go' and @type='image' and @onclick=' return validateSubmit(); ']");

                System.Threading.Thread.Sleep(3000);

                String isinNumber = setsmart.GetText("//table[3]/tbody/tr[26]/td[2]");

                setsmart.Click("link=LOGOUT");
                setsmart.Stop();

                //return isinNumber.Substring(isinNumber.IndexOf("Local") + 8, isinNumber.IndexOf("Foreign") - 9);
                return isinNumber.Substring(isinNumber.IndexOf("Local") + 7, isinNumber.Length - isinNumber.IndexOf("Local") - 7);

            }
            catch (SeleniumException e)
            {
                setsmart.Stop();
                logger.LogErrorAndRaiseException(e.ToString());
                return e.ToString();
            }
        }

        /**
         * Additional method, kill excel process
         * Return   :void
         * Parameter:Microsoft.Office.Interop.Excel.Application excelApp
         */
        private void KillExcelProcess(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            IntPtr t = new IntPtr(excelApp.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();
        }


        /**
        * Additional method, apply correct format for Organisation Name
        * Retrun: String
        * Parameter: String sourceStr
        */
        private String DWIssuerFormat(String sourceStr, String shortName)
        {

            bool isIncludeShortName = false;
            int index = sourceStr.IndexOf("PUBLIC COMPANY") - 1;
            if (index < 0)
            {
                index = sourceStr.IndexOf("PUBLIC") - 1;
            }

            String temp = sourceStr.Substring(0, index);
            temp = temp.ToLower();

            string[] strArray = temp.Split(" ".ToCharArray());
            string result = string.Empty;

            foreach (string s in strArray)
            {
                if (s.Substring(0, 1) == "(")
                {
                    isIncludeShortName = true;
                    result += "(" + s.Substring(1, 1).ToUpper() + s.Substring(2) + " ";
                }
                else
                {
                    result += s.Substring(0, 1).ToUpper() + s.Substring(1) + " ";
                }

            }
            if (isIncludeShortName)
            {
                return result + " PCL";
            }
            else
            {
                return result + "(" + shortName + ") PCL";
            }
        }

        /**
         * Additional method, apply correct format for IDN longname
         * Retrun: String
         * Parameter: String sourceStr
         */
        private String IDNLongNameFormat(String sourceStr)
        {
            if (sourceStr.Contains("."))
            {

                if (sourceStr.Substring(sourceStr.IndexOf('.') + 1, 1) == "0" && sourceStr.Substring(sourceStr.IndexOf('.') + 2, 1) == "0")
                {
                    return sourceStr.Substring(0, sourceStr.IndexOf('.'));
                }
                else if (sourceStr.Substring(sourceStr.IndexOf('.') + 2, 1) == "0")
                {
                    return sourceStr.Substring(0, sourceStr.IndexOf('.') + 2);
                }
                else
                {
                    return sourceStr;
                }
            }
            else
            {
                return sourceStr;
            }
        }

        /**
         * Additional method, change month from figure English expression
         * Retrun: String
         * Parameter: DateTime figureMonth
         */
        private String MonthTrans(DateTime figureMonth)
        {
            String month = "";
            month = figureMonth.ToString("MMM").ToString();
            //switch (figureMonth.Month)
            //{
            //    case 1:
            //        month = "Jan";
            //        break;
            //    case 2:
            //        month = "Feb";
            //        break;
            //    case 3:
            //        month = "Mar";
            //        break;
            //    case 4:
            //        month = "Apr";
            //        break;
            //    case 5:
            //        month = "May";
            //        break;
            //    case 6:
            //        month = "Jun";
            //        break;
            //    case 7:
            //        month = "Jul";
            //        break;
            //    case 8:
            //        month = "Aug";
            //        break;
            //    case 9:
            //        month = "Sep";
            //        break;
            //    case 10:
            //        month = "Oct";
            //        break;
            //    case 11:
            //        month = "Nov";
            //        break;
            //    case 12:
            //        month = "Dec";
            //        break;
            //}
            return month;
        }

    }//end of class
}
