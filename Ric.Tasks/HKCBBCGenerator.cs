using System;
using System.Collections;
using System.Collections.Generic;
using OpenQA.Selenium;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Globalization;
using System.Text;
using Ric.Util;

namespace Ric.Tasks
{
    public class HKCBBCGenerator
    {
        Ric.Core.Core coreObj = new Ric.Core.Core();
        private ISelenium selenium;
        int startPosition = 2;
        int pdf_disable = 0;
        private String fmSerialNumber = "0000";
        Hashtable issuerCodeHT = new Hashtable();
        List<HKRicTemplate> ricList = new List<HKRicTemplate>();
        List<HKRicTemplate> chineseList = new List<HKRicTemplate>();
        List<HKRicTemplate> pdfList = new List<HKRicTemplate>();

        const string CONFIGFILE_NAME = "Config\\HK\\HK_IssuerCode.xml";
        const string LOGFILE_NAME = "CBBCLog.txt";
        const string FOLDER_NAME = "CBBC";

        private int delayHour = 0;
        private string isKeepPDFFile = "Y";
        private int holidayCount = 0;
        private DateTime scheduleDate = DateTime.Now;

        public Ric.Core.Core CoreObj
        {
            get { return coreObj; }
        }
        public List<HKRicTemplate> RicList
        {
            get { return ricList; }
        }
        public List<HKRicTemplate> CHNList
        {
            get { return chineseList; }
        }

        public int DelayHour
        {
            set { delayHour = value; }
        }

        public String FMSerialNumber
        {
            set { fmSerialNumber = value; }
        }
        public DateTime ScheduleDate
        {
            set { scheduleDate = value; }
        }

        private string warrantIssuer = "";

        public string GetWarrantIssuer(string source)
        {
            if (!source.Equals(string.Empty))
            {
                warrantIssuer = ((IssuerCode)issuerCodeHT[source]).warrentIssuer.ToString();
            }
            return warrantIssuer;
        }

        public void StartGrabber()
        {

            //ReadConfigFile();
            //String code = "63917";
            //coreObj.txtAnalysis(code);
            System.Threading.Thread.Sleep(delayHour * 60 * 60 * 1000);

            DateTime start = DateTime.Now;
            SetupTest();

            //main method to capture data
            CBBCDataCapture();

            //Generate CBBC FM files
            int fileCount = (int)Math.Ceiling(ricList.Count / 20.0);
            for (int i = 0; i < fileCount; i++)
            {
                GenerateCBBCTemplate(i * 20);
                fmSerialNumber = coreObj.UpdateFMSerialNumber(fmSerialNumber);
            }



            ClearTest(isKeepPDFFile);

            DateTime end = DateTime.Now;
            TimeSpan runningtime = end.Subtract(start);

            coreObj.WriteLogFile("---------------------------------------------------------------------");
            coreObj.WriteLogFile("Generated Ric#:" + ricList.Count);
            coreObj.WriteLogFile("Execution Time:" + runningtime.ToString());

        }

        /**
         * Start selenium remote controller
         * 
         */
        private void StartRC()
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

                }
                selenium = new DefaultSelenium("localhost", 4444, "*iexplore", "http://www.hkex.com.hk");
                selenium.Start();
                selenium.UseXpathLibrary("javascript-xpath");
            }
            catch (System.Exception e)
            {
                MessageBox.Show("Selenium console not started." + e.ToString());
                System.Windows.Forms.Application.Exit();
            }

        }

        /**
         * Shutdown selenium remote controller
         * 
         */
        private void ShutDownRC()
        {
            try
            {

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
                MessageBox.Show("Error found in ClearTest:" + e.ToString());

            }
        }//end ClearTest

        /**
        * Grab Data from http://www.hkex.com.hk/eng/cbbc/newissue/newlaunch.htm
        *  
        */
        private void CBBCDataCapture()
        {

            int start_position = startPosition;

            try
            {
                selenium.OpenWindow("/eng/cbbc/newissue/newlaunch.htm", "CBBCListWindow");
                selenium.WaitForPopUp("CBBCListWindow", "3000");
                selenium.SelectWindow("CBBCListWindow");

                //Get English Content
                while (selenium.IsElementPresent("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[1]"))
                {
                    String launchDateStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[12]");
                    DateTime tempLaunchDateDT = DateTime.ParseExact(launchDateStr, "dd-MM-yyyy", null);
                    DateTime temp = coreObj.DateCalculate(scheduleDate, tempLaunchDateDT, holidayCount);

                    if (temp.Date == scheduleDate.Date)
                    {
                        HKRicTemplate hkRic = new HKRicTemplate();
                        hkRic.launchDateDT = DateTime.ParseExact(launchDateStr, "dd-MM-yyyy", null);
                        hkRic.ricCodeStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[2]");
                        hkRic.ricNameStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[3]");
                        hkRic.issuerIDStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[4]");
                        hkRic.underlyingStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[5]");
                        hkRic.bullBearStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[6]");
                        hkRic.boardLotStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[7]");
                        hkRic.strikeLevelStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[8]");
                        hkRic.callLevelStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[9]");
                        hkRic.entitlementRatioStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[10]");
                        hkRic.issueSizeStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[11]");
                        hkRic.clearingCommDateDT = DateTime.ParseExact(selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[13]"), "dd-MM-yyyy", null);
                        hkRic.listingDateDT = DateTime.ParseExact(selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[14]"), "dd-MM-yyyy", null);
                        hkRic.maturityDateDT = DateTime.ParseExact(selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[15]"), "dd-MM-yyyy", null);


                        //Get issue price
                        String url = selenium.GetAttribute("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[2]/a@href");
                        String id = url.Substring(url.IndexOf('=') + 1, 5);
                        String summaryUrl = "http://www.hkex.com.hk/eng/cbbc/cbbcsummary.asp?id=" + id;
                        selenium.OpenWindow(summaryUrl, "CBBCSummaryWindow" + start_position);
                        selenium.WaitForPopUp("CBBCSummaryWindow" + start_position, "30000");
                        selenium.SelectWindow("CBBCSummaryWindow" + start_position);


                        //For Equity Get underlying name
                        if (Char.IsDigit(hkRic.underlyingStr, 0))
                        {
                            hkRic.underlyingNameForStockStr = selenium.GetText("//table[@id='bodytable']/tbody/tr[3]/td[2]").Substring(8);
                        }

                        hkRic.issuePriceStr = selenium.GetText("//table[@id='bodytable']/tbody/tr[12]/td[2]").Substring(4);

                        selenium.Close();

                        ricList.Add(hkRic);


                        selenium.SelectWindow("CBBCListWindow");

                    }//end if

                    start_position++;

                }//end while


                ChineseNameCapture(start_position);

                //Get gearing and premium
                //search on page http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.asp
                DataCaptureFromPDF(0);

            }//end try
            catch (SeleniumException ex)
            {
                String errLog = ex.ToString();
                selenium.Close();
                selenium.Stop();
                CaptureRetry(start_position);

            }

        }//end CBBCDataCapture()

        private void CaptureRetry(int start_position)
        {

            selenium.Start();
            selenium.UseXpathLibrary("javascript-xpath");

            try
            {
                selenium.OpenWindow("/eng/cbbc/newissue/newlaunch.htm", "CBBCListWindow");
                selenium.WaitForPopUp("CBBCListWindow", "30000");
                selenium.SelectWindow("CBBCListWindow");

                //Get English Content
                while (selenium.IsElementPresent("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[1]"))
                {

                    String launchDateStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[12]");
                    DateTime tempLaunchDateDT = DateTime.ParseExact(launchDateStr, "dd-MM-yyyy", null);
                    DateTime temp = coreObj.DateCalculate(scheduleDate, tempLaunchDateDT, holidayCount);

                    if (temp.Date == scheduleDate.Date)
                    {
                        HKRicTemplate hkRic = new HKRicTemplate();
                        hkRic.launchDateDT = DateTime.ParseExact(launchDateStr, "dd-MM-yyyy", null);
                        hkRic.ricCodeStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[2]");
                        hkRic.ricNameStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[3]");
                        hkRic.issuerIDStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[4]");
                        hkRic.underlyingStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[5]");
                        hkRic.bullBearStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[6]");
                        hkRic.boardLotStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[7]");
                        hkRic.strikeLevelStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[8]");
                        hkRic.callLevelStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[9]");
                        hkRic.entitlementRatioStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[10]");
                        hkRic.issueSizeStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[11]");
                        hkRic.clearingCommDateDT = DateTime.ParseExact(selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[13]"), "dd-MM-yyyy", null);
                        hkRic.listingDateDT = DateTime.ParseExact(selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[14]"), "dd-MM-yyyy", null);
                        hkRic.maturityDateDT = DateTime.ParseExact(selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[15]"), "dd-MM-yyyy", null);

                        //Get issue price
                        String url = selenium.GetAttribute("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[2]/a@href");
                        String id = url.Substring(url.IndexOf('=') + 1, 5);
                        String summaryUrl = "http://www.hkex.com.hk/eng/cbbc/cbbcsummary.asp?id=" + id;
                        selenium.OpenWindow(summaryUrl, "CBBCSummaryWindow" + start_position);
                        selenium.WaitForPopUp("CBBCSummaryWindow" + start_position, "30000");
                        selenium.SelectWindow("CBBCSummaryWindow" + start_position);

                        //For Equity Get underlying name
                        if (Char.IsDigit(hkRic.underlyingStr, 0))
                        {
                            hkRic.underlyingNameForStockStr = selenium.GetText("//table[@id='bodytable']/tbody/tr[3]/td[2]").Substring(8);
                        }

                        hkRic.issuePriceStr = selenium.GetText("//table[@id='bodytable']/tbody/tr[12]/td[2]").Substring(4);

                        selenium.Close();

                        ricList.Add(hkRic);
                        selenium.SelectWindow("CBBCListWindow");

                    }//end if

                    start_position++;

                }//end while

                ChineseNameCapture(start_position);

                //Get gearing and premium
                //search on page http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.asp
                DataCaptureFromPDF(0);

            }//end try
            catch (SeleniumException ex)
            {
                String errLog = ex.ToString();
                selenium.Close();
                selenium.Stop();
                CaptureRetry(start_position);
            }
        }

        private void ChineseNameCapture(int lastRickIndex)
        {
            //Get Chinese Content
            HKRicTemplate chineseRic = new HKRicTemplate();
            int start_position = lastRickIndex - ricList.Count;

            selenium.Click("link=繁體");
            selenium.WaitForPageToLoad("600000");
            for (int i = 0; i < ricList.Count; i++)
            {
                //Get Chinese Name
                chineseRic.ricCHNNameStr = selenium.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[3]");
                chineseList.Add(chineseRic);
                start_position++;
            }

            //close chinese window
            selenium.Close();
            selenium.Stop();
        }

        /**
         * Get ISIN number from http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.asp
         * by specific official code
         * 
         */
        private void DataCaptureFromPDF(int index)
        {
            ISelenium pdfSearch = new DefaultSelenium("localhost", 4444, "*iexplore", "http://www.hkexnews.hk");
            int start_position = index;
            pdfSearch.Start();
            try
            {

                pdfSearch.Open("/listedco/listconews/advancedsearch/search_active_main.asp");
                pdfSearch.WaitForPageToLoad("60000");

                for (; start_position < ricList.Count - pdf_disable; start_position++)
                {
                    pdfSearch.Type("txt_stock_code", ricList[start_position].ricCodeStr);
                    pdfSearch.Click("//table[@id='Table1']/tbody/tr[3]/td/table[2]/tbody/tr/td/table/tbody/tr[14]/td[3]/label/a[1]/img");
                    pdfSearch.WaitForPageToLoad("30000");

                    String pdfUrl = pdfSearch.GetAttribute("//table[@id='Table4']/tbody/tr[8]/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td/a@href");
                    pdfList.Add(coreObj.PDFAnalysis(pdfUrl, ricList[start_position].ricCodeStr));

                    pdfSearch.Click("//table[@id='Table5']/tbody/tr/td/table/tbody/tr/td[1]/a/img");
                    pdfSearch.WaitForPageToLoad("30000");
                }

                pdfSearch.Close();
                pdfSearch.Stop();

            }
            catch (Exception ex)
            {
                String errLog = ex.ToString();
                coreObj.WriteLogFile(errLog);
                coreObj.WriteLogFile(ricList[start_position].ricCodeStr);
                coreObj.WriteLogFile("start_position=" + start_position);
                pdfSearch.Close();
                pdfSearch.Stop();

                DataCaptureFromPDF(start_position);

            }
        }

        /**
         * Generate FM template file
         * Return   : void
         */
        private void GenerateCBBCTemplate(int start)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();



            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct");
                return;
            }
            xlApp.Visible = false;

            try
            {

                Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                ///Create additional sheet if don't download PDF
                //Worksheet wSheet1 = null;
                //wBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //wSheet1 = (Worksheet)wBook.Worksheets[2];
                //wSheet1.Name = "PDFData";
                //wSheet1.Cells[1, 2] = "Gearing";
                //wSheet1.Cells[1, 3] = "Premium";

                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                wSheet.Name = "FM_HKCBBC";

                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 36;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 42;
                ((Range)wSheet.Columns["A:B", System.Type.Missing]).Font.Name = "Courier";
                xlApp.Cells.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;


                if (wSheet == null)
                {
                    coreObj.WriteLogFile("Worksheet could not be created. Check that your office installation and project references are correct.");
                }



                wSheet.Cells[1, 1] = "Please action the add of the following HK stock on TQS**";

                wSheet.Cells[3, 1] = "FM Serial Number:";
                wSheet.Cells[3, 2] = "HK" + DateTime.Now.Year.ToString().Substring(2) + "-" + fmSerialNumber;


                wSheet.Cells[4, 1] = "Effective Date:";
                ((Range)wSheet.Cells[4, 2]).NumberFormat = "@";
                wSheet.Cells[4, 2] = ricList[0].listingDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));



                int startLine = 6;

                int end = 0;
                if (start + 20 >= ricList.Count)
                {
                    end = ricList.Count;
                }
                else
                {
                    end = start + 20;
                }

                ///Create PDFData sheet if don't download PDF
                //int pdfStart = 2;
                //for (int j = start; j < end; j++)
                //{
                //    wSheet1.Cells[j + 2, 1] = ricList[j].ricCodeStr;
                //}

                //initial FM_HKCBBC sheet
                for (int i = start; i < end; i++)
                {

                    bool isIndex = false;
                    bool isBull = false;
                    bool isHKD = false;

                    //**************************For TQS***********************************

                    wSheet.Cells[startLine++, 1] = "+ADDITION+";
                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";
                    ((Range)wSheet.Cells[startLine, 1]).Font.Bold = true;
                    ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                    wSheet.Cells[startLine++, 1] = "(" + (i + 1) + ")";//dynamic increase
                    wSheet.Cells[startLine++, 1] = "**For TQS**";

                    startLine++;
                    wSheet.Cells[startLine, 1] = "Underlying RIC:";
                    wSheet.Cells[startLine++, 2] = ricList[i].ricCodeStr + ".HK";

                    wSheet.Cells[startLine, 1] = "Composite chain RIC:";
                    wSheet.Cells[startLine++, 2] = "0#" + ricList[i].ricCodeStr + ".HK";

                    wSheet.Cells[startLine, 1] = "Broker page RIC:";
                    wSheet.Cells[startLine++, 2] = ricList[i].ricCodeStr + "bk.HK";

                    wSheet.Cells[startLine, 1] = "Misc.Info page RIC:";
                    wSheet.Cells[startLine++, 2] = ricList[i].ricCodeStr + "MI.HK";

                    wSheet.Cells[startLine, 1] = "Displayname:";
                    wSheet.Cells[startLine++, 2] = ricList[i].ricNameStr;

                    wSheet.Cells[startLine, 1] = "Official Code:";
                    wSheet.Cells[startLine++, 2] = ricList[i].ricCodeStr;

                    wSheet.Cells[startLine, 1] = "Exchange Symbol:";
                    wSheet.Cells[startLine++, 2] = ricList[i].ricCodeStr;

                    wSheet.Cells[startLine, 1] = "Currency:";
                    wSheet.Cells[startLine++, 2] = "HKD";

                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine, 1] = "Recordtype:";
                    wSheet.Cells[startLine++, 2] = "097";

                    wSheet.Cells[startLine, 1] = "Spare_Ubytes8:";
                    wSheet.Cells[startLine++, 2] = "WRNT";

                    wSheet.Cells[startLine, 1] = "Underlying Chain RIC:";
                    wSheet.Cells[startLine++, 2] = "0#CBBC.HK";

                    wSheet.Cells[startLine, 1] = "Underlying Chain RIC:";
                    wSheet.Cells[startLine++, 2] = "0#WARRANTS.HK";

                    wSheet.Cells[startLine, 1] = "Warrant Type:";
                    wSheet.Cells[startLine++, 2] = "Callable " + ricList[i].bullBearStr + " Contracts";

                    wSheet.Cells[startLine, 1] = "Misc Info page Chain RIC:";
                    wSheet.Cells[startLine++, 2] = "0#MI.HK";

                    wSheet.Cells[startLine, 1] = "Lot Size:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0";
                    wSheet.Cells[startLine++, 2] = ricList[i].boardLotStr;

                    startLine++;
                    wSheet.Cells[startLine, 1] = "COI DSPLY_NMLL:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";



                    //For Bull
                    if (ricList[i].bullBearStr == "Bull")
                    {
                        isBull = true;
                        if (i + 1 <= chineseList.Count)
                        {
                            wSheet.Cells[startLine++, 2] = chineseList[i].ricCHNNameStr.Substring(0, 4)
                               + chineseList[i].ricCHNNameStr.Substring(chineseList[i].ricCHNNameStr.Length - 1)
                               + ricList[i].maturityDateDT.Month + "月" + "RC" + ricList[i].maturityDateDT.ToString("yy");
                        }
                        else
                        {
                            wSheet.Cells[startLine++, 2] = "null";
                        }
                    }
                    //For BEAR
                    else
                    {
                        if (i + 1 <= chineseList.Count)
                        {
                            wSheet.Cells[startLine++, 2] = chineseList[i].ricCHNNameStr.Substring(0, 4)
                               + chineseList[i].ricCHNNameStr.Substring(chineseList[i].ricCHNNameStr.Length - 1)
                               + ricList[i].maturityDateDT.Month + "月" + "RP" + ricList[i].maturityDateDT.ToString("yy");

                        }
                        else
                        {
                            wSheet.Cells[startLine++, 2] = "null";
                        }
                    }


                    wSheet.Cells[startLine++, 2] = "-----------------------------------------------------";

                    wSheet.Cells[startLine, 1] = "BCAST_REF:";
                    wSheet.Cells[startLine++, 2] = "n/a";

                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";

                    wSheet.Cells[startLine, 1] = "WNT_RATIO:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.0000000";
                    wSheet.Cells[startLine++, 2] = 1.0 / Convert.ToInt32(ricList[i].entitlementRatioStr);

                    wSheet.Cells[startLine, 1] = "STRIKE_PRC:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.000";

                    //For HKD
                    if (Char.IsLetter(ricList[i].strikeLevelStr, 0))
                    {
                        isHKD = true;
                        wSheet.Cells[startLine++, 2] = ricList[i].strikeLevelStr.Substring(4);
                    }
                    else
                    {
                        wSheet.Cells[startLine++, 2] = ricList[i].strikeLevelStr;
                    }


                    wSheet.Cells[startLine, 1] = "MATUR_DATE:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine++, 2] = ricList[i].maturityDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));

                    wSheet.Cells[startLine, 1] = "LONGLINK3:";
                    wSheet.Cells[startLine++, 2] = "t" + ricList[i].ricCodeStr + ".HK";


                    wSheet.Cells[startLine, 1] = "SPARE_SNUM13:";
                    wSheet.Cells[startLine++, 2] = "1";

                    wSheet.Cells[startLine, 1] = "GN_TX20_3:";
                    wSheet.Cells[startLine++, 2] = "[HK-\"WARRAN*\"]";


                    wSheet.Cells[startLine, 1] = "GN_TX20_6:";

                    //For Index
                    if (Char.IsLetter(ricList[i].underlyingStr, 0))
                    {
                        isIndex = true;
                        if (ricList[i].underlyingStr == "HSCEI")
                        {
                            wSheet.Cells[startLine++, 2] = "<.HSCE>";
                        }
                        else
                        {
                            wSheet.Cells[startLine++, 2] = "<." + ricList[i].underlyingStr + ">";
                        }
                    }
                    //For Equity
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "<" + ricList[i].underlyingStr.Substring(1) + ".HK>";
                    }

                    wSheet.Cells[startLine, 1] = "GN_TX20_7:";
                    wSheet.Cells[startLine++, 2] = "********************";

                    wSheet.Cells[startLine, 1] = "GN_TX20_10:";
                    wSheet.Cells[startLine++, 2] = "CBBC/" + ricList[i].bullBearStr.ToUpper();


                    wSheet.Cells[startLine, 1] = "GN_TX20_11:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0";
                    wSheet.Cells[startLine++, 2] = ricList[i].issueSizeStr;

                    wSheet.Cells[startLine, 1] = "GN_TX20_12 (Misc.Info):";
                    //For Index
                    if (isIndex)
                    {
                        wSheet.Cells[startLine++, 2] = "IDX~~~~~~<" + ricList[i].ricCodeStr + "MI.HK>";
                    }
                    //For Equity
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "HKD~~~~~~<" + ricList[i].ricCodeStr + "MI.HK>";
                    }

                    wSheet.Cells[startLine, 1] = "COUPON RATE:";
                    wSheet.Cells[startLine++, 2] = "n/a";

                    wSheet.Cells[startLine, 1] = "ISSUE PRICE:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.000";
                    wSheet.Cells[startLine++, 2] = ricList[i].issuePriceStr;

                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";

                    startLine++;
                    wSheet.Cells[startLine, 1] = "ROW80_13:";
                    wSheet.Cells[startLine++, 2] = "Callable " + ricList[i].bullBearStr + " Contracts";

                    //**************************For AFE***********************************
                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";

                    wSheet.Cells[startLine++, 1] = "**FOR AFE**";

                    wSheet.Cells[startLine, 1] = "GV1_FLAG:";
                    //For Bull
                    if (isBull)
                    {
                        wSheet.Cells[startLine++, 2] = "C";
                    }
                    //For Bear
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "P";
                    }

                    wSheet.Cells[startLine, 1] = "ISS_TP_FLG:";
                    //For Index
                    if (isIndex)
                    {
                        wSheet.Cells[startLine++, 2] = "I";
                    }
                    //For Equity
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "S";
                    }

                    wSheet.Cells[startLine, 1] = "RDM_CUR:";
                    wSheet.Cells[startLine++, 2] = "344";

                    wSheet.Cells[startLine, 1] = "LONGLINK14:";
                    //For Index
                    if (isIndex)
                    {
                        if (ricList[i].underlyingStr == "HSI")
                        {
                            wSheet.Cells[startLine++, 2] = ".HSI|HKD|1  <-- TQR INSERT DAILY";
                        }
                        else if (ricList[i].underlyingStr == "HSCEI")
                        {
                            wSheet.Cells[startLine++, 2] = ".HSCE|HKD|1  <-- TQR INSERT DAILY";
                        }
                        else if (ricList[i].underlyingStr == "DJI")
                        {
                            wSheet.Cells[startLine++, 2] = ".DJI|USD|1  <-- TQR INSERT DAILY";
                        }
                        else
                        {
                            wSheet.Cells[startLine++, 2] = "Index not equal HSI,HSCEI or DJI";
                        }

                    }
                    //For Equity
                    else
                    {
                        wSheet.Cells[startLine++, 2] = ricList[i].underlyingStr.Substring(1) + ".HK";
                    }

                    wSheet.Cells[startLine, 1] = "BOND_TYPE:";
                    wSheet.Cells[startLine++, 2] = "WARRANTS";

                    wSheet.Cells[startLine++, 1] = "LEG1_STR:";
                    wSheet.Cells[startLine++, 1] = "LEG2_STR:";
                    wSheet.Cells[startLine++, 1] = "GN_TXT24_1:";
                    wSheet.Cells[startLine++, 1] = "GN_TXT24_2:";

                    //**************************For NDA***********************************
                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";
                    wSheet.Cells[startLine++, 1] = "**For NDA**";
                    wSheet.Cells[startLine++, 1] = "New Organisation listing";
                    startLine++;
                    wSheet.Cells[startLine++, 1] = "Primary Listing (RIC):";

                    wSheet.Cells[startLine, 1] = "IDN Longname:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine++, 2] = IDNLongNameFormat(ricList[i], isIndex, isBull, isHKD);

                    wSheet.Cells[startLine++, 1] = "Organisation Name (DIRNAME):";
                    wSheet.Cells[startLine++, 1] = "Geographical Entity:";
                    wSheet.Cells[startLine++, 1] = "Organisation Type:";
                    wSheet.Cells[startLine++, 1] = "Alias (Previous Name):";
                    wSheet.Cells[startLine++, 1] = "Alias (General):";
                    wSheet.Cells[startLine++, 1] = "Issue Classification:";
                    wSheet.Cells[startLine++, 1] = "MSCI code:";
                    wSheet.Cells[startLine++, 1] = "Business Activity:";
                    startLine++;
                    wSheet.Cells[startLine++, 1] = "Existing Organisation listing";
                    startLine++;
                    wSheet.Cells[startLine++, 1] = "Primary Listing (RIC):";

                    wSheet.Cells[startLine, 1] = "Organisation Name (DIRNAME):";
                    wSheet.Cells[startLine++, 2] = ((IssuerCode)issuerCodeHT[ricList[i].issuerIDStr]).fullName;

                    wSheet.Cells[startLine, 1] = "Issue Classification:";
                    wSheet.Cells[startLine++, 2] = "WNT";

                    //**************************For WRT_CNR***********************************
                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";
                    wSheet.Cells[startLine++, 1] = "**For WRT_CNR**";

                    wSheet.Cells[startLine, 1] = "Gearing:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.00";
                    if (i >= pdfList.Count)
                    {
                        wSheet.Cells[startLine++, 2] = "0.00";
                    }
                    else
                    {
                        wSheet.Cells[startLine++, 2] = pdfList[i].gearStr;
                    }

                    //wSheet.Cells[startLine++, 2] = "=PDFData!B" + pdfStart;

                    wSheet.Cells[startLine, 1] = "Premium:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.00";
                    if (i >= pdfList.Count)
                    {
                        wSheet.Cells[startLine++, 2] = "0.00";
                    }
                    else
                    {
                        wSheet.Cells[startLine++, 2] = pdfList[i].premiumStr;
                    }
                    //wSheet.Cells[startLine++, 2] = "=PDFData!C" + pdfStart++;


                    wSheet.Cells[startLine, 1] = "Announcement Date:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine++, 2] = ricList[i].launchDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));

                    wSheet.Cells[startLine, 1] = "Payment Date:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine++, 2] = ricList[i].clearingCommDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));

                    wSheet.Cells[startLine, 1] = "Call Level:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.000";

                    //For HKD
                    if (isHKD)
                    {

                        wSheet.Cells[startLine++, 2] = ricList[i].callLevelStr.Substring(4);
                    }
                    else
                    {
                        wSheet.Cells[startLine++, 2] = ricList[i].callLevelStr;
                    }



                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";
                    startLine++;
                }


                xlApp.DisplayAlerts = false;
                xlApp.AlertBeforeOverwriting = false;

                String fullpath = coreObj.Log_Path + "\\" + coreObj.SubFolder + "\\" + "HK" + (DateTime.Now.Year).ToString().Substring(2) + "-" + fmSerialNumber + "_SEC";

                for (int i = start; i < end; i++)
                {
                    fullpath += "_" + ricList[i].ricCodeStr;
                }

                fullpath += "_" + ricList[start].listingDateDT.ToString("ddMMMyyyy", new CultureInfo("en-US")) + ".xls";

                wBook.SaveCopyAs(fullpath);


            }
            catch (Exception ex)
            {
                coreObj.WriteLogFile(ex.ToString());
            }
            finally
            {
                xlApp.Quit();
                coreObj.KillExcelProcess(xlApp);
            }
        }

        /**
        * Read config infomation and set value to vars
        * 
        */
        public void ReadConfigFile()
        {

            XmlDocument xmlDoc = new XmlDocument();

            try
            {
                xmlDoc.Load(CONFIGFILE_NAME);
                coreObj.PDF_Path = xmlDoc.SelectSingleNode("/*/pdf_path").InnerText;
                coreObj.Log_Path = xmlDoc.SelectSingleNode("/*/log_path").InnerText;
                coreObj.LogName = LOGFILE_NAME;
                coreObj.SubFolder = FOLDER_NAME;


                isKeepPDFFile = xmlDoc.SelectSingleNode("/*/IsKeepPDF").InnerText;
                startPosition = Convert.ToInt32(xmlDoc.SelectSingleNode("/*/cbbc_start_position").InnerText);
                holidayCount = Convert.ToInt32(xmlDoc.SelectSingleNode("/*/holidaycount").InnerText);
                pdf_disable = Convert.ToInt32(xmlDoc.SelectSingleNode("/*/pdf_disable").InnerText);

                XmlNode root = xmlDoc.SelectSingleNode("/*/HK_IssuerCode");

                for (int i = 1; i < root.ChildNodes.Count + 1; i++)
                {
                    String tempCode = xmlDoc.SelectSingleNode("/*/*/trans" + i + "/code").InnerText;
                    IssuerCode tempTrans = new IssuerCode();
                    tempTrans.fullName = xmlDoc.SelectSingleNode("/*/*/trans" + i + "/fullname").InnerText;
                    tempTrans.shortName = xmlDoc.SelectSingleNode("/*/*/trans" + i + "/shortname").InnerText;
                    tempTrans.warrentIssuer = xmlDoc.SelectSingleNode("/*/*/trans" + i + "/warrentIssuer").InnerText;
                    issuerCodeHT.Add(tempCode, tempTrans);
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString() + "Cannot find config file or file error, use default value");
                System.Windows.Forms.Application.Exit();

            }

        }

        public string organisationFullName(string issueIDStr)
        {
            issuerCodeHT = new Hashtable();
            ReadConfigFile();
            return ((IssuerCode)issuerCodeHT[issueIDStr]).fullName;
        }

        /**
         * Additional method, apply correct format for IDN longname
         * Retrun: String
         * Parameter: String sourceStr
         */
        public String IDNLongNameFormat(HKRicTemplate ricObj, bool isIndex, bool isBull, bool isHKD)
        {
            String idnLongName = "";
            //For Index
            if (isIndex)
            {
                if (ricObj.underlyingStr == "HSI")
                {
                    idnLongName = "HANG SENG@";
                }
                if (ricObj.underlyingStr == "HSCEI")
                {
                    idnLongName = "HANG SENG C E I@";
                }
                if (ricObj.underlyingStr == "DJI")
                {
                    idnLongName = "DJ INDU AVERAGE@";
                }
            }
            //For Stock
            else
            {
                idnLongName += ricObj.underlyingNameForStockStr + "@";
            }
            idnLongName += ((IssuerCode)issuerCodeHT[ricObj.issuerIDStr]).shortName + " ";
            idnLongName += ricObj.maturityDateDT.ToString("MMMyy", new CultureInfo("en-US")).ToUpper() + " ";


            //Attach Strike Price from Strike Level
            //For HKD
            if (isHKD)
            {

                idnLongName += ricObj.strikeLevelStr.Substring(4);
            }
            else
            {
                idnLongName += ricObj.strikeLevelStr;
            }

            idnLongName += " ";
            //For Bull
            if (isBull)
            {
                idnLongName += "C";
            }
            //For Bear
            else
            {
                idnLongName += "P";
            }

            if (isIndex)
            {
                idnLongName += "IR";
            }
            else
            {
                idnLongName += "R";
            }



            return idnLongName;
        }

        /**
         * SetupTest method call other method to read xml, create folder and start selenium RC
         * Retrun: void
         * 
         */
        private void SetupTest()
        {
            ReadConfigFile();
            coreObj.CreateDir(coreObj.Log_Path + "\\" + coreObj.SubFolder);
            coreObj.CreateDir(coreObj.PDF_Path + "\\" + coreObj.SubFolder);
            StartRC();
        }

        /**
         * ClearTest method call other method to read xml, create folder and start selenium RC
         * Retrun: void
         * Parameter: bool isKeepPDFFile
         */
        private void ClearTest(String isKeepPDFFile)
        {
            ShutDownRC();
            if (isKeepPDFFile == "N")
            {
                coreObj.DeleteTempDir(coreObj.PDF_Path + "\\" + coreObj.SubFolder);
            }
        }



    }
}
