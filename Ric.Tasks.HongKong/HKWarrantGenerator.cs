using System;
using System.Collections;
using System.Collections.Generic;
using Selenium;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.IO;
using System.Windows.Forms;
using System.Globalization;
using Ric.Shared;

namespace Ric.Tasks.HongKong
{
    public class HKWarrantGenerator
    {
        Core coreObj = new Core();
        private ISelenium selenium1;
        int startPosition = 2;
        int endPosition = 500;

        String fmSerialNumber = "0000";
        Hashtable issuerCodeHT = new Hashtable();
        //Store ric english info
        List<HKRicTemplate> ricList = new List<HKRicTemplate>();
        //Store ric chinese name
        List<HKRicTemplate> chineseList = new List<HKRicTemplate>();
        //Store gearing and premium get from pdf
        List<HKRicTemplate> pdfList = new List<HKRicTemplate>();

        const string CONFIGFILE_NAME = "Config\\HK\\HK_IssuerCode.xml";
        const string LOGFILE_NAME = "WarrantLog.txt";
        const string FOLDER_NAME = "Warrant";

        //Delay hours from user click start button to automation start
        private int delayHour = 0;
        //Flag whether keep or delete downloaded pdf files
        private string isKeepPDFFile = "N";
        //Number of holiday, include Satureday and Sunday if met Country holiday
        private int holidayCount = 0;
        private DateTime scheduleDate = DateTime.Now;

        public Core CoreObj
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

        ///<summary>
        ///StartGrabber method is out side interface, used to start automation
        ///</summary>
        ///<returns> void </returns>
        public void StartGrabber()
        {

            System.Threading.Thread.Sleep(delayHour * 60 * 60 * 1000);

            DateTime start = DateTime.Now;

            SetupTest();
            //main method to capture data
            WarrantDataCapture();

            int fileCount = (int)Math.Ceiling(ricList.Count / 20.0);
            for (int i = 0; i < fileCount; i++)
            {
                GenerateWarrantTemplate(i * 20);
                fmSerialNumber = coreObj.UpdateFMSerialNumber(fmSerialNumber);
            }



            ClearTest(isKeepPDFFile);

            DateTime end = DateTime.Now;
            TimeSpan runningtime = end.Subtract(start);

            coreObj.WriteLogFile("---------------------------------------------------------------------");
            coreObj.WriteLogFile("Generated Ric#:" + ricList.Count);
            coreObj.WriteLogFile("Execution Time:" + runningtime.ToString());

        }

        ///<summary>
        ///Start selenium remote controller
        ///</summary>
        ///<returns> void </returns>
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
                selenium1 = new DefaultSelenium("localhost", 4444, "*iexplore", "http://www.hkex.com.hk");
                selenium1.Start();
                selenium1.UseXpathLibrary("javascript-xpath");
            }
            catch (System.Exception e)
            {
                MessageBox.Show("Selenium console not started." + e.ToString());
                System.Windows.Forms.Application.Exit();
            }

        }

        ///<summary>
        ///Shutdown selenium remote controller
        ///</summary>
        ///<returns> void </returns>
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

        ///<summary>
        ///Grab warrant data from http://www.hkex.com.hk/eng/dwrc/newissue/newlaunch.htm
        ///</summary>
        ///<returns> void </returns>
        private void WarrantDataCapture()
        {

            int start_position = startPosition;

            try
            {
                selenium1.OpenWindow("/eng/dwrc/newissue/newlaunch.htm", "WarrantListWindow");
                selenium1.WaitForPopUp("WarrantListWindow", "30000");
                selenium1.SelectWindow("WarrantListWindow");

                //Get English Content
                //while (selenium1.IsElementPresent("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[1]"))
                while (start_position <= endPosition && selenium1.IsElementPresent("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[1]"))
                {

                    String launchDateStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[11]");
                    DateTime tempLaunchDateDT = DateTime.ParseExact(launchDateStr, "dd-MM-yyyy", null);
                    DateTime temp = coreObj.DateCalculate(scheduleDate, tempLaunchDateDT, holidayCount);

                    if (temp.Date == scheduleDate.Date)
                    {
                        HKRicTemplate hkRic = new HKRicTemplate();
                        hkRic.launchDateDT = DateTime.ParseExact(launchDateStr, "dd-MM-yyyy", null);
                        hkRic.ricCodeStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[2]");
                        hkRic.ricNameStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[3]");
                        hkRic.issuerIDStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[4]");
                        hkRic.underlyingStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[5]");
                        hkRic.callPutStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[6]");
                        hkRic.boardLotStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[7]");
                        hkRic.strikeLevelStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[8]");
                        hkRic.entitlementRatioStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[9]");
                        hkRic.issueSizeStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[10]");
                        hkRic.clearingCommDateDT = DateTime.ParseExact(selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[12]"), "dd-MM-yyyy", null);
                        hkRic.listingDateDT = DateTime.ParseExact(selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[13]"), "dd-MM-yyyy", null);
                        hkRic.maturityDateDT = DateTime.ParseExact(selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[14]"), "dd-MM-yyyy", null);


                        //Get issue price
                        String url = selenium1.GetAttribute("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[2]/a@href");
                        String id = url.Substring(url.IndexOf('=') + 1, 5);
                        String summaryUrl = "http://www.hkex.com.hk/eng/dwrc/dwsummary.asp?id=" + id;
                        selenium1.OpenWindow(summaryUrl, "WarrantSummaryWindow" + start_position);
                        selenium1.WaitForPopUp("WarrantSummaryWindow" + start_position, "30000");
                        selenium1.SelectWindow("WarrantSummaryWindow" + start_position);

                        //For Equity Get underlying name
                        if (Char.IsDigit(hkRic.underlyingStr, 0))
                        {
                            hkRic.underlyingNameForStockStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr[3]/td[2]").Substring(8);
                        }

                        hkRic.issuePriceStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr[11]/td[2]").Substring(4);

                        selenium1.Close();

                        ricList.Add(hkRic);


                        selenium1.SelectWindow("WarrantListWindow");

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
                selenium1.Close();
                selenium1.Stop();
                CaptureRetry(start_position);

            }

        }//end WarrantDataCapture()

        ///<summary>
        ///Recursive Method, start capture data from the breakpoint
        ///</summary>
        ///<param name="start_position">Position of ric that interrupted since time out</param>
        ///<returns> void </returns>
        private void CaptureRetry(int start_position)
        {

            selenium1.Start();
            selenium1.UseXpathLibrary("javascript-xpath");

            try
            {
                selenium1.OpenWindow("/eng/dwrc/newissue/newlaunch.htm", "WarrantListWindow");
                selenium1.WaitForPopUp("WarrantListWindow", "30000");
                selenium1.SelectWindow("WarrantListWindow");

                //Get English Content
                while (start_position <= endPosition && selenium1.IsElementPresent("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[1]"))
                {

                    String launchDateStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[11]");
                    DateTime tempLaunchDateDT = DateTime.ParseExact(launchDateStr, "dd-MM-yyyy", null);
                    DateTime temp = coreObj.DateCalculate(scheduleDate, tempLaunchDateDT, holidayCount);

                    if (temp.Date == scheduleDate.Date)
                    {
                        HKRicTemplate hkRic = new HKRicTemplate();
                        hkRic.launchDateDT = DateTime.ParseExact(launchDateStr, "dd-MM-yyyy", null);
                        hkRic.ricCodeStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[2]");
                        hkRic.ricNameStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[3]");
                        hkRic.issuerIDStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[4]");
                        hkRic.underlyingStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[5]");
                        hkRic.callPutStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[6]");
                        hkRic.boardLotStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[7]");
                        hkRic.strikeLevelStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[8]");
                        hkRic.entitlementRatioStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[9]");
                        hkRic.issueSizeStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[10]");
                        hkRic.clearingCommDateDT = DateTime.ParseExact(selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[12]"), "dd-MM-yyyy", null);
                        hkRic.listingDateDT = DateTime.ParseExact(selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[13]"), "dd-MM-yyyy", null);
                        hkRic.maturityDateDT = DateTime.ParseExact(selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[14]"), "dd-MM-yyyy", null);


                        //Get issue price
                        String url = selenium1.GetAttribute("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[2]/a@href");
                        String id = url.Substring(url.IndexOf('=') + 1, 5);
                        String summaryUrl = "http://www.hkex.com.hk/eng/dwrc/dwsummary.asp?id=" + id;
                        selenium1.OpenWindow(summaryUrl, "WarrantSummaryWindow" + start_position);
                        selenium1.WaitForPopUp("WarrantSummaryWindow" + start_position, "30000");
                        selenium1.SelectWindow("WarrantSummaryWindow" + start_position);

                        //For Equity Get underlying name
                        if (Char.IsDigit(hkRic.underlyingStr, 0))
                        {
                            hkRic.underlyingNameForStockStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr[3]/td[2]").Substring(8);
                        }

                        hkRic.issuePriceStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr[11]/td[2]").Substring(4);
                        selenium1.Close();

                        ricList.Add(hkRic);
                        selenium1.SelectWindow("WarrantListWindow");

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
                selenium1.Close();
                selenium1.Stop();
                CaptureRetry(start_position);
            }
        }

        ///<summary>
        ///Get all stocks' chinese name
        ///</summary>
        ///<param name="lastRickIndex">Position of last ric</param>
        ///<returns> void </returns>
        private void ChineseNameCapture(int lastRickIndex)
        {
            //Get Chinese Content
            HKRicTemplate chineseRic = new HKRicTemplate();
            int start_position = lastRickIndex - ricList.Count;

            selenium1.Click("link=繁體");
            selenium1.WaitForPageToLoad("600000");
            for (int i = 0; i < ricList.Count; i++)
            {
                //Get Chinese Name
                chineseRic.ricCHNNameStr = selenium1.GetText("//table[@id='bodytable']/tbody/tr/td/table/tbody[1]/tr[" + start_position + "]/td[3]");
                chineseList.Add(chineseRic);
                start_position++;
            }

            //close chinese window
            selenium1.Close();
            selenium1.Stop();
        }

        ///<summary>
        ///Recursive Method, recall this method when timeout or other errors happend
        ///Call PDFAnalysis method to download PDF from from http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.asp
        ///and extract data from PDF  by specific stock code
        ///</summary>
        ///<param name="index">ricList index</param>
        ///<returns> void </returns>
        private void DataCaptureFromPDF(int index)
        {
            ISelenium pdfSearch1 = new DefaultSelenium("localhost", 4444, "*iexplore", "http://www.hkexnews.hk");
            int start_position = index;
            pdfSearch1.Start();
            try
            {

                pdfSearch1.Open("/listedco/listconews/advancedsearch/search_active_main.asp");
                pdfSearch1.WaitForPageToLoad("60000");

                for (; start_position < ricList.Count; start_position++)
                {


                    pdfSearch1.Type("txt_stock_code", ricList[start_position].ricCodeStr);
                    pdfSearch1.Click("//table[@id='Table1']/tbody/tr[3]/td/table[2]/tbody/tr/td/table/tbody/tr[14]/td[3]/label/a[1]/img");
                    pdfSearch1.WaitForPageToLoad("30000");

                    String pdfUrl = pdfSearch1.GetAttribute("//table[@id='Table4']/tbody/tr[8]/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td/a@href");
                    pdfList.Add(coreObj.PDFAnalysis(pdfUrl, ricList[start_position].ricCodeStr));

                    pdfSearch1.Click("//table[@id='Table5']/tbody/tr/td/table/tbody/tr/td[1]/a/img");
                    pdfSearch1.WaitForPageToLoad("30000");
                }

                pdfSearch1.Close();
                pdfSearch1.Stop();

            }
            catch (Exception ex)
            {
                String errLog = ex.ToString();
                coreObj.WriteLogFile(errLog);
                coreObj.WriteLogFile(ricList[start_position].ricCodeStr);
                coreObj.WriteLogFile("start_position=" + start_position);
                pdfSearch1.Close();
                pdfSearch1.Stop();

                DataCaptureFromPDF(start_position);

            }
        }

        ///<summary>
        ///Generate FM template files for Warrant
        ///</summary>
        ///<param name="start">ricList index</param>
        ///<returns> void </returns>
        private void GenerateWarrantTemplate(int start)
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

                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                wSheet.Name = "FM_HK_WARRANT";

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

                //initial FM_HK_WARRANT sheet
                for (int i = start; i < end; i++)
                {

                    bool isIndex = false;
                    bool isStock = false;
                    bool isOil = false;
                    bool isCommodity = false;
                    bool isCurrency = false;
                    bool isCall = false;
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
                    //for Equity stock
                    if (Char.IsDigit(ricList[i].underlyingStr, 0))
                    {
                        isStock = true;
                        wSheet.Cells[startLine++, 2] = "0#" + ricList[i].underlyingStr.Substring(1) + "W.HK";
                    }
                    //for Index
                    else if (ricList[i].underlyingStr == "HSI" || ricList[i].underlyingStr == "HSCEI" || ricList[i].underlyingStr == "DJI")
                    {
                        isIndex = true;
                        if (ricList[i].underlyingStr == "HSCEI")
                        {
                            wSheet.Cells[startLine++, 2] = "0#.HSCEW.HK";
                        }
                        else
                        {
                            wSheet.Cells[startLine++, 2] = "0#." + ricList[i].underlyingStr + "W.HK";
                        }
                    }
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "";
                    }


                    wSheet.Cells[startLine, 1] = "Chain RIC:";
                    //for Equity stock
                    if (isStock)
                    {
                        wSheet.Cells[startLine++, 2] = "0#CWRTS.HK";
                    }
                    //for Index
                    else if (isIndex)
                    {
                        wSheet.Cells[startLine++, 2] = "";
                    }
                    //for Oil
                    else if (ricList[i].ricNameStr.Contains("OIL"))
                    {
                        isOil = true;
                        wSheet.Cells[startLine++, 2] = "0#OWRTS.HK";
                    }
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "0#MISCW.HK";
                    }

                    wSheet.Cells[startLine, 1] = "Chain RIC:";
                    wSheet.Cells[startLine++, 2] = "0#WARRANTS.HK";

                    wSheet.Cells[startLine, 1] = "Warrant Type:";
                    //for Equity stock
                    if (isStock)
                    {
                        wSheet.Cells[startLine++, 2] = "Equity Warrant";
                    }
                    //for Index
                    else if (isIndex)
                    {
                        if (ricList[i].underlyingStr == "HSI")
                        {
                            wSheet.Cells[startLine++, 2] = "Hang Seng Index Warrant";
                        }
                        else if (ricList[i].underlyingStr == "HSCEI")
                        {
                            wSheet.Cells[startLine++, 2] = "Hang Seng China Enterprises Index Warrant";
                        }
                        else//DJI
                        {
                            wSheet.Cells[startLine++, 2] = "DJ Industrial Average Index Warrant";
                        }

                    }
                    //for Oil
                    else if (isOil)
                    {
                        wSheet.Cells[startLine++, 2] = "Oil Warrant";
                    }
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "Miscellaneous Warrant";
                    }

                    wSheet.Cells[startLine, 1] = "Misc Info page Chain RIC:";
                    wSheet.Cells[startLine++, 2] = "0#MI.HK";

                    wSheet.Cells[startLine, 1] = "Lot Size:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0";
                    wSheet.Cells[startLine++, 2] = ricList[i].boardLotStr;

                    startLine++;
                    wSheet.Cells[startLine, 1] = "COI DSPLY_NMLL:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";

                    //For Call
                    if (ricList[i].callPutStr == "Call")
                    {
                        isCall = true;
                        //The last char of warrant name is a character
                        if (Char.IsLetter(ricList[i].ricNameStr, ricList[i].ricNameStr.Length - 1))
                        {
                            wSheet.Cells[startLine++, 2] = chineseList[i].ricCHNNameStr.Substring(0, 4)
                               + ricList[i].ricNameStr.Substring(ricList[i].ricNameStr.Length - 1)
                               + ricList[i].maturityDateDT.Month + "月" + "CW" + ricList[i].maturityDateDT.ToString("yy");
                        }
                        else
                        {
                            wSheet.Cells[startLine++, 2] = chineseList[i].ricCHNNameStr.Substring(0, 4)
                               + ricList[i].maturityDateDT.Month + "月" + "CW" + ricList[i].maturityDateDT.ToString("yy");
                        }

                    }
                    //For Put
                    else
                    {
                        if (Char.IsLetter(ricList[i].ricNameStr, ricList[i].ricNameStr.Length - 1))
                        {
                            wSheet.Cells[startLine++, 2] = chineseList[i].ricCHNNameStr.Substring(0, 4)
                               + ricList[i].ricNameStr.Substring(ricList[i].ricNameStr.Length - 1)
                               + ricList[i].maturityDateDT.Month + "月" + "PW" + ricList[i].maturityDateDT.ToString("yy");

                        }
                        else
                        {

                            wSheet.Cells[startLine++, 2] = chineseList[i].ricCHNNameStr.Substring(0, 4)
                               + ricList[i].maturityDateDT.Month + "月" + "PW" + ricList[i].maturityDateDT.ToString("yy");
                        }

                    }


                    wSheet.Cells[startLine++, 2] = "-----------------------------------------------------";

                    wSheet.Cells[startLine, 1] = "BCAST_REF:";
                    wSheet.Cells[startLine++, 2] = "n/a";

                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";

                    wSheet.Cells[startLine, 1] = "WNT_RATIO:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.0000000";
                    /*=======================================Error=======================================*/
                    //wSheet.Cells[startLine++, 2] = 1.0 / Convert.ToInt32(ricList[i].entitlementRatioStr);
                    wSheet.Cells[startLine++, 2] = 1.0 / Convert.ToDouble(ricList[i].entitlementRatioStr);

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

                    //for Equity stock
                    if (isStock)
                    {
                        wSheet.Cells[startLine++, 2] = "<" + ricList[i].underlyingStr.Substring(1) + ".HK>";
                    }
                    //for Index
                    else if (isIndex)
                    {
                        if (ricList[i].underlyingStr == "HSCEI")
                        {
                            wSheet.Cells[startLine++, 2] = "<.HSCE>";
                        }
                        else
                        {
                            wSheet.Cells[startLine++, 2] = "<." + ricList[i].underlyingStr + ">";
                        }
                    }
                    //for others
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "";
                    }

                    wSheet.Cells[startLine, 1] = "GN_TX20_7:";
                    wSheet.Cells[startLine++, 2] = "********************";

                    wSheet.Cells[startLine, 1] = "GN_TX20_10:";
                    wSheet.Cells[startLine++, 2] = "EU/" + ricList[i].callPutStr.ToUpper();


                    wSheet.Cells[startLine, 1] = "GN_TX20_11:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0";
                    wSheet.Cells[startLine++, 2] = ricList[i].issueSizeStr;

                    wSheet.Cells[startLine, 1] = "GN_TX20_12 (Misc.Info):";
                    //For Index
                    if (isIndex)
                    {
                        wSheet.Cells[startLine++, 2] = "IDX~~~~~~<" + ricList[i].ricCodeStr + "MI.HK>";
                    }
                    //For Stock
                    else if (isStock)
                    {
                        wSheet.Cells[startLine++, 2] = "HKD~~~~~~<" + ricList[i].ricCodeStr + "MI.HK>";
                    }
                    else//for other
                    {
                        wSheet.Cells[startLine++, 2] = "";
                    }

                    wSheet.Cells[startLine, 1] = "COUPON RATE:";
                    wSheet.Cells[startLine++, 2] = "n/a";

                    wSheet.Cells[startLine, 1] = "ISSUE PRICE:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.000";
                    wSheet.Cells[startLine++, 2] = ricList[i].issuePriceStr;

                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";

                    wSheet.Cells[startLine, 1] = "**FOR " + ricList[i].ricCodeStr + "bk.HK**";
                    startLine++;


                    wSheet.Cells[startLine, 1] = "ROW80_13:";
                    if (isStock)
                    {
                        wSheet.Cells[startLine++, 2] = "Warrant Type-Equity Warrant";
                    }
                    else if (isIndex)
                    {
                        wSheet.Cells[startLine++, 2] = "Warrant Type-Index Warrant";
                    }
                    else if (isOil)
                    {
                        wSheet.Cells[startLine++, 2] = "Warrant Type-Future Warrant";
                    }
                    else if (ricList[i].ricNameStr.Contains("GOLD") || ricList[i].ricNameStr.Contains("SILVER"))
                    {
                        isCommodity = true;
                        wSheet.Cells[startLine++, 2] = "Warrant Type-Commodity Warrant";
                    }
                    else
                    {
                        isCurrency = true;
                        wSheet.Cells[startLine++, 2] = "Warrant Type-Currency Warrant";
                    }





                    //**************************For AFE***********************************
                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";

                    wSheet.Cells[startLine++, 1] = "**FOR AFE**";

                    wSheet.Cells[startLine, 1] = "GV1_FLAG:";
                    //For Call
                    if (isCall)
                    {
                        wSheet.Cells[startLine++, 2] = "C";
                    }
                    //For Put
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
                    //For Stock
                    else if (isStock)
                    {
                        wSheet.Cells[startLine++, 2] = "S";
                    }
                    else//For Forex ricList[i].ricNameStr.Contains("EURUS/AUDUS/USYEN")
                    {
                        wSheet.Cells[startLine++, 2] = "F";
                    }

                    wSheet.Cells[startLine, 1] = "RDM_CUR:";
                    if (isStock || isIndex)
                    {
                        wSheet.Cells[startLine++, 2] = "344";
                    }
                    else if (isOil || isCommodity)
                    {
                        wSheet.Cells[startLine++, 2] = "840";
                    }
                    else if (isCurrency)
                    {
                        if (ricList[i].ricNameStr.Contains("YEN"))
                        {
                            wSheet.Cells[startLine++, 2] = "392";
                        }
                        else
                        {
                            wSheet.Cells[startLine++, 2] = "840";
                        }
                    }
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "";
                    }

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
                    //For Stock
                    else if (isStock)
                    {
                        wSheet.Cells[startLine++, 2] = ricList[i].underlyingStr.Substring(1) + ".HK";
                    }
                    else
                    {
                        wSheet.Cells[startLine++, 2] = "";
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
                    wSheet.Cells[startLine++, 2] = IDNLongNameFormat(ricList[i], isIndex, isStock, isCall, isHKD);

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
                    wSheet.Cells[startLine++, 2] = "WT";

                    //**************************For WRT_CNR***********************************
                    wSheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";
                    wSheet.Cells[startLine++, 1] = "**For WRT_CNR**";

                    wSheet.Cells[startLine, 1] = "Gearing:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.00";

                    /*=======================Error =======================*/
                    wSheet.Cells[startLine++, 2] = pdfList[i].gearStr;



                    wSheet.Cells[startLine, 1] = "Premium:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "0.00";
                    /*=======================Error =======================*/
                    wSheet.Cells[startLine++, 2] = pdfList[i].premiumStr;



                    wSheet.Cells[startLine, 1] = "Announcement Date:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine++, 2] = ricList[i].launchDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));

                    wSheet.Cells[startLine, 1] = "Payment Date:";
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine++, 2] = ricList[i].clearingCommDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));


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

        ///<summary>
        ///Read config infomation from HK_IssuerCode.xml and set default value to vars
        ///</summary>
        ///<returns> void</returns>
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
                startPosition = Convert.ToInt32(xmlDoc.SelectSingleNode("/*/warrant_start_position").InnerText);
                endPosition = Convert.ToInt32(xmlDoc.SelectSingleNode("/*/warrant_end_position").InnerText);
                holidayCount = Convert.ToInt32(xmlDoc.SelectSingleNode("/*/holidaycount").InnerText);

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

        ///<summary>
        ///Additional method, apply correct format for IDN longname
        ///</summary>
        ///<param name="ricObj">HKRicTemplate Object</param>
        ///<param name="isIndex">Flag to judge if one ric is a Index</param>
        ///<param name="isStock">Flag to judge if one ric is a Stock</param>
        ///<param name="isCall">Flag to judge if one ric is a Call or Put</param>
        ///<param name="isHKD">Flag to judge if one ric's Strike Price include HKD string</param>
        ///<returns> string:idnLongName </returns>
        public string IDNLongNameFormat(HKRicTemplate ricObj, bool isIndex, bool isStock, bool isCall, bool isHKD)
        {
            string idnLongName = "";
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
                //For Call
                if (isCall)
                {
                    idnLongName += "C";
                }
                //For Put
                else
                {
                    idnLongName += "P";
                }

                if (isIndex)
                {
                    idnLongName += "IW";
                }
                else
                {
                    idnLongName += "WT";
                }
            }
            //For Stock
            if (isStock)
            {
                idnLongName = ricObj.underlyingNameForStockStr + "@";
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
                //For Call
                if (isCall)
                {
                    idnLongName += "C";
                }
                //For Put
                else
                {
                    idnLongName += "P";
                }

                if (isIndex)
                {
                    idnLongName += "IW";
                }
                else
                {
                    idnLongName += "WT";
                }

            }

            return idnLongName;
        }

        ///<summary>
        ///SetupTest method call other method to read xml, create folder and start selenium RC
        ///</summary>
        ///<returns> void </returns>
        private void SetupTest()
        {
            ReadConfigFile();
            coreObj.CreateDir(coreObj.Log_Path + "\\" + coreObj.SubFolder);
            coreObj.CreateDir(coreObj.PDF_Path + "\\" + coreObj.SubFolder);
            StartRC();
        }

        ///<summary>
        ///ClearTest method used to shutdown RC and delete pdf files according configurable item in HK_IssuerCode.xml
        ///</summary>
        ///<param name="isKeepPDFFile">whether keep or delete downloaded pdf files</param>
        ///<returns> void </returns>
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
