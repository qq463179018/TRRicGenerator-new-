using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
////using Reuters.ProcessQuality.ContentAuto.Lib;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    public class HKCBBCFMSubGenerator
    {
        private StreamWriter swCBBCLog = null;
        private StreamWriter swChineseNameLog = null;

        private string fmSerialNumberCBBC = string.Empty;

        private Logger logger = null;
        private string outputPath = "";

        private HKFMAndBulkFileGenerator parent = null;
        private List<TaskResultEntry> taskResultList = new List<TaskResultEntry>();
        private readonly string cbbcMainPageUri = "http://www.hkex.com.hk/eng/cbbc/newissue/newlaunch.htm";

        public List<RicInfo> RicList = new List<RicInfo>();
        public List<RicInfo> RicChineseList = new List<RicInfo>();

        public HKCBBCFMSubGenerator(HKFMAndBulkFileGenerator parent)
        {
            this.parent = parent;
        }

        public void CBBCGenerate()
        {
            DateTime start = DateTime.Now;
            RicList = GetRicList();
            int fileCount = (int)Math.Ceiling(RicList.Count / 20.0);
            for (int i = 0; i < fileCount; i++)
            {
                GenerateCBBCTemplate(i * 20, RicList);
                fmSerialNumberCBBC = parent.UpdateFMSerialNumber(fmSerialNumberCBBC);
            }

            DateTime end = DateTime.Now;
            TimeSpan runningTime = end.Subtract(start);


            swCBBCLog.WriteLine("---------------------------------------------------------------------");
            swCBBCLog.WriteLine("Generated Ric#:" + RicList.Count);
            swCBBCLog.WriteLine("Execution Time:" + runningTime.ToString());
            swCBBCLog.Close();

            swChineseNameLog.WriteLine("RIC\t\tChineseName");
            foreach (RicInfo ric in RicChineseList)
            {
                swChineseNameLog.WriteLine(ric.Code + "\t\t" + ric.ChineseName);
            }
            swChineseNameLog.Close();
        }

        private void GenerateCBBCTemplate(int start, List<RicInfo> ricInfoList)
        {
            List<HKRicTemplate> ricTemplateList = new List<HKRicTemplate>();
            foreach (RicInfo ricInfo in ricInfoList)
            {
                HKRicTemplate ric = new HKRicTemplate(ricInfo, FMType.Cbbc);
                ricTemplateList.Add(ric);
            }
            try
            {
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    app.ExcelAppInstance.AlertBeforeOverwriting = false;

                    Workbook workbook = app.ExcelAppInstance.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                    if (worksheet == null)
                    {
                        throw new Exception("Worksheet could not be created. Check that your office installation and project references are correct.");
                    }
                    worksheet.Name = "FM_HKCBBC";

                    ((Range)worksheet.Columns["A", System.Type.Missing]).ColumnWidth = 36;
                    ((Range)worksheet.Columns["B", System.Type.Missing]).ColumnWidth = 42;
                    ((Range)worksheet.Columns["A:B", System.Type.Missing]).Font.Name = "Courier";
                    app.ExcelAppInstance.Cells.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    worksheet.Cells[1, 1] = "Please action the add of the following HK stock on TQS**";

                    worksheet.Cells[3, 1] = "FM Serial Number:";
                    //Serial Number
                    worksheet.Cells[3, 2] = "HK" + DateTime.Now.Year.ToString().Substring(2) + "-" + fmSerialNumberCBBC;
                    worksheet.Cells[4, 1] = "Effective Date:";
                    ((Range)worksheet.Cells[4, 2]).NumberFormat = "@";
                    worksheet.Cells[4, 2] = ricTemplateList[0].EffectiveDate;
                    int startLine = 6;
                    int end = 0;
                    if (start + 20 >= ricTemplateList.Count)
                    {
                        end = ricTemplateList.Count;
                    }
                    else
                    {
                        end = start + 20;
                    }

                    #region writer template file for each ric
                    for (int i = start; i < end; i++)
                    {
                        HKRicTemplate ricTemplate = ricTemplateList[i];
                        worksheet.Cells[startLine++, 1] = "+ADDITION+";
                        worksheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";
                        ((Range)worksheet.Cells[startLine, 1]).Font.Bold = true;
                        ((Range)worksheet.Cells[startLine, 1]).NumberFormat = "@";
                        worksheet.Cells[startLine++, 1] = "(" + (i + 1) + ")";//dynamic increase
                        worksheet.Cells[startLine++, 1] = "**For TQS**";
                        startLine++;
                        worksheet.Cells[startLine, 1] = "Underlying RIC:";

                        worksheet.Cells[startLine++, 2] = ricTemplate.UnderLyingRic;

                        worksheet.Cells[startLine, 1] = "Composite chain RIC:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.CompositeChainRic;
                        worksheet.Cells[startLine, 1] = "Broker page RIC:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.BrokerPageRic;

                        worksheet.Cells[startLine, 1] = "Misc.Info page RIC:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.MiscInfoPageRic;
                        worksheet.Cells[startLine, 1] = "Displayname:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.DisplayName;

                        worksheet.Cells[startLine, 1] = "Official Code:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.OfficicalCode;

                        worksheet.Cells[startLine, 1] = "Exchange Symbol:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.ExchangeSymbol;

                        worksheet.Cells[startLine, 1] = "Currency:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.Currency;

                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "@";
                        worksheet.Cells[startLine, 1] = "Recordtype:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.RecordType;

                        worksheet.Cells[startLine, 1] = "Spare_Ubytes8:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.SpareUbytes8;

                        worksheet.Cells[startLine, 1] = "Underlying Chain RIC:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.UnderlyingChainRic1;

                        worksheet.Cells[startLine, 1] = "Underlying Chain RIC:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.UnderlyingChainRic2;

                        worksheet.Cells[startLine, 1] = "Warrant Type:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.WarrantType;

                        worksheet.Cells[startLine, 1] = "Misc Info page Chain RIC:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.MiscInfoPageChainRic;

                        worksheet.Cells[startLine, 1] = "Lot Size:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "0";
                        worksheet.Cells[startLine++, 2] = ricTemplate.LotSize;

                        startLine++;
                        worksheet.Cells[startLine, 1] = "COI DSPLY_NMLL:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "@";
                        worksheet.Cells[startLine++, 2] = ricTemplate.ColDsplyNmll;

                        worksheet.Cells[startLine++, 2] = "-----------------------------------------------------";

                        worksheet.Cells[startLine, 1] = "BCAST_REF:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.BcastRef;

                        worksheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";

                        worksheet.Cells[startLine, 1] = "WNT_RATIO:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "0.0000000";
                        worksheet.Cells[startLine++, 2] = ricTemplate.WntRation;

                        worksheet.Cells[startLine, 1] = "STRIKE_PRC:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "0.000";
                        worksheet.Cells[startLine++, 2] = ricTemplate.StrikPrc;

                        worksheet.Cells[startLine, 1] = "MATUR_DATE:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "@";
                        worksheet.Cells[startLine++, 2] = ricTemplate.MaturDate;

                        worksheet.Cells[startLine, 1] = "LONGLINK3:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.LongLink3;

                        worksheet.Cells[startLine, 1] = "SPARE_SNUM13:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.SpareSnum13;

                        worksheet.Cells[startLine, 1] = "GN_TX20_3:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.GNTX20_3;

                        worksheet.Cells[startLine, 1] = "GN_TX20_6:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.GNTX20_6;

                        worksheet.Cells[startLine, 1] = "GN_TX20_7:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.GNTX20_7;

                        worksheet.Cells[startLine, 1] = "GN_TX20_10:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.GNTX20_10;


                        worksheet.Cells[startLine, 1] = "GN_TX20_11:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "0";
                        worksheet.Cells[startLine++, 2] = ricTemplate.GNTX20_11;

                        worksheet.Cells[startLine, 1] = "GN_TX20_12 (Misc.Info):";
                        worksheet.Cells[startLine++, 2] = ricTemplate.GNTX20_12;

                        worksheet.Cells[startLine, 1] = "COUPON RATE:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.CouponRate;

                        worksheet.Cells[startLine, 1] = "ISSUE PRICE:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "0.000";
                        worksheet.Cells[startLine++, 2] = ricTemplate.IssuePrice;

                        worksheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";

                        startLine++;
                        worksheet.Cells[startLine, 1] = "ROW80_13:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.Row80_13;
                        //**************************For AFE***********************************
                        worksheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";

                        worksheet.Cells[startLine++, 1] = "**FOR AFE**";

                        worksheet.Cells[startLine, 1] = "GV1_FLAG:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.GVFlag;

                        worksheet.Cells[startLine, 1] = "ISS_TP_FLG:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.IssTpFlg;

                        worksheet.Cells[startLine, 1] = "RDM_CUR:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.RdmCur;

                        worksheet.Cells[startLine, 1] = "LONGLINK14:";
                        worksheet.Cells[startLine++, 2] = ricTemplate.LongLink14;

                        worksheet.Cells[startLine, 1] = "BOND_TYPE:";
                        worksheet.Cells[startLine++, 2] = "WARRANTS";

                        worksheet.Cells[startLine++, 1] = "LEG1_STR:";
                        worksheet.Cells[startLine++, 1] = "LEG2_STR:";
                        worksheet.Cells[startLine++, 1] = "GN_TXT24_1:";
                        worksheet.Cells[startLine++, 1] = "GN_TXT24_2:";

                        //**************************For NDA***********************************
                        worksheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";
                        worksheet.Cells[startLine++, 1] = "**For NDA**";
                        worksheet.Cells[startLine++, 1] = "New Organisation listing";
                        startLine++;
                        worksheet.Cells[startLine++, 1] = "Primary Listing (RIC):";

                        worksheet.Cells[startLine, 1] = "IDN Longname:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "@";
                        worksheet.Cells[startLine++, 2] = ricTemplate.IdnLongName;
                        worksheet.Cells[startLine++, 1] = "Organisation Name (DIRNAME):";
                        worksheet.Cells[startLine++, 1] = "Geographical Entity:";
                        worksheet.Cells[startLine++, 1] = "Organisation Type:";
                        worksheet.Cells[startLine++, 1] = "Alias (Previous Name):";
                        worksheet.Cells[startLine++, 1] = "Alias (General):";
                        worksheet.Cells[startLine++, 1] = "Issue Classification:";
                        worksheet.Cells[startLine++, 1] = "MSCI code:";
                        worksheet.Cells[startLine++, 1] = "Business Activity:";
                        startLine++;
                        worksheet.Cells[startLine++, 1] = "Existing Organisation listing";
                        startLine++;
                        worksheet.Cells[startLine++, 1] = "Primary Listing (RIC):";

                        worksheet.Cells[startLine, 1] = "Organisation Name (DIRNAME):";
                        worksheet.Cells[startLine++, 2] = ricTemplate.OrgnizationName2;
                        worksheet.Cells[startLine, 1] = "Issue Classification:";
                        worksheet.Cells[startLine++, 2] = "WNT";

                        //**************************For WRT_CNR***********************************
                        worksheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";
                        worksheet.Cells[startLine++, 1] = "**For WRT_CNR**";

                        worksheet.Cells[startLine, 1] = "Gearing:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "0.00";

                        worksheet.Cells[startLine++, 2] = ricTemplate.Gearing;

                        worksheet.Cells[startLine, 1] = "Premium:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "0.00";

                        worksheet.Cells[startLine++, 2] = ricTemplate.Premium;

                        worksheet.Cells[startLine, 1] = "Announcement Date:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "@";
                        worksheet.Cells[startLine++, 2] = ricTemplate.AnnouncementDate;

                        worksheet.Cells[startLine, 1] = "Payment Date:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "@";
                        worksheet.Cells[startLine++, 2] = ricTemplate.PaymentDate;

                        worksheet.Cells[startLine, 1] = "Call Level:";
                        ((Range)worksheet.Cells[startLine, 2]).NumberFormat = "0.000";
                        worksheet.Cells[startLine++, 2] = ricTemplate.CallLevel;

                        worksheet.Cells[startLine++, 1] = "---------------------------------------------------------------------------------------------------";
                        startLine++;

                    }
                    #endregion
                    string fullPath = outputPath + "\\" + parent.RicTemplatePath + "\\" + parent.CbbcSubDir + "\\" + "HK" + (DateTime.Now.Year).ToString().Substring(2) + "-" + fmSerialNumberCBBC + "_SEC";
                    if (!Directory.Exists(Path.GetDirectoryName(fullPath)))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(fullPath));
                    }
                    for (int j = start; j < end; j++)
                    {
                        fullPath += "_" + ricInfoList[j].Code;
                    }
                    fullPath += ".xls";
                    fullPath += "_" + DateTime.ParseExact(ricInfoList[0].ListingDate, "dd-MM-yyyy", null).ToString("ddMMMyyyy", new CultureInfo("en-US")) + ".xls";

                    workbook.SaveCopyAs(fullPath);
                    taskResultList.Add(new TaskResultEntry("CBBC FM File", "CBBC FM File Path", fullPath));

                }
            }
            catch (Exception ex)
            {
                logger.Log("Generate CBBC FM Files Error: " + ex.Message);
            }

        }

        private HtmlAgilityPack.HtmlNodeCollection GetNewlyCBBCRicNode(string uri)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = WebClientUtil.GetHtmlDocument(uri, 180000);
            var node = htmlDoc.DocumentNode.SelectNodes("html/body/font/table/tbody/tr/td/table/tbody")[0].ChildNodes[1];
            HtmlAgilityPack.HtmlNodeCollection ricNodeList = node.SelectNodes("//td/table/tr/td/table/tbody/tr[@class='tr_normal']");
            if (ricNodeList == null || ricNodeList.Count == 0)
            {
                logger.Log("There's no newly launched CBBC. Please go to " + uri + "to have a check.");
            }
            return ricNodeList;
        }

        public List<RicInfo> GetRicChineseInfo()
        {
            List<RicInfo> cbbcRicChineseList = new List<RicInfo>();
            HtmlAgilityPack.HtmlNodeCollection ricNodeList = GetNewlyCBBCRicNode("http://www.hkex.com.hk/chi/cbbc/newissue/newlaunch_c.htm");
            int endPos = ricNodeList.Count < parent.CbbcEndPos ? ricNodeList.Count : parent.CbbcEndPos;
            for (int i = parent.CbbcStartPos - 1; i < endPos; i++)
            {
                HtmlAgilityPack.HtmlNode ricNode = ricNodeList[i];
                RicInfo ricInfo = new RicInfo();
                ricInfo.Code = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 1 + 1].InnerText);
                ricInfo.ChineseName = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 2 + 1].InnerText);
                cbbcRicChineseList.Add(ricInfo);
            }
            return cbbcRicChineseList;
        }

        //Get information from http://www.hkex.com.hk/eng/cbbc/cbbcsummary.asp?id={0}"
        private string[] GetUnderLyingForStockAndIssuePrice(string id)
        {
            string[] arr = { "", "" };
            string uri = string.Format("http://www.hkex.com.hk/eng/cbbc/cbbcsummary.asp?id={0}", id);
            HtmlAgilityPack.HtmlDocument htmlDoc = WebClientUtil.GetHtmlDocument(uri, 180000);
            var nodes = htmlDoc.DocumentNode.SelectNodes("//body/font/table/tbody/tr/td//table/tbody/tr/td/table/tbody/tr");
            foreach (HtmlAgilityPack.HtmlNode node in nodes)
            {
                if (node.ChildNodes[2 * 0 + 1].InnerText == "Issue Price (Initial Issue):")
                {
                    arr[1] = MiscUtil.GetCleanTextFromHtml(node.ChildNodes[2 * 1 + 1].InnerText).Substring(4);
                }

                else if (node.ChildNodes[2 * 0 + 1].InnerText == "Underlying:")
                {
                    arr[0] = MiscUtil.GetCleanTextFromHtml(node.ChildNodes[2 * 1 + 1].InnerText).Substring(8);
                }
            }
            return arr;
        }

        public List<RicInfo> GetRicList()
        {
            List<RicInfo> cbbcRicList = new List<RicInfo>();
            RicChineseList = GetRicChineseInfo();
            HtmlAgilityPack.HtmlNodeCollection ricNodeList = GetNewlyCBBCRicNode(cbbcMainPageUri);

            try
            {

                int endPos = ricNodeList.Count < parent.CbbcEndPos ? ricNodeList.Count : parent.CbbcEndPos;
                for (int i = parent.CbbcStartPos - 1; i < endPos; i++)
                {
                    HtmlAgilityPack.HtmlNode ricNode = ricNodeList[i];
                    RicInfo ricInfo = new RicInfo();
                    ricInfo.Code = ricNode.ChildNodes[2 * 1 + 1].InnerText;

                    ricInfo.Underlying = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 4 + 1].InnerText);

                    string id = ricNode.ChildNodes[2 * 1 + 1].ChildNodes[0].Attributes["href"].Value;
                    id = id.Substring(id.IndexOf('=') + 1);
                    id = id.Substring(0, id.IndexOf("'"));

                    // id = id.Substring(id.IndexOf('=') + 1, 5);
                    string[] arr = GetUnderLyingForStockAndIssuePrice(id);
                    if (Char.IsDigit(ricInfo.Underlying, 0))
                    {
                        ricInfo.UnderlyingNameForStock = arr[0];
                    }
                    ricInfo.IssuerPrice = arr[1];

                    ricInfo.Name = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 2 + 1].InnerText);
                    ricInfo.Issuer = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 3 + 1].InnerText);
                    ricInfo.BullBear = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 5 + 1].InnerText);
                    ricInfo.BoardLot = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 6 + 1].InnerText);
                    ricInfo.StrikeLevel = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 7 + 1].InnerText);
                    ricInfo.CallLevel = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 8 + 1].InnerText);
                    ricInfo.EntitlementRatio = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 9 + 1].InnerText);
                    ricInfo.TotalIssueSize = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 10 + 1].InnerText);
                    ricInfo.LauntchDate = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 11 + 1].InnerText);
                    ricInfo.ClearingCommencementDate = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 12 + 1].InnerText);
                    ricInfo.ListingDate = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 13 + 1].InnerText);
                    ricInfo.MaturityDate = MiscUtil.GetCleanTextFromHtml(ricNode.ChildNodes[2 * 14 + 1].InnerText);

                    //Get Chinese name information
                    foreach (RicInfo ric in RicChineseList)
                    {
                        if (ric.Code == ricInfo.Code)
                        {
                            ricInfo.ChineseName = ric.ChineseName;
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    //Get Gear and Premium from PDF
                    if (parent.yesOrNo.Equals("Yes"))//if download pdf
                        parent.PDFAnalysis(ricInfo, FMType.Cbbc);       //parent.PDFAnalysis(ricInfo, FMType.Warrant);

                    cbbcRicList.Add(ricInfo);
                }
            }
            catch (Exception ex)
            {
                string errInfo = ex.ToString();
            }
            return cbbcRicList;
        }

        public void Start()
        {
            CBBCGenerate();
        }

        public void Initialize(string outputPath, Logger logger, List<TaskResultEntry> taskResultList)
        {
            //this.logger = logger;
            this.RicList = parent.RicListCBBC;
            this.RicChineseList = parent.ChineseListCBBC;
            this.taskResultList = taskResultList;
            this.outputPath = outputPath;

            if (!Directory.Exists(outputPath + "\\" + parent.RicTemplatePath + "\\" + parent.CbbcSubDir))
            {
                Directory.CreateDirectory(outputPath + "\\" + parent.RicTemplatePath + "\\" + parent.CbbcSubDir);
            }

            string swCBBCLogPath = outputPath + "\\" + parent.RicTemplatePath + "\\" + parent.CbbcSubDir +
                "\\" + parent.CbbcLogName;
            swCBBCLog = new StreamWriter(swCBBCLogPath, false);
            swCBBCLog.AutoFlush = true;
            taskResultList.Add(new TaskResultEntry("CBBCLog.txt", "Log File for CBBC", swCBBCLogPath));
            string swChineseNameLogPath = outputPath + "\\" + parent.RicTemplatePath + "\\" + parent.CbbcSubDir
                + "\\" + "CBBCChineseNameLog_" + DateTime.Now.ToString("dd_MMM_yyyy", null) + ".txt";
            swChineseNameLog = new StreamWriter(swChineseNameLogPath, false);
            swChineseNameLog.AutoFlush = true;
            taskResultList.Add(new TaskResultEntry("CBBCChineseNameLog", "Chinese Name Log File for CBBC", swChineseNameLogPath));

            fmSerialNumberCBBC = parent.CbbcFmSn;

        }

        public void Cleanup()
        {
            if (swCBBCLog != null)
            {
                swCBBCLog.Close();
                swCBBCLog = null;
            }

            if (swChineseNameLog != null)
            {
                swChineseNameLog.Close();
                swChineseNameLog = null;
            }
        }
    }
}
