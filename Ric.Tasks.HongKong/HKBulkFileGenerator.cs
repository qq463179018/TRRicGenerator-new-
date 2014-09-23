using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Collections;
using Ric.Shared;

namespace Ric.Tasks.HongKong
{
    public class HKBulkFileGenerator
    {
        Core coreObj = new Core();
        HKCBBCGenerator cbbcGenerator = new HKCBBCGenerator();
        HKWarrantGenerator warrantGenerator = new HKWarrantGenerator();

        public void StartHKBulkFileGeneratorJob()
        {

            //Generate IAAdd.csv
            Generate_IAAdd_CSV(warrantGenerator.CoreObj.Log_Path, cbbcGenerator.RicList, warrantGenerator.RicList);

            //Generate QAAdd.csv
           Generate_QAAdd_CSV(warrantGenerator.CoreObj.Log_Path, cbbcGenerator.RicList, warrantGenerator.RicList);

            //Generate HKG_EQLB_CBBC.txt
            Generate_HKG_EQLB_CBBC(cbbcGenerator.CoreObj.Log_Path, cbbcGenerator.RicList, cbbcGenerator.CHNList);

            //Generate HKG_EQLB.txt
            Generate_HKG_EQLB(warrantGenerator.CoreObj.Log_Path, warrantGenerator.RicList, warrantGenerator.CHNList);

            //Generate HKG_EQLBMI.txt
            Generate_HKG_EQLBMI(warrantGenerator.CoreObj.Log_Path, cbbcGenerator.RicList, warrantGenerator.RicList);

            //Generate BK.txt
            Generate_BK(warrantGenerator.CoreObj.Log_Path, cbbcGenerator.RicList, warrantGenerator.RicList);

            //Generate MI.txt
            Generate_MI(warrantGenerator.CoreObj.Log_Path, cbbcGenerator.RicList, warrantGenerator.RicList);

            //Generate MAIN.txt
            //bulkFileGenerator.Generate_MAIN(warrantGenerator.CoreObj.Log_Path, cbbcGenerator.RicList, warrantGenerator.RicList);
        }

        public void Generate_IAAdd_CSV(string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList)
        {
            cbbcGenerator.ReadConfigFile();
            string formatDate = System.DateTime.Now.ToString("yyyyMMdd");
            string fullpath = filePath + "\\YS" + formatDate + "IAAdd.csv";

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
                wSheet.Name = "YS" + formatDate + "IAAdd";
                if (wSheet == null)
                {
                    coreObj.WriteLogFile("Worksheet could not be created. Check that your office installation and project references are correct.");
                }
                wSheet.Cells[1, 1] = "HONG KONG CODE";
                wSheet.Cells[1, 2] = "TYPE";
                wSheet.Cells[1, 3] = "CATEGORY";
                wSheet.Cells[1, 4] = "RCS ASSET CLASS";
                wSheet.Cells[1, 5] = "WARRANT ISSUER";

                for (int i = 0; i < cbbcList.Count; i++)
                {
                    wSheet.Cells[i + 2, 1] = cbbcList[i].ricCodeStr;
                    wSheet.Cells[i + 2, 2] = "DERIVATIVE";
                    wSheet.Cells[i + 2, 3] = "EIW";
                    wSheet.Cells[i + 2, 4] = "FXKNOCKOUT";
                    string organisationName = cbbcGenerator.organisationFullName(cbbcList[i].issuerIDStr);
                    wSheet.Cells[i + 2, 5] = cbbcGenerator.GetWarrantIssuer(organisationName);

                }

                for (int j = 0; j < warrantList.Count; j++)
                {
                    wSheet.Cells[j + cbbcList.Count + 2, 1] = warrantList[j].ricCodeStr;
                    wSheet.Cells[j + cbbcList.Count + 2, 2] = "DERIVATIVE";
                    wSheet.Cells[j + cbbcList.Count + 2, 3] = "EIW";
                    wSheet.Cells[j + cbbcList.Count + 2, 4] = "TRAD";
                    string organisationName = warrantGenerator.organisationFullName(warrantList[j].issuerIDStr);
                    wSheet.Cells[j + cbbcList.Count + 2, 5] = cbbcGenerator.GetWarrantIssuer(organisationName);
                }

                xlApp.DisplayAlerts = false;
                xlApp.AlertBeforeOverwriting = false;
                wBook.SaveAs(fullpath, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
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

        public void Generate_QAAdd_CSV(string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList)
        {
            int cbbcListCount = cbbcList.Count;
            int warrantListCount = warrantList.Count;
            string formatDate = System.DateTime.Now.ToString("yyyyMMdd");
            string fullpath = filePath + "\\YS" + formatDate + "QAAdd.csv";
            HKCBBCGenerator cbbcGenerator = new HKCBBCGenerator();
            HKWarrantGenerator warrantGenerator = new HKWarrantGenerator();
            cbbcGenerator.ReadConfigFile();
            warrantGenerator.ReadConfigFile();

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
                wSheet.Name = "YS" + formatDate + "QAAdd";
                if (wSheet == null)
                {
                    coreObj.WriteLogFile("Worksheet could not be created. Check that your office installation and project references are correct.");
                }
                wSheet.Cells[1, 1] = "RIC";
                wSheet.Cells[1, 2] = "TAG";
                wSheet.Cells[1, 3] = "ASSET COMMON NAME";
                wSheet.Cells[1, 4] = "ASSET SHORT NAME";
                wSheet.Cells[1, 5] = "CURRENCY";
                wSheet.Cells[1, 6] = "EXCHANGE";
                wSheet.Cells[1, 7] = "TYPE";
                wSheet.Cells[1, 8] = "CATEGORY";
                wSheet.Cells[1, 9] = "BASE ASSET";
                wSheet.Cells[1, 10] = "EXPIRY DATE";
                wSheet.Cells[1, 11] = "STRIKE PRICE";
                wSheet.Cells[1, 12] = "CALL PUT OPTION";
                wSheet.Cells[1, 13] = "ROUND LOT SIZE";
                wSheet.Cells[1, 14] = "TRADING SEGMENT";
                wSheet.Cells[1, 15] = "TICKER SYMBOL";
                wSheet.Cells[1, 16] = "DERIVATIVES FIRST TRADING DAY";
                wSheet.Cells[1, 17] = "ISSUE PRICE";
                wSheet.Cells[1, 18] = "WARRANT ISSUE QUANTITY";

                for (int i = 0; i < cbbcList.Count; i++)
                {
                    bool isIndex = false;
                    bool isBull = false;
                    bool isHKD = false;
                    if (Char.IsLetter(cbbcList[i].underlyingStr, 0))
                    {
                        isIndex = true;
                    }
                    if (cbbcList[i].bullBearStr == "Bull")
                    {
                        isBull = true;
                    }
                    if (Char.IsLetter(cbbcList[i].strikeLevelStr, 0))
                    {
                        isHKD = true;
                    }

                    wSheet.Cells[i + 2, 1] = cbbcList[i].ricCodeStr + ".HK";
                    wSheet.Cells[i + 2, 2] = "2037";
                    wSheet.Cells[i + 2, 3] = cbbcGenerator.IDNLongNameFormat(cbbcList[i], isIndex, isBull, isHKD).Replace('@', ' ');
                    wSheet.Cells[i + 2, 4] = cbbcList[i].ricNameStr.Replace('@', ' ');
                    wSheet.Cells[i + 2, 5] = "HKD";
                    wSheet.Cells[i + 2, 6] = "HKG";
                    wSheet.Cells[i + 2, 7] = "DERIVATIVE";
                    wSheet.Cells[i + 2, 8] = "EIW";
                    wSheet.Cells[i + 2, 9] = "";
                    ((Range)wSheet.Cells[i + 2, 10]).NumberFormat = "@";
                    wSheet.Cells[i + 2, 10] = cbbcList[i].maturityDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[i + 2, 11]).NumberFormat = "0.000";
                    if (Char.IsLetter(cbbcList[i].strikeLevelStr, 0))
                    {
                        wSheet.Cells[i + 2, 11] = cbbcList[i].strikeLevelStr.Substring(4);
                    }
                    else
                    {
                        wSheet.Cells[i + 2, 11] = cbbcList[i].strikeLevelStr;
                    }
                    if (cbbcList[i].bullBearStr.ToUpper().Equals("BULL"))
                    {
                        wSheet.Cells[i + 2, 12] = "CALL";
                    }
                    else if (cbbcList[i].bullBearStr.ToUpper().Equals("BEAR"))
                    {
                        wSheet.Cells[i + 2, 12] = "PUT";
                    }
                    ((Range)wSheet.Cells[i + 2, 13]).NumberFormat = "0";
                    wSheet.Cells[i + 2, 13] = cbbcList[i].boardLotStr;

                    wSheet.Cells[i + 2, 14] = "HKG:XHKG";
                    wSheet.Cells[i + 2, 15] = cbbcList[i].ricCodeStr;
                    wSheet.Cells[i + 2, 16] = cbbcList[i].listingDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[i + 2, 17]).NumberFormat = "0.000";
                    wSheet.Cells[i + 2, 17] = cbbcList[i].issuePriceStr;
                    ((Range)wSheet.Cells[i + 2, 18]).NumberFormat = "0";
                    wSheet.Cells[i + 2, 18] = cbbcList[i].issueSizeStr;

                }
                for (int j = 0; j < cbbcList.Count; j++)
                {
                    bool isIndex = false;
                    bool isBull = false;
                    bool isHKD = false;
                    if (Char.IsLetter(cbbcList[j].underlyingStr, 0))
                    {
                        isIndex = true;
                    }
                    if (cbbcList[j].bullBearStr == "Bull")
                    {
                        isBull = true;
                    }
                    if (Char.IsLetter(cbbcList[j].strikeLevelStr, 0))
                    {
                        isHKD = true;
                    }

                    wSheet.Cells[j + cbbcListCount + 2, 1] = cbbcList[j].ricCodeStr + "ta.HK";
                    wSheet.Cells[j + cbbcListCount + 2, 2] = "40115";
                    wSheet.Cells[j + cbbcListCount + 2, 3] = cbbcGenerator.IDNLongNameFormat(cbbcList[j], isIndex, isBull, isHKD).Replace('@', ' ');
                    wSheet.Cells[j + cbbcListCount + 2, 4] = cbbcList[j].ricNameStr.Replace('@', ' ');
                    wSheet.Cells[j + cbbcListCount + 2, 5] = "HKD";
                    wSheet.Cells[j + cbbcListCount + 2, 6] = "HKG";
                    wSheet.Cells[j + cbbcListCount + 2, 7] = "DERIVATIVE";
                    wSheet.Cells[j + cbbcListCount + 2, 8] = "EIW";
                    wSheet.Cells[j + cbbcListCount + 2, 9] = "";
                    ((Range)wSheet.Cells[j + cbbcListCount + 2, 10]).NumberFormat = "@";
                    wSheet.Cells[j + cbbcListCount + 2, 10] = cbbcList[j].maturityDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[j + cbbcListCount + 2, 11]).NumberFormat = "0.000";
                    if (Char.IsLetter(cbbcList[j].strikeLevelStr, 0))
                    {
                        wSheet.Cells[j + cbbcListCount + 2, 11] = cbbcList[j].strikeLevelStr.Substring(4);
                    }
                    else
                    {
                        wSheet.Cells[j + cbbcListCount + 2, 11] = cbbcList[j].strikeLevelStr;
                    }
                    if (cbbcList[j].bullBearStr.ToUpper().Equals("BULL"))
                    {
                        wSheet.Cells[j + cbbcListCount + 2, 12] = "CALL";
                    }
                    else if (cbbcList[j].bullBearStr.ToUpper().Equals("BEAR"))
                    {
                        wSheet.Cells[j + cbbcListCount + 2, 12] = "PUT";
                    }
                    ((Range)wSheet.Cells[j + cbbcListCount + 2, 13]).NumberFormat = "0";
                    wSheet.Cells[j + cbbcListCount + 2, 13] = cbbcList[j].boardLotStr;

                    wSheet.Cells[j + cbbcListCount + 2, 14] = "";
                    wSheet.Cells[j + cbbcListCount + 2, 15] = "";
                    wSheet.Cells[j + cbbcListCount + 2, 16] = cbbcList[j].listingDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[j + cbbcListCount + 2, 17]).NumberFormat = "0.000";
                    wSheet.Cells[j + cbbcListCount + 2, 17] = cbbcList[j].issuePriceStr;
                    ((Range)wSheet.Cells[j + cbbcListCount + 2, 18]).NumberFormat = "0";
                    wSheet.Cells[j + cbbcListCount + 2, 18] = cbbcList[j].issueSizeStr;

                }
                for (int k = 0; k < warrantList.Count; k++)
                {
                    bool isIndex = false;
                    bool isHKD = false;
                    bool isStock = false;
                    bool isCall = false;
                    if (warrantList[k].underlyingStr == "HSI" || warrantList[k].underlyingStr == "HSCEI" || warrantList[k].underlyingStr == "DJI")
                    {
                        isIndex = true;
                    }
                    if (Char.IsLetter(warrantList[k].strikeLevelStr, 0))
                    {
                        isHKD = true;
                    }
                    if (Char.IsDigit(warrantList[k].underlyingStr, 0))
                    {
                        isStock = true;
                    }

                    if (warrantList[k].callPutStr == "Call")
                    {
                        isCall = true;
                    }

                    wSheet.Cells[k + 2 * cbbcListCount + 2, 1] = warrantList[k].ricCodeStr + ".HK";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 2] = "2037";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 3] = warrantGenerator.IDNLongNameFormat(warrantList[k], isIndex, isStock, isCall, isHKD).Replace('@', ' ');
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 4] = warrantList[k].ricNameStr.Replace('@', ' ');
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 5] = "HKD";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 6] = "HKG";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 7] = "DERIVATIVE";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 8] = "EIW";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 9] = "";
                    ((Range)wSheet.Cells[k + 2 * cbbcListCount + 2, 10]).NumberFormat = "@";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 10] = warrantList[k].maturityDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[k + 2 * cbbcListCount + 2, 11]).NumberFormat = "0.000";
                    if (Char.IsLetter(warrantList[k].strikeLevelStr, 0))
                    {
                        wSheet.Cells[k + 2 * cbbcListCount + 2, 11] = warrantList[k].strikeLevelStr.Substring(4);
                    }
                    else
                    {
                        wSheet.Cells[k + 2 * cbbcListCount + 2, 11] = warrantList[k].strikeLevelStr;
                    }

                    wSheet.Cells[k + 2 * cbbcListCount + 2, 12] = warrantList[k].callPutStr.ToUpper();

                    ((Range)wSheet.Cells[k + 2 * cbbcListCount + 2, 13]).NumberFormat = "0";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 13] = warrantList[k].boardLotStr;

                    wSheet.Cells[k + 2 * cbbcListCount + 2, 14] = "HKG:XHKG";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 15] = warrantList[k].ricCodeStr;

                    wSheet.Cells[k + 2 * cbbcListCount + 2, 16] = warrantList[k].listingDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[k + 2 * cbbcListCount + 2, 17]).NumberFormat = "0.000";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 17] = warrantList[k].issuePriceStr;
                    ((Range)wSheet.Cells[k + 2 * cbbcListCount + 2, 18]).NumberFormat = "0";
                    wSheet.Cells[k + 2 * cbbcListCount + 2, 18] = warrantList[k].issueSizeStr;
                }
                for (int m = 0; m < warrantList.Count; m++)
                {
                    bool isIndex = false;
                    bool isHKD = false;
                    bool isStock = false;
                    bool isCall = false;
                    if (warrantList[m].underlyingStr == "HSI" || warrantList[m].underlyingStr == "HSCEI" || warrantList[m].underlyingStr == "DJI")
                    {
                        isIndex = true;
                    }
                    if (Char.IsLetter(warrantList[m].strikeLevelStr, 0))
                    {
                        isHKD = true;
                    }
                    if (Char.IsDigit(warrantList[m].underlyingStr, 0))
                    {
                        isStock = true;
                    }

                    if (warrantList[m].callPutStr == "Call")
                    {
                        isCall = true;
                    }

                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 1] = warrantList[m].ricCodeStr + "ta.HK";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 2] = "40115";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 3] = warrantGenerator.IDNLongNameFormat(warrantList[m], isIndex, isStock, isCall, isHKD).Replace('@', ' ');
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 4] = warrantList[m].ricNameStr.Replace('@', ' ');
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 5] = "HKD";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 6] = "HKG";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 7] = "DERIVATIVE";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 8] = "EIW";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 9] = "";
                    ((Range)wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 10]).NumberFormat = "@";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 10] = warrantList[m].maturityDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 11]).NumberFormat = "0.000";
                    if (Char.IsLetter(warrantList[m].strikeLevelStr, 0))
                    {
                        wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 11] = warrantList[m].strikeLevelStr.Substring(4);
                    }
                    else
                    {
                        wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 11] = warrantList[m].strikeLevelStr;
                    }

                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 12] = warrantList[m].callPutStr.ToUpper();

                    ((Range)wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 13]).NumberFormat = "0";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 13] = warrantList[m].boardLotStr;

                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 14] = "";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 15] = "";

                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 16] = warrantList[m].listingDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 17]).NumberFormat = "0.000";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 17] = warrantList[m].issuePriceStr;
                    ((Range)wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 18]).NumberFormat = "0";
                    wSheet.Cells[m + 2 * cbbcListCount + warrantListCount + 2, 18] = warrantList[m].issueSizeStr;
                }
                for (int n = 0; n < warrantList.Count; n++)
                {
                    bool isIndex = false;
                    bool isHKD = false;
                    bool isStock = false;
                    bool isCall = false;
                    if (warrantList[n].underlyingStr == "HSI" || warrantList[n].underlyingStr == "HSCEI" || warrantList[n].underlyingStr == "DJI")
                    {
                        isIndex = true;
                    }
                    if (Char.IsLetter(warrantList[n].strikeLevelStr, 0))
                    {
                        isHKD = true;
                    }
                    if (Char.IsDigit(warrantList[n].underlyingStr, 0))
                    {
                        isStock = true;
                    }

                    if (warrantList[n].callPutStr == "Call")
                    {
                        isCall = true;
                    }

                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 1] = warrantList[n].ricCodeStr + ".IXH";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 2] = "46111";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 3] = warrantGenerator.IDNLongNameFormat(warrantList[n], isIndex, isStock, isCall, isHKD).Replace('@', ' ');
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 4] = warrantList[n].ricNameStr.Replace('@', ' ');
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 5] = "HKD";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 6] = "IXH";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 7] = "DERIVATIVE";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 8] = "EIW";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 9] = "";
                    ((Range)wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 10]).NumberFormat = "@";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 10] = warrantList[n].maturityDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 11]).NumberFormat = "0.000";
                    if (Char.IsLetter(warrantList[n].strikeLevelStr, 0))
                    {
                        wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 11] = warrantList[n].strikeLevelStr.Substring(4);
                    }
                    else
                    {
                        wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 11] = warrantList[n].strikeLevelStr;
                    }

                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 12] = warrantList[n].callPutStr.ToUpper();

                    ((Range)wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 13]).NumberFormat = "0";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 13] = warrantList[n].boardLotStr;

                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 14] = "";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 15] = warrantList[n].ricCodeStr;

                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 16] = warrantList[n].listingDateDT.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    ((Range)wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 17]).NumberFormat = "0.000";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 17] = warrantList[n].issuePriceStr;
                    ((Range)wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 18]).NumberFormat = "0";
                    wSheet.Cells[n + 2 * cbbcListCount + 2 * warrantListCount + 2, 18] = warrantList[n].issueSizeStr;
                }

                xlApp.DisplayAlerts = false;
                xlApp.AlertBeforeOverwriting = false;
                wBook.SaveAs(fullpath, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
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
         * Generate HKG_EQLB_CBBC.txt file(for Main RICs of CBBC)
         * Return   : void
         * Parameter: string filePath, List<HKRicTemplate> ricList, List<HKRicTemplate> chineseList
         */
        public void Generate_HKG_EQLB_CBBC
            (string filePath, List<HKRicTemplate> ricList, List<HKRicTemplate> chineseList)
        {
            string fullpath = filePath + "\\HKG_EQLB_CBBC.txt";
            //string[] content = new string[ricList.Count];
            string[] content = new string[ricList.Count + 1];
            content[0] = "SYMBOL\tDSPLY_NAME\tRIC\tOFFCL_CODE\tEX_SYMBOL\tBCKGRNDPAG\tLOT_SIZE_A\tDSPLY_NMLL\tGV1_FLAG\tISS_TP_FLG\tRDM_CUR	MATUR_DATE\tSTRIKE_PRC\tWNT_RATIO\t#INSTMOD_MNEMONIC\tBCAST_REF\t#INSTMOD_LOT_SIZE_X\t#INSTMOD_SPARE_UBYTE3	EXL_NAME\t#INSTMOD_LONGLINK3\t#INSTMOD_GN_TX20_3\t#INSTMOD_GN_TX20_6\t#INSTMOD_GN_TX20_7\t#INSTMOD_GN_TX20_10\t#INSTMOD_GN_TX20_11\t#INSTMOD_GN_TX20_12\t#INSTMOD_BOND_TYPE\t#INSTMOD_LONGLINK2\t#INSTMOD_SPARE_SNUM13\t#INSTMOD_GEN_VAL4\t#INSTMOD_TDN_SYMBOL";
            for (int i = 0; i < ricList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(ricList[i].ricCodeStr + ".HK");
                sb.Append("\t");
                sb.Append(ricList[i].ricNameStr);
                sb.Append("\t");
                sb.Append(ricList[i].ricCodeStr + ".HK");
                sb.Append("\t");
                sb.Append(ricList[i].ricCodeStr);
                sb.Append("\t");
                sb.Append(ricList[i].ricCodeStr);
                sb.Append("\t");
                sb.Append("****");
                sb.Append("\t");
                sb.Append(ricList[i].boardLotStr.Replace(",", ""));
                sb.Append("\t");

                //For Bull
                if (ricList[i].bullBearStr == "Bull")
                {

                    sb.Append(chineseList[i].ricCHNNameStr.Substring(0, 4).Replace("恒", "恆")
                       + chineseList[i].ricCHNNameStr.Substring(chineseList[i].ricCHNNameStr.Length - 1)
                       + ricList[i].maturityDateDT.Month + "月" + "RC" + ricList[i].maturityDateDT.ToString("yy"));

                }
                //For BEAR
                else
                {

                    sb.Append(chineseList[i].ricCHNNameStr.Substring(0, 4).Replace("恒", "恆")
                       + chineseList[i].ricCHNNameStr.Substring(chineseList[i].ricCHNNameStr.Length - 1)
                       + ricList[i].maturityDateDT.Month + "月" + "RP" + ricList[i].maturityDateDT.ToString("yy"));

                }

                sb.Append("\t");

                if (ricList[i].bullBearStr == "Bull")
                {
                    sb.Append("C");
                }
                //For Bear
                else
                {
                    sb.Append("P");
                }
                sb.Append("\t");

                if (Char.IsLetter(ricList[i].underlyingStr, 0))
                {
                    sb.Append("I");
                }
                //For Equity
                else
                {
                    sb.Append("S");
                }
                sb.Append("\t");
                sb.Append("344");
                sb.Append("\t");
                sb.Append(ricList[i].maturityDateDT.ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                sb.Append("\t");
                //For HKD
                if (Char.IsLetter(ricList[i].strikeLevelStr, 0))
                {
                    sb.Append(ricList[i].strikeLevelStr.Substring(4));
                }
                else
                {
                    sb.Append(ricList[i].strikeLevelStr);
                }
                sb.Append("\t");

                //================warrant ratio
                sb.Append((1.0 / Convert.ToInt32(ricList[i].entitlementRatioStr)).ToString("0.0000000"));
                sb.Append("\t");
                sb.Append(ricList[i].ricCodeStr);
                sb.Append("\t");

                //For Index
                if (Char.IsLetter(ricList[i].underlyingStr, 0))
                {
                    if (ricList[i].underlyingStr == "HSCEI")
                    {
                        sb.Append(".HSCE");
                    }
                    else
                    {
                        sb.Append("." + ricList[i].underlyingStr);
                    }
                }
                //For Equity
                else
                {
                    sb.Append(ricList[i].underlyingStr.Substring(1) + ".HK");
                }
                sb.Append("\t");
                sb.Append(ricList[i].boardLotStr.Replace(",", ""));
                sb.Append("\t");
                sb.Append("14");
                sb.Append("\t");
                sb.Append("HKG_EQLB_CBBC");

                sb.Append("\t");

                sb.Append("t" + ricList[i].ricCodeStr + ".HK\t");//Append GN_TX20_3

                sb.Append("[HK-\"WARRAN*\"]\t");//Append GN_TX20_6

                //For Index
                if (Char.IsLetter(ricList[i].underlyingStr, 0))
                {
                    if (ricList[i].underlyingStr == "HSCEI")
                    {
                        sb.Append("<.HSCE>\t");
                    }
                    else
                    {
                        sb.Append("<." + ricList[i].underlyingStr + ">\t");
                    }
                }
                //For Equity
                else
                {
                    sb.Append("<" + ricList[i].underlyingStr.Substring(1) + ".HK>\t");
                }
                //Append GN_TX20_7, include 10 spaces after |
                sb.Append("********************\t|          ");
                //Append GN_TX20_10
                sb.Append("CBBC/" + ricList[i].bullBearStr.ToUpper() + "\t");
                //Append GN_TX20_11
                sb.Append(ricList[i].issueSizeStr.Replace(",", "") + "\t");

                //Append GN_TX20_12 (Misc.Info)
                //For Index
                if (Char.IsLetter(ricList[i].underlyingStr, 0))
                {
                    sb.Append("IDX     <" + ricList[i].ricCodeStr + "MI.HK>\t");
                }
                //For Equity
                else
                {
                    sb.Append("HKD     <" + ricList[i].ricCodeStr + "MI.HK>\t");
                }
                //Append BOND_TYPE
                sb.Append("WARRANTS\t");

                //Append LONGLINK2(equal LONGLINK14 in FM)
                //For Index
                if (Char.IsLetter(ricList[i].underlyingStr, 0))
                {
                    sb.Append("\t");
                }
                //For Equity
                else
                {
                    sb.Append(ricList[i].underlyingStr.Substring(1) + ".HK\t");
                }
                //Append SPARE_SNUM13
                sb.Append("1\t");

                //Append GEN_VAL4(equal Call Level in FM)
                //For HKD
                if (Char.IsLetter(ricList[i].strikeLevelStr, 0))
                {

                    sb.Append(ricList[i].callLevelStr.Substring(4));
                }
                else
                {
                    sb.Append(ricList[i].callLevelStr);
                }
                sb.Append("\t" + ricList[i].ricCodeStr);
                //content[i] = sb.ToString();      /*without the header*/
                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);

            }

            WriteTxtFile(fullpath, content);
        }

        /**
         * Generate HKG_EQLB.txt file (for Main RICs of warrants)
         * Return   : void
         * Parameter: string filePath, List<HKRicTemplate> ricList, List<HKRicTemplate> chineseList
         */
        public void Generate_HKG_EQLB
            (string filePath, List<HKRicTemplate> ricList, List<HKRicTemplate> chineseList)
        {
            string fullpath = filePath + "\\HKG_EQLB.txt";
            //string[] content = new string[ricList.Count];
            string[] content = new string[ricList.Count + 1];
            content[0] = "SYMBOL\tDSPLY_NAME\tRIC\tOFFCL_CODE\tEX_SYMBOL\tBCKGRNDPAG\tLOT_SIZE_A\tDSPLY_NMLL\tGV1_FLAG\tISS_TP_FLG\tRDM_CUR	MATUR_DATE\tSTRIKE_PRC\tWNT_RATIO\t#INSTMOD_BOND_TYPE\t#INSTMOD_MNEMONIC\tBCAST_REF\t#INSTMOD_LOT_SIZE_X\t#INSTMOD_SPARE_UBYTE3	EXL_NAME\tBCU\t#INSTMOD_LONGLINK3\t#INSTMOD_GN_TX20_3\t#INSTMOD_GN_TX20_6\t#INSTMOD_GN_TX20_7\t#INSTMOD_GN_TX20_10\t#INSTMOD_GN_TX20_11\t#INSTMOD_GN_TX20_12\t#INSTMOD_LONGLINK2\t#INSTMOD_SPARE_SNUM13\t#INSTMOD_TDN_SYMBOL";
            //content[0] = "";

            for (int i = 0; i < ricList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(ricList[i].ricCodeStr + ".HK");
                sb.Append("\t");
                sb.Append(ricList[i].ricNameStr);
                sb.Append("\t");
                sb.Append(ricList[i].ricCodeStr + ".HK");
                sb.Append("\t");
                sb.Append(ricList[i].ricCodeStr);
                sb.Append("\t");
                sb.Append(ricList[i].ricCodeStr);
                sb.Append("\t");
                sb.Append("****");
                sb.Append("\t");
                sb.Append(ricList[i].boardLotStr.Replace(",", ""));
                sb.Append("\t");

                //For Call
                if (ricList[i].callPutStr == "Call")
                {
                    //The last char of warrant name is a character
                    if (Char.IsLetter(ricList[i].ricNameStr, ricList[i].ricNameStr.Length - 1))
                    {
                        sb.Append(chineseList[i].ricCHNNameStr.Substring(0, 4).Replace("恒", "恆")
                           + ricList[i].ricNameStr.Substring(ricList[i].ricNameStr.Length - 1)
                           + ricList[i].maturityDateDT.Month + "月" + "CW" + ricList[i].maturityDateDT.ToString("yy"));
                    }
                    else
                    {
                        sb.Append(chineseList[i].ricCHNNameStr.Substring(0, 4).Replace("恒", "恆")
                           + ricList[i].maturityDateDT.Month + "月" + "CW" + ricList[i].maturityDateDT.ToString("yy"));
                    }

                }
                //For Put
                else
                {
                    if (Char.IsLetter(ricList[i].ricNameStr, ricList[i].ricNameStr.Length - 1))
                    {
                        sb.Append(chineseList[i].ricCHNNameStr.Substring(0, 4).Replace("恒", "恆")
                           + ricList[i].ricNameStr.Substring(ricList[i].ricNameStr.Length - 1)
                           + ricList[i].maturityDateDT.Month + "月" + "PW" + ricList[i].maturityDateDT.ToString("yy"));

                    }
                    else
                    {

                        sb.Append(chineseList[i].ricCHNNameStr.Substring(0, 4).Replace("恒", "恆")
                           + ricList[i].maturityDateDT.Month + "月" + "PW" + ricList[i].maturityDateDT.ToString("yy"));
                    }

                }

                sb.Append("\t");

                //For Call
                if (ricList[i].callPutStr == "Call")
                {
                    sb.Append("C");
                }
                //For Put
                else
                {
                    sb.Append("P");
                }
                sb.Append("\t");

                //For Index
                if (ricList[i].underlyingStr == "HSI" || ricList[i].underlyingStr == "HSCEI" || ricList[i].underlyingStr == "DJI")
                {
                    sb.Append("I");
                }
                //For Stock
                else if (Char.IsDigit(ricList[i].underlyingStr, 0))
                {
                    sb.Append("S");
                }
                else//For Forex ricList[i].ricNameStr.Contains("EURUS/AUDUS/USYEN")
                {
                    sb.Append("F");
                }
                sb.Append("\t");
                sb.Append("344");
                sb.Append("\t");
                sb.Append(ricList[i].maturityDateDT.ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                sb.Append("\t");
                //For HKD
                if (Char.IsLetter(ricList[i].strikeLevelStr, 0))
                {
                    sb.Append(Convert.ToDouble(ricList[i].strikeLevelStr.Substring(4)).ToString("0.000"));
                }
                else
                {
                    sb.Append(Convert.ToDouble(ricList[i].strikeLevelStr).ToString("0.000"));
                }
                sb.Append("\t");
                double enRatio = 1.0 / Convert.ToDouble(ricList[i].entitlementRatioStr);
                string x = enRatio.ToString();
                String entitlemenRatioStr = enRatio.ToString().Length >= 9 == true ? enRatio.ToString("0.0000000") : enRatio.ToString();
                sb.Append(entitlemenRatioStr);
                //sb.Append((1.0 / Convert.ToDouble(ricList[i].entitlementRatioStr)).ToString().Substring(0, 9));
                //sb.Append((1.0 / Convert.ToDouble(ricList[i].entitlementRatioStr)).ToString("0.0000000"));
                sb.Append("\t");
                sb.Append("WARRANTS");
                sb.Append("\t");
                sb.Append(ricList[i].ricCodeStr);
                sb.Append("\t");

                //For Index
                if (Char.IsLetter(ricList[i].underlyingStr, 0))
                {
                    if (ricList[i].underlyingStr == "HSCEI")
                    {
                        sb.Append(".HSCE");
                    }
                    else
                    {
                        sb.Append("." + ricList[i].underlyingStr);
                    }
                }
                //For Equity
                else
                {
                    sb.Append(ricList[i].underlyingStr.Substring(1) + ".HK");
                }

                sb.Append("\t");
                sb.Append(ricList[i].boardLotStr.Replace(",", ""));
                sb.Append("\t");

                //For Call
                if (ricList[i].callPutStr == "Call")
                {
                    sb.Append("5");
                }
                //For Put
                else
                {
                    sb.Append("6");
                }

                sb.Append("\t");
                sb.Append("HKG_EQLB");
                sb.Append("\t");
                //For stock
                if (Char.IsDigit(ricList[i].underlyingStr, 0))
                {
                    sb.Append("HKG_EQ_CWRTS");
                }

                sb.Append("\t");

                sb.Append("t" + ricList[i].ricCodeStr + ".HK\t");//Append GN_TX20_3

                sb.Append("[HK-\"WARRAN*\"]\t");//Append GN_TX20_6

                //for Equity stock
                if (Char.IsDigit(ricList[i].underlyingStr, 0))
                {
                    sb.Append("<" + ricList[i].underlyingStr.Substring(1) + ".HK>\t");
                }
                //for Index
                else if (ricList[i].underlyingStr == "HSI" || ricList[i].underlyingStr == "HSCEI" || ricList[i].underlyingStr == "DJI")
                {
                    if (ricList[i].underlyingStr == "HSCEI")
                    {
                        sb.Append("<.HSCE>\t");
                    }
                    else
                    {
                        sb.Append("<." + ricList[i].underlyingStr + ">\t");
                    }
                }
                //for others
                else
                {
                    sb.Append("\t");
                }
                //Append GN_TX20_7, include 12 spaces after |
                //for Equity stock
                if (Char.IsDigit(ricList[i].underlyingStr, 0))
                {
                    sb.Append("<" + ricList[i].underlyingStr.Substring(1) + "DIVCF.HK>\t|            ");
                }
                else
                {
                    sb.Append("********************\t|            ");
                }

                //Append GN_TX20_10
                sb.Append("EU/" + ricList[i].callPutStr.ToUpper() + "\t");
                //Append GN_TX20_11
                sb.Append(ricList[i].issueSizeStr.Replace(",", "") + "\t");

                //Append GN_TX20_12 (Misc.Info)
                //For Index
                if (Char.IsLetter(ricList[i].underlyingStr, 0))
                {
                    sb.Append("IDX     <" + ricList[i].ricCodeStr + "MI.HK>\t");
                }
                //For Equity
                else if (Char.IsDigit(ricList[i].underlyingStr, 0))
                {
                    sb.Append("HKD     <" + ricList[i].ricCodeStr + "MI.HK>\t");
                }
                else
                {
                    sb.Append("              \t");
                }
                //Append BOND_TYPE
                //sb.Append("WARRANTS\t");

                //Append LONGLINK2(equal LONGLINK14 in FM)
                //For Index
                if (Char.IsLetter(ricList[i].underlyingStr, 0))
                {
                    if (ricList[i].underlyingStr == "HSI")
                    {
                        sb.Append(".HSI|HKD|1\t");
                    }
                    else if (ricList[i].underlyingStr == "HSCEI")
                    {
                        sb.Append(".HSCE|HKD|1\t");
                    }
                    else if (ricList[i].underlyingStr == "DJI")
                    {
                        sb.Append(".DJI|USD|1\t");
                    }
                    else
                    {
                        sb.Append("       \t");
                    }

                }
                //For Stock
                else if (Char.IsDigit(ricList[i].underlyingStr, 0))
                {
                    sb.Append(ricList[i].underlyingStr.Substring(1) + ".HK\t");
                }
                else
                {
                    sb.Append("       \t");
                }
                //Append SPARE_SNUM13
                sb.Append("1\t");
                sb.Append(ricList[i].ricCodeStr);
                //content[i] = sb.ToString(); /*without the header*/
                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);

            }

            WriteTxtFile(fullpath, content);
        }

        /**
         * Generate HKG_EQLBMI.txt file (for MI RICs of Warrants&CBBC)
         * Return   : void
         * Parameter: string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList
         */
        public void Generate_HKG_EQLBMI
            (string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList)
        {
            string fullpath = filePath + "\\HKG_EQLBMI.txt";
            //string[] content = new string[cbbcList.Count + warrantList.Count];
            string[] content = new string[cbbcList.Count + warrantList.Count + 1];
            //content[0] = "SYMBOL\tDSPLY_NAME\tRIC\tEX_SYMBOL\t#INSTMOD_ROW80_1\tEXL_NAME\t#INSTMOD_ROW80_3\t#INSTMOD_ROW80_4\t#INSTMOD_ROW80_5\t#INSTMOD_ROW80_6\t#INSTMOD_ROW80_7\t#INSTMOD_ROW80_8\t#INSTMOD_ROW80_9\t#INSTMOD_ROW80_10\t#INSTMOD_ROW80_11\t#INSTMOD_ROW80_12\t#INSTMOD_ROW80_13\t#INSTMOD_ROW80_14\t#INSTMOD_ROW80_15\t#INSTMOD_ROW80_16\t#INSTMOD_ROW80_17\t#INSTMOD_ROW80_18\t#INSTMOD_ROW80_19\t#INSTMOD_ROW80_20\t#INSTMOD_ROW80_21";
            content[0] = "SYMBOL\tDSPLY_NAME\tRIC\tEX_SYMBOL\t#INSTMOD_ROW80_1\tEXL_NAME";

            for (int i = 0; i < cbbcList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(cbbcList[i].ricCodeStr + "MI.HK");
                sb.Append("\t");
                sb.Append(cbbcList[i].ricNameStr);
                sb.Append("\t");
                sb.Append(cbbcList[i].ricCodeStr + "MI.HK");
                sb.Append("\t");
                sb.Append(cbbcList[i].ricCodeStr + "MI");
                sb.Append("\t");
                //include 36 spaces at end
                sb.Append("Security Miscellaneous Information                                    ");
                sb.Append(cbbcList[i].ricCodeStr + "MI.HK");
                sb.Append("\t");
                sb.Append("HKG_EQLB_MI_PAGE");
                sb.Append("\t");
                ////new begin(Modified on 6-2)
                //sb.Append("ISIN                                    EIPO Start Date\t");
                //sb.Append("Instrument Type                         EIPO End Date\t");
                //sb.Append("Market                                  EIPO Start Time\t");
                //sb.Append("Sub-Market                              EIPO End Time\t");
                //sb.Append("Listing Date                            EIPO Price\t");
                //sb.Append("De-listing Date                         Spread Table\t");
                //sb.Append("Listing Status                          Shortselling Stock\t");
                //sb.Append("Trading Status                          Intra-day Shortselling Stock\t");
                //sb.Append("Stamp Duty                              Automatch Stock\t");
                //sb.Append("Test Stock                              CCASS Stock\t");
                //sb.Append("Dummy Stock\t");
                //sb.Append("--------------------------------------------------------------------------------\t");
                //sb.Append("Trading Start Time                      Trading End Time\t");
                //sb.Append("Session Type");
                ////new end
                //content[i] = sb.ToString();   /*without the header*/
                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            //int j = cbbcList.Count;          /*without the header*/
            int j = cbbcList.Count + 1;
            for (int i = 0; i < warrantList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(warrantList[i].ricCodeStr + "MI.HK");
                sb.Append("\t");
                sb.Append(warrantList[i].ricNameStr);
                sb.Append("\t");
                sb.Append(warrantList[i].ricCodeStr + "MI.HK");
                sb.Append("\t");
                sb.Append(warrantList[i].ricCodeStr + "MI");
                sb.Append("\t");
                //include 36 spaces at end
                sb.Append("Security Miscellaneous Information                                    ");
                sb.Append(warrantList[i].ricCodeStr + "MI.HK");
                sb.Append("\t");
                sb.Append("HKG_EQLB_MI_PAGE");
                sb.Append("\t");
                ////new begin
                //sb.Append("ISIN                                    EIPO Start Date\t");
                //sb.Append("Instrument Type                         EIPO End Date\t");
                //sb.Append("Market                                  EIPO Start Time\t");
                //sb.Append("Sub-Market                              EIPO End Time\t");
                //sb.Append("Listing Date                            EIPO Price\t");
                //sb.Append("De-listing Date                         Spread Table\t");
                //sb.Append("Listing Status                          Shortselling Stock\t");
                //sb.Append("Trading Status                          Intra-day Shortselling Stock\t");
                //sb.Append("Stamp Duty                              Automatch Stock\t");
                //sb.Append("Test Stock                              CCASS Stock\t");
                //sb.Append("Dummy Stock\t");
                //sb.Append("--------------------------------------------------------------------------------\t");
                //sb.Append("Trading Start Time                      Trading End Time\t");
                //sb.Append("Session Type");
                ////new end
                content[j++] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            WriteTxtFile(fullpath, content);
        }

        /**
        * Generate MI.txt file (for MI RICs of Warrants&CBBC)
        * Return   : void
        * Parameter: string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList
        */
        public void Generate_MI(string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList)
        {
            string fullpath = filePath + "\\MI.txt";
            string[] content = new string[cbbcList.Count + warrantList.Count + 1];
            //content[0] = "ROW80_3\tROW80_4\tROW80_5\tROW80_6\tROW80_7\tROW80_8\tROW80_9\tROW80_10\tROW80_11\tROW80_12\tROW80_13\tROW80_14\tROW80_15\tROW80_16\tROW80_17\tROW80_18\tROW80_19\tROW80_20\tROW80_21";
            content[0] = "HKSE.DAT;07-APR-2004 09:00:00;TPS;\r\nROW80_3;ROW80_4;ROW80_5;ROW80_6;ROW80_7;ROW80_8;ROW80_9;ROW80_10;ROW80_11;ROW80_12;ROW80_13;ROW80_14;ROW80_15;ROW80_16;ROW80_17;ROW80_18;ROW80_19;ROW80_20;ROW80_21;";
            for (int i = 0; i < cbbcList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(cbbcList[i].ricCodeStr + "MI.HK;");
                sb.Append("ISIN                                    EIPO Start Date;");
                sb.Append("Instrument Type                         EIPO End Date;");
                sb.Append("Market                                  EIPO Start Time;");
                sb.Append("Sub-Market                              EIPO End Time;");
                sb.Append("Listing Date                            EIPO Price;");
                sb.Append("De-listing Date                         Spread Table;");
                sb.Append("Listing Status                          Shortselling Stock;");
                sb.Append("Trading Status                          Intra-day Shortselling Stock;");
                sb.Append("Stamp Duty                              Automatch Stock;");
                sb.Append("Test Stock                              CCASS Stock;");
                sb.Append("Dummy Stock;");
                sb.Append("--------------------------------------------------------------------------------;");
                sb.Append("Trading Start Time                      Trading End Time;");
                sb.Append("Session Type;;;;;;");
                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);
            }
            int j = cbbcList.Count + 1;
            for (int i = 0; i < warrantList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(warrantList[i].ricCodeStr + "MI.HK;");
                sb.Append("ISIN                                    EIPO Start Date;");
                sb.Append("Instrument Type                         EIPO End Date;");
                sb.Append("Market                                  EIPO Start Time;");
                sb.Append("Sub-Market                              EIPO End Time;");
                sb.Append("Listing Date                            EIPO Price;");
                sb.Append("De-listing Date                         Spread Table;");
                sb.Append("Listing Status                          Shortselling Stock;");
                sb.Append("Trading Status                          Intra-day Shortselling Stock;");
                sb.Append("Stamp Duty                              Automatch Stock;");
                sb.Append("Test Stock                              CCASS Stock;");
                sb.Append("Dummy Stock;");
                sb.Append("--------------------------------------------------------------------------------;");
                sb.Append("Trading Start Time                      Trading End Time;");
                sb.Append("Session Type;;;;;;");
                content[j++] = sb.ToString();
                sb.Remove(0, sb.Length);
            }
            WriteTxtFile(fullpath, content);
        }

        /**
        * Generate BK.txt file (for bk RICs of warrants & CBBC)
        * Return   : void
        * Parameter: string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList
        */
        public void Generate_BK
            (string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList)
        {
            string fullpath = filePath + "\\BK.txt";
            string[] content = new string[cbbcList.Count + warrantList.Count + 1];
            content[0] = "RIC\t#INSTMOD_ROW80_1\t#INSTMOD_ROW80_2\t#INSTMOD_ROW80_13";


            for (int i = 0; i < cbbcList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(cbbcList[i].ricCodeStr + "bk.HK\t");

                sb.Append(cbbcList[i].ricNameStr);

                string startPosition = (17 - cbbcList[i].ricNameStr.Length).ToString();
                switch (startPosition)
                {
                    case "0": sb.Append(""); break;
                    case "1": sb.Append(" "); break;
                    case "2": sb.Append("  "); break;
                    case "3": sb.Append("   "); break;
                }


                sb.Append("<" + cbbcList[i].ricCodeStr + ".HK> <HKBK01>                                  " + cbbcList[i].ricCodeStr + "bk.HK\t");


                sb.Append("        BID                    ASK\t");

                if (cbbcList[i].bullBearStr == "Bull")
                {
                    sb.Append("Callable Bull Contracts");
                }
                else
                {
                    sb.Append("Callable Bear Contracts");
                }

                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);
            }
            int j = cbbcList.Count + 1;
            for (int i = 0; i < warrantList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(warrantList[i].ricCodeStr + "bk.HK\t");
                sb.Append(warrantList[i].ricNameStr);

                string startPosition = (17 - warrantList[i].ricNameStr.Length).ToString();
                switch (startPosition)
                {
                    case "0": sb.Append(""); break;
                    case "1": sb.Append(" "); break;
                    case "2": sb.Append("  "); break;
                    case "3": sb.Append("   "); break;
                }

                sb.Append("<" + warrantList[i].ricCodeStr + ".HK> <HKBK01>                                  " + warrantList[i].ricCodeStr + "bk.HK\t");


                sb.Append("        BID                    ASK\t");

                //For index
                if (Char.IsLetter(warrantList[i].underlyingStr, 0))
                {
                    sb.Append("Warrant Type-Index Warrant");
                }
                else if (Char.IsDigit(warrantList[i].underlyingStr, 0))//For stock
                {
                    sb.Append("Warrant Type-Equity Warrant");
                }
                else
                {
                    sb.Append("");
                }
                content[j++] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            WriteTxtFile(fullpath, content);
        }

        /**
         * WriteTxtFile method, generate text file
         * Retrun   : void
         * Parameter: string fullpath, string [] content
         */
        private void WriteTxtFile
            (string fullpath, string[] content)
        {
            try
            {
                //StreamReader sr = new StreamReader(fullpath);
                //string lineText = sr.ReadToEnd();
                //sr.Close();
                File.WriteAllLines(fullpath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                string errInfo = ex.ToString();
            }
        }

        /**
         * Generate MAIN.txt file (for Main RICs of warrants & CBBC)
         * Return   : void
         * Parameter: string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList
         */
        /*public void Generate_MAIN(string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList)
        {
            string fullpath = filePath + "\\MAIN.txt";
            string[] content = new string[cbbcList.Count + warrantList.Count + 1];
            content[0] = "RIC\tLONGLINK3\tGN_TX20_3\tGN_TX20_6\tGN_TX20_7\tGN_TX20_10\tGN_TX20_11\tGN_TX20_12\tBOND_TYPE\tLONGLINK2\tSPARE_SNUM13\tGEN_VAL4";

            //Get CBBC RIC
            for (int i = 0; i < cbbcList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(cbbcList[i].ricCodeStr + ".HK\t");
                //Append LONGLINK3
                sb.Append("t" + cbbcList[i].ricCodeStr + ".HK\t");
                //Append GN_TX20_3
                sb.Append("[HK-\"WARRAN*\"]\t");
                //Append GN_TX20_6
                //For Index
                if (Char.IsLetter(cbbcList[i].underlyingStr, 0))
                {
                    if (cbbcList[i].underlyingStr == "HSCEI")
                    {
                        sb.Append("<.HSCE>\t");
                    }
                    else
                    {
                        sb.Append("<." + cbbcList[i].underlyingStr + ">\t");
                    }
                }
                //For Equity
                else
                {
                    sb.Append("<" + cbbcList[i].underlyingStr.Substring(1) + ".HK>\t");
                }
                //Append GN_TX20_7, include 10 spaces after |
                sb.Append("********************\t|          ");
                //Append GN_TX20_10
                sb.Append("CBBC/" + cbbcList[i].bullBearStr.ToUpper() + "\t");
                //Append GN_TX20_11
                sb.Append(cbbcList[i].issueSizeStr + "\t");

                //Append GN_TX20_12 (Misc.Info)
                //For Index
                if (Char.IsLetter(cbbcList[i].underlyingStr, 0))
                {
                    sb.Append("IDX     <" + cbbcList[i].ricCodeStr + "MI.HK>\t");
                }
                //For Equity
                else
                {
                    sb.Append("HKD     <" + cbbcList[i].ricCodeStr + "MI.HK>\t");
                }
                //Append BOND_TYPE
                sb.Append("WARRANTS\t");

                //Append LONGLINK2(equal LONGLINK14 in FM)
                //For Index
                if (Char.IsLetter(cbbcList[i].underlyingStr, 0))
                {
                    sb.Append("\t");
                }
                //For Equity
                else
                {
                    sb.Append(cbbcList[i].underlyingStr.Substring(1) + ".HK\t");
                }
                //Append SPARE_SNUM13
                sb.Append("1\t");

                //Append GEN_VAL4(equal Call Level in FM)
                //For HKD
                if (Char.IsLetter(cbbcList[i].strikeLevelStr, 0))
                {

                    sb.Append(cbbcList[i].callLevelStr.Substring(4));
                }
                else
                {
                    sb.Append(cbbcList[i].callLevelStr);
                }

                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);
            }
            int j = cbbcList.Count + 1;

            //Get Warrants Ric
            for (int i = 0; i < warrantList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(warrantList[i].ricCodeStr + ".HK\t");
                //Append LONGLINK3
                sb.Append("t" + warrantList[i].ricCodeStr + ".HK\t");
                //Append GN_TX20_3
                sb.Append("[HK-\"WARRAN*\"]\t");
                //Append GN_TX20_6
                //for Equity stock
                if (Char.IsDigit(warrantList[i].underlyingStr, 0))
                {
                    sb.Append("<" + warrantList[i].underlyingStr.Substring(1) + ".HK>\t");
                }
                //for Index
                else if (warrantList[i].underlyingStr == "HSI" || warrantList[i].underlyingStr == "HSCEI" || warrantList[i].underlyingStr == "DJI")
                {
                    if (warrantList[i].underlyingStr == "HSCEI")
                    {
                        sb.Append("<.HSCE>\t");
                    }
                    else
                    {
                        sb.Append("<." + warrantList[i].underlyingStr + ">\t");
                    }
                }
                //for others
                else
                {
                    sb.Append("\t");
                }
                //Append GN_TX20_7, include 12 spaces after |
                //for Equity stock
                if (Char.IsDigit(warrantList[i].underlyingStr, 0))
                {
                    sb.Append("<" + warrantList[i].underlyingStr.Substring(1) + "DIVCF.HK>\t|            ");
                }
                else
                {
                    sb.Append("********************\t|            ");
                }

                //Append GN_TX20_10
                sb.Append("EU/" + warrantList[i].callPutStr.ToUpper() + "\t");
                //Append GN_TX20_11
                sb.Append(warrantList[i].issueSizeStr + "\t");

                //Append GN_TX20_12 (Misc.Info)
                //For Index
                if (Char.IsLetter(warrantList[i].underlyingStr, 0))
                {
                    sb.Append("IDX     <" + warrantList[i].ricCodeStr + "MI.HK>\t");
                }
                //For Equity
                else if (Char.IsDigit(warrantList[i].underlyingStr, 0))
                {
                    sb.Append("HKD     <" + warrantList[i].ricCodeStr + "MI.HK>\t");
                }
                else
                {
                    sb.Append("              \t");
                }
                //Append BOND_TYPE
                sb.Append("WARRANTS\t");

                //Append LONGLINK2(equal LONGLINK14 in FM)
                //For Index
                if (Char.IsLetter(warrantList[i].underlyingStr, 0))
                {
                    if (warrantList[i].underlyingStr == "HSI")
                    {
                        sb.Append(".HSI|HKD|1\t");
                    }
                    else if (warrantList[i].underlyingStr == "HSCEI")
                    {
                        sb.Append(".HSCE|HKD|1\t");
                    }
                    else if (warrantList[i].underlyingStr == "DJI")
                    {
                        sb.Append(".DJI|USD|1\t");
                    }
                    else
                    {
                        sb.Append("       \t");
                    }

                }
                //For Stock
                else if (Char.IsDigit(warrantList[i].underlyingStr, 0))
                {
                    sb.Append(warrantList[i].underlyingStr.Substring(1) + ".HK\t");
                }
                else
                {
                    sb.Append("       \t");
                }
                //Append SPARE_SNUM13
                sb.Append("1");

                content[j++] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            WriteTxtFile(fullpath, content);
        }*/

        /**
         * Generate MI.txt file (for MI RICs of Warrants&CBBC)
         * Return   : void
         * Parameter: string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList
         */
        /*public void Generate_MI(string filePath, List<HKRicTemplate> cbbcList, List<HKRicTemplate> warrantList)
        {
            string fullpath = filePath + "\\MI.txt";
            string[] content = new string[cbbcList.Count + warrantList.Count + 1];
            content[0] = "RIC\tROW80_3\tROW80_4\tROW80_5\tROW80_6\tROW80_7\tROW80_8\tROW80_9\tROW80_10\tROW80_11\tROW80_12\tROW80_13\tROW80_14\tROW80_15\tROW80_16\tROW80_17\tROW80_18\tROW80_19\tROW80_20\tROW80_21";
            for (int i = 0; i < cbbcList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(cbbcList[i].ricCodeStr + "MI.HK\t");
                sb.Append("ISIN                                    EIPO Start Date\t");
                sb.Append("Instrument Type                         EIPO End Date\t");
                sb.Append("Market                                  EIPO Start Time\t");
                sb.Append("Sub-Market                              EIPO End Time\t");
                sb.Append("Listing Date                            EIPO Price\t");
                sb.Append("De-listing Date                         Spread Table\t");
                sb.Append("Listing Status                          Shortselling Stock\t");
                sb.Append("Trading Status                          Intra-day Shortselling Stock\t");
                sb.Append("Stamp Duty                              Automatch Stock\t");
                sb.Append("Test Stock                              CCASS Stock\t");
                sb.Append("Dummy Stock\t");
                sb.Append("--------------------------------------------------------------------------------\t");
                sb.Append("Trading Start Time                      Trading End Time\t");
                sb.Append("Session Type");
                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);
            }
            int j = cbbcList.Count + 1;
            for (int i = 0; i < warrantList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(warrantList[i].ricCodeStr + "MI.HK\t");
                sb.Append("ISIN                                    EIPO Start Date\t");
                sb.Append("Instrument Type                         EIPO End Date\t");
                sb.Append("Market                                  EIPO Start Time\t");
                sb.Append("Sub-Market                              EIPO End Time\t");
                sb.Append("Listing Date                            EIPO Price\t");
                sb.Append("De-listing Date                         Spread Table\t");
                sb.Append("Listing Status                          Shortselling Stock\t");
                sb.Append("Trading Status                          Intra-day Shortselling Stock\t");
                sb.Append("Stamp Duty                              Automatch Stock\t");
                sb.Append("Test Stock                              CCASS Stock\t");
                sb.Append("Dummy Stock\t");
                sb.Append("--------------------------------------------------------------------------------\t");
                sb.Append("Trading Start Time                      Trading End Time\t");
                sb.Append("Session Type");
                content[j++] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            WriteTxtFile(fullpath, content);
        }*/
    }
}
