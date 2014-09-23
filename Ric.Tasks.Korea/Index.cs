using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.ComponentModel;
using System.Drawing;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
{
    #region used for KOREA_IndexGeneratorConfig
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_Index_ReadFile_CONFIG
    {
        public String WORKBOOK_PATH { get; set; }
        public String KOSPI_SHEETNAME { get; set; }
        public String KOSDAQ_SHEETNAME { get; set; }
        public String KRX_SHEETNAME { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_index_OrignalFile_CONFIG
    {
        public String WORKBOOK_PATH { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_Index_GeneratorFile_CONFIG
    {
        public String WORKBOOK_PATH { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_IndexGeneratorConfig
    {
        public String LOG_FILE_PATH { get; set; }
        public Korea_Index_ReadFile_CONFIG Korea_Index_ReadFile_CONFIG { get; set; }
        public Korea_KQorKSList_ReadFilePath_CONFIG Korea_KQorKSList_ReadFilePath_CONFIG { get; set; }
        public Korea_index_OrignalFile_CONFIG Korea_index_OrignalFile_CONFIG { get; set; }
        public Korea_Index_GeneratorFile_CONFIG Korea_Index_GeneratorFile_CONFIG { get; set; }
    }
    #endregion

    #region  this is the model for Index

    /// <summary>
    /// SheetName being used for save the everyone's sheet's name
    /// Ric,DisplayName,LocalLanguageName,FreeFloatingRate four item used for the common instrument which doesn't have the Add Items and Drop Items
    /// SheetName + ChainRic used for every sheet's content's title 
    /// IsAddItem ,IsDropItem ,AddItem ,DropItem used for which sheet contains Add Items or Drop Items
    /// </summary>
    class IndexTemplate
    {
        public List<FloatRateIndex> FRIList { get; set; }
        public List<LocalLanguageNameIndex> LLNIList { get; set; }
        public String WarrantKind { get; set; }
        public String SheetName { get; set; }
        public bool IsAddItem { get; set; }
        public bool IsDropItem { get; set; }
        public String ChainRIC { get; set; }
        public List<IndexAddItem> AddItemList { get; set; }
        public List<IndexDropItem> DropItemList { get; set; }
        public String SheetTitle { get; set; }
        public List<SectorIndex> SIList { get; set; }
    }

    class SectorIndex
    {
        public String SectorName { get; set; }
        public int Count { get; set; }
    }

    class LocalLanguageNameIndex
    {
        public String Ric { get; set; }
        public String DisplayName { get; set; }
        public String LocalLanguageName { get; set; }
    }

    class FloatRateIndex
    {
        public String Ric { get; set; }
        public String DisplayName { get; set; }
        public String LocalLanguageName { get; set; }
        public String FreeFloatingRate { get; set; }
    }

    class IndexAddItem
    {
        public String Ric { get; set; }
        public String LocalLanguageName { get; set; }
        public String DisplayName { get; set; }
    }

    class IndexDropItem
    {
        public String Ric { get; set; }
        public String LocalLanguageName { get; set; }
        public String DisplayName { get; set; }
    }

    class IndexListingList
    {
        public String KoreanNameForChainRIC { get; set; }
        public String EnglishNameForChainRIC { get; set; }
        public String IndexChainRIC { get; set; }
        public String Number { get; set; }
        public String One { get; set; }
        public String Two { get; set; }
        public String Three { get; set; }
    }
    #endregion

    public class Index : GeneratorBase
    {
        const string LOGFILE_NAME = "Index-Log.txt";
        private static readonly string CONFIG_FILE_PATH = ".\\Config\\Korea\\KOREA_IndexGenerator.config";
        private Hashtable kskqlistingHash = new Hashtable();
        private Hashtable KOSPIHash = new Hashtable();
        private Hashtable KOSDAQHash = new Hashtable();
        private Hashtable KRXHash = new Hashtable();
        private KOREA_IndexGeneratorConfig configObj = null;
        private Logger logger = null;

        protected override void Start()
        {
            StartIndexGeneratorJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(KOREA_IndexGeneratorConfig)) as KOREA_IndexGeneratorConfig;
            logger = new Logger(configObj.LOG_FILE_PATH + LOGFILE_NAME, Logger.LogMode.New);
        }

        public void StartIndexGeneratorJob()
        {
            ReadDataFromKSandKQListingEquityList_xls();
            ReadDataFromIndexListingList_xls();
            //GrabDataFromOrignalFile_xls();     KRA6531471B8
            ControlExcel();
        }

        private void ReadDataFromKSandKQListingEquityList_xls()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                Logger.Log("Excel Application could not be created! please check the referenec is correct!", Logger.LogType.Warning);
                return;
            }

            try
            {
                string ipath = configObj.Korea_KQorKSList_ReadFilePath_CONFIG.WORKBOOK_PATH;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = ExcelUtil.GetWorksheet(configObj.Korea_KQorKSList_ReadFilePath_CONFIG.WORKSHEET_NAME, wBook);
                if (wSheet == null)
                {
                    Logger.Log("Excel Worksheet could not be created! please check the referenec is correct!", Logger.LogType.Warning);
                    return;
                }
                int startLine = 2;
                while (wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null)
                {
                    if (wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                    {
                        KSorKQListingList listing = new KSorKQListingList();
                        if (wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null)
                            listing.Ric = ((Range)wSheet.Cells[startLine, 1]).Value2.ToString().Trim();
                        if (wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null)
                            listing.IDNDisplayName = ((Range)wSheet.Cells[startLine, 2]).Value2.ToString().Trim();
                        if (wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null)
                            listing.ISIN = ((Range)wSheet.Cells[startLine, 3]).Value2.ToString().Trim();
                        String ticker = "";
                        if (listing.Ric.IndexOf('.') > 0 && listing.Ric.Length == 9)
                        {
                            ticker = listing.Ric.Split('.')[0].Trim().ToString();
                            kskqlistingHash.Add(ticker, listing);
                            startLine++;
                        }
                        else
                        {
                            startLine++;
                            continue;
                        }
                    }
                }
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in ReadDataFromEquityMasterfile_xls  : " + ex.ToString(), Logger.LogType.Warning);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void ReadDataFromIndexListingList_xls()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                Logger.Log("Excel Application could not be created! please check the referenec is correct!", Logger.LogType.Warning);
                return;
            }

            try
            {
                string ipath = configObj.Korea_Index_ReadFile_CONFIG.WORKBOOK_PATH;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet_KOSPI = ExcelUtil.GetWorksheet(configObj.Korea_Index_ReadFile_CONFIG.KOSPI_SHEETNAME, wBook);
                Worksheet wSheet_KOSDAQ = ExcelUtil.GetWorksheet(configObj.Korea_Index_ReadFile_CONFIG.KOSDAQ_SHEETNAME, wBook);
                Worksheet wSheet_KRX = ExcelUtil.GetWorksheet(configObj.Korea_Index_ReadFile_CONFIG.KRX_SHEETNAME, wBook);
                if (wSheet_KOSPI == null || wSheet_KOSDAQ == null || wSheet_KRX == null)
                {
                    Logger.Log("Excel Worksheet could not be created! please check the referenec is correct!", Logger.LogType.Warning);
                    return;
                }
                int startLine = 2;
                startLine = GenerateKOSPIHash(wSheet_KOSPI, startLine);

                startLine = GenerateKOSDAQHash(wSheet_KOSDAQ, startLine);

                startLine = GenerateKRXHash(wSheet_KRX, startLine);

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in ReadDataFromIndexListingList_xls  : " + ex.ToString(), Logger.LogType.Warning);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private int GenerateKOSPIHash(Worksheet wSheet_KOSPI, int startLine)
        {
            try
            {
                while (wSheet_KOSPI.get_Range("A" + startLine, Type.Missing).Value2 != null && wSheet_KOSPI.get_Range("A" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                {
                    IndexListingList indexlisting = new IndexListingList();
                    indexlisting.KoreanNameForChainRIC = wSheet_KOSPI.get_Range("A" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSPI.get_Range("B" + startLine, Type.Missing).Value2 != null && wSheet_KOSPI.get_Range("B" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.EnglishNameForChainRIC = wSheet_KOSPI.get_Range("B" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSPI.get_Range("C" + startLine, Type.Missing).Value2 != null && wSheet_KOSPI.get_Range("C" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.IndexChainRIC = wSheet_KOSPI.get_Range("C" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSPI.get_Range("D" + startLine, Type.Missing).Value2 != null && wSheet_KOSPI.get_Range("D" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.Number = wSheet_KOSPI.get_Range("D" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSPI.get_Range("E" + startLine, Type.Missing).Value2 != null && wSheet_KOSPI.get_Range("E" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.One = wSheet_KOSPI.get_Range("E" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSPI.get_Range("F" + startLine, Type.Missing).Value2 != null && wSheet_KOSPI.get_Range("F" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.Two = wSheet_KOSPI.get_Range("F" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSPI.get_Range("G" + startLine, Type.Missing).Value2 != null && wSheet_KOSPI.get_Range("G" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.Three = wSheet_KOSPI.get_Range("G" + startLine, Type.Missing).Value2.ToString();

                    if (indexlisting.One != null && indexlisting.One != String.Empty && (!KOSPIHash.Contains(indexlisting.One)))
                        KOSPIHash.Add(indexlisting.One.ToUpper(), indexlisting);
                    if (indexlisting.Two != null && indexlisting.Two != String.Empty && (!KOSPIHash.Contains(indexlisting.Two)))
                        KOSPIHash.Add(indexlisting.Two.ToUpper(), indexlisting);
                    if (indexlisting.Three != null && indexlisting.Two != String.Empty && (!KOSPIHash.Contains(indexlisting.Three)))
                        KOSPIHash.Add(indexlisting.Three.ToUpper(), indexlisting);
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateKOSPIHash  : " + ex.ToString(), Logger.LogType.Error);
            }
            return startLine;
        }

        private int GenerateKOSDAQHash(Worksheet wSheet_KOSDAQ, int startLine)
        {
            startLine = 2;
            try
            {

                while (wSheet_KOSDAQ.get_Range("A" + startLine, Type.Missing).Value2 != null && wSheet_KOSDAQ.get_Range("A" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                {
                    IndexListingList indexlisting = new IndexListingList();
                    indexlisting.KoreanNameForChainRIC = wSheet_KOSDAQ.get_Range("A" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSDAQ.get_Range("B" + startLine, Type.Missing).Value2 != null && wSheet_KOSDAQ.get_Range("B" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.EnglishNameForChainRIC = wSheet_KOSDAQ.get_Range("B" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSDAQ.get_Range("C" + startLine, Type.Missing).Value2 != null && wSheet_KOSDAQ.get_Range("C" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.IndexChainRIC = wSheet_KOSDAQ.get_Range("C" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSDAQ.get_Range("D" + startLine, Type.Missing).Value2 != null && wSheet_KOSDAQ.get_Range("D" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.Number = wSheet_KOSDAQ.get_Range("D" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSDAQ.get_Range("E" + startLine, Type.Missing).Value2 != null && wSheet_KOSDAQ.get_Range("E" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.One = wSheet_KOSDAQ.get_Range("E" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSDAQ.get_Range("F" + startLine, Type.Missing).Value2 != null && wSheet_KOSDAQ.get_Range("F" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.Two = wSheet_KOSDAQ.get_Range("F" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KOSDAQ.get_Range("G" + startLine, Type.Missing).Value2 != null && wSheet_KOSDAQ.get_Range("G" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.Three = wSheet_KOSDAQ.get_Range("G" + startLine, Type.Missing).Value2.ToString();

                    if (indexlisting.One != null && indexlisting.One != String.Empty && (!KOSDAQHash.Contains(indexlisting.One)))
                        KOSDAQHash.Add(indexlisting.One.ToUpper(), indexlisting);
                    if (indexlisting.Two != null && indexlisting.Two != String.Empty && (!KOSDAQHash.Contains(indexlisting.Two)))
                        KOSDAQHash.Add(indexlisting.Two.ToUpper(), indexlisting);
                    if (indexlisting.Three != null && indexlisting.Three != String.Empty && (!KOSDAQHash.Contains(indexlisting.Three)))
                        KOSDAQHash.Add(indexlisting.Three.ToUpper(), indexlisting);
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateKOSDAQHash  : " + ex.ToString(), Logger.LogType.Error);
            }
            return startLine;
        }

        private int GenerateKRXHash(Worksheet wSheet_KRX, int startLine)
        {
            startLine = 2;
            try
            {
                while (wSheet_KRX.get_Range("A" + startLine, Type.Missing).Value2 != null && wSheet_KRX.get_Range("A" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                {
                    IndexListingList indexlisting = new IndexListingList();
                    indexlisting.KoreanNameForChainRIC = wSheet_KRX.get_Range("A" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KRX.get_Range("B" + startLine, Type.Missing).Value2 != null && wSheet_KRX.get_Range("B" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.EnglishNameForChainRIC = wSheet_KRX.get_Range("B" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KRX.get_Range("C" + startLine, Type.Missing).Value2 != null && wSheet_KRX.get_Range("C" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.IndexChainRIC = wSheet_KRX.get_Range("C" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KRX.get_Range("D" + startLine, Type.Missing).Value2 != null && wSheet_KRX.get_Range("D" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.Number = wSheet_KRX.get_Range("D" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KRX.get_Range("E" + startLine, Type.Missing).Value2 != null && wSheet_KRX.get_Range("E" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.One = wSheet_KRX.get_Range("E" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KRX.get_Range("F" + startLine, Type.Missing).Value2 != null && wSheet_KRX.get_Range("F" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.Two = wSheet_KRX.get_Range("F" + startLine, Type.Missing).Value2.ToString();
                    if (wSheet_KRX.get_Range("G" + startLine, Type.Missing).Value2 != null && wSheet_KRX.get_Range("G" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        indexlisting.Three = wSheet_KRX.get_Range("G" + startLine, Type.Missing).Value2.ToString();

                    if (indexlisting.One != null && indexlisting.One != String.Empty && (!KRXHash.Contains(indexlisting.One)))
                        KRXHash.Add(indexlisting.One.ToUpper(), indexlisting);
                    if (indexlisting.Two != null && indexlisting.Two != String.Empty && (!KRXHash.Contains(indexlisting.Two)))
                        KRXHash.Add(indexlisting.Two.ToUpper(), indexlisting);
                    if (indexlisting.Three != null && indexlisting.Three != String.Empty && (!KRXHash.Contains(indexlisting.Three)))
                        KRXHash.Add(indexlisting.Three.ToUpper(), indexlisting);
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateKRXHash  : " + ex.ToString(), Logger.LogType.Error);
            }
            return startLine;
        }

        /*-------------------------------------------------------test-----------------------------------------------*/

        private void ControlExcel()
        {
            ExcelApp excelApp = new ExcelApp(false, false);
            try
            {
                if (excelApp.ExcelAppInstance == null)
                {
                    Logger.Log("", Logger.LogType.Error);
                    return;
                }
                String ipath = configObj.Korea_index_OrignalFile_CONFIG.WORKBOOK_PATH;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                int sheetCounts = wBook.Worksheets.Count;
                for (var i = 1; i <= sheetCounts; i++)
                {
                    Worksheet wSheet = (Worksheet)wBook.Worksheets[i];
                    String sheetname = wSheet.Name;
                    if (!sheetname.Contains("섹터"))
                    {
                        String chainRic = "";
                        String[] sname_arr = sheetname.Split('(');
                        String str_name = sname_arr[0].Trim().ToString();
                        String sname = str_name.Contains("KOSPI") ? str_name.Replace("KOSPI", "") : (str_name.Contains("KOSDAQ") ? str_name.Replace("KOSDAQ", "") : (str_name.Contains("KRX") ? str_name.Replace("KRX", "") : str_name));
                        String str_instrument = sname_arr[(sname_arr.Length - 1)].Trim(new Char[] { ' ', ')' }).ToString();
                        switch (str_instrument)
                        {
                            case "변경종목":
                                wSheet.Name = "the change of " + sname;
                                break;
                            case "구성종목":
                                wSheet.Name = "the constituents of " + sname;
                                break;
                        }

                        String temp_str_name = String.Empty;
                        if (str_name.Contains("KOSPI") || str_name.Contains("코스피"))
                        {
                            temp_str_name = str_name.Contains("KOSPI") ? str_name.Replace("KOSPI", "").Trim().ToString().ToUpper().Replace(" ", "") : str_name.Replace("코스피", "").Trim().ToString().ToUpper().Replace(" ", "");
                            if (KOSPIHash.Contains(temp_str_name))
                                chainRic = ((IndexListingList)KOSPIHash[temp_str_name]).IndexChainRIC;
                        }
                        else if (str_name.Contains("KOSDAQ") || str_name.Contains("코스닥"))
                        {
                            temp_str_name = str_name.Contains("KOSDAQ") ? str_name.Replace("KOSDAQ", "").Trim().ToString().ToUpper().Replace(" ", "") : str_name.Replace("코스닥", "").Trim().ToString().ToUpper().Replace(" ", "");
                            if (KOSDAQHash.Contains(temp_str_name))
                                chainRic = ((IndexListingList)KOSDAQHash[temp_str_name]).IndexChainRIC;
                        }
                        else if (str_name.Contains("KRX"))
                        {
                            temp_str_name = str_name.Replace("KRX", "").Trim().ToString().ToUpper().Replace(" ", "");
                            if (KRXHash.Contains(temp_str_name))
                                chainRic = ((IndexListingList)KRXHash[temp_str_name]).IndexChainRIC;
                        }
                        else
                        {
                            temp_str_name = str_name.ToUpper().Replace(" ", "");
                            if (chainRic == string.Empty && KOSPIHash.Contains(temp_str_name))
                                chainRic = ((IndexListingList)KOSPIHash[temp_str_name]).IndexChainRIC;
                            else if (chainRic == string.Empty && KOSDAQHash.Contains(temp_str_name))
                                chainRic = ((IndexListingList)KOSDAQHash[temp_str_name]).IndexChainRIC;
                            else if (chainRic == string.Empty && KRXHash.Contains(temp_str_name))
                                chainRic = ((IndexListingList)KRXHash[temp_str_name]).IndexChainRIC;
                            else
                                chainRic = "There doesn't exists Index Chain RIC can match Key Word(sheet name) !";
                        }

                        int startLine = 4;
                        if (wSheet.get_Range("A" + startLine, Type.Missing).Value2 == null && wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null)
                            ((Range)wSheet.Columns["A", Type.Missing]).Delete(XlDeleteShiftDirection.xlShiftToLeft);

                        String title = wSheet.get_Range("A1", Type.Missing).Value2.ToString();
                        title = wSheet.Name + " " + chainRic;
                        wSheet.Cells[1, 1] = title;

                        String columnA3 = wSheet.get_Range("A" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("A" + (startLine - 1), Type.Missing).Value2.ToString() : null;
                        String columnB3 = wSheet.get_Range("B" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("B" + (startLine - 1), Type.Missing).Value2.ToString() : null;
                        String columnC3 = wSheet.get_Range("C" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("C" + (startLine - 1), Type.Missing).Value2.ToString() : null;
                        String columnD3 = wSheet.get_Range("D" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("D" + (startLine - 1), Type.Missing).Value2.ToString() : null;

                        if (columnA3 == "신규편입 종목" && columnC3 == "제외종목")
                        {
                            startLine = ModifyTheDataWithAddItemsAndDropItems_Worksheet(wSheet, startLine);
                        }
                        else if (columnA3 == "종목코드" && columnB3 == "종목명" && columnC3 == "시장구분" && (columnD3 == "유동비율(%)" || columnD3 == "유동비율"))
                        {
                            startLine = ModifyTheDataWithFreeFloatRate_Worksheet(wSheet, startLine);
                        }
                        else if (columnA3 == "종목코드" && columnB3 == "종목명")
                        {
                            startLine = ModifyTheDataWithLocalLanguageName_Worksheet(wSheet, startLine);
                        }
                        else
                        {

                        }
                    }
                    else
                    {
                        sheetname = sheetname.Replace("섹터", "Sector");  //KRX Sector(변경종목)
                        String chainRic = "";
                        String[] sname_arr = sheetname.Split('(');
                        String str_name = sname_arr[0].Trim().ToString();
                        String str_instrument = sname_arr[(sname_arr.Length - 1)].Trim(new Char[] { ' ', ')' }).ToString();
                        switch (str_instrument)
                        {
                            case "변경종목":
                                wSheet.Name = "the change of " + str_name;
                                break;
                            case "구성종목":
                                wSheet.Name = "the constituents of " + str_name;
                                break;
                        }

                        int startLine = 3;
                        if (wSheet.get_Range("A" + startLine, Type.Missing).Value2 == null && wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() == "섹터구분")
                            ((Range)wSheet.Columns["A", Type.Missing]).Delete(XlDeleteShiftDirection.xlShiftToLeft);

                        String columnA3 = wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                        String columnB3 = wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                        String columnC3 = wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                        String columnD3 = wSheet.get_Range("D" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                        String columnE3 = wSheet.get_Range("E" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("E" + startLine, Type.Missing).Value2.ToString().Trim() : null;

                        if (columnA3 == "섹터구분" && columnB3 == "신규편입 종목" && columnD3 == "제외종목")
                        {
                            ModifyTheSectorDataWithAddItemsAndDropItems_Worksheet(wSheet, ref chainRic, ref startLine);
                        }
                        else if (columnA3 == "섹터구분" && columnB3 == "종목코드" && columnC3 == "종목명" && columnD3 == "시장구분" && (columnE3 == "유동비율(%)" || columnE3 == "유동비율"))
                        {
                            ModifyTheSectorDataWithFreeFloatRate_Worksheet(wSheet, ref chainRic, ref startLine);
                        }
                    }
                    excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                    wBook.Save();
                }
            }
            catch (Exception ex)
            {
                Logger.Log("" + ex.ToString(), Logger.LogType.Warning);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="wSheet"></param>
        /// <param name="chainRic"></param>
        /// <param name="startLine"></param>
        private void ModifyTheSectorDataWithFreeFloatRate_Worksheet(Worksheet wSheet, ref String chainRic, ref int startLine)
        {
            wSheet.Cells[3, 1] = "sector";
            wSheet.Cells[3, 2] = "Ric";
            ((Range)wSheet.Columns["C", Type.Missing]).Insert(XlInsertShiftDirection.xlShiftToRight, null);
            wSheet.Cells[3, 3] = "Display Name";
            wSheet.Cells[3, 4] = "Local Language Name";
            ((Range)wSheet.Columns["E", Type.Missing]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
            wSheet.Cells[3, 5] = "Free Float Rate(%)";

            startLine = 4;
            while (wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() != String.Empty)
            {
                String sectorname = wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                if (sectorname != null)
                {
                    String parameter = sectorname.Contains("(") ? sectorname.Split('(')[0].Trim() : sectorname;
                    if (parameter != String.Empty)
                        chainRic = GetSectorChainRic(parameter);
                    wSheet.Cells[startLine, 1] = parameter + " " + chainRic;
                }

                String ticker = wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim();
                if (ticker.Length < 6)
                    for (var x = ticker.Length; x < 6; x++)
                        ticker = "0" + ticker;
                if (kskqlistingHash.Contains(ticker))
                {
                    var item = kskqlistingHash[ticker] as KSorKQListingList;
                    wSheet.Cells[startLine, 2] = item.Ric;
                    wSheet.Cells[startLine, 3] = item.IDNDisplayName;
                }
                startLine++;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wSheet"></param>
        /// <param name="chainRic"></param>
        /// <param name="startLine"></param>
        private void ModifyTheSectorDataWithAddItemsAndDropItems_Worksheet(Worksheet wSheet, ref String chainRic, ref int startLine)
        {
            wSheet.Cells[3, 1] = "sector";
            wSheet.Cells[3, 2] = "Add Items";
            wSheet.Cells[3, 4] = "Drop Items";

            String columnA4 = wSheet.get_Range("A" + (startLine + 1), Type.Missing).Value2 != null ? wSheet.get_Range("A" + (startLine + 1), Type.Missing).Value2.ToString().Trim() : null;
            String columnB4 = wSheet.get_Range("B" + (startLine + 1), Type.Missing).Value2 != null ? wSheet.get_Range("B" + (startLine + 1), Type.Missing).Value2.ToString().Trim() : null;
            String columnC4 = wSheet.get_Range("C" + (startLine + 1), Type.Missing).Value2 != null ? wSheet.get_Range("C" + (startLine + 1), Type.Missing).Value2.ToString().Trim() : null;
            String columnD4 = wSheet.get_Range("D" + (startLine + 1), Type.Missing).Value2 != null ? wSheet.get_Range("D" + (startLine + 1), Type.Missing).Value2.ToString().Trim() : null;
            String columnE4 = wSheet.get_Range("E" + (startLine + 1), Type.Missing).Value2 != null ? wSheet.get_Range("E" + (startLine + 1), Type.Missing).Value2.ToString().Trim() : null;

            if (columnB4 == "종목코드" && columnC4 == "종목명" && columnD4 == "종목코드" && columnE4 == "종목명")
            {
                wSheet.Cells[4, 2] = "Ric";
                ((Range)wSheet.Columns["C", Type.Missing]).Insert(XlInsertShiftDirection.xlShiftToRight, null);
                wSheet.Cells[4, 3] = "Display Name";
                wSheet.Cells[4, 4] = "Local Language Name";

                wSheet.Cells[4, 5] = "Ric";
                ((Range)wSheet.Columns["F", Type.Missing]).Insert(XlInsertShiftDirection.xlShiftToRight, null);
                wSheet.Cells[4, 6] = "Display Name";
                wSheet.Cells[4, 7] = "Local Language Name";

                startLine = 5;
                while (wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() != String.Empty)
                {
                    String sectorname = wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                    if (sectorname != null)
                    {
                        String parameter = sectorname.Contains("(") ? sectorname.Split('(')[0].Trim() : sectorname;
                        if (parameter != String.Empty)
                            chainRic = GetSectorChainRic(parameter);
                        wSheet.Cells[startLine, 1] = parameter + " " + chainRic;
                    }

                    String aticker = wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                    if (aticker.Length < 6)
                        for (var x = aticker.Length; x < 6; x++)
                            aticker = "0" + aticker;

                    if (kskqlistingHash.Contains(aticker))
                    {
                        var aitem = kskqlistingHash[aticker] as KSorKQListingList;
                        wSheet.Cells[startLine, 2] = aitem.Ric;
                        wSheet.Cells[startLine, 3] = aitem.IDNDisplayName;
                    }

                    String dticker = wSheet.get_Range("E" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("E" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                    if (dticker.Length < 6)
                        for (var x = dticker.Length; x < 6; x++)
                            dticker = "0" + dticker;

                    if (kskqlistingHash.Contains(dticker))
                    {
                        var ditem = kskqlistingHash[dticker] as KSorKQListingList;
                        wSheet.Cells[startLine, 5] = ditem.Ric;
                        wSheet.Cells[startLine, 6] = ditem.IDNDisplayName;
                    }
                    startLine++;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="str_name"></param>
        /// <returns></returns>
        private String GetSectorChainRic(String str_name)
        {
            String chainRic = String.Empty;
            String temp_str_name = String.Empty;
            if (str_name.Contains("KOSPI") || str_name.Contains("코스피"))
            {
                temp_str_name = str_name.Contains("KOSPI") ? str_name.Replace("KOSPI", "").Trim().ToString().ToUpper().Replace(" ", "") : str_name.Replace("코스피", "").Trim().ToString().ToUpper().Replace(" ", "");
                if (KOSPIHash.Contains(temp_str_name))
                    chainRic = ((IndexListingList)KOSPIHash[temp_str_name]).IndexChainRIC;
            }
            else if (str_name.Contains("KOSDAQ") || str_name.Contains("코스닥"))
            {
                temp_str_name = str_name.Contains("KOSDAQ") ? str_name.Replace("KOSDAQ", "").Trim().ToString().ToUpper().Replace(" ", "") : str_name.Replace("코스닥", "").Trim().ToString().ToUpper().Replace(" ", "");
                if (KOSDAQHash.Contains(temp_str_name))
                    chainRic = ((IndexListingList)KOSDAQHash[temp_str_name]).IndexChainRIC;
            }
            else if (str_name.Contains("KRX"))
            {
                temp_str_name = str_name.Replace("KRX", "").Trim().ToString().ToUpper().Replace(" ", "");
                if (KRXHash.Contains(temp_str_name))
                    chainRic = ((IndexListingList)KRXHash[temp_str_name]).IndexChainRIC;
            }
            else
            {
                temp_str_name = str_name.ToUpper().Replace(" ", "");
                if (chainRic == string.Empty && KOSPIHash.Contains(temp_str_name))
                    chainRic = ((IndexListingList)KOSPIHash[temp_str_name]).IndexChainRIC;
                else if (chainRic == string.Empty && KOSDAQHash.Contains(temp_str_name))
                    chainRic = ((IndexListingList)KOSDAQHash[temp_str_name]).IndexChainRIC;
                else if (chainRic == string.Empty && KRXHash.Contains(temp_str_name))
                    chainRic = ((IndexListingList)KRXHash[temp_str_name]).IndexChainRIC;
                else
                    chainRic = "There doesn't exists Index Chain RIC can match Key Word(sheet name) !";
            }
            return chainRic;
        }

        /// <summary>
        /// Modify the <common index> Worksheet's data which contains Add Items and Drop Items
        /// </summary>
        /// <param name="wSheet"></param>
        /// <param name="startLine"></param>
        /// <returns></returns>
        private int ModifyTheDataWithAddItemsAndDropItems_Worksheet(Worksheet wSheet, int startLine)
        {
            wSheet.Cells[3, 1] = "Add Items";
            wSheet.Cells[3, 3] = "Drop Items";
            String columnA4 = wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString() : null;
            String columnB4 = wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString() : null;
            String columnC4 = wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString() : null;
            String columnD4 = wSheet.get_Range("D" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString() : null;

            if (columnA4 == "종목코드" && columnB4 == "종목명" && columnC4 == "종목코드" && columnD4 == "종목명")
            {
                wSheet.Cells[4, 1] = "Ric";
                wSheet.Cells[4, 2] = "Local Language name";
                wSheet.Cells[4, 3] = "Ric";
                wSheet.Cells[4, 4] = "Local Language name";
                ((Range)wSheet.Columns["B", Type.Missing]).Insert(XlInsertShiftDirection.xlShiftToRight, null);
                wSheet.Cells[4, 2] = "Display Name";
                ((Range)wSheet.Columns["E", Type.Missing]).Insert(XlInsertShiftDirection.xlShiftToRight, null);
                wSheet.Cells[4, 5] = "Display Name";
                startLine = 5;

                while (wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                {
                    String aticker = wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim();
                    if (aticker.Length < 6)
                        for (var x = aticker.Length; x < 6; x++)
                            aticker = "0" + aticker;
                    if (kskqlistingHash.Contains(aticker))
                    {
                        var aitem = kskqlistingHash[aticker] as KSorKQListingList;
                        wSheet.Cells[startLine, 1] = aitem.Ric;
                        wSheet.Cells[startLine, 2] = aitem.IDNDisplayName;
                    }

                    String dticker = wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim();
                    if (dticker.Length < 6)
                        for (var x = dticker.Length; x < 6; x++)
                            dticker = "0" + dticker;
                    if (kskqlistingHash.Contains(dticker))
                    {
                        var ditem = kskqlistingHash[dticker] as KSorKQListingList;
                        wSheet.Cells[startLine, 4] = ditem.Ric;
                        wSheet.Cells[startLine, 5] = ditem.IDNDisplayName;
                    }
                    startLine++;
                }
            }
            return startLine;
        }

        /// <summary>
        /// Modify the <common index> Worksheet's data which Free Float Rate(%)
        /// </summary>
        /// <param name="wSheet"></param>
        /// <param name="startLine"></param>
        /// <returns></returns>
        private int ModifyTheDataWithFreeFloatRate_Worksheet(Worksheet wSheet, int startLine)
        {
            wSheet.Cells[3, 1] = "Ric";
            ((Range)wSheet.Columns["B", Type.Missing]).Insert(XlInsertShiftDirection.xlShiftToRight, null);
            wSheet.Cells[3, 2] = "Display Name";
            wSheet.Cells[3, 3] = "Local Language name";
            ((Range)wSheet.Columns["D", Type.Missing]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
            wSheet.Cells[3, 4] = "Free Float Rate(%)";
            startLine = 4;
            while (wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim() != String.Empty)
            {
                String ticker = wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim();
                if (ticker.Length < 6)
                    for (var x = ticker.Length; x < 6; x++)
                        ticker = "0" + ticker;

                if (kskqlistingHash.Contains(ticker))
                {
                    var item = kskqlistingHash[ticker] as KSorKQListingList;
                    wSheet.Cells[startLine, 1] = item.Ric;
                    wSheet.Cells[startLine, 2] = item.IDNDisplayName;
                }
                startLine++;
            }
            return startLine;
        }

        /// <summary>
        /// Modify the <common index> Worksheet's data which contains Local Language Name
        /// </summary>
        /// <param name="wSheet"></param>
        /// <param name="startLine"></param>
        /// <returns></returns>
        private int ModifyTheDataWithLocalLanguageName_Worksheet(Worksheet wSheet, int startLine)
        {
            wSheet.Cells[3, 1] = "Ric";
            ((Range)wSheet.Columns["B", Type.Missing]).Insert(XlInsertShiftDirection.xlShiftToRight, null);
            wSheet.Cells[3, 2] = "Display Name";
            wSheet.Cells[3, 3] = "Local Language name";
            startLine = 4;
            while (wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim() != String.Empty)
            {
                String ticker = wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim();
                if (ticker.Length < 6)
                    for (var x = ticker.Length; x < 6; x++)
                        ticker = "0" + ticker;

                if (kskqlistingHash.Contains(ticker))
                {
                    var item = kskqlistingHash[ticker] as KSorKQListingList;
                    wSheet.Cells[startLine, 1] = item.Ric;
                    wSheet.Cells[startLine, 2] = item.IDNDisplayName;
                }
                startLine++;
            }
            return startLine;
        }


        /*------------------------------------------------------- useless -----------------------------------------------*/

        #region   useless code

        private void GrabDataFromOrignalFile_xls()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = null;
            ExcelApp eap = null;
            try
            {

                excelApp = new ExcelApp(false, false);
                if (excelApp.ExcelAppInstance == null)
                {
                    Logger.Log("", Logger.LogType.Warning);
                    return;
                }
                eap = new ExcelApp(false, false);
                if (eap.ExcelAppInstance == null)
                {
                    Logger.Log("Excel application could not be created ! please check the installation and object reference correct !", Logger.LogType.Warning);
                    return;
                }
                String GeneratefilePath = configObj.Korea_Index_GeneratorFile_CONFIG.WORKBOOK_PATH;
                Workbook gBook = ExcelUtil.CreateOrOpenExcelFile(eap, GeneratefilePath);

                String ipath = configObj.Korea_index_OrignalFile_CONFIG.WORKBOOK_PATH;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                int count = wBook.Worksheets.Count;
                for (var i = 1; i <= count; i++)
                {
                    Worksheet wSheet = (Worksheet)wBook.Worksheets[i];
                    Worksheet gSheet = null;
                    IndexTemplate indexTemp = null;
                    if (wSheet != null)
                    {
                        indexTemp = new IndexTemplate();
                    }
                    string sheetname = wSheet.Name;
                    if (!sheetname.Contains("섹터"))
                    {
                        String chainRic = "";
                        String[] sname_arr = sheetname.Split('(');
                        String str_name = sname_arr[0].Trim().ToUpper().ToString().Replace(" ", "");
                        String sname = str_name.Contains("KOSPI") ? str_name.Replace("KOSPI", "") : (str_name.Contains("KOSDAQ") ? str_name.Replace("KOSDAQ", "") : (str_name.Contains("KRX") ? str_name.Replace("KRX", "") : str_name));
                        String str_instrument = sname_arr[(sname_arr.Length - 1)].Trim(new Char[] { ' ', ')' }).ToString();
                        switch (str_instrument)
                        {
                            case "변경종목":
                                indexTemp.SheetName = "the change of " + sname;
                                break;
                            case "구성종목":
                                indexTemp.SheetName = "the constituents of " + sname;
                                break;
                        }

                        if (str_name.Contains("KOSPI") || str_name.Contains("코스피"))
                        {
                            str_name = str_name.Contains("KOSPI") ? str_name.Replace("KOSPI", "").Trim().ToString() : str_name.Replace("코스피", "").Trim().ToString();
                            if (KOSPIHash.Contains(str_name))
                                chainRic = ((IndexListingList)KOSPIHash[str_name]).IndexChainRIC;
                        }
                        else if (str_name.Contains("KOSDAQ") || str_name.Contains("코스닥"))
                        {
                            str_name = str_name.Contains("KOSDAQ") ? str_name.Replace("KOSDAQ", "").Trim().ToString() : str_name.Replace("코스닥", "").Trim().ToString();
                            if (KOSDAQHash.Contains(str_name))
                                chainRic = ((IndexListingList)KOSDAQHash[str_name]).IndexChainRIC;
                        }
                        else if (str_name.Contains("KRX"))
                        {
                            str_name = str_name.Replace("KRX", "").Trim().ToString();
                            if (KRXHash.Contains(str_name))
                                chainRic = ((IndexListingList)KRXHash[str_name]).IndexChainRIC;
                        }
                        else
                        {
                            if (chainRic == string.Empty && KOSPIHash.Contains(str_name))
                                chainRic = ((IndexListingList)KOSPIHash[str_name]).IndexChainRIC;
                            else if (chainRic == string.Empty && KOSDAQHash.Contains(str_name))
                                chainRic = ((IndexListingList)KOSDAQHash[str_name]).IndexChainRIC;
                            else if (chainRic == string.Empty && KRXHash.Contains(str_name))
                                chainRic = ((IndexListingList)KRXHash[str_name]).IndexChainRIC;
                            else
                                chainRic = "There doesn't exists Index Chain RIC can match Key Word(sheet name) !";
                        }

                        int startLine = 4;
                        String title = wSheet.get_Range("A" + startLine, Type.Missing).Value2 == null ? wSheet.get_Range("B1", Type.Missing).Value2.ToString() : wSheet.get_Range("A1", Type.Missing).Value2.ToString();
                        title = indexTemp.SheetName + " " + chainRic;
                        indexTemp.SheetTitle = title;
                        //the columns A didn't contains data , columns B contains data
                        if (wSheet.get_Range("A" + startLine, Type.Missing).Value2 == null && wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null)
                        {
                            String columnsB3 = wSheet.get_Range("B" + (startLine - 1), Type.Missing).Value2.ToString().Trim();
                            String columnsD3 = wSheet.get_Range("D" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("D" + (startLine - 1), Type.Missing).Value2.ToString().Trim() : null;
                            String columnsC3 = wSheet.get_Range("C" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("C" + (startLine - 1), Type.Missing).Value2.ToString().Trim() : null;
                            String columnsE3 = wSheet.get_Range("E" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("E" + (startLine - 1), Type.Missing).Value2.ToString().Trim() : null;
                            String str_columnB = "B";
                            if (columnsB3 == "신규편입 종목" && columnsC3 == null && columnsD3 == "제외종목" && columnsE3 == null)
                            {
                                indexTemp.IsAddItem = true;
                                indexTemp.IsDropItem = true;
                                String columnsB4 = wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim();
                                String columnsC4 = wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                                String columnsD4 = wSheet.get_Range("D" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                                String columnsE4 = wSheet.get_Range("E" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("E" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                                if (columnsB4 == "종목코드" && columnsD4 == "종목코드" && columnsC4 == "종목명" && columnsE4 == "종목명")
                                {
                                    startLine = GrabDataWithAddItemsAndDropItems(wSheet, indexTemp, startLine, str_columnB);
                                    GenerateAddItemsAndDropItemsFileSheet_xls(gBook, i, indexTemp, gSheet);
                                }
                                eap.ExcelAppInstance.AlertBeforeOverwriting = false;
                                gBook.Save();
                            }
                            else if (columnsB3 == "종목코드" && columnsC3 == "종목명" && columnsD3 == "시장구분" && (columnsE3 == "유동비율(%)" || columnsE3 == "유동비율"))
                            {
                                indexTemp.IsAddItem = false;
                                indexTemp.IsDropItem = false;
                                startLine = GrabDataWithFreeFloatRateItems(wSheet, indexTemp, startLine, str_columnB);
                                GenerateFreeFloatRateItemsFileSheet_xls(gBook, i, indexTemp, gSheet);
                                eap.ExcelAppInstance.AlertBeforeOverwriting = false;
                                gBook.Save();
                            }
                            else if (columnsB3 == "종목코드" && columnsC3 == "종목명" && columnsD3 == null && columnsE3 == null)
                            {
                                indexTemp.IsAddItem = false;
                                indexTemp.IsDropItem = false;
                                startLine = GrabDataWithLocalLanguageNameItems(wSheet, indexTemp, startLine, str_columnB);

                                GenerateLocalLanguageNameItemFileSheet_xls(gBook, i, indexTemp, gSheet);
                                eap.ExcelAppInstance.AlertBeforeOverwriting = false;
                                gBook.Save();
                            }
                            else
                            {
                                continue;
                            }
                        }
                        //the columns A contains data
                        startLine = 4;
                        if (wSheet.get_Range("A" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                        {
                            String columnsA3 = wSheet.get_Range("A" + (startLine - 1), Type.Missing).Value2.ToString().Trim();
                            String columnsC3 = wSheet.get_Range("C" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("C" + (startLine - 1), Type.Missing).Value2.ToString().Trim() : null;
                            String columnsB3 = wSheet.get_Range("B" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("B" + (startLine - 1), Type.Missing).Value2.ToString().Trim() : null;
                            String columnsD3 = wSheet.get_Range("D" + (startLine - 1), Type.Missing).Value2 != null ? wSheet.get_Range("D" + (startLine - 1), Type.Missing).Value2.ToString().Trim() : null;
                            String str_columnA = "A";
                            if (columnsA3 == "신규편입 종목" && columnsB3 == null && columnsC3 == "제외종목" && columnsD3 == null)
                            {
                                indexTemp.IsAddItem = true;
                                indexTemp.IsDropItem = true;
                                String columnsA4 = wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim();
                                String columnsB4 = wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                                String columnsC4 = wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                                String columnsD4 = wSheet.get_Range("D" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                                if (columnsA4 == "종목코드" && columnsC4 == "종목코드" && columnsB4 == "종목명" && columnsD4 == "종목명")
                                {
                                    startLine = GrabDataWithAddItemsAndDropItems(wSheet, indexTemp, startLine, str_columnA);
                                    GenerateAddItemsAndDropItemsFileSheet_xls(gBook, i, indexTemp, gSheet);
                                }
                                eap.ExcelAppInstance.AlertBeforeOverwriting = false;
                                gBook.Save();
                            }
                            else if (columnsA3 == "종목코드" && columnsB3 == "종목명" && columnsC3 == "시장구분" && (columnsD3 == "유동비율(%)" || columnsD3 == "유동비율"))
                            {
                                indexTemp.IsAddItem = false;
                                indexTemp.IsDropItem = false;
                                startLine = GrabDataWithFreeFloatRateItems(wSheet, indexTemp, startLine, str_columnA);
                                GenerateFreeFloatRateItemsFileSheet_xls(gBook, i, indexTemp, gSheet);
                                eap.ExcelAppInstance.AlertBeforeOverwriting = false;
                                gBook.Save();
                            }
                            else if (columnsA3 == "종목코드" && columnsB3 == "종목명" && columnsC3 == null && columnsD3 == null)
                            {
                                indexTemp.IsAddItem = false;
                                indexTemp.IsDropItem = false;
                                startLine = GrabDataWithLocalLanguageNameItems(wSheet, indexTemp, startLine, str_columnA);
                                GenerateLocalLanguageNameItemFileSheet_xls(gBook, i, indexTemp, gSheet);
                                eap.ExcelAppInstance.AlertBeforeOverwriting = false;
                                gBook.Save();
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                    // this will be used for sector
                    else
                    {
                        sheetname = sheetname.Replace("섹터", "Sector");
                        String chainRic = "";
                        String[] sname_arr = sheetname.Split('(');
                        String str_name = sname_arr[0].Trim().ToUpper().ToString().Replace(" ", "");
                        String sname = str_name.Contains("KOSPI") ? str_name.Replace("KOSPI", "") : (str_name.Contains("KOSDAQ") ? str_name.Replace("KOSDAQ", "") : (str_name.Contains("KRX") ? str_name.Replace("KRX", "") : str_name));
                        String str_instrument = sname_arr[(sname_arr.Length - 1)].Trim(new Char[] { ' ', ')' }).ToString();
                        switch (str_instrument)
                        {
                            case "변경종목":
                                indexTemp.SheetName = "the change of " + str_name;
                                break;
                            case "구성종목":
                                indexTemp.SheetName = "the constituents of " + str_name;
                                break;
                        }

                        int startLine = 3;
                        String title = wSheet.get_Range("A" + startLine, Type.Missing).Value2 == null ? wSheet.get_Range("B1", Type.Missing).Value2.ToString() : wSheet.get_Range("A1", Type.Missing).Value2.ToString();
                        title = indexTemp.SheetName + " " + chainRic;
                        indexTemp.SheetTitle = title;
                        //the columns A didn't contains data , columns B contains data
                        if (wSheet.get_Range("A" + startLine, Type.Missing).Value2 == null && wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null)
                        {
                            String columnB3 = wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim();
                            String columnC3 = wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                            String columnD3 = wSheet.get_Range("D" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                            String columnE3 = wSheet.get_Range("E" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("E" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                            String columnF3 = wSheet.get_Range("F" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("F" + startLine, Type.Missing).Value2.ToString().Trim() : null;

                            if (columnB3 == "섹터구분" && columnC3 == "신규편입 종목" && columnD3 == null && columnE3 == "제외종목" && columnF3 == null)
                            {
                                indexTemp.IsAddItem = true;
                                indexTemp.IsDropItem = true;
                                startLine = GrabSectorDataWithAddItemsAndDropItemsFromWorksheet(wSheet, indexTemp, startLine);

                                GenerateSectorAddItemsAndDropItemsFileSheet_xls(gBook, i, gSheet, indexTemp);
                                eap.ExcelAppInstance.AlertBeforeOverwriting = false;
                                gBook.Save();

                            }
                            else if (columnB3 == "섹터구분" && columnC3 == "종목코드" && columnD3 == "종목명" && columnE3 == "시장구분" && (columnF3 == "유동비율(%)" || columnF3 == "유동비율"))
                            {
                                indexTemp.IsAddItem = false;
                                indexTemp.IsDropItem = false;
                                startLine = GrabSectorDataWithFreeFloatRateItemsFromWorksheet(wSheet, indexTemp, startLine);

                                GenerateSectorFreeFloatRateItemsFileSheet_xls(gBook, i, gSheet, indexTemp);


                                eap.ExcelAppInstance.AlertBeforeOverwriting = false;
                                gBook.Save();
                            }

                        }
                        else
                        {

                        }




                    }
                }

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GrabDataFromOrignalFile_xls" + ex.ToString(), Logger.LogType.Warning);
                return;
            }
            finally
            {
                excelApp.Dispose();
                eap.Dispose();
            }
        }

        private int GrabDataWithAddItemsAndDropItems(Worksheet wSheet, IndexTemplate indexTemp, int startLine, String str_column)
        {
            List<IndexAddItem> addList = new List<IndexAddItem>();
            List<IndexDropItem> dropList = new List<IndexDropItem>();
            startLine = 5;
            String column = str_column.Equals("A") ? "C" : "D";
            while (wSheet.get_Range(str_column + startLine, Type.Missing).Value2 != null && wSheet.get_Range(column + startLine, Type.Missing).Value2 != null)
            {
                if (wSheet.get_Range(str_column + startLine, Type.Missing).Value2.ToString() != String.Empty)
                {
                    IndexAddItem additem = new IndexAddItem();
                    IndexDropItem dropitem = new IndexDropItem();
                    String _aticker = str_column.Equals("A") ? wSheet.get_Range("A" + startLine, Type.Missing).Value2.ToString().Trim() : wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim();
                    String _alocalname = str_column.Equals("A") ? wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() : wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim();
                    if (_aticker.Length < 6)
                        for (var i = _aticker.Length; i < 6; i++)
                            _aticker = "0" + _aticker;
                    KSorKQListingList aitem = null;
                    if (kskqlistingHash.Contains(_aticker))
                        aitem = kskqlistingHash[_aticker] as KSorKQListingList;
                    additem.LocalLanguageName = _alocalname;
                    additem.Ric = aitem.Ric;
                    additem.DisplayName = aitem.IDNDisplayName;

                    String _dticker = str_column.Equals("A") ? wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim() : wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim();
                    String _dlocalname = str_column.Equals("A") ? wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim() : wSheet.get_Range("E" + startLine, Type.Missing).Value2.ToString().Trim();
                    if (_dticker.Length < 6)
                        for (var i = _dticker.Length; i < 6; i++)
                            _dticker = "0" + _dticker;
                    KSorKQListingList ditem = null;
                    if (kskqlistingHash.Contains(_dticker))
                        ditem = kskqlistingHash[_dticker] as KSorKQListingList;
                    dropitem.Ric = ditem.Ric;
                    dropitem.DisplayName = ditem.IDNDisplayName;
                    dropitem.LocalLanguageName = _dlocalname;
                    addList.Add(additem);
                    dropList.Add(dropitem);
                    startLine++;
                }
            }
            indexTemp.AddItemList = addList;
            indexTemp.DropItemList = dropList;
            return startLine;
        }

        private void GenerateAddItemsAndDropItemsFileSheet_xls(Workbook gBook, int i, IndexTemplate indexTemp, Worksheet gSheet)
        {
            try
            {
                gSheet = (Worksheet)gBook.Worksheets[i];
                gSheet.Name = indexTemp.SheetName;
                if (indexTemp.IsAddItem && indexTemp.IsDropItem)
                {
                    ((Range)gSheet.Columns["A", System.Type.Missing]).ColumnWidth = 12;
                    ((Range)gSheet.Columns["B", System.Type.Missing]).ColumnWidth = 25;
                    ((Range)gSheet.Columns["C", System.Type.Missing]).ColumnWidth = 25;
                    ((Range)gSheet.Columns["D", System.Type.Missing]).ColumnWidth = 12;
                    ((Range)gSheet.Columns["E", System.Type.Missing]).ColumnWidth = 25;
                    ((Range)gSheet.Columns["F", System.Type.Missing]).ColumnWidth = 25;
                    ((Range)gSheet.Columns["F", System.Type.Missing]).EntireColumn.AutoFit();
                    ((Range)gSheet.Columns["A:F", System.Type.Missing]).Font.Name = "Arial";

                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Size = 16;
                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Underline = true;
                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)gSheet.Rows[1, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    gSheet.Cells.get_Range("A1", "F1").MergeCells = true;
                    gSheet.Cells[1, 1] = indexTemp.SheetTitle;

                    ((Range)gSheet.Rows[3, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)gSheet.Rows[3, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ((Range)gSheet.Rows[3, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    gSheet.Cells.get_Range("A3", "C3").MergeCells = true;
                    gSheet.Cells[3, 1] = "Add Items";
                    gSheet.Cells.get_Range("D3", "F3").MergeCells = true;
                    gSheet.Cells[3, 4] = "Drop Items";

                    ((Range)gSheet.Rows[4, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)gSheet.Rows[4, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ((Range)gSheet.Rows[4, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    gSheet.Cells[4, 1] = "RIC";
                    gSheet.Cells[4, 2] = "Display name";
                    gSheet.Cells[4, 3] = "Local language name";
                    gSheet.Cells[4, 4] = "RIC";
                    gSheet.Cells[4, 5] = "Display name";
                    gSheet.Cells[4, 6] = "Local language name";

                    int rindex = 5;
                    for (var j = 0; j < indexTemp.AddItemList.Count; j++)
                    {
                        gSheet.Cells[rindex, 1] = indexTemp.AddItemList[j].Ric;
                        gSheet.Cells[rindex, 2] = indexTemp.AddItemList[j].DisplayName;
                        gSheet.Cells[rindex, 3] = indexTemp.AddItemList[j].LocalLanguageName;
                        gSheet.Cells[rindex, 4] = indexTemp.DropItemList[j].Ric;
                        gSheet.Cells[rindex, 5] = indexTemp.DropItemList[j].DisplayName;
                        gSheet.Cells[rindex, 6] = indexTemp.DropItemList[j].LocalLanguageName;
                        rindex++;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateAddItemsAndDropItemsFileSheet_xls  : " + ex.ToString(), Logger.LogType.Warning);
                return;
            }
        }

        private int GrabDataWithFreeFloatRateItems(Worksheet wSheet, IndexTemplate indexTemp, int startLine, String str_column)
        {
            startLine = 4;
            List<FloatRateIndex> RateIndexList = new List<FloatRateIndex>();

            while (wSheet.get_Range(str_column + startLine, Type.Missing).Value2 != null && wSheet.get_Range(str_column + startLine, Type.Missing).Value2.ToString() != string.Empty)
            {
                FloatRateIndex FRIndex = new FloatRateIndex();
                String ticker = wSheet.get_Range(str_column + startLine, Type.Missing).Value2.ToString().Trim();
                if (ticker.Length < 6)
                    for (var i = ticker.Length; i < 6; i++)
                        ticker = "0" + ticker;

                FRIndex.LocalLanguageName = str_column.Equals("A") ? wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() : wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim();
                FRIndex.FreeFloatingRate = str_column.Equals("A") ? wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim() : wSheet.get_Range("E" + startLine, Type.Missing).Value2.ToString().Trim();
                KSorKQListingList kskqlisting = null;
                if (kskqlistingHash.Contains(ticker))
                    kskqlisting = (KSorKQListingList)kskqlistingHash[ticker];
                else
                {
                    startLine++;
                    continue;
                }
                FRIndex.Ric = kskqlisting.Ric;
                FRIndex.DisplayName = kskqlisting.IDNDisplayName;
                RateIndexList.Add(FRIndex);
                startLine++;
            }
            indexTemp.FRIList = RateIndexList;
            return startLine;
        }

        private void GenerateFreeFloatRateItemsFileSheet_xls(Workbook gBook, int i, IndexTemplate indexTemp, Worksheet gSheet)
        {
            try
            {
                gSheet = (Worksheet)gBook.Worksheets[i];
                gSheet.Name = indexTemp.SheetName;
                if (indexTemp.FRIList.Count > 0)
                {
                    ((Range)gSheet.Columns["A", System.Type.Missing]).ColumnWidth = 12;
                    ((Range)gSheet.Columns["B", System.Type.Missing]).ColumnWidth = 25;
                    ((Range)gSheet.Columns["C", System.Type.Missing]).ColumnWidth = 25;
                    ((Range)gSheet.Columns["D", System.Type.Missing]).ColumnWidth = 20;
                    ((Range)gSheet.Columns["D", System.Type.Missing]).EntireColumn.AutoFit();
                    ((Range)gSheet.Columns["A:D", System.Type.Missing]).Font.Name = "Arial";

                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Size = 16;
                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Underline = true;
                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)gSheet.Rows[1, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    gSheet.Cells.get_Range("A1", "D1").MergeCells = true;
                    gSheet.Cells[1, 1] = indexTemp.SheetTitle;

                    ((Range)gSheet.Rows[3, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)gSheet.Rows[3, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ((Range)gSheet.Rows[3, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    gSheet.Cells[3, 1] = "RIC";
                    gSheet.Cells[3, 2] = "Display name";
                    gSheet.Cells[3, 3] = "Local language name";
                    gSheet.Cells[3, 4] = "free float rate";

                    int rindex = 4;
                    for (var j = 0; j < indexTemp.FRIList.Count; j++)
                    {
                        var item = indexTemp.FRIList[j] as FloatRateIndex;
                        gSheet.Cells[rindex, 1] = item.Ric;
                        gSheet.Cells[rindex, 2] = item.DisplayName;
                        gSheet.Cells[rindex, 3] = item.LocalLanguageName;
                        gSheet.Cells[rindex, 4] = item.FreeFloatingRate;
                        rindex++;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateFreeFloatRateItemsFileSheet_xls   : " + ex.ToString(), Logger.LogType.Warning);
                return;
            }
        }

        private int GrabDataWithLocalLanguageNameItems(Worksheet wSheet, IndexTemplate indexTemp, int startLine, String str_column)
        {
            startLine = 4;
            List<LocalLanguageNameIndex> LocalNameList = new List<LocalLanguageNameIndex>();
            while (wSheet.get_Range(str_column + startLine, Type.Missing).Value2 != null && wSheet.get_Range(str_column + startLine, Type.Missing).Value2.ToString().Trim() != string.Empty)
            {
                LocalLanguageNameIndex locallanguagenameIndex = new LocalLanguageNameIndex();
                String ticker = wSheet.get_Range(str_column + startLine, Type.Missing).Value2.ToString().Trim();
                if (ticker.Length < 6)
                    for (var i = ticker.Length; i < 6; i++)
                        ticker = "0" + ticker;
                KSorKQListingList kskqlisting = null;
                if (kskqlistingHash.Contains(ticker))
                    kskqlisting = kskqlistingHash[ticker] as KSorKQListingList;
                else
                {
                    startLine++;
                    continue;
                }
                locallanguagenameIndex.Ric = kskqlisting.Ric;
                locallanguagenameIndex.DisplayName = kskqlisting.IDNDisplayName;
                locallanguagenameIndex.LocalLanguageName = str_column.Equals("A") ? wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() : wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim();
                LocalNameList.Add(locallanguagenameIndex);
                startLine++;
            }
            indexTemp.LLNIList = LocalNameList;
            return startLine;
        }

        private void GenerateLocalLanguageNameItemFileSheet_xls(Workbook gBook, int i, IndexTemplate indexTemp, Worksheet gSheet)
        {
            try
            {
                gSheet = (Worksheet)gBook.Worksheets[i];
                gSheet.Name = indexTemp.SheetName;
                if (indexTemp.LLNIList.Count > 0)
                {
                    ((Range)gSheet.Columns["A", System.Type.Missing]).ColumnWidth = 12;
                    ((Range)gSheet.Columns["B", System.Type.Missing]).ColumnWidth = 25;
                    ((Range)gSheet.Columns["C", System.Type.Missing]).ColumnWidth = 25;
                    ((Range)gSheet.Columns["C", System.Type.Missing]).EntireColumn.AutoFit();
                    ((Range)gSheet.Columns["A:C", System.Type.Missing]).Font.Name = "Arial";

                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Size = 16;
                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Underline = true;
                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)gSheet.Rows[1, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ((Range)gSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    gSheet.Cells.get_Range("A1", "C1").MergeCells = true;
                    gSheet.Cells[1, 1] = indexTemp.SheetTitle;

                    ((Range)gSheet.Rows[3, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                    ((Range)gSheet.Rows[3, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ((Range)gSheet.Rows[3, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    gSheet.Cells[3, 1] = "RIC";
                    gSheet.Cells[3, 2] = "Display name";
                    gSheet.Cells[3, 3] = "Local language name";

                    int rindex = 4;
                    for (var j = 0; j < indexTemp.LLNIList.Count; j++)
                    {
                        var item = indexTemp.LLNIList[j] as LocalLanguageNameIndex;
                        gSheet.Cells[rindex, 1] = item.Ric;
                        gSheet.Cells[rindex, 2] = item.DisplayName;
                        gSheet.Cells[rindex, 3] = item.LocalLanguageName;

                        rindex++;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateLocalLanguageNameItemFileSheet_xls   : " + ex.ToString(), Logger.LogType.Warning);
                return;
            }
        }

        private int GrabSectorDataWithAddItemsAndDropItemsFromWorksheet(Worksheet wSheet, IndexTemplate indexTemp, int startLine)
        {
            String columnC4 = wSheet.get_Range("C" + (startLine + 1), Type.Missing).Value2 != null ? wSheet.get_Range("C" + (startLine + 1), Type.Missing).Value2.ToString().Trim() : null;
            String columnD4 = wSheet.get_Range("D" + (startLine + 1), Type.Missing).Value2 != null ? wSheet.get_Range("D" + (startLine + 1), Type.Missing).Value2.ToString().Trim() : null;
            String columnE4 = wSheet.get_Range("E" + (startLine + 1), Type.Missing).Value2 != null ? wSheet.get_Range("E" + (startLine + 1), Type.Missing).Value2.ToString().Trim() : null;
            String columnF4 = wSheet.get_Range("F" + (startLine + 1), Type.Missing).Value2 != null ? wSheet.get_Range("F" + (startLine + 1), Type.Missing).Value2.ToString().Trim() : null;

            if (columnC4 == "종목코드" && columnD4 == "종목명" && columnE4 == "종목코드" && columnF4 == "종목명")
            {
                startLine = 5;
                List<IndexAddItem> alist = new List<IndexAddItem>();
                List<IndexDropItem> dlist = new List<IndexDropItem>();
                List<SectorIndex> slist = new List<SectorIndex>();

                while (wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString() != String.Empty)
                {
                    IndexAddItem aitem = new IndexAddItem(); IndexDropItem ditem = new IndexDropItem(); SectorIndex sitem = null;

                    String cellname = wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                    if (cellname != null)
                    {
                        sitem = new SectorIndex();
                        sitem.Count = 0;
                        if (cellname.Contains("KOSPI"))
                        {
                            cellname = cellname.Replace("KOSPI", "").Trim().ToString();
                            if (KOSPIHash.Contains(cellname.ToUpper().Replace(" ", "")))
                                cellname = "KOSPI " + cellname + " " + ((IndexListingList)KOSPIHash[cellname.ToUpper()]).IndexChainRIC;
                            sitem.SectorName = cellname;
                        }
                        else if (cellname.Contains("KOSDAQ"))
                        {
                            cellname = cellname.Replace("KOSDAQ", "").Trim().ToString();
                            if (KOSDAQHash.Contains(cellname.ToUpper().Replace(" ", "")))
                                cellname = "KOSDAQ " + cellname + " " + ((IndexListingList)KOSDAQHash[cellname.ToUpper()]).IndexChainRIC;
                            sitem.SectorName = cellname;
                        }
                        else if (cellname.Contains("KRX"))
                        {
                            cellname = cellname.Replace("KRX", "").Trim().ToString();
                            if (KRXHash.Contains(cellname.ToUpper().Replace(" ", "")))
                                cellname = "KRX " + cellname + " " + ((IndexListingList)KRXHash[cellname.ToUpper()]).IndexChainRIC;
                            sitem.SectorName = cellname;
                        }
                        else
                        {
                            cellname = cellname.Trim().ToString();
                            if (KOSPIHash.Contains(cellname.ToUpper().Replace(" ", "")))
                                cellname = "KOSPI " + cellname + " " + ((IndexListingList)KOSPIHash[cellname.ToUpper()]).IndexChainRIC;
                            if (KOSDAQHash.Contains(cellname.ToUpper().Replace(" ", "")))
                                cellname = "KOSDAQ " + cellname + " " + ((IndexListingList)KOSDAQHash[cellname.ToUpper()]).IndexChainRIC;
                            if (KRXHash.Contains(cellname.ToUpper().Replace(" ", "")))
                                cellname = "KRX " + cellname + " " + ((IndexListingList)KRXHash[cellname.ToUpper()]).IndexChainRIC;
                            sitem.SectorName = cellname;
                        }

                        int no = startLine + 1;
                        while (wSheet.get_Range("B" + no, Type.Missing).Value2 == null && wSheet.get_Range("C" + no, Type.Missing).Value2 != null)
                        {
                            sitem.Count = ((no - startLine) + 1);
                            if (wSheet.get_Range("C" + (no + 1), Type.Missing).Value2 != null)
                                no++;
                            else
                                break;
                        }
                    }


                    String aticker = wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                    String alocalname = wSheet.get_Range("D" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                    String dticker = wSheet.get_Range("E" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("E" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                    String dlocalname = wSheet.get_Range("F" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("F" + startLine, Type.Missing).Value2.ToString().Trim() : null;

                    if (aticker.Length < 6)
                        for (var x = aticker.Length; x < 6; x++)
                            aticker = "0" + aticker;
                    if (kskqlistingHash.Contains(aticker))
                    {
                        aitem.Ric = ((KSorKQListingList)kskqlistingHash[aticker]).Ric;
                        aitem.DisplayName = ((KSorKQListingList)kskqlistingHash[aticker]).IDNDisplayName;
                    }
                    aitem.LocalLanguageName = alocalname;

                    if (dticker.Length < 6)
                        for (var x = dticker.Length; x < 6; x++)
                            dticker = "0" + dticker;
                    if (kskqlistingHash.Contains(dticker))
                    {
                        ditem.Ric = ((KSorKQListingList)kskqlistingHash[dticker]).Ric;
                        ditem.DisplayName = ((KSorKQListingList)kskqlistingHash[dticker]).IDNDisplayName;
                    }
                    ditem.LocalLanguageName = dlocalname;


                    alist.Add(aitem); dlist.Add(ditem);
                    if (sitem != null)
                        slist.Add(sitem);
                    if (wSheet.get_Range("C" + (startLine + 1), Type.Missing).Value2 != null)
                        startLine++;
                    else
                    {
                        //startLine++;
                        break;
                    }

                }
                indexTemp.AddItemList = alist;
                indexTemp.DropItemList = dlist;
                indexTemp.SIList = slist;
            }
            return startLine;
        }

        private void GenerateSectorAddItemsAndDropItemsFileSheet_xls(Workbook gBook, int i, Worksheet gSheet, IndexTemplate indexTemp)
        {
            gSheet = (Worksheet)gBook.Worksheets[i];
            gSheet.Name = indexTemp.SheetName;
            if (indexTemp.IsDropItem && indexTemp.IsAddItem)
            {
                ((Range)gSheet.Columns["A", System.Type.Missing]).ColumnWidth = 40;
                ((Range)gSheet.Columns["B", System.Type.Missing]).ColumnWidth = 12;
                ((Range)gSheet.Columns["C", System.Type.Missing]).ColumnWidth = 25;
                ((Range)gSheet.Columns["D", System.Type.Missing]).ColumnWidth = 25;
                ((Range)gSheet.Columns["E", System.Type.Missing]).ColumnWidth = 12;
                ((Range)gSheet.Columns["F", System.Type.Missing]).ColumnWidth = 25;
                ((Range)gSheet.Columns["G", System.Type.Missing]).ColumnWidth = 25;
                ((Range)gSheet.Columns["G", System.Type.Missing]).EntireColumn.AutoFit();
                ((Range)gSheet.Columns["A", System.Type.Missing]).VerticalAlignment = XlVAlign.xlVAlignCenter;
                ((Range)gSheet.Columns["A:G", System.Type.Missing]).Font.Name = "Arial";

                ((Range)gSheet.Rows[1, Type.Missing]).Font.Size = 16;
                ((Range)gSheet.Rows[1, Type.Missing]).Font.Underline = true;
                ((Range)gSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)gSheet.Rows[1, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ((Range)gSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                gSheet.Cells.get_Range("A1", "G1").MergeCells = true;
                gSheet.Cells[1, 1] = indexTemp.SheetTitle;

                ((Range)gSheet.Rows[3, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)gSheet.Rows[3, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ((Range)gSheet.Rows[3, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                gSheet.Cells.get_Range("A3", "A4").MergeCells = true;
                gSheet.Cells[3, 1] = "sector";
                gSheet.Cells.get_Range("B3", "D3").MergeCells = true;
                gSheet.Cells[3, 2] = "Add Items";
                gSheet.Cells.get_Range("E3", "G3").MergeCells = true;
                gSheet.Cells[3, 5] = "Drop Items";

                ((Range)gSheet.Rows[4, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)gSheet.Rows[4, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ((Range)gSheet.Rows[4, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                gSheet.Cells[4, 2] = "RIC";
                gSheet.Cells[4, 3] = "Display name";
                gSheet.Cells[4, 4] = "Local language name";
                gSheet.Cells[4, 5] = "RIC";
                gSheet.Cells[4, 6] = "Display name";
                gSheet.Cells[4, 7] = "Local language name";

                int rindex = 5;
                for (var x = 0; x < indexTemp.SIList.Count; x++)
                {
                    if (indexTemp.SIList[x].Count > 0)
                    {
                        gSheet.Cells.get_Range("A" + rindex, "A" + (rindex + indexTemp.SIList[x].Count - 1)).MergeCells = true;
                        gSheet.Cells[rindex, 1] = indexTemp.SIList[x].SectorName;
                        rindex = rindex + indexTemp.SIList[x].Count;
                    }
                    else
                    {
                        gSheet.Cells[rindex, 1] = indexTemp.SIList[x].SectorName;
                        rindex = (rindex + 1);
                    }
                }
                rindex = 5;
                for (var j = 0; j < indexTemp.AddItemList.Count; j++)
                {
                    gSheet.Cells[rindex, 2] = indexTemp.AddItemList[j].Ric;
                    gSheet.Cells[rindex, 3] = indexTemp.AddItemList[j].DisplayName;
                    gSheet.Cells[rindex, 4] = indexTemp.AddItemList[j].LocalLanguageName;
                    gSheet.Cells[rindex, 5] = indexTemp.DropItemList[j].Ric;
                    gSheet.Cells[rindex, 6] = indexTemp.DropItemList[j].DisplayName;
                    gSheet.Cells[rindex, 7] = indexTemp.DropItemList[j].LocalLanguageName;
                    rindex++;
                }
            }
        }

        private int GrabSectorDataWithFreeFloatRateItemsFromWorksheet(Worksheet wSheet, IndexTemplate indexTemp, int startLine)
        {
            startLine = 4;
            List<FloatRateIndex> FRIList = new List<FloatRateIndex>();
            List<SectorIndex> SIList = new List<SectorIndex>();
            while (wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null && wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString() != String.Empty)
            {
                FloatRateIndex friitem = new FloatRateIndex(); SectorIndex siitem = null;
                String cellname = wSheet.get_Range("B" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("B" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                if (cellname != null)
                {
                    siitem = new SectorIndex();
                    siitem.Count = 0;
                    if (cellname.Contains("KOSPI"))
                    {
                        cellname = cellname.Replace("KOSPI", "").Trim().ToString();
                        if (KOSPIHash.Contains(cellname.ToUpper().Replace(" ", "")))
                            cellname = "KOSPI " + cellname + " " + ((IndexListingList)KOSPIHash[cellname.ToUpper()]).IndexChainRIC;
                        siitem.SectorName = cellname;
                    }
                    else if (cellname.Contains("KOSDAQ"))
                    {
                        cellname = cellname.Replace("KOSDAQ", "").Trim().ToString();
                        if (KOSDAQHash.Contains(cellname.ToUpper().Replace(" ", "")))
                            cellname = "KOSDAQ " + cellname + " " + ((IndexListingList)KOSDAQHash[cellname.ToUpper()]).IndexChainRIC;
                        siitem.SectorName = cellname;
                    }
                    else if (cellname.Contains("KRX"))
                    {
                        cellname = cellname.Replace("KRX", "").Trim().ToString();
                        if (KRXHash.Contains(cellname.ToUpper().Replace(" ", "")))
                            cellname = "KRX " + cellname + " " + ((IndexListingList)KRXHash[cellname.ToUpper()]).IndexChainRIC;
                        siitem.SectorName = cellname;
                    }
                    else
                    {
                        cellname = cellname.Trim().ToString();
                        if (KOSPIHash.Contains(cellname.ToUpper().Replace(" ", "")))
                            cellname = "KOSPI " + cellname + " " + ((IndexListingList)KOSPIHash[cellname.ToUpper()]).IndexChainRIC;
                        if (KOSDAQHash.Contains(cellname.ToUpper().Replace(" ", "")))
                            cellname = "KOSDAQ " + cellname + " " + ((IndexListingList)KOSDAQHash[cellname.ToUpper()]).IndexChainRIC;
                        if (KRXHash.Contains(cellname.ToUpper().Replace(" ", "")))
                            cellname = "KRX " + cellname + " " + ((IndexListingList)KRXHash[cellname.ToUpper()]).IndexChainRIC;
                        siitem.SectorName = cellname;
                    }

                    int no = startLine + 1;
                    while (wSheet.get_Range("B" + no, Type.Missing).Value2 == null && wSheet.get_Range("C" + no, Type.Missing).Value2 != null)
                    {
                        siitem.Count = ((no - startLine) + 1);
                        if (wSheet.get_Range("C" + (no + 1), Type.Missing).Value2 != null)
                            no++;
                        else
                            break;
                    }
                }

                String ticker = wSheet.get_Range("C" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("C" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                if (ticker.Length < 6)
                    for (var z = ticker.Length; z < 6; z++)
                        ticker = "0" + ticker;

                friitem.Ric = kskqlistingHash.Contains(ticker) ? ((KSorKQListingList)kskqlistingHash[ticker]).Ric : "can not match the list with the ticker :" + ticker;
                friitem.DisplayName = kskqlistingHash.Contains(ticker) ? ((KSorKQListingList)kskqlistingHash[ticker]).IDNDisplayName : "can not match the list with the ticker :" + ticker;

                String localname = wSheet.get_Range("D" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("D" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                String freerate = wSheet.get_Range("F" + startLine, Type.Missing).Value2 != null ? wSheet.get_Range("F" + startLine, Type.Missing).Value2.ToString().Trim() : null;
                friitem.LocalLanguageName = localname;
                friitem.FreeFloatingRate = freerate;

                if (siitem != null)
                    SIList.Add(siitem);
                FRIList.Add(friitem);
                startLine++;
            }
            indexTemp.SIList = SIList;
            indexTemp.FRIList = FRIList;
            return startLine;
        }

        private void GenerateSectorFreeFloatRateItemsFileSheet_xls(Workbook gBook, int i, Worksheet gSheet, IndexTemplate indexTemp)
        {
            gSheet = (Worksheet)gBook.Worksheets[i];
            gSheet.Name = indexTemp.SheetName;
            if ((indexTemp.FRIList.Count > 0) && (indexTemp.SIList.Count > 0))
            {
                ((Range)gSheet.Columns["A", System.Type.Missing]).ColumnWidth = 25;
                ((Range)gSheet.Columns["B", System.Type.Missing]).ColumnWidth = 12;
                ((Range)gSheet.Columns["C", System.Type.Missing]).ColumnWidth = 25;
                ((Range)gSheet.Columns["D", System.Type.Missing]).ColumnWidth = 25;
                ((Range)gSheet.Columns["E", System.Type.Missing]).ColumnWidth = 20;
                ((Range)gSheet.Columns["E", System.Type.Missing]).EntireColumn.AutoFit();
                ((Range)gSheet.Columns["A", System.Type.Missing]).WrapText = true;
                ((Range)gSheet.Columns["A:E", System.Type.Missing]).Font.Name = "Arial";

                ((Range)gSheet.Rows[1, Type.Missing]).Font.Size = 16;
                ((Range)gSheet.Rows[1, Type.Missing]).Font.Underline = true;
                ((Range)gSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)gSheet.Rows[1, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ((Range)gSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                gSheet.Cells.get_Range("A1", "E1").MergeCells = true;
                gSheet.Cells[1, 1] = indexTemp.SheetTitle;

                ((Range)gSheet.Rows[3, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)gSheet.Rows[3, Type.Missing]).HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ((Range)gSheet.Rows[3, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                gSheet.Cells[3, 1] = "Sector";
                gSheet.Cells[3, 2] = "RIC";
                gSheet.Cells[3, 3] = "Display name";
                gSheet.Cells[3, 4] = "Local language name";
                gSheet.Cells[3, 5] = "free float rate";

                int rindex = 4;

                for (var x = 0; x < indexTemp.SIList.Count; x++)
                {
                    if (indexTemp.SIList[x].Count > 0)
                    {
                        gSheet.Cells.get_Range("A" + rindex, "A" + (rindex + indexTemp.SIList[x].Count - 1)).MergeCells = true;
                        gSheet.Cells[rindex, 1] = indexTemp.SIList[x].SectorName;
                        rindex = rindex + indexTemp.SIList[x].Count;

                    }
                    else
                    {
                        gSheet.Cells[rindex, 1] = indexTemp.SIList[x].SectorName;
                        rindex = (rindex + 1);
                    }
                }

                rindex = 4;
                for (var j = 0; j < indexTemp.FRIList.Count; j++)
                {
                    var item = indexTemp.FRIList[j] as FloatRateIndex;
                    gSheet.Cells[rindex, 2] = item.Ric;
                    gSheet.Cells[rindex, 3] = item.DisplayName;
                    gSheet.Cells[rindex, 4] = item.LocalLanguageName;
                    gSheet.Cells[rindex, 5] = item.FreeFloatingRate;
                    rindex++;
                }
            }
        }

        #endregion
    }

}
