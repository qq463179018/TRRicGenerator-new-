using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks
{
    #region used for KOREA_IndexGeneratorConfig
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaIndexReadFileConfig
    {
        public String WorkbookPath { get; set; }
        public String KospiSheetname { get; set; }
        public String KosdaqSheetname { get; set; }
        public String KrxSheetname { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaIndexOrignalFileConfig
    {
        public String WorkbookPath { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaIndexGeneratorFileConfig
    {
        public String WorkbookPath { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaIndexGeneratorConfig
    {
        public KoreaIndexReadFileConfig KoreaIndexReadFileConfig { get; set; }
        public KoreaKQorKsListReadFilePathConfig KoreaKQorKsListReadFilePathConfig { get; set; }
        public KoreaIndexOrignalFileConfig KoreaIndexOrignalFileConfig { get; set; }
        public KoreaIndexGeneratorFileConfig KoreaIndexGeneratorFileConfig { get; set; }
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
        private Hashtable kskqlistingHash = new Hashtable();
        private Hashtable KOSPIHash = new Hashtable();
        private Hashtable KOSDAQHash = new Hashtable();
        private Hashtable KRXHash = new Hashtable();
        private KoreaIndexGeneratorConfig configObj;

        protected override void Start()
        {
            StartIndexGeneratorJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KoreaIndexGeneratorConfig;
        }

        public void StartIndexGeneratorJob()
        {
            ReadDataFromKSandKQListingEquityList_xls();
            ReadDataFromIndexListingList_xls();
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
                string ipath = configObj.KoreaKQorKsListReadFilePathConfig.WorkbookPath;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = ExcelUtil.GetWorksheet(configObj.KoreaKQorKsListReadFilePathConfig.WorksheetName, wBook);
                if (wSheet == null)
                {
                    Logger.Log("Excel Worksheet could not be created! please check the referenec is correct!", Logger.LogType.Warning);
                    return;
                }
                int startLine = 2;
                while (wSheet.Range["A" + startLine, Type.Missing].Value2 != null)
                {
                    if (wSheet.Range["A" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                    {
                        KSorKQListingList listing = new KSorKQListingList();
                        if (wSheet.Range["A" + startLine, Type.Missing].Value2 != null)
                            listing.Ric = ((Range)wSheet.Cells[startLine, 1]).Value2.ToString().Trim();
                        if (wSheet.Range["B" + startLine, Type.Missing].Value2 != null)
                            listing.IDNDisplayName = ((Range)wSheet.Cells[startLine, 2]).Value2.ToString().Trim();
                        if (wSheet.Range["C" + startLine, Type.Missing].Value2 != null)
                            listing.ISIN = ((Range)wSheet.Cells[startLine, 3]).Value2.ToString().Trim();
                        String ticker;
                        if (listing.Ric.IndexOf('.') > 0 && listing.Ric.Length == 9)
                        {
                            ticker = listing.Ric.Split('.')[0].Trim().ToString();
                            kskqlistingHash.Add(ticker, listing);
                            startLine++;
                        }
                        else
                        {
                            startLine++;
                        }
                    }
                }
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in ReadDataFromEquityMasterfile_xls  : " + ex, Logger.LogType.Warning);
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
                string ipath = configObj.KoreaIndexReadFileConfig.WorkbookPath;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet_KOSPI = ExcelUtil.GetWorksheet(configObj.KoreaIndexReadFileConfig.KospiSheetname, wBook);
                Worksheet wSheet_KOSDAQ = ExcelUtil.GetWorksheet(configObj.KoreaIndexReadFileConfig.KosdaqSheetname, wBook);
                Worksheet wSheet_KRX = ExcelUtil.GetWorksheet(configObj.KoreaIndexReadFileConfig.KrxSheetname, wBook);
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
                Logger.Log("Error found in ReadDataFromIndexListingList_xls  : " + ex, Logger.LogType.Warning);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private int GenerateKOSPIHash(_Worksheet wSheet_KOSPI, int startLine)
        {
            try
            {
                while (wSheet_KOSPI.Range["A" + startLine, Type.Missing].Value2 != null && wSheet_KOSPI.Range["A" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                {
                    IndexListingList indexlisting = new IndexListingList
                    {
                        KoreanNameForChainRIC = wSheet_KOSPI.Range["A" + startLine, Type.Missing].Value2.ToString()
                    };
                    if (wSheet_KOSPI.Range["B" + startLine, Type.Missing].Value2 != null && wSheet_KOSPI.Range["B" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.EnglishNameForChainRIC = wSheet_KOSPI.Range["B" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSPI.Range["C" + startLine, Type.Missing].Value2 != null && wSheet_KOSPI.Range["C" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.IndexChainRIC = wSheet_KOSPI.Range["C" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSPI.Range["D" + startLine, Type.Missing].Value2 != null && wSheet_KOSPI.Range["D" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.Number = wSheet_KOSPI.Range["D" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSPI.Range["E" + startLine, Type.Missing].Value2 != null && wSheet_KOSPI.Range["E" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.One = wSheet_KOSPI.Range["E" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSPI.Range["F" + startLine, Type.Missing].Value2 != null && wSheet_KOSPI.Range["F" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.Two = wSheet_KOSPI.Range["F" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSPI.Range["G" + startLine, Type.Missing].Value2 != null && wSheet_KOSPI.Range["G" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.Three = wSheet_KOSPI.Range["G" + startLine, Type.Missing].Value2.ToString();

                    if (!string.IsNullOrEmpty(indexlisting.One) && (!KOSPIHash.Contains(indexlisting.One)))
                        KOSPIHash.Add(indexlisting.One.ToUpper(), indexlisting);
                    if (!string.IsNullOrEmpty(indexlisting.Two) && (!KOSPIHash.Contains(indexlisting.Two)))
                        KOSPIHash.Add(indexlisting.Two.ToUpper(), indexlisting);
                    if (indexlisting.Three != null && indexlisting.Two != String.Empty && (!KOSPIHash.Contains(indexlisting.Three)))
                        KOSPIHash.Add(indexlisting.Three.ToUpper(), indexlisting);
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateKOSPIHash  : " + ex, Logger.LogType.Error);
            }
            return startLine;
        }

        private int GenerateKOSDAQHash(_Worksheet wSheet_KOSDAQ, int startLine)
        {
            startLine = 2;
            try
            {

                while (wSheet_KOSDAQ.Range["A" + startLine, Type.Missing].Value2 != null && wSheet_KOSDAQ.Range["A" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                {
                    IndexListingList indexlisting = new IndexListingList
                    {
                        KoreanNameForChainRIC = wSheet_KOSDAQ.Range["A" + startLine, Type.Missing].Value2.ToString()
                    };
                    if (wSheet_KOSDAQ.Range["B" + startLine, Type.Missing].Value2 != null && wSheet_KOSDAQ.Range["B" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.EnglishNameForChainRIC = wSheet_KOSDAQ.Range["B" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSDAQ.Range["C" + startLine, Type.Missing].Value2 != null && wSheet_KOSDAQ.Range["C" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.IndexChainRIC = wSheet_KOSDAQ.Range["C" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSDAQ.Range["D" + startLine, Type.Missing].Value2 != null && wSheet_KOSDAQ.Range["D" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.Number = wSheet_KOSDAQ.Range["D" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSDAQ.Range["E" + startLine, Type.Missing].Value2 != null && wSheet_KOSDAQ.Range["E" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.One = wSheet_KOSDAQ.Range["E" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSDAQ.Range["F" + startLine, Type.Missing].Value2 != null && wSheet_KOSDAQ.Range["F" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.Two = wSheet_KOSDAQ.Range["F" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KOSDAQ.Range["G" + startLine, Type.Missing].Value2 != null && wSheet_KOSDAQ.Range["G" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.Three = wSheet_KOSDAQ.Range["G" + startLine, Type.Missing].Value2.ToString();

                    if (!string.IsNullOrEmpty(indexlisting.One) && (!KOSDAQHash.Contains(indexlisting.One)))
                        KOSDAQHash.Add(indexlisting.One.ToUpper(), indexlisting);
                    if (!string.IsNullOrEmpty(indexlisting.Two) && (!KOSDAQHash.Contains(indexlisting.Two)))
                        KOSDAQHash.Add(indexlisting.Two.ToUpper(), indexlisting);
                    if (!string.IsNullOrEmpty(indexlisting.Three) && (!KOSDAQHash.Contains(indexlisting.Three)))
                        KOSDAQHash.Add(indexlisting.Three.ToUpper(), indexlisting);
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateKOSDAQHash  : " + ex, Logger.LogType.Error);
            }
            return startLine;
        }

        private int GenerateKRXHash(_Worksheet wSheet_KRX, int startLine)
        {
            startLine = 2;
            try
            {
                while (wSheet_KRX.Range["A" + startLine, Type.Missing].Value2 != null && wSheet_KRX.Range["A" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                {
                    IndexListingList indexlisting = new IndexListingList
                    {
                        KoreanNameForChainRIC = wSheet_KRX.Range["A" + startLine, Type.Missing].Value2.ToString()
                    };
                    if (wSheet_KRX.Range["B" + startLine, Type.Missing].Value2 != null && wSheet_KRX.Range["B" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.EnglishNameForChainRIC = wSheet_KRX.Range["B" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KRX.Range["C" + startLine, Type.Missing].Value2 != null && wSheet_KRX.Range["C" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.IndexChainRIC = wSheet_KRX.Range["C" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KRX.Range["D" + startLine, Type.Missing].Value2 != null && wSheet_KRX.Range["D" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.Number = wSheet_KRX.Range["D" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KRX.Range["E" + startLine, Type.Missing].Value2 != null && wSheet_KRX.Range["E" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.One = wSheet_KRX.Range["E" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KRX.Range["F" + startLine, Type.Missing].Value2 != null && wSheet_KRX.Range["F" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.Two = wSheet_KRX.Range["F" + startLine, Type.Missing].Value2.ToString();
                    if (wSheet_KRX.Range["G" + startLine, Type.Missing].Value2 != null && wSheet_KRX.Range["G" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                        indexlisting.Three = wSheet_KRX.Range["G" + startLine, Type.Missing].Value2.ToString();

                    if (!string.IsNullOrEmpty(indexlisting.One) && (!KRXHash.Contains(indexlisting.One)))
                        KRXHash.Add(indexlisting.One.ToUpper(), indexlisting);
                    if (!string.IsNullOrEmpty(indexlisting.Two) && (!KRXHash.Contains(indexlisting.Two)))
                        KRXHash.Add(indexlisting.Two.ToUpper(), indexlisting);
                    if (!string.IsNullOrEmpty(indexlisting.Three) && (!KRXHash.Contains(indexlisting.Three)))
                        KRXHash.Add(indexlisting.Three.ToUpper(), indexlisting);
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in GenerateKRXHash  : " + ex, Logger.LogType.Error);
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
                String ipath = configObj.KoreaIndexOrignalFileConfig.WorkbookPath;
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
                        String str_name = sname_arr[0].Trim();
                        String sname = str_name.Contains("KOSPI") ? str_name.Replace("KOSPI", "") : (str_name.Contains("KOSDAQ") ? str_name.Replace("KOSDAQ", "") : (str_name.Contains("KRX") ? str_name.Replace("KRX", "") : str_name));
                        String str_instrument = sname_arr[(sname_arr.Length - 1)].Trim(new[] { ' ', ')' });
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
                            temp_str_name = str_name.Contains("KOSPI") ? str_name.Replace("KOSPI", "").Trim().ToUpper().Replace(" ", "") : str_name.Replace("코스피", "").Trim().ToUpper().Replace(" ", "");
                            if (KOSPIHash.Contains(temp_str_name))
                                chainRic = ((IndexListingList)KOSPIHash[temp_str_name]).IndexChainRIC;
                        }
                        else if (str_name.Contains("KOSDAQ") || str_name.Contains("코스닥"))
                        {
                            temp_str_name = str_name.Contains("KOSDAQ") ? str_name.Replace("KOSDAQ", "").Trim().ToUpper().Replace(" ", "") : str_name.Replace("코스닥", "").Trim().ToUpper().Replace(" ", "");
                            if (KOSDAQHash.Contains(temp_str_name))
                                chainRic = ((IndexListingList)KOSDAQHash[temp_str_name]).IndexChainRIC;
                        }
                        else if (str_name.Contains("KRX"))
                        {
                            temp_str_name = str_name.Replace("KRX", "").Trim().ToUpper().Replace(" ", "");
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
                        if (wSheet.Range["A" + startLine, Type.Missing].Value2 == null && wSheet.Range["B" + startLine, Type.Missing].Value2 != null)
                            ((Range)wSheet.Columns["A", Type.Missing]).Delete(XlDeleteShiftDirection.xlShiftToLeft);

                        String title = wSheet.Range["A1", Type.Missing].Value2.ToString();
                        title = wSheet.Name + " " + chainRic;
                        wSheet.Cells[1, 1] = title;

                        String columnA3 = wSheet.Range["A" + (startLine - 1), Type.Missing].Value2 != null ? wSheet.Range["A" + (startLine - 1), Type.Missing].Value2.ToString() : null;
                        String columnB3 = wSheet.Range["B" + (startLine - 1), Type.Missing].Value2 != null ? wSheet.Range["B" + (startLine - 1), Type.Missing].Value2.ToString() : null;
                        String columnC3 = wSheet.Range["C" + (startLine - 1), Type.Missing].Value2 != null ? wSheet.Range["C" + (startLine - 1), Type.Missing].Value2.ToString() : null;
                        String columnD3 = wSheet.Range["D" + (startLine - 1), Type.Missing].Value2 != null ? wSheet.Range["D" + (startLine - 1), Type.Missing].Value2.ToString() : null;

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
                    }
                    else
                    {
                        sheetname = sheetname.Replace("섹터", "Sector");  //KRX Sector(변경종목)
                        String chainRic = "";
                        String[] sname_arr = sheetname.Split('(');
                        String str_name = sname_arr[0].Trim();
                        String str_instrument = sname_arr[(sname_arr.Length - 1)].Trim(new[] { ' ', ')' });
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
                        if (wSheet.Range["A" + startLine, Type.Missing].Value2 == null && wSheet.Range["B" + startLine, Type.Missing].Value2 != null && wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim() == "섹터구분")
                            ((Range)wSheet.Columns["A", Type.Missing]).Delete(XlDeleteShiftDirection.xlShiftToLeft);

                        String columnA3 = wSheet.Range["A" + startLine, Type.Missing].Value2 != null ? wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim() : null;
                        String columnB3 = wSheet.Range["B" + startLine, Type.Missing].Value2 != null ? wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim() : null;
                        String columnC3 = wSheet.Range["C" + startLine, Type.Missing].Value2 != null ? wSheet.Range["C" + startLine, Type.Missing].Value2.ToString().Trim() : null;
                        String columnD3 = wSheet.Range["D" + startLine, Type.Missing].Value2 != null ? wSheet.Range["D" + startLine, Type.Missing].Value2.ToString().Trim() : null;
                        String columnE3 = wSheet.Range["E" + startLine, Type.Missing].Value2 != null ? wSheet.Range["E" + startLine, Type.Missing].Value2.ToString().Trim() : null;

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
                Logger.Log("" + ex, Logger.LogType.Warning);
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
        private void ModifyTheSectorDataWithFreeFloatRate_Worksheet(_Worksheet wSheet, ref String chainRic, ref int startLine)
        {
            wSheet.Cells[3, 1] = "sector";
            wSheet.Cells[3, 2] = "Ric";
            ((Range)wSheet.Columns["C", Type.Missing]).Insert(XlInsertShiftDirection.xlShiftToRight, null);
            wSheet.Cells[3, 3] = "Display Name";
            wSheet.Cells[3, 4] = "Local Language Name";
            ((Range)wSheet.Columns["E", Type.Missing]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
            wSheet.Cells[3, 5] = "Free Float Rate(%)";

            startLine = 4;
            while (wSheet.Range["B" + startLine, Type.Missing].Value2 != null && wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
            {
                String sectorname = wSheet.Range["A" + startLine, Type.Missing].Value2 != null ? wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim() : null;
                if (sectorname != null)
                {
                    String parameter = sectorname.Contains("(") ? sectorname.Split('(')[0].Trim() : sectorname;
                    if (parameter != String.Empty)
                        chainRic = GetSectorChainRic(parameter);
                    wSheet.Cells[startLine, 1] = parameter + " " + chainRic;
                }

                String ticker = wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim();
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
        private void ModifyTheSectorDataWithAddItemsAndDropItems_Worksheet(_Worksheet wSheet, ref String chainRic, ref int startLine)
        {
            wSheet.Cells[3, 1] = "sector";
            wSheet.Cells[3, 2] = "Add Items";
            wSheet.Cells[3, 4] = "Drop Items";

            String columnA4 = wSheet.Range["A" + (startLine + 1), Type.Missing].Value2 != null ? wSheet.Range["A" + (startLine + 1), Type.Missing].Value2.ToString().Trim() : null;
            String columnB4 = wSheet.Range["B" + (startLine + 1), Type.Missing].Value2 != null ? wSheet.Range["B" + (startLine + 1), Type.Missing].Value2.ToString().Trim() : null;
            String columnC4 = wSheet.Range["C" + (startLine + 1), Type.Missing].Value2 != null ? wSheet.Range["C" + (startLine + 1), Type.Missing].Value2.ToString().Trim() : null;
            String columnD4 = wSheet.Range["D" + (startLine + 1), Type.Missing].Value2 != null ? wSheet.Range["D" + (startLine + 1), Type.Missing].Value2.ToString().Trim() : null;
            String columnE4 = wSheet.Range["E" + (startLine + 1), Type.Missing].Value2 != null ? wSheet.Range["E" + (startLine + 1), Type.Missing].Value2.ToString().Trim() : null;

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
                while (wSheet.Range["B" + startLine, Type.Missing].Value2 != null && wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
                {
                    String sectorname = wSheet.Range["A" + startLine, Type.Missing].Value2 != null ? wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim() : null;
                    if (sectorname != null)
                    {
                        String parameter = sectorname.Contains("(") ? sectorname.Split('(')[0].Trim() : sectorname;
                        if (parameter != String.Empty)
                            chainRic = GetSectorChainRic(parameter);
                        wSheet.Cells[startLine, 1] = parameter + " " + chainRic;
                    }

                    String aticker = wSheet.Range["B" + startLine, Type.Missing].Value2 != null ? wSheet.Range["B" + startLine, Type.Missing].Value2.ToString().Trim() : null;
                    if (aticker.Length < 6)
                        for (var x = aticker.Length; x < 6; x++)
                            aticker = "0" + aticker;

                    if (kskqlistingHash.Contains(aticker))
                    {
                        var aitem = kskqlistingHash[aticker] as KSorKQListingList;
                        wSheet.Cells[startLine, 2] = aitem.Ric;
                        wSheet.Cells[startLine, 3] = aitem.IDNDisplayName;
                    }

                    String dticker = wSheet.Range["E" + startLine, Type.Missing].Value2 != null ? wSheet.Range["E" + startLine, Type.Missing].Value2.ToString().Trim() : null;
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
            String tempStrName;
            if (str_name.Contains("KOSPI") || str_name.Contains("코스피"))
            {
                tempStrName = str_name.Contains("KOSPI") ? str_name.Replace("KOSPI", "").Trim().ToUpper().Replace(" ", "") : str_name.Replace("코스피", "").Trim().ToUpper().Replace(" ", "");
                if (KOSPIHash.Contains(tempStrName))
                    chainRic = ((IndexListingList)KOSPIHash[tempStrName]).IndexChainRIC;
            }
            else if (str_name.Contains("KOSDAQ") || str_name.Contains("코스닥"))
            {
                tempStrName = str_name.Contains("KOSDAQ") ? str_name.Replace("KOSDAQ", "").Trim().ToUpper().Replace(" ", "") : str_name.Replace("코스닥", "").Trim().ToUpper().Replace(" ", "");
                if (KOSDAQHash.Contains(tempStrName))
                    chainRic = ((IndexListingList)KOSDAQHash[tempStrName]).IndexChainRIC;
            }
            else if (str_name.Contains("KRX"))
            {
                tempStrName = str_name.Replace("KRX", "").Trim().ToUpper().Replace(" ", "");
                if (KRXHash.Contains(tempStrName))
                    chainRic = ((IndexListingList)KRXHash[tempStrName]).IndexChainRIC;
            }
            else
            {
                tempStrName = str_name.ToUpper().Replace(" ", "");
                if (chainRic == string.Empty && KOSPIHash.Contains(tempStrName))
                    chainRic = ((IndexListingList)KOSPIHash[tempStrName]).IndexChainRIC;
                else if (chainRic == string.Empty && KOSDAQHash.Contains(tempStrName))
                    chainRic = ((IndexListingList)KOSDAQHash[tempStrName]).IndexChainRIC;
                else if (chainRic == string.Empty && KRXHash.Contains(tempStrName))
                    chainRic = ((IndexListingList)KRXHash[tempStrName]).IndexChainRIC;
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
        private int ModifyTheDataWithAddItemsAndDropItems_Worksheet(_Worksheet wSheet, int startLine)
        {
            wSheet.Cells[3, 1] = "Add Items";
            wSheet.Cells[3, 3] = "Drop Items";
            String columnA4 = wSheet.Range["A" + startLine, Type.Missing].Value2 != null ? wSheet.Range["A" + startLine, Type.Missing].Value2.ToString() : null;
            String columnB4 = wSheet.Range["B" + startLine, Type.Missing].Value2 != null ? wSheet.Range["B" + startLine, Type.Missing].Value2.ToString() : null;
            String columnC4 = wSheet.Range["C" + startLine, Type.Missing].Value2 != null ? wSheet.Range["C" + startLine, Type.Missing].Value2.ToString() : null;
            String columnD4 = wSheet.Range["D" + startLine, Type.Missing].Value2 != null ? wSheet.Range["D" + startLine, Type.Missing].Value2.ToString() : null;

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

                while (wSheet.Range["A" + startLine, Type.Missing].Value2 != null && wSheet.Range["A" + startLine, Type.Missing].Value2.ToString() != String.Empty)
                {
                    String aticker = wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim();
                    if (aticker.Length < 6)
                        for (var x = aticker.Length; x < 6; x++)
                            aticker = "0" + aticker;
                    if (kskqlistingHash.Contains(aticker))
                    {
                        var aitem = kskqlistingHash[aticker] as KSorKQListingList;
                        wSheet.Cells[startLine, 1] = aitem.Ric;
                        wSheet.Cells[startLine, 2] = aitem.IDNDisplayName;
                    }

                    String dticker = wSheet.Range["D" + startLine, Type.Missing].Value2.ToString().Trim();
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
            while (wSheet.Range["A" + startLine, Type.Missing].Value2 != null && wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
            {
                String ticker = wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim();
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
            while (wSheet.Range["A" + startLine, Type.Missing].Value2 != null && wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim() != String.Empty)
            {
                String ticker = wSheet.Range["A" + startLine, Type.Missing].Value2.ToString().Trim();
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
    }

}
