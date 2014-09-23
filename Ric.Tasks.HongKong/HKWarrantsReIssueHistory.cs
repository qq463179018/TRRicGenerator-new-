using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.ComponentModel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Globalization;
using System.Drawing;
using System.Text.RegularExpressions;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    #region [Configuration]
    [ConfigStoredInDB]
    class HKWarrantsReIssueHistoryConfig
    {
        [StoreInDB]
        [Category("GenerateResults")]
        [Description("GeneratedFilePath")]
        public string OutputPath { get; set; }

        [StoreInDB]
        [Category("HKFMAndBulkFileOutPutDir")]
        [Description("same path as task HKFMAndBulkFile")]
        public string InputPath { get; set; }

        [StoreInDB]
        [Category("WarrantIssueType")]
        [DefaultValue("0")]
        [Description("If you want to generate output files from pdf(s) you give. Please choose an announcement type.")]
        public WarrantIssueType WarrantIssueType { get; set; }
    }

    public enum WarrantIssueType { InitialIssue, FurtherIssue }//default from 0,1 ...

    class HKWarrantsReIssueHistory : GeneratorBase
    {
        private static HKWarrantsReIssueHistoryConfig configObj = null;
        private List<IssueAssetAddTemplate> listIAATemplate = new List<IssueAssetAddTemplate>();//generate csv file template
        private string hkIAAddFileName = string.Empty;
        private string hkQAAddFileName = string.Empty;
        private string issueAssAddPath = string.Empty;
        private string futherCBBCFilePath = string.Empty;
        private string futherDWRCFilePath = string.Empty;
        private string issueAssetReIssue = string.Empty;//generate csv 1
        private string wrtQuaHK = string.Empty;//generate csv 2
        private List<WrtQuaNotHK> listQuaNot = new List<WrtQuaNotHK>();
        private string wrtNotHK = string.Empty;//generate csv 3 

        protected override void Initialize()
        {
            configObj = Config as HKWarrantsReIssueHistoryConfig;
            hkIAAddFileName = Path.Combine(configObj.InputPath.Trim(), string.Format(@"HKRicTemplate\YS{0}IAAdd.csv", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyyMMdd")));
            hkQAAddFileName = Path.Combine(configObj.InputPath.Trim(), string.Format(@"HKRicTemplate\YS{0}QAAdd.csv", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyyMMdd")));
            issueAssAddPath = Path.Combine(configObj.OutputPath.Trim(), string.Format("Issue Asset Add_{0}.csv", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyyMMdd")));
            futherCBBCFilePath = Path.Combine(configObj.OutputPath.Trim(), string.Format(@"Download\Futher_dwrc{0}.xls", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyyMMdd")));
            futherDWRCFilePath = Path.Combine(configObj.OutputPath.Trim(), string.Format(@"Download\Futher_cbbc{0}.xls", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyyMMdd")));
            issueAssetReIssue = Path.Combine(configObj.OutputPath.Trim(), string.Format(@"Issue_Asset_ReIssue_{0}.csv", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyyMMdd")));
            wrtQuaHK = Path.Combine(configObj.OutputPath.Trim(), string.Format(@"WRT_QUA_{0}_hongkong.csv", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyyMMdd").ToUpper()));
            wrtNotHK = Path.Combine(configObj.OutputPath.Trim(), string.Format(@"WRT_NOT_{0}_hongkong.csv", DateTime.Now.ToUniversalTime().AddHours(+8).ToString("yyyyMMdd").ToUpper()));
        }
    #endregion

        protected override void Start()
        {
            if (configObj.WarrantIssueType.Equals(WarrantIssueType.InitialIssue))
            {
                if (File.Exists(hkIAAddFileName) && File.Exists(hkQAAddFileName))
                {
                    StartInitialJob();
                    return;
                }
                else
                {
                    string msg = String.Format("can't get files of HK FMAndBulkFileGenerator! hkIAAddFileName:{0}.hkQAAddFileName{1}", hkIAAddFileName, hkQAAddFileName);
                    Logger.Log(msg, Logger.LogType.Error);
                    MessageBox.Show("Run HK FMAndBulkFileGenerator First!");
                }
            }
            else
            {
                StartFurtherJob();
            }

        }

        #region [StartInitialJob]
        private void StartInitialJob()
        {
            try
            {
                GetHongKongCodeToList(listIAATemplate);
            }
            catch (Exception ex)
            {
                string msg = string.Format("get data from file: {0} error,msg: {1} ", hkIAAddFileName, ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                FillInListTemplate(listIAATemplate);
            }
            catch (Exception ex)
            {
                string msg = string.Format("get data from file: {0} error,msg: {1} ", hkQAAddFileName, ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                GenerateFile(listIAATemplate);
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate file error: {0} error,msg: {1} ", listIAATemplate.Count.ToString(), ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        private void GenerateFile(List<IssueAssetAddTemplate> listIAATemplate)
        {
            if (listIAATemplate == null || listIAATemplate.Count == 0)
            {
                string msg = string.Format("listIAATemplate is null or empty!");
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }

            ExcelApp excelApp = new ExcelApp(false, false);

            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, issueAssAddPath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;

                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                FillExcelTitle(wSheet);
                FillExcelBody(wSheet, listIAATemplate);
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
                TaskResultList.Add(new TaskResultEntry("HKWarrantsReIssueHistory", "InitialIssue", issueAssAddPath));
            }
            catch (Exception ex)
            {
                string msg = "Error found in NDA T&C file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void FillExcelBody(Worksheet wSheet, List<IssueAssetAddTemplate> item)
        {
            int startLine = 2;

            foreach (var tmp in item)
            {
                wSheet.Cells[startLine, 1] = tmp.HongKongCode;
                wSheet.Cells[startLine, 2] = tmp.Type;
                wSheet.Cells[startLine, 3] = tmp.Category;
                wSheet.Cells[startLine, 4] = tmp.WarrantIssuer;
                wSheet.Cells[startLine, 5] = tmp.RcsAssetClass;
                wSheet.Cells[startLine, 6] = tmp.WarrantIssueQuantity;
                wSheet.Cells[startLine, 7] = tmp.TranchePrice;
                ((Range)wSheet.Cells[startLine, 8]).NumberFormatLocal = "@";
                wSheet.Cells[startLine, 8] = tmp.TrancheListingDate.ToUpper();

                startLine++;
            }
        }

        private void FillExcelTitle(Worksheet wSheet)
        {
            ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 20;

            wSheet.Cells[1, 1] = "HONG KONG CODE";
            wSheet.Cells[1, 2] = "TYPE";
            wSheet.Cells[1, 3] = "CATEGORY";
            wSheet.Cells[1, 4] = "WARRANT ISSUER";
            wSheet.Cells[1, 5] = "RCS ASSET CLASS";
            wSheet.Cells[1, 6] = "WARRANT ISSUE QUANTITY";
            wSheet.Cells[1, 7] = "TRANCHE PRICE";
            wSheet.Cells[1, 8] = "TRANCHE LISTING DATE";

            ((Range)wSheet.Columns["A:H", System.Type.Missing]).Font.Name = "Arail";
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
        }

        private void FillInListTemplate(List<IssueAssetAddTemplate> listIAATemplate)
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, hkQAAddFileName);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 2, 1, ExcelLineWriter.Direction.Right))
                {
                    string ric = string.Empty;
                    string warrantIssueQuantity = string.Empty;//18
                    string tranche = string.Empty;//17
                    string trancheListingAate = string.Empty;//16

                    for (int i = 1; i <= lastUsedRow; i++)
                    {
                        ric = reader.ReadLineCellText();

                        if (ric == null || ric.Trim() == "")
                            continue;

                        foreach (var item in listIAATemplate)
                        {
                            if (!(item.HongKongCode.Trim() + ".HK").Equals(ric.Trim()))
                                continue;

                            reader.PlaceNext(reader.Row, 16);
                            trancheListingAate = DateTime.Parse(reader.ReadLineCellText()).ToString("dd-MMM-yyyy");
                            reader.PlaceNext(reader.Row, 17);
                            tranche = reader.ReadLineCellText();
                            reader.PlaceNext(reader.Row, 18);
                            warrantIssueQuantity = reader.ReadLineValue2();

                            if (trancheListingAate == null || tranche == null || warrantIssueQuantity == null)
                            {
                                string msg = String.Format("value (row,clo)=({0},{1}) is null!", reader.Row, reader.Col);
                                Logger.Log(msg, Logger.LogType.Error);
                                MessageBox.Show(msg);
                                continue;
                            }

                            item.TrancheListingDate = trancheListingAate;
                            item.TranchePrice = tranche;
                            item.WarrantIssueQuantity = warrantIssueQuantity;
                        }
                        reader.PlaceNext(reader.Row + 1, 1);
                    }
                }
            }
        }

        private void GetHongKongCodeToList(List<IssueAssetAddTemplate> listIAATemplate)
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, hkIAAddFileName);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 2, 1, ExcelLineWriter.Direction.Right))
                {
                    string ric = string.Empty;
                    string type = string.Empty;
                    string category = string.Empty;
                    string rcsAssrtClass = string.Empty;
                    string warrantIssue = string.Empty;

                    for (int i = 1; i <= lastUsedRow; i++)
                    {
                        ric = reader.ReadLineCellText();
                        reader.PlaceNext(reader.Row, 2);
                        type = reader.ReadLineCellText();
                        reader.PlaceNext(reader.Row, 3);
                        category = reader.ReadLineCellText();
                        reader.PlaceNext(reader.Row, 4);
                        rcsAssrtClass = reader.ReadLineCellText();
                        reader.PlaceNext(reader.Row, 5);
                        warrantIssue = reader.ReadLineCellText();

                        if (ric == null || type == null || category == null || rcsAssrtClass == null || warrantIssue == null)
                        {
                            string msg = String.Format("value (row,clo)=({0},{1}) is null!", reader.Row, reader.Col);
                            Logger.Log(msg, Logger.LogType.Error);
                            MessageBox.Show(msg);
                            continue;
                        }

                        //if (string.IsNullOrWhiteSpace(ric))
                        //    continue;

                        IssueAssetAddTemplate iaat = new IssueAssetAddTemplate();
                        iaat.HongKongCode = ric;
                        iaat.Type = type;
                        iaat.Category = category;
                        iaat.RcsAssetClass = rcsAssrtClass;
                        iaat.WarrantIssuer = warrantIssue;
                        listIAATemplate.Add(iaat);
                        reader.PlaceNext(reader.Row + 1, 1);
                    }
                }
            }
        }

        public class IssueAssetAddTemplate
        {
            public string HongKongCode { get; set; }         //HONG KONG CODE
            public string Type { get; set; }                 //TYPE
            public string Category { get; set; }             //CATEGORY
            public string WarrantIssuer { get; set; }        //WARRANT ISSUER
            public string RcsAssetClass { get; set; }        //RCS ASSET CLASS
            public string WarrantIssueQuantity { get; set; } //WARRANT ISSUE QUANTITY
            public string TranchePrice { get; set; }         //TRANCHE PRICE
            public string TrancheListingDate { get; set; }   //TRANCHE LISTING DATE
        }
        #endregion

        #region [StartFutherJop]
        private void StartFurtherJob()
        {
            try
            {
                DownloadFurtherFile();
            }
            catch (Exception ex)
            {
                string msg = string.Format("download file from website error. msg: {0} ", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                listQuaNot = GetQuaNot(futherCBBCFilePath, futherDWRCFilePath);

                if (listQuaNot == null || listQuaNot.Count == 0)
                {
                    string msg = string.Format("No Data On Next Work Day[{0}]", NextWorkDay().ToString("dd-MM-yyyy"));
                    MessageBox.Show(msg);
                    return;
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("get data from download csv error. msg: {0} ", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            GenerateFile(listQuaNot);//three csv

        }

        private void GenerateFile(List<WrtQuaNotHK> listQuaNot)
        {
            if (listQuaNot == null || listQuaNot.Count == 0)
            {
                string msg = string.Format("listIAATemplate is null or empty!");
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }

            try
            {
                GenerateIARI();
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate IARI csv file to local error. msg: {0} ", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }

            try
            {
                GenerateQUA();
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate QUA csv file to local error. msg: {0} ", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }

            try
            {
                GenerateNOT();
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate NOT csv file to local error. msg: {0} ", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        #region NOT
        private void GenerateNOT()
        {
            ExcelApp excelApp = new ExcelApp(false, false);

            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, wrtNotHK);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;

                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                FillExcelTitleNOT(wSheet);
                FillExcelBodyNOT(wSheet, listQuaNot);
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
                TaskResultList.Add(new TaskResultEntry("HKWarrantReIssueHistory", "FutureIssue", wrtNotHK));
            }
            catch (Exception ex)
            {
                string msg = "Error found in NOT file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void FillExcelBodyNOT(Worksheet wSheet, List<WrtQuaNotHK> item)
        {
            int startLine = 2;

            foreach (var tmp in item)
            {
                wSheet.Cells[startLine, 1] = tmp.LogicalKey;
                wSheet.Cells[startLine, 2] = tmp.SecondaryID;
                wSheet.Cells[startLine, 3] = tmp.SecondaryIDType;
                wSheet.Cells[startLine, 4] = tmp.Action;
                wSheet.Cells[startLine, 5] = tmp.Note1Type;
                wSheet.Cells[startLine, 6] = tmp.Note1;

                startLine++;
            }
        }

        private void FillExcelTitleNOT(Worksheet wSheet)
        {
            ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 50; ;

            wSheet.Cells[1, 1] = "Logical_Key";
            wSheet.Cells[1, 2] = "Secondary_ID";
            wSheet.Cells[1, 3] = "Secondary_ID_Type";
            wSheet.Cells[1, 4] = "Action";
            wSheet.Cells[1, 5] = "Note1_Type";
            wSheet.Cells[1, 6] = "Note1";

            ((Range)wSheet.Columns["A:F", System.Type.Missing]).Font.Name = "Arail";
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
        }
        #endregion

        #region QUA
        private void GenerateQUA()
        {
            ExcelApp excelApp = new ExcelApp(false, false);

            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, wrtQuaHK);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;

                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                FillExcelTitleQUA(wSheet);
                FillExcelBodyQUA(wSheet, listQuaNot);
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
                TaskResultList.Add(new TaskResultEntry("HKWarrantsReIssueHistory", "InitialIssue", wrtQuaHK));
            }
            catch (Exception ex)
            {
                string msg = "Error found in QUA file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void FillExcelBodyQUA(Worksheet wSheet, List<WrtQuaNotHK> item)
        {
            int startLine = 2;

            foreach (var tmp in item)
            {
                wSheet.Cells[startLine, 1] = tmp.LogicalKey;
                wSheet.Cells[startLine, 2] = tmp.SecondaryID;
                wSheet.Cells[startLine, 3] = tmp.SecondaryIDType;
                wSheet.Cells[startLine, 4] = tmp.EHIssueQuantity;
                wSheet.Cells[startLine, 5] = tmp.IssueQuantity;

                startLine++;
            }
        }

        private void FillExcelTitleQUA(Worksheet wSheet)
        {
            ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 30;
            ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 20;

            wSheet.Cells[1, 1] = "Logical_Key";
            wSheet.Cells[1, 2] = "Secondary_ID";
            wSheet.Cells[1, 3] = "Secondary_ID_Type";
            wSheet.Cells[1, 4] = "EH_Issue_Quantity";
            wSheet.Cells[1, 5] = "Issue_Quantity";

            ((Range)wSheet.Columns["A:E", System.Type.Missing]).Font.Name = "Arail";
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
        }
        #endregion

        #region [IARI]
        private void GenerateIARI()
        {
            ExcelApp excelApp = new ExcelApp(false, false);

            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, issueAssetReIssue);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;

                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                FillExcelTitleIARI(wSheet);
                FillExcelBodyIARI(wSheet, listQuaNot);
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
                TaskResultList.Add(new TaskResultEntry("HKWarrantReIssueHistory ", "FutrueIssue", issueAssetReIssue));
            }
            catch (Exception ex)
            {
                string msg = "Error found Generate IARA file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void FillExcelBodyIARI(Worksheet wSheet, List<WrtQuaNotHK> item)
        {
            int startLine = 2;

            foreach (var tmp in item)
            {
                wSheet.Cells[startLine, 1] = tmp.ISIN;
                wSheet.Cells[startLine, 2] = tmp.IssueQuantity;
                wSheet.Cells[startLine, 3] = tmp.TranchePrice;
                ((Range)wSheet.Cells[startLine, 4]).NumberFormatLocal = "@";
                wSheet.Cells[startLine, 4] = tmp.TrancheListingDate;

                startLine++;
            }
        }

        private void FillExcelTitleIARI(Worksheet wSheet)
        {
            ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 20;

            wSheet.Cells[1, 1] = "ISIN";
            wSheet.Cells[1, 2] = "WARRANT ISSUE QUANTITY";
            wSheet.Cells[1, 3] = "TRANCHE PRICE";
            wSheet.Cells[1, 4] = "TRANCHE LISTING DATE";

            ((Range)wSheet.Columns["A:D", System.Type.Missing]).Font.Name = "Arail";
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
        }
        #endregion

        #region [Read Download File]
        private List<WrtQuaNotHK> GetQuaNot(string futherCBBCFilePath, string futherDWRCFilePath)
        {
            List<WrtQuaNotHK> list = new List<WrtQuaNotHK>();

            if (File.Exists(futherCBBCFilePath))
            {
                ReadFile(list, futherCBBCFilePath);
            }
            else
            {
                string msg = string.Format("The file {0} is not exist in the config path", futherCBBCFilePath);
                Logger.Log(msg, Logger.LogType.Error);
            }

            if (File.Exists(futherDWRCFilePath))
            {
                ReadFile(list, futherDWRCFilePath);
            }
            else
            {
                string msg = string.Format("The file {0} is not exist in the config path", futherCBBCFilePath);
                Logger.Log(msg, Logger.LogType.Error);
            }
            return list;
        }

        private void ReadFile(List<WrtQuaNotHK> list, string path)
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, path);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 5, 7, ExcelLineWriter.Direction.Right))
                {
                    string code = string.Empty;
                    string listingDate = string.Empty;
                    string issuePrice = string.Empty;
                    string furtherIssueSize = string.Empty;

                    for (int i = 1; i <= lastUsedRow; i++)
                    {
                        listingDate = reader.ReadLineCellText();

                        if (listingDate != null && listingDate.Trim() != "")
                        {
                            if (NextWorkDay().ToString("M/d/yyyy").Equals(listingDate.Trim()))
                            {
                                reader.PlaceNext(reader.Row, 2);
                                code = reader.ReadLineCellText();
                                reader.PlaceNext(reader.Row, 8);
                                issuePrice = reader.ReadLineCellText();
                                reader.PlaceNext(reader.Row, 9);
                                furtherIssueSize = reader.ReadLineCellText().Replace(",", "");

                                if (code == null || issuePrice == null || furtherIssueSize == null)
                                {
                                    string msg = String.Format("value (row,clo)=({0},{1}) is null!", reader.Row, reader.Col);
                                    Logger.Log(msg, Logger.LogType.Error);
                                    MessageBox.Show(msg);
                                    continue;
                                }

                                WrtQuaNotHK wrt = new WrtQuaNotHK();

                                wrt.LogicalKey = list.Count + 1;
                                wrt.SecondaryID = code.Trim();
                                wrt.ISIN = GetISIN(wrt.SecondaryID);
                                wrt.SecondaryIDType = "HONG KONG CODE";
                                wrt.EHIssueQuantity = "N";
                                wrt.IssueQuantity = GetIssueQuantity(wrt.SecondaryID, furtherIssueSize);
                                wrt.Action = "I";
                                wrt.Note1Type = "O";
                                wrt.TranchePrice = issuePrice.Replace("HKD", "").Trim();
                                wrt.Note1 = string.Format("Further issue with issue price HKD {0} on {1}.", issuePrice.Replace("HKD", "").Trim(), NextWorkDay().ToString("dd-MMM-yyyy"));
                                wrt.TrancheListingDate = DateTime.Parse(listingDate).ToString("dd-MMM-yyyy");
                                list.Add(wrt);
                            }
                        }
                        reader.PlaceNext(reader.Row + 1, 7);
                    }
                }
            }
        }
        #endregion

        #region [Format Data]
        private string GetISIN(string str)
        {
            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(str + ".HK", "ISIN_CODE");
            Regex regex = new Regex(@"ISIN_CODE\s+(?<value>[0-9A-Z]+)");
            MatchCollection matches = regex.Matches(response);

            if (matches.Count > 0)
                return matches[0].Groups["value"].Value;
            else
                return "No ISIN In IDN";
        }

        private string GetIssueQuantity(string str, string qua)
        {
            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(str + ".HK", "AMT_ISSUE");
            Regex regex = new Regex(@"AMT_ISSUE\s+(?<value>\d+)");
            MatchCollection matches = regex.Matches(response);

            if (matches.Count > 0)
                return (Convert.ToInt32(qua) + Convert.ToInt32(matches[0].Groups["value"].Value)).ToString();
            else
                return str + "+No Issue In IDN";
        }

        private DateTime NextWorkDay()
        {
            if (DateTime.Now.ToUniversalTime().AddHours(+8).DayOfWeek.ToString() == "Saturday")
            {
                return DateTime.Now.ToUniversalTime().AddDays(+3).AddHours(+8);
            }

            if (DateTime.Now.ToUniversalTime().AddHours(+8).DayOfWeek.ToString() == "Sunday")
            {
                return DateTime.Now.ToUniversalTime().AddDays(+2).AddHours(+8);
            }

            return DateTime.Now.ToUniversalTime().AddDays(+1).AddHours(+8);
        }
        #endregion

        #region [Download File]
        private void DownloadFurtherFile()
        {
            WebClientUtil.DownloadFile("http://www.hkex.com.hk/eng/cbbc/furtherissue/fissue.xls", 1000, futherCBBCFilePath);
            WebClientUtil.DownloadFile("http://www.hkex.com.hk/eng/dwrc/furtherissue/fissue.xls", 1000, futherDWRCFilePath);
        }
        #endregion

        public class WrtQuaNotHK
        {
            public int LogicalKey { get; set; }             //Logical_Key	
            public string SecondaryID { get; set; }           //Secondary_ID	
            public string SecondaryIDType { get; set; }       //Secondary_ID_Type
            public string EHIssueQuantity { get; set; }       //EH_Issue_Quantity
            public string IssueQuantity { get; set; }         //Issue_Quantity
            public string Action { get; set; }             //Action	
            public string Note1Type { get; set; }         //Note1_Type	
            public string Note1 { get; set; }              //Note1
            public string ISIN { get; set; }
            public string TranchePrice { get; set; }          //TRANCHE PRICE	
            public string TrancheListingDate { get; set; }          //TRANCHE LISTING DATE
        }
        #endregion
    }
}