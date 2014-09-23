using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.ComponentModel;
using Ric.Util;
using Ric.Core;

namespace Ric.Tasks.HongKong
{
    public class HKISINUpdateConfig
    {
        public string DOWNLOAD_FILE_DIR { get; set; }
        public string ISIN_RECORD_FILE_PATH { get; set; }
        public string RIC_CONVS_TEMPLATE_V2_FILE_PATH { get; set; }
        public string WORKSHEET_NAME_TEMPLATE { get; set; }
        [Description("Date format should be like \"31Jan2012\"")]
        public string DATE { get; set; }
        public string LOG_FILE_PATH { get; set; }
    }


    //Used to describe the ric change events which listed in the file "DownloadRicChangeEvents_"
    public class HKRicChange
    {
        public string EffectiveDate { get; set; }
        public string RicWas { get; set; }
        public string RicNow { get; set; }
        public string ChangeType { get; set; }
        public string Country { get; set; }
        public string Exchanges { get; set; }
        public string AssetClass { get; set; }
        public string DescriptionWas { get; set; }
        public string DescriptionNow { get; set; }
        public string SummaryOfChange { get; set; }
        public string ISINWas { get; set; }
        public string ISINNow { get; set; }
        public string SecondMarketId { get; set; }
        public string SecondMarketWas { get; set; }
        public string SecondMarketNow { get; set; }
        public string RicSeqId { get; set; }
        public string ISIN { get; set; }

        public HKRicChange(string isin)
        {
            this.ISIN = isin;
        }

        public HKRicChange()
        { }
    }

    public class HKISINUpdator : GeneratorBase
    {
        private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_ISINUpdate.config";
        //private static readonly string DOWNLOAD_FILE_NAME_PREFIC = "DownloadRicChangeEvents";
        private static Dictionary<string, HKRicChange> addRicDic = new Dictionary<string, HKRicChange>();
        private static HKISINUpdateConfig configObj = null;
        //private static Logger logger = null;

        protected override void Start()
        {
            StartISINUpdate();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(HKISINUpdateConfig)) as HKISINUpdateConfig;
            //logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.New);
        }

        public void StartISINUpdate()
        {
            getAddRics();
            if (addRicDic.Keys.Count == 0)
            {
                Logger.Log("There's no ISIN needed to be updated!", Logger.LogType.Warning);
                return;
            }
            updateRicAddDicFromDownloadFile();
            updateTemplateV2File();
            addRicDic.Clear();
        }

        private void updateTemplateV2File()
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.RIC_CONVS_TEMPLATE_V2_FILE_PATH);
                Worksheet worksheet = ExcelUtil.GetWorksheet(configObj.WORKSHEET_NAME_TEMPLATE, workbook);
                if (worksheet == null)
                {
                    LogMessage(string.Format("There's no worksheet {0} in the file {1}", configObj.WORKSHEET_NAME_TEMPLATE, workbook.Name));
                }

                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 3, 1, ExcelLineWriter.Direction.Right))
                {
                    foreach (string key in addRicDic.Keys)
                    {
                        writer.WriteLine("Revise");//EventAction
                        writer.WriteLine(key.Remove(key.IndexOf('.')));//Ric SeqId
                        writer.WriteLine("Add"); //Change Type
                        writer.WriteLine(DateTime.Parse(addRicDic[key].EffectiveDate).ToString("ddMMMyy"));// Date
                        writer.WriteLine(""); //Description Was
                        writer.WriteLine(addRicDic[key].DescriptionNow); //Description Now
                        writer.WriteLine(""); //Ric Was
                        writer.WriteLine(addRicDic[key].RicNow);//RICNow
                        writer.WriteLine("");//ISINWas
                        writer.WriteLine(addRicDic[key].ISIN); //ISINNow
                        writer.WriteLine("Official Code");//2ndId
                        writer.WriteLine("");//2ndWas
                        writer.WriteLine(addRicDic[key].SecondMarketNow);//2ndNow
                        writer.WriteLine("");//ThomsonWas
                        writer.WriteLine("");//ThomsonNow
                        writer.WriteLine("");
                        writer.WriteLine(addRicDic[key].Exchanges);//Exchange
                        writer.WriteLine(addRicDic[key].AssetClass);//Asset
                        writer.PlaceNext(writer.Row + 1, 1);
                    }
                }

                //Run Macro
                worksheet.Activate();
                app.ExcelAppInstance.GetType().InvokeMember("Run",
                BindingFlags.Default | BindingFlags.InvokeMethod,
                null,
                app.ExcelAppInstance,
                new object[] { "FormatData" });

                string targetFileName = Path.Combine(Path.GetDirectoryName(configObj.RIC_CONVS_TEMPLATE_V2_FILE_PATH), configObj.DATE);
                targetFileName += Path.GetFileName(configObj.RIC_CONVS_TEMPLATE_V2_FILE_PATH);
                workbook.SaveCopyAs(targetFileName);
                workbook.Close(false, workbook.FullName, false);
                AddResult("Target file ", targetFileName, "The file has been updated with ISIN");
            }
        }
        private void getAddRics()
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.ISIN_RECORD_FILE_PATH);
                Range dateRange = null;
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    var worksheet = (Worksheet)workbook.Worksheets[i];
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                    // For SMF worksheet
                    if (i == 1 || i == 3)
                    {
                        while (true)
                        {
                            dateRange = ExcelUtil.GetRange(lastUsedRow, 4, worksheet);
                            if (dateRange.Text != null && dateRange.Text.ToString().Trim() != string.Empty)
                            {
                                if (DateTime.Parse(dateRange.Text.ToString().Trim()) == DateTime.Parse(configObj.DATE))
                                //if (DateTime.Parse(dateRange.Text.ToString().Trim()) == DateTime.Parse("2012/01/12"))
                                {
                                    addRicDic.Add(ExcelUtil.GetRange(lastUsedRow, 1, worksheet).Text.ToString(), new HKRicChange(ExcelUtil.GetRange(lastUsedRow, 3, worksheet).Text.ToString()));
                                    lastUsedRow--;
                                }

                                else
                                    break;
                            }
                            else
                                break;
                        }
                    }
                    else
                    {

                        while (true)
                        {
                            dateRange = ExcelUtil.GetRange(lastUsedRow, 3, worksheet);
                            if (dateRange.Text != null && dateRange.Text.ToString().Trim() != string.Empty)
                            {
                                DateTime date = DateTime.MinValue;
                                date = DateTime.Parse(dateRange.Text.ToString().Trim());
                                //if (i == 4 || i == 6)
                                //{
                                //    date = DateTime.ParseExact(dateRange.Text.ToString().Trim(), "dd/MM/yyyy", null);
                                //}

                                //else
                                //{
                                //    date = DateTime.Parse(dateRange.Text.ToString().Trim());
                                //}
                                if (date == DateTime.Parse(configObj.DATE))
                                //if (date == DateTime.Parse("2012/01/12"))
                                {
                                    addRicDic.Add(ExcelUtil.GetRange(lastUsedRow, 1, worksheet).Text.ToString(), new HKRicChange(ExcelUtil.GetRange(lastUsedRow, 2, worksheet).Text.ToString()));
                                    lastUsedRow--;
                                }
                                else
                                    break;
                            }
                            else
                                break;
                        }
                    }
                }
            }
        }
        private void updateRicAddDicFromDownloadFile()
        {
            string downloadFilePath = string.Empty;
            if (Directory.Exists(configObj.DOWNLOAD_FILE_DIR))
            {
                string[] files = Directory.GetFiles(configObj.DOWNLOAD_FILE_DIR, "DownloadRicChangeEvents*.xls", SearchOption.TopDirectoryOnly);
                if (files.Length != 0)
                {
                    if (files.Length > 1)
                    {
                        LogMessage("More than one downloadRicChangeEvents*.xls file found, please have a check, there should be only one download file. ");
                    }

                    else
                        downloadFilePath = files[0];
                    //foreach (string file in files)
                    //{
                    //    string[] arr = file.Split('_');
                    //    if (DateTime.ParseExact(arr[1],"yyyyMMdd",null).Equals(DateTime.Parse(configObj.DATE)))
                    //    {
                    //        downloadFilePath = file;
                    //    }
                    //}
                }
            }
            if (downloadFilePath == string.Empty)
            {
                LogMessage("There's no \"DownloadRicChangeEvents\" file under foler " + configObj.DOWNLOAD_FILE_DIR);
            }
            else
            {
                using (ExcelApp downloadApp = new ExcelApp(false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(downloadApp, downloadFilePath);
                    var worksheet = (Worksheet)workbook.Worksheets[2];
                    if (worksheet == null)
                    {
                        LogMessage("There's no sheet2 in the excel file " + workbook.FullName);
                    }

                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                    using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 2, 1, ExcelLineWriter.Direction.Right))
                    {
                        while (reader.Row <= lastUsedRow)
                        {
                            Range changeTypeRange = ExcelUtil.GetRange(reader.Row, 4, worksheet);
                            Range ricNowRange = ExcelUtil.GetRange(reader.Row, 3, worksheet);
                            if (changeTypeRange.Text != null && changeTypeRange.Text.ToString().Trim().ToUpper() == "ADD" && ricNowRange.Text != null)
                            {
                                string ricNow = ricNowRange.Text.ToString().Trim();
                                if (addRicDic.ContainsKey(ricNow))
                                {

                                    addRicDic[ricNow].EffectiveDate = reader.ReadLineCellText();//Effective Date
                                    addRicDic[ricNow].RicWas = reader.ReadLineCellText();//Ric Was
                                    addRicDic[ricNow].RicNow = reader.ReadLineCellText();//Ric Now
                                    addRicDic[ricNow].ChangeType = reader.ReadLineCellText();//Change Type
                                    addRicDic[ricNow].Country = reader.ReadLineCellText();//Country
                                    addRicDic[ricNow].Exchanges = reader.ReadLineCellText();//Exchanges
                                    addRicDic[ricNow].AssetClass = reader.ReadLineCellText();//Asset Class
                                    addRicDic[ricNow].DescriptionWas = reader.ReadLineCellText();
                                    addRicDic[ricNow].DescriptionNow = reader.ReadLineCellText();
                                    addRicDic[ricNow].SummaryOfChange = reader.ReadLineCellText();
                                    addRicDic[ricNow].ISINWas = reader.ReadLineCellText();
                                    //addRicDic[ricNow].ISINNow = reader.GetCellText();
                                    reader.ReadLineCellText();
                                    addRicDic[ricNow].SecondMarketId = reader.ReadLineCellText();
                                    addRicDic[ricNow].SecondMarketWas = reader.ReadLineCellText();
                                    addRicDic[ricNow].SecondMarketNow = reader.ReadLineCellText();
                                    addRicDic[ricNow].RicSeqId = reader.ReadLineCellText();
                                }

                                //reader.PlaceNext(reader.Row + 1, 1);
                            }
                            reader.PlaceNext(reader.Row + 1, 1);
                        }
                    }
                }
            }
        }

    }
}
