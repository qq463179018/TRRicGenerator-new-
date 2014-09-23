using System;
using System.Net;
using System.IO;
using System.Text;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing.Design;
using Microsoft.Office.Interop.Excel;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.Globalization;
using Ric.Core;
using Ric.Util;


namespace Ric.Tasks.Validation
{
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class ValidateNDAFMWithDSEFTPConfig
    {
        [StoreInDB]
        [Category("NumberOfFileType")]
        [Description("the number of the filetype on the ftp eg:2307")]
        public string TxtFileFromFtp { get; set; }

        [StoreInDB]
        [Category("FMPath")]
        [Description("the full path of the FM file")]
        public string FMFilePath { get; set; }
    }

    public class KoreaFMLine
    {
        public string UpdatedDate { get; set; }
        public string EffectiveDate { get; set; }
        public string RIC { get; set; }
        public string FM { get; set; }
        public string IDNDisplayName { get; set; }
        public string ISIN { get; set; }
        public string Ticker { get; set; }
        public string BCAST_REF { get; set; }
        public string QACommonName { get; set; }
        public string MatDate { get; set; }
        public string StrikePrice { get; set; }
        public string QuanityofWarrants { get; set; }
        public string IssuePrice { get; set; }
        public string IssueDate { get; set; }
        public string ConversionRatio { get; set; }
        public string Issuer { get; set; }
        public string Chain { get; set; }
        public string LastTradingDate { get; set; }
        public string KnockoutPrice { get; set; }
    }

    public class ValidateNDAFMWithDSEFTP : GeneratorBase
    {
        private string txtFileFromFtp = null;
        private string fMFilePath = null;
        List<KoreaFMLine> listKFM = new List<KoreaFMLine>();
        public ValidateNDAFMWithDSEFTPConfig configObj = null;
        DateTime dateValue;
        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as ValidateNDAFMWithDSEFTPConfig;
        }
        protected override void Start()
        {
            GetStreamToKoreaFMLine();
            UpdateFMToLocal();
        }

        /// <summary>
        /// Get Stream From Ftp
        /// </summary>
        public void GetStreamToKoreaFMLine()
        {
            try
            {
                txtFileFromFtp = configObj.TxtFileFromFtp;
                dateValue = DateTime.ParseExact(configObj.FMFilePath.Substring(configObj.FMFilePath.IndexOf("for") + 4, 11).Replace(" ", "/"), "dd/MMM/yyyy", new CultureInfo("en-US"), DateTimeStyles.None);
                txtFileFromFtp = @"ftp://ASIA2:ASIA2@ds1.rds.reuters.com//" + txtFileFromFtp + dateValue.Month + dateValue.Day + ".M";

                FtpWebRequest request =(FtpWebRequest)WebRequest.Create(txtFileFromFtp);
                WebProxy proxy = new WebProxy("10.40.14.56", 80);
                request.Proxy = proxy;
                WebResponse res = request.GetResponse();
                StreamReader sr = new StreamReader(res.GetResponseStream());
                string tmp = null;
                while ((tmp = sr.ReadLine()) != null)
                {
                    if (tmp.Length>805)
                    {
                        KoreaFMLine kFML = new KoreaFMLine();
                        kFML.RIC = tmp.Substring(2, 20);
                        kFML.ISIN = tmp.Substring(83, 12);
                        kFML.Ticker = tmp.Substring(411, 25);
                        kFML.QACommonName = tmp.Substring(22, 36);
                        kFML.MatDate = tmp.Substring(781, 8);
                        kFML.StrikePrice = tmp.Substring(790, 14);
                        listKFM.Add(kFML);
                    }

                }
            }
            catch (Exception)
            {
                Logger.LogErrorAndRaiseException("Cann't find file from Ftp");
            }
        }

        /// <summary>
        /// updateFmToLocal
        /// </summary>
        public void UpdateFMToLocal()
        {
            DateTime dateValue;
            fMFilePath = configObj.FMFilePath;
            if (File.Exists(fMFilePath))
            {
                ExcelApp app = new ExcelApp(false, false);
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, fMFilePath);
                Worksheet ws = workbook.Worksheets[1] as Worksheet;

                Range rangeISIN = ws.get_Range("G1");
                ExcelUtil.InsertBlankCols(rangeISIN, 1);
                ws.Cells[4, 7] = "ISIN_New";
                Range rangeTicker = ws.get_Range("I1");
                ExcelUtil.InsertBlankCols(rangeTicker, 1);
                ws.Cells[4, 9] = "Ticker_New";
                Range rangeQACommonName = ws.get_Range("L1");
                ExcelUtil.InsertBlankCols(rangeQACommonName, 1);
                ws.Cells[4, 12] = "QACommonName_New";
                Range rangeMatDate = ws.get_Range("N1");
                ExcelUtil.InsertBlankCols(rangeMatDate, 1);
                ws.Cells[4, 14] = "MatDate_New";
                Range rangeStrikePrice = ws.get_Range("P1");
                ExcelUtil.InsertBlankCols(rangeStrikePrice, 1);
                ws.Cells[4, 16] = "StrikePrice_New";

                foreach (KoreaFMLine fm in listKFM)
                {
                    for (int i = 4; ws.get_Range("C" + i).Value2 != null; i++)
                    {
                        if (fm.RIC.Trim() == ws.get_Range("C" + i).Value2.ToString())
                        {
                            ws.Cells[i, 7] = fm.ISIN.Trim();
                            ws.Cells[i, 9] = fm.Ticker.Trim();
                            ws.Cells[i, 12] = fm.QACommonName.Trim();
                            dateValue = DateTime.ParseExact(fm.MatDate.Trim(), "yyyyMMdd", new CultureInfo("en-US"), DateTimeStyles.None);
                            fm.MatDate = dateValue.ToString("dd-MMM-yy");
                            ws.Cells[i, 14] = fm.MatDate;
                            ws.Cells[i, 16] = fm.StrikePrice.Trim();
                            if (ws.get_Range("F" + i).Value2.ToString() != fm.ISIN.Trim())
                            {
                                ExcelUtil.GetRange(i, 7, ws).Interior.Color = System.Drawing.Color.FromArgb(0, 200, 0).ToArgb();
                            }
                            if (ws.get_Range("H" + i).Value2.ToString() != fm.Ticker.Trim())
                            {
                                ExcelUtil.GetRange(i, 9, ws).Interior.Color = System.Drawing.Color.FromArgb(0, 200, 0).ToArgb();
                            }
                            if (ws.get_Range("K" + i).Value2.ToString() != fm.QACommonName.Trim())
                            {
                                ExcelUtil.GetRange(i, 12, ws).Interior.Color = System.Drawing.Color.FromArgb(0, 200, 0).ToArgb();
                            }
                            if (ws.get_Range("M" + i).Value2.ToString() != fm.MatDate.Trim())
                            {
                                ExcelUtil.GetRange(i, 14, ws).Interior.Color = System.Drawing.Color.FromArgb(0, 200, 0).ToArgb();
                            }
                            if (ws.get_Range("O" + i).Value2.ToString() != fm.StrikePrice.Trim())
                            {
                                ExcelUtil.GetRange(i, 16, ws).Interior.Color = System.Drawing.Color.FromArgb(0, 200, 0).ToArgb();
                            }
                        }
                    }
                }
                app.ExcelAppInstance.AlertBeforeOverwriting = false;
                workbook.Save();
                workbook.Close();
                app.Dispose();
            }
            else
            {
                Logger.LogErrorAndRaiseException("FM file is not exist");
            }
        }
    }
}

