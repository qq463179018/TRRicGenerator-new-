using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using pdftron.PDF;
using Ric.Util;
using pdftron;
using System.IO;
using System.Net.Mail;
using System.ComponentModel;

namespace Ric.Tasks.Japan
{
    [ConfigStoredInDB]
    public class JPDataBackFillConfig
    {
        [StoreInDB]
        [Category("FolderPath")]
        [Description("bulk file file path")]
        public string OutputFolder { get; set; }

        [StoreInDB]
        [Category("FilePath")]
        [Description("pdf file path")]
        public string FilePath1 { get; set; }

        [StoreInDB]
        [Category("FilePath")]
        [Description("pdf file path")]
        public string FilePath2 { get; set; }

        [StoreInDB]
        [Category("FilePath")]
        [Description("pdf file path")]
        public string FilePath3 { get; set; }
    }

    public class JPDataBackFill : GeneratorBase
    {
        private static JPDataBackFillConfig config = null;
        public PdfAnalyzer pa = null;

        protected override void Initialize()
        {
            config = Config as JPDataBackFillConfig;
            pa = new PdfAnalyzer();
        }

        protected override void Start()
        {
            #region [create output folder]
            if ((config.OutputFolder + "").Trim().Length == 0)
            {
                LogMessage("Please input the output folder in the config page.");
                return;
            }

            if (!Directory.Exists(config.OutputFolder))
            {
                try
                {
                    Directory.CreateDirectory(config.OutputFolder);
                }
                catch (Exception ex)
                {
                    LogMessage(string.Format("create the outputfolder error.msg:{0}", ex.Message));
                    return;
                }
            }
            #endregion

            #region [extract type1]
            if (File.Exists(config.FilePath1))
            {
                try
                {
                    //StartExtractFile1();
                }
                catch (Exception ex)
                {
                    LogMessage(string.Format("extract file {0} error,msg:{1}", config.FilePath1, ex.Message));
                }
            }
            #endregion

            #region [extract type2]
            if (File.Exists(config.FilePath2))
            {
                try
                {
                    //StartExtract2();
                }
                catch (Exception ex)
                {
                    LogMessage(string.Format("extract file {0} error,msg:{1}", config.FilePath2, ex.Message));
                }
            }
            #endregion

            #region [extract type3]
            if (File.Exists(config.FilePath3))
            {
                try
                {
                    StartExtract3();
                }
                catch (Exception ex)
                {
                    LogMessage(string.Format("extract file {0} error,msg:{1}", config.FilePath3, ex.Message));
                }
            }
            #endregion
        }

        private void StartExtract3()
        {
            List<List<string>> bulkFileFilter = null;
            List<LineFound> bulkFile = null;
            pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
            PDFDoc doc = new PDFDoc(config.FilePath3);
            doc.InitSecurityHandler();
            string patternTitle = @"北弘";
            int page = 1;
            PdfString ricPosition = GetRicPosition(doc, patternTitle, page);
            if (ricPosition == null)
                return;

            ricPosition.Position.x1 = 106.8;
            ricPosition.Position.x2 = 99.12;
            ricPosition.Position.y1 = 29.5105;
            ricPosition.Position.y2 = 44.8734;
            string patternRic = @"\d{4}";
            //string patternValue = @"\-?(\,\.\d)*";
            string patternValue = @"(\-|\+)?\d+(\,|\.|\d)+";
            bulkFile = GetValue(doc, ricPosition, patternRic, patternValue);
            int indexOK = 0;
            bulkFileFilter = FilterBulkFile(bulkFile, indexOK);
            string filePath = Path.Combine(config.OutputFolder, string.Format("{0}.csv", Path.GetFileNameWithoutExtension(config.FilePath3)));

            if (File.Exists(filePath))
                File.Delete(filePath);

            XlsOrCsvUtil.GenerateStringCsv(filePath, bulkFileFilter);
            AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "type3");
        }

        private void StartExtract2()
        {
            List<List<string>> bulkFileFilter = null;
            List<LineFound> bulkFile = null;
            pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
            PDFDoc doc = new PDFDoc(config.FilePath2);
            doc.InitSecurityHandler();
            string patternTitle = @"コード";
            int page = 5;
            PdfString ricPosition = GetRicPosition(doc, patternTitle, page);
            if (ricPosition == null)
                return;

            ricPosition.Position.x1 = 106.8;
            ricPosition.Position.x2 = 99.12;
            ricPosition.Position.y1 = 29.5105;
            ricPosition.Position.y2 = 44.8734;
            string patternRic = @"\d{4}";
            //string patternValue = @"\-?(\,\.\d)*";
            string patternValue = @"(\-|\+)?\d+(\,|\.|\d)+";
            bulkFile = GetValue(doc, ricPosition, patternRic, patternValue);
            int indexOK = 0;
            bulkFileFilter = FilterBulkFile(bulkFile, indexOK);
            string filePath = Path.Combine(config.OutputFolder, string.Format("{0}.csv", Path.GetFileNameWithoutExtension(config.FilePath2)));

            if (File.Exists(filePath))
                File.Delete(filePath);

            XlsOrCsvUtil.GenerateStringCsv(filePath, bulkFileFilter);
            AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "type2");
        }

        private void StartExtractFile1()
        {
            List<List<string>> bulkFileFilter = null;
            List<LineFound> bulkFile = null;
            pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
            PDFDoc doc = new PDFDoc(config.FilePath1);
            doc.InitSecurityHandler();
            string patternTitle = @"コード";
            int page = 3;
            PdfString ricPosition = GetRicPosition(doc, patternTitle, page);
            if (ricPosition == null)
                return;

            ricPosition.Position.x1 = 345.78;
            ricPosition.Position.x2 = 337.62;
            ricPosition.Position.y1 = 49.67;
            ricPosition.Position.y2 = 65.98;
            string patternRic = @"\d{4}";
            string patternValue = @"(\-|\+)?\d+(\,|\.|\d)+";
            bulkFile = GetValue(doc, ricPosition, patternRic, patternValue);
            int indexOK = 0;
            bulkFileFilter = FilterBulkFile(bulkFile, indexOK);
            string filePath = Path.Combine(config.OutputFolder, string.Format("{0}.csv", Path.GetFileNameWithoutExtension(config.FilePath1)));

            if (File.Exists(filePath))
                File.Delete(filePath);

            XlsOrCsvUtil.GenerateStringCsv(filePath, bulkFileFilter);
            AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "type1");
        }

        private List<List<string>> FilterBulkFile(List<LineFound> bulkFile, int indexOK)
        {
            List<List<string>> result = new List<List<string>>();

            if (bulkFile == null || bulkFile.Count == 0)
            {
                Logger.Log("no value data extract from pdf");
                return null;
            }
            int count = bulkFile[indexOK].LineData.Count;

            List<string> line = null;
            foreach (var item in bulkFile)
            {
                if (item.LineData == null || item.LineData.Count <= 0)
                    continue;

                line = new List<string>();
                if (item.LineData.Count.CompareTo(count) == 0)
                {
                    foreach (var value in item.LineData)
                    {
                        line.Add(value.Words.ToString());
                    }
                }
                else
                {
                    line.Add(item.LineData[0].Words.ToString());
                    for (int i = 1; i < count; i++)
                    {
                        line.Add(string.Empty);
                    }
                }
                result.Add(line);
            }

            return result;
        }

        private List<LineFound> GetValue(PDFDoc doc, PdfString ricPosition, string patternRic, string patternValue)
        {
            List<LineFound> bulkFile = new List<LineFound>();
            try
            {
                List<string> line = new List<string>();
                List<PdfString> ric = null;

                //for (int i = 1; i < 10; i++)
                for (int i = 1; i < doc.GetPageCount(); i++)
                {
                    ric = pa.RegexExtractByPositionWithPage(doc, patternRic, i, ricPosition.Position);
                    foreach (var item in ric)
                    {
                        LineFound lineFound = new LineFound();
                        lineFound.Ric = item.Words.ToString();
                        lineFound.Position = item.Position;
                        lineFound.PageNumber = i;
                        lineFound.LineData = pa.RegexExtractByPositionWithPage(doc, patternValue, i, item.Position, PositionRect.X2);
                        bulkFile.Add(lineFound);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return bulkFile;
        }

        private PdfString GetRicPosition(PDFDoc doc, string pattern, int page)
        {
            try
            {
                List<PdfString> ricPosition = null;
                ricPosition = pa.RegexSearchByPage(doc, pattern, page);
                if (ricPosition == null || ricPosition.Count == 0)
                {
                    Logger.Log(string.Format("there is no ric title found by using pattern:{0} to find the ric title ,in the page:{1} of the pdf:{2}"));
                    return null;
                }

                return ricPosition[0];
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                throw;
            }
        }
    }

    struct LineFound
    {
        public string Ric { get; set; }
        public Rect Position { get; set; }
        public int PageNumber { get; set; }
        public List<PdfString> LineData { get; set; }
    }
}
