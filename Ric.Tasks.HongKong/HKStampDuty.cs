using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity.Validation;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Model;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    #region Config

    [ConfigStoredInDB]
    public class HKStampDutyConfig
    {
        [StoreInDB]
        [Category("Directory")]
        [Description("Where to save upload files.")]
        [DisplayName("Working directory")]
        public string WorkingDir { get; set; }
    }

    #endregion

    public class HKStampDuty : GeneratorBase
    {
        private class ResultFormat
        {
            public string Code { set; get; }
            public string ShortName { set; get; }
            public string SubjectToStampDuty { set; get; }
        }
        private static HKStampDutyConfig _configObj;
        private List<ResultFormat> resultForFile = new List<ResultFormat>();
        private HongKongModel _hongKongDataContext;

        #region GeneratorBase implementation

        protected override void Initialize()
        {
            base.Initialize();
            _configObj = Config as HKStampDutyConfig;
            _hongKongDataContext = new HongKongModel();
        }

        protected override void Start()
        {
            try
            {
                LogMessage("Unzip files in directory");
                UnZipFilesInDir(_configObj.WorkingDir);

                LogMessage("Getting list of file in folder");
                List<string> fileList = Directory.GetFiles(_configObj.WorkingDir).ToList();

                LogMessage("Retrieving Rics from scanned files");
                List<string> rics = GetRics(fileList);

                LogMessage("Query GATS to seek informations");
                Dictionary<string, string> res = PrepareRicsForGats(rics);

                LogMessage("Compare GATS result with own database");
                CompareData(res);

            }
            catch (Exception ex)
            {
                LogMessage(ex.Message, Logger.LogType.Error);
                LogMessage(ex.StackTrace, Logger.LogType.Error);
                throw;
            }
            finally
            {
                _hongKongDataContext.Dispose();
            }
        }

        #endregion

        #region Unzip files

        /// <summary>
        /// Scan directory and unzip all found zip files
        /// </summary>
        /// <param name="path"></param>
        private void UnZipFilesInDir(string path)
        {
            foreach (string file in Directory.GetFiles(path).ToList().Where(file => Path.GetExtension(file).ToUpper() == ".ZIP"))
            {
                string err;
                ZipUtil.UnzipFile(file, out err);
            }
        }

        #endregion

        /// <summary>
        /// From found results create a list of ResultFormat objects
        /// </summary>
        /// <param name="res"></param>
        private void FormatResults(Dictionary<string, string> res)
        {
            foreach (ResultFormat tmp in res.Select(resultLine => new ResultFormat
            {
                Code = resultLine.Key,
                ShortName = resultLine.Value,
                SubjectToStampDuty = (resultLine.Value == "2" || resultLine.Value == "13") ? "N" : "Y"
            }))
            {
                resultForFile.Add(tmp);
            }
        }

        /// <summary>
        /// Scan directory, find all .xls and .xlsx files
        /// </summary>
        /// <param name="fileList"></param>
        /// <returns></returns>
        private List<string> GetRics(IEnumerable<string> fileList)
        {
            var rics = new List<string>();
            foreach (List<string> ricInFile in from file in fileList
                                               let extend = Path.GetExtension(file).ToUpper()
                                               where extend.Equals(".XLS") || extend.Equals(".XLSX")
                                               select GetRicFromFile(file)
                                                   into ricInFile
                                                   where ricInFile != null
                                                   select ricInFile)
            {
                rics.AddRange(ricInFile);
            }
            return rics;
        }

        /// <summary>
        /// Find TAX_ELIGBL fids in GATS from given Rics
        /// </summary>
        /// <param name="rics"></param>
        /// <returns></returns>
        private Dictionary<string, string> QueryGats(string rics)
        {
            var gats = new GatsUtil(GatsUtil.Server.Elektron);
            var gatsValues = new Dictionary<string, string>();
            var rgxSpace = new Regex(@"\s+");
            try
            {
                string[] stringSeparators = { "\r\n" };
                char[] stringSeparators2 = { ' ' };

                string test = gats.GetGatsResponse(rics, "TAX_ELIGBL");

                string[] lines = test.Split(stringSeparators, StringSplitOptions.None);
                foreach (string[] lineTab in lines.Select(line => rgxSpace.Replace(line, " "))
                                                  .Select(formattedLine => formattedLine.Split(stringSeparators2))
                                                  .Where(lineTab => lineTab.Length > 2
                                                      && lineTab[2] != "").Where(lineTab => !gatsValues.ContainsKey(lineTab[0].Trim())))
                {
                    gatsValues.Add(lineTab[0].Trim(), lineTab[2].Trim());
                }
                return gatsValues;
            }
            catch (Exception ex)
            {
                LogMessage("Error in QueryGats", Logger.LogType.Error);
                throw new Exception("Error While using Gats: " + ex.Message);
            }
        }

        /// <summary>
        /// from ric list generate a ric string
        /// query Gats 2000 rics at a time to avoid overflow
        /// </summary>
        /// <param name="ricsToSearch"></param>
        /// <returns></returns>
        private Dictionary<string, string> PrepareRicsForGats(List<string> ricsToSearch)
        {
            var gatsRes = new Dictionary<string, string>();
            for (int i = 0; i < ricsToSearch.Count; i += 2000)
            {
                var sb = new StringBuilder();
                int tmpi = 0;
                foreach (string ric in ricsToSearch)
                {
                    if (tmpi >= i && tmpi <= (i + 2000))
                    {
                        sb.AppendFormat("{0},", ric);
                    }
                    tmpi++;
                }
                foreach (var item in QueryGats(sb.ToString()).Where(item => !gatsRes.ContainsKey(item.Key.Replace(".HK", ""))))
                {
                    gatsRes.Add(item.Key.Replace(".HK", ""), item.Value);
                }
            }
            return gatsRes;
        }

        /// <summary>
        /// Find all rics in xls file
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private List<string> GetRicFromFile(string file)
        {
            var excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                throw new Exception("Excel could not be started.");
            }

            try
            {
                string fileName = Path.GetFileName(file);
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, file);
                var wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    throw new Exception("Worksheet can not be accessed for file:" + fileName);
                }

                int col = 1;

                while ((ExcelUtil.GetRange(1, col, wSheet).Text.ToString().Trim() != "RIC"))
                {
                    col++;
                }

                int lastUsedRow = wSheet.UsedRange.Rows.Count;

                var ric = new List<string>();
                for (int i = 2; i <= lastUsedRow; i++)
                {
                    ric.Add(ExcelUtil.GetRange(i, col, wSheet).Text.ToString().Trim());
                }
                return ric;
            }
            catch
            {
                throw new Exception("Error found in reading ric list from " + Path.GetFileName(file));
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        #region File generation

        private void GenerateFile()
        {
            var equity = new List<List<string>>();
            var warrant = new List<List<string>>();
            //equity.Add(new[] { "RIC", "CAPITALIZATION STAMP DUTY" }.ToList());
            equity.Add(new[] { "RIC", "CAPITALISATION STAMP DUTY" }.ToList());
            warrant.Add(new[] { "RIC", "WARRANT STAMP DUTY" }.ToList());
            foreach (var result in resultForFile)
            {
                if (result.Code.Substring(0, 1) == "6" || result.ShortName.Contains("@"))
                {
                    var oneRes = new List<string> { result.Code + ".HK" };
                    while (oneRes[0].Length < 7)
                    {
                        oneRes[0] = "0" + oneRes[0];
                    }
                    oneRes.Add(result.SubjectToStampDuty);
                    warrant.Add(oneRes);
                }//warrant 
                else
                {
                    var oneRes = new List<string> { result.Code.TrimStart("0".ToArray()) + ".HK" };
                    while (oneRes[0].Length < 7)
                    {
                        oneRes[0] = "0" + oneRes[0];
                    }
                    oneRes.Add(result.SubjectToStampDuty);
                    equity.Add(oneRes);
                }
            }
            string equityPath = string.Format("{0}\\equity_{1}.csv", _configObj.WorkingDir, DateTime.Now.ToString("ddMM_hhmm", new CultureInfo("en-US")));
            string warrantPath = string.Format("{0}\\warrant_{1}.csv", _configObj.WorkingDir, DateTime.Now.ToString("ddMM_hhmm", new CultureInfo("en-US")));
            WriteToCsv(equityPath, equity);
            AddResult("Equity", equityPath, "file");
            WriteToCsv(warrantPath, warrant);
            AddResult("Warrant", warrantPath, "file");
        }

        public static void WriteToCsv(string filepath, List<List<string>> data)
        {
            var fs = new FileStream(filepath, FileMode.Create);
            var sw = new StreamWriter(fs, Encoding.UTF8);
            foreach (List<string> list in data)
            {
                for (int j = 0; j < list.Count; j++)
                {
                    if (j != list.Count - 1)
                        sw.Write(list[j] + ",");
                    else
                        sw.Write(list[j]);
                }
                sw.Write("\r\n");
            }
            sw.Close();
            fs.Close();
        }

        #endregion

        #region compare

        /// <summary>
        /// Compare grabbed values with Database
        /// 
        /// if already in database see if values match
        /// if not in database add it and put in result file
        /// 
        /// StampDuty values : 1 = Y And 2 = N
        /// 
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        private void CompareData(Dictionary<string, string> values)
        {
            var res = new Dictionary<string, string>();
            try
            {
                foreach (var code in values)
                {
                    ETI_HK_StampDuty result = (from stampDuty in _hongKongDataContext.ETI_HK_StampDuties
                                               where stampDuty.Ric == code.Key
                                               select stampDuty).SingleOrDefault();

                    if (result == null)
                    {
                        if (code.Value.Length < 3)
                        {
                            _hongKongDataContext.ETI_HK_StampDuties.Add(
                                new ETI_HK_StampDuty
                                {
                                    Ric = code.Key,
                                    SubjectToStampDuty = (code.Value == "1" || code.Value == "12") ? "Y" : "N",
                                    LastChange = DateTime.Now
                                });
                            res.Add(code.Key, code.Value);
                        }
                    }
                    else
                    {
                        if ((result.SubjectToStampDuty == "N" && (code.Value == "1" || code.Value == "12"))
                            || (result.SubjectToStampDuty == "Y" && (code.Value == "2" || code.Value == "13")))
                        {
                            res.Add(code.Key, code.Value);
                            result.SubjectToStampDuty = (code.Value == "1" || code.Value == "12") ? "Y" : "N";
                        }
                    }
                }
            }
            catch (DbEntityValidationException dbEx)
            {
                foreach (var validationError in dbEx.EntityValidationErrors.SelectMany(validationErrors => validationErrors.ValidationErrors))
                {
                    Trace.TraceInformation("Property: {0} Error: {1}", validationError.PropertyName,
                        validationError.ErrorMessage);
                }
            }
            LogMessage("Saving changes in own database");
            _hongKongDataContext.SaveChanges();

            FormatResults(res);
            LogMessage("Generate result files");
            GenerateFile();
        }

        #endregion
    }
}