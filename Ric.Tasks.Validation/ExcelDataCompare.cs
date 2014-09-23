using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using Ric.Util;
using Microsoft.Office.Interop.Excel;

namespace Ric.Tasks.Validation
{
    [ConfigStoredInDB]
    class ExcelDataCompareConfig
    {
        [StoreInDB]
        [Category("InputFilePath")]
        [Description("GetFieldsListPath like:D:/xxxx.xls")]
        public string InputFilePath { get; set; }

        [StoreInDB]
        [Category("OutputFileType")]
        [Description("OutputXlsOrCsvPath like:D:\\xxxxxx.xls(csv)\r\n 1.xls will need more time.\r\n 2.csv will need less time")]
        public FileType OutputFileType { get; set; }

        [StoreInDB]
        [Category("OutputFolder")]
        [Description("OutputXlsOrCsvFolder like:D:\\")]
        public string OutputFolder { get; set; }
    }

    class ExcelDataCompare : GeneratorBase
    {
        private static ExcelDataCompareConfig configObj = null;
        private Dictionary<string, List<string>> dicListFirst = new Dictionary<string, List<string>>();//show report file with first file Data
        private Dictionary<string, List<string>> dicListSecond = new Dictionary<string, List<string>>();
        private Dictionary<string, List<string>> dicListFirstDiffSecond = new Dictionary<string, List<string>>();
        private string outputFilePath = string.Empty;

        protected override void Initialize()
        {
            configObj = Config as ExcelDataCompareConfig;
        }

        private bool PreprarConfig()
        {
            bool configError = false;

            try
            {
                if (!File.Exists(configObj.InputFilePath))
                {
                    MessageBox.Show("Input File is not exist!");
                    return true;
                }

                if ((configObj.OutputFolder + "").Trim().Length == 0)
                {
                    MessageBox.Show("please input a output foder path");
                    return true;
                }

                if (Directory.Exists(configObj.OutputFolder))
                    Directory.CreateDirectory(configObj.OutputFolder);

                outputFilePath = Path.Combine(configObj.OutputFolder, string.Format("ReportDiff.{0}", configObj.OutputFileType.ToString()));
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                configError = true;
            }

            return configError;
        }

        protected override void Start()
        {
            if (PreprarConfig())
            {
                LogMessage("Error happened in configuration!");
                MessageBox.Show("Please check the configuration setting!");
                return;
            }

            LogMessage("start to read first sheet...");
            List<List<string>> listListDataFirst = GetDataFromFile(configObj.InputFilePath, 1);
            LogMessage(string.Format("first sheet count:{0}", listListDataFirst == null ? "0" : listListDataFirst.Count.ToString()));

            LogMessage("start to read second sheet...");
            List<List<string>> listListDataSecond = GetDataFromFile(configObj.InputFilePath, 2);
            LogMessage(string.Format("second sheet count:{0}", listListDataSecond == null ? "0" : listListDataSecond.Count.ToString()));

            if (listListDataSecond == null)
            {
                MessageBox.Show("the sheet count of input file >=2 !");
                return;
            }

            LogMessage("start to create index for first sheet...");
            dicListFirst = CreateIndexForListList(listListDataFirst);
            LogMessage(string.Format("first file index count:{0}", dicListFirst == null ? "0" : dicListFirst.Count.ToString()));

            LogMessage("start to create index for second sheet...");
            dicListSecond = CreateIndexForListList(listListDataSecond);
            LogMessage(string.Format("second file index count:{0}", dicListSecond == null ? "0" : dicListSecond.Count.ToString()));

            LogMessage("start to remove  exist data in second sheet...");
            dicListFirstDiffSecond = FormatDiffFirstInSecond(dicListFirst, dicListSecond);
            LogMessage(string.Format("output file index count:{0}", dicListFirstDiffSecond == null ? "0" : dicListFirstDiffSecond.Count.ToString()));

            LogMessage("start to generate output file...");
            GenerateOutputFile(dicListFirstDiffSecond, outputFilePath);
        }

        private void GenerateOutputFile(Dictionary<string, List<string>> dicList, string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);

                if (path.EndsWith(".csv"))
                    XlsOrCsvUtil.GenerateStringCsv(path, dicList);
                else
                    XlsOrCsvUtil.GenerateXls0rCsv(path, dicList);

                AddResult("ReportDiff", path, configObj.OutputFileType.ToString());
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private Dictionary<string, List<string>> FormatDiffFirstInSecond(Dictionary<string, List<string>> dicListFirst, Dictionary<string, List<string>> dicListSecond)
        {
            Dictionary<string, List<string>> result = new Dictionary<string, List<string>>();

            try
            {
                if (dicListFirst == null || dicListFirst.Count <= 1)
                    return result;

                if (dicListSecond == null || dicListSecond.Count <= 1)
                    return dicListFirst;

                string[] keysSecondDic = new string[dicListSecond.Keys.Count];
                dicListSecond.Keys.CopyTo(keysSecondDic, 0);

                string[] keysFirstDic = new string[dicListFirst.Keys.Count];
                dicListFirst.Keys.CopyTo(keysFirstDic, 0);

                //Dictionary<int, int> dicFieldsMapingFirstToSecond = GetSameFieldsMaping(dicListFirst[keysFirstDic[0]].ToArray(), dicListSecond[keysSecondDic[0]].ToArray());
                Dictionary<int, int> dicFieldsMapingFirstToSecond = GetSameFieldsMaping(dicListSecond[keysSecondDic[0]].ToArray(), dicListFirst[keysFirstDic[0]].ToArray());

                if (!dicListFirst.ContainsKey(keysSecondDic[0]))
                    return dicListFirst;

                //result.Add("title", dicListFirst[keysSecondDic[0]]);

                List<string> lineFirst = null;
                List<string> lineSecond = null;

                for (int i = 1; i < keysSecondDic.Length; i++)
                {
                    if (!dicListFirst.ContainsKey(keysSecondDic[i]))
                        continue;

                    lineFirst = dicListFirst[keysSecondDic[i]];
                    lineSecond = dicListSecond[keysSecondDic[i]];
                    dicListFirst[keysSecondDic[i]] = RemoveSameValue(lineFirst, lineSecond, dicFieldsMapingFirstToSecond);
                }

                return dicListFirst;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private List<string> RemoveSameValue(List<string> lineFirst, List<string> lineSecond, Dictionary<int, int> dicFieldsMapingFirstToSecond)
        {
            List<string> newLine = new List<string>();

            try
            {
                string valueSecond = string.Empty;

                newLine.Add(lineFirst[0]);

                for (int i = 1; i < lineFirst.Count; i++)
                {
                    if (!dicFieldsMapingFirstToSecond.ContainsKey(i))
                    {
                        newLine.Add(lineFirst[i]);
                        continue;
                    }

                    valueSecond = (lineSecond[dicFieldsMapingFirstToSecond[i]] + "").Trim();

                    if ((lineFirst[i] + "").Trim().Equals(valueSecond))
                        newLine.Add(string.Empty);
                    else
                        newLine.Add(lineFirst[i]);
                }

                return newLine;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private Dictionary<int, int> GetSameFieldsMaping(string[] listFieldsSecond, string[] listFieldsFirst)
        {
            Dictionary<int, int> dicMaping = new Dictionary<int, int>();

            try
            {
                string indexFields = string.Empty;

                for (int i = 1; i < listFieldsFirst.Length; i++)
                {
                    if ((listFieldsFirst[i] + "").Trim().Length == 0)
                        continue;

                    if (!listFieldsSecond.Contains(listFieldsFirst[i]))
                        continue;

                    indexFields = GetIndex(listFieldsFirst[i].Trim(), listFieldsSecond);

                    if ((indexFields + "").Trim().Length != 0)
                        dicMaping.Add(i, Convert.ToInt32(indexFields));
                }

                return dicMaping;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string GetIndex(string field, string[] listFieldsSecond)
        {
            string result = string.Empty;

            try
            {
                for (int i = 0; i < listFieldsSecond.Length; i++)
                {
                    if ((listFieldsSecond[i] + "").Trim().Equals(field))
                        return i.ToString();
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

            return result;
        }

        private Dictionary<string, List<string>> CreateIndexForListList(List<List<string>> listList)
        {
            Dictionary<string, List<string>> dicList = new Dictionary<string, List<string>>();

            try
            {
                if (listList == null || listList.Count <= 1)
                {
                    Logger.Log(string.Format("no data in workbook sheet! "), Logger.LogType.Error);
                    return null;
                }

                string key = string.Empty;

                foreach (var list in listList)
                {
                    if (list == null || (list[0] + "").Trim().Length == 0)
                        continue;

                    key = list[0].Trim();

                    if (dicList.ContainsKey(key))
                        key = string.Format("{0}(Repetition{1})", key, dicList.Count.ToString());

                    dicList.Add(key, list);
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

            return dicList;
        }

        private List<List<string>> GetDataFromFile(string path, int position)
        {
            try
            {
                if (!File.Exists(path))
                {
                    Logger.Log(string.Format("file :{0} is not esist.", path), Logger.LogType.Error);
                    return null;
                }

                ExcelApp app = new ExcelApp(false, false);
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.InputFilePath);

                if (position > workbook.Sheets.Count)
                {
                    Logger.Log(string.Format("Workbook sheet count:{0}, error file.", workbook.Sheets.Count.ToString()), Logger.LogType.Error);
                    return null;
                }

                return WorkbookExtension.ToList(workbook, position);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }
    }
}
