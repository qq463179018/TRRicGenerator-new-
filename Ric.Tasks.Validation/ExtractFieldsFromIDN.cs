using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using Ric.Util;
using System.Text.RegularExpressions;
using System.Threading;

namespace Ric.Tasks.Validation
{
    [ConfigStoredInDB]
    class ExtractFieldsFromIDNConfig
    {
        [StoreInDB]
        [Category("InputFilePath")]
        [Description("GetFieldsListPath like:D:\\xxxx.txt")]
        public string FieldsFilePath { get; set; }

        [StoreInDB]
        [Category("InputFilePath")]
        [Description("GetRicListPath like:D:\\xxxxxx.txt")]
        public string RicFilePath { get; set; }

        [StoreInDB]
        [Category("OutputFileType")]
        [Description("OutputXlsOrCsvPath like:D:\\xxxxxx.xls(csv)\r\n 1.xls will need more time.\r\n 2.csv will need less time")]
        public FileType OutputFileType { get; set; }

        [StoreInDB]
        [Category("OutputFolder")]
        [Description("OutputXlsOrCsvFolder like:D:\\")]
        public string OutputFolder { get; set; }
    }

    enum FileType : int { csv, xls, xlsx }

    class ExtractFieldsFromIDN : GeneratorBase
    {
        private static ExtractFieldsFromIDNConfig configObj = null;
        private List<string> listRic = new List<string>();
        private List<string> listFields = new List<string>();
        private Dictionary<string, List<string>> dicListOutput = new Dictionary<string, List<string>>();
        private Dictionary<string, List<string>> dicListOutputError = new Dictionary<string, List<string>>();
        private string allFieldsPath = string.Empty;
        private string reportFieldsPath = string.Empty;

        protected override void Initialize()
        {
            configObj = Config as ExtractFieldsFromIDNConfig;
        }

        protected override void Start()
        {
            if (PreprarConfig())
            {
                LogMessage("Error happened in configuration!");
                MessageBox.Show("Please check the input file path and output foder path!");
                return;
            }

            LogMessage(string.Format("start to read ric list..."));
            listRic = ReadFileToList(configObj.RicFilePath);
            LogMessage(string.Format("ric count:{0}", listRic == null ? "0" : listRic.Count.ToString()));

            LogMessage(string.Format("start to read field list..."));
            listFields = ReadFileToList(configObj.FieldsFilePath);
            LogMessage(string.Format("field count:{0}", listFields == null ? "0" : listFields.Count.ToString()));

            LogMessage(string.Format("start to query field in IDN..."));
            dicListOutput = QueryIDN(listRic);
            LogMessage(string.Format("all field count:{0}", dicListOutput == null ? "0" : dicListOutput.Count.ToString()));
            LogMessage(string.Format("report field count:{0}", dicListOutputError == null ? "0" : dicListOutputError.Count.ToString()));

            LogMessage(string.Format("start to generate all fields file..."));
            GenerateOutputFile(dicListOutput, allFieldsPath);
            LogMessage("Generate AllFields File Succeed!");
            AddResult("ALL Fields", allFieldsPath, "IDN(XlsOrCsv)");

            LogMessage(string.Format("start to generate report fields file..."));
            GenerateOutputFile(dicListOutputError, reportFieldsPath);
            LogMessage("Generate ReportFields File Succeed!");
            AddResult("Report Fields(Witchout *.*)", reportFieldsPath, "IDN(XlsOrCsv)");
        }

        private bool PreprarConfig()
        {
            bool configError = false;

            try
            {
                if (!Directory.Exists(configObj.OutputFolder))
                    Directory.CreateDirectory(configObj.OutputFolder);

                if (!File.Exists(configObj.FieldsFilePath) || !File.Exists(configObj.RicFilePath))
                {
                    MessageBox.Show("FieldsFilePath or RicFilePath is not exist!");
                    configError = true;
                }

                allFieldsPath = Path.Combine(configObj.OutputFolder, string.Format("AllFields.{0}", configObj.OutputFileType.ToString()));
                reportFieldsPath = Path.Combine(configObj.OutputFolder, string.Format("ReportFields.{0}", configObj.OutputFileType.ToString()));
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

        private void GenerateOutputFile(Dictionary<string, List<string>> dicList, string path)
        {
            try
            {
                try
                {
                    if (File.Exists(path))
                        File.Delete(path);
                }
                catch (Exception ex)
                {
                    Logger.Log(string.Format("delete alerday exist file: {0} error.msg:{1}", path, ex.Message), Logger.LogType.Error);
                }

                if (path.ToLower().EndsWith(".xls") || path.ToLower().EndsWith(".xlsx"))
                    XlsOrCsvUtil.GenerateXls0rCsv(path, dicList);
                else if (path.ToLower().EndsWith(".csv"))
                    XlsOrCsvUtil.GenerateStringCsv(path, dicList);
                else
                    MessageBox.Show("Output path must be end with .xls(x) or .csv");
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

        private void AddTitle(List<string> listFields, Dictionary<string, List<string>> dicListOutput)
        {
            try
            {
                List<string> listTitle = new List<string>();
                listTitle.Add("RIC");

                foreach (var item in listFields)
                {
                    if ((item + "").Trim().Length == 0)
                        continue;

                    listTitle.Add(item.Trim());
                }

                dicListOutput.Add("Title", listTitle);
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

        private Dictionary<string, List<string>> QueryIDN(List<string> listRic)
        {
            Dictionary<string, List<string>> dicListOutput = new Dictionary<string, List<string>>();
            AddTitle(listFields, dicListOutput);
            AddTitle(listFields, dicListOutputError);

            try
            {
                if (listRic == null || listRic.Count == 0)
                    return null;

                string strQuery = string.Empty;
                int count = listRic.Count;
                int fenMu = 2000;
                int qiuYu = count % fenMu;
                int qiuShang = count / fenMu;

                if (qiuShang > 0)
                {
                    for (int i = 0; i < qiuShang; i++)
                    {
                        for (int j = 0; j < fenMu; j++)
                        {
                            string strTmp = listRic[i * fenMu + j].ToString().Trim();

                            if (!string.IsNullOrEmpty(strTmp))
                            {
                                strQuery += string.Format(",{0}", strTmp);
                            }
                        }

                        if (strQuery.StartsWith(","))
                            strQuery = strQuery.Remove(0, 1);

                        GetDataFromGATS(strQuery, dicListOutput);
                        strQuery = string.Empty;
                    }
                }
                for (int i = qiuShang * fenMu; i < count; i++)
                {
                    string strTmp = listRic[i].ToString().Trim();

                    if (!string.IsNullOrEmpty(strTmp))
                    {
                        strQuery += string.Format(",{0}", strTmp);
                    }
                }

                if (strQuery.StartsWith(","))
                    strQuery = strQuery.Remove(0, 1);

                GetDataFromGATS(strQuery, dicListOutput);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return dicListOutput;
        }

        private void GetDataFromGATS(string strQuery, Dictionary<string, List<string>> dicListOutput)
        {
            try
            {
                int countKey = dicListOutput.Count + 1;
                string fids = GetFids(listFields);
                int fidsCount = GetFidsCount(fids);
                string pattern = @"\n(?<Ric>\S+)\s+(?<Key>\S+)\s+(?<Value>\S+)\r";
                //GatsUtil gats = new GatsUtil();
                SkitUtil skit = new SkitUtil();
                Thread.Sleep(2000);
                //string response = gats.GetGatsResponse(strQuery, fids);
                string response = skit.GetGatsResponse(strQuery, fids);
                Thread.Sleep(2000);

                if (string.IsNullOrEmpty(response))
                {
                    LogMessage("gats return no response . please check fields or ric.");
                    return;
                }

                Regex regex = new Regex(pattern);
                MatchCollection matches = regex.Matches(response);

                string ric = string.Empty;
                string key = string.Empty;
                string value = string.Empty;
                string lastRic = string.Empty;
                int countLine = 0;
                List<string> listLine = null;
                bool isError = false;

                foreach (Match match in matches)
                {
                    ric = match.Groups["Ric"].Value;
                    key = match.Groups["Key"].Value;
                    value = match.Groups["Value"].Value;

                    if (IsErrorValue(value))
                        isError = true;

                    if (countLine % (fidsCount - 1) == 0)
                    {
                        if (listLine != null && fidsCount == listLine.Count)
                        {
                            dicListOutput.Add(countKey.ToString(), listLine);

                            if (isError)
                            {
                                dicListOutputError.Add(countKey.ToString(), ClearRightValue(listLine));
                                isError = false;
                            }

                            countKey++;
                        }

                        listLine = new List<string>();
                        listLine.Add(ric);
                    }

                    listLine.Add(value);
                    countLine++;
                }

                dicListOutput.Add(countKey.ToString(), listLine);

                if (isError)
                {
                    dicListOutputError.Add(countKey.ToString(), ClearRightValue(listLine));
                    isError = false;
                }

                countKey++;
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

        private bool IsErrorValue(string value)
        {
            bool result = false;

            try
            {
                if ((value + "").Trim().Length < 2)
                    return true;

                if (!value.Contains("."))
                    return true;

                int lastPointIndex = value.LastIndexOf(".");

                if (value.Length != lastPointIndex + 2)
                    return true;
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

        private List<string> ClearRightValue(List<string> listLine)
        {
            List<string> list = new List<string>();

            try
            {
                for (int i = 0; i < listLine.Count; i++)
                {
                    string tmp = listLine[i];

                    if (i == 0 || IsErrorValue(tmp))
                        list.Add(tmp);
                    else
                        list.Add(string.Empty);
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

            return list;
        }

        private int GetFidsCount(string fids)
        {
            int fidsCount = 0;

            try
            {
                if ((fids + "").Trim().Length == 0)
                    return fidsCount;

                if (!fids.Contains(','))
                    fidsCount = 1;

                fidsCount = fids.Split(',').Length + 1;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return fidsCount;
        }

        private string GetFids(List<string> listFields)
        {
            string fids = string.Empty;

            try
            {
                if (listFields == null || listFields.Count == 0)
                    return fids;

                StringBuilder sb = new StringBuilder();

                foreach (var item in listFields)
                {
                    if ((item + "").Trim().Length == 0)
                        continue;

                    sb.AppendFormat("{0},", item.Trim());
                }

                sb.Length = sb.Length - 1;
                fids = sb.ToString();

            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return fids;
        }

        private List<string> ReadFileToList(string path)
        {
            List<string> result = new List<string>();

            try
            {
                if (!File.Exists(path))
                    return null;

                using (FileStream fs = new FileStream(path, FileMode.Open))
                {
                    using (StreamReader sr = new StreamReader(fs))
                    {
                        result = new List<string>(sr.ReadToEnd().Replace("\r\n", ",").Split(','));
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

            return result;
        }
    }
}
