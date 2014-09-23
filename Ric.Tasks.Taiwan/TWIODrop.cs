using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Taiwan
{
    #region Configuration

    [ConfigStoredInDB]
    public class TWIODropConfig
    {
        [StoreInDB]
        [Category("Folder")]
        [DisplayName("Working folder")]
        [Description("The path of the working folder, should end with a slash\n\nEg: C:/Mydrive/")]
        public string WorkingFolder { get; set; }

        [StoreInDB]
        [Category("Files")]
        [DisplayName("Future Chain")]
        [Description("The name of the file containing future chain RIC \n\nEg: chainRic.txt")]
        public string FutureChain { get; set; }

        [StoreInDB]
        [Category("Files")]
        [DisplayName("Future Spread Chain")]
        [Description("The name of the file containing future spread chain RIC \n\nEg: spreadChainRic.txt")]
        public string FutureSpreadChain { get; set; }

        [StoreInDB]
        [Category("Files")]
        [DisplayName("Option Chain")]
        [Description("The name of the file containing option chain RIC \n\nEg: optionChainRic.txt")]
        public string OptionChain { get; set; }
    }

    #endregion

    #region Task

    class TWIODrop : GeneratorBase
    {
        #region Declaration

        private static TWIODropConfig configObj;
        private List<string> spreadRics;
        private List<string> optionList;

        #endregion

        #region Interface implementation

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as TWIODropConfig;
        }

        protected override void Start()
        {
            List<string> bulkDropFileRic = FindFutureDropList();
            LogMessage("Find future drop list", Logger.LogType.Info);

            spreadRics = FindFutureSpreadChainList();
            LogMessage("Find future spread chain list", Logger.LogType.Info);
            var futureList = CompareRics(bulkDropFileRic, spreadRics);
            if (futureList != null)
            {
                bulkDropFileRic.AddRange(futureList);
            }

            LogMessage("Compare rics", Logger.LogType.Info);
            var findOptionList = FindOptionList();
            if (findOptionList != null)
            {
                bulkDropFileRic.AddRange(findOptionList);
            }

            LogMessage("Find option list", Logger.LogType.Info);
            WriteResultsInCsv(bulkDropFileRic);

            LogMessage("Write results", Logger.LogType.Info);
            WriteResultsInTxt(bulkDropFileRic);
        }

        #endregion

        #region Gats utility functions

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private IEnumerable<string> FindRicsToSearch(string filename, int pages = 0)
        {
            List<string> ricsToSearch = new List<string>();

            StreamReader file = new StreamReader(configObj.WorkingFolder + filename);
            string line;
            while ((line = file.ReadLine()) != null)
            {
                if (line.StartsWith("0#"))
                {
                    for (int pagenb = 0; pagenb <= pages; pagenb++)
                    {
                        ricsToSearch.Add(String.Format("{0}{1}", pagenb, line.Substring(1).Trim()));
                    }
                }
            }
            return ricsToSearch;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ricsToSearch"></param>
        /// <returns></returns>
        private string PrepareRicsForGats(IEnumerable<string> ricsToSearch)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string ric in ricsToSearch)
            {
                sb.AppendFormat("{0},", ric);
            }
            return sb.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rics"></param>
        /// <returns></returns>
        private List<string> QueryGats(string rics)
        {
            List<string> result = new List<string>();
            List<string> ricGroup = null;
            List<string> gatsValue = null;
            try
            {
                ricGroup = DivideGroup(rics);
                foreach (var item in ricGroup)
                {
                    if ((item + "").Trim().Length == 0)
                        continue;

                    gatsValue = QueryGatsByGroup(item);

                    if (gatsValue == null || gatsValue.Count == 0)
                        continue;

                    result.AddRange(gatsValue);
                }

                return result;
            }
            catch (Exception ex)
            {
                LogMessage("Error in QueryGats", Logger.LogType.Error);
                throw new Exception("Error While using Gats: " + ex.Message);
            }
        }

        private List<string> DivideGroup(string rics)
        {
            List<string> result = new List<string>();
            List<string> ricsList = null;

            try
            {
                ricsList = rics.Split(',').ToList();
                if (ricsList == null || ricsList.Count == 0)
                    return null;

                string strQuery = string.Empty;
                int count = ricsList.Count;
                int fenMu = 2000;
                int qiuYu = count % fenMu;
                int qiuShang = count / fenMu;

                if (qiuShang > 0)
                {
                    for (int i = 0; i < qiuShang; i++)
                    {
                        for (int j = 0; j < fenMu; j++)
                        {
                            string strTmp = ricsList[i * fenMu + j].ToString().Trim();

                            if (!string.IsNullOrEmpty(strTmp))
                            {
                                strQuery += string.Format(",{0}", strTmp);
                            }
                        }

                        strQuery = strQuery.Remove(0, 1);
                        result.Add(strQuery);
                        strQuery = string.Empty;
                    }
                }

                for (int i = qiuShang * fenMu; i < count; i++)
                {
                    string strTmp = ricsList[i].ToString().Trim();

                    if (!string.IsNullOrEmpty(strTmp))
                    {
                        strQuery += string.Format(",{0}", strTmp);
                    }
                }

                strQuery = strQuery.Remove(0, 1);
                result.Add(strQuery);

                return result;
            }
            catch (Exception ex)
            {
                LogMessage("Error in DivideGroup", Logger.LogType.Error);
                throw new Exception("Error While divide ric to group: " + ex.Message);
            }
        }

        private List<string> QueryGatsByGroup(string rics)
        {
            GatsUtil gats = new GatsUtil();
            List<string> gatsValues = new List<string>();
            Regex rgxSpace = new Regex(@"\s+");
            Regex rgxLongLink = new Regex(@"^LONGLINK", RegexOptions.IgnoreCase);
            try
            {
                string[] stringSeparators = { "\r\n" };
                char[] stringSeparators2 = { ' ' };

                string test = gats.GetGatsResponse(rics, "");

                string[] lines = test.Split(stringSeparators, StringSplitOptions.None);
                gatsValues.AddRange(from line in lines
                                    select rgxSpace.Replace(line, " ")
                                        into formattedLine
                                        select formattedLine.Split(stringSeparators2)
                                            into lineTab
                                            where lineTab.Length > 2
                                                && rgxLongLink.IsMatch(lineTab[1])
                                                && lineTab[2] != ""
                                                && !lineTab[2].EndsWith(".TW")
                                                && !lineTab[2].EndsWith("sp.TM")
                                            select lineTab[2]);
                return gatsValues;
            }
            catch (Exception ex)
            {
                LogMessage("Error in QueryGatsByGroup", Logger.LogType.Error);
                throw new Exception("Error While using Gats: " + ex.Message);
            }
        }

        #endregion

        #region Find Future drop list

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private List<string> FindFutureDropList()
        {
            try
            {
                IEnumerable<string> ricsToSearch = FindRicsToSearch(configObj.FutureChain, 5);
                List<string> ricsResult = QueryGats(PrepareRicsForGats(ricsToSearch));
                return FindOpint(PrepareRicsForGats(ricsResult));
            }
            catch (Exception)
            {
                LogMessage("Error in FindFutureDropList", Logger.LogType.Error);
            }
            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rics"></param>
        /// <returns></returns>
        private List<string> FindOpint(string rics)
        {
            GatsUtil gats = new GatsUtil();
            List<string> gatsValues = new List<string>();
            Regex rgxSpace = new Regex(@"\s+");
            try
            {
                string[] stringSeparators = { "\r\n" };
                char[] stringSeparators2 = { ' ' };

                string test = gats.GetGatsResponse(rics, "OPINT_1");

                string[] lines = test.Split(stringSeparators, StringSplitOptions.None);
                gatsValues.AddRange(from line in lines
                                    select rgxSpace.Replace(line, " ")
                                        into formattedLine
                                        select formattedLine.Split(stringSeparators2)
                                            into lineTab
                                            where lineTab.Length > 2
                                                && lineTab[2] == "0"
                                            select lineTab[0]);
                return gatsValues;
            }
            catch (Exception ex)
            {
                throw new Exception("Error While using Gats: " + ex.Message);
            }
        }

        #endregion

        #region Find Spread Rics

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private List<string> FindFutureSpreadChainList()
        {
            List<string> ricsResult = null;
            try
            {
                IEnumerable<string> ricsToSearch = FindRicsToSearch(configObj.FutureSpreadChain, 5);
                ricsResult = QueryGats(PrepareRicsForGats(ricsToSearch));
            }
            catch (Exception)
            {
                LogMessage("Error in FindFutureSpreadChainList", Logger.LogType.Error);
            }
            return ricsResult;
        }

        #endregion

        #region Compare droplist and spread rics

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dropList"></param>
        /// <param name="spreadrics"></param>
        /// <returns></returns>
        private IEnumerable<string> CompareRics(IEnumerable<string> dropList, List<string> spreadrics)
        {
            //List<string> result = new List<string>();

            try
            {
                return (from ric in dropList
                        from ric2 in spreadrics
                        where ric == (ric2.Substring(0, 5) + ric2.Substring(ric2.Length - 2, 2))
                            || ric == ric2.Substring(0, 7)
                        select ric2).ToList();

                //    string ric2 = string.Empty;
                //    foreach (var item in spreadrics)
                //    {
                //        if ((item + "").Trim().Length == 0)
                //            continue;

                //        if (dropList.Contains(item)
                //            || (item.Length >= 8 && dropList.Contains(item.Substring(0, 7))))
                //        {
                //            result.Add(item);
                //            continue;
                //        }

                //        int start = item.Length - 2;
                //        ric2 = string.Format("{0}{1}", item.Substring(0, 5), item.Substring(start, 2));
                //        if (dropList.Contains(ric2))
                //            result.Add(item);
                //    }

                //    return result;
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

        #endregion

        #region Find Options

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private IEnumerable<string> FindOptionList()
        {
            try
            {
                IEnumerable<string> ricsToSearch = FindRicsToSearch(configObj.OptionChain, 6);
                List<string> ricsResult = QueryGats(PrepareRicsForGats(ricsToSearch));

                return ricsResult == null ? null : FindMatchingRics(QueryGatsDisplayName(PrepareRicsForGats(ricsResult)));
            }
            catch (Exception)
            {
                LogMessage("Error in FindOptionList", Logger.LogType.Error);
            }
            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rics"></param>
        /// <returns></returns>
        private Dictionary<string, List<string>> QueryGatsDisplayName(string rics)
        {
            GatsUtil gats = new GatsUtil();
            Dictionary<string, List<string>> gatsValues = new Dictionary<string, List<string>>();
            Regex rgxSpace = new Regex(@"\s+");
            try
            {
                string[] stringSeparators = { "\r\n" };
                char[] stringSeparators2 = { ' ' };

                string test = gats.GetGatsResponse(rics, "DSPLY_NAME,OPINT_1");

                string[] lines = test.Split(stringSeparators, StringSplitOptions.None);
                foreach (string[] lineTab in lines
                        .Select(line => rgxSpace.Replace(line, " "))
                        .Select(formattedLine => formattedLine.Split(stringSeparators2))
                        .Where(lineTab => lineTab.Length > 2 && lineTab[2] != ""))
                {
                    if (!gatsValues.ContainsKey(lineTab[0]))
                    {
                        gatsValues.Add(lineTab[0], new List<string>());
                    }
                    string toAdd = "";
                    for (int count = 2; count < lineTab.Length; count++)
                    {
                        toAdd += lineTab[count] + " ";
                    }
                    gatsValues[lineTab[0]].Add(toAdd.TrimEnd());
                }
                return gatsValues;
            }
            catch (Exception ex)
            {
                LogMessage("Error while using Gats", Logger.LogType.Error);
                throw new Exception("Error While using Gats: " + ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rics"></param>
        /// <returns></returns>
        private List<string> FindMatchingRics(Dictionary<string, List<string>> rics)
        {
            List<string> results = new List<string>();
            foreach (KeyValuePair<string, List<string>> ric in rics)
            {
                if (ric.Value[0].EndsWith("C"))
                {
                    foreach (KeyValuePair<string, List<string>> ric2 in rics.Where(ric2 => ric2.Value[0] == ric.Value[0].Substring(0, ric.Key.Length) + " P")
                                                                            .Where(ric2 => ric2.Value[1] == "0" && ric.Value[1] == "0"
                                                                                && ric.Key.Substring(0, 5) == ric2.Key.Substring(0, 5)))
                    {
                        results.Add(ric.Key);
                        results.Add(ric2.Key);
                    }
                }
            }
            return results;
        }

        #endregion

        #region Write in CSV

        /// <summary>
        /// 
        /// </summary>
        /// <param name="results"></param>
        private void WriteResultsInCsv(IEnumerable<string> results)
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.WorkingFolder + "result.csv");
                try
                {
                    Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                    ExcelLineWriter lineWriter = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right);
                    lineWriter.PlaceNext(1, 1);
                    lineWriter.WriteLine("RIC");
                    lineWriter.WriteLine("RETIRE DATE");
                    int line = 2;
                    foreach (string entry in results)
                    {
                        lineWriter.PlaceNext(line, 1);
                        lineWriter.WriteLine(entry);
                        lineWriter.WriteLine(DateTime.Now.ToString("dd-MMM-yyyy"));
                        line++;
                    }
                }
                catch (Exception ex)
                {
                    LogMessage("Error while creating CSV", Logger.LogType.Error);
                    throw new Exception("Error when creating result CSV: " + ex.Message);
                }
                finally
                {
                    AddResult("TW IO Drop results", configObj.WorkingFolder + "result.csv", "file");
                    workbook.SaveAs(configObj.WorkingFolder + "result.csv", XlFileFormat.xlCSV);
                    workbook.Close();
                }
            }
        }

        #endregion

        #region Write in Txt

        /// <summary>
        /// Write result rics in T
        /// </summary>
        /// <param name="results"></param>
        private void WriteResultsInTxt(IEnumerable<string> results, string filename = "results.txt")
        {
            FileStream fs = null;
            StreamWriter sw = null;

            try
            {
                fs = new FileStream(configObj.WorkingFolder + filename, FileMode.Append);
                sw = new StreamWriter(fs, Encoding.UTF8);
                sw.WriteLine("RIC");
                foreach (string ric in results)
                {
                    sw.WriteLine(ric);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex;
                LogMessage(msg, Logger.LogType.Error);
                throw new FileNotFoundException("The file " + configObj.WorkingFolder + filename + " cannot be open, please close it first.");
            }
            finally
            {
                sw.Close();
                fs.Close();
                AddResult("TW IO Drop result file", configObj.WorkingFolder + filename, "file");
            }
        }

        #endregion
    }

    #endregion
}
