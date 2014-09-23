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
    public class TWIndexMaintenanceConfig
    {
        [StoreInDB]
        [Description("The path of the working folder, should end with a slash\n\nEg: C:\\Mydrive\\")]
        public string WorkingFolder { get; set; }

        [StoreInDB]
        [Description("The name of the file containing sector chain list \n\nEg: SectorTW.txt")]
        public string SectorChainList { get; set; }

        [StoreInDB]
        [DisplayName("Gtsm file name")]
        [Description("The name of the file containing gtsm list \n\nEg: gtsm.txt")]
        public string GtsmFile { get; set; }

        [StoreInDB]
        [DisplayName("Twse file name")]
        [Description("The name of the file containing twse list \n\nEg: twse.xls")]
        public string TwseFile { get; set; }
    }

    #endregion

    #region Task

    class TWIndexMaintenance : GeneratorBase
    {
        #region Declaration

        private static TWIndexMaintenanceConfig _configObj;
        private ExcelApp _app;
        private Workbook _workbook;

        #endregion

        #region Interface implementation

        protected override void Initialize()
        {
            base.Initialize();

            _configObj = Config as TWIndexMaintenanceConfig;
            _app = new ExcelApp(false, false);
            if (_app.ExcelAppInstance == null)
            {
                LogMessage("Excel cannot be started", Logger.LogType.Error);
            }
            _workbook = null;
        }

        protected override void Start()
        {
            try
            {
                Dictionary<string, string> twseList = GetTwseList();
                Dictionary<string, string> gtsmList = GetGtsmList();
                _workbook = ExcelUtil.CreateOrOpenExcelFile(_app, _configObj.WorkingFolder + _configObj.SectorChainList);

                LogMessage("Get Twse sector");
                Dictionary<string, string> sectorsTwse = GetTwseSector();
                Dictionary<string, string> resultsTwse = GetResultsTwse(twseList, sectorsTwse);
                Dictionary<string, string> oldTwse = GetDbList(sectorsTwse);

                LogMessage("Look for deletions in Twse sector");
                Dictionary<string, string> duplicatesDBTwse = FindDuplicates(oldTwse);
                OutPutDictionadyTest(duplicatesDBTwse, Path.Combine(_configObj.WorkingFolder, "IDNDuplicates(Twse sector).csv"));
                AddResult("IDNDuplicates", Path.Combine(_configObj.WorkingFolder, "IDNDuplicates(Twse sector).csv"), "CSV bulk file");

                Dictionary<string, string> duplicatesWebsiteTwse = FindDuplicates(resultsTwse);
                OutPutDictionadyTest(duplicatesWebsiteTwse, Path.Combine(_configObj.WorkingFolder, "WebsiteDuplicates(Twse sector).csv"));
                AddResult("WebsiteDuplicates", Path.Combine(_configObj.WorkingFolder, "WebsiteDuplicates(Twse sector).csv"), "CSV bulk file");

                Dictionary<string, string> WebsiteRemovedIDNTwse = CompareWithDb(resultsTwse, oldTwse);
                OutPutDictionadyTest(WebsiteRemovedIDNTwse, Path.Combine(_configObj.WorkingFolder, "WebsiteRemovedIDN(Twse sector).csv"));
                AddResult("WebsiteRemovedIDN", Path.Combine(_configObj.WorkingFolder, "WebsiteRemovedIDN(Twse sector).csv"), "CSV bulk file");

                Dictionary<string, string> IDNRemovedWebsiteTwse = CompareWithDb(oldTwse, resultsTwse);
                OutPutDictionadyTest(IDNRemovedWebsiteTwse, Path.Combine(_configObj.WorkingFolder, "IDNRemovedWebsite(Twse sector).csv"));
                AddResult("IDNRemovedWebsite", Path.Combine(_configObj.WorkingFolder, "IDNRemovedWebsite(Twse sector).csv"), "CSV bulk file");
                //LogMessage("Look for deletions in Twse sector");
                //OutPutDictionadyTest(resultsTwse, Path.Combine(_configObj.WorkingFolder, "resultsList(Twse sector).csv"));
                //OutPutDictionadyTest(oldTwse, Path.Combine(_configObj.WorkingFolder, "dbList(Twse sector).csv"));
                //Dictionary<string, string> updateTwse = CompareWithDb(resultsTwse, oldTwse);
                //OutPutDictionadyTest(updateTwse, Path.Combine(_configObj.WorkingFolder, "resultsListRemovedbList(Twse sector).csv"));
                //Dictionary<string, string> updateTwse1 = CompareWithDb(oldTwse, resultsTwse);
                //OutPutDictionadyTest(updateTwse1, Path.Combine(_configObj.WorkingFolder, "dbListRemoveresultsList(Twse sector).csv"));
                //WriteDeleteInTxt(updateTwse, true);
                //LogMessage("Look for additions in Twse sector");
                //WriteAddition(FindAddition(resultsTwse, oldTwse), String.Format("twse_additions_{0}.txt", DateTime.Now.ToString("ddMMMHHmm")));

                LogMessage("Get Gtsm sector");
                Dictionary<string, string> sectorsGtsm = GetGtsmSector();
                Dictionary<string, string> resultsGtsm = GetResultsGtsm(gtsmList, sectorsGtsm);
                Dictionary<string, string> oldGtsm = GetDbList(sectorsGtsm);

                LogMessage("Look for deletions in Gtsm sector");
                Dictionary<string, string> duplicatesDBGtsm = FindDuplicates(oldGtsm);
                OutPutDictionadyTest(duplicatesDBGtsm, Path.Combine(_configObj.WorkingFolder, "IDNDuplicates(Gtsm sector).csv"));
                AddResult("IDNDuplicates", Path.Combine(_configObj.WorkingFolder, "IDNDuplicates(Gtsm sector).csv"), "CSV bulk file");

                Dictionary<string, string> duplicatesWebsiteGtsm = FindDuplicates(resultsGtsm);
                OutPutDictionadyTest(duplicatesWebsiteGtsm, Path.Combine(_configObj.WorkingFolder, "WebsiteDuplicates(Gtsm sector).csv"));
                AddResult("WebsiteDuplicates", Path.Combine(_configObj.WorkingFolder, "WebsiteDuplicates(Gtsm sector).csv"), "CSV bulk file");

                Dictionary<string, string> WebsiteRemovedIDNGtsm = CompareWithDb(resultsGtsm, oldGtsm);
                OutPutDictionadyTest(WebsiteRemovedIDNGtsm, Path.Combine(_configObj.WorkingFolder, "WebsiteRemovedIDN(Gtsm sector).csv"));
                AddResult("WebsiteRemovedIDN", Path.Combine(_configObj.WorkingFolder, "WebsiteRemovedIDN(Gtsm sector).csv"), "CSV bulk file");

                Dictionary<string, string> IDNRemovedWebsiteGtsm = CompareWithDb(oldGtsm, resultsGtsm);
                OutPutDictionadyTest(IDNRemovedWebsiteGtsm, Path.Combine(_configObj.WorkingFolder, "IDNRemovedWebsite(Gtsm sector).csv"));
                AddResult("IDNRemovedWebsite", Path.Combine(_configObj.WorkingFolder, "IDNRemovedWebsite(Gtsm sector).csv"), "CSV bulk file");

                //LogMessage("Look for deletions in Gtsm sector");
                //OutPutDictionadyTest(resultsTwse, Path.Combine(_configObj.WorkingFolder, "resultsList(Gtsm sector).csv"));
                //OutPutDictionadyTest(oldTwse, Path.Combine(_configObj.WorkingFolder, "dbList(Gtsm sector).csv"));
                //Dictionary<string, string> updateGtsm = CompareWithDb(resultsGtsm, oldGtsm);
                //OutPutDictionadyTest(updateGtsm, Path.Combine(_configObj.WorkingFolder, "resultsListRemovedbList(Gtsm sector).csv"));
                //Dictionary<string, string> updateGtsm1 = CompareWithDb(oldGtsm, resultsGtsm);
                //OutPutDictionadyTest(updateGtsm1, Path.Combine(_configObj.WorkingFolder, "dbListRemoveresultsList(Gtsm sector).csv"));
                //WriteDeleteInTxt(updateGtsm);
                //LogMessage("Look for additions in Gtsm sector");
                //WriteAddition(FindAddition(resultsGtsm, oldGtsm), String.Format("gtsm_additions_{0}.txt", DateTime.Now.ToString("ddMMMHHmm")));
            }
            catch (Exception ex)
            {
                throw new Exception("Error while running task: " + ex.Message, ex);
            }
            finally
            {
                _workbook.Close();
                _app.Dispose();
            }
        }

        #endregion

        #region Searching DB

        private Dictionary<string, string> GetDbList(Dictionary<string, string> sectors)
        {
            string gatsRics = String.Empty;

            foreach (var indice in from sector in sectors
                                   from indice in sector.Value.Split(new[] { ',' }, StringSplitOptions.None)
                                   where !gatsRics.Contains(indice.Trim())
                                   select indice)
            {
                for (int index = 0; index < 80; index++)
                    gatsRics += String.Format("{0}#{1},", index, indice.Trim());
            }

            var gats = new GatsUtil(GatsUtil.Server.Idn);
            var gatsValues = new Dictionary<string, string>();
            var rgxSpace = new Regex(@"\s+");
            var rgxPre = new Regex(@"^LINK_", RegexOptions.IgnoreCase);

            string[] stringSeparators = { "\r\n" };
            char[] stringSeparators2 = { ' ' };

            string test = gats.GetGatsResponse(gatsRics, "LINK_1,LINK_2,LINK_3,LINK_4,LINK_5,LINK_6,LINK_7,LINK_8,LINK_9,LINK_10,LINK_11,LINK_12,LINK_13,LINK_14");

            string[] lines = test.Split(stringSeparators, StringSplitOptions.None);
            foreach (string[] lineTab in lines.Select(line => rgxSpace.Replace(line, " "))
                                              .Select(formattedLine => formattedLine.Split(stringSeparators2))
                                              .Where(lineTab => lineTab.Length > 2
                                                    && rgxPre.IsMatch(lineTab[1])
                                                    && lineTab[2] != ""
                                                    && !lineTab[2].StartsWith(".")))
            {
                if (!gatsValues.ContainsKey(lineTab[2].Trim()))
                {
                    gatsValues.Add(lineTab[2].Trim(), lineTab[0].Trim().Substring(lineTab[0].Trim().IndexOf("#") + 1));
                }
                else
                {
                    gatsValues[lineTab[2].Trim()] += "," + lineTab[0].Trim().Substring(lineTab[0].Trim().IndexOf("#") + 1);
                }
            }
            return gatsValues;
        }

        #endregion

        #region DB comparasion

        private static int CompareResultByLen(List<string> list1, List<string> list2)
        {
            if (list1.Count == list2.Count)
            {
                return 0;
            }
            if (list2.Count > list1.Count)
            {
                return 1;
            }
            return -1;
        }

        private List<List<string>> FindAddition(Dictionary<string, string> resultsList, Dictionary<string, string> dbList)
        {
            var additions = new Dictionary<string, List<string>>();
            var additionsList = new List<List<string>>();
            if (dbList == null)
            {
                return null;
            }
            foreach (var result in resultsList)
            {
                if (!dbList.ContainsKey(result.Key))
                {
                    foreach (var indice in result.Value.Split(new[] { ',' }, StringSplitOptions.None))
                    {
                        if (!additions.ContainsKey(indice.Trim()))
                        {
                            additions.Add(indice.Trim(), new List<string>());
                        }
                        additions[indice.Trim()].Add(result.Key.Trim());
                    }
                }
                else
                {
                    foreach (var indice in result.Value.Split(new[] { ',' }, StringSplitOptions.None).Where(indice => !dbList[result.Key].Contains(indice.Trim())))
                    {
                        if (!additions.ContainsKey(indice))
                        {
                            additions.Add(indice.Trim(), new List<string>());
                        }
                        additions[indice.Trim()].Add(result.Key.Trim());
                    }
                }
            }
            foreach (var addition in additions)
            {
                var newAdd = new List<string> { addition.Key };
                newAdd.AddRange(addition.Value);
                additionsList.Add(newAdd);
            }
            additionsList.Sort(CompareResultByLen);
            return additionsList;
        }

        /// <summary>
        /// find the duplicates value
        /// </summary>
        /// <param name="Dic"></param>
        /// <returns></returns>
        private Dictionary<string, string> FindDuplicates(Dictionary<string, string> Dic)
        {
            var result = new Dictionary<string, string>();

            if (Dic == null || Dic.Count == 0)
                return null;

            foreach (var item in Dic)
            {
                List<string> duplicateValue = FindDuplicatesLine(item.Value);
                if (duplicateValue == null || duplicateValue.Count == 0)
                    continue;

                string value = string.Empty;
                foreach (var itemDup in duplicateValue)
                {
                    value += string.Format(",{0}", itemDup);
                }
                value.Remove(0, 1);
                result.Add(item.Key, value);
            }

            return result;
        }

        private List<string> FindDuplicatesLine(string p)
        {
            List<string> resule = new List<string>();

            if ((p + "").Trim().Length == 0 || !p.Contains(","))
                return null;

            List<string> list = p.Split(',').ToList();
            List<string> tmp = new List<string>();
            foreach (var item in list)
            {
                if ((item + "").Trim().Length == 0)
                    continue;

                if (tmp.Contains(item))
                    resule.Add(item);
                else
                    tmp.Add(item);
            }

            return resule;
        }

        /// <summary>
        /// remove second from the first
        /// </summary>
        /// <param name="resultsDic"></param>
        /// <param name="dbList"></param>
        /// <returns></returns>
        private Dictionary<string, string> CompareWithDb(Dictionary<string, string> resultsDic, Dictionary<string, string> dbList)
        {

            var resultsList = CloneDic(resultsDic);
            var toUpdate = new Dictionary<string, string>();

            if (dbList == null || dbList.Count == 0)
                return resultsList;

            #region [Clear exist from dbList]
            List<string> lineDBListValue = null;
            List<string> lineResultListValue = null;
            foreach (var item in dbList)
            {
                if (!resultsList.ContainsKey(item.Key))
                    continue;

                lineDBListValue = TrimFormat(item.Value.Split(',').ToList());
                lineResultListValue = TrimFormat(resultsList[item.Key].Split(',').ToList());
                foreach (var lineItem in lineDBListValue)
                {
                    int count = lineResultListValue.Count;
                    for (int i = 0; i < count; i++)
                    {
                        if (!lineResultListValue.Contains(lineItem))
                            break;

                        lineResultListValue.Remove(lineItem);
                    }
                }

                if (lineResultListValue == null || lineResultListValue.Count == 0)
                {
                    resultsList[item.Key] = string.Empty;
                    continue;
                }

                string newLine = string.Empty;
                foreach (var listItem in lineResultListValue)
                {
                    newLine += string.Format(",{0}", listItem);
                }
                newLine.Remove(0, 1);
                resultsList[item.Key] = newLine;
            }
            #endregion

            foreach (var item in resultsList)
            {
                if ((item.Value + "").Trim().Length == 0)
                    continue;

                toUpdate.Add(item.Key, item.Value);
            }

            return toUpdate;
        }

        private Dictionary<string, string> CloneDic(Dictionary<string, string> resultsDic)
        {
            if (resultsDic == null)
                return null;

            var result = new Dictionary<string, string>();
            foreach (var item in resultsDic)
            {
                result.Add(item.Key, item.Value);
            }

            return result;
        }

        private List<string> TrimFormat(List<string> list)
        {
            var result = new List<string>();

            if (list == null || list.Count == 0)
                return list;

            foreach (var item in list)
            {
                if ((item + "").Trim().Length == 0)
                    continue;

                result.Add(item.Trim());
            }

            return result;
        }

        private void OutPutDictionadyTest(Dictionary<string, string> dic, string path)
        {
            try
            {
                string outputString = " ";
                StringBuilder sb = new StringBuilder();

                if (dic != null && dic.Count > 0)
                {
                    foreach (var item in dic)
                    {
                        sb.AppendFormat("{0},{1}\r\n", item.Key, item.Value);
                    }

                    outputString = sb.ToString();
                }

                FileUtil.WriteOutputFile(path, outputString);
            }
            catch (Exception ex)
            {
                string msg = string.Format("msg:{0}", ex.Message);
                LogMessage(msg);
                throw new Exception(msg);
            }
        }

        #endregion

        #region GetTwseList

        private Dictionary<string, string> GetTwseSector()
        {
            List<List<string>> twseSectorXls = _workbook.ToList();

            return twseSectorXls.Where(sector => sector[1] != null && sector[1] != "Industry" && sector[6] != null && sector[6] != "")
                                .ToDictionary(sector => sector[1], sector => sector[6]);
        }

        private Dictionary<string, string> GetTwseList()
        {
            var twseList = new Dictionary<string, string>();
            var twseWorkbook = ExcelUtil.CreateOrOpenExcelFile(_app, _configObj.WorkingFolder + _configObj.TwseFile);
            try
            {
                List<List<string>> twseValues = twseWorkbook.ToList();
                foreach (var twseValue in twseValues.Skip(10))
                {
                    twseList.Add(twseValue[0].Trim(), twseValue[3].Trim());
                }
            }
            catch (Exception ex)
            {
                string msg = "Error :" + ex;
                LogMessage(msg, Logger.LogType.Error);
            }
            return twseList;
        }

        #endregion

        #region GetGtsmList

        private Dictionary<string, string> GetGtsmSector()
        {
            List<List<string>> gtsmValues = _workbook.ToList(3);

            return gtsmValues.Where(sector => sector[0] != null && sector[0] != "" && sector[6] != null && sector[6] != "")
                             .ToDictionary(sector => sector[1].Trim(), sector => sector[6].Trim());
        }

        private Dictionary<string, string> GetGtsmList()
        {
            var gtsmList = new Dictionary<string, string>();

            var sr = new StreamReader(_configObj.WorkingFolder + _configObj.GtsmFile);
            string line = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                int index = 0;
                var charline = line.ToCharArray();
                var code = string.Empty;
                var sector = string.Empty;
                foreach (var linechar in charline)
                {
                    if (linechar >= '0' && linechar <= '9')
                    {
                        code += linechar;
                    }
                    else
                    {
                        break;
                    }
                    index++;
                }
                for (; charline[index] < '0' || charline[index] > '9'; index++)
                {

                }
                for (int newindex = index; newindex != index + 2; newindex++)
                {
                    sector += charline[newindex];
                }

                //gtsmList.Add(code, sector);
                gtsmList.Add(code, RemoveZeroStart(sector));
            }

            sr.Close();

            return gtsmList;
        }

        #endregion

        #region Get results

        private Dictionary<string, string> GetResultsTwse(Dictionary<string, string> twseList, Dictionary<string, string> twseSector)
        {
            var results = new Dictionary<string, string>();
            foreach (var twse in twseList)
            {
                string code = twse.Value;
                if (code.StartsWith("0"))
                {
                    code = code.Substring(1);
                }
                results.Add(twse.Key + ".TW", twseSector[code]);
            }
            return results;
        }

        private Dictionary<string, string> GetResultsGtsm(Dictionary<string, string> gtsmList, Dictionary<string, string> gtsmSector)
        {
            var results = new Dictionary<string, string>();
            foreach (var gtsm in gtsmList)
            {
                string code = gtsm.Value;
                code = code.Replace("Plastic", "Plastics").Replace("Electric Machinery", "Machinery").Replace("Electrical and Cable", "Ele Applicance & Cable");
                code = code.Replace("Iron and Steel", "Steek & Iron").Replace("Building Material and Construction", "Construction").Replace("Shipping and Transportation", "Transportation").Replace("Tourism", "Hotel & Tourism").Replace("Financial", "Bank & Insurance").Replace("Trading and Consumers Goods", "Department Stores").Replace("Other", "Others").Replace("Biotechnology and Medical Care", "Biotech & Medical Care").Replace("Gas and Electricity", "Oil, Gas and Electricity");
                code = code.Replace("Computer and Peripheral Equipment", "Computer & Peripheral Eq").Replace("Communications and Internet", "*Comms & Internet").Replace("Electronic Parts and Components", "Ele. Parts/Components").Replace("Electronic Products Distribution", "Ele.Products Distribution").Replace("Other Electronic Industries", "Other Electronic").Replace("Information Service", "*Information Services").Replace("Chemical", "*Chemical");
                if (gtsmSector.ContainsKey(code))
                {
                    results.Add(gtsm.Key + ".TWO", gtsmSector[code]);
                }
            }
            return results;
        }

        private string RemoveZeroStart(string code)
        {
            if (string.IsNullOrEmpty(code))
                return code;

            while (code.StartsWith("0"))
                code = code.Remove(0, 1);

            return code;
        }

        #endregion

        #region Write DB/Resulst/Deletions in Txt

        /// <summary>
        /// Write result rics in T
        /// </summary>
        /// <param name="toWrite"></param>
        /// <param name="firstTime"></param>
        private void WriteDeleteInTxt(Dictionary<string, string> toWrite, bool firstTime = false)
        {
            string filename = String.Format(firstTime ? "twse_deletions_{0}.txt" : "gtsm_deletions_{0}.txt", DateTime.Now.ToString("ddMMMHHmm"));

            var indiceRicDic = new Dictionary<string, List<string>>();

            foreach (var ric in toWrite)
            {
                foreach (var indice in ric.Value.Split(','))
                {
                    if (!indiceRicDic.ContainsKey(indice))
                    {
                        indiceRicDic.Add(indice, new List<string>());
                    }
                    indiceRicDic[indice].Add(ric.Key);
                }
            }

            var toWriteFile = new List<List<string>>();
            foreach (var indice in indiceRicDic)
            {
                var toAddList = new List<string> { indice.Key };
                toAddList.AddRange(indice.Value);
                toWriteFile.Add(toAddList);
            }

            WriteAddition(toWriteFile, filename);
        }

        /// <summary>
        /// Write result rics in T
        /// </summary>
        /// <param name="toWrite"></param>
        /// <param name="filename"></param>
        private void WriteAddition(List<List<string>> toWrite, string filename)
        {
            if (toWrite != null && toWrite.Count != 0)
            {
                var fs = new FileStream(_configObj.WorkingFolder + filename, FileMode.OpenOrCreate);
                var sw = new StreamWriter(fs, Encoding.UTF8);
                for (int i = 0; i < toWrite[0].Count; i++)
                {
                    var sb = new StringBuilder();
                    foreach (var col in toWrite.TakeWhile(col => i < col.Count))
                    {
                        if (i == 0 && filename.ToLower().Contains("twse"))
                        {
                            sb.AppendFormat("I_TAIW_{0}_CHAIN\t", col[0].Replace(".", ""));
                        }
                        else if (i == 0 && filename.ToLower().Contains("gtsm"))
                        {
                            sb.AppendFormat("I_OTCTWS_{0}_CHAIN\t", col[0].Replace(".", ""));
                        }
                        else
                        {
                            sb.Append(col[i] + "\t");
                        }
                    }
                    sw.WriteLine(sb.ToString());
                }
                sw.Close();
                fs.Close();
                string fileDesc = String.Empty;
                fileDesc += filename.Contains("twse") ? "Twse " : "Gtsm ";
                fileDesc += filename.Contains("additions") ? "additions " : "deletions ";
                fileDesc += "file";
                AddResult(fileDesc, _configObj.WorkingFolder + filename, "file");
            }
        }

        #endregion
    }

    #endregion
}
