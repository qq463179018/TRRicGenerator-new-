using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.China
{
    [ConfigStoredInDB]
    public class ChinaIndicesAdjustmentConfig
    {
        [StoreInDB]
        [DisplayName("Source Ric")]
        [Description("Full path of file 'RICs in 3000.xls' ")]
        public string SourceRicIn3000 { get; set; }

        [StoreInDB]
        [DisplayName("Worksheet source")]
        public string WorksheetSource { get; set; }

        [StoreInDB]
        [DisplayName("Target filename")]
        [Description("The target file ")]
        public string TargetIndicesAdjustmentFileName{get;set;}

        [StoreInDB]
        [DisplayName("Worksheet Add in")]
        public string WorksheetAddIn{get;set;}

        [StoreInDB]
        [DisplayName("Worksheet move out")]
        public string WorksheetMoveOut{get;set;}

        [StoreInDB]
        [DisplayName("CSI adjustment")]
        [Description("Full path of the file which contains all the add/move rics")]
        public string CsiAdjustment{ get; set; }

        [StoreInDB]
        [DisplayName("Chain number per sheet")]
        public int ChainNumPerSheet { get; set; }

    }

    public class ChinaIndicesAdjustment : GeneratorBase
    {
        private static ChinaIndicesAdjustmentConfig configObj;

        protected override void Start()
        {
            StartIndicesAdjustment();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as ChinaIndicesAdjustmentConfig;
        }

        public void StartIndicesAdjustment()
        {
            Dictionary<string, List<string>> newIndicesRics = GetNewIndiceRics();
            GenerateIndiceRicFile(newIndicesRics);
        }


        public void GenerateIndiceRicFile(Dictionary<string, List<string>> indicesRicDic)
        {
            using (ExcelApp app = new ExcelApp(false,false))
            {
                int startPos = 0;
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.TargetIndicesAdjustmentFileName);
                int sheetNum = (indicesRicDic.Keys.Count+configObj.ChainNumPerSheet-1)/configObj.ChainNumPerSheet;
                if (sheetNum > workbook.Worksheets.Count)
                {
                    workbook.Worksheets.Add(Missing.Value, Missing.Value, sheetNum - workbook.Worksheets.Count, Missing.Value);
                }

                for(int i=0; i<sheetNum;i++)
                {
                    var worksheet = workbook.Worksheets[i+1] as Worksheet;
                    startPos = i * configObj.ChainNumPerSheet;
                    int endPos = indicesRicDic.Keys.Count < (startPos + configObj.ChainNumPerSheet) ? indicesRicDic.Keys.Count : (startPos + configObj.ChainNumPerSheet);
                    WriterWorksheet(worksheet, indicesRicDic, startPos, endPos);
                }

                TaskResultList.Add(new TaskResultEntry("New Indices File", "", workbook.FullName));
                workbook.Close(true, workbook.FullName, true);
            }
        }

        public void WriterWorksheet(Worksheet worksheet,Dictionary<string, List<string>> indicesRicDic,int startPos, int endPos)
        {
            using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Down))
            {
                for (int i = startPos; i < endPos; i++)
                {
                    string key = indicesRicDic.Keys.ToList()[i];
                    writer.WriteLine("0#" + key);
                    writer.WriteLine(key);
                    foreach (string ric in indicesRicDic[key])
                    {
                        writer.WriteLine(ric);
                    }
                    writer.PlaceNext(1, writer.Col+1);
                }
            }
        }

        public Dictionary<string, List<string>> GetNewIndiceRics()
        {
            Dictionary<string, List<string>> newIndicesRics = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> addRicDic = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> moveRicDic = new Dictionary<string, List<string>>();
            using (ExcelApp app = new ExcelApp(false,false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.CsiAdjustment);
                var worksheetAdd = ExcelUtil.GetWorksheet(configObj.WorksheetAddIn, workbook);
                if (worksheetAdd == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", worksheetAdd.Name, workbook.FullName));
                }
                addRicDic = GetUpdateIndiceRics(worksheetAdd);

                var worksheetMove = ExcelUtil.GetWorksheet(configObj.WorksheetMoveOut, workbook);
                if (worksheetMove == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", worksheetMove.Name, workbook.FullName));
                }

                moveRicDic = GetUpdateIndiceRics(worksheetMove);
            }
            Dictionary<string, List<string>> sourceRicDic = GetSourceIndiceRics();

            foreach (string key in addRicDic.Keys)
            {
                if (!sourceRicDic.ContainsKey(key))
                {
                    Logger.Log(string.Format("Add indices: There's no such chain {0} in the source rics in 3000", key));
                }
                else
                {
                    sourceRicDic[key].AddRange(addRicDic[key]);
                    if (newIndicesRics.ContainsKey(key))
                    {
                        Logger.Log(string.Format("There's duplicate chain {0} in add chains", key));
                    }
                    else
                    {
                        sourceRicDic[key].Sort();
                        newIndicesRics.Add(key, sourceRicDic[key]);
                    }
                    //foreach (string ric in addRicDic[key])
                    //{
                    //    sourceRicDic[key].Add(ric);
                    //}
                }
            }

            foreach (string key in moveRicDic.Keys)
            {
                if (!sourceRicDic.ContainsKey(key))
                {
                    Logger.Log(string.Format("Move indices: There's no such chain {0} in the source rics in 3000", key));
                }
                else
                {
                    if (newIndicesRics.ContainsKey(key))
                    {
                        foreach (string ric in moveRicDic[key])
                        {
                            if (newIndicesRics[key].Contains(ric))
                            {
                                newIndicesRics[key].Remove(ric);
                            }
                            else
                            {
                                Logger.Log(string.Format("Move Indices: There's no such ric {0} in source rics in 3000 for chain {1}", ric,key));
                            }
                        }
                    }

                    else
                    {
                        foreach (string ric in moveRicDic[key])
                        {
                            if (sourceRicDic[key].Contains(ric))
                            {
                                sourceRicDic[key].Remove(ric);
                            }
                            else
                            {
                                Logger.Log(string.Format("Move Indices: There's no such ric {0} in source rics in 3000 for such {1}", ric, key));
                            }
                        }
                        sourceRicDic[key].Sort();
                        newIndicesRics.Add(key, sourceRicDic[key]);
                    }
                }
            }
            return newIndicesRics;
        }

        //Get Source Rics from 3000
        public Dictionary<string, List<string>> GetSourceIndiceRics()
        {
            Dictionary<string, List<string>> sourceIndiceRicDic = new Dictionary<string, List<string>>();
            using (ExcelApp app = new ExcelApp(false,false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.SourceRicIn3000);
                var worksheet = ExcelUtil.GetWorksheet(configObj.WorksheetSource, workbook);
                if (worksheet == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", worksheet.Name, workbook.FullName));
                }

                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                int lastUsedCol = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;

                List<string> rics = new List<string>();
                string chain = ExcelUtil.GetRange(1, 1, worksheet).Value2.ToString().Trim();

                for (int i = 0; i < lastUsedCol; i++)
                {
                    for (int j = 0; j < lastUsedRow; j++)
                    {
                        Range r = ExcelUtil.GetRange(j + 1, i + 1, worksheet);
                        if (r.Value2 != null && (r.Value2.ToString() != string.Empty))
                        {
                            string cellValue = r.Value2.ToString().Trim();
                            if (cellValue.StartsWith("."))
                            {
                                if (!sourceIndiceRicDic.ContainsKey(cellValue))
                                {
                                    if (j != 0)
                                    {
                                        try
                                        {
                                            if (!sourceIndiceRicDic.ContainsKey(chain))
                                            {
                                                sourceIndiceRicDic.Add(chain, rics);
                                                rics = null;
                                                rics = new List<string>();
                                                chain = cellValue;
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Logger.Log(string.Format("There's duplicate chain {0} in source rics in 3000: {1}", chain, ex.Message));
                                        }
                                    }
                                }

                                else
                                {
                                    Logger.Log(string.Format("There's duplicate chain {0} in the source files.",cellValue));
                                }
                            }
                            else
                            {
                                rics.Add(cellValue);
                            }

                        }
                    }
                }
                sourceIndiceRicDic.Add(chain, rics);

                workbook.Close(false, workbook.FullName, true);
            }

            return sourceIndiceRicDic;
        }

        //Get Add-in and Move-out Rics
        public Dictionary<string, List<string>> GetUpdateIndiceRics(Worksheet worksheet)
        {
            Dictionary<string, List<string>> updatedIndiceRicDic = new Dictionary<string, List<string>>();
            int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            List<string> rics = new List<string>();
            string chain = ExcelUtil.GetRange(2, 2, worksheet).Value2.ToString().Replace("0#","").Trim();

            for (int i = 2; i <= lastUsedRow; i++)
            {
                Range chainRange = ExcelUtil.GetRange(i, 2, worksheet);
                Range ricRange = ExcelUtil.GetRange(i, 1, worksheet);
                if (chainRange.Value2 == null || chainRange.Value2.ToString() == string.Empty || ricRange.Value2 == null || ricRange.Value2.ToString() == string.Empty)
                {
                    //Logger.Log(string.Format("There's a blank cell in Row {0} in worksheet {1}", i.ToString(), worksheet.Name));
                }
                else
                {
                    string cellValue = chainRange.Value2.ToString().Replace("0#","").Trim();
                    if (!updatedIndiceRicDic.ContainsKey(chain))
                    {
                        if (chain != cellValue)
                        {
                            updatedIndiceRicDic.Add(chain, rics);
                            rics = new List<string> {ricRange.Value2.ToString().Trim()};
                            chain = cellValue;
                        }
                        else
                        {
                            rics.Add(ricRange.Value2.ToString().Trim());
                        }
                    }

                    else
                    {
                        rics.Add(ricRange.Value2.ToString().Trim());
                    }
                }
            }

            if (updatedIndiceRicDic.ContainsKey(chain))
            {
                Logger.Log(string.Format("There's a duplicate chain {0} in worksheet {1}", chain, worksheet.Name));
            }

            else
            {
                updatedIndiceRicDic.Add(chain, rics);
            }

            return updatedIndiceRicDic;
        }
    }
}
