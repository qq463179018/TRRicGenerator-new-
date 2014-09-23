using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Threading;
using Ric.Db.Manager;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Validation
{
    [ConfigStoredInDB]
    public class ValidateXtronFMWithGATSConfig
    {
        [StoreInDB]
        [Category("Source File")]
        [Description("the full path of the FM file.")]
        public string FMFilePath { get; set; }
        
        [StoreInDB]
        [Category("Output")]
        [Description("the full path of the validation result file.")]
        public string ValidationResultFilePath{get;set;}

        [StoreInDB]
        [Category("Output")]
        [Description("the output directory.")]
        public string OutputDir { get; set; }

        [StoreInDB]
        [Category("Output")]
        [Description("the folder of the output rics.")]
        public string OutRicsFoldr { get; set; }

        [StoreInDB]
        [Category("Output")]
        [Description("the folder of the output ric fields.")]
        public string RicFieldsFoldr { get; set; }

        [StoreInDB]
        [Category("Source File")]
        [Description("the full path of the txt file which contain the ric fields of the chainric.")]
        public string RicFidsPath { get; set; }

        [StoreInDB]
        [Category("Source File")]
        [Description("the full path of the txt file which contain the fields of the ric.")]
        public string FidListPath { get; set; }

        [StoreInDB]
        [DefaultValue("10.40.247.103")]
        [Category("GATS Server")]
        [Description("GATS Server IP address.")]
        public string IP { get; set; }
    }
    public class ChainRicRecord
    {
        public List<RicRecord> RicList { get; set; }
        public MainRicRecord MainRic { get; set; }

        public ChainRicRecord(MainRicRecord mainRic)
        {
            MainRic = mainRic;
        }
    }

    public class RicRecord
    {
        public string Ric { get; set; }
        public string HaveFields { get; set; }
        public RicRecord(string ric, string haveFields)
        {
            Ric = ric;
            HaveFields = haveFields;
        }
    }

    public class MainRicRecord
    {
        public string Ric { get; set; }
        public string HaveFields { get; set; }
        public List<string> FieldsList { get; set; }

        public MainRicRecord(string ric)
        {
            Ric = ric;
            HaveFields = "0";
        }
    }

    public class ValidateXtronFMWithGATSGenerator : GeneratorBase
    {
        private ValidateXtronFMWithGATSConfig configObj = null;
        private string outputPath = "";
        private static readonly object syncRoot = new object();
        private List<ChainRicRecord> chainRicList = new List<ChainRicRecord>();
        private string[] fields = new string[9];
        private string gatsServer = string.Empty;

        protected override void Initialize()
        {
            base.Initialize();
            try
            {
                configObj = Config as ValidateXtronFMWithGATSConfig;
                outputPath = GetOutputFilePath();
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when initializing task... Ex: {0} .", ex.Message));
            }
            if (!string.IsNullOrEmpty(configObj.IP))
            {
                gatsServer = configObj.IP;
            }
            else
            {
                gatsServer = ConfigureOperator.GetGatsServer();
            }
        }

        protected override void Start()
        {
            try
            {
                StartValidateXtronFMWithGATSGenertor();
            }

            catch (Exception ex)
            {
                Logger.LogErrorAndRaiseException(string.Format("Ex: {0} .\nStack Trace: {1} .", ex.Message, ex.StackTrace));
            }
        }

        private void StartValidateXtronFMWithGATSGenertor()
        {
            List<string> ricList = new List<string>();
            try
            {
                ricList = GetMainRic(configObj.FMFilePath);
                foreach(string ric in ricList)
                {
                    int order = 0;
                    string chainRic = string.Format("{0}#{1}", order, ric);
                    string lineText;
                    do
                    {
                        string outRicsFilePath;
                        GetRics(chainRic, out outRicsFilePath);
                        StreamReader sr = new StreamReader(outRicsFilePath);
                        lineText = sr.ReadToEnd();
                        sr.Close();
                        order++;
                        chainRic = string.Format("{0}#{1}", order, ric);  
                    } while (!string.IsNullOrEmpty(lineText));
                }

                string outRicsFileDir = Path.Combine(configObj.OutputDir, configObj.OutRicsFoldr);
                string[] ricsFileFullPath = Directory.GetFiles(outRicsFileDir, "*.txt", SearchOption.AllDirectories);
                foreach (string ric in ricList)
                {
                    MainRicRecord mainRicRecord =new MainRicRecord(ric);
                    ChainRicRecord chainRicRecord=new ChainRicRecord(mainRicRecord);
                    chainRicRecord.RicList = new List<RicRecord>();
                    for (int i = 0; i < ricsFileFullPath.Length;i++ )
                    {
                        string ricsFileName = Path.GetFileName(ricsFileFullPath[i]);
                        if (ricsFileName.Contains(ric))
                        {
                            StreamReader sr = new StreamReader(ricsFileFullPath[i]);
                            string line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                string[] lineArray = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                if (lineArray.Length == 3)
                                {
                                    chainRicRecord.RicList.Add(new RicRecord(lineArray[2],"0"));
                                }
                            }
                            sr.Close();
                        }
                    }
                    chainRicList.Add(chainRicRecord);
                }

                GetRicFields(chainRicList);
                GenerateComparedAndValidationFile(chainRicList);
            }
            catch (Exception ex)
            {
                Logger.LogErrorAndRaiseException(string.Format("Error happens when generating output files. Ex: {0} .", ex.Message));
            }
        }

        /// <summary>
        /// Read FM file
        /// </summary>
        /// <param name="fmFilePath">The full path of the FM file.</param>
        /// <returns>The main Ric list.</returns>
        private List<string> GetMainRic(string fmFilePath)
        {
            List<string> ricList = new List<string>();

            using (ExcelApp app = new ExcelApp(false, false))
            {
                if (!File.Exists(fmFilePath))
                {
                    Logger.Log(string.Format("Can't find the FM file in the path {0} .", fmFilePath));
                    return ricList;
                }
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, fmFilePath);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                string ric=null;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 5, 3, ExcelLineWriter.Direction.Down))
                {
                    ric=reader.ReadLineCellText();
                    while(!string.IsNullOrEmpty(ric))
                    {
                        ricList.Add(ric);
                        ric = reader.ReadLineCellText();
                    }
                }
                workbook.Close(false, workbook.FullName, false);
            }

            return ricList;
        }

        /// <summary>
        /// Call GATs to get RIC list of Chain RIC
        /// </summary>
        /// <param name="chainRic">The ChainRic.</param>
        /// <param name="outRicsFilePath">The output file which contains all the rics of the chain ric.</param>
        private void GetRics(string chainRic, out string outRicsFilePath)
        {
            string outRicsFileName=string.Format("{0}_outrics.txt",chainRic);
            string outRicsFileDir = Path.Combine(configObj.OutputDir, configObj.OutRicsFoldr);
            if (!Directory.Exists(outRicsFileDir))
            {
                Directory.CreateDirectory(outRicsFileDir);
            }
            outRicsFilePath=Path.Combine(outRicsFileDir, outRicsFileName);
            string argument = string.Format("/c-quiet -dbout -raw_enum_vals  -ph {3} -pn IDN_SELECTFEED -rics \"{0}\" -lfid \"{1}\" -tee \"{2}\"", chainRic, configObj.RicFidsPath, outRicsFilePath, gatsServer);
            try
            {
                using (ProcessContext processContext = new ProcessContext(@"D:\\Data2XML\\Data2XML.exe", argument, "D:\\Data2XML"))
                {
                    processContext.ProcessInstance.StartInfo.RedirectStandardError = true;

                    processContext.ProcessInstance.Start();
                    //avoid deadlock
                    Logger.Log(processContext.ProcessInstance.StandardError.ReadToEnd());
                    processContext.ProcessInstance.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error happened when generating txt files. Ex: " + ex.Message);
            }
        }

        /// <summary>
        /// Get FID value of RIC
        /// </summary>
        /// <param name="chainRicList">The list which contains all the rics and their fids.</param>
        private void GetRicFields(List<ChainRicRecord> chainRicList)
        {
            string path1 = Path.Combine(configObj.OutputDir, configObj.RicFieldsFoldr);
            foreach (ChainRicRecord chainRicRecord in chainRicList)
            {
                string path2 = chainRicRecord.MainRic.Ric;
                chainRicRecord.MainRic.FieldsList = new List<string>();
                string chainRicFolderPath = Path.Combine(path1, path2);
                if (!Directory.Exists(chainRicFolderPath))
                {
                    Directory.CreateDirectory(chainRicFolderPath);
                }
                if (chainRicRecord.RicList.Count > 0)
                {
                    foreach (RicRecord ricRecord in chainRicRecord.RicList)
                    {
                        string ricFieldsFileName = string.Format("{0}.txt", ricRecord.Ric);
                        string ricFieldsFilePath = Path.Combine(chainRicFolderPath, ricFieldsFileName);
                        string argument = string.Format("/c-quiet -dbout -raw_enum_vals  -ph {3} -pn IDN_SELECTFEED -rics \"{0}\" -lfid \"{1}\" -tee \"{2}\"", ricRecord.Ric, configObj.FidListPath, ricFieldsFilePath, gatsServer);
                        try
                        {
                            using (ProcessContext processContext = new ProcessContext(@"D:\\Data2XML\\Data2XML.exe", argument, "D:\\Data2XML"))
                            {
                                processContext.ProcessInstance.StartInfo.RedirectStandardError = true;

                                processContext.ProcessInstance.Start();

                                //avoid deadlock
                                Logger.Log(processContext.ProcessInstance.StandardError.ReadToEnd());
                                processContext.ProcessInstance.WaitForExit();
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("Error happened when generating txt files. Ex: " + ex.Message);
                        }
                        StreamReader sr = new StreamReader(ricFieldsFilePath);
                        string line = sr.ReadLine();
                        if (line != null)
                        {
                            ricRecord.HaveFields = "1";
                        }
                        sr.Close();
                    }
                    if ("1".Equals(chainRicRecord.RicList[0].HaveFields))
                    {
                        chainRicRecord.MainRic.HaveFields = "1";
                    }
                }
                else
                {
                    string mainRic = chainRicRecord.MainRic.Ric;
                    string ricFieldsFileName = string.Format("{0}.txt", mainRic);
                    string ricFieldsFilePath = Path.Combine(chainRicFolderPath, ricFieldsFileName);
                    string argument = string.Format("/c-quiet -dbout -raw_enum_vals  -ph {3} -pn IDN_SELECTFEED -rics \"{0}\" -lfid \"{1}\" -tee \"{2}\"", mainRic, configObj.FidListPath, ricFieldsFilePath, gatsServer);
                    try
                    {
                        using (ProcessContext processContext = new ProcessContext(@"D:\\Data2XML\\Data2XML.exe", argument, "D:\\Data2XML"))
                        {
                            processContext.ProcessInstance.StartInfo.RedirectStandardError = true;

                            processContext.ProcessInstance.Start();

                            //avoid deadlock
                            Logger.Log(processContext.ProcessInstance.StandardError.ReadToEnd());
                            processContext.ProcessInstance.WaitForExit();
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log("Error happened when generating txt files. Ex: " + ex.Message);
                    }

                    StreamReader sr = new StreamReader(ricFieldsFilePath);
                    string line = sr.ReadLine();
                    if (line != null)
                    {
                        chainRicRecord.MainRic.HaveFields = "1";
                    }
                }

                if ("1".Equals(chainRicRecord.MainRic.HaveFields))
                {
                    string mainRic = chainRicRecord.MainRic.Ric;
                    string ricFieldsFileName = string.Format("{0}.txt", mainRic);
                    string ricFieldsFilePath = Path.Combine(chainRicFolderPath, ricFieldsFileName);
                    StreamReader sr = new StreamReader(ricFieldsFilePath);
                    string line;
                    int i = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        i++;
                        string[] lineArray = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (lineArray.Length >= 3)
                        {
                            StringBuilder sb = new StringBuilder();
                            for (int k = 2; k < lineArray.Length; k++)
                            {
                                sb.Append(" ");
                                sb.Append(lineArray[k]);
                            }
                            string fieldValue = sb.ToString().Substring(1);
                            SaveValue(i, fieldValue);
                        }
                        else if (lineArray.Length == 2)
                        {
                            SaveValue(i, "");
                        }
                    }
                    sr.Close();
                    for (int j = 0; j < fields.Length; j++)
                    {
                        chainRicRecord.MainRic.FieldsList.Add(fields[j]);
                        fields[j] = "";
                    }
                }
            }
        }

        private void SaveValue(int i,string value)
        {
            switch(i)
            {
                case 2:
                    fields[0]=value;
                    break;
                case 3:
                    fields[5]=value;
                    break;
                case 4:
                    fields[4] = value;
                    break;
                case 5:
                    fields[2] = value;
                    break;
                case 6:
                    fields[3] = value;
                    break;
                case 7:
                    fields[8] = value;
                    break;
                case 8:
                    fields[6] = value;
                    break;
                case 9:
                    fields[7] = value;
                    break;
                case 10:
                    fields[1] = value;
                    break;
                default:
                    break;
            }
        }

        private void GenerateComparedAndValidationFile(List<ChainRicRecord> chainRicList)
        {
            Logger.Log("Start to generate the compared file and the validation file.");
            using (ExcelApp xlapp = new ExcelApp(false, false))
            {
                MiscUtil.BackUpFile(configObj.FMFilePath);
                var workbook = ExcelUtil.CreateOrOpenExcelFile(xlapp, configObj.FMFilePath);
                var worksheet = ExcelUtil.GetWorksheet("Sheet1", workbook);
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    int[] addedColumnOrder = new int[] { 6, 8, 10, 12, 15, 17, 22, 26, 28 };
                    string[] addedColumnTitle = new string[] { "DSPLY_NAME", "OFFC_CODE2", "OFFCL_CODE", "BCAST_REF", "MATUR_DATE", "STRIKE_PRC", "WNT_RATIO", "GV2_DATE", "LONGLINK1" };
                    for (int i = 0; i < addedColumnOrder.Length; i++)
                    {
                        ExcelUtil.InsertBlankCols(ExcelUtil.GetRange(1, addedColumnOrder[i], worksheet), 1);
                        ExcelUtil.GetRange(4, addedColumnOrder[i], worksheet).Value2 = addedColumnTitle[i];
                        ExcelUtil.GetRange(4, addedColumnOrder[i], worksheet).Interior.Color = System.Drawing.Color.FromArgb(146, 208, 80).ToArgb(); 
                    }

                    writer.PlaceNext(5, 6);
                    foreach (ChainRicRecord chainRicRecord in chainRicList)
                    {
                        if ("1".Equals(chainRicRecord.MainRic.HaveFields))
                        {
                            for (int i = 0; i < addedColumnOrder.Length; i++)
                            {
                                ExcelUtil.GetRange(writer.Row, addedColumnOrder[i], worksheet).Value2 = chainRicRecord.MainRic.FieldsList[i];
                                string valueInFM = ExcelUtil.GetRange(writer.Row, addedColumnOrder[i] - 1, worksheet).Text.ToString();
                                string valueInGATS = ExcelUtil.GetRange(writer.Row, addedColumnOrder[i], worksheet).Text.ToString();
                                if (i == 4 || i==7)
                                {
                                    string[] dateArray = valueInFM.Split(new char[] { '-' });
                                    if (dateArray.Length == 3)
                                    {
                                        string formatedDate = string.Format("{0} {1} 20{2}", dateArray[0], dateArray[1].ToUpper(), dateArray[2]);
                                        if (!formatedDate.Equals(chainRicRecord.MainRic.FieldsList[i]))
                                        {
                                            ExcelUtil.GetRange(writer.Row, addedColumnOrder[i] - 1, worksheet).Interior.Color = System.Drawing.Color.FromArgb(149, 179, 215).ToArgb();
                                        }
                                    }
                                    else
                                    {
                                        if (!valueInFM.Equals(chainRicRecord.MainRic.FieldsList[i]))
                                        {
                                            ExcelUtil.GetRange(writer.Row, addedColumnOrder[i] - 1, worksheet).Interior.Color = System.Drawing.Color.FromArgb(149, 179, 215).ToArgb();
                                        }
                                    }
                                }
                                else
                                {
                                    if (!valueInFM.Equals(valueInGATS))
                                    {
                                        ExcelUtil.GetRange(writer.Row, addedColumnOrder[i] - 1, worksheet).Interior.Color = System.Drawing.Color.FromArgb(149, 179, 215).ToArgb();
                                    }
                                }
                            }
                        }
                        writer.PlaceNext(writer.Row + 1, 6);
                    }
                }
                workbook.Save();
                workbook.Close(false, workbook.FullName, false);

                var workbookValidation = ExcelUtil.CreateOrOpenExcelFile(xlapp, configObj.ValidationResultFilePath);
                var worksheet1 = ExcelUtil.GetWorksheet("Sheet1", workbookValidation);
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet1, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    string[] sheetTitle = new string[] { "MainRic", "Have Data(1 means Yes, 0 means No)", " Have Rics(1 means Yes, 0 means No)" };
                    for (int i = 0; i < sheetTitle.Length; i++)
                    {
                        writer.WriteLine(sheetTitle[i]);
                    }
                    writer.PlaceNext(2, 1);
                    int maxRicCount = 0;
                    foreach (ChainRicRecord chainRicRecord in chainRicList)
                    {
                        writer.WriteLine(chainRicRecord.MainRic.Ric);
                        writer.WriteLine(chainRicRecord.MainRic.HaveFields);
                        if (chainRicRecord.RicList.Count > 0)
                        {
                            writer.WriteLine("1");
                            chainRicRecord.RicList.RemoveAt(0);
                            maxRicCount = chainRicRecord.RicList.Count;
                            foreach (RicRecord ricRecord in chainRicRecord.RicList)
                            {
                                writer.WriteLine(string.Format("{0} : {1}",ricRecord.Ric, ricRecord.HaveFields));
                                if ("0".Equals(ricRecord.HaveFields))
                                {
                                    ExcelUtil.GetRange(writer.Row, writer.Col - 1, worksheet1).Interior.Color = System.Drawing.Color.FromArgb(149, 179, 215).ToArgb();
                                }
                            } 
                        }
                        else
                        {
                            writer.WriteLine("0");
                        }
                        writer.PlaceNext(writer.Row + 1, 1);
                    }

                    writer.PlaceNext(1, 4);
                    for (int num = 1; num <= maxRicCount; num++)
                    {
                        writer.WriteLine(string.Format("Ric{0}(1 means having data, 0 means not)", num));
                    }
                }
                workbookValidation.Save();
                workbookValidation.Close(false, workbookValidation.FullName, false);
            }
            Logger.Log("Finished generating the compared file and the validation file.");
        }
    }
}
