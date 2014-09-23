using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.ComponentModel;
using Ric.Util;
using Ric.Core;

namespace Ric.Tasks.HongKong
{
    public class HKFMTemplate
    {
        //public string FMSerialNumber { get; set; }

        //For TQS
        //public string EffectiveDate { get; set; }
        public string UnderLyingRic { get; set; }
        public string CompositeChainRic { get; set; }
        public string BrokerPageRic { get; set; }
        public string MiscInfoPageRic { get; set; }
        public string DisplayName { get; set; }
        public string OfficicalCode { get; set; }
        public string ExchangeSymbol { get; set; }
        public string Currency { get; set; }
        public string RecordType { get; set; }
        public string SpareUbytes8 { get; set; }
        public string UnderlyingChainRic1 { get; set; }
        public string UnderlyingChainRic2 { get; set; }

        public string ChainRic1 { get; set; }
        public string ChainRic2 { get; set; }
        public string ChainRic3 { get; set; }
        //public string WarrantType { get; set; }
        public string MiscInfoPageChainRic { get; set; }
        public string LotSize { get; set; }
        public string OfferPriceRange { get; set; }
        public string ColDsplyNmll { get; set; }
        public string BcastRef { get; set; }
        public string LongLink1_WarrantChainRic { get; set; }
        public string LongLink2_BrokerPageRic { get; set; }
        public string LongLink3_TasRic { get; set; }
        public string WntRation { get; set; }
        public string StrikPrc { get; set; }
        public string MaturDate { get; set; }
        //public string LongLink3 { get; set; }
        public string SpareSnum13 { get; set; }
        public string GNTX20_13 { get; set; }
        public string GNTX20_14 { get; set; }
        public string GNTX20_15 { get; set; }
        public string GNTX20_16 { get; set; }
        public string GNTX20_10 { get; set; }
        public string GNTX20_11 { get; set; }
        public string GNTX20_12 { get; set; }
        public string CouponRate { get; set; }
        //public string IssuePrice { get; set; }
        //public string Row80_13 { get; set; }
        public string GVFlag { get; set; }
        public string IssTpFlg { get; set; }
        public string RdmCur { get; set; }
        public string LongLink14 { get; set; }
        public string BondType { get; set; }
        public string Leg1Str { get; set; }
        public string Leg2Str { get; set; }
        public string GNTXT24_1 { get; set; }
        public string GNTXT24_2 { get; set; }
        public string ChainRic { get; set; }

        //For NDA
        public string NewOrgList { get; set; }
        public string PrimaryList { get; set; }
        public string IDNLongName { get; set; }
        public string LegalRegistName { get; set; }
        public string GeographyEntity { get; set; }
        public string OrgnizationType { get; set; }
        public string AliasPre { get; set; }
        public string AliasGen { get; set; }
        public string IssueClassification { get; set; }
        public string MSCICode { get; set; }
        public string BusinessActivity { get; set; }
        public string ExistingOrgList { get; set; }
        public string OrgnizationName1 { get; set; }
        public string OrgnizationName2 { get; set; }

        //For RA_ADG
        public string Ric { get; set; }
        public string LocalSectorClassification { get; set; }
        public string IndexRics { get; set; }
    }

    public class MainFileTemplate
    {
        public string Main_Symbol { get; set; }
        public string Main_Display_Name { get; set; }
        public string Main_Ric { get; set; }
        public string Main_Official_Code { get; set; }
        public string Main_Ex_Symbol { get; set; }
        public string Main_Background_Page { get; set; }
        public string Main_Lot_Size { get; set; }
        public string Main_Display_Nmll { get; set; }
        public string Main_Bcast_Ref { get; set; }
        public string Main_Instmod_Bond_Type { get; set; }
        public string Main_Instmod_Lot_Size { get; set; }
        public string Main_Instmod_Mnemonic { get; set; }
        public string Main_Instmod_Tdn_Symbol { get; set; }
        public string Main_Exl_Name { get; set; }
        public string Main_Instmod_Longlink2 { get; set; }
        public string Main_Instmod_Longlink3 { get; set; }
        public string Main_Instmod_Spare_Snum13 { get; set; }
        public string Main_Instmod_GNTX20_10 { get; set; }
        public string Main_Instmod_GNTX20_11 { get; set; }
        public string Main_Instmod_GNTX20_13 { get; set; }
        public string Main_Instmod_GNTX20_14 { get; set; }
        public string Main_Instmod_GNTX20_16 { get; set; }
        public string Main_Instmod_Dds_Lot_Size { get; set; }
        public string Main_Instmod_Tdn_Issuer_Name { get; set; }
        public string Main_Instmod_ISIN { get; set; }

        public MainFileTemplate()
        {
            this.Main_Symbol = "SYMBOL";
            this.Main_Display_Name = "DSPLY_NAME";
            this.Main_Ric = "RIC";
            this.Main_Official_Code = "OFFCL_CODE";
            this.Main_Ex_Symbol = "EX_SYMBOL";
            this.Main_Background_Page = "BCKGRNDPAG";
            this.Main_Lot_Size = "LOT_SIZE_A";
            this.Main_Display_Nmll = "DSPLY_NMLL";
            this.Main_Bcast_Ref = "BCAST_REF";
            this.Main_Instmod_Bond_Type = "#INSTMOD_BOND_TYPE";
            this.Main_Instmod_Lot_Size = "#INSTMOD_LOT_SIZE_X";
            this.Main_Instmod_Mnemonic = "#INSTMOD_MNEMONIC";
            this.Main_Instmod_Tdn_Symbol = "#INSTMOD_TDN_SYMBOL";
            this.Main_Exl_Name = "EXL_NAME";
            this.Main_Instmod_Longlink2 = "#INSTMOD_LONGLINK2";
            this.Main_Instmod_Longlink3 = "#INSTMOD_LONGLINK3";
            this.Main_Instmod_Spare_Snum13 = "#INSTMOD_SPARE_SNUM13";
            this.Main_Instmod_GNTX20_10 = "#INSTMOD_GN_TX20_10";
            this.Main_Instmod_GNTX20_11 = "#INSTMOD_GN_TX20_11";
            this.Main_Instmod_GNTX20_13 = "#INSTMOD_GN_TX20_13";
            this.Main_Instmod_GNTX20_14 = "#INSTMOD_GN_TX20_14";
            this.Main_Instmod_GNTX20_16 = "#INSTMOD_GN_TX20_16";
            this.Main_Instmod_Dds_Lot_Size = "#INSTMOD_#DDS_LOT_SIZE";
            this.Main_Instmod_Tdn_Issuer_Name = "#INSTMOD_TDN_ISSUER_NAME";
            this.Main_Instmod_ISIN = "#INSTMOD_#ISIN";
        }
    }

    public class MIFileTemplate
    {
        public string MI_Symbol { get; set; }
        public string MI_Display_Name { get; set; }
        public string MI_Ric { get; set; }
        public string MI_Ex_Symbol { get; set; }
        public string MI_Instmod_Row80_1 { get; set; }
        public string MI_Exl_Name { get; set; }

        public MIFileTemplate()
        {
            this.MI_Symbol = "SYMBOL";
            this.MI_Display_Name = "DSPLY_NAME";
            this.MI_Ric = "RIC";
            this.MI_Ex_Symbol = "EX_SYMBOL";
            this.MI_Instmod_Row80_1 = "#INSTMOD_ROW80_1";
            this.MI_Exl_Name = "EXL_NAME";
        }
    }

    public class FMFileFormate
    {
        public string Name { get; set; }
        public string FMSerialNumber { get; set; }
        public string Issue { get; set; }
        public string Bond { get; set; }
        public List<String> StockCode { get; set; }
        public string Date { get; set; }
    }

    public class ReadFMFiletoGenerateBulkFileConfig
    {
        [Description("descriptiondfmkdmgl,h;glj")]
        public string FM_SOURCE_FILE_PATH { get; set; }

        [Description("fdhskjfdjklgllh;g")]
        public string FM_SOURCE_FILE_SHEET_NAME { get; set; }

        public string MAIN_RESULT_FILE_PATH { get; set; }
        public string MI_RESULT_FILE_PATH { get; set; }
        public string LOG_FILE_PATH { get; set; }
    }

    public class ReadFMFiletoGenerateBulkFile : GeneratorBase
    {
        private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_ReadFMFiletoGenerateBulkFile.config";
        //private static Logger logger = null;
        private static ReadFMFiletoGenerateBulkFileConfig configObj = null;
        List<FMFileFormate> fileList = new List<FMFileFormate>();


        protected override void Initialize()
        {
            base.Initialize();
            try
            {
                configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(ReadFMFiletoGenerateBulkFileConfig)) as ReadFMFiletoGenerateBulkFileConfig;
            }
            catch (System.Exception ex)
            {
                Logger.Log("Error happens when initializing task... Ex: " + ex.Message);
            }
            if (configObj.FM_SOURCE_FILE_PATH == null)
            {
                configObj.FM_SOURCE_FILE_PATH = @"D:\";
            }
            if (configObj.FM_SOURCE_FILE_SHEET_NAME == null)
            {
                configObj.FM_SOURCE_FILE_SHEET_NAME = "Sheet1";
            }
            if (string.IsNullOrEmpty(configObj.MAIN_RESULT_FILE_PATH))
            {
                configObj.MAIN_RESULT_FILE_PATH = @"D:\HKRicTemplate\";
            }
            if (string.IsNullOrEmpty(configObj.MI_RESULT_FILE_PATH))
            {
                configObj.MI_RESULT_FILE_PATH = @"D:\HKRicTemplate\";
            }
            //logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.New);
        }

        protected override void Start()
        {
            StartReadFMFiletoGenerateBulkFileConfigJob();
        }

        public void StartReadFMFiletoGenerateBulkFileConfigJob()
        {
            if (Directory.Exists(configObj.FM_SOURCE_FILE_PATH))
            {
                getFiles(configObj.FM_SOURCE_FILE_PATH);

                #region Generate Main and MI files for per FM file

                foreach (FMFileFormate fileContent in fileList)
                {
                    //int countsInOneFMFile = 0;
                    string filePath = Path.Combine(configObj.FM_SOURCE_FILE_PATH, string.Format("{0}.xls", fileContent.Name));
                    List<HKFMTemplate> FMContentList = new List<HKFMTemplate>();

                    #region Get every FM file contents
                    using (ExcelApp app = new ExcelApp(false, false))
                    {
                        var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                        var worksheet = ExcelUtil.GetWorksheet(configObj.FM_SOURCE_FILE_SHEET_NAME, workbook);
                        if (worksheet == null)
                        {
                            String msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                            Logger.Log(msg, Logger.LogType.Error);
                        }
                        else
                        {
                            int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;

                            for (int i = 1; i <= lastUsedRow; )
                            {
                                if (ExcelUtil.GetRange(i, 1, worksheet).Text.ToString().Replace(" ", "").ToUpper() == "**FORTQS**")
                                {
                                    i += 2;
                                    HKFMTemplate FMContent = new HKFMTemplate();
                                    FMContent.UnderLyingRic = ExcelUtil.GetRange(i, 2, worksheet).Text.ToString();
                                    FMContent.DisplayName = ExcelUtil.GetRange(i + 4, 2, worksheet).Text.ToString().Split('<')[0].Trim().ToUpper();
                                    FMContent.OfficicalCode = ExcelUtil.GetRange(i + 5, 2, worksheet).Text.ToString();
                                    FMContent.LotSize = ExcelUtil.GetRange(i + 13, 2, worksheet).Text.ToString();
                                    FMContent.ColDsplyNmll = ExcelUtil.GetRange(i + 16, 2, worksheet).Text.ToString().Split('<')[0].Trim();
                                    FMContent.BcastRef = ExcelUtil.GetRange(i + 18, 2, worksheet).Text.ToString();
                                    FMContent.LongLink2_BrokerPageRic = ExcelUtil.GetRange(i + 21, 2, worksheet).Text.ToString().ToUpper();
                                    FMContent.LongLink3_TasRic = ExcelUtil.GetRange(i + 22, 2, worksheet).Text.ToString();
                                    FMContent.SpareSnum13 = ExcelUtil.GetRange(i + 23, 2, worksheet).Text.ToString();
                                    FMContent.GNTX20_10 = ExcelUtil.GetRange(i + 25, 2, worksheet).Text.ToString().Trim().PadRight(16) + ExcelUtil.GetRange(i + 26, 2, worksheet).Text.ToString().Trim();
                                    FMContent.GNTX20_11 = ExcelUtil.GetRange(i + 27, 2, worksheet).Text.ToString().Replace("~", " ").ToUpper();
                                    FMContent.GNTX20_13 = ExcelUtil.GetRange(i + 32, 2, worksheet).Text.ToString() + ExcelUtil.GetRange(i + 33, 2, worksheet).Text.ToString();
                                    FMContent.GNTX20_14 = ExcelUtil.GetRange(i + 35, 2, worksheet).Text.ToString() + ExcelUtil.GetRange(i + 36, 2, worksheet).Text.ToString();
                                    FMContent.GNTX20_16 = ExcelUtil.GetRange(i + 38, 2, worksheet).Text.ToString();
                                    FMContent.BondType = ExcelUtil.GetRange(i + 48, 2, worksheet).Text.ToString().ToUpper();
                                    FMContent.OrgnizationName1 = ExcelUtil.GetRange(i + 61, 2, worksheet).Text.ToString().ToUpper();
                                    //countsInOneFMFile++;
                                    FMContentList.Add(FMContent);
                                }
                                else
                                {
                                    i++;
                                    continue;
                                }
                            }
                        }
                    }
                    #endregion

                    List<string> EXL_NAME = new List<string>();

                    generateMainFile(configObj.MAIN_RESULT_FILE_PATH, FMContentList, EXL_NAME);
                    generateMIFile(configObj.MI_RESULT_FILE_PATH, FMContentList, EXL_NAME);
                }
                #endregion
            }
        }


        //Need to be improved, adding some validations, current logical is user maintain the FM files
        private void getFiles(String filesDir)
        {
            try
            {
                string[] files = System.IO.Directory.GetFiles(filesDir);
                foreach (string file in files)
                {
                    FMFileFormate formate = new FMFileFormate();
                    string[] count = file.Split('\\');
                    formate.Name = count[count.Length - 1].Split('.')[0];
                    string[] content = formate.Name.Split('_');
                    formate.FMSerialNumber = content[0].Split('-')[0];
                    formate.Issue = content[0].Split('-')[1];
                    formate.Bond = content[1];
                    if (content.Length > 4)
                    {
                        for (int i = 2; i < content.Length - 1; i++)
                            formate.StockCode.Add(content[i]);
                    }
                    formate.Date = content[content.Length - 1];
                    fileList.Add(formate);
                }
            }
            catch (System.Exception ex)
            {
                Logger.Log("Error happens when looking for files... Ex: " + ex.Message);
            }
        }

        private String genOutputFileName(String path, List<HKFMTemplate> FMContentList, String type)
        {
            string fileName = path;
            if (path.Last() == '\\')
            {
                if (type == "main")
                {
                    for (int i = 0; i < FMContentList.Count; i++)
                    {
                        fileName += FMContentList[i].OfficicalCode.ToString() + "_";
                    }
                    fileName += "HK_MAIN.txt";
                }
                if (type == "mi")
                {
                    for (int i = 0; i < FMContentList.Count; i++)
                    {
                        fileName += FMContentList[i].OfficicalCode.ToString() + "_";
                    }
                    fileName += "HK_MI.txt";
                }
            }
            else
            {
                fileName += path + "\\";
                if (type == "main")
                {
                    for (int i = 0; i < FMContentList.Count; i++)
                    {
                        fileName += FMContentList[i].OfficicalCode.ToString() + "_";
                    }
                    fileName += "HK_MAIN.txt";
                }
                if (type == "mi")
                {
                    for (int i = 0; i < FMContentList.Count; i++)
                    {
                        fileName += FMContentList[i].OfficicalCode.ToString() + "_";
                    }
                    fileName += "HK_MI.txt";
                }
            }
            return fileName;
        }

        private void generateMainFile(String mainFilePath, List<HKFMTemplate> FMContentList, List<string> EXL_NAME)
        {
            string filePath = genOutputFileName(mainFilePath, FMContentList, "main");
            StringBuilder sb = new StringBuilder();
            MainFileTemplate main = new MainFileTemplate();

            //sb.Append(main.Main_Symbol);
            sb.AppendFormat("{0}\t", main.Main_Symbol);
            sb.AppendFormat("{0}\t", main.Main_Display_Name);
            sb.AppendFormat("{0}\t", main.Main_Ric);
            sb.AppendFormat("{0}\t", main.Main_Official_Code);
            sb.AppendFormat("{0}\t", main.Main_Ex_Symbol);
            sb.AppendFormat("{0}\t", main.Main_Background_Page);
            sb.AppendFormat("{0}\t", main.Main_Lot_Size);
            sb.AppendFormat("{0}\t", main.Main_Display_Nmll);
            sb.AppendFormat("{0}\t", main.Main_Bcast_Ref);
            sb.AppendFormat("{0}\t", main.Main_Instmod_Bond_Type);
            sb.AppendFormat("{0}\t", main.Main_Instmod_Lot_Size.Replace(",", ""));
            sb.AppendFormat("{0}\t", main.Main_Instmod_Mnemonic);
            sb.AppendFormat("{0}\t", main.Main_Instmod_Tdn_Symbol);
            sb.AppendFormat("{0}\t", main.Main_Exl_Name);
            sb.AppendFormat("{0}\t", main.Main_Instmod_Longlink2);
            sb.AppendFormat("{0}\t", main.Main_Instmod_Longlink3);
            sb.AppendFormat("{0}\t", main.Main_Instmod_Spare_Snum13);
            sb.AppendFormat("{0}\t", main.Main_Instmod_GNTX20_10);
            sb.AppendFormat("{0}\t", main.Main_Instmod_GNTX20_11);
            sb.AppendFormat("{0}\t", main.Main_Instmod_GNTX20_13);
            sb.AppendFormat("{0}\t", main.Main_Instmod_GNTX20_14);
            sb.AppendFormat("{0}\t", main.Main_Instmod_GNTX20_16);
            sb.AppendFormat("{0}\t", main.Main_Instmod_Dds_Lot_Size);
            sb.AppendLine(main.Main_Instmod_Tdn_Issuer_Name);
            int i = 0;
            foreach (HKFMTemplate content in FMContentList)
            {
                sb.AppendFormat("{0}\t", content.UnderLyingRic);
                sb.AppendFormat("{0}\t", content.DisplayName);
                sb.AppendFormat("{0}\t", content.UnderLyingRic);
                sb.AppendFormat("{0}\t", content.OfficicalCode.Substring(1, content.OfficicalCode.Length - 1));
                sb.AppendFormat("{0}\t", content.UnderLyingRic.Split('.')[0]);
                sb.AppendFormat("{0}\t", "****");
                sb.AppendFormat("{0}\t", content.LotSize);
                sb.AppendFormat("{0}\t", content.ColDsplyNmll);
                sb.AppendFormat("{0}\t", content.BcastRef);
                sb.AppendFormat("{0}\t", content.BondType);
                sb.AppendFormat("{0}\t", content.LotSize);
                sb.AppendFormat("{0}\t", content.OfficicalCode.Substring(1, content.OfficicalCode.Length - 1));
                sb.AppendFormat("{0}\t", content.OfficicalCode.Substring(1, content.OfficicalCode.Length - 1));
                if (Int32.Parse(content.OfficicalCode) < 250)
                {
                    EXL_NAME.Add("HKG_EQB_1");
                    sb.AppendFormat("{0}\t", "HKG_EQB_1");
                }
                if ((Int32.Parse(content.OfficicalCode) >= 250) && (Int32.Parse(content.OfficicalCode) <= 1000))
                {
                    EXL_NAME.Add("HKG_EQB_2");
                    sb.AppendFormat("{0}\t", "HKG_EQB_2");
                }
                if (Int32.Parse(content.OfficicalCode) > 1000)
                {
                    EXL_NAME.Add("HKG_EQB_3");
                    sb.AppendFormat("{0}\t", "HKG_EQB_3");
                }
                sb.AppendFormat("{0}\t", content.LongLink2_BrokerPageRic);
                sb.AppendFormat("{0}\t", content.LongLink3_TasRic);
                sb.AppendFormat("{0}\t", content.SpareSnum13);
                sb.AppendFormat("{0}\t", content.GNTX20_10);
                sb.AppendFormat("{0}\t", content.GNTX20_11);
                sb.AppendFormat("{0}\t", content.GNTX20_13);
                sb.AppendFormat("{0}\t", content.GNTX20_14);
                sb.AppendFormat("{0}\t", content.GNTX20_16);
                sb.AppendFormat("{0}\t", content.LotSize);
                sb.AppendLine(content.OrgnizationName1);
                i++;
            }
            File.WriteAllText(filePath, sb.ToString(), Encoding.GetEncoding("gb2312"));
        }

        private void generateMIFile(String miFilePath, List<HKFMTemplate> FMContentList, List<string> EXL_NAME)
        {
            string filePath = genOutputFileName(miFilePath, FMContentList, "mi");
            StringBuilder sb = new StringBuilder();
            MIFileTemplate mi = new MIFileTemplate();
            sb.AppendFormat("{0}\t", mi.MI_Symbol);
            sb.AppendFormat("{0}\t", mi.MI_Display_Name);
            sb.AppendFormat("{0}\t", mi.MI_Ric);
            sb.AppendFormat("{0}\t", mi.MI_Ex_Symbol);
            sb.AppendFormat("{0}\t", mi.MI_Instmod_Row80_1);
            sb.AppendLine(mi.MI_Exl_Name);
            int j = 0;
            foreach (HKFMTemplate content in FMContentList)
            {
                sb.AppendFormat("{0}\t", content.UnderLyingRic.Split('.')[0] + "MI.HK");
                sb.AppendFormat("{0}\t", content.DisplayName);
                sb.AppendFormat("{0}\t", content.UnderLyingRic.Split('.')[0] + "MI.HK");
                sb.AppendFormat("{0}\t", content.UnderLyingRic.Split('.')[0] + "MI");
                sb.AppendFormat("{0}\t", "Security Miscellaneous Information                                     " + content.UnderLyingRic.Split('.')[0] + "MI.HK");
                sb.AppendLine(EXL_NAME[j] + "_MI_PAGE");
                j++;
            }
            File.WriteAllText(filePath, sb.ToString(), Encoding.GetEncoding("gb2312"));
        }
    }
}
