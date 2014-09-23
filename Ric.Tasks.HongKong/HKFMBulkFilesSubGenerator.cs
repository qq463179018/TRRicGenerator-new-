using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.IO;
using System.Data;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    public class HKFMBulkFilesSubGenerator
    {
        private readonly string issueCodeMapPath = ".\\Config\\HK\\HK_IssuerCode.xml";
        private readonly string holidayListFilePath = ".\\Config\\HK\\Holiday.xml";
        private static HK_IssuerCodeMap issuerCodeObj = null;
        private string fmSerialNumberWarrant = string.Empty;
        private List<DateTime> holidayList = null;
        private Logger logger = null;
        private string outputPath = "";
        private List<TaskResultEntry> taskResultList = new List<TaskResultEntry>();
        private HKFMAndBulkFileGenerator parent = null;
        public List<RicInfo> RicListCbbc = new List<RicInfo>();
        public List<RicInfo> ChineseListCbbc = new List<RicInfo>();
        public List<RicInfo> RicListWarrant = new List<RicInfo>();
        public List<RicInfo> ChineseListWarrant = new List<RicInfo>();
        private HK_IssuerCodeMap issuerCodeMap = new HK_IssuerCodeMap();


        public HKFMBulkFilesSubGenerator(HKFMAndBulkFileGenerator parent)
        {
            this.parent = parent;
        }

        public void Initialize(string outputPath, Logger logger, List<TaskResultEntry> taskResultList)
        {
            //this.logger = logger;
            this.taskResultList = taskResultList;
            this.outputPath = outputPath;

            issuerCodeObj = ConfigUtil.ReadConfig(issueCodeMapPath, typeof(HK_IssuerCodeMap)) as HK_IssuerCodeMap;
            if (File.Exists(holidayListFilePath))
            {
                holidayList = ConfigUtil.ReadConfig(holidayListFilePath, typeof(List<DateTime>)) as List<DateTime>;
            }
            else
            {
                holidayList = new List<DateTime>();
                holidayList.Add(DateTime.Parse("2012-01-01"));
                ConfigUtil.WriteConfig(holidayListFilePath, holidayList);
            }
        }

        public void Cleanup()
        {
        }

        public string GetColDsplyNmllInfo(RicInfo ricEnglishInfo)
        {
            string chineseName = ricEnglishInfo.ChineseName;
            string colDsplyNmll = chineseName.Substring(0, 4);
            char lastCharacter = chineseName[chineseName.Length - 1];
            if (lastCharacter >= 'A' && lastCharacter <= 'Z')
            {
                colDsplyNmll += chineseName[chineseName.Length - 1];
            }

            string[] arr = ricEnglishInfo.MaturityDate.Split('-');
            colDsplyNmll += int.Parse(arr[1]).ToString();
            colDsplyNmll += "月";
            if (ricEnglishInfo.BullBear == "Bull")
            {
                colDsplyNmll += "RC";
            }
            if (ricEnglishInfo.BullBear == "Call")
            {
                colDsplyNmll += "CW";
            }
            if (ricEnglishInfo.BullBear == "Put")
            {
                colDsplyNmll += "PW";
            }
            else
            {
                colDsplyNmll += "RP";
            }
            colDsplyNmll += arr[2].Substring(2, 2);
            return colDsplyNmll;
        }


        public string GetIDNLongName(RicInfo ric)
        {
            bool isIndex = false;
            bool isBull = false;
            bool isHKD = false;
            if (Char.IsLetter(ric.Underlying, 0))
            {
                isIndex = true;
            }
            if (ric.BullBear.ToLower().Contains("bull"))
            {
                isBull = true;
            }
            if (Char.IsLetter(ric.StrikeLevel, 0))
            {
                isHKD = true;
            }
            string idnLongName = string.Empty;
            //For Index
            if (isIndex)
            {
                if (ric.Underlying == "HSI")
                {
                    idnLongName = "HANG SENG@";
                }
                if (ric.Underlying == "HSCEI")
                {
                    idnLongName = "HANG SENG C E I@";
                }

                if (ric.Underlying == "DJI")
                {
                    idnLongName = "DJ INDU AVERAGE@";
                }
            }

            //For stock
            else
            {
                idnLongName += ric.UnderlyingNameForStock + "@";
            }

            idnLongName += GetIssuerName(ric.Issuer, issuerCodeObj)[0];
            idnLongName += " ";
            DateTime maturittDateTime = DateTime.ParseExact(ric.MaturityDate, "dd-MM-yyyy", null);
            idnLongName += maturittDateTime.ToString("MMMyy", new CultureInfo("en-US")).ToUpper() + " ";

            //For HKD
            if (isHKD)
            {
                idnLongName += ric.StrikeLevel.Trim().Substring(4);
            }
            else
            {
                idnLongName += ric.StrikeLevel.Trim();
            }
            idnLongName += " ";

            if (isBull)
            {
                idnLongName += "C";
            }
            else
            {
                idnLongName += "P";
            }
            if (isIndex)
            {
                idnLongName += "IR";
            }
            else
            {
                idnLongName += "R";
            }
            return idnLongName;
        }

        public string GetIDNLongNameForWarrant(RicInfo ricObj, bool isIndex, bool isStock, bool isCall, bool isHKD, HK_IssuerCodeMap issuerCodeMap)
        {
            string idnLongName = "";
            //For Index
            if (isIndex)
            {
                if (ricObj.Underlying == "HSI")
                {
                    idnLongName = "HANG SENG@";
                }
                if (ricObj.Underlying == "HSCEI")
                {
                    idnLongName = "HANG SENG C E I@";
                }
                if (ricObj.Underlying == "DJI")
                {
                    idnLongName = "DJ INDU AVERAGE@";
                }

                idnLongName += GetIssuerName(ricObj.Issuer, issuerCodeMap)[0];
                idnLongName += " ";
                DateTime maturittDateTime = DateTime.ParseExact(ricObj.MaturityDate, "dd-MM-yyyy", null);
                idnLongName += maturittDateTime.ToString("MMMyy", new CultureInfo("en-US")).ToUpper() + " ";

                //Attach Strike Price from Strike Level
                //For HKD
                if (isHKD)
                {
                    idnLongName += ricObj.StrikeLevel.Substring(4);
                }
                else
                {
                    idnLongName += ricObj.StrikeLevel;
                }

                idnLongName += " ";
                //For Call
                if (isCall)
                {
                    idnLongName += "C";
                }
                //For Put
                else
                {
                    idnLongName += "P";
                }

                if (isIndex)
                {
                    idnLongName += "IW";
                }
                else
                {
                    idnLongName += "WT";
                }
            }
            //For Stock
            if (isStock)
            {
                idnLongName = ricObj.UnderlyingNameForStock + "@";
                idnLongName += GetIssuerName(ricObj.Issuer, issuerCodeMap)[0];
                DateTime maturittDateTime = DateTime.ParseExact(ricObj.MaturityDate, "dd-MM-yyyy", null);
                idnLongName += maturittDateTime.ToString("MMMyy", new CultureInfo("en-US")).ToUpper() + " ";
                //Attach Strike Price from Strike Level
                //For HKD
                if (isHKD)
                {
                    idnLongName += ricObj.StrikeLevel.Substring(4);
                }
                else
                {
                    idnLongName += ricObj.StrikeLevel;
                }

                idnLongName += " ";
                //For Call
                if (isCall)
                {
                    idnLongName += "C";
                }
                //For Put
                else
                {
                    idnLongName += "P";
                }

                if (isIndex)
                {
                    idnLongName += "IW";
                }
                else
                {
                    idnLongName += "WT";
                }

            }

            return idnLongName;
        }

        public string[] GetIssuerName(string issuerID, HK_IssuerCodeMap issuerCodeMap)
        {
            string[] arr = { "", "", "" };
            foreach (Trans tran in issuerCodeMap.Map)
            {
                if (issuerID == tran.Code)
                {
                    arr[0] = tran.ShortName;
                    arr[1] = tran.FullName;
                    arr[2] = tran.WarrentIssuer;
                    break;
                }
                else
                {
                    continue;
                }
            }
            if (arr[0] == string.Empty || arr[1] == string.Empty)
            {
                logger.Log("\n There's no issuer for " + issuerID);
            }
            return arr;
        }

        public void StartHKBulkFileGeneratorJob()
        {
            //Generate IAAdd.csv
            GenerateIAAddCSV(outputPath + "\\" + parent.RicTemplatePath, RicListCbbc, RicListWarrant);
            //Generate QAAdd.csv
            GenerateQAAddCSV(outputPath + "\\" + parent.RicTemplatePath, RicListCbbc, RicListWarrant);
            //Generate HKG_EQLB_CBBC.txt
            GenerateHKGEQLBCBBC(outputPath + "\\" + parent.RicTemplatePath, RicListCbbc, ChineseListCbbc);

            //Generate HKG_EQLB.txt
            GenerateHKGEQLB(outputPath + "\\" + parent.RicTemplatePath, RicListWarrant, ChineseListWarrant);//changed for new future underlying

            //Generate HKG_EQLBMI.txt
            GenerateHKGEQLBMI(outputPath + "\\" + parent.RicTemplatePath, RicListCbbc, RicListWarrant);

            //Generate BK.txt
            //GenerateBK(outputPath + "\\" + parent.RicTemplatePath, RicListCbbc, RicListWarrant);      //this file was useless anymore
            //Generate MI.txt
            //GenerateMI(outputPath + "\\" + parent.RicTemplatePath, RicListCbbc, RicListWarrant);      //this file was useless anymore
            //Generate MI_FMSddMMMyyyy.csv 
            //GenerateMIFMS(outputPath + "\\" + parent.RicTemplatePath, RicListCbbc, RicListWarrant);   //this file was useless anymore

            //Generate IS_FMddMMMyyyy.csv
            GenerateISFMS(outputPath + "\\" + parent.RicTemplatePath, RicListCbbc, RicListWarrant);
        }

        /// <summary>
        /// Generate the IS_FMS.csv file for FID 1675- #INSTMOD_GN_TX20_11
        /// </summary>
        /// <param name="filePath">output folder and filename</param>
        /// <param name="cbbcList">CBBC source data</param>
        /// <param name="warrantList">Warrant source data</param>
        public void GenerateISFMS(string filePath, List<RicInfo> cbbcList, List<RicInfo> warrantList)
        {
            List<string> isFmiTitle = new List<string>() { "RIC", "DOMAIN", "GN_TX20_11", "AMT_ISSUE" };
            string fullPath = filePath + "\\IS_FMS" + DateTime.Now.ToString("ddMMMyyyy") + ".csv";

            System.Data.DataTable dt = GenerateCsvTitle(isFmiTitle);

            List<RicInfo> ricList = new List<RicInfo>();
            ricList.AddRange(cbbcList);
            ricList.AddRange(warrantList);
            for (int i = 0; i < ricList.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = ricList[i].Code + ".HK";
                dr[1] = "MARKET_PRICE";
                dr[2] = ricList[i].TotalIssueSize.Replace(",", "");
                dr[3] = ricList[i].TotalIssueSize.Replace(",", "");
                dt.Rows.Add(dr);
            }

            GenerateCSV(fullPath, "IS_FMI", dt);

        }

        /// <summary>
        /// Generate MI_FMS.csv file (for MI RICs of Warrants&CBBC).
        /// </summary>
        /// <param name="filePath">output folder and filename</param>
        /// <param name="cbbcList">CBBC source data</param>
        /// <param name="warrantList">Warrant source data</param>
        public void GenerateMIFMS(string filePath, List<RicInfo> cbbcList, List<RicInfo> warrantList)
        {
            List<string> miFmiTitle = new List<string>() { "RIC", "DOMAIN", "ROW80_3", "ROW80_4", "ROW80_5", "ROW80_6", "ROW80_7", "ROW80_8",
                                                               "ROW80_9", "ROW80_10", "ROW80_11", "ROW80_12", "ROW80_13", "ROW80_14", "ROW80_15", "ROW80_16"};
            List<string> miFmiData = new List<string>() { "MI.HK", "MARKET_PRICE", "ISIN                                    EIPO Start Date", 
                                                          "Instrument Type                         EIPO End Date", "Market                                  EIPO Start Time",
                                                          "Sub-Market                              EIPO End Time", "Listing Date                            EIPO Price", 
                                                          "De-listing Date                         Spread Table", "Listing Status                          Shortselling Stock", 
                                                          "Trading Status                          Intra-day Shortselling Stock", "Stamp Duty                              Automatch Stock", 
                                                          "Test Stock                              CCASS Stock", "Dummy Stock", "--------------------------------------------------------------------------------", 
                                                          "Trading Start Time                      Trading End Time", "Session Type"};
            string fullpath = filePath + "\\MI_FMS" + DateTime.Now.ToString("ddMMMyyyy") + ".csv";
            System.Data.DataTable dt = GenerateCsvTitle(miFmiTitle);

            List<RicInfo> ricList = new List<RicInfo>();
            ricList.AddRange(cbbcList);
            ricList.AddRange(warrantList);
            for (int i = 0; i < ricList.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = ricList[i].Code + "MI.HK";
                for (int j = 1; j < miFmiData.Count; j++)
                {
                    dr[j] = miFmiData[j];
                }
                dt.Rows.Add(dr);
            }

            GenerateCSV(fullpath, "MI_FMI", dt);

        }

        public void GenerateIAAddCSV(string filePath, List<RicInfo> cbbcList, List<RicInfo> warrantList)
        {
            List<string> iAAddTitle = new List<string>() { "HONG KONG CODE", "TYPE", "CATEGORY", "RCS ASSET CLASS", "WARRANT ISSUER" };

            string formatDate = System.DateTime.Now.ToString("yyyyMMdd");
            string fullPath = filePath + "\\YS" + formatDate + "IAAdd.csv";

            System.Data.DataTable dt = GenerateCsvTitle(iAAddTitle);

            for (int i = 0; i < cbbcList.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = cbbcList[i].Code;
                dr[1] = "DERIVATIVE";
                dr[2] = "EIW";
                dr[3] = "FXKNOCKOUT";
                dr[4] = GetIssuerName(cbbcList[i].Issuer, issuerCodeObj)[2];
                dt.Rows.Add(dr);
            }
            for (int i = 0; i < warrantList.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = warrantList[i].Code;
                dr[1] = "DERIVATIVE";
                dr[2] = "EIW";
                dr[3] = "TRAD";
                dr[4] = GetIssuerName(warrantList[i].Issuer, issuerCodeObj)[2];
                dt.Rows.Add(dr);
            }

            GenerateCSV(fullPath, "IAAdd", dt);

        }
        /// <summary>
        /// Generate QAAdd.csv file
        /// </summary>
        /// <param name="filePath">output folder and filename</param>
        /// <param name="cbbcList">CBBC source data</param>
        /// <param name="warrantList">Warrant source data</param>
        public void GenerateQAAddCSV(string filePath, List<RicInfo> cbbcList, List<RicInfo> warrantList)
        {
            List<string> qAAddTitle = new List<string>() { "RIC", "TAG", "ASSET COMMON NAME", "ASSET SHORT NAME", "CURRENCY", "EXCHANGE", "TYPE", "CATEGORY", "BASE ASSET", 
                                                           "EXPIRY DATE", "STRIKE PRICE", "CALL PUT OPTION", "ROUND LOT SIZE", "TRADING SEGMENT", "TICKER SYMBOL", 
                                                           "DERIVATIVES FIRST TRADING DAY", "WARRANT ISSUE PRICE", "WARRANT ISSUE QUANTITY", "WARRANT STAMP DUTY" };

            string formatDate = System.DateTime.Now.ToString("yyyyMMdd");
            string fullPath = filePath + "\\YS" + formatDate + "QAAdd.csv";

            System.Data.DataTable dt = GenerateCsvTitle(qAAddTitle);

            dt = GenerateQAAddData(cbbcList, 1, dt);
            dt = GenerateQAAddData(cbbcList, 2, dt);
            dt = GenerateQAAddData(warrantList, 3, dt);
            dt = GenerateQAAddData(warrantList, 4, dt);
            //dt = GenerateQAAddData(warrantList, 5, dt);

            GenerateCSV(fullPath, "QAAdd", dt);

        }

        /// <summary>
        /// Generate HKG_EQLB_CBBC.txt file(for Main RICs of CBBC).
        /// </summary>
        /// <param name="filePath">output folder and filename.</param>
        /// <param name="ricList">ric source data.</param>
        /// <param name="chineseList">chinese source data.</param>
        public void GenerateHKGEQLBCBBC(string filePath, List<RicInfo> ricList, List<RicInfo> chineseList)
        {
            string fullPath = filePath + "\\HKG_EQLB_CBBC.txt";
            string[] content = new string[ricList.Count + 1];
            content[0] = "SYMBOL\tDSPLY_NAME\tRIC\tOFFCL_CODE\tEX_SYMBOL\tBCKGRNDPAG\tLOT_SIZE_A\tDSPLY_NMLL\tGV1_FLAG\tISS_TP_FLG\tRDM_CUR	MATUR_DATE\tSTRIKE_PRC"
                + "\tWNT_RATIO\t#INSTMOD_MNEMONIC\tBCAST_REF\t#INSTMOD_LOT_SIZE_X\t#INSTMOD_SPARE_UBYTE3	EXL_NAME\t#INSTMOD_GN_TX20_3\t#INSTMOD_GN_TX20_6"
                + "\t#INSTMOD_GN_TX20_7\t#INSTMOD_GN_TX20_10\t#INSTMOD_GN_TX20_12\t#INSTMOD_BOND_TYPE\t#INSTMOD_LONGLINK2\t#INSTMOD_SPARE_SNUM13\t#INSTMOD_GEN_VAL4"
                + "\t#INSTMOD_TDN_SYMBOL\t#INSTMOD_SPARE_DT1\t#INSTMOD_#DDS_LOT_SIZE\t#INSTMOD_#DDS_GN_TX20_25\tCOMMNT\t#INSTMOD_GEN_VAL2\t#INSTMOD_LONGLINK4\t#INSTMOD_LONGLINK6"
                + "\t#INSTMOD_STRIKE_RAT\t#INSTMOD_UNDERLYING\t#INSTMOD_RELNEWS\t#INSTMOD_UNDLY_TYPE\t#INSTMOD_PROT_PRC\t#INSTMOD_EXPIR_DATE";

            for (int i = 0; i < ricList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                bool isExistInDB = false;
                HKUnderlyingInfo underlyingInfo = HKUnderlyingManager.SelectUnderlyingInfoByUnderlying(ricList[i].Underlying);

                if (underlyingInfo != null)
                    isExistInDB = true;

                sb.Append(ricList[i].Code + ".HK");
                sb.Append("\t");
                sb.Append(ricList[i].Name);
                sb.Append("\t");
                sb.Append(ricList[i].Code + ".HK");
                sb.Append("\t");
                sb.Append(ricList[i].Code);
                sb.Append("\t");
                sb.Append(ricList[i].Code);
                sb.Append("\t");
                sb.Append("****");
                sb.Append("\t");
                sb.Append(ricList[i].BoardLot.Replace(",", ""));
                sb.Append("\t");


                // For Bull
                DateTime maturityDateDT = DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null);
                if (ricList[i].BullBear == "Bull")
                {
                    sb.Append(ricList[i].ChineseName.Substring(0, 4).Replace("恒", "恆")
                       + ricList[i].ChineseName.Substring(ricList[i].ChineseName.Length - 1)
                       + maturityDateDT.Month + "月" + "RC" + maturityDateDT.ToString("yy"));

                }
                //For BEAR
                else
                {

                    sb.Append(ricList[i].ChineseName.Substring(0, 4).Replace("恒", "恆")
                       + ricList[i].ChineseName.Substring(ricList[i].ChineseName.Length - 1)
                       + maturityDateDT.Month + "月" + "RP" + maturityDateDT.ToString("yy"));

                }

                sb.Append("\t");

                if (ricList[i].BullBear == "Bull")
                {
                    sb.Append("C");
                }
                //For Bear
                else
                {
                    sb.Append("P");
                }
                sb.Append("\t");

                if (Char.IsLetter(ricList[i].Underlying, 0))
                {
                    sb.Append("I");
                }
                //For Equity
                else
                {
                    sb.Append("S");
                }
                sb.Append("\t");
                sb.Append("344");
                sb.Append("\t");
                sb.Append(DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                sb.Append("\t");
                //For HKD
                if (Char.IsLetter(ricList[i].StrikeLevel, 0))
                {
                    sb.Append(ricList[i].StrikeLevel.Substring(4));
                }
                else
                {
                    sb.Append(ricList[i].StrikeLevel);
                }
                sb.Append("\t");

                //================warrant ratio
                sb.Append((1.0 / Convert.ToInt32(ricList[i].EntitlementRatio)).ToString("0.0000000"));
                sb.Append("\t");
                sb.Append(ricList[i].Code);
                sb.Append("\t");
                //BCAST_REF
                //For Index
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.BCAST_REF);
                }
                else
                {
                    if (Char.IsLetter(ricList[i].Underlying, 0))
                    {
                        if (ricList[i].Underlying == "HSCEI")
                        {
                            sb.Append(".HSCE");
                        }
                        else
                        {
                            sb.Append("." + ricList[i].Underlying);
                        }
                    }
                    //For Equity
                    else
                    {
                        sb.Append(ricList[i].Underlying.Substring(1) + ".HK");
                    }
                }
                sb.Append("\t");
                sb.Append(ricList[i].BoardLot.Replace(",", ""));
                sb.Append("\t");
                sb.Append("14");
                sb.Append("\t");
                sb.Append("HKG_EQLB_CBBC");

                sb.Append("\t");

                //sb.Append("t" + ricList[i].Code + ".HK\t");

                //Append GN_TX20_3
                sb.Append("[HK-\"WARRAN*\"]\t");

                //Append GN_TX20_6
                //For Index
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_GN_TX20_6 + "\t");
                }
                else
                {
                    if (Char.IsLetter(ricList[i].Underlying, 0))
                    {
                        if (ricList[i].Underlying == "HSCEI")
                        {
                            sb.Append("<.HSCE>\t");
                        }
                        else
                        {
                            sb.Append("<." + ricList[i].Underlying + ">\t");
                        }
                    }
                    //For Equity
                    else
                    {
                        sb.Append("<" + ricList[i].Underlying.Substring(1) + ".HK>\t");
                    }
                }
                //Append GN_TX20_7, include 10 spaces after |
                sb.Append("********************\t|          ");
                //Append GN_TX20_10
                sb.Append("CBBC/" + ricList[i].BullBear.ToUpper() + "\t");

                //Append GN_TX20_12 (Misc.Info)
                //For Index
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_GN_TX20_12 + "     <" + ricList[i].Code + "MI.HK>\t");
                }
                else
                {
                    if (Char.IsLetter(ricList[i].Underlying, 0))
                    {
                        sb.Append("IDX     <" + ricList[i].Code + "MI.HK>\t");
                    }
                    //For Equity
                    else
                    {
                        sb.Append("HKD     <" + ricList[i].Code + "MI.HK>\t");
                    }
                }
                //Append BOND_TYPE
                sb.Append("WARRANTS\t");

                //Append LONGLINK2(equal LONGLINK14 in FM)
                //For Index
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_LONGLINK2 + "\t");
                }
                else
                {
                    if (Char.IsLetter(ricList[i].Underlying, 0))
                    {
                        sb.Append("\t");
                    }
                    //For Equity
                    else
                    {
                        sb.Append(ricList[i].Underlying.Substring(1) + ".HK\t");
                    }
                }
                //Append SPARE_SNUM13
                sb.Append("1\t");

                //Append GEN_VAL4(equal Call Level in FM)
                //For HKD
                if (Char.IsLetter(ricList[i].StrikeLevel, 0))
                {

                    sb.Append(ricList[i].CallLevel.Substring(4));
                }
                else
                {
                    sb.Append(ricList[i].CallLevel);
                }
                sb.Append("\t" + ricList[i].Code);

                sb.Append("\t" + MiscUtil.GetLastTradingDay(DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null), holidayList, 1).ToString("dd/MM/yyyy"));
                sb.Append("\t");
                sb.Append(ricList[i].BoardLot.Replace(",", ""));
                sb.Append("\t");
                if (Char.IsLetter(ricList[i].Underlying, 0))
                {
                    sb.Append("I");
                }
                //For Equity
                else
                {
                    sb.Append("S");
                }
                sb.Append("\t");

                sb.Append("Callable " + ricList[i].BullBear + " Contracts");
                sb.Append("\t");

                //For migration phase 1 update
                if (Char.IsLetter(ricList[i].StrikeLevel, 0))
                {
                    sb.Append(ricList[i].StrikeLevel.Substring(4));
                }
                else
                {
                    sb.Append(ricList[i].StrikeLevel);
                }
                sb.Append("\t");

                //Append for the field #INSTMOD_LONGLINK4
                sb.Append("<" + ricList[i].Code + "MI.HK>");
                sb.Append("\t");

                //Append for the field #INSTMOD_LONGLINK6
                sb.Append("******************");
                sb.Append("\t");

                //Append for the field #INSTMOD_STRIKE_RAT
                sb.Append((1.0 / Convert.ToInt32(ricList[i].EntitlementRatio)).ToString("0.0000000"));
                sb.Append("\t");

                //Append for the field #INSTMOD_UNDERLYING
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_UNDERLYING);
                }
                else
                {
                    if (Char.IsLetter(ricList[i].Underlying, 0))
                    {
                        if (ricList[i].Underlying == "HSCEI")
                        {
                            sb.Append("<.HSCE>");
                        }
                        else
                        {
                            sb.Append("<." + ricList[i].Underlying + ">");
                        }
                    }
                    else
                    {
                        sb.Append("<" + ricList[i].Underlying.Substring(1) + ".HK>");
                    }
                }
                sb.Append("\t");

                //Append for the field #INSTMOD_RELNEWS
                sb.Append("[HK-\"WARRAN*\"]");
                sb.Append("\t");

                //Append for the field #INSTMOD_TAS_RIC
                //sb.Append("t" + ricList[i].Code + ".HK");
                //sb.Append("\t");

                //Append for the field #INSTMOD_UNDLY_TYPE
                if (Char.IsLetter(ricList[i].Underlying, 0))
                {
                    sb.Append("3");
                }
                else
                {
                    sb.Append("1");
                }
                sb.Append("\t");

                //Append for the field #INSTMOD_PROT_PRC
                if (Char.IsLetter(ricList[i].StrikeLevel, 0))
                {

                    sb.Append(ricList[i].CallLevel.Substring(4));
                }
                else
                {
                    sb.Append(ricList[i].CallLevel);
                }
                sb.Append("\t");

                //Append the field #INSTMOD_EXPIR_DATE
                sb.Append(DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));

                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            WriteTxtFile(fullPath, content);
            taskResultList.Add(new TaskResultEntry("HKG_EQLB_CBBC.txt", "HKG_EQLB_CBBC.txt File Path", fullPath));
        }

        /// <summary>
        /// Generate HKG_EQLB.txt file (for Main RICs of Warrants).
        /// </summary>
        /// <param name="filePath">output folder and filename.</param>
        /// <param name="ricList">Ric source data.</param>
        /// <param name="chineseList">Chinese source data.</param>
        public void GenerateHKGEQLB(string filePath, List<RicInfo> ricList, List<RicInfo> chineseList)
        {
            string fullPath = filePath + "\\HKG_EQLB.txt";
            string[] content = new string[ricList.Count + 1];
            content[0] = "SYMBOL\tDSPLY_NAME\tRIC\tOFFCL_CODE\tEX_SYMBOL\tBCKGRNDPAG\tLOT_SIZE_A\tDSPLY_NMLL\tGV1_FLAG\tISS_TP_FLG\tRDM_CUR	MATUR_DATE\tSTRIKE_PRC"
                + "\tWNT_RATIO\t#INSTMOD_BOND_TYPE\t#INSTMOD_MNEMONIC\tBCAST_REF\t#INSTMOD_LOT_SIZE_X\t#INSTMOD_SPARE_UBYTE3	EXL_NAME\tBCU\t#INSTMOD_GN_TX20_3"
                + "\t#INSTMOD_GN_TX20_6\t#INSTMOD_GN_TX20_7\t#INSTMOD_GN_TX20_10\t#INSTMOD_GN_TX20_12\t#INSTMOD_LONGLINK2\t#INSTMOD_SPARE_SNUM13\t#INSTMOD_TDN_SYMBOL\t#INSTMOD_SPARE_DT1"
                + "\t#INSTMOD_#DDS_LOT_SIZE \t#INSTMOD_#DDS_PUTCALLIND\t#INSTMOD_#DDS_GN_TX20_25\tCOMMNT\t#INSTMOD_GEN_VAL2\t#INSTMOD_LOTSZUNITS\t#INSTMOD_LONGLINK4\t#INSTMOD_LONGLINK6"
                + "\t#INSTMOD_STRIKE_RAT\t#INSTMOD_LSTTRDDATE\t#INSTMOD_UNDERLYING\t#INSTMOD_RELNEWS\t#INSTMOD_UNDLY_TYPE\t#INSTMOD_EXPIR_DATE";

            for (int i = 0; i < ricList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                bool isExistInDB = false;
                HKUnderlyingInfo underlyingInfo = HKUnderlyingManager.SelectUnderlyingInfoByUnderlying(ricList[i].Underlying);

                if (underlyingInfo != null)
                    isExistInDB = true;

                sb.Append(ricList[i].Code + ".HK");
                sb.Append("\t");
                sb.Append(ricList[i].Name);
                sb.Append("\t");
                sb.Append(ricList[i].Code + ".HK");
                sb.Append("\t");
                sb.Append(ricList[i].Code);
                sb.Append("\t");
                sb.Append(ricList[i].Code);
                sb.Append("\t");
                sb.Append("****");
                sb.Append("\t");
                sb.Append(ricList[i].BoardLot.Replace(",", ""));
                sb.Append("\t");

                //For Call
                if (ricList[i].BullBear == "Call")
                {
                    //The last char of warrant name is a character
                    if (Char.IsLetter(ricList[i].Name, ricList[i].Name.Length - 1))
                    {
                        sb.Append(ricList[i].ChineseName.Substring(0, 4).Replace("恒", "恆")
                           + ricList[i].Name.Substring(ricList[i].Name.Length - 1)
                           + DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).Month + "月" + "CW" + DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).ToString("yy"));
                    }
                    else
                    {
                        sb.Append(ricList[i].ChineseName.Substring(0, 4).Replace("恒", "恆")
                           + DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).Month + "月" + "CW" + DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).ToString("yy"));
                    }

                }
                //For Put
                else
                {
                    if (Char.IsLetter(ricList[i].Name, ricList[i].Name.Length - 1))
                    {
                        sb.Append(ricList[i].ChineseName.Substring(0, 4).Replace("恒", "恆")
                           + ricList[i].Name.Substring(ricList[i].Name.Length - 1)
                           + DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).Month + "月" + "PW" + DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).ToString("yy"));

                    }
                    else
                    {

                        sb.Append(ricList[i].ChineseName.Substring(0, 4).Replace("恒", "恆")
                           + DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).Month + "月" + "PW" + DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).ToString("yy"));
                    }

                }

                sb.Append("\t");

                //For Call
                if (ricList[i].BullBear == "Call")
                {
                    sb.Append("C");
                }
                //For Put
                else
                {
                    sb.Append("P");
                }
                sb.Append("\t");

                //For Index
                if (ricList[i].Underlying == "HSI" || ricList[i].Underlying == "HSCEI" || ricList[i].Underlying == "DJI")
                {
                    sb.Append("I");
                }
                //For Stock
                else if (Char.IsDigit(ricList[i].Underlying, 0))
                {
                    sb.Append("S");
                }
                else//For Forex ricList[i].Name.Contains("EURUS/AUDUS/USYEN")
                {
                    sb.Append("F");
                }
                sb.Append("\t");
                sb.Append("344");
                sb.Append("\t");
                sb.Append(DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                sb.Append("\t");
                //For HKD
                if (Char.IsLetter(ricList[i].StrikeLevel, 0))
                {
                    sb.Append(Convert.ToDouble(ricList[i].StrikeLevel.Substring(4)).ToString("0.000"));
                }
                else
                {
                    sb.Append(Convert.ToDouble(ricList[i].StrikeLevel).ToString("0.000"));
                }
                sb.Append("\t");
                double enRatio = 1.0 / Convert.ToDouble(ricList[i].EntitlementRatio);
                string x = enRatio.ToString();
                String entitlemenRatioStr = enRatio.ToString().Length >= 9 == true ? enRatio.ToString("0.0000000") : enRatio.ToString();
                sb.Append(entitlemenRatioStr);
                sb.Append("\t");
                sb.Append("WARRANTS");
                sb.Append("\t");
                sb.Append(ricList[i].Code);
                sb.Append("\t");
                //BCAST_REF
                //For Index
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.BCAST_REF);
                }
                else
                {
                    if (Char.IsLetter(ricList[i].Underlying, 0))
                    {
                        if (ricList[i].Underlying == "HSCEI")
                        {
                            sb.Append(".HSCE");
                        }
                        else
                        {
                            sb.Append("." + ricList[i].Underlying);
                        }
                    }
                    //For Equity
                    else
                    {
                        sb.Append(ricList[i].Underlying.Substring(1) + ".HK");
                    }
                }

                sb.Append("\t");
                sb.Append(ricList[i].BoardLot.Replace(",", ""));
                sb.Append("\t");

                //For Call
                if (ricList[i].BullBear == "Call")
                {
                    sb.Append("5");
                }
                //For Put
                else
                {
                    sb.Append("6");
                }

                sb.Append("\t");
                sb.Append("HKG_EQLB");
                sb.Append("\t");
                //For stock
                if (Char.IsDigit(ricList[i].Underlying, 0))
                {
                    sb.Append("HKG_EQ_CWRTS");
                }

                sb.Append("\t");

                //sb.Append("t" + ricList[i].Code + ".HK\t");

                //Append GN_TX20_3
                sb.Append("[HK-\"WARRAN*\"]\t");

                //Append GN_TX20_6
                //for Equity stock
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_GN_TX20_6 + "\t");
                }
                else
                {
                    if (Char.IsDigit(ricList[i].Underlying, 0))
                    {
                        sb.Append("<" + ricList[i].Underlying.Substring(1) + ".HK>\t");
                    }
                    //for Index
                    else if (ricList[i].Underlying == "HSI" || ricList[i].Underlying == "HSCEI" || ricList[i].Underlying == "DJI")
                    {
                        if (ricList[i].Underlying == "HSCEI")
                        {
                            sb.Append("<.HSCE>\t");
                        }
                        else
                        {
                            sb.Append("<." + ricList[i].Underlying + ">\t");
                        }
                    }
                    //for others
                    else
                    {
                        sb.Append("\t");
                    }
                }
                //Append GN_TX20_7, include 12 spaces after |
                //for Equity stock
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_GN_TX20_7 + "\t");
                }
                else
                {
                    if (Char.IsDigit(ricList[i].Underlying, 0))
                    {
                        sb.Append("<" + ricList[i].Underlying.Substring(1) + "DIVCF.HK>\t|            ");
                    }
                    else
                    {
                        sb.Append("********************\t|            ");
                    }
                }

                //Append GN_TX20_10
                sb.Append("EU/" + ricList[i].BullBear.ToUpper() + "\t");

                //Append GN_TX20_12 (Misc.Info)
                //For Index
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_GN_TX20_12 + "     <" + ricList[i].Code + "MI.HK>\t");
                }
                else
                {
                    if (Char.IsLetter(ricList[i].Underlying, 0))
                    {
                        sb.Append("IDX     <" + ricList[i].Code + "MI.HK>\t");
                    }
                    //For Equity
                    else if (Char.IsDigit(ricList[i].Underlying, 0))
                    {
                        sb.Append("HKD     <" + ricList[i].Code + "MI.HK>\t");
                    }
                    else
                    {
                        sb.Append("              \t");
                    }
                }

                //Append LONGLINK2(equal LONGLINK14 in FM)
                //For Index
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_LONGLINK2 + "\t");
                }
                else
                {
                    if (Char.IsLetter(ricList[i].Underlying, 0))
                    {
                        if (ricList[i].Underlying == "HSI")
                        {
                            sb.Append(".HSI|HKD|1\t");
                        }
                        else if (ricList[i].Underlying == "HSCEI")
                        {
                            sb.Append(".HSCE|HKD|1\t");
                        }
                        else if (ricList[i].Underlying == "DJI")
                        {
                            sb.Append(".DJI|USD|1\t");
                        }
                        else
                        {
                            sb.Append("       \t");
                        }

                    }
                    //For Stock
                    else if (Char.IsDigit(ricList[i].Underlying, 0))
                    {
                        sb.Append(ricList[i].Underlying.Substring(1) + ".HK\t");
                    }
                    else
                    {
                        sb.Append("       \t");
                    }
                }

                //Append SPARE_SNUM13
                sb.Append("1\t");
                sb.Append(ricList[i].Code);
                sb.Append("\t" + MiscUtil.GetLastTradingDay(DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null), holidayList, 4).ToString("dd/MM/yyyy"));
                sb.Append("\t");
                sb.Append(ricList[i].BoardLot.Replace(",", ""));
                sb.Append("\t");
                //For Call
                if (ricList[i].BullBear == "Call")
                {
                    sb.Append("5");
                }
                //For Put
                else
                {
                    sb.Append("6");
                }

                sb.Append("\t");

                //For Index
                if (ricList[i].Underlying == "HSI" || ricList[i].Underlying == "HSCEI" || ricList[i].Underlying == "DJI")
                {
                    sb.Append("I");
                }
                //For Stock
                else if (Char.IsDigit(ricList[i].Underlying, 0))
                {
                    sb.Append("S");
                }
                else//For Forex ricList[i].Name.Contains("EURUS/AUDUS/USYEN")
                {
                    sb.Append("F");
                }
                sb.Append("\t");
                sb.Append("Equity Warrant");
                sb.Append("\t");

                //For migration phase 1 update
                if (Char.IsLetter(ricList[i].StrikeLevel, 0))
                {
                    sb.Append(Convert.ToDouble(ricList[i].StrikeLevel.Substring(4)).ToString("0.000"));
                }
                else
                {
                    sb.Append(Convert.ToDouble(ricList[i].StrikeLevel).ToString("0.000"));
                }
                sb.Append("\t");

                // Append the field #INSTMOD_LOTSZUNITS
                if (Char.IsLetter(ricList[i].Underlying, 0))
                {
                    sb.Append("INDX");
                }
                else if (Char.IsDigit(ricList[i].Underlying, 0))
                {
                    sb.Append("HKD");
                }
                else
                {
                    sb.Append("   ");
                }
                sb.Append("\t");

                // Append the field #INSTMOD_LONGLINK4
                if (Char.IsLetter(ricList[i].Underlying, 0))
                {
                    sb.Append("<" + ricList[i].Code + "MI.HK>");
                }
                else if (Char.IsDigit(ricList[i].Underlying, 0))
                {
                    sb.Append("<" + ricList[i].Code + "MI.HK>");
                }
                else
                {
                    sb.Append("");
                }
                sb.Append("\t");

                // Append the field #INSTMOD_LONGLINK6
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_LONGLINK6);
                }
                else
                {
                    if (Char.IsDigit(ricList[i].Underlying, 0))
                    {
                        sb.Append("<" + ricList[i].Underlying.Substring(1) + "DIVCF.HK>");
                    }
                    else
                    {
                        sb.Append("******************");
                    }
                }
                sb.Append("\t");

                // Append the field #INSTMOD_STRIKE_RAT
                sb.Append(entitlemenRatioStr);
                sb.Append("\t");

                // Append the field #INSTMOD_LSTTRDDATE
                sb.Append(MiscUtil.GetLastTradingDay(DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null), holidayList, 4).ToString("dd/MM/yyyy"));
                sb.Append("\t");

                // Append the field #INSTMOD_UNDERLYING
                if (isExistInDB)
                {
                    sb.Append(underlyingInfo.INSTMOD_UNDERLYING + "\t");
                }
                else
                {
                    if (Char.IsDigit(ricList[i].Underlying, 0))
                    {
                        sb.Append("<" + ricList[i].Underlying.Substring(1) + ".HK>\t");
                    }
                    else if (ricList[i].Underlying == "HSI" || ricList[i].Underlying == "HSCEI" || ricList[i].Underlying == "DJI")
                    {
                        if (ricList[i].Underlying == "HSCEI")
                        {
                            sb.Append("<.HSCE>\t");
                        }
                        else
                        {
                            sb.Append("<." + ricList[i].Underlying + ">\t");
                        }
                    }
                    else
                    {
                        sb.Append("\t");
                    }
                }

                // Append the field #INSTMOD_RELNEWS
                sb.Append("[HK-\"WARRAN*\"]\t");

                // Append the field #INSTMOD_TAS_RIC
                //sb.Append("t" + ricList[i].Code + ".HK\t");

                // Append the field #INSTMOD_UNDLY_TYPE
                if (ricList[i].Underlying == "HSI" || ricList[i].Underlying == "HSCEI" || ricList[i].Underlying == "DJI")
                {
                    sb.Append("3");
                }
                else if (Char.IsDigit(ricList[i].Underlying, 0))
                {
                    sb.Append("1");
                }
                else
                {
                    sb.Append("7");
                }
                sb.Append("\t");

                //Append the field #INSTMOD_EXPIR_DATE
                sb.Append(DateTime.ParseExact(ricList[i].MaturityDate, "dd-MM-yyyy", null).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));

                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);

            }

            WriteTxtFile(fullPath, content);
            taskResultList.Add(new TaskResultEntry("HKG_EQLB.txt", "HKG_EQLB.txt File Path", fullPath));
        }

        /// <summary>
        /// Generate HKG_EQLBMI.txt file (for MI RICs of Warrants&CBBC).
        /// </summary>
        /// <param name="filePath">output folder and filename.</param>
        /// <param name="cbbcList">Cbbc source data.</param>
        /// <param name="warrantList">Warrant source data.</param>
        public void GenerateHKGEQLBMI(string filePath, List<RicInfo> cbbcList, List<RicInfo> warrantList)
        {
            string fullPath = filePath + "\\HKG_EQLBMI.txt";
            string[] content = new string[cbbcList.Count + warrantList.Count + 1];
            content[0] = "SYMBOL\tDSPLY_NAME\tRIC\tEX_SYMBOL\t#INSTMOD_ROW80_1\tEXL_NAME";

            for (int i = 0; i < cbbcList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(cbbcList[i].Code + "MI.HK");
                sb.Append("\t");
                sb.Append(cbbcList[i].Name);
                sb.Append("\t");
                sb.Append(cbbcList[i].Code + "MI.HK");
                sb.Append("\t");
                sb.Append(cbbcList[i].Code + "MI");
                sb.Append("\t");
                //include 36 spaces at end
                sb.Append("Security Miscellaneous Information                                    ");
                sb.Append(cbbcList[i].Code + "MI.HK");
                sb.Append("\t");
                sb.Append("HKG_EQLB_MI_PAGE");
                sb.Append("\t");
                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            int j = cbbcList.Count + 1;
            for (int i = 0; i < warrantList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(warrantList[i].Code + "MI.HK");
                sb.Append("\t");
                sb.Append(warrantList[i].Name);
                sb.Append("\t");
                sb.Append(warrantList[i].Code + "MI.HK");
                sb.Append("\t");
                sb.Append(warrantList[i].Code + "MI");
                sb.Append("\t");
                //include 36 spaces at end
                sb.Append("Security Miscellaneous Information                                    ");
                sb.Append(warrantList[i].Code + "MI.HK");
                sb.Append("\t");
                sb.Append("HKG_EQLB_MI_PAGE");
                sb.Append("\t");
                content[j++] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            WriteTxtFile(fullPath, content);
            taskResultList.Add(new TaskResultEntry("HKG_EQLBMI.txt", "HKG_EQLBMI.txt File Path", fullPath));
        }

        /// <summary>
        /// Generate MI.txt file (for MI RICs of Warrants&CBBC).
        /// </summary>
        /// <param name="filePath">output folder and filename.</param>
        /// <param name="cbbcList">Cbbc source data.</param>
        /// <param name="warrantList">Warrant source data.</param>
        public void GenerateMI(string filePath, List<RicInfo> cbbcList, List<RicInfo> warrantList)
        {
            string fullPath = filePath + "\\MI.txt";
            string[] content = new string[cbbcList.Count + warrantList.Count + 1];
            content[0] = "HKSE.DAT;07-APR-2004 09:00:00;TPS;\r\nROW80_3;ROW80_4;ROW80_5;ROW80_6;ROW80_7;ROW80_8;ROW80_9;ROW80_10;ROW80_11;ROW80_12;ROW80_13;ROW80_14;ROW80_15;ROW80_16;";
            for (int i = 0; i < cbbcList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(cbbcList[i].Code + "MI.HK;");
                sb.Append("ISIN                                    EIPO Start Date;");
                sb.Append("Instrument Type                         EIPO End Date;");
                sb.Append("Market                                  EIPO Start Time;");
                sb.Append("Sub-Market                              EIPO End Time;");
                sb.Append("Listing Date                            EIPO Price;");
                sb.Append("De-listing Date                         Spread Table;");
                sb.Append("Listing Status                          Shortselling Stock;");
                sb.Append("Trading Status                          Intra-day Shortselling Stock;");
                sb.Append("Stamp Duty                              Automatch Stock;");
                sb.Append("Test Stock                              CCASS Stock;");
                sb.Append("Dummy Stock;");
                sb.Append("--------------------------------------------------------------------------------;");
                sb.Append("Trading Start Time                      Trading End Time;");
                sb.Append("Session Type;");
                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);
            }
            int j = cbbcList.Count + 1;
            for (int i = 0; i < warrantList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(warrantList[i].Code + "MI.HK;");
                sb.Append("ISIN                                    EIPO Start Date;");
                sb.Append("Instrument Type                         EIPO End Date;");
                sb.Append("Market                                  EIPO Start Time;");
                sb.Append("Sub-Market                              EIPO End Time;");
                sb.Append("Listing Date                            EIPO Price;");
                sb.Append("De-listing Date                         Spread Table;");
                sb.Append("Listing Status                          Shortselling Stock;");
                sb.Append("Trading Status                          Intra-day Shortselling Stock;");
                sb.Append("Stamp Duty                              Automatch Stock;");
                sb.Append("Test Stock                              CCASS Stock;");
                sb.Append("Dummy Stock;");
                sb.Append("--------------------------------------------------------------------------------;");
                sb.Append("Trading Start Time                      Trading End Time;");
                sb.Append("Session Type;");
                content[j++] = sb.ToString();
                sb.Remove(0, sb.Length);
            }
            WriteTxtFile(fullPath, content);
            taskResultList.Add(new TaskResultEntry("MI.txt", "MT.txt File Path", fullPath));
        }

        /// <summary>
        /// Generate BK.txt file (for bk RICs of warrants & CBBC).
        /// </summary>
        /// <param name="filePath">output folder and filename.</param>
        /// <param name="cbbcList">Cbbc source data.</param>
        /// <param name="warrantList">Warrant source data.</param>
        public void GenerateBK(string filePath, List<RicInfo> cbbcList, List<RicInfo> warrantList)
        {
            string fullPath = filePath + "\\BK.txt";
            string[] content = new string[cbbcList.Count + warrantList.Count + 1];
            content[0] = "RIC\t#INSTMOD_ROW80_1\t#INSTMOD_ROW80_2\t#INSTMOD_ROW80_13";


            for (int i = 0; i < cbbcList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(cbbcList[i].Code + "bk.HK\t");

                sb.Append(cbbcList[i].Name);

                string startPosition = (17 - cbbcList[i].Name.Length).ToString();
                switch (startPosition)
                {
                    case "0": sb.Append(""); break;
                    case "1": sb.Append(" "); break;
                    case "2": sb.Append("  "); break;
                    case "3": sb.Append("   "); break;
                }


                sb.Append("<" + cbbcList[i].Code + ".HK> <HKBK01>                                  " + cbbcList[i].Code + "bk.HK\t");


                sb.Append("        BID                    ASK\t");

                if (cbbcList[i].BullBear == "Bull")
                {
                    sb.Append("Callable Bull Contracts");
                }
                else
                {
                    sb.Append("Callable Bear Contracts");
                }

                content[i + 1] = sb.ToString();
                sb.Remove(0, sb.Length);
            }
            int j = cbbcList.Count + 1;
            for (int i = 0; i < warrantList.Count; i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(warrantList[i].Code + "bk.HK\t");
                sb.Append(warrantList[i].Name);

                string startPosition = (17 - warrantList[i].Name.Length).ToString();
                switch (startPosition)
                {
                    case "0": sb.Append(""); break;
                    case "1": sb.Append(" "); break;
                    case "2": sb.Append("  "); break;
                    case "3": sb.Append("   "); break;
                }

                sb.Append("<" + warrantList[i].Code + ".HK> <HKBK01>                                  " + warrantList[i].Code + "bk.HK\t");


                sb.Append("        BID                    ASK\t");

                //For index
                if (Char.IsLetter(warrantList[i].Underlying, 0))
                {
                    sb.Append("Warrant Type-Index Warrant");
                }
                else if (Char.IsDigit(warrantList[i].Underlying, 0))//For stock
                {
                    sb.Append("Warrant Type-Equity Warrant");
                }
                else
                {
                    sb.Append("");
                }
                content[j++] = sb.ToString();
                sb.Remove(0, sb.Length);
            }

            WriteTxtFile(fullPath, content);
            taskResultList.Add(new TaskResultEntry("BK.txt", "BK.txt File Path", fullPath));
        }

        private void WriteTxtFile(string fullpath, string[] content)
        {
            try
            {
                File.WriteAllLines(fullpath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                string errInfo = ex.ToString();
            }
        }

        protected System.Data.DataTable GenerateCsvTitle(List<string> csvTitle)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            foreach (string title in csvTitle)
            {
                dt.Columns.Add(title);
            }

            DataRow dr1 = dt.NewRow();
            for (int i = 0; i < csvTitle.Count; i++)
            {
                dr1[i] = csvTitle[i];
            }
            dt.Rows.Add(dr1);

            return dt;
        }

        protected System.Data.DataTable GenerateQAAddData(List<RicInfo> sourceList, int index, System.Data.DataTable dt)
        {
            for (int i = 0; i < sourceList.Count; i++)
            {
                string tempBullBear = "";
                string tempLongName = "";
                if (index == 1 || index == 2)
                {
                    tempLongName = GetIDNLongName(sourceList[i]).Replace('@', ' ');

                    if (sourceList[i].BullBear.ToUpper().Equals("BULL"))
                    {
                        tempBullBear = "CALL";
                    }
                    else if (sourceList[i].BullBear.ToUpper().Equals("BEAR"))
                    {
                        tempBullBear = "PUT";
                    }
                }
                else
                {
                    bool isIndex = false;
                    bool isHKD = false;
                    bool isStock = false;
                    bool isCall = false;
                    if (sourceList[i].Underlying == "HSI" || sourceList[i].Underlying == "HSCEI" || sourceList[i].Underlying == "DJI")
                    {
                        isIndex = true;
                    }
                    if (Char.IsLetter(sourceList[i].StrikeLevel, 0))
                    {
                        isHKD = true;
                    }
                    if (Char.IsDigit(sourceList[i].Underlying, 0))
                    {
                        isStock = true;
                    }

                    if (sourceList[i].BullBear == "Call")
                    {
                        isCall = true;
                    }
                    tempLongName = GetIDNLongNameForWarrant(sourceList[i], isIndex, isStock, isCall, isHKD, issuerCodeObj).Replace('@', ' ');
                    tempBullBear = sourceList[i].BullBear.ToUpper();
                }

                DataRow dr = dt.NewRow();
                if (index == 1 || index == 3)
                {
                    dr[0] = sourceList[i].Code + ".HK";
                    dr[1] = "2037";
                    dr[13] = "HKG:XHKG";
                    dr[14] = sourceList[i].Code;
                    dr[18] = "N";
                    dr[5] = "HKG";
                }
                else if (index == 2 || index == 4)
                {
                    dr[0] = sourceList[i].Code + "ta.HK";
                    dr[1] = "40115";
                    dr[13] = "";
                    dr[14] = "";
                    dr[5] = "HKG";
                }
                else
                {
                    dr[0] = sourceList[i].Code + ".IXH";
                    dr[1] = "46111";
                    dr[13] = "";
                    dr[14] = sourceList[i].Code;
                    dr[5] = "IXH";
                }

                dr[2] = tempLongName;
                dr[3] = sourceList[i].Name.Replace('@', ' ');
                dr[4] = "HKD";

                dr[6] = "DERIVATIVE";
                dr[7] = "EIW";
                dr[8] = "";
                dr[9] = DateTime.ParseExact(sourceList[i].MaturityDate, "dd-MM-yyyy", null).ToString("dd-MMM-yy", new CultureInfo("en-US"));

                if (Char.IsLetter(sourceList[i].StrikeLevel, 0))
                {
                    dr[10] = sourceList[i].StrikeLevel.Substring(4);
                }
                else
                {
                    dr[10] = sourceList[i].StrikeLevel;
                }

                dr[11] = tempBullBear;
                dr[12] = sourceList[i].BoardLot;
                dr[15] = DateTime.ParseExact(sourceList[i].ListingDate, "dd-MM-yyyy", null).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                dr[16] = sourceList[i].IssuerPrice;
                dr[17] = sourceList[i].TotalIssueSize;

                dt.Rows.Add(dr);
            }

            return dt;
        }

        protected void GenerateCSV(string filePath, string taskName, System.Data.DataTable dataTable)
        {
            ExcelApp xlApp = new ExcelApp(false, false);
            if (xlApp.ExcelAppInstance == null)
            {
                String msg = "Excel could not be started. Check that your office installation and project reference correct!!!";
                logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                Workbook wBook = xlApp.ExcelAppInstance.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];

                if (wSheet == null)
                {
                    throw new Exception("Worksheet could not be created. Check that your office installation and project references are correct.");
                }

                ExcelLineWriter writer = new ExcelLineWriter(wSheet, 1, 1, ExcelLineWriter.Direction.Right);

                foreach (DataRow dr in dataTable.Rows)
                {
                    foreach (DataColumn dc in dataTable.Columns)
                    {
                        writer.WriteLine(dr[dc]);
                    }
                    writer.PlaceNext(writer.Row + 1, 1);
                }

                xlApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.SaveAs(filePath, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                taskResultList.Add(new TaskResultEntry(taskName + ".csv", taskName + ".csv File Path", filePath));
            }
            catch (Exception ex)
            {
                String msg = "Error found in write to csv file : " + filePath + ex.ToString();
                logger.Log(msg, Logger.LogType.Error);
                return;
            }
            finally
            {
                xlApp.Dispose();
            }
        }

        internal void Start()
        {
            StartHKBulkFileGeneratorJob();
        }
    }
}
