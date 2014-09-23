using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Util;
using Ric.Core;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Drawing.Design;
using Ric.Db.Manager;
using Ric.Db.Info;
using HtmlAgilityPack;

namespace Ric.Tasks
{
    public enum FMType
    {
        Cbbc,
        Warrant
    }
    public class HKRicTemplate
    {
        private static readonly string issuerCodeMapPath = ".\\Config\\HK\\HK_IssuerCode.xml";
        private static HK_IssuerCodeMap issuerCodeObj;
        public static string ErrorMsg = string.Empty;

        public string EffectiveDate { get; set; }
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
        public string WarrantType { get; set; }
        public string MiscInfoPageChainRic { get; set; }
        public string LotSize { get; set; }
        public string ColDsplyNmll { get; set; }
        public string BcastRef { get; set; }
        public string WntRation { get; set; }
        public string StrikPrc { get; set; }
        public string MaturDate { get; set; }
        public string LongLink3 { get; set; }
        public string SpareSnum13 { get; set; }
        public string GNTX20_3 { get; set; }
        public string GNTX20_6 { get; set; }
        public string GNTX20_7 { get; set; }
        public string GNTX20_10 { get; set; }
        public string GNTX20_11 { get; set; }
        public string GNTX20_12 { get; set; }
        public string CouponRate { get; set; }
        public string IssuePrice { get; set; }
        public string Row80_13 { get; set; }
        public string GVFlag { get; set; }
        public string IssTpFlg { get; set; }
        public string RdmCur { get; set; }
        public string LongLink14 { get; set; }
        public string BondType { get; set; }
        public string Leg1Str { get; set; }
        public string Leg2Str { get; set; }
        public string GNTXT24_1 { get; set; }
        public string GNTXT24_2 { get; set; }

        //For NDA
        public string NewOrgList { get; set; }
        public string PrimaryList { get; set; }
        public string IdnLongName { get; set; }
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

        //For WRT_CNR
        public string Gearing { get; set; }
        public string Premium { get; set; }
        public string AnnouncementDate { get; set; }
        public string PaymentDate { get; set; }
        public string CallLevel { get; set; }

        public HKRicTemplate(RicInfo ricInfo,FMType fmType)
        {
            bool isIndex = false;
            bool isCall = false;
            bool isHKD = false;
            bool isStock = false;
            bool isOil = false;
            bool isCommodity = false;
            if (ricInfo.Underlying == "HSI" || ricInfo.Underlying == "HSCEI" || ricInfo.Underlying == "DJI")
            {
                isIndex = true;
            }

            if (Char.IsDigit(ricInfo.Underlying, 0))
            {
                isStock = true;
            }
            if(ricInfo.Name.Contains("OIL"))
            {
                isOil=true;
            }

            if (ricInfo.BullBear.ToLower().Contains("call"))
            {
                isCall = true;
            }
            if (ricInfo.Name.Contains("GOLD") || ricInfo.Name.Contains("SILVER"))
            {
                isCommodity = true;
            }
            if (Char.IsLetter(ricInfo.StrikeLevel, 0))
            {
                isHKD = true;
            }

            issuerCodeObj = ConfigUtil.ReadConfig(issuerCodeMapPath, typeof(HK_IssuerCodeMap)) as HK_IssuerCodeMap;
            DateTime effectiveDate = DateTime.ParseExact(ricInfo.ListingDate, "dd-MM-yyyy", null);
            EffectiveDate = effectiveDate.ToString("dd-MMM-yy");
            UnderLyingRic = ricInfo.Code + ".HK";
            CompositeChainRic = "0#" + ricInfo.Code + ".HK";
            BrokerPageRic = ricInfo.Code + "bk.HK";
            MiscInfoPageRic = ricInfo.Code + "MI.HK";
            DisplayName = ricInfo.Name;
            OfficicalCode = ricInfo.Code;
            ExchangeSymbol = ricInfo.Code;
            Currency = "HKD";
            RecordType = "097";
            SpareUbytes8 = "WRNT";
            #region For CBBC only
            //For CBBC only
            if (fmType == FMType.Cbbc)
            {
                IdnLongName = GetIDNLongName(ricInfo);
                UnderlyingChainRic1 = "0#CBBC.HK";
                UnderlyingChainRic2 = "0#WARRANTS.HK";
                GNTX20_10 = "CBBC/" + ricInfo.BullBear;
                GNTX20_10 = GNTX20_10.ToUpper();
                Row80_13 = "Callable " + ricInfo.BullBear + " Contracts";
                WarrantType = "Callable " + ricInfo.BullBear + " Contracts";
                RdmCur = "344";
                if (ricInfo.BullBear.ToLower().Contains("bear"))
                {
                    GVFlag = "P";
                }
                if (ricInfo.BullBear.ToLower().Contains("bull"))
                {
                    GVFlag = "C";
                }
                if (Char.IsLetter(ricInfo.StrikeLevel, 0))
                {
                    CallLevel = ricInfo.CallLevel.Substring(4);
                }
                else
                {
                    CallLevel = ricInfo.CallLevel;
                }
                if (Char.IsLetter(ricInfo.Underlying, 0))
                {
                    GNTX20_12 = "IDX~~~~~~<" + ricInfo.Code + "MI.HK>";
                    IssTpFlg = "I";
                    if (ricInfo.Underlying == "HSCEI")
                    {
                        GNTX20_6 = "<.HSCE";
                        LongLink14 = ".HSCE|HKD|1  <-- TQR INSERT DAILY";

                    }
                    else if (ricInfo.Underlying == "HSI")
                    {
                        LongLink14 = ".HSI|HKD|1  <-- TQR INSERT DAILY";
                        GNTX20_6 = "<." + ricInfo.Underlying + ">";
                    }

                    else if (ricInfo.Underlying == "DJI")
                    {
                        LongLink14 = ".DJI|USD|1  <-- TQR INSERT DAILY";
                        GNTX20_6 = "<." + ricInfo.Underlying + ">";
                    }
                    else
                    {
                        GNTX20_6 = "<." + ricInfo.Underlying + ">";
                        LongLink14 = "Index not equal HSI,HSCEI or DJI";
                    }


                }
                else
                {
                    GNTX20_12 = "HKD~~~~~~<" + ricInfo.Code + "MI.HK>";
                    IssTpFlg = "S";
                    GNTX20_6 = "<" + ricInfo.Underlying + ".HK>";
                    LongLink14 = ricInfo.Underlying.Substring(1) + ".HK";
                }
            }
            #endregion

            #region For Warrant only
            else
            {
                GNTX20_10 = "EU/" + ricInfo.BullBear;
                GNTX20_10 = GNTX20_10.ToUpper();
                Row80_13 = "Equity Warrant";
                IdnLongName = GetIDNLongNameForWarrant(ricInfo, isIndex, isStock, isCall, isHKD, issuerCodeObj);
                if (isCall)
                {
                    GVFlag = "C";
                }
                else
                {
                    GVFlag = "P";
                }
                if(isStock)
                {
                    UnderlyingChainRic1 = "0#" + ricInfo.Underlying.Substring(1) + "W.HK";
                    ChainRic1 = "0#CWRTS.HK";
                    WarrantType = "Equity Warrant";
                    GNTX20_6 = "<" + ricInfo.Underlying.Substring(1) + ".HK>";
                    GNTX20_12="HKD~~~~~~<" + ricInfo.Code + "MI.HK>";
                    IssTpFlg = "S";
                    RdmCur = "344";
                    LongLink14=ricInfo.Underlying.Substring(1) + ".HK";
                }

                else if (isIndex)
                {                    
                    ChainRic1 = "";
                    Row80_13 = "Index Warrant";
                    GNTX20_12 = "IDX~~~~~~<" + ricInfo.Code + "MI.HK>";
                    RdmCur = "344";
                    if (ricInfo.Underlying == "HSCEI")
                    {
                        UnderlyingChainRic1 = "0#.HSCEW.HK";
                        WarrantType = "Hang Seng China Enterprises Index Warrant";
                        GNTX20_6 = "<.HSCE>";
                        LongLink14=".HSCE|HKD|1  <-- TQR INSERT DAILY";
                    }
                    else if (ricInfo.Underlying == "HSI")
                    {
                        WarrantType = "Hang Seng Index Warrant";
                        UnderlyingChainRic1 = "0#." + ricInfo.Underlying + "W.HK";
                        GNTX20_6 = "<." + ricInfo.Underlying + ">";
                        LongLink14=".HSI|HKD|1  <-- TQR INSERT DAILY";
                    }
                    else
                    {
                        if(ricInfo.Underlying=="DJI")
                        {
                            LongLink14=".DJI|USD|1  <-- TQR INSERT DAILY";
                        }
                        else
                        {
                            LongLink14="Index not equal HSI,HSCEI or DJI";
                        }
                        UnderlyingChainRic1 = "0#." + ricInfo.Underlying + "W.HK";
                        WarrantType="DJ Industrial Average Index Warrant";
                    }
                    IssTpFlg = "I";
                }
                else if (isOil)
                {
                    ChainRic1 = "0#OWRTS.HK";
                    WarrantType = "Oil Warrant";
                    Row80_13 = "Future Warrant";
                    RdmCur = "840";
                }
                else if (isCommodity)
                {
                    Row80_13 = "Commodity Warrant";
                    UnderlyingChainRic1 = string.Empty;
                    ChainRic1 = "0#OWRTS.HK";
                    WarrantType = "Miscellaneous Warrant";
                    RdmCur = "840";
                }
                else
                {
                    //isCurrency = true;
                    if(ricInfo.Name.Contains("YEN"))
                    {
                        RdmCur = "392";
                    }
                    UnderlyingChainRic1 = string.Empty;
                    ChainRic1 = "0#OWRTS.HK";
                    WarrantType = "Miscellaneous Warrant";
                    Row80_13 = "Currency Warrant";
                    IssTpFlg = "F";
                    RdmCur = string.Empty;
                    LongLink14=string.Empty;
                }
            }
            #endregion

            ChainRic2 = "0#WARRANTS.HK";

            MiscInfoPageChainRic = "0#MI.HK";
            LotSize = ricInfo.BoardLot;
            //ColDsplyNmll
            ColDsplyNmll = GetColDsplyNmllInfo(ricInfo);
            BcastRef = "n/a";
            WntRation = (1.0 / Convert.ToDouble(ricInfo.EntitlementRatio)).ToString();

            if (Char.IsLetter(ricInfo.StrikeLevel, 0))
            {
                StrikPrc = ricInfo.StrikeLevel.Substring(4);
            }
            else
            {
                StrikPrc = ricInfo.StrikeLevel;
            }

            DateTime maturDate = DateTime.ParseExact(ricInfo.MaturityDate, "dd-MM-yyyy", null);
            MaturDate = maturDate.ToString("dd-MMM-yy");

            LongLink3 = "t" + ricInfo.Code + ".HK";
            SpareSnum13 = "1";
            GNTX20_3 = "[HK-\"WARRAN*\"]";
            GNTX20_7 = "********************";

            GNTX20_11 = ricInfo.TotalIssueSize;
            CouponRate = "n/a";
            IssuePrice = ricInfo.IssuerPrice;
            BondType = "WARRANTS";
            Leg1Str = string.Empty;
            Leg2Str = string.Empty;
            GNTXT24_1 = string.Empty;
            GNTXT24_2 = string.Empty;

            NewOrgList = string.Empty;
            PrimaryList = string.Empty;
            OrgnizationName1 = string.Empty;
            GeographyEntity = string.Empty;
            OrgnizationType = string.Empty;
            AliasPre = string.Empty;
            AliasGen = string.Empty;
            IssueClassification = string.Empty;
            MSCICode = string.Empty;
            BusinessActivity = string.Empty;
            ExistingOrgList = string.Empty;
            PrimaryList = string.Empty;
            OrgnizationName2 = GetIssuerName(ricInfo.Issuer, issuerCodeObj)[1];
            IssueClassification = "WNT";
            //For WRT_CNR
            Gearing = ricInfo.Gear;
            Premium = ricInfo.Premium;
            DateTime announcementData = DateTime.ParseExact(ricInfo.LauntchDate, "dd-MM-yyyy", null);
            AnnouncementDate = announcementData.ToString("dd-MMM-yy");
            PaymentDate = DateTime.ParseExact(ricInfo.ClearingCommencementDate, "dd-MM-yyyy", null).ToString("dd-MMM-yy");
        }



        //Get ColDsplyNmll:  ColDsplyNmll = Underlying Chinses name + Issuer Chinses name+ Letter if any+ Month + RP/RC +Year
        public string GetColDsplyNmllInfo(RicInfo ricEnglishInfo)
        {        
            string chineseName = ricEnglishInfo.ChineseName;
            string colDsplyNmll = chineseName.Substring(0, 4);
            char lastCharacter = chineseName[chineseName.Length - 1];
            if (lastCharacter>='A'&&lastCharacter<='Z')
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
            else if (ricEnglishInfo.BullBear == "Call")
            {
                colDsplyNmll += "CW";
            }
            else if (ricEnglishInfo.BullBear == "Put")
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
            idnLongName+=" ";
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
            string[] arr = {"",""};
            foreach (Trans tran in issuerCodeMap.Map)
            {
                if (issuerID == tran.Code)
                {
                    arr[0] = tran.ShortName;
                    arr[1] = tran.FullName;

                    break;
                }
                else
                {
                    continue;
                }
            }
            if (arr[0] == string.Empty||arr[1]==string.Empty)
            {
                ErrorMsg += "\n There's no issuer for ";
                ErrorMsg += issuerID;
            }
            return arr;
        }
    }

    public class RicInfo
    {
        //Info got from http://www.hkex.com.hk/eng/cbbc/newissue/newlaunch.htm
        public string Code { get; set; }
        public string Name { get; set; }
        public string Issuer { get; set; }
        public string Underlying { get; set; }
        public string BullBear { get; set; }
        public string BoardLot { get; set; }
        public string StrikeLevel { get; set; }
        public string CallLevel { get; set; }
        public string EntitlementRatio { get; set; }
        public string TotalIssueSize { get; set; }
        public string LauntchDate { get; set; }
        public string ClearingCommencementDate { get; set; }
        public string ListingDate { get; set; }
        public string MaturityDate { get; set; }

        public string Gear { get; set; }
        public string Premium { get; set; }

        //Info got from http://www.hkex.com.hk/chi/cbbc/newissue/newlaunch_c.htm
        public string ChineseName { get; set; }

        //Info got from http://www.hkex.com.hk/eng/cbbc/cbbcsummary.asp?id=?
        public string UnderlyingNameForStock{get;set;}
        public string IssuerPrice{get;set;}
    }

    public class HK_IssuerCodeMap
    {
        public List<Trans> Map { get; set; }
    }

    public class Trans
    {
        public string Code { get; set; }
        public string FullName { get; set; }
        public string ShortName { get; set; }
        public string WarrentIssuer { get; set; }
    }

    //Configuration for CBBC/Warrant and bulk file generate 
    [ConfigStoredInDB]
    public class HKFMAndBulkFileGeneratorConfig
    {
        //Properties for HKCBBCGenerator
        [StoreInDB]
        [Category("CBBC")]
        [DisplayName("Start position")]
        public string CbbcStartPos { get; set; }

        [StoreInDB]
        [Category("CBBC")]
        [DisplayName("End position")]
        public string CbbcEndPos { get; set; }

        //Properties for HKWarrantGenerator
        [StoreInDB]
        [Category("Warrant")]
        [DisplayName("Start position")]
        public string WarrantStartPos { get; set; }

        [StoreInDB]
        [Category("Warrant")]
        [DisplayName("End position")]
        public string WarrantEndPos { get; set; }

        [StoreInDB]
        [DisplayName("Output directory")]
        [DefaultValue("D:")]
        [Description("It's the output root path.")]
        public string OutPutDir { get; set; }
    }

    public class HKFMAndBulkFileGenerator : GeneratorBase
    {
        private HKFMAndBulkFileGeneratorConfig configObj = null;
       

        private HKCBBCFMSubGenerator cbbcGenerator = null;
        private HKWarrantFMSubGenerator warrantGenerator = null;
        private HKFMBulkFilesSubGenerator bulkFileGenerator = null;

        private readonly string pdfFileDir = "HKPDFFile";
        private readonly List<string> taskList = new List<string>() { "CBBCGenerator", "WarrantGenerator", "BulkFileGenerator" };
        private readonly bool isKeepPdf = true;

        public readonly string RicTemplatePath = "HKRicTemplate";
        public readonly string CbbcSubDir = "CBBC";
        public readonly string WarrantSubDir = "Warrant";
        public readonly string CbbcLogName = "CBBCLog.txt";
        public readonly string WarrantLogName = "WarrantLog.txt";
        public readonly string CbbcFmSn = "1000";
        public readonly string WarrantFmSn = "2000";
        public int CbbcStartPos
        {
            get;
            private set;
        }
        public int CbbcEndPos
        {
            get;
            private set;
        }
        public int WarrantStartPos
        {
            get;
            private set;
        }
        public int WarrantEndPos
        {
            get;
            private set;
        }

        public List<RicInfo> RicListCBBC = new List<RicInfo>();
        public List<RicInfo> ChineseListCBBC = new List<RicInfo>();

        public List<RicInfo> RicListWarrant = new List<RicInfo>();
        public List<RicInfo> ChineseListWarrant = new List<RicInfo>();

        protected override void Start()
        {
            StartFMAndBulkFileGeneratorJob();
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as HKFMAndBulkFileGeneratorConfig;
            try
            {
                CbbcStartPos = int.Parse(configObj.CbbcStartPos);
                CbbcEndPos = int.Parse(configObj.CbbcEndPos);
                WarrantStartPos = int.Parse(configObj.WarrantStartPos);
                WarrantEndPos = int.Parse(configObj.WarrantEndPos);
            }
            catch (Exception)
            {
                Logger.Log("The value of CBBC Positions or Warrant Positions must be a digital!");
            }

            cbbcGenerator = new HKCBBCFMSubGenerator(this);
            warrantGenerator = new HKWarrantFMSubGenerator(this);
            bulkFileGenerator = new HKFMBulkFilesSubGenerator(this);

            if (!Directory.Exists(configObj.OutPutDir))
            {
                Directory.CreateDirectory(configObj.OutPutDir);
            }

            cbbcGenerator.Initialize(configObj.OutPutDir, Logger, TaskResultList);
            warrantGenerator.Initialize(configObj.OutPutDir, Logger, TaskResultList);
            bulkFileGenerator.Initialize(configObj.OutPutDir, Logger, TaskResultList);

        }

        protected override void Cleanup()
        {
            
            base.Cleanup();
            cbbcGenerator.Cleanup();
            warrantGenerator.Cleanup();
            cbbcGenerator = null;
            warrantGenerator = null;
            bulkFileGenerator = null;

            if (isKeepPdf == false)
            {
                DeleteTempDir(configObj.OutPutDir + "\\" + pdfFileDir);
            }
        }    
        // * Delete local temp folder and all sub folders and files under it
        public void DeleteTempDir(String dir)
        {

            try
            {
                if (Directory.GetDirectories(dir).Length == 0 && Directory.GetFiles(dir).Length == 0)
                {
                    Directory.Delete(dir);
                    return;
                }
                foreach (string var in Directory.GetDirectories(dir))
                {
                    DeleteTempDir(var);
                }
                foreach (string var in Directory.GetFiles(dir))
                {

                    File.SetAttributes(var, FileAttributes.Normal);
                    File.Delete(var);
                }
                Directory.Delete(dir);
            }
            catch (Exception)
            { }
        }

        public void StartFMAndBulkFileGeneratorJob()
        {
            if (taskList == null || taskList.Count == 0)
            {
                Logger.LogErrorAndRaiseException("Please select at least one job: CBBCGenerator, WarrantGenerator or BulkFileGenerator");
            }

            for (int i = 0; i < taskList.Count; i++)
            {
                if (taskList[i].Contains("CBBC"))
                {
                    cbbcGenerator.Start();
                    RicListCBBC = cbbcGenerator.RicList;
                    ChineseListCBBC = cbbcGenerator.RicChineseList;
                }
                else if (taskList[i].Contains("Warrant"))
                {
                    warrantGenerator.Start();
                    RicListWarrant = warrantGenerator.RicList;
                    ChineseListWarrant = warrantGenerator.RicChineseList;
                }
                else
                {
                    bulkFileGenerator.RicListCbbc = RicListCBBC;
                    bulkFileGenerator.RicListWarrant = RicListWarrant;
                    bulkFileGenerator.ChineseListCbbc = ChineseListCBBC;
                    bulkFileGenerator.ChineseListWarrant = ChineseListWarrant;
                    bulkFileGenerator.Start();
                }
            }

            string date = DateTime.Now.ToString("yyyy_MMM_dd");
            HKRicNumInfo cbbcRicNumInfo = new HKRicNumInfo(date, RicListCBBC.Count, RicListWarrant.Count);            
            HKRicNumManager ricManager = new HKRicNumManager();
            bool isSuccess = false;
            if (ricManager.GetByDate(date) == null)
            {
                isSuccess = ricManager.Insert(cbbcRicNumInfo);
            }
            else
            {
                isSuccess = ricManager.ModifyByDate(date, RicListCBBC.Count, RicListWarrant.Count);
            }
        }

        public void UpdateDatabase()
        {
            HKRicNumInfo cbbcRicNumInfo = new HKRicNumInfo();
        }

        #region Get Information from PDF

        public void PDFAnalysis(RicInfo ric, FMType fmType)
        {
            string ricCode = ric.Code;
            int position = 0;
            string txtPath = GetPDFToTxtFilePath(ricCode,fmType);

            System.Threading.Thread.Sleep(3000);
            try
            {
                if (txtPath == null)
                {
                    throw new Exception();
                }

                String gearStr = "";
                String premiumStr = "";
                String lineText = "";
                StreamReader sr = new StreamReader(txtPath);
                lineText = sr.ReadToEnd();
                sr.Close();


                //Get position of stock
                int stockIndex = lineText.IndexOf("Stock code");
                int stockLength = "Stock code".Length;
                if (stockIndex < 0)
                {
                    stockIndex = lineText.IndexOf("Stock Code");
                }


                string codeStr = SearchKeyValue(stockIndex, stockLength, lineText);

                position = codeStr.IndexOf(ricCode) / 5 + 1;

                //Get gearing value
                int gearingIndex = lineText.IndexOf("\n Gearing* ");
                int gearLength = "\n Gearing* ".Length;
                if (gearingIndex < 0)
                {
                    gearingIndex = lineText.IndexOf("\nGearing*");
                    gearLength = "\nGearing*".Length;
                }
                if (gearingIndex < 0)
                {
                    gearingIndex = lineText.IndexOf("\n Gearing *");
                    gearLength = "\n Gearing *".Length;
                }
                if (gearingIndex < 0)
                {
                    gearingIndex = lineText.IndexOf("\nGearing     *");
                    gearLength = "\n Gearing     *".Length;
                }
                if (gearingIndex < 0)
                {
                    gearingIndex = lineText.IndexOf("\nGearing *");
                    gearLength = "\nGearing *".Length;
                }

                if (gearingIndex < 0)
                {
                    gearingIndex = lineText.IndexOf("Gearing*");
                    gearLength = "Gearing*".Length;
                }
                gearStr = SearchKeyValue(gearingIndex, gearLength, lineText);

                //Get premium value
                int premiumIndex = lineText.IndexOf("Premium*");
                int premiumLength = "Premium*".Length;
                if (premiumIndex < 0)
                {
                    premiumIndex = lineText.IndexOf("Premium *");
                    premiumLength = "Premium *".Length;
                }

                premiumStr = SearchKeyValue(premiumIndex, premiumLength, lineText);

                String[] gearArr = gearStr.Split('x');
                String[] premiumArr = premiumStr.Split('%');

                ric.Gear = gearArr[position - 1];
                ric.Premium = premiumArr[position - 1];

            }//end try
            catch (Exception)
            {
                Logger.Log("PDF analysis failed for " + ricCode + "! Action: Need manually input gearing and premium ", Logger.LogType.Warning);
            }
        }

        public String SearchKeyValue(int keyPosition, int keyLength, String sourceStr)
        {
            StringBuilder valueStr = new StringBuilder();
            Char[] sourceChar = sourceStr.ToCharArray();
            int position = keyPosition + keyLength;
            int index = position;

            while (sourceChar[position] != '\r')
            {
                if (sourceChar[position] != ' ')
                {
                    valueStr.Append(sourceChar[position]);
                }
                position++;
            }
            return valueStr.ToString();
        }

        public string GetPDFUrl(string id)
        {
            string pdfUrl = string.Empty;
            string postData = getPostData(id);
            string Uri = "http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.aspx";

            try
            {
                string pageSource = WebClientUtil.GetPageSource(Uri, 24000, postData);
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(pageSource);
                HtmlAgilityPack.HtmlNode pdfLinkNode = htmlDoc.DocumentNode.SelectSingleNode("//a[contains(@href, '.pdf')]");
                if (pdfLinkNode == null)
                {
                    Logger.Log("There's no PDF file for ric " + id);
                    return null;
                }
                else
                {
                    pdfUrl = "http://www.hkexnews.hk";
                    pdfUrl += pdfLinkNode.Attributes["href"].Value;
                }
            }
            catch (Exception ex)
            {
                string errInfo = ex.ToString();
            }
            return pdfUrl;
        }

        public string GetPDFToTxtFilePath(string codeId, FMType fmType)
        {
            string txtFilePath = string.Empty;
            string pdfUrl = GetPDFUrl(codeId);
            string pdfFilePath = string.Empty;
            if (pdfUrl == null)
            {
                return null;
            }
            if (fmType == FMType.Cbbc)
            {
                pdfFilePath = configObj.OutPutDir + "\\" + pdfFileDir + "\\" + "CBBC" + "\\" + codeId + ".pdf";
            }
            else
            {
                pdfFilePath = configObj.OutPutDir + "\\" + pdfFileDir + "\\" + "Warrant" + "\\" + codeId + ".pdf";
            }

            if (!Directory.Exists(Path.GetDirectoryName(pdfFilePath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(pdfFilePath));
            }
            WebClientUtil.DownloadFile(pdfUrl, 24000, pdfFilePath);

            txtFilePath = PDFToTxt(codeId, pdfFilePath,fmType);
            return txtFilePath;
        }

        //Transfer PDF to TXT file
        public String PDFToTxt(String ricCode, String pdfFilePath, FMType fmType)
        {
            string command = "pdftotext.exe";
            string txtPath = string.Empty;
            if (fmType == FMType.Cbbc)
            {

                txtPath = configObj.OutPutDir + "\\" + pdfFileDir + "\\" + "CBBC" + "\\" + ricCode + ".txt";
            }
            else
            {
                txtPath = configObj.OutPutDir + "\\" + pdfFileDir + "\\" + "Warrant" + "\\" + ricCode + ".txt";
            }

            if (!Directory.Exists(Path.GetDirectoryName(txtPath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(txtPath));
            }
            string parameters = "-layout -enc UTF-8 -q " + pdfFilePath + " " + txtPath;
            try
            {
                using (ProcessContext p = new ProcessContext(command, parameters, "."))
                {
                    p.ProcessInstance.Start();
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorAndRaiseException("Error accured when transferring PDF to txt. " + ex.Message);
            }

            return txtPath;
        }
        #endregion

        private string getPostData(string id)
        {
            DateTime today = DateTime.Now;
            DateTime yesterday = today.AddDays(-3);
            string postDataTemplate = "txt_stock_code={0}&sel_DateOfReleaseFrom_y={1}&sel_DateOfReleaseFrom_m={2}&sel_DateOfReleaseFrom_d={3}&sel_DateOfReleaseTo_y={4}&sel_DateOfReleaseTo_m={5}&sel_DateOfReleaseTo_d={6}&sel_tier_1=-2&sel_tier_2_group=-2&sel_tier_2=-2&IsFromNewList=False";
            string postData =  string.Format(postDataTemplate,
              id,yesterday.Year,yesterday.ToString("MM"),yesterday.ToString("dd"),today.Year,today.ToString("MM"),today.ToString("dd"));
            return postData;
        }

        //Update FMSerialNumber
        public String UpdateFMSerialNumber(String fmSerialNumber)
        {
            int number = 0;
            if (fmSerialNumber.Substring(0, 1) != "0")
            {
                fmSerialNumber = (Convert.ToInt32(fmSerialNumber) + 1).ToString();
            }
            else if (fmSerialNumber.Substring(1, 1) != "0")
            {
                number = Convert.ToInt32(fmSerialNumber.Substring(1));
                if (number == 999)
                {
                    fmSerialNumber = (number + 1).ToString();
                }
                else
                {
                    fmSerialNumber = "0" + (number + 1).ToString();
                }
            }
            else if (fmSerialNumber.Substring(2, 1) != "0")
            {
                number = Convert.ToInt32(fmSerialNumber.Substring(2));
                if (number == 99)
                {
                    fmSerialNumber = "0" + (number + 1).ToString();
                }
                else
                {
                    fmSerialNumber = "00" + (number + 1).ToString();
                }
            }
            else
            {
                number = Convert.ToInt32(fmSerialNumber.Substring(3));
                if (number == 9)
                {
                    fmSerialNumber = "00" + (number + 1).ToString();
                }
                else
                {
                    fmSerialNumber = "000" + (number + 1).ToString();
                }
            }
            return fmSerialNumber;
        }//end UpdateFMSerialNumber
    }
}
