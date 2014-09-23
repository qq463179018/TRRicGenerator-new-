﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Design;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks
{
    #region Entity class being used for deliver data in the application

    public class RICUpgrade
    {
        public string Effective_date { get; set; }
        public string RIC_Old { get; set; }
        public string RIC_New { get; set; }
        public string TAG_Old { get; set; }
        public string TAG_New { get; set; }
        public string Currency { get; set; }
        public string QA_Common_Name { get; set; }
        public string QA_Short_Name { get; set; }
        public string IA_Common_Name { get; set; }
        public string Korean_Code { get; set; }
        public string ISIN { get; set; }
        public string Country_Headquarters { get; set; }
        public string Legal_Name { get; set; }
        public string Korean_Name { get; set; }
        public string Edcoid_Old { get; set; }
        public string Edcoid_New { get; set; }
        public string Old_MSCI { get; set; }
        public string FTSE { get; set; }
        public string Korean_Scheme { get; set; }
        public string Record_Type { get; set; }
        public string KOSDAQ_Chians_Old { get; set; }
        public string KSE_Chians_New { get; set; }
        public string Issue_Classfication { get; set; }
        public string Lot_Size_Old { get; set; }
        public string Lot_Size_New { get; set; }
    }

    public class NameChangeTemplate
    {
        public string UpdateDate { get; set; }
        public string EffectiveDate { get; set; }
        public string RIC { get; set; }
        public string OldRIC { get; set; }
        public string NewRIC { get; set; }
        public string ISIN { get; set; }
        public string OldISIN { get; set; }
        public string NewISIN { get; set; }
        public string Ticker { get; set; }
        public string OldTicker { get; set; }
        public string NewTicker { get; set; }
        public string OldLegalName { get; set; }
        public string NewLegalName { get; set; }
        public string OldDisplayName { get; set; }
        public string NewDisplayName { get; set; }
        public string KoreaName { get; set; }
        public bool isRevised { get; set; }
    }   

    public class CompanyWarrantDropTemplate
    {
        public string EffectiveDate { get; set; }
        public string RIC { get; set; }
        public string ISIN { get; set; }
        public string QACommonName { get; set; }
        public string QAShortName { get; set; }
        public string ForIACommonName { get; set; }
        public string LegalName { get; set; }
        public string KoreanName { get; set; }
    }

    public class DropTemplate
    {
        public string EffectiveDate { get; set; }
        public string RIC { get; set; }
        public string ISIN { get; set; }
        public string QAShortName { get; set; }
        public string LegalName { get; set; }
        //search key word
        public string KoreaName { get; set; }
        public string Market { get; set; }
        public string Type { get; set; }
        public bool isRevised { get; set; }
        public string AnnouncementTime { get; set; }

        public DropTemplate()
        {
            EffectiveDate = string.Empty;
            isRevised = false;
            Market = string.Empty;
        }
    }

    public class SPCRAdjustmentTemplate
    {
        public string UpdateDate { get; set; }
        public string KoreaName { get; set; }
        public string EffectiveDate { get; set; }
        public string RIC { get; set; }
        public string ISIN { get; set; }
        public string StrikePrice { get; set; }
        public string ConversionRatio { get; set; }
        public string QACommonName { get; set; }
    }

    public class CompanyWarrantList
    {
        public string Ric { get; set; }
        public string Display_Name { get; set; }
        public string ISIN { get; set; }
        public string Conversion_ratio { get; set; }
        public string Exercise_Price { get; set; }
    }

    public class KSorKQListingList
    {
        public string Ric { get; set; }
        public string ISIN { get; set; }
        public string IDNDisplayName { get; set; }
    }

    public class ETFListingList
    {
        public string RIC { get; set; }
        public string IDNDisplayName { get; set; }
        public string ISIN { get; set; }
    }

    public class FurtherIssueModel
    {
        public string Updated_Date { get; set; }
        public string Effective_Date { get; set; }
        public string Old_Ric { get; set; }
        public string New_Ric { get; set; }
        public string Old_Isin { get; set; }
        public string New_Isin { get; set; }
        public string Old_Ticker { get; set; }
        public string New_Ticker { get; set; }
        public string Old_Quanity { get; set; }
        public string New_Quanity { get; set; }
    }

    #endregion

    #region KOREA_ELWFMGeneratorConfig used for ELW instrument
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaElwfmReadFilePathConfig
    {
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaElwfm1GenerateFileConfig
    {
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public int DelayTime { get; set; }
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaElwfm2GenerateFileConfig
    {
        public string PdfPath { get; set; }
        public string TextPath { get; set; }
        public string SplitDate { get; set; }
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaElwfmDropGenerateFileConfig
    {
        public DateTime StartDate { get; set; }
        public string WorkbookPath { get; set; }
        public string WorksheetNaem { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaElwfmSingleSearchGenerateFileConfig
    {
        public string InputISIN { get; set; }
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaElwfmFurtherIssuerGenerateFileConfig
    {
        public string SplitDate { get; set; }
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaElwfmGeneratorConfig
    {
        public KoreaElwfmReadFilePathConfig KoreaElwfmReadFilePathConfig { get; set; }
        public KoreaElwfm1GenerateFileConfig KoreaElwfm1GenerateFileConfig { get; set; }
        public KoreaElwfm2GenerateFileConfig KoreaElwfm2GenerateFileConfig { get; set; }
        public KoreaElwfmDropGenerateFileConfig KoreaElwfmDropGenerateFileConfig { get; set; }
        public KoreaElwfmSingleSearchGenerateFileConfig KoreaElwfmSingleSearchGenerateFileConfig { get; set; }
        public KoreaElwfmFurtherIssuerGenerateFileConfig KoreaElwfmFurtherIssuerGenerateFileConfig { get; set; }
    }
    #endregion

    #region KOREA_EquityGeneratorConfig used for Equity and ETF 、REITs instrument
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaEquityGeneratorConfig
    {
        public KoreaEquityReadFileConfig KoreaEquityReadFileConfig { get; set; }
        public KoreaKQorKsListReadFilePathConfig KoreaKQorKsListReadFilePathConfig { get; set; }
        public KoreaPeoGeneratorFileConfig KoreaPeoGeneratorFileConfig { get; set; }
        public KoreaPeoAddRicUpgradeGeneratorFileConfig KoreaPeoAddRicUpgradeGeneratorFileConfig { get; set; }
        public KoreaReitGeneratorFileConfig KoreaReitGeneratorFileConfig { get; set; }
        public KoreaEtfGeneratorFileConfig KoreaEtfGeneratorFileConfig { get; set; }
        public KoreaNameChangeGeneratorFileConfig KoreaNameChangeGeneratorFileConfig { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaEquityReadFileConfig
    {
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaKQorKsListReadFilePathConfig
    {
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaPeoGeneratorFileConfig
    {
        public string StartDate { get; set; }
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaPeoAddRicUpgradeGeneratorFileConfig
    {
        public string StartDate { get; set; }
        public string PeoAddWorkbookPath { get; set; }
        public string PeoAddWorksheetNaem { get; set; }
        public string RicUpgradeWorkbookPath { get; set; }
        public string RicUpgradeWorksheetNaem { get; set; }
    }


    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaReitGeneratorFileConfig
    {
        public string StartDate { get; set; }
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaEtfGeneratorFileConfig
    {
        public string StartDate { get; set; }
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaNameChangeGeneratorFileConfig
    {
        public string StartDate { get; set; }
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }
    #endregion

    #region KOREA_DROP_Config used for Company Warrant Add、Change、Drop and <EquityDrop 、BC Drop  (include in class Drop )>

    [TypeConverter(typeof(ExpandableObjectConverter))]
    [ConfigStoredInDB]
    public class KoreaDropConfig
    {
        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string StartDate { get; set; }

        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\Drop\\GEDA\\")]
        [Description("Path for saving generated GEDA files (Name_Change)\nE.g. C:\\Korea_Auto\\Equity\\Name_Change\\GEDA\\ ")]
        public string GEDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\Drop\\NDA\\")]
        [Description("Path for saving generated NDA files (Name_Change)\nE.g. C:\\Korea_Auto\\Equity\\Name_Change\\NDA\\ ")]
        public string NDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\Drop\\FM\\")]
        [Description("Path for saving generated FM files (Name_Change)\nE.g. C:\\Korea_Auto\\Equity\\Name_Change\\FM\\ ")]
        public string FM { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }

        public KoreaDropConfig()
        {
            StartDate = DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd");
            EndDate = DateTime.Today.ToString("yyyy-MM-dd");
        }

    }


    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_DROP_Config
    {
        public string KoreaDropStartDate { get; set; }
        public string KoreaDropEndDate { get; set; }
        public KoreaDropCompanyWarrantReadFileConfig KoreaDropCompanyWarrantReadFileConfig { get; set; }
        public KoreaDropKQorKsListReadFilePathConfig KoreaDropKQorKsListReadFilePathConfig { get; set; }
        public Korea_CompanyWarrant_DropGenerator_CONFIG KoreaCompanyWarrantDropGeneratorConfig { get; set; }
        public Korea_Equity_DropGenerator_CONFIG KoreaEquityDropGeneratorConfig { get; set; }
        public Korea_BC_DropGenerator_CONFIG KoreaBcDropGeneratorConfig { get; set; }
        public KoreaDropEtfListingItemsReadFilePathConfig KoreaDropEtfListingItemsReadFilePathConfig { get; set; }
        public Korea_ETF_DropGenerator_CONFIG KoreaEtfDropGeneratorConfig { get; set; }

        public List<string> AlertMailToList { get; set; }

        public List<string> AlertMailCCList { get; set; }

        public List<string> AlertMailSignatureInformationList { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaDropCompanyWarrantReadFileConfig
    {
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaDropEtfListingItemsReadFilePathConfig
    {
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaDropKQorKsListReadFilePathConfig
    {
        public string WorkbookPath { get; set; }
        public string WorksheetName { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_CompanyWarrant_DropGenerator_CONFIG
    {
        public string startDate { get; set; }
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
        public string NDA_FILE_PATH { get; set; }
        public string GEDA_FILE_PATH { get; set; }

    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_Equity_DropGenerator_CONFIG
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_ETF_DropGenerator_CONFIG
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_BC_DropGenerator_CONFIG
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    #endregion

    #region  KOREA_PEOGeneratorConfig for PEO

    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_PEOGeneratorConfig
    {
        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\Pre_IPO\\GEDA\\")]
        [Description("Path for saving generated GEDA files (Name_Change)\nE.g. C:\\Korea_Auto\\Equity\\Name_Change\\GEDA\\ ")]
        public string GEDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\Pre_IPO\\NDA\\")]
        [Description("Path for saving generated NDA files (Name_Change)\nE.g. C:\\Korea_Auto\\Equity\\Name_Change\\NDA\\ ")]
        public string NDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\Pre_IPO\\FM\\")]
        [Description("Path for saving generated FM files (Name_Change)\nE.g. C:\\Korea_Auto\\Equity\\Name_Change\\FM\\ ")]
        public string FM { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }
    }

    [Description("")]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_ReadDataFromEquityMasterfileConfig
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    #endregion

    #region  KOREA_GenerateGEDAFile
    public class GEDAFileTemplate
    {
        public string Geda_Symbol { get; set; }
        public string Geda_Display_Name { get; set; }
        public string Geda_Ric { get; set; }
        public string Geda_Official_Code { get; set; }
        public string Geda_Ex_Symbol { get; set; }
        public string Geda_Background_Page { get; set; }
        public string Geda_Lot_Size { get; set; }
        public string Geda_Display_Nmll { get; set; }
        public string Geda_Bcast_Ref { get; set; }
        public string Geda_Instmod_Lot_Size { get; set; }
        public string Geda_Instmod_Mnemonic { get; set; }
        public string Geda_Instmod_Tdn_Symbol { get; set; }
        public string Geda_Exl_Name { get; set; }
        public string Geda_Instmod_Dds_Lot_Size { get; set; }
        public string Geda_Instmod_Tdn_Issuer_Name { get; set; }
        public string Geda_Instmod_ISIN { get; set; }
        public string Geda_BCU { get; set; }

        public GEDAFileTemplate()
        {
            Geda_Symbol = "SYMBOL";
            Geda_Display_Name = "DSPLY_NAME";
            Geda_Ric = "RIC";
            Geda_Official_Code = "OFFCL_CODE";
            Geda_Ex_Symbol = "EX_SYMBOL";
            Geda_Background_Page = "BCKGRNDPAG";
            Geda_Lot_Size = "LOT_SIZE_A";
            Geda_Display_Nmll = "DSPLY_NMLL";
            Geda_Bcast_Ref = "BCAST_REF";
            Geda_Instmod_Lot_Size = "#INSTMOD_LOT_SIZE_X";
            Geda_Instmod_Mnemonic = "#INSTMOD_MNEMONIC";
            Geda_Instmod_Tdn_Symbol = "#INSTMOD_TDN_SYMBOL";
            Geda_Exl_Name = "EXL_NAME";
            Geda_Instmod_Dds_Lot_Size = "#INSTMOD_#DDS_LOT_SIZE";
            Geda_Instmod_Tdn_Issuer_Name = "#INSTMOD_TDN_ISSUER_NAME";
            Geda_Instmod_ISIN = "#INSTMOD_#ISIN";
            Geda_BCU = "BCU";
        }
    }
    #endregion

    #region KOREA_GenerateNDAFile
    public class NDAFileTemplate
    {
        public string Nda_Ric { get; set; }
        public string Nda_Tag { get; set; }
        public string Nda_Base_Asset { get; set; }
        public string Nda_Ticker_Symbol { get; set; }
        public string Nda_Asset_Short_Name { get; set; }
        public string Nda_Asset_Common_Name { get; set; }
        public string Nda_Type { get; set; }
        public string Nda_Category { get; set; }
        public string Nda_Currency { get; set; }
        public string Nda_Exchange { get; set; }
        public string Nda_Equity_First_Trading_Day { get; set; }
        public string Nda_Round_Lot_Size { get; set; }

        public NDAFileTemplate()
        {
            Nda_Ric = "RIC";
            Nda_Tag = "TAG";
            Nda_Base_Asset = "BASE ASSET";
            Nda_Ticker_Symbol = "TICKER SYMBOL";
            Nda_Asset_Short_Name = "ASSET SHORT NAME";
            Nda_Asset_Common_Name = "ASSET COMMON NAME";
            Nda_Type = "TYPE";
            Nda_Category = "CATEGORY";
            Nda_Currency = "CURRENCY";
            Nda_Exchange = "EXCHANGE";
            Nda_Equity_First_Trading_Day = "EQUITY FIRST TRADING DAY";
            Nda_Round_Lot_Size = "ROUND LOT SIZE";
        }
    }
    #endregion

    #region  KOREA_ADDGeneratorConfig for ADD (include the PEO ADD , ETF ADD and REITs ADD)

    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_ADDGeneratorConfig
    {
        [Category("1.Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string StartDate { get; set; }

        [Category("1.Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("2.PEO")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\PEO_ADD\\FM\\")]
        [Description("Path for saving generated FM files\nE.g. C:\\Korea_Auto\\Equity\\PEO_ADD\\FM\\ ")]
        public string PEOFM { get; set; }

        [StoreInDB]
        [Category("2.PEO")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\PEO_ADD\\GEDA\\")]
        [Description("Path for saving generated GEDA files\nE.g. C:\\Korea_Auto\\Equity\\PEO_ADD\\GEDA\\ ")]
        public string PEOGEDA { get; set; }

        [StoreInDB]
        [Category("2.PEO")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\PEO_ADD\\NDA\\")]
        [Description("Path for saving generated NDA files\nE.g. C:\\Korea_Auto\\Equity\\PEO_ADD\\NDA\\ ")]
        public string PEONDA { get; set; }

        [StoreInDB]
        [Category("4.REIT")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\REIT\\FM\\")]
        [Description("Path for saving generated FM files\nE.g. C:\\Korea_Auto\\Equity\\REIT\\FM\\ ")]
        public string REITFM { get; set; }

        [StoreInDB]
        [Category("4.REIT")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\REIT\\GEDA\\")]
        [Description("Path for saving generated GEDA files\nE.g. C:\\Korea_Auto\\Equity\\REIT\\GEDA\\ ")]
        public string REITGEDA { get; set; }

        [StoreInDB]
        [Category("4.REIT")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\REIT\\NDA\\")]
        [Description("Path for saving generated NDA files\nE.g. C:\\Korea_Auto\\Equity\\REIT\\NDA\\ ")]
        public string REITNDA { get; set; }

        [StoreInDB]
        [Category("3.ETF")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\ETF\\FM\\")]
        [Description("Path for saving generated FM files\nE.g. C:\\Korea_Auto\\Equity\\ETF\\FM\\ ")]
        public string ETFFM { get; set; }

        [StoreInDB]
        [Category("3.ETF")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\ETF\\GEDA\\")]
        [Description("Path for saving generated GEDA files\nE.g. C:\\Korea_Auto\\Equity\\ETF\\GEDA\\ ")]
        public string ETFGEDA { get; set; }

        [StoreInDB]
        [Category("3.ETF")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\ETF\\NDA\\")]
        [Description("Path for saving generated NDA files\nE.g. C:\\Korea_Auto\\Equity\\ETF\\NDA\\ ")]
        public string ETFNDA { get; set; }

        [StoreInDB]
        [Category("5.BC")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\BC\\FM\\")]
        [Description("Path for saving generated FM files\nE.g. C:\\Korea_Auto\\Equity\\BC\\FM\\ ")]
        public string BCFM { get; set; }

        [StoreInDB]
        [Category("5.BC")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\BC\\GEDA\\")]
        [Description("Path for saving generated GEDA files\nE.g. C:\\Korea_Auto\\Equity\\BC\\GEDA\\ ")]
        public string BCGEDA { get; set; }

        [StoreInDB]
        [Category("5.BC")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\BC\\NDA\\")]
        [Description("Path for saving generated NDA files\nE.g. C:\\Korea_Auto\\Equity\\BC\\NDA\\ ")]
        public string BCNDA { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("6.Mail")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("6.Mail")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("6.Mail")]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }

        //KSKQ Listing file . Need to revise.
        //public Korea_ReadDataFromKSKQListingListConfig Korea_ReadDataFromKSKQListingListConfig { get; set; }
        //public Korea_RICUpgrade_GeneratorFileConfig Korea_RICUpgrade_GeneratorFileConfig { get; set; }

        public KOREA_ADDGeneratorConfig()
        {
            StartDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            EndDate = DateTime.Now.ToString("yyyy-MM-dd");
        }

    }

    /// <summary>
    /// this will be used for all which need to read KS or KQ Listing items
    /// </summary>    
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_ReadDataFromKSKQListingListConfig
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_RICUpgrade_GeneratorFileConfig
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }



    #endregion

    #region  KOREA_ELWStrikeConversionGeneratorConfig be used for ELW Strike price and Conversion ratio Adjustment
    public class KOREA_ELWStrikeConversionGeneratorConfig
    {
        public string Korea_SPCRAdjustment_StartDate { get; set; }
        public string Korea_SPCRAdjustment_EndDate { get; set; }

        public Korea_SPCRAdjustment_GeneratorFileConfig Korea_SPCRAdjustment_GeneratorFileConfig { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_SPCRAdjustment_GeneratorFileConfig
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }
    #endregion

    #region KOREA_NameChangeGeneratorConfig be used for Name Change

    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_NameChangeGeneratorConfig
    {
        [Category("Date")]
        [DisplayName("Start date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string StartDate { get; set; }

        [Category("Date")]
        [DisplayName("End date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\Name_Change\\GEDA\\")]
        [Description("Path for saving generated GEDA files (Name_Change)\nE.g. C:\\Korea_Auto\\Equity\\Name_Change\\GEDA\\ ")]
        public string GEDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\Name_Change\\NDA\\")]
        [Description("Path for saving generated NDA files (Name_Change)\nE.g. C:\\Korea_Auto\\Equity\\Name_Change\\NDA\\ ")]
        public string NDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Equity\\Name_Change\\FM\\")]
        [Description("Path for saving generated FM files (Name_Change)\nE.g. C:\\Korea_Auto\\Equity\\Name_Change\\FM\\ ")]
        public string FM { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [DisplayName("Recipients")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [DisplayName("Recipients (Cc)")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [DisplayName("Signature")]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }

        public KOREA_NameChangeGeneratorConfig()
        {
            StartDate = DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd");
            EndDate = DateTime.Today.ToString("yyyy-MM-dd");
        }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_NameChange_GenerateFileConfig
    {

        public string WORKSHEET_NAME { get; set; }
    }


    #endregion

    #region  KOREA_ELWFM1GeneratorConfig be used for ELW FM1

    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig
    {
        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string StartDate { get; set; }
        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("ELW")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\ELW_FM1\\FM\\")]
        [Description("Path for saving generated FM files.(ELW FM1)\nE.g. C:\\Korea_Auto\\ELW\\ELW_FM1\\FM\\ ")]
        public string FM { get; set; }

        [StoreInDB]
        [Category("ELW")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\ELW_FM1\\BulkFile\\")]
        [Description("Path for saving generated GEDA and NDA files.(ELW FM1)\nE.g. C:\\Korea_Auto\\ELW\\ELW_FM1\\BulkFile\\ ")]
        public string BulkFile { get; set; }

        [StoreInDB]
        [Category("ELW")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\TAG and PILC.xls")]
        [Description("Full path of TAG and PILC.xls. \nE.g. C:\\Korea_Auto\\ELW\\TAG and PILC.xls ")]
        public string TagPilcFile { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Signature for E-mail")]
        public List<string> Signature { get; set; }

        [StoreInDB]
        [Category("New Underlying")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\New_URIC\\")]
        [Description("Path for saving GEDA files of new underlying.\nE.g.C:\\Korea_Auto\\ELW_FM\\New_URIC\\)")]
        public string GEDA_NewUnderlying { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("New Underlying")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> NewUnderlying_MailCC { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("New Underlying")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> NewUnderlying_MailTo { get; set; }

        [StoreInDB]
        [Category("New Underlying")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Signature for E-mail")]
        public List<string> NewUnderlying_Signature { get; set; }


        public KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig()
        {
            StartDate = MiscUtil.GetLastBusinessDay(0, DateTime.Today).ToString("yyyy-MM-dd");
            EndDate = MiscUtil.GetLastBusinessDay(0, DateTime.Today).ToString("yyyy-MM-dd");
        }
    }

    #endregion

    #region KOREA_ELWFM2AndFurtherIssuerGeneratorConfig be used for ELW FM2 and ELW Further Issuer

    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREA_ELWFM2AndFurtherIssuerGeneratorConfig
    {

        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string StartDate { get; set; }

        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("ELW And KOBA")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\ELW_FM2\\FM\\")]
        [Description("Path for saving generated FM files.(ELW FM2)\nE.g. C:\\Korea_Auto\\ELW\\ELW_FM2\\FM\\ ")]
        public string FM { get; set; }

        [StoreInDB]
        [Category("ELW And KOBA")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\ELW_FM2\\BulkFile\\")]
        [Description("Path for saving generated GEDA and NDA files.(ELW FM2)\nE.g. C:\\Korea_Auto\\ELW\\ELW_FM2\\BulkFile\\ ")]
        public string BulkFile { get; set; }

        [StoreInDB]
        [Category("ELW And KOBA")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\TAG and PILC.xls")]
        [Description("Full path of TAG and PILC.xls. \nE.g. C:\\Korea_Auto\\ELW\\TAG and PILC.xls ")]
        public string TagPilcFile { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("ELW And KOBA")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("ELW And KOBA")]
        [Description("Signature for E-mail")]
        public List<string> Signature { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("ELW And KOBA")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; } 

        [StoreInDB]
        [Category("Further Issuer")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\ELW_FM2\\Further Issuer\\FM\\")]
        [Description("Path for saving generated FM files.(Further Issuer)\nE.g. C:\\Korea_Auto\\ELW\\ELW_FM2\\Further Issuer\\FM\\ ")] 
        public string FM_FurtherIssuer { get; set; }    

        [StoreInDB]
        [GroupValue]
        [Category("Further Issuer")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> FurtherIssuer_MailCC { get; set; }

        [StoreInDB]
        [Category("Further Issuer")]
        [Description("Signature for E-mail")]
        public List<string> FurtherIssuer_Signature { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Further Issuer")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> FurtherIssuer_MailTo { get; set; }

        [StoreInDB]
        [Category("New Underlying")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\New_URIC\\")]
        [Description("Path for saving GEDA files of new underlying.\nE.g.C:\\Korea_Auto\\ELW_FM\\New_URIC\\)")]
        public string GEDA_NewUnderlying { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("New Underlying")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> NewUnderlying_MailCC { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("New Underlying")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> NewUnderlying_MailTo { get; set; }

        [StoreInDB]
        [Category("New Underlying")]
        [Description("Signature for E-mail")]
        public List<string> NewUnderlying_Signature { get; set; }
                
        [Category("Run from pdf")]
        [DefaultValue("0")]
        [Description("If you want to generate output files from pdf(s) you give. Please choose an announcement type.")]
        public ElwAnnounceType AnnouncementType { get; set; }

        [StoreInDB]
        [Category("Run from pdf")]
        [DefaultValue("C:\\Korea_Auto\\ELW_FM\\ELW_FM2\\PDF\\")]
        [Description("Folder for the given pdf(s). \nE.g. C:\\Korea_Auto\\ELW_FM\\ELW_FM2\\PDF\\ ")]
        public string PdfPath { get; set; }

        public KOREA_ELWFM2AndFurtherIssuerGeneratorConfig()
        {
            StartDate = DateTime.Now.ToString("yyyy-MM-dd");
            EndDate = DateTime.Now.ToString("yyyy-MM-dd");
        }      
    }

    public enum ElwAnnounceType : int
    {
        None = 0,
        FM2_ELW = 1,
        FurtherIssuer = 2
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_FM2_ReadDataFromMasterfileConfig
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_FM2_GenerateFileConfig
    {
        public string FM { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_FurtherIssuer_GenerateFileConfig
    {
        public string FurtherFM { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_FM2_AppendDataToKOBAFileConfig
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_FM2_GenerateKOBAFileConfig
    {

        public string WORKSHEET_NAME { get; set; }
    }

    #endregion

    #region KOREA CompanyWarrant Config be used for Company Warrant Add/Change/Drop

    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREACompanyWarrantAddGeneratorConfig
    {
        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string StartDate { get; set; }

        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Company_Warrant\\Warrant_ADD\\GEDA\\")]
        [Description("Path for saving generated GEDA files (Company Warrant Add)\nE.g. C:\\Korea_Auto\\Company_Warrant\\Warrant_ADD\\GEDA\\ ")]
        public string GEDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Company_Warrant\\Warrant_ADD\\NDA\\")]
        [Description("Path for saving generated NDA files (Company Warrant Add)\nE.g. C:\\Korea_Auto\\Company_Warrant\\Warrant_ADD\\NDA\\ ")]
        public string NDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Company_Warrant\\Warrant_ADD\\FM\\")]
        [Description("Path for saving generated FM files (Company Warrant Add)\nE.g. C:\\Korea_Auto\\Company_Warrant\\Warrant_ADD\\FM\\ ")]
        public string FM { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }

        public KOREACompanyWarrantAddGeneratorConfig()
        {
            StartDate = DateTime.Now.ToString("yyyy-MM-dd");
            EndDate = DateTime.Now.ToString("yyyy-MM-dd");
        }
    }

    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREACompanyWarrantChangeGeneratorConfig
    {
        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string StartDate { get; set; }

        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Company_Warrant\\Warrant_CHANGE\\FM\\")]
        [Description("Path for saving generated FM files (Company Warrant Change)\nE.g. C:\\Korea_Auto\\Company_Warrant\\Warrant_CHANGE\\FM\\ ")]
        public string FM { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Company_Warrant\\Warrant_CHANGE\\GEDA\\")]
        [Description("Path for saving generated GEDA files (Company Warrant Change)\nE.g. C:\\Korea_Auto\\Company_Warrant\\Warrant_CHANGE\\GEDA\\ ")]
        public string GEDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Company_Warrant\\Warrant_CHANGE\\NDA\\")]
        [Description("Path for saving generated NDA files (Company Warrant Change)\nE.g. C:\\Korea_Auto\\Company_Warrant\\Warrant_CHANGE\\NDA\\ ")]
        public string NDA { get; set; }

        [StoreInDB]
        //[GroupValue]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        //[GroupValue]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }

        public KOREACompanyWarrantChangeGeneratorConfig()
        {
            StartDate = DateTime.Now.ToString("yyyy-MM-dd");
            EndDate = DateTime.Now.ToString("yyyy-MM-dd");
        }
    }


    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KoreaCompanyWarrantDropGeneratorConfig
    {
        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string StartDate { get; set; }

        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Company_Warrant\\Warrant_DROP\\FM\\")]
        [Description("Path for saving generated FM files (Company Warrant Change)\nE.g. C:\\Korea_Auto\\Company_Warrant\\Warrant_CHANGE\\FM\\ ")]
        public string FM { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Company_Warrant\\Warrant_DROP\\GEDA\\")]
        [Description("Path for saving generated GEDA files (Company Warrant Change)\nE.g. C:\\Korea_Auto\\Company_Warrant\\Warrant_CHANGE\\GEDA\\ ")]
        public string GEDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Company_Warrant\\Warrant_DROP\\NDA\\")]
        [Description("Path for saving generated NDA files (Company Warrant Change)\nE.g. C:\\Korea_Auto\\Company_Warrant\\Warrant_CHANGE\\NDA\\ ")]
        public string NDA { get; set; }

        [StoreInDB]
        //[GroupValue]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        //[GroupValue]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }

        public KoreaCompanyWarrantDropGeneratorConfig()
        {
            StartDate = DateTime.Now.ToString("yyyy-MM-dd");
            EndDate = DateTime.Now.ToString("yyyy-MM-dd");
        }

    }
    #endregion

    #region KOREA_RightsGeneratorConfig be used for Rights Add & Drop
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class KOREARightsGeneratorConfig
    {
        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string StartDate { get; set; }

        [Category("Date")]
        [Description("Announcement pulished Date\nDate Format: yyyy-mm-dd\nE.g. 2000-01-01")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Rights\\Add\\FM\\")]
        [Description("Path for saving generated FM files (Rights Add)\nE.g. C:\\Korea_Auto\\Rights\\Add\\FM\\ ")]
        public string RightsAddFM { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Rights\\Add\\GEDA\\")]
        [Description("Path for saving generated GEDA files (Rights Add)\nE.g. C:\\Korea_Auto\\Rights\\Add\\GEDA\\ ")]
        public string RightsAddGEDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Rights\\Add\\NDA\\")]
        [Description("Path for saving generated NDA files (Rights Add)\nE.g. C:\\Korea_Auto\\Rights\\Add\\NDA\\ ")]
        public string RightsAddNDA { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Rights\\Drop\\FM\\")]
        [Description("Path for saving generated FM files (Rights Drop)\nE.g. C:\\Korea_Auto\\Rights\\Drop\\FM\\ ")]
        public string RightsDropFM { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Rights\\Drop\\GEDA\\")]
        [Description("Path for saving generated GEDA files (Rights Drop)\nE.g. C:\\Korea_Auto\\Rights\\Drop\\GEDA\\ ")]
        public string RightsDropGEDA { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [GroupValue]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }

        public KOREARightsGeneratorConfig()
        {
            StartDate = DateTime.Now.ToString("yyyy-MM-dd");
            EndDate = DateTime.Now.ToString("yyyy-MM-dd");
        }

    }
    #endregion

    #region KOREA_ELWSearchByISINGeneratorConfig be used for ELW FM Search By ISIN
    public class KOREA_ELWSearchByISINGeneratorConfig
    {
        public List<String> ISINList { get; set; }

        public Korea_SearchByISIN_GenerateFileConfig Korea_SearchByISIN_GenerateFileConfig { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Korea_SearchByISIN_GenerateFileConfig
    {
        public string WORKBOOK_PATH { get; set; }
        public string WORKSHEET_NAME { get; set; }
    }

    #endregion

}
