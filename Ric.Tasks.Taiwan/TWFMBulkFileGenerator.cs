using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks.Taiwan
{
    [ConfigStoredInDB]
    public class TWFMBulkFileGeneratorConfig
    {
        [StoreInDB]
        [DisplayName("Fm file path")]
        [DefaultValue("D:\\TW\\TW+WARRANT+ADD+FM.xls")]
        [Description("It's the FM file path. The format should be like \"D:\\TW\\TW+WARRANT+ADD+FM.xls\", this field must be set.")]
        public string AddFmFilePath { get; set; }
    }

    public class TWFMBulkFileGenerator : GeneratorBase
    {
        private TWFMBulkFileGeneratorConfig configObj;
        private string outputPath = "";

        protected override void Initialize()
        {
            base.Initialize();
            try
            {
                configObj = Config as TWFMBulkFileGeneratorConfig;
                outputPath = GetOutputFilePath();
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when initializing task... Ex: {0} .", ex.Message));
            }
        }

        protected override void Start()
        {
            try
            {
                StartFMBulkFileGenertor();
            }

            catch (Exception ex)
            {
                Logger.LogErrorAndRaiseException(string.Format("Ex: {0} .\nStack Trace: {1} .", ex.Message, ex.StackTrace));
            }
        }

        private void StartFMBulkFileGenertor()
        {
            var allFMs = new List<TWFMTemplate>();
            var allFMsNew = new List<TWFMTemplate>();
            try
            {
                GenerateEmaFile();//11Nov2013 EMA + TW + WARRANT + ADD + FM.xls
                allFMs = GetFMTemplatesAndGenerateNDARevisedWntFM(configObj.AddFmFilePath);//TW_NDARevisedWntFM.xls
                GenerateNDAIAAddCSVFile(outputPath, allFMs);//20131106IAAdd.csv 
                UpdateNewFMProperties(allFMs, allFMsNew);
                UpdateFMProperties(allFMs);
                GenerateIDNFiles(outputPath, allFMs);//MAIN_RIC_TW.txt
                GenerateCsvByNewFMProperties(outputPath, allFMsNew);
                GenerateNDAQAAddCSVFile(outputPath, allFMs);
                GenerateISINTxtFile(outputPath, allFMs);//ISIN.txt
                GenerateBCUTxtFile(outputPath, allFMs);//BCU.txt
                GenerateTQRicTxtFile(outputPath, allFMs);//TQRIC INSERT.txt
                GenerateTickeLadderAddCsvFile(outputPath, allFMs);//[yyyyddmm]TickerLadderAdd.csv
                GenerateLotLadderAddCsvFiel(outputPath, allFMs);//[yyyyddmm]LotLadderAdd.csv
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating output files. Ex: {0} .", ex.Message));
            }
        }

        private void GenerateLotLadderAddCsvFiel(string fileDir, IEnumerable<TWFMTemplate> allFMs)
        {
            string strLotLadderAddFilePath = Path.Combine(fileDir, string.Format("{0}LotLadderAdd.csv", DateTime.Now.ToString("yyyyMMdd")));
            try
            {
                using (var app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, strLotLadderAddFilePath);
                    var worksheet = workbook.Worksheets[1] as Worksheet;
                    using (var writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                    {
                        writer.WriteLine("RIC");
                        writer.WriteLine("LOT NOT APPLICABLE");
                        writer.WriteLine("LOT LADDER NAME");
                        writer.WriteLine("LOT EFFECTIVE FROM");
                        writer.WriteLine("LOT EFFECTIVE TO");
                        writer.WriteLine("LOT PRICE INDICATOR");
                        writer.PlaceNext(writer.Row + 1, 1);
                        int logical_Key = 1;
                        foreach (TWFMTemplate tem in allFMs)
                        {
                            writer.WriteLine(tem.Ric);
                            writer.WriteLine("N");
                            writer.WriteLine("LOT_LADDER_EQTY_<1000>");
                            writer.WriteLine(DateTime.ParseExact(tem.EffectiveDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US"))
.ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));
                            writer.WriteLine("");
                            writer.WriteLine("ORDER");
                            writer.PlaceNext(writer.Row + 1, 1);
                            logical_Key++;
                        }
                    }
                    worksheet.UsedRange.NumberFormat = "@";
                    workbook.SaveAs(workbook.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating Lot Ladder Add Csv file. Ex: {0} .", ex.Message));
            }
            AddResult("Lot Ladder Add", strLotLadderAddFilePath, "nda");
        }


        private void GenerateTickeLadderAddCsvFile(string fileDir, IEnumerable<TWFMTemplate> allFMs)
        {
            string strTickerLadderAddFilePath = Path.Combine(fileDir, string.Format("{0}TickerLadderAdd.csv", DateTime.Now.ToString("yyyyMMdd")));
            try
            {
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, strTickerLadderAddFilePath);
                    Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                    using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                    {
                        writer.WriteLine("RIC");
                        writer.WriteLine("TICK NOT APPLICABLE");
                        writer.WriteLine("TICK LADDER NAME");
                        writer.WriteLine("TICK EFFECTIVE FROM");
                        writer.WriteLine("TICK EFFECTIVE TO");
                        writer.WriteLine("TICK PRICE INDICATOR");
                        writer.PlaceNext(writer.Row + 1, 1);
                        int logical_Key = 1;
                        foreach (TWFMTemplate tem in allFMs)
                        {
                            writer.WriteLine(tem.Ric);
                            writer.WriteLine("N");
                            writer.WriteLine("TICK_LADDER_TAI_2");
                            writer.WriteLine(DateTime.ParseExact(tem.EffectiveDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US"))
        .ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));
                            writer.WriteLine("");
                            writer.WriteLine("ORDER");
                            writer.PlaceNext(writer.Row + 1, 1);
                            logical_Key++;
                        }
                    }
                    worksheet.UsedRange.NumberFormat = "@";
                    workbook.SaveAs(workbook.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating Tick Ladder Add txt file. Ex: {0} .", ex.Message));
            }
            AddResult("Tick Ladder add", strTickerLadderAddFilePath, "nda");
        }

        //generate new .csv file 
        private void GenerateCsvByNewFMProperties(string fileDir, IEnumerable<TWFMTemplate> allFMsNew)
        {
            string fileName = DateTime.Now.ToString("ddMMMyyyy") + " EMA + TW + WARRANT + ADD + FM.csv";
            string filePath = Path.Combine(fileDir, fileName);
            try
            {
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                    Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                    using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                    {
                        writer.WriteLine("Logical_Key");
                        writer.WriteLine("Secondary_ID");
                        writer.WriteLine("Secondary_ID_Type");
                        writer.WriteLine("Warrant_Title");
                        writer.WriteLine("Issuer_OrgId");
                        writer.WriteLine("Issue_Date");
                        writer.WriteLine("Country_Of_Issue");
                        writer.WriteLine("Governing_Country");
                        writer.WriteLine("Announcement_Date");
                        writer.WriteLine("Payment_Date");
                        writer.WriteLine("Underlying_Type");
                        writer.WriteLine("Clearinghouse1_OrgId");
                        writer.WriteLine("Clearinghouse2_OrgId");
                        writer.WriteLine("Clearinghouse3_OrgId");
                        writer.WriteLine("Guarantor");
                        writer.WriteLine("Guarantor_Type");
                        writer.WriteLine("Guarantee_Type");
                        writer.WriteLine("Incr_Exercise_Lot");
                        writer.WriteLine("Min_Exercise_Lot");
                        writer.WriteLine("Max_Exercise_Lot");
                        writer.WriteLine("Rt_Page_Range");
                        writer.WriteLine("Underwriter1_OrgId");
                        writer.WriteLine("Underwriter1_Role");
                        writer.WriteLine("Underwriter2_OrgId");
                        writer.WriteLine("Underwriter2_Role");
                        writer.WriteLine("Underwriter3_OrgId");
                        writer.WriteLine("Underwriter3_Role");
                        writer.WriteLine("Underwriter4_OrgId");
                        writer.WriteLine("Underwriter4_Role");
                        writer.WriteLine("Exercise_Style");
                        writer.WriteLine("Warrant_Type");
                        writer.WriteLine("Expiration_Date");
                        writer.WriteLine("Registered_Bearer_Code");
                        writer.WriteLine("Price_Display_Type");
                        writer.WriteLine("Private_Placement");
                        writer.WriteLine("Coverage_Type");
                        writer.WriteLine("Warrant_Status");
                        writer.WriteLine("Status_Date");
                        writer.WriteLine("Redemption_Method");
                        writer.WriteLine("Issue_Price");
                        writer.WriteLine("Issue_Quantity");
                        writer.WriteLine("Issue_Currency");
                        writer.WriteLine("Issue_Price_Type");
                        writer.WriteLine("Issue_Spot_Price");
                        writer.WriteLine("Issue_Spot_Currency");
                        writer.WriteLine("Issue_Spot_FX_Rate");
                        writer.WriteLine("Issue_Delta");
                        writer.WriteLine("Issue_Elasticity");
                        writer.WriteLine("Issue_Gearing");
                        writer.WriteLine("Issue_Premium");
                        writer.WriteLine("Issue_Premium_PA");
                        writer.WriteLine("Denominated_Amount");
                        writer.WriteLine("Exercise_Begin_Date");
                        writer.WriteLine("Exercise_End_Date");
                        writer.WriteLine("Offset_Number");
                        writer.WriteLine("Period_Number");
                        writer.WriteLine("Offset_Frequency");
                        writer.WriteLine("Offset_Calendar");
                        writer.WriteLine("Period_Calendar");
                        writer.WriteLine("Period_Frequency");
                        writer.WriteLine("RAF_Event_Type");
                        writer.WriteLine("Exercise_Price");
                        writer.WriteLine("Exercise_Price_Type");
                        writer.WriteLine("Warrants_Per_Underlying");
                        writer.WriteLine("Underlying_FX_Rate");
                        writer.WriteLine("Underlying_RIC");
                        writer.WriteLine("Underlying_Item_Quantity");
                        writer.WriteLine("Units");
                        writer.WriteLine("Cash_Currency");
                        writer.WriteLine("Delivery_Type");
                        writer.WriteLine("Settlement_Type");
                        writer.WriteLine("Settlement_Currency");
                        writer.WriteLine("Underlying_Group");
                        writer.WriteLine("Country1_Code");
                        writer.WriteLine("Coverage1_Type");
                        writer.WriteLine("Country2_Code");
                        writer.WriteLine("Coverage2_Type");
                        writer.WriteLine("Country3_Code");
                        writer.WriteLine("Coverage3_Type");
                        writer.WriteLine("Country4_Code");
                        writer.WriteLine("Coverage4_Type");
                        writer.WriteLine("Country5_Code");
                        writer.WriteLine("Coverage5_Type");
                        writer.WriteLine("Note1_Type");
                        writer.WriteLine("Note1");
                        writer.WriteLine("Note2_Type");
                        writer.WriteLine("Note2");
                        writer.WriteLine("Note3_Type");
                        writer.WriteLine("Note3");
                        writer.WriteLine("Note4_Type");
                        writer.WriteLine("Note4");
                        writer.WriteLine("Note5_Type");
                        writer.WriteLine("Note5");
                        writer.WriteLine("Note6_Type");
                        writer.WriteLine("Note6");
                        writer.WriteLine("Exotic1_Parameter");
                        writer.WriteLine("Exotic1_Value");
                        writer.WriteLine("Exotic1_Begin_Date");
                        writer.WriteLine("Exotic1_End_Date");
                        writer.WriteLine("Exotic2_Parameter");
                        writer.WriteLine("Exotic2_Value");
                        writer.WriteLine("Exotic2_Begin_Date");
                        writer.WriteLine("Exotic2_End_Date");
                        writer.WriteLine("Exotic3_Parameter");
                        writer.WriteLine("Exotic3_Value");
                        writer.WriteLine("Exotic3_Begin_Date");
                        writer.WriteLine("Exotic3_End_Date");
                        writer.WriteLine("Exotic4_Parameter");
                        writer.WriteLine("Exotic4_Value");
                        writer.WriteLine("Exotic4_Begin_Date");
                        writer.WriteLine("Exotic4_End_Date");
                        writer.WriteLine("Exotic5_Parameter");
                        writer.WriteLine("Exotic5_Value");
                        writer.WriteLine("Exotic5_Begin_Date");
                        writer.WriteLine("Exotic5_End_Date");
                        writer.WriteLine("Exotic6_Parameter");
                        writer.WriteLine("Exotic6_Value");
                        writer.WriteLine("Exotic6_Begin_Date");
                        writer.WriteLine("Exotic6_End_Date");
                        writer.WriteLine("Event_Type1");
                        writer.WriteLine("Event_Period_Number1");
                        writer.WriteLine("Event_Calendar_Type1");
                        writer.WriteLine("Event_Frequency1");
                        writer.WriteLine("Event_Type2");
                        writer.WriteLine("Event_Period_Number2");
                        writer.WriteLine("Event_Calendar_Type2");
                        writer.WriteLine("Event_Frequency2");
                        writer.WriteLine("Exchange_Code1");
                        writer.WriteLine("Incr_Trade_Lot1");
                        writer.WriteLine("Min_Trade_Lot1");
                        writer.WriteLine("Min_Trade_Amount1");
                        writer.WriteLine("Exchange_Code2");
                        writer.WriteLine("Incr_Trade_Lot2");
                        writer.WriteLine("Min_Trade_Lot2");
                        writer.WriteLine("Min_Trade_Amount2");
                        writer.WriteLine("Exchange_Code3");
                        writer.WriteLine("Incr_Trade_Lot3");
                        writer.WriteLine("Min_Trade_Lot3");
                        writer.WriteLine("Min_Trade_Amount3");
                        writer.WriteLine("Exchange_Code4");
                        writer.WriteLine("Incr_Trade_Lot4");
                        writer.WriteLine("Min_Trade_Lot4");
                        writer.WriteLine("Min_Trade_Amount4");
                        writer.WriteLine("Attached_To_Id");
                        writer.WriteLine("Attached_To_Id_Type");
                        writer.WriteLine("Attached_Quantity");
                        writer.WriteLine("Attached_Code");
                        writer.WriteLine("Detachable_Date");
                        writer.WriteLine("Bond_Exercise");
                        writer.WriteLine("Bond_Price_Percentage");
                        writer.PlaceNext(writer.Row + 1, 1);
                        //Main Ric
                        int logical_Key = 1;
                        foreach (TWFMTemplate tem in allFMsNew)
                        {
                            writer.WriteLine(logical_Key);
                            writer.WriteLine(tem.OffcCode2);
                            writer.WriteLine("ISIN");
                            writer.WriteLine(tem.IDNLongName.Replace(",", ""));
                            writer.WriteLine(tem.Issuer_OrgId);
                            writer.WriteLine(tem.IssueDate);
                            writer.WriteLine("TWN");
                            writer.WriteLine("TWN");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine(tem.LocalSectorClassification);//K
                            writer.WriteLine("118750");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("1000");
                            writer.WriteLine("1000");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine(tem.GEN_TEXT16);//AD
                            writer.WriteLine(tem.PutCallInd);//AE
                            writer.WriteLine(tem.MaturDate);//AF
                            writer.WriteLine("K");
                            writer.WriteLine("D");
                            writer.WriteLine("N");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine(tem.TotalSharesOutstanding);
                            writer.WriteLine(tem.IssuePrice);//AO
                            writer.WriteLine(tem.Currency);//AP
                            writer.WriteLine("A");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("1000");
                            writer.WriteLine(tem.Exercise_Begin_Date);//BA
                            writer.WriteLine(tem.Exercise_End_Date);//BB
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine(tem.StrikePrc);//BJ
                            writer.WriteLine("A");
                            writer.WriteLine(tem.WntRatio);//BL
                            writer.WriteLine("");
                            writer.WriteLine(tem.UnderlyingRIC);//BN
                            writer.WriteLine("1");
                            writer.WriteLine(tem.Units);//BP
                            writer.WriteLine("");
                            writer.WriteLine(tem.ISS_TP_FLG);//BR
                            writer.WriteLine(tem.GN_TXT16_2);//BS
                            writer.WriteLine("TWD");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("CAP");
                            writer.WriteLine(tem.CapPrice);//CS
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("TAI");
                            writer.WriteLine("1000");
                            writer.WriteLine("1000");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.PlaceNext(writer.Row + 1, 1);
                            logical_Key++;
                        }
                    }
                    worksheet.UsedRange.NumberFormat = "@";
                    workbook.SaveAs(workbook.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating NDQAAAddCSVFile file, Ex: {0} ", ex.StackTrace));
            }
            AddResult("FM new csv file", filePath, "fm");
        }

        //update new list
        private void UpdateNewFMProperties(IEnumerable<TWFMTemplate> allFMs, List<TWFMTemplate> allFMsNew)
        {
            foreach (TWFMTemplate twfm in allFMs)
            {
                var tem = new TWFMTemplate
                {
                    Ric = twfm.Ric,
                    IssueDate = TWHelper.DateStringForm(twfm.IssueDate, "dd/MM/yyyy"),
                    IssuePrice = twfm.IssuePrice,
                    CapPrice = float.Parse(twfm.CapPrice).ToString(),
                    EffectiveDate = TWHelper.DateStringForm(twfm.EffectiveDate, "dd/MM/yyyy"),
                    MaturDate = TWHelper.DateStringForm(twfm.MaturDate, "dd/MM/yyyy"),
                    Exercise_End_Date = TWHelper.DateStringForm(twfm.MaturDate, "dd/MM/yyyy"),
                    GEN_TEXT16 = twfm.GEN_TEXT16,
                    DisplayName = twfm.DisplayName,
                    OfficialCode = twfm.OfficialCode,
                    ExchangeSymbol = twfm.ExchangeSymbol,
                    OffcCode2 = twfm.OffcCode2,
                    Currency = twfm.Currency,
                    RecordType = twfm.RecordType,
                    ChainRic = twfm.ChainRic,
                    PositionInChain = twfm.PositionInChain,
                    LotSize = twfm.LotSize,
                    CoiDisplyNmll = twfm.CoiDisplyNmll,
                    CoiSectorChain = twfm.CoiSectorChain,
                    BcastRef = twfm.BcastRef,
                    StrikePrc = float.Parse(twfm.StrikePrc).ToString(),
                    ConvFac = twfm.ConvFac,
                    Isin = twfm.Isin,
                    IDNLongName = twfm.IDNLongName,
                    IssueClassification = twfm.IssueClassification,
                    PrimaryListing = twfm.PrimaryListing,
                    OrganisationName = twfm.OrganisationName,
                    UnderlyingRIC = twfm.UnderlyingRIC,
                    IssuedCompanyName = twfm.IssuedCompanyName,
                    CompositeChainRic = twfm.CompositeChainRic,
                    LongLink1 = twfm.LongLink1,
                    LongLink2 = twfm.LongLink2,
                    LongLink3 = twfm.LongLink3,
                    LongLink4 = twfm.LongLink4,
                    LongLink5 = twfm.LongLink5,
                    LongLink6 = twfm.LongLink6,
                    LongLink7 = twfm.LongLink7,
                    LongLink8 = twfm.LongLink8,
                    LongLink9 = twfm.LongLink9,
                    BondType = twfm.BondType,
                    PutCallInd = (twfm.PutCallInd.Trim().StartsWith("C") ? "Cap Call" : "Cap Put"),
                    ISS_TP_FLG = twfm.ISS_TP_FLG,
                    Units = (twfm.ISS_TP_FLG.Trim() == "S" ? "shr" : "idx"),
                    GN_TXT16_2 = twfm.GN_TXT16_2,
                    Longlink1_Issuer = twfm.Longlink1_Issuer,
                    Longlink2_MenuPage = twfm.Longlink2_MenuPage,
                    Gearing = twfm.Gearing,
                    Premium = twfm.Premium,
                    LONGLINK1_TAS_RIC = twfm.LONGLINK1_TAS_RIC,
                    LONKLINK2_WT_Chain = twfm.LONKLINK2_WT_Chain,
                    LONKLINK3_Tech_Ric = twfm.LONKLINK3_Tech_Ric,
                    LONKLINK4_ValueAdded_Ric = twfm.LONKLINK4_ValueAdded_Ric,
                    Issuer_OrgId = twfm.Issuer_OrgId
                };
                tem.Ric = twfm.Ric;

                tem.LocalSectorClassification = (twfm.LocalSectorClassification.Trim() == "index warrants" ? "INDEX" : "STOCK");//K Underlying_Type
                tem.IndexRic = twfm.IndexRic;
                tem.Exercise_Begin_Date = (twfm.GEN_TEXT16.Trim() == "A" ? tem.EffectiveDate : tem.MaturDate);//BA Exercise_Begin_Date
                string[] s2 = twfm.WntRatio.Split(':');
                if (s2.Length > 1)
                {
                    tem.WntRatio = ((float.Parse(s2[0])) / (float.Parse(s2[1]))).ToString();//BL Warrants_Per_Underlying
                }
                string[] strTol = twfm.TotalSharesOutstanding.Trim().Split(' ')[0].Split(',');
                string tmpTol = strTol.Aggregate("", (current, t) => current + t);

                tem.TotalSharesOutstanding = tmpTol;//AN Issue_Quantity

                allFMsNew.Add(tem);
            }
        }

        private void UpdateFMProperties(IEnumerable<TWFMTemplate> allFMs)
        {
            foreach (TWFMTemplate tem in allFMs.Where(tem => !string.IsNullOrEmpty(tem.Ric)))
            {
                if (tem.Ric.EndsWith("TWO"))
                {
                    tem.Properties.IsTWO = true;
                }

                string code = tem.Ric.Remove(tem.Ric.IndexOf('.'));
                if (code.EndsWith("B") || code.EndsWith("C"))
                {
                    tem.Properties.IsCBBC = true;
                }

                if (tem.LocalSectorClassification.Contains("index") || tem.LocalSectorClassification.Contains("Index"))
                {
                    tem.Properties.IsIndex = true;
                }
            }
        }

        private string GetBCU(TWFMTemplate tem)
        {
            if ((tem.Ric + "").Trim().Length >= 6)
            {
                if (tem.Ric.Substring(5, 1).Equals("X") || tem.Ric.Substring(5, 1).Equals("Y"))
                {
                    return string.Format("TW_ISSUE_CBBCYT,TAIW_CBBC3,TAIW_CBBC,TAIW_XCBBC,TAIW_EQ_{0}_CBBC,TAIW_EQ_{0}_REL", tem.BcastRef.Substring(0, tem.BcastRef.IndexOf('.')));
                }
            }

            string BCU = "";
            char indexNum = '0';
            for (int i = 0; i < tem.Ric.Length - 1; i++)
            {
                if (tem.Ric[i] != '0')
                {
                    indexNum = tem.Ric[i];
                    break;
                }
            }
            if (tem.Properties.IsTWO == false)
            {
                if (tem.Properties.IsCBBC)
                {
                    if (tem.Properties.IsIndex)
                    {
                        //BCU += string.Format("TAIW_CBBC,TAIW_INX_TWII_CBBC,TAIW_CBBC{0},OTCTWS_EQ_{0}_REL", indexNum);
                        BCU += string.Format("TAIW_CBBC,TAIW_INX_TWII_CBBC,TAIW_CBBC{0}", indexNum);
                    }
                    else
                    {
                        //BCU += "TAIW_CBBC,TAIW_EQ_";
                        //BCU += tem.BcastRef.Substring(0, tem.BcastRef.IndexOf('.'));
                        //BCU += "_REL,TAIW_EQ_";
                        //BCU += tem.BcastRef.Substring(0, tem.BcastRef.IndexOf('.'));
                        //BCU += string.Format("_CBBC,TAIW_CBBC{0}", indexNum);
                        BCU += string.Format("TAIW_CBBC,TAIW_INX_TWII_CBBC,TAIW_CBBC{0},TAIW_EQ_{0}_REL,TAIW_EQ_{0}_CBBC", indexNum);
                    }
                }
                else
                {
                    if (tem.Properties.IsIndex)
                    {
                        BCU += string.Format("TAIW_EQ_CWNT,TAIW_EQ_CWNT{0},TAIW_EQ_IWNT", indexNum);
                    }
                    else
                    {
                        BCU += "TAIW_EQ_SWNT,TAIW_EQ_CWNT,TAIW_EQ_";
                        BCU += tem.BcastRef.Substring(0, tem.BcastRef.IndexOf('.'));
                        BCU += "_REL,TAIW_EQ_";
                        BCU += tem.BcastRef.Substring(0, tem.BcastRef.IndexOf('.'));
                        BCU += string.Format("_WT,TAIW_EQ_CWNT{0}", indexNum);
                    }
                }
            }
            else
            {
                if (!tem.Properties.IsCBBC)
                {
                    BCU +=
                        string.Format("OTCTWS_EQ_SWNT,OTCTWS_EQ_CWNT,OTCTWS_EQ_{0}_REL,OTCTWS_EQ_{0}_WT,OTCTWS_EQ_CWNT{1}",
                        tem.BcastRef.Replace(".TWO", ""), tem.Ric[1]);
                }
                else
                {
                    BCU +=
                        string.Format("OTCTWS_CBBC,OTCTWS_EQ_{0}_REL,OTCTWS_EQ_{0}_CBBC,OTCTWS_CBBC{1}",
                        tem.BcastRef.Replace(".TWO", ""), tem.Ric[1]);
                }
            }
            return BCU;
        }

        private void GenerateTQRicTxtFile(string fileDir, IEnumerable<TWFMTemplate> allFMs)
        {
            string warrantAddTQRicTxtFilePath = Path.Combine(fileDir, "TQRIC INSERT.txt");
            string[] content = new string[3];
            string part1CBBC = "TAI_STRIKEPRICE.DAT;07-JAN-2008 09:00:00;TPS;";
            part1CBBC += "\r\n";
            part1CBBC += "STRIKE_PRC;CONV_FAC;SECTOR_ID;GEN_VAL6;";
            part1CBBC += "\r\n";
            string part2 = "TAI_STRIKEPRICE.DAT;07-JAN-2008 09:00:00;TPS;\r\n";
            part2 += "STRIKE_PRC;CONV_FAC;SECTOR_ID;";
            part2 += "\r\n";
            string part3 = "HKSE.DAT;07-APR-2004 09:00:00;TPS;\r\n";
            part3 += "LONGLINK1;LONGLINK2;";
            part3 += "\r\n";

            foreach (TWFMTemplate tem in allFMs)
            {
                File.WriteAllLines(warrantAddTQRicTxtFilePath, content, Encoding.UTF8);
                if (tem.Properties.IsCBBC)
                {
                    part1CBBC += string.Format("{0};{1};{2};DUM;{3};\r\n", tem.Ric, tem.StrikePrc, tem.ConvFac, tem.CapPrice);
                }
                else
                {
                    try
                    {
                        part2 += string.Format("{0};{1};{2};DUM;\r\n", tem.Ric, tem.StrikePrc, tem.ConvFac);
                        string s = GetStrForPart3ForTQRic(tem);
                        part3 += string.Format("{0}\r\n", s);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(string.Format("Error happens when generating TQRIC txt file. Ex: {0} .", ex.Message));
                    }
                }
            }
            part1CBBC += "\r\n\r\n";
            part2 += "\r\n\r\n";
            content[0] = part1CBBC;
            content[1] = part2;
            content[2] = part3;
            try
            {
                File.WriteAllLines(warrantAddTQRicTxtFilePath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating TQRIC txt file. Ex: {0} .", ex.Message));
            }
            AddResult("BCU file", warrantAddTQRicTxtFilePath, "bcu");
        }

        private string GetStrForPart3ForTQRic(TWFMTemplate tem)
        {
            string text = "";
            if (tem.Properties.IsIndex)
            {
                if (tem.Properties.IsTWO)
                {
                }
                else
                {
                    text = string.Format("{0};.TWII;TW/WTS1;", tem.LongLink4);
                }
            }
            else
            {
                text = string.Format(tem.Properties.IsTWO ? "{0};{1};TWO/WTS01;" : "{0};{1};TW/WTS1;", tem.LongLink4, tem.BcastRef);
            }
            return text;
        }

        private void GenerateBCUTxtFile(string fileDir, IEnumerable<TWFMTemplate> allFMs)
        {
            string warrantAddBCUTxtFilePath = Path.Combine(fileDir, "BCU.txt");
            string content = "RIC\tBCU\r\n";
            foreach (TWFMTemplate tem in allFMs)
            {
                content += string.Format("{0}\t", tem.Ric);
                content += string.Format(GetBCU(tem));
                content += "\r\n";
            }
            try
            {
                File.WriteAllText(warrantAddBCUTxtFilePath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating BCU txt file. Ex: {0} .", ex.Message));
            }
            AddResult("BCU file", warrantAddBCUTxtFilePath, "bcu");
        }

        private void GenerateISINTxtFile(string fileDir, IEnumerable<TWFMTemplate> allFMs)
        {
            string warrantAddISINTxtFilePath = Path.Combine(fileDir, "ISIN.txt");
            string content = "RIC\t#INSTMOD_OFFC_CODE2\r\n";
            foreach (TWFMTemplate tem in allFMs)
            {
                content += string.Format("{0}\t", tem.Ric);
                content += string.Format("{0}", tem.OffcCode2);
                content += "\r\n";
            }
            try
            {
                File.WriteAllText(warrantAddISINTxtFilePath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating ISIN txt file. Ex: {0} .", ex.Message));
            }
            AddResult("ISIN file", warrantAddISINTxtFilePath, "isin");
        }

        private void GenerateEmaFile()
        {
            string fmFilePath = configObj.AddFmFilePath;
            string fileName = DateTime.Now.ToString("ddMMMyyyy") + " EMA + TW + WARRANT + ADD + FM.xls";
            String destFile = Path.Combine(outputPath, fileName);
            File.Copy(fmFilePath, destFile, true);
            using (ExcelApp app = new ExcelApp(false, false))
            {

                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, destFile);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                int lastUsedColumn = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
                Range range;
                Range rangeToRemove = ExcelUtil.GetRange(1, 50, lastUsedRow, lastUsedColumn, worksheet);
                rangeToRemove.Clear();
                for (int i = 2; i <= lastUsedRow; i++)
                {
                    range = ExcelUtil.GetRange("AR" + i, worksheet);
                    if (string.IsNullOrEmpty(range.Text.ToString()))
                        break;
                    range.Value = range.Text.ToString().Substring(0, 1).Equals("C") ? "CA_CALL" : "PU_PUT";
                }
                workbook.Save();
                workbook.Close(false, workbook.FullName, false);
            }
        }

        private List<TWFMTemplate> GetFMTemplatesAndGenerateNDARevisedWntFM(string fmFilePath)
        {
            string NDARevisedWntFMPath = Path.Combine(outputPath, "TW_NDARevisedWntFM.xls");
            List<TWFMTemplate> fmList = new List<TWFMTemplate>();
            using (ExcelApp app = new ExcelApp(false, false))
            {
                if (!File.Exists(fmFilePath))
                {
                    Logger.Log(string.Format("Can't find the FM file in the path {0} .", fmFilePath));
                    return fmList;
                }
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, fmFilePath);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                using (ExcelLineWriter reader = new ExcelLineWriter(worksheet, 2, 1, ExcelLineWriter.Direction.Right))
                {
                    for (int i = 1; i <= lastUsedRow; i++)
                    {
                        string ric = reader.ReadLineCellText();
                        if (!string.IsNullOrEmpty(ric))
                        {
                            TWFMTemplate tem = new TWFMTemplate();
                            tem.Ric = ric;//A
                            tem.IssueDate = reader.ReadLineCellText();//B
                            tem.IssuePrice = reader.ReadLineCellText();//C
                            tem.CapPrice = reader.ReadLineCellText();//D
                            tem.EffectiveDate = reader.ReadLineCellText();
                            tem.DisplayName = reader.ReadLineCellText();
                            tem.OfficialCode = reader.ReadLineCellText();
                            tem.ExchangeSymbol = reader.ReadLineCellText();//H
                            tem.OffcCode2 = reader.ReadLineCellText();
                            tem.Currency = reader.ReadLineCellText();
                            tem.RecordType = reader.ReadLineCellText();
                            tem.ChainRic = reader.ReadLineCellText();
                            tem.PositionInChain = reader.ReadLineCellText();
                            tem.LotSize = reader.ReadLineCellText();
                            tem.CoiDisplyNmll = reader.ReadLineCellText();
                            tem.CoiSectorChain = reader.ReadLineCellText();
                            tem.BcastRef = reader.ReadLineCellText();
                            tem.WntRatio = reader.ReadLineCellText();//R
                            tem.StrikePrc = reader.ReadLineCellText();
                            tem.MaturDate = reader.ReadLineCellText();
                            tem.ConvFac = reader.ReadLineCellText();
                            tem.Isin = reader.ReadLineCellText();
                            tem.IDNLongName = reader.ReadLineCellText();
                            tem.IssueClassification = reader.ReadLineCellText();
                            tem.PrimaryListing = reader.ReadLineCellText();//Y
                            tem.OrganisationName = reader.ReadLineCellText();
                            tem.UnderlyingRIC = reader.ReadLineCellText();
                            tem.IssuedCompanyName = reader.ReadLineCellText();
                            tem.Ric = reader.ReadLineCellText();
                            tem.LocalSectorClassification = reader.ReadLineCellText();
                            tem.IndexRic = reader.ReadLineCellText();
                            tem.TotalSharesOutstanding = reader.ReadLineCellText();
                            tem.CompositeChainRic = reader.ReadLineCellText();
                            tem.LongLink1 = reader.ReadLineCellText();//AH
                            tem.LongLink2 = reader.ReadLineCellText();
                            tem.LongLink3 = reader.ReadLineCellText();
                            tem.LongLink4 = reader.ReadLineCellText();
                            tem.LongLink5 = reader.ReadLineCellText();
                            tem.LongLink6 = reader.ReadLineCellText();
                            tem.LongLink7 = reader.ReadLineCellText();
                            tem.LongLink8 = reader.ReadLineCellText();
                            tem.LongLink9 = reader.ReadLineCellText();
                            tem.BondType = reader.ReadLineCellText();
                            tem.PutCallInd = reader.ReadLineCellText();//AR
                            tem.ISS_TP_FLG = reader.ReadLineCellText();
                            tem.GEN_TEXT16 = reader.ReadLineCellText();
                            tem.GN_TXT16_2 = reader.ReadLineCellText();
                            tem.Longlink1_Issuer = reader.ReadLineCellText();
                            tem.Longlink2_MenuPage = reader.ReadLineCellText();
                            tem.Gearing = reader.ReadLineCellText();
                            tem.Premium = reader.ReadLineCellText();//AY
                            tem.LONGLINK1_TAS_RIC = reader.ReadLineCellText();
                            tem.LONKLINK2_WT_Chain = reader.ReadLineCellText();
                            tem.LONKLINK3_Tech_Ric = reader.ReadLineCellText();
                            tem.LONKLINK4_ValueAdded_Ric = reader.ReadLineCellText();//BC
                            fmList.Add(tem);
                        }
                        reader.PlaceNext(reader.Row + 1, 1);
                    }
                }

                for (int i = 2; i <= lastUsedRow + 1; i++)
                {

                    if (ExcelUtil.GetRange(i, 44, worksheet).Text.ToString().StartsWith("C"))
                    {
                        worksheet.Cells[i, 44] = "CA_CALL";
                    }

                    if (ExcelUtil.GetRange(i, 44, worksheet).Text.ToString().StartsWith("P"))
                    {
                        worksheet.Cells[i, 44] = "PU_PUT";
                    }
                }

                ExcelUtil.GetRange(1, 50, lastUsedRow, 60, worksheet).Clear();
                worksheet.UsedRange.NumberFormat = "@";
                workbook.SaveAs(NDARevisedWntFMPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddResult("NDA Revised FM File", NDARevisedWntFMPath, "fm");
                workbook.Close(false, workbook.FullName, false);
            }
            return fmList;
        }

        #region IDN File Generator
        private void GenerateIDNFiles(string fileDir, IEnumerable<TWFMTemplate> allFMs)
        {
            string warrantIDNFilePath = Path.Combine(fileDir, "MAIN_RIC_TW.txt");
            string content = "SYMBOL\tDSPLY_NAME\tRIC\tOFFCL_CODE\tEX_SYMBOL\tEXPIR_DATE\tDSPLY_NMLL\tBCAST_REF\t#INSTMOD_GEN_TEXT16\t#INSTMOD_GN_TXT16_2\t#INSTMOD_PUTCALLIND\t#INSTMOD_TDN_SYMBOL\t#INSTMOD_MNEMONIC\tEXL_NAME\r\n";
            foreach (TWFMTemplate tem in allFMs)
            {
                content += string.Format("{0}\t", tem.Ric);
                content += string.Format("{0}\t", tem.DisplayName);
                content += string.Format("{0}\t", tem.Ric);
                content += string.Format("{0}\t", tem.OfficialCode);
                content += string.Format("{0}\t", tem.ExchangeSymbol);
                content += string.Format("{0}\t", tem.MaturDate);
                content += string.Format("{0}\t", tem.CoiDisplyNmll);
                content += string.Format("{0}\t", tem.BcastRef);
                content += string.Format("{0}\t", tem.GEN_TEXT16);
                content += string.Format("{0}\t", tem.GN_TXT16_2);
                content += string.Format("{0}\t", ForamtInstmodPutCallIND(tem.PutCallInd));
                if (tem.Properties.IsTWO)
                {
                    content += string.Format("{0}\t", "");
                }
                else
                {
                    content += string.Format("{0}\t", tem.OfficialCode);
                }
                content += string.Format("{0}\t", tem.OfficialCode);
                if (IsSpecailRic(tem.Ric))
                {
                    content += string.Format("{0}\t", "TAIW_CBBC");
                }
                else
                {
                    content += string.Format("{0}\t",
                                            (tem.Properties.IsCBBC ? "TAIW_CBBC" : (tem.LocalSectorClassification.Contains("index") ? "TAIW_INX_WNT" :
                                            (tem.Properties.IsTWO ? "OTCTWS_WNT" : "TAIW_EQLB_WNT"))));
                }
                content += "\r\n";
            }
            try
            {
                File.WriteAllText(warrantIDNFilePath, content, Encoding.UTF8);
            }

            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating IDN files. EX: {0}", ex.Message));
            }
            AddResult("IDN File", warrantIDNFilePath, "idn");
        }

        private bool IsSpecailRic(string str)
        {
            bool result = false;

            if ((str + "").Trim().Length < 1)
                return false;

            try
            {
                string comChar = string.Empty;
                int index;
                if (!str.Contains("."))
                    index = str.Length - 1;
                else
                    index = str.IndexOf(".") - 1;

                comChar = str.Substring(index, 1);
                if (comChar.Equals("X") || comChar.Equals("Y"))
                    result = true;
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

        private string ForamtInstmodPutCallIND(string str)
        {
            if ((str + "").Trim().Length < 1)
                return string.Empty;

            string startChar = string.Empty;

            try
            {
                startChar = str.Trim().Substring(0, 1);

                if (startChar.Equals("P"))
                    return "PU_PUT";

                if (startChar.Equals("C"))
                    return "CA_CALL";
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return string.Empty;
        }
        #endregion

        #region NDA File Generator
        private void GenerateNDAIAAddCSVFile(string fileDir, IEnumerable<TWFMTemplate> allFMs)
        {
            string filePath = Path.Combine(fileDir, string.Format("{0}IAAdd.csv", DateTime.Now.ToString("yyyyMMdd")));
            try
            {
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                    Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                    using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                    {
                        writer.WriteLine("TAIWAN CODE");
                        writer.WriteLine("TYPE");
                        writer.WriteLine("CATEGORY");
                        writer.WriteLine("RCS ASSET CLASS");
                        writer.WriteLine("WARRANT ISSUER");
                        writer.WriteLine("ISIN");
                        writer.WriteLine("WARRANT ISSUE QUANTITY");
                        writer.PlaceNext(writer.Row + 1, 1);
                        foreach (TWFMTemplate tem in allFMs)
                        {
                            writer.WriteLine(tem.OfficialCode);
                            writer.WriteLine("DERIVATIVE");
                            writer.WriteLine("EIW");
                            writer.WriteLine((tem.OfficialCode.EndsWith("C") || tem.OfficialCode.EndsWith("B")) ? "BARWNT" : "TRAD");
                            try
                            {
                                string warrIss = TWIssueManager.GetByEnglishFullName(tem.OrganisationName).WarrantIssuer;
                                writer.WriteLine(warrIss);//Warrant Issuer
                                tem.Issuer_OrgId = warrIss;
                            }
                            catch (Exception ex)
                            {
                                Logger.Log(string.Format("Error happens when generating IAAdd CSV file. EX: {0}", ex.Message));
                            }
                            writer.WriteLine(tem.OffcCode2);
                            writer.WriteLine(tem.TotalSharesOutstanding.ToLower().Replace(",", "").Replace(" units", ""));
                            writer.PlaceNext(writer.Row + 1, 1);
                        }
                    }
                    worksheet.UsedRange.NumberFormat = "@";
                    workbook.SaveAs(filePath, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating NDAIAAddCSVFile file, Ex: {0} ", ex.StackTrace));
            }
            AddResult("NDA IAAdd csv file", filePath, "nda");
        }

        private void GenerateNDAQAAddCSVFile(string fileDir, List<TWFMTemplate> allFMs)
        {
            string filePath = Path.Combine(fileDir, string.Format("{0}QAAdd.csv", DateTime.Now.ToString("yyyyMMdd")));
            try
            {
                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                    Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                    using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                    {
                        writer.WriteLine("RIC");
                        writer.WriteLine("TAG");
                        writer.WriteLine("ASSET COMMON NAME");
                        writer.WriteLine("ASSET SHORT NAME");
                        writer.WriteLine("CURRENCY");
                        writer.WriteLine("EXCHANGE");
                        writer.WriteLine("TYPE");
                        writer.WriteLine("CATEGORY");
                        writer.WriteLine("BASE ASSET");
                        writer.WriteLine("EXPIRY DATE");
                        writer.WriteLine("STRIKE PRICE");
                        writer.WriteLine("CALL PUT OPTION");
                        writer.WriteLine("ROUND LOT SIZE");
                        writer.WriteLine("TRADING SEGMENT");
                        writer.WriteLine("TICKER SYMBOL");
                        writer.WriteLine("DERIVATIVES FIRST TRADING DAY");
                        writer.WriteLine("WARRANT ISSUE PRICE");
                        writer.PlaceNext(writer.Row + 1, 1);
                        //Main Ric
                        foreach (TWFMTemplate tem in allFMs)
                        {
                            writer.WriteLine(tem.Ric);
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "2132" : "2042");
                            writer.WriteLine(tem.IDNLongName.Replace("@", " "));
                            writer.WriteLine(tem.DisplayName);
                            writer.WriteLine("TWD");
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "TWO" : "TAI");
                            writer.WriteLine("DERIVATIVE"); //Type
                            writer.WriteLine("EIW");
                            writer.WriteLine(string.Format("ISIN:{0}", tem.Isin));
                            writer.WriteLine(DateTime.ParseExact(tem.MaturDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US"))
                                    .ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine(tem.StrikePrc);
                            writer.WriteLine(tem.PutCallInd.StartsWith("C") ? "CALL" : "PUT");
                            writer.WriteLine(tem.LotSize.Replace(" Units", ""));
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "TWO:ROCO" : "TAI:XTAI");
                            writer.WriteLine(tem.OfficialCode);
                            writer.WriteLine(DateTime.ParseExact(tem.EffectiveDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US")).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine(tem.IssuePrice);
                            writer.PlaceNext(writer.Row + 1, 1);
                        }
                        //stat Ric
                        foreach (TWFMTemplate tem in allFMs)
                        {
                            writer.WriteLine(tem.LongLink3);
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "46432" : "40938");
                            writer.WriteLine(tem.IDNLongName.Replace("@", " "));
                            writer.WriteLine(tem.DisplayName);
                            writer.WriteLine("TWD");
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "TWO" : "TAI");
                            writer.WriteLine("DERIVATIVE"); //Type
                            writer.WriteLine("EIW");
                            writer.WriteLine(string.Format("ISIN:{0}", tem.Isin));
                            writer.WriteLine(DateTime.ParseExact(tem.MaturDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US"))
                                    .ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine(tem.StrikePrc);
                            writer.WriteLine(tem.PutCallInd.StartsWith("C") ? "CALL" : "PUT");
                            writer.WriteLine(tem.LotSize.Replace(" Units", ""));
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "TWO:ROCO" : "TAI:XTAI");
                            writer.WriteLine(tem.OfficialCode);
                            writer.WriteLine(DateTime.ParseExact(tem.EffectiveDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US")).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine("");
                            writer.PlaceNext(writer.Row + 1, 1);
                        }
                        //va Ric
                        foreach (TWFMTemplate tem in allFMs.Where(tem => !tem.Properties.IsCBBC))
                        {
                            writer.WriteLine(tem.LongLink4);
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "2162" : "2161");
                            writer.WriteLine(tem.IDNLongName.Replace("@", " "));
                            writer.WriteLine(tem.DisplayName);
                            writer.WriteLine("TWD");
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "TWO" : "TAI");
                            writer.WriteLine("DERIVATIVE"); //Type
                            writer.WriteLine("EIW");
                            writer.WriteLine(string.Format("ISIN:{0}", tem.Isin));
                            writer.WriteLine(DateTime.ParseExact(tem.MaturDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US"))
                                .ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine(tem.StrikePrc);
                            writer.WriteLine(tem.PutCallInd.StartsWith("C") ? "CALL" : "PUT");
                            writer.WriteLine(tem.LotSize.Replace(" Units", ""));
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "TWO:ROCO" : "TAI:XTAI");
                            writer.WriteLine("");
                            writer.WriteLine(DateTime.ParseExact(tem.EffectiveDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US")).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine("");
                            writer.PlaceNext(writer.Row + 1, 1);
                        }
                        //ta Ric
                        foreach (TWFMTemplate tem in allFMs)
                        {
                            writer.WriteLine(tem.LongLink2);
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "40121" : "40120");
                            writer.WriteLine(tem.IDNLongName.Replace("@", " "));
                            writer.WriteLine(tem.DisplayName);
                            writer.WriteLine("TWD");
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "TWO" : "TAI");
                            writer.WriteLine("DERIVATIVE"); //Type
                            writer.WriteLine("EIW");
                            writer.WriteLine(string.Format("ISIN:{0}", tem.Isin));
                            writer.WriteLine(DateTime.ParseExact(tem.MaturDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US"))
                                    .ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine(DateTime.ParseExact(tem.EffectiveDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US")).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine("");
                            writer.PlaceNext(writer.Row + 1, 1);
                        }
                        //f Ric
                        foreach (TWFMTemplate tem in allFMs.Where(tem => tem.Ric.EndsWith("TWO")))
                        {
                            writer.WriteLine(string.Format("{0}f.TWO", tem.OfficialCode));
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "46469" : "40120");
                            writer.WriteLine(tem.IDNLongName.Replace("@", " "));
                            writer.WriteLine(tem.DisplayName);
                            writer.WriteLine("TWD");
                            writer.WriteLine(tem.Ric.EndsWith(".TWO") ? "TWO" : "TAI");
                            writer.WriteLine("DERIVATIVE"); //Type
                            writer.WriteLine("EIW");
                            writer.WriteLine(string.Format("ISIN:{0}", tem.Isin));
                            writer.WriteLine(DateTime.ParseExact(tem.MaturDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US"))
                                .ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine("");
                            writer.WriteLine(DateTime.ParseExact(tem.EffectiveDate, "dd-MMM-yy", new System.Globalization.CultureInfo("en-US")).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US")));//Expiry Date
                            writer.WriteLine("");
                            writer.PlaceNext(writer.Row + 1, 1);
                        }
                    }
                    worksheet.UsedRange.NumberFormat = "@";
                    workbook.SaveAs(workbook.FullName, XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating NDQAAAddCSVFile file, Ex: {0} ", ex.StackTrace));
            }
            AddResult("NDA QAAdd csv file", filePath, "nda");
        }
        #endregion
    }
}
