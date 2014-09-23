using Ric.FileLib.Attribute;

namespace Ric.FileLib.Entry
{
    /// <summary>
    /// Basic representation of a Nda bulk file entry
    /// </summary>
    public class TcEntry : AEntry
    {
        [TitleName("Logical_Key")]
        public string LogicalKey { get; set; }

        [TitleName("Secondary_ID")]
        public string SecondaryId { get; set; }

        [TitleName("Secondary_ID_Type")]
        public string SecondaryIdType { get; set; }

        [TitleName("Warrant_Title")]
        public string WarrantTitle { get; set; }

        [TitleName("Issuer_OrgId")]
        public string IssuerOrgId { get; set; }

        [TitleName("Issue_Date")]
        public string IssueDate { get; set; }

        [TitleName("Country_Of_Issue")]
        public string CountryOfIssue { get; set; }

        [TitleName("Governing_Country")]
        public string GoverningCountry { get; set; }

        [TitleName("Announcement_Date")]
        public string AnnouncementDate { get; set; }

        [TitleName("Payment_Date")]
        public string PaymentDate { get; set; }

        [TitleName("Underlying_Type")]
        public string UnderlyingType { get; set; }

        [TitleName("Clearinghouse1_OrgId")]
        public string ClearingHouse1OrdId { get; set; }

        [TitleName("Clearinghouse2_OrgId")]
        public string ClearingHouse2OrdId { get; set; }

        [TitleName("Clearinghouse3_OrgId")]
        public string ClearingHouse3OrdId { get; set; }

        [TitleName("Guarantor")]
        public string Guarantor { get; set; }

        [TitleName("Guarantor_Type")]
        public string GuarantorType { get; set; }

        [TitleName("Guarantee_Type")]
        public string GuaranteeType { get; set; }

        [TitleName("Incr_Exercise_Lot")]
        public string IncrExerciseLot { get; set; }

        [TitleName("Min_Exercise_Lot")]
        public string MinExerciseLot { get; set; }

        [TitleName("Max_Exercise_Lot")]
        public string MaxExerciseLot { get; set; }

        [TitleName("Rt_Page_Range")]
        public string RtPageRange { get; set; }

        [TitleName("Underwriter1_OrgId")]
        public string Underwriter1OrgId { get; set; }

        [TitleName("Underwriter1_Role")]
        public string Underwriter1Role { get; set; }

        [TitleName("Underwriter2_OrgId")]
        public string Underwriter2OrgId { get; set; }

        [TitleName("Underwriter2_Role")]
        public string Underwriter2Role { get; set; }

        [TitleName("Underwriter3_OrgId")]
        public string Underwriter3OrgId { get; set; }

        [TitleName("Underwriter3_Role")]
        public string Underwriter3Role { get; set; }

        [TitleName("Underwriter4_OrgId")]
        public string Underwriter4OrgId { get; set; }

        [TitleName("Underwriter4_Role")]
        public string Underwriter4Role { get; set; }

        [TitleName("Exercise_Style")]
        public string ExerciseStyle { get; set; }

        [TitleName("Warrant_Type")]
        public string WarrantType { get; set; }

        [TitleName("Expiration_Date")]
        public string ExpirationDate { get; set; }

        [TitleName("Registered_Bearer_Code")]
        public string RegisteredBearerCode { get; set; }

        [TitleName("Price_Display_Type")]
        public string PriceDisplayType { get; set; }

        [TitleName("Private_Placement")]
        public string PrivatePlacement { get; set; }

        [TitleName("Coverage_Type")]
        public string CoverageType { get; set; }

        [TitleName("Warrant_Status")]
        public string WarrantStatus { get; set; }

        [TitleName("Status_Date")]
        public string StatusDate { get; set; }

        [TitleName("Redemption_Method")]
        public string RedemptionMethod { get; set; }

        [TitleName("Issue_Quantity")]
        public string IssueQuantity { get; set; }

        [TitleName("Issue_Price")]
        public string IssuePrice { get; set; }

        [TitleName("Issue_Currency")]
        public string IssueCurrency { get; set; }

        [TitleName("Issue_Price_Type")]
        public string IssuePriceType { get; set; }

        [TitleName("Issue_Spot_Price")]
        public string IssueSpotPrice { get; set; }

        [TitleName("Issue_Spot_Currency")]
        public string IssueSpotCurrency { get; set; }

        [TitleName("Issue_Spot_FX_Rate")]
        public string IssueSpotFxRate { get; set; }

        [TitleName("Issue_Delta")]
        public string IssueDelta { get; set; }

        [TitleName("Issue_Elasticity")]
        public string IssueElasticity { get; set; }

        [TitleName("Issue_Gearing")]
        public string IssueGearing { get; set; }

        [TitleName("Issue_Premium")]
        public string IssuePremium { get; set; }

        [TitleName("Issue_Premium_PA")]
        public string IssuePremiumPa { get; set; }

        [TitleName("Denominated_Amount")]
        public string DenominatedAmount { get; set; }

        [TitleName("Exercise_Begin_Date")]
        public string ExerciseBeginDate { get; set; }

        [TitleName("Exercise_End_Date")]
        public string ExerciseEndDate { get; set; }

        [TitleName("Offset_Number")]
        public string OffsetNumber { get; set; }

        [TitleName("Period_Number")]
        public string PeriodNumber { get; set; }

        [TitleName("Offset_Frequency")]
        public string OffsetFrequency { get; set; }

        [TitleName("Offset_Calendar")]
        public string OffsetCalendar { get; set; }

        [TitleName("Period_Calendar")]
        public string PeriodCalendar { get; set; }

        [TitleName("Period_Frequency")]
        public string PeriodFrequency { get; set; }

        [TitleName("RAF_Event_Type")]
        public string RafEventType { get; set; }

        [TitleName("Exercise_Price")]
        public string ExercisePrice { get; set; }

        [TitleName("Exercise_Price_Type")]
        public string ExercisePriceType { get; set; }

        [TitleName("Warrants_Per_Underlying")]
        public string WarrantsPerUnderlying { get; set; }

        [TitleName("Underlying_FX_Rate")]
        public string UnderlyingFxRate { get; set; }

        [TitleName("Underlying_RIC")]
        public string UnderlyingRic { get; set; }

        [TitleName("Underlying_Item_Quantity")]
        public string UnderlyingItemQuantity { get; set; }

        [TitleName("Units")]
        public string Units { get; set; }

        [TitleName("Cash_Currency")]
        public string CashCurrency { get; set; }

        [TitleName("Delivery_Type")]
        public string DeliveryType { get; set; }

        [TitleName("Settlement_Type")]
        public string SettlementType { get; set; }

        [TitleName("Settlement_Currency")]
        public string SettlementCurrency { get; set; }

        [TitleName("Underlying_Group")]
        public string UnderlyingGroup { get; set; }

        [TitleName("Country1_Code")]
        public string Country1Code { get; set; }

        [TitleName("Coverage1_Type")]
        public string Coverage1Type { get; set; }

        [TitleName("Country2_Code")]
        public string Country2Code { get; set; }

        [TitleName("Coverage2_Type")]
        public string Coverae2Type { get; set; }

        [TitleName("Country3_Code")]
        public string Country3Code { get; set; }

        [TitleName("Coverage3_Type")]
        public string Coverage3Type { get; set; }

        [TitleName("Country4_Code")]
        public string Country4Code { get; set; }

        [TitleName("Coverage4_Type")]
        public string Coverage4Type { get; set; }

        [TitleName("Country5_Code")]
        public string Country5Code { get; set; }

        [TitleName("Coverage5_Type")]
        public string Coverage5Type { get; set; }

        [TitleName("Note1_Type")]
        public string Note1Type { get; set; }

        [TitleName("Note1")]
        public string Note1 { get; set; }

        [TitleName("Note2_Type")]
        public string Note2Type { get; set; }

        [TitleName("Note2")]
        public string Note2 { get; set; }

        [TitleName("Note3_Type")]
        public string Note3Type { get; set; }

        [TitleName("Note3")]
        public string Note3 { get; set; }

        [TitleName("Note4_Type")]
        public string Note4Type { get; set; }

        [TitleName("Note4")]
        public string Note4 { get; set; }

        [TitleName("Note5_Type")]
        public string Note5Type { get; set; }

        [TitleName("Note5")]
        public string Note5 { get; set; }

        [TitleName("Note6_Type")]
        public string Note6Type { get; set; }

        [TitleName("Note6")]
        public string Note6 { get; set; }

        [TitleName("Exotic1_Parameter")]
        public string Exotic1Parameter { get; set; }

        [TitleName("Exotic1_Value")]
        public string Exotic1Value { get; set; }

        [TitleName("Exotic1_Begin_Date")]
        public string Exotic1BeginDate { get; set; }

        [TitleName("Exotic1_End_Date")]
        public string Exotic1EndDate { get; set; }

        [TitleName("Exotic2_Parameter")]
        public string Exotic2Parameter { get; set; }

        [TitleName("Exotic2_Value")]
        public string Exotic2Value { get; set; }

        [TitleName("Exotic2_Begin_Date")]
        public string Exotic2BeginDate { get; set; }

        [TitleName("Exotic2_End_Date")]
        public string Exotic2EndDate { get; set; }

        [TitleName("Exotic3_Parameter")]
        public string Exotic3Parameter { get; set; }

        [TitleName("Exotic3_Value")]
        public string Exotic3Value { get; set; }

        [TitleName("Exotic3_Begin_Date")]
        public string Exotic3BeginDate { get; set; }

        [TitleName("Exotic3_End_Date")]
        public string Exotic3EndDate { get; set; }

        [TitleName("Exotic4_Parameter")]
        public string Exotic4Parameter { get; set; }

        [TitleName("Exotic4_Value")]
        public string Exotic4Value { get; set; }

        [TitleName("Exotic4_Begin_Date")]
        public string Exotic4BeginDate { get; set; }

        [TitleName("Exotic4_End_Date")]
        public string Exotic4EndDate { get; set; }

        [TitleName("Exotic5_Parameter")]
        public string Exotic5Parameter { get; set; }

        [TitleName("Exotic5_Value")]
        public string Exotic5Value { get; set; }

        [TitleName("Exotic5_Begin_Date")]
        public string Exotic5BeginDate { get; set; }

        [TitleName("Exotic5_End_Date")]
        public string Exotic5EndDate { get; set; }

        [TitleName("Exotic6_Parameter")]
        public string Exotic6Parameter { get; set; }

        [TitleName("Exotic6_Value")]
        public string Exotic6Value { get; set; }

        [TitleName("Exotic6_Begin_Date")]
        public string Exotic6BeginDate { get; set; }

        [TitleName("Exotic6_End_Date")]
        public string Exotic6EndDate { get; set; }

        [TitleName("Event_Type1")]
        public string EventType1 { get; set; }

        [TitleName("Event_Period_Number1")]
        public string EventPeriodNumber1 { get; set; }

        [TitleName("Event_Calendar_Type1")]
        public string EventCalendarType1 { get; set; }

        [TitleName("Event_Frequency1")]
        public string EventFrequency1 { get; set; }

        [TitleName("Event_Type2")]
        public string EventType2 { get; set; }

        [TitleName("Event_Period_Number2")]
        public string EventPeriodNumber2 { get; set; }

        [TitleName("Event_Calendar_Type2")]
        public string EventCalendarType2 { get; set; }

        [TitleName("Event_Frequency2")]
        public string EventFrequency2 { get; set; }

        [TitleName("Exchange_Code1")]
        public string ExchangeCode1 { get; set; }

        [TitleName("Incr_Trade_Lot1")]
        public string IncrTradeLot1 { get; set; }

        [TitleName("Min_Trade_Lot1")]
        public string MinTradelot1 { get; set; }

        [TitleName("Min_Trade_Amount1")]
        public string MinTradeAmount1 { get; set; }

        [TitleName("Exchange_Code2")]
        public string ExchangeCode2 { get; set; }

        [TitleName("Incr_Trade_Lot2")]
        public string IncrTradeLot2 { get; set; }

        [TitleName("Min_Trade_Lot2")]
        public string MinTradeLot2 { get; set; }

        [TitleName("Min_Trade_Amount2")]
        public string MinTradeAmount2 { get; set; }

        [TitleName("Exchange_Code3")]
        public string ExchangeCode3 { get; set; }

        [TitleName("Incr_Trade_Lot3")]
        public string IncrTradeLot3 { get; set; }

        [TitleName("Min_Trade_Lot3")]
        public string MinTradeLot3 { get; set; }

        [TitleName("Min_Trade_Amount3")]
        public string MinTradeAmount3 { get; set; }

        [TitleName("Exchange_Code4")]
        public string ExchangeCode4 { get; set; }

        [TitleName("Incr_Trade_Lot4")]
        public string IncrTradeLot4 { get; set; }

        [TitleName("Min_Trade_Lot4")]
        public string MinTradeLot4 { get; set; }

        [TitleName("Min_Trade_Amount4")]
        public string MinTradeAmount4 { get; set; }

        [TitleName("Attached_To_Id")]
        public string AttachedToId { get; set; }

        [TitleName("Attached_To_Id_Type")]
        public string AttachedToIdType { get; set; }

        [TitleName("Attached_Quantity")]
        public string AttachedQuantity { get; set; }

        [TitleName("Attached_Code")]
        public string AttachedCode { get; set; }

        [TitleName("Detachable_Date")]
        public string DetachableDate { get; set; }

        [TitleName("Bond_Exercise")]
        public string BondExercise { get; set; }

        [TitleName("Bond_Price_Percentage")]
        public string BondPricePercentage { get; set; }
    }
}
