using System;

namespace Ric.Tasks.Taiwan
{

    public class TWWarrantBaseInfo : IComparable<TWWarrantBaseInfo>
    {
        public string WarrantCode { get; set; }
        public string WarrantNameAbb { get; set; }
        public string Type { get; set; }//TWO, TW       

        public int CompareTo(TWWarrantBaseInfo other)
        {
            return this.WarrantCode.CompareTo(other.WarrantCode);
        }       
    }

    public class TWWarrant : TWWarrantBaseInfo
    {
        //KGI
        public string IssueDate { get; set; }
        public string IssuePrice { get; set; }
        public string TargetCode { get; set; }
        public string KGIType { get; set; }
        public string OrigContracePrice { get; set; }
        public string KGIShortrName { set; get; }
        public string NewContactPrice { get; set; }
        public string IssueSum { get; set; }
        public string IssuerOrgName { get; set; }
        public string WarrantType { get; set; }
        public string WarrantEnglishNameAbb { get; set; }
        public string ListingDate { get; set; }
        public string FinalTradingDay { get; set; }
        public string ExpireDay { get; set; }
        public string ContractStartDay { get; set; }
        public string ContractExpireDay { get; set; }
        public string OrigCeilingPrice { get; set; }
        public string OrigLowerPrice { get; set; }
        public string PaymentType { get; set; }
        public string Delegate { get; set; }
        public string ChiEngNameAbb { get; set; }
        public string NewTargetSum { get; set; }
        public bool isCBBC { get; set; }
        public bool isIndex { get; set; }
        public bool isETF { get; set; }
        public string callPut { get; set; }
        public string ChineseShortName { get; set; }
        public string SettlementIndicator { get; set; }

        public TWWarrant()
        {
            this.isCBBC = false;
            this.isIndex = false;
            this.isETF = false;
            this.callPut = "C";
        }   
    }

    public class TWFMTemplate
    {
        public string Ric { get; set; }
        public string IssueDate { get; set; }
        public string IssuePrice { get; set; }
        public string CapPrice { get; set; }
        public string EffectiveDate { get; set; }
        public string DisplayName { get; set; }
        public string OfficialCode { get; set; }
        public string ExchangeSymbol { get; set; }
        public string OffcCode2 { get; set; }
        public string Currency { get; set; }
        public string RecordType { get; set; }
        public string ChainRic { get; set; }
        public string PositionInChain { get; set; }
        public string LotSize { get; set; }
        public string CoiDisplyNmll { get; set; }
        public string CoiSectorChain { get; set; }
        public string BcastRef { get; set; }
        public string WntRatio { get; set; }
        public string StrikePrc { get; set; }
        public string MaturDate { get; set; }
        public string ConvFac { get; set; }
        public string Isin { get; set; }
        public string IDNLongName { get; set; }
        public string IssueClassification { get; set; }
        public string PrimaryListing { get; set; }
        public string OrganisationName { get; set; }
        public string UnderlyingRIC { get; set; }
        public string IssuedCompanyName { get; set; }
        public string LocalSectorClassification { get; set; }
        public string IndexRic { get; set; }
        public string TotalSharesOutstanding { get; set; }
        public string CompositeChainRic { get; set; }
        public string LongLink1 { get; set; }
        public string LongLink2 { get; set; }
        public string LongLink3 { get; set; }
        public string LongLink4 { get; set; }
        public string LongLink5 { get; set; }
        public string LongLink6 { get; set; }
        public string LongLink7 { get; set; }
        public string LongLink8 { get; set; }
        public string LongLink9 { get; set; }
        public string BondType { get; set; }
        public string PutCallInd { get; set; }
        public string ISS_TP_FLG { get; set; }
        public string GEN_TEXT16 { get; set; }
        public string GN_TXT16_2 { get; set; }
        public string Longlink1_Issuer { get; set; }
        public string Longlink2_MenuPage { get; set; }
        public string Gearing { get; set; }
        public string Premium { get; set; }
        public string LONGLINK1_TAS_RIC { get; set; }
        public string LONKLINK2_WT_Chain { get; set; }
        public string LONKLINK3_Tech_Ric { get; set; }
        public string LONKLINK4_ValueAdded_Ric { get; set; }
        public string Issuer_OrgId { get; set; }//new add column 
        public string Units { get; set; }//new add column 
        public string Exercise_Begin_Date { get; set; }//new add column 
        public string Exercise_End_Date { get; set; }//new add column 
        public TWWarrantProperty Properties { get; set; }
        public TWFMTemplate()
        {
            this.Currency = "TWD";
            this.RecordType = "97";
            this.PositionInChain = "by alpha order";
            this.LotSize = "1000 Units";
            this.IssueClassification = "WT";
            this.PrimaryListing = "N/A";
            this.IndexRic = "N/A";
            this.ISS_TP_FLG = "S";
            this.GN_TXT16_2 = "C";
            this.Longlink1_Issuer = "N/A";
            this.Ric = string.Empty;
            this.IssueDate = string.Empty;//发行日期
            this.IssuePrice = string.Empty;
            this.CapPrice = string.Empty;
            this.EffectiveDate = string.Empty;
            this.DisplayName = string.Empty;
            this.OfficialCode = string.Empty;
            this.ExchangeSymbol = string.Empty;
            this.OffcCode2 = string.Empty;
            this.ChainRic = string.Empty;
            this.CoiDisplyNmll = string.Empty;
            this.CoiSectorChain = string.Empty;
            this.BcastRef = string.Empty;
            this.WntRatio = string.Empty;
            this.StrikePrc = string.Empty;
            this.MaturDate = string.Empty;
            this.ConvFac = string.Empty;
            this.Isin = string.Empty;
            this.IDNLongName = string.Empty;
            this.OrganisationName = string.Empty;
            this.UnderlyingRIC = string.Empty;
            this.IssuedCompanyName = string.Empty;
            this.LocalSectorClassification = string.Empty;
            this.TotalSharesOutstanding = string.Empty;
            this.CompositeChainRic = string.Empty;
            this.LongLink1 = string.Empty;
            this.LongLink2 = string.Empty;
            this.LongLink3 = string.Empty;
            this.LongLink4 = string.Empty;
            this.LongLink5 = string.Empty;
            this.LongLink6 = string.Empty;
            this.LongLink7 = string.Empty;
            this.LongLink8 = string.Empty;
            this.LongLink9 = string.Empty;
            this.BondType = string.Empty;
            this.PutCallInd = string.Empty;
            this.GEN_TEXT16 = string.Empty;
            this.Longlink2_MenuPage = string.Empty;
            this.Gearing = string.Empty;
            this.Premium = string.Empty;
            this.LONGLINK1_TAS_RIC = string.Empty;
            this.LONKLINK2_WT_Chain = string.Empty;
            this.LONKLINK3_Tech_Ric = string.Empty;
            this.LONKLINK4_ValueAdded_Ric = string.Empty;
            this.Issuer_OrgId = string.Empty;//new add column
            this.Properties = new TWWarrantProperty();
        }
    }

    public class TWWarrantProperty
    {
        public bool IsTWO { get; set; }
        public bool IsCBBC { get; set; }
        public bool IsIndex { get; set; }

        public TWWarrantProperty()
        {
            this.IsTWO = false;
            this.IsCBBC = false;
            this.IsIndex = false;
        }
    }
}
