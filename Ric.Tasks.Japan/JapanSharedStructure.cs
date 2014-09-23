using System;
using System.Collections.Generic;

namespace Ric.Tasks.Japan
{
    /// <summary>
    /// Name map 
    /// </summary>
    public class NameMap
    {
        public string OriginalName { get; set; }
        public string JapaneseName { get; set; }
        public string EnglistName { get; set; }
        public string Ric { get; set; }
        public string ShortName { get; set; }
        public NameMap()
        {
            OriginalName = JapaneseName = EnglistName = ShortName = Ric = string.Empty;
        }
    }

    public class CompanyInfo
    {
        public string OriginalName { get; set; }
        public string JapaneseName { get; set; }
        public string EnglishName { get; set; }
        public string ShortEnglishName { get; set; }
        public string Ric { get; set; }
    }

    public class OSETransSingleCompanyInfo
    {
        public CompanyInfo CompanyInfo { get; set; }
        public string TransSum { get; set; }
        public string Volume_OP_INT { get; set; }
        public OSETransSingleCompanyInfo()
        {
            CompanyInfo = new CompanyInfo();
        }
    }

    public class OSETransSector
    {
        public List<OSETransSingleCompanyInfo> SellCompanyInfoList { get; set; }
        public List<OSETransSingleCompanyInfo> BuyCompanyInfoList { get; set; }
        public DateTime Date1 { get; set; }
        public DateTime Date2 { get; set; }
        public string TransSum { get; set; }
        public string StrikePrice { get; set; }

        public OSETransSector()
        {
            SellCompanyInfoList = new List<OSETransSingleCompanyInfo>();
            BuyCompanyInfoList = new List<OSETransSingleCompanyInfo>();
        }
    }

    public class SecurityOptionSector
    {
        public string Type { get; set; }//CALL, PUT
        public List<OSETransSingleCompanyInfo> SellCompanyInfoList { get; set; }
        public List<OSETransSingleCompanyInfo> BuyCompanyInfoList { get; set; }
        public DateTime TradeDate { get; set; }
        public DateTime ContractMonthCode { get; set; }
        public OSETransSingleCompanyInfo TradedCompany { get; set; }
    }
}
