using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace Ric.Db.Manager
{
    public class RightsTemplate
    {
        public string AddEffectiveDate { get; set; }
        public string DropEffectiveDate { get; set; }
        public string Ticker { get; set; }
        public string RIC { get; set; }
        public string Currency { get; set; }
        public string QACommonName { get; set; }
        public string QAShortName { get; set; }
        public string KoreaCode { get; set; }
        public string ISIN { get; set; }
        public string CountryHeadquarters { get; set; }
        public string LegalName { get; set; }
        public string KoreaName { get; set; }
        public string Edcoid { get; set; }
        public string OldMSCI { get; set; }
        public string RBSS { get; set; }
        public string KoreaScheme { get; set; }
        public string QuantityOfRights { get; set; }
        public string StrikePrice { get; set; }
        public string RecordType { get; set; }
        public string KOSPIChainRIC { get; set; }
        public string PositionInChain { get; set; }
        public string IssueClassification { get; set; }
        public string LotSize { get; set; }
        /// <summary>
        /// Temporary variables 
        /// </summary>
        public string TempVar { get; set; }

        public RightsTemplate()
        {
            Ticker = string.Empty;
            RIC = string.Empty;
            Currency = string.Empty;
            QACommonName = string.Empty;
            QAShortName = string.Empty;
            KoreaCode = string.Empty;
            ISIN = string.Empty;
            CountryHeadquarters = string.Empty;
            LegalName = string.Empty;
            KoreaName = string.Empty;
            Edcoid = string.Empty;
            OldMSCI = string.Empty;
            RBSS = string.Empty;
            KoreaScheme = string.Empty;
            QuantityOfRights = string.Empty;
            StrikePrice = string.Empty;
            RecordType = string.Empty;
            KOSPIChainRIC = string.Empty;
            PositionInChain = string.Empty;
            IssueClassification = string.Empty;
            LotSize = string.Empty;
        }
    }

    public class KoreaRightsManager : ManagerBase
    {
        private const string ETI_KOREA_RIGHTS_TABLE_NAME = "ETI_Korea_Rights";

        public static int UpdateRights(RightsTemplate right)
        {
            string condition = string.Format("where RIC = '{0}' and EffectiveDateAdd = '{1}'", right.RIC, right.AddEffectiveDate);
            DataTable dt = Select(ETI_KOREA_RIGHTS_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null)
            {
                return 0;
            }

            string updateDate = DateTime.Now.ToString("yyyy-MM-dd");
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    dr["UpdateDate"] = updateDate;
                    dr["EffectiveDateAdd"] = right.AddEffectiveDate;
                    dr["EffectiveDateDrop"] = right.DropEffectiveDate;
                    dr["RIC"] = right.RIC;
                    dr["Currency"] = right.Currency;
                    dr["QACommonName"] = right.QACommonName;
                    dr["QAShortName"] = right.QAShortName;
                    dr["KoreaCode"] = right.KoreaCode;
                    dr["ISIN"] = right.ISIN;
                    dr["CountryHeadquarters"] = right.CountryHeadquarters;
                    dr["LegalName"] = right.LegalName;
                    dr["KoreaName"] = right.KoreaName;
                    dr["Edcoid"] = right.Edcoid;
                    dr["OldMSCI"] = right.OldMSCI;
                    dr["RBSS"] = right.RBSS;
                    dr["KoreaScheme"] = right.KoreaScheme;
                    dr["QuantityOfRights"] = right.QuantityOfRights;
                    dr["StrikePrice"] = right.StrikePrice;
                    dr["RecordType"] = right.RecordType;
                    dr["KOSPIChainRIC"] = right.KOSPIChainRIC;
                    dr["PositionInChain"] = right.PositionInChain;
                    dr["IssueClassification"] = right.IssueClassification;
                    dr["LotSize"] = right.LotSize;
                }
            }
            else
            {
                DataRow dr = dt.NewRow();
                dr["UpdateDate"] = updateDate;
                dr["EffectiveDateAdd"] = right.AddEffectiveDate;
                dr["EffectiveDateDrop"] = right.DropEffectiveDate;
                dr["RIC"] = right.RIC;
                dr["Currency"] = right.Currency;
                dr["QACommonName"] = right.QACommonName;
                dr["QAShortName"] = right.QAShortName;
                dr["KoreaCode"] = right.KoreaCode;
                dr["ISIN"] = right.ISIN;
                dr["CountryHeadquarters"] = right.CountryHeadquarters;
                dr["LegalName"] = right.LegalName;
                dr["KoreaName"] = right.KoreaName;
                dr["Edcoid"] = right.Edcoid;
                dr["OldMSCI"] = right.OldMSCI;
                dr["RBSS"] = right.RBSS;
                dr["KoreaScheme"] = right.KoreaScheme;
                dr["QuantityOfRights"] = right.QuantityOfRights;
                dr["StrikePrice"] = right.StrikePrice;
                dr["RecordType"] = right.RecordType;
                dr["KOSPIChainRIC"] = right.KOSPIChainRIC;
                dr["PositionInChain"] = right.PositionInChain;
                dr["IssueClassification"] = right.IssueClassification;
                dr["LotSize"] = right.LotSize;
                dt.Rows.Add(dr);
            }
            return UpdateDbTable(dt, ETI_KOREA_RIGHTS_TABLE_NAME);
        }

        public static int UpdateRights(List<RightsTemplate> rights)
        {
            if (rights == null || rights.Count == 0)
            {
                return 0;
            }
            int i = 0;
            foreach (RightsTemplate item in rights)
            {
                i = i + UpdateRights(item);
            }
            return i;
        }
    }
}
