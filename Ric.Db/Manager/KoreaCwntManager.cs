using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace Ric.Db.Manager
{
    public class CompanyWarrantTemplate
    {
        public string EffectiveDate { get; set; }
        public string Ticker { get; set; }
        public string RIC { get; set; }
        public string Currency { get; set; }
        public string QACommonName { get; set; }
        public string QAShortName { get; set; }
        public string ConversionRatio { get; set; }
        public string KoreaCode { get; set; }
        public string ISIN { get; set; }
        public string PreviousExercisePrice { get; set; }
        public string ExercisePrice { get; set; }
        public string OldQuantity { get; set; }
        public string QuantityOfWarrants { get; set; }
        public string ExercisePeriod { get; set; }
        public string ExpiryDate { get; set; }
        public string CountryHeadquarters { get; set; }
        public string LegalName { get; set; }
        public string KoreanName { get; set; }
        public string WarrantStyle { get; set; }
        public string Edcoid { get; set; }
        public string RecordType { get; set; }
        public string IssueClassification { get; set; }
        public string SettlementType { get; set; }
        public string LotSize { get; set; }
        public string IssueDate { get; set; }
        public string IssuerORGID { get; set; }
        public string ForIACommonName { get; set; }
        public string Status { get; set; }
        public List<string> ChangeItems { get; set; }
        public DateTime AnouncementDate { get; set; }
    }

    public class KoreaCwntManager : ManagerBase
    {
        private static string ETI_KOREA_CWNT_TABLE_NAME = "ETI_Korea_CompanyWarrant";

        public static List<CompanyWarrantTemplate> SelectWarrantByEffectiveDateChange(string effectiveDate)
        {
            string condition = string.Format("where EffectiveDateChange = '{0}' and ChangeType = '1'", effectiveDate);
            DataTable dt = Select(ETI_KOREA_CWNT_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            List<CompanyWarrantTemplate> warrants = new List<CompanyWarrantTemplate>();

            foreach (DataRow dr in dt.Rows)
            {
                CompanyWarrantTemplate item = new CompanyWarrantTemplate();
                item.RIC = Convert.ToString(dr["RIC"]).Trim();
                item.ISIN = Convert.ToString(dr["ISIN"]).Trim();
                item.QACommonName = Convert.ToString(dr["QACommonName"]).Trim();             
                item.ExercisePrice = Convert.ToString(dr["ExercisePrice"]).Trim();
                item.ForIACommonName = Convert.ToString(dr["ForIACommonName"]).Trim();                
                warrants.Add(item);
            }
            return warrants;
        }
    }
}
