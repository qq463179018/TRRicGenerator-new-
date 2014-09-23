using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections;
using System.Globalization;
using System.Data.SqlClient;

namespace Ric.Db.Manager
{
    public class WarrantTemplate
    {
        //the update date as the format is DD-MMM-YY
        public string UpdateDate { get; set; }
        //at the year
        public string EffectiveDate { get; set; }
        //tr2_td4 use the number only.skip the first English letter in 단축코드+".KS"
        public string RIC { get; set; }
        //all 1
        public string FM { get; set; }
        //tr4_td2   Select 5 letters of종목영문명 from left, find this 5 letters in excel Issuer, column B, if match, displayed the corresponding letters in excel ISSUER, column C.&The fourth letter to seventh of 단축코드&If주가지수="KOSPI200"displayed the KOSPI,주가지수not available ,find the주권발행기관_1 , use this name to the same value in Underlying ,column A, displayed the corresponding letters in excel Underlying,&The right letter of “종목영문
        public string IDNDisplayName { get; set; }
        //tr2_td2    "표준코드" from website
        public string ISIN { get; set; }
        //tr2_td4    단축코드,Number only
        public string Ticker { get; set; }
        //tr11_td4   주가지수,if not avaliable,find the 주권발행기관_1,and the value in Underlying
        public string BCASTREF { get; set; }
        //,LEFT(K2,4))&" "&RIGHT(종목영문약명,1)&IF(주가지수="KOSPI200","IW","WNT")
        public string QACommonName { get; set; }
        //행사종료일 tr22_td4     행사종료일,change the format
        public string MatDate { get; set; }
        //tr20_td4     행사지수/가격
        public string StrikePrice { get; set; }
        //발행증권수 tr6_td4      발행증권수
        public string QuanityofWarrants { get; set; }
        //발행단가(원) tr6_td2
        public string IssuePrice { get; set; }
        //발행일 change the format tr5_td2
        public string IssueDate { get; set; }
        //전환비율(워런트당) tr7_td2
        public string ConversionRatio { get; set; }
        //종목영문명 upper   tr3_td2(大写)
        public string Issuer { get; set; }
        //종목약명 tr2_td4
        public string KoreaWarrantName { get; set; }
        public string Chain { get; set; }
        public string LastTradingDate { get; set; }
        //tr21_td2 조기종료발생기준가격/지수(원/P)  
        public string KnockOutPrice { get; set; }
        public string OrgMatDate { get; set; }
        public string OrgIssueDate { get; set; }
        public string UnderlyingKoreaName { get; set; }
        public string IssuerKoreaName { get; set; }
        public string CallOrPut { get; set; }
        public bool IsKOBA { get; set; }

        public WarrantTemplate()
        {
            UpdateDate = string.Empty;
            EffectiveDate = string.Empty;
            RIC = string.Empty;
            FM = string.Empty;
            IDNDisplayName = string.Empty;
            ISIN = string.Empty;
            Ticker = string.Empty;
            BCASTREF = string.Empty;
            QACommonName = string.Empty;
            MatDate = string.Empty;
            StrikePrice = string.Empty;
            QuanityofWarrants = string.Empty;
            IssuePrice = string.Empty;
            IssueDate = string.Empty;
            ConversionRatio = string.Empty;
            Issuer = string.Empty;
            KoreaWarrantName = string.Empty;
            Chain = string.Empty;
            KnockOutPrice = string.Empty;
        }

    }

    public class PilcTemplate
    {
        public string QACommonName { set; get; }
        public string IACommonName { set; get; }
        public string ExpiryDate { set; get; }
        public string RIC { get; set; }
        public string PILC { get; set; }
        public string TAG { get; set; }
        public string ShortNameD { get; set; }
        public string StrikePrice { get; set; }
        public string ExercisePriceDeep { get; set; }
        public string ExerciseQuantity { get; set; }
        public string UndlyAssetQuantity { get; set; }
        public string FirstTradingDay { get; set; }
        public string IssueQuantitySOFA { get; set; }
        public string IssuePriceDeep { get; set; }
        public string IssueQuantityDeep { get; set; }
    }

    public class ELWFMDropModel
    {
        public string OrgSource { get; set; }
        public string UpdateDate { get; set; }
        public string EffectiveDate { get; set; }
        public string RIC { get; set; }
        public string Type { get; set; }
        public string IDNDisplayName { get; set; }
        public string ISIN { get; set; }
        public string Ticker { get; set; }
        public string MaturityDate { get; set; }
        public string Comment { get; set; }
        public string Publisher { get; set; }

        /*Temporary Variable*/
        public string Issuername { get; set; }
        public string Num { get; set; }

        public ELWFMDropModel()
        {
            OrgSource = string.Empty;
            RIC = string.Empty;
            Type = string.Empty;
            IDNDisplayName = string.Empty;
            ISIN = string.Empty;
            Ticker = string.Empty;
            Comment = string.Empty;
            Publisher = string.Empty;
        }
    }


    public class KoreaELWManager : ManagerBase
    {
        private const string ETI_KOREA_ELW_TABLE_NAME = "ETI_Korea_ELW";
        private const string ETI_KOREA_TAG_PILC_TABLE_NAME = "ETI_Korea_TagPilc";
        private const string ETI_KOREA_ELW_DROP_TABLE_NAME = "ETI_Korea_ELW_Drop";
        private const string ETI_KOREA_KOBA_TABLE_NAME = "ETI_Korea_KOBA";

        public static int InsertELWBak(List<WarrantTemplate> elws)
        {
            DataTable dt = Select(ETI_KOREA_ELW_TABLE_NAME);
            if (dt == null)
            {
                return 0;
            }

            foreach (WarrantTemplate elw in elws)
            {
                string effectiveDate = elw.EffectiveDate;
                if (elw.EffectiveDate.Length == 4)
                {
                    effectiveDate += "-01-01";
                }

                DataRow dr = dt.NewRow();
                if (!string.IsNullOrEmpty(elw.UpdateDate))
                {
                    dr["UpdateDate"] = elw.UpdateDate;
                }
                if (!string.IsNullOrEmpty(effectiveDate))
                {
                    dr["EffectiveDate"] = effectiveDate;
                }
                dr["RIC"] = elw.RIC;
                dr["FM"] = elw.FM;
                dr["IDNDisplayName"] = elw.IDNDisplayName;
                dr["ISIN"] = elw.ISIN;
                dr["Ticker"] = elw.Ticker;
                dr["BCAST_REF"] = elw.BCASTREF;
                dr["QACommonName"] = elw.QACommonName;
                if (!string.IsNullOrEmpty(elw.MatDate))
                {
                    dr["MatDate"] = elw.MatDate;
                }
                dr["StrikePrice"] = elw.StrikePrice;
                dr["QuantityOfWarrant"] = elw.QuanityofWarrants;
                dr["IssuePrice"] = elw.IssuePrice;
                if (!string.IsNullOrEmpty(elw.IssueDate))
                {
                    dr["IssueDate"] = elw.IssueDate;
                }
                dr["ConversionRatio"] = elw.ConversionRatio;
                dr["Issuer"] = elw.Issuer;
                dr["KoreaWarrantName"] = elw.KoreaWarrantName;
                dr["Chain"] = elw.Chain;
                if (!string.IsNullOrEmpty(elw.LastTradingDate))
                {
                    dr["LastTradingDate"] = elw.LastTradingDate;
                }
                dr["EquityType"] = "ELW";
                dt.Rows.Add(dr);
            }
            return UpdateDbTable(dt, ETI_KOREA_ELW_TABLE_NAME);
        }

        public static int InsertELW(List<WarrantTemplate> elws)
        {
            int updateRows = 0;
            foreach (WarrantTemplate elw in elws)
            {
                updateRows += InsertELW(elw);
            }
            return updateRows;
        }

        public static int InsertELW(WarrantTemplate elw)
        {
            string condition = string.Format("where RIC = '{0}' and FM = '1'", elw.RIC);
            DataTable dt = Select(ETI_KOREA_ELW_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null)
            {
                return 0;
            }
            string effectiveDate = elw.EffectiveDate;
            if (elw.EffectiveDate.Length == 4)
            {
                effectiveDate += "-01-01";
            }

            if (dt.Rows.Count > 0)
            {

                foreach (DataRow dr in dt.Rows)
                {
                    if (!string.IsNullOrEmpty(elw.UpdateDate))
                    {
                        dr["UpdateDate"] = elw.UpdateDate;
                    }
                    if (!string.IsNullOrEmpty(effectiveDate))
                    {
                        dr["EffectiveDate"] = effectiveDate;
                    }
                    dr["RIC"] = elw.RIC;
                    dr["FM"] = elw.FM;
                    dr["IDNDisplayName"] = elw.IDNDisplayName;
                    dr["ISIN"] = elw.ISIN;
                    dr["Ticker"] = elw.Ticker;
                    dr["BCAST_REF"] = elw.BCASTREF;
                    dr["QACommonName"] = elw.QACommonName;
                    if (!string.IsNullOrEmpty(elw.MatDate))
                    {
                        dr["MatDate"] = elw.MatDate;
                    }
                    dr["StrikePrice"] = elw.StrikePrice;
                    dr["QuantityOfWarrant"] = elw.QuanityofWarrants;
                    dr["IssuePrice"] = elw.IssuePrice;
                    if (!string.IsNullOrEmpty(elw.IssueDate))
                    {
                        dr["IssueDate"] = elw.IssueDate;
                    }
                    dr["ConversionRatio"] = elw.ConversionRatio;
                    dr["Issuer"] = elw.Issuer;
                    dr["KoreaWarrantName"] = elw.KoreaWarrantName;
                    dr["Chain"] = elw.Chain;
                    if (!string.IsNullOrEmpty(elw.LastTradingDate))
                    {
                        dr["LastTradingDate"] = elw.LastTradingDate;
                    }
                    dr["EquityType"] = "ELW";
                }
            }
            else
            {
                DataRow dr = dt.NewRow();
                if (!string.IsNullOrEmpty(elw.UpdateDate))
                {
                    dr["UpdateDate"] = elw.UpdateDate;
                }
                if (!string.IsNullOrEmpty(effectiveDate))
                {
                    dr["EffectiveDate"] = effectiveDate;
                }
                dr["RIC"] = elw.RIC;
                dr["FM"] = elw.FM;
                dr["IDNDisplayName"] = elw.IDNDisplayName;
                dr["ISIN"] = elw.ISIN;
                dr["Ticker"] = elw.Ticker;
                dr["BCAST_REF"] = elw.BCASTREF;
                dr["QACommonName"] = elw.QACommonName;
                if (!string.IsNullOrEmpty(elw.MatDate))
                {
                    dr["MatDate"] = elw.MatDate;
                }
                dr["StrikePrice"] = elw.StrikePrice;
                dr["QuantityOfWarrant"] = elw.QuanityofWarrants;
                dr["IssuePrice"] = elw.IssuePrice;
                if (!string.IsNullOrEmpty(elw.IssueDate))
                {
                    dr["IssueDate"] = elw.IssueDate;
                }
                dr["ConversionRatio"] = elw.ConversionRatio;
                dr["Issuer"] = elw.Issuer;
                dr["KoreaWarrantName"] = elw.KoreaWarrantName;
                dr["Chain"] = elw.Chain;
                if (!string.IsNullOrEmpty(elw.LastTradingDate))
                {
                    dr["LastTradingDate"] = elw.LastTradingDate;
                }
                dr["EquityType"] = "ELW";
                dt.Rows.Add(dr);
            }
            return UpdateDbTable(dt, ETI_KOREA_ELW_TABLE_NAME);
        
        }

        public static Hashtable SelectPILC()
        {
            DataTable dt = Select(ETI_KOREA_TAG_PILC_TABLE_NAME);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }

            Hashtable pilcs = new Hashtable();
            foreach (DataRow dr in dt.Rows)
            {
                PilcTemplate pilc = new PilcTemplate();

                if (dr["IACommonNameD"] != null)
                {
                    pilc.IACommonName = Convert.ToString(dr["IACommonNameD"]);
                }
                if (dr["QACommonNameD"] != null)
                {
                    pilc.IACommonName = Convert.ToString(dr["QACommonNameD"]);
                }
                if (dr["PILC"] != null)
                {
                    pilc.IACommonName = Convert.ToString(dr["PILC"]);
                }
                if (dr["RIC"] != null)
                {
                    string ric = Convert.ToString(dr["RIC"]);
                    pilcs.Add(ric, pilc);
                }
            }
            return pilcs;
        }

        public static int InsertELWDrop(List<ELWFMDropModel> elws, string rics)
        {
            string where = string.Format("where RIC in ({0})", rics);

            DataTable dt = Select(ETI_KOREA_ELW_DROP_TABLE_NAME, new string[] { "*" }, where);
            if (dt == null)
            {
                return 0;
            }
            else if (dt.Rows.Count > 0)
            {
                DeleteExistedElwDrop(rics);
            }

            foreach (ELWFMDropModel elw in elws)
            {

                DataRow dr = dt.NewRow();
                dr["DropUpdateDate"] = elw.UpdateDate;
                dr["DropEffectiveDate"] = elw.EffectiveDate;
                dr["RIC"] = elw.RIC;
                dr["Type"] = elw.Type;
                dr["IDNDisplayName"] = elw.IDNDisplayName;
                dr["ISIN"] = elw.ISIN;
                dr["Ticker"] = elw.Ticker;
                if (!string.IsNullOrEmpty(elw.MaturityDate))
                {
                    dr["MaturityDate"] = elw.MaturityDate;
                }
                dr["Comment"] = elw.Comment;
                dr["CompanyName"] = elw.OrgSource;
                dr["Publisher"] = elw.Publisher;
                dt.Rows.Add(dr);

            }
            return UpdateDbTable(dt, ETI_KOREA_ELW_DROP_TABLE_NAME);
        }

        private static void DeleteExistedElwDrop(string rics)
        {
            string sql = string.Format("delete from {0} where RIC in ({1})", ETI_KOREA_ELW_DROP_TABLE_NAME, rics);
            using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
            {
                if (conn.State != System.Data.ConnectionState.Open)
                {
                    conn.Open();
                }

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    comm.CommandText = sql;
                    comm.ExecuteNonQuery();
                }
            }
        }

        public static Hashtable SelectELWFM1(string rics)
        {
            string condition = string.Format("where FM = '1' and RIC in ({0})", rics);
            DataTable dt = Select(ETI_KOREA_ELW_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }

            Hashtable fmOne = new Hashtable();
            foreach (DataRow dr in dt.Rows)
            {
                WarrantTemplate elw = new WarrantTemplate();
                elw.RIC = Convert.ToString(dr["RIC"]);
                elw.IDNDisplayName = Convert.ToString(dr["IDNDisplayName"]);
                elw.ISIN = Convert.ToString(dr["ISIN"]);
                elw.Ticker = Convert.ToString(dr["Ticker"]);
                elw.BCASTREF = Convert.ToString(dr["BCAST_REF"]);
                elw.QACommonName = Convert.ToString(dr["QACommonName"]);
                if (!string.IsNullOrEmpty(Convert.ToString(dr["MatDate"])))
                {
                    elw.MatDate = Convert.ToDateTime(dr["MatDate"]).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                }
                elw.StrikePrice = Convert.ToString(dr["StrikePrice"]);
                elw.QuanityofWarrants = Convert.ToString(dr["QuantityOfWarrant"]);
                elw.IssuePrice = Convert.ToString(dr["IssuePrice"]);
                if (!string.IsNullOrEmpty(Convert.ToString(dr["IssueDate"])))
                {
                    elw.IssueDate = Convert.ToDateTime(dr["IssueDate"]).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                }
                elw.ConversionRatio = Convert.ToString(dr["ConversionRatio"]);
                elw.Issuer = Convert.ToString(dr["Issuer"]);
                elw.KoreaWarrantName = Convert.ToString(dr["KoreaWarrantName"]);
                elw.Chain = Convert.ToString(dr["Chain"]);
                fmOne.Add(elw.RIC, elw);
            }
            return fmOne;
        }

        public static int DeleteELWFM1(string rics)
        {
            string sql = string.Format("delete from {0} where FM = '1' and RIC in ({1})", ETI_KOREA_ELW_TABLE_NAME, rics);
            using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
            {
                if (conn.State != System.Data.ConnectionState.Open)
                {
                    conn.Open();
                }

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    comm.CommandText = sql;
                    return comm.ExecuteNonQuery();
                }
            }
        }

        public static int DeleteELWFM2(string rics)
        {
            string sql = string.Format("delete from {0} where FM = '2' and RIC in ({1})", ETI_KOREA_ELW_TABLE_NAME, rics);
            using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
            {
                if (conn.State != System.Data.ConnectionState.Open)
                {
                    conn.Open();
                }

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    comm.CommandText = sql;
                    return comm.ExecuteNonQuery();
                }
            }
        }

        public static int InsertKOBA(List<WarrantTemplate> koba)
        {
            DataTable dt = Select(ETI_KOREA_ELW_TABLE_NAME);
            if (dt == null)
            {
                return 0;
            }

            foreach (WarrantTemplate item in koba)
            {
                string effectiveDate = item.EffectiveDate;
                if (item.EffectiveDate.Length == 4)
                {
                    effectiveDate += "-01-01";
                }

                DataRow dr = dt.NewRow();
                if (!string.IsNullOrEmpty(item.UpdateDate))
                {
                    dr["UpdateDate"] = item.UpdateDate;
                }
                if (!string.IsNullOrEmpty(effectiveDate))
                {
                    dr["EffectiveDate"] = effectiveDate;
                }
                dr["RIC"] = item.RIC;
                dr["FM"] = item.FM;
                dr["IDNDisplayName"] = item.IDNDisplayName;
                dr["ISIN"] = item.ISIN;
                dr["Ticker"] = item.Ticker;
                dr["BCAST_REF"] = item.BCASTREF;
                dr["QACommonName"] = item.QACommonName;
                if (!string.IsNullOrEmpty(item.MatDate))
                {
                    dr["MatDate"] = item.MatDate;
                }
                dr["StrikePrice"] = item.StrikePrice;
                dr["QuantityOfWarrant"] = item.QuanityofWarrants;
                dr["IssuePrice"] = item.IssuePrice;
                if (!string.IsNullOrEmpty(item.IssueDate))
                {
                    dr["IssueDate"] = item.IssueDate;
                }
                dr["ConversionRatio"] = item.ConversionRatio;
                dr["Issuer"] = item.Issuer;
                dr["KoreaWarrantName"] = item.KoreaWarrantName;
                dr["Chain"] = item.Chain;
                if (!string.IsNullOrEmpty(item.LastTradingDate))
                {
                    dr["LastTradingDate"] = item.LastTradingDate;
                }
                dr["KnockOutPrice"] = item.KnockOutPrice;
                dr["EquityType"] = "KOBA";

                dt.Rows.Add(dr);
            }
            return UpdateDbTable(dt, ETI_KOREA_ELW_TABLE_NAME);
        }



    }
}
