using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Ric.Db.Info;
using System.Data.SqlClient;

namespace Ric.Db.Manager
{
    public class KoreaEquityManager : ManagerBase
    {
        const string ETI_KOREA_EQUITY_TABLE_NAME = "ETI_Korea_Equity";
        
        public static bool ExistsFMOne(string ticker)
        {
            string condition = string.Format("where Ticker = '{0}' and FM = '1'", ticker);
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static bool ExsitsFMTwo(string ric, string isin)
        {
            string condition = string.Format("where RIC = '{0}' and ISIN = '{1}' and FM = '2' and Status = 'Active'", ric, isin);
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Check if the table contains a same display name for an equity.
        /// </summary>
        /// <param name="displayName">display name</param>
        /// <returns>true or false</returns>
        public static bool ExistDisplayName(string displayName, string ric)
        {
            if (ric.Contains("."))
            {
                ric = ric.Split('.')[0];
            }
            string condition = @"where IDNDisplayName = '" + displayName + "' and Ticker <> '" + ric + "'";
            DataTable dt = ManagerBase.Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Check if the table contains a same korea name for an equity.
        /// </summary>
        /// <param name="koreaName">korea name</param>
        /// <returns>true or false</returns>
        public static bool ExistKoreaName(string koreaName, string ric)
        {
            string condition = @"where KoreaName = '" + koreaName + "' and RIC <> '" + ric + "'";
            DataTable dt = ManagerBase.Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static int UpdateEquity(KoreaEquityInfo equity)
        {
            string condition = string.Format("where RIC = '{0}' and FM = '{1}' and Status = 'Active'", equity.RIC, equity.FM);           
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null)
            {
                return 0;
            }
            string effectiveDate = equity.EffectiveDate;
            if (equity.EffectiveDate.Length == 4)
            {
                effectiveDate = equity.EffectiveDate + "-01-01";
            }

            if (dt.Rows.Count > 0)
            {

                foreach (DataRow dr in dt.Rows)
                {
                    dr["UpdateDate"] = equity.UpdateDate;
                    dr["EffectiveDateAdd"] = effectiveDate;
                    dr["RIC"] = equity.RIC;
                    dr["Type"] = equity.Type;
                    dr["RecordType"] = equity.RecordType;
                    dr["IDNDisplayName"] = equity.IDNDisplayName;
                    dr["ISIN"] = equity.ISIN;
                    dr["BCAST_REF"] = equity.BcastRef;
                    dr["LegalName"] = equity.LegalName;
                    dr["KoreaName"] = equity.KoreaName;
                    dr["LotSize"] = equity.Lotsize;                    
                    dr["Status"] = equity.Status;
                }
            }
            else
            {
                DataRow dr = dt.NewRow();
                dr["UpdateDate"] = equity.UpdateDate;
                dr["EffectiveDateAdd"] = effectiveDate;
                dr["RIC"] = equity.RIC;
                dr["Type"] = equity.Type;
                dr["RecordType"] = equity.RecordType;
                dr["IDNDisplayName"] = equity.IDNDisplayName;
                dr["ISIN"] = equity.ISIN;
                dr["BCAST_REF"] = equity.BcastRef;
                dr["LegalName"] = equity.LegalName;
                dr["KoreaName"] = equity.KoreaName;
                dr["LotSize"] = equity.Lotsize;                
                dr["Ticker"] = equity.Ticker;
                dr["FM"] = equity.FM;
                dr["Status"] = equity.Status;
                dt.Rows.Add(dr);
            }
            return UpdateDbTable(dt, ETI_KOREA_EQUITY_TABLE_NAME);
        }

        public static int UpdateEquity(List<KoreaEquityInfo> equity)
        {
            if (equity == null || equity.Count == 0)
            {
                return 0;
            }
            int i = 0;
            foreach (KoreaEquityInfo item in equity)
            {
                i = i + UpdateEquity(item);
            }
            return i;
        }

        public static KoreaEquityInfo SelectEquityFMOne(string ric)
        {
            string condition = string.Format("where RIC = '{0}' and FM = 1", ric);
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            DataRow dr = dt.Rows[0];
            KoreaEquityInfo itemFm1 = new KoreaEquityInfo();
            itemFm1.LegalName = Convert.ToString(dr["LegalName"]).Trim();
            itemFm1.KoreaName = Convert.ToString(dr["KoreaName"]).Trim();
            itemFm1.IDNDisplayName = Convert.ToString(dr["IDNDisplayName"]).Trim();     
            itemFm1.ISIN = Convert.ToString(dr["ISIN"]).Trim();
            return itemFm1;
        }

        public static List<KoreaEquityInfo> SelectEquityByEffectiveDateChange(string effectiveDate)
        {
            string condition = string.Format("where EffectiveDateChange = '{0}' and FM = '2'", effectiveDate);
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            List<KoreaEquityInfo> equities = new List<KoreaEquityInfo>();

            foreach (DataRow dr in dt.Rows)
            {
                KoreaEquityInfo item = new KoreaEquityInfo();
                item.RIC = Convert.ToString(dr["RIC"]).Trim();
                item.Type = Convert.ToString(dr["Type"]).Trim();
                item.LegalName = Convert.ToString(dr["LegalName"]).Trim();
                item.KoreaName = Convert.ToString(dr["KoreaName"]).Trim();
                item.IDNDisplayName = Convert.ToString(dr["IDNDisplayName"]).Trim();
                item.ISIN = Convert.ToString(dr["ISIN"]).Trim();
                equities.Add(item);
            }
            return equities;
        }

        public static KoreaEquityInfo SelectEquityByIsin(string isin)
        {
            string condition = string.Format("where ISIN = '{0}' and FM ='2' and Status = 'Active'", isin);
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            DataRow dr = dt.Rows[0];
            KoreaEquityInfo item = new KoreaEquityInfo();
            item.RIC = Convert.ToString(dr["RIC"]).Trim();
            item.LegalName = Convert.ToString(dr["LegalName"]).Trim();
            item.KoreaName = Convert.ToString(dr["KoreaName"]).Trim();
            item.IDNDisplayName = Convert.ToString(dr["IDNDisplayName"]).Trim();
            item.Type = Convert.ToString(dr["Type"]).Trim();
            item.ISIN = isin;
            return item;        
        }

        public static bool DelistedEquity(string isin)
        {
            string condition = string.Format("where ISIN = '{0}' and Status = 'De-Active'", isin);
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static KoreaEquityInfo SelectEquityByRic(string ric)
        {
            string condition = string.Format("where RIC = '{0}' and FM ='2' and Status = 'Active'", ric);
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            DataRow dr = dt.Rows[0];
            KoreaEquityInfo item = new KoreaEquityInfo();
            item.ISIN = Convert.ToString(dr["ISIN"]).Trim();
            item.LegalName = Convert.ToString(dr["LegalName"]).Trim();            
            item.IDNDisplayName = Convert.ToString(dr["IDNDisplayName"]).Trim();           
            item.Type = Convert.ToString(dr["Type"]).Trim();
            item.RIC = ric;
            return item;        
        }

        public static int DeleteEquityFMOne(string ric)
        {
            using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
            {
                if (conn.State != System.Data.ConnectionState.Open)
                {
                    conn.Open();
                }

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    comm.CommandText = string.Format("delete from {0} where RIC = '{1}' and FM = '1'", ETI_KOREA_EQUITY_TABLE_NAME, ric);
                    int rowAffected = comm.ExecuteNonQuery();
                    return rowAffected;
                }
            }
        }

        public static int UpdateNameChange(string updateSql)
        {
            using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
            {
                if (conn.State != System.Data.ConnectionState.Open)
                {
                    conn.Open();
                }

                using (SqlCommand comm = new SqlCommand())
                {
                    comm.Connection = conn;
                    comm.CommandText = string.Format("update {0} set {1}", ETI_KOREA_EQUITY_TABLE_NAME, updateSql);
                    int rowAffected = comm.ExecuteNonQuery();
                    return rowAffected;
                }
            }        
        }

        public static int UpdateDrop(string ric, string effectiveDate)
        {
            using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
            {
                if (conn.State != System.Data.ConnectionState.Open)
                {
                    conn.Open();
                }

                using (SqlCommand comm = new SqlCommand())
                {
                    string newRic = @"D^" + ric;
                    string updateDate = DateTime.Today.ToString("yyyy-MM-dd");
                    comm.Connection = conn;
                    comm.CommandText = string.Format("update {0} set RIC = '{1}', EffectiveDateDrop = '{2}', UpdateDateDrop = '{3}', Status = 'De-Active' where RIC = '{4}'", ETI_KOREA_EQUITY_TABLE_NAME, newRic, effectiveDate, updateDate, ric);
                    int rowAffected = comm.ExecuteNonQuery();
                    return rowAffected;
                }
            }   
        }

        public static bool ExistsFmOneCode(string code)
        {
            string condition = string.Format("where LinkCode = '{0}'", code);
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static List<KoreaEquityInfo> SelectEquityByDate(string updateDate)
        {
            string condition = string.Format("where UpdateDate = '{0}'", updateDate);
            DataTable dt = Select(ETI_KOREA_EQUITY_TABLE_NAME, new string[] { "*" }, condition);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            List<KoreaEquityInfo> equities = new List<KoreaEquityInfo>();

            foreach (DataRow dr in dt.Rows)
            {
                KoreaEquityInfo item = new KoreaEquityInfo();
                item.RIC = Convert.ToString(dr["RIC"]).Trim();               
                item.Ticker = Convert.ToString(dr["Ticker"]).Trim();
                item.IDNDisplayName = Convert.ToString(dr["IDNDisplayName"]).Trim();
                item.ISIN = Convert.ToString(dr["ISIN"]).Trim();
                item.BcastRef = Convert.ToString(dr["BCAST_REF"]).Trim();
                equities.Add(item);
            }
            return equities;
        }
    }
}
