using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using Ric.Db.Info;


namespace Ric.Db.Manager
{
    public class TWUnderlyingNameManager : ManagerBase
    {
        private const string ETI_TW_UNDERLYING_TABLE_NAME = "ETI_TW_UNDERLYING_NAME";

        public static TWUnderlyingNameInfo GetByChiEngName(string chineseDisplay)
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
                    comm.CommandText = string.Format("select * from ETI_TW_UNDERLYING_NAME where ChineseDisplay=N'{0}' or ChineseChain=N'{0}' ", chineseDisplay.Trim());                   

                    using (SqlDataReader dr = comm.ExecuteReader())
                    {
                        if (dr.HasRows && dr.Read())
                        {
                            TWUnderlyingNameInfo info = new TWUnderlyingNameInfo();
                            info.ChineseChain = Convert.ToString(dr["ChineseChain"]);
                            info.EnglishDisplay = Convert.ToString(dr["EnglishDisplay"]);
                            info.OrganizationName = Convert.ToString(dr["OrganizationName"]);
                            info.ChineseDisplay = Convert.ToString(dr["ChineseDisplay"]);
                            info.UnderlyingRIC = Convert.ToString(dr["Code"]);
                            return info;
                        }

                        else
                        {
                            return null;
                            //throw new Exception(string.Format("Cannot find TWUnderlyingNameInfo object with ChiENgName: {0} in Table TWUnderlyingName", chineseDisplay));
                        }
                    }
                }
            }
        }

        public static TWUnderlyingNameInfo GetByRIC(string itemRIC)
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
                    comm.CommandText = "select * from ETI_TW_UNDERLYING_NAME where UnderlyingRIC=" + itemRIC.Trim();

                    using (SqlDataReader dr = comm.ExecuteReader())
                    {
                        if (dr.HasRows && dr.Read())
                        {
                            TWUnderlyingNameInfo info = new TWUnderlyingNameInfo();
                            info.ChineseChain = Convert.ToString(dr["ChineseChain"]);
                            info.EnglishDisplay = Convert.ToString(dr["EnglishDisplay"]);
                            info.OrganizationName = Convert.ToString(dr["OrganizationName"]);
                            info.UnderlyingRIC = Convert.ToString(dr["Code"]);
                            info.ChineseDisplay = Convert.ToString(dr["ChineseDisplay"]);
                            return info;
                        }

                        else
                        {
                            return null;
                            //throw new Exception(string.Format("Cannot find TWUnderlyingNameInfo object with UnderlyingRIC: {0} in Table TWUnderlyingName", itemRIC));
                        }
                    }
                }
            }
        }

        public static bool ExistUnderlying(string chineseDisplay)
        {
            if (string.IsNullOrEmpty(chineseDisplay))
            {
                return false;
            }
            string where = string.Format("where ChineseDisplay = N'{0}' or ChineseChain = N'{0}'", chineseDisplay);
            System.Data.DataTable dt = Select(ETI_TW_UNDERLYING_TABLE_NAME, new string[] { "*" }, where);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static int InsertNewUnderlying(TWUnderlyingNameInfo underlying)
        {
            if (underlying == null)
            {
                return 0;
            }
            string where = string.Format("where Code = '{0}'", underlying.UnderlyingRIC);
            System.Data.DataTable dt = Select(ETI_TW_UNDERLYING_TABLE_NAME, new string[] { "*" }, where);
            if (dt == null || dt.Rows.Count > 0)
            {
                return 0;
            }
            System.Data.DataRow dr = dt.NewRow();
            dr["Code"] = underlying.UnderlyingRIC;
            dr["OrganizationName"] = underlying.OrganizationName;
            dr["EnglishDisplay"] = underlying.EnglishDisplay;
            dr["ChineseDisplay"] = underlying.ChineseDisplay;
            dr["ChineseChain"] = underlying.ChineseChain;
       
            dt.Rows.Add(dr);
            return UpdateDbTable(dt, ETI_TW_UNDERLYING_TABLE_NAME);
        }

        public static TWUnderlyingNameInfo GetByCode(string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }    
            int codeNum;
            if (!(int.TryParse(code, out codeNum)))
            {
                return null;
            }

            string where = string.Format("where Code like '{0}.%'", code);
            System.Data.DataTable dt = Select(ETI_TW_UNDERLYING_TABLE_NAME, new string[] { "*" }, where);
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            System.Data.DataRow dr = dt.Rows[0];
            TWUnderlyingNameInfo info = new TWUnderlyingNameInfo();
            info.ChineseChain = Convert.ToString(dr["ChineseChain"]);
            info.EnglishDisplay = Convert.ToString(dr["EnglishDisplay"]);
            info.OrganizationName = Convert.ToString(dr["OrganizationName"]);
            info.UnderlyingRIC = Convert.ToString(dr["Code"]);
            info.ChineseDisplay = Convert.ToString(dr["ChineseDisplay"]);
            return info;
        }
    }
}
