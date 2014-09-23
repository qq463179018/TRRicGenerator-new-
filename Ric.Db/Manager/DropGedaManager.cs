using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace Ric.Db.Manager
{
    public class DropGedaManager : ManagerBase
    {
        const string ETI_DROP_GEDA_TABLE_NAME = "ETI_DROP_GEDA";

        public static List<string> SelectDrop(string effectiveDate)
        {
            DataTable dt = Select(ETI_DROP_GEDA_TABLE_NAME, new string[] { "*" }, "where EffectiveDate ='" + effectiveDate + "'");
            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }
            return (from DataRow r in dt.Rows select Convert.ToString(r["RIC"]).Trim()).ToList();
        }

        public static int UpdateDrop(string effectiveDate, string ric, string taskId)
        {
            DataTable dt = Select(ETI_DROP_GEDA_TABLE_NAME, new string[] { "*" }, "where ric ='" + ric + "' and TaskId = " + taskId);
            if (dt == null)
            {
                return 0;
            }

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow r in dt.Rows)
                {
                    r["EffectiveDate"] = effectiveDate;
                }
            }
            else
            {
                DataRow row = dt.NewRow();
                row["RIC"] = ric;
                row["TaskId"] = taskId;
                row["EffectiveDate"] = effectiveDate;
                dt.Rows.Add(row);
            }

            return UpdateDbTable(dt, ETI_DROP_GEDA_TABLE_NAME);
        }
        
        public static int DeleteDrop(string effectiveDate)
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
                    comm.CommandText = "delete from "+ ETI_DROP_GEDA_TABLE_NAME +" where EffectiveDate = '" + effectiveDate + "'";                                   
                    int rowAffected = comm.ExecuteNonQuery();
                    return rowAffected;
                }
            }
        }
    }
}
