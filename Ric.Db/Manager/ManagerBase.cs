using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Ric.Db.Config;
using Ric.Util;

namespace Ric.Db.Manager
{
    public abstract class ManagerBase : IDisposable
    {
        public static readonly string ConfigPath = "DbConfig.xml";
        private static DbConfig _config = null;

        public static DbConfig Config
        {
            get { return _config ?? (_config = ConfigUtil.ReadConfig(ConfigPath, typeof (DbConfig)) as DbConfig); }
            protected set { _config = value; }
        }

        public static DataTable GetDataTable(string tableName)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlCommand comm = new SqlCommand())
                    {
                        comm.Connection = conn;
                        comm.CommandText = "select * from " + tableName;

                        using (SqlDataAdapter da = new SqlDataAdapter(comm))
                        {
                            DataTable retTable = new DataTable(tableName);
                            da.Fill(retTable);
                            return retTable;
                        }
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static DataTable GetDataTable(DbTable table)
        {
            return GetDataTable(table.Name);
        }

        public static int UpdateDbTable(DataTable table, string tableName)
        {
            using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
            {
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }

                using (SqlCommand comm = new SqlCommand("select * from " + tableName, conn))
                {
                    comm.Connection = conn;

                    using (SqlDataAdapter da = new SqlDataAdapter(comm))
                    {
                        SqlCommandBuilder scb = new SqlCommandBuilder(da);
                        return da.Update(table);
                    }
                }
            }

        }

        public static DataTable Select(string tableName, string[] columnNames, string where)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlCommand comm = new SqlCommand())
                    {
                        string commandText = columnNames.Aggregate("select ", (current, column) => current + (column + ","));
                        commandText=commandText.TrimEnd(new[] { ','})+" ";
                        commandText += "from " + tableName;
                        if (!string.IsNullOrEmpty(where))
                        {
                            commandText += " " + where;
                        }
                        comm.Connection = conn;
                        comm.CommandText = commandText;

                        using (SqlDataAdapter da = new SqlDataAdapter(comm))
                        {
                            DataTable retTable = new DataTable(tableName);
                            da.Fill(retTable);
                            return retTable;
                        }
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static DataTable Select(string tableName, string[] columnNames)
        {
            return Select(tableName, columnNames, string.Empty);
        }

        public static DataTable Select(string tableName)
        {
            return Select(tableName, new[]{ "*"}, string.Empty);
        }

        #region IDisposable Members

        public void Dispose()
        {
        }

        #endregion
    }

    enum TableInfoType
    {
        UnderlyingCodeInfo,
        HKRicNumInfo,
        KoreaRicNumInfo
    }
}
