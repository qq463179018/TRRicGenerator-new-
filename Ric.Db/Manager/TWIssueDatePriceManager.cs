using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Info;
using System.Data.SqlClient;

namespace Ric.Db.Manager
{
    public class TWIssueDatePriceManager:ManagerBase
    {
        public static TWIssueDatePriceInfo GetByWarrantName(string warrantName)
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
                    comm.CommandText = "select * from ETI_TW_ISSUE_DATE_PRICE where WarrantName= N'"+ warrantName +"'";
                   
                    using (SqlDataReader dr = comm.ExecuteReader())
                    {
                        if (dr.HasRows && dr.Read())
                        {
                            TWIssueDatePriceInfo info = new TWIssueDatePriceInfo();
                            info.IssueDate = Convert.ToString(dr["IssueDate"]);
                            info.IssuePrice = Convert.ToString(dr["IssuePrice"]);
                            info.ShortName = Convert.ToString(dr["ShortName"]);
                            info.WarrantName = warrantName;
                            return info;
                        }

                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }

        public static bool Insert(TWIssueDatePriceInfo issueDateIssuePriceInfo)
        {
            try
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
                        comm.CommandText = "insert into ETI_TW_ISSUE_DATE_PRICE(WarrantName,ShortName,IssueDate,IssuePrice) values(@WarrantName,@ShortName,@IssueDate,@IssuePrice)";
                        comm.Parameters.Add(new SqlParameter("@WarrantName", issueDateIssuePriceInfo.WarrantName));
                        comm.Parameters.Add(new SqlParameter("@ShortName", issueDateIssuePriceInfo.ShortName));
                        comm.Parameters.Add(new SqlParameter("@IssueDate", issueDateIssuePriceInfo.IssueDate));
                        comm.Parameters.Add(new SqlParameter("@IssuePrice", issueDateIssuePriceInfo.IssuePrice));
                        int rowAffected = comm.ExecuteNonQuery();

                        if (rowAffected == 0)
                        {
                            return false;
                        }

                        return true;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool Update(TWIssueDatePriceInfo issueDateIssuePriceInfo)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlCommand comm = new SqlCommand())
                    {
                        string warrantName = issueDateIssuePriceInfo.WarrantName;
                        comm.Connection = conn;
                        comm.CommandText = "update ETI_TW_ISSUE_DATE_PRICE set ShortName=@ShortName ,IssueDate=@IssueDate,IssuePrice=@IssuePrice where WarrantName= N'" + warrantName + "'";
                        comm.Parameters.Add(new SqlParameter("@ShortName", issueDateIssuePriceInfo.ShortName));
                        comm.Parameters.Add(new SqlParameter("@IssueDate", issueDateIssuePriceInfo.IssueDate));
                        comm.Parameters.Add(new SqlParameter("@IssuePrice", issueDateIssuePriceInfo.IssuePrice));
                       
                        int rowAffected = comm.ExecuteNonQuery();

                        if (rowAffected == 0)
                        {
                            return false;
                        }

                        return true;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static List<TWIssueDatePriceInfo> GetAll()
        {
            List<TWIssueDatePriceInfo> infoList = new List<TWIssueDatePriceInfo>();

            try
            {
                using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlCommand comm = new SqlCommand("select * from ETI_TW_ISSUE_DATE_PRICE", conn))
                    {
                        using (SqlDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    TWIssueDatePriceInfo info = new TWIssueDatePriceInfo();
                                    info.WarrantName = Convert.ToString(dr["WarrantName"]);
                                    info.ShortName = Convert.ToString(dr["ShortName"]);
                                    info.IssueDate = Convert.ToString(dr["IssueDate"]);
                                    info.IssuePrice = Convert.ToString(dr["IssuePrice"]);
                                    infoList.Add(info);
                                }
                            }
                        }
                    }
                }

                return infoList;
            }
            catch (Exception)
            {
                return null;
            }
        }

    }
}
