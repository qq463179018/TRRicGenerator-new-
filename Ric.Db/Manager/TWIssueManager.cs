using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using Ric.Db.Info;


namespace Ric.Db.Manager
{
    public class TWIssueManager : ManagerBase
    {
        private const string ETI_TW_ISSUE_TABLE_NAME = "ETI_TW_ISSUE_INFO";

        public bool ModifyIssuerByChineseName(string ChineseName, string ChineseShortName, string EnglishShortName)
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
                        comm.CommandText = string.Format("update {0} set CHNShort=@ChineseShortName,ENGShort=@EnglishShortName where ChineseName=N'{1}'", ETI_TW_ISSUE_TABLE_NAME, ChineseName); 
                        comm.Parameters.Add(new SqlParameter("@ChineseShortName", ChineseShortName));
                        comm.Parameters.Add(new SqlParameter("@EnglishShortName", EnglishShortName));                       
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

        public static TWIssueInfo GetByEnglishFullName(string EnglishFullName)
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
                    comm.CommandText = "select * from ETI_TW_ISSUE_INFO where EnglishFullName='" + EnglishFullName + "'";
                  
                    using (SqlDataReader dr = comm.ExecuteReader())
                    {
                        if (dr.HasRows && dr.Read())
                        {
                            TWIssueInfo info = new TWIssueInfo();                           
                            info.ChineseShortName = Convert.ToString(dr["ChineseShortName"]);
                            info.EnglishShortName = Convert.ToString(dr["EnglishShortName"]);
                            info.EnglishBriefName = Convert.ToString(dr["EnglishBriefName"]);
                            info.EnglishName = Convert.ToString(dr["EnglishName"]);
                            info.EnglishFullName = Convert.ToString(dr["EnglishFullName"]);                           
                            info.WarrantIssuer = Convert.ToString(dr["WarrantIssuer"]);
                            info.IssueCode = Convert.ToString(dr["IssueCode"]);
                            return info;
                        }

                        else
                        {
                            throw new Exception(string.Format("Cannot find TWUnderlyingNameInfo object with ChiENgName: {0} in Table TWIssuer", EnglishFullName));
                        }
                    }
                }
            }
        }

        public static TWIssueInfo GetByChineseShortName(string ChineseShortName)
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
                        comm.CommandText = "select * from ETI_TW_ISSUE_INFO where ChineseShortName=N'" + ChineseShortName + "'";                       

                        using (SqlDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows && dr.Read())
                            {
                                TWIssueInfo info = new TWIssueInfo();                                
                                info.ChineseShortName = Convert.ToString(dr["ChineseShortName"]);
                                info.EnglishShortName = Convert.ToString(dr["EnglishShortName"]);
                                info.EnglishBriefName = Convert.ToString(dr["EnglishBriefName"]);
                                info.EnglishName = Convert.ToString(dr["EnglishName"]);
                                info.EnglishFullName = Convert.ToString(dr["EnglishFullName"]);                                
                                info.WarrantIssuer = Convert.ToString(dr["WarrantIssuer"]);
                                info.IssueCode = Convert.ToString(dr["IssueCode"]);
                                return info;
                            }

                            else
                            {
                                throw new Exception(string.Format("Cannot find TWUnderlyingNameInfo object with ChiENgName: {0} in Table TWIssuer", ChineseShortName));
                            }
                        }
                    }
                }
            }

        public List<TWIssueInfo> GetAll()
        {
            List<TWIssueInfo> infoList = new List<TWIssueInfo>();
            try
            {
                using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlCommand comm = new SqlCommand("select * from ETI_TW_ISSUE_INFO", conn))
                    {
                        using (SqlDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    TWIssueInfo info = new TWIssueInfo();                                  
                                    info.ChineseShortName = Convert.ToString(dr["ChineseShortName"]);
                                    info.EnglishShortName = Convert.ToString(dr["EnglishShortName"]);
                                    info.EnglishBriefName = Convert.ToString(dr["EnglishBriefName"]);
                                    info.EnglishName = Convert.ToString(dr["EnglishName"]);
                                    info.EnglishFullName = Convert.ToString(dr["EnglishFullName"]);                                    
                                    info.IssueCode = Convert.ToString(dr["IssueCode"]);
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

        public static bool ExistChineseName(string chineseName)
        {
            if (string.IsNullOrEmpty(chineseName))
            {
                return false;
            }
            string where = string.Format("where ChineseShortName = N'{0}'", chineseName);
            System.Data.DataTable dt = Select(ETI_TW_ISSUE_TABLE_NAME, new string[] { "*" }, where);
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        public static int InsertNewIssuer(TWIssueInfo issuer)
        {
            if (issuer == null)
            {
                return 0;
            }
            string where = string.Format("where ChineseShortName = N'{0}'", issuer.ChineseShortName);
            System.Data.DataTable dt = Select(ETI_TW_ISSUE_TABLE_NAME, new string[] { "*" }, where);
            if (dt == null || dt.Rows.Count > 0)
            {
                return 0;
            }
            System.Data.DataRow dr = dt.NewRow();
            dr["ChineseShortName"] = issuer.ChineseShortName;
            dr["EnglishShortName"] = issuer.EnglishShortName;
            dr["EnglishBriefName"] = issuer.EnglishBriefName;
            dr["EnglishName"] = issuer.EnglishName;
            dr["EnglishFullName"] = issuer.EnglishFullName;
            dr["IssueCode"] = issuer.IssueCode;
            dr["WarrantIssuer"] = issuer.WarrantIssuer;
            dt.Rows.Add(dr);
            return UpdateDbTable(dt, ETI_TW_ISSUE_TABLE_NAME);
        }
    }
}
