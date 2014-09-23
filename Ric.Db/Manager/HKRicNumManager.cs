using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Info;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;

namespace Ric.Db.Manager
{
    public class HKRicNumManager : ManagerBase
    {
        public bool ModifyCBBCNumByDate(string date, int cbbcRicNum)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (MySqlCommand comm = new MySqlCommand())
                    {
                        comm.Connection = conn;
                        comm.CommandText = "update HKRicNumInfo set CBBCRicNum=@CBBCRicNum where Date=@Date";
                        comm.Parameters.Add(new MySqlParameter("@CBBCRicNum", cbbcRicNum));
                        comm.Parameters.Add(new MySqlParameter("@Date", date));

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

        public bool ModifyWarrantNumByDate(string date, int warrantRicNum)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (MySqlCommand comm = new MySqlCommand())
                    {
                        comm.Connection = conn;
                        comm.CommandText = "update HKRicNumInfo set WarrantRicNum=@WarrantRicNum where Date=@Date";
                        comm.Parameters.Add(new MySqlParameter("@WarrantRicNum", warrantRicNum));
                        comm.Parameters.Add(new MySqlParameter("@Date", date));

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

        public bool ModifyByDate(string date, int cbbcRicNum, int warrantRicNum)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (MySqlCommand comm = new MySqlCommand())
                    {
                        comm.Connection = conn;
                        comm.CommandText = "update HKRicNumInfo set  WarrantRicNum=@WarrantRicNum,CBBCRicNum=@CBBCRicNum where Date=@Date";
                        comm.Parameters.Add(new MySqlParameter("@CBBCRicNum", cbbcRicNum));
                        comm.Parameters.Add(new MySqlParameter("WarrantRicNum", warrantRicNum));
                        comm.Parameters.Add(new MySqlParameter("@Date", date));

                        int rowAffected = comm.ExecuteNonQuery();

                        if (rowAffected == 0)
                        {
                            return false;
                        }

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                string errMsg = ex.Message;
                return false;
            }
        }

        public bool Insert(HKRicNumInfo info)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (MySqlCommand comm = new MySqlCommand())
                    {
                        comm.Connection = conn;
                        comm.CommandText = "insert into HKRicNumInfo(Date,CBBCRicNum,WarrantRicNum) values(@Date,@CBBCRicNum,@WarrantRicNum)";
                        comm.Parameters.Add(new MySqlParameter("@Date", info.LauchDate));
                        comm.Parameters.Add(new MySqlParameter("@CBBCRicNum", info.CBBCRicNum));
                        comm.Parameters.Add(new MySqlParameter("@WarrantRicNum", info.WarrantRicNum));

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

        public bool Delete(string date)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (MySqlCommand comm = new MySqlCommand())
                    {
                        comm.Connection = conn;
                        comm.CommandText = "delete from HKRicNumInfo where  Date=@Date";
                        comm.Parameters.Add(new MySqlParameter("@Date", date));
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

        public HKRicNumInfo GetByDate(string date)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (MySqlCommand comm = new MySqlCommand())
                    {
                        comm.Connection = conn;
                        comm.CommandText = "select * from HKRicNumInfo where Date=@Date";
                        comm.Parameters.Add(new MySqlParameter("@Date", date));

                        using (MySqlDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows && dr.Read())
                            {
                                HKRicNumInfo info = new HKRicNumInfo();
                                info.LauchDate = Convert.ToString(dr["Date"]);
                                info.CBBCRicNum = Convert.ToInt32(dr["CBBCRicNum"]);
                                info.WarrantRicNum = Convert.ToInt32(dr["WarrantRicNum"]);
                                return info;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }

            return null;
        }

        public List<HKRicNumInfo> GetAll()
        {
            List<HKRicNumInfo> infoList = new List<HKRicNumInfo>();
            try
            {
                using (MySqlConnection conn = new MySqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (MySqlCommand comm = new MySqlCommand("select * from HKRicNumInfo", conn))
                    {
                        using (MySqlDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    HKRicNumInfo info = new HKRicNumInfo();
                                    info.LauchDate = Convert.ToString(dr["Date"]);
                                    info.CBBCRicNum = Convert.ToInt32(dr["CBBCRicNum"]);
                                    info.WarrantRicNum = Convert.ToInt32(dr["WarrantRicNum"]);
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
