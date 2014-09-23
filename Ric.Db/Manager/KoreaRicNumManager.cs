using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Db.Info;
using System.Data.SqlClient;

namespace Ric.Db.Manager
{
    public class KoreaRicNumManager : ManagerBase
    {
        public bool ModifyWarrantAddNumByDate(string date, int warrantAddRicNum)
        {//ETI_Korea_ELW_Number
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
                        comm.CommandText = "update ETI_Korea_ELW_Number set WarrantAddRicNum=@WarrantAddRicNum where Date=@Date";
                        comm.Parameters.Add(new SqlParameter("@WarrantAddRicNum", warrantAddRicNum));
                        comm.Parameters.Add(new SqlParameter("@Date", date));

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

        public bool ModifyWarrantDropNumByDate(string date, int warrantDropRicNum)
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
                        comm.CommandText = "update ETI_Korea_ELW_Number set WarrantDropRicNum=@WarrantDropRicNum where Date=@Date";
                        comm.Parameters.Add(new SqlParameter("@WarrantDropRicNum", warrantDropRicNum));
                        comm.Parameters.Add(new SqlParameter("@Date", date));

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

        public bool ModifyByDate(string date, int warrantAddRicNum, int warrantDropRicNum)
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
                        comm.CommandText = "update ETI_Korea_ELW_Number set  WarrantAddRicNum=@WarrantAddRicNum,WarrantDropRicNum=@WarrantDropRicNum where Date=@Date";
                        comm.Parameters.Add(new SqlParameter("@WarrantAddRicNum", warrantAddRicNum));
                        comm.Parameters.Add(new SqlParameter("@WarrantDropRicNum", warrantDropRicNum));
                        comm.Parameters.Add(new SqlParameter("@Date", date));

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

        public bool Insert(KoreaRicNumInfo info)
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
                        comm.CommandText = "insert into ETI_Korea_ELW_Number(Date,WarrantAddRicNum,WarrantDropRicNum) values(@Date,@AddRicNum,@DropRicNum)";
                        comm.Parameters.Add(new SqlParameter("@Date", info.LauchDate));
                        comm.Parameters.Add(new SqlParameter("@AddRicNum", info.WarrantAddRicNum));
                        comm.Parameters.Add(new SqlParameter("@DropRicNum", info.WarrantDropRicNum));

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
                using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlCommand comm = new SqlCommand())
                    {
                        comm.Connection = conn;
                        comm.CommandText = "delete from ETI_Korea_ELW_Number where  Date=@Date";
                        comm.Parameters.Add(new SqlParameter("@Date", date));
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

        public static KoreaRicNumInfo GetByDate(string date)
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
                        comm.CommandText = "select * from ETI_Korea_ELW_Number where Date=@Date";
                        comm.Parameters.Add(new SqlParameter("@Date", date));

                        using (SqlDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows && dr.Read())
                            {
                                KoreaRicNumInfo info = new KoreaRicNumInfo();
                                info.LauchDate = Convert.ToString(dr["Date"]);
                                info.WarrantAddRicNum = Convert.ToInt32(dr["WarrantAddRicNum"]);
                                info.WarrantDropRicNum = Convert.ToInt32(dr["WarrantDropRicNum"]);
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

        public List<KoreaRicNumInfo> GetAll()
        {
            List<KoreaRicNumInfo> infoList = new List<KoreaRicNumInfo>();
            try
            {
                using (SqlConnection conn = new SqlConnection(Config.ConnectionString))
                {
                    if (conn.State != System.Data.ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlCommand comm = new SqlCommand("select * from ETI_Korea_ELW_Number", conn))
                    {
                        using (SqlDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    KoreaRicNumInfo info = new KoreaRicNumInfo();
                                    info.LauchDate = Convert.ToString(dr["Date"]);
                                    info.WarrantAddRicNum = Convert.ToInt32(dr["WarrantAddRicNum"]);
                                    info.WarrantDropRicNum = Convert.ToInt32(dr["WarrantDropRicNum"]);
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
