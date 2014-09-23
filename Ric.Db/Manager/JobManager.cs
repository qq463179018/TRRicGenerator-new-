using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;
using Ric.Db.Info;

namespace Ric.Db.Manager
{
    public class JobManager : ManagerBase
    {
        public List<JobInfo> GetAllJobs()
        {
            List<JobInfo> jobList = new List<JobInfo>();

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
                        comm.CommandText = "select * from JobInfo";

                        using (MySqlDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    JobInfo job = new JobInfo
                                    {
                                        JobId = Convert.ToInt32(dr["JobId"]),
                                        JobName = Convert.ToString(dr["JobName"]),
                                        JobSequence = Convert.ToString(dr["JobSequence"]),
                                        MailCCRecipients = Convert.ToString(dr["MailCCRecipients"]),
                                        AssignTo = Convert.ToString(dr["AssignTo"])
                                    };
                                    jobList.Add(job);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }

            return jobList;
        }
    }
}
