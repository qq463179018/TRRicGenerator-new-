using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class KoreaDropGedaConfig
    {
        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\Korea_Auto\\Korea_Drop\\")]
        [Description("Folder for saving generated GDEA file.")]
        public string BulkFile { get; set;}
    }

    /// <summary>
    /// This task can get the drop records of today in database and generate GEDA file named as 'KR_DROP_yyyyMMdd.txt'. 
    /// </summary>
    public class KoreaDropGeda : GeneratorBase
    {
        private KoreaDropGedaConfig configObj;

        protected override void Initialize()
        {
            configObj = Config as KoreaDropGedaConfig;
            if (string.IsNullOrEmpty(configObj.BulkFile))
            {
                configObj.BulkFile = GetOutputFilePath();
            }

            TaskResultList.Add(new TaskResultEntry("Log", "Log", Logger.FilePath));           
        }

        protected override void Start()
        {
            List<string> rics = GetTodayDropFromDb();

            string msg = string.Format("{0} RICs will drop today.", rics.Count);
            Logger.Log(msg);

            if (rics.Count > 0)
            {
                GenerateGedaFile(rics);
            }
        }

        private List<string> GetTodayDropFromDb()
        {
            string tablename = string.Format("fn_GetEtiKoreaDropInfo('{0}')", DateTime.Today.ToString("yyyy-MM-dd"));
            DataTable dt = ManagerBase.Select(tablename, new[] {"RIC"}, "order by InstrumentType");
            if (dt == null)
            {
                string msg = "Error found in getting today's drop records from database. Please check the database.";
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception(msg);
            }
            return
                (from DataRow dr in dt.Rows select dr["RIC"].ToString().Replace("D", "").Replace("^", "").Trim()).ToList();
        }

        private void GenerateGedaFile(List<string> rics)
        {
            try
            {
                string fileName = string.Format("KR_DROP_{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
               
                string filePath = Path.Combine(configObj.BulkFile, fileName);

                FileUtil.WriteOutputFile(filePath, rics, "RIC", WriteMode.Overwrite);

                TaskResultList.Add(new TaskResultEntry("Output Folder", "Output Folder", configObj.BulkFile));
           
                TaskResultList.Add(new TaskResultEntry(fileName, fileName, filePath, FileProcessType.GEDA_BULK_CHAIN_RIC_DELETION));

                Logger.Log("Generate drop GEDA file. OK!");
            }
            catch (Exception ex)
            {
                string msg = "Error found in generating GEDA drop file. Error:" + ex.Message;
                Logger.Log(msg, Logger.LogType.Error);
                throw ex;
            }
        }

    }
}
