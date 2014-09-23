using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.ComponentModel;
using Ric.Db.Manager;
using System.Data;
using System.IO;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
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
        private KoreaDropGedaConfig configObj = null;

        protected override void Initialize()
        {
            configObj = Config as KoreaDropGedaConfig;
            if (string.IsNullOrEmpty(configObj.BulkFile))
            {
                configObj.BulkFile = GetOutputFilePath();
            }

            AddResult("Log",Logger.FilePath,"Log");           
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
            List<string> rics = new List<string>();
          
                string tablename = string.Format("fn_GetEtiKoreaDropInfo('{0}')", DateTime.Today.ToString("yyyy-MM-dd"));
                System.Data.DataTable dt = ManagerBase.Select(tablename, new string[] { "RIC" }, "order by InstrumentType");
                if (dt == null)
                {
                    string msg = "Error found in getting today's drop records from database. Please check the database.";
                    Logger.Log(msg, Logger.LogType.Error);
                    throw new Exception(msg);
                    //Error
                }
               
                foreach (DataRow dr in dt.Rows)
                {
                    string ric = dr["RIC"].ToString().Replace("D", "").Replace("^", "").Trim();
                    rics.Add(ric);
                }

                return rics; 
        }

        private void GenerateGedaFile(List<string> rics)
        {
            try
            {
                string fileName = string.Format("KR_DROP_{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
               
                string filePath = Path.Combine(configObj.BulkFile, fileName);

                FileUtil.WriteOutputFile(filePath, rics, "RIC", WriteMode.Overwrite);

                AddResult("Output Folder",configObj.BulkFile,"Output Folder");
           
                AddResult(fileName,filePath,fileName);

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
