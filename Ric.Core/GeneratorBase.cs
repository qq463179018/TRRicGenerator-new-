using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using Ric.Core.Events;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;
using Microsoft.Office.Core;

namespace Ric.Core
{
    public class Property
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public abstract class GeneratorBase : IDisposable
    {
        protected Core CoreObj = new Core();
        public List<TaskResultEntry> TaskResultList = new List<TaskResultEntry>();
        public string LogMsg = string.Empty;
        protected Logger Logger;
        protected object Config;
        protected bool IsEikonExcelDisable = false;
        public event EventHandler<LogEventArgs> Log;
        public event EventHandler<ResultEventArgs> Result;
        public event EventHandler<PropEventArgs> Props;

        protected GeneratorBase()
        {
            Logger = new Logger(GetLogFileFullName(), Logger.LogMode.New);

            InitConfig();
            Initialize();
        }

        #region Events

        protected void LogMessage(string message)
        {
            if (Log != null)
            {
                Log(this, new LogEventArgs(message));
            }
            Logger.Log(message);
        }

        protected void LogMessage(string message, Logger.LogType type)
        {
            if (Log != null)
            {
                Log(this, new LogEventArgs(message, type));
            }
            Logger.Log(message, type);
        }

        protected void AddResult(string fileName, string filePath, string fileType)
        {
            if (Result != null)
            {
                Result(this, new ResultEventArgs(fileName, filePath, fileType));
            }
            //LogMessage("New file created");
            TaskResultList.Add(new TaskResultEntry
            {
                Name = fileName,
                Description = fileType,
                Path = filePath
            });
        }

        protected void AddProp(Dictionary<string, string> prop)
        {
            Props(this, new PropEventArgs(prop));
        }

        #endregion

        protected string TaskTypeName
        {
            get
            {
                return GetType().Name;
            }
        }

        public string TaskId
        {
            get;
            private set;
        }

        public string TaskName
        {
            get;
            private set;
        }

        public int MarketId
        {
            get;
            private set;
        }

        public string MarketName
        {
            get;
            private set;
        }

        public virtual bool IsPrerequisiteMeet()
        {
            return true;
        }

        public virtual void OnPrerequisiteNotMeet()
        {
        }

        public virtual void ExecuteUnderPrerequisiteNotMeet()
        {
        }

        protected virtual void Initialize()
        {
        }

        protected virtual void Cleanup()
        {
        }

        protected void InitConfig()
        {
            TaskInfo task = TaskManager.GetTaskByType(GetType().FullName, RunTimeContext.Context.DatabaseContext);

            if (task == null)
            {
                return;
            }

            TaskId = task.TaskId.ToString(CultureInfo.InvariantCulture);
            TaskName = task.TaskName;
            MarketId = task.MarketId;
            MarketName = task.MarketName;

            Config = RunTimeContext.Context.ConfigStore.GetConfig(Convert.ToInt16(TaskId, CultureInfo.InvariantCulture));

            if (Config != null) return;
            Config = ConfigBuilder.CreateConfigInstance(Type.GetType(task.ConfigType), Convert.ToInt16(TaskId, CultureInfo.InvariantCulture));
            RunTimeContext.Context.ConfigStore.StoreConfig(Convert.ToInt16(TaskId, CultureInfo.InvariantCulture), Config);
        }

        protected virtual void Start()
        {
        }

        public void StartGenerator()
        {
            ExcelApp excelApp = null;
            COMAddIn comAddIn = null;
            if (IsEikonExcelDisable)
            {
                SetEikonExcelDisable(ref excelApp, ref comAddIn);
            }

            try
            {
                Start();
                SaveTaskResult();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (IsEikonExcelDisable)
                {
                    SetEikonExcelAble(ref excelApp, ref comAddIn);
                }
                AddResult("Log file", Logger.FilePath, "log");
            }
        }

        private void SetEikonExcelDisable(ref ExcelApp excelApp, ref COMAddIn comAddIn)
        {
            excelApp = new ExcelApp(false, false);
            if (excelApp == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            COMAddIns addIns = excelApp.ExcelAppInstance.COMAddIns;
            foreach (COMAddIn item in addIns)
            {
                //Disable Eikon Excel.
                if (item.ProgId.Trim().Equals("PowerlinkCOMAddIn.COMAddIn"))
                {
                    item.Connect = false;
                    comAddIn = item;
                }
            }
        }

        private void SetEikonExcelAble(ref ExcelApp excelApp, ref COMAddIn comAddIn)
        {
            if (excelApp == null)
            {
                return;
            }
            excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
            excelApp.ExcelAppInstance.DisplayAlerts = false;
            if (comAddIn != null)
            {
                comAddIn.Connect = true;
            }
            excelApp.Dispose();
        }

        public void SaveTaskResult()
        {
            ConfigUtil.WriteConfig(GetResultFileFullName(), TaskResultList);
        }

        public static List<TaskResultEntry> LoadTaskResult(Type type)
        {
            List<TaskResultEntry> result = null;

            try
            {
                result = ConfigUtil.ReadConfig(GetResultFileFullName(type.Name), typeof(List<TaskResultEntry>)) as List<TaskResultEntry>;
            }
            catch
            {
            }

            return result;
        }

        protected string GetResultFileFullName()
        {
            return GetResultFileFullName(TaskTypeName);
        }

        protected static string GetResultFileFullName(string typeName)
        {
            string path = CreateFolderUnderCurrentPath("Result");

            return Path.Combine(path, string.Format(CultureInfo.InvariantCulture, "{0}.xml", typeName));
        }

        protected string GetLogFileFullName()
        {
            string path = CreateFolderUnderCurrentPath(
               string.Format(@"{0}\{1}\{2}", "Log", TaskTypeName, DateTime.Now.ToString("yyyyMMdd")));

            return Path.Combine(path, string.Format("{0}.log", TaskTypeName));
        }

        protected string GetOutputFilePath()
        {
            return CreateFolderUnderCurrentPath(
                string.Format(@"{0}\{1}\{2}", "Output", TaskTypeName, DateTime.Now.ToString("yyyyMMdd")));
        }

        protected static string GetCurrentPath()
        {
            return Thread.GetDomain().BaseDirectory;
        }

        public static string GetAppDataPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ETI_Auto");
        }

        protected static string CreateFolderUnderCurrentPath(string folderName)
        {
            string path = Path.Combine(GetAppDataPath(), folderName);

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }

        #region IDisposable Members

        public virtual void Dispose()
        {
            Cleanup();
        }

        #endregion IDisposable Members
    }

    public class TaskResultEntry
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public string Path { get; set; }

        public FileProcessType ProcessType { get; set; }

        public MailToSend Mail { get; set; }

        public TaskResultEntry()
        {
        }

        public TaskResultEntry(string name, string desc, string path, MailToSend mail)
        {
            if (mail != null && !string.IsNullOrEmpty(mail.MailSubject))
            {
                ProcessType = FileProcessType.SendMail;
            }
            Name = name;
            Description = desc;
            Path = path;
            Mail = mail;
        }

        public TaskResultEntry(string name, string desc, string path, FileProcessType processType)
        {
            Name = name;
            Description = desc;
            Path = path;
            ProcessType = processType;
        }

        public TaskResultEntry(string name, string desc, string path)
        {
            Name = name;
            Description = desc;
            Path = path;
        }

        //public void ProcessFile(string user, string pwd, bool ifTest)
        //{
        //    if (ProcessType.ToString().Contains("GEDA"))
        //    {
        //        GEDAConnection.BulkLoadFileToGEDA(Path, ProcessType.ToString().Remove(0, 5), ifTest, user, pwd);
        //    }
        //    else if (ProcessType.ToString().Contains("NDA"))
        //    {
        //        string ftpFilePath = NDAConnection.UploadFileToFTP(Path);
        //    }
        //    else if (ProcessType.ToString().Contains("VAP"))
        //    {
        //    }
        //    else if (ProcessType == FileProcessType.SendMail && Mail != null)
        //    {
        //        string err = string.Empty;
        //        try
        //        {
        //            using (OutlookApp app = new OutlookApp())
        //            {
        //                OutlookUtil.CreateAndSendMail(app, Mail, out err);
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            throw (new Exception("Error happens when sending mail: " + Mail.MailSubject + " Ex: " + ex.Message));
        //        }
        //    }
        //    else
        //    {
        //        throw (new Exception(("No Need to be Processed by Production Team.")));
        //    }
        //}
    }

    public enum FileProcessType
    {
        Other,
        Log,
        NDA,
        SendMail,
        GEDA_BULK_RIC_CHANGE,
        GEDA_BULK_RIC_CREATION,
        GEDA_BULK_RIC_UPDATE,
        GEDA_BULK_RIC_DELETE,
        GEDA_BULK_CHAIN_RIC_CREATION,
        GEDA_BULK_CHAIN_RIC_UPDATE,
        GEDA_BULK_CHAIN_RIC_DELETION,
        GEDA_BULK_CHAIN_RIC_CONS_ADD,
        GEDA_BULK_CHAIN_RIC_CONS_REPLACE,
        GEDA_BULK_CHAIN_RIC_CONS_DELETION,
        GEDA_BULK_SIC_RIC_CHANGE,
        GEDA_BULK_RIC_QUERY,
        GEDA_BULK_SIC_QUERY,
        GEDA_BULK_SYMBOL_QUERY,
        GEDA_BULK_GROUP_MODFIER_CREATION,
        GEDA_BULK_GROUP_MODFIER_UPDATE,
        GEDA_BULK_GROUP_MODFIER_DELETION,
        GEDA_BULK_CHAIN_GROUP_MOD_CREATION,
        GEDA_BULK_CHAIN_GROUP_MOD_UPDATE,
        GEDA_BULK_CHAIN_GROUP_MOD_DELETION,
        GEDA_BULK_COPY_EXCHANGE_LIST,
        GEDA_BULK_COPY_CHAIN_EXCHANGE_LIST,
        GEDA_BULK_EXCHANGE_LIST_UPDATE,
        GEDA_BULK_EXCHANGE_LIST_DELETION,
        GEDA_RIC_IN_GEDA,
        GEDA_RIC_NOT_IN_MDI,
        GEDA_RIC_NOT_IN_GQS,
        GEDA_RIC_IN_RA,
        GEDA_RIC_NOT_IN_RA,
        GEDA_BULK_RIC_MOVE,
        GEDA_BULK_LOAD_BCU,
        GEDA_BULK_UPDATE_HD,
        GEDA_BULK_DELETE_HD,
        GEDA_BULK_BCU_DELETE,
        GEDA_BULK_BCU_CON_REPLACE,
        GEDA_BULK_BCU_CON_DELETE,
        GEDA_TEST
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class StoreInDBAttribute : Attribute
    {

    }

    [AttributeUsage(AttributeTargets.Property)]
    public class GroupValueAttribute : Attribute
    {

    }

    [AttributeUsage(AttributeTargets.Class)]
    public class ConfigStoredInDBAttribute : Attribute
    {

    }
}