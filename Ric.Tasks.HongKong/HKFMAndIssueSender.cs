using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.IO;
using System.ComponentModel;
using System.Drawing.Design;
using Ric.Util;
using Ric.Core;

namespace Ric.Tasks.HongKong
{
    public class HKFMAndIssueSenderConfig
    {
        public HKFMSenderConfig FM_SENDER_CONFIG { get; set; }
        public HKIssueSenderConfig ISSUE_SENDER_CONFIG { get; set; }

        public string LOG_FILE_PATH { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class HKFMSenderConfig
    {
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> TO_TYPE_RECIPIENTS { get; set; }
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> CC_TYPE_RECIPIENTS { get; set; }

        public string FM_FILE_DIR { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class HKIssueSenderConfig
    {
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> TO_TYPE_RECIPIENTS { get; set; }
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> CC_TYPE_RECIPIENTS { get; set; }

        public string ISSUE_FILE_DIR { get; set; }
        public string CBBC_ADD_FM_FILE_DIR { get; set; }
    }

    public class HKFMAndIssueSender : GeneratorBase
    {
        private static string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_FMAndIssueSender.config";
        //private static Logger logger = null;
        private static HKFMAndIssueSenderConfig configObj = null;

        protected override void Start()
        {
            Logger.Log("Start HK FM Sender Job: ");
            StartFMSenderJob();
            Logger.Log("HK FM Sener Job Done.");
            Logger.Log("Start Issue and CBBC Add FM File Sender Job: ");
            StartIssueSenderJob();
            Logger.Log("HK Issue and CBBC Add FM File Sender JOb Done.");
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(HKFMAndIssueSenderConfig)) as HKFMAndIssueSenderConfig;
            //logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.New);
        }

        public void StartFMSenderJob()
        {
            string errorMsg = string.Empty;
            List<string> attachedFileList = new List<string>();
            string subject = "HKFM_";
            subject += MiscUtil.ParseDateTime(DateTime.Now);
            subject += "_2";
            string mailBody = @"Hi, 

Please find  the FM for today in the attached folder.

Should you have any questions, please feel free to contact me.  ";

            string attachedFile = GetFMFileZip(configObj.FM_SENDER_CONFIG.FM_FILE_DIR);
            attachedFileList.Add(attachedFile);
            using (OutlookApp app = new OutlookApp())
            {
                OutlookUtil.CreateAndSendMail(app, attachedFileList, subject, configObj.FM_SENDER_CONFIG.TO_TYPE_RECIPIENTS, configObj.FM_SENDER_CONFIG.CC_TYPE_RECIPIENTS, mailBody, out errorMsg);
            }

        }

        public void StartIssueSenderJob()
        {
            string errorMsg = string.Empty;
            string mailBody = string.Empty;
            string subject = string.Empty;
            bool isIssueFileExist = false;
            List<string> attachedFileList = GetAttachedFileList(out isIssueFileExist);

            if (isIssueFileExist)
            {
                mailBody = @"Dear  All ,
 
 
There are one file for CBBC and 2 files for further issue today.   
 
Should you have any questions, please feel free to contact me.  ";
            }

            else
            {
                mailBody = @"Dear  All ,
 
 
There are one file for CBBC and no files for further issue today.   
 
Should you have any questions, please feel free to contact me.  ";
            }

            subject = "Daily files from East Asis " + MiscUtil.ParseDateTimeWithBlank(DateTime.Now) + " (HONG KONG)";

            using (OutlookApp app = new OutlookApp())
            {
                OutlookUtil.CreateAndSendMail(app, attachedFileList, subject, configObj.ISSUE_SENDER_CONFIG.TO_TYPE_RECIPIENTS, configObj.ISSUE_SENDER_CONFIG.CC_TYPE_RECIPIENTS, mailBody, out errorMsg);
            }
        }

        public List<string> GetAttachedFileList(out bool isIssueFileExist)
        {
            List<string> attachedFileList = new List<string>();
            attachedFileList.Add(GetCBBCAddFilePath(configObj.ISSUE_SENDER_CONFIG.CBBC_ADD_FM_FILE_DIR));
            List<string> issueFileList = GetIssueFilePathList(configObj.ISSUE_SENDER_CONFIG.ISSUE_FILE_DIR);
            if (issueFileList.Count == 0)
            {
                isIssueFileExist = false;
            }
            else
            {
                isIssueFileExist = true;
                foreach (string issueFilePath in issueFileList)
                {
                    attachedFileList.Add(issueFilePath);
                }
            }
            return attachedFileList;
        }

        //Get FM Files
        public string GetFMFileZip(string fileDir)
        {
            string zipFilePath = Path.Combine(fileDir, MiscUtil.ParseDateTime(DateTime.Now));
            zipFilePath += ".zip";
            string errorMsg = string.Empty;
            if (string.IsNullOrEmpty(fileDir))
            {
                LogMessage("FM File Directory Can't be null or Empty.");
            }
            if (!Directory.Exists(fileDir))
            {
                LogMessage("Direcory " + fileDir + "doesn't Exist.");
            }

            string[] FMFilePath = Directory.GetFiles(fileDir, "*.xls");
            bool zipResult = ZipUtil.ZipFile(FMFilePath, zipFilePath, out errorMsg);
            if (!zipResult)
            {
                LogMessage("Zip FM Files Failed. ERROR: " + errorMsg);
            }
            return zipFilePath;
        }

        //Get CBBC add FM file
        public string GetCBBCAddFilePath(string dir)
        {
            string CBBCAddFilePath = string.Empty;
            if (string.IsNullOrEmpty(dir))
            {
                LogMessage("Dir can't be Null or empty.");
            }
            if (!Directory.Exists(dir))
            {
                LogMessage("Directory " + dir + "doesn't exist.");
            }

            string[] filePathArr = Directory.GetFiles(dir, "*.csv", SearchOption.AllDirectories);
            string fileName = "WRT_ADD_" + MiscUtil.ParseDateTime(DateTime.Now) + "_hongkongturbo.csv";
            if (filePathArr == null || filePathArr.Length < 1)
            {
                LogMessage("There's no CBBC add FM file under directory " + dir);
            }
            List<string> cbbcFMFileList = new List<string>();
            foreach (string filePath in filePathArr)
            {
                if (Path.GetFileName(filePath) == fileName)
                {
                    cbbcFMFileList.Add(filePath);
                }
            }

            if (cbbcFMFileList.Count == 0)
            {
                LogMessage("There's no CBBC add FM file under directory " + dir);
            }

            if (cbbcFMFileList.Count > 1)
            {
                Logger.Log("There're " + cbbcFMFileList.Count.ToString() + "CBBC add FM files under " + dir, Logger.LogType.Warning);
            }

            else
            {
                CBBCAddFilePath = cbbcFMFileList[0];
            }
            return CBBCAddFilePath;
        }

        //Get Issue file if any
        public List<string> GetIssueFilePathList(string dir)
        {
            List<string> issueFileList = new List<string>();
            if (string.IsNullOrEmpty(dir))
            {
                LogMessage("Dir can't be Null or empty.");
            }
            if (!Directory.Exists(dir))
            {
                LogMessage("Directory " + dir + "doesn't exist.");
            }

            string[] csvFileArr = Directory.GetFiles(dir, "*.csv", SearchOption.TopDirectoryOnly);
            foreach (string filePath in csvFileArr)
            {
                if (Path.GetFileNameWithoutExtension(filePath).Contains(MiscUtil.ParseDateTime(DateTime.Now)))
                {
                    issueFileList.Add(filePath);
                }
            }
            return issueFileList;
        }
    }
}
