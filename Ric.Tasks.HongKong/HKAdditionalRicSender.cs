using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.IO;
using System.ComponentModel;
using System.Drawing.Design;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    public class HKAdditonalRicSenderConfig
    {
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> TO_TYPE_RECIPIENTS { get; set; }
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> CC_TYPE_RECIPIENTS { get; set; }
        public string CBBC_FILE_PATH { get; set; }
        public string WARRANT_FILE_PATH { get; set; }

        public string LOG_FILE_PATH { get; set; }

        public string NAME_CHANGE_INFO { get; set; }

    }
    public class HKAdditionalRicSender : GeneratorBase
    {
        private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_AdditionalRicSender.config";
        //private static Logger logger = null;
        private static HKAdditonalRicSenderConfig configObj = null;

        protected override void Start()
        {
            StartAdditionalRicSenderJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(HKAdditonalRicSenderConfig)) as HKAdditonalRicSenderConfig;
            ////logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.New);
        }

        public void StartAdditionalRicSenderJob()
        {
            string errorMsg = string.Empty;
            //FM READY_2 CBBC  15 WARRANTS  1 NAME CHANGE  2 PTs.msg
            string subject = "FM READY_ ";
            subject += GetAdditionalRicNumStr(configObj.CBBC_FILE_PATH);
            subject += " CBBC  ";
            subject += GetAdditionalRicNumStr(configObj.WARRANT_FILE_PATH);
            subject += " WARRANTS  ";
            subject += configObj.NAME_CHANGE_INFO;

            using (OutlookApp app = new OutlookApp())
            {
                OutlookUtil.CreateAndSendMail(app, subject, configObj.TO_TYPE_RECIPIENTS, configObj.CC_TYPE_RECIPIENTS, "", out errorMsg);
            }
            Logger.Log(errorMsg);
        }

        public string GetAdditionalRicNumStr(string filePath)
        {
            string additionalRicNumStr = string.Empty;
            if (string.IsNullOrEmpty(filePath))
            {
                LogMessage("File path can't be null or empty.");
            }
            else if (!File.Exists(filePath))
            {
                LogMessage("File " + filePath + " doesn't exist.");
            }
            else
            {
                using (StreamReader sr = new StreamReader(filePath))
                {
                    while (sr.ReadLine() != null)
                    {
                        string line = sr.ReadLine();
                        if (line.Contains("Generated Ric#:"))
                        {
                            additionalRicNumStr = line.Replace("Generated Ric#:", "");
                            break;
                        }
                    }
                }
            }
            if (additionalRicNumStr == string.Empty)
            {
                LogMessage("There's no additional ric in file " + filePath);
            }
            return additionalRicNumStr;
        }
    }

}
