using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.IO;
using System.Drawing.Design;
using System.ComponentModel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.HongKong
{
    public class HKDelistSenderConfig
    {
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> TO_TYPE_RECIPIENTS { get; set; }
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> CC_TYPE_RECIPIENTS { get; set; }
        public string LAST_DELIST_FILE_DIR { get; set; }
        public string LOG_FILE_PATH { get; set; }

        public int HOLIDAY_COUNT { get; set; }
    }
    public class DelistSender: GeneratorBase
    {
        private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_DelistSender.config";
        //private static Logger logger = null;
        private static HKDelistSenderConfig configObj = null;

        protected override void Start()
        {
            StartDelistSenderJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(HKDelistSenderConfig)) as HKDelistSenderConfig;
            //logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.New);
        }

        public void StartDelistSenderJob()
        {
            string errorMsg = string.Empty;
            string subject = "HKFM_";
            subject += ParseDateTime(GetNextBusinessDay(configObj.HOLIDAY_COUNT));
            subject += "_1";

            string mailBody = @"Dear  all, 
 
 
Please find attached today’s last minute FM.  
 
Should you have any questions, please feel free to contact me.";

            List<string> attachedFileList = GetDelistFileList(configObj.LAST_DELIST_FILE_DIR);

            using (OutlookApp app = new OutlookApp())
            {
                OutlookUtil.CreateAndSendMail(app, attachedFileList, subject, configObj.TO_TYPE_RECIPIENTS, configObj.CC_TYPE_RECIPIENTS, mailBody, out errorMsg);
            }
            Logger.Log(errorMsg);
           
        }

        public List<string> GetDelistFileList(string dir)
        {
            List<string> DelistFileList = new List<string>();
            if (string.IsNullOrEmpty(dir))
            {
                LogMessage("There's no such directory "+dir);
            }

            string[] fileArr = Directory.GetFiles(dir, "*.xls", SearchOption.TopDirectoryOnly);
            if (fileArr == null || fileArr.Length < 1)
            {
                LogMessage("There's no xls file under " + dir);
            }

            else
            {
                foreach (string fileName in fileArr)
                {
                    if (fileName.ToUpper().Contains("DELETE")&&IsNextBusinessDay(GetFileTime(fileName),configObj.HOLIDAY_COUNT))
                    {
                        DelistFileList.Add(Path.GetFullPath(fileName));
                    }
                }
            }
            return DelistFileList;
        }


        public string GetFileTime(string fileName)
        {
            string fileTime = string.Empty;
            fileName = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrEmpty(fileName))
            {
                LogMessage("File name can't be null or empty.");
            }

            string[] namePart = fileName.Split('_');
            if (namePart.Length != 7)
            {
                LogMessage("Please check the fileName for file " + fileName);
            }

            else
            {
                fileTime = namePart[6];
            }
            return fileTime;
        }

        public DateTime GetNextBusinessDay(int holidayCount)
        {
            DateTime currentDate = DateTime.Now;
            DateTime nextBusinessDay;
            if (currentDate.DayOfWeek == DayOfWeek.Friday)
            {
                nextBusinessDay = currentDate.AddDays(3+holidayCount);
            }

            if (currentDate.DayOfWeek == DayOfWeek.Saturday)
            {
                nextBusinessDay = currentDate.AddDays(2+holidayCount);
            }

            else
            {
                nextBusinessDay = currentDate.AddDays(1+holidayCount);
            }

            return nextBusinessDay;
        }

        
        public bool IsNextBusinessDay(string time, int holidayCount)
        {
            string nextBusinessDayStr = ParseDateTime(GetNextBusinessDay(configObj.HOLIDAY_COUNT));
            if (time.ToUpper() == nextBusinessDayStr.ToUpper())
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string ParseDateTime(DateTime dateTime)
        {
            string dateTimeStr = string.Empty;
            string[] month = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            string temp = dateTime.ToString("dd_MM_yyyy");
            string[] tempArr = temp.Split('_');
            dateTimeStr = tempArr[0];
            dateTimeStr += month[int.Parse(tempArr[1])];
            dateTimeStr += tempArr[2];
            return dateTimeStr;
        }
    }
}
