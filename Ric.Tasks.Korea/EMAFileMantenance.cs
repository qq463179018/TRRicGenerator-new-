using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Selenium;
using System.Xml;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using HtmlAgilityPack;
using System.Web;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.Drawing.Design;
using System.ComponentModel;
using Ric.Db.Manager;
using Ric.Util;
using Ric.Core;
namespace Ric.Tasks.Korea
{
    public class KOREA_SEMAFileMantenanceConfig
    {
        public string SEND_FILE_DATE { get; set; }
        public string LOG_FILE_NAME{ get; set; }
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> ALERT_MAIL_TO_LIST { get; set; }
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string>ALERT_MAIL_CC_LIST { get; set; }
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> ALERT_MAIL_SIGNATURE_INFORMATION_LIST { get; set; }
    }
    public class EMAFileMantenance:GeneratorBase
    {
        private static readonly string CONFIGFILE_NAME = ".\\Config\\Korea\\KOREA_EMAFileMantenance.config";
        private KOREA_SEMAFileMantenanceConfig configObj = null;
        private Logger logger = null;
        
        protected override void Start()
        {
            StartEMAFileMentanence();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(CONFIGFILE_NAME, typeof(KOREA_SEMAFileMantenanceConfig)) as KOREA_SEMAFileMantenanceConfig;
            logger = new Logger(configObj.LOG_FILE_NAME, Logger.LogMode.New);
        }
        /// <summary>
        /// 获得所有将发送的文件的名称。
        /// </summary>
        /// <param name="dir"></param>
        /// <returns></returns>
        public List<string> GetAllTheEMAFile(string dir)
        {
            if (!Directory.Exists(dir))
            {
                //logger.LogErrorAndRaiseException(string.Format("{0} doesn't exit",dir));
                Logger.Log("No email file to send today !");
                 return null;
            }
            return Directory.GetFiles(dir).ToList();            
        }

        /// <summary>
        /// 发送邮件功能。
        /// </summary>
        public void StartSendMail()
        {
            string inscrubed = "";
            inscrubed = inscrubed + "\r\n\r\n\r\n";
            //inscrubed = inscrubed + configObj.ALERT_MAIL_INSCRIBED_PERSON + "\r\n";
            //inscrubed = inscrubed + configObj.ALERT_MAIL_INSCRIBED_PSITION + "\r\n";
            //inscrubed = inscrubed + "\r\nThomson Reuters\r\n\r\nPhone: " + configObj.ALERT_MAIL_INSCRIBED_PHONE + "\r\n";
            //inscrubed = inscrubed + configObj.ALERT_MAIL_INSCRIBED_EMAIL_ADRESS + "\r\nthomsonreuters.com\r\n";
            for (int i = 0; i < configObj.ALERT_MAIL_SIGNATURE_INFORMATION_LIST.Count; i++)
            {
                inscrubed = inscrubed + configObj.ALERT_MAIL_SIGNATURE_INFORMATION_LIST[i] + "\r\n";
            }

            MailToSend mail = new MailToSend();
            string sendFileDate = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            string today = DateTime.Today.ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            string mailSubject = "Daily files from East Asia " + today + " (South Korea)"; 
            
            if (!string.IsNullOrEmpty(configObj.SEND_FILE_DATE))
            {
                sendFileDate = configObj.SEND_FILE_DATE;
            }

            //string emaFileDir = configObj.EMA_FILE_BASE_DIR + "\\" + sendFileDate;
           // string emaFileDir = ConfigureOperator.getEmaFileSaveDir()+"\\"  + sendFileDate; 
            string emaFileDir = Path.Combine(ConfigureOperator.GetEmaFileSaveDir(),sendFileDate);
            List<string> fileList = GetAllTheEMAFile(emaFileDir);
            if (fileList == null)
            {
                return;
            }
            mail.ToReceiverList.AddRange(configObj.ALERT_MAIL_TO_LIST);
            mail.MailSubject = mailSubject;
            mail.CCReceiverList.AddRange(configObj.ALERT_MAIL_CC_LIST);

            string strADD = "\nFile for ADD.\n";
            string strXLSX = "\nThe excel file is for your checking on the mature date (compare with the date in the csv file dd/mm/yy).\n";
            string strISIN = "\nFiles for Change such as ISIN, Issue amount, issue price and strike price change and so on.\n";
            int countADDFile = 0;
            int countXLSXFile = 0;
            int countISINFile = 0;

            foreach (string file in fileList)
            {
                string fileName = Path.GetFileName(file);
                string fileExtension = Path.GetExtension(file);
                mail.AttachFileList.Add(file);
                if (fileExtension  == ".xls" || fileExtension==".xlsx")
                {
                    countXLSXFile++;
                    strXLSX = strXLSX + "   " + fileName + "\n";
                }
                else
                {
                    if (fileName .Split('.').Length>1&& fileName.Split('_')[1] == "ADD")
                    {
                        countADDFile++;
                        strADD = strADD + "   " + fileName + "\n";
                    }
                    else
                    {
                        countISINFile++;
                        strISIN = strISIN + "   " + fileName + "\n";
                    }
                }                
            }

            if (countADDFile == 0)
            {
                strADD = strADD + "No such files today ! \n";
            }
            if (countISINFile == 0)
            {
                strISIN = strISIN + "No such files today ! \n";
            }
            if(countXLSXFile == 0)
            {
                strXLSX = strXLSX + "No such files today ! \n";
            }
            mail.MailBody = "Hi, colleagues,\n   Please find the files for South Korea today." + "\n" + strADD + strXLSX + strISIN +"\nPlease  let me know if you have any concerns.\n\n"+inscrubed;

            //string err = "Send EMA mail error !";
            //using (OutlookApp app = new OutlookApp())
            //{
            //    OutlookUtil.CreateAndSendMail(app, mail, out err);
            //}
           // AddResult(sendFileDate,emaFileDir,"EMA File Folder Path");
            TaskResultList.Add(new TaskResultEntry("EMAFile", "EMA File Today", emaFileDir,mail));
        }
        
        ///
        ///  删除7天前的所有文件。
        ///
        public void DeletOverDateFile()
        {
            
            string sdir = ConfigureOperator.GetEmaFileSaveDir() +"\\";
            
            DirectoryInfo dir = new DirectoryInfo(sdir);
            DateTime SevenDaysAgo = DateTime.Now.AddDays(-7);
            if (configObj.SEND_FILE_DATE != "")
            {
                SevenDaysAgo = DateTime.ParseExact(configObj.SEND_FILE_DATE, "yyyy-MM-dd", null);
            }
            foreach (var subDir in Directory.GetDirectories(ConfigureOperator.GetEmaFileSaveDir()))
            {
                string dirName = Path.GetFileName(subDir);

                try
                {
                    DateTime curDate = DateTime.ParseExact(dirName, "yyyy-MM-dd", null);
                    if (curDate < SevenDaysAgo)
                    {
                        Directory.Delete(subDir, true);
                    }
                }

                catch (Exception ex)
                {
                    Logger.Log(string.Format("Error happens when trying to delete the files seven days ago,  please check the file folder name, correct folder name should be in the format \"yyyy-MM-dd\". Ex: {0}", ex.Message));
                }
            }
        }

        public void StartEMAFileMentanence()
        {
            StartSendMail();
            DeletOverDateFile();
        }
   
    }
}
