using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class KOREA_AutoSendFMEmailConfig
    {
        [StoreInDB]
        [DisplayName("Send file date")]
        public string SendFileDate { get; set; }

        [StoreInDB]
        [Category("Alert mail")]
        [DisplayName("Recipients")]
        public List<string> AlertMailToList { get; set; }

        [StoreInDB]
        [Category("Alert mail")]
        [DisplayName("Recipients (Cc)")]
        public List<string> AlertMailCCList { get; set; }

        [StoreInDB]
        [Category("Alert mail")]
        [DisplayName("Signature")]
        public List<string> AlertMailSignatureInformationList { get; set; }
    }

    class AutoSendFMEmail : GeneratorBase
    {
        private KOREA_AutoSendFMEmailConfig configObj;

        protected override void Start()
        {
            StartEMAFileMentanence();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREA_AutoSendFMEmailConfig;
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
                //Logger.LogErrorAndRaiseException(string.Format("{0} doesn't exit",dir));
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
            inscrubed = configObj.AlertMailSignatureInformationList.Aggregate(inscrubed, (current, t) => current + t + "\r\n");

            MailToSend mail = new MailToSend();
            string sendFileDate = DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            string today = DateTime.Today.ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
            string mailSubject = "Daily files from East Asia " + today + " (South Korea)";

            if (!string.IsNullOrEmpty(configObj.SendFileDate))
            {
                sendFileDate = configObj.SendFileDate;
            }

            string emaFileDir = Path.Combine(ConfigureOperator.GetEmaFileSaveDir(), sendFileDate);
            List<string> fileList = GetAllTheEMAFile(emaFileDir);
            if (fileList == null)
            {
                return;
            }
            mail.ToReceiverList.AddRange(configObj.AlertMailToList);
            mail.MailSubject = mailSubject;
            mail.CCReceiverList.AddRange(configObj.AlertMailCCList);

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
                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {
                    countXLSXFile++;
                    strXLSX = strXLSX + "   " + fileName + "\n";
                }
                else
                {
                    if (fileName.Split('.').Length > 1 && fileName.Split('_')[1] == "ADD")
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
            if (countXLSXFile == 0)
            {
                strXLSX = strXLSX + "No such files today ! \n";
            }
            mail.MailBody = "PEO:\t" + "Hi, colleagues,\n   Please find the files for South Korea today." + "\n" + strADD + strXLSX + strISIN + "\nPlease  let me know if you have any concerns.\n\n" + inscrubed;

            TaskResultList.Add(new TaskResultEntry("EMAFile", "EMA File Today", emaFileDir, mail));
        }

        ///
        ///  删除7天前的所有文件。
        ///
        public void DeletOverDateFile()
        {

            string sdir = ConfigureOperator.GetEmaFileSaveDir() + "\\";

            DirectoryInfo dir = new DirectoryInfo(sdir);
            DateTime SevenDaysAgo = DateTime.Now.AddDays(-7);
            if (configObj.SendFileDate != "")
            {
                SevenDaysAgo = DateTime.ParseExact(configObj.SendFileDate, "yyyy-MM-dd", null);
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
