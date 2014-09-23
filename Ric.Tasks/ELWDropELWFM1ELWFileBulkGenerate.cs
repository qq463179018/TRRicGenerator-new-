using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Util;
using Ric.Db.Info;
using Ric.Db.Manager;
using System.IO;
using System.Globalization;
using Ric.Core;

namespace Ric.Tasks
{
    public class ELWDropELWFM1ELWFileBulkGenerate : GeneratorBase
    {
        public static string filename = string.Empty;
        private KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig configObj = null;

        protected override void Start()
        {
            try
            {
                StartJob();
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
                Logger.Log(ex.StackTrace);
                throw ex;
            }
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as KOREA_ELWFM1ELWDropAndFileBulkGeneratorConfig;

        }

        private void StartJob()
        {
            string date = DateTime.Today.ToString("yyyy-MM-dd");
            KoreaRicNumInfo todayInfo = KoreaRicNumManager.GetByDate(date);
            if (todayInfo != null)
            {
                string msg = string.Format("Please notice the task has been done today. Add number: {0}. Drop number: {1}.", todayInfo.WarrantAddRicNum, todayInfo.WarrantDropRicNum);
                Logger.Log(msg, Logger.LogType.Warning);                
            }                  
            

            int addNum = StartELWFM1AndFileBulkGenerateJob();
            int dropNum = StartELWDropJob();

            if ((addNum + dropNum) == 0)
            {
                string msg = "No ADD or DROP data grabbed today.";
                Logger.Log(msg);
                return;
            }

            string emaFolder = CreateFM1FileEmaDir();
            TaskResultList.Insert(0, new TaskResultEntry("EMA Folder", "EMA Folder", emaFolder));
            TaskResultList.Insert(0, new TaskResultEntry("BulkFile Folder", "BulkFile Folder", configObj.BulkFile));
            TaskResultList.Insert(0, new TaskResultEntry("FM Folder", "FM Folder", configObj.FM));
            TaskResultList.Insert(0, new TaskResultEntry("LOG", "LOG", Logger.FilePath));
                       

            KoreaRicNumInfo warrantRicNumInfo = new KoreaRicNumInfo(date, addNum, dropNum);
            KoreaRicNumManager ricManager = new KoreaRicNumManager();

            if (todayInfo == null)
            {
                ricManager.Insert(warrantRicNumInfo);
            }

            else
            {
                ricManager.ModifyByDate(date, addNum, dropNum);
            }

            string filename = "Korea FM for " + DateTime.Today.ToString("dd-MMM-yyyy", new CultureInfo("en-US")).Replace("-", " ") + " (Morning).xls";
            string ipath = Path.Combine(configObj.FM, filename);
            TaskResultList.Add(new TaskResultEntry(Path.GetFileNameWithoutExtension(ipath), "FM File", ipath, CreatFm1Mail(addNum, dropNum)));

            
        }

        private MailToSend CreatFm1Mail(int addNum, int dropNum)
        {
            MailToSend mail = new MailToSend();
            StringBuilder mailbodyBuilder = new StringBuilder();
            string filename = "Korea FM for " + DateTime.Today.ToString("dd-MMM-yyyy", new CultureInfo("en-US")).Replace("-", " ") + " (Morning).xls";
            string ipath = Path.Combine(configObj.FM, filename);
            mail.MailSubject = "KR FM [ELW IPO & DROP] wef " + DateTime.Today.ToString("dd-MMM-yyyy");
            mailbodyBuilder.Append("ELW IPO: ");
            mailbodyBuilder.Append(addNum);
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("DROP: ");

            mailbodyBuilder.Append(dropNum);

            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("Effective Date: ");

            mailbodyBuilder.Append(DateTime.Today.ToString("dd-MMM-yyyy"));
            mailbodyBuilder.Append("\r\n\r\n");
            mail.MailBody = mailbodyBuilder.ToString();

            mail.MailBody += "\r\n\r\n\r\n\r\n";
            foreach (var term in configObj.Signature)
            {
                mail.MailBody += term + "\r\n";
            }
            // List<string> 
            mail.ToReceiverList.AddRange(configObj.MailTo);
            mail.CCReceiverList.AddRange(configObj.MailCC);
            mail.AttachFileList.Add(ipath);
            return mail;

        }

        private int StartELWFM1AndFileBulkGenerateJob()
        {
            Logger.Log("Start ELW FM1 job.");
            ELWFMFirstPart FM = new ELWFMFirstPart(this.TaskResultList, Logger);
            return FM.StartELWFM1AndBulkFileJob(configObj);
        }

        private int StartELWDropJob()
        {
            Logger.Log("Start ELW Drop job.");
            ELWDrop drop = new ELWDrop(this.TaskResultList, Logger);
            return drop.StartELWDropJob(configObj);
        }

        private string CreateFM1FileEmaDir()
        {
            string dir = Path.Combine(ConfigureOperator.GetEmaFileSaveDir(), DateTime.Today.ToString("yyyy-MM-dd", new CultureInfo("en-US")));
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            return dir;
        }
    }
}
