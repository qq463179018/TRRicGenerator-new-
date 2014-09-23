using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Ric.Core;
using Ric.Db.Info;
using Ric.Db.Manager;
using Ric.Util;

namespace Ric.Tasks
{

    [ConfigStoredInDB]
    public class KoreaChangeGedaNdaConfig
    {
        [StoreInDB]
        [Category("Path")]
        [DisplayName("Bulk file")]
        [DefaultValue("C:\\Korea_Auto\\Korea_Change\\")]
        [Description("Folder for saving generated GDEA and NDA files.")]
        public string BulkFile { get; set; }
    }

    /// <summary>
    /// This task can get the tomorrow's change records in database and generate GEDA and NDA files. 
    /// </summary>
    public class KoreaChangeGedaNda : GeneratorBase
    {
        #region Fields

        private KoreaChangeGedaNdaConfig configObj = null;

        private List<CompanyWarrantTemplate> priceChange = null;

        private List<KoreaEquityInfo> nameChange = null;

        private string outputFolder = string.Empty;

        private string effectiveDate = string.Empty; 

        #endregion

        #region Intialize and Start

        protected override void Initialize()
        {
            configObj = Config as KoreaChangeGedaNdaConfig;
            if (string.IsNullOrEmpty(configObj.BulkFile))
            {
                configObj.BulkFile = GetOutputFilePath();
            }

            outputFolder = Path.Combine(configObj.BulkFile, DateTime.Today.ToString("yyyy-MM-dd"));
            
            TaskResultList.Add(new TaskResultEntry("Log", "Log", Logger.FilePath));  
        }

        /// <summary>
        /// Get next trading day's name change info.
        /// Get next trading day's company warrant change info.
        /// Generate files.
        /// </summary>
        protected override void Start()
        {       
            
            InitialzeEffectiveDate();

            StartNameChangePart();

            StartPriceChangePart();

        }

        /// <summary>
        /// Find next trading day as effective date.
        /// </summary>
        private void InitialzeEffectiveDate()
        {
            List<DateTime> holidayList = HolidayManager.SelectHoliday(MarketId);

            effectiveDate = MiscUtil.GetNextTradingDay(DateTime.Today, holidayList, 1).ToString("yyyy-MM-dd", new CultureInfo("en-US"));
        }

        #endregion

        #region Equity Name Change Part

        /// <summary>
        /// Select equities from database and generate GEDA and NDA files.
        /// </summary>
        private void StartNameChangePart()
        {
            Logger.Log("Start Equity Name Change Part:");
            nameChange = KoreaEquityManager.SelectEquityByEffectiveDateChange(effectiveDate);
            if (nameChange == null)
            {
                string msg = string.Format("No equity name change records will effect on {0}.", effectiveDate);
                Logger.Log(msg);
                return;
            }
            else
            {
                string msg = string.Format("{0} equity name change record(s) will effect on {1}.", nameChange.Count, effectiveDate);
                Logger.Log(msg);
            }
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            GenerateNameChangeGedaFile();
            GenerateNameChangeNdaQAFile();
            GenerateNameChangeNdaIAFile();
        }

        /// <summary>
        /// Generate name change GEDA file.
        /// </summary>
        private void GenerateNameChangeGedaFile()
        {
            string fileName = string.Format("KR_EQ_NameChange_{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
            string filePath = Path.Combine(outputFolder, fileName);

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            List<List<string>> data = new List<List<string>>();
            List<string> title = new List<string>() { "RIC", "DSPLY_NAME", "DSPLY_NMLL", "#INSTMOD_TDN_ISSUER_NAME" };

            foreach (KoreaEquityInfo item in nameChange)
            {
                StringBuilder newLegalName = new StringBuilder(item.LegalName);
                for (int i = 0; i < newLegalName.Length; i++)
                {
                    if (char.IsLower(newLegalName[i]))
                    {
                        newLegalName[i] = char.ToUpper(newLegalName[i]);
                    }
                }

                List<string> content = new List<string>();
                content.Add(item.RIC);
                content.Add(item.IDNDisplayName);
                content.Add(item.KoreaName);
                content.Add(newLegalName.ToString());
                data.Add(content);
            }

            FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Overwrite);
            Logger.Log("Generated name change GEDA file.");
            AddResult(fileName, filePath, "idn");
            //TaskResultList.Add(new TaskResultEntry(fileName, fileName, filePath, FileProcessType.GEDA_BULK_RIC_CHANGE));
        }
        
        /// <summary>
        /// Generate Name Change Nda QA File.
        /// </summary>
        private void GenerateNameChangeNdaQAFile()
        {
            string fileName = string.Format("KR_EQName{0}QAChg.csv", DateTime.Today.ToString("yyyyMMdd"));
            string filePath = Path.Combine(outputFolder, fileName);

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            List<List<string>> data = new List<List<string>>();
            List<string> title = new List<string>() { "RIC", "ASSET SHORT NAME", "ASSET COMMON NAME" };
            string[] rics = new string[6] { ".", "F.", "S.", "stat.", "ta.", "bl." };

            foreach (KoreaEquityInfo item in nameChange)
            {
                string suffix = item.Type != null ? (" " + item.Type) : "";
                for (int i = 0; i < 6; i++)
                {
                    List<string> content = new List<string>();
                    string ric = item.RIC.Split('.')[0] + rics[i] + item.RIC.Split('.')[1];
                    content.Add(ric);
                    content.Add(item.IDNDisplayName);
                    content.Add(item.IDNDisplayName + suffix);
                    data.Add(content);
                }
            }

            FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Overwrite);
            Logger.Log("Generated name change NDA QA file.");
            AddResult(fileName, filePath, "nda");
            //TaskResultList.Add(new TaskResultEntry(fileName, fileName, filePath, FileProcessType.NDA));
        }

        /// <summary>
        /// Generate Name Change Nda IA File.
        /// </summary>
        private void GenerateNameChangeNdaIAFile()
        {

            string fileName = string.Format("KR_EQName{0}IAChg.csv", DateTime.Today.ToString("yyyyMMdd"));
            string filePath = Path.Combine(outputFolder, fileName);

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            List<List<string>> data = new List<List<string>>();
            List<string> title = new List<string>() { "ISIN", "ASSET COMMON NAME" };

            foreach (KoreaEquityInfo item in nameChange)
            {
                string suffix = FormatNdaSuffix(item);
                string assetCommonName = FormatAssetCommonName(item);
                List<string> content = new List<string>();
                content.Add(item.ISIN);
                content.Add(assetCommonName + suffix);
                data.Add(content);
            }

            FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Overwrite);
            Logger.Log("Generated name change NDA IA file.");
            AddResult(fileName, filePath, "nda");
            //TaskResultList.Add(new TaskResultEntry(fileName, fileName, filePath, FileProcessType.NDA));
        }

        /// <summary>
        /// Format NDA IA suffix of column "ASSET COMMON NAME"
        /// </summary>
        /// <param name="type">equity type</param>
        /// <returns>suffix</returns>
        private string FormatNdaSuffix(KoreaEquityInfo item)
        {
            string type = item.Type;
            if (string.IsNullOrEmpty(type))
            {
                string msg = string.Format("Name change item {0} do not have a type. Please confirm.", item.RIC);
                Logger.Log(msg, Logger.LogType.Warning);
                return "";
            }
            string suffix = "";

            if (type.Equals("ORD"))
            {
                suffix = " Ord Shs";
            }
            else if (type.Equals("PRF"))
            {
                suffix = " Prf Shs";
            }
            else if (type.Equals("KDR"))
            {
                suffix = " KDR";
            }
            else
            {
                string msg = string.Format("Name change for {0} occurs. RIC:{1}. Please format the suffix of column: 'ASSET COMMON NAME' in NDA IA file. ", item.Type, item.RIC);
                Logger.Log(msg, Logger.LogType.Warning);
            }
            return suffix;
        }

        /// <summary>
        /// Format equity's ASSET COMMON NAME in NDA IA file.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private string FormatAssetCommonName(KoreaEquityInfo item)
        {
            string company = ClearCoLtdForName(item.LegalName);
            if (item.Type.Equals("PRF"))
            {
                int index = 0;
                string ending = null;
                if (string.IsNullOrEmpty(item.PrfEnd))
                {
                    Regex regex = new Regex("[0-9]+P");
                    Match match = regex.Match(company);
                    if (match.Success)
                    {
                        index = match.Index;
                        ending = match.Value;
                    }
                }
                else
                {
                    index = company.IndexOf(item.PrfEnd);
                    ending = item.PrfEnd;
                }
                company = company.Replace(ending, "");
                FormatAssetNameNoEnd(ref company);
                if (!string.IsNullOrEmpty(ending))
                {
                    company = company.Insert(index, ending);
                }
            }
            else
            {
                FormatAssetNameNoEnd(ref company);
            }
            return company;
        }

        /// <summary>
        /// Format the company name with rule: first char in each word should be uppercase, others are lowercase.
        /// </summary>
        /// <param name="company"></param>
        private void FormatAssetNameNoEnd(ref string company)
        {
            StringBuilder newDisPlayName = new StringBuilder(company);

            for (int i = 0; i < newDisPlayName.Length; i++)
            {
                if (i == 0)
                {
                    if (char.IsLower(newDisPlayName[i]))
                        newDisPlayName[i] = char.ToUpper(newDisPlayName[i]);
                }
                else if (newDisPlayName[i - 1] == ' ')
                {
                    if (char.IsLower(newDisPlayName[i]))
                        newDisPlayName[i] = char.ToUpper(newDisPlayName[i]);
                    
                }
                else
                {
                    if (char.IsUpper(newDisPlayName[i]))
                        newDisPlayName[i] = char.ToLower(newDisPlayName[i]);
                }
            }
            company = newDisPlayName.ToString();
        }

        /// <summary>
        /// Remove the infos of company like CO LTD CORP INC CORPARATION
        /// </summary>
        /// <param name="legalName">full name</param>
        /// <returns>name without company infos</returns>
        public string ClearCoLtdForName(string legalName)
        {
            legalName = legalName.ToUpper();
            List<string> names = legalName.Split(new char[] { ' ', ',', '.' }).ToList();
            string result = "";
            names.Remove("CO");
            names.Remove("LTD");
            names.Remove("INC");
            names.Remove("CORP");
            names.Remove("COMPANY");
            names.Remove("LIMITED");
            names.Remove("CORPORATION");
            foreach (string name in names)
            {
                if (name == "" || name == " ")
                {
                    continue;
                }
                result += name + " ";
            }
            return result.TrimEnd();
        }

        #endregion

        #region Company Warrant Change Part

        /// <summary>
        /// Select company warrants from database and generate GEDA and NDA files.
        /// </summary>
        private void StartPriceChangePart()
        {
            Logger.Log("Start Company Warrant Change Part:");

            priceChange = KoreaCwntManager.SelectWarrantByEffectiveDateChange(effectiveDate);

            if (priceChange == null)
            {
                string msg = string.Format("No company warrant change record(s) will effect on {0}.", effectiveDate);
                Logger.Log(msg);
                return;
            }
            else
            {
                string msg = string.Format("{0} company warrant change record(s) will effect on {1}.", priceChange.Count, effectiveDate);
                Logger.Log(msg);
            }
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }
            GeneratePriceChangeGedaFile();
            GeneratePriceChangeNdaQAFile();
            GeneratePriceChangeNdaIAFile();
        }

        /// <summary>
        /// Generate Price Change Geda File
        /// </summary>
        private void GeneratePriceChangeGedaFile()
        {
            string fileName = string.Format("KR_CWRTS_Change_{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
            string filePath = Path.Combine(outputFolder, fileName);

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            List<List<string>> data = new List<List<string>>();
            List<string> title = new List<string>() { "RIC", "#INSTMOD_STRIKE_PRC" };

            foreach (CompanyWarrantTemplate item in priceChange)
            {
                List<string> content = new List<string>();
                content.Add(item.RIC);
                content.Add(item.ExercisePrice);
                data.Add(content);
            }

            FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Overwrite);
            Logger.Log("Generated company warrant change GEDA file.");
            AddResult(fileName, filePath, "idn");
            //TaskResultList.Add(new TaskResultEntry(fileName, fileName, filePath, FileProcessType.GEDA_BULK_RIC_CHANGE));
        }

        /// <summary>
        /// Generate Price Change Nda QA File
        /// </summary>
        private void GeneratePriceChangeNdaQAFile()
        {
            string fileName = string.Format("KR_CWRTS_{0}QAChg.csv", DateTime.Today.ToString("yyyyMMdd"));
            string filePath = Path.Combine(outputFolder, fileName);

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            List<List<string>> data = new List<List<string>>();
            List<string> title = new List<string>() { "RIC", "ASSET COMMON NAME", "STRIKE PRICE" };
            string[] rics = new string[2] { ".", "F." };

            foreach (CompanyWarrantTemplate item in priceChange)
            {
                for (int i = 0; i < 2; i++)
                {
                    List<string> content = new List<string>();
                    string ric = item.RIC.Split('.')[0] + rics[i] + item.RIC.Split('.')[1];
                    content.Add(ric);
                    content.Add(item.QACommonName);
                    content.Add(item.ExercisePrice);
                    data.Add(content);
                }
            }

            FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Overwrite);
            Logger.Log("Generated company warrant change NDA QA file.");
            AddResult(fileName, filePath, "nda");
            //TaskResultList.Add(new TaskResultEntry(fileName, fileName, filePath, FileProcessType.NDA));
        }

        /// <summary>
        /// Generate Price Change Nda IA File
        /// </summary>
        private void GeneratePriceChangeNdaIAFile()
        {
            string fileName = string.Format("KR_CWRTS_{0}IAChg.csv", DateTime.Today.ToString("yyyyMMdd"));
            string filePath = Path.Combine(outputFolder, fileName);
           
            List<List<string>> data = new List<List<string>>();
            List<string> title = new List<string>() { "ISIN", "ASSET COMMON NAME" };

            foreach (CompanyWarrantTemplate item in priceChange)
            {
                List<string> content = new List<string>();
                content.Add(item.ISIN);
                content.Add(item.ForIACommonName);
                data.Add(content);
            }
            FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Overwrite);
            Logger.Log("Generated company warrant change NDA IA file.");
            AddResult(fileName, filePath, "nda");
            //TaskResultList.Add(new TaskResultEntry(fileName, fileName, filePath, FileProcessType.NDA));
        }

        #endregion
    }
}
