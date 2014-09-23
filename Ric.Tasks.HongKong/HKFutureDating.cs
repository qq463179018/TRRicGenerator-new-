using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;
using System.Windows;
using System.IO;
using Ric.Util;
using Microsoft.Office.Interop.Excel;

namespace Ric.Tasks.HongKong
{
    [ConfigStoredInDB]
    class HKFutureDatingConfig
    {
        [StoreInDB]
        [Category("IA Template")]
        [DisplayName("IA Template File Path.")]
        [Description("F:\\work\\xxx.csv")]
        public string IATemplatePath { get; set; }

        [StoreInDB]
        [Category("BulkFile")]
        [DisplayName("Bulk File Output Folder")]
        [Description("F:\\work\\xxx")]
        public string OutputFolder { get; set; }

        [StoreInDB]
        [Category("HShareIPO")]
        [DisplayName("SourcePath")]
        [Description("F:\\work\\xxx.xls")]
        public string SourcePathHShareIPO { get; set; }
        [StoreInDB]
        [Category("HShareIPO")]
        [DisplayName("TemplatePath")]
        [Description("F:\\work\\xxx.csv")]
        public string TemplatePathHShareIPO { get; set; }

        [StoreInDB]
        [Category("IPO")]
        [DisplayName("SourcePath")]
        [Description("F:\\work\\xxx.xls")]
        public string SourcePathIPO { get; set; }
        [StoreInDB]
        [Category("IPO")]
        [DisplayName("TemplatePath")]
        [Description("F:\\work\\xxx.csv")]
        public string TemplatePathIPO { get; set; }

        [StoreInDB]
        [Category("Parellel trading")]
        [DisplayName("SourcePath")]
        [Description("F:\\work\\xxx.xls")]
        public string SourcePathParellel { get; set; }
        [StoreInDB]
        [Category("Parellel trading")]
        [DisplayName("TemplatePath")]
        [Description("F:\\work\\xxx.csv")]
        public string TemplatePathParellel { get; set; }

        [StoreInDB]
        [Category("RIGHTS")]
        [DisplayName("SourcePath")]
        [Description("F:\\work\\xxx.xls")]
        public string SourcePathRIGHTS { get; set; }
        [StoreInDB]
        [Category("RIGHTS")]
        [DisplayName("TemplatePath")]
        [Description("F:\\work\\xxx.csv")]
        public string TemplatePathRIGHTS { get; set; }
        [StoreInDB]
        [Category("RIGHTS")]
        [DisplayName("EffectiveToValue")]
        [Description("EffectiveTo in IA&QA file ")]
        public string EffectiveTo { get; set; }

        [Category("FileCategory")]
        [DisplayName("FileType")]
        [Description("select file type to run.")]
        public FileCategory FileType { get; set; }
    }

    enum FileCategory : int { AllType, HShareIPO, IPO, ParellelTrading, RIGHTS }

    class HKFutureDating : GeneratorBase
    {
        private static HKFutureDatingConfig configObj;
        private List<List<string>> iaTemplate = null;
        protected override void Initialize()
        {
            configObj = Config as HKFutureDatingConfig;
        }

        protected override void Start()
        {
            if (!CheckConfig())
                return;

            #region [HShareIPO]
            if (configObj.FileType.Equals(FileCategory.AllType) || configObj.FileType.Equals(FileCategory.HShareIPO))
            {
                LogMessage(string.Format("start to generate {0} bulk file.", FileCategory.HShareIPO.ToString()));
                string folder = Path.Combine(configObj.OutputFolder, Path.Combine("HShareIPO", DateTime.Now.ToString("MM-dd-yyyy")));
                List<List<string>> source = ReadExcel(configObj.SourcePathHShareIPO, 1);
                List<List<string>> qaTemplate = ReadExcel(configObj.TemplatePathHShareIPO, 1);
                SourceTemplate sourceTemplate = FormatSource(source);
                FillQATemplate(qaTemplate, sourceTemplate, FileCategory.HShareIPO);
                FillIATemplate(iaTemplate, sourceTemplate, FileCategory.HShareIPO);

                string pathQA = Path.Combine(folder, string.Format("Quote_Future_Add_{0}(HShareIPO).csv", sourceTemplate.OfficialCode));
                GenerateFile(pathQA, qaTemplate);
                AddResult("Quote_Future_Add_HShareIPO.csv", pathQA, "qa bulk file");

                string pathIA = Path.Combine(folder, string.Format("Issue_Future_Add_{0}(HShareIPO).csv", sourceTemplate.OfficialCode));
                GenerateFile(pathIA, iaTemplate);
                AddResult("Issue_Future_Add_HShareIPO.csv", pathIA, "ia bulk file");
            }
            #endregion

            #region [IPO]
            if (configObj.FileType.Equals(FileCategory.AllType) || configObj.FileType.Equals(FileCategory.IPO))
            {
                LogMessage(string.Format("start to generate {0} bulk file.", FileCategory.IPO.ToString()));
                string folder = Path.Combine(configObj.OutputFolder, Path.Combine("IPO", DateTime.Now.ToString("MM-dd-yyyy")));
                List<List<string>> source = ReadExcel(configObj.SourcePathIPO, 1);
                List<List<string>> qaTemplate = ReadExcel(configObj.TemplatePathIPO, 1);
                SourceTemplate sourceTemplate = FormatSource(source);
                FillQATemplate(qaTemplate, sourceTemplate, FileCategory.IPO);
                FillIATemplate(iaTemplate, sourceTemplate, FileCategory.IPO);

                string pathQA = Path.Combine(folder, string.Format("Quote_Future_Add_{0}(IPO).csv", sourceTemplate.OfficialCode));
                GenerateFile(pathQA, qaTemplate);
                AddResult("Quote_Future_Add_IPO.csv", pathQA, "qa bulk file");

                string pathIA = Path.Combine(folder, string.Format("Issue_Future_Add_{0}(IPO).csv", sourceTemplate.OfficialCode));
                GenerateFile(pathIA, iaTemplate);
                AddResult("Issue_Future_Add_IPO.csv", pathIA, "ia bulk file");
            }
            #endregion

            #region [ParaellelTrading]
            if (configObj.FileType.Equals(FileCategory.AllType) || configObj.FileType.Equals(FileCategory.ParellelTrading))
            {
                LogMessage(string.Format("start to generate {0} bulk file.", FileCategory.ParellelTrading.ToString()));
                string folder = Path.Combine(configObj.OutputFolder, Path.Combine("ParellelTrading", DateTime.Now.ToString("MM-dd-yyyy")));
                List<List<string>> source = ReadExcel(configObj.SourcePathParellel, 1);
                List<List<string>> qaTemplate = ReadExcel(configObj.TemplatePathParellel, 1);
                SourceTemplate sourceTemplate = FormatSource(source);
                FillQATemplate(qaTemplate, sourceTemplate, FileCategory.ParellelTrading);
                FillIATemplate(iaTemplate, sourceTemplate, FileCategory.ParellelTrading);

                string pathQA = Path.Combine(folder, string.Format("Quote_Future_Add_{0}(Parellel trading).csv", sourceTemplate.OfficialCodeNew));
                GenerateFile(pathQA, qaTemplate);
                AddResult("Quote_Future_Add_Parellel trading.csv", pathQA, "qa bulk file");

                string pathIA = Path.Combine(folder, string.Format("Issue_Future_Add_{0}(Parellel trading)(Parellel trading).csv", sourceTemplate.OfficialCodeNew));
                GenerateFile(pathIA, iaTemplate);
                AddResult("Issue_Future_Add_Parellel trading.csv", pathIA, "ia bulk file");
            }
            #endregion

            #region [RIGHTS]
            if (configObj.FileType.Equals(FileCategory.AllType) || configObj.FileType.Equals(FileCategory.RIGHTS))
            {
                LogMessage(string.Format("start to generate {0} bulk file.", FileCategory.RIGHTS.ToString()));
                string folder = Path.Combine(configObj.OutputFolder, Path.Combine("RIGHTS", DateTime.Now.ToString("MM-dd-yyyy")));
                List<List<string>> source = ReadExcel(configObj.SourcePathRIGHTS, 1);
                List<List<string>> qaTemplate = ReadExcel(configObj.TemplatePathRIGHTS, 1);
                SourceTemplate sourceTemplate = FormatSource(source);
                FillQATemplate(qaTemplate, sourceTemplate, FileCategory.RIGHTS);
                FillIATemplate(iaTemplate, sourceTemplate, FileCategory.RIGHTS);

                string pathQA = Path.Combine(folder, string.Format("Quote_Future_Add_{0}(RIGHTS).csv", sourceTemplate.OfficialCode));
                GenerateFile(pathQA, qaTemplate);
                AddResult("Quote_Future_Add_RIGHTS.csv", pathQA, "qa bulk file");

                string pathIA = Path.Combine(folder, string.Format("Issue_Future_Add_{0}(RIGHTS).csv", sourceTemplate.OfficialCode));
                GenerateFile(pathIA, iaTemplate);
                AddResult("Issue_Future_Add_RIGHTS.csv", pathIA, "ia bulk file");
            }
            #endregion
        }

        private void FillIATemplate(List<List<string>> iaTemplate, SourceTemplate sourceTemplate, FileCategory category)
        {
            try
            {
                if (sourceTemplate == null)
                {
                    LogMessage("there is no data in the ParellelTrading source file.");
                    return;
                }

                List<List<string>> propertyName = GetPropertyName(iaTemplate);
                if (propertyName == null || propertyName.Count == 0)
                {
                    LogMessage("ParellelTrading QATemplate is emptyt.");
                    return;
                }

                for (int i = 0; i < propertyName[0].Count; i++)
                {
                    iaTemplate[i + 1][4] = GetIAEffectiveTo(category);
                    iaTemplate[i + 1][3] = GetIAEffectiveDate(sourceTemplate, category);
                    iaTemplate[i + 1][2] = GetIAPropertyValue(sourceTemplate, propertyName[0][i], iaTemplate[i + 1][2], category);
                    iaTemplate[i + 1][5] = GetIAChangeOffset(sourceTemplate, propertyName[0][i], category);
                    iaTemplate[i + 1][6] = GetIAChangeTrigger(sourceTemplate, propertyName[0][i], category);
                }

            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private string GetIAEffectiveTo(FileCategory category)
        {
            if (!category.Equals(FileCategory.RIGHTS))
                return string.Empty;

            if ((configObj.EffectiveTo + "").Trim().Length == 0)
                return string.Empty;

            return configObj.EffectiveTo;
        }

        private void FillQATemplate(List<List<string>> qaTemplate, SourceTemplate sourceTemplate, FileCategory category)
        {
            try
            {
                if (sourceTemplate == null)
                {
                    LogMessage("there is no data in the ParellelTrading source file.");
                    return;
                }

                List<List<string>> propertyName = GetPropertyName(qaTemplate);
                if (propertyName == null || propertyName.Count == 0)
                {
                    LogMessage("ParellelTrading is emptyt.");
                    return;
                }

                Dictionary<int, string> endString = new Dictionary<int, string>();
                endString.Add(0, ".HK");

                if (category.Equals(FileCategory.RIGHTS))
                {
                    endString.Add(1, "ta.HK");
                    //endString.Add(2, ".HS");
                }
                else
                {
                    endString.Add(1, "stat.HK");
                    endString.Add(2, "ta.HK");
                    endString.Add(3, ".HS");
                }

                int index;
                for (int i = 0; i < propertyName.Count; i++)
                {
                    for (int j = 0; j < propertyName[i].Count; j++)
                    {
                        index = GetIndex(i, j, propertyName);
                        qaTemplate[index][0] = string.Format("{0}{1}", GetQAUnderlying(sourceTemplate, category), endString[i]);
                        qaTemplate[index][2] = GetQAPropertyValue(sourceTemplate, propertyName[i][j], endString[i], category);
                        qaTemplate[index][3] = GetQAEffectiveDate(sourceTemplate, category);
                        qaTemplate[index][4] = GetQAEffectiveTo(propertyName[i][j], category);
                        qaTemplate[index][5] = GetQAChangeOffset(sourceTemplate, propertyName[i][j], category);
                        qaTemplate[index][6] = GetQAChangeTrigger(sourceTemplate, propertyName[i][j], category);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private string GetQAEffectiveTo(string protyName, FileCategory category)
        {
            if (!category.Equals(FileCategory.RIGHTS))
                return string.Empty;

            if ((configObj.EffectiveTo + "").Trim().Length == 0)
                return string.Empty;

            if ((protyName + "").Trim().Length == 0 || (protyName + "").Trim().Equals("RIC"))
                return configObj.EffectiveTo;

            return string.Empty;
        }

        private string GetIAPropertyValue(SourceTemplate sourceTemplate, string propertyName, string propertyValueInTemplate, FileCategory category)
        {
            try
            {
                if ((propertyName + "").Trim().Length == 0)
                    return GetIAEmptyPropertyName(sourceTemplate, category);

                if (propertyName.Contains("HONG KONG CODE"))
                    return GetIAOfficialCode(sourceTemplate, category);

                if (propertyName.Contains("ASSET COMMON NAME"))
                    return GetIADisplayname(sourceTemplate, category);

                if (propertyName.Contains("RCS ASSET CLASS"))
                    return FileCategory.RIGHTS.Equals(category) ? "RTS" : propertyValueInTemplate;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return string.Empty;
        }

        private string GetQAPropertyValue(SourceTemplate sourceTemplate, string propertyName, string endString, FileCategory category)
        {
            try
            {
                if ((propertyName + "").Trim().Length == 0)
                    return string.Empty;

                if (propertyName.Contains("RIC"))
                    return string.Format("{0}{1}", GetQAUnderlying(sourceTemplate, category), endString);

                if (propertyName.Contains("ASSET COMMON NAME"))
                    return GetQADisplayname(sourceTemplate, category);

                if (propertyName.Contains("ROUND LOT SIZE"))
                    return GetQALotSize(sourceTemplate, category);

                if (propertyName.Contains("TICKER SYMBOL"))
                    return GetQAOfficialCode(sourceTemplate, category);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return string.Empty;
        }

        #region [QA]
        private string GetQAChangeTrigger(SourceTemplate sourceTemplate, string propertyName, FileCategory category)
        {
            if (category.Equals(FileCategory.HShareIPO) || category.Equals(FileCategory.IPO) || category.Equals(FileCategory.RIGHTS))
            {
                if ((propertyName + "").Trim().Length == 0)
                    return string.Empty;
                else
                    return "PEO";
            }

            return string.Empty;
        }

        private string GetQAChangeOffset(SourceTemplate sourceTemplate, string propertyName, FileCategory category)
        {
            if (category.Equals(FileCategory.RIGHTS))
            {
                if ((propertyName + "").Contains("RIC"))
                    return "23";
                else
                    return string.Empty;
            }

            return string.Empty;
        }

        private string GetQAEmptyPropertyName(SourceTemplate sourceTemplate, FileCategory category)
        {
            return string.Empty;
        }

        private string GetQAUnderlying(SourceTemplate sourceTemplate, FileCategory category)
        {
            //if (category.Equals(FileCategory.ParellelTrading))
            //    return sourceTemplate.UnderlyingRICNew;
            //else
            //    return sourceTemplate.UnderlyingRIC;

            if ((sourceTemplate.UnderlyingRICNew + "").Trim().Length != 0)
                return sourceTemplate.UnderlyingRICNew;
            else if ((sourceTemplate.UnderlyingRIC + "").Trim().Length != 0)
                return sourceTemplate.UnderlyingRIC;
            else
            {
                LogMessage("cannot found the UnderlyingRICNew/UnderlyingRIC in the source file.");
                return string.Empty;
            }
        }

        private string GetQAOfficialCode(SourceTemplate sourceTemplate, FileCategory category)
        {
            //if (category.Equals(FileCategory.ParellelTrading))
            //    return sourceTemplate.OfficialCodeNew;
            //else
            //    return sourceTemplate.OfficialCode;

            if ((sourceTemplate.OfficialCodeNew + "").Trim().Length != 0)
                return sourceTemplate.OfficialCodeNew;
            else if ((sourceTemplate.OfficialCode + "").Trim().Length != 0)
                return sourceTemplate.OfficialCode;
            else
            {
                LogMessage("cannot found the OfficialCodeNew/OfficialCode in the source file.");
                return string.Empty;
            }
        }

        private string GetQALotSize(SourceTemplate sourceTemplate, FileCategory category)
        {
            //if (category.Equals(FileCategory.ParellelTrading))
            //    return sourceTemplate.LotSizeNew;
            //else
            //    return sourceTemplate.LotSize;

            if ((sourceTemplate.LotSizeNew + "").Trim().Length != 0)
                return sourceTemplate.LotSizeNew;
            else if ((sourceTemplate.LotSize + "").Trim().Length != 0)
                return sourceTemplate.LotSize;
            else
            {
                LogMessage("cannot found the LotSizeNew/LotSize in the source file.");
                return string.Empty;
            }
        }

        private string GetQADisplayname(SourceTemplate sourceTemplate, FileCategory category)
        {
            if (category.Equals(FileCategory.HShareIPO))
                return string.Format("{0} ORD H", sourceTemplate.Displayname.Replace("<--NOT YET CFM", "").Trim());

            if (category.Equals(FileCategory.IPO))
                return string.Format("{0} ORD", sourceTemplate.Displayname.Replace("<--NOT YET CFM", "").Trim());

            if (category.Equals(FileCategory.ParellelTrading))
                return string.Format("{0} ORD (TEMP)", sourceTemplate.Displayname.Replace("<--NOT YET CFM", "").Trim());

            if (category.Equals(FileCategory.RIGHTS))
                return sourceTemplate.Displayname.Replace("<--NOT YET CFM", "").Trim();

            return string.Empty;
        }

        private string GetQAEffectiveDate(SourceTemplate sourceTemplate, FileCategory category)
        {
            return sourceTemplate.EffectiveDate;
        }
        #endregion

        #region [IA]
        private string GetIAChangeTrigger(SourceTemplate sourceTemplate, string propertyName, FileCategory category)
        {
            if (category.Equals(FileCategory.HShareIPO) || category.Equals(FileCategory.IPO) || category.Equals(FileCategory.RIGHTS))
            {
                if ((propertyName + "").Trim().Length == 0)
                    return string.Empty;
                else
                    return "PEO";
            }

            return string.Empty;
        }

        private string GetIAChangeOffset(SourceTemplate sourceTemplate, string propertyName, FileCategory category)
        {
            //if (category.Equals(FileCategory.RIGHTS))
            //{
            //    if ((propertyName + "").Trim().Length == 0)
            //        return string.Empty;
            //    else
            //        return "23";
            //}

            return string.Empty;
        }

        private string GetIAEmptyPropertyName(SourceTemplate sourceTemplate, FileCategory category)
        {
            return string.Empty;
        }

        private string GetIAUnderlying(SourceTemplate sourceTemplate, FileCategory category)
        {
            //if (category.Equals(FileCategory.ParellelTrading))
            //    return sourceTemplate.UnderlyingRICNew;
            //else
            //    return sourceTemplate.UnderlyingRIC;

            if ((sourceTemplate.UnderlyingRICNew + "").Trim().Length != 0)
                return sourceTemplate.UnderlyingRICNew;
            else if ((sourceTemplate.UnderlyingRIC + "").Trim().Length != 0)
                return sourceTemplate.UnderlyingRIC;
            else
            {
                LogMessage("cannot found the UnderlyingRICNew/UnderlyingRIC in the source file.");
                return string.Empty;
            }
        }

        private string GetIAOfficialCode(SourceTemplate sourceTemplate, FileCategory category)
        {
            //if (category.Equals(FileCategory.ParellelTrading))
            //    return sourceTemplate.OfficialCodeNew;
            //else
            //    return sourceTemplate.OfficialCode;

            if ((sourceTemplate.OfficialCodeNew + "").Trim().Length != 0)
                return sourceTemplate.OfficialCodeNew;
            else if ((sourceTemplate.OfficialCode + "").Trim().Length != 0)
                return sourceTemplate.OfficialCode;
            else
            {
                LogMessage("cannot found the OfficialCodeNew/OfficialCode in the source file.");
                return string.Empty;
            }
        }

        private string GetIALotSize(SourceTemplate sourceTemplate, FileCategory category)
        {
            //if (category.Equals(FileCategory.ParellelTrading))
            //    return sourceTemplate.LotSizeNew;
            //else
            //    return sourceTemplate.LotSize;

            if ((sourceTemplate.LotSizeNew + "").Trim().Length != 0)
                return sourceTemplate.LotSizeNew;
            else if ((sourceTemplate.LotSize + "").Trim().Length != 0)
                return sourceTemplate.LotSize;
            else
            {
                LogMessage("cannot found the LotSizeNew/LotSize in the source file.");
                return string.Empty;
            }
        }

        private string GetIADisplayname(SourceTemplate sourceTemplate, FileCategory category)
        {
            if (category.Equals(FileCategory.HShareIPO))
                return string.Format("{0} Ord Shs H", sourceTemplate.LegalRegisteredName.Replace("Company", "").Replace("Limited", "").Trim());

            if (category.Equals(FileCategory.IPO))
                return string.Format("{0} Ord Shs", sourceTemplate.LegalRegisteredName.Replace("Company", "").Replace("Limited", "").Trim());

            if (category.Equals(FileCategory.ParellelTrading))
                return string.Format("{0} Ord Shs (Temp)", sourceTemplate.LegalRegisteredName.Replace("Company", "").Replace("Limited", "").Trim());

            if (category.Equals(FileCategory.RIGHTS))
                return string.Format("{0} Rights", sourceTemplate.LegalRegisteredName.Replace("Company", "").Replace("Limited", "").Trim());

            return string.Empty;
        }

        private string GetIAEffectiveDate(SourceTemplate sourceTemplate, FileCategory category)
        {
            return sourceTemplate.EffectiveDate;
        }
        #endregion

        private int GetIndex(int lineCount, int valueCount, List<List<string>> propertyName)
        {
            int result = valueCount + 1;

            try
            {
                for (int i = 0; i < lineCount; i++)
                    result += propertyName[i].Count;

                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return 0;
            }
        }

        private void GenerateFile(string path, List<List<string>> template)
        {
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(path)))
                    Directory.CreateDirectory(Path.GetDirectoryName(path));

                if (File.Exists(path))
                    File.Delete(path);

                XlsOrCsvUtil.GenerateStringCsv(path, template);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private List<List<string>> GetPropertyName(List<List<string>> qaTemplate)
        {
            List<List<string>> result = new List<List<string>>();
            List<string> line = null;

            try
            {
                for (int i = 1; i < qaTemplate.Count; i++)
                {
                    if (i == 1)
                        line = new List<string>();

                    if (line.Contains(qaTemplate[i][1]))
                    {
                        result.Add(line);
                        line = new List<string>();
                    }

                    line.Add(qaTemplate[i][1]);
                }

                result.Add(line);
                return result;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private SourceTemplate FormatSource(List<List<string>> listSource)
        {
            SourceTemplate sourceTemplate = new SourceTemplate();

            try
            {
                foreach (var line in listSource)
                {
                    if (line == null || line.Count < 2)
                        continue;

                    if ((line[0] + "").Trim().Length == 0)
                        continue;

                    string name = line[0].Trim();
                    string value = (line[1] + "").Replace("<-- NOT YET CFM", "").Trim();
                    if (name.Contains("Effective Date:") && value.Length != 0)
                        sourceTemplate.EffectiveDate = ConvertDateNumber(value);

                    if (name.Contains("Effective Date (NEW):") && value.Length != 0)
                        sourceTemplate.EffectiveDateNew = ConvertDateNumber(value);

                    if (name.Contains("Underlying RIC:") && value.Length != 0)
                        sourceTemplate.UnderlyingRIC = value.Substring(0, value.Length - 3);

                    if (name.Contains("Underlying RIC (NEW):") && value.Length != 0)
                        sourceTemplate.UnderlyingRICNew = value.Substring(0, value.Length - 3);

                    if (name.Contains("Displayname:") && value.Length != 0)
                        sourceTemplate.Displayname = value;

                    if (name.Contains("Displayname (NEW):") && value.Length != 0)
                        sourceTemplate.DisplaynameNew = value;

                    if (name.Contains("Official Code:") && value.Length != 0)
                        sourceTemplate.OfficialCode = value;

                    if (name.Contains("Official Code (NEW):") && value.Length != 0)
                        sourceTemplate.OfficialCodeNew = value;

                    if (name.Contains("Lot Size:") && value.Length != 0)
                        sourceTemplate.LotSize = value;

                    if (name.Contains("Lot Size (NEW):") && value.Length != 0)
                        sourceTemplate.LotSizeNew = value;

                    if (name.Contains("Legal Registered Name:") && value.Length != 0)
                        sourceTemplate.LegalRegisteredName = value;

                    if (name.Contains("Legal Registered Name (NEW):") && value.Length != 0)
                        sourceTemplate.LegalRegisteredNameNew = value;
                }

                return sourceTemplate;
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return null;
            }
        }

        private string ConvertDateNumber(string value)
        {
            try
            {
                return Convert.ToDateTime("1900-01-01").AddDays(Convert.ToInt32(value.Trim()) - 2).ToString("dd-MMM-yy");
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);

                return string.Empty;
            }
        }

        private bool CheckConfig()
        {
            if ((configObj.IATemplatePath + "").Trim().Length == 0 || !File.Exists(configObj.IATemplatePath))
            {
                LogMessage("IATemplate path invalid .");
                return false;
            }

            if ((configObj.OutputFolder + "").Trim().Length == 0)
            {
                LogMessage("Output folder invalid .");
                return false;
            }

            if (configObj.FileType.Equals(FileCategory.AllType) || configObj.FileType.Equals(FileCategory.IPO))
            {
                if (!CheckPath(configObj.SourcePathIPO, configObj.TemplatePathIPO))
                {
                    LogMessage("IPO config setting invalid .");
                    return false;
                }

            }

            if (configObj.FileType.Equals(FileCategory.AllType) || configObj.FileType.Equals(FileCategory.ParellelTrading))
            {
                if (!CheckPath(configObj.SourcePathParellel, configObj.TemplatePathParellel))
                {
                    LogMessage("ParellelTading config setting invalid .");
                    return false;
                }
            }

            if (configObj.FileType.Equals(FileCategory.AllType) || configObj.FileType.Equals(FileCategory.RIGHTS))
            {
                if (!CheckPath(configObj.SourcePathRIGHTS, configObj.TemplatePathRIGHTS))
                {
                    LogMessage("RIGHTS config setting invalid .");
                    return false;
                }
            }

            try
            {
                if (!Directory.Exists(configObj.OutputFolder))
                    Directory.CreateDirectory(configObj.OutputFolder);

                iaTemplate = ReadExcel(configObj.IATemplatePath, 1);
            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
                return false;
            }

            return true;
        }

        private List<List<string>> ReadExcel(string path, int position)
        {
            try
            {
                using (ExcelApp excelApp = new ExcelApp(false, false))
                {
                    return ExcelUtil.CreateOrOpenExcelFile(excelApp, path).ToList(position);
                }

            }
            catch (Exception ex)
            {
                string msg = string.Format("\r\n	     ClassName:  {0}\r\n	     MethodName: {1}\r\n	     Message:    {2}",
                                            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(),
                                            System.Reflection.MethodBase.GetCurrentMethod().Name,
                                            ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }

            return null;
        }

        private bool CheckPath(string source, string template)
        {
            if ((source + "").Trim().Length == 0 || !File.Exists(source))
                return false;

            if ((template + "").Trim().Length == 0 || !File.Exists(source))
                return false;

            return true;
        }
    }

    class SourceTemplate
    {
        public string EffectiveDate { get; set; }
        public string UnderlyingRIC { get; set; }
        public string Displayname { get; set; }
        public string OfficialCode { get; set; }
        public string LotSize { get; set; }
        public string LegalRegisteredName { get; set; }

        public string EffectiveDateNew { get; set; }
        public string UnderlyingRICNew { get; set; }
        public string DisplaynameNew { get; set; }
        public string OfficialCodeNew { get; set; }
        public string LotSizeNew { get; set; }
        public string LegalRegisteredNameNew { get; set; }
    }
}
