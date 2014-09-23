using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Taiwan
{
    #region Configuration
    [ConfigStoredInDB]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class TGODailyRICCreationRolloverConfig
    {
        [StoreInDB]
        [DisplayName("Text file path")]
        [Description("GeneratedTxtFilePath like: G\\Japan")]
        public string TxtFilePath { get; set; }
    }
    #endregion

    #region IDNTGOBulkFileEntity
    public class IDNTGOBulk
    {
        public string SYMBOL { get; set; }
        public string DSPLY_NAME { get; set; }
        public string RIC { get; set; }
        public string OFFCL_CODE { get; set; }
        public string EX_SYMBOL { get; set; }
        public string EXPIR_DATE { get; set; }
        public string CONTR_MNTH { get; set; }
        public string CONTR_SIZE { get; set; }
        public string STRIKE_PRC { get; set; }
        public string PUTCALLIND { get; set; }
        public string BCKGRNDPAG { get; set; }
        public string DSPLY_NMLL { get; set; }
        public string X_INST_TITLE { get; set; }
        public string X_80CHAR { get; set; }
        public string INSTMOD_PUT_CALL { get; set; }
        public string EXL_NAME { get; set; }
        public string BCU { get; set; }
        public string INSTMOD_PROV_SYMB { get; set; }
    }
    #endregion

    class TGODailyRICCreationRollover : GeneratorBase
    {
        #region Declaration
        private static TGODailyRICCreationRolloverConfig configObj = null;
        private string tgtm = "#TG*.TM";
        private string patternTGTM = @"\bLONGLINK\d+\b\s+\b(?<RIC>TG\d{4,5}[A-Z]{1}\d{1}\.TM)\r\n\S+";
        private Dictionary<string, IDNTGOBulk> dicIDNTGOBulkTGTM = new Dictionary<string, IDNTGOBulk>();
        private const string strTGFc1 = "TGFc1";
        private string patternTGFc1 = @"\bSETTLE\b\s+\b(?<SETTLE>\d+)(\r\nTGFc1|\.\d+\r\nTGFc1)\b";
        private Dictionary<string, IDNTGOBulk> dicIDNTGOBulkTGFc1 = new Dictionary<string, IDNTGOBulk>();
        private List<IDNTGOBulk> listIDNTGOBulkTGFc1ROLLOV = new List<IDNTGOBulk>();
        private Dictionary<string, string> dicCalls = new Dictionary<string, string>();
        private Dictionary<string, string> dicPuts = new Dictionary<string, string>();
        private List<string> listMonthAndYearFromIDNTGOBulkTGTM = new List<string>();
        private string txtFileName = string.Empty;
        private string txtFileNameMissing = string.Empty;
        private string txtFilePath = string.Empty;
        private string expirDate = string.Empty;
        private string expirDateFormed = string.Empty;
        private string monthYear = string.Empty;
        protected override void Initialize()
        {
            base.Initialize();
            configObj = Config as TGODailyRICCreationRolloverConfig;
            txtFilePath = configObj.TxtFilePath.ToString().Trim();
            dicCalls.Add("A", "Jan");
            dicCalls.Add("B", "Feb");
            dicCalls.Add("C", "Mar");
            dicCalls.Add("D", "Apr");
            dicCalls.Add("E", "May");
            dicCalls.Add("F", "Jun");
            dicCalls.Add("G", "Jul");
            dicCalls.Add("H", "Aug");
            dicCalls.Add("I", "Sep");
            dicCalls.Add("J", "Oct");
            dicCalls.Add("K", "Nov");
            dicCalls.Add("L", "Dec");
            dicPuts.Add("M", "Jan");
            dicPuts.Add("N", "Feb");
            dicPuts.Add("O", "Mar");
            dicPuts.Add("P", "Apr");
            dicPuts.Add("Q", "May");
            dicPuts.Add("R", "Jun");
            dicPuts.Add("S", "Jul");
            dicPuts.Add("T", "Aug");
            dicPuts.Add("U", "Sep");
            dicPuts.Add("V", "Oct");
            dicPuts.Add("W", "Nov");
            dicPuts.Add("X", "Dec");
            txtFileName = DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + "_TGO.txt";
            txtFileNameMissing = DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + "_ROLLOV.txt";
            expirDate = Interaction.InputBox("Click OK and input EXPIR_DATE in the following input box will generate rollover file", "Whether to generate rollover file?", "", 400, 320).ToString();
            if (!string.IsNullOrEmpty(expirDate))
            {
                expirDateFormed = TWHelper.DateStringForm(expirDate, "dd-MMM-yy");
                monthYear = expirDateFormed.Substring(expirDateFormed.IndexOf('-') + 1, 3).ToUpper() + expirDateFormed.Substring(expirDateFormed.Length - 1, 1);
            }
        }
        #endregion

        protected override void Start()
        {
            try
            {
                UseListTGTM(tgtm);
                FilldicCallAndPuts(dicIDNTGOBulkTGTM, listMonthAndYearFromIDNTGOBulkTGTM);
                GetTodaySETTLEFromGATSTolistIDNTGOBulkTGFc1(strTGFc1, patternTGFc1, dicIDNTGOBulkTGFc1, listMonthAndYearFromIDNTGOBulkTGTM, listIDNTGOBulkTGFc1ROLLOV);
                GenerateTxtlistIDNTGOBulkTGFc1ROLLOVER(listIDNTGOBulkTGFc1ROLLOV, txtFilePath, txtFileNameMissing);
                ComparelistIDNTGOBulkTGTMWithlistIDNTGOBulkTGFc1(dicIDNTGOBulkTGTM, dicIDNTGOBulkTGFc1);
                GenerateTxtlistIDNTGOBulkTGFc1TGO(dicIDNTGOBulkTGFc1, txtFilePath, txtFileName);
            }
            catch (Exception e)
            {
                LogMessage("Error happened: " + e.Message, Logger.LogType.Error);
            }
        }

        #region GenerateTxtlistIDNTGOBulkTGFc1TGO
        private void GenerateTxtlistIDNTGOBulkTGFc1TGO(Dictionary<string, IDNTGOBulk> dicIDNTGOBulkTGFc1, string txtFilePath, string txtFileName)
        {
            string warrantAddISINTxtFilePath = Path.Combine(txtFilePath, txtFileName);
            string content = "SYMBOL\tDSPLY_NAME\tRIC\tOFFCL_CODE\tEX_SYMBOL\tEXPIR_DATE\tCONTR_MNTH\tCONTR_SIZE\tSTRIKE_PRC\tPUTCALLIND\tBCKGRNDPAG\tDSPLY_NMLL\tX_INST_TITLE\tX_80CHAR\t#INSTMOD_PUT_CALL\tEXL_NAME\tBCU\t#INSTMOD_PROV_SYMB\r\n";
            Dictionary<string, IDNTGOBulk>.ValueCollection dicIDNTGOBulkTGFc1Col = dicIDNTGOBulkTGFc1.Values;
            foreach (IDNTGOBulk idntgobulk in dicIDNTGOBulkTGFc1Col)
            {
                content += string.Format("{0}\t", idntgobulk.SYMBOL);
                content += string.Format("{0}\t", idntgobulk.DSPLY_NAME);
                content += string.Format("{0}\t", idntgobulk.RIC);
                content += string.Format("{0}\t", idntgobulk.OFFCL_CODE);
                content += string.Format("{0}\t", idntgobulk.EX_SYMBOL);
                content += string.Format("{0}\t", idntgobulk.EXPIR_DATE);
                content += string.Format("{0}\t", idntgobulk.CONTR_MNTH);
                content += string.Format("{0}\t", idntgobulk.CONTR_SIZE);
                content += string.Format("{0}\t", idntgobulk.STRIKE_PRC);
                content += string.Format("{0}\t", idntgobulk.PUTCALLIND);
                content += string.Format("{0}\t", idntgobulk.BCKGRNDPAG);
                content += string.Format("{0}\t", idntgobulk.DSPLY_NMLL);
                content += string.Format("{0}\t", idntgobulk.X_INST_TITLE);
                content += string.Format("{0}\t", idntgobulk.X_80CHAR);
                content += string.Format("{0}\t", idntgobulk.INSTMOD_PUT_CALL);
                content += string.Format("{0}\t", idntgobulk.EXL_NAME);
                content += string.Format("{0}\t", idntgobulk.BCU);
                content += string.Format("{0}\t", idntgobulk.INSTMOD_PROV_SYMB);
                content += "\r\n";
            }
            try
            {
                File.WriteAllText(warrantAddISINTxtFilePath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LogMessage(string.Format("Error happens when generating txt file. Ex: {0} .", ex.Message), Logger.LogType.Error);
            }
            AddResult("File for txt bulk load", warrantAddISINTxtFilePath, "txt file");
        }
        #endregion

        #region GenerateTxtlistIDNTGOBulkTGFc1ROLLOVER
        private void GenerateTxtlistIDNTGOBulkTGFc1ROLLOVER(List<IDNTGOBulk> listIDNTGOBulkTGFc1ROLLOV, string txtFilePath, string txtFileNameMissing)
        {
            if (!string.IsNullOrEmpty(expirDateFormed))
            {
                string warrantAddISINTxtFilePath = Path.Combine(txtFilePath, txtFileNameMissing);
                string content = "SYMBOL\tDSPLY_NAME\tRIC\tOFFCL_CODE\tEX_SYMBOL\tEXPIR_DATE\tCONTR_MNTH\tCONTR_SIZE\tSTRIKE_PRC\tPUTCALLIND\tBCKGRNDPAG\tDSPLY_NMLL\tX_INST_TITLE\tX_80CHAR\t#INSTMOD_PUT_CALL\tEXL_NAME\tBCU\t#INSTMOD_PROV_SYMB\r\n";
                Dictionary<string, IDNTGOBulk>.ValueCollection dicIDNTGOBulkTGFc1Col = dicIDNTGOBulkTGFc1.Values;
                foreach (IDNTGOBulk idntgobulk in listIDNTGOBulkTGFc1ROLLOV)
                {
                    content += string.Format("{0}\t", idntgobulk.SYMBOL);
                    content += string.Format("{0}\t", idntgobulk.DSPLY_NAME);
                    content += string.Format("{0}\t", idntgobulk.RIC);
                    content += string.Format("{0}\t", idntgobulk.OFFCL_CODE);
                    content += string.Format("{0}\t", idntgobulk.EX_SYMBOL);
                    content += string.Format("{0}\t", idntgobulk.EXPIR_DATE);
                    content += string.Format("{0}\t", idntgobulk.CONTR_MNTH);
                    content += string.Format("{0}\t", idntgobulk.CONTR_SIZE);
                    content += string.Format("{0}\t", idntgobulk.STRIKE_PRC);
                    content += string.Format("{0}\t", idntgobulk.PUTCALLIND);
                    content += string.Format("{0}\t", idntgobulk.BCKGRNDPAG);
                    content += string.Format("{0}\t", idntgobulk.DSPLY_NMLL);
                    content += string.Format("{0}\t", idntgobulk.X_INST_TITLE);
                    content += string.Format("{0}\t", idntgobulk.X_80CHAR);
                    content += string.Format("{0}\t", idntgobulk.INSTMOD_PUT_CALL);
                    content += string.Format("{0}\t", idntgobulk.EXL_NAME);
                    content += string.Format("{0}\t", idntgobulk.BCU);
                    content += string.Format("{0}\t", idntgobulk.INSTMOD_PROV_SYMB);
                    content += "\r\n";
                }
                try
                {
                    File.WriteAllText(warrantAddISINTxtFilePath, content, Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    LogMessage(string.Format("Error happens when generating txt file. Ex: {0} .", ex.Message), Logger.LogType.Error);
                }
                AddResult("File for txt bulk load", warrantAddISINTxtFilePath, "txt file");
            }
        }
        #endregion

        #region FilldicCallAndPuts
        private void FilldicCallAndPuts(Dictionary<string, IDNTGOBulk> dicIDNTGOBulkTGTM, List<string> listMonthAndYearFromIDNTGOBulkTGTM)
        {
            Dictionary<string, IDNTGOBulk>.KeyCollection keyCol = dicIDNTGOBulkTGTM.Keys;
            string monthAndYear = string.Empty;
            foreach (string key in keyCol)
            {
                monthAndYear = key.Substring(6, 2);
                if (!listMonthAndYearFromIDNTGOBulkTGTM.Contains(monthAndYear))
                {
                    listMonthAndYearFromIDNTGOBulkTGTM.Add(monthAndYear);
                }
            }
        }
        #endregion

        #region ComparelistIDNTGOBulkTGTMWithlistIDNTGOBulkTGFc1
        private void ComparelistIDNTGOBulkTGTMWithlistIDNTGOBulkTGFc1(Dictionary<string, IDNTGOBulk> dicIDNTGOBulkTGTM, Dictionary<string, IDNTGOBulk> dicIDNTGOBulkTGFc1)
        {
            string[] KeyCol = new string[dicIDNTGOBulkTGFc1.Keys.Count];
            dicIDNTGOBulkTGFc1.Keys.CopyTo(KeyCol, 0);
            foreach (string dicTGFc1Key in KeyCol)
            {
                if (dicIDNTGOBulkTGTM.ContainsKey(dicTGFc1Key))
                {
                    dicIDNTGOBulkTGFc1.Remove(dicTGFc1Key);
                }
            }
        }
        #endregion

        #region GetTodaySETTLEFromGATSTolistIDNTGOBulkTGFc1
        private void GetTodaySETTLEFromGATSTolistIDNTGOBulkTGFc1(string strTGFc1, string patternTGFc1, Dictionary<string, IDNTGOBulk> dicIDNTGOBulkTGFc1, List<string> listMonthAndYearFromIDNTGOBulkTGTM, List<IDNTGOBulk> listIDNTGOBulkTGFc1ROLLOV)
        {
            try
            {
                GatsUtil gats = new GatsUtil();
                string response = gats.GetGatsResponse(strTGFc1, null);
                Regex regex = new Regex(patternTGFc1);
                MatchCollection matches = regex.Matches(response);
                string price = string.Empty;
                List<int> listPrice = new List<int>();
                int high = 0;
                if (matches.Count == 1)
                {
                    int settle = Convert.ToInt32(matches[0].Groups["SETTLE"].Value);
                    if (settle > 0)
                    {
                        IDNTGOBulk idntgobulk = null;
                        IDNTGOBulk idntgobulkROLLOV = null;
                        high = settle - settle % 100;
                        listPrice.Add(high);
                        int reduce = high;
                        int add = high;
                        for (int i = 0; i < 5; i++)
                        {
                            if (reduce <= 2000)
                            {
                                reduce -= 25;
                            }
                            else if (reduce > 2000 && reduce <= 4000)
                            {
                                reduce -= 50;
                            }
                            else
                            {
                                reduce -= 100;
                            }
                            listPrice.Add(reduce);
                        }
                        for (int i = 0; i < 5; i++)
                        {
                            if (add <= 2000)
                            {
                                add += 25;
                            }
                            else if (add > 2000 && add <= 4000)
                            {
                                add += 50;
                            }
                            else
                            {
                                add += 100;
                            }
                            listPrice.Add(add);
                        }
                        foreach (int item in listPrice)
                        {
                            price = item.ToString();
                            Dictionary<string, string>.KeyCollection dicCallsKeys = dicCalls.Keys;
                            Dictionary<string, string>.KeyCollection dicPutsKeys = dicPuts.Keys;
                            foreach (string monthAndYear in listMonthAndYearFromIDNTGOBulkTGTM)
                            {
                                idntgobulk = new IDNTGOBulk();
                                idntgobulk.SYMBOL = "TGO" + (price.Length == 4 ? ("0" + price + monthAndYear) : price + monthAndYear);
                                idntgobulk.DSPLY_NAME = "TG " + (dicCalls.ContainsKey(monthAndYear.Substring(0, 1)) ? dicCalls[monthAndYear.Substring(0, 1)].ToUpper() + monthAndYear.Substring(1, 1) + " " + price + " C" : dicPuts[monthAndYear.Substring(0, 1)].ToUpper() + monthAndYear.Substring(1, 1) + " " + price + " P");
                                idntgobulk.RIC = "TG" + price + monthAndYear + ".TM";
                                idntgobulk.OFFCL_CODE = idntgobulk.SYMBOL;
                                idntgobulk.EX_SYMBOL = idntgobulk.SYMBOL;
                                idntgobulk.EXPIR_DATE = "	 ";
                                idntgobulk.CONTR_MNTH = dicCalls.ContainsKey(monthAndYear.Substring(0, 1)) ? dicCalls[monthAndYear.Substring(0, 1)].ToUpper() + monthAndYear.Substring(1, 1) : dicPuts[monthAndYear.Substring(0, 1)].ToUpper() + monthAndYear.Substring(1, 1);
                                idntgobulk.CONTR_SIZE = "187.5";
                                idntgobulk.STRIKE_PRC = price;
                                idntgobulk.PUTCALLIND = (dicCalls.ContainsKey(monthAndYear.Substring(0, 1)) ? "CA_CALL" : "PU_PUT");
                                idntgobulk.BCKGRNDPAG = "TM01";
                                idntgobulk.DSPLY_NMLL = idntgobulk.DSPLY_NAME;
                                idntgobulk.X_INST_TITLE = "I";
                                idntgobulk.X_80CHAR = "1";
                                idntgobulk.INSTMOD_PUT_CALL = (dicCalls.ContainsKey(monthAndYear.Substring(0, 1)) ? "C" : "P") + "_EU";
                                idntgobulk.EXL_NAME = "TAIFO_OPT_TG";
                                idntgobulk.BCU = "TAIFO_OPT_TG,TAIFO_OPT_TG_" + (dicCalls.ContainsKey(monthAndYear.Substring(0, 1)) ? "C" : "P");
                                idntgobulk.INSTMOD_PROV_SYMB = idntgobulk.SYMBOL;
                                dicIDNTGOBulkTGFc1.Add(idntgobulk.RIC, idntgobulk);
                            }
                            if (!string.IsNullOrEmpty(expirDateFormed))
                            {
                                List<string> keysFromPutAndCall = new List<string>();
                                keysFromPutAndCall.Add(dicCalls.FirstOrDefault(q => q.Value.ToUpper() == monthYear.Substring(0, 3)).Key.ToString());
                                keysFromPutAndCall.Add(dicPuts.FirstOrDefault(q => q.Value.ToUpper() == monthYear.Substring(0, 3)).Key.ToString());
                                string cORp = "C";
                                foreach (string keyMonth in keysFromPutAndCall)
                                {
                                    idntgobulkROLLOV = new IDNTGOBulk();
                                    idntgobulkROLLOV.SYMBOL = "TGO" + (price.Length == 4 ? ("0" + price + keyMonth + monthYear.Substring(3, 1)) : price + keyMonth + monthYear.Substring(3, 1));
                                    idntgobulkROLLOV.DSPLY_NAME = "TG " + monthYear + " " + price + " " + cORp;
                                    idntgobulkROLLOV.RIC = "TG" + price + keyMonth + monthYear.Substring(3, 1) + ".TM";
                                    idntgobulkROLLOV.OFFCL_CODE = idntgobulkROLLOV.SYMBOL;
                                    idntgobulkROLLOV.EX_SYMBOL = idntgobulkROLLOV.SYMBOL;
                                    idntgobulkROLLOV.EXPIR_DATE = expirDateFormed;
                                    idntgobulkROLLOV.CONTR_MNTH = monthYear;
                                    idntgobulkROLLOV.CONTR_SIZE = "187.5";
                                    idntgobulkROLLOV.STRIKE_PRC = price;
                                    idntgobulkROLLOV.PUTCALLIND = cORp.Equals("C") ? "CA_CALL" : "PU_PUT";
                                    idntgobulkROLLOV.BCKGRNDPAG = "TM01";
                                    idntgobulkROLLOV.DSPLY_NMLL = idntgobulkROLLOV.DSPLY_NAME;
                                    idntgobulkROLLOV.X_INST_TITLE = "I";
                                    idntgobulkROLLOV.X_80CHAR = "1";
                                    idntgobulkROLLOV.INSTMOD_PUT_CALL = cORp + "_EU";
                                    idntgobulkROLLOV.EXL_NAME = "TAIFO_OPT_TG";
                                    idntgobulkROLLOV.BCU = "TAIFO_OPT_TG,TAIFO_OPT_TG_" + cORp;
                                    idntgobulkROLLOV.INSTMOD_PROV_SYMB = idntgobulkROLLOV.SYMBOL;
                                    listIDNTGOBulkTGFc1ROLLOV.Add(idntgobulkROLLOV);
                                    cORp = "P";
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("No Value of SETTLE is 0 ! ");
                    }
                }
            }
            catch (Exception e)
            {
                LogMessage("error has happened when running GetTodaySETTLEFromGATS() : " + e.Message, Logger.LogType.Error);
            }
        }
        #endregion

        #region UseListTGTM
        private void UseListTGTM(string tgtm)
        {
            try
            {
                int low = 0;
                int dicRunAfter = dicIDNTGOBulkTGTM.Count;
                int dicRunBefor = -1;
                string strTGTM = string.Empty;
                while (dicRunAfter != dicRunBefor)
                {
                    dicRunBefor = dicIDNTGOBulkTGTM.Count;
                    strTGTM = low + tgtm;
                    GetDataFromGATSTolistIDNTGOBulTGTM(strTGTM, patternTGTM, dicIDNTGOBulkTGTM);
                    low++;
                    dicRunAfter = dicIDNTGOBulkTGTM.Count;
                }
            }
            catch (Exception e)
            {
                LogMessage("error has generated : " + e.Message, Logger.LogType.Error);
            }
        }
        #endregion

        #region GetDataFromGATSTolistIDNTGOBulTGTM
        private void GetDataFromGATSTolistIDNTGOBulTGTM(string strTGTM, string patternTGTM, Dictionary<string, IDNTGOBulk> dicIDNTGOBulkTGTM)
        {
            try
            {
                GatsUtil gats = new GatsUtil();
                string response = gats.GetGatsResponse(strTGTM, null);
                IDNTGOBulk idntgobulk = null;
                Regex regex = new Regex(patternTGTM);
                MatchCollection matches = regex.Matches(response);
                foreach (Match match in matches)
                {
                    idntgobulk = new IDNTGOBulk();
                    idntgobulk.RIC = match.Groups["RIC"].Value;
                    if (!dicIDNTGOBulkTGTM.ContainsKey(idntgobulk.RIC))
                    {
                        dicIDNTGOBulkTGTM.Add(idntgobulk.RIC, idntgobulk);
                    }
                }
            }
            catch (Exception e)
            {
                LogMessage("error happened : " + e.Message, Logger.LogType.Error);
            }
        }
        #endregion
    }
}
