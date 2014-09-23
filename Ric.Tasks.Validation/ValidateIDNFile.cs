using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Validation
{
    #region Configuration
    [ConfigStoredInDB]
    class ValidateIDNFileConfig
    {
        [StoreInDB]
        [Category("FilePath")]
        [Description("GetFilePath like:D:/tmp.txt")]
        public string TxtFilePath { get; set; }
    }
    #endregion

    #region Description
    class ValidateIDNFile : GeneratorBase
    {
        private static ValidateIDNFileConfig configObj = null;
        List<string> listRic = null;
        List<List<string>> listListRic = new List<List<string>>();
        public string strFilePath = string.Empty;
        public string strPatternGATS = string.Empty;
        public string strSaveFilePath = string.Empty;
        protected override void Initialize()
        {
            configObj = Config as ValidateIDNFileConfig;
            strPatternGATS = @"\s+OFFCL_CODE\s+(?<OFFCL_CODE>\S+)\s+\r\n(?<RIC>\S+)\s+PROV_SYMB\s+(?<PROV_SYMB>\S+)\s+";
            strFilePath = configObj.TxtFilePath;
            strSaveFilePath = Path.Combine(strFilePath.Substring(0, strFilePath.LastIndexOf("\\")), DateTime.Now.ToString("yyyy-MMM-dd HH-mm-ss") + ".txt");
        }
    #endregion

        protected override void Start()
        {
            listRic = ReadFileToList(strFilePath);
            GetDataFromGatsToListList(listRic, strPatternGATS, listListRic);
            GenerateFile(listListRic, strSaveFilePath);
        }

        private void GetDataFromGatsToListList(List<string> listRic, string strPatternGATS, List<List<string>> listListRic)
        {
            if (listRic == null || listRic.Count == 0)
            {
                return;
            }
            if (listRic.Count == 1 && string.IsNullOrEmpty(listRic[0].ToString().Trim()))
            {
                return;
            }
            string strQuery = string.Empty;
            int count = listRic.Count;
            int fenMu = 2000;
            int qiuYu = count % fenMu;
            int qiuShang = count / fenMu;
            if (qiuShang > 0)
            {
                for (int i = 0; i < qiuShang; i++)
                {
                    for (int j = 0; j < fenMu; j++)
                    {
                        string strTmp = listRic[i * fenMu + j].ToString().Trim();
                        if (!string.IsNullOrEmpty(strTmp))
                        {
                            strQuery += string.Format(",{0}", strTmp);
                        }
                    }
                    strQuery = strQuery.Remove(0, 1);
                    GetDataFromGATS(strQuery, listListRic, strPatternGATS);
                    strQuery = string.Empty;
                }
            }
            for (int i = qiuShang * fenMu; i < count; i++)
            {
                string strTmp = listRic[i].ToString().Trim();
                if (!string.IsNullOrEmpty(strTmp))
                {
                    strQuery += string.Format(",{0}", strTmp);
                }
            }
            strQuery = strQuery.Remove(0, 1);
            GetDataFromGATS(strQuery, listListRic, strPatternGATS);
        }

        private void GetDataFromGATS(string strQuery, List<List<string>> listListRic, string strPatternGATS)
        {
            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(strQuery, "PROV_SYMB,OFFCL_CODE");
            if (!string.IsNullOrEmpty(response))
            {
                Regex regex = new Regex(strPatternGATS);
                MatchCollection matches = regex.Matches(response);
                string ric = string.Empty;
                string prov = string.Empty;
                string off = string.Empty;
                foreach (Match match in matches)
                {
                    ric = match.Groups["RIC"].Value;
                    prov = match.Groups["PROV_SYMB"].Value;
                    off = match.Groups["OFFCL_CODE"].Value;
                    List<string> li = new List<string>();
                    li.Add(ric);
                    li.Add(prov);
                    li.Add(off);
                    listListRic.Add(li);
                }
            }
            else
            {
                Logger.Log("Too many ric");
                throw new Exception("Input too many RICs a time,GATS dead. ");
            }
        }

        private void GenerateFile(List<List<string>> listList, string strSaveFilePath)
        {
            if (listList != null)
            {
                string content = string.Empty;
                foreach (var str in listList)
                {
                    content += string.Format("{0}\t", str[0]);
                    content += string.Format("{0}\t", str[1]);
                    content += string.Format("{0}\t", str[2]);
                    content += "\r\n";
                }
                try
                {
                    File.WriteAllText(strSaveFilePath, content);
                    TaskResultList.Add(new TaskResultEntry("IDN file", "IDN", strSaveFilePath));
                }
                catch (Exception ex)
                {
                    Logger.Log(string.Format("Error happens when generating file. Ex: {0} .", ex.Message));
                }
            }
        }

        public List<string> ReadFileToList(string filePath)
        {
            if (File.Exists(filePath))
            {
                List<string> tmp = null;
                using (FileStream fs = new FileStream(filePath, FileMode.Open))
                {
                    using (StreamReader sr = new StreamReader(fs))
                    {
                        tmp = new List<string>(sr.ReadToEnd().Replace("\r\n", ",").Split(','));
                        return tmp;
                    }
                }
            }
            return null;
        }
    }
}
