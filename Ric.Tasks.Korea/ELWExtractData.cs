using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.ComponentModel;
using System.IO;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
{
    #region Configuration
    [ConfigStoredInDB]
    class ELWExtractDataConfig
    {
        [StoreInDB]
        [Category("FilePath")]
        [Description("GetFilePath like:D:\tmp.txt")]
        public string TxtFilePath { get; set; }
        [StoreInDB]
        [Category("OutPutPath")]
        [Description("GetFilePath like:D:\\ELW")]
        public string OutPutPath { get; set; }
    }
    #endregion

    #region Description
    class ELWExtractData : GeneratorBase
    {
        private static ELWExtractDataConfig configObj = null;
        List<string> listRic = null;
        List<string> listRicError = new List<string>();
        List<string> listRicNoResponse = new List<string>();
        public string strTxtFilePath = string.Empty;
        public string strOutPutPath = string.Empty;
        List<ELWExtractEntity> listELWEntity = null;
        public DateTime endDate = DateTime.Now.ToUniversalTime().AddHours(+8);
        protected override void Initialize()
        {
            configObj = Config as ELWExtractDataConfig;
            strTxtFilePath = configObj.TxtFilePath;
            strOutPutPath = configObj.OutPutPath;
        }
    #endregion

        protected override void Start()
        {
            listRic = ReadFileToList(strTxtFilePath);//get ric from txt
            listELWEntity = ExtractDataFromWebsite(listRic);//get xls entity from website
            GenerateOutPutFile(listELWEntity, listRicNoResponse, listRicError, strOutPutPath);
        }

        #region OutPutResultFile
        /// <summary>
        /// step 1: if extract data output to xls file
        /// step 2: if generate error ouput to txt file
        /// step 3: if generate no response output txt file
        /// </summary>
        private void GenerateOutPutFile(List<ELWExtractEntity> listELWEntity, List<string> listRicNoResponse, List<string> listRicError, string strOutPutPath)
        {
            if (listELWEntity != null && listELWEntity.Count > 0)
            {
                GenerateXls(listELWEntity, strOutPutPath);
            }
            //this file is about website no response when query a ric on website 
            if (listRicNoResponse != null && listRicNoResponse.Count > 0)
            {
                string strFilePath = Path.Combine(strOutPutPath, DateTime.Today.ToString("yyyy-MM-dd") + "NoResponsePageRIC.txt");
                GenerateTxt(listRicNoResponse, strFilePath);
                AddResult("ELWExtractData",strFilePath,"NoResponsePageRICTxtFile");
            }
            //this file is about website generate error when query a ric on website
            if (listRicError != null && listRicError.Count > 0)
            {
                string strFilePath = Path.Combine(strOutPutPath, DateTime.Today.ToString("yyyy-MM-dd") + "ExtractErrorRIC.txt");
                GenerateTxt(listRicError, strFilePath);
                AddResult("ELWExtractData",strFilePath,"ExtractErrorRICTxtFile");
            }
        }

        private void GenerateTxt(List<string> list, string filePath)
        {
            string content = string.Empty;
            foreach (var str in list)
            {
                content += string.Format("\r\n{0}", str);
            }
            content = content.Remove(0, 1);
            try
            {
                File.WriteAllText(filePath, content);
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("Error happens when generating file. Ex: {0} .", ex.Message));
            }
        }

        private void GenerateXls(List<ELWExtractEntity> listELWEntity, string strOutPutPath)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
            }
            try
            {
                string strFileName = DateTime.Today.ToString("yyyy-MM-dd") + "_ELW.xls";
                string strFilePath = Path.Combine(strOutPutPath, strFileName);
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, strFilePath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                }
                //first table left on ELW website
                wSheet.Cells[1, 1] = "발행기관코드";
                wSheet.Cells[1, 2] = "표준코드";
                wSheet.Cells[1, 3] = "한글종목명";
                wSheet.Cells[1, 4] = "금융상품구분";
                wSheet.Cells[1, 5] = "상장여부";
                wSheet.Cells[1, 6] = "상장일";
                wSheet.Cells[1, 7] = "발행수량(워런트)";
                wSheet.Cells[1, 8] = "발행단가";
                wSheet.Cells[1, 9] = "발행통화";
                wSheet.Cells[1, 10] = "발행형태";
                wSheet.Cells[1, 11] = "표준/비표준";
                //first table right on ELW website
                wSheet.Cells[1, 12] = "발행기관명";
                wSheet.Cells[1, 13] = "단축코드";
                wSheet.Cells[1, 14] = "영문종목명";
                wSheet.Cells[1, 15] = "발행회차(회)";
                wSheet.Cells[1, 16] = "활성여부";
                wSheet.Cells[1, 17] = "상장폐지일";
                wSheet.Cells[1, 18] = "발행일";
                wSheet.Cells[1, 19] = "만기일";
                wSheet.Cells[1, 20] = "발행구분";
                wSheet.Cells[1, 21] = "전환비율";
                wSheet.Cells[1, 22] = "권리형태";
                //second table on ELW website
                wSheet.Cells[1, 23] = "기초자산종류";
                wSheet.Cells[1, 24] = "주권발행기관";
                wSheet.Cells[1, 25] = "주가지수종류";
                wSheet.Cells[1, 26] = "기초자산기타";
                //third table on ELW website
                wSheet.Cells[1, 27] = "권리유형";
                wSheet.Cells[1, 28] = "특이발행조건";
                wSheet.Cells[1, 29] = "권리행사방식";
                //fourth table on ELW website
                wSheet.Cells[1, 30] = "CFI";

                int startLine = 2;
                foreach (var tmp in listELWEntity)
                {
                    //first table left on ELW website
                    wSheet.Cells[startLine, 1] = tmp.IssuingAuthorityCongenial;
                    wSheet.Cells[startLine, 2] = tmp.StandardCongenial;
                    wSheet.Cells[startLine, 3] = tmp.KoreanProjectName;
                    wSheet.Cells[startLine, 4] = tmp.FinancialProducts;
                    wSheet.Cells[startLine, 5] = tmp.ListedOrNot;
                    wSheet.Cells[startLine, 6] = tmp.Listed;
                    wSheet.Cells[startLine, 7] = tmp.IssueNumber;
                    wSheet.Cells[startLine, 8] = tmp.ReleaseTheUnitPrice;
                    wSheet.Cells[startLine, 9] = tmp.Money;
                    wSheet.Cells[startLine, 10] = tmp.ReleaseForm;
                    wSheet.Cells[startLine, 11] = tmp.StandardNonStandard;
                    //first table left on ELW website
                    wSheet.Cells[startLine, 12] = tmp.IssuingAuthority;
                    wSheet.Cells[startLine, 13] = tmp.ShortenTheCongenial;
                    wSheet.Cells[startLine, 14] = tmp.TheProjectNameEnglish;
                    wSheet.Cells[startLine, 15] = tmp.ToIssue;
                    wSheet.Cells[startLine, 16] = tmp.WhetherTheActivity;
                    wSheet.Cells[startLine, 17] = tmp.ListedUntil;
                    wSheet.Cells[startLine, 18] = tmp.TheDate;
                    wSheet.Cells[startLine, 19] = tmp.TheExpirationOfThe;
                    wSheet.Cells[startLine, 20] = tmp.IssueToDistinguish;
                    wSheet.Cells[startLine, 21] = tmp.ConversionRatio;
                    wSheet.Cells[startLine, 22] = tmp.ThePowerForm;
                    //second table on ELW website
                    wSheet.Cells[startLine, 23] = tmp.UnderlyingAssetTypes;
                    wSheet.Cells[startLine, 24] = tmp.SovereignIssuingAuthority;
                    wSheet.Cells[startLine, 25] = tmp.StockIndexTypes;
                    wSheet.Cells[startLine, 26] = tmp.UnderlyingAssetGuitar;
                    //third table on ELW website
                    wSheet.Cells[startLine, 27] = tmp.TheRightType;
                    wSheet.Cells[startLine, 28] = tmp.TheIssueOfTheSpecialConditions;
                    wSheet.Cells[startLine, 29] = tmp.TheExerciseOfTheRightWay;
                    //fourth table on ELW website
                    wSheet.Cells[1, 30] = tmp.CFI;

                    startLine++;
                }
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
                AddResult("ELWExtractData",strFilePath,"XLS");
            }
            catch (Exception ex)
            {
                string msg = "Error found in NDA T&C file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }
        #endregion

        #region ExtractData
        /// <summary>
        /// step 1: check config info
        /// step 2: get postdata from website
        /// step 3: extract data from website
        /// </summary>
        private List<ELWExtractEntity> ExtractDataFromWebsite(List<string> listRic)
        {
            List<ELWExtractEntity> listELWEntity = new List<ELWExtractEntity>();
            //string strStartDate = startDate.ToString("yyyyMMdd");
            string strStartDate = "";
            string strEndDate = endDate.ToString("yyyyMMdd");
            if (listRic == null || listRic.Count == 0)
            {
                Logger.Log(string.Format("can't get ric from txt file. please check txt file."));
                return null;
            }
            if (listRic.Count == 1 && listRic[0].Trim() == "")
            {
                Logger.Log(string.Format("txt file is empty."));
                return null;
            }
            foreach (string strRic in listRic)
            {
                if (string.IsNullOrEmpty(strRic.Trim()))
                {
                    Logger.Log(string.Format("This ric is invalid."));
                    continue;
                }
                ELWExtractEntity elwTmp = GetELWExtractEntityStepOne(strStartDate, strEndDate, strRic);
                if (elwTmp == null)
                {
                    Logger.Log(string.Format("error when search Ric{0} from {1} to {2}.", strRic, strStartDate, strEndDate));
                    continue;
                }
                listELWEntity.Add(elwTmp);
            }
            return listELWEntity;
        }

        private ELWExtractEntity GetELWExtractEntityStepOne(string strStartDate, string strEndDate, string strRic)
        {
            string strUrl = @"http://isin.krx.co.kr/srch/srch.do?method=srchList";
            string strPostData = string.Format("std_cd_grnt_start_dd={0}&std_cd_grnt_end_dd={1}&std_cd_grnt=1&searchRadio2=11&searchRadio1=11&searchRadio=11&list_start_dd=&list_end_dd=&listRadio=1&isu_start_dd=&isu_end_dd=&isuRadio=1&isur_cd=&isur_nm={2}&com_nm=", strStartDate, strEndDate, strRic);
            string strPageSource = string.Empty;
            ELWExtractEntity tmp = new ELWExtractEntity();
            try
            {
                AdvancedWebClient wc = new AdvancedWebClient();
                HtmlDocument htc = new HtmlDocument();
                strPageSource = WebClientUtil.GetPageSource(wc, strUrl, 300000, strPostData);
                if (string.IsNullOrEmpty(strPageSource))
                {
                    Logger.Log(string.Format("return response is null,when query ric:{0}", strRic));
                    if (!listRicNoResponse.Contains(strRic))
                        listRicNoResponse.Add(strRic);
                    return null;
                }
                htc.LoadHtml(strPageSource);
                HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                HtmlNode table = tables[1];
                HtmlNodeCollection trs = table.SelectNodes(".//tr");
                if (trs.Count == 1)
                {
                    Logger.Log(string.Format("can't get ric in strPageSource,when query ric:{0}", strRic));
                    if (!listRicNoResponse.Contains(strRic))
                        listRicNoResponse.Add(strRic);
                    return null;
                }
                else if (trs.Count >= 2)
                {
                    string std_cd = trs[1].SelectNodes(".//td")[1].InnerText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();
                    string isu_nm = trs[1].SelectNodes(".//td")[3].InnerText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();
                    tmp = GetGetELWExtractEntityStepTwo(std_cd, isu_nm);
                }
                return tmp;
            }
            catch (Exception ex)
            {
                if (!listRicError.Contains(strRic))
                    listRicError.Add(strRic);
                Logger.Log(string.Format("Error found in function: {0}. Exception message: {1}", "GetELWExtractEntityStepOne", ex.Message));
                return null;
            }
        }

        private ELWExtractEntity GetGetELWExtractEntityStepTwo(string std_cd, string isu_nm)
        {
            string strUrl = @"http://isin.krx.co.kr/srch/srch.do?method=srchPopup11";
            string stdcd_type = "11";
            string mod_del_cd = "";
            string pershr_isu_prc = "";
            string isu_shrs = "";
            string strPostData = string.Format("stdcd_type={0}&std_cd={1}&mod_del_cd={2}&isu_nm={3}&pershr_isu_prc={4}&isu_shrs={5}", stdcd_type, std_cd, mod_del_cd, isu_nm, pershr_isu_prc, isu_shrs);
            string strPageSource = string.Empty;
            ELWExtractEntity tmp = new ELWExtractEntity();
            try
            {
                AdvancedWebClient wc = new AdvancedWebClient();
                HtmlDocument htc = new HtmlDocument();
                strPageSource = WebClientUtil.GetPageSource(wc, strUrl, 300000, strPostData);
                if (string.IsNullOrEmpty(strPageSource))
                {
                    Logger.Log(string.Format("return response is null,when query ric:{0}", std_cd));
                    if (!listRicNoResponse.Contains(std_cd))
                        listRicNoResponse.Add(std_cd);
                    return null;
                }
                htc.LoadHtml(strPageSource);
                HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");
                if (tables.Count < 5)
                {
                    Logger.Log(string.Format("tables.count<5,so missing data in the pageSource ,current ric:", isu_nm));
                    if (!listRicNoResponse.Contains(isu_nm))
                        listRicNoResponse.Add(isu_nm);
                    return null;
                }
                HtmlNode table1 = tables[1];
                HtmlNodeCollection trs1 = table1.SelectNodes(".//tr");//11*2=22
                HtmlNode table2 = tables[2];
                HtmlNodeCollection trs2 = table2.SelectNodes(".//tr");//0*2+1*1+2*1=4
                HtmlNode table3 = tables[3];
                HtmlNodeCollection trs3 = table3.SelectNodes(".//tr");//3*1=3
                HtmlNode table4 = tables[4];
                HtmlNodeCollection trs4 = table4.SelectNodes(".//tr");//1*1=1
                if (trs1.Count < 11 || trs2.Count < 3 || trs3.Count < 3 || trs4.Count < 1)
                {
                    Logger.Log(string.Format("trs1.count is too small,so missing data in the pageSource ,current ric:", isu_nm));
                    if (!listRicNoResponse.Contains(isu_nm))
                        listRicNoResponse.Add(isu_nm);
                    return null;
                }
                //first table left on ELW website
                tmp.IssuingAuthorityCongenial = FormatInnerText(trs1[0].SelectNodes(".//td")[0].InnerText);
                tmp.StandardCongenial = FormatInnerText(trs1[1].SelectNodes(".//td")[0].InnerText);
                tmp.KoreanProjectName = FormatInnerText(trs1[2].SelectNodes(".//td")[0].InnerText);
                tmp.FinancialProducts = FormatInnerText(trs1[3].SelectNodes(".//td")[0].InnerText);
                tmp.ListedOrNot = FormatInnerText(trs1[4].SelectNodes(".//td")[0].InnerText);
                tmp.Listed = FormatInnerText(trs1[5].SelectNodes(".//td")[0].InnerText);
                tmp.IssueNumber = FormatInnerText(trs1[6].SelectNodes(".//td")[0].InnerText);
                tmp.ReleaseTheUnitPrice = FormatInnerText(trs1[7].SelectNodes(".//td")[0].InnerText);
                tmp.Money = FormatInnerText(trs1[8].SelectNodes(".//td")[0].InnerText);
                tmp.ReleaseForm = FormatInnerText(trs1[9].SelectNodes(".//td")[0].InnerText);
                tmp.StandardNonStandard = FormatInnerText(trs1[10].SelectNodes(".//td")[0].InnerText);
                //first table right on ELW website
                tmp.IssuingAuthority = FormatInnerText(trs1[0].SelectNodes(".//td")[1].InnerText);
                tmp.ShortenTheCongenial = FormatInnerText(trs1[1].SelectNodes(".//td")[1].InnerText);
                tmp.TheProjectNameEnglish = FormatInnerText(trs1[2].SelectNodes(".//td")[1].InnerText);
                tmp.ToIssue = FormatInnerText(trs1[3].SelectNodes(".//td")[1].InnerText);
                tmp.WhetherTheActivity = FormatInnerText(trs1[4].SelectNodes(".//td")[1].InnerText);
                tmp.ListedUntil = FormatInnerText(trs1[5].SelectNodes(".//td")[1].InnerText);
                tmp.TheDate = FormatInnerText(trs1[6].SelectNodes(".//td")[1].InnerText);
                tmp.TheExpirationOfThe = FormatInnerText(trs1[7].SelectNodes(".//td")[1].InnerText);
                tmp.IssueToDistinguish = FormatInnerText(trs1[8].SelectNodes(".//td")[1].InnerText);
                tmp.ConversionRatio = FormatInnerText(trs1[9].SelectNodes(".//td")[1].InnerText);
                tmp.ThePowerForm = FormatInnerText(trs1[10].SelectNodes(".//td")[1].InnerText);
                //second table on ELW website
                tmp.UnderlyingAssetTypes = FormatInnerText(trs2[0].SelectNodes(".//td")[0].InnerText);
                tmp.SovereignIssuingAuthority = FormatInnerText(trs2[0].SelectNodes(".//td")[1].InnerText);
                tmp.StockIndexTypes = FormatInnerText(trs2[1].SelectNodes(".//td")[0].InnerText);
                tmp.UnderlyingAssetGuitar = FormatInnerText(trs2[2].SelectNodes(".//td")[0].InnerText);
                //third table on ELW website
                tmp.TheRightType = FormatInnerText(trs3[0].SelectNodes(".//td")[0].InnerText);
                tmp.TheIssueOfTheSpecialConditions = FormatInnerText(trs3[0].SelectNodes(".//td")[1].InnerText);
                tmp.TheExerciseOfTheRightWay = FormatInnerText(trs3[1].SelectNodes(".//td")[0].InnerText);
                //fourth table on ELW website
                tmp.CFI = FormatInnerText(trs4[0].SelectNodes(".//td")[0].InnerText);
                return tmp;
            }
            catch (Exception ex)
            {
                if (!listRicError.Contains(std_cd))
                    listRicError.Add(std_cd);
                Logger.Log(string.Format("Error found in function: {0}. Exception message: {1}", "GetELWExtractEntityStepOne", ex.Message));
                return null;
            }
        }

        private string FormatInnerText(string iText)
        {
            return iText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();
        }
        #endregion

        #region ReadTxtFile
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
            Logger.Log(string.Format("error when read txt file,please check configuration"));
            return null;
        }
        #endregion
    }

    #region XlsEntity
    class ELWExtractEntity
    {
        //first table left on ELW website
        public string IssuingAuthorityCongenial { get; set; }//1
        public string StandardCongenial { get; set; }//2                  B
        public string KoreanProjectName { get; set; }//3
        public string FinancialProducts { get; set; }//4
        public string ListedOrNot { get; set; }//5
        public string Listed { get; set; }//6
        public string IssueNumber { get; set; }//7                         G
        public string ReleaseTheUnitPrice { get; set; }//8                H
        public string Money { get; set; }//9
        public string ReleaseForm { get; set; }//10
        public string StandardNonStandard { get; set; }//11
        //first table right on ELW website
        public string IssuingAuthority { get; set; }//12
        public string ShortenTheCongenial { get; set; }//13               M
        public string TheProjectNameEnglish { get; set; }//14            N
        public string ToIssue { get; set; }//15
        public string WhetherTheActivity { get; set; }//16
        public string ListedUntil { get; set; }//17
        public string TheDate { get; set; }//18                           R
        public string TheExpirationOfThe { get; set; }//19              S
        public string IssueToDistinguish { get; set; }//20
        public string ConversionRatio { get; set; }//21                 U
        public string ThePowerForm { get; set; }//22
        //second table on ELW website
        public string UnderlyingAssetTypes { get; set; }//23
        public string SovereignIssuingAuthority { get; set; }//24
        public string StockIndexTypes { get; set; }//25
        public string UnderlyingAssetGuitar { get; set; }//26
        //third table on ELW website
        public string TheRightType { get; set; }//27    AA
        public string TheIssueOfTheSpecialConditions { get; set; }//28  
        public string TheExerciseOfTheRightWay { get; set; }//29
        //fourth table on ELW website
        public string CFI { get; set; }//30
    }
    #endregion
}
