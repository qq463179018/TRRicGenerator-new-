using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.ComponentModel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using HtmlAgilityPack;
using System.Net;
using Ric.Db.Manager;
using Ric.Db.Info;
using System.Drawing;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.Korea
{
    #region Configuration
    [ConfigStoredInDB]
    class KoreaIPOConfig
    {
        [StoreInDB]
        [Category("Date")]
        [Description("Start Search Date eg:yyyy-MM-dd")]
        public string StartDate { get; set; }

        [StoreInDB]
        [Category("Date")]
        [Description("End Search Date eg:yyyy-MM-dd")]
        public string EndDate { get; set; }

        [StoreInDB]
        [Category("FilePath")]
        [Description("Dowanload And Generate File Path eg:D:\tmp.txt")]
        public string FilePath { get; set; }
    }
    #endregion

    class KoreaIPO : GeneratorBase
    {
        #region Description
        private static KoreaIPOConfig configObj = null;
        private string strStartDate = string.Empty;
        private string strEndDate = string.Empty;
        private string strFilePath = string.Empty;
        private string strDownLoadFilePath = string.Empty;

        List<FMELWEntity> listFMELW = new List<FMELWEntity>();
        List<string> listError = new List<string>();//error isin and issue anthority     post's problem
        List<string> listNoResponse = new List<string>();//no response isin and issue anthority   website's problem

        protected override void Initialize()
        {
            configObj = Config as KoreaIPOConfig;
            strStartDate = configObj.StartDate.Replace("-", "").Trim();
            strEndDate = configObj.EndDate.Replace("-", "").Trim();
            strFilePath = configObj.FilePath;
            strDownLoadFilePath = Path.Combine(strFilePath, string.Format("{0}To{1}_{2}", strStartDate, strEndDate, "KoreaIPO.xls"));
        }
        #endregion

        protected override void Start()
        {
            DownloadFile(strStartDate, strEndDate, strDownLoadFilePath);
            listFMELW = ReadFile(strDownLoadFilePath, listFMELW);
            ExtractDataStepOne(listFMELW);
            FormatEntityData(listFMELW);
            GenerateResultFile(strFilePath, listFMELW);
        }

        private void DownloadFile(string strStartDate, string strEndDate, string strFilePath)
        {
            string post = File.ReadAllText(@"Config\Korea\KoreaIPOPostData.txt", Encoding.UTF8);
            string url = "http://isin.krx.co.kr/srch/srch.do?method=srchDownElwList";
            post = string.Format(post, strStartDate, strEndDate);

            try
            {
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version11;
                request.Timeout = 100000;
                request.UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)";
                request.Method = "POST";
                request.KeepAlive = true;
                request.Headers["Accept-Language"] = "en-US";
                request.Accept = "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*";
                request.ContentType = "multipart/form-data; boundary=---------------------------7de29a81c0e16";
                byte[] buf = Encoding.UTF8.GetBytes(post);
                request.ContentLength = buf.Length;
                request.GetRequestStream().Write(buf, 0, buf.Length);
                HttpWebResponse httpResponse = (HttpWebResponse)request.GetResponse();
                Stream page3 = httpResponse.GetResponseStream();

                using (Stream file = File.Create(strFilePath))
                {
                    byte[] buffer = new byte[8 * 1024];
                    int len;
                    int offset = 0;

                    while ((len = page3.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        file.Write(buffer, 0, len);
                        offset += len;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("can't download file message:{0}", ex.ToString());
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        private void GenerateResultFile(string strFilePath, List<FMELWEntity> listFMELW)
        {
            if (listFMELW != null && listFMELW.Count > 0)
            {
                GenerateXls(listFMELW, strFilePath);
            }

            //this file is about website no response when query a ric on website 
            if (listNoResponse != null && listNoResponse.Count > 0)
            {
                string noResponsePath = Path.Combine(strFilePath, string.Format("{0}To{1}_{2}", strStartDate, strEndDate, "NoResponsePageRIC.txt"));
                GenerateTxt(listNoResponse, strFilePath);
                AddResult("ELWExtractData",strFilePath,"NoResponsePageRICTxtFile");
            }

            //this file is about website generate error when query a ric on website
            if (listError != null && listError.Count > 0)
            {
                string errorPath = Path.Combine(strFilePath, string.Format("{0}To{1}_{2}", strStartDate, strEndDate, "ExtractErrorRIC.txt"));
                GenerateTxt(listError, strFilePath);
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

        private void GenerateXls(List<FMELWEntity> listFMELW, string strFilePath)
        {
            if (listFMELW == null || listFMELW.Count == 0)
            {
                Logger.Log("listFMELW is null or empty!", Logger.LogType.Warning);
                return;
            }

            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);

            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                string filePath = Path.Combine(strFilePath, string.Format("{0}To{1}_{2}", strStartDate, strEndDate, "ELW.xls"));
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;

                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                FillExcelTitle(wSheet);
                FillExcelBody(wSheet, listFMELW);
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
                AddResult("KoreaIPO",filePath,"ExtractXLS");
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

        private void FillExcelBody(Worksheet wSheet, List<FMELWEntity> listFMELW)
        {
            int startLine = 2;

            foreach (var tmp in listFMELW)
            {
                if (!tmp.ReleaseForm.Equals("공모"))
                {
                    continue;
                }

                ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "dd-MMM-yy";
                wSheet.Cells[startLine, 1] = tmp.UpdatedDate;
                wSheet.Cells[startLine, 2] = tmp.EffectiveDate;
                wSheet.Cells[startLine, 3] = tmp.RIC;
                wSheet.Cells[startLine, 4] = tmp.FM;
                wSheet.Cells[startLine, 5] = tmp.IDNDisplayName;
                wSheet.Cells[startLine, 6] = tmp.ISIN;
                wSheet.Cells[startLine, 7] = tmp.Ticker;
                wSheet.Cells[startLine, 8] = tmp.BCAST_REF;
                wSheet.Cells[startLine, 9] = tmp.QACommonName;
                ((Range)wSheet.Cells[startLine, 10]).NumberFormat = "dd-MMM-yy";
                wSheet.Cells[startLine, 10] = tmp.MatDate;
                wSheet.Cells[startLine, 11] = tmp.StrikePrice;
                wSheet.Cells[startLine, 12] = tmp.QuanityOfWarrants;
                wSheet.Cells[startLine, 13] = tmp.IssuePrice;
                ((Range)wSheet.Cells[startLine, 14]).NumberFormat = "dd-MMM-yy";
                wSheet.Cells[startLine, 14] = tmp.IssueDate;
                wSheet.Cells[startLine, 15] = tmp.ConversionRatio;
                wSheet.Cells[startLine, 16] = tmp.Issuer;
                wSheet.Cells[startLine, 17] = tmp.KoreaWarrantName;
                wSheet.Cells[startLine, 18] = tmp.Chain;
                wSheet.Cells[startLine, 19] = tmp.LastTradingDate;
                wSheet.Cells[startLine, 20] = tmp.KnockOutPrice;
                startLine++;
            }
        }

        private void FillExcelTitle(Worksheet wSheet)
        {
            ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 25;
            ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 25;
            ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 60;
            ((Range)wSheet.Columns["Q", System.Type.Missing]).ColumnWidth = 30;
            ((Range)wSheet.Columns["R", System.Type.Missing]).ColumnWidth = 60;
            ((Range)wSheet.Columns["S", System.Type.Missing]).ColumnWidth = 16;
            ((Range)wSheet.Columns["T", System.Type.Missing]).ColumnWidth = 16;

            wSheet.Cells[1, 1] = "Updated Date";
            wSheet.Cells[1, 2] = "Effective Date";
            wSheet.Cells[1, 3] = "RIC";
            wSheet.Cells[1, 4] = "FM";
            wSheet.Cells[1, 5] = "IDN Display Name";
            wSheet.Cells[1, 6] = "ISIN";
            wSheet.Cells[1, 7] = "Ticker";
            wSheet.Cells[1, 8] = "BCAST_REF";
            wSheet.Cells[1, 9] = "QA Common Name";
            wSheet.Cells[1, 10] = "Mat Date";
            wSheet.Cells[1, 11] = "Strike Price";
            wSheet.Cells[1, 12] = "Quanity of Warrants";
            wSheet.Cells[1, 13] = "Issue Price";
            wSheet.Cells[1, 14] = "Issue Date";
            wSheet.Cells[1, 15] = "Conversion Ratio";
            wSheet.Cells[1, 16] = "Issuer";
            wSheet.Cells[1, 17] = "Korea Warrant Name";
            wSheet.Cells[1, 18] = "Chain";
            wSheet.Cells[1, 19] = "Last Trading Date";
            wSheet.Cells[1, 20] = "Knock-out Price";
            ((Range)wSheet.Columns["A:T", System.Type.Missing]).Font.Name = "Arail";
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
        }

        private void FormatEntityData(List<FMELWEntity> listFMELW)
        {
            if (listFMELW == null || listFMELW.Count == 0)
            {
                string msg = "there is not useful data in the file";
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }

            string updateDate = DateTime.Now.ToString("dd-MMM-yy");
            string updateYear = DateTime.Now.ToString("yyyy");
            int lengthTicker = 0;

            try
            {
                foreach (var fm in listFMELW)
                {
                    lengthTicker = fm.Ticker.Length;
                    DateTime matDate = DateTime.Parse(fm.MatDate);
                    DateTime issueDate = DateTime.Parse(fm.IssueDate);
                    fm.Ticker = fm.Ticker.Substring(lengthTicker - 6, 6);
                    fm.UpdatedDate = updateDate;
                    fm.EffectiveDate = updateYear;
                    fm.RIC = fm.Ticker + ".KS";
                    fm.FM = "1";
                    string bcastRef = string.Empty;

                    if (string.IsNullOrEmpty(fm.XSovereignIssuingAuthority.Trim()))
                    {
                        bcastRef = fm.YStockIndexTypes.Trim();
                    }
                    else
                    {
                        bcastRef = fm.XSovereignIssuingAuthority.Trim();
                        bcastRef = bcastRef.Replace("(주)", "").Replace("구성주식비율: 1", "");
                    }

                    KoreaUnderlyingInfo underlying = KoreaUnderlyingManager.SelectUnderlying(bcastRef);
                    KoreaIssuerInfo issuer = KoreaIssuerManager.SelectIssuerByIssuerCode2(fm.Ticker.Trim().Substring(0, 2));
                    fm.IDNDisplayName += issuer.IssuerCode4;
                    fm.IDNDisplayName += fm.Ticker.Substring(2, 4);
                    fm.IDNDisplayName += underlying.IDNDisplayNamePart;

                    if (fm.KoreaWarrantName.Trim() == "콜")
                    {
                        fm.IDNDisplayName += "C";    //  "C" ***************** KBIS  +4019 + underlying.IDNDisplayNamePart  +C
                    }
                    else
                    {
                        fm.IDNDisplayName += "P";  // "P"**************** KBIS  +4019 + underlying.IDNDisplayNamePart    +P
                    }

                    fm.BCAST_REF = underlying.UnderlyingRIC;
                    fm.QACommonName = " ";
                    fm.MatDate = matDate.ToString("dd-MMM-yy");
                    fm.StrikePrice = "";
                    fm.IssueDate = issueDate.ToString("dd-MMM-yy");
                    fm.Issuer = fm.Issuer.ToUpper();
                    string koreanWarrantName = fm.KoreaWarrantName;
                    fm.KoreaWarrantName = string.Empty;
                    fm.KoreaWarrantName += issuer.KoreaIssuerName;
                    fm.KoreaWarrantName += fm.Ticker.Substring(2, 4);
                    fm.KoreaWarrantName += underlying.KoreaNameDrop;
                    fm.KoreaWarrantName += koreanWarrantName;

                    if (fm.BCAST_REF.Trim() == ".KS200")
                    {
                        fm.Chain = "0#WARRANTS.KS, 0#ELW.KS, 0#.KS200W.KS";
                    }
                    else
                    {
                        fm.Chain = "0#WARRANTS.KS, 0#ELW.KS, 0#CELW.KS,0#" + fm.BCAST_REF.Substring(0, 6) + "W.KS";
                    }

                    fm.LastTradingDate = "";
                    fm.KnockOutPrice = "";
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("error happened when format result entity :{0}", ex.Message);
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void ExtractDataStepOne(List<FMELWEntity> listFMELW)
        {
            if (listFMELW == null || listFMELW.Count == 0)
            {
                string msg = "there is not useful data in the file";
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }

            foreach (var fm in listFMELW)
            {
                ExtractDataStepTwo(fm);
            }
        }

        private List<FMELWEntity> ReadFile(string filePath, List<FMELWEntity> listFMELW)
        {
            List<FMELWEntity> list = new List<FMELWEntity>();
            if (!File.Exists(filePath))
            {
                string msg = string.Format("the file [{0}] is not exist.", filePath);
                Logger.Log(msg, Logger.LogType.Error);
                return null;
            }

            try
            {
                using (ExcelApp eApp = new ExcelApp(false, false))
                {
                    Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(eApp, filePath);
                    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;

                    if (wSheet == null)
                    {
                        string msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                        Logger.Log(msg, Logger.LogType.Error);
                        return null;
                    }

                    //wSheet.Name = "FM";
                    int lastUsedRow = wSheet.UsedRange.Row + wSheet.UsedRange.Rows.Count - 1;//193=1+193-1

                    using (ExcelLineWriter reader = new ExcelLineWriter(wSheet, 2, 2, ExcelLineWriter.Direction.Right))
                    {
                        while (reader.Row <= lastUsedRow)
                        {
                            FMELWEntity fm = new FMELWEntity();
                            fm.ISIN = reader.ReadLineCellText();//web.StandardCongenial

                            reader.PlaceNext(reader.Row, 4);
                            fm.IssuingAuthority = reader.ReadLineValue2();

                            reader.PlaceNext(reader.Row + 1, 2);

                            if (!string.IsNullOrEmpty(fm.ISIN.Trim()) && !string.IsNullOrEmpty(fm.IssuingAuthority.Trim()))
                            {
                                list.Add(fm);
                            }
                        }
                    }
                }
                return list;
            }
            catch (Exception ex)
            {
                string msg = string.Format("read xls file error :{0}", ex.ToString());
                Logger.Log(msg);
                return null;
            }
        }

        private void ExtractDataStepTwo(FMELWEntity fm)
        {
            string isu_nm = fm.IssuingAuthority.Trim();
            string std_cd = fm.ISIN.Trim();
            string strUrl = @"http://isin.krx.co.kr/srch/srch.do?method=srchPopup11";
            string stdcd_type = "11";
            string mod_del_cd = "";
            string pershr_isu_prc = "";
            string isu_shrs = "";
            string strPostData = string.Format("stdcd_type={0}&std_cd={1}&mod_del_cd={2}&isu_nm={3}&pershr_isu_prc={4}&isu_shrs={5}", stdcd_type, std_cd, mod_del_cd, isu_nm, pershr_isu_prc, isu_shrs);
            string strPageSource = string.Empty;
            //string strReleaseForm = string.Empty;
            try
            {
                AdvancedWebClient wc = new AdvancedWebClient();
                HtmlDocument htc = new HtmlDocument();
                strPageSource = WebClientUtil.GetPageSource(wc, strUrl, 300000, strPostData);

                if (string.IsNullOrEmpty(strPageSource))
                {
                    Logger.Log(string.Format("return response is null,when query ric:{0}", std_cd));

                    if (!listNoResponse.Contains(std_cd))
                        listNoResponse.Add(std_cd);

                    return;
                }

                htc.LoadHtml(strPageSource);
                HtmlNodeCollection tables = htc.DocumentNode.SelectNodes(".//table");

                if (tables.Count < 5)
                {
                    Logger.Log(string.Format("tables.count<5,so missing data in the pageSource ,current ric:", isu_nm));

                    if (!listNoResponse.Contains(isu_nm))
                        listNoResponse.Add(isu_nm);

                    return;
                }

                HtmlNode table1 = tables[1];
                HtmlNodeCollection trs1 = table1.SelectNodes(".//tr");//11*2=22
                HtmlNode table2 = tables[2];
                HtmlNodeCollection trs2 = table2.SelectNodes(".//tr");//0*2+1*1+2*1=4
                HtmlNode table3 = tables[3];
                HtmlNodeCollection trs3 = table3.SelectNodes(".//tr");//3*1=3

                if (trs1.Count < 11 || trs2.Count < 3 || trs3.Count < 3)
                {
                    Logger.Log(string.Format("trs1.count is too small,so missing data in the pageSource ,current ric:", isu_nm));

                    if (!listNoResponse.Contains(isu_nm))
                        listNoResponse.Add(isu_nm);

                    return;
                }

                fm.ReleaseForm = FormatInnerText(trs1[10].SelectNodes(".//td")[0].InnerText);
                //strReleaseForm = FormatInnerText(trs1[9].SelectNodes(".//td")[0].InnerText);
                //if (!strReleaseForm.Equals("공모"))
                //{
                //    return;
                //}

                //first table left on ELW website 
                fm.QuanityOfWarrants = FormatInnerText(trs1[6].SelectNodes(".//td")[0].InnerText);//quanity of warrants
                fm.IssuePrice = FormatInnerText(trs1[7].SelectNodes(".//td")[0].InnerText);//issue price

                //first table right on ELW website
                fm.Ticker = FormatInnerText(trs1[1].SelectNodes(".//td")[1].InnerText);//ticker
                fm.Issuer = FormatInnerText(trs1[2].SelectNodes(".//td")[1].InnerText);//issuer
                fm.IssueDate = FormatInnerText(trs1[6].SelectNodes(".//td")[1].InnerText);//issuer date
                fm.MatDate = FormatInnerText(trs1[7].SelectNodes(".//td")[1].InnerText);//mat date
                fm.ConversionRatio = FormatInnerText(trs1[9].SelectNodes(".//td")[1].InnerText);//conversion ratio

                //second table on ELW website
                fm.XSovereignIssuingAuthority = FormatInnerText(trs2[0].SelectNodes(".//td")[1].InnerText);//X
                fm.YStockIndexTypes = FormatInnerText(trs2[1].SelectNodes(".//td")[0].InnerText);//Y

                //third table on ELW website
                fm.KoreaWarrantName = FormatInnerText(trs3[0].SelectNodes(".//td")[0].InnerText);//korea warrant name
            }
            catch (Exception ex)
            {
                if (!listError.Contains(std_cd))
                    listError.Add(std_cd);

                Logger.Log(string.Format("Error found in function: {0}. Exception message: {1}", "GetELWExtractEntityStepOne", ex.Message));
                return;
            }
        }

        private string FormatInnerText(string iText)
        {
            return iText.Replace("&nbsp;", "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Trim();
        }
    }

    class FMELWEntity
    {
        public string UpdatedDate { get; set; }
        public string EffectiveDate { get; set; }
        public string RIC { get; set; }
        public string FM { get; set; }
        public string IDNDisplayName { get; set; }
        public string Ticker { get; set; }
        public string BCAST_REF { get; set; }
        public string QACommonName { get; set; }
        public string MatDate { get; set; }
        public string StrikePrice { get; set; }
        public string QuanityOfWarrants { get; set; }
        public string IssuePrice { get; set; }
        public string IssueDate { get; set; }
        public string ConversionRatio { get; set; }
        public string Issuer { get; set; }
        public string KoreaWarrantName { get; set; }
        public string Chain { get; set; }
        public string LastTradingDate { get; set; }
        public string KnockOutPrice { get; set; }
        public string IssuingAuthority { get; set; }  //isu_nm(the second query post data value)//
        public string ISIN { get; set; }                    //2              B //std_cd(the second query post data value) ==fm.webStandardCongenial web:B
        public string XSovereignIssuingAuthority { get; set; }//24        X     
        public string YStockIndexTypes { get; set; }//25                  Y 
        public string ReleaseForm { get; set; }//judge invalid row
    }
}
