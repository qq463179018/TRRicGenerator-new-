using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Ric.Util;
using System.Collections;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.IO;
using HtmlAgilityPack;
using Ric.Core;

namespace Ric.Tasks
{
    public class ELWStrikeConversion : GeneratorBase
    {
        const string LOGFILE_NAME = "Adjustment-Log.txt";
        private static readonly String CONFIGFILE_NAME = ".\\Config\\Korea\\KOREA_ELWStrikeConversionGenerator.Config";
        private List<SPCRAdjustmentTemplate> spcrList = new List<SPCRAdjustmentTemplate>();
        private Hashtable KoreaIssuerNameHash = null;
        private Hashtable koreaUnderlyingHash = null;
        KOREA_ELWStrikeConversionGeneratorConfig configObj = null;
        Logger Logger = null;

        protected override void Start()
        {
            StartELWStrikeConversionJob();
        }

        protected override void Initialize()
        {
            base.Initialize();
            configObj = ConfigUtil.ReadConfig(CONFIGFILE_NAME, typeof(KOREA_ELWStrikeConversionGeneratorConfig)) as KOREA_ELWStrikeConversionGeneratorConfig;
            Logger = new Logger(configObj.LOG_FILE_PATH + LOGFILE_NAME, Logger.LogMode.New);
        }

        public void StartELWStrikeConversionJob()
        {
            LoadIssuerAndUnderlyingData();
            GrabDataFromWebpage();
            FormatAdjustmentData();
            GenerateAdjustmentFile_xls();
        }

        private void GrabDataFromWebpage()
        {
            String startDate = configObj.Korea_SPCRAdjustment_StartDate.Trim().ToString();
            String endDate = configObj.Korea_SPCRAdjustment_EndDate.Trim().ToString();
            //method=searchDisclosureByStockTypeSub&currentPageSize=15&pageIndex=1&menuIndex=3&orderIndex=1&forward=disclosurebystocktype_sub&elwIsuCd=&elwUly=&lpMbr=&corpNameList=&marketType=&fromData=2011-06-01&toData=2011-12-01&reportNm=&elwRsnClss=09003
            //method=searchDisclosureByStockTypeSub&currentPageSize=100&pageIndex=1&menuIndex=3&orderIndex=1&forward=disclosurebystocktype_sub&elwIsuCd=&elwUly=&lpMbr=&corpNameList=&marketType=&fromData=2011-06-01&toData=2011-12-01&reportNm=&elwRsnClss=09003
            //method=searchDisclosureByStockTypeSub&currentPageSize=100&pageIndex=1&menuIndex=3&orderIndex=1&forward=disclosurebystocktype_sub&elwIsuCd=&elwUly=&lpMbr=&corpNameList=&marketType=&fromData=2011-06-02&toData=2011-12-02&reportNm=&elwRsnClss=09003            
            String postData = String.Format("method=searchDisclosureByStockTypeSub&currentPageSize=300&pageIndex=1&menuIndex=3&orderIndex=1&forward=disclosurebystocktype_sub&elwIsuCd=&elwUly=&lpMbr=&corpNameList=&marketType=&fromData={0}&toData={1}&reportNm=&elwRsnClss=09003", startDate, endDate);
            String uri = "http://kind.krx.co.kr/disclosure/disclosurebystocktype.do";
            AdvancedWebClient wc = new AdvancedWebClient();
            String pageSource = String.Empty;
            pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, postData);
            //pageSource = WebClientUtil.GetPageSource(wc, uri, 300000, postData);
            HtmlDocument htc = new HtmlDocument();
            if (!String.IsNullOrEmpty(pageSource))
                htc.LoadHtml(pageSource);
            if (htc != null)
            {
                HtmlNodeCollection nodeCollections = htc.DocumentNode.SelectNodes(".//div[@id='menu1']/table/tbody/tr");
                int count = nodeCollections.Count;
                if (count > 0)
                {
                    for (var i = 0; i < count; i++)
                    {
                        HtmlNode node = nodeCollections[i].SelectSingleNode(".//td[4]/a");
                        String attribute = String.Empty;
                        if (node != null)
                            attribute = node.Attributes["onclick"].Value.Trim().ToString();
                        if (!String.IsNullOrEmpty(attribute))
                        {
                            attribute = attribute.Split('(')[1].Split(',')[0].Trim(new Char[] { ' ', '\'' }).ToString();
                            //http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno=20111201000283&docno=&viewerhost=&viewerport=
                            uri = String.Format("http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno={0}&docno=&viewerhost=&viewerport=", attribute/*"20111201000283"*/);
                            HtmlDocument doc = new HtmlDocument();
                            pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, null);
                            if (!String.IsNullOrEmpty(pageSource))
                                doc.LoadHtml(pageSource);
                            String parameter = String.Empty;
                            if (doc != null)
                                //doc.DocumentNode.SelectSingleNode(".//div[@id='pWrapper']/div[@id='pContArea']/form[@name='frm']/select[@id='mainDocId']/option[2]")
                                parameter = doc.DocumentNode.SelectSingleNode(".//select[@id='mainDocId']/option[2]").Attributes["value"].Value.Trim().ToString();

                            attribute = attribute.Insert(4, "/").Insert(7, "/").Insert(10, "/");
                            //http://kind.krx.co.kr/external/2011/12/01/000283/20111201000641/99858.htm
                            uri = String.Format("http://kind.krx.co.kr/external/{0}/{1}/99858.htm", attribute, parameter);
                            doc = null;
                            doc = WebClientUtil.GetHtmlDocument(uri, 300000, null);
                            if (doc != null)
                            {
                                String judgeStr = doc.DocumentNode.SelectSingleNode("//table[@id='XFormD1_Form0_RepeatTable0']/tbody/tr[1]/td[2]/span").InnerText;
                                if (!String.IsNullOrEmpty(judgeStr))
                                {
                                    if (!judgeStr.Equals("확정"))
                                        continue;
                                    else
                                    {
                                        HtmlNodeCollection trNodes = doc.DocumentNode.SelectNodes("//tr");
                                        int amount = trNodes.Count;
                                        HtmlNode eNode = trNodes[(amount - 1)];
                                        String edate = String.Empty;
                                        if (eNode != null)
                                        {
                                            edate = eNode.SelectSingleNode(".//td[2]").InnerText.Trim().ToString().Replace(" ", "");
                                            int pos = edate.IndexOf("부터");
                                            edate = edate.Substring((pos - 11), 11);
                                            Char[] arr = edate.ToCharArray();
                                            if (arr[0] < 47 && arr[0] > 58)
                                            {
                                                edate = edate.Substring(1);
                                                arr = edate.ToCharArray();
                                                if (arr[0] < 47 && arr[0] > 58)
                                                    edate = edate.Substring(1);
                                            }

                                            edate = Convert.ToDateTime(edate).ToString("yyyy-MMM-dd", new CultureInfo("en-US")).ToUpper();
                                        }

                                        for (var x = 6; x < (amount - 3); x++)
                                        {
                                            SPCRAdjustmentTemplate spcr = new SPCRAdjustmentTemplate();
                                            HtmlNode itemNode = trNodes[x];
                                            if (itemNode != null)
                                            {
                                                String isin = itemNode.SelectSingleNode(".//td[1]").InnerText.Trim().ToString();
                                                String kname = itemNode.SelectSingleNode(".//td[2]").InnerText.Trim().ToString();
                                                String sprice = itemNode.SelectSingleNode(".//td[7]").InnerText.Trim().ToString();
                                                sprice = sprice.Contains(",") ? sprice.Replace(",", "") : sprice;
                                                String cratio = itemNode.SelectSingleNode(".//td[9]").InnerText.Trim().ToString();
                                                spcr.EffectiveDate = edate;
                                                spcr.ISIN = isin;
                                                spcr.KoreaName = kname;
                                                spcr.StrikePrice = sprice;
                                                spcr.ConversionRatio = cratio;
                                                spcrList.Add(spcr);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void LoadIssuerAndUnderlyingData()
        {
            try
            {
                KoreaIssuerNameHash = new Hashtable();
                koreaUnderlyingHash = new Hashtable();
                string codeMapPath = ".\\Config\\Korea\\KOREA_ELWFMCodeMap.xml";
                ClassCodeMap codeMapObj = ConfigUtil.ReadConfig(codeMapPath, typeof(ClassCodeMap)) as ClassCodeMap;


                List<IssuerCodeConversion> iList = codeMapObj.IssuerCodeMap.ToList();
                List<UnderlyingCodeConversion> uList = codeMapObj.UnderlyingCodeMap.ToList();

                foreach (var item in iList)
                    if (!KoreaIssuerNameHash.Contains(item.DSPLY_NMLL))
                        KoreaIssuerNameHash.Add(item.DSPLY_NMLL, item);

                foreach (var item in uList)
                    if (!koreaUnderlyingHash.Contains(item.DSPLY_NMLL))
                        koreaUnderlyingHash.Add(item.DSPLY_NMLL, item);
            }
            catch (Exception ex)
            {
                String msg = "Error found in LoadIssuerAndUnderlyingData()   : \r\n"+ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }
        }

        private void FormatAdjustmentData()
        {
            try
            {
                if (spcrList.Count > 0)
                {
                    for (var i = 0; i < spcrList.Count; i++)
                    {
                        var spcr = spcrList[i] as SPCRAdjustmentTemplate;
                        Char[] temp_array = spcr.KoreaName.ToCharArray();
                        String str_issuername = String.Empty;
                        foreach (var item in temp_array)
                        {
                            int asciiCode = (int)item;
                            if (asciiCode > 47 && asciiCode < 58)
                                break;
                            str_issuername += item.ToString();
                        }

                        int no = 0;
                        for (var j = 0; j < temp_array.Length; j++)
                        {
                            int asciiCode = (int)temp_array[j];
                            if (asciiCode > 47 && asciiCode < 58)
                            {
                                no = j;
                                break;
                            }
                        }

                        String CharLen = spcr.KoreaName.Substring(no, 4).Trim().ToString();
                        String issuer_ShortName = String.Empty;
                        if (KoreaIssuerNameHash.Contains(str_issuername))
                        {
                            spcr.RIC = ((IssuerCodeConversion)KoreaIssuerNameHash[str_issuername]).no + CharLen + ".KS";
                            issuer_ShortName = ((IssuerCodeConversion)KoreaIssuerNameHash[str_issuername]).shortname;
                        }
                        else
                            issuer_ShortName = "***";

                        String matDate = GetMatDate(spcr.ISIN);
                        matDate = Convert.ToDateTime(matDate).ToString("MMM-yy", new CultureInfo("en-US")).Replace("-", "").ToUpper();

                        String str_underlying = spcr.KoreaName.Substring((no + 4)).Trim().ToString();
                        String dsplyNmll = str_underlying.Substring(0, (str_underlying.Length - 1)).Trim();

                        if (dsplyNmll == "KOSPI200")
                            dsplyNmll = "코스피";
                        if (dsplyNmll == "스탠차")
                            dsplyNmll = "스탠다드차타드";
                        if (dsplyNmll == "IBK")
                            dsplyNmll = "아이비케이";
                        if (dsplyNmll == "HMC")
                            dsplyNmll = "에이치엠씨";
                        if (dsplyNmll == "KB")
                            dsplyNmll = "케이비";

                        String porc = str_underlying.Substring(str_underlying.Length - 1).Trim().ToString();
                        porc = porc.Equals("콜") ? "C" : (porc.Equals("풋") ? "P" : "*");

                        String underlying_QACommonName = String.Empty;
                        if (koreaUnderlyingHash.Contains(dsplyNmll))
                            underlying_QACommonName = ((UnderlyingCodeConversion)koreaUnderlyingHash[dsplyNmll]).QACommonName;
                        else
                            underlying_QACommonName = "***";

                        String strikePirce = spcr.StrikePrice.Length > 4 ? spcr.StrikePrice.Substring(0, 4).Trim().ToString() : spcr.StrikePrice;
                        String type = dsplyNmll.Equals("KSOPI200") ? "IW" : "WNT";
                        String QACommonName = underlying_QACommonName + " " + issuer_ShortName + " " + matDate + " " + strikePirce + " " + porc + type;
                        spcr.QACommonName = QACommonName.Trim().ToString().ToUpper();
                    }
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in FormatAdjustmentData()     : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        private String GetMatDate(String isin)
        {
            String matdate = String.Empty;
            try
            {
                //                          http://isin.krx.co.kr/jsp/BA_VW021.jsp?isu_cd=KRA631193158&modi=f&req_no=201105240094
                String uri = String.Format("http://isin.krx.co.kr/jsp/BA_VW021.jsp?isu_cd={0}&modi=f&req_no=", isin);
                String pageSource = WebClientUtil.GetPageSource(uri, 300000);
                if (!String.IsNullOrEmpty(pageSource))
                {
                    HtmlDocument htc = new HtmlDocument();
                    htc.LoadHtml(pageSource);
                    if (htc != null)
                    {
                        HtmlNode node = htc.DocumentNode.SelectSingleNode(".//tr[6]/td[4]");
                        if (node != null)
                            matdate = node.InnerText.Trim().ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                String msg = "Error found in GetMatDate()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return matdate;
        }

        private void GenerateAdjustmentFile_xls()
        {
            if (spcrList.Count > 0)
            {
                ExcelApp excelApp = new ExcelApp(false, false);
                if (excelApp.ExcelAppInstance == null) { }
                try
                {
                    String filename = String.Format("KR FM (ELW Adjustment) Strike Price and Conversion Ratio (wef {0}).xls", spcrList[0].EffectiveDate);
                    String ipath = "C:\\Korea_Auto\\ELW_FM\\ELW_Adjustment\\" + filename;
                    Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                    Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet1", wBook);
                    if (wSheet == null) { }

                    if (wSheet.get_Range("C1", Type.Missing).Value2 == null)
                    {
                        ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 15;
                        ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                        ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
                        ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 15;
                        ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 20;
                        ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 20;
                        ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 35;
                        ((Range)wSheet.Columns["A:G", System.Type.Missing]).Font.Name = "Arial";
                        ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                        ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                        ((Range)wSheet.Rows[1, Type.Missing]).WrapText = true;
                        wSheet.Cells[1, 1] = "Update Date";
                        wSheet.Cells[1, 2] = "Effective Date";
                        wSheet.Cells[1, 3] = "RIC";
                        wSheet.Cells[1, 4] = "ISIN";
                        wSheet.Cells[1, 5] = "Strike Price(After Adjustment)";
                        wSheet.Cells[1, 6] = "Conversion Ratio(After Adjustment)";
                        wSheet.Cells[1, 7] = "QA Common Name";
                    }

                    int startLine = 2;

                    foreach (var item in spcrList)
                    {
                        wSheet.Cells[startLine, 1] = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                        wSheet.Cells[startLine, 2] = item.EffectiveDate;
                        wSheet.Cells[startLine, 3] = item.RIC;
                        wSheet.Cells[startLine, 4] = item.ISIN;
                        wSheet.Cells[startLine, 5] = item.StrikePrice;
                        wSheet.Cells[startLine, 6] = item.ConversionRatio;
                        wSheet.Cells[startLine, 7] = item.QACommonName;
                        startLine++;
                    }

                    excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                    wBook.Save();
                }
                catch (Exception ex)
                {
                    String msg = "Error found in GenerateAdjustmentFile_xls()   : \r\n" + ex.ToString();
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }
                finally
                {
                    excelApp.Dispose();
                }
            }
        }
                
    }
}
