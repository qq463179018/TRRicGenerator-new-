using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.Threading;
using HtmlAgilityPack;
using ICSharpCode.SharpZipLib.Zip;
using System.Windows.Forms;
using pdftron;
using PdfTronWrapper;
using Ric.Db.Manager;
using Ric.Db.Info;
//using ETI.Core;
using System.Data;
using Ric.Util;
using Ric.Core;

namespace Ric.Tasks.Korea
{
    public class ELWFMSecondPart : GeneratorBase
    {
        #region Data models

        private class ChangeddataModel
        {
            public string Secondary_ID { set; get; }
            public string Issue_Price { set; get; }
            public string Warrants_Per_Underlying { set; get; }
            public string Underlying_RIC { set; get; }
            public bool Secondary_ID_Changed { set; get; }
            public string QACommonName { set; get; }
            public string IDNDisplayName { set; get; }
            public string StrikePrice { set; get; }
            public string MatDate { set; get; }
            public string RIC { set; get; }
            public string CPOption { set; get; }
            public string Warrant_Title { set; get; }
            public string Warrant_Type { set; get; }
            public ChangeddataModel()
            {
                this.Issue_Price = string.Empty;
                this.Secondary_ID = string.Empty;
                this.Underlying_RIC = string.Empty;
                this.Warrants_Per_Underlying = string.Empty;
                this.QACommonName = string.Empty;
                this.IDNDisplayName = string.Empty;
                this.StrikePrice = string.Empty;
                this.MatDate = string.Empty;
                this.RIC = string.Empty;
                this.CPOption = string.Empty;
                this.Warrant_Title = string.Empty;
                this.Warrant_Type = string.Empty;
                this.Secondary_ID_Changed = false;
            }
            public override string ToString()
            {
                string code = "";
                if (string.IsNullOrEmpty(QACommonName))
                {
                    code += "0";
                }
                else
                {
                    code += "1";
                }
                if (string.IsNullOrEmpty(IDNDisplayName))
                {
                    code += "0";
                }
                else
                {
                    code += "1";
                }
                if (string.IsNullOrEmpty(MatDate))
                {
                    code += "0";
                }
                else
                {
                    code += "1";
                }
                if (string.IsNullOrEmpty(StrikePrice))
                {
                    code += "0";
                }
                else
                {
                    code += "1";
                }
                return code;
            }
        }

        private class ReferenceListUnderLyingModel
        {
            public string UnderlyingRIC { set; get; }
            public string UnderlyingCompanyName { set; get; }
            public string NDATCUnderlyingTitle { set; get; }
        }

        private class ReferenceListIssueModel
        {
            public string RIC { set; get; }
            public string IssuerCompanyName { set; get; }
            public string NDAIssuerORGID { set; get; }
            public string NDATCIssuerTitle { set; get; }
        }

        private class EmaQuantityChange
        {
            public string ISAN { set; get; }
            public string QuanityofWarrants { set; get; }
        }

        private class EmaPriceChange
        {
            public string ISIN { get; set; }
            public string ExcercisePrice { get; set; }
            public string WarrantsPerUnderlying { get; set; }

            public EmaPriceChange()
            {
                ExcercisePrice = string.Empty;
                WarrantsPerUnderlying = string.Empty;
            }
        }

        #endregion

        #region Fields
        private List<WarrantTemplate> fmSecondList = new List<WarrantTemplate>();

        private List<WarrantTemplate> kobaList = new List<WarrantTemplate>();

        private Hashtable fmSecondHash = new Hashtable();

        private Hashtable pilcHash = new Hashtable();

        private Hashtable emaFileHash = null;

        private Hashtable ndaFileHash = null;

        private Hashtable gedaFileHash = null;

        private Hashtable priceChange = new Hashtable();

        private List<WarrantTemplate> noFm1List = new List<WarrantTemplate>();

        private List<EmaQuantityChange> quanityofWarrantsList = null;

        Hashtable refernceList = new Hashtable();

        private bool fm2HasChangedValue = false;

        private KOREA_ELWFM2AndFurtherIssuerGeneratorConfig configObj = null;

        //These two tags will be marked true to judge whether FM2 or FutherIssue exist
        //The value was set in DownLoadFile()
        //Used in StartJob()
        private bool isFM2 = false;

        private bool isFutherIssue = false;

        private List<KoreaUnderlyingInfo> newUnderLying = new List<KoreaUnderlyingInfo>();

        private string downloadPdf = string.Empty;

        private string downloadPdfFurtherIssuer = string.Empty;

        private string originalBulkFilePath = string.Empty;

        #endregion

        #region Intialize and Start Methods

        protected override void Start()
        {
            DateTime startTime = DateTime.Now;
            StartJob();
            TimeSpan ts = DateTime.Now.Subtract(startTime);
            double count = ts.TotalMinutes;
            Logger.Log("Total Minutes:" + count.ToString(), Logger.LogType.Info);
            configObj.BulkFile = originalBulkFilePath;
            //AddResult("LOG FILE",logger.FilePath,"LOG");
        }

        protected override void Initialize()
        {
            //pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
            configObj = Config as KOREA_ELWFM2AndFurtherIssuerGeneratorConfig;
        }

        private void InitializeConfig()
        {
            if (!string.IsNullOrEmpty(configObj.BulkFile))
            {
                originalBulkFilePath = configObj.BulkFile;
                configObj.BulkFile = Path.Combine(configObj.BulkFile, DateTime.Today.ToString("yyyy-MM-dd"));
            }
            else
            {
                configObj.BulkFile = GetOutputFilePath() + "\\BulkFile\\";
            }
            if (!Directory.Exists(configObj.BulkFile))
            {
                Directory.CreateDirectory(configObj.BulkFile);
            }
            if (string.IsNullOrEmpty(configObj.FM))
            {
                configObj.FM = GetOutputFilePath() + "\\FM\\";
            }
            downloadPdf = Path.GetDirectoryName(configObj.FM) + "\\PDF\\";

            if (string.IsNullOrEmpty(configObj.FM_FurtherIssuer))
            {
                configObj.FM_FurtherIssuer = GetOutputFilePath() + "\\FM_FurtherIssuer\\";
            }
            downloadPdfFurtherIssuer = Path.GetDirectoryName(configObj.FM_FurtherIssuer) + "\\PDF\\";
        }

        public void StartJob()
        {
            InitializeConfig();
            try
            {
                if (configObj.AnnouncementType.Equals(ElwAnnounceType.None))
                {
                    StartDownloadPDFFileJob();
                }
                else if (configObj.AnnouncementType.Equals(ElwAnnounceType.FM2_ELW))
                {
                    isFM2 = true;
                    downloadPdf = configObj.PdfPath;
                }
                else if (configObj.AnnouncementType.Equals(ElwAnnounceType.FurtherIssuer))
                {
                    isFutherIssue = true;
                    downloadPdfFurtherIssuer = configObj.PdfPath;
                }

                if (isFutherIssue && isFM2)
                {
                    StartELWFurtherIssuerJob();
                    StartELWFMSecondPartJob();
                }
                else
                {
                    if (isFutherIssue)
                        StartELWFurtherIssuerJob();
                    if (isFM2)
                        StartELWFMSecondPartJob();
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
                throw (ex);
            }
        }

        public void StartDownloadPDFFileJob()
        {
            DeleteUselessFileAndFolder();
            DownLoadFile();
        }

        public void StartELWFMSecondPartJob()
        {
            try
            {
                GrabAllDataFromPdfFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            GetReferenceFromDB();

            if (kobaList.Count > 0)
            {
                AddResult("KOBA FM Folder", configObj.FM, "KOBA FM Folder");

                GrabKOBADataFromISINWebsit();
                KOBAListDataFormat();
                GenerateKOBAFmFile();
                AppendDataToKOBADb();
                GenerateKobaEmaGedaNdaAndCreatEmail();
            }

            if (fmSecondList.Count > 0)
            {
                AddResult("ELW FM2 Folder", configObj.FM, "ELW FM2 Folder");
                AddResult("BulkFile Folder", configObj.BulkFile, "BulkFile Folder");

                pilcHash = ReadReferenceTableOfPlic();
                //GetTagAndPilcFromDb();
                FM2ListDataFormat();
                string rics = GetELWFmOneFromDb();
                quanityofWarrantsList = new List<EmaQuantityChange>();
                emaFileHash = new Hashtable();
                ndaFileHash = new Hashtable();
                gedaFileHash = new Hashtable();
                //Compare Master file and generate FMSecond file
                GenerateFMSecondPartFile();
                GenerateEmaGedaNdaAndCreatEmail();
                AppendDataToELWDb(rics);
            }
            GenerateNdaTickLotFile();
            if (newUnderLying.Count > 0)
            {
                GenerateNewUnderlyingFiles();
            }
        }

        public void StartELWFurtherIssuerJob()
        {
            FurtherIssue fissuer = new FurtherIssue();
            fissuer.StartFurtherIssuerJob(configObj, TaskResultList, downloadPdfFurtherIssuer);
        }

        #endregion

        #region  Download PDF File Job

        private void DownLoadFile()
        {
            try
            {

                string uri = string.Format("http://www.krx.co.kr/m6/m6_2/m6_2_2/JHPKOR06002_02_01.jsp");
                string pageSource = string.Empty;
                pageSource = WebClientUtil.GetPageSource(null, uri, 30000, null, Encoding.UTF8);
                HtmlAgilityPack.HtmlDocument htc = new HtmlAgilityPack.HtmlDocument();
                if (!string.IsNullOrEmpty(pageSource))
                    htc.LoadHtml(pageSource);
                if (htc != null)
                {
                    HtmlNodeCollection trNodes = htc.DocumentNode.SelectNodes("//table")[1].SelectNodes(".//tr[@class='row1']");

                    int count = trNodes.Count;
                    string startDate = configObj.StartDate;
                    for (var i = 0; i < count; i++)
                    {
                        var item = trNodes[i] as HtmlNode;
                        string date = item.SelectSingleNode(".//td[3]").InnerText.Trim().ToString();
                        date = Convert.ToDateTime(date).ToString("yyyy-MM-dd");
                        if (date.Equals(startDate))
                        {
                            string title = item.SelectSingleNode(".//td[1]").InnerText;
                            //this is for ELW FM2 <Change> pdf file download
                            if (title.Contains("주식워런트증권 신규상장") || (title.Contains("ELW") && title.Contains("신규상장"))) // && title.Contains("상장일")// 
                            {
                                string type = "FM2";
                                SavaDownloadFile(i, item, type, null);
                                isFM2 = true;
                            }   //this is for Further Issuer pdf file download
                            else if (title.Contains("주식워런트증권 추가상장") && title.Contains("상장일"))
                            {
                                string type = "FurtherIssuer";
                                SavaDownloadFile(i, item, type, null);
                                isFutherIssue = true;
                            }
                            else if (title.Contains("상장일"))
                            {
                                SavaDownloadFile(i, item, null, null);
                            }
                            else
                            {
                                if (i == (count - 1))
                                {
                                    if (date.Equals(startDate))
                                        DownLoadFile(uri, 2);
                                    else break;
                                }
                                else continue;
                            }

                            if (i == (count - 1))
                            {
                                if (date.Equals(startDate))
                                    DownLoadFile(uri, 2);
                                else break;
                            }
                            else continue;
                        }
                        else
                        {
                            double seconds = Convert.ToDateTime(date).Subtract(Convert.ToDateTime(startDate)).TotalSeconds;
                            if (seconds > 0)
                            {
                                if (i == (count - 1))
                                {
                                    if (seconds >= 0)
                                        DownLoadFile(uri, 2);
                                    else break;
                                }
                                else continue;
                            }
                            else
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in DownLoadFDFFile()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        private void DownLoadFile(string uri, int pageIndex)
        {
            string postData = string.Format("cur_page={0}&ntc_cd=&ntc_seq_cd=&search_type=&search_content=&search_word=", pageIndex.ToString());
            try
            {
                string pageSource = string.Empty;
                pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, postData);
                HtmlAgilityPack.HtmlDocument htc = new HtmlAgilityPack.HtmlDocument();
                if (!string.IsNullOrEmpty(pageSource))
                    htc.LoadHtml(pageSource);
                if (htc != null)
                {
                    HtmlNodeCollection trNodes = htc.DocumentNode.SelectNodes("//table")[1].SelectNodes(".//tr[@class='row1']");
                    int count = trNodes.Count;
                    for (var i = 0; i < count; i++)
                    {
                        var item = trNodes[i] as HtmlNode;
                        string date = item.SelectSingleNode(".//td[3]").InnerText;
                        date = Convert.ToDateTime(date).ToString("yyyy-MM-dd");
                        string startDate = configObj.StartDate;
                        if (date.Equals(startDate))
                        {
                            string title = item.SelectSingleNode(".//td[1]").InnerText;
                            if (title.Contains("주식워런트증권 신규상장") && title.Contains("상장일"))
                            {
                                string type = "FM2";
                                SavaDownloadFile(i, item, type, pageIndex.ToString());
                            }
                            else if (title.Contains("주식워런트증권 추가상장") && title.Contains("상장일"))
                            {
                                string type = "FurtherIssuer";
                                SavaDownloadFile(i, item, type, pageIndex.ToString());
                            }
                            else
                            {
                                if (i == (count - 1))
                                {
                                    if (date.Equals(startDate))
                                        DownLoadFile(uri, pageIndex++);
                                    else break;
                                }
                                else continue;
                            }
                        }
                        else
                        {
                            double seconds = Convert.ToDateTime(date).Subtract(Convert.ToDateTime(startDate)).TotalSeconds;
                            if (seconds > 0)
                            {
                                if (i == (count - 1))
                                {
                                    if (date.Equals(startDate))
                                        DownLoadFile(uri, pageIndex++);
                                    else break;
                                }
                                else continue;
                            }
                            else
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in DownLoadFDFFile()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        private void SavaDownloadFile(int i, HtmlNode item, string type, string index)
        {
            try
            {
                string attribute = item.SelectSingleNode(".//td[1]/a").Attributes["onclick"].Value.ToString().Trim();
                string param1 = attribute.Split(',')[0].Split('(')[1].Trim(new char[] { ' ', ',', '(', '\'' }).ToString();
                string param2 = attribute.Split(',')[1].Split(')')[0].Trim(new char[] { ' ', ',', '(', '\'' }).ToString();
                string url = string.Format("http://www.krx.co.kr/por_kor/m6/m6_2/m6_2_2/JHPKOR06002_02_02.jsp?ntc_seq_cd={0}&ntc_cd={1}", param1, param2);
                //                          http://www.krx.co.kr/por_kor/m6/m6_2/m6_2_2/JHPKOR06002_02_02.jsp?ntc_seq_cd={0}&ntc_cd={1}
                string source = string.Empty;
                source = WebClientUtil.GetPageSource(url, 300000, null);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                if (!string.IsNullOrEmpty(source))
                    doc.LoadHtml(source);
                if (doc != null)
                {
                    if (string.IsNullOrEmpty(type))
                    {
                        string innertitle = doc.DocumentNode.SelectSingleNode(".//table/tr[4]/td").InnerText.Trim();
                        if (!string.IsNullOrEmpty(innertitle))
                        {
                            if (innertitle.Contains("주식워런트증권") && innertitle.Contains("신규상장"))
                            {
                                type = "FM2";
                            }
                            else if (innertitle.Contains("주식워런트증권") && innertitle.Contains("추가상장"))
                            {
                                type = "FurtherIssuer";
                            }
                            else return;
                        }
                    }

                    string filetype = doc.DocumentNode.SelectSingleNode(".//table/tr[5]/td/a").InnerText.Trim();
                    filetype = filetype.Substring(filetype.Length - 3);
                    string site = doc.DocumentNode.SelectSingleNode(".//table/tr[5]/td/a").Attributes["href"].Value;

                    //http://www.krx.co.kr/por_kor/servlets/FileDownload?type=intro_file0010&noti_no=7874&seq=9557
                    site = site.Trim().ToString().Replace("amp;", "");
                    url = "http://www.krx.co.kr" + site;
                    string filename = string.Empty;
                    if (!string.IsNullOrEmpty(index))
                        filename = index + i.ToString() + "." + filetype;
                    else
                        filename = i.ToString() + "." + filetype;
                    string ipath = string.Empty;
                    if (type.Equals("FM2"))
                        ipath = downloadPdf + filename;                //"C:\\Korea_Auto\\ELW_FM\\ELW_FM2\\PDF\\" + filename;
                    else
                        ipath = downloadPdfFurtherIssuer + filename;      //"C:\\Korea_Auto\\ELW_FM\\ELW_FM_Further_Issuer\\PDF\\" + filename;
                    WebClientUtil.DownloadFile(url, 300000, ipath);
                    if (filetype.Equals("zip"))
                    {
                        try
                        {
                            //ipath = ipath.Replace(".zip", "");
                            //String[] pdfarr = Directory.GetFiles(ipath);
                            //int count = pdfarr.Length;

                            FileInfo fi = new FileInfo(ipath);
                            using (ZipInputStream stream = new ZipInputStream(fi.OpenRead()))
                            {
                                string foldername = string.Empty;
                                if (string.IsNullOrEmpty(foldername))
                                    foldername = fi.FullName.Replace(fi.Extension, string.Empty);
                                //Directory.CreateDirectory(foldername);
                                foldername = downloadPdf;
                                ZipEntry ze = null;
                                while ((ze = stream.GetNextEntry()) != null)
                                {
                                    int size = 204800000;
                                    byte[] data = new byte[size];
                                    String[] s = ze.Name.Split('\\');
                                    if (s.Length > 1)
                                    {
                                        StringBuilder sb = new StringBuilder(foldername);

                                        int x = 0;
                                        while (x < s.Length - 1)
                                        {
                                            sb.Append('\\');
                                            sb.Append(s[x++]);
                                        }
                                        Directory.CreateDirectory(sb.ToString());
                                    }

                                    string outfile = foldername + ze.Name;


                                    using (FileStream fs = new FileStream(outfile, FileMode.Create))
                                    {
                                        while (true)
                                        {
                                            size = stream.Read(data, 0, data.Length);

                                            if (size > 0)
                                            {
                                                fs.Write(data, 0, size);
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        fs.Flush();
                                        fs.Close();
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            string msg = "Get pdf file from zip folder cause error , please check the zip folder ,make sure there contains file in folder.   : \r\n" + ex.ToString();
                            Logger.Log(msg, Logger.LogType.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in SaveDownloadFile()    :\r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void DeleteUselessFileAndFolder()
        {
            string msg = string.Empty;
            try
            {
                string FMPath = downloadPdf;
                if (!string.IsNullOrEmpty(FMPath))
                {
                    //FMPath = FMPath.Substring(0, FMPath.Length - 4).Trim().ToString();
                    Common.DeleteUselessFileAndFolder(FMPath);
                }
                string FIPath = downloadPdfFurtherIssuer;
                if (!string.IsNullOrEmpty(FIPath))
                {
                    //FIPath = FIPath.Substring(0, FIPath.Length - 4).Trim().ToString();
                    Common.DeleteUselessFileAndFolder(FIPath);
                }
            }
            catch (Exception ex)
            {
                msg = "\r\n    Error found in DeleteUselessFileAndFolder()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        //Parse all pdf files, store data into FM2List
        private void GrabAllDataFromPdfFile()
        {
            ArrayList fileList = new ArrayList();
            GetAllPDFFilesName(fileList);
            List<FreeTable> table = null;
            //used to store all source data from pdf
            fmSecondList = new List<WarrantTemplate>();

            foreach (var item in fileList)
            {
                table = GrabTableFromPdf(item.ToString());
                //  Thread.Sleep(5000);
                GrabDataFromTable(table);
            }
            int count = fmSecondList.Count;
        }

        private void GetAllPDFFilesName(ArrayList list)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("ko-KR");
            string fpath = downloadPdf; //"D:\\Korea_Auto\\ELW_FM\\ELW_FM2\\PDF\\";
            DirectoryInfo dir = new DirectoryInfo(fpath);
            if (dir.Exists == true && dir.GetFileSystemInfos() != null && dir.GetFileSystemInfos().Length != 0)
            {

                foreach (FileSystemInfo fsi in dir.GetFileSystemInfos())
                {
                    string filename = string.Empty;
                    if (fsi is FileInfo)
                    {
                        filename = fsi.Name;
                        if (filename.Contains(".pdf"))
                            list.Add(filename);

                    }
                    else
                    {
                        filename = fsi.FullName;
                    }
                }
            }
        }

        private void GrabDataFromTable(List<FreeTable> tables)
        {
            try
            {

                foreach (var table in tables)
                {
                    WarrantTemplate koreaTemp1 = new WarrantTemplate();
                    WarrantTemplate koreaTemp2 = new WarrantTemplate();
                    GetIssuer(koreaTemp1, koreaTemp2, table);                 //Issuer  , to judge wheather the issuer is KOBA
                    GetKoreaName(koreaTemp1, koreaTemp2, table);                //Exchange_Warrant_Name
                    GetISIN(koreaTemp1, koreaTemp2, table);                     //ISIN
                    GetTicker(koreaTemp1, koreaTemp2, table);                   //Ticker
                    GetQuantity(koreaTemp1, koreaTemp2, table);                 //Quantity_of_Warrant
                    GetIssuerPrice(koreaTemp1, koreaTemp2, table);              //Issuer_Price
                    GetConvationRatio(koreaTemp1, koreaTemp2, table);           //Convation_Ratio
                    GetStrikePrice(koreaTemp1, koreaTemp2, table);              //Strike_Price
                    GetMatureDate(koreaTemp1, koreaTemp2, table);               //Mature_date
                    GetLastTradingDate(koreaTemp1, koreaTemp2, table);          //Last_Trading_date
                    GetEffectiveDate(koreaTemp1, koreaTemp2, table);            //Effective_date
                    GetUnderlyingKoreaName(koreaTemp1, koreaTemp2, table);      //UnderlyingKoreaName
                    koreaTemp1.IDNDisplayName = "";
                    if (koreaTemp1.Ticker != null && koreaTemp1.Ticker != string.Empty)
                    {
                        if (koreaTemp1.IsKOBA)
                        {
                            GetKnockOutPrice(koreaTemp1, koreaTemp2, table);            //if issuer is KOBA , Knock_Out_Price              
                            kobaList.Add(koreaTemp1);
                        }
                        else
                            fmSecondList.Add(koreaTemp1);
                    }

                    if (koreaTemp2.Ticker != null && koreaTemp2.Ticker != string.Empty)
                    {
                        if (koreaTemp2.IsKOBA)
                        {
                            GetKnockOutPrice(koreaTemp1, koreaTemp2, table);            //if issuer is KOBA , Knock_Out_Price              
                            kobaList.Add(koreaTemp2);
                        }
                        else
                            fmSecondList.Add(koreaTemp2);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabDataFromText()       : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        #endregion

        #region PDF to TEXT
        private static void WriteFreeTableToTxt(string txtPath, string spaceType, List<FreeTable> tableList)
        {
            FileStream fs = new FileStream(txtPath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            try
            {
                foreach (var table in tableList)
                {
                    foreach (var row in table)
                    {
                        sw.Write(spaceType);
                        foreach (var cell in row)
                        {
                            sw.Write(cell.Value + spaceType);
                        }
                        sw.Write("\r\n");
                    }
                    sw.Write("******************\r\n");
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                sw.Close();
                fs.Close();
            }
        }
        private List<FreeTable> PDFTronTransfer(string pdfPath)
        {
            pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
            TableLocator locator = new TableLocator(pdfPath);
            LocateConfiguration config = new LocateConfiguration();
            config.TableEndFirstLetterRegex = @"\d";
            config.TableEndRegex = @"\d{1,}/\d{1,}";
            config.TableNameNearbyFirstLetterRegex = "상";
            config.TableNameNearbyRegex = "상장일";
            List<TablePos> tablePosList = locator.GetMultiTablePos(".*?신규상장\\s*내역", config);
            List<FreeTable> tableList = new List<FreeTable>();
            foreach (var tablePos in tablePosList)
            {
                FreeTable table = TableExtractor.Extract(locator.pdfDoc, tablePos);
                tableList.Add(table);

            }
            return tableList;
        }

        //Get tablelist from PDF directly
        private List<FreeTable> GrabTableFromPdf(string name)
        {
            string txtPath = string.Empty;
            List<FreeTable> tableList = null;
            try
            {

                System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("ko-KR");
                string pdfPath = downloadPdf + name;                      //commonObj.Log_Path + "\\" + commonObj.SubFolder + "\\" + pdfFolder + "\\" + name;
                tableList = PDFTronTransfer(pdfPath);
            }
            catch (Exception ex)
            {
                string msg = "Error found in PDFToText     : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
            return tableList;
        }
        #endregion

        #region Get each data from pdf table to WarrantTemplate.

        private void GetKoreaName(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[2][1].Value);
                koreaTemp1.KoreaWarrantName = pre[2][1].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[2][2].Value);
                if (m.Success)
                {
                    koreaTemp2.KoreaWarrantName = pre[2][2].Value;
                }

            }
            catch (Exception ex)
            {
                string msg = "Error found in GetKoreaName()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetISIN(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[7][2].Value);
                koreaTemp1.ISIN = pre[7][2].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[7][3].Value);
                if (m.Success)
                {
                    koreaTemp2.ISIN = pre[7][3].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetISIN()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetTicker(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[8][1].Value);
                koreaTemp1.Ticker = pre[8][1].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[8][2].Value);
                if (m.Success)
                {
                    koreaTemp2.Ticker = pre[8][2].Value;
                }
                else
                {
                    koreaTemp2.Ticker = null;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetTicker()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetQuantity(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[9][2].Value);
                koreaTemp1.QuanityofWarrants = pre[9][2].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[9][3].Value);
                if (m.Success)
                {
                    koreaTemp2.QuanityofWarrants = pre[9][3].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetQuantity()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetIssuerPrice(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[10][1].Value);
                koreaTemp1.IssuePrice = pre[10][1].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[10][2].Value);
                if (m.Success)
                {
                    koreaTemp2.IssuePrice = pre[10][2].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetIssuerPrice()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetConvationRatio(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[13][1].Value);
                koreaTemp1.ConversionRatio = pre[13][1].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[13][2].Value);
                if (m.Success)
                {
                    koreaTemp2.ConversionRatio = pre[13][2].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetConvationRatio()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetStrikePrice(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[16][1].Value);
                koreaTemp1.StrikePrice = pre[16][1].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[16][2].Value);
                if (m.Success)
                {
                    koreaTemp2.StrikePrice = pre[16][2].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetStrikePice()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetMatureDate(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[21][2].Value);
                koreaTemp1.MatDate = pre[21][2].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[21][3].Value);
                if (m.Success)
                {
                    koreaTemp2.MatDate = pre[21][3].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetMatureDate()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetLastTradingDate(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[25][1].Value);
                koreaTemp1.LastTradingDate = pre[25][1].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[25][2].Value);
                if (m.Success)
                {
                    koreaTemp2.LastTradingDate = pre[25][2].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetLastTradingDate()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetEffectiveDate(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[0][1].Value);
                koreaTemp1.EffectiveDate = pre[0][1].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[0][2].Value);
                if (m.Success)
                {
                    koreaTemp2.EffectiveDate = pre[0][2].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetEffectiveDate()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetUnderlyingKoreaName(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<name>[^\(\)]*)");
                Match m = r.Match(pre[12][1].Value);
                koreaTemp1.UnderlyingKoreaName = m.Groups["name"].Value;


                r = new Regex(@"(?<name>[^\(\)]*)");
                m = r.Match(pre[12][2].Value);
                if (m.Success)
                {
                    koreaTemp2.UnderlyingKoreaName = m.Groups["name"].Value;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetUnderlyingKoreaName()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetIssuer(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[3][1].Value);
                koreaTemp1.Issuer = pre[3][1].Value;


                r = new Regex(@"(?<issuer>.{1,})");
                m = r.Match(pre[3][2].Value);
                if (m.Success)
                {
                    koreaTemp2.Issuer = pre[3][2].Value;
                    if (pre[1][3].Value.Contains("조기종료") && pre[2][2].Value.Contains("조기종료") && pre[3][2].Value.Contains("KNOCK-OUT"))
                        koreaTemp2.IsKOBA = true;
                    else
                        koreaTemp2.IsKOBA = false;
                }
                if (pre[1][2].Value.Contains("조기종료") && pre[2][1].Value.Contains("조기종료") && pre[3][1].Value.Contains("KNOCK-OUT"))
                {
                    koreaTemp1.IsKOBA = true;
                }
                else
                {
                    koreaTemp1.IsKOBA = false;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetIssuer()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GetKnockOutPrice(WarrantTemplate koreaTemp1, WarrantTemplate koreaTemp2, FreeTable pre)
        {
            try
            {
                Regex r = new Regex(@"(?<issuer>\d{1,})");
                Match m = r.Match(pre[17][1].Value);
                koreaTemp1.KnockOutPrice = pre[17][1].Value;


                r = new Regex(@"(?<issuer>\d{1,})");
                m = r.Match(pre[17][2].Value);
                if (m.Groups.Count > 0)
                {
                    koreaTemp2.KnockOutPrice = pre[17][2].Value;
                }

            }
            catch (Exception ex)
            {
                string msg = "Error found in GetKnockOutPrice()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private string VariableFormat(string pre, int pos)
        {
            string result = "";
            char[] carr = pre.ToCharArray();
            while (carr[pos] != '\n')
            {
                result += carr[pos];
                if ((pos + 1) >= carr.Length)
                    break;
                pos++;
            }
            result = result.Trim(new char[] { ' ', '\n', '\r' }).ToString();
            return result;
        }

        private string[] GenerateArray(string var)
        {
            int index = 0;
            char[] carr = var.ToCharArray();
            for (var i = 0; i < carr.Length; i++)
            {
                if (carr[i] == ' ' && carr[(i + 1)] == ' ')
                    index = i;
            }
            string temp1 = var.Substring(0, (index - 1)).Trim().ToString();
            string temp2 = var.Substring(index).Trim().ToString();
            String[] result = new string[] { temp1, temp2 };
            return result;
        }

        #endregion

        #region KOBA work logic
        private void GrabKOBADataFromISINWebsit()
        {
            try
            {
                foreach (var item in kobaList)
                {
                    string uri = string.Format("http://isin.krx.co.kr/jsp/BA_VW021.jsp?isu_cd={0}&modi=f&req_no=", item.ISIN);
                    HtmlAgilityPack.HtmlDocument htc = new HtmlAgilityPack.HtmlDocument();
                    string pageSource = WebClientUtil.GetDynamicPageSource(uri, 300000, null);
                    if (!string.IsNullOrEmpty(pageSource))
                        htc.LoadHtml(pageSource);
                    if (htc != null)
                    {
                        HtmlNode table = htc.DocumentNode.SelectNodes("//table")[2];
                        string str_issuer = table.SelectSingleNode(".//tr[4]/td[2]").InnerText.Trim().ToString();
                        string str_issuer_date = table.SelectSingleNode(".//tr[6]/td[2]").InnerText.Trim().ToString();
                        item.IssueDate = Convert.ToDateTime(str_issuer_date).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                        item.Issuer = str_issuer;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabKOBADataFromISINWebsit()     : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void KOBAListDataFormat()
        {
            try
            {
                foreach (var item in kobaList)
                {
                    if (item.Ticker != null)
                    {
                        item.UpdateDate = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                        item.EffectiveDate = Convert.ToDateTime(item.EffectiveDate).ToString("yyyy-MMM-dd", new CultureInfo("en-US"));
                        item.FM = "1";
                        item.Ticker = item.Ticker.Substring(1).ToString();
                        item.RIC = item.Ticker + ".KS";
                        item.MatDate = Convert.ToDateTime(item.MatDate).ToString("yyyy-MMM-dd");
                        string _strike_price = item.StrikePrice.Contains(',') ? item.StrikePrice.Replace(",", "") : item.StrikePrice;
                        if (_strike_price.IndexOf('.') > 0)
                        {
                            _strike_price = _strike_price.TrimEnd(new char[] { '0', ' ' }).ToString();
                            if (_strike_price.Split('.')[1] == string.Empty)
                                _strike_price = _strike_price.Replace(".", "");
                        }
                        item.StrikePrice = _strike_price;
                        string _conversion_ratio = item.ConversionRatio.IndexOf('.') > 0 ? item.ConversionRatio.TrimEnd(new char[] { '0', ' ' }).ToString() : item.ConversionRatio;
                        if (string.IsNullOrEmpty(_conversion_ratio.Split('.')[1]) || _conversion_ratio.Split('.').Length < 1)
                            _conversion_ratio = _conversion_ratio.Replace(".", "").Trim().ToString();
                        item.ConversionRatio = _conversion_ratio;
                        item.QuanityofWarrants = item.QuanityofWarrants.Replace(",", "").ToString();
                        item.IssuePrice = item.IssuePrice.IndexOf('.') > 0 ? item.IssuePrice.Replace(",", "").Split('.')[0].ToString() : item.IssuePrice.Replace(",", "").ToString();
                        item.LastTradingDate = Convert.ToDateTime(item.LastTradingDate).ToString("dd-MMM-yy");
                        item.KoreaWarrantName = item.KoreaWarrantName.Contains("(주)") ? item.KoreaWarrantName.Replace("(주)", "") : item.KoreaWarrantName;
                        string _knockout_price = item.KnockOutPrice.Contains(',') ? item.KnockOutPrice.Replace(",", "").Trim().ToString() : item.KnockOutPrice;
                        if (_knockout_price.IndexOf('.') > 0)
                        {
                            _knockout_price = _knockout_price.Contains('.') ? _knockout_price.TrimEnd(new char[] { '0', ' ' }).ToString() : _knockout_price;
                            if (_knockout_price.Split('.')[1] == string.Empty)
                                _knockout_price = _knockout_price.Replace(".", "");
                        }
                        item.KnockOutPrice = _knockout_price;
                        //combine
                        string PorC = item.KoreaWarrantName.Substring((item.KoreaWarrantName.Length - 1));
                        item.CallOrPut = PorC == "콜" ? "CALL" : "PUT";
                        string last = item.CallOrPut == "CALL" ? "C" : "P";

                        char[] str_issuer = item.KoreaWarrantName.ToCharArray();    //노무라1960KOSPI200조기종료콜
                        string ikoreaname = "";
                        foreach (var x in str_issuer)
                        {
                            if (x > 47 && x < 58) break;
                            ikoreaname += x.ToString();
                        }
                        item.IssuerKoreaName = ikoreaname;

                        //IDN Display Name
                        KoreaIssuerInfo issuer = KoreaIssuerManager.SelectIssuer(item.IssuerKoreaName);
                        string sname = "*****";
                        string cname = "*****";
                        string underlying_ric = "*****";
                        if (issuer != null)
                        {
                            sname = issuer.IssuerCode4;

                            string NumLen = item.Ticker.Substring(2, 4);
                            string underly = "*****";
                            KoreaUnderlyingInfo underlying = KoreaUnderlyingManager.SelectUnderlying(item.UnderlyingKoreaName, KoreaNameType.KoreaNameForFM2);
                            if (underlying == null)
                            {
                                Logger.Log("Can not find underlying info for ELW FM2:" + item.UnderlyingKoreaName + ". Please input the ISIN.", Logger.LogType.Warning);
                                string isin = InputISIN.Prompt(item.UnderlyingKoreaName, "Korea Name For FM2");
                                if (!string.IsNullOrEmpty(isin))
                                {
                                    Logger.Log("User input ISIN:" + isin);
                                    underlying = KoreaUnderlyingManager.SelectUnderlyingByISIN(isin);
                                    if (underlying != null)
                                    {
                                        KoreaUnderlyingManager.UpdateKoreaNameFM2(item.UnderlyingKoreaName, isin);
                                        Logger.Log(string.Format("Update KoreaNameForFM2 as {0} for ISIN:{1}.", item.UnderlyingKoreaName, isin));
                                    }
                                    else
                                    {
                                        underlying = NewUnderlying.GrabNewUnderlyingInfo(item.UnderlyingKoreaName, isin);
                                        if (underlying != null)
                                        {
                                            string issuerName = item.IssuerKoreaName.Substring(4, item.IssuerKoreaName.Length - 5);
                                            //Regex regex = new Regex(@"\D+");
                                            //MatchCollection m = regex.Matches(item.Issuer_Korea_Name);
                                            //string issuerName = m[1].Value.Substring(0, m[1].Length - 1);
                                            underlying.UnderlyingName = issuerName;
                                            KoreaUnderlyingManager.UpdateUnderlying(underlying);
                                            newUnderLying.Add(underlying);
                                            Logger.Log("Insert new underlying to database successfully. RIC:" + underlying.UnderlyingRIC);
                                        }
                                        else
                                        {
                                            Logger.Log("Can not get new underlying info with ISIN:" + isin, Logger.LogType.Error);
                                        }
                                    }
                                }
                            }
                            if (underlying != null)
                            {
                                underly = underlying.IDNDisplayNamePart;
                                underlying_ric = underlying.UnderlyingRIC;
                                cname = underlying.QACommonNamePart;
                            }
                            string CharLen = item.UnderlyingKoreaName == "KOSPI200" ? "KOSPI" : underly;
                            item.IDNDisplayName = sname + NumLen + CharLen + "KO" + last;
                            if (item.IDNDisplayName.Contains("*****"))
                            {
                                item.IDNDisplayName = "**********";
                            }
                            item.BCASTREF = underlying_ric.ToUpper();

                            string mtime = Convert.ToDateTime(item.MatDate).ToString("MMM-yy", new CultureInfo("en-US")).Replace("-", "").ToUpper();
                            string price = item.StrikePrice.IndexOf('.') > 0 ? item.StrikePrice.Split('.')[0].Trim().ToString() : item.StrikePrice;
                            price = price.Length > 4 ? price.Substring(0, 4) : price;
                            string slast = item.UnderlyingKoreaName == "KOSPI200" ? "IW" : "WNT";
                            item.QACommonName = cname + " " + sname + " " + mtime + " " + price + " " + "KO " + last + slast;
                            item.QACommonName = item.QACommonName.ToUpper();

                            //Chain
                            item.Chain = "0#KOBA.KS";
                        }
                    }
                    else continue;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in KOBAListDataFormat     : \r\n" + ex.ToString(); Logger.Log(msg, Logger.LogType.Warning);
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GenerateKOBAFmFile()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            try
            {
                string filename = "Korea FM (KOBA Add)" + DateTime.Today.ToString("dd-MMM-yy").Replace("-", " ") + " (Afternoon).xls";
                string fpath = Path.Combine(configObj.FM, filename);
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, fpath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                GenerateKOBAExcelFileTitle(wSheet);

                int startLine = 5;
                LoopPrintKOBAData(wSheet, startLine, "common");

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                string msg = "Error found in GenerateKOBAFile_xls : " + ex.StackTrace + "  : " + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void GenerateKOBAExcelFileTitle(Worksheet wSheet)
        {
            if (wSheet.get_Range("C1", Type.Missing).Value2 == null)
            {
                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 5;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 35;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["Q", System.Type.Missing]).AutoFit();    //this is chain column 
                ((Range)wSheet.Columns["R", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["S", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["A:S", System.Type.Missing]).Font.Name = "Arial";


                ((Range)wSheet.Rows[4, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                ((Range)wSheet.Cells[1, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.PaleGreen);
                wSheet.Cells[1, 1] = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                ((Range)wSheet.Cells[3, 1]).Font.Underline = System.Drawing.FontStyle.Underline;
                ((Range)wSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                wSheet.Cells[3, 1] = "KOBA ADD 2";

                wSheet.Cells[4, 1] = "Updated Date";
                wSheet.Cells[4, 2] = "Effective Date";
                wSheet.Cells[4, 3] = "RIC";
                wSheet.Cells[4, 4] = "FM";
                wSheet.Cells[4, 5] = "IDN Display Name";
                wSheet.Cells[4, 6] = "ISIN";
                wSheet.Cells[4, 7] = "Ticker";
                wSheet.Cells[4, 8] = "BCAST_REF";
                wSheet.Cells[4, 9] = "QA Common Name";
                wSheet.Cells[4, 10] = "Mat Date";
                wSheet.Cells[4, 11] = "Strike Price";
                wSheet.Cells[4, 12] = "Quanity of Warrants";
                wSheet.Cells[4, 13] = "Issue Price";
                wSheet.Cells[4, 14] = "Issue Date";
                wSheet.Cells[4, 15] = "Conversion Ratio";
                wSheet.Cells[4, 16] = "Issuer";
                wSheet.Cells[4, 17] = "Chain";
                wSheet.Cells[4, 18] = "Last Trading Date";
                wSheet.Cells[4, 19] = "Kncok-out Price";
            }
        }

        private void GenerateExcelFileTitle(Worksheet wSheet)
        {
            if (wSheet.get_Range("C1", Type.Missing).Value2 == null)
            {
                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 18;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 5;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 35;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 10;
                ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 65;
                ((Range)wSheet.Columns["Q", System.Type.Missing]).AutoFit();    //this is chain column 
                ((Range)wSheet.Columns["R", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["S", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["A:S", System.Type.Missing]).Font.Name = "Arial";

                ((Range)wSheet.Rows[4, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                ((Range)wSheet.Cells[1, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.PaleGreen);
                wSheet.Cells[1, 1] = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                ((Range)wSheet.Cells[3, 1]).Font.Underline = System.Drawing.FontStyle.Underline;
                ((Range)wSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                wSheet.Cells[3, 1] = "WARRANT ADD 2";

                wSheet.Cells[4, 1] = "Updated Date";
                wSheet.Cells[4, 2] = "Effective Date";
                wSheet.Cells[4, 3] = "RIC";
                wSheet.Cells[4, 4] = "FM";
                wSheet.Cells[4, 5] = "IDN Display Name";
                wSheet.Cells[4, 6] = "ISIN";
                wSheet.Cells[4, 7] = "Ticker";
                wSheet.Cells[4, 8] = "BCAST_REF";
                wSheet.Cells[4, 9] = "QA Common Name";
                wSheet.Cells[4, 10] = "Mat Date";
                wSheet.Cells[4, 11] = "Strike Price";
                wSheet.Cells[4, 12] = "Quanity of Warrants";
                wSheet.Cells[4, 13] = "Issue Price";
                wSheet.Cells[4, 14] = "Issue Date";
                wSheet.Cells[4, 15] = "Conversion Ratio";
                wSheet.Cells[4, 16] = "Issuer";
                wSheet.Cells[4, 17] = "Chain";
                wSheet.Cells[4, 18] = "Last Trading Date";
                wSheet.Cells[4, 19] = "Kncok-out Price";
            }
        }

        private void LoopPrintKOBAData(Worksheet wSheet, int startLine, string type)
        {
            try
            {
                foreach (var item in kobaList)
                {
                    ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                    wSheet.Cells[startLine, 1] = item.UpdateDate;
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine, 2] = Convert.ToDateTime(item.EffectiveDate).ToString("dd-MMM-yy");
                    wSheet.Cells[startLine, 3] = item.RIC;
                    wSheet.Cells[startLine, 4] = item.FM;
                    wSheet.Cells[startLine, 5] = item.IDNDisplayName;
                    wSheet.Cells[startLine, 6] = item.ISIN;
                    ((Range)wSheet.Cells[startLine, 7]).NumberFormat = "@";
                    wSheet.Cells[startLine, 7] = item.Ticker;
                    wSheet.Cells[startLine, 8] = item.BCASTREF;
                    wSheet.Cells[startLine, 9] = item.QACommonName;
                    //((Range)wSheet.Cells[startLine, 10]).NumberFormat = "@";
                    wSheet.Cells[startLine, 10] = Convert.ToDateTime(item.MatDate).ToString("dd-MMM-yy");
                    wSheet.Cells[startLine, 11] = item.StrikePrice;
                    wSheet.Cells[startLine, 12] = item.QuanityofWarrants;
                    wSheet.Cells[startLine, 13] = item.IssuePrice;
                    //((Range)wSheet.Cells[startLine, 14]).NumberFormat = "@";
                    wSheet.Cells[startLine, 14] = Convert.ToDateTime(item.IssueDate).ToString("dd-MMM-yy");
                    wSheet.Cells[startLine, 15] = item.ConversionRatio;
                    wSheet.Cells[startLine, 16] = item.Issuer;
                    if (type.Equals("master"))
                    {
                        wSheet.Cells[startLine, 17] = item.KoreaWarrantName;
                        wSheet.Cells[startLine, 18] = item.Chain;
                        //((Range)wSheet.Cells[startLine, 19]).NumberFormat = "@";
                        //wSheet.Cells[startLine, 19] = item.LastTradingDate;
                        wSheet.Cells[startLine, 19] = "";
                        wSheet.Cells[startLine, 20] = item.KnockOutPrice;
                    }
                    else
                    {
                        wSheet.Cells[startLine, 17] = item.Chain;
                        //((Range)wSheet.Cells[startLine, 18]).NumberFormat = "@";
                        //wSheet.Cells[startLine, 18] = item.LastTradingDate;
                        wSheet.Cells[startLine, 18] = "";
                        wSheet.Cells[startLine, 19] = item.KnockOutPrice;
                    }
                    startLine++;
                }
                //startLine++;
                //wSheet.Cells[startLine, 1] = "EQUITY ADD";
                //((Range)wSheet.Cells[startLine, 1]).Font.Underline = System.Drawing.FontStyle.Underline;
                //wSheet.Cells[startLine + 2,1] = "- End -";
            }
            catch (Exception ex)
            {
                string msg = "Error found in LoopPrintKOBAData : " + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }
        #endregion

        #region FM2 work logic

        private void FM2ListDataFormat()
        {
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                if (fmSecondList.Count > 0)
                    foreach (var item in fmSecondList)
                    {
                        if (item.Ticker != null)
                        {
                            item.UpdateDate = DateTime.Today.ToString("dd-MMM-yy", new CultureInfo("en-US"));
                            item.EffectiveDate = Convert.ToDateTime(item.EffectiveDate).ToString("yyyy-MMM-dd");
                            item.FM = "2";
                            item.Ticker = item.Ticker.Substring(1).ToString();
                            item.RIC = item.Ticker + ".KS";
                            item.MatDate = Convert.ToDateTime(item.MatDate).ToString("yyyy-MMM-dd");
                            string _strike_price = item.StrikePrice.Contains(',') ? item.StrikePrice.Replace(",", "") : item.StrikePrice;
                            if (_strike_price.IndexOf('.') > 0)
                            {
                                _strike_price = _strike_price.TrimEnd(new char[] { '0', ' ' }).ToString();
                                if (_strike_price.Split('.')[1] == string.Empty)
                                    _strike_price = _strike_price.Replace(".", "");
                            }
                            item.StrikePrice = _strike_price;
                            item.QuanityofWarrants = item.QuanityofWarrants.Replace(",", "").ToString();
                            item.IssuePrice = item.IssuePrice.IndexOf('.') > 0 ? item.IssuePrice.Replace(",", "").Split('.')[0].ToString() : item.IssuePrice.Replace(",", "").ToString();
                            item.LastTradingDate = Convert.ToDateTime(item.LastTradingDate).ToString("dd-MMM-yy");
                            item.KoreaWarrantName = item.KoreaWarrantName.Contains("(주)") ? item.KoreaWarrantName.Replace("(주)", "") : item.KoreaWarrantName;
                        }
                        else
                            continue;
                    }
            }
            catch (Exception ex)
            {
                string msg = "Error found in FM2ListDataFormat()     : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void GenerateFMSecondPartFile()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            try
            {
                string filename = "Korea FM for " + DateTime.Today.ToString("dd-MMM-yy").Replace("-", " ") + " (Afternoon).xls";
                string ipath = Path.Combine(configObj.FM, filename);// commonObj.Log_Path + "\\" + commonObj.SubFolder + "\\" + filename;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Worksheet could not be created. Check that your office installation and project reference are correct!";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }
                GenerateExcelFileTitle(wSheet);

                int startLine = 5;

                LoopPrintELWFMSecond_xls(wSheet, startLine, 17);

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
            }
            catch (Exception ex)
            {
                string msg = "Error found in FM2CompareWithFM1_GenerateFMSecondFile_xls : " + ex.StackTrace + "     :   " + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void LoopPrintELWFMSecond_xls(Worksheet wSheet, int startLine, int Y_index)
        {
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                foreach (var item in fmSecondList)
                {
                    //if the RIC exist in FM1 master file
                    if (fmSecondHash.Contains(item.RIC))
                    {
                        //get FM1 data, assign to WarrantTemplate object
                        WarrantTemplate kw = fmSecondHash[item.RIC] as WarrantTemplate;
                        //format cell and assign value to each cell
                        ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                        wSheet.Cells[startLine, 1] = item.UpdateDate;
                        ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                        wSheet.Cells[startLine, 2] = Convert.ToDateTime(item.EffectiveDate).ToString("dd-MMM-yy");
                        wSheet.Cells[startLine, 3] = item.RIC;
                        wSheet.Cells[startLine, 4] = item.FM;
                        wSheet.Cells[startLine, 5] = kw.IDNDisplayName;
                        ((Range)wSheet.Cells[startLine, 6]).NumberFormat = "@";
                        wSheet.Cells[startLine, 6] = kw.ISIN;
                        ((Range)wSheet.Cells[startLine, 7]).NumberFormat = "@";
                        wSheet.Cells[startLine, 7] = kw.Ticker;
                        wSheet.Cells[startLine, 8] = kw.BCASTREF;
                        wSheet.Cells[startLine, 9] = kw.QACommonName;
                        ((Range)wSheet.Cells[startLine, 10]).NumberFormat = "@";
                        wSheet.Cells[startLine, 10] = Convert.ToDateTime(kw.MatDate).ToString("dd-MMM-yy");
                        wSheet.Cells[startLine, 11] = kw.StrikePrice;
                        wSheet.Cells[startLine, 12] = kw.QuanityofWarrants;
                        wSheet.Cells[startLine, 13] = kw.IssuePrice;
                        ((Range)wSheet.Cells[startLine, 14]).NumberFormat = "@";
                        wSheet.Cells[startLine, 14] = Convert.ToDateTime(kw.IssueDate).ToString("dd-MMM-yy");
                        wSheet.Cells[startLine, 15] = kw.ConversionRatio;
                        wSheet.Cells[startLine, 16] = kw.Issuer;
                        if (Y_index == 17)
                        {
                            wSheet.Cells[startLine, 17] = kw.Chain;
                            ((Range)wSheet.Cells[startLine, 18]).NumberFormat = "@";
                            //wSheet.Cells[startLine, 18] = item.LastTradingDate;
                            wSheet.Cells[startLine, 18] = "";
                        }
                        if (Y_index == 18)
                        {
                            wSheet.Cells[startLine, 17] = kw.KoreaWarrantName;
                            wSheet.Cells[startLine, 18] = kw.Chain;
                            ((Range)wSheet.Cells[startLine, 19]).NumberFormat = "@";
                            //wSheet.Cells[startLine, 19] = item.LastTradingDate;
                            wSheet.Cells[startLine, 19] = "";
                        }

                        item.IDNDisplayName = kw.IDNDisplayName;
                        item.IssueDate = kw.IssueDate;
                        item.Chain = kw.Chain;
                        item.BCASTREF = kw.BCASTREF;
                        item.QACommonName = kw.QACommonName;

                        //start compare FM2 and FM1 value

                        if (kw.Ticker != item.Ticker)
                        {
                            fm2HasChangedValue = true;
                            wSheet.Cells[startLine, 7] = item.Ticker;
                            kw.Ticker = item.Ticker;
                            ((Range)wSheet.Cells[startLine, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                        }
                        if (kw.ISIN != item.ISIN)
                        {
                            fm2HasChangedValue = true;
                            wSheet.Cells[startLine, 6] = item.ISIN;
                            kw.ISIN = item.ISIN;
                            if (emaFileHash[kw.ISIN] == null)
                            {
                                emaFileHash[kw.ISIN] = new ChangeddataModel();
                                ((ChangeddataModel)emaFileHash[kw.ISIN]).Secondary_ID = kw.ISIN;
                                ((ChangeddataModel)emaFileHash[kw.ISIN]).Secondary_ID_Changed = true;
                            }
                            ((Range)wSheet.Cells[startLine, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                        }


                        if (kw.MatDate != item.MatDate)
                        {
                            string _Mat_date = Convert.ToDateTime(item.MatDate).ToString("dd-MMM-yy");
                            fm2HasChangedValue = true;
                            wSheet.Cells[startLine, 10] = _Mat_date;
                            if (ndaFileHash[kw.RIC] == null)
                            {
                                ndaFileHash[kw.RIC] = new ChangeddataModel();
                                ((ChangeddataModel)ndaFileHash[kw.RIC]).RIC = kw.RIC;

                            }
                            ((ChangeddataModel)ndaFileHash[kw.RIC]).MatDate = item.MatDate;

                            ((Range)wSheet.Cells[startLine, 10]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

                            string mt = Convert.ToDateTime(item.MatDate).ToString("MMM-yy").Replace("-", "").ToUpper();
                            int count = kw.QACommonName.Split(' ').Length;
                            string qa_common_name_mat_date = kw.QACommonName.Split(' ')[(count - 3)].ToString().ToUpper();
                            if (qa_common_name_mat_date != mt)
                            {
                                kw.QACommonName = kw.QACommonName.Replace(qa_common_name_mat_date, mt).ToUpper();
                                wSheet.Cells[startLine, 9] = kw.QACommonName;
                                item.QACommonName = kw.QACommonName;
                                ((ChangeddataModel)ndaFileHash[kw.RIC]).QACommonName = kw.QACommonName;
                                ((Range)wSheet.Cells[startLine, 9]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                            }

                            //if Mat_Date change without ISIN change, need generate KR_ELW_ ddMMMyyyy_FM2_MD_N_update.txt 
                            if (kw.ISIN == item.ISIN)
                            {
                                if (gedaFileHash[kw.RIC] == null)
                                {
                                    gedaFileHash[kw.RIC] = new ChangeddataModel();
                                    ((ChangeddataModel)gedaFileHash[kw.RIC]).RIC = kw.RIC;
                                    ((ChangeddataModel)gedaFileHash[kw.RIC]).IDNDisplayName = kw.IDNDisplayName;
                                    ((ChangeddataModel)gedaFileHash[kw.RIC]).Warrant_Title = kw.KoreaWarrantName;
                                    ((ChangeddataModel)gedaFileHash[kw.RIC]).MatDate = DateTime.Parse(item.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
                                }
                            }

                        }

                        if (kw.QuanityofWarrants != item.QuanityofWarrants)
                        {
                            fm2HasChangedValue = true;
                            wSheet.Cells[startLine, 12] = item.QuanityofWarrants;
                            kw.QuanityofWarrants = item.QuanityofWarrants;
                            EmaQuantityChange wrt = new EmaQuantityChange();
                            wrt.ISAN = kw.ISIN;
                            wrt.QuanityofWarrants = kw.QuanityofWarrants;
                            quanityofWarrantsList.Add(wrt);
                            ((Range)wSheet.Cells[startLine, 12]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                        }

                        if (kw.StrikePrice != item.StrikePrice)
                        {
                            fm2HasChangedValue = true;
                            wSheet.Cells[startLine, 11] = item.StrikePrice;
                            if (ndaFileHash[kw.RIC] == null)
                            {
                                ndaFileHash[kw.RIC] = new ChangeddataModel();
                                ((ChangeddataModel)ndaFileHash[kw.RIC]).RIC = kw.RIC;
                            }
                            ((ChangeddataModel)ndaFileHash[kw.RIC]).StrikePrice = item.StrikePrice;
                            if (priceChange[item.ISIN] == null)
                            {
                                EmaPriceChange changeItem = new EmaPriceChange();
                                changeItem.ISIN = item.ISIN;
                                changeItem.ExcercisePrice = item.StrikePrice;
                                priceChange.Add(item.ISIN, changeItem);
                            }

                            ((Range)wSheet.Cells[startLine, 11]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                            #region as QACommonName is not exist in DB ELW

                            //string strike_p = item.StrikePrice.IndexOf('.') > 0 ? item.StrikePrice.Split('.')[0] : item.StrikePrice;
                            //strike_p = strike_p.Length > 4 == true ? strike_p.Substring(0, 4) : strike_p;
                            //int count = kw.QACommonName.Split(' ').Length;
                            //string qa_common_name_Num = "";
                            //qa_common_name_Num = kw.QACommonName.Split(' ')[(count - 2)].ToString();

                            //if (strike_p != qa_common_name_Num)
                            //{
                            //    kw.QACommonName = kw.QACommonName.Replace(qa_common_name_Num, strike_p).ToUpper();
                            //    wSheet.Cells[startLine, 9] = kw.QACommonName;
                            //    item.QACommonName = kw.QACommonName;
                            //    ((ChangeddataModel)ndaFileHash[kw.RIC]).QACommonName = kw.QACommonName;
                            //    ((Range)wSheet.Cells[startLine, 9]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                            //}




                            //int qaCommonNameIndex = kw.QACommonName.IndexOf("  ");
                            //string QACommonNamePart1 = kw.QACommonName.Substring(0, qaCommonNameIndex).Trim();
                            //int IDNDisplayNameIndex = kw.IDNDisplayName.IndexOf(kw.Ticker.Substring(2, 4));
                            //string qaCommonNamePart2 = kw.IDNDisplayName.Substring(0, IDNDisplayNameIndex);
                            //int matDateIndex = kw.MatDate.IndexOf("-");
                            //string qaCommonNamePart3 = kw.MatDate.Substring(matDateIndex + 1, 3).ToUpper() + kw.MatDate.Substring(matDateIndex - 2, 2);
                            //string qaCommonNamePart4 = kw.StrikePrice.Length >= 4 ? kw.StrikePrice.Substring(0, 4) : kw.StrikePrice.Substring(0, 3);
                            //int qaCommonNameLength = kw.QACommonName.Length;
                            //string qaCommonNamePart5 = kw.QACommonName.Substring(qaCommonNameIndex, qaCommonNameLength - qaCommonNameIndex);
                            //kw.QACommonName = string.Format("{0} {1} {2} {3} {4}", QACommonNamePart1, qaCommonNamePart2, qaCommonNamePart3, qaCommonNamePart4);



                            wSheet.Cells[startLine, 9] = kw.QACommonName;
                            ((Range)wSheet.Cells[startLine, 9]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                            #endregion
                        }

                        if (kw.IssuePrice != item.IssuePrice)
                        {
                            fm2HasChangedValue = true;
                            wSheet.Cells[startLine, 13] = item.IssuePrice;
                            kw.IssuePrice = item.IssuePrice;
                            if (emaFileHash[kw.ISIN] == null)
                            {
                                emaFileHash[kw.ISIN] = new ChangeddataModel();
                                ((ChangeddataModel)emaFileHash[kw.ISIN]).Secondary_ID = kw.ISIN;
                            }
                            ((ChangeddataModel)emaFileHash[kw.ISIN]).Issue_Price = kw.IssuePrice;
                            ((Range)wSheet.Cells[startLine, 13]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                        }
                        Double o_ratio;
                        Double n_ratio;
                        //if (Convert.ToDouble(kw.Conversion_Ratio) != Convert.ToDouble(item.Conversion_Ratio))
                        if (Double.TryParse(kw.ConversionRatio, out o_ratio) && Double.TryParse(item.ConversionRatio, out n_ratio))
                        {

                            if (o_ratio != n_ratio)
                            {
                                fm2HasChangedValue = true;
                                wSheet.Cells[startLine, 15] = n_ratio.ToString();
                                kw.ConversionRatio = n_ratio.ToString();
                                if (priceChange[item.ISIN] == null)
                                {
                                    EmaPriceChange changeItem = new EmaPriceChange();
                                    changeItem.ISIN = item.ISIN;
                                    priceChange.Add(item.ISIN, changeItem);
                                }
                                ((EmaPriceChange)priceChange[item.ISIN]).WarrantsPerUnderlying = CalculateWarrantsPerUnderlying(item.ConversionRatio);
                                if (emaFileHash[kw.ISIN] != null)
                                {
                                    emaFileHash[kw.ISIN] = new ChangeddataModel();
                                    ((ChangeddataModel)emaFileHash[kw.ISIN]).Secondary_ID = kw.ISIN;
                                    ((ChangeddataModel)emaFileHash[kw.ISIN]).Warrants_Per_Underlying = kw.ConversionRatio;
                                }
                                ((Range)wSheet.Cells[startLine, 15]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                            }
                        }

                        if (kw.KoreaWarrantName != item.KoreaWarrantName)
                        {
                            fm2HasChangedValue = true;
                            if (Y_index == 18)
                            {
                                wSheet.Cells[startLine, 17] = item.KoreaWarrantName;
                                ((Range)wSheet.Cells[startLine, 17]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                            }
                        }
                        /*======================================== PUBLIC ========================================*/
                        string line = item.KoreaWarrantName;
                        string _call_put = line.Substring((line.Length - 1), 1);
                        if (_call_put == "콜")
                            _call_put = "C";
                        else
                            _call_put = "P";
                        item.CallOrPut = _call_put;
                        string _ric = "***.***";
                        string _qa_name = "***";
                        string _idn_name = "***";

                        KoreaUnderlyingInfo underlying = KoreaUnderlyingManager.SelectUnderlying(item.UnderlyingKoreaName, KoreaNameType.KoreaNameForFM2);
                        if (underlying == null)
                        {
                            Logger.Log("Can not find underlying info for ELW FM2:" + item.UnderlyingKoreaName + "Please input the ISIN.", Logger.LogType.Warning);
                            string isin = InputISIN.Prompt(item.UnderlyingKoreaName, "Korea Name For FM2");
                            if (!string.IsNullOrEmpty(isin))
                            {
                                underlying = KoreaUnderlyingManager.SelectUnderlyingByISIN(isin);
                                if (underlying != null)
                                {
                                    KoreaUnderlyingManager.UpdateKoreaNameFM2(item.UnderlyingKoreaName, isin);
                                }
                                else
                                {
                                    underlying = NewUnderlying.GrabNewUnderlyingInfo(item.UnderlyingKoreaName, isin);
                                    if (underlying != null)
                                    {
                                        string issuerName = item.IssuerKoreaName.Substring(4, item.IssuerKoreaName.Length - 5);
                                        //Regex regex = new Regex(@"\D+");
                                        //MatchCollection m = regex.Matches(item.Issuer_Korea_Name);
                                        //string issuerName = m[1].Value.Substring(0, m[1].Length - 1);
                                        underlying.UnderlyingName = issuerName;
                                        KoreaUnderlyingManager.UpdateUnderlying(underlying);
                                        newUnderLying.Add(underlying);
                                        Logger.Log("Insert new underlying to database successfully. RIC:" + underlying.UnderlyingRIC);

                                        if (!refernceList.Contains(underlying.UnderlyingRIC))
                                        {
                                            ReferenceListUnderLyingModel referenceUnderlying = new ReferenceListUnderLyingModel();
                                            referenceUnderlying.UnderlyingRIC = underlying.UnderlyingRIC;
                                            referenceUnderlying.NDATCUnderlyingTitle = underlying.NDATCUnderlyingTitle;
                                            referenceUnderlying.UnderlyingCompanyName = "***";
                                            refernceList.Add(referenceUnderlying.UnderlyingRIC, referenceUnderlying);
                                        }
                                    }
                                    else
                                    {
                                        Logger.Log("Can not get new underlying info with ISIN:" + isin, Logger.LogType.Error);
                                    }
                                }
                            }
                        }
                        if (underlying != null)
                        {
                            _ric = underlying.UnderlyingRIC;
                            item.BCASTREF = _ric;
                            _qa_name = underlying.QACommonNamePart;
                            _idn_name = underlying.IDNDisplayNamePart;
                        }

                        /*==========================IDN Display Name==========================*/
                        ModifyIDNDisplayName(wSheet, startLine, kw, _call_put, _idn_name, _ric, item);

                        /*==========================Bcast Ref==========================*/
                        ModifyBcastRef(wSheet, startLine, kw, _ric);

                        /*==========================Chain==========================*/
                        ModifyQACommonName(wSheet, startLine, item, kw, _call_put, _qa_name);

                        /*==========================QA Common Name==========================*/
                        ModifyChain(wSheet, startLine, item, kw, Y_index);


                        startLine++;
                    }
                    else
                    {
                        noFm1List.Add(item);
                        fmSecondHash.Add(item.RIC, item);
                        GrabAndFormatData(item);
                        ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                        wSheet.Cells[startLine, 1] = item.UpdateDate;
                        ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                        wSheet.Cells[startLine, 2] = Convert.ToDateTime(item.EffectiveDate).ToString("dd-MMM-yy");
                        wSheet.Cells[startLine, 3] = item.RIC;
                        ((Range)wSheet.Cells[startLine, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                        wSheet.Cells[startLine, 4] = item.FM;

                        wSheet.Cells[startLine, 5] = item.IDNDisplayName;
                        ((Range)wSheet.Cells[startLine, 6]).NumberFormat = "@";
                        wSheet.Cells[startLine, 6] = item.ISIN;
                        ((Range)wSheet.Cells[startLine, 7]).NumberFormat = "@";
                        wSheet.Cells[startLine, 7] = item.Ticker;
                        wSheet.Cells[startLine, 8] = item.BCASTREF;
                        wSheet.Cells[startLine, 9] = item.QACommonName;
                        ((Range)wSheet.Cells[startLine, 10]).NumberFormat = "@";
                        wSheet.Cells[startLine, 10] = item.MatDate;
                        wSheet.Cells[startLine, 11] = item.StrikePrice;
                        wSheet.Cells[startLine, 12] = item.QuanityofWarrants;
                        wSheet.Cells[startLine, 13] = item.IssuePrice;
                        ((Range)wSheet.Cells[startLine, 14]).NumberFormat = "@";
                        wSheet.Cells[startLine, 14] = item.IssueDate;
                        wSheet.Cells[startLine, 15] = item.ConversionRatio;
                        wSheet.Cells[startLine, 16] = item.Issuer;
                        if (Y_index == 17)
                        {
                            wSheet.Cells[startLine, 17] = item.Chain;
                            ((Range)wSheet.Cells[startLine, 18]).NumberFormat = "@";
                            wSheet.Cells[startLine, 18] = item.LastTradingDate;
                        }
                        if (Y_index == 18)
                        {
                            wSheet.Cells[startLine, 17] = item.KoreaWarrantName;
                            wSheet.Cells[startLine, 18] = item.Chain;
                            ((Range)wSheet.Cells[startLine, 19]).NumberFormat = "@";
                            wSheet.Cells[startLine, 19] = item.LastTradingDate;
                        }
                        startLine++;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in LoopPrintELWFMSecond_xls : " + ex.ToString() + "\r\n";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        /// <summary>
        /// Calculate column:Warrants_Per_Underlying in EMA PRC file. Use warrant's ConvertionRatio
        /// </summary>
        /// <param name="convertionRatio">warrant's ConvertionRatio</param>
        /// <returns>value of Warrants_Per_Underlying</returns>
        private string CalculateWarrantsPerUnderlying(string convertionRatio)
        {
            string result = "";
            double division = 0;
            double tempRatio = 0;
            if (convertionRatio == null)
            {
                return result;
            }

            if (convertionRatio.Contains("%"))
            {
                convertionRatio = convertionRatio.Replace("%", "");
                if (Double.TryParse(convertionRatio, out tempRatio))
                {
                    tempRatio = tempRatio / 100;
                }
            }
            else if (!Double.TryParse(convertionRatio, out tempRatio))
            {
                return result;
            }

            if (tempRatio != 0)
            {
                division = 1 / tempRatio;
                division = Math.Round(division, 5);

                result = Convert.ToString(division);
                return result;
            }
            return result;
        }

        private void GrabAndFormatData(WarrantTemplate item)
        {
            try
            {
                string uri = string.Format("http://isin.krx.co.kr/jsp/BA_VW021.jsp?isu_cd={0}&modi=f&req_no=", item.ISIN);
                HtmlAgilityPack.HtmlDocument htc = new HtmlAgilityPack.HtmlDocument();
                AdvancedWebClient wc = new AdvancedWebClient();

                string pageSource = WebClientUtil.GetPageSource(wc, uri, 300000, null);
                if (!string.IsNullOrEmpty(pageSource))
                    htc.LoadHtml(pageSource);
                if (htc != null)
                {
                    HtmlNode table = htc.DocumentNode.SelectNodes("//table")[2];
                    string str_issuer = table.SelectSingleNode(".//tr[4]/td[2]").InnerText.Trim().ToString();
                    string str_issuer_date = table.SelectSingleNode(".//tr[6]/td[2]").InnerText.Trim().ToString();
                    item.IssueDate = Convert.ToDateTime(str_issuer_date).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    item.Issuer = str_issuer;
                }

                item.MatDate = Convert.ToDateTime(item.MatDate).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                /* ---------------------------------------------------------------------------------------------------- */

                //combine
                string PorC = item.KoreaWarrantName.Substring((item.KoreaWarrantName.Length - 1));
                item.CallOrPut = PorC == "콜" ? "CALL" : "PUT";
                string last = item.CallOrPut.Equals("CALL") ? "C" : "P";

                char[] str_issuer_arr = item.KoreaWarrantName.ToCharArray();
                string ikoreaname = string.Empty;
                foreach (var x in str_issuer_arr)
                {
                    if (x > 47 && x < 58) break;
                    ikoreaname += x.ToString();
                }

                item.IssuerKoreaName = ikoreaname;

                //IDN Display Name
                KoreaIssuerInfo issuer = KoreaIssuerManager.SelectIssuer(item.IssuerKoreaName);
                string sname = "***";
                string underlying_ric = "***";
                string cname = "***";
                if (issuer != null)
                {
                    sname = issuer.IssuerCode4;
                    string NumLen = item.Ticker.Substring(2, 4);
                    string underly = "***";
                    if (!string.IsNullOrEmpty(item.UnderlyingKoreaName))
                    {
                        KoreaUnderlyingInfo underlying = KoreaUnderlyingManager.SelectUnderlying(item.UnderlyingKoreaName, KoreaNameType.KoreaNameForFM2);
                        if (underlying == null)
                        {
                            Logger.Log("Can not find underlying info for ELW FM2:" + item.UnderlyingKoreaName + "Please input the ISIN.", Logger.LogType.Warning);
                            string isin = InputISIN.Prompt(item.UnderlyingKoreaName, "Korea Name For FM2");
                            if (!string.IsNullOrEmpty(isin))
                            {
                                underlying = KoreaUnderlyingManager.SelectUnderlyingByISIN(isin);
                                if (underlying != null)
                                {
                                    KoreaUnderlyingManager.UpdateKoreaNameFM2(item.UnderlyingKoreaName, isin);
                                }
                                else
                                {
                                    underlying = NewUnderlying.GrabNewUnderlyingInfo(item.UnderlyingKoreaName, isin);
                                    if (underlying != null)
                                    {
                                        string issuerName = item.IssuerKoreaName.Substring(4, item.IssuerKoreaName.Length - 5);
                                        //Regex regex = new Regex(@"\D+");
                                        //MatchCollection m = regex.Matches(item.Korea_Warrant_Name);
                                        //string issuerName = m[1].Value.Substring(0, m[1].Length - 1);
                                        underlying.UnderlyingName = issuerName;
                                        underlying.KoreaNameFM2 = item.UnderlyingKoreaName;
                                        KoreaUnderlyingManager.UpdateUnderlying(underlying);
                                        newUnderLying.Add(underlying);
                                        Logger.Log("Insert new underlying to database successfully. RIC:" + underlying.UnderlyingRIC);
                                    }
                                    else
                                    {
                                        Logger.Log("Can not get new underlying info with ISIN:" + isin, Logger.LogType.Error);
                                    }
                                }
                            }
                        }
                        if (underlying != null)
                        {
                            underly = underlying.IDNDisplayNamePart;
                            underlying_ric = underlying.UnderlyingRIC;
                            cname = underlying.QACommonNamePart;
                        }
                    }
                    string CharLen = item.UnderlyingKoreaName.Equals("KOSPI200") ? "KOSPI" : underly;
                    item.IDNDisplayName = sname + NumLen + CharLen + last;
                    if (item.IDNDisplayName.Contains("***"))
                        item.IDNDisplayName = "******";
                    else
                        item.IDNDisplayName = item.IDNDisplayName.ToUpper();

                    item.BCASTREF = underlying_ric.ToUpper();

                    //QA Common Name
                    string mtime = Convert.ToDateTime(item.MatDate).ToString("MMM-yy", new CultureInfo("en-US")).Replace("-", "").ToUpper();
                    string price = item.StrikePrice.Contains(".") ? item.StrikePrice.Split('.')[0] : item.StrikePrice;
                    price = price.Length >= 4 ? price.Substring(0, 4) : price;
                    string slast = item.UnderlyingKoreaName == "KOSPI200" ? "IW" : "WNT";
                    string qacommonname = cname + " " + sname + " " + mtime + " " + price + " " + last + slast;
                    item.QACommonName = qacommonname.ToUpper();

                    //Chain
                    string chain = string.Empty;
                    if (!item.BCASTREF.Contains("***"))
                        chain = item.UnderlyingKoreaName.Equals("KOSPI200") ? ("0#WARRANTS.KS, 0#ELW.KS, 0#.KS200W.KS") : "0#WARRANTS.KS, 0#ELW.KS, 0#CELW.KS, 0#" + item.BCASTREF.Split('.')[0] + "W." + item.BCASTREF.Split('.')[1];
                    else
                        chain = "***";
                    item.Chain = chain.ToUpper();
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabAndFormatData()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void ModifyIDNDisplayName(Worksheet wSheet, int startLine, WarrantTemplate kw, string _call_put, string _idn_name, string _ric, WarrantTemplate item)
        {
            try
            {
                string _new_idn_display_name = "";
                string CharLen = kw.IDNDisplayName.Substring(0, 8).ToString();
                _new_idn_display_name = CharLen + _idn_name + _call_put;
                item.IDNDisplayName = _new_idn_display_name;
                if (kw.IDNDisplayName != _new_idn_display_name)
                {

                    wSheet.Cells[startLine, 5] = _new_idn_display_name;
                    ((Range)wSheet.Cells[startLine, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

                    kw.IDNDisplayName = _new_idn_display_name;
                    if (ndaFileHash[kw.RIC] == null)
                    {
                        ndaFileHash[kw.RIC] = new ChangeddataModel();
                        ((ChangeddataModel)ndaFileHash[kw.RIC]).RIC = kw.RIC;
                    }
                    ((ChangeddataModel)ndaFileHash[kw.RIC]).IDNDisplayName = kw.IDNDisplayName;

                    if (_call_put == "C")
                    {
                        ((ChangeddataModel)ndaFileHash[kw.RIC]).CPOption = "CALL";
                    }
                    else
                    {
                        ((ChangeddataModel)ndaFileHash[kw.RIC]).CPOption = "PUT";
                    }

                    //fill warrant_type and warrant_title once IDN name changed.
                    if (emaFileHash[kw.ISIN] == null)
                    {
                        emaFileHash[kw.ISIN] = new ChangeddataModel();
                        ((ChangeddataModel)emaFileHash[kw.ISIN]).Secondary_ID = kw.ISIN;
                    }
                    ((ChangeddataModel)emaFileHash[kw.ISIN]).Warrant_Type = ((ChangeddataModel)ndaFileHash[kw.RIC]).CPOption;


                    string temp_warrant_title = string.Empty;
                    string ricCode = kw.RIC.Substring(0, 2);
                    if (refernceList[ricCode] != null)
                    {
                        temp_warrant_title += ((ReferenceListIssueModel)refernceList[ricCode]).NDATCIssuerTitle;
                    }

                    temp_warrant_title += "/ ";

                    if (refernceList[_ric] != null)
                    {
                        temp_warrant_title += ((ReferenceListUnderLyingModel)refernceList[_ric]).NDATCUnderlyingTitle;
                    }


                    if (kw.QACommonName.Substring(kw.QACommonName.Length - 2, 2) == "IW")
                        temp_warrant_title += " INDEX ";
                    else
                        temp_warrant_title += " SHS ";
                    temp_warrant_title += ((ChangeddataModel)ndaFileHash[kw.RIC]).CPOption;
                    temp_warrant_title += " WTS " + DateTime.Parse(item.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
                    ((ChangeddataModel)emaFileHash[kw.ISIN]).Warrant_Title = temp_warrant_title;


                    //if IDN_Display_Name change without ISIN change, need generate KR_ELW_ ddMMMyyyy_FM2_MD_N_update.txt 
                    if (kw.ISIN == item.ISIN)
                    {
                        if (gedaFileHash[kw.RIC] == null)
                        {
                            gedaFileHash[kw.RIC] = new ChangeddataModel();
                            ((ChangeddataModel)gedaFileHash[kw.RIC]).RIC = kw.RIC;
                            ((ChangeddataModel)gedaFileHash[kw.RIC]).IDNDisplayName = _new_idn_display_name;
                            ((ChangeddataModel)gedaFileHash[kw.RIC]).Warrant_Title = item.KoreaWarrantName;
                            ((ChangeddataModel)gedaFileHash[kw.RIC]).MatDate = DateTime.Parse(kw.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
                        }
                        else
                        {
                            ((ChangeddataModel)gedaFileHash[kw.RIC]).IDNDisplayName = _new_idn_display_name;
                            ((ChangeddataModel)gedaFileHash[kw.RIC]).Warrant_Title = item.KoreaWarrantName;
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in ModifyTheIDNDiplayName      : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        private void ModifyBcastRef(Worksheet wSheet, int startLine, WarrantTemplate kw, string _ric)
        {
            try
            {
                if (kw.BCASTREF != _ric)
                {
                    wSheet.Cells[startLine, 8] = _ric;
                    kw.BCASTREF = _ric;
                    if (emaFileHash[kw.ISIN] == null)
                    {
                        emaFileHash[kw.ISIN] = new ChangeddataModel();
                        ((ChangeddataModel)emaFileHash[kw.ISIN]).Secondary_ID = kw.ISIN;
                    }
                    ((ChangeddataModel)emaFileHash[kw.ISIN]).Underlying_RIC = kw.BCASTREF;

                    ((Range)wSheet.Cells[startLine, 8]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in ModifyTheBcastRef      : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
                return;
            }
        }

        private void ModifyQACommonName(Worksheet wSheet, int startLine, WarrantTemplate item, WarrantTemplate kw, string _call_put, string _qa_name)
        {
            try
            {
                //QA Common Name
                int IDNDisplayNameIndex = kw.IDNDisplayName.IndexOf(kw.Ticker.Substring(2, 4));
                string issuerCode4 = item.IDNDisplayName.Substring(0, IDNDisplayNameIndex);
                KoreaUnderlyingInfo underlying = KoreaUnderlyingManager.SelectUnderlying(item.UnderlyingKoreaName);
                string slast = "WNT";

                if (underlying != null && underlying.UnderlyingRIC.StartsWith("."))
                    slast = "IW";

                string mtime = Convert.ToDateTime(item.MatDate).ToString("MMM-yy", new CultureInfo("en-US")).Replace("-", "").ToUpper();
                string price = item.StrikePrice.Contains(".") ? item.StrikePrice.Split('.')[0] : item.StrikePrice;
                price = price.Length >= 4 ? price.Substring(0, 4) : price;
                string last = item.CallOrPut;
                item.QACommonName = string.Format("{0} {1} {2} {3} {4}{5}", _qa_name, issuerCode4, mtime, price, last, slast);

                wSheet.Cells[startLine, 9] = item.QACommonName;
                kw.QACommonName = item.QACommonName;
                if (ndaFileHash[kw.RIC] == null)
                {
                    ndaFileHash[kw.RIC] = new ChangeddataModel();
                    ((ChangeddataModel)ndaFileHash[kw.RIC]).RIC = kw.RIC;
                }
                ((ChangeddataModel)ndaFileHash[kw.RIC]).QACommonName = kw.QACommonName;
                ((Range)wSheet.Cells[startLine, 9]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

                //String[] qa_array = kw.QACommonName.Split(' ');
                //int len = qa_array.Length;
                //string _underlying_qa = "";
                //string _last = "";
                //_last = item.UnderlyingKoreaName == "KOSPI200" ? _call_put + "IW" : _call_put + "WNT";
                //_underlying_qa = _qa_name + " " + qa_array[(len - 4)] + " " + qa_array[(len - 3)] + " " + qa_array[(len - 2)] + " " + _last;
                //_underlying_qa = _underlying_qa.ToUpper();
                //item.QACommonName = _underlying_qa;

                //if (kw.QACommonName != _underlying_qa)
                //{
                //    wSheet.Cells[startLine, 9] = _underlying_qa;
                //    kw.QACommonName = _underlying_qa;
                //    if (ndaFileHash[kw.RIC] == null)
                //    {
                //        ndaFileHash[kw.RIC] = new ChangeddataModel();
                //        ((ChangeddataModel)ndaFileHash[kw.RIC]).RIC = kw.RIC;
                //    }
                //    ((ChangeddataModel)ndaFileHash[kw.RIC]).QACommonName = kw.QACommonName;
                //    ((Range)wSheet.Cells[startLine, 9]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                //}
            }
            catch (Exception ex)
            {
                string msg = "Error found in ModifyTheQACommonName     : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void ModifyChain(Worksheet wSheet, int startLine, WarrantTemplate item, WarrantTemplate kw, int Y_index)
        {
            try
            {
                string _chain = "";
                if (item.UnderlyingKoreaName != "KOSPI200")
                {
                    _chain = "0#WARRANTS.KS, 0#ELW.KS, 0#CELW.KS, 0#" + kw.BCASTREF.Split('.')[0] + "W." + kw.BCASTREF.Split('.')[1];
                }
                else
                {
                    _chain = "0#WARRANTS.KS, 0#ELW.KS, 0#.KS200W.KS";
                }
                item.Chain = _chain;
                if (kw.Chain != _chain)
                {
                    wSheet.Cells[startLine, Y_index] = _chain;
                    kw.Chain = _chain;
                    ((Range)wSheet.Cells[startLine, Y_index]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("Error found in Modify the Chain. UnderlyingName:{0}. UnderlyingRIC:{1}. ELW RIC:{2}.", item.UnderlyingKoreaName, kw.BCASTREF, item.RIC);
                Logger.Log(msg, Logger.LogType.Warning);
                Logger.Log(ex.Message, Logger.LogType.Error);
            }
        }
        #endregion

        #region FM2 EMA file and NDA file and GEDA file

        private void GenerateEmaGedaNdaAndCreatEmail()
        {
            GeneretaFmSecondPartEmaFile();
            if (noFm1List.Count > 0)
            {
                GeneretaNoFm1EmaFile();
                GenerateNDAQAFile(configObj.BulkFile);
                GenerateNDAIAFile(configObj.BulkFile);
            }
            GenerateQuanityofWarrantsCsv();
            GeneratePriceChangeCsv();
            GeneretaFmSecondPartGedaFiles();
            GenerateFmSecondPartNdaFiles();
            GenerateFmSecondPartEmail();
        }

        private void GeneratePriceChangeCsv()
        {
            if (priceChange.Count == 0)
            {
                return;
            }

            string path = GeneratePriceChangeCsvEmaFilePath();
            List<List<string>> emaRes = new List<List<string>>();
            List<string> title = new List<string>() { "Logical_Key", "Secondary_ID", "Secondary_ID_Type", "Action", "Exotic1_Parameter", "Exotic1_Value", 
                                               "EH_Exercise_Price", "Exercise_Price", "EH_Warrants_Per_Underlying", "Warrants_Per_Underlying" };
            int line = 0;
            if (File.Exists(path))
            {
                line = ReadLogic(path);
            }
            ArrayList priceChangeArr = new ArrayList(priceChange.Keys);
            priceChangeArr.Sort();
            for (int i = 0; i < priceChangeArr.Count; i++)
            {
                string isin = priceChangeArr[i].ToString();
                EmaPriceChange changeItem = priceChange[isin] as EmaPriceChange;
                List<string> temp = new List<string>();
                temp.Add("" + (i + 1 + line));
                temp.Add(changeItem.ISIN);
                temp.Add("ISIN");
                temp.Add("");
                temp.Add("");
                temp.Add("");
                if (string.IsNullOrEmpty(changeItem.ExcercisePrice))
                {
                    temp.Add("");
                    temp.Add("");
                }
                else
                {
                    temp.Add("N");
                    temp.Add(changeItem.ExcercisePrice);
                }
                if (string.IsNullOrEmpty(changeItem.WarrantsPerUnderlying))
                {
                    temp.Add("");
                    temp.Add("");
                }
                else
                {
                    temp.Add("N");
                    temp.Add(changeItem.WarrantsPerUnderlying);
                }

                emaRes.Add(temp);
            }
            FileUtil.WriteOutputFile(path, emaRes, title, WriteMode.Append);
            AddResult(Path.GetFileNameWithoutExtension(path), path, "EMA File");

        }

        private string GeneratePriceChangeCsvEmaFilePath()
        {
            string dir = ConfigureOperator.GetEmaFileSaveDir();
            string sendDir = dir + "\\" + DateTime.Now.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            if (!Directory.Exists(sendDir))
                Directory.CreateDirectory(sendDir);

            string path = sendDir + @"\WRT_PRC_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_Korea.csv";
            return path;
        }

        private int ReadLogic(string filePth)
        {
            StreamReader readFileAll = new StreamReader(filePth);
            int count = 0;
            while (!readFileAll.EndOfStream)
            {
                readFileAll.ReadLine();
                count++;
            }
            readFileAll.Close();
            return count - 1;

        }
        private string GenerateQuanityofWarrantsCsvEmaFilePath()
        {
            string dir = ConfigureOperator.GetEmaFileSaveDir();
            string sendDir = dir + "\\" + DateTime.Now.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            if (!Directory.Exists(sendDir))
                Directory.CreateDirectory(sendDir);

            string path = sendDir + @"\WRT_QUA_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_Korea.csv";
            return path;
        }
        private void GenerateQuanityofWarrantsCsv()
        {
            if (quanityofWarrantsList.Count == 0)
            {
                return;
            }
            string path = GenerateQuanityofWarrantsCsvEmaFilePath();
            List<List<string>> emaRes = new List<List<string>>();
            string[] headStrs = new string[] { "Logical_Key", "Secondary_ID", "Secondary_ID_Type", "EH_Issue_Quantity", "Issue_Quantity" };
            int line = 0;
            if (File.Exists(path))
            {
                line = ReadLogic(path);
            }
            emaRes.Add(headStrs.ToList());
            for (int i = 0; i < quanityofWarrantsList.Count; i++)
            {
                List<string> temp = new List<string>();
                temp.Add("" + (i + 1 + line));
                temp.Add(quanityofWarrantsList[i].ISAN);
                temp.Add("ISIN");
                temp.Add("N");
                temp.Add(quanityofWarrantsList[i].QuanityofWarrants);
                emaRes.Add(temp);
            }
            if (emaRes.Count > 1)
            {
                if (File.Exists(path))
                {
                    emaRes.RemoveAt(0);//remove head
                    OperateExcel.WriteToCSV(path, emaRes, FileMode.Append);
                }
                else
                {
                    OperateExcel.WriteToCSV(path, emaRes, FileMode.Create);
                }

                AddResult(Path.GetFileNameWithoutExtension(path), path, "EMA File");
            }
        }
        private string GenerateSecondPartEmaFilePath()
        {
            string dir = ConfigureOperator.GetEmaFileSaveDir();
            string sendDir = dir + "\\" + DateTime.Now.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            if (!Directory.Exists(sendDir))
                Directory.CreateDirectory(sendDir);

            string path = sendDir + @"\WRT_MOD_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_Korea.csv";
            return path;
        }
        private void GeneretaFmSecondPartEmaFile()
        {

            List<List<string>> emaRes = new List<List<string>>();
            string[] headStrs = new string[] { "Logical_Key", "Secondary_ID", "Secondary_ID_Type", "Warrant_Title", "Issuer_OrgId", "Issue_Date", "Country_Of_Issue", "Governing_Country", "Announcement_Date", "Payment_Date", "Underlying_Type", "Action1", "Clearinghouse1_OrgId", "Action2", "Guarantor", "Guarantor_Type", "Guarantee_Type", "Incr_Exercise_Lot", "Min_Exercise_Lot", "Max_Exercise_Lot", "Rt_Page_Range", "Action3", "Underwriter1_OrgId", "Underwriter1_Role", "Exercise_Style", "Warrant_Type", "EH_Expiration_Date", "Expiration_Date", "Registered_Bearer_Code", "Price_Display_Type", "Private_Placement", "Coverage_Type", "Warrant_Status", "Status_Date", "Redemption_Method", "EH_Issue_Quantity", "Issue_Quantity", "Issue_Price", "Issue_Currency", "Issue_Price_Type", "Issue_Spot_Price", "Issue_Spot_Currency", "Issue_Spot_FX_Rate", "Issue_Delta", "Issue_Elasticity", "Issue_Gearing", "Issue_Premium", "Issue_Premium_Pa", "Denominated_Amount", "EH_Begin_Date", "Exercise_Begin_Date", "EH_End_Date", "Exercise_End_Date", "Offset_Number", "Period_Number", "Offset_Frequency", "Offset_Calendar", "Period_Calendar", "Period_Frequency", "RAF_Event_Type", "EH_Exercise_Price", "Exercise_Price", "Exercise_Price_Type", "EH_Exercise_Quantity", "Warrants_Per_Underlying", "Underlying_FX_Rate", "Action4", "Underlying_RIC", "EH_Undly_Quantity", "Underlying_Item_Quantity", "Units", "Cash_Currency", "Delivery_Type", "Settlement_Type", "Settlement_Currency", "Underlying_Group", "Action5", "Group_Long_Name", "Group_Short_Name", "Group_Mini_Name", "Action6", "Country1_Code", "Coverage1_Type", "Action7", "Note1_Type", "Note1", "Action8", "Exotic1_Parameter", "Exotic1_Value", "Exotic1_Begin_Date", "Exotic1_End_Date", "Action9", "Event_Type1", "Period_Number1", "Calendar_Type1", "Frequency1", "Action10", "Exchange_Code", "Incr_Trade_Lot", "Min_Trade_Lot", "Min_Trade_Amount", "Action11", "Attatched_To_ID", "Attacted_To_ID_Type", "Attached_Quantity", "Attached_Code", "Detachable_Date", "Bond_Exercise", "Bond_Price_Percentage" };
            emaRes.Add(headStrs.ToList());
            IEnumerator enumrator = emaFileHash.Values.GetEnumerator();
            string path = GenerateSecondPartEmaFilePath();
            int count = 1;

            while (enumrator.MoveNext())
            {
                emaRes.Add((new string[109]).ToList());
                emaRes[count][0] = (count).ToString();
                emaRes[count][1] = ((ChangeddataModel)enumrator.Current).Secondary_ID;
                emaRes[count][2] = "ISIN";
                //warrant_title
                emaRes[count][3] = ((ChangeddataModel)enumrator.Current).Warrant_Title;
                //warrant_type
                emaRes[count][25] = ((ChangeddataModel)enumrator.Current).Warrant_Type;
                emaRes[count][37] = ((ChangeddataModel)enumrator.Current).Issue_Price;
                if (!string.IsNullOrEmpty(((ChangeddataModel)enumrator.Current).Warrants_Per_Underlying))
                    emaRes[count][64] = (1.0 / Convert.ToDouble(((ChangeddataModel)enumrator.Current).Warrants_Per_Underlying)).ToString();
                emaRes[count][67] = ((ChangeddataModel)enumrator.Current).Underlying_RIC;
                count++;
            }
            if (emaRes.Count > 1)
            {
                OperateExcel.WriteToCSV(path, emaRes);
                AddResult(Path.GetFileNameWithoutExtension(path), path, "EMA File");
            }
        }
        private string GenerateNoFm1EmaFilePath()
        {
            string dir = ConfigureOperator.GetEmaFileSaveDir();
            string sendDir = dir + "\\" + DateTime.Now.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            if (!Directory.Exists(sendDir))
                Directory.CreateDirectory(sendDir);
            string path = sendDir + @"\WRT_ADD_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_Korea_Afternoon.csv";
            return path;
        }
        private void GeneretaNoFm1EmaFile()
        {
            string filePath = GenerateNoFm1EmaFilePath();
            List<List<string>> emaRes = new List<List<string>>();
            string[] headStrs = new string[] {
                "Logical_Key","Secondary_ID","Secondary_ID_Type","Warrant_Title","Issuer_OrgId","Issue_Date","Country_Of_Issue","Governing_Country","Announcement_Date","Payment_Date","Underlying_Type","Clearinghouse1_OrgId","Clearinghouse2_OrgId","Clearinghouse3_OrgId","Guarantor","Guarantor_Type","Guarantee_Type","Incr_Exercise_Lot","Min_Exercise_Lot","Max_Exercise_Lot","Rt_Page_Range","Underwriter1_OrgId","Underwriter1_Role","Underwriter2_OrgId","Underwriter2_Role","Underwriter3_OrgId","Underwriter3_Role","Underwriter4_OrgId","Underwriter4_Role","Exercise_Style","Warrant_Type","Expiration_Date","Registered_Bearer_Code","Price_Display_Type","Private_Placement","Coverage_Type","Warrant_Status","Status_Date","Redemption_Method","Issue_Quantity","Issue_Price","Issue_Currency","Issue_Price_Type","Issue_Spot_Price","Issue_Spot_Currency","Issue_Spot_FX_Rate","Issue_Delta","Issue_Elasticity","Issue_Gearing","Issue_Premium","Issue_Premium_PA","Denominated_Amount","Exercise_Begin_Date","Exercise_End_Date","Offset_Number","Period_Number","Offset_Frequency","Offset_Calendar","Period_Calendar","Period_Frequency","RAF_Event_Type","Exercise_Price","Exercise_Price_Type","Warrants_Per_Underlying","Underlying_FX_Rate","Underlying_RIC","Underlying_Item_Quantity","Units","Cash_Currency","Delivery_Type","Settlement_Type","Settlement_Currency","Underlying_Group","Country1_Code","Coverage1_Type","Country2_Code","Coverage2_Type","Country3_Code","Coverage3_Type","Country4_Code","Coverage4_Type","Country5_Code","Coverage5_Type","Note1_Type","Note1","Note2_Type","Note2","Note3_Type","Note3","Note4_Type","Note4","Note5_Type","Note5","Note6_Type","Note6","Exotic1_Parameter","Exotic1_Value","Exotic1_Begin_Date","Exotic1_End_Date","Exotic2_Parameter","Exotic2_Value","Exotic2_Begin_Date","Exotic2_End_Date","Exotic3_Parameter","Exotic3_Value","Exotic3_Begin_Date","Exotic3_End_Date","Exotic4_Parameter","Exotic4_Value","Exotic4_Begin_Date","Exotic4_End_Date","Exotic5_Parameter","Exotic5_Value","Exotic5_Begin_Date","Exotic5_End_Date","Exotic6_Parameter","Exotic6_Value","Exotic6_Begin_Date","Exotic6_End_Date","Event_Type1","Event_Period_Number1","Event_Calendar_Type1","Event_Frequency1","Event_Type2","Event_Period_Number2","Event_Calendar_Type2","Event_Frequency2","Exchange_Code1","Incr_Trade_Lot1","Min_Trade_Lot1","Min_Trade_Amount1","Exchange_Code2","Incr_Trade_Lot2","Min_Trade_Lot2","Min_Trade_Amount2","Exchange_Code3","Incr_Trade_Lot3","Min_Trade_Lot3","Min_Trade_Amount3","Exchange_Code4","Incr_Trade_Lot4","Min_Trade_Lot4","Min_Trade_Amount4","Attached_To_Id","Attached_To_Id_Type","Attached_Quantity","Attached_Code","Detachable_Date","Bond_Exercise","Bond_Price_Percentage"
            };
            emaRes.Add(headStrs.ToList());
            int count = 1;

            try
            {

                foreach (var term in noFm1List)
                {
                    emaRes.Add((new string[150]).ToList());
                    emaRes[count][0] = (count).ToString();
                    emaRes[count][1] = term.ISIN;
                    emaRes[count][2] = "ISIN";
                    emaRes[count][5] = DateTime.Parse(term.IssueDate).ToString("dd/MM/yyyy", new CultureInfo("en-US"));
                    emaRes[count][6] = "KOR";
                    emaRes[count][7] = "KOR";
                    if (term.QACommonName.Substring(term.QACommonName.Length - 2, 2).Equals("IW"))
                        emaRes[count][10] = "INDEX";
                    else
                        emaRes[count][10] = "STOCK";
                    emaRes[count][17] = "10";
                    emaRes[count][18] = "10";
                    emaRes[count][29] = "E";
                    if (term.IDNDisplayName.Substring(term.IDNDisplayName.Length - 1, 1).Equals("C"))
                        emaRes[count][30] = "Call";
                    else
                        emaRes[count][30] = "Put";
                    emaRes[count][31] = DateTime.Parse(term.MatDate).ToString("dd/MM/yyyy", new CultureInfo("en-US"));
                    emaRes[count][32] = "R";
                    emaRes[count][33] = "D";
                    emaRes[count][39] = term.QuanityofWarrants;
                    emaRes[count][40] = term.IssuePrice;
                    emaRes[count][41] = "KRW";
                    emaRes[count][42] = "A";

                    emaRes[count][51] = "10";
                    emaRes[count][52] = DateTime.Parse(term.MatDate).ToString("dd/MM/yyyy", new CultureInfo("en-US"));
                    emaRes[count][53] = DateTime.Parse(term.MatDate).ToString("dd/MM/yyyy", new CultureInfo("en-US"));


                    emaRes[count][61] = term.StrikePrice;//Exercise_Price need to be confirm 

                    emaRes[count][62] = "A";
                    emaRes[count][63] = 1.0 / Convert.ToDouble(term.ConversionRatio) + "";
                    emaRes[count][65] = term.BCASTREF;
                    emaRes[count][66] = "1";
                    if (term.QACommonName.Substring(term.QACommonName.Length - 2, 2).Equals("IW"))
                        emaRes[count][67] = "idx";
                    else
                        emaRes[count][67] = "shr";
                    if (term.QACommonName.Substring(term.QACommonName.Length - 2, 2).Equals("IW"))
                        emaRes[count][69] = "I";
                    else
                        emaRes[count][69] = "S";
                    emaRes[count][70] = "C";
                    emaRes[count][71] = "KRW";
                    emaRes[count][83] = "T";
                    emaRes[count][84] = "Last Trading Day is " + DateTime.Parse(term.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")) + ".";
                    emaRes[count][85] = "S";
                    // emaRes[count][86] = "Note2";//Note2 need to be confirm 
                    if (term.RIC.Split(".".ToArray())[1].Equals("KS"))
                        emaRes[count][127] = "KSC";
                    else
                        emaRes[count][127] = "KOE";
                    emaRes[count][128] = "10";
                    emaRes[count][129] = "10";
                    string code = term.RIC.Substring(0, 2);
                    var issueRefer = refernceList[code] as ReferenceListIssueModel;
                    var uderLyingRefer = refernceList[term.BCASTREF] as ReferenceListUnderLyingModel;
                    emaRes[count][3] = issueRefer.NDATCIssuerTitle + "/" + " " + uderLyingRefer.NDATCUnderlyingTitle + " " + (emaRes[count][10] == "STOCK" ? "SHS" : emaRes[count][10]) + " " + emaRes[count][30].ToUpper() + " " + "WTS" + " " + DateTime.Parse(term.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")).ToUpper();
                    emaRes[count][4] = issueRefer.NDAIssuerORGID;

                    string RR = emaRes[count][65] + " " + emaRes[count][30];
                    string note2 = "";
                    switch (RR)
                    {
                        case ".KS200 Call":
                            note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Closing level on last trade date - Strike level)* 100";
                            break;
                        case ".KS200 Put":
                            note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Strike level - Closing level on last trade date)* 100";
                            break;
                        case ".N225 Call":
                            note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Closing level on last trade date - Strike level)* Conversion ratio";
                            break;
                        case ".N225 Put":
                            note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Strike level - Closing level on last trade date)* Conversion ratio";
                            break;
                        case ".HSI Call":
                            note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Average closing level during the last 5 trading days(including last trading day) - Strike level)* Conversion ratio";
                            break;
                        case ".HSI Put":
                            note2 = "Settlement amount per 10 Warrants(1 LOT) = 10* (Strike level - Average closing level during the last 5 trading days(including last trading day))* Conversion ratio";
                            break;
                        default:
                            note2 = "Average closing price of the underlying stock during the last 5 trading days(including last trading day) before expiry.";
                            break;
                    }
                    emaRes[count][86] = note2;
                    count++;
                }

            }
            catch (Exception ex)
            {
                string errInfo = ex.ToString();
                int ss = count;
            }
            if (emaRes.Count > 1)
            {
                OperateExcel.WriteToCSV(filePath, emaRes);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "EMA File");
            }
        }
        //need to be confirm
        private void GenerateFM2BcuUpdate(string gedaDir)
        {
            List<List<string>> res = new List<List<string>>();
            List<string> head = new List<string>();
            string filePath = Path.Combine(gedaDir, "KR_ELW_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_ADD_FM2_BCU_UPDATE.txt");
            head.Add("RIC");
            head.Add("\t");
            head.Add("BCU");
            res.Add(head);
            foreach (var item in fmSecondList)
            {
                List<string> tempRes = new List<string>();
                if (fmSecondHash.Contains(item.RIC))
                {
                    WarrantTemplate kw = fmSecondHash[item.RIC] as WarrantTemplate;
                    string chain = kw.Chain;
                    string ric = kw.RIC;
                    string ticker = kw.BCASTREF.Substring(0, 6);
                    string bcu = "KSE_EQ_WARRANTS,KSE_EQB_ELW,";
                    //MessageBox.Show(chain.Substring(chain.Length - 2, 2));
                    if (chain.Contains("0#.KS200W.KS"))
                    {
                        bcu += "KSE_INDEX_KS200W";
                    }
                    else
                    {
                        bcu += "KSE_EQ_CELW,";
                        if (chain.Substring(chain.Length - 2, 2).Equals("KS"))
                        {
                            bcu += "KSE_STOCK_";
                            bcu += ticker + "W";
                        }
                        else if (chain.Substring(chain.Length - 2, 2).Equals("KQ"))
                        {
                            bcu += "KOSDAQ_STOCK_";
                            bcu += ticker + "W";
                        }
                    }
                    tempRes.Add(ric);
                    tempRes.Add("\t");
                    tempRes.Add(bcu);
                    res.Add(tempRes);
                }

            }
            if (res.Count > 1)
            {
                OperateExcel.WriteToTXT(filePath, res);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "GEDA File");
            }

        }
        private List<string> GenerateOneIsanRes(string isan)
        {
            List<string> item = new List<string>();
            foreach (var fm2 in fmSecondList)
            {
                if (fm2.ISIN == isan)
                {
                    WarrantTemplate temp = fmSecondHash[fm2.RIC] as WarrantTemplate;
                    string ticker = temp.BCASTREF.Substring(0, 6);

                    item.Add(temp.RIC);
                    item.Add("\t");
                    item.Add(temp.IDNDisplayName);
                    item.Add("\t");
                    item.Add(temp.RIC);
                    item.Add("\t");
                    item.Add(temp.Ticker);
                    item.Add("\t");
                    item.Add(temp.ISIN);
                    item.Add("\t");
                    item.Add(temp.BCASTREF);
                    item.Add("\t");
                    item.Add("KSE_EQB_ELW");
                    item.Add("\t");
                    item.Add(temp.Ticker);
                    item.Add("\t");
                    item.Add(temp.ISIN);
                    item.Add("\t");
                    item.Add("J" + temp.Ticker);
                    item.Add("\t");
                    item.Add(fm2.KoreaWarrantName);
                    item.Add("\t");
                    item.Add(DateTime.Parse(fm2.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                    item.Add("\t");
                    item.Add("ELW/" + temp.Ticker);
                    item.Add("\t");
                    string chain = temp.Chain;
                    string bcu = "KSE_EQ_WARRANTS,KSE_EQB_ELW,";
                    //MessageBox.Show(chain.Substring(chain.Length - 2, 2));
                    if (chain.Contains("0#.KS200W.KS"))
                    {
                        bcu += "KSE_INDEX_KS200W";
                    }
                    else
                    {
                        bcu += "KSE_EQ_CELW,";
                        if (chain.Substring(chain.Length - 2, 2).Equals("KS"))
                        {
                            bcu += "KSE_STOCK_";
                            bcu += ticker + "W";
                        }
                        else if (chain.Substring(chain.Length - 2, 2).Equals("KQ"))
                        {
                            bcu += "KOSDAQ_STOCK_";
                            bcu += ticker + "W";
                        }
                    }
                    item.Add(bcu);
                }
            }
            return item;
        }
        private void GenerateISINChangeRicCreation(string gedaDir)
        {
            List<List<string>> res = new List<List<string>>();
            string filePath = gedaDir + "\\KSE_ELW_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_ISIN_CHANGE_RIC_RECREATION.txt";
            string[] headStrs = new string[] { "SYMBOL", "\t", "DSPLY_NAME", "\t", "RIC", "\t", "OFFCL_CODE", "\t", "EX_SYMBOL", "\t", "BCAST_REF", "\t", "EXL_NAME", "\t", "#INSTMOD_TDN_SYMBOL", "\t", "#INSTMOD_#ISIN", "\t", "#INSTMOD_MNEMONIC", "\t", "DSPLY_NMLL", "\t", "MATUR_DATE", "\t", "#INSTMOD_LONGLINK2", "\t", "BCU" };
            string dropFilePath = gedaDir + "\\KSE_ELW_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_ISIN_CHANGE_RIC_DROP.txt";
            List<List<string>> dropRes = new List<List<string>>();
            res.Add(headStrs.ToList());//head
            dropRes.Add(new string[] { "RIC" }.ToList());//Drop head
            IEnumerator enumrator = emaFileHash.Values.GetEnumerator();
            while (enumrator.MoveNext())
            {
                if (((ChangeddataModel)enumrator.Current).Secondary_ID_Changed)
                {
                    res.Add(GenerateOneIsanRes(((ChangeddataModel)enumrator.Current).Secondary_ID));
                    List<string> oneDropRes = new List<string>();
                    oneDropRes.Add(GenerateOneIsanRes(((ChangeddataModel)enumrator.Current).Secondary_ID)[0]);//Ric for Drop
                    dropRes.Add(oneDropRes);
                }

            }
            foreach (var temp in noFm1List)
            {
                res.Add(GenerateOneIsanRes(temp.ISIN));
            }
            if (res.Count > 1)
            {
                OperateExcel.WriteToTXT(filePath, res);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "GEDA File(ADD)");

            }
            if (dropRes.Count > 1)
            {
                OperateExcel.WriteToTXT(dropFilePath, dropRes);
                AddResult(Path.GetFileNameWithoutExtension(dropFilePath), dropFilePath, "GEDA File(DROP)");
            }
        }
        private void GenerateBcuRicRemove(string gedaDir)
        {
            List<List<string>> res = new List<List<string>>();
            List<string> head = new List<string>();
            string filePath = gedaDir + "\\KR_ELW_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_FM2_BCU_RIC_REMOVE.txt";
            head.Add("KSE_EQ_IPOELW_CHAIN");
            res.Add(head);
            foreach (var item in fmSecondList)
            {
                List<string> tempRes = new List<string>();
                tempRes.Add(item.RIC);
                res.Add(tempRes);
            }
            if (res.Count > 1)
            {
                OperateExcel.WriteToTXT(filePath, res);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "GEDA File(BCU_RIC_REMOVE)");
            }
        }
        private void GeneretaFmSecondPartGedaFiles()
        {
            string gedaDir = configObj.BulkFile;
            if (!Directory.Exists(gedaDir))
                Directory.CreateDirectory(gedaDir);
            GenerateFM2BcuUpdate(gedaDir);
            GenerateISINChangeRicCreation(gedaDir);
            GenerateBcuRicRemove(gedaDir);

            if (gedaFileHash.Count > 0)
                GenerateFM2MdnUpdate(gedaDir);
        }

        // Generate KR_ELW_ ddMMMyyyy_FM2_MD_N_update.txt 
        private void GenerateFM2MdnUpdate(string gedaDir)
        {
            List<List<string>> gedaRes = new List<List<string>>();
            string[] head = new string[] { "RIC\t", "DSPLY_NAME\t", "DSPLY_NMLL\t", "MATUR_DATE\t" };
            string filePath = gedaDir + "\\KR_ELW_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_FM2_MD_N_UPDATE.txt";

            gedaRes.Add(head.ToList());
            IEnumerator enumrator = gedaFileHash.Values.GetEnumerator();
            while (enumrator.MoveNext())
            {
                List<string> rowList = new List<string>();
                rowList.Add(((ChangeddataModel)enumrator.Current).RIC);
                rowList.Add("\t");
                rowList.Add(((ChangeddataModel)enumrator.Current).IDNDisplayName);
                rowList.Add("\t");
                rowList.Add(((ChangeddataModel)enumrator.Current).Warrant_Title);
                rowList.Add("\t");
                rowList.Add(((ChangeddataModel)enumrator.Current).MatDate);
                gedaRes.Add(rowList);

            }
            OperateExcel.WriteToTXT(filePath, gedaRes);
            AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "GEDA File(MD_N_UPDATE)");

        }

        private void GenerateQAChgFirstTradeDate(string NdaDir)
        {
            string filePath = string.Format("{0}\\RX{1}QAChgFirstTradeDate.csv", NdaDir, DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")));
            List<List<string>> res = new List<List<string>>();
            List<string> head = new List<string>();
            head.Add("RIC");
            head.Add("DERIVATIVES FIRST TRADING DAY");
            head.Add("ASSET COMMON NAME");
            head.Add("STRIKE PRICE");
            res.Add(head);
            foreach (var item in fmSecondList)
            {
                List<string> oneRes = new List<string>();
                oneRes.Add(item.RIC);
                oneRes.Add(DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                oneRes.Add(item.QACommonName);
                oneRes.Add(item.StrikePrice);
                res.Add(oneRes);
                oneRes = new List<string>();
                oneRes.Add(item.RIC.Replace(".KS", "F.KS"));
                oneRes.Add(DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                oneRes.Add(item.QACommonName);
                oneRes.Add(item.StrikePrice);
                res.Add(oneRes);

            }
            if (res.Count > 1)
            {
                OperateExcel.WriteToCSV(filePath, res);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "NDA QA File");
            }
        }
        private List<string> GenerateQAChgCommonNameHead(ChangeddataModel term)
        {
            List<string> head = new List<string>();
            head.Add("RIC");
            if (!string.IsNullOrEmpty(term.QACommonName))
            {
                if (!head.Contains("ASSET COMMON NAME"))
                {
                    head.Add("ASSET COMMON NAME");
                }
            }
            if (!string.IsNullOrEmpty(term.IDNDisplayName))
            {
                if (!head.Contains("ASSET SHORT NAME"))
                {
                    head.Add("ASSET SHORT NAME");
                    head.Add("CALL PUT OPTION");
                }
            }
            if (!string.IsNullOrEmpty(term.StrikePrice))
            {
                if (!head.Contains("STRIKE PRICE"))
                {
                    head.Add("STRIKE PRICE");
                }
            }
            if (!string.IsNullOrEmpty(term.MatDate))
            {
                if (!head.Contains("EXPIRY DATE"))
                {
                    head.Add("EXPIRY DATE");
                }
            }
            return head;
        }
        private Hashtable GenerateQAHash()
        {
            Hashtable QAresHash = new Hashtable();
            IEnumerator enumrator = ndaFileHash.Values.GetEnumerator();
            while (enumrator.MoveNext())
            {
                var term = (ChangeddataModel)enumrator.Current;
                List<string> head = null;
                List<List<string>> res = null;
                if (!QAresHash.Contains(term.ToString()))
                {
                    QAresHash.Add(term.ToString(), new List<List<string>>());
                    res = QAresHash[term.ToString()] as List<List<string>>;
                    head = GenerateQAChgCommonNameHead(term);
                    res.Add(head);
                }
                res = QAresHash[term.ToString()] as List<List<string>>;
                List<string> oneRes = new List<string>();
                head = res[0];
                oneRes.Add(term.RIC);

                if (head.Contains("ASSET COMMON NAME"))
                {
                    oneRes.Add(term.QACommonName);
                }
                if (head.Contains("ASSET SHORT NAME"))
                {
                    oneRes.Add(term.IDNDisplayName);
                    oneRes.Add(term.CPOption);
                }
                if (head.Contains("STRIKE PRICE"))
                {
                    oneRes.Add(term.StrikePrice);
                }
                if (head.Contains("EXPIRY DATE"))
                {
                    oneRes.Add(DateTime.Parse(term.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                }
                res.Add(oneRes);

                oneRes = new List<string>();
                oneRes.Add(term.RIC.Replace(".KS", "F.KS"));
                if (head.Contains("ASSET COMMON NAME"))
                {
                    oneRes.Add(term.QACommonName);
                }
                if (head.Contains("ASSET SHORT NAME"))
                {
                    oneRes.Add(term.IDNDisplayName);
                    oneRes.Add(term.CPOption);
                }
                if (head.Contains("STRIKE PRICE"))
                {
                    oneRes.Add(term.StrikePrice);
                }
                if (head.Contains("EXPIRY DATE"))
                {
                    oneRes.Add(DateTime.Parse(term.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                }
                res.Add(oneRes);

            }
            return QAresHash;
        }
        private void GenerateQAChgCommonName(string NdaDir)
        {
            Hashtable qaRes = GenerateQAHash();
            IEnumerator enumrator = qaRes.Values.GetEnumerator();
            int count = 0;
            while (enumrator.MoveNext())
            {
                var term = (List<List<string>>)enumrator.Current;
                string filePath = string.Format("{0}\\Type{1}_RX{2}QAChgQACommonName.csv", NdaDir, count++, DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")));
                OperateExcel.WriteToCSV(filePath, term);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "NDA QA File");
            }

        }



        private void GetReferenceFromDB()
        {
            string issueTableName = "ETI_Korea_Issuer";
            string underlyingTableName = "ETI_Korea_Underlying";
            System.Data.DataTable issueTable = ManagerBase.Select(issueTableName);
            foreach (DataRow issuer in issueTable.Rows)
            {
                ReferenceListIssueModel referenceIssuer = new ReferenceListIssueModel();
                referenceIssuer.RIC = Convert.ToString(issuer["IssuerCode2"]);
                referenceIssuer.NDAIssuerORGID = Convert.ToString(issuer["NDAIssuerOrgid"]);
                referenceIssuer.NDATCIssuerTitle = Convert.ToString(issuer["NDATCIssuerTitle"]);
                referenceIssuer.IssuerCompanyName = Convert.ToString(issuer["BodyGroupCommonName"]);
                refernceList.Add(referenceIssuer.RIC, referenceIssuer);
            }
            System.Data.DataTable underlyingTable = ManagerBase.Select(underlyingTableName);
            foreach (DataRow underlying in underlyingTable.Rows)
            {
                ReferenceListUnderLyingModel referenceUnderlying = new ReferenceListUnderLyingModel();
                referenceUnderlying.UnderlyingRIC = Convert.ToString(underlying["UnderlyingRIC"]);
                referenceUnderlying.NDATCUnderlyingTitle = Convert.ToString(underlying["NDATCUnderlyingTitle"]);
                referenceUnderlying.UnderlyingCompanyName = Convert.ToString(underlying["BodyGroupCommonName"]);
                refernceList.Add(referenceUnderlying.UnderlyingRIC, referenceUnderlying);
            }
        }

        private string GenerateIACommonName(ChangeddataModel term)
        {
            string res = string.Empty;
            WarrantTemplate wt = fmSecondHash[term.RIC] as WarrantTemplate;
            string underlyingRIC = ((WarrantTemplate)fmSecondHash[term.RIC]).BCASTREF;
            string Code = term.RIC.Substring(0, 2);
            string IssuerCompanyName = string.Empty;
            string UnderlyingCompanyName = string.Empty;
            string matDate = DateTime.Parse((string.IsNullOrEmpty(term.MatDate) ? wt.MatDate : term.MatDate)).ToString("ddMMMyy", new CultureInfo("en-US"));
            string call = wt.IDNDisplayName.Substring(wt.IDNDisplayName.Length - 1, 1);
            call = call.Equals("C") ? "Call" : "Put";
            if (refernceList[Code] != null)
            {
                IssuerCompanyName = ((ReferenceListIssueModel)refernceList[Code]).IssuerCompanyName;
            }
            else
            {
                string msg = string.Format("No IssuerCompanyName Could Be Found. RIC:{0}. IssuerCode:{2}.", term.RIC, Code);
                Logger.Log(msg, Logger.LogType.Error);
                IssuerCompanyName = "***";
                //throw (new Exception("No IssuerCompanyName Could Be Found"));
            }
            if (refernceList[underlyingRIC] != null)
            {
                UnderlyingCompanyName = ((ReferenceListUnderLyingModel)refernceList[underlyingRIC]).UnderlyingCompanyName;
            }
            else
            {
                string msg = string.Format("No UnderlyingCompanyName Could Be Found. RIC:{0}. UnderlyingRIC:{1}.", term.RIC, underlyingRIC);
                Logger.Log(msg, Logger.LogType.Error);
                UnderlyingCompanyName = "***";
                //throw (new Exception("No UnderlyingCompanyName Could Be Found"));
            }
            //if (refernceList[Code] != null)
            //{
            //    IssuerCompanyName = ((ReferenceListIssueModel)refernceList[Code]).IssuerCompanyName;
            //}
            //else
            //{
            //    throw (new Exception("No IssuerCompanyName Could Be Found"));
            //}
            //if (refernceList[underlyingRIC] != null)
            //{
            //    UnderlyingCompanyName = ((ReferenceListUnderLyingModel)refernceList[underlyingRIC]).UnderlyingCompanyName;
            //}
            //else
            //{
            //    throw (new Exception("No UnderlyingCompanyName Could Be Found"));
            //}
            res = IssuerCompanyName + " " + call + " " + (string.IsNullOrEmpty(term.StrikePrice) ? wt.StrikePrice : term.StrikePrice) + " " + "KRW" + " " + UnderlyingCompanyName + " " + DateTime.Parse(matDate).ToString("ddMMMyy", new CultureInfo("en-US"));
            return res;
        }

        private void GetTagAndPilcFromDb()
        {
            pilcHash = KoreaELWManager.SelectPILC();
            if (pilcHash == null || pilcHash.Count == 0)
            {
                string msg = "Can not get TAG and PILC information from database. Please check!";
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception(msg);
            }
        }

        private Hashtable ReadReferenceTableOfPlic()
        {
            Hashtable tableOfPlic = new Hashtable();
            string path = configObj.TagPilcFile;
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel applcation could not be created ,please check your office installation is corrected !!";
                Logger.Log(msg, Logger.LogType.Error);
                throw (new Exception(msg));
            }
            try
            {
                string fpath = configObj.TagPilcFile;  // "C:\\Korea_FM\\ELW_FM1\\" + filename;
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, fpath);
                Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet 1", wBook);
                if (wSheet == null)
                {
                    string msg = string.Format("worksheet{0} couldn't be found !!", "Issuer");
                    Logger.Log(msg, Logger.LogType.Error);
                    throw (new Exception(msg));
                }
                for (int i = 2; ; i++)
                {
                    string ric = ((Range)wSheet.Cells[i, 2]).Text.ToString();
                    if (!string.IsNullOrEmpty(ric))
                    {
                        string plic = ((Range)wSheet.Cells[i, 3]).Text.ToString();
                        tableOfPlic[ric] = plic;
                    }
                    else
                    {
                        break;
                    }
                }
                wBook.Save();
            }
            catch (Exception ex)
            {
                string msg = "Error found when read TAG and PILC.xls.";
                Logger.Log(msg, Logger.LogType.Error);
                Logger.Log(ex.Message, Logger.LogType.Error);
                Logger.Log(ex.StackTrace, Logger.LogType.Error);
                throw ex;
            }
            finally
            {
                excelApp.Dispose();
            }
            return tableOfPlic;
        }

        private void GenerateIAChgCommonNameCsv(string NdaDir)
        {
            string filePath = string.Format("{0}\\RX{1}IAChgIACommonName.csv", NdaDir, DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")));
            List<List<string>> res = new List<List<string>>();
            List<string> head = new List<string>();
            //head.Add("PILC");
            head.Add("ISIN");
            head.Add("ASSET COMMON NAME");
            res.Add(head);
            IEnumerator enumrator = ndaFileHash.Values.GetEnumerator();
            while (enumrator.MoveNext())
            {
                var term = (ChangeddataModel)enumrator.Current;
                List<string> oneRes = new List<string>();
                string isin = (fmSecondHash[term.RIC] as WarrantTemplate).ISIN;
                string pilc = string.Empty;
                if (pilcHash.Contains(term.RIC))
                {
                    pilc = pilcHash[term.RIC] as string;
                }
                //oneRes.Add(pilc);
                oneRes.Add(isin);
                string IACommonName = GenerateIACommonName(term);
                oneRes.Add(IACommonName);
                res.Add(oneRes);

            }
            if (res.Count > 1)
            {
                OperateExcel.WriteToCSV(filePath, res);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "NDA IA File");
            }

        }
        private void GenerateFmSecondPartNdaFiles()
        {
            string ndaDir = configObj.BulkFile;
            if (!Directory.Exists(ndaDir))
                Directory.CreateDirectory(ndaDir);
            GenerateQAChgFirstTradeDate(ndaDir);
            GenerateQAChgCommonName(ndaDir);
            GenerateIAChgCommonNameCsv(ndaDir);
        }

        private MailToSend CreateMail()
        {
            MailToSend mail = new MailToSend();
            StringBuilder mailbodyBuilder = new StringBuilder();
            string filename = "Korea FM for " + DateTime.Today.ToString("dd-MMM-yy").Replace("-", " ") + " (Afternoon).xls";
            string ipath = Path.Combine(configObj.FM, filename);
            mail.MailSubject = "KR FM [ELW ADD] wef " + (fmSecondList.Count > 0 ? fmSecondList[0].EffectiveDate : "");
            mailbodyBuilder.Append("ELW ADD:                                ");
            mailbodyBuilder.Append(fmSecondList.Count);
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("Effective Date:                 ");
            mailbodyBuilder.Append(fmSecondList.Count > 0 ? fmSecondList[0].EffectiveDate : "");
            mailbodyBuilder.Append("\r\n\r\n");
            if (fm2HasChangedValue)
            {
                mailbodyBuilder.Append("Please be noticed some value has been changed and highlighted in the attached file. Please refer to the attached file for the details. ");
            }
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
        private void GenerateFmSecondPartEmail()
        {
            string filename = "Korea FM for " + DateTime.Today.ToString("dd-MMM-yy").Replace("-", " ") + " (Afternoon).xls";
            string ipath = Path.Combine(configObj.FM, filename);
            AddResult(Path.GetFileNameWithoutExtension(ipath), ipath, "FM File");
        }
        #endregion

        #region Koba EMA file and NDA file and GEDA file

        private void GenerateKobaEmaGedaNdaAndCreatEmail()
        {
            GeneretaKobaGedaFile();
            GenerateKobaNdaFiles();
            GeneretaKobaEmaFile();
            GenerateKobaEmail();
        }


        private string GenerateKobaEmaFilePath()
        {
            string dir = ConfigureOperator.GetEmaFileSaveDir();
            string sendDir = dir + "\\" + DateTime.Now.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            if (!Directory.Exists(sendDir))
                Directory.CreateDirectory(sendDir);

            string path = sendDir + @"\KOBA_WRT_ADD_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_Korea.csv";
            return path;
        }
        private void GeneretaKobaEmaFile()
        {
            string filePath = GenerateKobaEmaFilePath();
            List<List<string>> emaRes = new List<List<string>>();
            string[] headStrs = new string[] {
                "Logical_Key","Secondary_ID","Secondary_ID_Type","Warrant_Title","Issuer_OrgId","Issue_Date","Country_Of_Issue","Governing_Country","Announcement_Date","Payment_Date","Underlying_Type","Clearinghouse1_OrgId","Clearinghouse2_OrgId","Clearinghouse3_OrgId","Guarantor","Guarantor_Type","Guarantee_Type","Incr_Exercise_Lot","Min_Exercise_Lot","Max_Exercise_Lot","Rt_Page_Range","Underwriter1_OrgId","Underwriter1_Role","Underwriter2_OrgId","Underwriter2_Role","Underwriter3_OrgId","Underwriter3_Role","Underwriter4_OrgId","Underwriter4_Role","Exercise_Style","Warrant_Type","Expiration_Date","Registered_Bearer_Code","Price_Display_Type","Private_Placement","Coverage_Type","Warrant_Status","Status_Date","Redemption_Method","Issue_Quantity","Issue_Price","Issue_Currency","Issue_Price_Type","Issue_Spot_Price","Issue_Spot_Currency","Issue_Spot_FX_Rate","Issue_Delta","Issue_Elasticity","Issue_Gearing","Issue_Premium","Issue_Premium_PA","Denominated_Amount","Exercise_Begin_Date","Exercise_End_Date","Offset_Number","Period_Number","Offset_Frequency","Offset_Calendar","Period_Calendar","Period_Frequency","RAF_Event_Type","Exercise_Price","Exercise_Price_Type","Warrants_Per_Underlying","Underlying_FX_Rate","Underlying_RIC","Underlying_Item_Quantity","Units","Cash_Currency","Delivery_Type","Settlement_Type","Settlement_Currency","Underlying_Group","Country1_Code","Coverage1_Type","Country2_Code","Coverage2_Type","Country3_Code","Coverage3_Type","Country4_Code","Coverage4_Type","Country5_Code","Coverage5_Type","Note1_Type","Note1","Note2_Type","Note2","Note3_Type","Note3","Note4_Type","Note4","Note5_Type","Note5","Note6_Type","Note6","Exotic1_Parameter","Exotic1_Value","Exotic1_Begin_Date","Exotic1_End_Date","Exotic2_Parameter","Exotic2_Value","Exotic2_Begin_Date","Exotic2_End_Date","Exotic3_Parameter","Exotic3_Value","Exotic3_Begin_Date","Exotic3_End_Date","Exotic4_Parameter","Exotic4_Value","Exotic4_Begin_Date","Exotic4_End_Date","Exotic5_Parameter","Exotic5_Value","Exotic5_Begin_Date","Exotic5_End_Date","Exotic6_Parameter","Exotic6_Value","Exotic6_Begin_Date","Exotic6_End_Date","Event_Type1","Event_Period_Number1","Event_Calendar_Type1","Event_Frequency1","Event_Type2","Event_Period_Number2","Event_Calendar_Type2","Event_Frequency2","Exchange_Code1","Incr_Trade_Lot1","Min_Trade_Lot1","Min_Trade_Amount1","Exchange_Code2","Incr_Trade_Lot2","Min_Trade_Lot2","Min_Trade_Amount2","Exchange_Code3","Incr_Trade_Lot3","Min_Trade_Lot3","Min_Trade_Amount3","Exchange_Code4","Incr_Trade_Lot4","Min_Trade_Lot4","Min_Trade_Amount4","Attached_To_Id","Attached_To_Id_Type","Attached_Quantity","Attached_Code","Detachable_Date","Bond_Exercise","Bond_Price_Percentage"
            };
            emaRes.Add(headStrs.ToList());
            int count = 1;

            foreach (var term in kobaList)
            {
                emaRes.Add((new string[150]).ToList());
                emaRes[count][0] = (count).ToString();
                emaRes[count][1] = term.ISIN;
                emaRes[count][2] = "ISIN";
                emaRes[count][5] = DateTime.Parse(term.IssueDate).ToString("dd/MM/yyyy", new CultureInfo("en-US"));
                emaRes[count][6] = "KOR";
                emaRes[count][7] = "KOR";
                if (term.QACommonName.Substring(term.QACommonName.Length - 2, 2).Equals("IW"))
                    emaRes[count][10] = "INDEX";
                else
                    emaRes[count][10] = "STOCK";
                emaRes[count][17] = "10";
                emaRes[count][18] = "10";
                emaRes[count][29] = "E";
                if (term.IDNDisplayName.Substring(term.IDNDisplayName.Length - 1, 1).Equals("C"))
                    emaRes[count][30] = "Down and Out Call";
                else
                    emaRes[count][30] = "Up and Out Put";
                emaRes[count][31] = DateTime.Parse(term.MatDate).ToString("dd/MM/yyyy", new CultureInfo("en-US"));
                emaRes[count][32] = "R";
                emaRes[count][33] = "D";
                emaRes[count][39] = term.QuanityofWarrants;
                emaRes[count][40] = term.IssuePrice;
                emaRes[count][41] = "KRW";
                emaRes[count][42] = "A";

                emaRes[count][51] = "10";
                emaRes[count][52] = DateTime.Parse(term.MatDate).ToString("dd/MM/yyyy", new CultureInfo("en-US"));
                emaRes[count][53] = DateTime.Parse(term.MatDate).ToString("dd/MM/yyyy", new CultureInfo("en-US"));


                emaRes[count][61] = term.StrikePrice;//Exercise_Price need to be confirm 

                emaRes[count][62] = "A";
                emaRes[count][63] = 1.0 / Convert.ToDouble(term.ConversionRatio) + "";
                emaRes[count][65] = term.BCASTREF;
                emaRes[count][66] = "1";
                if (term.QACommonName.Substring(term.QACommonName.Length - 2, 2).Equals("IW"))
                    emaRes[count][67] = "idx";
                else
                    emaRes[count][67] = "sht";
                if (term.QACommonName.Substring(term.QACommonName.Length - 2, 2).Equals("IW"))
                    emaRes[count][69] = "I";
                else
                    emaRes[count][69] = "S";
                emaRes[count][70] = "C";
                emaRes[count][71] = "KRW";
                emaRes[count][83] = "T";
                emaRes[count][84] = "Last Trading Day is " + DateTime.Parse(term.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")) + ".";
                emaRes[count][85] = "S";
                // emaRes[count][86] = "Note2";//Note2 need to be confirm 
                emaRes[count][95] = "BAR";
                emaRes[count][96] = term.KnockOutPrice;
                if (term.RIC.Split(".".ToArray())[1].Equals("KS"))
                    emaRes[count][127] = "KSC";
                else
                    emaRes[count][127] = "KOE";
                emaRes[count][128] = "10";
                emaRes[count][129] = "10";

                string code = term.RIC.Substring(0, 2);
                var issueRefer = refernceList[code] as ReferenceListIssueModel;
                var uderLyingRefer = refernceList[term.BCASTREF] as ReferenceListUnderLyingModel;
                emaRes[count][3] = issueRefer.NDATCIssuerTitle + "/" + " " + uderLyingRefer.NDATCUnderlyingTitle + " " + (emaRes[count][10] == "STOCK" ? "SHS" : emaRes[count][10]) + " " + emaRes[count][30].ToUpper() + " " + "WTS" + " " + DateTime.Parse(term.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")).ToUpper();
                emaRes[count][4] = issueRefer.NDAIssuerORGID;
                count++;
            }
            if (emaRes.Count > 1)
            {
                OperateExcel.WriteToCSV(filePath, emaRes);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "EMA File(KOBA)");
            }
        }


        private List<string> GenerateKobaOneIsanRes(string isan)
        {
            List<string> item = new List<string>();
            foreach (var temp in kobaList)
            {
                if (temp.ISIN.Equals(isan))
                {
                    item.Add(temp.RIC);
                    item.Add("\t");
                    item.Add(temp.IDNDisplayName);
                    item.Add("\t");
                    item.Add(temp.RIC);
                    item.Add("\t");
                    item.Add(temp.Ticker);
                    item.Add("\t");
                    item.Add(temp.ISIN);
                    item.Add("\t");
                    item.Add(temp.BCASTREF);
                    item.Add("\t");
                    item.Add("KSE_EQB_ELW");
                    item.Add("\t");
                    item.Add(temp.Ticker);
                    item.Add("\t");
                    item.Add(temp.ISIN);
                    item.Add("\t");
                    item.Add("J" + temp.Ticker);
                    item.Add("\t");
                    item.Add(temp.KoreaWarrantName);
                    item.Add("\t");
                    item.Add(DateTime.Parse(temp.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                    item.Add("\t");
                    item.Add("KSE_EQB_ELW");
                }


            }
            return item;
        }
        private void GenerateKobaCreation(string gedaDir)
        {
            List<List<string>> res = new List<List<string>>();
            string filePath = gedaDir + "\\KR_KOBA_ADD_" + DateTime.Now.ToString("ddMMMyyyy", new CultureInfo("en-US")) + "_FM_CREATE.txt";
            string[] headStrs = new string[] { "SYMBOL", "\t", "DSPLY_NAME", "\t", "RIC", "\t", "OFFCL_CODE", "\t", "EX_SYMBOL", "\t", "BCAST_REF", "\t", "EXL_NAME", "\t", "#INSTMOD_TDN_SYMBOL", "\t", "#INSTMOD_#ISIN", "\t", "#INSTMOD_MNEMONIC", "\t", "DSPLY_NMLL", "\t", "MATUR_DATE", "\t", "BCU" };
            res.Add(headStrs.ToList());
            foreach (var term in kobaList)
            {
                res.Add(GenerateKobaOneIsanRes(term.ISIN));
            }
            if (res.Count > 1)
            {
                OperateExcel.WriteToTXT(filePath, res);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "GEDA File(KOBA)");
            }
        }
        private void GeneretaKobaGedaFile()
        {
            string gedaDir = configObj.BulkFile;
            if (!Directory.Exists(gedaDir))
                Directory.CreateDirectory(gedaDir);
            // generateKoba_BCU_UPDATE(gedaDir);
            GenerateKobaCreation(gedaDir);
        }


        private void GenerateKobaIA(string ndaDir)
        {
            string filePath = string.Format("{0}\\KOBA_{1}IAChg.csv", ndaDir, DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")));
            List<List<string>> res = new List<List<string>>();
            string[] head = new string[]
            {
                "ISIN",
                "TYPE",
                "CATEGORY",
                "WARRANT ISSUER",
                "RCS ASSET CLASS",
                "WARRANT ISSUE QUANTITY"
            };
            res.Add(head.ToList());
            foreach (var term in kobaList)
            {
                List<string> oneRes = new List<string>();
                oneRes.Add(term.ISIN);
                oneRes.Add("DERIVATIVE");
                oneRes.Add("EIW");
                string Code = term.RIC.Substring(0, 2);
                string orgId = (refernceList[Code] as ReferenceListIssueModel).NDAIssuerORGID;
                oneRes.Add(orgId);
                oneRes.Add("FXKNOCKOUT");
                oneRes.Add(term.QuanityofWarrants);
                res.Add(oneRes);
            }

            OperateExcel.WriteToCSV(filePath, res);
            AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "NDA IA File(KOBA)");
        }
        private void GenerateKobaQA(string ndaDir)
        {
            string filePath = string.Format("{0}\\KOBA_{1}QAChg.csv", ndaDir, DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")));
            List<List<string>> res = new List<List<string>>();
            string[] head = new string[]
            {
                    "RIC",
                    "TAG",
                    "TYPE",
                    "CATEGORY",
                    "EXCHANGE",
                    "CURRENCY",
                    "ASSET COMMON NAME",
                    "ASSET SHORT NAME",
                    "CALL PUT OPTION",
                    "STRIKE PRICE",
                    "WARRANT ISSUE PRICE",
                    "ROUND LOT SIZE",
                    "EXPIRY DATE",
                    "TICKER SYMBOL",
                    "BASE ASSET"
            };

            res.Add(head.ToList());
            foreach (var term in kobaList)
            {
                List<string> oneRes = new List<string>();
                oneRes.Add(term.RIC);
                oneRes.Add("46429");
                oneRes.Add("DERIVATIVE");
                oneRes.Add("EIW");
                if (term.RIC.Split(".".ToArray())[1].Equals("KS"))
                    oneRes.Add("KSC");
                else
                    oneRes.Add("KOE");
                oneRes.Add("KRW");
                oneRes.Add(term.QACommonName);
                oneRes.Add(term.IDNDisplayName);
                oneRes.Add(term.IDNDisplayName.Substring(term.IDNDisplayName.Length - 1, 1).Equals("C") ? "CALL" : "PUT");
                oneRes.Add(term.StrikePrice);
                oneRes.Add(term.IssuePrice);
                oneRes.Add("10");
                oneRes.Add(DateTime.Parse(term.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                oneRes.Add(term.Ticker);
                oneRes.Add("ISIN:" + term.ISIN);
                res.Add(oneRes);


                oneRes = new List<string>();
                oneRes.Add(term.RIC.Replace(".KS", "F.KS"));
                oneRes.Add("44398");
                oneRes.Add("DERIVATIVE");
                oneRes.Add("EIW");
                if (term.RIC.Split(".".ToArray())[1].Equals("KS"))
                    oneRes.Add("KSC");
                else
                    oneRes.Add("KOE");
                oneRes.Add("KRW");
                oneRes.Add(term.QACommonName);
                oneRes.Add(term.IDNDisplayName);
                oneRes.Add(term.IDNDisplayName.Substring(term.IDNDisplayName.Length - 1, 1).Equals("C") ? "CALL" : "PUT");
                oneRes.Add(term.StrikePrice);
                oneRes.Add(term.IssuePrice);
                oneRes.Add("10");
                oneRes.Add(DateTime.Parse(term.MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US")));
                oneRes.Add(term.Ticker);
                oneRes.Add("ISIN:" + term.ISIN);
                res.Add(oneRes);
            }
            OperateExcel.WriteToCSV(filePath, res);
            AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "NDA QA File(KOBA)");
        }
        private void GenerateKobaNdaFiles()
        {
            string ndaDir = configObj.BulkFile;
            if (!Directory.Exists(ndaDir))
                Directory.CreateDirectory(ndaDir);
            GenerateKobaIA(ndaDir);
            GenerateKobaQA(ndaDir);
        }

        private void GenerateNdaTickLotFile()
        {
            try
            {
                List<string> tickTitle = new List<string>(){"RIC", "TICK NOT APPLICABLE", "TICK LADDER NAME", 
                                                      "TICK EFFECTIVE FROM", "TICK EFFECTIVE TO", "TICK PRICE INDICATOR" };
                List<string> lotTitle = new List<string>(){"RIC", "LOT NOT APPLICABLE", "LOT LADDER NAME", 
                                                      "LOT EFFECTIVE FROM", "LOT EFFECTIVE TO", "LOT PRICE INDICATOR" };
                string today = DateTime.Today.ToString("yyyyMMdd", new CultureInfo("en-US"));
                string filePathTick = Path.Combine(configObj.BulkFile, "TickAdd_ELW_KOBA_" + today + ".csv");
                string filePathLot = Path.Combine(configObj.BulkFile, "LotAdd_ELW_KOBA_" + today + ".csv");
                List<List<string>> tickContent = new List<List<string>>();
                List<List<string>> lotContent = new List<List<string>>();
                List<WarrantTemplate> mergeElwKoba = kobaList;
                mergeElwKoba.AddRange(fmSecondList);
                foreach (WarrantTemplate item in mergeElwKoba)
                {
                    List<string> tickRecord = new List<string>();
                    List<string> lotRecord = new List<string>();
                    tickRecord.Add(item.RIC);
                    lotRecord.Add(item.RIC);
                    tickRecord.Add("N");
                    lotRecord.Add("N");
                    tickRecord.Add("TICK_LADDER_<5>");
                    lotRecord.Add("LOT_LADDER_EQTY_<10>");
                    string effectiveDate = DateTime.Parse(item.EffectiveDate).ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US"));
                    tickRecord.Add(effectiveDate);
                    lotRecord.Add(effectiveDate);
                    tickRecord.Add("");
                    lotRecord.Add("");
                    tickRecord.Add("ORDER");
                    lotRecord.Add("CLOSE");
                    tickContent.Add(tickRecord);
                    lotContent.Add(lotRecord);
                }
                FileUtil.WriteOutputFile(filePathTick, tickContent, tickTitle, WriteMode.Overwrite);
                FileUtil.WriteOutputFile(filePathLot, lotContent, lotTitle, WriteMode.Overwrite);
                AddResult(Path.GetFileName(filePathTick), filePathTick, "NDA Tick Add File");
                AddResult(Path.GetFileName(filePathLot), filePathLot, "NDA Lot Add File");

                Logger.Log("Generate NDA tick and lot add files successfully.", Logger.LogType.Info);
            }
            catch (Exception ex)
            {
                Logger.Log("Error found in generating NDA Tick and Lot file. \r\n" + ex.ToString(), Logger.LogType.Error);
            }
        }


        private MailToSend CreatKobaMail()
        {
            MailToSend mail = new MailToSend();
            StringBuilder mailbodyBuilder = new StringBuilder();
            string filename = "Korea FM (KOBA Add)" + DateTime.Today.ToString("dd-MMM-yy").Replace("-", " ") + " (Afternoon).xls";
            string ipath = Path.Combine(configObj.FM, filename);
            mail.MailSubject = "KR FM [KOBA ADD] wef " + DateTime.Today.ToString("dd-MMM-yy");
            mailbodyBuilder.Append("KOBA ADD:                                ");
            mailbodyBuilder.Append(kobaList.Count);
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("Effective Date:                 ");

            mailbodyBuilder.Append(kobaList.Count > 0 ? kobaList[0].EffectiveDate : "");
            mailbodyBuilder.Append("\r\n\r\n");
            mail.MailBody = mailbodyBuilder.ToString();

            mail.MailBody += "\r\n\r\n\r\n\r\n";
            foreach (var term in configObj.Signature)
            {
                mail.MailBody += term + "\r\n";
            }
            mail.ToReceiverList.AddRange(configObj.MailTo);
            mail.CCReceiverList.AddRange(configObj.MailCC);
            mail.AttachFileList.Add(ipath);
            return mail;

        }
        private void GenerateKobaEmail()
        {
            string filename = "Korea FM (KOBA Add)" + DateTime.Today.ToString("dd-MMM-yy").Replace("-", " ") + " (Afternoon).xls";
            string fpath = Path.Combine(configObj.FM, filename);
            AddResult(Path.GetFileNameWithoutExtension(fpath), fpath, "FM File(KOBA)");
        }
        #endregion

        #region No fm1 QA and IA

        //NDA IA File
        private void GenerateNDAIAFile(string ndaDir)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
            }

            try
            {
                string filePath = string.Format("{0}\\FM2_NO_FM1_{1}IAChg.csv", ndaDir, DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")));
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet1", wBook);

                //wSheet.Name = "kr" + DateTime.Today.ToString("yyyy-MM-dd").Replace("-", "") + "IAWntAdd";

                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 17;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 12;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 20;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 25;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 35;
                ((Range)wSheet.Columns["A:E", System.Type.Missing]).Font.Name = "Arial";
                ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                }
                wSheet.Cells[1, 1] = "ISIN";
                wSheet.Cells[1, 2] = "TYPE";
                wSheet.Cells[1, 3] = "CATEGORY";
                wSheet.Cells[1, 4] = "WARRANT ISSUER";
                wSheet.Cells[1, 5] = "RCS ASSET CLASS";
                wSheet.Cells[1, 6] = "WARRANT ISSUE QUANTITY";

                int startLine = 2;
                LoopPrintNDAIAFile(wSheet, startLine, noFm1List);

                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.SaveAs(wBook.FullName, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "NDA IA File");
            }
            catch (Exception ex)
            {
                string msg = "Error found in NDA IA file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private void LoopPrintNDAIAFile(Worksheet wSheet, int startLine, List<WarrantTemplate> lists)
        {
            try
            {
                for (var i = 0; i < lists.Count; i++)
                {
                    wSheet.Cells[startLine, 1] = lists[i].ISIN;
                    wSheet.Cells[startLine, 2] = "DERIVATIVE";
                    wSheet.Cells[startLine, 3] = "EIW";
                    KoreaIssuerInfo issuer = KoreaIssuerManager.SelectIssuer(lists[i].IssuerKoreaName);
                    if (issuer != null)
                        wSheet.Cells[startLine, 4] = issuer.NDAIssuerOrgid;
                    else
                        wSheet.Cells[startLine, 4] = "出现异常信息... ...";

                    wSheet.Cells[startLine, 5] = "TRAD";
                    wSheet.Cells[startLine, 6] = lists[i].QuanityofWarrants;
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in LoopPrintNDAIAFile()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private void GenerateNDAQAFile(string ndaDir)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference are correct !!!";
                Logger.Log(msg, Logger.LogType.Error);
            }

            try
            {
                string filePath = string.Format("{0}\\FM2_NO_FM1_{1}QAChg.csv", ndaDir, DateTime.Now.ToString("yyyyMMdd", new CultureInfo("en-US")));
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, filePath);
                Worksheet wSheet = ExcelUtil.GetWorksheet("Sheet1", wBook);
                if (wSheet == null)
                {
                    string msg = "Excel Worksheet could not be started. Check that your office installation and project reference are correct !!!";
                    Logger.Log(msg, Logger.LogType.Error);
                }
                //wSheet.Name = "kr" + DateTime.Today.ToString("yyyy-MM-dd").Replace("-", "") + "QAWntAdd";

                GenerateQAExcelTitle(wSheet);

                int startLine = 2;
                LoopPrintNDAQAFile(wSheet, startLine, noFm1List, "QA");
                LoopPrintNDAQAFile(wSheet, startLine + noFm1List.Count, noFm1List, "QAF");
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.SaveAs(wBook.FullName, XlFileFormat.xlCSV, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                AddResult(Path.GetFileNameWithoutExtension(filePath), filePath, "NDA QA File");
            }
            catch (Exception ex)
            {
                string msg = "Error found in NDA QA file :" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private static void GenerateQAExcelTitle(Worksheet wSheet)
        {
            ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 13;
            ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 10;
            ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 12;
            ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 12;
            ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 12;
            ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 45;
            ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 25;
            ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["K", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["L", System.Type.Missing]).ColumnWidth = 20;
            ((Range)wSheet.Columns["M", System.Type.Missing]).ColumnWidth = 17;
            ((Range)wSheet.Columns["N", System.Type.Missing]).ColumnWidth = 13;
            ((Range)wSheet.Columns["O", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["P", System.Type.Missing]).ColumnWidth = 15;
            ((Range)wSheet.Columns["A:P", System.Type.Missing]).Font.Name = "Arial";
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            ((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

            wSheet.Cells[1, 1] = "RIC";
            wSheet.Cells[1, 2] = "TAG";
            wSheet.Cells[1, 3] = "TYPE";
            wSheet.Cells[1, 4] = "CATEGORY";
            wSheet.Cells[1, 5] = "EXCHANGE";
            wSheet.Cells[1, 6] = "CURRENCY";
            wSheet.Cells[1, 7] = "ASSET COMMON NAME";
            wSheet.Cells[1, 8] = "ASSET SHORT NAME";
            wSheet.Cells[1, 9] = "CALL PUT OPTION";
            wSheet.Cells[1, 10] = "STRIKE PRICE";
            wSheet.Cells[1, 11] = "WARRANT ISSUE PRICE";
            wSheet.Cells[1, 12] = "ROUND LOT SIZE";
            wSheet.Cells[1, 13] = "EXPIRY DATE";
            wSheet.Cells[1, 14] = "TICKER SYMBOL";
            wSheet.Cells[1, 15] = "BASE ASSET";
        }

        private void LoopPrintNDAQAFile(Worksheet wSheet, int startLine, List<WarrantTemplate> lists, string type)
        {
            try
            {
                for (var i = 0; i < lists.Count; i++)
                {
                    wSheet.Cells[startLine, 1] = lists[i].RIC;
                    if (type.Equals("QA"))
                    {
                        wSheet.Cells[startLine, 1] = lists[i].RIC;
                        wSheet.Cells[startLine, 2] = "46429";
                    }
                    else if (type.Equals("QAF"))
                    {
                        string ric = lists[i].RIC.Insert(6, "F");
                        wSheet.Cells[startLine, 1] = ric;
                        wSheet.Cells[startLine, 2] = "44398";
                    }
                    wSheet.Cells[startLine, 3] = "DERIVATIVE";
                    wSheet.Cells[startLine, 4] = "EIW";
                    wSheet.Cells[startLine, 5] = "KSC";
                    wSheet.Cells[startLine, 6] = "KRW";
                    wSheet.Cells[startLine, 7] = lists[i].QACommonName;
                    wSheet.Cells[startLine, 8] = lists[i].IDNDisplayName;
                    wSheet.Cells[startLine, 9] = lists[i].CallOrPut;
                    wSheet.Cells[startLine, 10] = lists[i].StrikePrice;
                    wSheet.Cells[startLine, 11] = lists[i].IssuePrice;
                    wSheet.Cells[startLine, 12] = "10";
                    ((Range)wSheet.Cells[startLine, 13]).NumberFormat = "@";
                    wSheet.Cells[startLine, 13] = Convert.ToDateTime(lists[i].MatDate).ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
                    wSheet.Cells[startLine, 14] = lists[i].Ticker;
                    wSheet.Cells[startLine, 15] = "ISIN:" + lists[i].ISIN;
                    startLine++;
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in LoopPrintNDAQAFile()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        #endregion

        #region Database Operation
        private string GetELWFmOneFromDb()
        {
            string condition = string.Empty;
            foreach (var item in fmSecondList)
            {
                condition += string.Format("'{0}',", item.RIC);
            }
            condition = condition.Substring(0, condition.Length - 1);

            fmSecondHash = KoreaELWManager.SelectELWFM1(condition);
            //if (fmSecondHash == null)
            //{
            //    string msg = "Can not access to ELW Master database! Please check!";
            //    throw new Exception(msg);
            //}
            return condition;
        }

        private void AppendDataToKOBADb()
        {
            int row = KoreaELWManager.InsertKOBA(kobaList);
            string msge = string.Format("Updated {0} KOBA record(s) in database.", row);
            Logger.Log(msge);
        }

        private void AppendDataToELWDb(string rics)
        {
            int row = KoreaELWManager.InsertELW(fmSecondList);
            string msge = string.Format("Updated {0} ELW FM2 record(s) in database.", row);
            Logger.Log(msge);

            row = KoreaELWManager.DeleteELWFM1(rics);
            msge = string.Format("Deleted {0} ELW FM1 record(s) in database.", row);
            Logger.Log(msge);
        }
        #endregion

        #region New Underlying Logic
        /// <summary>
        /// Generate three GEDA files for new underlying
        /// e.g.
        ///	1.File name: UNDERLYING_CHAIN_UPLOAD_0#028150W.KQ.txt
        ///	2.File name: CHAIN_CONST_ADD_ 0#028150W.KQ.txt
        ///	3.File name: SUPERCHAIN_CONST_ADD_ 0#UNDLY.KQ.txt or KS
        /// </summary>
        private void GenerateNewUnderlyingFiles()
        {
            //From config
            string filePath = configObj.GEDA_NewUnderlying;
            FileUtil.CreateDirectory(filePath);
            AddResult("New Underlying Folder", configObj.GEDA_NewUnderlying, "New Underlying Folder");
            bool superKS = false;
            bool superKQ = false;
            string superKSFileName = "SUPERCHAIN_CONST_ADD_0#UNDLY.KS.txt";
            string superKQFileName = "SUPERCHAIN_CONST_ADD_0#UNDLY.KQ.txt";
            foreach (KoreaUnderlyingInfo newItem in newUnderLying)
            {
                string[] ricSpilt = newItem.UnderlyingRIC.Split('.');
                string modifyRic = newItem.UnderlyingRIC.Split('.')[0] + "W." + newItem.UnderlyingRIC.Split('.')[1];
                string fileName = Path.Combine(filePath, "UNDERLYING_CHAIN_UPLOAD_0#" + modifyRic + ".txt");

                string ricChainToFill = "KSE_STOCK_" + ricSpilt[0] + "W_CHAIN";
                string ksOrkqStrToFill = "STQS6\tKSE";
                string exchangeToFill = "KO";
                string mrnToFill = "287";
                string rdnExchidToFill = "156";
                string rdnExchd2ToFill = "156";
                string prodPermToFill = "3104";
                string superChain = "KSE_EQ_UNDLY_CHAIN";
                string ricToFill = ricSpilt[0];
                string ksOrkqTE = "STQS6";
                if (ricSpilt[1] == "KQ")
                {
                    ricChainToFill = "KOSDAQ_STOCK_" + ricSpilt[0] + "W_CHAIN";
                    ksOrkqStrToFill = "STQSR\tKOSDAQ";
                    exchangeToFill = "KQ";
                    mrnToFill = "144";
                    rdnExchidToFill = "0";
                    rdnExchd2ToFill = "380";
                    prodPermToFill = "4084";
                    superChain = "KOSDAQ_EQ_UNDLY_CHAIN";
                    ksOrkqTE = "STQSR";
                    superKQ = true;
                }
                else
                {
                    superKS = true;
                }

                string chainUploadData = "FILENAME\t" + ricChainToFill + "\t" + ksOrkqStrToFill + "\r\n" +
                                       "CHAIN_RIC\t0#" + modifyRic + "\r\n" +
                                       "LINK_ROOT\t" + modifyRic + "\r\n" +
                                       "LOCAL_LANGUAGE\tLL_KOREAN\r\n" +
                                       "DISPLAY_NAME\t" + newItem.NDATCUnderlyingTitle + "\r\n" +
                                       "DISPLAY_NMLL\t" + newItem.CompanyName + "\r\n" +
                                       "RDNDISPLAY\t244\r\n" +
                                       "EXCHANGE\t" + exchangeToFill + "\r\n" +
                                       "ISSUE\tLINK\r\n" +
                                       "SEND\tIMSOUT\r\n" +
                                       "MRN\t" + mrnToFill + "\r\n" +
                                       "MRV\t0\r\n" +
                                       "INH_RANK\tTRUE\r\n" +
                                       "TPL_VER\t2.02\r\n" +
                                       "RDN_EXCHID\t" + rdnExchidToFill + "\r\n" +
                                       "RDN_EXCHD2\t" + rdnExchd2ToFill + "\r\n" +
                                       "TPL_NUM\t85\r\n" +
                                       "RECORDTYPE\t104\r\n" +
                                       "PROD_PERM\t" + prodPermToFill + "\r\n" +
                                       "CURRENCY\t410\r\n" +
                                       "RECKEY\tRK_SYMBOL\r\n" +
                                       "EMAIL_GROUP_ID\tgrpcntmarketdatastaff@thomsonreuters.com\r\nEND";
                string chainTitle = "RIC\tBCU\r\n";
                string chainConst = chainTitle + newItem.UnderlyingRIC + "\t" + ricChainToFill;
                string superChainData = "0#" + modifyRic + "\t" + superChain + "\r\n";

                File.WriteAllText(fileName, chainUploadData);
                AddResult(Path.GetFileName(fileName), fileName, Path.GetFileNameWithoutExtension(fileName));
                fileName = Path.Combine(filePath, "CHAIN_CONST_ADD_0#" + modifyRic + ".txt");
                File.WriteAllText(fileName, chainConst);
                AddResult(Path.GetFileName(fileName), fileName, Path.GetFileNameWithoutExtension(fileName));
                fileName = Path.Combine(filePath, "SUPERCHAIN_CONST_ADD_0#UNDLY." + ricSpilt[1] + ".txt");
                if (!File.Exists(fileName))
                {
                    superChainData = chainTitle + superChainData;
                }
                File.AppendAllText(fileName, superChainData);
                SendNewUnderlyingMail(ksOrkqTE, ricChainToFill, rdnExchd2ToFill, modifyRic);
            }
            if (superKQ)
            {
                AddResult(superKQFileName, Path.Combine(filePath, superKQFileName), superKQFileName);
            }
            if (superKS)
            {
                AddResult(superKSFileName, Path.Combine(filePath, superKSFileName), superKQFileName);
            }
        }

        /// <summary>
        /// Generate two new underlying mails.
        /// 1. LXL update
        /// 2. Pls add Delay Chain  
        /// </summary>
        /// <param name="ksOrkqTE">TE mark for KQ or KS</param>
        /// <param name="chain">part of mail content</param>
        /// <param name="exchangeID">exchange ID for a new underlying</param>
        /// <param name="modifyRic">ric</param>
        private void SendNewUnderlyingMail(string ksOrkqTE, string chain, string exchangeID, string modifyRic)
        {
            string filePath = Path.Combine(configObj.GEDA_NewUnderlying, "Mail");
            FileUtil.CreateDirectory(filePath);

            string mailBody = "<p>Hi BJG Central DBA,</p><p>Below chain has been created and TE success in " + ksOrkqTE
                               + ".</p><p>Please update the following:</p><p>BCU to be added:</p>"
                               + "<table style=\"border-collapse:collapse;border:none; font-family: 'Arial';font-size: 12px;\"><tr><td style=\"border: solid #000 1px;\" >BCU</td><td style=\"border: solid #000 1px;\">Action Date</td></tr>"
                             + "<tr><td style=\"border: solid #000 1px;\">" + chain + "</td><td style=\"border: solid #000 1px;\">ASAP</td></tr></table>";

            CreatMailAndSave("LXL update", mailBody, Path.Combine(filePath, "LXL update" + modifyRic + ".msg"));
            mailBody = "<p>Hi BJG Central DBA,</p><p>Please help to build below delay chain under delay PE 5229.</p><p>The Exchange ID is " + exchangeID + ".</p><p>0#" + modifyRic + "</p>";
            CreatMailAndSave("Pls add Delay Chain", mailBody, Path.Combine(filePath, "Pls add Delay Chain" + modifyRic + ".msg"));
        }

        /// <summary>
        /// Create mail and save it to local disk. Users can check the mail content.
        /// </summary>
        /// <param name="mailSubject">mail subject</param>
        /// <param name="mailBody">mail body</param>
        /// <param name="filePath">path to save</param>
        private void CreatMailAndSave(string mailSubject, string mailBody, string filePath)
        {
            MailToSend mail = new MailToSend();
            mail.ToReceiverList.AddRange(configObj.NewUnderlying_MailTo);
            mail.CCReceiverList.AddRange(configObj.NewUnderlying_MailCC);
            mail.MailSubject = mailSubject;
            string signature = string.Join("<br>", configObj.NewUnderlying_Signature.ToArray());
            //mail.MailBody += signature;

            mail.MailHtmlBody = "<div style=\"font-family: 'Arial';font-size: 10pt;\">" + mailBody;
            mail.MailHtmlBody += "<br>" + signature + "</div>";
            string err = string.Empty;
            using (OutlookApp outlookApp = new OutlookApp())
            {
                OutlookUtil.SaveMail(outlookApp, mail, out err, filePath);
            }
            AddResult(mail.MailSubject + ".msg", filePath, mail.MailSubject);
        }
        #endregion
    }
}
