using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Text;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.China
{
    [ConfigStoredInDB]
    public class ChinaIndexQCConfig
    {
        [StoreInDB]
        [Category("Uri")]
        [DisplayName("SSE Base Uri")]
        public string SseBaseUri { get; set; }

        [StoreInDB]
        [Category("Uri")]
        [DisplayName("CSI Base Uri")]
        public string CsiBaseUri { get; set; }

        [StoreInDB]
        [Category("Uri")]
        [DisplayName("SZSE Base Uri")]
        public string SzseBaseUri { get; set; }

        [StoreInDB]
        [DisplayName("Target file directory")]
        public string TargetFileDir { get; set; }

        [StoreInDB]
        [DisplayName("Chain number per sheet")]
        public int ChainNumPerSheet { get; set; }
    }

    public class ChinaIndex
    {
        public string ChineseName { get; set; }
        public string Chain { get; set; }
        public List<string> RicList { get; set; }
        public string OfficeCode { get; set; }
        public string SourceUrl { get; set; }

        public ChinaIndex()
        {
            ChineseName = Chain = OfficeCode = SourceUrl = string.Empty;
            RicList = new List<string>();
        }
    }
    

    public class ChinaIndexQC : GeneratorBase
    {
        private static ChinaIndexQCConfig configObj;

        private static List<ChinaIndex> sseIndexList = new List<ChinaIndex>();
        private static List<ChinaIndex> csiIndexList = new List<ChinaIndex>();
        private static List<ChinaIndex> szseIndexList = new List<ChinaIndex>();

        protected override void Start()
        {
            getSSEIndexList();
            generateFile(sseIndexList, "SSEResult.xls");
            getCSIIndexList();
            generateFile(csiIndexList, "CSIResult.xls");
            getSZSEIndexList();
            generateFile(szseIndexList, "SZSEResult.xls");
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as ChinaIndexQCConfig;
        }


        #region SSE Related
        private void getSSEIndexList()
        {
            var doc = WebClientUtil.GetHtmlDocument(string.Format("{0}/sseportal/index/cn/common/index_list_new.shtml", configObj.SseBaseUri), 180000,"",Encoding.GetEncoding("gb2312"));
            var nodeList = doc.DocumentNode.SelectNodes("//td[@class='table2']//a");
            foreach (HtmlNode node in nodeList)
            {
                //http://www.sse.com.cn/sseportal/index/cn/i000010/intro.
                if (node.ParentNode.Attributes["align"] == null)
                {
                    ChinaIndex index = new ChinaIndex();
                    index.ChineseName = MiscUtil.GetCleanTextFromHtml(node.InnerText);
                    index.SourceUrl = string.Format("{0}{1}",configObj.SseBaseUri, node.Attributes["href"].Value);
                    //http://www.sse.com.cn/sseportal/index/cn/i000010/intro.shtml
                    //http://www.sse.com.cn/sseportal/index/cn/i000010/const_list.shtml
                    index.SourceUrl = index.SourceUrl.Replace("intro.shtml", "const_list.shtml");
                    sseIndexList.Add(index);
                }
            }

            updateSSEIndexListWithRic();
        }

        private void updateSSEIndexListWithRic()
        {
            foreach (ChinaIndex index in sseIndexList)
            {
                var htmlDoc = WebClientUtil.GetHtmlDocument(index.SourceUrl, 180000, "",Encoding.GetEncoding("gb2312"));
                var nodeList = htmlDoc.DocumentNode.SelectNodes("//td[@class='table3']//a");
                foreach (HtmlNode node in nodeList)
                {
                    string officialCode = MiscUtil.GetCleanTextFromHtml(node.InnerText);
                    officialCode = officialCode.Substring(officialCode.IndexOf("(") + 1, (officialCode.IndexOf(")") - officialCode.IndexOf("(") - 1));
                    index.RicList.Add(generateRic(officialCode));  
                }
            }
        }

        #endregion

        #region CSI Related
        private void getCSIIndexList()
        {
            //string pageSource = WebClientUtil.GetPageSource(null, string.Format("{0}/sseportal/csiportal/xzzx/queryindexdownloadlist.do?type=1", configObj.CSI_BASE_URI), 180000, "", Encoding.GetEncoding("gb2312"));
            var doc = WebClientUtil.GetHtmlDocument(string.Format("{0}/sseportal/csiportal/xzzx/queryindexdownloadlist.do?type=1", configObj.CsiBaseUri), 180000, "", Encoding.GetEncoding("gb2312"));
            var nodeList = doc.DocumentNode.SelectNodes("//tr[@align='center']");
            var nodeList2 = doc.DocumentNode.SelectNodes("//tr[@class='list-div-table-header']");

            foreach (HtmlNode node in nodeList2)
            {
                nodeList.Append(node);
            }

            for(int i =0;i<nodeList.Count;i++)
            {
                ChinaIndex index = new ChinaIndex();
                index.ChineseName = MiscUtil.GetCleanTextFromHtml(nodeList[i].ChildNodes[2 * 0 + 1].ChildNodes[2].InnerText);
                try
                {
                    if (nodeList[i].ChildNodes[2 * 5 + 1].ChildNodes.Count == 3)
                    {
                        index.SourceUrl = configObj.CsiBaseUri+MiscUtil.GetCleanTextFromHtml(nodeList[i].ChildNodes[2 * 5 + 1].ChildNodes[1].Attributes["href"].Value);
                        index.OfficeCode = getOfficialCodeForCSI(index.SourceUrl);
                        index.RicList = getRicsForCSI(index.SourceUrl);
                        csiIndexList.Add(index);
                    }
                    else
                    {
                        Logger.Log(string.Format("There's no '成分股列表'in the exchange web site for {0}",index.ChineseName));
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(string.Format("There's error when parsing information for {0}. Exception message: {1} ",index.ChineseName,ex.Message));
                }
            }
        }

        private List<string> getRicsForCSI(string xlsFileUrl)
        {
            List<string> ricList = new List<string>();
            string xlsFilePath = targetDownloadFileDir();
            xlsFilePath += "\\CSI";
            if (!Directory.Exists(xlsFilePath))
            {
                Directory.CreateDirectory(xlsFilePath);
            }

            string[] subLinkNode = xlsFileUrl.Split('/');
            xlsFilePath += "\\"+subLinkNode[subLinkNode.Length - 1];
            WebClientUtil.DownloadFile(xlsFileUrl, 180000, xlsFilePath);
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, xlsFilePath);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                for (int i = 2; i <= lastUsedRow; i++)
                {
                    Range aRange = ExcelUtil.GetRange(i, 1, worksheet);
                    if (aRange != null && aRange.Text.ToString().Trim() != string.Empty)
                    {
                        string officialCode = aRange.Text.ToString();
                        int temp = -1;
                        if (int.TryParse(officialCode, out temp))
                        {
                            officialCode = temp.ToString("D6");
                        }
                        ricList.Add(generateRic(officialCode));
                    }
                }
            }
            return ricList;
        }

        private string getOfficialCodeForCSI(string source)
        {
            string officialCode = string.Empty;
            string[] subNode = source.Split('/');
            officialCode = subNode[subNode.Length - 1].Substring(0, 6);
            return officialCode ;            
        }

        #endregion

        #region SZSE Related
        private void getSZSEIndexList()
        { 
             string pageSource = WebClientUtil.GetPageSource(configObj.SzseBaseUri, 180000, "");
             var doc = WebClientUtil.GetHtmlDocument(string.Format("{0}/main/marketdata/hqcx/zsybg/", configObj.SzseBaseUri), 180000, "", Encoding.GetEncoding("gb2312"));
             string szseIndexSourceFileUrl = MiscUtil.GetCleanTextFromHtml(doc.DocumentNode.SelectNodes("//td[@align='right']/a")[0].Attributes["href"].Value);
             downloadAndParseIndexFile(string.Format("{0}{1}",configObj.SzseBaseUri,szseIndexSourceFileUrl));
             updateSZSEIndexListWithRic();
        }


        private void updateSZSEIndexListWithRic()
        {
            foreach (ChinaIndex index in szseIndexList)
            {
                //http://www.szse.cn/szseWeb/FrontController.szse?ACTIONID=8&CATALOGID=1747&TABKEY=tab1&ENCODE=1&ZSDM=399328
                string url = string.Format("{0}/szseWeb/FrontController.szse?ACTIONID=8&CATALOGID=1747&TABKEY=tab1&ENCODE=1&ZSDM={1}", configObj.SzseBaseUri, index.Chain);
                //string pageSource = WebClientUtil.GetPageSource(null, url, 180000, "", Encoding.GetEncoding("gb2312"));
                string ricFilePath = targetDownloadFileDir()+"\\SZSE\\";
                ricFilePath+=string.Format("{0}.xls",index.Chain);
                WebClientUtil.DownloadFile(url, 180000,ricFilePath);

                using (ExcelApp app = new ExcelApp(false, false))
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, ricFilePath);
                    var worksheet = workbook.Worksheets[1] as Worksheet;
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                    for (int i = 2; i <= lastUsedRow; i++)
                    {
                        Range aRange = ExcelUtil.GetRange(i, 1, worksheet);
                        if (aRange != null && !string.IsNullOrEmpty(aRange.Text.ToString().Trim()))
                        {
                            index.RicList.Add(generateRic(aRange.Text.ToString().Trim()));
                        }
                    }
                }
            }
        }

        private void downloadAndParseIndexFile(string url)
        {
            string szseIndexFilePath = targetDownloadFileDir();
            szseIndexFilePath += "\\SZSE";
            if (!Directory.Exists(szseIndexFilePath))
            {
                Directory.CreateDirectory(szseIndexFilePath);
            }
            szseIndexFilePath += "\\Index.xls";

            WebClientUtil.DownloadFile(url, 180000, szseIndexFilePath);
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, szseIndexFilePath);
                var worksheet = workbook.Worksheets[1] as Worksheet;
                int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                for (int i = 2; i <= lastUsedRow; i++)
                {
                    Range aRange = ExcelUtil.GetRange(i, 1, worksheet);
                    if (aRange != null && aRange.Text.ToString().Trim() != string.Empty)
                    {
                        ChinaIndex index = new ChinaIndex();
                        index.Chain = aRange.Text.ToString().Trim();
                        index.ChineseName = ExcelUtil.GetRange(i, 2, worksheet).Text.ToString();
                        szseIndexList.Add(index);
                    }
                }
            }
        }

        private string targetDownloadFileDir()
        {
            string dir = @"D:\CHN_INDEX_";
            dir += DateTime.Now.ToString("ddMMMyy");
            return dir;
        }
        #endregion

        private string generateRic(string officialCode)
        {
            if (officialCode.StartsWith("6"))
            {
                officialCode += ".SS";
            }

            else if (officialCode.StartsWith("300") || officialCode.StartsWith("002"))
            {
                officialCode += ".SZ";
            }
            else if (officialCode.StartsWith("122"))
            {
                officialCode = string.Format("CN{0}=SS", officialCode);
            }

            return officialCode;
        }

        private void generateFile(List<ChinaIndex> indexList, string fileName)
        {
            if (indexList.Count == 0)
            {
                Logger.Log("No item in the index list.");
                return;
            }

            using (ExcelApp app = new ExcelApp(false,false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app,string.Format("{0}\\{1}",configObj.TargetFileDir,fileName));
                int sheetNum = (indexList.Count + configObj.ChainNumPerSheet-1) / configObj.ChainNumPerSheet;
                int startPos = 0;
                if(sheetNum>workbook.Worksheets.Count)
                {
                    workbook.Worksheets.Add(Missing.Value, Missing.Value, sheetNum - workbook.Worksheets.Count, Missing.Value);
                }
                for (int i = 0; i < sheetNum; i++)
                {
                    var worksheet = workbook.Worksheets[i + 1] as Worksheet;
                    
                    startPos = i * configObj.ChainNumPerSheet;
                    int endPos = indexList.Count < (startPos + configObj.ChainNumPerSheet) ? indexList.Count : (startPos + configObj.ChainNumPerSheet);
                    WriterWorksheet(worksheet, indexList, startPos, endPos);
                }
                TaskResultList.Add(new TaskResultEntry(fileName,"",workbook.FullName));
                workbook.Save();
                workbook.Close(true, workbook.FullName, true);

            }
        }

        private void WriterWorksheet(Worksheet worksheet, List<ChinaIndex> indexList, int startPos, int endPos)
        {
            using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Down))
            {
                for (int i = startPos; i < endPos; i++)
                {
                    writer.WriteLine(indexList[i].Chain);
                    writer.WriteLine(indexList[i].ChineseName);
                    foreach (string ric in indexList[i].RicList)
                    {
                        writer.WriteLine(ric);
                    }
                    writer.PlaceNext(1, writer.Col + 1);
                }
            }
        }
    }
}
