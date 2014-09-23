using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks.China
{
    [ConfigStoredInDB]
    public class ChinaISINConfig
    {
        [StoreInDB]
        [DisplayName("Source file name")]
        [Description("Full path of file 'China Masterfile 2008 v5.xls' ")]
        public string SourceFilePath { get; set; }

        [StoreInDB]
        [DisplayName("Target isin file name")]
        [Description("Target file name which contain ISIN information ")]
        public string TargetIsinFileName{get;set;}

        [StoreInDB]
        [Category("Worksheet")]
        [DisplayName("Equity")]
        public string EquityWorksheet { get; set; }

        [StoreInDB]
        [Category("Worksheet")]
        [DisplayName("Bond")]
        public string BondWorksheet { get; set; }

        [StoreInDB]
        [DisplayName("Query Url")]
        public string QueryUrl{get;set;}

    }

    public class RicISINInfo
    {
        public string OfficialCode { get; set; }
        public string Name { get; set; }
        public string ISIN { get; set; }
        public string type { get; set; }
        public string ric { get; set; }
        public RicISINInfo()
        {
            OfficialCode = Name = ISIN = type = ric = string.Empty;
        }
    }

    public class ChinaISINGenerator : GeneratorBase
    {
        private static ChinaISINConfig configObj;

        protected override void Start()
        {
            StartISINGenerator();
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as ChinaISINConfig;
        }

        public void StartISINGenerator()
        {
            List<RicISINInfo> newGenerateRicISINList = GetAllNewGenerateISINList();
            GenerateISINFile(newGenerateRicISINList);

        }

        //Get the official code of which ISIN needs updated
        public List<RicISINInfo> GetAllNewGenerateISINList()
        {
            List<RicISINInfo> ricInfoList = new List<RicISINInfo>();
            Range nameRange;
            Range isinRange;
            Range ricRange;
            int lastUsedRow;
            File.Copy(configObj.SourceFilePath, NewTargetFilePath(configObj.SourceFilePath));

            using (ExcelApp app = new ExcelApp(false,false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, configObj.SourceFilePath);

                //Get Equity ric from "Equity" worksheet
                var equityWorksheet = ExcelUtil.GetWorksheet(configObj.EquityWorksheet, workbook);
                if (equityWorksheet == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.EquityWorksheet, workbook.Name));
                }

                lastUsedRow = equityWorksheet.UsedRange.Row + equityWorksheet.UsedRange.Rows.Count - 1;
                for (int i = 1; i <= lastUsedRow; i++)
                {
                    nameRange = ExcelUtil.GetRange(i, 20, equityWorksheet);
                    isinRange = ExcelUtil.GetRange(i, 19, equityWorksheet);
                    ricRange = ExcelUtil.GetRange(i, 3, equityWorksheet);
                    if ((nameRange.Value2 != null && nameRange.Value2.ToString() != string.Empty) && (isinRange.Value2 == null || isinRange.Value2.ToString() == string.Empty))
                    {
                        RicISINInfo ricInfo = new RicISINInfo
                        {
                            type = "Equity",
                            OfficialCode = ExcelUtil.GetRange(i, 6, equityWorksheet).Value2.ToString(),
                            Name = nameRange.Value2.ToString()
                        };
                        ricInfo.ISIN = GetISINFromCode(ricInfo.OfficialCode);
                        ricInfo.ric = ricRange.Value2.ToString();
                        if (ricInfo.ISIN != string.Empty)
                        {
                            isinRange.Value2 = ricInfo.ISIN;
                            ricInfoList.Add(ricInfo);
                        }
                    }
                }



                //Get bond ric from "Bond" worksheet

                List<RicISINInfo> allBondISINList = GetAllISINForGoverBond();
                var bondWorksheet = ExcelUtil.GetWorksheet(configObj.BondWorksheet, workbook);
                if (bondWorksheet == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("Cannot get worksheet {0} from workbook {1}", configObj.BondWorksheet, workbook.Name));
                }

                lastUsedRow = bondWorksheet.UsedRange.Row + bondWorksheet.UsedRange.Rows.Count - 1;
                for (int j = 1; j <= lastUsedRow; j++)
                {
                    nameRange = ExcelUtil.GetRange(j, 22, bondWorksheet);
                    isinRange = ExcelUtil.GetRange(j, 21, bondWorksheet);
                    ricRange = ExcelUtil.GetRange(j,3,bondWorksheet);
                    if ((nameRange.Value2 != null && nameRange.Value2.ToString() != string.Empty) && ((isinRange.Value2 == null) || (isinRange.Value2.ToString() == string.Empty)))
                    {
                        RicISINInfo ricInfo = new RicISINInfo
                        {
                            type = "Bond",
                            Name = nameRange.Value2.ToString(),
                            OfficialCode = ExcelUtil.GetRange(j, 5, bondWorksheet).Value2.ToString(),
                            ric = ricRange.Value2.ToString()
                        };
                        ricInfo.ISIN = !ricInfo.Name.Contains("國債") ? GetISINFromCode(ricInfo.OfficialCode) : GetISINFromName(ricInfo.Name, allBondISINList);

                        if (ricInfo.ISIN != string.Empty)
                        {
                            ricInfoList.Add(ricInfo);
                            isinRange.Value2 = ricInfo.ISIN;
                        }
                    }
                }
                TaskResultList.Add(new TaskResultEntry("Updated Source File", "All the newly generated ISIN had been updated into the file.", workbook.FullName));
                workbook.Close(true, workbook.FullName, true);
            }

            return ricInfoList;
        }

        public string NewTargetFilePath(string sourceFilePath)
        {
            string dir = Path.GetDirectoryName(sourceFilePath);

            dir = Path.Combine(dir, DateTime.Now.ToString("dd_MMM_yyyy"));
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            string newFilePath = Path.Combine(dir,Path.GetFileName(sourceFilePath));
            if (File.Exists(newFilePath))
            {
                File.Delete(newFilePath);
            }
            return newFilePath;
        }


        public void GenerateISINFile(List<RicISINInfo> ricInfoList)
        {
            using (ExcelApp app = new ExcelApp(false,false))
            {
                string targetFilePath = NewTargetFilePath(Path.Combine(Path.GetDirectoryName(configObj.SourceFilePath), configObj.TargetIsinFileName));
                if (File.Exists(targetFilePath))
                {
                    File.Delete(targetFilePath);
                }
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, targetFilePath);
                var worksheet = (Worksheet)workbook.Worksheets[1];
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 1, 1, ExcelLineWriter.Direction.Right))
                {
                    writer.WriteLine("Ric");
                    writer.WriteLine("Official Code");
                    writer.WriteLine("ISIN");
                    writer.WriteLine("Name");
                    writer.WriteLine("Type");
                    writer.PlaceNext(writer.Row + 1, 1);
                    foreach (RicISINInfo ricISINInfo in ricInfoList)
                    {
                        writer.WriteLine(ricISINInfo.ric);
                        ExcelUtil.GetRange(writer.Row, writer.Col, worksheet).NumberFormat = "@";
                        writer.WriteLine(ricISINInfo.OfficialCode);
                        writer.WriteLine(ricISINInfo.ISIN);
                        writer.WriteLine(ricISINInfo.Name);
                        writer.WriteLine(ricISINInfo.type);
                        writer.PlaceNext(writer.Row + 1, 1);
                    }
                }
                workbook.Save();
                TaskResultList.Add(new TaskResultEntry("Newly Generated ISIN File", "The file contains all the newly generated ISIN.", targetFilePath));
            }
        }

        //http://www.chinaclear.cn/isin/isinbiz/queryCommon.do?m=querycommon
        //http://www.chinaclear.cn/isin/isinbiz/queryCommon.do
        public string GetISINFromCode(string officialCode)
        {
            string isin = string.Empty;
            string postData = string.Format("m=querycommon&currentPage=0&securitiesCode={0}&isin=&securitiesNameCL=&query=%B2%E9%D1%AF",officialCode);
            var htmlDoc = WebClientUtil.GetHtmlDocument(configObj.QueryUrl, 18000, postData);
            htmlDoc.DocumentNode.SelectNodes("//body//table//tr//td//table//tr//td//form//table//tr//td[@class='isnr']");
            var htmlDocumentNodes = htmlDoc.DocumentNode.SelectNodes("//body//table//tr//td//table//tr//td//table//tr//td[@class='isnr']");

            //TODO: To Find a Better Way to Get the ISIN 
            if (htmlDocumentNodes.Count == 11)
            {
                isin = MiscUtil.GetCleanTextFromHtml(htmlDocumentNodes[8].InnerText);
            }
            return isin;
        }


        //Get all the newly generated ISIN for goverment bond
        //http://www.chinaclear.cn/isin/user/userApplyLogin.do?m=enter
        //http://www.chinaclear.cn/isin/user/userApplyLogin.do?m=queryFast
        public List<RicISINInfo> GetAllISINForGoverBond()
        {
            AdvancedWebClient wc = new AdvancedWebClient();
            string postData = "loginName=&password=&securitiesName=%BC%C7%D5%CB%CA%BD%B8%BD%CF%A2&securitiesCode=";
            string url = "http://www.chinaclear.cn/isin/user/userApplyLogin.do?m=enter";
            string pageSource = WebClientUtil.GetPageSource(wc, url, 18000, postData);
            pageSource = WebClientUtil.GetPageSource(wc, "http://www.chinaclear.cn/isin/user/userApplyLogin.do?m=queryFast", 18000, "",Encoding.GetEncoding("gb2312"));
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(pageSource);
            var nodeList = htmlDoc.DocumentNode.SelectNodes("//tr[@class='td5']");
            return nodeList.Select(node => new RicISINInfo
            {
                ISIN = MiscUtil.GetCleanTextFromHtml(node.ChildNodes[1*2 + 1].InnerText), Name = MiscUtil.GetCleanTextFromHtml(node.ChildNodes[2*2 + 1].InnerText)
            }).ToList();
        }

        public string GetISINFromName(string bondName, List<RicISINInfo> isinList)
        {
            string isin = string.Empty;
            foreach (RicISINInfo ricISIN in isinList.Where(ricISIN => ricISIN.Name == bondName))
            {
                isin = ricISIN.ISIN;
            }
            return isin;
        }
    }
}
