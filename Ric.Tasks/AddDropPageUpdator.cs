using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Ric.Util;
using Ric.Db.Manager;
using System.Data;
using System.Windows.Forms;
using Ric.Core;

namespace Ric.Tasks
{
    /// <summary>
    /// Items which can be configured by users
    /// </summary>
    [ConfigStoredInDB]
    public class KOREA_AddDropChangeConfig
    {
        //[StoreInDB]
        //[DefaultValue("C:\\Korea_Auto\\ChangePage\\FM\\")]
        //[Description("Directory which contains all kinds of FM files. E.g. C:\\Korea_Auto\\ChangePage\\FM\\")]
        //public string FM { get; set; }

        [StoreInDB]
        [DefaultValue("C:\\Korea_Auto\\ChangePage\\RIC_Convs_Template_V3.xls")]
        [Description("Full path of RIC Conversion Template excel file. E.g. C:\\Korea_Auto\\ChangePage\\RIC_Convs_Template_V3.xls ")]
        public string TemplatePath { get; set; }
        
    }

    /// <summary>
    /// Information which needs to be updated, which are column names of final add& drop upload file
    /// </summary>
    public class AddDropChangeInfo
    {
        public string EventAction { get; set; }
        public string RicSeqId { get; set; }
        public string ChangeType { get; set; }
        public DateTime Date { get; set; }
        public string DescriptionWas { get; set; }
        public string DescriptionNow { get; set; }
        public string RicWas { get; set; }
        public string RicNow { get; set; }
        public string ISINWas { get; set; }
        public string ISINNow { get; set; }
        public string SecondID {get;set;}
        public string SecondWas { get; set; }
        public string SecondNow { get; set; }
        public string ThomsonWas { get; set; }
        public string ThomsonNow { get; set; }
        public string Exchange { get; set; }
        public string Asset { get; set; }

        public AddDropChangeInfo()
        {
            this.EventAction = "Create";
            this.SecondID = "Official Code";
            this.RicSeqId = this.ChangeType = this.DescriptionNow = this.DescriptionWas = this.RicWas = this.RicNow = this.ISINNow = this.ISINWas = 
                                this.SecondWas = this.ThomsonNow = this.ThomsonWas = this.ThomsonNow = this.Exchange = this.Asset = string.Empty;
        }
    }

    /// <summary>
    /// Add drop change page file updator related
    /// </summary>
    public class AddDropPageUpdator : GeneratorBase
    {       
        private static KOREA_AddDropChangeConfig configObj = null;        

        protected override void Start()
        {
            //For DB version
            UpdateAddDropPageFromDb();
            //UpdateAddDropPage(configObj.FM);
        }

        protected override void Initialize()
        {
            configObj = Config as KOREA_AddDropChangeConfig;
            //configObj.FM = @"D:\test\PageUpdate\FM";
            //configObj.TemplatePath = @"D:\test\PageUpdate\Out\RIC_Convs_Template_V3.xls";
            //if (string.IsNullOrEmpty(configObj.FM))
            //{
            //    string msg = "'FM' in configuration can't be blank!";
            //    MessageBox.Show(msg, "Error");
            //    Logger.Log(msg, Logger.LogType.Error);
            //    throw new Exception(msg);
            //}
            if (string.IsNullOrEmpty(configObj.TemplatePath))
            {
                string msg = "'TemplatePath' in configuration can't be blank!";
                MessageBox.Show(msg, "Error");
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception(msg);
            }
            if (!File.Exists(configObj.TemplatePath))
            {
                string msg = string.Format("Can't find file {0}. Please check the path.", configObj.TemplatePath);
                MessageBox.Show(msg, "Error");
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception(msg);
            }
            Logger.Log("Initialize OK!");

            TaskResultList.Add(new TaskResultEntry("Log", "Log", Logger.FilePath)); 
        }

        private void UpdateAddDropPage(string filePath)
        {
            List<AddDropChangeInfo> addDropEventList = GetAllAdddropEventList(filePath);            

            if (addDropEventList.Count > 0)
            {
                UpdateAddDropFile(configObj.TemplatePath, addDropEventList);
            }
        }

        private void UpdateAddDropPageFromDb()
        {
            string tablename = string.Format("fn_GetEtiKoreaPageUpdateInfo('{0}')", DateTime.Today.ToString("yyyy-MM-dd"));
            System.Data.DataTable dt = ManagerBase.Select(tablename);

            if (dt == null)
            {
                string msg = "Error found when get add/change/drop information from database.";
                Logger.Log(msg, Logger.LogType.Error);
                return;            
            }


            if(dt.Rows.Count == 0)
            {
                string msg = "No add/change/drop record updated today.";
                Logger.Log(msg);
                return;
            }

            string msg1 = string.Format("{0} add/change/drop record(s) updated today.", dt.Rows.Count);
            Logger.Log(msg1);

            GeneratePageUpdateFile(configObj.TemplatePath, dt);
        }

        private void GeneratePageUpdateFile(string templateFilePath, System.Data.DataTable dt)
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, templateFilePath);
                var worksheet = ExcelUtil.GetWorksheet("INPUT SHEET", workbook);
                if (worksheet == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("There's no worksheet: {0}", worksheet.Name));
                }
                ((Range)worksheet.Columns["D"]).NumberFormat = "@";
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 3, 1, ExcelLineWriter.Direction.Right))
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        writer.WriteLine("Create");
                        writer.WriteLine("");
                        writer.WriteLine(Convert.ToString(dr["ChangeType"]));
                        string effectiveDate = Convert.ToDateTime(dr["EffectiveDate"]).ToString("ddMMMyy");
                        writer.WriteLine(effectiveDate);
                        writer.WriteLine(Convert.ToString(dr["DescriptionWas"]));
                        writer.WriteLine(Convert.ToString(dr["DescriptionNow"]));
                        string ricWas = (Convert.ToString(dr["RICWas"])).Replace("D^", "").Replace("^", "");
                        writer.WriteLine(ricWas);
                        writer.WriteLine(Convert.ToString(dr["RICNow"]));
                        writer.WriteLine(Convert.ToString(dr["ISINWas"]));
                        writer.WriteLine(Convert.ToString(dr["ISINNow"]));
                        writer.WriteLine("Official Code");
                        writer.WriteLine(Convert.ToString(dr["SecondWas"]));
                        writer.WriteLine(Convert.ToString(dr["SecondNow"]));
                        writer.WriteLine("");
                        writer.WriteLine("");
                        writer.WriteLine("");
                       //Exchange
                        string exchange = Convert.ToString(dr["RICNow"]);
                        if (string.IsNullOrEmpty(exchange))
                        {
                            exchange = Convert.ToString(dr["RICWas"]);
                        }
                        if (exchange.Contains(".KS"))
                        {
                            exchange = "KSC";
                        }
                        else if(exchange.Contains(".KQ"))
                        {
                            exchange = "KOE";
                        }
                        else if (exchange.Contains(".KN"))
                        {
                            exchange = "KNX";
                        }
                        writer.WriteLine(exchange);
                        string asset = Convert.ToString(dr["Asset"]);
                        if (asset.Equals("KDR"))
                        {
                            asset = "DRC";
                        }
                        writer.WriteLine(asset);
                        writer.PlaceNext(writer.Row + 1, 1);
                    }
                }

                //Run Macros
                app.ExcelAppInstance.GetType().InvokeMember("Run",
                    BindingFlags.Default | BindingFlags.InvokeMethod,
                    null,
                    app.ExcelAppInstance,
                    new object[] { "FormatData" });

                string targetFilePath = Path.GetDirectoryName(templateFilePath);
                targetFilePath = Path.Combine(targetFilePath, DateTime.Today.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(targetFilePath))
                { 
                    Directory.CreateDirectory(targetFilePath);
                }

                targetFilePath += "\\Result_Add_Drop_Upload_File.xls";
                workbook.SaveCopyAs(targetFilePath);
                TaskResultList.Add(new TaskResultEntry("Result File", "", targetFilePath));
                workbook.Close(false, templateFilePath, false);
            }
        }

        private void UpdateAddDropFile(string templateFilePath, List<AddDropChangeInfo> addDropEventList)
        {
            using (ExcelApp app = new ExcelApp(false, false))
            {
                var workbook = ExcelUtil.CreateOrOpenExcelFile(app, templateFilePath);
                var worksheet = ExcelUtil.GetWorksheet("INPUT SHEET", workbook);
                if (worksheet == null)
                {
                    Logger.LogErrorAndRaiseException(string.Format("There's no worksheet: {0}",worksheet.Name));
                }
                ((Range)worksheet.Columns["D"]).NumberFormat = "@";
                using (ExcelLineWriter writer = new ExcelLineWriter(worksheet, 3, 1, ExcelLineWriter.Direction.Right))
                {
                    foreach (AddDropChangeInfo changeInfo in addDropEventList)
                    {
                        writer.WriteLine(changeInfo.EventAction);
                        writer.WriteLine(changeInfo.RicSeqId);
                        writer.WriteLine(changeInfo.ChangeType);                        
                        writer.WriteLine(changeInfo.Date.ToString("ddMMMyy"));
                        writer.WriteLine(changeInfo.DescriptionWas);
                        writer.WriteLine(changeInfo.DescriptionNow);
                        writer.WriteLine(changeInfo.RicWas);
                        writer.WriteLine(changeInfo.RicNow);
                        writer.WriteLine(changeInfo.ISINWas);
                        writer.WriteLine(changeInfo.ISINNow);
                        writer.WriteLine(changeInfo.SecondID);
                        writer.WriteLine(changeInfo.SecondWas);
                        writer.WriteLine(changeInfo.SecondNow);
                        writer.WriteLine(changeInfo.ThomsonWas);
                        writer.WriteLine(changeInfo.ThomsonNow);
                        writer.WriteLine("");
                        writer.WriteLine(changeInfo.Exchange);
                        writer.WriteLine(changeInfo.Asset);
                        writer.PlaceNext(writer.Row + 1, 1);
                    }
                }

                //Run Macros
                app.ExcelAppInstance.GetType().InvokeMember("Run",
                    BindingFlags.Default | BindingFlags.InvokeMethod,
                    null,
                    app.ExcelAppInstance,
                    new object[] { "FormatData" });

                string targetFilePath = Path.GetDirectoryName(templateFilePath);
                targetFilePath = Path.Combine(targetFilePath, DateTime.Today.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(targetFilePath))
                {
                    Directory.CreateDirectory(targetFilePath);
                }
                targetFilePath += "\\Result_Add_Drop_Upload_File.xls";
                workbook.SaveCopyAs(targetFilePath);
                TaskResultList.Add(new TaskResultEntry("Result File", "", targetFilePath));
                //workbook.Save();
                workbook.Close(false, templateFilePath, false);
            }
        }
        private List<AddDropChangeInfo> GetAllAdddropEventList(string fmFileDir)
        {
            List<AddDropChangeInfo> addDropList = new List<AddDropChangeInfo>();
            List<string> fileList = new List<string>();
            string supportedExtensions = "*.xls,*.xlsx";
            foreach (string file in Directory.GetFiles(fmFileDir, "*.*", SearchOption.TopDirectoryOnly).Where(s => supportedExtensions.Contains(Path.GetExtension(s).ToLower())))
            {                              
                if (File.GetCreationTime(file).ToString("ddMMMyy") == DateTime.Now.ToString("ddMMMyy"))
                {
                    fileList.Add(file);  
                }
            }
            if (fileList.Count == 0)
            {
                string msg = string.Format("No qulified FM file found in folder:{0}", "FM");//configObj.FM);
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception(msg);
            }

            using (ExcelApp app = new ExcelApp(false, false))
            {
                foreach (string filePath in fileList)
                {
                    var workbook = ExcelUtil.CreateOrOpenExcelFile(app, filePath);
                    var worksheet = workbook.Worksheets[1] as Worksheet;
                    int lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                    string fileName = Path.GetFileNameWithoutExtension(filePath).ToUpper();
                    if (fileName.Contains("AFTERNOON"))
                    {
                        addDropList.AddRange(GetELWAddEventList(worksheet, lastUsedRow));
                        continue;
                    }

                    if (fileName.Contains("MORNING"))
                    {
                        addDropList.AddRange(GetELWDropEventList(worksheet, lastUsedRow));
                        continue;
                    }

                    if (fileName.ToUpper().Contains("RIGHT ADD"))
                    {
                        addDropList.AddRange(GetRightsAndCompanyWarrantAddEventList(worksheet, lastUsedRow, "RTS"));
                        continue;
                    }

                    if (fileName.Contains("COMPANY WARRANT ADD"))
                    {
                        addDropList.AddRange(GetRightsAndCompanyWarrantAddEventList(worksheet, lastUsedRow, "WNT"));
                        continue;
                    }

                    if (fileName.Contains("ADD") && (!fileName.Contains("AFTERNOON")) &&(!fileName.Contains("RIGHT")) &&(!fileName.Contains("COMPANY")))
                    {
                        addDropList.AddRange(GetPeoPrfReitCefAddEventList(worksheet, lastUsedRow));
                        continue;
                    }
                    //if ((fileName.Contains("KR FM") && fileName.Contains("PEO ADD")) ||
                    //            fileName.Contains("KR FM") && fileName.Contains("ETF") && fileName.Contains("ADD") ||
                    //            fileName.Contains("Korea FM") && fileName.Contains("REIT") && fileName.Contains("ADD")) 
                    //{
                        //addDropList.AddRange(getPeoPrfReitCefAddEventList(worksheet, lastUsedRow));
                        //continue;
                    //}

                    if (fileName.Contains("DROP"))
                    {
                        if (fileName.Contains("PRF"))
                        {
                            addDropList.AddRange(GetPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "PRF"));
                            continue;
                        }
                        else if (fileName.Contains("ETF"))
                        {
                            //worksheet = workbook.Worksheets[1] as Worksheet;
                            lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                            addDropList.AddRange(GetPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "ETF"));
                            continue;
                        }
                        else if (fileName.Contains("CEF"))
                        {
                            addDropList.AddRange(GetPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "CEF"));
                            continue;
                        }

                        else if (fileName.Contains("REIT"))
                        {
                            addDropList.AddRange(GetPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "REI"));
                            continue;
                        }

                        else if (fileName.Contains("RIGHT"))
                        {
                            addDropList.AddRange(GetRightsAndCompanyWarrantDropEventList(worksheet, lastUsedRow, "RTS"));
                            continue;
                        }

                        else if (fileName.Contains("COMPANY"))
                        {
                           // worksheet = ExcelUtil.GetWorksheet("Deletion", workbook);
                            lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                            addDropList.AddRange(GetRightsAndCompanyWarrantDropEventList(worksheet, lastUsedRow, "WNT"));
                            continue;
                        }
                        else
                        {
                            addDropList.AddRange(GetPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "ORD"));
                            continue;
                        }

                    }
                    //if (fileName.Contains("KR FM (DROP) REQUEST"))
                    //{
                    //    addDropList.AddRange(getPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "ORD"));
                    //    continue;
                    //}

                    //if (fileName.Contains("KR FM (PRF DROP)"))
                    //{
                    //    addDropList.AddRange(getPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "PRF"));
                    //    continue;
                    //}
                    //if (fileName.Contains("KR_FM (ETF DROP) REQUEST"))
                    //{
                    //    worksheet = ExcelUtil.GetWorksheet("Deletion", workbook);
                    //    lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                    //    addDropList.AddRange(getPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "ETF"));
                    //    continue;
                    //}
                    //if (fileName.Contains("KR FM (CEF drop) Request"))
                    //{
                    //    addDropList.AddRange(getPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "CEF"));
                    //    continue;
                    //}
                    //if (fileName.Contains("KR FM (REIT drop) Request"))
                    //{
                    //    addDropList.AddRange(getPeoPrfReitCefDropEventList(worksheet, lastUsedRow, "REI"));
                    //    continue;
                    //}
                    //DRC

                    //if (fileName.Contains("KR FM (Right DROP) Request"))
                    //{
                    //    addDropList.AddRange(getRightsAndCompanyWarrantDropEventList(worksheet, lastUsedRow, "RTS"));
                    //    continue;
                    //}
                    //if (fileName.Contains("(Company Warrant Drop)"))
                    //{
                    //    worksheet = ExcelUtil.GetWorksheet("Deletion", workbook);
                    //    lastUsedRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                    //    addDropList.AddRange(getRightsAndCompanyWarrantDropEventList(worksheet, lastUsedRow, "WNT"));
                    //    continue;
                    //}
                    if (fileName.Contains("NAME CHANGE"))
                    {
                        addDropList.AddRange(GetNameChangeEventList(worksheet, lastUsedRow));
                        continue;
                    }
                }
            }
            return addDropList;
        }

        //get elw add event list:  file name Korea FM for (dd-mm-yy) (Afternoon).xls
        private List<AddDropChangeInfo> GetELWAddEventList(Worksheet worksheet, int lastUsedRow)
        {
            List<AddDropChangeInfo> elwAddEventList = new List<AddDropChangeInfo>();
            for (int i = 5; i <= lastUsedRow; i++)
            {
                if (ExcelUtil.GetRange(i, 1, worksheet).Text != null && ExcelUtil.GetRange(i, 1, worksheet).Text.ToString() != string.Empty)
                {
                    AddDropChangeInfo changeInfo = new AddDropChangeInfo();
                    changeInfo.ChangeType = "Add";
                    changeInfo.Date = DateTime.ParseExact(ExcelUtil.GetRange(i, 2, worksheet).Text.ToString().Trim(), "dd-MMM-yy", null);
                    //changeInfo.DescriptionWas = "";
                    changeInfo.DescriptionNow = ExcelUtil.GetRange(i, 5, worksheet).Text.ToString().Trim();
                    //changeInfo.RicWas = "";
                    changeInfo.RicNow = ExcelUtil.GetRange(i, 3, worksheet).Text.ToString().Trim();
                    //changeInfo.ISINWas = "";
                    changeInfo.ISINNow = ExcelUtil.GetRange(i, 6, worksheet).Text.ToString().Trim();
                    //changeInfo.SecondWas = "";
                    changeInfo.SecondNow = ExcelUtil.GetRange(i, 7, worksheet).Text.ToString().Trim();
                    changeInfo.Exchange = "KSC";
                    changeInfo.Asset = "WNT";
                    elwAddEventList.Add(changeInfo);
                }
            }
            return elwAddEventList;
        }

        //get elw drop event list: FM for (dd-mm-yy) (Morning).xls
        private List<AddDropChangeInfo> GetELWDropEventList(Worksheet worksheet, int lastUsedRow)
        {
            List<AddDropChangeInfo> elwDropEventList = new List<AddDropChangeInfo>();
            int startPos = 1;
            for (int i = startPos; i <= lastUsedRow; i++)
            {
                if (ExcelUtil.GetRange(i, 1, worksheet).Text != null && ExcelUtil.GetRange(i, 1, worksheet).Text.ToString().ToUpper() == "DROP")
                {
                    startPos = i;
                    break;
                }
            }
            for (int i = startPos+2; i <= lastUsedRow; i++)
            {
                if (ExcelUtil.GetRange(i, 1, worksheet).Text != null && ExcelUtil.GetRange(i, 1, worksheet).Text.ToString() != string.Empty)
                {
                    AddDropChangeInfo changeInfo = new AddDropChangeInfo();
                    changeInfo.ChangeType = "Delete";
                    changeInfo.Date = DateTime.ParseExact(ExcelUtil.GetRange(i, 2, worksheet).Text.ToString().Trim(), "dd-MMM-yy", null);
                    changeInfo.DescriptionWas = ExcelUtil.GetRange(i, 5, worksheet).Text.ToString().Trim();
                    //changeInfo.DescriptionNow = "";
                    changeInfo.RicWas = ExcelUtil.GetRange(i, 3, worksheet).Text.ToString().Trim();
                    //changeInfo.RicNow = "";
                    changeInfo.ISINWas = ExcelUtil.GetRange(i, 6, worksheet).Text.ToString().Trim();
                    //changeInfo.ISINNow = "";
                    changeInfo.SecondWas = ExcelUtil.GetRange(i, 7, worksheet).Text.ToString().Trim();
                    //changeInfo.SecondNow = "";
                    changeInfo.Exchange = "KSC";
                    changeInfo.Asset = "WNT";
                    elwDropEventList.Add(changeInfo);
                }
            }
            return elwDropEventList;
        }

        //get PEO, PRF, REIT, ETF, CEF, ADD event list:  Sample File: KR FM (PEO ADD) _ 079980.KS(wef 2012-Feb-23).xlsx and so on
        private List<AddDropChangeInfo> GetPeoPrfReitCefAddEventList(Worksheet worksheet, int lastUsedRow)
        {
            List<AddDropChangeInfo> PEOPRFREITCefAddEventList = new List<AddDropChangeInfo>();
            int startRow = 1;
            for (int i = 1; i <= lastUsedRow; i++)
            {
                if (ExcelUtil.GetRange(i, 1, worksheet).Text != null && ExcelUtil.GetRange(i, 1, worksheet).Text.ToString() == "Updated Date")
                {
                    startRow = i + 1;
                    break;
                }
            }
            for (int i = startRow; i <= lastUsedRow; i++)
            {
                if (ExcelUtil.GetRange(i, 2, worksheet).Text != null && ExcelUtil.GetRange(i, 2, worksheet).Text.ToString() != string.Empty)
                {
                    AddDropChangeInfo changeInfo = new AddDropChangeInfo();
                    changeInfo.ChangeType = "Add";
                    changeInfo.Date = DateTime.ParseExact(ExcelUtil.GetRange(i, 2, worksheet).Text.ToString().Trim(), "dd-MMM-yy", null);
                    //changeInfo.DescriptionWas = "";
                    changeInfo.DescriptionNow = ExcelUtil.GetRange(i, 7, worksheet).Text.ToString().Trim();
                    //changeInfo.RicWas = "";
                    changeInfo.RicNow = ExcelUtil.GetRange(i, 3, worksheet).Text.ToString().Trim();
                    //changeInfo.ISINWas = "";
                    changeInfo.ISINNow = ExcelUtil.GetRange(i, 8, worksheet).Text.ToString().Trim();
                    //changeInfo.SecondWas = "";
                    changeInfo.SecondNow = ExcelUtil.GetRange(i, 9, worksheet).Text.ToString().Trim();
                    if (changeInfo.RicNow.EndsWith(".KS"))
                    {
                        changeInfo.Exchange = "KSC";
                    }
                    else if (changeInfo.RicNow.EndsWith(".KQ"))
                    {
                        changeInfo.Exchange = "KOE"; 
                    }
                    else if (changeInfo.RicNow.EndsWith(".KN"))
                    {
                        changeInfo.Exchange = "KNX";
                    }

                    if (ExcelUtil.GetRange(i, 4, worksheet).Text.ToString().Trim() == "REIT")
                    {
                        changeInfo.Asset = "REI";
                    }
                    else if (ExcelUtil.GetRange(i, 4, worksheet).Text.ToString().Trim() == "KDR")
                    {
                        changeInfo.Asset = "DRC";
                    }
                    else
                    {
                        changeInfo.Asset = ExcelUtil.GetRange(i, 4, worksheet).Text.ToString().Trim();
                    }

                    PEOPRFREITCefAddEventList.Add(changeInfo);
                }
            }
            return PEOPRFREITCefAddEventList;
        }

        //get company warrant add event list
        //Sample files: KR FM(Company Warrant ADD)Request_102940W.KQ (wef 2011-SEP-09) .xls 
        //FM(Right Add)_027970_r.KS (wef 2012-FEB-28).xls
        private List<AddDropChangeInfo> GetRightsAndCompanyWarrantAddEventList(Worksheet worksheet, int lastUsedRow, string type)
        {
            List<AddDropChangeInfo> RightsAndcompamyWarrantddEventList = new List<AddDropChangeInfo>();
            int currentRow = 3;
            while(currentRow<=lastUsedRow)
            {
                if (ExcelUtil.GetRange(currentRow, 1, worksheet).Text != null && ExcelUtil.GetRange(currentRow, 1, worksheet).Text.ToString() != string.Empty)
                {
                    AddDropChangeInfo changeInfo = new AddDropChangeInfo();
                    changeInfo.ChangeType = "Add";
                    changeInfo.Date = DateTime.ParseExact(ExcelUtil.GetRange(currentRow, 3, worksheet).Text.ToString().Trim(), "yyyy-MMM-dd", null);
                    //changeInfo.DescriptionWas = "";
                    changeInfo.DescriptionNow = ExcelUtil.GetRange(currentRow + 4, 3, worksheet).Text.ToString().Trim();
                    //changeInfo.RicWas = "";
                    changeInfo.RicNow = ExcelUtil.GetRange(currentRow + 1, 3, worksheet).Text.ToString().Trim();
                    //changeInfo.ISINWas = "";
                    changeInfo.ISINNow = ExcelUtil.GetRange(currentRow + 7, 3, worksheet).Text.ToString().Trim();
                    //changeInfo.SecondWas = "";
                    changeInfo.SecondNow = ExcelUtil.GetRange(currentRow + 6, 3, worksheet).Text.ToString().Trim();
                    if (changeInfo.RicNow.EndsWith(".KS"))
                    {
                        changeInfo.Exchange = "KSC";
                    }
                    else if (changeInfo.RicNow.EndsWith(".KQ"))
                    {
                        changeInfo.Exchange = "KOE";
                    }
                    else if (changeInfo.RicNow.EndsWith(".KN"))
                    {
                        changeInfo.Exchange = "KNX";
                    }
                    changeInfo.Asset = type;
                    RightsAndcompamyWarrantddEventList.Add(changeInfo);
                    currentRow += 24;
                }
                else
                {
                    break;
                }
            }
            return RightsAndcompamyWarrantddEventList;
        }

        //get PEO, PRF, REIT, ETF, CEF, drop
        //Sample files: KR FM (Drop) Request_134000.KS (wef 2012-Mar-02).xls
        //KR FM (Right DROP) Request_ 027970_r.KS(wef 2012-MAR-07).xls
        //KR_FM (ETF Drop) Request_097730.KS (wef 2009-AUG-11).xls
        private List<AddDropChangeInfo> GetPeoPrfReitCefDropEventList(Worksheet worksheet, int lastUsedRow, string type)
        {
            List<AddDropChangeInfo> peoPrfReitCefDropEventList = new List<AddDropChangeInfo>();
            int currentRow = 3;
            while (currentRow <= lastUsedRow)
            {
                AddDropChangeInfo changeInfo = new AddDropChangeInfo();
                changeInfo.ChangeType = "Delete";
                changeInfo.Date = DateTime.ParseExact(ExcelUtil.GetRange(currentRow, 3, worksheet).Text.ToString().Trim(), "yyyy-MMM-dd", null);
                changeInfo.DescriptionWas = ExcelUtil.GetRange(currentRow+3, 3, worksheet).Text.ToString().Trim();
                //changeInfo.DescriptionNow = "";
                changeInfo.RicWas = ExcelUtil.GetRange(currentRow + 1, 3, worksheet).Text.ToString().Trim();
                //changeInfo.RicNow = "";
                changeInfo.ISINWas = ExcelUtil.GetRange(currentRow + 2, 3, worksheet).Text.ToString().Trim();
                //changeInfo.ISINNow = "";
                changeInfo.SecondWas = changeInfo.RicWas.Remove(changeInfo.RicWas.IndexOf('.'));
                //changeInfo.SecondNow = "";
                if (changeInfo.RicWas.EndsWith(".KS"))
                {
                    changeInfo.Exchange = "KSC";
                }
                else if (changeInfo.RicWas.EndsWith(".KQ"))
                {
                    changeInfo.Exchange = "KOE";
                }
                else if (changeInfo.RicNow.EndsWith(".KN"))
                {
                    changeInfo.Exchange = "KNX";
                }

                //TO DO: Depends on the file name
                changeInfo.Asset = type;
                peoPrfReitCefDropEventList.Add(changeInfo);

                currentRow+=7;
                while ((ExcelUtil.GetRange(currentRow, 1, worksheet).Text == null || ExcelUtil.GetRange(currentRow, 1, worksheet).Text.ToString().Trim()==string.Empty) && currentRow<lastUsedRow)
                {
                    currentRow++;
                }
            }
            return peoPrfReitCefDropEventList;
        }

        //get rights and company warrant drop event list
        //Sample files: KR FM (Right DROP) Request_ 027970_r.KS(wef 2012-MAR-07).xls 
        //KR FM (Company Warrant Drop) Request_068630W.KQ (wef 2011-Sep-28).xls  
        private List<AddDropChangeInfo> GetRightsAndCompanyWarrantDropEventList(Worksheet worksheet, int lastUsedRow, string type)
        {
            List<AddDropChangeInfo> rightsAndCompanyWarrantsDropEventList = new List<AddDropChangeInfo>();
            int currentRow = 3;
            while (currentRow <= lastUsedRow)
            {
                if (ExcelUtil.GetRange(currentRow, 3, worksheet).Text != null && ExcelUtil.GetRange(currentRow, 3, worksheet).Text.ToString() != string.Empty)
                {
                    AddDropChangeInfo changeInfo = new AddDropChangeInfo();
                    changeInfo.ChangeType = "Delete";
                    changeInfo.Date = DateTime.ParseExact(ExcelUtil.GetRange(currentRow, 3, worksheet).Text.ToString().Trim(), "yyyy-MMM-dd", null);
                    changeInfo.DescriptionWas = ExcelUtil.GetRange(currentRow + 3, 3, worksheet).Text.ToString().Trim();
                    //changeInfo.DescriptionNow = "";
                    changeInfo.RicWas = ExcelUtil.GetRange(currentRow + 1, 3, worksheet).Text.ToString().Trim();
                    //changeInfo.RicNow = "";
                    changeInfo.ISINWas = ExcelUtil.GetRange(currentRow + 2, 3, worksheet).Text.ToString().Trim();
                    //changeInfo.ISINNow = "";
                    Regex r = new Regex("\\d{8}");
                    Match m = r.Match(changeInfo.ISINWas);
                    changeInfo.SecondWas = m.Value;
                    //changeInfo.SecondNow = "";
                    if (changeInfo.RicWas.EndsWith(".KS"))
                    {
                        changeInfo.Exchange = "KSC";
                    }
                    else if (changeInfo.RicWas.EndsWith(".KQ"))
                    {
                        changeInfo.Exchange = "KOE";
                    }
                    else if (changeInfo.RicNow.EndsWith(".KN"))
                    {
                        changeInfo.Exchange = "KNX";
                    }
                    changeInfo.Asset = type;
                    rightsAndCompanyWarrantsDropEventList.Add(changeInfo);
                    currentRow += 11;
                }
                else
                { 
                    break;
                }
            }
            return rightsAndCompanyWarrantsDropEventList;
        }

        //get name change event list: 
        //Sample file: KR FM(Name Change)Korea FM_123410.KQ (wef 02-MAR-12).xlsx
        private List<AddDropChangeInfo> GetNameChangeEventList(Worksheet worksheet, int lastUsedRow)
        {
            List<AddDropChangeInfo> nameChangeEventList = new List<AddDropChangeInfo>();
            for (int i = 2; i <= lastUsedRow; i++)
            {
                AddDropChangeInfo changeInfo = new AddDropChangeInfo();
                changeInfo.ChangeType = "Change";
                changeInfo.Date = DateTime.ParseExact(ExcelUtil.GetRange(i, 2, worksheet).Text.ToString().Trim(), "dd-MMM-yy", null);
                changeInfo.DescriptionWas = ExcelUtil.GetRange(i, 11, worksheet).Text.ToString().Trim();
                changeInfo.DescriptionNow = ExcelUtil.GetRange(i, 12, worksheet).Text.ToString().Trim();
                changeInfo.RicWas = ExcelUtil.GetRange(i, 3, worksheet).Text.ToString().Trim();
                changeInfo.RicNow = ExcelUtil.GetRange(i, 4, worksheet).Text.ToString().Trim();
                changeInfo.ISINWas = ExcelUtil.GetRange(i, 5, worksheet).Text.ToString().Trim();
                changeInfo.ISINNow = ExcelUtil.GetRange(i, 6, worksheet).Text.ToString().Trim();
                changeInfo.SecondWas = ExcelUtil.GetRange(i, 7, worksheet).Text.ToString().Trim();
                changeInfo.SecondNow = ExcelUtil.GetRange(i, 8, worksheet).Text.ToString().Trim();
                if (changeInfo.RicWas.EndsWith(".KS"))
                {
                    changeInfo.Exchange = "KSC";
                }
                else if (changeInfo.RicWas.EndsWith(".KQ"))
                {
                    changeInfo.Exchange = "KOE";
                }
                else if (changeInfo.RicNow.EndsWith(".KN"))
                {
                    changeInfo.Exchange = "KNX";
                }

                string ric = changeInfo.RicNow.Remove(changeInfo.RicNow.IndexOf('.'));
                if(ric.EndsWith("5")||ric.EndsWith("7")||ric.EndsWith("9"))
                {
                    changeInfo.Asset = "PRF";
                }
                else
                {
                    changeInfo.Asset = "ORD";
                }
                nameChangeEventList.Add(changeInfo);
            }
            return nameChangeEventList;
        }
    }    
}
