using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.ComponentModel;
using System.Threading;
using Ric.Util;
using Ric.Core;

namespace Ric.Tasks.HongKong
{

    public class HKSPIndexCheckConfig
    {
        [Category("LogPath")]
        [Description("This log file will record everyday info whether same or different.")]
        public SPIndexCheckLogPath SP_INDEX_CHECK_LOG_PATH { get; set; }

        [Category("ErrorLogPath")]
        [Description("Record the error information.")]
        public string LOG_FILE_PATH { get; set; }

        [Category("Save Path")]
        public string HKL_FILE_PATH { get; set; }

        [Category("Save Path")]
        public string HKG_FILE_PATH { get; set; }

        [Description("Please write the HKL & HKG data from the Xtra3000 into this file.")]
        public string XTRA3000_FILE_PATH { get; set; }
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class SPIndexCheckLogPath
    {
        public string LOG_PATH { get; set; }
        public string SUB_FOLDER { get; set; }
        public string FILE_NAME { get; set; }
    }

    public class HKSPIndexCheck : GeneratorBase
    {
        private const string spHKL = ".SPHKL";
        private const string spHKG = ".SPHKGEM";
        private List<string> hkLRemoveItems = new List<string>();
        private List<string> hkLAddItems = new List<string>();
        private List<string> hkGRemoveItems = new List<string>();
        private List<string> hkGAddItems = new List<string>();
        private Worksheet xtraHKLSheet = null;
        private Worksheet xtraHKGSheet = null;
        DateTime date = DateTime.Now;

        private static readonly string CONFIG_FILE_PATH = ".\\Config\\HK\\HK_SPIndexCheck.config";
        private static HKSPIndexCheckConfig configObj = null;
        //private static Logger logger = null;

        protected override void Start()
        {
            StartCodeIndexCheck();
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = ConfigUtil.ReadConfig(CONFIG_FILE_PATH, typeof(HKSPIndexCheckConfig)) as HKSPIndexCheckConfig;

            //logger = new Logger(configObj.LOG_FILE_PATH, Logger.LogMode.New);
        }

        public void StartCodeIndexCheck()
        {
            //Core coreObj = new Core();
            //coreObj.Log_Path = configObj.SP_INDEX_CHECK_LOG_PATH.LOG_PATH;
            //coreObj.SubFolder = configObj.SP_INDEX_CHECK_LOG_PATH.SUB_FOLDER;
            //coreObj.LogName = configObj.SP_INDEX_CHECK_LOG_PATH.FILE_NAME;
            string messageInfo = "";

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            string hklUrl = "ftp://Re2512:eR2180@edx.standardandpoors.com/Inbox/" + TimeUtil.GetYYYYMMDD(date) + "_SPHKL_ADJ.SPC";
            string hkgUrl = "ftp://Re2512:eR2180@edx.standardandpoors.com/Inbox/" + TimeUtil.GetYYYYMMDD(date) + "_SPHKG_ADJ.SPC";

            List<string> hkLRicCode = GetRicCode(xlApp, hklUrl);
            List<string> hkGRicCode = GetRicCode(xlApp, hkgUrl);
            if ((hkLRicCode != null && hkLRicCode.Count > 0) || (hkGRicCode != null && hkGRicCode.Count > 0))
            {
                ReadXtraInfoAndCompare(xlApp, hkLRicCode, hkGRicCode);

                if (hkLAddItems.Count == 0 && hkLRemoveItems.Count == 0 && hkGAddItems.Count == 0 && hkGRemoveItems.Count == 0)
                {
                    messageInfo = TimeUtil.GetYYYYMMDD(date) + "  " + "There is no added and removed data!";
                    //coreObj.WriteLogFile(messageInfo);
                    Logger.Log("There is no added and removed data!");
                }
                else
                {
                    if (hkLAddItems.Count > 0 || hkLRemoveItems.Count > 0)
                    {
                        GenerateHKLFile(xlApp);
                    }
                    if (hkGAddItems.Count > 0 || hkGRemoveItems.Count > 0)
                    {
                        GenerateHKGFile(xlApp);
                    }
                    messageInfo = TimeUtil.GetYYYYMMDD(date) + "  " + "Files Generated!";
                    //coreObj.WriteLogFile(messageInfo);
                    Logger.Log("Files Generated! Please go for Z:\\Hong Kong\\FM\\Today to check.");
                }
            }
            else
            {
                LogMessage("Can't get today's SPHKL_ADJ.SPC and SPHKG_ADJ.SPC from FTP!");
            }
        }

        public string ReturnCode(List<string> codeItems)
        {
            string code = "";
            if (codeItems != null && codeItems.Count > 0)
            {
                for (int index = 0; index < codeItems.Count; index++)
                {
                    code += " , " + "<" + codeItems[index] + ">";
                }
            }
            return code.Trim().TrimStart(',').Trim();
        }

        private void GenerateHKLFile(Microsoft.Office.Interop.Excel.Application xlApp)
        {
            string folderPath = configObj.HKL_FILE_PATH;
            //string folderPath = @"D:\zhang fan\SP";
            //string folderPath = @"Z:\Hong Kong\FM\Today";
            string fileName = "HK" + TimeUtil.shortYear + "-_CHANGE_S&P HKL Index_" + TimeUtil.GetFormatDate(date) + ".xls";
            Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
            wSheet.Cells[1, 1] = "Please action the following add/drop of HK stocks in 0#.SPHKL chain RIC on TQS.";
            wSheet.Cells[3, 1] = "FM Serial Number:";
            wSheet.Cells[3, 2] = "HK" + TimeUtil.shortYear;
            wSheet.Cells[4, 1] = "Effective Date:";
            wSheet.Cells[4, 2] = TimeUtil.GetEffectiveDate(date);
            wSheet.Cells[6, 1] = "Action Time:";
            wSheet.Cells[6, 2] = "Immediately";
            wSheet.Cells[8, 1] = "+Amendment of component Rics in Chain+";
            wSheet.Cells[9, 1] = "---------------------------------------------------------------------------------------------------";
            wSheet.Cells[10, 1] = "(1) <0#.SPHKL>";
            wSheet.Cells[11, 1] = "---------------------------------------------------------------------------------------------------";
            wSheet.Cells[12, 1] = "Chain Ric:";
            wSheet.Cells[12, 2] = "0#.SPHKL";
            wSheet.Cells[13, 1] = "Display name:";
            wSheet.Cells[13, 2] = "S&P/HKEx LC";
            wSheet.Cells[14, 1] = "Ric to be added:";
            wSheet.Cells[14, 2] = ReturnCode(hkLAddItems);
            wSheet.Cells[15, 1] = "Ric to be removed:";
            wSheet.Cells[15, 2] = ReturnCode(hkLRemoveItems);
            wSheet.Cells[16, 1] = "---------------------------------------------------------------------------------------------------";
            wBook.SaveAs(folderPath + "\\" + fileName, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            wBook.Close(Missing.Value, Missing.Value, Missing.Value);
        }

        private void GenerateHKGFile(Microsoft.Office.Interop.Excel.Application xlApp)
        {
            string folderPath = configObj.HKG_FILE_PATH;
            //string folderPath = @"D:\zhang fan\SP";
            //string folderPath = @"Z:\Hong Kong\FM\Today";
            string fileName = "HK" + TimeUtil.shortYear + "-_CHANGE_S&P GEM Index_" + TimeUtil.GetFormatDate(date) + ".xls";
            Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
            wSheet.Cells[1, 1] = "Please action the following add/drop of HK stocks in 0#.SPHKGEM chain RIC on TQS.";
            wSheet.Cells[3, 1] = "FM Serial Number:";
            wSheet.Cells[3, 2] = "HK" + TimeUtil.shortYear;
            wSheet.Cells[4, 1] = "Effective Date:";
            wSheet.Cells[4, 2] = TimeUtil.GetEffectiveDate(date);
            wSheet.Cells[6, 1] = "Action Time:";
            wSheet.Cells[6, 2] = "Immediately";
            wSheet.Cells[8, 1] = "+Amendment of component Rics in Chain+";
            wSheet.Cells[9, 1] = "---------------------------------------------------------------------------------------------------";
            wSheet.Cells[10, 1] = "(1) <0#.SPHKGEM>";
            wSheet.Cells[11, 1] = "---------------------------------------------------------------------------------------------------";
            wSheet.Cells[12, 1] = "Chain Ric:";
            wSheet.Cells[12, 2] = "0#.SPHKGEM";
            wSheet.Cells[13, 1] = "Display name:";
            wSheet.Cells[13, 2] = "S&P/HKEx GEM COMPONENTS";
            wSheet.Cells[14, 1] = "Ric to be added:";
            wSheet.Cells[14, 2] = ReturnCode(hkGAddItems);
            wSheet.Cells[15, 1] = "Ric to be removed:";
            wSheet.Cells[15, 2] = ReturnCode(hkGRemoveItems);
            wSheet.Cells[16, 1] = "---------------------------------------------------------------------------------------------------";
            wBook.SaveAs(folderPath + "\\" + fileName, XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            wBook.Close(Missing.Value, Missing.Value, Missing.Value);
        }

        private void CompareHKCode(Worksheet xtraSheet, List<string> ricCode, string hkLOrG)
        {
            int rowIndex = 1;
            int ricIndex = 0;
            string cellValue = "";
            Range range = ExcelUtil.GetRange(rowIndex, 1, xtraSheet);
            while (range.get_Value(Missing.Value) != null)
            {
                cellValue = range.get_Value(Missing.Value).ToString();
                if (ricIndex >= ricCode.Count)
                {
                    break;
                }
                if (cellValue.Equals(hkLOrG))
                    rowIndex++;
                else
                {
                    if (ricCode[ricIndex].Equals(cellValue))
                    {
                        ricIndex++;
                        rowIndex++;
                    }
                    else
                    {
                        List<string> romoveCodes = CheckIfRemove(cellValue, ricIndex, ricCode);
                        if (romoveCodes == null || romoveCodes.Count == 0)
                        {
                            AddItems(cellValue, hkLOrG);
                            rowIndex++;
                        }
                        else
                        {
                            RomoveItems(romoveCodes, 0, hkLOrG);
                            ricIndex = ricIndex + romoveCodes.Count;
                        }
                    }
                }
                range = ExcelUtil.GetRange(rowIndex, 1, xtraSheet);
            }

            if (range.get_Value(Missing.Value) != null)
            {
                while (range.get_Value(Missing.Value) != null)
                {
                    cellValue = range.get_Value(Missing.Value).ToString();
                    AddItems(cellValue, hkLOrG);
                    rowIndex++;
                    range = ExcelUtil.GetRange(rowIndex, 1, xtraSheet);
                }
            }
            if (ricIndex < ricCode.Count)
            {
                RomoveItems(ricCode, ricIndex, hkLOrG);
            }
        }

        private void AddItems(string value, string hkLOrG)
        {
            switch (hkLOrG)
            {
                case spHKL:
                    hkLAddItems.Add(value);
                    break;
                case spHKG:
                    hkGAddItems.Add(value);
                    break;
            }
        }

        private void RomoveItems(List<string> romoveCodes, int start, string hkLOrG)
        {
            for (int romoveIndex = start; romoveIndex < romoveCodes.Count; romoveIndex++)
            {
                switch (hkLOrG)
                {
                    case spHKL:
                        hkLRemoveItems.Add(romoveCodes[romoveIndex]);
                        break;
                    case spHKG:
                        hkGRemoveItems.Add(romoveCodes[romoveIndex]);
                        break;
                }
            }
        }

        private List<string> CheckIfRemove(string cellValue, int ricIndex, List<string> ricCode)
        {
            List<string> romoveItem = new List<string>();
            for (int i = ricIndex; i < ricCode.Count; i++)
            {
                if (ricCode[i].Equals(cellValue))
                {
                    break;
                }
                romoveItem.Add(ricCode[i]);

            }
            if (romoveItem.Count == (ricCode.Count - ricIndex))
            {
                romoveItem.Clear();
            }
            return romoveItem;
        }

        public string GetWebPage(string url)
        {
            string content = "";
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(url);
                request.Method = WebRequestMethods.Ftp.DownloadFile;

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream);
                content = reader.ReadToEnd();
                reader.Close();
                response.Close();

            }
            catch (Exception e)
            {
                if (e.ToString().Contains("Not Found"))
                {
                    LogMessage("  Not found today's file from FTP!:" + e);
                }
            }
            return content;
        }

        public Worksheet CopyTextIntoSheet(Worksheet wSheet, string content)
        {
            if (content != "")
            {
                //Clipboard.SetData("Text", content);
                //wSheet.Paste(Missing.Value, Missing.Value);

                string[] lines = content.Split('\n');
                int rowIndex = 1;
                foreach (string rowData in lines)
                {
                    int colIndex = 1;
                    string[] colData = rowData.Split('\t');
                    foreach (string colItem in colData)
                    {
                        wSheet.Cells[rowIndex, colIndex] = colItem;
                        colIndex++;
                    }
                    rowIndex++;
                }

            }
            return wSheet;
        }

        public List<string> GetRicCode(Microsoft.Office.Interop.Excel.Application xlApp, string url)
        {
            Workbook wBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
            string content = GetWebPage(url);

            wSheet = CopyTextIntoSheet(wSheet, content);
            List<string> ricCode = new List<string>();
            if (wSheet != null)
            {
                int usedRange = wSheet.UsedRange.Rows.Count;
                Range range = ExcelUtil.GetRange(7, 4, wSheet);
                if (range.get_Value(Missing.Value) != null)
                {
                    for (int rowIndex = 7; rowIndex < usedRange; rowIndex++)
                    {
                        range = ExcelUtil.GetRange(rowIndex, 4, wSheet);
                        if (range.get_Value(Missing.Value) == null || range.get_Value(Missing.Value).ToString().Equals(string.Empty))
                            break;

                        string cellValue = range.get_Value(Missing.Value).ToString();
                        ricCode.Add(cellValue);
                    }
                }
                ricCode.Sort();
            }
            wBook.Application.DisplayAlerts = false;
            wBook.Close(Missing.Value, Missing.Value, Missing.Value);
            return ricCode;
        }

        public void ReadXtraInfoAndCompare(Microsoft.Office.Interop.Excel.Application xlApp, List<string> hkLRicCode, List<string> hkGRicCode)
        {
            Workbook lBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            xtraHKLSheet = (Worksheet)lBook.Worksheets[1];
            Workbook gBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            xtraHKGSheet = (Worksheet)gBook.Worksheets[1];

            try
            {
                string fileName = configObj.XTRA3000_FILE_PATH;
                //string fileName = ".\\Config\\HK\\Xtra3000SPData.txt";
                StreamReader sr = new StreamReader(fileName);
                string text = sr.ReadToEnd();
                int hkLPos = text.IndexOf(".SPHKL");
                int hkGPos = text.IndexOf(".SPHKGEM");
                if (hkLPos < hkGPos)
                {
                    CopyTextIntoSheet(xtraHKLSheet, text.Substring(hkLPos, hkGPos - hkLPos));
                    CopyTextIntoSheet(xtraHKGSheet, text.Substring(hkGPos, text.Length - hkGPos));

                    //Clipboard.SetData("Text", text.Substring(hkLPos, hkGPos));
                    //xtraHKLSheet.Paste(Missing.Value, Missing.Value);
                    //Clipboard.SetData("Text", text.Substring(hkGPos, text.Length - hkGPos));
                    //xtraHKGSheet.Paste(Missing.Value, Missing.Value);

                }
                else
                {
                    CopyTextIntoSheet(xtraHKGSheet, text.Substring(hkGPos, hkLPos - hkGPos));
                    CopyTextIntoSheet(xtraHKLSheet, text.Substring(hkLPos, text.Length - hkLPos));

                    //Clipboard.SetData("Text", text.Substring(hkGPos, hkLPos));
                    //xtraHKGSheet.Paste(Missing.Value, Missing.Value);
                    //Clipboard.SetData("Text", text.Substring(hkLPos, text.Length - hkLPos));
                    //xtraHKLSheet.Paste(Missing.Value, Missing.Value);
                }

                CompareHKCode(xtraHKLSheet, hkLRicCode, spHKL);
                CompareHKCode(xtraHKGSheet, hkGRicCode, spHKG);
            }
            catch (Exception e)
            {
                LogMessage("error msg:" + e);
            }
            finally
            {
                lBook.Close(Missing.Value, Missing.Value, Missing.Value);
                gBook.Close(Missing.Value, Missing.Value, Missing.Value);
            }

        }
    }
}
