using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Globalization;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using pdftron;
using PdfTronWrapper;
using Ric.Db.Manager;
using Ric.Core;
using Ric.Util;
namespace Ric.Tasks.Korea
{
    public class FurtherIssue : GeneratorBase
    {
        private Hashtable furtherIssue_hash = new Hashtable();
        private List<FurtherIssueModel> orgList = new List<FurtherIssueModel>();
        public string splitStyle = DateTime.Today.ToString("yyyy-MM-dd");
        private KOREA_ELWFM2AndFurtherIssuerGeneratorConfig configObj = null;
        private string downloadPdfFurtherIssuer = string.Empty;

        public void StartFurtherIssuerJob(KOREA_ELWFM2AndFurtherIssuerGeneratorConfig obj, List<TaskResultEntry> taskResultList, string downloadPdfPath)
        {
            downloadPdfFurtherIssuer = downloadPdfPath;
            configObj = obj;
            this.TaskResultList = taskResultList;
            PDFToTXT();
        }

        private void PDFToTXT()
        {
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ko-KR");
                string ipath = downloadPdfFurtherIssuer;


                DirectoryInfo dir = new DirectoryInfo(ipath);
                ArrayList list = new ArrayList();
                string txt = "";
                foreach (FileSystemInfo fsi in dir.GetFileSystemInfos())
                {
                    string fileName = "";
                    if (fsi is FileInfo)
                    {
                        fileName = fsi.Name;
                        if (fileName.Contains(".pdf"))
                            list.Add(fileName);
                    }
                    else
                    {
                        fileName = fsi.FullName;
                        // throw new Exception("else >>>>" + fileName);
                    }
                }

                string txtPath = Path.Combine(ipath, "TXT");
                if (!Directory.Exists(txtPath))
                {
                    Directory.CreateDirectory(txtPath);
                }

                foreach (var item in list)
                {
                    txt = TransferToTxT(item.ToString(), txtPath);
                    System.Threading.Thread.Sleep(1500);
                    GrabDataFromTXT(txt);
                }

                List<FurtherIssueModel> fList = FurtherIssueDataFormat();
                PrintFurtherIssueToExcel(fList);
                generateEmaCsv(fList);
            }
            catch (Exception ex)
            {
                string msg = "Error found in PDFToTXT()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
        }

        #region PDF to TEXT
        private string TransferToTxT(string name, string txtPath)
        {
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ko-KR");
                string pdfPath = downloadPdfFurtherIssuer + name;
                txtPath = Path.Combine(txtPath, name.Replace("pdf", "txt"));
                pdftronTransfer(pdfPath, txtPath);

            }
            catch (Exception ex)
            {
                string msg = "Error found in ThransferToTxT()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
            return txtPath;
        }
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
                    //sw.Write("******************\r\n");
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
        private static void pdftronTransfer(string pdfPath, string txtPath)
        {
            //pdftron.PDFNet.Initialize("Reuters Technology China Ltd.(thomsonreuters.com):CPU:1::W:AMC(20121010):AD5EE33F2505D1CAF1B425461F9C92BAA89204FA0AD8AAA17E07887EF0FA");
            TableLocator locator = new TableLocator(pdfPath);
            LocateConfiguration config = new LocateConfiguration();
            config.TableEndFirstLetterRegex = @"\d";
            config.TableEndRegex = @"\d{1,}/\d{1,}";
            config.TableNameNearbyFirstLetterRegex = "상";
            config.TableNameNearbyRegex = "상장일";
            List<TablePos> tablePosList = locator.GetMultiTablePos(".*?추가상장\\s*내역", config);
            List<FreeTable> tableList = new List<FreeTable>();
            foreach (var tablePos in tablePosList)
            {
                FreeTable table = TableExtractor.Extract(locator.pdfDoc, tablePos);
                tableList.Add(table);

            }
            WriteFreeTableToTxt(txtPath, " ", tableList);

        }
        #endregion

        private void GrabDataFromTXT(string txt)
        {
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture =
                   new System.Globalization.CultureInfo("en-US");

                string fullTxT = "";
                StreamReader sr = new StreamReader(txt);
                fullTxT = sr.ReadToEnd();
                sr.Dispose();
                sr.Close();

                int pageCount = Regex.Matches(fullTxT, "상장일\\s{0,}(\\d|-){1,}").Count;
                for (int i = 0; i < pageCount; i++)
                {
                    FurtherIssueModel fim1 = new FurtherIssueModel();
                    FurtherIssueModel fim2 = new FurtherIssueModel();
                    Get_Isin(fullTxT, fim1, fim2, i);
                    Get_Effective_Date(fullTxT, fim1, fim2, i);
                    Get_Ticker(fullTxT, fim1, fim2, i);
                    Get_Old_Quanity(fullTxT, fim1, fim2, i);
                    Get_New_Quanity(fullTxT, fim1, fim2, i);
                    orgList.Add(fim1);
                    orgList.Add(fim2);
                }
            }
            catch (Exception ex)
            {
                string msg = "Error found in GrabDataFromTXT()    : " + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
            }
        }

        private List<FurtherIssueModel> FurtherIssueDataFormat()
        {
            List<FurtherIssueModel> fList = new List<FurtherIssueModel>();
            for (var i = 0; i < orgList.Count; i++)
            {
                if (orgList[i].New_Isin == null)
                {
                    orgList.Remove(orgList[i]);
                    i--;
                }
                else
                {
                    FurtherIssueModel fim = new FurtherIssueModel();
                    fim.New_Isin = orgList[i].New_Isin;
                    fim.Old_Isin = orgList[i].New_Isin;
                    fim.Effective_Date = orgList[i].Effective_Date;
                    fim.Updated_Date = System.Convert.ToDateTime(DateTime.Today).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    fim.New_Ticker = orgList[i].New_Ticker.ToString().Substring(1, (orgList[i].New_Ticker.Length - 1));
                    fim.Old_Ticker = fim.New_Ticker;
                    fim.New_Ric = orgList[i].New_Ticker.ToString().Substring(1, (orgList[i].New_Ticker.Length - 1)) + ".KS";
                    fim.Old_Ric = fim.New_Ric;
                    string n_quantity = orgList[i].New_Quanity.Trim().ToString().Replace(",", "");
                    string o_quantity = orgList[i].Old_Quanity.Trim().ToString().Replace(",", "");
                    o_quantity = (System.Convert.ToInt32(n_quantity) - System.Convert.ToInt32(o_quantity)).ToString();
                    fim.New_Quanity = n_quantity;
                    fim.Old_Quanity = o_quantity;
                    fList.Add(fim);
                }
            }
            return fList;
        }

        private void PrintFurtherIssueToExcel(List<FurtherIssueModel> fList)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            ExcelApp excelApp = new ExcelApp(false, false);
            if (excelApp.ExcelAppInstance == null)
            {
                string msg = "Excel could not be started. Check that your office installation and project reference correct!!!";
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }

            try
            {
                string filename = "KR FM (Further Issue) _ ";
                foreach (var item in fList)
                    filename += item.New_Ric + ",";
                filename = filename.Trim(new char[] { ' ', ',' }) + "(wef " + Convert.ToDateTime(fList[0].Effective_Date).ToString("yyyy-MMM-dd", new CultureInfo("en-US")) + ").xls";

                //rename file if the length of file name more than 218 chars
                if (filename.Length > 218)
                    filename = "KR FM (Further Issue) _ (wef " + Convert.ToDateTime(fList[0].Effective_Date).ToString("yyyy-MMM-dd", new CultureInfo("en-US")) + ").xls";

                string ipath = Path.Combine(configObj.FM_FurtherIssuer, filename);
                Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(excelApp, ipath);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (wSheet == null)
                {
                    string msg = "Error found in PrintFurtherIssueToExcel :(WorkSheet could not be created. Check that your office installation and project reference correct!!!)";
                    Logger.Log(msg, Logger.LogType.Error);
                    return;
                }

                GenerateExcelFileTitle(wSheet);

                int startLine = 5;
                foreach (var item in fList)
                {
                    ((Range)wSheet.Cells[startLine, 1]).NumberFormat = "@";
                    wSheet.Cells[startLine, 1] = item.Updated_Date;
                    ((Range)wSheet.Cells[startLine, 2]).NumberFormat = "@";
                    wSheet.Cells[startLine, 2] = Convert.ToDateTime(item.Effective_Date).ToString("dd-MMM-yy", new CultureInfo("en-US"));
                    wSheet.Cells[startLine, 3] = item.Old_Ric;
                    wSheet.Cells[startLine, 4] = item.New_Ric;
                    wSheet.Cells[startLine, 5] = item.Old_Isin;
                    wSheet.Cells[startLine, 6] = item.New_Isin;
                    ((Range)wSheet.Cells[startLine, 7]).NumberFormat = "@";
                    wSheet.Cells[startLine, 7] = item.Old_Ticker;
                    ((Range)wSheet.Cells[startLine, 8]).NumberFormat = "@";
                    wSheet.Cells[startLine, 8] = item.New_Ticker;
                    wSheet.Cells[startLine, 9] = item.Old_Quanity;
                    wSheet.Cells[startLine, 10] = item.New_Quanity;
                    startLine++;
                }
                excelApp.ExcelAppInstance.AlertBeforeOverwriting = false;
                wBook.Save();
                AddResult(Path.GetFileNameWithoutExtension(ipath), ipath, "");
            }
            catch (Exception ex)
            {
                string msg = "Error found in PrintFurtherIssueToExcel : " + ex.ToString();
                Logger.Log(msg, Logger.LogType.Error);
                return;
            }
            finally
            {
                excelApp.Dispose();
            }
        }

        private static void GenerateExcelFileTitle(Worksheet wSheet)
        {
            if (wSheet.get_Range("D1", Type.Missing).Value2 == null)
            {
                ((Range)wSheet.Columns["A", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["B", System.Type.Missing]).ColumnWidth = 15;
                ((Range)wSheet.Columns["C", System.Type.Missing]).ColumnWidth = 13;
                ((Range)wSheet.Columns["D", System.Type.Missing]).ColumnWidth = 13;
                ((Range)wSheet.Columns["E", System.Type.Missing]).ColumnWidth = 16;
                ((Range)wSheet.Columns["F", System.Type.Missing]).ColumnWidth = 16;
                ((Range)wSheet.Columns["G", System.Type.Missing]).ColumnWidth = 13;
                ((Range)wSheet.Columns["H", System.Type.Missing]).ColumnWidth = 13;
                ((Range)wSheet.Columns["I", System.Type.Missing]).ColumnWidth = 24;
                ((Range)wSheet.Columns["J", System.Type.Missing]).ColumnWidth = 24;
                ((Range)wSheet.Columns["A:R", System.Type.Missing]).Font.Name = "Arial";
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Rows[4, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);

                ((Range)wSheet.Cells[3, 1]).Font.Bold = System.Drawing.FontStyle.Bold;
                ((Range)wSheet.Cells[3, 1]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                wSheet.Cells[3, 1] = "CHANGE";
                ((Range)wSheet.Cells[3, 1]).Font.Underline = System.Drawing.FontStyle.Underline;

                wSheet.Cells[4, 1] = "Updated Date";
                wSheet.Cells[4, 2] = "Effective Date";
                wSheet.Cells[4, 3] = "Old RIC";
                ((Range)wSheet.Cells[4, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                wSheet.Cells[4, 4] = "New RIC";
                ((Range)wSheet.Cells[4, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                wSheet.Cells[4, 5] = "Old ISIN";
                ((Range)wSheet.Cells[4, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                wSheet.Cells[4, 6] = "New ISIN";
                ((Range)wSheet.Cells[4, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                wSheet.Cells[4, 7] = "Old Ticker";
                ((Range)wSheet.Cells[4, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                wSheet.Cells[4, 8] = "New Ticker";
                ((Range)wSheet.Cells[4, 8]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
                wSheet.Cells[4, 9] = "Old Quanity  Of Warrant";
                ((Range)wSheet.Cells[4, 9]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);
                wSheet.Cells[4, 10] = "New Quanity  Of Warrant";
                ((Range)wSheet.Cells[4, 10]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Pink);
            }
        }

        private void Get_Ticker(string txtConetnt, FurtherIssueModel fim1, FurtherIssueModel fim2, int pageNum)
        {
            try
            {
                Regex r = new Regex("\\s{1,}단축코드\\s{1,}(?<ticker>.*?)\\s{0,}\r\n");
                MatchCollection mc = r.Matches(txtConetnt);
                string str_ticker = mc[pageNum].Groups["ticker"].Value;
                String[] ticker_array = str_ticker.Split(' ');
                string left_ticker = "";
                string right_ticker = "";
                if (ticker_array.Count() > 1)
                {
                    left_ticker = ticker_array[0].Trim().ToString();
                    right_ticker = ticker_array[(ticker_array.Count() - 1)].Trim().ToString();
                }
                else
                    left_ticker = ticker_array[0].Trim().ToString();

                if (left_ticker != string.Empty && right_ticker != string.Empty)
                {
                    fim1.New_Ticker = left_ticker;
                    fim2.New_Ticker = right_ticker;
                }
                else
                    fim1.New_Ticker = left_ticker;
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetTicker()    : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void Get_New_Quanity(string txtConetnt, FurtherIssueModel fim1, FurtherIssueModel fim2, int pageNum)
        {
            try
            {
                Regex r = new Regex("\\s{0,}누적발행수량\\s{1,}(?<NewQuanity>.*?)\\s{0,}\r\n");
                MatchCollection mc = r.Matches(txtConetnt);
                string str_new_quanity = mc[pageNum].Groups["NewQuanity"].Value;
                //int pos_new_quanity = arr[i].IndexOf("발   누적발행수량");
                //int pos_new_quanity_next = arr[i].IndexOf("행   기존추가발행총액합계");
                //string str_new_quanity = arr[i].Substring((pos_new_quanity + 20), (pos_new_quanity_next - pos_new_quanity - 20)).Trim().ToString();
                String[] new_quanity_array = str_new_quanity.Split(' ');
                string left_new_quanity = "";
                string right_new_quanity = "";
                if (new_quanity_array.Count() > 1)
                {
                    left_new_quanity = new_quanity_array[0].Trim().ToString();
                    right_new_quanity = new_quanity_array[(new_quanity_array.Count() - 1)].Trim().ToString();
                }
                else
                    left_new_quanity = new_quanity_array[0].Trim().ToString();

                if (left_new_quanity != string.Empty && right_new_quanity != string.Empty)
                {
                    fim1.New_Quanity = left_new_quanity;
                    fim2.New_Quanity = right_new_quanity;
                }
                else
                    fim1.New_Quanity = left_new_quanity;
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetNewQuantity()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void Get_Old_Quanity(string txtConetnt, FurtherIssueModel fim1, FurtherIssueModel fim2, int pageNum)
        {
            try
            {
                Regex r = new Regex("\\s{0,}추가발행수량\\s{1,}(?<OldQuantity>.*?)\\s{0,}\r\n");
                MatchCollection mc = r.Matches(txtConetnt);
                string str_old_quanity = mc[pageNum].Groups["OldQuantity"].Value;
                String[] old_quanity_array = str_old_quanity.Split(' ');
                string left_old_quanity = "";
                string right_old_quanity = "";
                if (old_quanity_array.Count() > 1)
                {
                    left_old_quanity = old_quanity_array[0].Trim().ToString();
                    right_old_quanity = old_quanity_array[(old_quanity_array.Count() - 1)].Trim().ToString();
                }
                else
                    left_old_quanity = old_quanity_array[0].Trim().ToString();

                if (left_old_quanity != string.Empty && right_old_quanity != string.Empty)
                {
                    fim1.Old_Quanity = left_old_quanity;
                    fim2.Old_Quanity = right_old_quanity;
                }
                else
                    fim1.Old_Quanity = left_old_quanity;
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetOldQuantity()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void Get_Effective_Date(string txtConetnt, FurtherIssueModel fim1, FurtherIssueModel fim2, int pageNum)
        {
            try
            {
                Regex r = new Regex("\\s{0,}상장일\\s{0,}(?<EffectiveDate>.*?)\\s{0,}\r\n");
                MatchCollection mc = r.Matches(txtConetnt);
                string str_effective_date = mc[pageNum].Groups["EffectiveDate"].Value;
                //int pos_effective_date = arr[i].IndexOf("      상장일");
                //int pos_effective_date_next = arr[i].IndexOf("\r\n    종목명");
                //string str_effective_date = arr[i].Substring((pos_effective_date + 12), (pos_effective_date_next - pos_effective_date - 12)).Trim().ToString();
                String[] effective_date_array = Regex.Split(str_effective_date, " ", RegexOptions.IgnoreCase);// str_effective_date.Split(' ');
                string left_effective_date = "";
                string right_effective_date = "";
                if (effective_date_array.Count() > 1)
                {
                    left_effective_date = effective_date_array[0].Trim().ToString();
                    right_effective_date = effective_date_array[(effective_date_array.Count() - 2)].Trim().ToString();
                }
                else
                    left_effective_date = effective_date_array[0].Trim().ToString();

                if (left_effective_date != string.Empty && right_effective_date != string.Empty)
                {
                    fim1.Effective_Date = left_effective_date;
                    fim2.Effective_Date = right_effective_date;
                }
                else
                    fim1.Effective_Date = left_effective_date;
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetEffectiveDate()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }

        private void Get_Isin(string txtConetnt, FurtherIssueModel fim1, FurtherIssueModel fim2, int pageNum)
        {
            try
            {
                Regex r = new Regex("\\s{0,}표준코드\\s{0,}(?<Isin>.*?)\\s{0,}\r\n");
                MatchCollection mc = r.Matches(txtConetnt);
                string str_isin = mc[pageNum].Groups["Isin"].Value;
                //int pos_Isin = arr[i].IndexOf("코   표준코드");
                //int pos_ticker = arr[i].IndexOf("드   단축코드");

                //string str_isin = arr[i].Substring((pos_Isin + 13), (pos_ticker - pos_Isin - 13)).Trim().ToString();
                String[] isin_array = str_isin.Split(' ');
                string left_isin = "";
                string right_isin = "";
                if (isin_array.Count() > 1)
                {
                    left_isin = isin_array[0].Trim().ToString();
                    right_isin = isin_array[(isin_array.Count() - 1)].Trim().ToString();
                }
                else
                    left_isin = isin_array[0].Trim().ToString();

                if (left_isin != string.Empty && right_isin != string.Empty)
                {
                    fim1.New_Isin = left_isin;
                    fim2.New_Isin = right_isin;
                }
                else
                    fim1.New_Isin = left_isin;
            }
            catch (Exception ex)
            {
                string msg = "Error found in GetISIN()   : \r\n" + ex.ToString();
                Logger.Log(msg, Logger.LogType.Warning);
            }
        }
        private string getEmaFileDir(List<FurtherIssueModel> fList)
        {
            string mainDir = ConfigureOperator.GetEmaFileSaveDir();
            string fileDir = DateTime.Now.ToString("yyyy-MM-dd", new CultureInfo("en-US"));
            string path = string.Format("{0}\\{1}", mainDir, fileDir);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }
        private MailToSend creatMail(List<FurtherIssueModel> fList)
        {
            MailToSend mail = new MailToSend();
            StringBuilder mailbodyBuilder = new StringBuilder();
            string filename = "KR FM (Further Issue) _ ";
            foreach (var item in fList)
                filename += item.New_Ric + ",";
            filename = filename.Trim(new char[] { ' ', ',' }) + "(wef " + Convert.ToDateTime(fList[0].Effective_Date).ToString("yyyy-MMM-dd", new CultureInfo("en-US")) + ").xls";
            string ipath = Path.Combine(configObj.FM_FurtherIssuer, filename);

            mail.MailSubject = filename;
            mailbodyBuilder.Append("Further Issue:   ");
            foreach (var item in fList)
            {
                mailbodyBuilder.Append(item.New_Ric + ",");
            }
            mailbodyBuilder.Append("\r\n");
            mailbodyBuilder.Append("Effective Date:  ");
            mailbodyBuilder.Append(fList.Count != 0 ? fList[0].Effective_Date : "");
            mailbodyBuilder.Append("\r\n");
            mail.MailBody = mailbodyBuilder.ToString();
            mail.MailBody += "\r\n\r\n\r\n\r\n";
            foreach (var term in configObj.FurtherIssuer_Signature)
            {
                mail.MailBody += term + "\r\n";
            }
            mail.ToReceiverList.AddRange(configObj.FurtherIssuer_MailTo);
            mail.CCReceiverList.AddRange(configObj.FurtherIssuer_MailCC);
            mail.AttachFileList.Add(ipath);
            return mail;

        }
        private int readLogic(string filePth)
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
        private void generateEmaCsv(List<FurtherIssueModel> fList)
        {
            string emaFileDir = getEmaFileDir(fList);
            //need to confirm
            string csvFileName = string.Format("{0}\\WRT_QUA_{1}_Korea.csv", emaFileDir, DateTime.Today.ToString("ddMMMyyyy", new CultureInfo("en-US")));
            List<List<string>> csvResList = new List<List<string>>();
            List<string> head = new List<string>();
            int line = 0;
            head.Add("Logical_Key");
            head.Add("Secondary_ID");
            head.Add("Secondary_ID_Type");
            head.Add("EH_Issue_Quantity");
            head.Add("Issue_Quantity");
            csvResList.Add(head);
            if (File.Exists(csvFileName))
            {
                line = readLogic(csvFileName);
            }
            for (int i = 0; i < fList.Count; i++)
            {
                List<string> term = new List<string>();
                term.Add((i + 1 + line).ToString());
                term.Add(fList[i].Old_Isin);
                term.Add("ISIN");
                term.Add("N");
                term.Add(fList[i].New_Quanity);
                csvResList.Add(term);
            }
            if (File.Exists(csvFileName))
            {
                csvResList.RemoveAt(0);//remove head
                OperateExcel.WriteToCSV(csvFileName, csvResList, FileMode.Append);
            }
            else
            {
                OperateExcel.WriteToCSV(csvFileName, csvResList, FileMode.Create);
            }
            AddResult(Path.GetFileNameWithoutExtension(csvFileName), csvFileName, "");


        }

    }
}
