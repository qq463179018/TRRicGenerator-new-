using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using Ric.Util;
using System.Windows.Forms;


namespace Ric.Core
{
    public class Core
    {

        [DllImport("user32.dll")]
        private static extern void GetWindowThreadProcessId(IntPtr hWnd, out int k);

        private String log_path = @"D:\HKRicTemplate";
        private String subFolder = "";
        private String logName = "";
        private String pdf_path = @"D:\HKPDFFile";

        public string Log_Path
        {
            get { return log_path; }
            set { log_path = value; }
        }

        public string SubFolder
        {
            get { return subFolder; }
            set { subFolder = value; }
        }

        public string LogName
        {
            get { return logName; }
            set { logName = value; }
        }

        public string PDF_Path
        {
            get { return pdf_path; }
            set { pdf_path = value; }
        }


        /**
        * Create local folder
        * Retrun: void
        * Parameter: String fullPath 
        */
        public void CreateDir(String fullPath)
        {

            try
            {
                if (!Directory.Exists(fullPath))
                {
                    DirectoryInfo dir = new DirectoryInfo(fullPath);
                    dir.Create();
                }
                else
                {
                    DeleteTempDir(fullPath);
                    DirectoryInfo dir = new DirectoryInfo(fullPath);
                    dir.Create();
                }
            }
            catch (Exception ex)
            {
                String errInfo = ex.ToString();
            }
        }

        /**
         * Delete local temp folder and all sub folders and files under it
         * Retrun: void
         * Parameter: String dir
         * 
         */
        public void DeleteTempDir(String dir)
        {

            try
            {
                if (Directory.GetDirectories(dir).Length == 0 && Directory.GetFiles(dir).Length == 0)
                {
                    Directory.Delete(dir);
                    return;
                }
                foreach (string var in Directory.GetDirectories(dir))
                {
                    DeleteTempDir(var);
                }
                foreach (string var in Directory.GetFiles(dir))
                {

                    File.SetAttributes(var, FileAttributes.Normal);
                    File.Delete(var);
                }
                Directory.Delete(dir);
            }
            catch (Exception ex)
            {
                String errInfo = ex.ToString();
            }
        }

        /**
         * Input err info when delete local folder fail
         * 
         */
        public void WriteLogFile(String info)
        {
            String fullpath = log_path + "\\" + subFolder + "\\" + logName;
            try
            {
                FileStream logFile = null;
                StreamWriter sw = null;

                if (File.Exists(fullpath))
                {
                    logFile = new FileStream(fullpath, FileMode.Open, FileAccess.Write);
                    //move file pointer to the end
                    logFile.Seek(0, SeekOrigin.End);
                    sw = new StreamWriter(logFile);
                    sw.WriteLine(info);


                }
                else
                {
                    logFile = new FileStream(fullpath, FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(logFile);
                    sw.WriteLine(info);
                }
                sw.Close();
                logFile.Close();
            }
            catch (Exception ex)
            {
                string errInfo = ex.ToString();
            }
        }

        /**
         * Additional method, kill excel process
         * Return   :void
         * Parameter:Microsoft.Office.Interop.Excel.Application excelApp
         */
        public void KillExcelProcess(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            IntPtr t = new IntPtr(excelApp.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();
        }

        /**
         * Additional method, update FM serial number
         * Retrun: void
         */
        public String UpdateFMSerialNumber(String fmSerialNumber)
        {
            int number = 0;
            if (fmSerialNumber.Substring(0, 1) != "0")
            {
                fmSerialNumber = (Convert.ToInt32(fmSerialNumber) + 1).ToString();
            }
            else if (fmSerialNumber.Substring(1, 1) != "0")
            {
                number = Convert.ToInt32(fmSerialNumber.Substring(1));
                if (number == 999)
                {
                    fmSerialNumber = (number + 1).ToString();
                }
                else
                {
                    fmSerialNumber = "0" + (number + 1).ToString();
                }
            }
            else if (fmSerialNumber.Substring(2, 1) != "0")
            {
                number = Convert.ToInt32(fmSerialNumber.Substring(2));
                if (number == 99)
                {
                    fmSerialNumber = "0" + (number + 1).ToString();
                }
                else
                {
                    fmSerialNumber = "00" + (number + 1).ToString();
                }
            }
            else
            {
                number = Convert.ToInt32(fmSerialNumber.Substring(3));
                if (number == 9)
                {
                    fmSerialNumber = "00" + (number + 1).ToString();
                }
                else
                {
                    fmSerialNumber = "000" + (number + 1).ToString();
                }
            }
            return fmSerialNumber;
        }//end UpdateFMSerialNumber

        /**
         * Additional method,download PDF file from web
         * Retrun: String
         * Parameter: String pdfUrl
         */
        private String PDFDownload(String pdfUrl, String ricCode)
        {
            WebClient pdfClient = new WebClient();
            String pdfFilePath = pdf_path + "\\" + subFolder + "\\" + ricCode + ".pdf";
            try
            {
                pdfClient.DownloadFile(pdfUrl, pdfFilePath);
                return PDFToTxt(ricCode, pdfFilePath);
            }
            catch (Exception ex)
            {
                String logerror = ex.ToString();
                return "PDFDownload Error";
            }
        }

        /**
         * Additional method,transfer PDF to TXT file
         * Retrun: String
         * Parameter: String fileName, String pdfFilePath
         */
        private String PDFToTxt(String ricCode, String pdfFilePath)
        {
            String command = "pdftotext.exe";
            String txtPath = pdf_path + "\\" + subFolder + "\\" + ricCode + ".txt";
            String parameters = "-layout -enc UTF-8 -q " + pdfFilePath + " " + txtPath;
            System.Diagnostics.Process.Start(command, parameters);

            return txtPath;
        }

        /**
         * Additional method, get gearing and premium from PDF
         * Retrun: void
         * Parameter: String pdfUrl, int position, HKRicTemplate hkRic
         */
        public HKRicTemplate PDFAnalysis(String pdfUrl, String ricCode)
        {
            HKRicTemplate hkRic = new HKRicTemplate();
            hkRic.gearStr = "0.00";
            hkRic.premiumStr = "0.00";
            int position = 0;
            String txtPath = PDFDownload(pdfUrl, ricCode);
            System.Threading.Thread.Sleep(3000);

            if (txtPath == "PDFDownload Error")
            {
                throw new Exception(txtPath);
            }
            else
            {
                try
                {
                    String gearStr = "";
                    String premiumStr = "";
                    String lineText = "";
                    StreamReader sr = new StreamReader(txtPath);
                    lineText = sr.ReadToEnd();
                    sr.Close();


                    //Get position of stock
                    int stockIndex = lineText.IndexOf("Stock code");
                    int stockLength = "Stock code".Length;
                    if (stockIndex < 0)
                    {
                        stockIndex = lineText.IndexOf("Stock Code");
                    }


                    string codeStr = SearchKeyValue(stockIndex, stockLength, lineText);

                    position = codeStr.IndexOf(ricCode) / 5 + 1;

                    //Get gearing value
                    int gearingIndex = lineText.IndexOf("\n Gearing* ");
                    int gearLength = "\n Gearing* ".Length;
                    if (gearingIndex < 0)
                    {
                        gearingIndex = lineText.IndexOf("\nGearing*");
                        gearLength = "\nGearing*".Length;
                    }
                    if (gearingIndex < 0)
                    {
                        gearingIndex = lineText.IndexOf("\n Gearing *");
                        gearLength = "\n Gearing *".Length;
                    }
                    if (gearingIndex < 0)
                    {
                        gearingIndex = lineText.IndexOf("\nGearing     *");
                        gearLength = "\n Gearing     *".Length;
                    }
                    if (gearingIndex < 0)
                    {
                        gearingIndex = lineText.IndexOf("\nGearing *");
                        gearLength = "\nGearing *".Length;
                    }

                    if (gearingIndex < 0)
                    {
                        gearingIndex = lineText.IndexOf("Gearing*");
                        gearLength = "Gearing*".Length;
                    }
                    gearStr = SearchKeyValue(gearingIndex, gearLength, lineText);

                    //Get premium value
                    int premiumIndex = lineText.IndexOf("Premium*");
                    int premiumLength = "Premium*".Length;
                    if (premiumIndex < 0)
                    {
                        premiumIndex = lineText.IndexOf("Premium *");
                        premiumLength = "Premium *".Length;
                    }

                    premiumStr = SearchKeyValue(premiumIndex, premiumLength, lineText);

                    String[] gearArr = gearStr.Split('x');
                    String[] premiumArr = premiumStr.Split('%');

                    hkRic.gearStr = gearArr[position - 1];
                    hkRic.premiumStr = premiumArr[position - 1];

                    return hkRic;

                }//end try
                catch (Exception ex)
                {
                    String errLog = ex.ToString();
                    WriteLogFile("PDF analysis failed for " + ricCode + "! Action: Need manually input gearing and premium ");
                    return hkRic;
                }
            }

        }//PDFAnalysis

        /**
         * Additional method, SearchKeyValue
         * Retrun: String
         * Parameter: int keyPosition, int keyLength, String sourceStr
         */
        private String SearchKeyValue(int keyPosition, int keyLength, String sourceStr)
        {
            StringBuilder valueStr = new StringBuilder();
            Char[] sourceChar = sourceStr.ToCharArray();
            int position = keyPosition + keyLength;
            int index = position;

            while (sourceChar[position] != '\r')
            {
                if (sourceChar[position] != ' ')
                {
                    valueStr.Append(sourceChar[position]);
                }
                position++;
            }
            return valueStr.ToString();
        }

        /**
         * Additional method, calculate date
         * Retrun: DateTime
         * Parameter: DateTime sDate, DateTime launchDate
         */
        public DateTime DateCalculate(DateTime sDate, DateTime launchDate, int holidayCount)
        {
            DateTime temp;
            if (holidayCount == 0)
            {
                if (sDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    temp = launchDate.AddDays(2);
                }
                else if (sDate.DayOfWeek == DayOfWeek.Monday)
                {
                    temp = launchDate.AddDays(3);
                }
                else
                {
                    temp = launchDate.AddDays(1);
                }
            }
            else
            {
                temp = launchDate.AddDays(holidayCount + 1);
            }

            return temp;
        }

        /// <summary>
        /// test method
        /// </summary>
        /// <param name="ricCode"></param>
        public void txtAnalysis(string ricCode)
        {
            string gearStr = "";
            string premiumStr = "";
            int position = 0;

            try
            {
                String txtPath = @"D:\HKPDFFile\CBBC\" + ricCode + ".txt";
                String lineText = "";
                StreamReader sr = new StreamReader(txtPath);
                lineText = sr.ReadToEnd();
                sr.Close();

                //Get position of stock
                int stockIndex = lineText.IndexOf("Stock code");
                int stockLength = "Stock code".Length;
                if (stockIndex < 0)
                {
                    stockIndex = lineText.IndexOf("Stock Code");
                }

                string codeStr = SearchKeyValue(stockIndex, stockLength, lineText);

                position = codeStr.IndexOf(ricCode) / 5 + 1;

                //Get gearing value
                int gearingIndex = lineText.IndexOf("\n Gearing* ");
                int gearLength = "\n Gearing* ".Length;
                if (gearingIndex < 0)
                {
                    gearingIndex = lineText.IndexOf("\nGearing*");
                    gearLength = "\nGearing*".Length;
                }

                gearStr = SearchKeyValue(gearingIndex, gearLength, lineText);

                //Get premium value
                int premiumIndex = lineText.IndexOf("Premium*");
                int premiumLength = "Premium*".Length;
                if (premiumIndex < 0)
                {
                    premiumIndex = lineText.IndexOf("premium*");
                }

                premiumStr = SearchKeyValue(premiumIndex, premiumLength, lineText);

                String[] gearArr = gearStr.Split('x');
                String[] premiumArr = premiumStr.Split('%');

                gearStr = gearArr[position - 1];
                premiumStr = premiumArr[position - 1];

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
        }//end txtAnalysis

    }
}
