using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;

namespace Ric.Util
{

    public class ExcelApp : IDisposable
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private Application excelAppInstance = null;

        public Application ExcelAppInstance
        {
            get
            {
                return excelAppInstance;
            }
        }

        public ExcelApp(bool visible)
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            //Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("zh-HK");
            excelAppInstance = new Microsoft.Office.Interop.Excel.Application();
            excelAppInstance.Visible = visible;
        }

        public ExcelApp(bool visible, bool displayAlerts)
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            //Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("zh-HK");
            excelAppInstance = new Microsoft.Office.Interop.Excel.Application();
            excelAppInstance.Visible = visible;
            excelAppInstance.DisplayAlerts = displayAlerts;
        }

        public ExcelApp(bool visible, bool displayAlerts, System.Globalization.CultureInfo cultureInfo)
        {
            Thread.CurrentThread.CurrentCulture = cultureInfo;
            //Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("zh-HK");
            excelAppInstance = new Microsoft.Office.Interop.Excel.Application();
            excelAppInstance.Visible = visible;
            excelAppInstance.DisplayAlerts = displayAlerts;
        }


        #region IDisposable Members

        public void Dispose()
        {
            if (excelAppInstance != null)
            {
                int processId;
                GetWindowThreadProcessId(excelAppInstance.Hwnd, out processId);

                excelAppInstance.Quit();

                //  Thread.Sleep(5000);

                Process p = Process.GetProcessById(processId);

                if (p != null)
                {
                    try
                    {
                        p.Kill();
                    }
                    catch (Exception ex) { throw (ex); }
                }
            }
        }

        #endregion
    }

    public class ExcelLineWriter : IDisposable
    {
        public enum Direction
        {
            Down,
            Right
        }

        public Worksheet WorksheetInstance { get; private set; }
        public int Row { get; private set; }
        public int Col { get; private set; }
        public Direction DirectionType { get; private set; }

        public ExcelLineWriter(Worksheet worksheet, int row, int col, Direction direction)
        {
            Reset(worksheet, row, col, direction);
        }

        public void Reset(Worksheet worksheet, int row, int col, Direction direction)
        {
            if (worksheet == null)
            {
                throw (new Exception("Cannot create a writer for the worksheet is null."));
            }

            this.WorksheetInstance = worksheet;
            this.Row = row;
            this.Col = col;
            this.DirectionType = direction;
        }

        public void PlaceNext(int row, int col)
        {
            this.Row = row;
            this.Col = col;
        }

        public void PlaceNextAndWriteLine(int row, int col, object line)
        {
            PlaceNext(row, col);
            WriteLine(line);
        }

        public void MoveNext()
        {
            if (DirectionType == Direction.Down)
            {
                Row++;
            }
            else if (DirectionType == Direction.Right)
            {
                Col++;
            }
        }

        public void WriteInPlace(object line)
        {
            Range range = ExcelUtil.GetRange(Row, Col, this.WorksheetInstance);
            range.NumberFormat = "@";
            range.Value2 = line;
        }

        public void WriteLine(object line)
        {
            WriteInPlace(line);
            MoveNext();
        }

        public string ReadLineValue2()
        {
            Range range = ExcelUtil.GetRange(Row, Col, this.WorksheetInstance);
            if (range.Value2 != null)
            {
                MoveNext();
                return range.Value2.ToString().Trim();
            }
            else
            {
                MoveNext();
                return string.Empty;
            }
        }

        public string ReadLineCellText()
        {
            Range range = ExcelUtil.GetRange(Row, Col, this.WorksheetInstance);
            if (range.Text != null)
            {
                MoveNext();
                return range.Text.ToString().Trim();
            }
            else
            {
                MoveNext();
                return string.Empty;
            }
        }

        #region IDisposable Members

        public void Dispose()
        {
        }

        #endregion
    }

    public static class WorkbookExtension
    {
        public static List<List<string>> ToList(this Workbook workbook, int position = 1)
        {
            Worksheet worksheet = workbook.Worksheets[position] as Worksheet;
            Range last = GetLastRange(workbook, position);
            Range all = GetFullRange(workbook, position);
            int totalRow = last.Row;
            int totalColumn = last.Column;
            object[,] values = (object[,])all.Value2;
            List<List<string>> fullWorksheet = new List<List<string>>();

            for (int row = 1; row <= totalRow; row++)
            {
                List<string> lineList = new List<string>();
                for (int column = 1; column <= totalColumn; column++)
                {
                    if (Convert.ToString(values[row, column]) == null)
                    {
                        lineList.Add("");
                    }
                    else
                    {
                        lineList.Add(Convert.ToString(values[row, column]));
                    }
                }
                fullWorksheet.Add(lineList);
            }
            return fullWorksheet;
        }
        public static int GetLastColumn(this Workbook workbook, int position = 1)
        {
            Worksheet worksheet = workbook.Worksheets[position] as Worksheet;
            return GetLastRange(workbook, position).Column;
        }
        public static int GetLastRow(this Workbook workbook, int position = 1)
        {
            Worksheet worksheet = workbook.Worksheets[position] as Worksheet;
            return GetLastRange(workbook, position).Row;
        }
        public static Range GetLastRange(this Workbook workbook, int position = 1)
        {
            Worksheet worksheet = workbook.Worksheets[position] as Worksheet;
            return worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
        }
        public static Range GetFullRange(this Workbook workbook, int position = 1)
        {
            Worksheet worksheet = workbook.Worksheets[position] as Worksheet;
            Range last = GetLastRange(workbook, position);
            return worksheet.get_Range("A1", last);
        }
    }

    public class ExcelUtil
    {
        [DllImport("user32.dll")]
        private static extern void GetWindowThreadProcessId(IntPtr hWnd, out int k);

        public static void ExcelWorksheetToTsv(string excelFilePath, string worksheetName, string textFilePath)
        {
            using (var app = new ExcelApp(false, false))
            {
                var wb = ExcelUtil.CreateOrOpenExcelFile(app, excelFilePath);
                var ws = ExcelUtil.GetWorksheet(worksheetName, wb);
                ws.SaveAs(textFilePath, XlFileFormat.xlUnicodeText,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            }
        }

        public static Worksheet GetWorksheet(string filePath, string worksheetName)
        {
            Worksheet worksheet = null;
            Workbook workbook = null;
            using (ExcelApp app = new ExcelApp(false, false))
            {
                try
                {
                    workbook = CreateOrOpenExcelFile(app, filePath);

                    if (!string.IsNullOrEmpty(worksheetName))
                    {
                        worksheet = workbook.Worksheets[1] as Worksheet;
                    }

                    else
                    {
                        worksheet = GetWorksheet(worksheetName, workbook);
                    }

                    if (worksheet == null)
                    {
                        throw new Exception("Cannot find worksheet " + worksheetName);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Error happened when getting worksheet {0} from {1}. Exception message is: {2}", worksheetName, filePath, ex.Message));
                }
            }

            return worksheet;
        }
        public static Workbook CreateOrOpenExcelFile(ExcelApp excelApp, string path)
        {
            Workbook curWorkbook = null;
            path = Path.GetFullPath(path);
            try
            {
                if (File.Exists(path))
                {
                    curWorkbook = excelApp.ExcelAppInstance.Workbooks.Open(path, 0, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    return curWorkbook;
                }

                curWorkbook = excelApp.ExcelAppInstance.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                if (!Directory.Exists(Path.GetDirectoryName(path)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                }
                curWorkbook.SaveAs(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                     XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            }

            catch (Exception ex)
            {
                throw new Exception(string.Format("There's error when openning/ creating excel file {0}. Exception message: {1}", path, ex.Message));
            }

            return curWorkbook;
        }

        public static void InsertBlankRows(Range where, int rowCount)
        {
            try
            {
                while (rowCount-- > 0)
                {
                    where.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, Missing.Value);
                }
            }
            catch (Exception ex)
            {
                string errInfo = ex.ToString();
            }
        }

        public static void InsertBlankCols(Range where, int colCount)
        {
            while (colCount-- > 0)
            {
                where.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, Missing.Value);
            }
        }
        /// <summary>
        /// Delete entire rows
        /// </summary>
        /// <param name="where">Range where the rows will be deleted</param>
        /// <param name="rowCount">The number indicates how many rows will be deleted</param>
        public static void DeleteRows(Range where, int rowCount)
        {

            while (rowCount-- > 0)
            {
                where.EntireRow.Delete(XlInsertShiftDirection.xlShiftDown);
            }
        }

        public static void RangeDeleteRows(int row, int col, Worksheet ws)
        {
            Range cols = ExcelUtil.GetRange(row, col, ws);
            DeleteRows(cols, col);
        }

        public static Range GetRange(int row, int col, Worksheet worksheet)
        {
            return GetRange(row, col, row, col, worksheet);
        }

        public static Range GetRange(string rangeStr, Worksheet worksheet)
        {
            string[] frags = rangeStr.Split(':');
            if (frags.Length == 2)
            {
                return worksheet.get_Range(frags[0], frags[1]);
            }
            return worksheet.Range[rangeStr, rangeStr];
        }

        public static Range GetRange(int row1, int col1, int row2, int col2, Worksheet worksheet)
        {
            return worksheet.Range[worksheet.Cells[row1, col1], worksheet.Cells[row2, col2]];
        }

        public static Worksheet GetWorksheet(string name, Workbook workbook)
        {
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == name)
                {
                    return ws;
                }
            }
            return null;
        }

        public static Worksheet AddWorksheet(string name, Workbook workbook)
        {
            workbook.Worksheets.Add(Missing.Value, Missing.Value, 1, Missing.Value);
            Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
            worksheet.Name = name;
            return worksheet;
        }
    }
}
