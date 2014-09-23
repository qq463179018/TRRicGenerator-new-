using System;
using System.IO;
using System.Linq;
using System.Windows.Input;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Ric.Core;
using Ric.Ui.ViewModel;
using Ric.Util;
using System.Drawing;
using Xceed.Wpf.Toolkit;

namespace Ric.Ui.Commands.Admin
{
    internal class CreateReportCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public CreateReportCommand(AdminWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return true;
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            string folderName = String.Empty;
            using (var fileDialog = new FolderBrowserDialog())
            {
                if (DialogResult.OK == fileDialog.ShowDialog())
                {
                    folderName = fileDialog.SelectedPath;
                }
            }

            var reportPath = Path.Combine(folderName, String.Format("RicReport_{0:MM-dd}.xlsx", DateTime.Now));
            var app = new ExcelApp(true, false);
            Workbook workbook = ExcelUtil.CreateOrOpenExcelFile(app, reportPath);
            Worksheet worksheetFirst = workbook.Worksheets[1] as Worksheet;
            worksheetFirst.Name = "general";
            worksheetFirst.Range["B3", "D3"].Merge();
            worksheetFirst.Cells[3, 2] = "ETI Ric Generator report";
            worksheetFirst.Cells[5, 2] = "from:";
            worksheetFirst.Cells[5, 3] = "to:";
            worksheetFirst.Cells[6, 2] = _viewModel.StartDate.ToString("MM/dd/yy");
            worksheetFirst.Cells[6, 3] = _viewModel.EndDate.ToString("MM/dd/yy");
            ((Range)worksheetFirst.Columns["B:C"]).ColumnWidth = 15;


            workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[1]);

            Worksheet worksheetSecond = workbook.Worksheets[2] as Worksheet;
            worksheetSecond.Name = _viewModel.MarketFilterIndex == 0 ? "All markets" : _viewModel.MarketFilter.Name;
            
            ((Range)worksheetSecond.Columns["A"]).ColumnWidth = 25;
            ((Range)worksheetSecond.Columns["B:C"]).ColumnWidth = 20;
            ((Range)worksheetSecond.Columns["D:G"]).ColumnWidth = 15;
            worksheetSecond.Cells[1, 1] = "Name";
            worksheetSecond.Cells[1, 2] = "Market";
            worksheetSecond.Cells[1, 3] = "Developer";
            worksheetSecond.Cells[1, 4] = "Success";
            worksheetSecond.Cells[1, 5] = "Fails";
            worksheetSecond.Cells[1, 6] = "Average time";
            worksheetSecond.Cells[1, 7] = "Success percent";

            int row = 2;

            foreach (var reportTask in
                     (from newReport in _viewModel.FilteredReport
                      let successPercent = (newReport.Successed + newReport.Failed)== 0 ? 0 : ((float)newReport.Successed / (float)(newReport.Successed + newReport.Failed)) * 100.0
                      orderby successPercent descending, newReport.Successed descending
                      select newReport))
            {
                worksheetSecond.Cells[row, 1] = reportTask.Task.Name;
                worksheetSecond.Cells[row, 2] = reportTask.Task.Market.Name;
                worksheetSecond.Cells[row, 3] = reportTask.Task.Owner.Surname + " " + reportTask.Task.Owner.Familyname;
                worksheetSecond.Cells[row, 4] = reportTask.Successed;
                worksheetSecond.Cells[row, 5] = reportTask.Failed;
                worksheetSecond.Cells[row, 6] = reportTask.AverageTime;
                worksheetSecond.Cells[row, 7] = (reportTask.Successed + reportTask.Failed)== 0 ? 0 : ((float)reportTask.Successed / (float)(reportTask.Successed + reportTask.Failed)) * 100.0;
                row++;
            }

            Range toTable = worksheetSecond.Range["A1", "G" + (row - 1)];

            toTable.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, toTable, System.Type.Missing, XlYesNoGuess.xlYes, System.Type.Missing).Name = "Table1";
            toTable.Select();
            //toTable.Worksheet.ListObjects["Table1"].TableStyle = "TableStyleMedium15";

            Range toCond = worksheetSecond.Range["G2", "G" + (row - 1)];

            FormatCondition cond1 = (FormatCondition)toCond.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlGreater, 66);
            cond1.Interior.PatternColorIndex = Constants.xlAutomatic;
            cond1.Interior.Color = ColorTranslator.ToWin32(Color.LawnGreen);

            FormatCondition cond2 = (FormatCondition)toCond.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlBetween, 34, 66);
            cond2.Interior.PatternColorIndex = Constants.xlAutomatic;
            cond2.Interior.Color = ColorTranslator.ToWin32(Color.Orange);

            FormatCondition cond3 = (FormatCondition)toCond.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, 33);
            cond3.Interior.PatternColorIndex = Constants.xlAutomatic;
            cond3.Interior.Color = ColorTranslator.ToWin32(Color.Red);

            //toTable.;
            worksheetFirst.Activate();
            workbook.Save();
            workbook.Close();
            app.Dispose();
        }

        #endregion
    }
}
