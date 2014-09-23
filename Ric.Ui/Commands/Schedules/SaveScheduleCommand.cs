using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;
using MessageBox = Xceed.Wpf.Toolkit.MessageBox;
using Ric.Ui.View;

namespace Ric.Ui.Commands.Schedules
{
    internal class SaveScheduleCommand : ICommand
    {
        private ScheduleWindowViewModel _viewModel;

        public SaveScheduleCommand(ScheduleWindowViewModel viewModel)
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
            var toChange = (from schedule in RunTimeContext.Context.DatabaseContext.Schedules
                where schedule.Id == _viewModel.SelectedSchedule.Id
                select schedule).SingleOrDefault();

            string dayString = string.Empty;
            ScheduleWindow sw = (ScheduleWindow)parameter;
            if ((bool)sw.MoncheckBox.IsChecked)
            {
                dayString += "Mon,";    
            }
            if ((bool)sw.TuecheckBox.IsChecked)
            {
                dayString += "Tue,";
            }
            if ((bool)sw.WencheckBox.IsChecked)
            {
                dayString += "Wen,";
            }
            if ((bool)sw.ThucheckBox.IsChecked)
            {
                dayString += "Thu,";
            }
            if ((bool)sw.FricheckBox.IsChecked)
            {
                dayString += "Fri,";
            }
            if ((bool)sw.SatcheckBox.IsChecked)
            {
                dayString += "Sat,";
            }
            if ((bool)sw.SuncheckBox.IsChecked)
            {
                dayString += "Sun,";
            } 
            dayString = dayString.TrimEnd(new char[]{','});
            if (toChange != null)
            {
                toChange = _viewModel.SelectedSchedule;
            }
            else
            {
                _viewModel.SelectedSchedule.DayOfWeek = dayString;
                RunTimeContext.Context.DatabaseContext.Schedules.Add(_viewModel.SelectedSchedule);
            }
            RunTimeContext.Context.DatabaseContext.SaveChanges();
            MessageBox.Show("New schedule saved.");
            var window = parameter as Window;
            if (window != null) window.Close();
        }

        #endregion
    }
}
