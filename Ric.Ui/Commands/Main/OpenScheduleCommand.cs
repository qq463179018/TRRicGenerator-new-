using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Ric.Db.Manager;
using Ric.Ui.ViewModel;
using Ric.Core;
using Ric.Ui.View;
using Ric.Db.Model;

namespace Ric.Ui.Commands.Main
{
    internal class OpenScheduleCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public OpenScheduleCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedTask != null);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            var swd = new ScheduleWindow(new Schedule
            {
                Date = DateTime.Now.AddHours(2),
                Count = 1,
                Interval = 1,
                Frequency = ScheduleFrequency.Workday,
                TaskId = _viewModel.SelectedTask.Id,
                UserId = RunTimeContext.Context.CurrentUser.Id,
            });
            swd.ShowDialog();
            _viewModel.ScheduledList = new ObservableCollection<Schedule>(ScheduleManager.GetScheduledTask(RunTimeContext.Context.CurrentUser,
                RunTimeContext.Context.DatabaseContext));
        }

        #endregion
    }
}
