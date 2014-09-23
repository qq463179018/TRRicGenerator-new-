using System;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Model;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class AddScheduleCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public AddScheduleCommand(MainWindowViewModel viewModel)
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
            //RunTimeContext.Context.DatabaseContext.Schedules.InsertOnSubmit(new Schedule
            //{
            //    Count = 0,
            //    Interval = 0,
            //    TaskId = _viewModel.TaskSchedule.Id,
            //    UserId = RunTimeContext.Context.CurrentUser.Id,
            //    Date = _viewModel.DateSchedule,
            //    Frequency = Db.Manager.ScheduleType.Workday
            //});
            //RunTimeContext.Context.DatabaseContext.SubmitChanges();
        }

        #endregion
    }
}
