using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
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
            if (toChange != null)
            {
                toChange = _viewModel.SelectedSchedule;
            }
            else
            {
                RunTimeContext.Context.DatabaseContext.Schedules.Add(_viewModel.SelectedSchedule);
            }
            RunTimeContext.Context.DatabaseContext.SaveChanges();
        }

        #endregion
    }
}
