using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class RemoveScheduleCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public RemoveScheduleCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedSchedule != null);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            var toDelete = (from schedule in RunTimeContext.Context.DatabaseContext.Schedules
                            where schedule.Id == _viewModel.SelectedSchedule.Id
                            select schedule).Single();
            _viewModel.ScheduledList.Remove(toDelete);
            RunTimeContext.Context.DatabaseContext.Schedules.Remove(toDelete);
            RunTimeContext.Context.DatabaseContext.SaveChanges();
        }

        #endregion
    }
}
