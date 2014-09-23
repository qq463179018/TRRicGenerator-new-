using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Admin
{
    internal class ResetTaskCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public ResetTaskCommand(AdminWindowViewModel viewModel)
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
            //Find all the runs associated with the task
            var toDeleteRuns = (from run in RunTimeContext.Context.DatabaseContext.Runs
                where run.TaskId == _viewModel.SelectedTask.Id
                select run);
            // then delete
            RunTimeContext.Context.DatabaseContext.Runs.RemoveRange(toDeleteRuns);
            RunTimeContext.Context.DatabaseContext.SaveChanges();
        }

        #endregion
    }
}
