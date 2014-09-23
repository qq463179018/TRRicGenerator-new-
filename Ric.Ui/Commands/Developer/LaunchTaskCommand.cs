using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Model;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Developer
{
    internal class LaunchTaskCommand : ICommand
    {
        private DeveloperWindowViewModel _viewModel;

        public LaunchTaskCommand(DeveloperWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedTask != null &&
                _viewModel.SelectedTask.Status == TaskStatus.InDev);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            var toChangeTask = (from task in RunTimeContext.Context.DatabaseContext.Tasks
                where task.Id == _viewModel.SelectedTask.Id
                select task).SingleOrDefault();
            if (toChangeTask != null)
            {
                toChangeTask.Status = TaskStatus.Active;
            }
            RunTimeContext.Context.DatabaseContext.SaveChanges();
        }

        #endregion
    }
}
