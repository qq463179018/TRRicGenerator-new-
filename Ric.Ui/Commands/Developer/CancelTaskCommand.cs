using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Developer
{
    internal class CancelTaskCommand : ICommand
    {
        private DeveloperWindowViewModel _viewModel;

        public CancelTaskCommand(DeveloperWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedTask != null
                && !_viewModel.IsReadOnly);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.IsReadOnly = true;
            _viewModel.SelectedTask = _viewModel.TaskList[_viewModel.TaskIndex];
        }

        #endregion
    }
}
