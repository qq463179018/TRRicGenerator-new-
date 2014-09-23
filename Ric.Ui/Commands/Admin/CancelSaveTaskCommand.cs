using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Admin
{
    internal class CancelSaveTaskCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public CancelSaveTaskCommand(AdminWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedTask != null
                && !_viewModel.IsTaskReadOnly);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.IsTaskReadOnly = true;
            _viewModel.SelectedTask = _viewModel.TaskList[_viewModel.TaskIndex];
        }

        #endregion
    }
}
