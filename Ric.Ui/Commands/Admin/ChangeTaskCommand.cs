using System;
using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Admin
{
    internal class ChangeTaskCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public ChangeTaskCommand(AdminWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedTask != null
                && _viewModel.IsTaskReadOnly);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.IsTaskReadOnly = false;
        }

        #endregion
    }
}
