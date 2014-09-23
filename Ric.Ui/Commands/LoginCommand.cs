using System;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Model;
using Ric.Ui.View;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class LoginCommand : ICommand
    {
        private LoadingScreen _viewModel;

        public LoginCommand(LoadingScreen viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.LoginBox.Text.Length > 3);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.TryLogin(_viewModel.LoginBox.Text);
        }

        #endregion
    }
}
