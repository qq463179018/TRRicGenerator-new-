using System;
using System.Windows;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.View;
using Ric.Ui.ViewModel;
using Ric.Util;

namespace Ric.Ui.Commands
{
    internal class LogoutCommand : ICommand
    {

        public LogoutCommand()
        {

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
            //var loginWindow = new LoadingScreen(true);
            //loginWindow.Show();
            //var window = parameter as Window;
            //if (window != null) window.Close();
        }
        #endregion
    }
}
