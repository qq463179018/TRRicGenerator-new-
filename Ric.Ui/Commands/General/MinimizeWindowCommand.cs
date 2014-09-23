using System;
using System.Windows;
using System.Windows.Input;

namespace Ric.Ui.Commands.General
{
    internal class MinimizeWindowCommand : ICommand
    {
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
            var window = parameter as Window;
            if (window != null) window.WindowState = WindowState.Minimized;
        }
        #endregion
    }
}
