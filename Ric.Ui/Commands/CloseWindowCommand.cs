using System;
using System.Windows;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;
using Ric.Util;

namespace Ric.Ui.Commands
{
    internal class CloseWindowCommand : ICommand
    {

        public CloseWindowCommand()
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
            var window = parameter as Window;
            if (window != null) window.Close();
        }
        #endregion
    }
}
