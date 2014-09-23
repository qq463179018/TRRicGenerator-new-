using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Ric.Ui.View;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Log
{
    internal class OpenLogCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public OpenLogCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (true);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            //var logwindow = new LogWindows();
            //logwindow.Show();
        }

        #endregion
    }
}
