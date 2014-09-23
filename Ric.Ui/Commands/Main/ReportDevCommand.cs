using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using Ric.Db.Model;
using Ric.Ui.View;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Main
{
    internal class ReportDevCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public ReportDevCommand(MainWindowViewModel viewModel)
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
            var errorwd = new ErrorWindow(_viewModel.SelectedTask);
            errorwd.ShowDialog();
        }

        #endregion
    }
}
