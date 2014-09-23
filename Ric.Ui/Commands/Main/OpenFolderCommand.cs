using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Ric.Ui.ViewModel;
using Ric.Core;
using Ric.Ui.View;
using System.Diagnostics;
using System.IO;

namespace Ric.Ui.Commands.Main
{
    internal class OpenFolderCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public OpenFolderCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedResult != null);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            try
            {
                Process.Start(Path.GetDirectoryName(_viewModel.SelectedResult.FilePath));
            }
            catch (Exception)
            {
                MessageBox.Show("Error happened when trying to open file");
            }
        }

        #endregion
    }
}
