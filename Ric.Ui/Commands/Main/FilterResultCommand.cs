using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Main
{
    internal class FilterResultCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public FilterResultCommand(MainWindowViewModel viewModel)
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
            _viewModel.ResultFilter = _viewModel.SelectedTask.Name;
            _viewModel.GridVisibility = Visibility.Visible;
        }

        #endregion
    }
}
