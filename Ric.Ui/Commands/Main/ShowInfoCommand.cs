using System;
using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Main
{
    internal class ShowInfoCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public ShowInfoCommand(MainWindowViewModel viewModel)
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
            if (_viewModel.GridWidth)
            {
                _viewModel.CollapseGrid();
            }
            else
            {
                _viewModel.ShowGrid();
            }
        }

        #endregion
    }
}
