using System;
using System.Windows;
using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Main
{
    internal class ShowQueueCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public ShowQueueCommand(MainWindowViewModel viewModel)
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
            if (_viewModel.QueueVisibility == Visibility.Visible)
            {
                _viewModel.CollapseQueue();
            }
            else
            {
                _viewModel.ShowQueue();
            }
        }

        #endregion
    }
}
