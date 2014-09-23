using System;
using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class StartQueueCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public StartQueueCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.QueueList.Count > 0 && !_viewModel.TaskRunning);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.Start();
        }

        #endregion
    }
}
