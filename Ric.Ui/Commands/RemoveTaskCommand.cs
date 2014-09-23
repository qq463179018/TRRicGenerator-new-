using System;
using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class RemoveTaskCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public RemoveTaskCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (!(_viewModel.QueueIndex == 0 && _viewModel.TaskRunning));
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.RemoveTaskFromQueue();
        }

        #endregion
    }
}
