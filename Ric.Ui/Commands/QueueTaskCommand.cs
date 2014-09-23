using System;
using System.Linq;
using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class QueueTaskCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public QueueTaskCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (!(from task in _viewModel.QueueList
                        where task.Id == _viewModel.SelectedTask.Id
                        select task).Any()
                    && _viewModel.SelectedTask != null);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.QueueTask();
        }

        #endregion
    }
}
