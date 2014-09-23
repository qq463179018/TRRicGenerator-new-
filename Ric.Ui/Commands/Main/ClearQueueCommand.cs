using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using Ric.Db.Model;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Main
{
    internal class ClearQueueCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public ClearQueueCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (!_viewModel.TaskRunning
                    && _viewModel.QueueList.Count > 0);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
           _viewModel.QueueList = new ObservableCollection<Task>();
        }

        #endregion
    }
}
