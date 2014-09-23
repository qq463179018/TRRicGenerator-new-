using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Db.Model;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class ShowMarketTaskCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public ShowMarketTaskCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return true;
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.TaskList =
                new ObservableCollection<Task>(TaskManager.GetTaskByGroupMarket(_viewModel.CurrentUser,
                    RunTimeContext.Context.DatabaseContext));
            _viewModel.SelectedTask = _viewModel.TaskList.FirstOrDefault();
        }

        #endregion
    }
}
