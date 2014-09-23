using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Ui.View;
using Ric.Ui.ViewModel;
using Ric.Db.Model;

namespace Ric.Ui.Commands
{
    internal class ChangeTaskListCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public ChangeTaskListCommand(MainWindowViewModel viewModel)
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
            //if (_viewModel.ShowOnlyMarketTasks)
            //{
            //    _viewModel.TaskList =
            //        new ObservableCollection<Task>(TaskManager.GetTaskByGroup(_viewModel.CurrentUser,
            //            RunTimeContext.Context.DatabaseContext));
            //    _viewModel.ShowOnlyMarketTasks = false;

            //}
            //else
            //{
            //    _viewModel.TaskList =
            //        new ObservableCollection<Task>(TaskManager.GetTaskByGroupMarket(_viewModel.CurrentUser,
            //            RunTimeContext.Context.DatabaseContext));
            //    _viewModel.ShowOnlyMarketTasks = true;
            //}
        }

        #endregion
    }
}
