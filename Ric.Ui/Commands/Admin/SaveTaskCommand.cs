﻿using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Admin
{
    internal class SaveTaskCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public SaveTaskCommand(AdminWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedTask != null
                && !_viewModel.IsTaskReadOnly);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            var oldValueTask = (from task in RunTimeContext.Context.DatabaseContext.Tasks
                where task.Id == _viewModel.SelectedTask.Id
                select task).Single();

            oldValueTask.Name = _viewModel.Title;
            oldValueTask.Description = _viewModel.Description;
            oldValueTask.MarketId = _viewModel.TaskMarket.Id;
            oldValueTask.OwnerId = _viewModel.Dev.Id;
            oldValueTask.Status = _viewModel.Taskstatus;
            oldValueTask.ManualTime = _viewModel.TaskManualTime;

            RunTimeContext.Context.DatabaseContext.SaveChanges();

            _viewModel.IsTaskReadOnly = true;
            //_viewModel.SetTaskList();
        }

        #endregion
    }
}
