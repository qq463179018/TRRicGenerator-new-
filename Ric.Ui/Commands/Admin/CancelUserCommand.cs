﻿using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Developer
{
    internal class CancelUserCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public CancelUserCommand(AdminWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedUser != null
                && !_viewModel.IsUserReadOnly);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.IsUserReadOnly = true;
            _viewModel.SelectedTask = _viewModel.TaskList[_viewModel.TaskIndex];
        }

        #endregion
    }
}
