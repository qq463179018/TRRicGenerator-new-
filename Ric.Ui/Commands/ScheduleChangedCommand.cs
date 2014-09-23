using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Model;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class ScheduleChangedCommand : ICommand
    {
        private ScheduleWindowViewModel _viewModel;

        public ScheduleChangedCommand(ScheduleWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.TaskSchedule != null);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            _viewModel.SelectedSchedule = parameter as Schedule;
        }

        #endregion
    }
}
