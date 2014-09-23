using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Model;
using Ric.Ui.View;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class ChangeScheduleCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public ChangeScheduleCommand(MainWindowViewModel viewModel)
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
            ScheduleWindow swd = new ScheduleWindow(_viewModel.SelectedSchedule);
            swd.ShowDialog();
        }

        #endregion
    }
}
