using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class MoveUpQueueCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public MoveUpQueueCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.QueueIndex != 0);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            
        }

        #endregion
    }
}
