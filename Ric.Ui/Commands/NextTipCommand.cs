using System;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;
using Ric.Util;

namespace Ric.Ui.Commands
{
    internal class NextTipCommand : ICommand
    {
        private TipWindowViewModel _viewModel;

        public NextTipCommand(TipWindowViewModel viewModel)
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
            
        }

        #endregion
    }
}
