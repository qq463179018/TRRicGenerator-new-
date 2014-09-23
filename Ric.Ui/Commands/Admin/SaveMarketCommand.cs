using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Admin
{
    internal class SaveMarketCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public SaveMarketCommand(AdminWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedMarket != null);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            var oldValueMarket = (from market in RunTimeContext.Context.DatabaseContext.Markets
                where market.Id == _viewModel.SelectedMarket.Id
                select market).Single();

            // check if changes, not sure how to do that
            //if (oldValueTask != _viewModel.SelectedTask)
            //{

            oldValueMarket.Name = _viewModel.MarketName;
            oldValueMarket.ManagerId = _viewModel.MarketManagerUser.Id;

            RunTimeContext.Context.DatabaseContext.SaveChanges();

            //_viewModel.SetTaskList();
            //_viewModel.IsNeedSave = true;
            //_viewModel.IsRoTitle = true;
            //_viewModel.IsRoDescription = true;
            //}
        }

        #endregion
    }
}
