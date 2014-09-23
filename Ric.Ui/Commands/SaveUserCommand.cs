using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands
{
    internal class SaveUserCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public SaveUserCommand(AdminWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedUser != null);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            var oldValueUser = (from user in RunTimeContext.Context.DatabaseContext.Users
                                where user.Id == _viewModel.SelectedUser.Id
                                select user).Single();

            // check if changes, not sure how to do that
            //if (oldValueTask != _viewModel.SelectedTask)
            //{

            oldValueUser.Familyname = _viewModel.UserFamilyname;
            oldValueUser.Surname = _viewModel.UserSurname;
            oldValueUser.Email = _viewModel.UserEmail;
            oldValueUser.Group = _viewModel.Usergroup;
            oldValueUser.MainMarketId = _viewModel.UserMainMarket.Id;

            RunTimeContext.Context.DatabaseContext.SaveChanges();
        }

        #endregion
    }
}
