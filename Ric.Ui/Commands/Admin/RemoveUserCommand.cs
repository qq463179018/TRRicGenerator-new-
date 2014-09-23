using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Admin
{
    internal class RemoveUserCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public RemoveUserCommand(AdminWindowViewModel viewModel)
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
            var toDelete = (from user in RunTimeContext.Context.DatabaseContext.Users
                            where user.Id == _viewModel.SelectedUser.Id
                            select user).Single();
            _viewModel.UserList.Remove(toDelete);
            RunTimeContext.Context.DatabaseContext.Users.Remove(toDelete);
            RunTimeContext.Context.DatabaseContext.SaveChanges();
            _viewModel.SelectedUser = _viewModel.UserList.First();
        }

        #endregion
    }
}
