using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Model;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Developer
{
    internal class CreateUserCommand : ICommand
    {
        private AdminWindowViewModel _viewModel;

        public CreateUserCommand(AdminWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (true);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            //
            // create a new basic user
            //
            RunTimeContext.Context.DatabaseContext.Users.Add(new User
            {
                WinUser = "Username",
                GedaUser = "ETIASIA",
                GedaPassword = "ETIASIA",
                Group = UserGroup.User,
                Status = UserStatus.Active,
                MainMarketId = 2,
                Email = "email@thomsonreuters.com",
                ManagerId = 1,
                Surname = "surname",
                Familyname = "familyName"
            });
            RunTimeContext.Context.DatabaseContext.SaveChanges();
            _viewModel.SetUserList();
            _viewModel.SelectedUser = _viewModel.UserList.Last();
            _viewModel.IsUserReadOnly = false;
        }

        #endregion
    }
}
