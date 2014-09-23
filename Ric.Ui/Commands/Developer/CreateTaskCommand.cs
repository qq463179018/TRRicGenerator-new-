using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Model;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Developer
{
    internal class CreateTaskCommand : ICommand
    {
        private DeveloperWindowViewModel _viewModel;

        public CreateTaskCommand(DeveloperWindowViewModel viewModel)
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
            // todo
            // create a new basic task
            //
            var toAddTask = new Task
            {
                Name = "New task",
                Description = "test task",
                Status = TaskStatus.InDev,
                MarketId = 1,
                OwnerId = RunTimeContext.Context.CurrentUser.Id
            };
            RunTimeContext.Context.DatabaseContext.Tasks.Add(toAddTask);
            RunTimeContext.Context.DatabaseContext.SaveChanges();
            _viewModel.SetTaskList();
            _viewModel.SelectedTask = _viewModel.TaskList.Last();
        }

        #endregion
    }
}
