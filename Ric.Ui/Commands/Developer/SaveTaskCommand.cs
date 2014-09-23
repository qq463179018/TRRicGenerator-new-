using System;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;

namespace Ric.Ui.Commands.Developer
{
    internal class SaveTaskCommand : ICommand
    {
        private DeveloperWindowViewModel _viewModel;

        public SaveTaskCommand(DeveloperWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedTask != null
                && !_viewModel.IsReadOnly);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            var oldValueTask = (from task in RunTimeContext.Context.DatabaseContext.Tasks
                where task.Id == _viewModel.SelectedTask.Id
                select task).Single();

            oldValueTask.Name = _viewModel.Title;
            oldValueTask.Description = _viewModel.Description;
            oldValueTask.MarketId = _viewModel.TaskMarket.Id;
            oldValueTask.ManualTime = _viewModel.TaskManualTime;
            oldValueTask.ConfigType = _viewModel.ConfigurationType;
            oldValueTask.GeneratorType = _viewModel.TaskType;

            RunTimeContext.Context.DatabaseContext.SaveChanges();

            _viewModel.IsReadOnly = true;
            _viewModel.SetTaskList();
        }

        #endregion
    }
}
