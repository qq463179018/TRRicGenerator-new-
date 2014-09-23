using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Ric.Ui.ViewModel;
using Ric.Core;
using Ric.Ui.View;

namespace Ric.Ui.Commands.Main
{
    internal class OpenConfigCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public OpenConfigCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return (_viewModel.SelectedTask != null);
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            object config = null;

            if (ConfigBuilder.IsConfigStoredInDB(Type.GetType(_viewModel.SelectedTask.ConfigType)))
            {
                if ((config = RunTimeContext.Context.ConfigStore.GetConfig(_viewModel.SelectedTask.Id)) == null)
                {
                    config = ConfigBuilder.CreateConfigInstance(Type.GetType(_viewModel.SelectedTask.ConfigType), _viewModel.SelectedTask.Id);
                }
            }
            var configWin = new ConfigWindow(config, _viewModel.SelectedTask);
            configWin.ShowDialog();
        }

        #endregion
    }
}
