using System;
using System.Windows;
using Ric.Core;
using Ric.Db.Model;

namespace Ric.Ui.View
{
    /// <summary>
    /// Interaction logic for ConfigWindow.xaml
    /// </summary>
    public partial class ConfigWindow
    {
        private Task selectedTask;

        public ConfigWindow(object config, Task task)
        {
            InitializeComponent();
            SetConfigWindow(config, task);
        }

        public void SetConfigWindow(object config, Task task)
        {
            selectedTask = task;
            PropertyGrid.SelectedObject = config;
            PropertyGrid.SelectedObjectType = Type.GetType(task.ConfigType);
            Title = String.Format("{0} configuration", selectedTask.Name);
        }

        /// <summary>
        /// todo transform in command
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Save(object sender, RoutedEventArgs e)
        {
            ConfigBuilder.UpdateConfigProperty(Type.GetType(selectedTask.ConfigType),
                        PropertyGrid.SelectedObject, RunTimeContext.Context.CurrentUser.Id, selectedTask.Id);

            RunTimeContext.Context.ConfigStore.StoreConfig(selectedTask.Id, PropertyGrid.SelectedObject);
            Xceed.Wpf.Toolkit.MessageBox.Show("Configuration successfully saved!");
            Close();
        }

        private void Button_Close(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
