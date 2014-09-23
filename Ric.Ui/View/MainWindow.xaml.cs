using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;
using System.Xml;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Db.Model;
using Ric.Ui.ViewModel;
using System.Windows.Media;
using System.Deployment.Application;

namespace Ric.Ui.View
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindowViewModel Myvm { get; set; }

        public MainWindow(List<Task> tasks)
        {

#if DEBUG
            string path = @".\Themes\TestTheme.xaml";

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                // Read in ResourceDictionary File
                ResourceDictionary dic =
                   (ResourceDictionary)XamlReader.Load(fs);
                Application.Current.Resources = dic;
                // Clear any previous dictionaries loaded
                //Application.Current.Resources.MergedDictionaries.Clear();
                // Add in newly loaded Resource Dictionary
                //Application.Current.Resources.MergedDictionaries.Add(dic);
                //this.textBlock.Text = "RIC Generator - Production";
            }
#endif

            InitializeComponent();
            Myvm = new MainWindowViewModel(tasks, RunTimeContext.Context.CurrentUser.Group);
            DataContext = Myvm;
        }

        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
#if DEBUG
            this.textBlock.Text = "RIC Generator - Staging";
            this.textBlock.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#87A7DD"));
#endif

            lbVersion.Content = GetDeploymentVersion();
            //var tips = new TipsWindow();
            //tips.Show();
        }

        private string GetDeploymentVersion()
        {
            string version = string.Empty;
            try
            {
                version = string.Format("Version:{0}", ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString());
                return version;
            }
            catch
            {
                return "Version:";
            }
        }

        private void OpenConfigContextMenu_OnClick(object sender, RoutedEventArgs e)
        {
            var selectedTask = TaskList.SelectedItem as Task;

            object config = null;
            if (selectedTask == null) return;
            if (ConfigBuilder.IsConfigStoredInDB(Type.GetType(selectedTask.ConfigType)))
            {
                if ((config = RunTimeContext.Context.ConfigStore.GetConfig(selectedTask.Id)) == null)
                {
                    config = ConfigBuilder.CreateConfigInstance(Type.GetType(selectedTask.ConfigType), selectedTask.Id);
                }
            }
            var configWin = new ConfigWindow(config, selectedTask);
            configWin.ShowDialog();
        }

        private void ManagerItem_OnClick(object sender, RoutedEventArgs e)
        {
            var mwd = new ManagerWindow();
            mwd.ShowDialog();
        }

        private void AdminItem_OnClick(object sender, RoutedEventArgs e)
        {
            var awd = new AdminWindow();
            awd.ShowDialog();
        }

        private void DeveloperItem_OnClick(object sender, RoutedEventArgs e)
        {
            var dwd = new DeveloperWindow();
            dwd.ShowDialog();
        }

        private void ScheduleContextMenu_OnClick(object sender, RoutedEventArgs e)
        {
            var swd = new ScheduleWindow(new Schedule
            {
                Date = DateTime.Now.AddHours(2),
                Count = 1,
                Interval = 1,
                Frequency = ScheduleFrequency.Workday,
                TaskId = Myvm.SelectedTask.Id,
                UserId = RunTimeContext.Context.CurrentUser.Id,
            });
            swd.ShowDialog();
            Myvm.ScheduledList =
                new ObservableCollection<Schedule>(ScheduleManager.GetScheduledTask(RunTimeContext.Context.CurrentUser,
                    RunTimeContext.Context.DatabaseContext));
        }

        private void MainWindow_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void ResultList_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Process.Start(Myvm.SelectedResult.FilePath);
            }
            catch (Exception)
            {
                MessageBox.Show("Error happened when trying to open file");
            }
        }

        private void OnTaskListDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (Myvm.TaskRunning)
            {
                Myvm.QueueTask();
            }
            else
            {
                Myvm.StartTask();
            }
        }

        private void DatabaseItem_OnClick(object sender, RoutedEventArgs e)
        {
            var databaseWindow = new DatabaseWindow();
            databaseWindow.Show();
        }

        private void ConfigFolderItem_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                string configFolder = Path.Combine(Environment.CurrentDirectory, "Config");
                if (Directory.Exists(configFolder))
                    Process.Start(configFolder);
                else
                    MessageBox.Show(string.Format("{0} is not exist.", configFolder));
            }
            catch (Exception)
            {
                MessageBox.Show("Error happened when trying to open config folder.");
            }
        }

        private void ETIAutoFolderItem_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                string etiAutoFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ETI_Auto");
                if (Directory.Exists(etiAutoFolder))
                    Process.Start(etiAutoFolder);
                else
                    MessageBox.Show(string.Format("{0} is not exist.", etiAutoFolder));
            }
            catch (Exception)
            {
                MessageBox.Show("Error happened when trying to open ETI Auto folder");
            }
        }
    }
}