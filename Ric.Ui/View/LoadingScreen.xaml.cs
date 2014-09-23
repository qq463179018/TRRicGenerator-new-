using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;
using System.Xml;
using Ric.Core;
using Ric.Db;
using Ric.Db.Manager;
using Ric.Db.Model;
using Ric.Ui.Commands;

namespace Ric.Ui.View
{
    /// <summary>
    /// Interaction logic for LoadingScreen.xaml
    /// </summary>
    public partial class LoadingScreen
    {
        #region Properties

        private User _logonuser;
        private List<Task> _tasklist = new List<Task>();
        public Visibility LoginVisibility { get; set; }
        public bool AutomaticLogin { get; set; }
        public string Username { get; set; }

        #endregion

        #region Commands

        public ICommand LoginCommand { get; private set; }

        #endregion

        public LoadingScreen()
        {
            InitializeComponent();
            DataContext = this;
            LoginVisibility = Visibility.Hidden;
            AutomaticLogin = true;
            LoginCommand = new LoginCommand(this);
            Username = "";
        }

        public LoadingScreen(bool needLogin = false)
        {
            InitializeComponent();
            DataContext = this;
            LoginVisibility = Visibility.Hidden;
            if (needLogin)
            {
                LoginVisibility = Visibility.Visible;
                AutomaticLogin = false;
            }
            LoginCommand = new LoginCommand(this);
            Username = "";
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            _tasklist = TaskManager.GetTaskByGroup(_logonuser, RunTimeContext.Context.DatabaseContext);
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            LoadingBlock.Text += "Tasks loaded\r\n";
            var mainwd = new MainWindow(_tasklist);
            mainwd.Show();
            Close();
        }

        private void LoadingScreen_OnLoaded(object sender, RoutedEventArgs e)
        {
            if (AutomaticLogin)
                TryLogin();
        }

        public void TryLogin(string username = null)
        {
            try
            {
                RunTimeContext.Context.DatabaseContext = new EtiRicGeneratorEntities();
                if (username == null)
                {
                    LogonAsCurrentUser();
                }
                else
                {
                    LogonAsDifferentUser(username);
                }
                LoadingBlock.Text += "Welcome " + _logonuser + "\r\n";
                LoadingBlock.Text += "Loading Database\r\n";
                var backgroundWorker = new BackgroundWorker();
                backgroundWorker.DoWork += worker_DoWork;
                backgroundWorker.RunWorkerCompleted += worker_RunWorkerCompleted;
                backgroundWorker.RunWorkerAsync();
            }
            catch (Exception)
            {
                 //MessageBox.Show(ex.StackTrace);
                 //MessageBox.Show("Wrong username, please login again " + Environment.UserName);
                MessageBox.Show(
                    "You are not yet in the user database, please ask an Admin to add you if you want to be able to use" +
                    "Ric generator.\r\n\r\n" +
                    "Your Username is " + 
                    Environment.UserName);
            }
        }

        private void LogonAsDifferentUser(string username)
        {
            _logonuser = Auth.GetEtiUser(username, RunTimeContext.Context.DatabaseContext);
            RunTimeContext.Context.CurrentUser = _logonuser;
        }

        private void LogonAsCurrentUser()
        {
            _logonuser = Auth.GetEtiUser(RunTimeContext.Context.DatabaseContext);
            RunTimeContext.Context.CurrentUser = _logonuser;
        }
    }
}
