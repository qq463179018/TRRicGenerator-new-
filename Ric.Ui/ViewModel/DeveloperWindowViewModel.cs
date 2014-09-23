using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Db.Model;
using Ric.Ui.Commands.Developer;
using System.Linq;
using Ric.Ui.Commands.General;

namespace Ric.Ui.ViewModel
{
    public class DeveloperWindowViewModel : BaseViewModel
    {
        #region properties

        #region Tasks

        private ObservableCollection<Task> _taskList;
        public ObservableCollection<Task> TaskList
        {
            get { return _taskList; }
            set
            {
                _taskList = value;
                NotifyPropertyChanged("TaskList");
            }
        }

        private ObservableCollection<Run> _runList;
        public ObservableCollection<Run> RunList
        {
            get { return _runList; }
            set
            {
                _runList = value;
                NotifyPropertyChanged("RunList");
            }
        }

        private Task _selectedTask;
        public Task SelectedTask
        {
            get { return _selectedTask; }
            set
            {
                _selectedTask = value;
                OnTaskChanged(value);
                NotifyPropertyChanged("SelectedTask");
            }
        }

        private int _taskIndex;
        public int TaskIndex
        {
            get { return _taskIndex; }
            set
            {
                if (_taskIndex != value)
                {
                    _taskIndex = value;
                    NotifyPropertyChanged("TaskIndex");
                }
            }
        }

        #endregion

        #region SelectedTask properties


        private string _description;
        public string Description
        {
            get { return _description; }
            set
            {
                if (_description != value)
                {
                    _description = value;
                    NotifyPropertyChanged("Description");
                }
            }
        }

        private string _title;
        public string Title
        {
            get { return _title; }
            set
            {
                if (_title != value)
                {
                    _title = value;
                    NotifyPropertyChanged("Title");
                }
            }
        }

        private string _configurationType;

        public string ConfigurationType
        {
            get { return _configurationType; }
            set
            {
                _configurationType = value;
                NotifyPropertyChanged("ConfigurationType");
            }
        }

        private string _taskType;

        public string TaskType
        {
            get { return _taskType; }
            set
            {
                _taskType = value;
                NotifyPropertyChanged("TaskType");
            }
        }
        private int _taskManualTime;

        public int TaskManualTime
        {
            get { return _taskManualTime; }
            set
            {
                _taskManualTime = value;
                NotifyPropertyChanged("TaskManualTime");
            }
        }

        private Market _taskMarket;
        public Market TaskMarket
        {
            get { return _taskMarket; }
            set
            {
                _taskMarket = value;
                NotifyPropertyChanged("TaskMarket");
            }
        }

        #endregion

        private ObservableCollection<Market> _marketList;
        public ObservableCollection<Market> MarketList
        {
            get { return _marketList; }
            set
            {
                _marketList = value;
                NotifyPropertyChanged("MarketList");
            }
        }

        private int boderTextBox;

        public int BoderTextBox
        {
            get { return boderTextBox; }
            set
            {
                boderTextBox = value;
                NotifyPropertyChanged("BoderTextBox");
            }
        }
        

        private bool isReadOnly;

        public bool IsReadOnly
        {
            get { return isReadOnly; }
            set
            {
                isReadOnly = value;
                BoderTextBox = value ? 0 : 1;
                IsEnable = !value;
                NotifyPropertyChanged("IsReadOnly");
            }
        }

        private bool isEnable;

        public bool IsEnable
        {
            get { return isEnable; }
            set
            {
                isEnable = value;
                NotifyPropertyChanged("IsEnable");
            }
        }

        #endregion

        #region Commands

        public ICommand ChangeTaskCommand { get; private set; }
        public ICommand SaveTaskCommand { get; private set; }
        public ICommand CancelSaveTaskCommand { get; private set; }

        public ICommand CreateTaskCommand { get; private set; }
        public ICommand LaunchTaskCommand { get; private set; }

        public ICommand CloseCommand { get; private set; }
        public ICommand MaximizeCommand { get; private set; }
        public ICommand MinimizeCommand { get; private set; }

        #endregion

        public DeveloperWindowViewModel()
        {
            MarketList = new ObservableCollection<Market>(MarketManager.GetAllMarkets(RunTimeContext.Context.DatabaseContext));
            SetTaskList();
            
            IsReadOnly = true;
            
            SelectedTask = TaskList.FirstOrDefault();

            ChangeTaskCommand = new ChangeTaskCommand(this);
            SaveTaskCommand = new SaveTaskCommand(this);
            CancelSaveTaskCommand = new CancelTaskCommand(this);

            CreateTaskCommand = new CreateTaskCommand(this);
            LaunchTaskCommand = new LaunchTaskCommand(this);

            CloseCommand = new CloseWindowCommand();
            MaximizeCommand = new MaximizeWindowCommand();
            MinimizeCommand = new MinimizeWindowCommand();
        }

        public void OnTaskChanged(Task newTask)
        {
            if (newTask != null)
            {
                Title = newTask.Name;
                Description = newTask.Description;
                ConfigurationType = newTask.ConfigType;
                TaskType = newTask.GeneratorType;
                TaskMarket = newTask.Market;
                TaskManualTime = newTask.ManualTime;
            }
        }

        public void SetTaskList()
        {
            TaskList = new ObservableCollection<Task>(TaskManager.GetTaskByOwner(RunTimeContext.Context.CurrentUser, RunTimeContext.Context.DatabaseContext));
        }
    }
}
