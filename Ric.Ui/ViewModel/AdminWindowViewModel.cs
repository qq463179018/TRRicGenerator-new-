using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows.Input;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Db.Model;
using Ric.Ui.Commands.Admin;
using Ric.Ui.Commands.Developer;
using Ric.Ui.Commands.General;
using ChangeTaskCommand = Ric.Ui.Commands.Admin.ChangeTaskCommand;
using Run = Ric.Db.Model.Run;
using SaveTaskCommand = Ric.Ui.Commands.Admin.SaveTaskCommand;

namespace Ric.Ui.ViewModel
{
    public class ReportTask
    {
        public Task Task { get; set; }
        public int Failed { get; set; }
        public int Successed { get; set; }
        public float AverageTime { get; set; }
        public float RunTime { get; set; }
    }

    public class AdminWindowViewModel : BaseViewModel
    {
        #region Properties

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

        private string _weekTimeSaved;

        public string WeekTimeSaved
        {
            get { return _weekTimeSaved; }
            set
            {
                _weekTimeSaved = value;
                NotifyPropertyChanged("WeekTimeSaved");
            }
        }

        private string _monthTimeSaved;

        public string MonthTimeSaved
        {
            get { return _monthTimeSaved; }
            set
            {
                _monthTimeSaved = value;
                NotifyPropertyChanged("MonthTimeSaved");
            }
        }

        private string _overallTimeSaved;

        public string OverallTimeSaved
        {
            get { return _overallTimeSaved; }
            set
            {
                _overallTimeSaved = value;
                NotifyPropertyChanged("OverallTimeSaved");
            }
        }

        private string _averageTime;

        public string AverageTime
        {
            get { return _averageTime; }
            set
            {
                _averageTime = value;
                NotifyPropertyChanged("AverageTime");
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

        private User _dev;
        public User Dev
        {
            get { return _dev; }
            set
            {
                _dev = value;
                NotifyPropertyChanged("Dev");
            }
        }

        private int _failedRuns;

        public int FailedRuns
        {
            get { return _failedRuns; }
            set
            {
                _failedRuns = value;
                NotifyPropertyChanged("FailedRuns");
            }
        }

        #endregion

        #region Users

        private ObservableCollection<User> _userList;
        public ObservableCollection<User> UserList
        {
            get { return _userList; }
            set
            {
                _userList = value;
                NotifyPropertyChanged("UserList");
            }
        }

        private User _selectedUser;
        public User SelectedUser
        {
            get { return _selectedUser; }
            set
            {
                _selectedUser = value;
                OnUserChanged(value);
                NotifyPropertyChanged("SelectedUser");
            }
        }

        private int _userIndex;
        public int UserIndex
        {
            get { return _userIndex; }
            set
            {
                if (_userIndex != value)
                {
                    _userIndex = value;
                    NotifyPropertyChanged("UserIndex");
                }
            }
        }

        #endregion

        #region SelectedUser properties

        private string _userSurname;
        public string UserSurname
        {
            get { return _userSurname; }
            set
            {
                if (_userSurname != value)
                {
                    _userSurname = value;
                    NotifyPropertyChanged("UserSurname");
                }
            }
        }

        private string _userFamilyname;
        public string UserFamilyname
        {
            get { return _userFamilyname; }
            set
            {
                if (_userFamilyname != value)
                {
                    _userFamilyname = value;
                    NotifyPropertyChanged("UserFamilyname");
                }
            }
        }

        private string _userEmail;
        public string UserEmail
        {
            get { return _userEmail; }
            set
            {
                if (_userEmail != value)
                {
                    _userEmail = value;
                    NotifyPropertyChanged("UserEmail");
                }
            }
        }

        private string _userWin;
        public string UserWin
        {
            get { return _userWin; }
            set
            {
                if (_userWin != value)
                {
                    _userWin = value;
                    NotifyPropertyChanged("UserWin");
                }
            }
        }

        private Market _userMainMarket;
        public Market UserMainMarket
        {
            get { return _userMainMarket; }
            set
            {
                _userMainMarket = value;
                NotifyPropertyChanged("UserMainMarket");
            }
        }

        private UserGroup _usergroup;

        public UserGroup Usergroup
        {
            get { return _usergroup; }
            set
            {
                _usergroup = value;
                NotifyPropertyChanged("Usergroup");
            }
        }

        public IEnumerable<UserGroup> UserGroupValues
        {
            get
            {
                return Enum.GetValues(typeof(UserGroup))
                    .Cast<UserGroup>();
            }
        }

        private TaskStatus _taskstatus;

        public TaskStatus Taskstatus
        {
            get { return _taskstatus; }
            set
            {
                _taskstatus = value;
                NotifyPropertyChanged("Taskstatus");
            }
        }

        public IEnumerable<TaskStatus> TaskStatusValues
        {
            get
            {
                return Enum.GetValues(typeof(TaskStatus))
                    .Cast<TaskStatus>();
            }
        }

        #endregion

        #region Market

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

        private Market _selectedMarket;
        public Market SelectedMarket
        {
            get { return _selectedMarket; }
            set
            {
                _selectedMarket = value;
                OnMarketChanged(value);
                NotifyPropertyChanged("SelectedMarket");
            }
        }

        private int _marketIndex;
        public int MarketIndex
        {
            get { return _marketIndex; }
            set
            {
                if (_marketIndex != value)
                {
                    _marketIndex = value;
                    NotifyPropertyChanged("MarketIndex");
                }
            }
        }

        #endregion

        #region SelectedMarket properties

        private string _marketName;
        public string MarketName
        {
            get { return _marketName; }
            set
            {
                if (_marketName != value)
                {
                    _marketName = value;
                    NotifyPropertyChanged("MarketName");
                }
            }
        }

        private User _marketManagerUser;
        public User MarketManagerUser
        {
            get { return _marketManagerUser; }
            set
            {
                _marketManagerUser = value;
                NotifyPropertyChanged("MarketManagerUser");
            }
        }

        #endregion

        #region Report

        private DateTime _startDate;

        public DateTime StartDate
        {
            get { return _startDate; }
            set
            {
                _startDate = value;
                NotifyPropertyChanged("StartDate");
                NotifyPropertyChanged("FilteredReport");
            }
        }

        private DateTime _endDate;

        public DateTime EndDate
        {
            get { return _endDate; }
            set
            {
                _endDate = value;
                NotifyPropertyChanged("EndDate");
                NotifyPropertyChanged("FilteredReport");
            }
        }

        private Market _marketFilter;

        public Market MarketFilter
        {
            get { return _marketFilter; }
            set
            {
                _marketFilter = value;
                NotifyPropertyChanged("MarketFilter");
            }
        }

        private int _marketFilterIndex;

        public int MarketFilterIndex
        {
            get { return _marketFilterIndex; }
            set
            {
                _marketFilterIndex = value;
                NotifyPropertyChanged("MarketFilterIndex");
                NotifyPropertyChanged("FilteredReport");
            }
        }

        private Task _taskFilter;

        public Task TaskFilter
        {
            get { return _taskFilter; }
            set
            {
                _taskFilter = value;
                NotifyPropertyChanged("TaskFilter");
            }
        }

        private int _taskFilterIndex;

        public int TaskFilterIndex
        {
            get { return _taskFilterIndex; }
            set
            {
                _taskFilterIndex = value;
                NotifyPropertyChanged("TaskFilterIndex");
                NotifyPropertyChanged("FilteredReport");
            }
        }

        private User _devFilter;

        public User DevFilter
        {
            get { return _devFilter; }
            set
            {
                _devFilter = value;
                NotifyPropertyChanged("DevFilter");
            }
        }

        private int _devFilterIndex;

        public int DevFilterIndex
        {
            get { return _devFilterIndex; }
            set
            {
                _devFilterIndex = value;
                NotifyPropertyChanged("DevFilterIndex");
                NotifyPropertyChanged("FilteredReport");
            }
        }

        public IEnumerable<Task> FilteredTasks
        {
            get
            {
                if (MarketFilterIndex == 0 && DevFilterIndex == 0)
                    return (from task in TaskList
                            select task);
                if (MarketFilterIndex == 0)
                    return (from task in TaskList
                            where task.OwnerId == DevFilter.Id
                            select task);
                if (DevFilterIndex == 0)
                    return (from task in TaskList
                            where task.MarketId == MarketFilter.Id
                            select task);
                return (from task in TaskList
                        where task.OwnerId == DevFilter.Id
                        where task.MarketId == MarketFilter.Id
                        select task);
            }
        }

        public IEnumerable<ReportTask> FilteredReport
        {
            get
            {
                return (from reportTask in FilteredTasks
                    let allRuns = (from run in RunTimeContext.Context.DatabaseContext.Runs 
                                   where run.TaskId == reportTask.Id
                                    && run.Date.Value.CompareTo(StartDate) > 0 
                                    && run.Date.Value.CompareTo(EndDate) < 0 
                                   select run)
                    select new ReportTask
                    {
                        Task = reportTask, 
                        Failed = (from run in allRuns 
                                  where run.Result == TaskResult.Fail 
                                  select run).Count(), 
                        Successed = (from run in allRuns 
                                     where run.Result == TaskResult.Success 
                                     select run).Count()
                    }).ToList();
            }
        }

        #endregion

        private bool _isTaskReadOnly;

        public bool IsTaskReadOnly
        {
            get { return _isTaskReadOnly; }
            set
            {
                _isTaskReadOnly = value;
                TaskBorderTextBox = value ? 0 : 1;
                IsTaskEnable = !value;
                NotifyPropertyChanged("IsTaskReadOnly");
            }
        }

        private bool _isTaskEnable;

        public bool IsTaskEnable
        {
            get { return _isTaskEnable; }
            set
            {
                _isTaskEnable = value;
                NotifyPropertyChanged("IsTaskEnable");
            }
        }


        private int _taskBorderTextBox;

        public int TaskBorderTextBox
        {
            get { return _taskBorderTextBox; }
            set
            {
                _taskBorderTextBox = value;
                NotifyPropertyChanged("TaskBorderTextBox");
            }
        }

        private bool _isUserReadOnly;

        public bool IsUserReadOnly
        {
            get { return _isUserReadOnly; }
            set
            {
                _isUserReadOnly = value;
                UserBorderTextBox = value ? 0 : 1;
                IsUserEnable = !value;
                NotifyPropertyChanged("IsUserReadOnly");
            }
        }

        private bool _isUserEnable;

        public bool IsUserEnable
        {
            get { return _isUserEnable; }
            set
            {
                _isUserEnable = value;
                NotifyPropertyChanged("IsUserEnable");
            }
        }

        private int _userBorderTextBox;

        public int UserBorderTextBox
        {
            get { return _userBorderTextBox; }
            set
            {
                _userBorderTextBox = value;
                NotifyPropertyChanged("UserBorderTextBox");
            }
        }

        private bool _isMarketReadOnly;

        public bool IsMarketReadOnly
        {
            get { return _isMarketReadOnly; }
            set
            {
                _isMarketReadOnly = value;
                MarketBorderTextBox = value ? 0 : 1;
                IsMarketEnable = !value;
                NotifyPropertyChanged("IsMarketReadOnly");
            }
        }

        private bool _isMarketEnable;

        public bool IsMarketEnable
        {
            get { return _isMarketEnable; }
            set
            {
                _isMarketEnable = value;
                NotifyPropertyChanged("IsMarketEnable");
            }
        }

        private int _marketBorderTextBox;

        public int MarketBorderTextBox
        {
            get { return _marketBorderTextBox; }
            set
            {
                _marketBorderTextBox = value;
                NotifyPropertyChanged("MarketBorderTextBox");
            }
        }


        private bool _isNeedSave;
        public bool IsNeedSave
        {
            get { return _isNeedSave; }
            set
            {
                if (_isNeedSave != value)
                {
                    _isNeedSave = value;
                    NotifyPropertyChanged("IsNeedSave");
                }
            }
        }

        private double _percentSuccess;

        public double PercentSuccess
        {
            get { return _percentSuccess; }
            set
            {
                _percentSuccess = value;
                NotifyPropertyChanged("PercentSuccess");
            }
        }

        public List<User> DevList { get; set; }

        #endregion

        #region Commands

        public ICommand SaveTaskCommand { get; private set; }
        public ICommand CancelSaveTaskCommand { get; private set; }
        public ICommand ChangeTaskCommand { get; private set; }

        public ICommand SaveMarketCommand { get; private set; }

        public ICommand ContactDeveloperCommand { get; private set; }
        public ICommand ResetTaskCommand { get; private set; }

        public ICommand CreateReportCommand { get; private set; }

        public ICommand CloseCommand { get; private set; }
        public ICommand MaximizeCommand { get; private set; }
        public ICommand MinimizeCommand { get; private set; }

        public ICommand CreateUserCommand { get; private set; }
        public ICommand ChangeUserCommand { get; private set; }
        public ICommand SaveUserCommand { get; private set; }
        public ICommand CancelUserCommand { get; private set; }
        public ICommand RemoveUserCommand { get; private set; }

        #endregion

        public AdminWindowViewModel()
        {
            RunList = new ObservableCollection<Run>();

            IsTaskReadOnly = true;
            IsUserReadOnly = true;
            IsMarketReadOnly = true;

            SaveTaskCommand = new SaveTaskCommand(this);
            SaveMarketCommand = new SaveMarketCommand(this);

            ChangeTaskCommand = new ChangeTaskCommand(this);
            CancelSaveTaskCommand = new CancelSaveTaskCommand(this);
            ContactDeveloperCommand = new ContactDeveloperCommand(this);
            ResetTaskCommand = new ResetTaskCommand(this);
            
            CreateReportCommand = new CreateReportCommand(this);

            CloseCommand = new CloseWindowCommand();
            MaximizeCommand = new MaximizeWindowCommand();
            MinimizeCommand = new MinimizeWindowCommand();

            CreateUserCommand = new CreateUserCommand(this);
            ChangeUserCommand = new ChangeUserCommand(this);
            SaveUserCommand = new SaveUserCommand(this);
            CancelUserCommand = new CancelUserCommand(this);
            RemoveUserCommand = new RemoveUserCommand(this);

            TaskIndex = 0;
            UserIndex = 0;
            MarketIndex = 0;

            DevList = new List<User>(UserManager.GetUserByGroup(UserGroup.Dev, RunTimeContext.Context.DatabaseContext));
            
            MarketList = new ObservableCollection<Market>(MarketManager.GetAllMarkets(RunTimeContext.Context.DatabaseContext));
            TaskList = new ObservableCollection<Task>(TaskManager.GetAllTasks(RunTimeContext.Context.DatabaseContext));
            SetUserList();
            
            SelectedTask = TaskList.FirstOrDefault();
            SelectedUser = UserList.FirstOrDefault();
            SelectedMarket = MarketList.FirstOrDefault();

            StartDate = DateTime.Now.AddMonths(-1);
            EndDate = DateTime.Now;

            TaskFilter = TaskList.FirstOrDefault();
            MarketFilter = MarketList.FirstOrDefault();
            DevFilter = DevList.FirstOrDefault();
        }

        public void OnTaskChanged(Task newTask)
        {
            if (newTask != null)
            {
                Title = newTask.Name;
                Description = newTask.Description;
                TaskMarket = newTask.Market;
                Taskstatus = newTask.Status;
                TaskManualTime = newTask.ManualTime;
                Dev = newTask.Owner;
                FailedRuns = StatsManager.GetFailedRuns(newTask, RunTimeContext.Context.DatabaseContext);
                PercentSuccess = (Convert.ToDouble(StatsManager.GetSuccessfullRunsByTask(newTask.Id, 100000, RunTimeContext.Context.DatabaseContext).Count))
                    / Convert.ToDouble(StatsManager.GetRunsByTask(newTask.Id, RunTimeContext.Context.DatabaseContext).Count)
                    * 100;
                WeekTimeSaved = TaskManager.GetTimeSaved(newTask, -7, RunTimeContext.Context.DatabaseContext).ToString(CultureInfo.InvariantCulture);
                MonthTimeSaved = TaskManager.GetTimeSaved(newTask, -30, RunTimeContext.Context.DatabaseContext).ToString(CultureInfo.InvariantCulture);
                OverallTimeSaved = TaskManager.GetTimeSaved(newTask, -100000, RunTimeContext.Context.DatabaseContext).ToString(CultureInfo.InvariantCulture);

                AverageTime = TaskManager.AverageRunningTime(newTask, RunTimeContext.Context.DatabaseContext).ToString(CultureInfo.InvariantCulture);
            }
        }

        public void OnUserChanged(User newUser)
        {
            if (newUser != null)
            {
                UserSurname = newUser.Surname;
                UserFamilyname = newUser.Familyname;
                UserEmail = newUser.Email;
                Usergroup = newUser.Group;
                UserMainMarket = newUser.MainMarket;
                UserWin = newUser.WinUser;
            }
        }

        public void OnMarketChanged(Market newMarket)
        {
            if (newMarket != null)
            {
                MarketManagerUser = newMarket.Manager;
                MarketName = newMarket.Name;
            }
        }

        public void SetUserList()
        {
            UserList = new ObservableCollection<User>(UserManager.GetAllUsers(RunTimeContext.Context.DatabaseContext));
        }
    }
}
