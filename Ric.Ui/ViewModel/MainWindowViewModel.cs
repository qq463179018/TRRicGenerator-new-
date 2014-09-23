using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Ric.Core;
using Ric.Core.Events;
using Ric.Db.Manager;
using Ric.Db.Model;
using Ric.Ui.Commands;
using Ric.Ui.Commands.General;
using Ric.Ui.Commands.Log;
using Ric.Ui.Commands.Main;
using Ric.Ui.Model;
using Ric.Ui.View;

namespace Ric.Ui.ViewModel
{
    public class ResultTask
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string Type { get; set; }
        public string Task { get; set; }
    }

    public class ScheduledTaskInfo
    {
        public int Retry { get; set; }
        public int Interval { get; set; }
        public DateTime ScheduleTime { get; set; }
    }

    public class MainWindowViewModel : BaseViewModel
    {
        #region Properties

        #region User

        public User CurrentUser { get; set; }

        #endregion

        #region Queue

        private ObservableCollection<Task> _queueList;
        public ObservableCollection<Task> QueueList
        {
            get { return _queueList; }
            set
            {
                _queueList = value;
                NotifyPropertyChanged("QueueList");
            }
        }

        private int _queueIndex;
        public int QueueIndex
        {
            get { return _queueIndex; }
            set
            {
                _queueIndex = value;
                NotifyPropertyChanged("QueueIndex");
            }
        }

        private Visibility _queueVisibility;
        public Visibility QueueVisibility
        {
            get { return _queueVisibility; }
            set
            {
                _queueVisibility = value;
                NotifyPropertyChanged("QueueVisibility");
            }
        }

        #endregion

        #region Task

        public Task RunningTask { get; set; }

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

        public IEnumerable<Task> FilteredTasks
        {
            get
            {
                return (from task in TaskList
                        where task.Name.ToLower().Contains(TaskFilter.ToLower())
                            || Equals(task.Market.Name.ToLower(), TaskFilter.ToLower())
                            || Equals(task.Market.Abbreviation.ToLower(), TaskFilter.ToLower())
                        select task);
            }    
        } 

        private Task _selectedTask;

        public Task SelectedTask
        {
            get { return _selectedTask; }
            set
            {
                _selectedTask = value;
                NotifyPropertyChanged("SelectedTask");
            }
        }

        private int _taskIndex;
        public int TaskIndex
        {
            get { return _taskIndex; }
            set
            {
                _taskIndex = value;
                NotifyPropertyChanged("TaskIndex");
            }
        }

        private bool _taskRunning;

        public bool TaskRunning
        {
            get { return _taskRunning; }
            set
            {
                _taskRunning = value;
                NotifyPropertyChanged("TaskRunning");
            }
        }

        #endregion

        #region Results

        private ObservableCollection<ResultTask> _resultList;
        public ObservableCollection<ResultTask> ResultList
        {
            get { return _resultList; }
            set
            {
                _resultList = value;
                NotifyPropertyChanged("ResultList");
            }
        }

        private ResultTask _selectedResult;
        public ResultTask SelectedResult
        {
            get { return _selectedResult; }
            set
            {
                _selectedResult = value;
                NotifyPropertyChanged("SelectedResult");
            }
        }

        public IEnumerable<ResultTask> FilteredResults
        {
            get
            {
                return (from result in ResultList
                        where result.FileName.ToLower().Contains(ResultFilter.ToLower())
                            || Equals(result.Type.ToLower(), ResultFilter.ToLower())
                            || Equals(result.Task.ToLower(), ResultFilter.ToLower())
                            || Equals(result.FilePath.ToLower(), ResultFilter.ToLower())
                        select result);
            }
        } 
        #endregion

        #region Schedule

        private ObservableCollection<Schedule> _scheduleList;
        public ObservableCollection<Schedule> ScheduledList
        {
            get { return _scheduleList; }
            set
            {
                _scheduleList = value;
                NotifyPropertyChanged("ScheduledList");
            }
        }

        private Schedule _selectedSchedule;

        public Schedule SelectedSchedule
        {
            get { return _selectedSchedule; }
            set
            {
                _selectedSchedule = value;
                NotifyPropertyChanged("SelectedSchedule");
            }
        }

        private Dictionary<int, ScheduledTaskInfo> CurrentlyScheduledTasks { get; set; }


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

        #endregion

        #region Filters

        private string _taskFilter;
        public string TaskFilter
        {
            get { return _taskFilter; }
            set
            {
                _taskFilter = value;
                NotifyPropertyChanged("FilteredTasks");
                NotifyPropertyChanged("TaskFilter");
            }
        }

        private string _resultFilter;
        public string ResultFilter
        {
            get { return _resultFilter; }
            set
            {
                _resultFilter = value;
                NotifyPropertyChanged("FilteredResults");
                NotifyPropertyChanged("ResultFilter");
            }
        }

        #endregion

        #region Visibility

        private Visibility _isAdmin;

        public Visibility IsAdmin
        {
            get { return _isAdmin; }
            set
            {
                _isAdmin = value;
                NotifyPropertyChanged("IsAdmin");
            }
        }

        private Visibility _isManager;

        public Visibility IsManager
        {
            get { return _isManager; }
            set
            {
                _isManager = value;
                NotifyPropertyChanged("IsManager");
            }
        }

        private Visibility _isDev;

        public Visibility IsDev
        {
            get { return _isDev; }
            set
            {
                _isDev = value;
                NotifyPropertyChanged("IsDev");
            }
        }

        private Visibility _gridVisibility;

        public Visibility GridVisibility
        {
            get { return _gridVisibility; }
            set
            {
                _gridVisibility = value;
                NotifyPropertyChanged("GridVisibility");
            }
        }

        #endregion

        #region Log

        private ObservableCollection<Log> _log;

        public ObservableCollection<Log> Log
        {
            get { return _log; }
            set
            {
                _log = value;
                NotifyPropertyChanged("Log");
            }
        }

        #endregion

        #region Results

        public ObservableCollection<Dictionary<string, string>> Props { get; set; }

        public int FailedInARow { get; set; }

        private TaskResult _lastResult;

        #endregion

        private bool _gridWidth;
        public bool GridWidth
        {
            get { return _gridWidth; }
            set
            {
                _gridWidth = value;
                NotifyPropertyChanged("GridWidth");
            }
        }

        #endregion

        #region Commands

        // All commands, see in Commands/ folder
        //
        // Could use RelayCommand instead
        // like for RunCommand : 
        // public ICommand RunCommand = new RelayCommand(Start);
        // (would need to import Microsoft implementation of RelayCommand first)

        public ICommand MaximizeCommand { get; private set; }
        public ICommand MinimizeCommand { get; private set; }
        public ICommand RunCommand { get; private set; }
        public ICommand QueueCommand { get; private set; }
        public ICommand RemoveCommand { get; private set; }
        public ICommand CloseCommand { get; private set; }
        public ICommand ClearQueueCommand { get; private set; }
        public ICommand MoveUpQueueCommand { get; private set; }
        public ICommand MoveDownQueueCommand { get; private set; }
        public ICommand StartQueueCommand { get; private set; }
        public ICommand AddScheduleCommand { get; private set; }
        public ICommand RemoveScheduleCommand { get; private set; }
        public ICommand ChangeScheduleCommand { get; private set; }
        public ICommand ShowAllTaskCommand { get; private set; }
        public ICommand ShowMarketTaskCommand { get; private set; }
        public ICommand LogoutCommand { get; private set; }

        public ICommand ShowInfoCommand { get; private set; }
        public ICommand ShowQueueCommand { get; private set; }

        public ICommand FilterResultCommand { get; private set; }

        public ICommand OpenConfigCommand { get; private set; }
        public ICommand OpenScheduleCommand { get; private set; }

        public ICommand OpenFileCommand { get; private set; }
        public ICommand OpenFolderCommand { get; private set; }

        public ICommand OpenLogCommand { get; private set; }

        public ICommand ReportDevCommand { get; private set; }

        #endregion

        public MainWindowViewModel(List<Task> tasks, UserGroup group)
        {
            GridWidth = false;

            // Set the current user
            CurrentUser = RunTimeContext.Context.CurrentUser;

            #region Lists

            // The list of scheduled tasks (bind to the Schedule list in UI)
            ScheduledList = new ObservableCollection<Schedule>(ScheduleManager.GetScheduledTask(RunTimeContext.Context.CurrentUser,
                RunTimeContext.Context.DatabaseContext));

            // Result list (bind to the result panel in UI)
            ResultList = new ObservableCollection<ResultTask>();

            // Property list, to use with FileLib, /!\ not implemented now :(
            Props = new ObservableCollection<Dictionary<string, string>>();

            // Task list, with all the task, given by the loading screen as constructor parameter
            TaskList = tasks == null ? new ObservableCollection<Task>() : new ObservableCollection<Task>(tasks);

            // The list of queued tasks (bind to the queue panel in UI)
            QueueList = new ObservableCollection<Task>();

            // The list of markets (Bind to the combobox with the list of markets in UI)
            MarketList = new ObservableCollection<Market>(MarketManager.GetAllMarkets(RunTimeContext.Context.DatabaseContext));

            #endregion

            #region Selected

            // Set the market to the main market of the user
            SelectedMarket = CurrentUser.MainMarket;

            // Set the selected task to the first one of the list
            SelectedTask = FilteredTasks.FirstOrDefault();

            #endregion

            #region User right

            // Those three properties are used to be bind with some UI elements like menu
            // If you bind the IsVisibility property of a control with one of those, only the group of person with the appropriate
            // rights will be able to see it
            // eg : Visibility="{Binding IsAdmin}" <-- only Admin level people will see

            // Do the user have Admin Access ?
            IsAdmin = (@group == UserGroup.Dev || @group == UserGroup.Admin) ? Visibility.Visible : Visibility.Collapsed;

            // Manager ?
            IsManager = @group == UserGroup.Manager ? Visibility.Visible : Visibility.Collapsed;

            // Or maybe a Developer ?
            IsDev = @group == UserGroup.Dev ? Visibility.Visible : Visibility.Collapsed;

            #endregion

            #region Commands

            OpenConfigCommand = new OpenConfigCommand(this);
            OpenScheduleCommand = new OpenScheduleCommand(this);

            OpenFileCommand = new OpenFileCommand(this);
            OpenFolderCommand = new OpenFolderCommand(this);

            OpenLogCommand = new OpenLogCommand(this);

            RunCommand = new RunTaskCommand(this);
            QueueCommand = new QueueTaskCommand(this);
            RemoveCommand = new RemoveTaskCommand(this);
            CloseCommand = new CloseWindowCommand();
            MinimizeCommand = new MinimizeWindowCommand();
            MaximizeCommand = new MaximizeWindowCommand();
            LogoutCommand = new LogoutCommand();

            ClearQueueCommand = new ClearQueueCommand(this);
            MoveDownQueueCommand = new MoveDownQueueCommand(this);
            MoveUpQueueCommand = new MoveUpQueueCommand(this);

            StartQueueCommand = new StartQueueCommand(this);
            AddScheduleCommand = new AddScheduleCommand(this);
            RemoveScheduleCommand = new RemoveScheduleCommand(this);
            ChangeScheduleCommand = new ChangeScheduleCommand(this);
            ShowAllTaskCommand = new ShowAllTaskCommand(this);
            ShowMarketTaskCommand = new ShowMarketTaskCommand(this);

            ShowInfoCommand = new ShowInfoCommand(this);
            ShowQueueCommand = new ShowQueueCommand(this);

            FilterResultCommand = new FilterResultCommand(this);

            ReportDevCommand = new ReportDevCommand(this);

            #endregion

            // Number of time in a row a task failed
            // if 3 the error window will appear automatically
            FailedInARow = 0;

            #region Hide other regions

            GridWidth = true;
            //CollapseQueue();
            //CollapseGrid();

            #endregion

            // A simple bool to see if a task is running at the moment or not
            TaskRunning = false;

            // The result filter, chaging it will change automatically the task list
            ResultFilter = "";

            // The log List
            Log = new ObservableCollection<Log>();

            // The list of currently scheduled tasks, use for rerun if failed and keep track of current running status
            CurrentlyScheduledTasks = new Dictionary<int, ScheduledTaskInfo>();

            #region Timer

            // Create new DispatcherTimer and attach event handler to it
            var launchScheduleTimer = new DispatcherTimer();
            launchScheduleTimer.Tick += LaunchScheduleTimer_Tick;
            launchScheduleTimer.Interval = new TimeSpan(0, 1, 0);
            launchScheduleTimer.Start();

            #endregion
        }

        /// <summary>
        /// Function launched every Tick (every minute) to see if some task 
        /// was scheduled and need to be run 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LaunchScheduleTimer_Tick(object sender, EventArgs e)
        {
            // First looking for schedule tasks already ran but failed and need to re-run
            foreach (var toAdd in
                from schedule in CurrentlyScheduledTasks
                where schedule.Value.ScheduleTime.Hour == DateTime.Now.Hour
                      && schedule.Value.ScheduleTime.Minute == DateTime.Now.Minute
                select
                    new
                    {
                        toScheduleTask =
                            TaskManager.GetTaskById(schedule.Key, RunTimeContext.Context.DatabaseContext),
                        taskSchedule = schedule.Value
                    }

                )
            {
                SelectedTask = toAdd.toScheduleTask;
                QueueList.Add(toAdd.toScheduleTask);
                if (!TaskRunning)
                {
                    Start();
                }
            }

            // Now looking for task to add in the CurrentlyScheduledTasks list and run them
            foreach (var toAdd in
                from schedule in ScheduledList
                where schedule.Date.Hour == DateTime.Now.Hour
                      && schedule.Date.Minute == DateTime.Now.Minute
                select
                    new
                    {
                        toScheduleTask =
                            TaskManager.GetTaskById(schedule.TaskId, RunTimeContext.Context.DatabaseContext),
                        taskSchedule = schedule
                    }

                )
            {
                if (string.IsNullOrEmpty(toAdd.taskSchedule.DayOfWeek))
                {
                    toAdd.taskSchedule.DayOfWeek = string.Empty;
                }
                int dayOfWeekSign = 0;//schedule时间是否符合设定的week时间的标志位
                string[] dayString = toAdd.taskSchedule.DayOfWeek.Split(',');
                string nowDayOfWeek = DateTime.Now.DayOfWeek.ToString();
                nowDayOfWeek = nowDayOfWeek.Substring(0, 3);
                if (dayString.Length == 1 && string.IsNullOrEmpty(dayString[0]))
                {
                    dayOfWeekSign = 1;
                }
                foreach (string item in dayString)
                {
                    if (nowDayOfWeek.Equals(item))
                    {
                        dayOfWeekSign = 1;
                        break;
                    }
                }
                if (dayOfWeekSign == 0)
                {
                    break;
                }




                SelectedTask = toAdd.toScheduleTask;
                CurrentlyScheduledTasks.Add(toAdd.toScheduleTask.Id,
                    new ScheduledTaskInfo
                    {
                        Interval = toAdd.taskSchedule.Interval.Value,
                        Retry = toAdd.taskSchedule.Count.Value,
                        ScheduleTime = toAdd.taskSchedule.Date
                    });
                PrintMessage(String.Format("[Scheduler] {0} is scheduled to run at {1:HH:mm tt}, adding to queue now", toAdd.toScheduleTask.Name, toAdd.taskSchedule.Date));
                QueueList.Add(toAdd.toScheduleTask);
                if (!TaskRunning)
                {
                    Start();
                }
            }
        }

        /// <summary>
        /// Creates a new Background worker, attach event handler functions and run it
        /// </summary>
        public void Start()
        {
            var worker = new BackgroundWorker { WorkerReportsProgress = true };
            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerAsync();
        }

        #region Commands

        // Those functions explain themselves, just read the names

        internal void StartTask()
        {
            QueueTask();
            Start();
        }

        internal void QueueTask()
        {
            QueueList.Add(SelectedTask);
            ShowQueue();
        }

        internal void RemoveTaskFromQueue()
        {
            try
            {
                QueueList.RemoveAt(QueueIndex);
            }
            catch { }
        }

        #endregion

        #region Worker


        /// <summary>
        /// The main function of the worker, creates the Task instance and run it
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            //Task runningTask = e.Argument as Task;
            TaskRunning = true;
            var failed = false;
            while (QueueList.Count != 0)
            {
                var runningTask = QueueList.First();
                RunningTask = runningTask;
                var startTime = DateTime.Now;
                try
                {
                    (sender as BackgroundWorker).ReportProgress(0, new KeyValuePair<Task, string>(runningTask, ""));
                    var taskGenerator =
                        Activator.CreateInstance(Type.GetType(runningTask.GeneratorType)) as GeneratorBase;
                    taskGenerator.Log += Add_Message;
                    taskGenerator.Result += Add_Result;
                    taskGenerator.Props += Add_Prop;
                    taskGenerator.StartGenerator();
                    (sender as BackgroundWorker).ReportProgress(1, new KeyValuePair<Task, string>(runningTask, ""));
                }
                catch (Exception ex)
                {

                    var messageBuilder = new StringBuilder();
#if DEBUG
                    // In debug mode shows everything
                    if (ex.InnerException != null)
                    {
                        messageBuilder.Append("INNER ERROR MESSAGE:\r\n-------------------------------\r\n\r\n");
                        messageBuilder.Append(ex.InnerException.Message);
                        messageBuilder.Append("\r\n\r\nINNER STACK TRACE:\r\n-------------------------------\r\n\r\n");
                        messageBuilder.Append(ex.InnerException.StackTrace);
                        messageBuilder.Append("\r\n\r\n");
                    }
                    messageBuilder.Append("ERROR MESSAGE:\r\n-------------------------------\r\n\r\n");
                    messageBuilder.Append(ex.Message);
                    messageBuilder.Append("\r\n\r\nSTACK TRACE:\r\n-------------------------------\r\n\r\n");
                    messageBuilder.Append(ex.StackTrace);
#else
                    // In release just shows the exception message
                    messageBuilder.Append(ex.Message);
#endif
                    (sender as BackgroundWorker).ReportProgress(-1,
                        new KeyValuePair<Task, string>(runningTask, messageBuilder.ToString()));
                    failed = true;
                }
                finally
                {
                    _lastResult = failed ? TaskResult.Fail : TaskResult.Success;
                    var endTime = DateTime.Now;
                    var finishedTaskId = QueueList[0].Id;
                    e.Result = new Run
                    {
                        Date = DateTime.Now,
                        Duration = (int) (endTime - startTime).TotalSeconds,
                        Result = _lastResult,
                        TaskId = finishedTaskId,
                        UserId = CurrentUser.Id
                    };
                    PrintMessage(String.Format("[{0}] {1} in {2:mm\\mss\\sff}", QueueList[0].Name, _lastResult, (endTime - startTime)));

                    // If scheduled task see if failed
                    // so it need to be reschedule X minutes later Y times
                    if (CurrentlyScheduledTasks.ContainsKey(finishedTaskId))
                    {
                        if (_lastResult == TaskResult.Success)
                        {
                            CurrentlyScheduledTasks.Remove(finishedTaskId);
                            PrintMessage(String.Format("[Scheduler] {0} ran successfully, no retries are scheduled.", QueueList[0].Name));
                        }
                        else
                        {
                            CurrentlyScheduledTasks[finishedTaskId].Retry -= 1;
                            if (CurrentlyScheduledTasks[finishedTaskId].Retry == 0)
                            {
                                CurrentlyScheduledTasks.Remove(finishedTaskId);
                                PrintMessage(String.Format("[Scheduler] {0} task failed, no re-run will be attempted.", QueueList[0].Name), Logger.LogType.Warning);
                            
                            }
                            else
                            {
                                CurrentlyScheduledTasks[finishedTaskId].ScheduleTime = 
                                    CurrentlyScheduledTasks[finishedTaskId].ScheduleTime.Add(new TimeSpan(0,
                                        CurrentlyScheduledTasks[finishedTaskId].Interval, 0));
                                PrintMessage(String.Format("[Scheduler] {0} task failed, the task will re-run automatically in {1} minutes", QueueList[0].Name, CurrentlyScheduledTasks[finishedTaskId].Interval), Logger.LogType.Warning);
                            }
                        }
                    }

                    // Refresh Queue list
                    var newQueueTasks = new ObservableCollection<Task>(QueueList);
                    newQueueTasks.RemoveAt(0);
                    QueueList = newQueueTasks;

                }
            }
            TaskRunning = false; 
        }

        /// <summary>
        /// Call when task is running and change status
        /// mainly used to print messages depending the output
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
      
            var progress = e.ProgressPercentage;
            var pair = (KeyValuePair<Task, string>) (e.UserState);
            var message = pair.Value;

            switch (progress)
            {
                case 0: // Task starting
                    PrintMessage(String.Format("[{0}] Starting task", RunningTask.Name));
                    break;
                case 1: // Task Successed
                    FailedInARow = 0;
                    break;
                case -1: // Task Failed
                    if (FailedInARow++ == 2)
                    {
                        var errorwd = new ErrorWindow(pair.Key);
                        errorwd.ShowDialog();
                    }
#if DEBUG
                    // When debugging show a big messagebox
                    MessageBox.Show(message);
#else
                    // In release mode just show message in log panel
                    PrintMessage(String.Format("[{0}] Task failed with following error message: {1}", RunningTask.Name, message), Logger.LogType.Error);
#endif
                    break;
            }
        }


        /// <summary>
        /// Call when the task is completed
        /// Add a Run in the database
        /// If the user is a Dev just ignore it to avoid wrong statistics
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (CurrentUser.Group != UserGroup.Dev)
            {
                StatsManager.AddRunInfo(e.Result as Run, RunTimeContext.Context.DatabaseContext);
            }
        }

        #endregion

        #region Event handlers
        
        /// <summary>
        /// If no LogType is given, use Info by default
        /// </summary>
        /// <param name="message"></param>
        public void PrintMessage(string message)
        {
            PrintMessage(message, Logger.LogType.Info);
        }

        /// <summary>
        /// Add a given messge to the LogList to appear live in the log panel of the UI
        /// Gives a different color depending on the LogType, but this feature not used at the moment
        /// </summary>
        /// <param name="message"></param>
        /// <param name="type"></param>
        public void PrintMessage(string message, Logger.LogType type)
        {
            var showColor = String.Empty;
            if (type == Logger.LogType.Error)
            {
                showColor = "Red";
            }
            else if (type == Logger.LogType.Warning)
            {
                showColor = "Orange";
            }
            else
            {
                showColor = "Green";
            }
            Log = new ObservableCollection<Log>(Log) {new Log
            {
               Message = message, 
               ColorText = showColor
            }};    
        }

        /// <summary>
        /// Event handler when adding a log message
        /// Call PrintMessage to print the given message live to the log panel of the UI
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Add_Message(object sender, LogEventArgs e)
        {
            PrintMessage(String.Format("[{0}] {1}", RunningTask.Name, e.Message), e.LogType);
        }

        /// <summary>
        /// Event handler when task add result
        /// Add new ResultTask object in the list of result (will appear live on the MainWindow)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Add_Result(object sender, ResultEventArgs e)
        {
            var newResultList = new ObservableCollection<ResultTask>(ResultList)
            {
                new ResultTask 
                {
                    FilePath = e.FilePath, 
                    FileName = e.FileName, 
                    Type = e.Filetype, 
                    Task = RunningTask.Name
                }
            };
            ResultList = newResultList;
            ResultFilter = ResultFilter;
        }
        
        /// <summary>
        /// Event handler if the task add a property
        /// To use with FileLib
        /// /!\ Not implemented yet :(
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Add_Prop(object sender, PropEventArgs e)
        {
            var newPropList = new ObservableCollection<Dictionary<string, string>>(Props) {e.Props};
            Props = newPropList;
        }

        #endregion

        #region Change Visibility

        internal void ShowQueue()
        {
            QueueVisibility = Visibility.Visible;
        }

        internal void CollapseQueue()
        {
            QueueVisibility = Visibility.Collapsed;
        }

        internal void ShowGrid()
        {
            GridWidth = true;
        }

        internal void CollapseGrid()
        {
            GridWidth = false;
        }

        #endregion
        
        /// <summary>
        /// When selected market is changed in the droplist
        /// Set the filter to the name of the selected market (to trigger filter)
        /// Then make first task of the result the selected task
        /// </summary>
        /// <param name="newMarket"></param>
        internal void OnMarketChanged(Market newMarket)
        {
            TaskFilter = newMarket.Name;
            SelectedTask = FilteredTasks.FirstOrDefault();
        }
    }
}
