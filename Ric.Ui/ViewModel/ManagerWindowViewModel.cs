using System.Collections.ObjectModel;
using Ric.Core;
using Ric.Db.Manager;
using Ric.Db.Model;

namespace Ric.Ui.ViewModel
{
    class ManagerWindowViewModel : BaseViewModel
    {
        #region properties

        #region User

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

        #region Task

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
                if (_taskIndex != value)
                {
                    _taskIndex = value;
                    NotifyPropertyChanged("TaskIndex");
                }
            }
        }

        #endregion

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

        private bool _isRoEmail;
        public bool IsRoEmail
        {
            get { return _isRoEmail; }
            set
            {
                if (_isRoEmail != value)
                {
                    _isRoEmail = value;
                    NotifyPropertyChanged("IsRoEmail");
                }
            }
        }


        #endregion

        #region Commands

        //public ICommand SaveCommand { get; private set; }

        #endregion

        public ManagerWindowViewModel()
        {
            UserList = new ObservableCollection<User>(UserManager.GetUserNamesByManager(RunTimeContext.Context.CurrentUser, RunTimeContext.Context.DatabaseContext));
            TaskList = new ObservableCollection<Task>(TaskManager.GetTasksByManager(RunTimeContext.Context.CurrentUser, RunTimeContext.Context.DatabaseContext));
            //SaveCommand = new SaveUserCommand(this);
            UserIndex = 0;
            TaskIndex = 0;
        }
    }
}
