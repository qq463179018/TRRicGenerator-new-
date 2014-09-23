using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using Ric.Db.Model;
using Ric.Ui.Commands;

namespace Ric.Ui.ViewModel
{
    class ErrorWindowViewModel : BaseViewModel
    {
        private string _warningText;
        public string WarningText
        {
            get { return _warningText; }
            set
            {
                if (_warningText != value)
                {
                    _warningText = value;
                    NotifyPropertyChanged("WarningText");
                }
            }
        }

        private string _messageText;
        public string MessageText
        {
            get { return _messageText; }
            set
            {
                if (_messageText != value)
                {
                    _messageText = value;
                    NotifyPropertyChanged("MessageText");
                }
            }
        }

        public Task ProblemTask { get; set; }

        #region Commands

        public ICommand SendMessageCommand { get; private set; }

        #endregion

        public ErrorWindowViewModel(Task problemTask)
        {
            ProblemTask = problemTask;
            SendMessageCommand = new SendMessageCommand(this);
            WarningText = String.Format("Task {0} seems to have some problems.\r\n\r\nThe developer in charge of this task is {1} {2}.\r\n" +
                                        "You can write a message on the box below then click on \"Send message\"", problemTask.Name, problemTask.Owner.Surname, problemTask.Owner.Familyname);
        }
    }
}
