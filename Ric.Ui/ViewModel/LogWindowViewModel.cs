using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows.Input;
using Ric.Db.Model;
using Ric.Ui.Commands;
using Ric.Ui.Commands.General;
using Ric.Ui.Model;

namespace Ric.Ui.ViewModel
{
    public class LogWindowViewModel : BaseViewModel
    {
        #region properties

        private ObservableCollection<Log> _log;

        public ObservableCollection<Log> Logs
        {
            get { return _log; }
            set
            {
                _log = value;
                NotifyPropertyChanged("Logs");
            }
        }

        #endregion

        #region Commands

        public ICommand CloseCommand { get; private set; }
        public ICommand MaximizeCommand { get; private set; }
        public ICommand MinimizeCommand { get; private set; }

        #endregion

        public LogWindowViewModel()
        {
            #region Commands

            CloseCommand = new CloseWindowCommand();
            MaximizeCommand = new MaximizeWindowCommand();
            MinimizeCommand = new MinimizeWindowCommand();

            #endregion

            Logs = new ObservableCollection<Log>();
        }
    }
}
