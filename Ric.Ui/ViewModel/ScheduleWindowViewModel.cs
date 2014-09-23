using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using Ric.Db.Model;
using Ric.Ui.Commands.Schedules;

namespace Ric.Ui.ViewModel
{
    public class ScheduleWindowViewModel : BaseViewModel
    {
        #region properties

        private DateTime dateSchedule;
        public DateTime DateSchedule
        {
            get { return dateSchedule; }
            set
            {
                dateSchedule = value;
                NotifyPropertyChanged("DateSchedule");
            }
        }

        private Task taskSchedule;
        public Task TaskSchedule
        {
            get { return taskSchedule; }
            set
            {
                taskSchedule = value;
                NotifyPropertyChanged("TaskSchedule");
            }
        }

        private Schedule selectedSchedule;
        public Schedule SelectedSchedule
        {
            get { return selectedSchedule; }
            set
            {
                selectedSchedule = value;
                if (value != null)
                    DateSchedule = selectedSchedule.Date;
                NotifyPropertyChanged("SelectedSchedule");
            }
        }

        private ScheduleFrequency _frequencySchedule;

        public ScheduleFrequency FrequencySchedule
        {
            get { return _frequencySchedule; }
            set
            {
                _frequencySchedule = value;
                NotifyPropertyChanged("FrequencySchedule");
            }
        }

        public IEnumerable<ScheduleFrequency> ScheduleTypesValues
        {
            get
            {
                return Enum.GetValues(typeof(ScheduleFrequency))
                    .Cast<ScheduleFrequency>();
            }
        }

        #endregion

        #region Commands

        public ICommand SaveScheduleCommand { get; private set; }

        #endregion

        public ScheduleWindowViewModel(Schedule toChangeSchedule)
        {
            SelectedSchedule = toChangeSchedule;
            FrequencySchedule = SelectedSchedule.Frequency;
            DateSchedule = DateTime.Now.AddHours(2);
            SaveScheduleCommand = new SaveScheduleCommand(this);
        }
    }
}
