using System.Windows;
using System.Windows.Input;
using Ric.Db.Model;
using Ric.Ui.ViewModel;

namespace Ric.Ui.View
{
	/// <summary>
	/// Interaction logic for ScheduleWindow.xaml
	/// </summary>
	public partial class ScheduleWindow
	{
		public ScheduleWindowViewModel Schedulevm { get; set; }

		public ScheduleWindow(Schedule toChangeSchedule)
		{
			InitializeComponent();
            Schedulevm = new ScheduleWindowViewModel(toChangeSchedule);
			DataContext = Schedulevm;
            if (!string.IsNullOrEmpty(toChangeSchedule.DayOfWeek))
            {
                string[] day = toChangeSchedule.DayOfWeek.Split(',');
                foreach (string item in day)
                {
                    switch (item)
                    {
                        case "Mon":
                            MoncheckBox.IsChecked = true;
                            break;   
                        case "Tue":
                            TuecheckBox.IsChecked = true;
                            break;
                        case "Wen":
                            WencheckBox.IsChecked = true;
                            break;
                        case "Thu":
                            ThucheckBox.IsChecked = true;
                            break;
                        case "Fri":
                            FricheckBox.IsChecked = true;
                            break;
                        case "Sat":
                            SatcheckBox.IsChecked = true;
                            break;
                        case "Sun":
                            SuncheckBox.IsChecked = true;
                            break;
                        default:
                            break;
                    }
                }
            }
            

		}

        private void ScheduleWindow_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void Button_Close(object sender, RoutedEventArgs e)
        {
            Close();
        }
	}
}