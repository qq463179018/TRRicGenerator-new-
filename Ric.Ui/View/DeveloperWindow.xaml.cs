using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.View
{
    /// <summary>
    /// Interaction logic for DeveloperWindow.xaml
    /// </summary>
    public partial class DeveloperWindow
    {
        public DeveloperWindowViewModel Myvm { get; set; }



        public DeveloperWindow()
        {
            InitializeComponent();
            Myvm = new DeveloperWindowViewModel();
            DataContext = Myvm;
        }

        private void DeveloperWindow_OnMouseLeftButton(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
    }
}
