using System.Windows.Input;
using Ric.Ui.ViewModel;

namespace Ric.Ui.View
{
    /// <summary>
    /// Interaction logic for AdminWindow.xaml
    /// </summary>
    public partial class AdminWindow
    {
        public AdminWindowViewModel Adminvm { get; set; }

        public AdminWindow()
        {
            InitializeComponent();
            Adminvm = new AdminWindowViewModel();
            DataContext = Adminvm;
        }

        private void AdminWindow_OnMouseLeftButtonDownWindow_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
    }
}
