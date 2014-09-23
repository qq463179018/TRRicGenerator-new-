using Ric.Ui.ViewModel;

namespace Ric.Ui.View
{
    /// <summary>
    /// Interaction logic for ManagerWindow.xaml
    /// </summary>
    public partial class ManagerWindow
    {
        private ManagerWindowViewModel Managervm { get; set; }

        public ManagerWindow()
        {
            InitializeComponent();
            Managervm = new ManagerWindowViewModel();
            DataContext = Managervm;
        }
    }
}
