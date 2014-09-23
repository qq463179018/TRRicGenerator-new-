using System.Windows;
using Ric.Ui.ViewModel;

namespace Ric.Ui.View
{
    /// <summary>
    /// Interaction logic for TipsWindow.xaml
    /// </summary>
    public partial class TipsWindow
    {
        public TipsWindow()
        {
            InitializeComponent();
            var vm = new TipWindowViewModel();
            DataContext = vm;
        }

        private void CloseButton_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
