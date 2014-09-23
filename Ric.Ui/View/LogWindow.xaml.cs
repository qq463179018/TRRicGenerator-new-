using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Ric.Ui.ViewModel;

namespace Ric.Ui.View
{
    /// <summary>
    /// Interaction logic for LogWindows.xaml
    /// </summary>
    public partial class LogWindow
    {
        public LogWindowViewModel Myvm { get; set; }

        public LogWindow()
        {
            InitializeComponent();
            Myvm = new LogWindowViewModel();
            DataContext = Myvm;
        }

        private void LogWindow_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
    }
}
