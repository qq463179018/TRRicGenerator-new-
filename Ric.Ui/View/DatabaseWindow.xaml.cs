using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.Entity;
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
using Ric.Core;
using Ric.Db.Model;
using Ric.Ui.ViewModel;

namespace Ric.Ui.View
{
    /// <summary>
    /// Interaction logic for LogWindows.xaml
    /// </summary>
    public partial class DatabaseWindow
    {
        public DatabaseViewModel Databasevm { get; set; }

        public HongKongModel Model { get; set; }

        public DatabaseWindow()
        {
            InitializeComponent();
            Databasevm = new DatabaseViewModel();
            DataContext = Databasevm;
            Model = new HongKongModel();
        }

        private void LogWindow_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void Selector_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Databasevm.SelectedTable == "Stamp Duty")
            {
                Model.ETI_HK_StampDuties.Load();

                DatabaseGrid.ItemsSource = Model.ETI_HK_StampDuties.Local.ToList();
            }
            else if (Databasevm.SelectedTable == "Trading News Expire Date")
            {
                Model.ETI_HK_TradingNews_ExpireDates.Load();

                DatabaseGrid.ItemsSource = Model.ETI_HK_TradingNews_ExpireDates.Local.ToList();
                //DatabaseGrid.ItemsSource =
                //    (from entry in Model.ETI_HK_TradingNews_ExpireDates
                //        select entry).ToList();
            }
            else if (Databasevm.SelectedTable == "Trading News Exl Name")
            {
                DatabaseGrid.ItemsSource =
                    (from entry in Model.ETI_HK_TradingNews_ExlNames
                        select entry).ToList();
            }

        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            int numberChanged = Model.SaveChanges();
            MessageBox.Show(String.Format("Changes in the database saved successfully.\r\n\r\n{0} rows affected.",
                numberChanged));
            DatabaseGrid.Items.Refresh();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
