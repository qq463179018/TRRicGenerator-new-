using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Security;
using System.Text;
using System.Windows.Input;
using Ric.Db.Model;
using Ric.Ui.Commands;
using Ric.Ui.Commands.General;
using Ric.Ui.Model;
using Ric.Ui.View;

namespace Ric.Ui.ViewModel
{
    public class DatabaseViewModel : BaseViewModel
    {
        #region properties

        public List<string> AvailableTables { get; set; }

        public string SelectedTable { get; set; }

        #endregion

        #region Commands

        public ICommand CloseCommand { get; private set; }
        public ICommand MaximizeCommand { get; private set; }
        public ICommand MinimizeCommand { get; private set; }

        #endregion

        public DatabaseViewModel()
        {
            #region Commands

            CloseCommand = new CloseWindowCommand();
            MaximizeCommand = new MaximizeWindowCommand();
            MinimizeCommand = new MinimizeWindowCommand();

            #endregion

            AvailableTables = new List<string>
            {
                "Stamp Duty",
                "Trading News Expire Date",
                "Trading News Exl Name"
            };

            SelectedTable = AvailableTables.FirstOrDefault();
        }
    }
}
