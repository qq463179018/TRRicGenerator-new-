using System.Windows.Input;
using Ric.Ui.Commands;

namespace Ric.Ui.ViewModel
{
    class TipWindowViewModel : BaseViewModel
    {

        private string _messageText;
        public string MessageText
        {
            get { return _messageText; }
            set
            {
                if (_messageText != value)
                {
                    _messageText = value;
                    NotifyPropertyChanged("MessageText");
                }
            }
        }

        //public Tip CurrentTip { get; set; }

        #region Commands

        public ICommand NextTipCommand { get; private set; }

        #endregion

        public TipWindowViewModel()
        {
            NextTipCommand = new NextTipCommand(this);
        }
    }
}
