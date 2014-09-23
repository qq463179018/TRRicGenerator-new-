using System;
using System.Windows;
using System.Windows.Input;
using Ric.Core;
using Ric.Ui.ViewModel;
using Ric.Util;

namespace Ric.Ui.Commands
{
    internal class SendMessageCommand : ICommand
    {
        private ErrorWindowViewModel _viewModel;

        public SendMessageCommand(ErrorWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        #region Icommand implementation

        bool ICommand.CanExecute(object parameter)
        {
            return true;
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        void ICommand.Execute(object parameter)
        {
            var window = parameter as Window;

            var message = _viewModel.MessageText;
            var devEmail = _viewModel.ProblemTask.Owner.Email;
            var outlookApp = new OutlookApp();
            string err;
            var mail = new MailToSend
            {
                MailBody = message,
                MailSubject = _viewModel.ProblemTask.Name + " Problem."
            };
            mail.ToReceiverList.Add(devEmail);
            if (RunTimeContext.Context.CurrentUser.Manager != null)
            {
                mail.CCReceiverList.Add(RunTimeContext.Context.CurrentUser.Manager.Email);
            }
            OutlookUtil.CreateAndSendMail(outlookApp,  mail, out err);
            MessageBox.Show("The email was sent successfully!");
            if (window != null)
            {
                window.Close();
            }
        }

        #endregion
    }
}
