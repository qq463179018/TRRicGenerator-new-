using System;

namespace Ric.Core.Events
{
    public class LogEventArgs : EventArgs
    {
        public string Message { get; set; }

        public Logger.LogType LogType { get; set; }

        public LogEventArgs(string logMessage)
        {
            Message = logMessage;
            LogType = Logger.LogType.Info;
        }

        public LogEventArgs(string logMessage, Logger.LogType logType)
        {
            Message = logMessage;
            LogType = logType;
        }
    }
}