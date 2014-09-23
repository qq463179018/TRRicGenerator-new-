using System;
using System.IO;
using System.Text;

namespace Ric.Core
{
    public class Logger
    {
        private static String filename = null;
        private static String dir = null;

        public enum LogMode
        {
            New,
            Append,
            Overwrite
        }

        public enum LogType
        {
            Info,
            Warning,
            Error,
            Other
        }

        public string FilePath { get; private set; }

        private StreamWriter sw = null;

        public Logger(string filePath, LogMode mode)
        {
            bool append = (mode == LogMode.Append);
            if (mode == LogMode.New)
            {
                dir = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                filename = Path.GetFileNameWithoutExtension(filePath);
                filename += "_";
                filename += DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                filename += "_";
                filename += Guid.NewGuid();
                filename += Path.GetExtension(filePath);
                filePath = Path.Combine(dir, filename);
            }

            FilePath = filePath;
            if (sw == null)
            {
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                sw = new StreamWriter(FilePath, append) {AutoFlush = true};
            }
        }

        ~Logger()
        {
            if (sw != null)
            {
                try
                {
                    //sw.Dispose();
                }
                catch (Exception) { }
            }
        }

        public void LogErrorAndRaiseException(string msg)
        {
            Log(msg, LogType.Error);
            throw new Exception(msg);
        }

        public void LogErrorAndRaiseException(string msg, Exception innerException)
        {
            Log(msg, LogType.Error);
            throw new Exception(msg, innerException);
        }

        public void Log(string msg)
        {
            Log(msg, LogType.Info);
        }

        public void Log(string msg, LogType logType)
        {
            if (sw == null) return;
            StringBuilder sb = new StringBuilder();
            sb.Append(logType);
            sb.Append(",");
            sb.Append(DateTime.Now);
            sb.Append(",");
            sb.Append(msg);
            sw.WriteLine(sb.ToString());
        }
    }
}