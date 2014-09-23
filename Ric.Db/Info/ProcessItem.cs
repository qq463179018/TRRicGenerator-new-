using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Info
{
    public class ProcessItem
    {
        private string content = string.Empty;
        private string url = string.Empty;
        private DateTime captureDate;
        private List<ProcessException> Exceptions = new List<ProcessException>();

        public string Content
        {
            get { return content; }
            set { content = value; }
        }
        public string Url
        {
            get { return url; }
            set { url = value; }
        }
        public DateTime CaptureDate
        {
            get { return captureDate; }
            set { captureDate = value; }
        }
        public bool HasExceptions
        {
            get
            {
                return Exceptions.Count > 0;
            }
        }

        public void AddException(ProcessException processExcetpion)
        {
            Exceptions.Add(processExcetpion);
        }
    }

    public class ProcessException : Exception
    {
        public ProcessException(string message)
            : base(message) { }

        public ProcessException(string message, Exception innerException)
            : base(message, innerException) { }
    }
}
