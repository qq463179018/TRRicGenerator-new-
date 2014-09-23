using System;
using System.Diagnostics;

namespace Ric.Core
{
    public class ProcessContext : IDisposable
    {
        public Process ProcessInstance = null;

        public ProcessContext(string file, string arg, string path)
        {
            ProcessStartInfo psi = new ProcessStartInfo(file, arg)
            {
                WorkingDirectory = path,
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false
            };
            ProcessInstance = new Process {StartInfo = psi};
        }

        public ProcessContext(string file, string arg)
        {
            ProcessStartInfo psi = new ProcessStartInfo(file, arg);
            ProcessInstance = new Process {StartInfo = psi};
        }

        public ProcessContext(ProcessStartInfo psi)
        {
            ProcessInstance = new Process {StartInfo = psi};
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (ProcessInstance != null)
            {
                try
                {
                    ProcessInstance.Kill();
                }
                catch (Exception)
                {
                }
                ProcessInstance = null;
            }
        }

        #endregion IDisposable Members
    }
}