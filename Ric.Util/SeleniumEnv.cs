using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Selenium;

namespace Ric.Util
{
    public class SeleniumEnv : IDisposable
    {
        private string baseUri = string.Empty;
        public string BaseUri
        {
            get
            {
                return baseUri;
            }
        }

        private ISelenium seleniumInstance = null;

        public ISelenium SeleniumInstance
        {
            get
            {
                return this.seleniumInstance;
            }
        }

        public SeleniumEnv(string baseUri)
        {
            this.baseUri = baseUri;

            startServer();
        }

        //Start RC server
        private void startServer()
        {
            string filename = "java";
            string parameters = @" -jar selenium-server.jar";
            System.Diagnostics.Process[] allProcess = System.Diagnostics.Process.GetProcessesByName("java");
            try
            {
                if (allProcess.Length == 0)
                {
                    System.Diagnostics.Process.Start(filename, parameters);
                    System.Threading.Thread.Sleep(3000);

                }
                seleniumInstance = new DefaultSelenium("localhost", 4444, "*iexplore", this.baseUri);
                seleniumInstance.Start();
                seleniumInstance.UseXpathLibrary("javascript-xpath");
            }
            catch (System.Exception e)
            {
                Exception friendlyEx = new Exception("Selenium server not started.", e);
                throw friendlyEx;
            }

        }

        #region IDisposable Members

        public void Dispose()
        {
            shutDownServer();
        }


        //Shut down RC server
        private void shutDownServer()
        {
            try
            {
                if (seleniumInstance != null)
                {
                    seleniumInstance.Stop();
                    seleniumInstance.ShutDownSeleniumServer();
                }

                System.Diagnostics.Process[] allProcess = System.Diagnostics.Process.GetProcessesByName("java");
                if (allProcess.Length != 0)
                {
                    for (int i = 0; i < allProcess.Length; i++)
                    {
                        allProcess[i].Kill();
                    }
                }
            }
            catch (SeleniumException e)
            {
                Exception friendlyEx = new Exception("Selenium sever shut down failed.", e);
                throw friendlyEx;
            }
        }

        #endregion
    }
}
