using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using Selenium;
using System.Windows.Forms;
//using Reuters.ProcessQuality.ContentAuto.Lib;

namespace Ric.Tasks.HongKong
{
    
    public class ConnectionConfig
    {
        public static void SetUpAutoConfigScript(bool enable)
        {
            const string DefaultConnectionSettings = "DefaultConnectionSettings";
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\Connections", true);
            if (key == null)
                return;

            byte[] connectionSettingsValues = (byte[])key.GetValue(DefaultConnectionSettings);
            if (enable)
                connectionSettingsValues[8] = Convert.ToByte(5);
            else
                connectionSettingsValues[8] = Convert.ToByte(1);

            key.SetValue(DefaultConnectionSettings, connectionSettingsValues, RegistryValueKind.Binary);
            key.Close();
        }
    }

    public class UploadFileIntoGeda
    {
        private ISelenium selenium;

        public void gedaOperation()
        {

            uploadFile(0);
            ShutDownRC();
        }

        private void uploadFile(int i)
        {
            StartRC();
            i++;
            try
            {
                selenium.Open("/GEDA15/index.html");
                selenium.WaitForPageToLoad("50000");
                selenium.Type("username", "YQ_LI");
                selenium.Type("password", "Reuters5");
                //selenium.Click("ext-gen27");
                selenium.Click("//button[@class=' x-btn-text'][text()='Login']");
                selenium.WaitForPageToLoad("50000");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("css=img.x-tree-ec-icon.x-tree-elbow-end-plus");
                System.Threading.Thread.Sleep(2000);
                selenium.DoubleClick("//ul[@class='x-tree-node-ct']/li[@class='x-tree-node']/descendant::a[@class='x-tree-node-anchor']/span[text()='Message Editor']");
                selenium.SelectFrame("IFFMEditor");
                System.Threading.Thread.Sleep(4000);

                //Upload the first file
                selenium.Click("ext-gen163");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("//div[@class='x-combo-list-inner']/descendant::div[contains(@class, 'x-combo-list-item')][text()='HK_BULK']");
                selenium.Click("ext-gen168");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("css=#ext-gen182 > div.x-combo-list-item");
                selenium.Click("//button[@class=' x-btn-text search'][text()='Search']");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("//button[@class=' x-btn-text import'][text()='Upload File']");
                System.Threading.Thread.Sleep(2000);
                selenium.Type("form-file", "D:\\HKRicTemplate\\HKG_EQLB.txt");
                System.Threading.Thread.Sleep(3000);
                //Upload
                selenium.Click("//button[@class=' x-btn-text'][text()='Upload']");
                System.Threading.Thread.Sleep(2000);
                //Ok
                selenium.Click("//button[@class=' x-btn-text'][text()='OK']");
                System.Threading.Thread.Sleep(3000);

                //Upload the second file
                selenium.Click("ext-gen163");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("//div[@class='x-combo-list-inner']/descendant::div[contains(@class, 'x-combo-list-item')][text()='HK_CBBC1']");
                selenium.Click("ext-gen168");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("css=#ext-gen182 > div.x-combo-list-item");
                selenium.Click("//button[@class=' x-btn-text search'][text()='Search']");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("//button[@class=' x-btn-text import'][text()='Upload File']");
                System.Threading.Thread.Sleep(2000);
                selenium.Type("form-file", "D:\\HKRicTemplate\\HKG_EQLB_CBBC.txt");
                System.Threading.Thread.Sleep(2000);
                //Upload
                selenium.Click("//button[@class=' x-btn-text'][text()='Upload']");
                System.Threading.Thread.Sleep(2000);
                //Ok
                selenium.Click("//button[@class=' x-btn-text'][text()='OK']");
                System.Threading.Thread.Sleep(3000);

                //Upload the third file
                selenium.Click("ext-gen163");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("//div[@class='x-combo-list-inner']/descendant::div[contains(@class, 'x-combo-list-item')][text()='HK_EQLBMI']");
                selenium.Click("ext-gen168");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("css=#ext-gen182 > div.x-combo-list-item");
                selenium.Click("//button[@class=' x-btn-text search'][text()='Search']");
                System.Threading.Thread.Sleep(2000);
                selenium.Click("//button[@class=' x-btn-text import'][text()='Upload File']");
                System.Threading.Thread.Sleep(2000);
                selenium.Type("form-file", "D:\\HKRicTemplate\\HKG_EQLBMI.txt");
                System.Threading.Thread.Sleep(2000);
                //Upload
                selenium.Click("//button[@class=' x-btn-text'][text()='Upload']");
                System.Threading.Thread.Sleep(2000);
                //Ok
                selenium.Click("//button[@class=' x-btn-text'][text()='OK']");

                selenium.Close();
                selenium.Stop();

            }
            catch (SeleniumException ex)
            {
                if (ex.Message.Contains("Timed out") && i < 4)
                {

                    uploadFile(i);

                }

                MessageBox.Show(ex.Message);
                selenium.Close();
                selenium.Stop();
            }
        }

        private void StartRC()
        {

            string filename = "java";
            string parameters = @" -jar selenium-server.jar";
            System.Diagnostics.Process[] allProcess = System.Diagnostics.Process.GetProcessesByName("java");
            try
            {

                System.Diagnostics.Process.Start(filename, parameters);
                System.Threading.Thread.Sleep(5000);

                selenium = new DefaultSelenium("localhost", 4444, "*iexplore", "http://196.11.63.237:8080/");
                selenium.Start();
                //selenium.UseXpathLibrary("javascript-xpath");
            }
            catch (System.Exception e)
            {
                MessageBox.Show("Selenium console not started." + e.ToString());
                System.Windows.Forms.Application.Exit();
            }

        }

        private void ShutDownRC()
        {
            try
            {

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
                MessageBox.Show("Error found in ShutDownRC:" + e.ToString());

            }
        }
    }
}
