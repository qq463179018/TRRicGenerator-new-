using System;
using System.IO;
using System.Diagnostics;

namespace Ric.Util
{
    public class GatsConfig
    {
        public string IdnServerIp { get; set; }
        public string ElektronServerIp { get; set; }
    }

    public class GatsUtil
    {
        public enum Server
        {
            Idn,
            Elektron
        }
        private const string GatsToolName = "Tools\\Data2XML\\Data2XML.exe";
        private const string GatsToolPath = "Tools\\Data2XML";
        private const string GatsConfig = "Tools\\Data2XML\\GatsConfig.xml";
        private const int Timeout = 20;
        public string ServerIp { get; set; }
        public Server ServerFeed { get; set; }

        public GatsUtil()
        {
            InitializeServerIp(Server.Idn);
            InitializeFileRequire();
            ConnectGats();
        }

        public GatsUtil(Server server)
        {
            InitializeServerIp(server);
            InitializeFileRequire();
            ConnectGats();
        }

        /// <summary>
        /// Get GATS server IP from config file.
        /// </summary>
        private void InitializeServerIp(Server server)
        {
            GatsConfig conf = ConfigUtil.ReadConfig(GatsConfig, typeof(GatsConfig)) as GatsConfig;
            if (server == Server.Idn)
            {
                ServerIp = conf.IdnServerIp;
                ServerFeed = Server.Idn;
            }
            else
            {
                ServerIp = conf.ElektronServerIp;
                ServerFeed = Server.Elektron;
            }
            if (string.IsNullOrEmpty(ServerIp))
            {
                string msg = "Can not get GATS server IP from config file.";
                throw new Exception(msg);
            }
        }

        /// <summary>
        /// Check if the tool is existed.
        /// </summary>
        private void InitializeFileRequire()
        {
            if (!File.Exists(GatsToolName))
            {
                string msg = string.Format("Can not found GATS tool. Please check below path. {0}", GatsToolName);
                throw new FileNotFoundException(msg);
            }            
        }

        /// <summary>
        /// Test if GATS can be accessed.
        /// </summary>     
        public void ConnectGats()
        {
            string command = string.Format("data2xml.exe -ph {0} -pn ELEKTRON_AD -rics \".KS200\" -fids \"DSPLY_NAME\" -quiet -dbout", ServerIp);
            Process gatsProcess = new Process
            {
                StartInfo =
                {
                    FileName = GatsToolName,
                    WorkingDirectory = GatsToolPath,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    Arguments = command
                }
            };

            int retry = 3;
            bool success = false;
            while (!success && retry-- > 0)
            {                
                gatsProcess.Start();
                success = gatsProcess.WaitForExit(Timeout * 1000);
                if (!success)
                {
                    gatsProcess.Kill();
                }
            }
            if(!success)
            {
                throw new Exception("Can not connect to GATS. GATS returns no reponse.");
            }           
        }

        /// <summary>
        /// Give GATS a command line. Get the response.
        /// </summary>   
        /// <param name="command">command</param>
        /// <returns>response</returns>
        public string GetGatsResponse(string command)
        {
            Process gatsProcess = new Process
            {
                StartInfo =
                {
                    FileName = GatsToolName,
                    WorkingDirectory = GatsToolPath,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                }
            };

            try
            {  
                int retry = 3;
                bool success = false;
                string response = null;
                while (!success && retry-- > 0)
                {
                    gatsProcess.StartInfo.Arguments = command;
                    gatsProcess.Start();
                    response = gatsProcess.StandardOutput.ReadToEnd();
                    success = gatsProcess.WaitForExit(Timeout * 1000);
                    if (!success)
                    {
                        gatsProcess.Kill();
                    }
                }              
               
                return response;
            }
            catch
            {
                return null;
            }
        }

        private string GetServerFeed()
        {
            return ServerFeed == Server.Idn ? "IDN_RDF" : "ELEKTRON_AD";
        }

        /// <summary>
        /// Give the specific rics and fids to query. 
        /// fids can be null.
        /// </summary>
        /// <param name="rics">rics, with comma sperated</param>
        /// <param name="fids">fids, with comma sperated</param>
        /// <returns>GATS reponse</returns>
        public string GetGatsResponse(string rics, string fids)
        {
            if (string.IsNullOrEmpty(rics))
            {
                return null;
            }
            string command = string.Format("data2xml.exe -quiet -dbout -raw_enum_vals -ph {0} -pn \"{1}\" -rics \"{2}\"", ServerIp, GetServerFeed(), rics);
            if (!string.IsNullOrEmpty(fids))
            {
                command += " -fids \"{0}\"";
                command = string.Format(command, fids);
            }

            return GetGatsResponse(command);
        }


        
    }
}

