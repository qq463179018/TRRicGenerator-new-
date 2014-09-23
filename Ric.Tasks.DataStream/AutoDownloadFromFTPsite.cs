using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;
using Ric.Util;
using System.Net;
using System.IO;
using HtmlAgilityPack;

namespace Ric.Tasks.DataStream
{
    [ConfigStoredInDB]
    public class AutoDownloadFromFTPsiteConfig
    {

        [Category("Date")]
        [DisplayName("Date")]
        [Description("Date format: yyyyMMdd. E.g. 20141206")]
        public string Date { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\AutoDownloadFromFTPsite\\")]
        [DisplayName("Output path")]
        public string OutputPath { get; set; }

        public AutoDownloadFromFTPsiteConfig()
        {
            Date = DateTime.Now.ToUniversalTime().ToString("ddMMyy");
        }
    }

    class AutoDownloadFromFTPsite : GeneratorBase   
    {
        private AutoDownloadFromFTPsiteConfig ConfigObj = null;
        private string GetFileInfo()
        {
            try
            {
                string url = @"ftp://datastream:suY0l2IF@ock-ftp2.ilx.net/pub/";

                FtpWebRequest reqFTP = (FtpWebRequest)FtpWebRequest.Create(url);
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                StreamReader sr = new StreamReader(response.GetResponseStream());
                string st = sr.ReadToEnd();

                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(st);
                string href = string.Format("DS_fut_file.{0}", ConfigObj.Date);
                href = string.Format(".//a[@href='{0}']", href);

                HtmlNodeCollection node = doc.DocumentNode.SelectNodes(href);
                if (node.Count > 0)
                {
                    string filename = node[0].InnerText;
                    url = url + filename;
                    reqFTP = (FtpWebRequest)FtpWebRequest.Create(url);
                    response = (FtpWebResponse)reqFTP.GetResponse();
                    sr = new StreamReader(response.GetResponseStream());
                    st = sr.ReadToEnd();
                    return st;
                }
                else
                {
                    LogMessage("No file downloaded!");
                    return null;
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("Get file info error.ex:{0}", ex.Message);
                Logger.Log(msg);
                LogMessage("No file downloaded!");
                return null;
            }
            
        }

        protected override void Initialize()
        {
            ConfigObj = Config as AutoDownloadFromFTPsiteConfig;
        }

        protected override void Start()
        {
            string fileContext = GetFileInfo();
            if (!string.IsNullOrEmpty(fileContext))
            {
                GenerateTargetFile(fileContext);
            }
            
        }

        private void GenerateTargetFile(string fileContext)
        {
            try
            {
                string targetFile = ConfigObj.OutputPath +"\\" + "DS_fut_file." + ConfigObj.Date;
                if (File.Exists(targetFile))
                {
                    File.Delete(targetFile);
                }
                //File.Create(targetFile);
                FileStream fs = new FileStream(targetFile,FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(fileContext);
                sw.Close();
                fs.Close();
                
            }
            catch (System.Exception ex)
            {
                string msg = string.Format("Generate target file error.ex:{0}", ex.Message);
                Logger.Log(msg);
            }
        }
    }
}
