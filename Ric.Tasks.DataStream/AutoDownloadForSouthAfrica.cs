using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.Net;
using System.IO;
using HtmlAgilityPack;
using Ric.Util;

namespace Ric.Tasks.DataStream
{
    class AutoDownloadForSouthAfrica : GeneratorBase
    {
        #region Field
        private string url = @"https://www.jse.co.za/downloadable-files?RequestNode=/Safex/EdmStats";
        private AutoDownloadForSouthAficaConfig ConfigObj = null;
        private string Date = null;
        private string FileFolder = null;
        private string FileName = null;
        private CookieContainer cookies = new CookieContainer();
        #endregion

        protected override void Initialize()
        {
            ConfigObj = Config as AutoDownloadForSouthAficaConfig;
            Date = ConfigObj.Date;
            FileFolder = ConfigObj.OutputPath;
            FileName = "Fullstats" + Date + ".xls";
            Logger.Log("Initialize...OK!");
        }

        protected override void Start()
        {
            StartJob();
        }

        private void StartJob()
        {
            string fileUrl = GetFileUrl();
            DownLoadFile(@"https://www.jse.co.za/" + fileUrl);
        }


        private string LandWebSite()
        {
            try
            {
                string st = WebClientUtil.GetPageSource(url, 300000);
                return st;
            }
            catch (System.Exception ex)
            {

                Logger.Log(string.Format("Failed to land the website.{0}", ex.Message));
                return null;
            }
            
        }

        private string GetFileUrl()
        {
            try
            {
                string st = LandWebSite();

                if (string.IsNullOrEmpty(st))
                {
                    return null;
                }

                string fileUrl = null;
                HtmlDocument html = new HtmlDocument();
                html.LoadHtml(st);

                HtmlNodeCollection nodeCollections = html.DocumentNode.SelectNodes("//div[@class = 'documents__detailed left']");
                foreach (HtmlNode node in nodeCollections)
                {
                    HtmlNode a = node.SelectSingleNode(".//a");
                    string fileName = a.InnerText;
                    if (fileName.Equals(FileName))
                    {
                        fileUrl = a.Attributes["href"].Value;
                        break;
                    }
                }
                return fileUrl;
            }
            catch (System.Exception ex)
            {
                Logger.Log(string.Format("Failed to get file url.{0}", ex.Message));
                return null;
            }
        }

        private void DownLoadFiles(string url, string fileName)
        {
            try
            {
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 300000;
                request.Method = "GET";
                request.CookieContainer = cookies;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream content = response.GetResponseStream();
                using (Stream file = File.Create(fileName))
                {

                    byte[] buffer = new byte[8 * 1024];
                    int len;
                    int offset = 0;

                    while ((len = content.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        file.Write(buffer, 0, len);
                        offset += len;
                    }
                }
            }
            catch (System.Exception ex)
            {
                Logger.Log(string.Format("Failed to download file.{0}", ex.Message));
            }      
        }

        private void DownLoadFile(string fileUrl)
        {
            string fileName = FileFolder + "\\" + FileName;
            DownLoadFiles(fileUrl, fileName);
        }
    }
}
