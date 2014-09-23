using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.Core;
using System.ComponentModel;

namespace Ric.Tasks.DataStream
{
    [ConfigStoredInDB]
    public class AutoDownloadForSouthAficaConfig
    {

        [Category("Date")]
        [DisplayName("Date")]
        [Description("Date format: yyyyMMdd. E.g. 20141206")]
        public string Date { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("C:\\AutoDownloadForSouthAfrica\\")]
        [DisplayName("Output path")]
        public string OutputPath { get; set; }

        public AutoDownloadForSouthAficaConfig()
        {
            Date = DateTime.Now.ToUniversalTime().ToString("yyyyMMdd");
        }
    }
}
