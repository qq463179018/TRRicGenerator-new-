using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Drawing.Design;
using Ric.Core;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class DataStreamRicCreationWithFileDownloadConfig
    {
        [StoreInDB]
        [Category("Username")]
        [DefaultValue("icw.marketdata@reuters.com")]
        [Description("Username for login to the website.")]
        public string Username { get; set; }

        [StoreInDB]
        [Category("Password")]
        [DefaultValue("Germany@8624")]
        [Description("Password for login to the website.")]
        public string Password { get; set; }

        [StoreInDB]
        [Category("Path")]
        [DefaultValue("D:\\DataStream\\RIC_Creation\\")]
        [Description("Path to save generated output file. E.g.D:\\DataStream\\RIC_Creation\\")]
        public string OutputPath { get; set; }

        [StoreInDB]
        [Category("EmailAccount")]
        [Description("Account name which used to search the target mail, like: \"UC169XXX\"")]
        public string AccountName { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Recepients list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailTo { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Mail CC list\nMail address should contain full information\nE.g. xxx.xxx@thomsonreuters.com")]
        public List<string> MailCC { get; set; }

        [StoreInDB]
        [Category("Mail")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Description("Signature for E-mail")]
        public List<string> MailSignature { get; set; }
    }
}
