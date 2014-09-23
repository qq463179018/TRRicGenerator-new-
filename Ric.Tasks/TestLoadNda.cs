using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using Ric.Core;
using Ric.FileLib;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class TestClassConfig2
    {
        [Category("Information")]
        [DisplayName("File name")]
        [StoreInDB]
        public string ExcelName { get; set; }

        [StoreInDB]
        public string ResultFolder { get; set; }

        [StoreInDB]
        public List<string> TestList { get; set; }
    }

    public class TestClass2 : GeneratorBase
    {
        private TestClassConfig2 configObj = null;

        protected override void Start()
        {
            Thread.Sleep(2000);

            var props = new List<Dictionary<string, string>>
            {
                new Dictionary<string, string>
                {
                    {"ric", "12345"},
                    {"commonname", "first company"}
                },
                new Dictionary<string, string>
                {
                    {"ric", "33322"},
                    {"commonname", "other company"}
                }
            };

            var myNda = new Nda();
            LogMessage("create nda from template");
            myNda.Load(@"C:\Users\websiting\Documents\ricpresentation\nda.csv");

            Thread.Sleep(2000);
            //myNda.Content[0].
            //LogMessage("save nda in : " + configObj.ExcelName);
            //LogMessage("task done");
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as TestClassConfig2;
        }
    }
}
