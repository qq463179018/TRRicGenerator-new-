using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using Ric.Core;
using Ric.FileLib;
using Ric.FileLib.Entry;
using Ric.FormatLib;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class TestClassConfig3
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

    public class TestClass3 : GeneratorBase
    {
        private TestClassConfig3 configObj = null;

        protected override void Start()
        {
            Thread.Sleep(2000);
            LogMessage("task start");

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
            var test = typeof(Nda);
            LogMessage("Load nda from File");
            myNda.Load(@"C:\Users\websiting\Documents\ricpresentation\result.xls");
            myNda.Save(@"C:\Users\websiting\Documents\ricpresentation\result4.xls");
            Thread.Sleep(2000);
            //LogMessage("the first ric is : " + myNda.Content[0].Ric);
            ////foreach (NdaEntry testentry in myNda)
            ////{
            ////    var commoname = testentry.AssetCommonName;
            ////}

            //var test42 = from NdaEntry entry in myNda
            //            where entry.Ric == "ORD"
            //            select entry;
            
            //LogMessage("[DYNAMIC]the first ric is : " + myNda.DynamicContent[0].RIC);
            LogMessage("task done");
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as TestClassConfig3;
        }
    }
}
