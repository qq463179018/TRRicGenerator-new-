using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using Ric.Core;
using Ric.FileLib;
using Ric.FileLib.Entry;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;
using System;
using System.Drawing;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class TestClassConfig
    {
        [StoreInDB]
        public string ExcelName { get; set; }

        [StoreInDB]
        public string ResultFolder { get; set; }

        [StoreInDB]
        public int TestInt { get; set; }

        [StoreInDB]
        public List<string> TestList { get; set; }

        [StoreInDB]
        public DateTime test2 { get; set; }

        [StoreInDB]
        public bool test3 { get; set; }

        //[StoreInDB]
        //public Color test4 { get; set; }

        //[StoreInDB]
        //public TimeSpan test5 { get; set; }

        //[StoreInDB]
        //[ExpandableObject()]
        //[Category("Information")]
        //[DisplayName("Test other class 2")]
        //public AnotherTest otherClass { get; set; }

        //public TestClassConfig()
        //{
        //    otherClass = new AnotherTest();
        //}
    }

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class AnotherTest
    {
        public string TestExpand { get; set; }
    }

    public class TestClass : GeneratorBase
    {
        private TestClassConfig configObj = null;

        protected override void Start()
        {
            Thread.Sleep(2000);
            LogMessage("I do that");
            LogMessage("and that");
            Thread.Sleep(2000);
            AddResult("result file", @"C:\workspace\branches\newUi\Documentation\TaskGuidelines.docx", "nda");

            //var myNda = new Nda();
            //myNda.Load(@"C:\Users\Aurelien\Documents\test.xlsx");

            //foreach (NdaEntry entry in myNda)
            //{
            //    LogMessage("The Ric is " + entry.Ric);
            //}



            Thread.Sleep(2000);
            LogMessage("last thing");
            //LogMessage("Excel name is " + configObj.ExcelName);
            LogMessage("writing text from task");
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as TestClassConfig;
        }
    }
}
