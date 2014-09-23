using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using Ric.Core;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class TestClassConfig4
    {
        [StoreInDB]
        public string ExcelName { get; set; }

        [StoreInDB]
        public string ResultFolder { get; set; }

        [StoreInDB]
        public List<string> TestList { get; set; }

    }

    public class TestClass4 : GeneratorBase
    {
        private TestClassConfig4 configObj = null;

        protected override void Start()
        {
            throw new Exception("suce");
            //AddResult(@"C:\Users\websiting\Documents\ricpresentation\nda.csv");
        }

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as TestClassConfig4;
        }
    }
}
