using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Globalization;
using Ric.Util;
using pdftron;
using PdfTronWrapper;
using Ric.Db.Manager;
using Ric.Core;

namespace Ric.Tasks
{
    [ConfigStoredInDB]
    public class GedaTestConfig
    {
        [StoreInDB]
        public string path { set; get; }
        public string type { set; get; }

        public GedaTestConfig()
        {
            type = "GEDA_TEST";
        }
    }

    public class GedaTest : GeneratorBase
    {
        GedaTestConfig testConfig;

        protected override void Initialize()
        {
            base.Initialize();

            testConfig = Config as GedaTestConfig;
        }

        protected override void Start()
        {
            string path = testConfig.path;

            Logger.Log("test log");

            FileProcessType type = (FileProcessType)(Enum.Parse(typeof(FileProcessType), testConfig.type));
            TaskResultList.Add(new TaskResultEntry(Path.GetFileNameWithoutExtension(testConfig.path), "", testConfig.path, type));
        }
    }
}
