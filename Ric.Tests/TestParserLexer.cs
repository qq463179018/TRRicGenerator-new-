using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ric.Tasks.Corax;
using System.Text.RegularExpressions;

namespace Ric.Tests
{
    [TestClass]
    public class TestParserLexer
    {
        [TestMethod]
        public void StartTest()
        {
            string test = "sfds55d66s76455dfg32434eertertre";
            string aaaa= (new Regex(@"\d+")).Replace(test,"");
            string qq = "";
            foreach (Match item in (new Regex(@"\d+")).Matches(test))
            {
                string tmp = item.Groups[0].ToString();
                test = test.Replace(tmp, "");
                qq = qq + "--" + tmp;

            }

        }
    }
}
