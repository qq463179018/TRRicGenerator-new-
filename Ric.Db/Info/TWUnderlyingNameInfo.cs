using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Info
{
    public class TWUnderlyingNameInfo
    {
        public string UnderlyingRIC { get; set; } //Code
        public string OrganizationName { get; set; }
        public string EnglishDisplay { get; set; }
        public string ChineseDisplay { get; set; }
        public string ChineseChain { get; set; }

        public TWUnderlyingNameInfo()
        { }
    }
}
