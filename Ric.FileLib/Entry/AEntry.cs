using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.FileLib.Attribute;

namespace Ric.FileLib.Entry
{
    /// <summary>
    /// Basic file entry representation
    /// All specific file entry classes
    /// must inherit from it
    /// </summary>
    public abstract class AEntry
    {
        [TitleName("RIC")]
        public string Ric { get; set; }
    }
}
