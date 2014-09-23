using System;

namespace Ric.FileLib.Attribute
{
    /// <summary>
    /// Custom attrivute used for conversion
    /// between NdaEntry parameter and field title.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class TitleName : System.Attribute
    {
        public TitleName(string name)
        {
            Name = name;
        }
        public string Name { get; private set; }
    }
}
