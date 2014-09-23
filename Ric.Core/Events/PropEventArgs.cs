using System;
using System.Collections.Generic;

namespace Ric.Core.Events
{
    public class PropEventArgs : EventArgs
    {
        public Dictionary<string, string> Props { get; set; }

        public PropEventArgs(Dictionary<string, string> props)
        {
            Props = props;
        }

    }
}