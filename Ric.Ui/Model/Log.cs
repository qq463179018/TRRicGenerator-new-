using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Ric.Ui.Model
{
    public class Log
    {
        public string Message { get; set; }

        public string ColorText { get; set; }

        public Log(string message, string colorText)
        {
            Message = message;
            ColorText = colorText;
        }

        public Log()
        {

        }
    }
}
