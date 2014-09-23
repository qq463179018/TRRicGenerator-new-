using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pdftron.PDF;

namespace PdfTronWrapper.Utility
{
    internal static class PdftronPointExtension
    {
        public static bool IsBelow(this Point point, LinePos lowRange)
        {
            if (lowRange == null)
                return true;

            return point.y <= lowRange.AxisValue;
        }
    }
}
