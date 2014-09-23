using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace PdfTronWrapper
{
    public class FreeTableRow : List<FreeTableCell>
    {
        public object Tag { get; set; }

        public override string ToString()
        {
            return this.Aggregate(string.Empty, (n, m) => n + m.Value);
        }
    }
}
