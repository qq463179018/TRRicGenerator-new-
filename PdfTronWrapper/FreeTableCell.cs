using System;
using System.Collections.Generic;

namespace PdfTronWrapper
{
    public class FreeTableCell
    {
        public string Value { get; set; }

        public int RowSpan { get; set; }

        public int ColSpan { get; set; }

        public object Tag { get; set; }

        public FreeTableCell Clone()
        {
            return (FreeTableCell)MemberwiseClone();
        }

        public override string ToString()
        {
            return Value ?? string.Empty;
        }
    }
}
