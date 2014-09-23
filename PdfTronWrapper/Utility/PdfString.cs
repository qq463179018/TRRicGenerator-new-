using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pdftron.PDF;

namespace PdfTronWrapper.Utility
{
    public class PdfString : IComparable<PdfString>
    {
        public string Words { get; set; }
        public Rect Position { get; set; }
        public int PageNumber { get; set; }

        public PdfString(string str, Rect pos, int pageNum)
        {
            this.Words = str;
            this.Position = pos;
            this.PageNumber = pageNum;
        }

        public override string ToString()
        {
            return this.Words;
        }

        public int CompareTo(PdfString other)
        {
            if (Math.Max(this.Position.y1, this.Position.y2) > Math.Max(other.Position.y1, other.Position.y2))
                return -1;

            if (Math.Max(this.Position.y1, this.Position.y2) < Math.Max(other.Position.y1, other.Position.y2))
                return 1;

            if (Math.Min(this.Position.x1, this.Position.x2) < Math.Min(other.Position.x1, other.Position.x2))
                return -1;

            if (Math.Min(this.Position.x1, this.Position.x2) > Math.Min(other.Position.x1, other.Position.x2))
                return 1;

            return 0;
        }
    }

}
