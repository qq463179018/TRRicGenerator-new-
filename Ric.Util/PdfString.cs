using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pdftron.PDF;

namespace Ric.Util
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

        public static PdfStringComparer GetComparer()
        {
            return new PdfString.PdfStringComparer();
        }

        public int CompareTo(PdfString rhs)
        {
            return this.Position.x2.CompareTo(rhs.Position.x2);
        }

        /// <summary>
        ///comparer by Horizontal or vertical 
        /// </summary>
        /// <param name="rhs"></param>
        /// <param name="which">comparer type</param>
        /// <returns></returns>
        public int CompareTo(PdfString rhs, PdfString.PdfStringComparer.ComparisonType which)
        {
            switch (which)
            {
                case PdfStringComparer.ComparisonType.Horizontal:
                    return this.Position.x2.CompareTo(rhs.Position.x2);
                case PdfStringComparer.ComparisonType.Vertical:
                    return this.Position.y2.CompareTo(rhs.Position.y2);
            }

            return 0;
        }

        /// <summary>
        /// for comparer custom comparer type
        /// </summary>
        public class PdfStringComparer : IComparer<PdfString>
        {
            private PdfString.PdfStringComparer.ComparisonType whichComparison;

            public enum ComparisonType { Horizontal, Vertical };

            public int Compare(PdfString lhs, PdfString rhs)
            {
                return lhs.CompareTo(rhs, whichComparison);
            }

            public PdfString.PdfStringComparer.ComparisonType WhichComparison
            {
                get { return whichComparison; }
                set { whichComparison = value; }
            }
        }
    }
}
