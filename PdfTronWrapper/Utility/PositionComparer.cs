//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:Providing the function to define which of the two rectangle appears firstly.
//-----
//-----------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pdftron.PDF;

namespace PdfTronWrapper.Utility
{
    /// <summary>
    /// Define which of the two rectangle appears firstly.
    /// </summary>
    internal class PositionComparer
    {
        //public static PDFDoc PdfDoc;

        //public static int PageNumber;

        public static int CompareRectPos(Rect rectA, Rect rectB)
        {
            if (rectA == null || rectB == null ||
                (rectA.x1 == rectB.x1 &&
                rectA.x2 == rectB.x2 &&
                rectA.y1 == rectB.y1 &&
                rectA.y2 == rectB.y2)
                )
                return 0;

            return GenerateCompareResult(CompareYValue(rectA.y1, rectB.y1),
                                                               CompareYValue(rectA.x1, rectB.x1));
        }

        static int GenerateCompareResult(int firstCompareResult, int secondCompareResult)
        {
            if (firstCompareResult != 0)
                return firstCompareResult;
            return secondCompareResult;
        }

        static int CompareYValue(double y1, double y2)
        {
            if (y1 == y2)
                return 0;

            return y1 > y2 ? -1 : 1;
        }

        static int CompareXValue(double x1, double x2)
        {
            if (x1 == x2)
                return 0;

            return x1 > x2 ? 1 : -1;
        }
    }
}
