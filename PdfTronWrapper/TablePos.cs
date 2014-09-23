//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:The function of the classes is helping to locate the data area of financial tables.
//-----
//-----------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pdftron.PDF;
using PdfTronWrapper.TableBorder;

namespace PdfTronWrapper
{
    public class TablePos
    {
        /// <summary>
        /// The page number of the pdf page.
        /// </summary>
        public int PageNum
        {
            get;
            set;
        }

        public string TableNameText { get; set; }

        /// <summary>
        /// The rectangle scale of the table’s content
        /// </summary>
        public pdftron.PDF.Rect TableRect
        {
            get;
            set;
        }

        public SortedDictionary<double, FormLineList>HorizontialLines
        {
            get;
            set;
        }

        public SortedDictionary<double, FormLineList> VerticalLines
        {
            get;
            set;
        }

        public double Height
        {
            get
            {
                double height = 0;
                if (HorizontialLines.Count > 1)
                {
                    height = HorizontialLines.First().Value[0].GetDistance(HorizontialLines.Last().Value[0]);
                }
                return height;
            }
        }
    }

    public class LinePos : ICloneable, IComparable
    {
        public LinePos()
        {
            PageNum = 1;
            AxisValue = 0;
            AxisValueWithLineHeight = 0;
            TrimText = "";
            IsEstimated = false;
        }

        public bool IsEstimated 
        {
            get; set; 
        }

        /// <summary>
        /// The page number of the line.
        /// </summary>
        public int PageNum
        {
            get;
            set;
        }

        /// <summary>
        /// The value of Axis
        /// </summary>
        public double AxisValue
        {
            get;
            set;
        }

        /// <summary>
        /// The value of Axis with LineHeight added.
        /// </summary>
        public double AxisValueWithLineHeight
        {
            get;
            set;
        }

        /// <summary>
        /// The text in the position
        /// </summary>
        public string TrimText
        {
            get;
            set;
        }

        public bool IsBetween(LinePos lowRange, LinePos maxRange)
        {
            return (lowRange == null || this > lowRange) &&
                (maxRange == null || this < maxRange);
        }

        public static bool operator >(LinePos range1, LinePos range2)
        {
            return range1.CompareTo(range2) > 0;
        }

        public static bool operator >=(LinePos range1, LinePos range2)
        {
            return range1.CompareTo(range2) >= 0;
        }

        public static bool operator <(LinePos range1, LinePos range2)
        {
            return range1.CompareTo(range2) < 0;
        }

        public static bool operator <=(LinePos range1, LinePos range2)
        {
            return range1.CompareTo(range2) <= 0;
        }

        #region ICloneable Members

        public object Clone()
        {
            return new LinePos
            {
                PageNum = PageNum,
                AxisValue = AxisValue,
                TrimText = TrimText,
                AxisValueWithLineHeight = AxisValueWithLineHeight
            };
        }

        #endregion

        #region IComparable Members

        public int CompareTo(object obj)
        {
            LinePos linePos = obj as LinePos;
            if (PageNum < linePos.PageNum)
                return -1;

            if (PageNum == linePos.PageNum)
            {
                if (AxisValue == linePos.AxisValue)
                    return 0;

                return AxisValue < linePos.AxisValue ? 1 : -1;
            }

            if (PageNum > linePos.PageNum)
                return 1;

            return 0;
        }

        #endregion
    }
}
