//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:A series of data block composited a data column.
//-----
//-----------------------------------------------------------------------------------------------------------------------


using System;

namespace PdfTronWrapper.TableBorder
{
    public class DataColumn : IComparable
    {

        DataType _type = DataType.None;

        public DataType ColumnType
        {
            get
            {
                if (_type == DataType.None)
                {
                    _type=TextBlocks.DataType;
                }
                return _type;
            }
        }


        public double LeftBound { get; set; }

        public double RightBound { get; set; }

        public double LeftDataBound { get; set; }

        public double RightDataBound { get; set; }

        public double DataWidth
        {
            get
            {
                return RightDataBound - LeftDataBound;
            }
        }

        public DataBlockList TextBlocks {get;set;}

        public static double ColumnBlockIntersectFloatValue = 3;

        /// <summary>
        /// Indicate whether the data column is intersected with the other.
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public bool IsIntersect(DataColumn column)
        {
            return IsBetween(LeftDataBound, column.LeftDataBound, column.RightDataBound, ColumnBlockIntersectFloatValue) ||
                IsBetween(RightDataBound, column.LeftDataBound, column.RightDataBound, ColumnBlockIntersectFloatValue) ||

                IsBetween(column.LeftDataBound, LeftDataBound, RightDataBound, ColumnBlockIntersectFloatValue) ||
                IsBetween(column.RightDataBound, LeftDataBound, RightDataBound, ColumnBlockIntersectFloatValue);
        }

        /// <summary>
        /// Indicate whether the data column is intersected with a data block.
        /// </summary>
        /// <param name="block"></param>
        /// <returns></returns>
        public bool IsIntersect(DataBlock block)
        {
            return IsBetween(LeftDataBound, block.LeftBound, block.RightBound) ||
                IsBetween(RightDataBound, block.LeftBound, block.RightBound);
        }

        public static bool IsBetween(double num, double num1, double num2,double floatValue)
        {
            double maxNum = Math.Max(num1, num2);
            double minNum = Math.Min(num1, num2);
            return num >= minNum+floatValue && num <= maxNum-floatValue;
        }

        public static bool IsBetween(double num, double num1, double num2)
        {
            return IsBetween(num, num1, num2, 0);
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            DataColumn column = obj as DataColumn;
            if (LeftDataBound == column.LeftDataBound)
                return 0;
            return LeftDataBound < column.LeftDataBound ? -1 : 1;
        }

        #endregion
    }
}
