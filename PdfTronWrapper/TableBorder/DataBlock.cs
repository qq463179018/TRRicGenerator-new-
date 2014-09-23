//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:Indicate a series of connective letters in a pdf page.
//-----
//-----------------------------------------------------------------------------------------------------------------------


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using pdftron.PDF;

namespace PdfTronWrapper.TableBorder
{
    /// <summary>
    /// Indicate a series of connective letters in a pdf page.
    /// </summary>
    public class DataBlock : IComparable
    {
        string numberRegexString = @"-?(\d+,)*(\d+\.)?\d+|^-$";

        public DataType DataType
        {
            get
            {
                Regex numberRegex = new Regex(numberRegexString);
                MatchCollection match = numberRegex.Matches(Text);
                return match.Count == 1 ? DataType.Number : DataType.Text;
            }
        }

        public string Text { get; set; }
        public double LeftBound { get; set; }
        public double RightBound { get; set; }
        public HorBlocksInfo HorBlockDicKey { get; set; }
        public double LeftBlockDicKey { get; set; }
        public double RightBlockDicKey { get; set; }
        public double TopBound { get; set; }
        public double BottomBound { get; set; }
        public double Width
        {
            get
            {
                return RightBound - LeftBound;
            }
        }
        public double Height{ get; set; }
        public bool IsLeft { get; set; }
        public bool IsRight { get; set; }

        public bool InRegion(Rect rect)
        {
            return rect.Contains(HorCenter, VerCenter);
        }

        public double HorCenter
        {
            get
            {
                return (LeftBound + RightBound) / 2;
            }
        }

        public double VerCenter
        {
            get
            {
                return (TopBound + BottomBound) / 2;
            }
        }

        public List<int> ColumnIndexes { get; set; }

        public bool IsIntersect(List<DataColumn> textColumns)
        {
            return textColumns.Exists(column => IsIntersect(column));
        }

        /// <summary>
        /// Set the column indexes of the text blocks
        /// </summary>
        /// <param name="textColumns"></param>
        public void SetColumnIndexes(List<DataColumn> textColumns)
        {
            List<int> indexes = textColumns.Where(column => IsIntersect(column)).Select(column => textColumns.IndexOf(column)).ToList();
            ColumnIndexes = indexes;
        }

        public bool IsIntersect(DataColumn textColumn)
        {
            double floatValue = DataColumn.ColumnBlockIntersectFloatValue;
            return DataColumn.IsBetween(LeftBound, textColumn.LeftDataBound, textColumn.RightDataBound) ||
                DataColumn.IsBetween(RightBound, textColumn.LeftDataBound, textColumn.RightDataBound) ||

                DataColumn.IsBetween(textColumn.LeftDataBound, LeftBound, RightBound) ||
                DataColumn.IsBetween(textColumn.RightDataBound, LeftBound, RightBound);
        }

        /// <summary>
        /// Get the text columns which is intersected with the data block.
        /// </summary>
        /// <param name="textColumns"></param>
        /// <returns></returns>
        public List<DataColumn> GetIntersectColumns(List<DataColumn> textColumns)
        {
            return textColumns.Where(column => IsIntersect(column) || column.IsIntersect(this)).ToList();
        }

        public bool IsHorCenterBetween(double leftValue, double rightValue)
        {
            double centerValue=(LeftBound+RightBound)/2;
            return leftValue < centerValue && centerValue < rightValue;
        }

        public bool IsVerCenterBetween(double bottomValue, double topValue)
        {
            double centerValue = (TopBound + BottomBound) / 2;
            return bottomValue < centerValue && centerValue < topValue;
        }

        public bool IsHorScaleContains(double xValue, double horPosFloatValue)
        {
            return (LeftBound < xValue && Math.Abs(LeftBound-xValue)>horPosFloatValue)
                && (xValue < RightBound &&Math.Abs(xValue-RightBound)>horPosFloatValue );
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            try
            {
                DataBlock textBlock = obj as DataBlock;
                if (textBlock.LeftBound == LeftBound &&
                    textBlock.RightBound == RightBound &&
                    textBlock.TopBound == TopBound &&
                    textBlock.BottomBound == BottomBound)
                    return true;
                return false;
            }
            catch (Exception ex)
            {
                string err = ex.ToString();
                throw ex;
            }
        }

        /// <summary>
        /// Get two lines which is nearby the data block.
        /// </summary>
        /// <param name="horizontialLines"></param>
        /// <returns></returns>
        public FormLineList[] GetNearbyTwoLines(SortedDictionary<double, FormLineList> horizontialLines)
        {
            double averageXValue=(BottomBound+TopBound)/2;
            FormLineList[] nearbyLines = new  FormLineList[2] { null, null };
            for (int i = 0; i < horizontialLines.Count - 1; i++)
            {
                FormLineList currentLines = horizontialLines[horizontialLines.Keys.ToArray()[i]];
                FormLineList nextLines = horizontialLines[horizontialLines.Keys.ToArray()[i + 1]];
                if (currentLines.Exists(currentLine=>
                    averageXValue > currentLine.StartPoint.y && currentLine.IsCover(this)))
                {
                    nearbyLines[0] = currentLines;
                }
                if (nearbyLines[1] == null &&
                    nextLines.Exists(nextLine=>
                 averageXValue <= nextLine.StartPoint.y && nextLine.IsCover(this)))
                {
                    nearbyLines[1] = nextLines;
                }
                if (nearbyLines.All(line => line != null))
                {
                    break;
                }
            }
            return nearbyLines;
        }

        public void SetBlockDicKey(double key, bool isLeft)
        {
            if (isLeft)
            {
                LeftBlockDicKey = key;
            }
            else
            {
                RightBlockDicKey = key;
            }
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            DataBlock block = obj as DataBlock;
            if (BottomBound == block.BottomBound)
            {
                if (LeftBound == block.LeftBound)
                {
                    if (RightBound == block.RightBound)
                        return 0;
                    return RightBound > block.RightBound ? 1 : -1;
                }
                return LeftBound > block.LeftBound ? 1 : -1;
            }
            return BottomBound > block.BottomBound ? 1 : -1;
        }

        #endregion
    }
}
