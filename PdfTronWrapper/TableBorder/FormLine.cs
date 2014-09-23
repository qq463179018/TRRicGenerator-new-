//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:The line of financial sheet.
//-----
//-----------------------------------------------------------------------------------------------------------------------


using System;
using System.Collections.Generic;
using System.Linq;
using pdftron.PDF;

namespace PdfTronWrapper.TableBorder
{
    public class FormLine : IComparable, ICloneable
    {
        public Point StartPoint { get; set; }
        public Point EndPoint { get; set; }
        public bool IsTransverseLine { get; set; }
        public bool IsExistent { get; set; }
        public double Length
        {
            get
            {
                return IsTransverseLine ? Math.Abs(StartPoint.x - EndPoint.x) :
                    Math.Abs(StartPoint.y - EndPoint.y);
            }
        }

        public Point CenterPoint
        {
            get
            {
                return new Point
                {
                    x = (StartPoint.x + EndPoint.x) / 2,
                    y = (StartPoint.y + EndPoint.y) / 2
                };
            }
        }

        public FormLine(Point fromPoint, Point toPoint,bool isExistent)
        {
            IsTransverseLine = Math.Abs(fromPoint.y - toPoint.y) < Math.Abs(fromPoint.x - toPoint.x);
            
            Point _fromPoint=new Point{x=fromPoint.x,y=fromPoint.y};
            Point _toPoint=new Point{x=toPoint.x,y=toPoint.y};

            SetPoints(_fromPoint, _toPoint);
            IsExistent = isExistent;
        }

        void SetPoints(Point fromPoint, Point toPoint)
        {
            Point[] sortedPoints = SortPoints(fromPoint, toPoint);
            StartPoint = sortedPoints[0];
            EndPoint = sortedPoints[1];
        }

        Point[] SortPoints(Point fromPoint, Point toPoint)
        {
            Point startPoint, endPoint;
            Point[] sortedPoints=new Point[2];
            if (IsTransverseLine)
            {
                startPoint = fromPoint.x < toPoint.x ? fromPoint : toPoint;
                endPoint = fromPoint.x < toPoint.x ? toPoint : fromPoint;
            }
            else
            {
                startPoint = fromPoint.y < toPoint.y ? fromPoint : toPoint;
                endPoint = fromPoint.y < toPoint.y ?  toPoint: fromPoint;
            }
            sortedPoints[0] = startPoint;
            sortedPoints[1] = endPoint;
            return sortedPoints;
        }

        /// <summary>
        /// Generate a horizontial line by some parameters.
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="x2"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public static FormLine GenerateHorizontialLine(double x1, double x2, double y)
        {
            return new FormLine(
                new Point(x1, y),
                new Point(x2, y),
                false);
        }

        /// <summary>
        /// Generate a vertical line by some parameters.
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="x2"></param>
        /// <param name="y"></param>
        /// <returns></returns>        
        public static FormLine GenerateVerticalLine(double y1, double y2, double x)
        {
            return new FormLine(
                new Point(x, y1),
                new Point(x, y2),
                false);
        }

        public bool IsCover(DataBlock textBlock)
        {
            return StartPoint.x <= textBlock.LeftBound && EndPoint.x >= textBlock.RightBound;
        }

        public double GetDistance(FormLine formLine)
        {
            if (IsTransverseLine != formLine.IsTransverseLine)
                return 0;

            double curValue=IsTransverseLine?StartPoint.y:StartPoint.x;
            double otherValue=IsTransverseLine?formLine.StartPoint.y:formLine.StartPoint.x;

            double distance=Math.Abs(curValue-otherValue);
            return distance;
        }

        public double GetDistanceOfNearestExtremePoints(FormLine formLine)
        {
            List<Point> extremePoints=new List<Point>{StartPoint,EndPoint,formLine.StartPoint,formLine.EndPoint};
            List<double> axisValues=IsTransverseLine?extremePoints.Select(point=>point.x).ToList():
                extremePoints.Select(point => point.y).ToList();
            double distance = axisValues.Max() - axisValues.Min() - Length - formLine.Length;
            return distance < 0 ? 0 : distance;
        }

        public void MergeLine(FormLine formLine)
        {
            List<Point> extremePoints = new List<Point> { StartPoint, EndPoint, formLine.StartPoint, formLine.EndPoint };
            List<double> axisValues = IsTransverseLine ? extremePoints.Select(point => point.x).ToList() :
                extremePoints.Select(point => point.y).ToList();
            if (IsTransverseLine)
            {
                StartPoint.x = axisValues.Min();
                EndPoint.x = axisValues.Max();
            }
            else
            {
                StartPoint.y = axisValues.Min();
                EndPoint.y = axisValues.Max();
            }
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            int result = 0;
            FormLine formLine = obj as FormLine;
            if (IsTransverseLine == formLine.IsTransverseLine)
            {
                if (IsTransverseLine)
                {
                    result = Compare(StartPoint.y, formLine.StartPoint.y);
                    if (result == 0)
                    {
                        result = Compare(StartPoint.x, formLine.StartPoint.x);
                        if (result == 0)
                        {
                            result = Compare(EndPoint.x, formLine.EndPoint.x);
                        }
                    }
                }
                else
                {
                    result = Compare(EndPoint.x, formLine.EndPoint.x);
                    if (result == 0)
                    {
                        result = Compare( formLine.EndPoint.y,EndPoint.y);
                        if (result == 0)
                        {
                            result = Compare(formLine.StartPoint.y, StartPoint.y);
                        }
                    }
                }
            }
            return result;
        }

        int Compare(double value1, double value2)
        {
            if (value1 == value2)
                return 0;
            return value1 > value2 ? 1 : -1;
        }

        #endregion

        #region ICloneable Members

        public object Clone()
        {
            return new FormLine(new Point
                {
                    x=StartPoint.x,
                    y=StartPoint.y
                },
                new Point
                {
                    x = EndPoint.x,
                    y = EndPoint.y
                },IsExistent
            );
        }

        #endregion
    }
}
