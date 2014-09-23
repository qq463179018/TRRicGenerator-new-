//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:Providing the function to find out the data area of a financial table.
//-----
//-----------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PdfTronWrapper;
using pdftron.PDF;
using pdftron;
using pdftron.Common;
using PdfTronWrapper.TableBorder;

namespace PdfTronWrapper.Utility
{
    internal class FormLineSearcher
    {

        #region Private Fields

        /// <summary>
        /// A helper for using pdftron
        /// </summary>
        PdfTronHelper pdfTronHelper;

        /// <summary>
        /// The pdf document object to be operated
        /// </summary>
        PDFDoc _pdfDoc;

        /// <summary>
        /// The min difference value of axis to distinct two points
        /// </summary>
        double lengthError = 4;

        /// <summary>
        /// The number of the pdf page which is operated currently.
        /// </summary>
        int _pageNum;

        /// <summary>
        /// The top bound of the operated scale on a pdf page.
        /// </summary>
        LinePos topBound = new LinePos() { AxisValue = -1 };

        /// <summary>
        /// The bottom bound of the operated scale on a pdf page.
        /// </summary>
        LinePos bottomBound = new LinePos() { AxisValue = -1 };

        /// <summary>
        /// The point object when the page is travelsaling.
        /// </summary>
        Point curPoint;

        /// <summary>
        /// Horizontal lines in a pdf page
        /// </summary>
        SortedDictionary<double, FormLineList> horizontalLines;

        /// <summary>
        /// Vertical lines in a pdf page
        /// </summary>
        SortedDictionary<double, FormLineList> verticalLines;

        LinePos _startPos, _endPos;

        double[] pageSize;

        #endregion

        #region Pubic Methods

        public SortedDictionary<double, FormLineList>[] GetFormLines(int pageNum, LinePos startPos, LinePos endPos
            ,bool isSubsequentPage,SortedDictionary<double, FormLineList> lastPageVerticalLines)
        {
            //Get the information of lines and extreme points by travelsaling the page
            horizontalLines = new SortedDictionary<double, FormLineList>();
            verticalLines = new SortedDictionary<double, FormLineList>();
            _pageNum = pageNum;
            _startPos = startPos;
            _endPos = endPos;
            
            Page page = _pdfDoc.GetPage(pageNum);
            pageSize = PdfTronHelper.GetPageSize(page);
            topBound.AxisValue = startPos == null ? pageSize[1] : startPos.AxisValue;
            bottomBound.AxisValue = endPos == null ? 0 : endPos.AxisValue;
            ii = 0;
            using (ElementReader page_reader = new ElementReader())
            {
                page_reader.Begin(page);
                ProcessElements(page_reader);
            }
            RemoveLittleLines(horizontalLines);
            RemoveLittleLines(verticalLines);
            //Remove short lines.
            Rect posRect = GetTablePosRect(_pageNum, _startPos, _endPos,isSubsequentPage);
            RemoveTooShortAndTooLongLines(page, horizontalLines, true, posRect);
            
            bool isNotNeedGenerateVerLines=
            isSubsequentPage && lastPageVerticalLines != null && verticalLines.Count == lastPageVerticalLines.Count;
            if (!isNotNeedGenerateVerLines)
            {
                RemoveTooShortAndTooLongLines(page, verticalLines, false, posRect);
            }
            //Generate drawed lines.
            Rect areaRect;

            bool existRealRect = horizontalLines.Count > 1 && verticalLines.Count > 1 && IsRect(posRect,horizontalLines,verticalLines);
            areaRect = existRealRect ? GenerateRectByLines() : posRect;
            if (existRealRect || isSubsequentPage)
            {
                RemoveLinesNotInRect(areaRect, horizontalLines, verticalLines);
                isNotNeedGenerateVerLines =
            isSubsequentPage && lastPageVerticalLines != null && verticalLines.Count == lastPageVerticalLines.Count;
            }
            else
            {
                posRect = GetTablePosRect(_pageNum, _startPos, _endPos, true);
            }
            FormLineGenerator.RemoveSpareNearLines(verticalLines, false);
            if (!isNotNeedGenerateVerLines)
            {
                isNotNeedGenerateVerLines =
                isSubsequentPage && lastPageVerticalLines != null && verticalLines.Count == lastPageVerticalLines.Count;
            }
            FormLineGenerator lineGenerator = new FormLineGenerator(page, areaRect);
            SortedDictionary<double, FormLineList>[] lines = lineGenerator.GetFormLines(existRealRect, horizontalLines, verticalLines, isNotNeedGenerateVerLines);
            return lines;
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor for TableBoundSearcher class
        /// </summary>
        /// <param name="pdfDoc">The pdf document object to be operated</param>
        public FormLineSearcher(PDFDoc pdfDoc)
        {
            _pdfDoc = pdfDoc;
            pdfTronHelper = new PdfTronHelper(pdfDoc);
        }

        #endregion

        #region Private Methods

        public static void RemoveLinesNotInRect(Rect rect, params SortedDictionary<double, FormLineList>[] lines)
        {
            foreach (SortedDictionary<double, FormLineList> lineDic in lines)
            {
                for (int j = 0; j < lineDic.Count; j++)
                {
                    FormLineList lineList = lineDic[lineDic.Keys.ToArray()[j]];
                    for (int i = 0; i < lineList.Count; )
                    {
                        if (!rect.ContainFormLine(lineList[i]))
                        {
                            lineList.RemoveAt(i);
                        }
                        else
                        {
                            i++;
                        }
                    }
                }
                GenericMethods<double, FormLineList>.RemoveZeroAmountValueItems(lineDic);
            }
        }

        Rect GenerateRectByLines()
        {
            double maxHorLength = horizontalLines.Max(line => line.Value.Max(x => x.Length));
            IEnumerable<Point> horExtremePoints = horizontalLines.SelectMany(line => line.Value)
                .Where(line=>line.Length>maxHorLength*0.9)
                .SelectMany(line =>
            {
                return new List<Point> { line.StartPoint, line.EndPoint };
            });
            IEnumerable<double> horXValues = horExtremePoints.Select(point => point.x);

            double maxVerLength = verticalLines.Max(line => line.Value.Max(x => x.Length));
            IEnumerable<Point> verExtremePoints = verticalLines.SelectMany(line => line.Value)
                .Where(line => line.Length > maxVerLength * 0.9)
                .SelectMany(line =>
            {
                return new List<Point> { line.StartPoint, line.EndPoint };
            });
            IEnumerable<double> verYValues = verExtremePoints.Select(point => point.y);

            double x1 = horXValues.Min();
            double y1 = verYValues.Min();
            double x2 = horXValues.Max();
            double y2 = verYValues.Max();
            Rect rect = new Rect(x1 - 3, y1 - 3, x2 + 3, y2 + 3);
            return rect;
        }

        Rect GetTablePosRect(int pageNum, LinePos topPos, LinePos bottomPos,bool isMatchNum)
        {
            double topAxisValue = topPos == null ? pdfTronHelper.GetTopPosOfPage(pageNum).AxisValue : topPos.AxisValue;
            double bottomPosAxisValue=bottomPos==null?0:bottomPos.AxisValue;
            if (bottomPosAxisValue == 0)
            {
                bottomPosAxisValue =
                    pdfTronHelper.GetBottomPosOfPage(pageNum, isMatchNum, bottomPosAxisValue, topAxisValue).AxisValue;
            }
            double[] leftRightTextBounds = pdfTronHelper.GetLeftRightTextBounds(pageNum);

            Rect rect = new Rect(leftRightTextBounds[0] - 10, bottomPosAxisValue, leftRightTextBounds[1] + 10, topAxisValue);
            return rect;
        }

        void RemoveLittleLines(SortedDictionary<double, FormLineList> diclines)
        {
            foreach (double key in diclines.Keys)
            {
                FormLineList lines = diclines[key];
                lines.Where(line => line.Length < 1).ToList().ForEach(line => lines.Remove(line));
            }
            diclines.Where(pair => pair.Value.Count == 0).ToList().ForEach(pair => diclines.Remove(pair.Key));
        }

        /// <summary>
        /// Remove short lines from line dictionary.
        /// </summary>
        /// <param name="diclines">Line dictionary</param>
        void RemoveTooShortAndTooLongLines(Page page, SortedDictionary<double, FormLineList> diclines, bool isHorizontial,Rect posRect)
        {
            double[] pageSize = PdfTronHelper.GetPageSize(page);

            double maxLength = isHorizontial ? pageSize[0] : pageSize[1];
            diclines.Where(pair => pair.Value.Exists(line => Math.Abs(line.Length - maxLength) < 3))
                .Select(pair => pair.Key).ToList()
                .ForEach(key => diclines.Remove(key));
            FormLineList _lines = new FormLineList(diclines.SelectMany(pair => pair.Value).ToList());

            if (_lines.Count > 1)
            {
                if (isHorizontial)
                {
                    double[] textLeftRightXValue = pdfTronHelper.GetLeftRightTextBounds(page);
                    maxLength = (textLeftRightXValue[1] - textLeftRightXValue[0]);
                    diclines.Where(x => x.Value.Sum(line => line.Length) < maxLength * 0.5
                   ).Select(x => x.Key).ToList().ForEach(key => diclines.Remove(key));
                    foreach (double key in diclines.Keys.ToArray())
                    {
                        FormLineList lines = diclines[key];
                        if (lines.Count < 2)
                        {
                            continue;
                        }
                        double _maxLength = lines.Max(line => line.Length);
                        FormLine maxLengthLine = lines.Find(line => line.Length == _maxLength);
                        lines.Where(line => line.Length < (_maxLength * 0.7)).ToList().ForEach(line => lines.Remove(line));
                    }
                    FormLineList templines = new FormLineList(diclines.SelectMany(pair => pair.Value).ToList());

                    if (templines.Count > 1)
                    {
                        maxLength = templines.Select(line => line.Length).Max();
                        double scale = 0.4;
                        double minLength = maxLength * scale;
                        IEnumerable<double> shortLineKeys = diclines.Where(
                        x => x.Value.Sum(line => line.Length) < minLength
                        ).Select(x => x.Key);
                        shortLineKeys.ToList().ForEach(key => diclines.Remove(key));
                    }
                }
                else
                {
                    maxLength = posRect.Height();
                    if (posRect.Height() < 300)
                    {
                        maxLength = _lines.Select(line => line.Length).Max();
                    }
                    double minLength = maxLength * 0.4;

                    if (minLength < 9)
                        minLength = 9;

                    IEnumerable<double> shortLineKeys = diclines.Where(
                        x => x.Value.Sum(line => line.Length) < minLength
                        ).Select(x => x.Key);
                    shortLineKeys.ToList().ForEach(key => diclines.Remove(key));
                }
            }
            else
            {
                diclines.Clear();
            }
        }

        /// <summary>
        /// Check whether the lines can composite a rectangle.
        /// </summary>
        /// <returns>If the lines can composite a rectangle,return true;Otherwise,return false;</returns>
        public static bool IsRect(Rect posRect, SortedDictionary<double, FormLineList> _horizontalLines, SortedDictionary<double, FormLineList> _verticalLines)
        {
            double errorScale = 0.2;

            double horInstance = _horizontalLines.Last().Key - _horizontalLines.First().Key;

            if (horInstance < posRect.Height() * 0.5)
                return false;

            bool firstLineValidResult = Math.Abs((_verticalLines.Last().Value.Sum(line => line.Length) / horInstance) - 1) < errorScale;
            bool lastLineValidResult = Math.Abs((_verticalLines.First().Value.Sum(line=>line.Length) / horInstance) - 1) < errorScale;
            if (!firstLineValidResult || !lastLineValidResult)
                return false;

            double verInstance = _verticalLines.Last().Key - _verticalLines.First().Key;

            if (verInstance < posRect.Width() * 0.7)
                return false;

            firstLineValidResult = Math.Abs((_horizontalLines.Last().Value.Sum(line=>line.Length)/ verInstance) - 1) < errorScale;
            lastLineValidResult = Math.Abs((_horizontalLines.First().Value.Sum(line => line.Length) / verInstance) - 1) < errorScale;
            return firstLineValidResult && lastLineValidResult;
        }

        /// <summary>
        /// Process the elements of a pdf page.
        /// </summary>
        /// <param name="reader">ElementReader object</param>
        void ProcessElements(ElementReader reader)
        {
            Element element;
            while ((element = reader.Next()) != null)
            {
                Element.Type elementType=element.GetType();
                switch (elementType)
                {
                    case Element.Type.e_path:          // Process path data...
                        ProcessPath(element);
                        break;
                }
                reader.ClearChangeList();
            }
        }

        /// <summary>
        /// Deal the operation fo drawing line.
        /// </summary>
        /// <param name="fromPoint">One extreme point of the line</param>
        /// <param name="toPoint">The other extreme point of the line</param>
        void DealDrawLine(Point fromPoint, Point toPoint)
        {
            if (!ValidateInScale(fromPoint) || !ValidateInScale(toPoint))// || IsSamePoint(fromPoint, toPoint))
            {
                return;
            }

            FormLine newLine = new FormLine(fromPoint, toPoint,true);
            AddLine(newLine);
        }

        /// <summary>
        /// Binary search
        /// </summary>
        /// <param name="low">The low index of the array</param>
        /// <param name="high">The high index of the array</param>
        /// <param name="Find_Name">The number to search</param>
        /// <param name="nums">The number array</param>
        /// <returns>Search result</returns>
        double HalfFind(int low, int high, double Find_Name, double[] nums)
        {
            while (low <= high)
            {
                int mid = (low + high) / 2;
                if (Math.Abs(nums[mid] - Find_Name) <= lengthError)
                    return nums[mid];
                else
                {
                    if (Find_Name > nums[mid])
                        low = mid + 1;
                    else high = mid - 1;
                    return HalfFind(low, high, Find_Name, nums);
                }
            }

            List<CustomPair<double, double>> pairs = new List<CustomPair<double, double>>();
            for (int i = 0; i < nums.Length; i++)
            {
                double differValue=Math.Abs(nums[i]-Find_Name);
                if (differValue <= lengthError)
                {
                    pairs.Add(new CustomPair<double, double>() { Key = differValue, Value = nums[i] });
                }
            }
            if (pairs.Count > 0)
            {
                pairs.Sort();
                return pairs[0].Value;
            }
            return -1;
        }

        /// <summary>
        /// Add line information to list.
        /// </summary>
        /// <param name="newLine">A line object</param>
        void AddLine(FormLine newLine)
        {
            if (newLine.IsTransverseLine)
            {
                AddLineToList(ref newLine, ref horizontalLines);
            }
            else
                AddLineToList(ref newLine, ref verticalLines);
        }

        /// <summary>
        /// Add a new line to line dictionary.
        /// </summary>
        /// <param name="newLine">A new line.</param>
        /// <param name="lineDic">Line dictionary</param>
        void AddLineToList(ref FormLine newLine, ref SortedDictionary<double, FormLineList> lineDic)
        {
            double newAxisValue = newLine.IsTransverseLine ? newLine.StartPoint.y : newLine.StartPoint.x;
            double[] existAxisValues = lineDic.Keys.ToArray();
            double findResult = HalfFind(0, existAxisValues.Length - 1, newAxisValue, existAxisValues);

            if (findResult == -1)
            {
                FormLineList newList = new FormLineList() { newLine };
                lineDic.Add(newAxisValue, newList);
            }
            else
            {
                FormLineList lineList = lineDic[findResult];
                lineList.Add(newLine);

                FormLineList matchLines = new FormLineList();
                foreach (FormLine line in lineList)
                    if (HasRepeatPart(line, newLine))
                        matchLines.Add(line);

                MergeLines(matchLines, lineList);
            }
        }

        /// <summary>
        /// Merge two line to one line.
        /// </summary>
        /// <param name="mergeLines">Line to be merged.</param>
        void  MergeLines(FormLineList mergeLines, FormLineList sameLevelLines)
        {
            if (mergeLines == null || mergeLines.Count < 2)
                return;

            FormLine firstLine = mergeLines[0];
            bool isHorizontal = firstLine.IsTransverseLine;
            if (isHorizontal)
            {
                List<double> xValues = mergeLines.Select(line => line.StartPoint.x).
                    Concat(mergeLines.Select(line => line.EndPoint.x)).ToList();
                firstLine.StartPoint.x = xValues.Min();
                firstLine.EndPoint.x = xValues.Max();
            }
            else
            {
                List<double> yValues = mergeLines.Select(line => line.StartPoint.y).
                    Concat(mergeLines.Select(line => line.EndPoint.y)).ToList();
                firstLine.StartPoint.y = yValues.Min();
                firstLine.EndPoint.y = yValues.Max();
            }
            for (int i = 1; i < mergeLines.Count; i++)
                sameLevelLines.Remove(mergeLines[i]);
        }

        /// <summary>
        /// Indicate whether two line is on the same level.
        /// </summary>
        /// <param name="line1">One line</param>
        /// <param name="line2">The other line</param>
        /// <returns>If two line is on the same level,return true;Otherwise,return false.</returns>
        bool OnSameLine(FormLine line1, FormLine line2)
        {
            Func<double, double, bool> func = (value1, value2) =>
            {
                return Math.Abs(value1 - value2) < lengthError;
            };
            if (line1.IsTransverseLine)
            {
                return func(line1.EndPoint.y, line2.StartPoint.y);
            }
            return func(line1.EndPoint.x, line2.StartPoint.x);
        }

        /// <summary>
        /// Indicate whether two line has repeat part.
        /// </summary>
        /// <param name="line1">One line</param>
        /// <param name="line2">The other line</param>
        /// <returns>If two line has repeat part.,return true;Otherwise,return false.</returns>
        bool HasRepeatPart(FormLine line1, FormLine line2)
        {
            if (line1.IsTransverseLine)
            {
                return IsBetween(line1.StartPoint.x, line2.StartPoint.x, line2.EndPoint.x) ||
                    IsBetween(line1.EndPoint.x, line2.StartPoint.x, line2.EndPoint.x) ||
                    IsBetween(line2.StartPoint.x, line1.StartPoint.x, line1.EndPoint.x) ||
                    IsBetween(line2.EndPoint.x, line1.StartPoint.x, line1.EndPoint.x);
            }
            return IsBetween(line1.StartPoint.y, line2.StartPoint.y, line2.EndPoint.y) ||
                IsBetween(line1.EndPoint.y, line2.StartPoint.y, line2.EndPoint.y) ||
                IsBetween(line2.StartPoint.y, line1.StartPoint.y, line1.EndPoint.y) ||
                IsBetween(line2.EndPoint.y, line1.StartPoint.y, line1.EndPoint.y);
        }

        /// <summary>
        /// Indicate whether a number is between num1 and num2.
        /// </summary>
        /// <param name="num">number</param>
        /// <param name="num1">number</param>
        /// <param name="num2">number</param>
        /// <returns>If two line has repeat part.,return true;Otherwise,return false.</returns>
        bool IsBetween(double num, double num1, double num2)
        {
            double maxNum = Math.Max(num1, num2);
            double minNum = Math.Min(num1, num2);
            return (num >= minNum || minNum - num < lengthError) &&
                (num <= maxNum || num - maxNum < lengthError);
        }

        /// <summary>
        /// Indicate whether the line is horizontal
        /// </summary>
        /// <param name="fromPoint">One extreme point.</param>
        /// <param name="toPoint">The other extreme point.</param>
        /// <returns>If  the line is horizontal,return true;Otherwise,return false.</returns>
        bool IsHorizontal(Point fromPoint, Point toPoint)
        {
            double horDifference = Math.Abs(fromPoint.y - toPoint.y);
            double verDifference = Math.Abs(fromPoint.x - toPoint.x);
            return horDifference < verDifference && horDifference < lengthError;
        }

        /// <summary>
        /// Indicate whether the point is in the scale.
        /// </summary>
        /// <param name="point">A point.</param>
        /// <returns>If  the point is in the scale,return true;Otherwise,return false.</returns>
        bool ValidateInScale(Point point)
        {
            return (topBound.AxisValue == -1 || point.IsBelow(topBound)) &&
                (bottomBound.AxisValue == -1 || !point.IsBelow(bottomBound));
        }

        /// <summary>
        /// Indicate whether the points is the same one.
        /// </summary>
        /// <param name="point1">One point.</param>
        /// <param name="point2">The other point.</param>
        /// <returns>If the points is the same one,return true;Otherwise,return false.</returns>
        bool IsSamePoint(Point point1, Point point2)
        {
            return Math.Abs(point1.x - point2.x) < lengthError && 
                Math.Abs(point1.y - point2.y) < lengthError;
        }
        int ii= 0;
        /// <summary>
        /// Process a path
        /// </summary>
        /// <param name="path">A path object which is processed.</param>
        void ProcessPath(Element path)
        {
            if (!path.IsFilled() && !path.IsStroked())
                return;

            try
            {
                pdftron.PDF.PathData pathData = path.GetPathData();
                double[] data = pathData.points;
                ii++;
                if (ii == 627)
                {
                }
                if (verticalLines.Count >0 && verticalLines.First().Value.Count == 2)
                {
                }

                Matrix2D matrix2D = path.GetCTM();
                //Mark
                Matrix2D pageMatrix2D = _pdfDoc.GetPage(_pageNum).GetDefaultMatrix();
                //Matrix2D exchangeMatrix = new Matrix2D(Math.Abs(pageMatrix2D.m_a),
                //    Math.Abs(pageMatrix2D.m_b), Math.Abs(pageMatrix2D.m_c), Math.Abs(pageMatrix2D.m_d), 0, 0);

                Matrix2D exchangeMatrix = new Matrix2D(pageMatrix2D.m_a,
                    pageMatrix2D.m_b, pageMatrix2D.m_c, pageMatrix2D.m_d, 0, 0);
                for (int i = 0; i < data.Length; i += 2)
                {
                    matrix2D.Mult(ref data[i], ref data[i + 1]);
                }

                int data_sz = data.Length;

                byte[] opr = pathData.operators;
                int opr_sz = opr.Length;

                int opr_itr = 0, opr_end = opr_sz;
                int data_itr = 0, data_end = data_sz;
                double x1, y1, x2, y2, x3, y3;
                for (; opr_itr < opr_end; ++opr_itr)
                {
                    switch ((pdftron.PDF.PathData.PathSegmentType)((int)opr[opr_itr]))
                    {
                        case pdftron.PDF.PathData.PathSegmentType.e_moveto:
                            x1 = data[data_itr]; ++data_itr;
                            y1 = data[data_itr]; ++data_itr;
                            pageMatrix2D.Mult(ref x1, ref y1);
                            curPoint = new Point(x1, y1);
                            break;
                        case pdftron.PDF.PathData.PathSegmentType.e_lineto:
                            x1 = data[data_itr]; ++data_itr;
                            y1 = data[data_itr]; ++data_itr;
                            pageMatrix2D.Mult(ref x1, ref y1);
                            Point point = new Point(x1, y1);
                            DealDrawLine(curPoint, point);
                            curPoint = point;
                            break;
                        case pdftron.PDF.PathData.PathSegmentType.e_cubicto:
                            x1 = data[data_itr++];
                            y1 = data[data_itr++];
                            x2 = data[data_itr++];
                            y2 = data[data_itr++];
                            x3 = data[data_itr++];
                            y3 = data[data_itr++];
                            break;
                        case pdftron.PDF.PathData.PathSegmentType.e_rect:
                            {
                                x1 = data[data_itr++];
                                y1 = data[data_itr++];
                                pageMatrix2D.Mult(ref x1, ref y1);
                                double w = data[data_itr++];
                                double h = data[data_itr++];
                                exchangeMatrix.Mult(ref w, ref h);
                                x2 = x1 + w;
                                y2 = y1;
                                x3 = x2;
                                y3 = y1 + h;
                                double x4 = x1;
                                double y4 = y3;
                                Point point1 = new Point(x1, y1);
                                Point point2 = new Point(x2, y2);
                                Point point3 = new Point(x3, y3);
                                Point point4 = new Point(x4, y4);
                                if (Math.Abs(w) > lengthError)
                                {
                                    DealDrawLine(point1, point2);
                                    if (Math.Abs(h) > lengthError)
                                    {
                                        DealDrawLine(point4, point3);
                                    }
                                }
                                if (Math.Abs(h) > lengthError)
                                {
                                    DealDrawLine(point3, point2);
                                    if (Math.Abs(w) > lengthError)
                                    {
                                        DealDrawLine(point4, point1);
                                    }
                                }
                                if (Math.Abs(w) <= lengthError && Math.Abs(h) <= lengthError)
                                {
                                    if (w > h)
                                    {
                                        DealDrawLine(point1, point2);
                                    }
                                    else
                                    {
                                        DealDrawLine(point4, point1);
                                    }
                                }
                                break;
                            }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        #endregion

    }
}
