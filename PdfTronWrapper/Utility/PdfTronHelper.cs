//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:A helper class for pdftron operating
//-----
//-----------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using pdftron.Common;
using pdftron.PDF;

using PdfTronWrapper.TableBorder;
using PdfTronWrapper;
using System.Text;

namespace PdfTronWrapper.Utility
{
    public class PdfTronHelper
    {

        #region Private Fields

        /// <summary>
        /// The pdf document object to be operated
        /// </summary>
        PDFDoc _pdfDoc;

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor for PdfTronHelper class
        /// </summary>
        /// <param name="_pdfDoc">The pdf document object to be operated</param>
        public PdfTronHelper(PDFDoc pdfDoc)
        {
            _pdfDoc = pdfDoc;
        }

        #endregion

        #region Pubic Methods

        public static double[] GetPageSize(Page page)
        {
            Rect cropBoxRect = page.GetCropBox();
            Matrix2D pageMatrix2D = page.GetDefaultMatrix();

            Matrix2D exchangeMatrix = new Matrix2D(Math.Abs(pageMatrix2D.m_a),
                Math.Abs(pageMatrix2D.m_b), Math.Abs(pageMatrix2D.m_c), Math.Abs(pageMatrix2D.m_d), 0, 0);
            double width = cropBoxRect.Width();
            double height = cropBoxRect.Height();
            exchangeMatrix.Mult(ref width, ref height);
            return new double[2] { width, height };
        }

        /// <summary>
        /// Get the pos of the bottom bound of a pdf page.
        /// </summary>
        /// <param name="pageNum">The number of the pdf page.</param>
        /// <returns>Return the pos of the bottom bound of a pdf page.</returns>
        public LinePos GetBottomPosOfPage(int pageNum,bool isMatchNum,double lowPos,double highPos)
        {
            if (pageNum > _pdfDoc.GetPageCount() || pageNum < 1)
                return null;

            LinePos linePos = new LinePos() { PageNum = pageNum };

            linePos.AxisValue = linePos.AxisValueWithLineHeight =
                GetBottomTextYValue(pageNum,
                lowPos==-1?0:lowPos, 
                highPos==-1?GetTopPosOfPage(pageNum).AxisValue:highPos,isMatchNum);

            return linePos;
        }

        /// <summary>
        /// Get the pos of the top bound of a pdf page.
        /// </summary>
        /// <param name="pageNum">The number of the pdf page.</param>
        /// <returns>Return the pos of the top bound of a pdf page.</returns>
        public LinePos GetTopPosOfPage(int pageNum)
        {
            double[] pageSize = GetPageSize(pageNum);
            LinePos linePos = new LinePos() { PageNum = pageNum };
            double axisValue = 0;
            axisValue = pageSize[1];
            linePos.AxisValueWithLineHeight = linePos.AxisValue = axisValue;
            return linePos;
        }

        /// <summary>
        /// Search a result and return
        /// </summary>
        /// <param name="searcher">A TextSearch object</param>
        /// <param name="lowRange">The low range of the search scale</param>
        /// <param name="highRange">The high range of the search scale</param>
        /// <param name="isDone">Indicate whether the search is over</param>
        /// <returns>A list of LinePos object found in two pages</returns>
        public List<LinePos> Search(string regex, LinePos lowRange, LinePos highRange, ref int startPage, bool isSearchUp = false)
        {
            int pageNum = 0;
            String resultStr = "", ambientStr = "";
            Highlights hlts = new Highlights();
            List<LinePos> linePoses = new List<LinePos>();
            TextSearch searcher = InitAndBeginSearch(regex, startPage, startPage, isSearchUp);
            while (true)
            {
                pdftron.PDF.TextSearch.ResultCode resultCode;
                resultCode = searcher.Run(ref pageNum, ref  resultStr, ref ambientStr, hlts);
                if (resultCode == TextSearch.ResultCode.e_found)
                {
                    LinePos linePos = GetLinePos(regex, hlts, lowRange);
                    if (linePos == null)
                    {
                        continue;
                    }
                    if (linePos.IsBetween(lowRange, highRange))
                        linePoses.Add(linePos);
                }
                if (resultCode == TextSearch.ResultCode.e_done)
                {
                    startPage += isSearchUp ? -1 : 1;
                    if (linePoses.Count > 0)
                    {
                        break;
                    }
                    else
                    {
                        int highPage = GetHighPage(highRange);
                        int lowPage = GetLowPage(lowRange);
                        if ((!isSearchUp && startPage > highPage) || (isSearchUp && startPage < lowPage))
                        {
                            break;
                        }
                        searcher = InitAndBeginSearch(regex, startPage, startPage, isSearchUp);
                    }
                }
            }
            linePoses.Sort();
            return linePoses;
        }

        /// <summary>
        /// Search by the condition provided by parameters and return the first result.
        /// </summary>
        /// <param name="regex">The regex to match the text.</param>
        /// <param name="lowRange">The low range of the search scale</param>
        /// <param name="highRange">The high range of the search scale</param>
        /// <returns>If there are some search result,return the first one;Otherwise,return null.</returns>
        public LinePos Search(string regex, LinePos lowRange, LinePos highRange)
        {
            bool isDone = false;
            int startPage = GetLowPage(lowRange);

            while (!isDone)
            {
                List<LinePos> linePoses = Search(regex, lowRange, highRange, ref startPage);
                if (linePoses.Count > 0)
                    return linePoses[0];
                isDone = startPage > GetHighPage(highRange);
            }

            return null;
        }

        public void RevertTransportRect(int pageNum, Rect bounds)
        {
            Page page = _pdfDoc.GetPage(pageNum);
            Matrix2D matrix = page.GetDefaultMatrix().Inverse();
            ApplyMatrixToRect(matrix, bounds);
        }

        public int GetLowPage(LinePos lowPos)
        {
            return lowPos == null ? 1 : lowPos.PageNum;
        }

        public double[] GetLeftRightTextBounds(int pageNum)
        {
            PageTextExtractor pageProcess = null;

            if (!_pageProcessCache.TryGetValue(pageNum, out pageProcess))
            {
                pageProcess = new PageTextExtractor(_pdfDoc.GetPage(pageNum));

                _pageProcessCache.Add(pageNum, pageProcess);
            }

            return pageProcess.LeftRightBounds;
        }

        public double[] GetLeftRightTextBounds(Page page)
        {
            int pageNumber = page.GetIndex();
            return GetLeftRightTextBounds(pageNumber);
        }

        /// <summary>
        /// Validate the distance between two position.
        /// </summary>
        /// <param name="posUpon">The high positon.</param>
        /// <param name="posBelow">The low position.</param>
        /// <param name="maxInstance">The max distance allowed between two positions.</param>
        /// <returns>Return true if the distance is in the max distance,otherwise return false.</returns>
        public bool ValidateInstance(LinePos posUpon, LinePos posBelow, double maxInstance)
        {
            int pageSpan = posBelow.PageNum - posUpon.PageNum;
            if (pageSpan > 1)
                return false;
            if (pageSpan == 1)
            {
                double pageEndAxisValue =0;
                double toPageEndDistance = Math.Abs(posUpon.AxisValue - pageEndAxisValue);
                toPageEndDistance -= GetBlankAreaHeight(posUpon);
                double fromPageStartDistance = Math.Abs(posBelow.AxisValue - GetTopPosOfPage(posBelow.PageNum).AxisValue);
                return toPageEndDistance + fromPageStartDistance <= maxInstance + 40;
            }
            return Math.Abs(posBelow.AxisValue - posUpon.AxisValue) <= maxInstance;
        }

        public string GetPDFContent()
        {
            StringBuilder stringBuilder = new StringBuilder();
            int pageCount = _pdfDoc.GetPageCount();
            for (int i = 1; i <= pageCount; i++)
            {
                string pageContent = GetPdfPageContent(_pdfDoc, i);
                stringBuilder.Append(pageContent);
            }
            return stringBuilder.ToString();
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Get the  scale of hightlights text.
        /// </summary>
        /// <param name="hightLights">The hightlights information</param>
        /// <param name="_pdfDoc">The pdf document object</param>
        /// <returns>The line scale of hightlights text</returns>
        List<Rect> GetLineRect(Highlights hightLights)
        {
            List<Rect> rects = GetHightlightRect(hightLights);
            hightLights.Begin(_pdfDoc);
            double[] leftRightTextBounds = GetLeftRightTextBounds(hightLights.GetCurrentPageNumber());
            for (int i = 0; i < rects.Count; i++)
            {
                rects[i] = new Rect(leftRightTextBounds[0], rects[i].y1, leftRightTextBounds[1], rects[i].y2);
            }
            return rects;
        }

        /// <summary>
        /// Get the scale of hightlights text.
        /// </summary>
        /// <param name="hightLights">The hightlights information</param>
        /// <param name="_pdfDoc">The pdf document object</param>
        /// <returns>The scale of hightlights text</returns>
        List<Rect> GetHightlightRect(Highlights hightLights)
        {
            List<Rect> rects = new List<Rect>();
            hightLights.Begin(_pdfDoc);
            int pageNumber = hightLights.GetCurrentPageNumber();
            while (hightLights.HasNext())
            {
                Page page = _pdfDoc.GetPage(pageNumber);
                Matrix2D matrix = page.GetDefaultMatrix();

                double[] quads = hightLights.GetCurrentQuads();
                for (int i = 0; i < quads.Length; i += 2)
                {
                    matrix.Mult(ref quads[i], ref quads[i + 1]);
                }

                int quad_count = quads.Length / 8;
                for (int i = 0; i < quad_count; ++i)
                {
                    //assume each quad is an axis-aligned rectangle
                    int offset = 8 * i;
                    double[] xValues = new double[4] { quads[offset + 0], quads[offset + 2], quads[offset + 4], quads[offset + 6] };
                    double[] yValues = new double[4] { quads[offset + 1], quads[offset + 3], quads[offset + 5], quads[offset + 7] };
                    double x1 = xValues.Min();
                    double x2 = xValues.Max();
                    double y1 = yValues.Min();
                    double y2 = yValues.Max();
                    rects.Add(new Rect(x1, y1, x2, y2));
                }
                hightLights.Next();
            }
            rects.Sort(PositionComparer.CompareRectPos);
            return rects;
        }

        /// <summary>
        /// Get the axis value of a rectangle's bound on horizontal direction.
        /// </summary>
        /// <param name="isUpBound">Indicate whether the bound is upper boundary</param>
        /// <param name="roatatedAngle">The angle the pdf page has been rotated</param>
        /// <param name="lineRect">A rectangle</param>
        /// <returns>The axis value of a rectangle's bound on horizontal direction</returns>
        double GetAxisBoundValue(bool isUpBound, Rect lineRect)
        {
            return isUpBound ? lineRect.y2 : lineRect.y1;
        }

        string ReadTextFromRect(int pageNum, Rect bbox)
        {
            PageTextExtractor pageProcess = null;

            if (!_pageProcessCache.TryGetValue(pageNum, out pageProcess))
            {
                pageProcess = new PageTextExtractor(_pdfDoc.GetPage(pageNum));

                _pageProcessCache.Add(pageNum, pageProcess);
            }

            return pageProcess.SearchText(bbox);
        }

        int GetHighPage(LinePos highPos)
        {
            return highPos == null ? _pdfDoc.GetPageCount() : highPos.PageNum;
        }

        /// <summary>
        /// Deal the find results when some record is found out.
        /// </summary>
        /// <param name="hlts">The hightlights information</param>
        /// <param name="_pdfDoc">The pdf document object</param>
        /// <param name="lowRange">The beginning location for finding</param>
        /// <param name="verifyRegex">The regex for veritying key words</param>
        /// <returns>If the result is defined,return a LinePos object,otherwise return null</returns>
        LinePos GetLinePos(string regex, Highlights hlts, LinePos lowRange)
        {
            List<Rect> lineRects = GetLineRect(hlts);

            int pageNum = hlts.GetCurrentPageNumber();

            foreach (Rect lineRect in lineRects)
            {
                if (pageNum <= GetLowPage(lowRange) && !IsBelow(lineRect, lowRange))
                {
                    continue;
                }

                string text = ReadTextFromRect(pageNum, lineRect);

                string trimText = text.RemoveBlankSpace();

                if (!Regex.IsMatch(trimText, regex))
                {
                    continue;
                }

                LinePos linePos = new LinePos()
                {
                    PageNum = pageNum,
                    AxisValue = GetAxisBoundValue(true, lineRect),
                    TrimText = trimText,
                    AxisValueWithLineHeight = GetAxisBoundValue(false, lineRect),
                };
                return linePos;
            }
            return null;
        }

        /// <summary>
        /// Initilize a TextSearch object and begin search.
        /// </summary>
        /// <param name="regex">The regex for searching</param>
        /// <param name="startPage">The number of start page for searching</param>
        /// <param name="endPage">The number of end page for searching</param>
        /// <returns>A TestSearch object.</returns>
        TextSearch InitAndBeginSearch(string regex, int startPage, int endPage, bool isToUpSearch)
        {
            TextSearch searcher = new TextSearch();
            int searchMode = (int)(pdftron.PDF.TextSearch.SearchMode.e_reg_expression |
                                                pdftron.PDF.TextSearch.SearchMode.e_highlight);
            if (isToUpSearch)
            {
                searchMode = searchMode | (int)pdftron.PDF.TextSearch.SearchMode.e_search_up;
            }
            searcher.Begin(_pdfDoc, regex, searchMode, startPage, endPage);
            return searcher;
        }

        bool IsBelow(Point point, double axisValue)
        {
            if (axisValue == -1)
                return true;

            return point.y <= axisValue;
        }

        /// <summary>
        /// Get the size of a pdf page,
        /// the first number in the number array is width,
        /// and the second is height.
        /// </summary>
        /// <param name="pageNum">Pdf page number</param>
        /// <returns>The size of a pdf page</returns>
        double[] GetPageSize(int pageNum)
        {
            Page page = _pdfDoc.GetPage(pageNum);
            return GetPageSize(page);
        }

        /// <summary>
        /// Judge whether the position of lineRect is below lowRange.
        /// </summary>
        /// <param name="rotateAngle">The angle current pdf page has been rotated</param>
        /// <param name="lineRect">A rectangle</param>
        /// <param name="lowRange">The position of level</param>
        /// <returns>Return true if the position of lineRect is below lowRange,otherwise return false</returns>
        bool IsBelow(Rect lineRect, LinePos lowRange)
        {
            if (lowRange == null)
                return true;

            return lineRect.y1 <= lowRange.AxisValue;
        }

        string GetPdfPageContent(PDFDoc pdfDoc, int pageNumber)
        {
            List<PdfString> matchFuncLines = new List<PdfString>();

            Page page = pdfDoc.GetPage(pageNumber);
            if (page == null) return null;

            TextExtractor txt = new TextExtractor();
            txt.Begin(page);

            TextExtractor.Line line;
            TextExtractor.Word word;

            string lineString = null;
            StringBuilder stringBuilder = new StringBuilder();

            for (line = txt.GetFirstLine(); line.IsValid(); line = line.GetNextLine())
            {
                if (line.GetNumWords() == 0)
                {
                    continue;
                }
                lineString = null;
                for (word = line.GetFirstWord(); word.IsValid(); word = word.GetNextWord())
                {
                    int sz = word.GetStringLen();
                    if (sz == 0) continue;

                    lineString += word.GetString();
                }
                if (string.IsNullOrEmpty(lineString)) continue;
                stringBuilder.Append(lineString);
            }
            txt.Dispose();
            return stringBuilder.ToString();
        }

        Dictionary<int, PageTextExtractor> _pageProcessCache = new Dictionary<int, PageTextExtractor>();

        void ApplyMatrixToRect(Matrix2D matrix, Rect bounds)
        {
            double x1, x2, y1, y2;
            x1 = bounds.x1;
            x2 = bounds.x2;
            y1 = bounds.y1;
            y2 = bounds.y2;
            matrix.Mult(ref x1, ref y1);
            matrix.Mult(ref x2, ref y2);
            bounds.x1 = Math.Min(x1, x2);
            bounds.x2 = Math.Max(x1, x2);
            bounds.y1 = Math.Min(y1, y2);
            bounds.y2 = Math.Max(y1, y2);
        }

        double GetBlankAreaHeight(LinePos uponPos)
        {
            int pageNum = uponPos.PageNum;
            double bottomAxisValue = 0;
            double uponValue = GetBottomTextYValue(pageNum, bottomAxisValue,uponPos.AxisValue,false);
            return Math.Abs(uponValue - bottomAxisValue);
        }

        double GetBottomTextYValue(int pageNum, double bottomAxisValue, double topAxisValue, bool isMatchNum)
        {
            Page curPage = _pdfDoc.GetPage(pageNum);
            List<TextAndAxisValue> values = new List<TextAndAxisValue>();
            using (ElementReader pageReader = new ElementReader())
            {
                pageReader.Begin(curPage);
                GetTextElement(pageNum, pageReader, values, bottomAxisValue, topAxisValue, isMatchNum);
            }
            Merge(values);
            Filtrate(values);
            double bottomValue = bottomAxisValue;
            if (values.Count > 0)
            {
                double value = values.Min(x => x.yValue) - 15;
                if (value > 0)
                {
                    bottomValue = value;
                }
            }
            return bottomValue;
        }

        void Merge(List<TextAndAxisValue> TextAndAxisValues)
        {
            for (int i = 0; i < TextAndAxisValues.Count; i++)
            {
                TextAndAxisValue current = TextAndAxisValues[i];
                for (int j = i + 1; j < TextAndAxisValues.Count; )
                {
                    TextAndAxisValue next = TextAndAxisValues[j];
                    if (Math.Abs(current.yValue - next.yValue) < 3)
                    {
                        if (current.xValue > next.xValue)
                        {
                            current.text = next.text + current.text;
                        }
                        else
                        {
                            current.text += next.text;
                        }
                        TextAndAxisValues.RemoveAt(j);
                    }
                    else
                    {
                        j++;
                    }
                }
            }
        }

        void Filtrate(List<TextAndAxisValue> TextAndAxisValues)
        {
            string regex = @"^(\d+|第\d+页(共\d+页)?|-\d+-)$";
            for (int i = 0; i < TextAndAxisValues.Count; )
            {
                TextAndAxisValue current = TextAndAxisValues[i];
                if (Regex.IsMatch(current.text.RemoveBlankSpace(), regex))
                {
                    TextAndAxisValues.RemoveAt(i);
                }
                else
                {
                    i++;
                }
            }
        }

        void GetTextElement(int pageNum, ElementReader reader, List<TextAndAxisValue> values, double bottomAxisValue, double topAixsValue, bool isMatchNum)
        {
            Element element;
            while ((element = reader.Next()) != null)
            {
                Element.Type elementType = element.GetType();
                switch (elementType)
                {
                    case Element.Type.e_text:
                        ProcessText(pageNum, element, values, bottomAxisValue, topAixsValue, isMatchNum);
                        break;
                    case Element.Type.e_text_begin:
                        GetTextElement(pageNum, reader, values, bottomAxisValue, topAixsValue, isMatchNum);
                        break;
                    case Element.Type.e_text_end:
                        return;
                    case Element.Type.e_form:
                        reader.FormBegin();
                        GetTextElement(pageNum, reader, values, bottomAxisValue, topAixsValue, isMatchNum);
                        reader.End();
                        break;
                }
                reader.ClearChangeList();
            }
        }

        void ProcessText(int pageNum, Element element, List<TextAndAxisValue> values, double bottomAxisValue, double topAxisValue, bool isMatchNum)
        {
            Rect bbox = new Rect();
            element.GetBBox(bbox);

            string text = element.GetTextString();

            if (element.GetType() == Element.Type.e_text &&
                !string.IsNullOrEmpty(element.GetTextString().RemoveBlankSpace()))
            {
                Point leftBottomPoint = GetCooridinateValueOfLeftBottomCorner(pageNum,bbox);
                double horAxisValue = leftBottomPoint.y;
                double verAxisValue = leftBottomPoint.x;
                if (
                    IsBelow(leftBottomPoint, topAxisValue) &&
                    !IsBelow(leftBottomPoint,bottomAxisValue) &&
                    (!isMatchNum || Regex.IsMatch(text, @"\d")))
                {
                    TextAndAxisValue textAndAxisValue = new TextAndAxisValue();
                    textAndAxisValue.text = text;
                    textAndAxisValue.yValue = horAxisValue;
                    textAndAxisValue.xValue = verAxisValue;
                    values.Add(textAndAxisValue);
                }
            }
        }

        Point GetCooridinateValueOfLeftBottomCorner(int pageNum, Rect bbox)
        {
            Rect _bbox = bbox.Copy();
            Matrix2D matrix = _pdfDoc.GetPage(pageNum).GetDefaultMatrix();
            ApplyMatrixToRect(matrix, _bbox);
            return new Point(_bbox.x1, _bbox.y1);
        }

        #endregion

    }

    class TextAndAxisValue
    {
        public string text;
        public double xValue;
        public double yValue;
    }
}
