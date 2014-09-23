using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pdftron.Common;
using pdftron.PDF;

namespace PdfTronWrapper.Utility
{
    class PageTextExtractor
    {
        double _pageHeight;

        Matrix2D _pageMatrix;

        int _charCode, _charUnicode;

        PdfLineOfText[] _textLines;

        double right = -1, left = -1;

        public PageTextExtractor(Page page)
        {
            PdfPage = page;
            ExtractText();
        }

        #region class properties

        public Page PdfPage { get; private set; }

        public double[] LeftRightBounds
        {
            get
            {
                if (left == -1)
                {
                    left = 0;
                    List<PdfLineOfText> lines = _textLines.Where(line => line != null).ToList();
                    if (lines.Count() > 0)
                    {
                        List<double> leftValues = lines.Where(line => line.Left > -1).Select(line => line.Left).ToList();
                        if (leftValues.Count > 0)
                            left = leftValues.Min();
                    }

                    right = 0;
                    if (lines.Count() > 0)
                    {
                        List<double> RightValues = lines.Where(line => line.Right > -1).Select(line => line.Right).ToList();
                        if (RightValues.Count > 0)
                            right = RightValues.Max();
                    }
                }
                return new double[2] { left, right };
            }
        }

        #endregion

        /// <summary>
        /// search text in PDF page,a less strict search method.
        /// text that intersects with the area will also be selected.
        /// </summary>
        /// <param name="area"></param>
        /// <returns></returns>
        public string SearchText(Rect area)
        {
            area.Normalize();

            if (area.y2 < 0 || area.y1 > _pageHeight) return string.Empty;

            if (area.y1 < 0) area.y1 = 0;

            if (area.y2 > _pageHeight) area.y2 = _pageHeight;

            var start = (int)(_pageHeight - area.y2) >> 2;
            var end = ((int)(_pageHeight - area.y1) >> 2) + 2;

            StringBuilder sb = new StringBuilder();
            var rect = new System.Windows.Rect(area.x1, area.y1, area.Width(), area.Height());

            for (int i = start; i < end; i++)
            {
                var line = _textLines[i];

                if (line == null) continue;

                var first = line.FirstChar;

                if (first.Top + 2 > area.y2) continue;
                if (first.Bottom < area.y1 + 1) break;

                line.SearchIntersectsText(rect, sb);
            }

            return sb.ToString();
        }

        /// <summary>
        /// search text int PDF page,a more strict search method.
        /// text that only intersect with the area will be discarded.
        /// </summary>
        public string SearchTextWithStrictMode(System.Windows.Rect area)
        {
            if (area.Bottom < 0 || area.Top > _pageHeight) return string.Empty;

            if (area.Top < 0) area.Y = 0;

            if (area.Bottom > _pageHeight) area.Height -= (area.Bottom - _pageHeight);

            var start = (int)(_pageHeight - area.Bottom) >> 2;
            var end = ((int)(_pageHeight - area.Top) >> 2) + 2;

            StringBuilder sb = new StringBuilder();

            for (int i = start; i < end; i++)
            {
                var line = _textLines[i];

                if (line == null) continue;

                var first = line.FirstChar;

                if (first.Bottom < area.Top + 1) break;
                if (area.Bottom < first.Top + 4) continue;

                line.SearchTextExact(area.Left, area.Right, sb);
            }

            return sb.ToString();
        }

        #region text extract functions

        /// <summary>
        /// extract text in the pdf page
        /// </summary>
        void ExtractText()
        {
            _charCode = -1;

            _pageMatrix = PdfPage.GetDefaultMatrix();

            _pageHeight = PdfPage.GetPageHeight();

            _textLines = new PdfLineOfText[((int)_pageHeight >> 2) + 8];

            using (ElementReader page_reader = new ElementReader())
            {
                page_reader.Begin(PdfPage);
                ProcessElements(page_reader);
            }

            PdfLineOfText preLt = null;

            for (int i = 0; i < _textLines.Length; i++)
            {
                var lt = _textLines[i];

                if (preLt != null && lt != null && preLt.FirstChar.Top - lt.FirstChar.Top < 3)
                {
                    preLt.AddPdfChars(lt.Chars);

                    _textLines[i] = lt = null;
                }

                preLt = lt;
            }
        }

        /// <summary>
        /// process all elements on a PDF page
        /// </summary>
        void ProcessElements(ElementReader reader)
        {
            Element element;

            while ((element = reader.Next()) != null)
            {
                switch (element.GetType())
                {
                    case Element.Type.e_text:           // Process text strings...
                        ProcessTextElement(element);
                        break;
                    case Element.Type.e_form:           // Process form XObjects
                        reader.FormBegin();
                        ProcessElements(reader);
                        reader.End();
                        break;
                }
            }
        }

        // prcoess text element
        void ProcessTextElement(Element element)
        {
            double x, y;

            var text = element.GetTextString();

            if (text.Trim().Length == 0) return;

            var matrix = element.GetCTM();

            matrix.Concat(_pageMatrix.m_a, _pageMatrix.m_b, _pageMatrix.m_c, _pageMatrix.m_d, _pageMatrix.m_h, _pageMatrix.m_v);

            matrix *= element.GetTextMatrix();

            var gs = element.GetGState();

            var font = gs.GetFont();

            double font_size = Math.Abs(gs.GetFontSize() * Math.Sqrt(matrix.m_b * matrix.m_b + matrix.m_d * matrix.m_d));// font_size * font_sz_scale_factor;

            // remove watermark
            if (font_size > 32 && gs.GetTextRenderMode() == GState.TextRenderingMode.e_stroke_text) return;

            int index = -1;

            var chs = new PdfChar[text.Length];

            for (CharIterator itr = element.GetCharIterator(); itr.HasNext(); itr.Next())
            {
                index++;
                x = itr.Current().x;
                y = itr.Current().y;
                matrix.Mult(ref x, ref y);

                var ch = text[index];

                if (ch == 65533)
                {
                    //in some pdf files,we can't find the right unicode for some characters by using pdftron,
                    //so we record a pdf charcode with a right unicode value,and then computing the unicode for those bad characters.
                    if (_charCode > 0) ch = (char)(_charUnicode + itr.Current().char_code - _charCode);
                }
                else if (ch >= 0x4e00)
                {
                    _charUnicode = ch;
                    _charCode = itr.Current().char_code;
                }

                chs[index] = new PdfChar(ch, new System.Windows.Rect(x, y, font.GetWidth(itr.Current().char_code) * 0.001 * font_size, font_size));
            }

            var pchar = chs[0];

            var dis = _pageHeight - pchar.Top;

            index = (int)dis >> 2;

            var textLine = _textLines[index];

            if (textLine == null) _textLines[index] = textLine = new PdfLineOfText();

            textLine.AddPdfChars(chs);
        }

        #endregion

        #region PdfChar&PdfLineOfText

        /// <summary>
        /// character info in pdf page
        /// </summary>
        struct PdfChar
        {
            public char _char;

            public System.Windows.Rect charBox;

            public PdfChar(char ch, System.Windows.Rect bbox)
            {
                this._char = ch;
                this.charBox = bbox;
            }

            public double Left
            {
                get { return charBox.Left; }
            }

            public double Right
            {
                get { return charBox.Right; }
            }

            public double Top
            {
                get { return charBox.Top; }
            }

            public double Bottom
            {
                get { return charBox.Bottom; }
            }

            public bool IsSpace
            {
                get { return _char == ' ' || _char == '　'; }
            }

            public bool IntersectsWith(System.Windows.Rect rect)
            {
                return charBox.IntersectsWith(rect);
            }

            public readonly static PdfChar Empty = new PdfChar(char.MinValue, System.Windows.Rect.Empty);
        }

        /// <summary>
        /// text line info
        /// </summary>
        class PdfLineOfText
        {
            bool _sorted = false;

            List<PdfChar> _list = new List<PdfChar>();

            public double Left
            {
                get
                {
                    SortAndCleanup();
                    return _list.First(ch => !ch.IsSpace).Left;

                }
            }

            public double Right
            {
                get
                {
                    SortAndCleanup();
                    return _list.Last(ch => !ch.IsSpace).Right;
                }
            }

            public PdfChar FirstChar
            {
                get { return _list[0]; }
            }

            public IEnumerable<PdfChar> Chars
            {
                get { return _list; }
            }

            /// <summary>
            /// sort chars and clean overlapping chars
            /// </summary>
            public void SortAndCleanup()
            {
                if (_sorted) return;

                _sorted = true;

                _list.Sort((n, m) => (int)(n.Left - m.Left));

                var preChar = PdfChar.Empty;

                bool exists = false;

                for (int i = 0; i < _list.Count; i++)
                {
                    var curChar = _list[i];

                    if (curChar._char == preChar._char && curChar.Left - preChar.Left < 2)
                    {
                        exists = true;
                        _list[i] = PdfChar.Empty;
                    }

                    preChar = curChar;
                }

                if (exists) _list.RemoveAll(n => n._char == char.MinValue);
            }

            /// <summary>
            /// save characters
            /// </summary>
            public void AddPdfChars(IEnumerable<PdfChar> chars)
            {
                _list.AddRange(chars);
            }

            //
            public void SearchTextExact(double x1, double x2, StringBuilder sb)
            {
                if (!_sorted) this.SortAndCleanup();

                var start = GetSearchStart(x1);

                if (start < 0) return;

                for (int i = start; i < _list.Count; i++)
                {
                    var curCh = _list[i];

                    if (curCh.Right < x2 || x2 - curCh.Left > 2)
                    {
                        sb.Append(curCh._char);
                    }
                    else break;
                }
            }

            public void SearchTextExact(double x1, double x2, StringBuilder sb, ref System.Windows.Rect bbox)
            {
                if (!_sorted) this.SortAndCleanup();

                var start = GetSearchStart(x1);

                if (start < 0) return;

                var i = start;

                while (i < _list.Count && char.IsWhiteSpace(_list[i]._char)) i++;

                for (; i < _list.Count; i++)
                {
                    var curCh = _list[i];

                    if (curCh.Right < x2 || x2 - curCh.Left > 2)
                    {
                        sb.Append(curCh._char);
                        bbox.Union(curCh.charBox);
                    }
                    else break;
                }
            }

            /// <summary>
            /// search text intersecting with bbox
            /// </summary>
            public string SearchIntersectsText(System.Windows.Rect bbox, StringBuilder sb)
            {
                if (!_sorted) this.SortAndCleanup();

                var start = GetSearchStart(bbox.X);

                if (start < 0) return string.Empty;

                for (int i = start; i < _list.Count; i++)
                {
                    if (_list[i].IntersectsWith(bbox)) sb.Append(_list[i]._char);
                }

                return sb.ToString();
            }

            // get start index about x value
            int GetSearchStart(double x)
            {
                double diff;

                int low = 0, hi = _list.Count - 1, half, start = -1;

                while (low <= hi)
                {
                    half = (low + hi) >> 1;

                    diff = _list[half].Left + 2.11 - x;

                    if (diff > 0)
                    {
                        start = half;
                        hi = half - 1;
                    }
                    else
                    {
                        low = half + 1;
                    }
                }

                return start;
            }

            /// <summary>
            /// get continous text blocks in a region
            /// </summary>
            public void GetTextBlocks(List<PdfTextBlock> results, double x1, double x2, double y1, int rowId)
            {
                if (!_sorted) this.SortAndCleanup();

                var start = GetSearchStart(x1);

                if (start < 0 || _list[start].Bottom < y1 + 1) return;

                double left = 0, right = 0, charSpacing = 0;

                var sb = new StringBuilder();

                for (int index = start; index < _list.Count; index++)
                {
                    var pch = _list[index];

                    if (pch.Left + 2 > x2) break;

                    switch (sb.Length)
                    {
                        case 0:
                            if (!pch.IsSpace)
                            {
                                left = pch.Left;
                                charSpacing = -101;
                                sb.Append(pch._char);
                            }
                            continue;
                        case 1:
                            if (charSpacing < -100)
                            {
                                charSpacing = pch.Left - right;
                            }
                            break;
                        default:
                            if (Math.Abs(pch.Left - right - charSpacing) > 1 && pch.Left - right > 10)
                            {
                                results.Add(new PdfTextBlock(sb.ToString(), new System.Windows.Rect(left, 0, right - left, 0)) { RowID = rowId });

                                sb.Length = 0;

                                left = pch.Left;
                            }
                            break;
                    }

                    right = pch.Right;

                    if (!pch.IsSpace) sb.Append(pch._char);
                }

                if (sb.Length > 0)
                {
                    results.Add(new PdfTextBlock(sb.ToString(), new System.Windows.Rect(left, 0, right - left, 0)) { RowID = rowId });
                }
            }

            public override string ToString()
            {
                if (!_sorted) this.SortAndCleanup();

                StringBuilder sb = new StringBuilder();

                foreach (var pchar in _list) sb.Append(pchar._char);

                return sb.ToString();
            }
        }

        class PdfTextBlock
        {
            string _text;

            public PdfTextBlock(string text)
                : this(text, System.Windows.Rect.Empty)
            {

            }

            public PdfTextBlock(string text, System.Windows.Rect bbox)
            {
                _text = text;
                BBox = bbox;
            }

            public string Text
            {
                get { return _text ?? string.Empty; }
                set { _text = value; }
            }

            public int RowID { get; set; }

            public double Left
            {
                get { return BBox.Left; }
            }

            public double Right
            {
                get { return BBox.Right; }
            }

            public System.Windows.Rect BBox { get; set; }

            public override string ToString()
            {
                return Text;
            }
        }

        #endregion
    }
}