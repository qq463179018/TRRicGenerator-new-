using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pdftron.PDF;

namespace Ric.Util
{
    public enum PositionRect { X1, Y1, X2, Y2 }

    public class PdfAnalyzer
    {
        public List<PdfString> RegexSearchAllPages(PDFDoc doc, string pattern)
        {
            return RegexSearch(doc, pattern, false, -1, -1, true);
        }

        public List<PdfString> RegexSearchByPageRange(PDFDoc doc, string pattern, int startPage, int endPage)
        {
            if (endPage > doc.GetPageCount())
                throw new Exception("endPage out of MaxRange of pdf.");

            if (startPage < 1)
                throw new Exception("startPage out of MixRange of pdf.");

            if (startPage > endPage)
                throw new Exception("pageRange is invalid.");

            return RegexSearch(doc, pattern, false, startPage, endPage, true);
        }

        public List<PdfString> RegexSearchByPage(PDFDoc doc, string pattern, int pageIndex)
        {
            if (pageIndex > doc.GetPageCount())
                throw new Exception("pageIndex out of MaxRange of pdf.");

            if (pageIndex < 1)
                throw new Exception("pageIndex out of MixRange of pdf.");

            return RegexSearch(doc, pattern, false, pageIndex, pageIndex, true);
        }

        public List<PdfString> RegexSearch(PDFDoc doc, string pattern, bool ifWholeWord, int startPage, int endPage, bool ignoreCase)
        {
            List<PdfString> result = new List<PdfString>();
            Int32 page_num = 0;
            string result_str = "";
            string ambient_string = "";
            Highlights hlts = new Highlights();

            Int32 mode = (Int32)(TextSearch.SearchMode.e_reg_expression | TextSearch.SearchMode.e_highlight);
            if (ifWholeWord) mode |= (Int32)TextSearch.SearchMode.e_whole_word;
            if (ignoreCase) mode |= (Int32)TextSearch.SearchMode.e_case_sensitive;

            int pageCount = doc.GetPageCount();
            if (endPage > pageCount) endPage = pageCount;

            TextSearch txt_search = new TextSearch();
            txt_search.Begin(doc, pattern, mode, startPage, endPage);

            while (true)
            {
                TextSearch.ResultCode code = txt_search.Run(ref page_num, ref result_str, ref ambient_string, hlts);

                if (code == TextSearch.ResultCode.e_found)
                {
                    hlts.Begin(doc);
                    double[] box = null;
                    string temp = result_str;

                    while (hlts.HasNext())
                    {
                        box = hlts.GetCurrentQuads();
                        if (box.Length != 8)
                        {
                            hlts.Next();
                            continue;
                        }

                        result.Add(new PdfString(result_str, new Rect(box[0], box[1], box[4], box[5]), page_num));
                        hlts.Next();
                    }
                }
                else if (code == TextSearch.ResultCode.e_done)
                {
                    break;
                }
            }
            return result;
        }

        public List<PdfString> RegexExtractByPositionWithPage(PDFDoc doc, string pattern, int pageIndex, Rect rect, PositionRect positionRect = PositionRect.Y2, double range = 2.0)
        {
            return GetNearbyPdfString(RegexSearchByPage(doc, pattern, pageIndex), rect, positionRect, range);
        }

        public List<PdfString> GetNearbyPdfString(List<PdfString> pdfStringAll, Rect rect, PositionRect positionRect, double range)
        {
            List<PdfString> pdfStringFilter = new List<PdfString>();

            foreach (var pdf in pdfStringAll)
            {
                if (pdf == null)
                    continue;

                if (GetRange(pdf.Position, rect, positionRect) > range)
                    continue;

                pdfStringFilter.Add(pdf);
            }

            PdfString.PdfStringComparer comparerType = PdfString.GetComparer();

            if (positionRect.Equals(PositionRect.Y2) || positionRect.Equals(PositionRect.Y1))
                comparerType.WhichComparison = PdfString.PdfStringComparer.ComparisonType.Horizontal;
            else if (positionRect.Equals(PositionRect.X1) || positionRect.Equals(PositionRect.X2))
                comparerType.WhichComparison = PdfString.PdfStringComparer.ComparisonType.Vertical;

            pdfStringFilter.Sort(comparerType);
            return pdfStringFilter;
        }

        private double GetRange(Rect pdf, Rect title, PositionRect positionRect)
        {
            switch (positionRect)
            {
                case PositionRect.X1:
                    return Math.Abs(pdf.x1 - title.x1);
                case PositionRect.X2:
                    return Math.Abs(pdf.x2 - title.x2);
                case PositionRect.Y1:
                    return Math.Abs(pdf.y1 - title.y1);
                case PositionRect.Y2:
                    return Math.Abs(pdf.y2 - title.y2);
                default:
                    throw new Exception("calculation range of pdfstring with title error.");
            }
        }

        public List<PdfString> RegexExtractByPositionWithAllPage(PDFDoc doc, string pattern, Rect rect, PositionRect positionRect = PositionRect.Y2, double range = 2.0)
        {
            return GetNearbyPdfString(RegexSearchAllPages(doc, pattern), rect, positionRect, range);
        }

        public List<PdfString> RegexExtractByPositionWithRangePage(PDFDoc doc, string pattern, int startPage, int endPage, Rect rect, PositionRect positionRect = PositionRect.Y2, double range = 2.0)
        {
            return GetNearbyPdfString(RegexSearchByPageRange(doc, pattern, startPage, endPage), rect, positionRect, range);
        }
    }
}
