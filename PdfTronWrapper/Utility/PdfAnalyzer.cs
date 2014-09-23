using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pdftron.PDF;

namespace PdfTronWrapper.Utility
{
    public class PdfAnalyzer
    {
        private int startPage = -1;
        private int endPage = -1;
        private int minPage = 1;

        public List<PdfString> RegexSearchAllPages(PDFDoc doc, string pattern)
        {
            return RegexSearch(doc, pattern, false, startPage, endPage, true);
        }

        public List<PdfString> RegexSearchByPageRange(PDFDoc doc, string pattern, int startPage, int endPage)
        {
            if (endPage > doc.GetPageCount())
                throw new Exception("endPage out of MaxRange of pdf.");

            if (startPage < minPage)
                throw new Exception("startPage out of MixRange of pdf.");

            if (startPage < endPage)
                throw new Exception("pageRange is invalid.");

            return RegexSearch(doc, pattern, false, startPage, endPage, true);
        }

        public List<PdfString> RegexSearchByPage(PDFDoc doc, string pattern, int pageIndex)
        {
            if (pageIndex > doc.GetPageCount())
                throw new Exception("pageIndex out of MaxRange of pdf.");

            if (pageIndex < minPage)
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
    }
}
