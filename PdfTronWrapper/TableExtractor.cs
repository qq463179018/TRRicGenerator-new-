using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PdfTronWrapper.Utility;

namespace PdfTronWrapper
{
    public class TableExtractor
    {
        #region extract PDF table

        /// <summary>
        /// extract pdf table data
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="tablePos"></param>
        /// <returns></returns>
        public static FreeTable Extract(pdftron.PDF.PDFDoc doc, TablePos tablePos)
        {
            PageTextExtractor pdfPageProcess = new PageTextExtractor(doc.GetPage(tablePos.PageNum));

            var vLines = new List<System.Windows.Rect>(tablePos.VerticalLines.Count);
            var hLines = new List<System.Windows.Rect>(tablePos.HorizontialLines.Count);

            foreach (var kv in tablePos.VerticalLines)
            {
                var top = kv.Value.Min(n => n.StartPoint.y);
                var bottom = kv.Value.Max(n => n.EndPoint.y);
                vLines.Add(new System.Windows.Rect(kv.Key, top, 0, bottom - top));
            }

            foreach (var kv in tablePos.HorizontialLines)
            {
                var left = kv.Value.Min(n => n.StartPoint.x);
                var right = kv.Value.Max(n => n.EndPoint.x);
                hLines.Add(new System.Windows.Rect(left, kv.Key, right - left, 0));
            }

            if (vLines.Count == 0 || hLines.Count == 0)
            {
                throw new Exception("vertical or horizontal lines is empty");
            }

            hLines.Reverse();

            FreeTable freeTable = new FreeTable();

            for (int i = 1; i < hLines.Count; i++)
            {
                var y = hLines[i].Y;

                var y2 = hLines[i - 1].Y;

                var prej = 0;

                var x = hLines[i - 1].Left;

                while (prej < vLines.Count)
                {
                    if (vLines[prej].X > x - 5) break;

                    prej++;
                }

                var x1 = vLines[prej].X;

                FreeTableRow row = new FreeTableRow();

                for (int j = prej + 1; j < vLines.Count; j++)
                {
                    var vline = vLines[j];

                    if (vline.Bottom - y > 5 || j == vLines.Count - 1)
                    {
                        int ii = i;

                        for (; ii < hLines.Count - 1; ii++)
                        {
                            if (vline.X - hLines[ii].Left > 5) break;
                        }

                        var data = pdfPageProcess.SearchTextWithStrictMode(
                            new System.Windows.Rect(x1, hLines[ii].Y, vline.X - x1, y2 - hLines[ii].Y));

                        row.Add(new FreeTableCell() { Value = data, ColSpan = j - prej, RowSpan = ii - i + 1 });

                        prej = j;

                        x1 = vline.X;
                    }
                }

                freeTable.Add(row);
            }

            return freeTable;
        }

        #endregion
    }
}
