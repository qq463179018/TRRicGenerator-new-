//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:A data element part represent the objects after data element divide by backspace.
//-----
//-----------------------------------------------------------------------------------------------------------------------


using System;
using System.Collections.Generic;
using System.Text;
using pdftron.Common;
using pdftron.PDF;
using PdfTronWrapper.Utility;

namespace PdfTronWrapper.TableBorder
{
    internal class DataElementPart
    {
        public DataElementPart()
        {
            Text = "";
        }

        public double LeftBound { get; set; }
        public double Width { get; set; }
        public double RightBound
        {
            get
            {
                return LeftBound + Width;
            }
        }
        public string Text { get; set; }

        /// <summary>
        /// Get the parts of the element ergodiced from pdf page.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="pageDefaultMatrix"></param>
        /// <returns></returns>
        public static List<DataElementPart> GetElementParts(Element element, Matrix2D pageDefaultMatrix)
        { 
            List<DataElementPart> parts = new List<DataElementPart>();
            int index = -1;

            var gs = element.GetGState();
            var font = gs.GetFont();
            string text = element.GetTextString();
            var mtx = element.GetCTM() * element.GetTextMatrix();
            double font_size = Math.Abs(gs.GetFontSize() * Math.Sqrt(mtx.m_b * mtx.m_b + mtx.m_d * mtx.m_d));
            for (CharIterator itr = element.GetCharIterator(); itr.HasNext(); itr.Next())
            {
                index++;
                if (text[index].IsBlankSpace())
                {
                    continue;
                }
                double x = itr.Current().x;
                double y = itr.Current().y;
                int charCode = itr.Current().char_code;
                double charWidth = font.GetWidth(charCode);
                char chars = (char)charCode;
               
                //if (charCode >= 48 && charCode <= 57)//|| chars==',' || chars=='.')
                //    charWidth = 500;
                double w =charWidth * 0.001 * font_size;
                double h = font_size;
                mtx.Mult(ref x, ref y);
                pageDefaultMatrix.Mult(ref x, ref y);
                parts.Add(new DataElementPart
                {
                    LeftBound=x,Width=w,Text=text[index].ToString()
                });
            }
            Merge(parts);
            return parts;
        }

        static void Merge(List<DataElementPart> parts)
        {
            for (int i = 0; i < parts.Count - 1; )
            {
                DataElementPart current = parts[i];
                DataElementPart next = parts[i + 1];
                if (next.LeftBound - current.RightBound > current.Width/current.Text.Length * 0.8)
                {
                    i++;
                }
                else
                {
                    current.Width = next.RightBound - current.LeftBound;
                    current.Text += next.Text;
                    parts.RemoveAt(i + 1);
                }
            }
        }

    }
}
