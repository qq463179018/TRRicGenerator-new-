//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:To extend the element's function.
//-----
//-----------------------------------------------------------------------------------------------------------------------


using System;
using pdftron.Common;
using pdftron.PDF;
using PdfTronWrapper.Utility;

namespace PdfTronWrapper.TableBorder
{
    internal static class ElementExtension
    {

        /// <summary>
        /// Get the horizontial bound of a element.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="pageDefaultMatrix"></param>
        /// <returns></returns>
        public static double[] GetHorBound(this Element element, Matrix2D pageDefaultMatrix)
        {
            double[] bound = new double[2];
            bool leftBreak = false;

            var gs = element.GetGState();
            var font = gs.GetFont();
            string text = element.GetTextString();
            var mtx = element.GetCTM() * element.GetTextMatrix();
            double font_size = Math.Abs(gs.GetFontSize() * Math.Sqrt(mtx.m_b * mtx.m_b + mtx.m_d * mtx.m_d));
            int charIndex = 0;
            for (CharIterator itr = element.GetCharIterator(); itr.HasNext(); itr.Next())
            {
                double x = itr.Current().x;
                double y = itr.Current().y;
                double w = font.GetWidth(itr.Current().char_code) * 0.001 * font_size;
                int charCode = itr.Current().char_code;
                char ch =text[charIndex];
                mtx.Mult(ref x, ref y);
                pageDefaultMatrix.Mult(ref x, ref y);
                if (!leftBreak &&  !ch.IsBlankSpace())
                {
                    leftBreak = true;
                    bound[0] = x;
                }
                if (!ch.IsBlankSpace())
                {
                    bound[1] = x + w - bound[0];
                }
                charIndex++;
            }
            return bound;
        }

        public static CustomElement GenerateCustomElement(this Element element, Matrix2D matrix)
        {
            return new CustomElement
            {
                TextData = element.GetTextString(),
                BBoxRect = element.GetBBoxAfterMatrixTranslate(matrix)
            };
        }

        /// <summary>
        /// Get bound scale of the element after applying the matrix translation.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="matrix"></param>
        /// <returns></returns>
        public static Rect GetBBoxAfterMatrixTranslate(this Element element, Matrix2D matrix)
        {
            Rect elementBound = new Rect();
            element.GetBBox(elementBound);
            elementBound.ApplyMatrix(matrix);
            double[] horBounds = element.GetHorBound(matrix);
            elementBound.x1 = horBounds[0];
            elementBound.x2 = horBounds[0] + horBounds[1];
            return elementBound;
        }
    }
}
