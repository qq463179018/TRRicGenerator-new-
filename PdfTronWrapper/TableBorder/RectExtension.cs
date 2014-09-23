//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:Extend methods for class Rect.
//-----
//-----------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pdftron.PDF;
using pdftron.Common;

namespace PdfTronWrapper.TableBorder
{
    internal static class RectExtension
    {
        public static Rect Copy(this Rect rect)
        {
            return new Rect
            {
                x1 = rect.x1,
                x2 = rect.x2,
                y1 = rect.y1,
                y2 = rect.y2,
            };
        }

        public static void ApplyMatrix(this Rect rect, Matrix2D matrix)
        {
            double _x1, _x2, _y1, _y2;
            _x1 = rect.x1;
            _x2 = rect.x2;
            _y1 = rect.y1;
            _y2 = rect.y2;
            matrix.Mult(ref _x1, ref _y1);
            matrix.Mult(ref _x2, ref _y2);
            rect.x1 = Math.Min(_x1, _x2);
            rect.x2 = Math.Max(_x1, _x2);
            rect.y1 = Math.Min(_y1, _y2);
            rect.y2 = Math.Max(_y1, _y2);
        }

        public static void RevertMatrix(this Rect rect, Matrix2D matrix)
        {
            Matrix2D inverseMatrix = matrix.Inverse();
            rect.ApplyMatrix(inverseMatrix);
        }

        public static bool ContainFormLine(this Rect rect, FormLine formLine)
        {
            bool isContain = true;
            try
            {
                Point centerPoint=formLine.CenterPoint;
                if(!rect.Contains(centerPoint.x,centerPoint.y))
                    isContain = false;
            }
            catch (Exception ex)
            {
                string err = ex.ToString();
                throw ex;
            }
            return isContain;
        }

        public static bool IsSame(this Rect rect, Rect otherRect, double error)
        {
            bool isSame = true;
            if (!IsSameValue(rect.x1, otherRect.x1, error) ||
                    !IsSameValue(rect.x2, otherRect.x2, error) ||
                    !IsSameValue(rect.y1, otherRect.y1, error) ||
                    !IsSameValue(rect.y2, otherRect.y2, error))
            {
                isSame = false;
            }
            return isSame;
        }

        static bool IsSameValue(double value1, double value2, double error)
        {
            return Math.Abs(value1 - value2) < error;
        }

    }
}
