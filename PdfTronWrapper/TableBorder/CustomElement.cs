//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:To record the information of text block in the pdf page.
//-----
//-----------------------------------------------------------------------------------------------------------------------


using pdftron.PDF;

namespace PdfTronWrapper.TableBorder
{
    /// <summary>
    /// To record the information of text block in the pdf page.
    /// </summary>
    internal class CustomElement
    {
        /// <summary>
        /// The text of the element.
        /// </summary>
        public string TextData { get; set; }

        /// <summary>
        /// The rectangle region of the element.
        /// </summary>
        public Rect BBoxRect { get; set; }

        /// <summary>
        /// Indicate whether one element overlap the other.
        /// </summary>
        /// <param name="otherElement">The other element</param>
        /// <returns>If the elements overlap each other,return true;Otherwise,return false.</returns>
        public bool IsOverlap(CustomElement otherElement)
        {
            return TextData.Equals(otherElement.TextData) && BBoxRect.IsSame(otherElement.BBoxRect, 1);
        }

    }
}
