//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:Extend methods for class char.
//-----
//-----------------------------------------------------------------------------------------------------------------------


namespace PdfTronWrapper.Utility
{
    internal static class CharExtension
    {
        public static bool IsBlankSpace(this char ch)
        {
            return ch==' '||ch==' '||ch=='　';
        }
    }
}
