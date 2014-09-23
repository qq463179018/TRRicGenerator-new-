//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:Customized keyvaluepair class with the comparing function.
//-----
//-----------------------------------------------------------------------------------------------------------------------


using System;

namespace PdfTronWrapper.TableBorder
{
    /// <summary>
    /// Customized keyvaluepair class with the comparing function.
    /// </summary>
    /// <typeparam name="K">Data type of the key.</typeparam>
    /// <typeparam name="V">Data type of the value.</typeparam>
    internal class CustomPair<K, V> : IComparable where K : IComparable
    {
        public K Key;
        public V Value;

        #region IComparable Members

        public int CompareTo(object obj)
        {
            CustomPair<K, V> pair = obj as CustomPair<K, V>;
            return Key.CompareTo(pair.Key);
        }

        #endregion
    }
}
