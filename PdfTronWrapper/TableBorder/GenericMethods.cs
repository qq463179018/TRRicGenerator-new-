//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:A series of generic methods.
//-----
//-----------------------------------------------------------------------------------------------------------------------

using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace PdfTronWrapper.TableBorder
{
    internal class GenericMethods<T, K> where K : IList
    {
        public static void RemoveZeroAmountValueItems(SortedDictionary<T, K> dictionary)
        {
            dictionary.Where(pair => pair.Value.Count == 0)
                .Select(pair => pair.Key).ToList()
                .ForEach(key => dictionary.Remove(key));
        }
    }
}
