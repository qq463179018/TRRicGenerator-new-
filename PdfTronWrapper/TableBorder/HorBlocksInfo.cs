//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:To record the horizontial infomation of a series of data blocks.
//-----
//-----------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdfTronWrapper.TableBorder
{
    public class HorBlocksInfo : IComparable
    {
        public double TopBound { get; set; }
        public double BottomBound { get; set; }

        public double CenterYValue
        {
            get
            {
                return (TopBound + BottomBound) / 2;
            }
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            HorBlocksInfo horBlocksInfo = obj as HorBlocksInfo;
            if (horBlocksInfo.BottomBound == BottomBound)
                return 0;
            return BottomBound > horBlocksInfo.BottomBound ? 1 : -1;
        }

        public bool IsUpon(double yValue)
        {
            return (TopBound + BottomBound) / 2 > yValue;
        }

        public bool IsBelow(double yValue)
        {
            return (TopBound + BottomBound) / 2 < yValue;
        }



        #endregion
    }
}
