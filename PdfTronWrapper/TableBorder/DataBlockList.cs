//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:A series of data block.
//-----
//-----------------------------------------------------------------------------------------------------------------------


using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PdfTronWrapper.TableBorder
{
    public class DataBlockList : List<DataBlock>
    {
        string chineseLetterRegex = "[\u4e00-\u9fa5]+";

        public double TopBound
        {
            get
            {
                return Count > 0 ?
                    this.Select(block => block.TopBound).Max() :
                    -1;
            }
        }

        public double BottomBound
        {
            get
            {
                return Count > 0 ?
                    this.Select(block => block.BottomBound).Min() :
                    -1;
            }
        }

        public double Hight
        {
            get { return TopBound - BottomBound; }
        }

        public new bool Remove(DataBlock textBlock)
        {
            bool result = false;
            foreach (DataBlock _textBlock in this)
            {
                if (_textBlock.Equals(textBlock))
                {
                    result = base.Remove(_textBlock); 
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// Get the number block of the block list whose position is lowest.
        /// </summary>
        /// <param name="horBlocks"></param>
        /// <returns></returns>
        public DataBlock GetLastNumBlock(SortedDictionary<HorBlocksInfo, DataBlockList> horBlocks)
        {
            DataBlock lastNumBlock = null;

            Sort();
            foreach (DataBlock block in this)
            {
                if (block.DataType==DataType.Number)
                {
                    lastNumBlock = block;
                    break;
                }
            }
            return lastNumBlock;
        }

        public DataType DataType
        {
            get
            {
                bool isNumberColumn = false;
                int numberCount = 0;
                int chineseCount = 0;
                foreach (DataBlock block in this)
                {
                    if (block.DataType == DataType.Number)
                    {
                        numberCount++;
                    }
                    if (Regex.IsMatch(block.Text, chineseLetterRegex))
                    {
                        chineseCount++;
                    }
                }
                isNumberColumn = numberCount>=chineseCount;
                return isNumberColumn?DataType.Number:DataType.Text;
            }
        }

        public DataBlockList()
        {
        }

        public DataBlockList(List<DataBlock> blocks)
        {
            foreach (DataBlock block in blocks)
            {
                Add(block);
            }
        }

        public List<DataBlock> GetBlocksOfScale(double bottomBound, double topBound)
        {
            List < DataBlock > blocks = this.Where(block =>
            {
                double medialHorPos = (block.TopBound + block.BottomBound) / 2;
                return medialHorPos > bottomBound && medialHorPos < topBound;
            }).Distinct().ToList();
            return blocks;
        }

        /// <summary>
        /// Get two lines which is nearby the data blocks.
        /// </summary>
        /// <param name="horizontialLines"></param>
        /// <returns></returns>
        public FormLineList[] GetNearbyTwoLines(SortedDictionary<double, FormLineList> horizontialLines)
        {
            List<FormLineList> bottomLines = new List<FormLineList>();
            List<FormLineList>topLines = new List<FormLineList>();
            foreach (DataBlock block in this)
            {
                FormLineList[] formLines = block.GetNearbyTwoLines(horizontialLines);
                if (formLines != null && formLines[0] != null && formLines[1] != null)
                {
                    bottomLines.Add(formLines[0]);
                    topLines.Add(formLines[1]);
                }
            }
            if (bottomLines.Count > 0)
            {
                int minDistanceIndex = 0;
                double minDistance = 0;
                for (int i = 0; i < bottomLines.Count; i++)
                {
                    double distance = bottomLines[i][0].GetDistance(topLines[i][0]);
                    if (i == 0)
                    {
                        minDistance = distance;
                    }
                    else
                    {
                        if (distance < minDistance)
                        {
                            minDistance = distance;
                            minDistanceIndex = i;
                        }
                    }
                }
                return new FormLineList[2] { bottomLines[minDistanceIndex], topLines[minDistanceIndex] };
            }
            return new FormLineList[2] { null, null };
        }

        public void SetBlockDicKey(double key, bool isLeft)
        {
            this.ForEach(block=>block.SetBlockDicKey(key,isLeft));
        }
    }
}
