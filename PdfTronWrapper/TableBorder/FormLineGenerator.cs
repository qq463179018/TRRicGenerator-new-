//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:Providing the function to draw form lines for pdf page.
//-----
//-----------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using pdftron.Common;
using pdftron.PDF;
using PdfTronWrapper.Utility;

namespace PdfTronWrapper.TableBorder
{
    internal class FormLineGenerator
    {

        #region Constractor

        /// <summary>
        /// Constractor Function
        /// </summary>
        /// <param name="pdfPage">Pdf page object</param>
        /// <param name="rectRegion">Rectangle region for the scale of text content</param>
        public FormLineGenerator(Page pdfPage, Rect rectRegion)
        {
            PdfPage = pdfPage;
            RectRegion = rectRegion.Copy();
            _defaultMatrix = pdfPage.GetDefaultMatrix();
            rectRegion.ApplyMatrix(_defaultMatrix);
        }

        #endregion

        #region Private Fields

        /// <summary>
        /// Text blocks dictionary.
        /// </summary>
        public SortedDictionary<HorBlocksInfo, DataBlockList> _dicHorTextBlocks;
        /// <summary>
        /// Text block groups divided by the blocks' left bound.
        /// </summary>
        public SortedDictionary<double, DataBlockList> _dicLeftVerTextBlocks;
        /// <summary>
        /// Text block groups divided by the blocks' right bound.
        /// </summary>
        public SortedDictionary<double, DataBlockList> _dicRightVerTextBlocks;
        /// <summary>
        /// Text columns.
        /// </summary>
        public List<DataColumn> _columns;
        /// <summary>
        /// The default matrix of pdf page.
        /// </summary>
        Matrix2D _defaultMatrix;

        double horizontal_Text_SeparationDistance = 5;
        double vertical_Text_SeparationDistance = 3;

        /// <summary>
        /// The text blocks which are removed provisionally.
        /// </summary>
        DataBlockList removedBlocks;

        //The position scale of text content
        double minLeftBound = -1;
        double maxRightBound = -1;
        double maxTopBound = -1;
        double minBottomBound = -1;

        /// <summary>
        /// Indicate whether the horizontal lines has been generated completely.
        /// </summary>
        bool isHorLinesGenerated = false;

        //The existed lines of pdf page.
        SortedDictionary<double, FormLineList> _existedHorLines;
        SortedDictionary<double, FormLineList> _existedVerLines;

        /// <summary>
        /// Indicate whether need to generate  vertical lines.
        /// </summary>
        bool _isNotNeedGenerateVerLines;

        /// <summary>
        /// Record the page elements which have been processed.
        /// </summary>
        List<CustomElement> _processedElements;

        /// <summary>
        /// Indicate whether the current lines can constitute a rectangle.
        /// </summary>
        bool _isExistRect;

        #endregion

        #region Public Properties

        /// <summary>
        /// The pdf page object which is processing.
        /// </summary>
        public Page PdfPage { get; set; }

        /// <summary>
        /// The rectangle region of processing area. 
        /// </summary>
        public Rect RectRegion { get; set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Get form lines from specify region of a pdf page.
        /// </summary>
        /// <param name="existRealRect">Indicate whether exist a rectangle from in the pdf page.</param>
        /// <param name="horLines">The horizontial lines in the specify region.</param>
        /// <param name="verLines">The vertical lines in the specify region.</param>
        /// <param name="isNotNeedGenerateVerLines">Indicate whether need regenerate vertical lines.</param>
        /// <returns>Completed horizontial and vertical lines for the specify region.</returns>
        public SortedDictionary<double, FormLineList>[] GetFormLines(bool existRealRect,
            SortedDictionary<double, FormLineList> horLines, SortedDictionary<double, FormLineList> verLines,
            bool isNotNeedGenerateVerLines)
        {
            _existedHorLines = horLines;
            _existedVerLines = verLines;
            _isExistRect = existRealRect;
            _isNotNeedGenerateVerLines = isNotNeedGenerateVerLines;
            GetTextBlocks();
            RemoveTooLongAndTooShortTextBlocks();
            if (_dicHorTextBlocks.Count == 0)
            {
                return null;
            }
            NarrowTopBoundOfRegion(verLines);
            RemoveBlocksDividedByVerLines(verLines);
            GenerateLeftRightDicBlocks();
            DivideBlocksIntoGroups();
            NarrowBottomBoundOfRegion(horLines, verLines);
            RemoveNeitherLeftNorRightBlocks(_dicLeftVerTextBlocks, _dicRightVerTextBlocks);
            GenerateColumns();
            if (_columns.Count < 2)
                return null;
            SortedDictionary<double, FormLineList>[] lines = GenericLines(horLines, verLines, isNotNeedGenerateVerLines);
            return lines;
        }

        #endregion

        #region Main Sub Methods

        /// <summary>
        /// Ergodic the paf page,get connective text blocks.
        /// </summary>
        void GetTextBlocks()
        {
            _dicHorTextBlocks = new SortedDictionary<HorBlocksInfo, DataBlockList>();

            using (ElementReader pageReader = new ElementReader())
            {
                pageReader.Begin(PdfPage);
                ProcessElements(pageReader);
            }
        }

        /// <summary>
        /// Remove too long and too short text blocks after ergodicing the pdf page.
        /// </summary>
        void RemoveTooLongAndTooShortTextBlocks()
        {
            removedBlocks = new DataBlockList();
            GetTextBound();
            double contentWidth = maxRightBound - minLeftBound;
            double maxWidth = contentWidth / 2;
            foreach (KeyValuePair<HorBlocksInfo, DataBlockList> pair in _dicHorTextBlocks)
            {
                DataBlockList blockList = pair.Value;
                for (int i = 0; i < blockList.Count; )
                {
                    if (blockList[i].Width > maxWidth || blockList[i].Text.Length == 1)
                    {
                        removedBlocks.Add(blockList[i]);
                        blockList.RemoveAt(i);
                    }
                    else
                    {
                        i++;
                    }
                }
            }
            GenericMethods<HorBlocksInfo, DataBlockList>.RemoveZeroAmountValueItems(_dicHorTextBlocks);
        }

        /// <summary>
        /// Narrow the top bound of text region
        /// </summary>
        /// <param name="verLines">The vertical lines in the text region.</param>
        void NarrowTopBoundOfRegion(SortedDictionary<double, FormLineList> verLines)
        {
            if (_isExistRect || verLines.Count == 0)
                return;

            double topYValueOfVerLines = verLines.SelectMany(lines => lines.Value.Select(line => line.EndPoint.y)).Max();
            double bottomYValueOfVerLines = verLines.SelectMany(lines => lines.Value.Select(line => line.StartPoint.y)).Min();

            RectRegion.y2 = topYValueOfVerLines + 5;
            RectRegion.y1 = bottomYValueOfVerLines - 5;
            _dicHorTextBlocks.Keys.Where(key => key.CenterYValue > topYValueOfVerLines ||
                 key.CenterYValue < bottomYValueOfVerLines
                ).ToList().ForEach(key => _dicHorTextBlocks.Remove(key));
            if (_isNotNeedGenerateVerLines)
            {
                FormLineSearcher.RemoveLinesNotInRect(RectRegion, _existedHorLines);
            }
            else
            {
                FormLineSearcher.RemoveLinesNotInRect(RectRegion, _existedHorLines, verLines);
            }
            if (_existedHorLines.Count > 1 && verLines.Count > 1)
            {
                _isExistRect = FormLineSearcher.IsRect(RectRegion, _existedHorLines, verLines);
            }
            Remove_NotInRegion_RemovedBlocks();
        }

        /// <summary>
        /// Remove blocks which isn't int the text region after narrowing the region.
        /// </summary>
        void Remove_NotInRegion_RemovedBlocks()
        {
            removedBlocks.Where(block => !block.InRegion(RectRegion)).ToList().
                ForEach(block => removedBlocks.Remove(block));
        }

        /// <summary>
        /// Remove blocks which are divided by vertical line.
        /// </summary>
        /// <param name="verLines">The vertical lines in the text region.</param>
        void RemoveBlocksDividedByVerLines(SortedDictionary<double, FormLineList> verLines)
        {
            List<double> xValues = verLines.Keys.ToList();
            if (xValues.Count == 0)
            {
                return;
            }
            foreach (KeyValuePair<HorBlocksInfo, DataBlockList> pair in _dicHorTextBlocks)
            {
                DataBlockList blockList = pair.Value;
                for (int i = 0; i < blockList.Count; )
                {
                    if (xValues.Exists(xValue => blockList[i].IsHorScaleContains(xValue, horizontal_Text_SeparationDistance)))
                    {
                        removedBlocks.Add(blockList[i]);
                        blockList.RemoveAt(i);
                    }
                    else
                    {
                        i++;
                    }
                }
            }
            GenericMethods<HorBlocksInfo, DataBlockList>.RemoveZeroAmountValueItems(_dicHorTextBlocks);
        }

        /// <summary>
        /// Divide blocks into left and right groups.
        /// </summary>
        void DivideBlocksIntoGroups()
        {
            foreach (KeyValuePair<HorBlocksInfo, DataBlockList> pair in _dicHorTextBlocks.Reverse())
            {
                foreach (DataBlock textBlock in pair.Value)
                {
                    ConfirmSide(textBlock);
                }
            }
        }

        /// <summary>
        /// Confirm the side of the data block,left or right or none.
        /// </summary>
        /// <param name="dataBlock"></param>
        void ConfirmSide(DataBlock dataBlock)
        {
            double leftVerKey = dataBlock.LeftBlockDicKey;
            double rightVerKey = dataBlock.RightBlockDicKey;
            DataBlockList rightVerBlocks = _dicRightVerTextBlocks[rightVerKey];
            DataBlockList leftVerBlocks = _dicLeftVerTextBlocks[leftVerKey];
            int side = GetBlockSide(dataBlock, leftVerBlocks, rightVerBlocks);
            if (side == -1)
            {
                dataBlock.IsLeft = true;
                dataBlock.IsRight = false;
                bool result = rightVerBlocks.Remove(dataBlock);

                if (rightVerBlocks.Count == 0)
                {
                    _dicRightVerTextBlocks.Remove(rightVerKey);
                }
            }
            if (side == 1)
            {
                dataBlock.IsRight = true;
                dataBlock.IsLeft = false;
                leftVerBlocks.Remove(dataBlock);
                if (leftVerBlocks.Count == 0)
                {
                    _dicLeftVerTextBlocks.Remove(leftVerKey);
                }
            }
            if (side == 0)
            {
                dataBlock.IsLeft = false;
                dataBlock.IsRight = false;
            }
        }

        /// <summary>
        /// Get the side of data block.
        /// </summary>
        /// <param name="dataBlock">The data block.</param>
        /// <param name="leftVerBlocks">Left block group which contain the data block.</param>
        /// <param name="rightVerBlocks">Right block group which contain the data block.</param>
        /// <returns></returns>
        int GetBlockSide(DataBlock dataBlock, DataBlockList leftVerBlocks, DataBlockList rightVerBlocks)
        {
            int side = 0;
            if (leftVerBlocks.Count > rightVerBlocks.Count)
            {
                side = -1;
            }
            if (rightVerBlocks.Count > leftVerBlocks.Count)
            {
                side = 1;
            }
            if (rightVerBlocks.Count == leftVerBlocks.Count &&
                rightVerBlocks.Count > 1)
            {
                if (rightVerBlocks.DataType != leftVerBlocks.DataType)
                {
                    side = leftVerBlocks.DataType == dataBlock.DataType ? -1 : 1;
                }
                else
                {
                    side = rightVerBlocks.DataType == DataType.Number ? 1 : -1;
                }
            }
            return side;
        }

        /// <summary>
        /// Narrow the bottom bound of text region
        /// </summary>
        /// <param name="horLines">The horizontial lines in the text region.</param>
        /// <param name="verLines">The vertical lines in the text region.</param>
        void NarrowBottomBoundOfRegion(SortedDictionary<double, FormLineList> horLines, SortedDictionary<double, FormLineList> verLines)
        {
            if (_isExistRect)
                return;
            //Get all the text blocks whose data type are number
            List<DataBlockList> blockLists =
            _dicLeftVerTextBlocks.Concat(_dicRightVerTextBlocks)
            .Where(pair => pair.Value.DataType == DataType.Number)
            .Select(pair => pair.Value).ToList();
            if (blockLists.Count == 0)
                return;
            double bottomBound = RectRegion.y1;
            int maxAmount = blockLists.Select(blocks => blocks.Count()).Max();
            //Get the lowest text block.
            DataBlock lastNumBlock = null;
            foreach (DataBlockList blockList in blockLists)
            {
                DataBlock _lastNumBlock = blockList.GetLastNumBlock(_dicHorTextBlocks);
                if (_lastNumBlock != null)
                {
                    if (lastNumBlock == null || _lastNumBlock.BottomBound < bottomBound)
                    {
                        bottomBound = _lastNumBlock.BottomBound;
                        lastNumBlock = _lastNumBlock;
                    }
                }
            }
            //Get new bottom bound of rectangle region.
            if (lastNumBlock != null)
            {
                DataBlockList blockList = GetBelowIntersectBlocks(lastNumBlock);
                if (blockList != null)
                    bottomBound = blockList.BottomBound;
            }
            //Narrow the rectangle region by new bottom bound.
            if (bottomBound >= RectRegion.y1 && bottomBound < RectRegion.y2)
            {
                if (verLines.Count == 0 || bottomBound < verLines.SelectMany(x => x.Value).Select(line => line.StartPoint.y).Min())
                {
                    RectRegion.y1 = bottomBound - 5;
                }
                else
                {
                    RectRegion.y1 = verLines.SelectMany(x => x.Value).Select(line => line.StartPoint.y).Min() - 2;
                }
                _dicHorTextBlocks.Where(pair => pair.Key.IsBelow(RectRegion.y1))
                    .Select(pair => pair.Key).ToList()
                    .ForEach(key =>
                    {
                        _dicHorTextBlocks[key].ForEach(block =>
                        {
                            if (removedBlocks.Contains(block))
                            {
                                removedBlocks.Remove(block);
                            }
                        });
                        _dicHorTextBlocks.Remove(key);
                    });
                Remove_NotInRegion_RemovedBlocks();
                GenerateLeftRightDicBlocks();
                DivideBlocksIntoGroups();
                RemoveNeitherLeftNorRightBlocks(_dicLeftVerTextBlocks, _dicRightVerTextBlocks);
                if (_isNotNeedGenerateVerLines)
                {
                    FormLineSearcher.RemoveLinesNotInRect(RectRegion, horLines);
                }
                else
                {
                    FormLineSearcher.RemoveLinesNotInRect(RectRegion, horLines, verLines);
                }
            }
        }

        /// <summary>
        /// Get the block which is below it and is intersect with it in horizontial direction.
        /// </summary>
        /// <param name="dataBlock">it</param>
        /// <returns></returns>
        DataBlockList GetBelowIntersectBlocks(DataBlock dataBlock)
        {
            DataBlockList blockList = null;
            int index = 0;
            foreach (KeyValuePair<HorBlocksInfo, DataBlockList> pair in _dicHorTextBlocks)
            {
                if (pair.Value.Contains(dataBlock) && index > 0)
                {
                    HorBlocksInfo belowInfo = _dicHorTextBlocks.Keys.ToList()[index - 1];
                    double belowRowTop = belowInfo.TopBound;
                    if (belowRowTop > pair.Key.BottomBound)
                    {
                        blockList = _dicHorTextBlocks[belowInfo];
                        break;
                    }
                }
                index++;
            }
            return blockList;
        }

        /// <summary>
        /// Generate data columns by blocks.
        /// </summary>
        void GenerateColumns()
        {
            List<DataColumn> rightTextColumns = GenerateVerticalColumns(_dicRightVerTextBlocks);
            RemoveBlocksOfNotRightColumn(rightTextColumns);
            RemoveLeftBlocksCoverRightColumn(rightTextColumns);
            List<DataColumn> leftTextColumns = GenerateVerticalColumns(_dicLeftVerTextBlocks);
            _columns = rightTextColumns.Concat(leftTextColumns).ToList();
            GenerateColumnsFromRemovedBlocks();
            _columns.Sort();
            ExtendColumnBound();
            RevertRemovedBlocks();
            GenerateColumnIndexesOfBlocks();
        }

        /// <summary>
        /// Generate lines for the text in the special rect region.
        /// </summary>
        /// <param name="existHorLines"></param>
        /// <param name="existVerLines"></param>
        /// <param name="isNotNeedGenerateVerLines"></param>
        /// <returns></returns>
        SortedDictionary<double, FormLineList>[] GenericLines(SortedDictionary<double, FormLineList> existHorLines,
            SortedDictionary<double, FormLineList> existVerLines, bool isNotNeedGenerateVerLines)
        {
            SortedDictionary<double, FormLineList> horizontialLines = GenerateHorizontialLines(existHorLines);
            JoinLittleDistanceLinesTogether(horizontialLines, true);
            SortedDictionary<double, FormLineList> verticalLines = existVerLines;

            if (!isNotNeedGenerateVerLines)
            {
                verticalLines = GenerateVerticalLines(horizontialLines, existVerLines);
            }
            JoinLittleDistanceLinesTogether(verticalLines, false);
            return new SortedDictionary<double, FormLineList>[2] { horizontialLines, verticalLines };
        }

        #endregion

        #region Generate Columns

        void ExtendColumnBound()
        {
            for (int i = 0; i < _columns.Count - 1; i++)
            {
                DataColumn column = _columns[i];
                DataColumn nextColumn = _columns[i + 1];
                if (i == 0)
                {
                    column.LeftBound = column.LeftDataBound;
                }
                double bound = (column.RightDataBound + nextColumn.LeftDataBound) / 2;
                column.RightBound = nextColumn.LeftBound = bound;
                if (i == _columns.Count - 2)
                {
                    nextColumn.RightBound = nextColumn.RightDataBound;
                }
            }
        }

        List<DataColumn> GenerateColumns(SortedDictionary<double, DataBlockList> dic)
        {
            List<DataColumn> columns =
            dic.Select(pair => new DataColumn
            {
                LeftDataBound = pair.Value.Select(block => block.LeftBound).Min(),
                RightDataBound = pair.Value.Select(block => block.RightBound).Max(),
                TextBlocks = pair.Value
            }).ToList();
            return columns;
        }

        List<DataColumn> GenerateVerticalColumns(SortedDictionary<double, DataBlockList> dic)
        {
            List<DataColumn> columns = new List<DataColumn>();
            IEnumerable<double> yValues = _dicHorTextBlocks.Keys.Select(key => key.CenterYValue);
            foreach (double yValue in yValues)
            {
                foreach (KeyValuePair<double, DataBlockList> pair in dic)
                {
                    //Find same level text block.
                    double rightXValue = pair.Key;
                    DataBlockList blockList = pair.Value;
                    DataBlock sameLevelBlock = blockList.FirstOrDefault(block => block.BottomBound < yValue && block.TopBound > yValue);
                    if (sameLevelBlock == null)
                    {
                        continue;
                    }

                    //Judge whether exist intersect column.
                    List<DataColumn> intersectColumns = columns.Where(column => sameLevelBlock.IsIntersect(column)).ToList();
                    if (intersectColumns.Count > 1)
                    {
                        removedBlocks.Add(sameLevelBlock);
                        blockList.Remove(sameLevelBlock);
                        continue;
                    }
                    //Generate new column or update bound info of existed column.
                    if (intersectColumns.Count == 0)
                    {
                        DataColumn newColumn = new DataColumn
                        {
                            LeftDataBound = sameLevelBlock.LeftBound,
                            RightDataBound = sameLevelBlock.RightBound,
                            TextBlocks = new DataBlockList { sameLevelBlock }
                        };
                        columns.Add(newColumn);
                    }
                    else
                    {
                        DataColumn existedColumn = intersectColumns[0];
                        existedColumn.LeftDataBound = Math.Min(existedColumn.LeftDataBound, sameLevelBlock.LeftBound);
                        existedColumn.RightDataBound = Math.Max(existedColumn.RightDataBound, sameLevelBlock.RightBound);
                        existedColumn.TextBlocks.Add(sameLevelBlock);
                    }
                }
            }
            GenericMethods<double, DataBlockList>.RemoveZeroAmountValueItems(dic);
            return columns;
        }

        void GenerateColumnsFromRemovedBlocks()
        {
            removedBlocks.Sort();
            List<DataBlock> dealedBlocks = new List<DataBlock>();

            for (int i = 0; i < removedBlocks.Count; i++)
            {
                DataBlock removedBlock = removedBlocks[i];
                if (dealedBlocks.Exists(_block => _block.Equals(removedBlock)))
                    continue;
                List<DataColumn> intersectColumns = removedBlock.GetIntersectColumns(_columns);
                if (intersectColumns.Count == 0)
                {
                    DataColumn newColumn = new DataColumn
                    {
                        LeftDataBound = removedBlock.LeftBound,
                        RightDataBound = removedBlock.RightBound,
                        TextBlocks = new DataBlockList { removedBlock }
                    };
                    _columns.Add(newColumn);
                }
                else
                {
                    if (intersectColumns.Count == 1)
                    {
                        intersectColumns[0].LeftDataBound = Math.Min(intersectColumns[0].LeftDataBound, removedBlock.LeftBound);
                        intersectColumns[0].RightDataBound = Math.Max(intersectColumns[0].RightDataBound, removedBlock.RightBound);
                        intersectColumns[0].TextBlocks.Add(removedBlock);
                    }
                    else
                    {
                        intersectColumns.ForEach(column =>
                        column.TextBlocks.Add(removedBlock));
                    }
                }
                dealedBlocks.Add(removedBlock);
            }
        }

        void GenerateColumnIndexesOfBlocks()
        {
            foreach (KeyValuePair<HorBlocksInfo, DataBlockList> pair in _dicHorTextBlocks)
            {
                DataBlockList blocks = pair.Value;
                foreach (DataBlock block in blocks)
                {
                    block.SetColumnIndexes(_columns);
                }
            }
        }

        #endregion

        #region Generate Lines

        SortedDictionary<double, FormLineList> GenerateHorizontialLines(SortedDictionary<double, FormLineList> existLines)
        {
            if (isHorLinesGenerated)
            {
                CheckBoundHorizontialLines(existLines);
                return existLines;
            }

            SortedDictionary<double, FormLineList> lines = new SortedDictionary<double, FormLineList>();
            GetTextBound();
            RemoveSpareNearLines(existLines, true);
            CheckBoundHorizontialLines(existLines);
            bool isComplete = IsHorizontialLinesComplete(existLines);
            if (isComplete)
            {
                lines = existLines;
            }
            else
            {
                double[] pageSize = PdfTronHelper.GetPageSize(PdfPage);

                double leftX = minLeftBound - 3 > 0 ? minLeftBound - 3 : 0;
                double rightX = maxRightBound + 3 < pageSize[0] ? maxRightBound + 3 : pageSize[0];
                double topY = maxTopBound + 3 < pageSize[1] ? maxTopBound + 3 : pageSize[1];
                double bottomY = minBottomBound - 3 > 0 ? minBottomBound - 3 : 0;

                if (minBottomBound < RectRegion.y1 - 10)
                {
                    bottomY = RectRegion.y1 - 5;
                }

                List<double> yValues = new List<double>();
                List<HorBlocksInfo> horBlocksInfos = _dicHorTextBlocks.Keys.ToList();
                lines.Add(bottomY, new FormLineList { FormLine.GenerateHorizontialLine(leftX, rightX, bottomY) });

                int topIndex = _dicHorTextBlocks.Count - 1;
                if (existLines.Count > 1)
                {
                    topIndex = GetTopRowIndex_NeedDrawHorizontialLine(existLines);
                }
                for (int i = 0; i < topIndex; i++)
                {
                    DataBlockList curRowBlocks = _dicHorTextBlocks[horBlocksInfos[i]];
                    DataBlockList highRowBlocks = _dicHorTextBlocks[horBlocksInfos[i + 1]];
                    double curRowTop = curRowBlocks.TopBound;
                    double highRowBottom = highRowBlocks.BottomBound;
                    int internalLineIndex = i;
                    if (curRowTop > highRowBottom && IsOneToTwo(ref internalLineIndex))
                    {
                        if (internalLineIndex == i)
                        {
                            FormLineList seperateLines = GenerateSeperateHorizontialLines(internalLineIndex);
                            if (seperateLines.Count > 0)
                            {
                                lines.Add(seperateLines[0].StartPoint.y, seperateLines);
                            }
                        }
                    }
                    else
                    {
                        internalLineIndex++;
                        if (internalLineIndex > horBlocksInfos.Count - 2 || !IsOneToTwo(ref internalLineIndex))
                        {
                            double yVlaue = (curRowTop + highRowBottom) / 2;
                            lines.Add(yVlaue, new FormLineList { FormLine.GenerateHorizontialLine(leftX, rightX, yVlaue) });
                        }
                    }

                    //Remove spare hoizontial lines
                    if (lines.Count > 3)
                    {
                        RemoveSpareHoriontialLines(lines);
                    }
                }
                lines.Add(topY, new FormLineList { FormLine.GenerateHorizontialLine(leftX, rightX, topY) });
                existLines.Keys.ToList().ForEach(key => lines.Add(key, existLines[key]));
                RemoveSpareNearLines(lines, true);
            }
            return lines;
        }

        SortedDictionary<double, FormLineList> GenerateVerticalLines(SortedDictionary<double, FormLineList> horizontialLines,
            SortedDictionary<double, FormLineList> existVerticalLines)
        {
            RemoveSpareNearLines(existVerticalLines, false);

            existVerticalLines.Where(pair =>
            {
                double amountLenght = pair.Value.Sum(line => line.Length);
                return amountLenght < (horizontialLines.Last().Key - horizontialLines.First().Key) * 0.45;
            }).Select(pair => pair.Key).ToList()
            .ForEach(key => existVerticalLines.Remove(key));

            double[] horBound = GetStartEndXValueOfHorizontialLines(horizontialLines);
            double leftX = horBound[0];
            double rightX = horBound[1];
            double topY = horizontialLines.Last().Value[0].StartPoint.y;
            double bottomY = horizontialLines.First().Value[0].StartPoint.y;

            List<KeyValuePair<double, FormLineList>> existLines = GetExistVerticalLine(existVerticalLines, leftX, _columns[0].LeftDataBound);
            if (existLines.Count == 0 && !existVerticalLines.ContainsKey(leftX))
            {
                existVerticalLines.Add(leftX, new FormLineList { FormLine.GenerateVerticalLine(bottomY, topY, leftX) });
            }
            existLines = GetExistVerticalLine(existVerticalLines, _columns.Last().RightDataBound, rightX);
            if (existLines.Count == 0 && !existVerticalLines.ContainsKey(rightX))
            {
                existVerticalLines.Add(rightX, new FormLineList { FormLine.GenerateVerticalLine(bottomY, topY, rightX) });
            }

            //if (IsVerticalLinesComplete(existVerticalLines))
            return existVerticalLines;

            //SortedDictionary<double, FormLineList> lines = new SortedDictionary<double, FormLineList>();
            //for (int i = 0; i < _columns.Count - 1; i++)
            //{
            //    existLines = GetExistVerticalLine(existVerticalLines, _columns[i], _columns[i + 1]);
            //    if (existLines != null && existLines.Count > 0)
            //    {
            //        existLines.ForEach(pair =>
            //        {
            //            if (!lines.ContainsKey(pair.Key))
            //                lines.Add(pair.Key, pair.Value);
            //            existVerticalLines.Remove(pair.Key);
            //        });
            //    }
            //    else
            //    {
            //        double x = (_columns[i].RightDataBound + _columns[i + 1].LeftDataBound) / 2;
            //        FormLine line = FormLine.GenerateVerticalLine(bottomY, topY, x);
            //        FormLineList _lines = GetDividedLines(line, i, i + 1, horizontialLines);
            //        if (_lines.Count > 0)
            //        {
            //            if (!lines.ContainsKey(x))
            //                lines.Add(x, _lines);
            //        }
            //    }
            //}

            //existVerticalLines.Keys.ToList().ForEach(key =>
            //{
            //    if (!lines.ContainsKey(key))
            //    {
            //        lines.Add(key, existVerticalLines[key]);
            //    }
            //});
            //RemoveSpareNearLines(lines, false);

            //return lines;
        }

        List<KeyValuePair<double, FormLineList>> GetExistVerticalLine(SortedDictionary<double, FormLineList> existLines, DataColumn leftColumn, DataColumn rightColumn)
        {
            List<KeyValuePair<double, FormLineList>> _lines = GetExistVerticalLine(existLines, leftColumn.RightDataBound, rightColumn.LeftDataBound);
            if (_lines.Count > 0)
            {
                return _lines;
            }
            _lines = GetExistVerticalLine(existLines, leftColumn.LeftDataBound + leftColumn.DataWidth * 0.3, rightColumn.LeftDataBound);
            return _lines;
        }

        List<KeyValuePair<double, FormLineList>> GetExistVerticalLine(SortedDictionary<double, FormLineList> existLines, double leftBound, double rightBound)
        {
            List<KeyValuePair<double, FormLineList>> lines = new List<KeyValuePair<double, FormLineList>>();
            foreach (KeyValuePair<double, FormLineList> pair in existLines)
            {
                double xValue = pair.Key;
                if (xValue < leftBound - horizontal_Text_SeparationDistance || xValue > rightBound + horizontal_Text_SeparationDistance)
                {
                    continue;
                }

                double amountLength = pair.Value.Sum(line => line.Length);
                double dataHeight = _dicHorTextBlocks.Last().Key.TopBound - _dicHorTextBlocks.First().Key.BottomBound;
                if (amountLength < dataHeight * 0.8)
                {
                    continue;
                }
                lines.Add(pair);
            }
            return lines;
        }

        FormLineList GetDividedLines(FormLine line, int leftColIndex, int rightColIndex, SortedDictionary<double, FormLineList> horizontialLines)
        {
            FormLineList lines = new FormLineList();
            List<double[]> removeSegments = new List<double[]>();
            for (int i = 0; i < _dicHorTextBlocks.Count; i++)
            {
                HorBlocksInfo key = _dicHorTextBlocks.Keys.ToList()[i];
                if (_dicHorTextBlocks[key].Exists(
                    block => new List<int> { leftColIndex, rightColIndex }
                        .All(index => block.ColumnIndexes.Contains(index))))
                {
                    FormLineList[] nearbyLines = _dicHorTextBlocks[key].GetNearbyTwoLines(horizontialLines);
                    if (nearbyLines[0] == null || nearbyLines[1] == null)
                    {
                        continue;
                    }
                    if (removeSegments.Count > 0 &&
                        removeSegments[removeSegments.Count - 1][1] == nearbyLines[0][0].StartPoint.y)
                    {
                        removeSegments[removeSegments.Count - 1][1] = nearbyLines[1][0].StartPoint.y;
                    }
                    else
                    {
                        double[] segment = new double[2] { nearbyLines[0][0].StartPoint.y, nearbyLines[1][0].StartPoint.y };
                        if (!removeSegments.Exists(seg => seg[0] == segment[0] && seg[1] == segment[1]))
                        {
                            removeSegments.Add(segment);
                        }
                    }
                }
            }
            List<double[]> lineSegments = GetLineSegments(line, removeSegments);
            lines = new FormLineList(lineSegments.Select(segment => FormLine.GenerateVerticalLine(segment[0], segment[1], line.StartPoint.x)).ToList());
            return lines;
        }

        List<double[]> GetLineSegments(FormLine line, List<double[]> removeSegments)
        {
            List<double[]> lineSegments = new List<double[]>()
            {
                new double[2] 
                { 
                    line.IsTransverseLine? line.StartPoint.x: line.StartPoint.y, 
                    line.IsTransverseLine? line.EndPoint.x: line.EndPoint.y
                }
            };
            while (removeSegments.Count > 0)
            {
                foreach (double[] removeSegment in removeSegments)
                {
                    double[] containSegment = lineSegments.FirstOrDefault(
                        segment => segment[0] <= removeSegment[0] &&
                            segment[1] >= removeSegment[1]);
                    if (containSegment != null)
                    {
                        List<double[]> newSegments = new List<double[]>
                        {
                            new double[2]{containSegment[0],removeSegment[0]},
                            new double[2]{removeSegment[1],containSegment[1]},
                        };
                        newSegments.ForEach(segment =>
                        {
                            if (segment[0] != segment[1])
                            {
                                lineSegments.Add(segment);
                            }
                        });
                        lineSegments.Remove(containSegment);
                        removeSegments.Remove(removeSegment);
                        break;
                    }
                    else
                    {
                        removeSegments.Remove(removeSegment);
                        break;
                    }
                }
            }

            return lineSegments;
        }

        FormLineList GenerateSeperateHorizontialLines(int horLineIndex)
        {
            FormLineList lines = new FormLineList();
            List<HorBlocksInfo> horBlockInfos = _dicHorTextBlocks.Keys.ToList();
            double lowRowTop = horBlockInfos[horLineIndex - 1].TopBound;
            double lowRowBottom = horBlockInfos[horLineIndex - 1].BottomBound;
            double highRowTop = horBlockInfos[horLineIndex + 1].TopBound;
            double highRowBottom = horBlockInfos[horLineIndex + 1].BottomBound;
            if (lowRowTop <= highRowBottom)
            {
                double yValue = (lowRowTop + highRowBottom) / 2;
                FormLine line = new FormLine(
                    new Point(_columns[0].LeftBound, yValue),
                    new Point(_columns[_columns.Count - 1].RightBound, yValue),
                    false);

                List<double[]> removeSegments = new List<double[]>();
                foreach (DataColumn column in _columns)
                {
                    if (column.ColumnType == DataType.Text ||
                        column.TextBlocks.GetBlocksOfScale(lowRowBottom, highRowTop).Count < 2)
                        removeSegments.Add(new double[2] { column.LeftBound, column.RightBound });
                }

                List<double[]> lineSegments = GetLineSegments(line, removeSegments);
                lines = new FormLineList(lineSegments.Select(
                    segment => FormLine.GenerateHorizontialLine(segment[0], segment[1], line.StartPoint.y)
                    ).ToList());
            }
            return lines;
        }

        void RemoveSpareHoriontialLines(SortedDictionary<double, FormLineList> horizontialLines)
        {
            double[] keys = horizontialLines.Keys.ToArray();
            int count = horizontialLines.Count;
            double lowestYValue = keys[count - 4];
            double lowerYValue = keys[count - 3];
            double higherYValue = keys[count - 2];
            double highestYValue = keys[count - 1];

            List<DataBlock> lowBlocks = GetBlocksBetweenHorizontialLines(lowestYValue, lowerYValue);
            List<DataBlock> middleBlocks = GetBlocksBetweenHorizontialLines(lowerYValue, higherYValue);
            List<DataBlock> highBlocks = GetBlocksBetweenHorizontialLines(higherYValue, highestYValue);

            if (lowBlocks.Count > 0 && highBlocks.Count > 0 && middleBlocks.Count > 0)
            {
                List<int> numberColumnIndexes = _columns.Where(col => col.ColumnType == DataType.Number)
                    .Select(col => _columns.IndexOf(col)).ToList();
                List<int> textColumnIndexes = _columns.Where(col => col.ColumnType == DataType.Text)
                    .Select(col => _columns.IndexOf(col)).ToList();

                if (lowBlocks.Concat(highBlocks).All
                    (block => block.ColumnIndexes.All
                        (index => textColumnIndexes.Contains(index))) &&

                    middleBlocks.All
                    (block => block.DataType == DataType.Number && block.ColumnIndexes.All
                        (index => numberColumnIndexes.Contains(index))))
                {
                    horizontialLines.Remove(lowerYValue);
                    horizontialLines.Remove(higherYValue);
                }
            }
        }

        void JoinLittleDistanceLinesTogether(SortedDictionary<double, FormLineList> lines, bool isHorizontial)
        {
            int pairCount = lines.Count;
            double minDistance = isHorizontial ? 18 : 9;
            for (int i = 0; i < lines.Count; i++)
            {
                FormLineList _lines = lines[lines.Keys.ToArray()[i]];
                if (_lines.Count < 2)
                {
                    continue;
                }

                _lines.Sort();
                for (int j = 0; j < _lines.Count - 1; )
                {
                    if (_lines[j].GetDistanceOfNearestExtremePoints(_lines[j + 1]) < minDistance)
                    {
                        _lines[j].MergeLine(_lines[j + 1]);
                        _lines.RemoveAt(j + 1);
                    }
                    else
                    {
                        j++;
                    }
                }
            }
        }

        int GetTopRowIndex_NeedDrawHorizontialLine(SortedDictionary<double, FormLineList> existLines)
        {
            int index = _dicHorTextBlocks.Count - 1;
            List<HorBlocksInfo> horBlocksInfos = _dicHorTextBlocks.Keys.ToList();
            List<double> existLineYValues = existLines.Keys.ToList();
            if (existLines.Count > 2)
            {
                for (int i = existLines.Count - 1; i > 0; i--)
                {
                    int j = i - 1;
                    if (existLineYValues[i] - existLineYValues[j] > 100)
                    {
                        for (int k = _dicHorTextBlocks.Count - 1; k > 0; k--)
                        {
                            if (horBlocksInfos[k].BottomBound >= existLineYValues[i] && horBlocksInfos[k - 1].TopBound <= existLineYValues[i])
                            {
                                index = k - 1;
                                break;
                            }
                        }
                        break;
                    }
                }
            }

            return index;
        }

        public static void RemoveSpareNearLines(SortedDictionary<double, FormLineList> lines, bool isHorizontial)
        {
            for (int i = 0; i < lines.Count - 1; )
            {
                double[] keys = lines.Keys.ToArray();
                double leftKey = keys[i];
                double rightKey = keys[i + 1];
                if (rightKey - leftKey < (isHorizontial ? 7 : 14))
                {
                    double leftAmountLenght = lines[leftKey].Sum(line => line.Length);
                    double rightAmountLenght = lines[rightKey].Sum(line => line.Length);
                    if (rightAmountLenght > leftAmountLenght)
                    {
                        lines.Remove(leftKey);
                    }
                    else
                    {
                        lines.Remove(rightKey);
                    }
                }
                else
                {
                    i++;
                }
            }
        }

        void CheckBoundHorizontialLines(SortedDictionary<double, FormLineList> existLines)
        {
            HorBlocksInfo bottomHorInfo = _dicHorTextBlocks.First().Key;
            double bottomBlocksCenterYValue = bottomHorInfo.CenterYValue;
            if (existLines.Count == 0 || bottomBlocksCenterYValue < existLines.First().Key)
            {
                double yValue = bottomHorInfo.BottomBound;
                existLines.Add(yValue,
                    new FormLineList{
                        FormLine.GenerateHorizontialLine(minLeftBound,maxRightBound,yValue)
                    });
            }

            HorBlocksInfo topHorInfo = _dicHorTextBlocks.Last().Key;
            double topBlocksCenterYValue = topHorInfo.CenterYValue;
            if (existLines.Count == 0 || topBlocksCenterYValue > existLines.Last().Key)
            {
                double yValue = topHorInfo.TopBound;
                existLines.Add(yValue,
                    new FormLineList{
                        FormLine.GenerateHorizontialLine(minLeftBound,maxRightBound,yValue)
                    });
            }
        }

        #endregion

        #region Private Methods

        void RevertRemovedBlocks()
        {
            List<DataBlock> dealedBlocks = new List<DataBlock>();
            removedBlocks.ForEach(block =>
                {
                    if (dealedBlocks.Exists(_block => _block.Equals(block)))
                        return;

                    HorBlocksInfo horInfo = block.HorBlockDicKey;
                    if (_dicHorTextBlocks.ContainsKey(horInfo))
                    {
                        if (!_dicHorTextBlocks[horInfo].Contains(block))
                        {
                            _dicHorTextBlocks[horInfo].Add(block);
                        }
                    }
                    else
                    {
                        _dicHorTextBlocks.Add(horInfo, new DataBlockList { block });
                    }

                    dealedBlocks.Add(block);
                });
        }

        void SortLeftRightBlocks()
        {
            Action<SortedDictionary<double, DataBlockList>> sortMethod = dic =>
                {
                    foreach (KeyValuePair<double, DataBlockList> pair in dic)
                    {
                        pair.Value.Sort();
                    }
                };
            sortMethod(_dicLeftVerTextBlocks);
            sortMethod(_dicRightVerTextBlocks);
        }

        bool IsExistBlocksBetweenHorizontialLines(FormLineList lines1, FormLineList lines2)
        {
            List<double> yValues = new List<double> { lines1[0].StartPoint.y, lines2[0].StartPoint.y };
            double lines1LeftX = lines1.Select(line => line.StartPoint.x).Min();
            double lines1RightX = lines1.Select(line => line.EndPoint.x).Max();
            double lines2LeftX = lines2.Select(line => line.StartPoint.x).Min();
            double lines2RightX = lines2.Select(line => line.EndPoint.x).Max();
            double leftX = Math.Max(lines1LeftX, lines2LeftX);
            double rightX = Math.Min(lines1RightX, lines2RightX);
            double minNum = yValues.Min();
            double maxNum = yValues.Max();
            return _dicHorTextBlocks.Keys.ToList().Exists(horBlockInfo =>
            {
                double medialNum = horBlockInfo.CenterYValue;
                return medialNum > minNum && medialNum < maxNum
                    && _dicHorTextBlocks[horBlockInfo].Exists(block => block.IsHorCenterBetween(leftX, rightX));
            })

            ||

           removedBlocks.Exists(block => block.IsHorCenterBetween(leftX, rightX) && block.IsVerCenterBetween(minNum, maxNum));
        }

        List<DataBlock> GetBlocksBetweenHorizontialLines(double y1, double y2)
        {
            List<double> nums = new List<double> { y1, y2 };
            double minNum = nums.Min();
            double maxNum = nums.Max();
            return _dicHorTextBlocks.Where(pair =>
            {
                double medialNum = (pair.Key.BottomBound + pair.Key.TopBound) / 2;
                return medialNum > minNum && medialNum < maxNum;
            }).SelectMany(pair => pair.Value).ToList();
        }

        bool IsHorizontialLinesComplete(SortedDictionary<double, FormLineList> existLines)
        {
            bool isComplete;
            if (existLines.Count < 2)
            {
                isComplete = false;
            }
            else
            {
                double distanceBetweenFirstLastLines = existLines.First().Value[0].GetDistance(existLines.Last().Value[0]);
                double minHeight = (_dicHorTextBlocks.Last().Key.TopBound - _dicHorTextBlocks.First().Key.BottomBound) * 0.8;
                if (distanceBetweenFirstLastLines < minHeight)
                {
                    isComplete = false;
                }
                else
                {
                    isComplete = Validate_NeighbouringHorizontialLines_Distance(existLines);
                }
            }
            return isComplete;
        }

        bool Validate_NeighbouringHorizontialLines_Distance(SortedDictionary<double, FormLineList> horizontialLines)
        {
            bool result = true;
            List<double> distinces = new List<double>();
            for (int index = 0; index < horizontialLines.Count - 1; index++)
            {
                FormLine currentLine = horizontialLines[horizontialLines.Keys.ToArray()[index]][0];
                FormLine nextLine = horizontialLines[horizontialLines.Keys.ToArray()[index + 1]][0];

                double distance = currentLine.GetDistance(nextLine);
                double maxTimeNum = 5;
                if (distinces.Exists(item => item / distance > maxTimeNum || distance / item > maxTimeNum))
                {
                    result = false;
                    continue;
                }
                distinces.Add(distance);
            }
            if (result)
            {
                double averageDistance = distinces.Average();
                result = averageDistance < 16 * 3;
            }
            return result;
        }

        bool Validate_NeighbouringVerticalLines_Distance(SortedDictionary<double, FormLineList> vertiaclLines)
        {
            if (vertiaclLines.Count < 3)
                return false;

            if (vertiaclLines.Count == _columns.Count + 1)
                return true;

            List<double> xValues = vertiaclLines.Keys.ToList();
            for (int i = 0; i < xValues.Count - 1; i++)
            {
                double leftXValue = xValues[i];
                double rightXValue = xValues[i + 1];
                int rowCount_BlocksAmountLargeThan2 = 0;
                int rowCount_BlocksAmountLargeThan0 = 0;
                foreach (KeyValuePair<HorBlocksInfo, DataBlockList> pair in _dicHorTextBlocks)
                {
                    int amount = 0;
                    DataBlockList blocks = pair.Value;
                    foreach (DataBlock block in blocks)
                    {
                        if (block.IsHorCenterBetween(leftXValue, rightXValue))
                        {
                            amount++;
                            if (amount > 1)
                            {
                                rowCount_BlocksAmountLargeThan2++;
                                break;
                            }
                        }
                    }
                    if (amount > 0)
                    {
                        rowCount_BlocksAmountLargeThan0++;
                    }
                }
                if (rowCount_BlocksAmountLargeThan0 > 3 && rowCount_BlocksAmountLargeThan2 > rowCount_BlocksAmountLargeThan0 * 0.33)
                {
                    return false;
                }
            }
            return true;
        }

        bool IsVerticalLinesComplete(SortedDictionary<double, FormLineList> existLines)
        {
            bool isComplete;
            if (existLines.Count < 2)
            {
                isComplete = false;
            }
            else
            {
                double distanceBetweenFirstLastLines = existLines.First().Value[0].GetDistance(existLines.Last().Value[0]);
                double minWidth = (maxRightBound - minLeftBound) * 0.8;
                if (distanceBetweenFirstLastLines < minWidth)
                {
                    isComplete = false;
                }
                else
                {
                    isComplete = Validate_NeighbouringVerticalLines_Distance(existLines);
                }
            }
            return isComplete;
        }

        bool IsOneToTwo(ref int index)
        {
            List<HorBlocksInfo> horBlocksInfos = _dicHorTextBlocks.Keys.ToList();
            bool result = false;

            int blockLineCount = _dicHorTextBlocks.Count;

            if (blockLineCount > 2)
            {
                bool isBottomRow = index == 0;
                int curRowIndex = index;

                if (isBottomRow)
                {
                    curRowIndex = index + 1;
                }

                result = OneColumnToTwo(_dicHorTextBlocks, curRowIndex);
                if (result)
                {
                    index = curRowIndex;
                }
            }
            return result;
        }

        public static bool OneColumnToTwo(SortedDictionary<HorBlocksInfo, DataBlockList> horBlocks, int curRowIndex)
        {
            List<HorBlocksInfo> horBlocksInfos = horBlocks.Keys.ToList();
            bool result = false;

            if (curRowIndex > 0 && curRowIndex < horBlocksInfos.Count - 1)
            {
                DataBlockList curRowBlocks = horBlocks[horBlocksInfos[curRowIndex]];
                DataBlockList lowRowBlocks = horBlocks[horBlocksInfos[curRowIndex - 1]];
                DataBlockList highRowBlocks = horBlocks[horBlocksInfos[curRowIndex + 1]];
                double curRowTop = curRowBlocks.TopBound;
                double curRowBottom = curRowBlocks.BottomBound;
                double highRowBottom = highRowBlocks.BottomBound;
                double lowRowTop = lowRowBlocks.TopBound;
                double upIntersectHeight = curRowTop > highRowBottom ? curRowTop - highRowBottom : 0;
                double downIntersectHeight = lowRowTop > curRowBottom ? lowRowTop - curRowBottom : 0;

                if (upIntersectHeight > 0 && downIntersectHeight > 0 &&
                    upIntersectHeight + downIntersectHeight > horBlocks[horBlocksInfos[curRowIndex]].Hight * 0.25)
                {
                    result = true;
                }
            }
            return result;
        }

        double[] GetStartEndXValueOfHorizontialLines(SortedDictionary<double, FormLineList> horizontialLines)
        {
            double[] xValues = new double[2];
            double maxLength = horizontialLines.SelectMany(x => x.Value).Select(line => line.Length).Max();
            IEnumerable<FormLine> longLines = horizontialLines.SelectMany(x => x.Value).Where(line => Math.Abs(line.Length - maxLength) < 10);
            xValues[0] = longLines.Select(line => line.StartPoint.x).Min();
            xValues[1] = longLines.Select(line => line.EndPoint.x).Max();
            return xValues;
        }

        void GetTextBound()
        {
            List<DataBlock> blocks = _dicHorTextBlocks.SelectMany(pair => pair.Value).ToList();
            if (blocks.Count > 0)
            {
                minLeftBound = blocks.Select(block => block.LeftBound).Min();
                maxRightBound = blocks.Select(block => block.RightBound).Max();
            }
            else
            {
                minLeftBound = RectRegion.x1;
                maxRightBound = RectRegion.x2;
            }
            if (_existedHorLines.Count > 0)
            {
                double horLinesMinLeftXValue = _existedHorLines.SelectMany(pair => pair.Value.Select(line => line.StartPoint.x)).Min();
                double horLinesMaxRightXValue = _existedHorLines.SelectMany(pair => pair.Value.Select(line => line.EndPoint.x)).Max();
                minLeftBound = Math.Min(minLeftBound, horLinesMinLeftXValue);
                maxRightBound = Math.Max(maxRightBound, horLinesMaxRightXValue);
            }
            if (blocks.Count > 0)
            {
                maxTopBound = blocks.Select(block => block.TopBound).Max();
                minBottomBound = blocks.Select(block => block.BottomBound).Min();
            }
            else
            {
                maxTopBound = RectRegion.y2;
                minBottomBound = RectRegion.y1;
            }
        }

        void RemoveLeftBlocksCoverRightColumn(List<DataColumn> rightColumns)
        {
            foreach (KeyValuePair<double, DataBlockList> pair in _dicLeftVerTextBlocks)
            {
                DataBlockList blockList = pair.Value;
                for (int i = 0; i < blockList.Count; )
                {
                    if (blockList[i].IsIntersect(rightColumns))
                    {
                        removedBlocks.Add(blockList[i]);
                        blockList.RemoveAt(i);
                    }
                    else
                        i++;
                }
            }
            GenericMethods<double, DataBlockList>.RemoveZeroAmountValueItems(_dicLeftVerTextBlocks);
        }

        void RemoveBlocksOfNotRightColumn(List<DataColumn> textColumns)
        {
            for (int i = 0; i < textColumns.Count; )
            {
                DataColumn column = textColumns[i];
                if (column.ColumnType != DataType.Number)
                {
                    removedBlocks.AddRange(column.TextBlocks);
                    textColumns.RemoveAt(i);
                }
                else
                {
                    i++;
                }
            }
        }

        void RemoveNeitherLeftNorRightBlocks(params SortedDictionary<double, DataBlockList>[] dics)
        {
            if (_dicHorTextBlocks.Count <= 1)
                return;

            bool isFirst = true;
            foreach (SortedDictionary<double, DataBlockList> dic in dics)
            {
                foreach (KeyValuePair<double, DataBlockList> pair in dic)
                {
                    double key = pair.Key;
                    DataBlockList blockList = pair.Value;
                    for (int i = 0; i < blockList.Count; )
                    {
                        if (!blockList[i].IsLeft && !blockList[i].IsRight)
                        {
                            if (isFirst)
                            {
                                removedBlocks.Add(blockList[i]);
                            }
                            blockList.RemoveAt(i);
                        }
                        else
                            i++;
                    }
                }
                GenericMethods<double, DataBlockList>.RemoveZeroAmountValueItems(dic);
                isFirst = false;
            }
        }

        void ProcessElements(ElementReader reader)
        {
            _processedElements = new List<CustomElement>();
            Element element;
            while ((element = reader.Next()) != null)
            {
                Element.Type elementType = element.GetType();
                switch (elementType)
                {
                    case Element.Type.e_text:
                        ProcessText(element);
                        break;
                    case Element.Type.e_text_begin:
                        ProcessElements(reader);
                        break;
                    case Element.Type.e_text_end:
                        return;
                    case Element.Type.e_form:
                        reader.FormBegin();
                        ProcessElements(reader);
                        reader.End();
                        break;
                }
                reader.ClearChangeList();
            }
        }

        void ProcessText(Element element)
        {
            CustomElement customElement = element.GenerateCustomElement(_defaultMatrix);

            if (_processedElements.Exists(item => customElement.IsOverlap(item)))
            {
                return;
            }
            _processedElements.Add(customElement);

            //Transform element bound.
            Rect elementBound = element.GetBBoxAfterMatrixTranslate(_defaultMatrix);
            double[] horBound = element.GetHorBound(_defaultMatrix);
            elementBound.x1 = horBound[0];
            elementBound.x2 = horBound[0] + horBound[1];
            //Not in the scale
            if (!RectRegion.Contains((elementBound.x1 + elementBound.x2) / 2,
                (elementBound.y1 + elementBound.y2) / 2)) return;
            //Check text is null string.
            if (string.IsNullOrEmpty(element.GetTextString().Trim()))
                return;
            AddTextBlocks(elementBound, element);
        }

        void AddTextBlocks(Rect elementBound, Element element)
        {
            double y1 = elementBound.y1;
            double height = elementBound.Height();

            //Get or create key-value pair for _dicHorTextBlocks
            DataBlockList textBlocks;

            KeyValuePair<HorBlocksInfo, DataBlockList> keyValuePair;
            bool isExist = IsExistPair(_dicHorTextBlocks, out keyValuePair, y1);
            HorBlocksInfo horBlocksInfo = isExist ? keyValuePair.Key : null;
            if (!isExist)
            {
                horBlocksInfo = new HorBlocksInfo
                {
                    BottomBound = y1,
                    TopBound = y1 + height
                };
                textBlocks = new DataBlockList();
                _dicHorTextBlocks.Add(horBlocksInfo, textBlocks);
            }
            else
            {
                if (y1 < horBlocksInfo.BottomBound)
                    horBlocksInfo.BottomBound = y1;
                if (horBlocksInfo.TopBound < y1 + height)
                    horBlocksInfo.TopBound = y1 + height;

                textBlocks = _dicHorTextBlocks[horBlocksInfo];
            }
            //Add textblock object to textblock list
            AddAndMergeBlocks(horBlocksInfo, textBlocks, element);
        }

        void AddAndMergeBlocks(HorBlocksInfo horBlocksInfo, DataBlockList originalBlocks, Element element)
        {
            List<DataElementPart> parts = DataElementPart.GetElementParts(element, _defaultMatrix);
            int index = 0;
            List<char> lastConnectChars = new List<char> { '、', '(' };
            List<char> nextConnectChars = new List<char> { '、', ')', '.' };

            foreach (DataElementPart elementPart in parts)
            {
                if (originalBlocks.Count > 0)
                {
                    DataBlock lastBlock = originalBlocks.Last();


                    if ((elementPart.LeftBound >= lastBlock.LeftBound && elementPart.LeftBound <= lastBlock.RightBound)
                        || IsSameAxisValue(lastBlock.RightBound, elementPart.LeftBound, true)
                        || lastConnectChars.Exists(ch => lastBlock.Text.EndsWith(ch.ToString()))
                        || nextConnectChars.Exists(ch => elementPart.Text.StartsWith(ch.ToString())))
                    {
                        lastBlock.RightBound = elementPart.RightBound;
                        lastBlock.Text += elementPart.Text;
                        index++;
                        continue;
                    }
                }

                //Add block to Horizontial dictionary
                DataBlock newBlock = new DataBlock
                {
                    LeftBound = elementPart.LeftBound,
                    RightBound = elementPart.RightBound,
                    BottomBound = horBlocksInfo.BottomBound,
                    TopBound = horBlocksInfo.TopBound,
                    Text = elementPart.Text,
                    HorBlockDicKey = horBlocksInfo
                };
                originalBlocks.Add(newBlock);
                index++;
            }
        }

        void GenerateLeftRightDicBlocks()
        {
            _dicLeftVerTextBlocks = new SortedDictionary<double, DataBlockList>();
            _dicRightVerTextBlocks = new SortedDictionary<double, DataBlockList>();
            _dicHorTextBlocks.Values.ToList().
                ForEach(blocks => blocks.ForEach(
                    block =>
                    {
                        AddVerBlocks(block, true);
                        AddVerBlocks(block, false);
                    }));
            SortLeftRightBlocks();
        }

        bool IsExistPair(SortedDictionary<double, DataBlockList> dictionary,
            out KeyValuePair<double, DataBlockList> pair, double key, bool isHorizontial)
        {
            List<CustomPair<double, KeyValuePair<double, DataBlockList>>> pairs =
                new List<CustomPair<double, KeyValuePair<double, DataBlockList>>>();
            foreach (KeyValuePair<double, DataBlockList> _pair in dictionary)
            {
                if (IsSameAxisValue(_pair.Key, key, isHorizontial))
                {
                    double newKey = Math.Abs(_pair.Key - key);
                    pairs.Add(new CustomPair<double, KeyValuePair<double, DataBlockList>>
                    {
                        Key = newKey,
                        Value = _pair
                    });
                }
            }
            if (pairs.Count > 0)
            {
                pairs.Sort();
                pair = pairs[0].Value;
                return true;
            }
            pair = new KeyValuePair<double, DataBlockList>();
            return false;
        }

        bool IsExistPair(SortedDictionary<HorBlocksInfo, DataBlockList> dictionary,
            out KeyValuePair<HorBlocksInfo, DataBlockList> pair, double key)
        {
            List<CustomPair<double, KeyValuePair<HorBlocksInfo, DataBlockList>>> pairs =
                new List<CustomPair<double, KeyValuePair<HorBlocksInfo, DataBlockList>>>();
            foreach (KeyValuePair<HorBlocksInfo, DataBlockList> _pair in dictionary)
            {
                if (IsSameAxisValue(_pair.Key.BottomBound, key, false))
                {
                    pairs.Add(new CustomPair<double, KeyValuePair<HorBlocksInfo, DataBlockList>>
                    {
                        Key = Math.Abs(_pair.Key.BottomBound - key),
                        Value = _pair
                    });
                }
            }
            if (pairs.Count > 0)
            {
                pairs.Sort();
                pair = pairs[0].Value;
                return true;
            }
            pair = new KeyValuePair<HorBlocksInfo, DataBlockList>();
            return false;
        }

        void RemoveVerBlocks(DataBlock findBlock, bool isLeft)
        {
            SortedDictionary<double, DataBlockList> blockDictionary = isLeft ? _dicLeftVerTextBlocks : _dicRightVerTextBlocks;
            double key = isLeft ? findBlock.LeftBlockDicKey : findBlock.RightBlockDicKey;

            DataBlockList textBlocks = blockDictionary[key];
            textBlocks.Remove(findBlock);
            if (textBlocks.Count == 0)
            {
                blockDictionary.Remove(key);
            }
        }

        void AddVerBlocks(DataBlock newBlock, bool isLeft)
        {
            SortedDictionary<double, DataBlockList> dicBlocks =
                isLeft ? _dicLeftVerTextBlocks : _dicRightVerTextBlocks;
            double newBoundValue = isLeft ? newBlock.LeftBound : newBlock.RightBound;
            //Organize by right bound
            KeyValuePair<double, DataBlockList> keyValuePair;
            DataBlockList verBlocks;
            double axisXValue;
            bool isExist = IsExistPair(dicBlocks, out keyValuePair, newBoundValue, true);
            if (!isExist)
            {
                axisXValue = newBoundValue;
                newBlock.SetBlockDicKey(axisXValue, isLeft);
                verBlocks = new DataBlockList { newBlock };
                dicBlocks.Add(axisXValue, verBlocks);
            }
            else
            {
                axisXValue = keyValuePair.Key;
                newBlock.SetBlockDicKey(axisXValue, isLeft);
                verBlocks = keyValuePair.Value;
                verBlocks.Add(newBlock);
                if (newBoundValue > axisXValue)
                {
                    dicBlocks.Remove(axisXValue);
                    double updateKey = (newBoundValue + axisXValue) / 2;
                    if (!dicBlocks.ContainsKey(updateKey))
                    {
                        dicBlocks.Add(updateKey, verBlocks);
                    }
                    else
                    {
                        dicBlocks[updateKey].AddRange(verBlocks);
                    }
                    dicBlocks[updateKey].SetBlockDicKey(updateKey, isLeft);
                }
            }
        }

        bool IsSameAxisValue(double value1, double value2, bool isHorizontial)
        {
            double floatValue = isHorizontial ? horizontal_Text_SeparationDistance : vertical_Text_SeparationDistance;
            return Math.Abs(value1 - value2) <= floatValue;
        }

        #endregion

    }
}