//-----------------------------------------------------------------------------------------------------------------------
//-----
//-----Author:MaShaoming
//-----
//-----Description:Providing the function to locate the data area of a financial table.
//-----
//-----------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using pdftron.PDF;
using pdftron.PDF.Annots;
using PdfTronWrapper.TableBorder;
using PdfTronWrapper.Utility;

namespace PdfTronWrapper
{
    /// <summary>
    /// Get the content scale of the table by pdftron.
    /// </summary>
    public class TableLocator : IDisposable
    {

        #region Private Fields

        /// <summary>
        /// The pdf file path.
        /// </summary>
        private string _pdfFile;

        /// <summary>
        /// Pdf document object.
        /// </summary>
        public PDFDoc pdfDoc;

        /// <summary>
        /// A helper for using pdftron library.
        /// </summary>
        public PdfTronHelper pdfTronHelper { get; private set; }

        private LocateConfiguration _locateConfiguration;

        #endregion

        #region Public Proprities

        public bool IsConsolidated { get; set; }

        public bool ConsolidatedFound { get; set; }

        public bool ParentFound { get; set; }

        private int maxInstance_Between_TableName_And_TabelNameNearbyInfo =120;
        public int MaxInstanceBetweenTableNameAndTabelNameNearbyInfo
        {
            get
            {
                return maxInstance_Between_TableName_And_TabelNameNearbyInfo;
            }
            set
            {
                maxInstance_Between_TableName_And_TabelNameNearbyInfo = value;
            }
        }

        #endregion

        #region Pubic Methods

        public List<TablePos> GetMultiTablePos(string tableName, LocateConfiguration locateConfiguration)
        {
            List<TablePos> tablePoses = new List<TablePos>();
            _locateConfiguration = locateConfiguration;
            LinePos linePos = null;
            List<LinePos> tableNamePoses = new List<LinePos>();
            List<LinePos> tableEndPoses = new List<LinePos>();
            LinePos tableNamePos = null;
            do
            {
                LinePos linePosCopy = linePos == null ? null : (LinePos)linePos.Clone();
                tableNamePos = GetSearchTextResult(tableName, tableName, ref linePos, true);
                if (tableNamePos == null)
                {
                    linePos = linePosCopy;
                    tableNamePos = GetSearchTextResult(tableName.GetFirstLetter( true), tableName, ref linePos, true);
                }
                if (tableNamePos != null)
                {
                    string tableEndRegex = locateConfiguration.TableEndRegex;
                    string tableEndFirstLetterRegex = locateConfiguration.TableEndFirstLetterRegex;
                    linePosCopy = linePos == null ? null : (LinePos)linePos.Clone();
                    LinePos tableEndPos = GetSearchTextResult(tableEndRegex ,tableEndRegex, ref linePos, false);
                    if (tableEndPos == null)
                    {
                        linePos = linePosCopy;
                        tableEndPos = GetSearchTextResult(tableEndFirstLetterRegex, tableEndRegex, ref linePos, false);
                    }
                    if (tableEndPos == null)
                    {
                        tableEndPos = pdfTronHelper.GetBottomPosOfPage(tableNamePos.PageNum, false, -1, -1);
                    }
                    tableNamePoses.Add(tableNamePos);
                    tableEndPoses.Add(tableEndPos);
                }
                else
                {
                    break;
                }
            } while (tableNamePos != null);

            for (int i = 0; i < tableNamePoses.Count; i++)
            {
                List<TablePos> tempPoses = GetTablePoses(tableNamePoses[i], tableEndPoses[i]);
                tablePoses.AddRange(tempPoses);
            }

            return tablePoses;
        }

        public List<TablePos> GetOnlyTablePos(string tableName, LocateConfiguration locateConfiguration, bool isChineseTableName)
        {
            List<TablePos> tablePoses = new List<TablePos>();
            LinePos linePos = null;
            _locateConfiguration = locateConfiguration;

            LinePos tableNamePos = GetSearchTextResult(tableName, tableName, ref linePos, true);
            if (tableNamePos == null)
            {
                linePos = null;
                tableNamePos = GetSearchTextResult(tableName.GetFirstLetter(isChineseTableName), tableName, ref linePos, true);
            }
            if (tableNamePos != null)
            {
                string tableEndRegex = locateConfiguration.TableEndRegex;
                string tableEndFirstLetterRegex = locateConfiguration.TableEndFirstLetterRegex;
                LinePos tableEndPos = GetSearchTextResult(tableEndRegex, tableEndRegex, ref linePos, false);
                if (tableEndPos == null)
                {
                    linePos = null;
                    tableEndPos = GetSearchTextResult(tableEndFirstLetterRegex, tableEndRegex, ref linePos, false);
                }
                if (tableEndPos == null)
                {
                    tableEndPos = pdfTronHelper.GetBottomPosOfPage(tableNamePos.PageNum, false, -1, -1);
                }
                tablePoses = GetTablePoses(tableNamePos, tableEndPos);
            }
            return tablePoses;
        }

        #endregion

        #region Private Methods

        private LinePos GetSearchTextResult(string searchText, string verifyText, ref LinePos lowRange, bool isTableName)
        {
            try
            {
                bool isDone = false;
                int startPage = pdfTronHelper.GetLowPage(lowRange);
                while (!isDone)
                {
                    List<LinePos> linePoses = pdfTronHelper.Search(searchText, lowRange, null, ref startPage);
                    if (linePoses.Count > 0)
                    {
                        foreach (LinePos linePos in linePoses)
                        {
                            if (!Regex.IsMatch(linePos.TrimText, verifyText))
                            {
                                continue;
                            }

                            if (isTableName)
                            {
                                string tableNameNearbyRegex = _locateConfiguration.TableNameNearbyRegex;
                                string tableNameNearbyFirstLetterRegex = _locateConfiguration.TableNameNearbyFirstLetterRegex;
                                LinePos linePosCopy = (LinePos)linePos.Clone();
                                string nearByRegexExpression=tableNameNearbyRegex;
                                LinePos tableNearbyPos = GetSearchTextResult(nearByRegexExpression, nearByRegexExpression, ref linePosCopy, false);
                                if (tableNearbyPos == null)
                                {
                                    LinePos lowRangeCopy2 = (LinePos)linePos.Clone();
                                    tableNearbyPos = GetSearchTextResult(tableNameNearbyFirstLetterRegex, nearByRegexExpression, ref lowRangeCopy2, false);
                                }
                                if (tableNearbyPos == null || !ValidateTableNameNearbyPos(linePos, tableNearbyPos))
                                {
                                    continue;
                                }
                            }

                            lowRange = new LinePos()
                            {
                                PageNum = linePos.PageNum,
                                AxisValue = linePos.AxisValue + GetError(true, linePos.PageNum),
                                TrimText = linePos.TrimText
                            };

                            linePos.AxisValue = linePos.AxisValueWithLineHeight;
                            return linePos;
                        }
                    }
                    isDone = startPage > pdfDoc.GetPageCount();
                }
            }
            catch (Exception ex)
            {
                string err = ex.ToString();
                throw;
            }

            return null;
        }

        private bool ValidateTableNameNearbyPos(LinePos tableNamePos, LinePos nearbyPos)
        {
            return nearbyPos != null && ValidateTableNameUnitInstance(tableNamePos, nearbyPos);
        }

        private bool ValidateTableNameUnitInstance(LinePos tableNamePos, LinePos unitPos)
        {
            return pdfTronHelper.ValidateInstance(tableNamePos, unitPos, maxInstance_Between_TableName_And_TabelNameNearbyInfo);
        }

        private List<TablePos> GetTablePoses(LinePos tableNamePos, LinePos tableEndPos)
        {
            List<TablePos> tablePoses = new List<TablePos>();

            tableNamePos.AxisValue = tableNamePos.AxisValueWithLineHeight;
            //There is only one page
            if (tableEndPos.PageNum == tableNamePos.PageNum)
            {
                TablePos tablePos = GetTablePos(tableNamePos.PageNum, tableNamePos, tableEndPos, false, false, null);
                if (tablePos != null)
                {
                    tablePoses.Add(tablePos);
                }
            }
            //The page amount is over one
            else
            {
                //Get start page
                //Mark
                TablePos startTablePos = null;
                LinePos endPos = pdfTronHelper.GetBottomPosOfPage(tableNamePos.PageNum, false, -1, -1);

                startTablePos = GetTablePos(tableNamePos.PageNum, tableNamePos, endPos, false, false, null);

                if (startTablePos != null)
                {
                    tablePoses.Add(startTablePos);
                }
                bool isStartTablePosNull = startTablePos == null;
                //Get medial page
                for (int i = tableNamePos.PageNum + 1; i < tableEndPos.PageNum; i++)
                {
                    TablePos intervalTablePos = GetTablePos(i, null, pdfTronHelper.GetBottomPosOfPage(i, false, -1, -1),
                        false,
                        !isStartTablePosNull, isStartTablePosNull ? null : startTablePos.VerticalLines);
                    if (intervalTablePos != null)
                    {
                        tablePoses.Add(intervalTablePos);
                    }
                }
                //Get end page
                int endPageNum = tableEndPos.PageNum;
                TablePos endTablePos = GetTablePos(endPageNum, null, tableEndPos, true,
                        !isStartTablePosNull, isStartTablePosNull ? null : startTablePos.VerticalLines);
                if (endTablePos != null)
                {
                    tablePoses.Add(endTablePos);
                }
            }
            return tablePoses;
        }

        private TablePos GetTablePos(int pageNum, LinePos startPos, LinePos endPos, bool isEndTablePos, bool isSubsequentPage, SortedDictionary<double, FormLineList> lastPageVerticalLines)
        {
            FormLineSearcher searcher = new FormLineSearcher(pdfDoc);
            if (endPos != null)
            {
                endPos.AxisValue = endPos.AxisValueWithLineHeight;
            }
            SortedDictionary<double, FormLineList>[] formLines = searcher.GetFormLines(pageNum, startPos, endPos, isSubsequentPage, lastPageVerticalLines);
            if (formLines == null)
                return null;
            TablePos tablePos;

            tablePos = new TablePos
            {
                PageNum = pageNum,
                HorizontialLines = formLines[0],
                VerticalLines = formLines[1]
            };
            //RevertAxisTransform(tablePos);
            return tablePos;
        }

        private double GetError(bool isStart, int pageNum)
        {
            double coefficient = isStart ? -1 : 1;
            double errorPixelNum = 3 * coefficient;
            return errorPixelNum;
        }

        public string BookMark(List<TablePos> tablePoses)
        {
            ClearBookMarks(pdfDoc);
            tablePoses.ForEach(tablePos =>
            CreateTablePosHighlight(tablePos, pdfDoc));
            string path = System.IO.Path.GetDirectoryName(_pdfFile);
            var name = System.IO.Path.GetFileNameWithoutExtension(_pdfFile);
            string BookmarkFilePath = System.IO.Path.Combine(path, name + "_bookmark.pdf");
            pdfDoc.Save(BookmarkFilePath, 0);
            return BookmarkFilePath;
        }

        void CreateTablePosHighlight(TablePos tablePos, PDFDoc pdfDoc)
        {
            double lineWidth = 3;

            Page page = pdfDoc.GetPage(tablePos.PageNum);

            foreach (FormLine line in tablePos.HorizontialLines.Concat(tablePos.VerticalLines).SelectMany(pair => pair.Value))
            {
                lineWidth = line.IsExistent ? 3 : 1;

                Rect rect = line.IsTransverseLine ? new Rect(line.StartPoint.x, line.StartPoint.y, line.EndPoint.x, line.StartPoint.y + lineWidth) :
                    new Rect(line.StartPoint.x, line.StartPoint.y, line.StartPoint.x + lineWidth, line.EndPoint.y);
                PdfTronHelper pdfTronHelper = new PdfTronHelper(pdfDoc);
                pdfTronHelper.RevertTransportRect(tablePos.PageNum, rect);
                CreateHighlight(pdfDoc, page, rect);
            }
        }

        void CreateHighlight(PDFDoc pdfDoc, Page page, Rect rect)
        {
            var highLight = Highlight.Create(pdfDoc, rect);
            highLight.RefreshAppearance();
            page.AnnotPushBack(highLight);
        }

        void ClearBookMarks(PDFDoc pdfDoc)
        {
            Bookmark firstBookMark = pdfDoc.GetFirstBookmark();
            while (firstBookMark != null)
            {
                if (firstBookMark.IsValid())
                {
                    firstBookMark.Delete();
                    firstBookMark = pdfDoc.GetFirstBookmark();
                }
                else
                {
                    firstBookMark = null;
                }
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="pdfFile">The pdf file to be operating</param>
        public TableLocator(string pdfFile)
        {
            _pdfFile = pdfFile;
            pdfDoc = new PDFDoc(_pdfFile);
            pdfDoc.InitSecurityHandler();
            pdfTronHelper = new PdfTronHelper(pdfDoc);
            IsConsolidated = true;
            ConsolidatedFound = true;
            ParentFound = true;
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (pdfDoc != null)
            {
                pdfDoc.Close();
                pdfDoc.Dispose();
            }

            pdfDoc = null;
        }

        #endregion
    }
}
