using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace PdfTronWrapper
{
    public class FreeTable : List<FreeTableRow>
    {
        public string TableName { get; set; }

        public object Tag { get; set; }

        public int FindRowIndex(Regex regex, bool last = false)
        {
            return last ?
                this.FindLastIndex(row => row.Exists(cell => regex.IsMatch(cell.Value))) :
                this.FindIndex(row => row.Exists(cell => regex.IsMatch(cell.Value)));
        }

        /// <summary>
        /// convert to datatable;
        /// this function will work fine if you confirm the table is a regular table
        /// </summary>
        /// <param name="dataTable"></param>
        public DataTable ConvertToDataTable()
        {
            if (this.Count == 0) return null;

            int colCount = this.Max(n => n.Count);

            DataTable dataTable = new DataTable();

            dataTable.TableName = TableName;

            for (int i = 0; i < colCount; i++)
            {
                dataTable.Columns.Add();
            }

            foreach (FreeTableRow webRow in this)
            {
                var index = 0;

                var empty = true;

                var dataRow = dataTable.NewRow();

                foreach (var cell in webRow)
                {
                    if (empty && !string.IsNullOrEmpty(cell.Value.Trim()))
                    {
                        empty = false;
                    }

                    dataRow[index++] = cell.Value;
                }

                if (!empty)
                {
                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }
    }
}
