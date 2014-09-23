using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Ric.Core;
using Ric.Util;

namespace Ric.Tasks
{
    public class HKShared
    {
        //Update the worksheet format to fit some requiements: 
        //
        //1. Delete A column;
        //2. Delete D column;
        //3. Insert blank column after F column;
        //4. Insert black row before 11 row.

        public static void UpdateWorksheetFormat(Worksheet worksheet)
        {
            if (worksheet == null)
            {
                throw new Exception(String.Format("Cannot Find the worksheet : {0}", worksheet));
            }
            RangeDeleteCols(1, 1, worksheet);
            RangeDeleteCols(1, 4, worksheet);
            Range insertCols = ExcelUtil.GetRange(1,7,worksheet);
            ExcelUtil.InsertBlankCols(insertCols,1);
            Range insertRows = ExcelUtil.GetRange(11,1,worksheet);
            ExcelUtil.InsertBlankRows(insertRows, 1);
            //RangeDeleteCols(1, 9, worksheet);
        }

        public static void RangeDeleteCols(int row, int col, Worksheet ws) 
        {
            Range cols = ExcelUtil.GetRange(row, col, ws);
            DeleteCols(cols,col);
        }

        public static void DeleteCols(Range where, int colCount) 
        {
            if (colCount-- > 0) 
            {
                where.EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
            }            
        }
    }
}
