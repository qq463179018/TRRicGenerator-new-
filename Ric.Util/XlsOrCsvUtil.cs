using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Ric.Util
{
    public class XlsOrCsvUtil
    {
        public static bool GenerateXls0rCsv(string path, Dictionary<string, List<string>> dicList)
        {
            if (dicList == null || dicList.Count <= 1)//title must exist while no data in file
            {
                string msg = string.Format("no data need to generate");
                throw new Exception(msg);
            }

            using (ExcelApp app = new ExcelApp(false, false))
            {
                try
                {
                    Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(app, path);
                    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                    FillExcel(wSheet, dicList);
                    app.ExcelAppInstance.AlertBeforeOverwriting = false;
                    wBook.Save();
                    return true;
                }
                catch (Exception ex)
                {
                    string msg = string.Format("generate XlsOrCsv file error ,msg:{0}", ex.ToString());
                    throw new Exception(msg);
                }
            }
        }

        public static bool GenerateXls0rCsv(string path, List<List<string>> listList)
        {
            if (listList == null || listList.Count <= 1)//title must exist while no data in file
            {
                string msg = string.Format("no data need to generate");
                throw new Exception(msg);
            }

            using (ExcelApp app = new ExcelApp(false, false))
            {
                try
                {
                    Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(app, path);
                    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                    FillExcel(wSheet, listList);
                    app.ExcelAppInstance.AlertBeforeOverwriting = false;
                    wBook.Save();
                    return true;
                }
                catch (Exception ex)
                {
                    string msg = string.Format("generate XlsOrCsv file error ,msg:{0}", ex.ToString());
                    throw new Exception(msg);
                }
            }
        }

        private static void FillExcel(Worksheet wSheet, List<List<string>> listList)
        {
            SetTitle(wSheet, listList[0]);

            for (int i = 1; i < listList.Count; i++)
            {
                for (int j = 0; j < listList[i].Count; j++)
                {
                    ((Range)wSheet.Cells[i + 1, j + 1]).NumberFormatLocal = "@";
                    wSheet.Cells[i + 1, j + 1] = listList[i][j];
                }
            }
        }

        private static void FillExcel(Worksheet wSheet, Dictionary<string, List<string>> dic)
        {
            int rowCount = dic.Count;
            SetTitle(wSheet, dic.Values.ToList()[0]);

            for (int i = 1; i < rowCount; i++)
            {
                var list = dic.Values.ToList()[i];
                for (int j = 0; j < list.Count; j++)
                {
                    ((Range)wSheet.Cells[i + 1, j + 1]).NumberFormatLocal = "@";
                    wSheet.Cells[i + 1, j + 1] = list[j];
                }
            }
        }

        private static void SetTitle(Worksheet wSheet, List<string> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                ((Range)wSheet.Columns[ToName(i), System.Type.Missing]).ColumnWidth = 20;
                wSheet.Cells[1, i + 1] = list[i];
            }

            //((Range)wSheet.Columns["A:" + ToName(list.Count - 1), System.Type.Missing]).Font.Name = "Arail";//set style of XlsOrCsv
            //((Range)wSheet.Rows[1, Type.Missing]).Font.Bold = System.Drawing.FontStyle.Bold;
            //((Range)wSheet.Rows[1, Type.Missing]).Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
        }

        public static string ToName(int index)
        {
            if (index < 0)
                throw new Exception("invalid parameter");

            List<string> chars = new List<string>();

            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + (int)'A')).ToString());
                index = (int)((index - index % 26) / 26);
            }
            while (index > 0);

            return String.Join(string.Empty, chars.ToArray());
        }

        public static bool GenerateStringCsv(string path, Dictionary<string, List<string>> dicList)
        {
            string result = string.Empty;
            if (dicList == null || dicList.Count <= 1)//title must exist while no data in file
            {
                string msg = string.Format("no data need to generate");
                throw new Exception(msg);
            }

            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (var item in dicList.Values.ToList())
                {
                    foreach (var str in item)
                        sb.AppendFormat("{0},", str.Replace(",", ""));

                    sb.Length = sb.Length - 1;
                    sb.Append("\r\n");
                }

                File.WriteAllText(path, sb.ToString());
                return true;
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate StringCsv file error ,msg:{0}", ex.ToString());
                throw new Exception(msg);
            }
        }

        public static bool GenerateStringCsv(string path, List<List<string>> listList)
        {
            string result = string.Empty;
            if (listList == null || listList.Count <= 1)//title must exist while no data in file
            {
                string msg = string.Format("no data need to generate");
                throw new Exception(msg);
            }

            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (var item in listList)
                {
                    foreach (var str in item)
                        sb.AppendFormat("{0},", str.Replace(",", ""));

                    sb.Length = sb.Length - 1;
                    sb.Append("\r\n");
                }

                File.WriteAllText(path, sb.ToString());
                return true;
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate StringCsv file error ,msg:{0}", ex.ToString());
                throw new Exception(msg);
            }
        }



        //use template to generate List<List<string>>
        public static bool GenerateXls0rCsv(string path, List<object> entities, List<string> title)
        {
            if (entities == null || entities.Count <= 0)
            {
                string msg = string.Format("no data need to generate");
                throw new Exception(msg);
            }

            try
            {
                //Workbook wBook = ExcelUtil.CreateOrOpenExcelFile(app, path);
                //Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                //FillExcel(wSheet, listList);
                //app.ExcelAppInstance.AlertBeforeOverwriting = false;
                //wBook.Save();

                /*
                CODE_2,反射获取对象数据

                    StringBuilder commandText = new StringBuilder(" insert into ");
                    Type type = obj.GetType();
                    string tableName = type.Name;//表名称
                    PropertyInfo[] pros = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);//所有字段名称
                    StringBuilder fieldStr = new StringBuilder();//拼接需要插入数据库的字段
                    StringBuilder paramStr = new StringBuilder();//拼接每个字段对应的参数
                    int len = pros.Length;
                    if (!"".Equals(identityName) && null != identityName) param = new SqlParameter[len-1];//如果有自动增长的字段,则该字段不需要SqlParameter
                    int paramLIndex = 0;
                    for (int i = 0; i < len; i++)
                    {
  　　                string fieldName = pros[i].Name;
                      if (!fieldName.ToUpper().Equals(identityName.ToUpper()))
                         {
                            //非自动增长字段才加入SQL语句
    　　　　                fieldStr.Append(fieldName);
                            string paramName = "@" + fieldName;//SQL语句的字段名称和参数名称保持一致
                            paramStr.Append(paramName);
                            if (i < (len - 1))
                               {
          　　                   fieldStr.Append(",");//参数和字段用逗号隔开
                                 paramStr.Append(",");
                               }
                           object val = type.GetProperty(fieldName).GetValue(obj, null);// 根据属性名称获取当前属性的值
                           if (val == null) 
                                val = DBNull.Value;//如果该值为空的话,则将其转化为数据库的NULL
                           param[paramLIndex] = new SqlParameter(fieldName, val);//给每个参数赋值
                           paramLIndex++;
                          }
                        }
                 **/

                return true;
            }
            catch (Exception ex)
            {
                string msg = string.Format("generate XlsOrCsv file error ,msg:{0}", ex.ToString());
                throw new Exception(msg);
            }
        }
    }
}
