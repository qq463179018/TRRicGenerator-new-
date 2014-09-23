using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Globalization;

namespace Ric.Util
{
    public class FileUtil
    {
        public static bool IsFolderExist(string folderPath)
        {
            bool isExist = false;
            if (Directory.Exists(folderPath))
            {
                isExist = true;
            }
            return isExist;
        }

        public static List<string> GetTodayFileNameFromFolder(string folderPath)
        {
            string formatToday = DateTime.Now.ToString("ddMMyyyy", new CultureInfo("en-US"));
            List<string> fileName = new List<string>();

            if (IsFolderExist(folderPath))
            {
                DirectoryInfo warrentFolder = new DirectoryInfo(folderPath);
                foreach (FileInfo file in warrentFolder.GetFiles())
                {
                    string creationDate = file.LastWriteTime.ToString("ddMMyyyy", new CultureInfo("en-US"));
                    if (file.Extension.Equals(".xls") && creationDate.Equals(formatToday))
                    {
                        fileName.Add(file.Name);

                    }
                }
            }
            return fileName;
        }
        public static void  CopyDirectory(string srcDir, string desDir,string[] ignoreFiles )
        {
            string folderName = srcDir.Substring(srcDir.LastIndexOf("\\") + 1);

            string desfolderdir = desDir + "\\" + folderName;

            if (desDir.LastIndexOf("\\") == (desDir.Length - 1))
            {
                desfolderdir = desDir + folderName;
            }
            string[] filenames = Directory.GetFileSystemEntries(srcDir);

            foreach (string file in filenames)// 遍历所有的文件和目录
            {
                if (Directory.Exists(file))// 先当作目录处理如果存在这个目录就递归Copy该目录下面的文件
                {

                    string currentdir = desfolderdir + "\\" + file.Substring(file.LastIndexOf("\\") + 1);
                    if (!Directory.Exists(currentdir))
                    {
                        Directory.CreateDirectory(currentdir);
                    }

                    CopyDirectory(file, desfolderdir,ignoreFiles);
                }

                else // 否则直接copy文件
                {
                    string srcfileName = file.Substring(file.LastIndexOf("\\") + 1);

                    srcfileName = desfolderdir + "\\" + srcfileName;

                    if (!Directory.Exists(desfolderdir))
                    {
                        Directory.CreateDirectory(desfolderdir);
                    }
                    if(!ignoreFiles.Contains(file.Substring(file.LastIndexOf("\\") + 1)))
                    File.Copy(file, srcfileName,true);
                }
            }//foreach
        }//function end


        /// <summary>
        /// For some files, the data need to write after the last record if file exsited, the WriteMode is Append.
        /// For some files, if the file exsits, backup it and create a new file, WriteMode is Overwrite .        
        /// If file is txt, seperate with '\t'
        /// If file is csv, seperate with ','
        /// </summary>
        /// <param name="filepath">file path</param>
        /// <param name="data">data to write</param>
        /// <param name="title">columns' name of file</param>
        /// <param name="filetype">output file type</param>
        public static void WriteOutputFile(string filepath, List<List<string>> data, List<string> title, WriteMode mode)
        {
            if (!File.Exists(filepath))
            {
                string dir = Path.GetDirectoryName(filepath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                if (title != null && title.Count > 0)
                {
                    data.Insert(0, title);
                }
            }
            else
            {
                if (mode.Equals(WriteMode.Overwrite))
                {
                    MiscUtil.BackUpFile(filepath);
                    File.Delete(filepath);
                    if (title != null && title.Count > 0)
                    {
                        data.Insert(0, title);
                    }
                }
            }
            WriteToTXT(filepath, data);
        }

        /// <summary>
        /// Write data to file.
        /// For some files, the data need to write after the last record if file exsited, the WriteMode is Append.
        /// For some files, if the file exsits, backup it and create a new file, WriteMode is Overwrite .        
        /// </summary>
        /// <param name="filepath">file path</param>
        /// <param name="data">data to write, contain seperators(eg. '\t' or ',') in each string, without '\r\n'</param>
        /// <param name="title">columns title</param>
        /// <param name="mode">if overwrite</param>
        public static void WriteOutputFile(string filepath, string[] data, string title, WriteMode mode)
        {
            List<string> content = data.ToList();
            WriteOutputFile(filepath, content, title, mode);
        }

        /// <summary>
        /// Write data to file.
        /// For some files, the data need to write after the last record if file exsited, the WriteMode is Append.
        /// For some files, if the file exsits, backup it and create a new file, WriteMode is Overwrite .        
        /// </summary>
        /// <param name="filepath">file path</param>
        /// <param name="data">data to write, contain seperators(eg. '\t' or ',') in each string, without '\r\n'</param>
        /// <param name="title">columns title</param>
        /// <param name="mode">if overwrite</param>
        public static void WriteOutputFile(string filepath, List<string> data, string title, WriteMode mode)
        {
            if (!File.Exists(filepath))
            {
                string dir = Path.GetDirectoryName(filepath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                if (!string.IsNullOrEmpty(title))
                {
                    data.Insert(0, title);
                }
            }
            else
            {
                if (mode.Equals(WriteMode.Overwrite))
                {
                    MiscUtil.BackUpFile(filepath);
                    File.Delete(filepath);
                    if (!string.IsNullOrEmpty(title))
                    {
                        data.Insert(0, title);
                    }
                }
            }
            WriteToTXT(filepath, data);
        }

        /// <summary>
        /// Write data to file.
        /// If file exsits, it will back up and overwrite it. 
        /// </summary>
        /// <param name="filepath">file path</param>
        /// <param name="fullContent">full content to write to file</param>
        public static void WriteOutputFile(string filepath, string fullContent)
        {
            if (File.Exists(filepath))
            {
                MiscUtil.BackUpFile(filepath);
            }
            else
            {
                string dir = Path.GetDirectoryName(filepath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
            }
            File.WriteAllText(filepath, fullContent, Encoding.UTF8);
        }

        /// <summary>
        /// Write data to txt file.
        /// Open file if it exsits and seeks to the end of the file, or creates a new file.
        /// </summary>
        /// <param name="filepath">file path</param>
        /// <param name="data">data to write</param>       
        protected static void WriteToTXT(string filepath, List<List<string>> data)
        {
            string dir = Path.GetDirectoryName(filepath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            char seperator = '\t';
            if (Path.GetExtension(filepath).ToUpper() == ".CSV")
            {
                seperator = ',';
            }
            FileStream fs = new FileStream(filepath, FileMode.Append);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    sw.Write(data[i][j]);
                    if (j != data[i].Count - 1)
                    {
                        sw.Write(seperator);
                    }
                }
                sw.Write("\r\n");
            }
            sw.Close();
            fs.Close();
        }

        /// <summary>
        /// Write list data to file.
        /// Open file if it exsits and seeks to the end of the file, or creates a new file.
        /// </summary>
        /// <param name="filepath">file path</param>
        /// <param name="data"></param>
        protected static void WriteToTXT(string filepath, List<string> data)
        {
            string dir = Path.GetDirectoryName(filepath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            FileStream fs = new FileStream(filepath, FileMode.Append);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            for (int i = 0; i < data.Count; i++)
            {
                sw.WriteLine(data[i]);
            }
            sw.Close();
            fs.Close();
        }

        /// <summary>
        /// Write single line to file.
        /// If file is ".csv", seperate items with ','.
        /// If file is ".txt", seperate items with '\t'
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="data"></param>
        public static void WriteSingleLine(string filepath, List<string> data)
        {
            WriteSingleLine(filepath, data, null);
        }

        /// <summary>
        /// Write single record to file.
        /// </summary>
        /// <param name="filepath">file path</param>
        /// <param name="record">record to write</param>
        public static void WriteSingleLine(string filepath, string record)
        {
            WriteSingleLine(filepath, record, null);
        }

        /// <summary>
        /// Write a record to file.
        /// If file not exsits, firstly write the column title to file.
        /// else, only write the record.
        /// </summary>
        /// <param name="filepath">file path</param>
        /// <param name="record">record to write</param>
        /// <param name="title">columns' name</param>
        public static void WriteSingleLine(string filepath, string record, string title)
        {
            string content = "";
            if (!File.Exists(filepath))
            {
                string dir = Path.GetDirectoryName(filepath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                if (title != null)
                {
                    content = title + "\r\n";
                }
            }
            content += record + "\r\n";
            File.AppendAllText(filepath, content, Encoding.UTF8);
        }

        /// <summary>
        /// Write one record to file.
        /// If file not exsits, firstly write the column title to file.
        /// else, only write the record.
        /// </summary>
        /// <param name="filepath">file path</param>
        /// <param name="record">record to write</param>
        /// <param name="title">column title</param>
        public static void WriteSingleLine(string filepath, List<string> record, List<string> title)
        {
            string content = "";
            char seperator = '\t';
            if (Path.GetExtension(filepath).ToUpper() == ".CSV")
            {
                seperator = ',';
            }
            if (!File.Exists(filepath))
            {
                string dir = Path.GetDirectoryName(filepath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                if (title != null)
                {
                    content = FormatDataWithSeperator(title, seperator);
                }
            }
            content += FormatDataWithSeperator(record, seperator);
            File.AppendAllText(filepath, content, Encoding.UTF8);
        }

        /// <summary>
        /// Format given data with seperator.
        /// </summary>
        /// <param name="data">data</param>
        /// <param name="seperator">seperator</param>
        /// <returns>formated string</returns>
        public static string FormatDataWithSeperator(List<string> data, char seperator)
        {
            string line = "";
            foreach (string record in data)
            {
                line += record + seperator;
            }
            line = line.Substring(0, line.Length - 1) + "\r\n";
            return line;
        }

        public static void CreateDirectory(string directory)
        {
            if (Path.HasExtension(directory))
            {
                directory = Path.GetDirectoryName(directory);
            }
            if (!Directory.Exists(directory) && directory != null)
            {
                Directory.CreateDirectory(directory);
            }
        }

    }    

    public enum WriteMode
    {
        Overwrite,
        Append
    }

}
