using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace Ric.Util
{
    public class ZipUtil
    {

        //Zip file
        public static bool ZipFile(string[] filePathArr, string zipFilePath, out string err)
        {
            err = string.Empty;
            if (filePathArr == null || filePathArr.Length == 0)
            {
                err = "There's no file needed to be zipped.";
                return false;
            }

            if(!Directory.Exists(Path.GetDirectoryName(zipFilePath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(zipFilePath));
            }

            try
            {
                using (ZipOutputStream s = new ZipOutputStream(File.Create(zipFilePath)))
                {
                    s.SetLevel(9);
                    byte[] buffer = new byte[4096];
                    foreach (string file in filePathArr)
                    {
                        ZipEntry entry = new ZipEntry(Path.GetFileName(file));
                        entry.DateTime = DateTime.Now;
                        s.PutNextEntry(entry);
                        using (FileStream fs = File.OpenRead(file))
                        {
                            int sourceBytes;
                            do
                            {
                                sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                s.Write(buffer, 0, sourceBytes);
                            } while (sourceBytes > 0);
                        }
                    }
                    s.Finish();
                    s.Close();
                }
            }

            catch (Exception ex)
            {
                err = ex.Message;
                return false;
            }
            return true;
        }

        //Unzip file
        public static bool UnZipFile(string zipFilePath, string unZipDir, out string err)
        {
            err = string.Empty;
            if (!File.Exists(zipFilePath))
            {
                err = zipFilePath + " doesn't exist";
                return false;
            }

            if (unZipDir == null || unZipDir == string.Empty)
            {
                unZipDir = Path.GetDirectoryName(zipFilePath);
            }

            if (!Directory.Exists(unZipDir))
            {
                Directory.CreateDirectory(unZipDir);
            }

            try
            {
                using (ZipInputStream s = new ZipInputStream(File.OpenRead(zipFilePath)))
                {
                    ZipEntry theEntry;
                    while ((theEntry = s.GetNextEntry()) != null)
                    {
                        string dirName = Path.GetDirectoryName(theEntry.Name);
                        string fileName = Path.GetFileName(theEntry.Name);
                        if (dirName.Length > 0)
                        {
                            Directory.CreateDirectory(unZipDir + "\\" + dirName);
                        }
                        if (!dirName.EndsWith("//"))
                            dirName += "//";
                        if (fileName != string.Empty)
                        {
                            using (FileStream sw = File.Create(unZipDir + "\\" + theEntry.Name))
                            {
                                int size = 2048;
                                byte[] data = new byte[2048];
                                while (true)
                                {
                                    size = s.Read(data, 0, data.Length);
                                    if (size > 0)
                                    {
                                        sw.Write(data, 0, size);
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                err = ex.Message;
                return false;
            }

            return true;
        }

        //Unzip file with default path
        public static bool UnzipFile(string zipFilePath, out string err)
        {
            return UnZipFile(zipFilePath, string.Empty, out err);
        }
    }
}
