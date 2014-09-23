using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Ric.Util
{
    public class LinqUtil
    {
        public static string GetHeader(string path)
        {
            using (var sr = File.OpenText(path))
            {
                if (!sr.EndOfStream)
                {
                    return sr.ReadLine();
                }
            }
            return null;
        }

        public static IEnumerable<string> GetLines(string path, bool skipHeader)
        {
            using (var sr = File.OpenText(path))
            {
                if (skipHeader && !sr.EndOfStream)
                {
                    sr.ReadLine();
                }
                while (!sr.EndOfStream)
                {
                    yield return sr.ReadLine();
                }
            }
        }

        public static void Dump<T>(IEnumerable<T> results, string path)
        {
            Dump(results, path, "\t");
        }

        public static void Dump<T>(IEnumerable<T> results, string path, string delimiter)
        {
            var piList = typeof(T).GetProperties().ToList();
            var header = string.Join(delimiter, piList.Select(x => x.Name).ToArray());
            using (var sw = new StreamWriter(path, false))
            {
                sw.WriteLine(header);
                foreach (var result in results)
                {
                    var line = string.Join(delimiter, piList.Select(x => x.GetValue(result, null).ToString()).ToArray());
                    sw.WriteLine(line);
                }
            }
        }
    }
}