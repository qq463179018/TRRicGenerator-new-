using System;

namespace Ric.Core.Events
{
    public class ResultEventArgs : EventArgs
    {
        public string FileName { get; set; }
        public string Filetype { get; set; }
        public string FilePath { get; set; } 

        public ResultEventArgs(string filename, string filepath, string filetype)
        {
            FileName = filename;
            Filetype = filetype;
            FilePath = filepath;
        }

    }
}