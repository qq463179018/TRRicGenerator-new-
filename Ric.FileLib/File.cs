using System.Collections.Generic;
using Ric.FileLib.Entry;
using Ric.FileLib.Enum;
using Ric.FormatLib;
using Ric.FileLib.Exception;

namespace Ric.FileLib
{
    /// <summary>
    /// File type representation
    /// file is for any file possible.
    /// It doesn't have any rules
    /// Entry is null by default so if you don't add your
    /// own type only dynamic entry are available
    /// </summary>
    public class File : AFile
    {
        #region Constructor

        /// <summary>
        /// File constructor
        /// </summary>
        /// <param name="format">
        /// Format.Horizontal by default
        /// </param>
        /// <param name="mode">
        /// file mode
        /// </param>>
        public File(Format format, FileMode mode)
        {
            Initialize(format, mode);
        }

        /// <summary>
        /// File constructor
        /// </summary>
        /// <param name="format"></param>
        public File(Format format)
        {
            Initialize(format, FileMode.ReadWrite);
        }

        /// <summary>
        /// File constructor
        /// </summary>
        /// <param name="mode"></param>
        public File(FileMode mode)
        {
            Initialize(Format.Vertical, mode);
        }

        /// <summary>
        /// File constructor
        /// </summary>
        public File()
        {
            Initialize(Format.Vertical, FileMode.ReadWrite);
        }

        #endregion

        #region Initialization
        
        /// <summary>
        /// Initialization function for files
        /// </summary>
        /// <param name="format"></param>
        /// <param name="mode"></param>
        private void Initialize(Format format, FileMode mode)
        {
            Titles = new List<string>();
            DynamicContent = new List<dynamic>();
            Content = new List<AEntry>();
            EntryType = null;

            ChooseMode(mode);
            ChooseFormat(format);
        }

        #endregion

        #region Overriding saving 

        /// <summary>
        /// Overriding the Save function
        /// </summary>
        public override void Save()
        {
            if (Mode == FileMode.ReadOnly)
            {
                throw new FileLibException("File is Read Only, cannot save");
            }

            FileFormat.Save(Path);
        }

        /// <summary>
        /// Overriding the Save function
        /// </summary>
        /// <param name="path"></param>
        public override void Save(string path)
        {
            if (Mode == FileMode.ReadOnly)
            {
                throw new FileLibException("File is Read Only, cannot save");
            }

            FileFormat.Save(path);
        }

        #endregion

        #region Rules

        // This kind of file doesn't have any rule by default

        #endregion
    }
}
