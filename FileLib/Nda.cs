using System.Collections.Generic;
using Ric.FileLib.Entry;
using Ric.FormatLib;
using FileMode = Ric.FileLib.Enum.FileMode;

namespace Ric.FileLib
{
    /// <summary>
    /// Nda file representation
    /// </summary>
    public class Nda : AFile
    {
        #region Constructor

        /// <summary>
        /// Nda constructor
        /// </summary>
        /// <param name="format">
        /// Format.Horizontal by default
        /// </param>
        /// <param name="mode">
        /// file mode
        /// </param>>
        public Nda(Format format, FileMode mode)
        {
            Initialize(format, mode);
        }

        public Nda(Format format)
        {
            Initialize(format, FileMode.ReadWrite);
        }

        public Nda(FileMode mode)
        {
            Initialize(Format.Horizontal, mode);
        }

        public Nda()
        {
            Initialize(Format.Horizontal, FileMode.ReadWrite);
        }

        #endregion

        #region Initialization

        private void Initialize(Format format, FileMode mode)
        {
            Titles = new List<string>();
            DynamicContent = new List<dynamic>();
            Content = new List<AEntry>();
            EntryType = typeof(NdaEntry);

            ChooseMode(mode);
            ChooseFormat(format);
        }

        #endregion

        #region Rules

        public List<string> testchecktitles(List<string> titles)
        {
            return titles;
        }

        #endregion
    }

}
