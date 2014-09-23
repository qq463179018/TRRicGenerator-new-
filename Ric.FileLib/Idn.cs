using System.Collections.Generic;
using Ric.FileLib.Entry;
using Ric.FileLib.Enum;
using Ric.FormatLib;

namespace Ric.FileLib
{
    public class Idn : AFile
    {
        #region Constructor

        /// <summary>
        /// Idn constructor
        /// </summary>
        /// <param name="format">
        /// Format.Horizontal by default
        /// </param>
        /// <param name="mode">
        /// file mode
        /// </param>>
        public Idn(Format format, FileMode mode)
        {
            Initialize(format, mode);
        }

        /// <summary>
        /// Idn constructor
        /// </summary>
        /// <param name="format"></param>
        public Idn(Format format)
        {
            Initialize(format, FileMode.ReadWrite);
        }

        /// <summary>
        /// Idn constructor
        /// </summary>
        /// <param name="mode"></param>
        public Idn(FileMode mode)
        {
            Initialize(Format.Horizontal, mode);
        }

        /// <summary>
        /// Idn constructor
        /// </summary>
        public Idn()
        {
            Initialize(Format.Horizontal, FileMode.ReadWrite);
        }

        #endregion

        #region Initialization

        /// <summary>
        /// Initialization function for Idn files
        /// </summary>
        /// <param name="format"></param>
        /// <param name="mode"></param>
        private void Initialize(Format format, FileMode mode)
        {
            Titles = new List<string>();
            DynamicContent = new List<dynamic>();
            Content = new List<AEntry>();
            EntryType = typeof(IdnEntry);

            ChooseMode(mode);
            ChooseFormat(format);
        }

        #endregion
    }
}
