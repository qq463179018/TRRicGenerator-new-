using System.Collections.Generic;
using Ric.FileLib.Entry;
using Ric.FileLib.Enum;
using Ric.FormatLib;

namespace Ric.FileLib
{
    public class Fm : AFile
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
        public Fm(Format format, FileMode mode)
        {
            Initialize(format, mode);
        }

        public Fm(Format format)
        {
            Initialize(format, FileMode.ReadWrite);
        }

        public Fm(FileMode mode)
        {
            Initialize(Format.Vertical, mode);
        }

        public Fm()
        {
            Initialize(Format.Vertical, FileMode.ReadWrite);
        }

        #endregion

        #region Initialization

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
