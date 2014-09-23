using System.Collections.Generic;
using Ric.FileLib.Entry;
using Ric.FileLib.Enum;
using Ric.FileLib.Exception;
using Ric.FormatLib;

namespace Ric.FileLib
{
    /// <summary>
    /// File Maintenance type representation
    /// Can be used with vertical or horizontal format
    /// </summary>
    public class Fm : AFile
    {
        #region Constructor

        /// <summary>
        /// Fm constructor
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

        /// <summary>
        /// Fm constructor
        /// </summary>
        /// <param name="format"></param>
        public Fm(Format format)
        {
            Initialize(format, FileMode.ReadWrite);
        }

        /// <summary>
        /// Fm constructor
        /// </summary>
        /// <param name="mode"></param>
        public Fm(FileMode mode)
        {
            Initialize(Format.Vertical, mode);
        }

        /// <summary>
        /// Fm constructor
        /// </summary>
        public Fm()
        {
            Initialize(Format.Vertical, FileMode.WriteOnly);
        }

        #endregion

        #region Initialization

        /// <summary>
        /// Initialization function for Fm files
        /// </summary>
        /// <param name="format"></param>
        /// <param name="mode"></param>
        private void Initialize(Format format, FileMode mode)
        {
            if (format == Format.Vertical && mode != FileMode.WriteOnly)
            {
                throw new FileLibException("Vertical Fm does not support other mode than WriteOnly");
                //ToDo support ReadWrite for Vertical Fm (if find a nice way)
            }
            Titles = new List<string>();
            DynamicContent = new List<dynamic>();
            Content = new List<AEntry>();
            EntryType = typeof(FmEntry);

            ChooseMode(mode);
            ChooseFormat(format);
        }

        #endregion

        #region Rules

        #endregion

    }
}
