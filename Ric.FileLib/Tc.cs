using System.Collections.Generic;
using Ric.FileLib.Entry;
using Ric.FileLib.Enum;
using Ric.FormatLib;

namespace Ric.FileLib
{
    /// <summary>
    /// T&C file representation
    /// </summary>
    public class Tc : AFile
    {
        #region Constructor

        /// <summary>
        /// Tc constructor
        /// </summary>
        /// <param name="format">
        /// Format.Horizontal by default
        /// </param>
        /// <param name="mode">
        /// file mode
        /// </param>>
        public Tc(Format format, FileMode mode)
        {
            Initialize(format, mode);
        }

        /// <summary>
        /// Tc constructor
        /// </summary>
        /// <param name="format"></param>
        public Tc(Format format)
        {
            Initialize(format, FileMode.ReadWrite);
        }

        /// <summary>
        /// Tc constructor
        /// </summary>
        /// <param name="mode"></param>
        public Tc(FileMode mode)
        {
            Initialize(Format.Horizontal, mode);
        }

        /// <summary>
        /// Tc constructor
        /// </summary>
        public Tc()
        {
            Initialize(Format.Horizontal, FileMode.ReadWrite);
        }

        #endregion

        #region Initialization

        /// <summary>
        /// Initialization function for T&C files
        /// </summary>
        /// <param name="format"></param>
        /// <param name="mode"></param>
        private void Initialize(Format format, FileMode mode)
        {
            Titles = new List<string>();
            DynamicContent = new List<dynamic>();
            Content = new List<AEntry>();
            EntryType = typeof(TcEntry);

            ChooseMode(mode);
            ChooseFormat(format);
        }

        #endregion

        #region Rules

        #endregion
    }
}
