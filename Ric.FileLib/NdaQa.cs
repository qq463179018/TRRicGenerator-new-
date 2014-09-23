using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ric.FileLib.Enum;
using Ric.FormatLib;

namespace Ric.FileLib
{
    /// <summary>
    /// NdaQa file reprensentation
    /// </summary>
    public class NdaQa : Nda
    {
        #region Constructors

        /// <summary>
        /// Nda Qa constructor
        /// </summary>
        /// <param name="format">
        /// Format.Horizontal by default
        /// </param>
        /// <param name="mode">
        /// file mode
        /// </param>>
        public NdaQa(Format format, FileMode mode)
        {
            Initialize(format, mode);
        }

        /// <summary>
        /// Nda constructor
        /// </summary>
        /// <param name="format"></param>
        public NdaQa(Format format)
        {
            Initialize(format, FileMode.ReadWrite);
        }

        /// <summary>
        /// Nda constructor
        /// </summary>
        /// <param name="mode"></param>
        public NdaQa(FileMode mode)
        {
            Initialize(Format.Horizontal, mode);
        }

        /// <summary>
        /// Nda constructor
        /// </summary>
        public NdaQa()
        {
            Initialize(Format.Horizontal, FileMode.ReadWrite);
        }

        #endregion

        #region Rules

        #endregion

    }
}
