using System;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Various flags affecting worksheet compatibility.
    /// </summary>
    [Flags]
    public enum WorksheetCompatibilityFlags
    {
        /// <summary>
        /// </summary>
        NONE = 0,
        
        /// <summary>
        /// Allows creating worksheet without restriction on name length, present in previous versions.
        /// If not set, worksheet length is restricted to 31 characters.
        /// Not set by default.
        /// </summary>
        RELAX_WORKSHEET_LENGTH_CONSTRAINT = 1 << 0,
    }
}