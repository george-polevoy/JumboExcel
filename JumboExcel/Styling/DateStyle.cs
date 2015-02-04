using System.Drawing;
using JumboExcel.Formatting;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Date style.
    /// </summary>
    public struct DateStyle
    {
        /// <summary>
        /// Shared cell style.
        /// </summary>
        internal readonly CellStyle cellStyle;

        /// <summary>
        /// Constructor.
        /// </summary>
        public DateStyle(DateTimeFormat format, Font font, Border border, Color? fillColor = default (Color?))
        {
            cellStyle = new CellStyle(font, border, fillColor, (format ?? DateTimeFormat.DateDMmm).FormatCode);
        }
    }
}