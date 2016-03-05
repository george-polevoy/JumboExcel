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
        /// Style.
        /// </summary>
        internal readonly CellStyle cellStyle;

        /// <summary>
        /// Constructor.
        /// </summary>
        public DateStyle(DateTimeFormat format, Font font, Border border = Border.NONE, Color? fillColor = default (Color?), Alignment alignment = null)
        {
            cellStyle = new CellStyle(font, border, fillColor, (format ?? DateTimeFormat.DateDMmm).FormatCode, alignment);
        }
    }
}