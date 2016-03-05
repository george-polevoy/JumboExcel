using System.Drawing;
using JumboExcel.Formatting;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Number style.
    /// </summary>
    public struct NumberStyle
    {
        /// <summary>
        /// Shared cell style.
        /// </summary>
        internal readonly CellStyle cellStyle;

        /// <summary>
        /// Constructor.
        /// </summary>
        public NumberStyle(NumberFormat format, Font font = null, Border border = Border.NONE, Color? fillColor = default (Color?), Alignment alignment = null)
        {
            cellStyle = new CellStyle(font, border, fillColor, (format ?? NumberFormat.Default).FormatCode, alignment);
        }
    }
}