using System.Drawing;
using JumboExcel.Formatting;

namespace JumboExcel.Styling
{
    /// <summary>
    /// String style.
    /// </summary>
    public struct StringStyle
    {
        /// <summary>
        /// Shared cell style.
        /// </summary>
        internal readonly CellStyle cellStyle;

        /// <summary>
        /// Constructor.
        /// </summary>
        public StringStyle(Font font, Border border, Color? fillColor = default (Color?))
        {
            cellStyle =  new CellStyle(font, border, fillColor, CommonValueFormat.String.FormatCode);
        }
    }
}