using System.Drawing;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Boolean style.
    /// </summary>
    public struct BooleanStyle
    {
        /// <summary>
        /// Shared cell style.
        /// </summary>
        internal readonly CellStyle cellStyle;

        /// <summary>
        /// Constructor.
        /// </summary>
        public BooleanStyle(Font font, Border border, Color? fillColor = default(Color?))
        {
            cellStyle = new CellStyle(font, border, fillColor);
        }
    }
}