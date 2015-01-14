using System.Drawing;
using JumboExcel.Formatting;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Date style.
    /// </summary>
    public struct DateStyleDefinition
    {
        /// <summary>
        /// Shared cell style.
        /// </summary>
        internal readonly CellStyleDefinition CellStyleDefinition;

        /// <summary>
        /// Constructor.
        /// </summary>
        public DateStyleDefinition(DateTimeFormat format, FontDefinition fontDefinition, BorderDefinition borderDefinition, Color? fillColor = default (Color?))
        {
            CellStyleDefinition = new CellStyleDefinition(fontDefinition, borderDefinition, fillColor, (format ?? DateTimeFormat.DateDMmm).FormatCode);
        }
    }
}