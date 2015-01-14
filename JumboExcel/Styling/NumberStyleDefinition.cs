using System.Drawing;
using JumboExcel.Formatting;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Number style.
    /// </summary>
    public struct NumberStyleDefinition
    {
        /// <summary>
        /// Shared cell style.
        /// </summary>
        internal readonly CellStyleDefinition CellStyleDefinition;

        /// <summary>
        /// Constructor.
        /// </summary>
        public NumberStyleDefinition(NumberFormat format, FontDefinition fontDefinition, BorderDefinition borderDefinition, Color? fillColor = default (Color?))
        {
            CellStyleDefinition = new CellStyleDefinition(fontDefinition, borderDefinition, fillColor, (format ?? NumberFormat.Default).FormatCode);
        }
    }
}