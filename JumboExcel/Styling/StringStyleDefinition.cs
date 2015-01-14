using System.Drawing;
using JumboExcel.Formatting;

namespace JumboExcel.Styling
{
    /// <summary>
    /// String style.
    /// </summary>
    public struct StringStyleDefinition
    {
        /// <summary>
        /// Shared cell style.
        /// </summary>
        internal readonly CellStyleDefinition CellStyleDefinition;

        /// <summary>
        /// Constructor.
        /// </summary>
        public StringStyleDefinition(FontDefinition fontDefinition, BorderDefinition borderDefinition, Color? fillColor = default (Color?))
        {
            CellStyleDefinition =  new CellStyleDefinition(fontDefinition, borderDefinition, fillColor, CommonValueFormat.String.FormatCode);
        }
    }
}