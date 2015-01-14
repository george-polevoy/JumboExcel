using System.Drawing;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Boolean style.
    /// </summary>
    public struct BooleanStyleDefinition
    {
        /// <summary>
        /// Shared cell style.
        /// </summary>
        internal readonly CellStyleDefinition CellStyleDefinition;

        /// <summary>
        /// Constructor.
        /// </summary>
        public BooleanStyleDefinition(FontDefinition fontDefinition, BorderDefinition borderDefinition, Color? fillColor = default(Color?))
        {
            CellStyleDefinition = new CellStyleDefinition(fontDefinition, borderDefinition, fillColor);
        }
    }
}