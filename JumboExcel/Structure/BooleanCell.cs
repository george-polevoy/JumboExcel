using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents boolean cell.
    /// </summary>
    public sealed class BooleanCell : ValueCell<bool>
    {
        /// <summary>
        /// Style for the cell.
        /// </summary>
        public BooleanStyle Style { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="value">Value.</param>
        /// <param name="style">Style.</param>
        public BooleanCell(bool? value, BooleanStyle style = default (BooleanStyle)) : base(value)
        {
            Style = style;
        }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
