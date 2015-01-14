using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents boolean cell.
    /// </summary>
    public sealed class BooleanCellElement : ValueCell<bool>
    {
        /// <summary>
        /// Style for the cell.
        /// </summary>
        public BooleanStyleDefinition Style { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="value">Value.</param>
        /// <param name="style">Style.</param>
        public BooleanCellElement(bool? value, BooleanStyleDefinition style = default (BooleanStyleDefinition)) : base(value)
        {
            Style = style;
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
