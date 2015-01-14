using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    public sealed class BooleanCellElement : ValueCell<bool>
    {
        /// <summary>
        /// Style for the cell.
        /// </summary>
        public BooleanStyleDefinition Style { get; private set; }

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