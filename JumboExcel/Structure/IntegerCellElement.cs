using System;
using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents nullable cell.
    /// </summary>
    public sealed class IntegerCellElement : NumberCell<Int64>
    {
        /// <summary>
        /// Constructor, taking a nullable value and style definition.
        /// </summary>
        /// <param name="number">Value or null.</param>
        /// <param name="style">Style.</param>
        public IntegerCellElement(Int64? number, NumberStyleDefinition style = null)
            : base(number, style)
        {
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
