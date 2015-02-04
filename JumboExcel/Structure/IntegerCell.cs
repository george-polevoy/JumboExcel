using System;
using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents nullable cell.
    /// </summary>
    public sealed class IntegerCell : NumberCell<Int64>
    {
        /// <summary>
        /// Constructor, taking a nullable value and style definition.
        /// </summary>
        /// <param name="number">Value or null.</param>
        /// <param name="style">Style.</param>
        public IntegerCell(Int64? number, NumberStyle style = default(NumberStyle))
            : base(number, style)
        {
        }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
