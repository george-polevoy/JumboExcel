using System;
using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents a cell, holding a decimal value.
    /// </summary>
    public sealed class DecimalCell : NumberCell<decimal>
    {
        public DecimalCell(Decimal? number, NumberStyle style = default(NumberStyle)) : base(number, style)
        {
        }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
