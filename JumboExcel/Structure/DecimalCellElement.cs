using System;
using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents a cell, holding a decimal value.
    /// </summary>
    public sealed class DecimalCellElement : NumberCell<decimal>
    {
        public DecimalCellElement(Decimal? number, NumberStyleDefinition style = default(NumberStyleDefinition)) : base(number, style)
        {
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
