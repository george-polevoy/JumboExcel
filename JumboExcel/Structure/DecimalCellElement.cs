using System;
using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    public sealed class DecimalCellElement : NumberCell<decimal>
    {
        public DecimalCellElement(Decimal? value, NumberStyleDefinition style = null) : base(value, style)
        {
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
