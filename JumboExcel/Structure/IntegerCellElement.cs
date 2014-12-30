using System;
using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    public class IntegerCellElement : NumberCell<Int64>
    {
        public IntegerCellElement(Int64? value, NumberStyleDefinition style = null)
            : base(value, style)
        {
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
