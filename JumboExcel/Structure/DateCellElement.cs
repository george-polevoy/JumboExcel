using System;
using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    public sealed class DateCellElement : ValueCell<DateTime>
    {
        public DateStyleDefinition DateStyle { get; private set; }

        public DateCellElement(DateTime value, DateStyleDefinition dateStyle) : base(value)
        {
            DateStyle = dateStyle;
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
