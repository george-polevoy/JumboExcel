using System;
using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents a cell, holding a DateTime value.
    /// </summary>
    public sealed class DateTimeCellElement : ValueCell<DateTime>
    {
        /// <summary>
        /// Style for the cell.
        /// </summary>
        public DateStyleDefinition Style { get; private set; }

        /// <summary>
        /// Constructor, taking nullable dateTime and style.
        /// </summary>
        /// <param name="dateTime">Value to display in the cell.</param>
        /// <param name="style">Optional Style.</param>
        public DateTimeCellElement(DateTime dateTime, DateStyleDefinition style = default(DateStyleDefinition)) : base(dateTime)
        {
            Style = style;
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
