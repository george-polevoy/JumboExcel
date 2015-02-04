using System;
using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents a cell, holding a DateTime value.
    /// </summary>
    public sealed class DateTimeCell : ValueCell<DateTime>
    {
        /// <summary>
        /// Style for the cell.
        /// </summary>
        public DateStyle Style { get; private set; }

        /// <summary>
        /// Constructor, taking nullable dateTime and style.
        /// </summary>
        /// <param name="dateTime">Value to display in the cell.</param>
        /// <param name="style">Optional Style.</param>
        public DateTimeCell(DateTime? dateTime, DateStyle style = default(DateStyle)) : base(dateTime)
        {
            Style = style;
        }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
