using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents an inline string. For strings that represent short values, and are likely to appear only a few times.
    /// For large values that are likely to repeat among the worksheets, use <see cref="SharedString"/>.
    /// </summary>
    /// <remarks>http://stackoverflow.com/questions/6468783/what-is-the-difference-between-cellvalues-inlinestring-and-cellvalues-string-in</remarks>
    public sealed class InlineString : CellElement
    {
        /// <summary>
        /// Value.
        /// </summary>
        private readonly string value;

        /// <summary>
        /// Style.
        /// </summary>
        public StringStyle Style { get; private set; }

        public InlineString(string value, StringStyle style = default(StringStyle))
        {
            this.value = value;
            Style = style;
        }

        /// <summary>
        /// Value.
        /// </summary>
        public string Value { get { return value; } }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }

        public override string ToString()
        {
            return string.Format("Value: {0}, Style: {1}", value, Style);
        }
    }
}
