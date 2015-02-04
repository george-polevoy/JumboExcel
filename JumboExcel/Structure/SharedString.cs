using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents shared string cell. Use this class to represent lengthy strings, that a likely to appear more then once in the document.
    /// For string that represent short values, and are likely to appear only a few times, use <see cref="InlineString"/>
    /// This representation uses internal shared string string table structure, which saves memory in case of repeated strings.
    /// </summary>
    /// <remarks>http://stackoverflow.com/questions/6468783/what-is-the-difference-between-cellvalues-inlinestring-and-cellvalues-string-in</remarks>
    public sealed class SharedString : CellElement
    {
        /// <summary>
        /// Value.
        /// </summary>
        private readonly string value;

        /// <summary>
        /// Style.
        /// </summary>
        public StringStyle Style { get; private set; }

        public SharedString(string value, StringStyle style = default(StringStyle))
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
