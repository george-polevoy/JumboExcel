using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents shared string. Use this class to represent lengthy strings, that a likely to appear more then once in the document. For string values that a likely to be unique, use <see cref="InlineStringElement"/>
    /// This representation uses internal shared string string table structure, which saves memory in case of repeated strings.
    /// </summary>
    /// <remarks>http://stackoverflow.com/questions/6468783/what-is-the-difference-between-cellvalues-inlinestring-and-cellvalues-string-in</remarks>
    public class SharedStringElement : CellElement
    {
        /// <summary>
        /// Value.
        /// </summary>
        private readonly string value;

        public SharedStringElement(string value, StringStyleDefinition style = default(StringStyleDefinition))
        {
            this.value = value;
            Style = style;
        }

        public string Value { get { return value; } }

        /// <summary>
        /// Style.
        /// </summary>
        public StringStyleDefinition Style { get; private set; }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }

        public override string ToString()
        {
            return string.Format("Value: {0}, Style: {1}", value, Style);
        }
    }
}
