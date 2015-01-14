using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents inline string. Use this class to represent cells with uncommon values. For large values that are likely to repeat among the worksheets, use <see cref="SharedStringElement"/>.
    /// </summary>
    /// <remarks>http://stackoverflow.com/questions/6468783/what-is-the-difference-between-cellvalues-inlinestring-and-cellvalues-string-in</remarks>
    public class InlineStringElement : CellElement
    {
        /// <summary>
        /// Value.
        /// </summary>
        private readonly string value;

        /// <summary>
        /// Style.
        /// </summary>
        public StringStyleDefinition Style { get; set; }

        public InlineStringElement(string value, StringStyleDefinition style = default(StringStyleDefinition))
        {
            this.value = value;
            Style = style;
        }

        /// <summary>
        /// Value.
        /// </summary>
        public string Value { get { return value; } }

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
