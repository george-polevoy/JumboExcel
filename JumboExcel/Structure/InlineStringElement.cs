using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents inline string.
    /// </summary>
    /// <remarks>http://stackoverflow.com/questions/6468783/what-is-the-difference-between-cellvalues-inlinestring-and-cellvalues-string-in</remarks>
    public class InlineStringElement : CellElement
    {
        public override string ToString()
        {
            return string.Format("Value: {0}, Style: {1}", value, Style);
        }

        private readonly string value;
        public StringStyleDefinition Style { get; set; }
        public string Value { get { return value; } }

        public InlineStringElement(string value, StringStyleDefinition style = default(StringStyleDefinition))
        {
            this.value = value;
            Style = style;
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
