using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents shared string.
    /// </summary>
    /// <remarks>http://stackoverflow.com/questions/6468783/what-is-the-difference-between-cellvalues-inlinestring-and-cellvalues-string-in</remarks>
    public class SharedStringElement : CellElement
    {
        public override string ToString()
        {
            return string.Format("Value: {0}, Style: {1}", value, Style);
        }

        private readonly string value;
        public SharedStringStyleDefinition Style { get; private set; }
        public string Value { get { return value; } }

        public SharedStringElement(string value, SharedStringStyleDefinition style = null)
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
