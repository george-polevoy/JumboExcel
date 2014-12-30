using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    public class InlineStringElement : CellElement
    {
        public override string ToString()
        {
            return string.Format("Value: {0}, Style: {1}", value, Style);
        }

        private readonly string value;
        public StringStyleDefinition Style { get; set; }
        public string Value { get { return value; } }

        public InlineStringElement(string value, StringStyleDefinition style = null)
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
