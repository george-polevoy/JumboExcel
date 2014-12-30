using JumboExcel.Styling;

namespace JumboExcel.Structure
{
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
