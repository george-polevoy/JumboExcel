using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    public class EmptyCellElement : CellElement
    {
        public StringStyleDefinition StyleDefinition { get; set; }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.VisitEmptyCell();
        }

        public EmptyCellElement(StringStyleDefinition styleDefinition = null)
        {
            StyleDefinition = styleDefinition;
        }
    }
}
