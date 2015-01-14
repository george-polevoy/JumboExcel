namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents empty cell.
    /// </summary>
    public sealed class EmptyCellElement : CellElement
    {
        private static readonly EmptyCellElement instance = new EmptyCellElement();

        public static EmptyCellElement Instance { get { return instance; } }

        private EmptyCellElement()
        {
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.VisitEmptyCell();
        }
    }
}
