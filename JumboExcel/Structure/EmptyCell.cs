namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents an empty cell. Yielding an instance of this class as one of the cells in a <see cref="Row"/>, effectively skips one cell.
    /// </summary>
    public sealed class EmptyCell : CellElement
    {
        private static readonly EmptyCell instance = new EmptyCell();

        public static EmptyCell Instance { get { return instance; } }

        private EmptyCell()
        {
        }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.VisitEmptyCell();
        }
    }
}
