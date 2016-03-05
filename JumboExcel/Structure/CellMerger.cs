namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents a part of a merged cell.
    /// </summary>
    public abstract class CellMerger : CellElement
    {
        /// <summary>
        /// Constructs a part of a merged cell.
        /// </summary>
        /// <param name="innerElement">
        /// Cell element that will be output. If the resulting cell is offset from the anchor cell then only the style of this element will be used.
        /// </param>
        protected CellMerger(CellElement innerElement)
        {
            InnerElement = innerElement;
        }

        /// <summary>
        /// Gets the column's zero based index of the anchor cell.
        /// </summary>
        /// <param name="currentColumn">Column index of the element as it's renderd.</param>
        /// <remarks>The instance itself does not contain it's position in document, because an instance could be reused in different coordinates,
        /// so the only way to know the anchor coordinates is to compute it during rendering.</remarks>
        public abstract int GetAnchorColumn(int currentColumn);

        /// <summary>
        /// Gets the row's zero based index of the anchor cell.
        /// </summary>
        /// <param name="currentRow">Row index of the element as it's renderd.</param>
        /// <remarks>The instance itself does not contain it's position in document, because an instance could be reused in different coordinates,
        /// so the only way to know the anchor coordinates is to compute it during rendering.</remarks>
        public abstract int GetAnchorRow(int currentRow);

        /// <summary>
        /// The actual cell element that will be output.
        /// The contents will be lost in case the cell is actually merging to the other cell, in this case only the cell's style will be used.
        /// </summary>
        public CellElement InnerElement { get; private set; }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}