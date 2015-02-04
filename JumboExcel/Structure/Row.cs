using System.Collections.Generic;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents single row.
    /// </summary>
    public sealed class Row : RowLevelElement
    {
        /// <summary>
        /// Child cell elements.
        /// </summary>
        private readonly IEnumerable<CellElement> cells;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="cells">Child cell elements.</param>
        public Row(IEnumerable<CellElement> cells)
        {
            this.cells = cells;
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="cells">Child cell elements.</param>
        public Row(params CellElement[] cells)
        {
            this.cells = cells;
        }

        /// <summary>
        /// Child cell elements.
        /// </summary>
        public IEnumerable<CellElement> Cells
        {
            get { return cells; }
        }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
