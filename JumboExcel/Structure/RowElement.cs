using System.Collections.Generic;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents single row.
    /// </summary>
    public class RowElement : RowLevelElement
    {
        /// <summary>
        /// Child cell elements.
        /// </summary>
        private readonly IEnumerable<CellElement> cellElements;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="cellElements">Child cell elements.</param>
        public RowElement(IEnumerable<CellElement> cellElements)
        {
            this.cellElements = cellElements;
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="cellElements">Child cell elements.</param>
        public RowElement(params CellElement[] cellElements)
        {
            this.cellElements = cellElements;
        }

        /// <summary>
        /// Child cell elements.
        /// </summary>
        public IEnumerable<CellElement> CellElements
        {
            get { return cellElements; }
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
