using System.Collections.Generic;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents row group, which can be collapsed.
    /// </summary>
    public class RowGroupElement : RowLevelElement
    {
        /// <summary>
        /// Child row level elements.
        /// </summary>
        private readonly IEnumerable<RowLevelElement> rowElements;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="rowElements">Child row level elements.</param>
        public RowGroupElement(IEnumerable<RowLevelElement> rowElements)
        {
            this.rowElements = rowElements;
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="rowElements">Child row level elements.</param>
        public RowGroupElement(params RowLevelElement[] rowElements)
        {
            this.rowElements = rowElements;
        }

        /// <summary>
        /// Child row level elements.
        /// </summary>
        public IEnumerable<RowLevelElement> RowElements { get { return rowElements; } }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}