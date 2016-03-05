using System.Collections.Generic;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents row group, which can be collapsed.
    /// </summary>
    public sealed class RowGroup : RowLevelElement
    {
        /// <summary>
        /// Child row level elements.
        /// </summary>
        private readonly IEnumerable<RowLevelElement> rowElements;

        /// <summary>
        /// Display elements as collapsed initially.
        /// </summary>
        public bool Collapse { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="rowElements">Child row level elements.</param>
        /// <param name="collapse">Show child rows as collapsed initially.</param>
        public RowGroup(IEnumerable<RowLevelElement> rowElements, bool collapse = false)
        {
            Collapse = collapse;
            this.rowElements = rowElements;
        }

        /// <summary>
        /// Child row level elements.
        /// </summary>
        public IEnumerable<RowLevelElement> RowElements
        {
            get { return rowElements; }
        }

        internal override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}