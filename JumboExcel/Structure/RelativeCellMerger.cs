using System;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents a part of a merged cell.
    /// </summary>
    public class RelativeCellMerger : CellMerger
    {
        /// <summary>
        /// Constructs a part of a merged cell.
        /// </summary>
        /// <param name="innerElement">Cell element which will be actually written.
        /// If the resulting cell is offset from the anchor cell, then only the style of this element will be used.</param>
        /// <param name="rowOffset">Positive row offset from the anchor cell.</param>
        /// <param name="columnOffset">Positive column offset from the anchor cell.</param>
        public RelativeCellMerger(CellElement innerElement, int rowOffset, int columnOffset)
            : base(innerElement)
        {
            if (rowOffset < 0)
                throw new ArgumentOutOfRangeException("rowOffset", rowOffset, "Must be non-negative.");

            if (columnOffset < 0)
                throw new ArgumentOutOfRangeException("columnOffset", rowOffset, "Must be non-negative.");

            this.rowOffset = rowOffset;
            this.columnOffset = columnOffset;
        }

        /// <summary>
        /// Positive row offset from the anchor cell.
        /// </summary>
        readonly int rowOffset;

        /// <summary>
        /// Positive column offset from the anchor cell.
        /// </summary>
        readonly int columnOffset;

        public override int GetAnchorColumn(int currentColumn)
        {
            return currentColumn - columnOffset;
        }

        public override int GetAnchorRow(int currentRow)
        {
            return currentRow - rowOffset;
        }
    }
}