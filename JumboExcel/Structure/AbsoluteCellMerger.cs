using System;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Merges cells based on absolute coordinates of the ancor cell.
    /// </summary>
    public class AbsoluteCellMerger : CellMerger
    {
        /// <summary>
        /// Constructs the merger accepting the content element and the absolute coodrinates of the anchor cell.
        /// </summary>
        /// <param name="innerElement">Cell element which will be actually written.
        /// If the resulting cell is offset from the anchor cell, then only the style of this element will be used.</param>
        /// <param name="anchorRow">Zero based index of the anchor cell's row.</param>
        /// <param name="anchorColumn">Zero based index of the anchor cell's column.</param>
        /// <remarks>It's important to provide the coorectly styled <paramref name="innerElement"/>
        /// for styling the borders of the resulting merged cell on the edges of the merged cell.</remarks>
        public AbsoluteCellMerger(CellElement innerElement, int anchorRow, int anchorColumn)
            : base(innerElement)
        {
            if (innerElement == null)
                throw new ArgumentNullException("innerElement");
            this.anchorColumn = anchorColumn;
            this.anchorRow = anchorRow;
        }

        /// <summary>
        /// Zero based index of the anchor cell's row.
        /// </summary>
        readonly int anchorRow;

        /// <summary>
        /// Zero based index of the anchor cell's column.
        /// </summary>
        readonly int anchorColumn;

        public override int GetAnchorColumn(int currentColumn)
        {
            return anchorColumn;
        }

        public override int GetAnchorRow(int currentRow)
        {
            return anchorRow;
        }
    }
}
