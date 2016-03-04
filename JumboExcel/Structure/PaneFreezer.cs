using System;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Designates the frozen panes.
    /// </summary>
    public class PaneFreezer
    {
        /// <summary>
        /// Constructs a pane freezer element used to designate frozen panes.
        /// A zero or negative index of row/column means that the pane will not be generated in this direction.
        /// </summary>
        /// <param name="rowIndex">Zero based row index of top left corner of the scrolling region.</param>
        /// <param name="columnIndex">Zero based column index of top left corner of the scrolling region.</param>
        public PaneFreezer(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0)
                throw new ArgumentOutOfRangeException("rowIndex", rowIndex, "Must be non-negative.");

            if (columnIndex < 0)
                throw new ArgumentOutOfRangeException("columnIndex", columnIndex, "Must be non-negative.");

            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        /// <summary>
        /// Zero based row index of top left corner of the scrolling region.
        /// </summary>
        public int RowIndex { get; private set; }

        /// <summary>
        /// Zero based column index of top left corner of the scrolling region.
        /// </summary>
        public int ColumnIndex { get; private set; }
    }
}