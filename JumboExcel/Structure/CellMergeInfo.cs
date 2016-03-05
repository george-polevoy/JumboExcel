namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents information about the anchor and the merging cells.
    /// </summary>
    internal sealed class CellMergeInfo
    {
        /// <summary>
        /// Upper left corner of the merged cells range.
        /// </summary>
        public CellRef UpperLeft { get; private set; }

        /// <summary>
        /// Lower right corner of the merged cells range.
        /// </summary>
        public CellRef LowerRight { get; private set; }

        /// <summary>
        /// Constructs the merge information from the two cell references.
        /// </summary>
        /// <param name="upperLeft">Upper left cell reference.</param>
        /// <param name="lowerRight">Lower right cell reference.</param>
        public CellMergeInfo(CellRef upperLeft, CellRef lowerRight)
        {
            UpperLeft = upperLeft;
            LowerRight = lowerRight;
        }
    }
}
