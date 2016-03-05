using System;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents position of the cell.
    /// </summary>
    internal sealed class CellRef : IEquatable<CellRef>
    {
        /// <summary>
        /// Constructs the cell reference from the row and column zero based indices.
        /// </summary>
        /// <param name="row">Zero based row index.</param>
        /// <param name="column">Zero based column index.</param>
        public CellRef(int row, int column)
        {
            Row = row;
            Column = column;
        }

        /// <summary>
        /// Zero based row index.
        /// </summary>
        public int Row { get; private set; }

        /// <summary>
        /// Zero based column index.
        /// </summary>
        public int Column { get; private set; }

        public override string ToString()
        {
            return string.Format("Row: {0}, Column: {1}", Row, Column);
        }

        public bool Equals(CellRef other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Row == other.Row && Column == other.Column;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            return obj is CellRef && Equals((CellRef)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (Row * 397) ^ Column;
            }
        }

        public static bool operator ==(CellRef left, CellRef right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(CellRef left, CellRef right)
        {
            return !Equals(left, right);
        }
    }
}