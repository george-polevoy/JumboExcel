using System;
using System.Text;

namespace JumboExcel
{
    /// <summary>
    /// Auxiliary, Excel specific functions.
    /// </summary>
    static class ExcelHelper
    {
        /// <summary>
        /// Formats the cell address for Excel.
        /// </summary>
        /// <param name="row">Zero based row index.</param>
        /// <param name="column">Zero based column index.</param>
        /// <returns>Returns the cell's address in Excel format.</returns>
        public static string CellRef(int row, int column)
        {
            var charBuffer = new StringBuilder(10);
            var c = column + 1;
            do
            {
                int digit;
                c = Math.DivRem(c - 1, 26, out digit);
                charBuffer.Insert(0, (char)('A' + digit));
            } while (c > 0);
            charBuffer.Append(row + 1);
            return charBuffer.ToString();
        }
    }
}
