using System.Collections.Generic;

namespace JumboExcel.Formatting
{
    /// <summary>
    /// Represents integer number format.
    /// </summary>
    public sealed class IntegerFormat : NumberFormat
    {
        private IntegerFormat(int id, string formatCode) : base(id, formatCode)
        {
        }

        /// <summary>
        /// Predefined number format <c>0</c>. Renders integer number part.
        /// </summary>
        public static readonly NumberFormat General = new IntegerFormat(1, "0");

        /// <summary>
        /// Predefined number format <c>#,##0 ;(#,##0)</c>.
        /// </summary>
        /// <remarks>Negative number is in round brackets. Positive number is aligned right and shifted one character left, to accomodate the closing bracket of negative numbers.</remarks>
        public static readonly NumberFormat AccountingAmount = new IntegerFormat(37, "#,##0 ;(#,##0)");

        /// <summary>
        /// Predefined number format <c>#,##0 ;[Red](#,##0)</c>.
        /// </summary>
        /// <remarks>Negative number is in round brackets, colored red. Positive number is aligned right and shifted one character left, to accomodate the closing bracket of negative numbers.</remarks>
        public static readonly NumberFormat AccountingAmountColored = new IntegerFormat(38, "#,##0 ;[Red](#,##0)");

        /// <summary>
        /// Predefined number format <c>#,##0</c>.
        /// </summary>
        public static readonly IntegerFormat IntegerWithSeparator = new IntegerFormat(3, "#,##0");

        /// <summary>
        /// Get all integer formats.
        /// </summary>
        public static IEnumerable<NumberFormat> GetIntegerFormats()
        {
            return new[]
            {
                General,
                IntegerWithSeparator,
                AccountingAmount,
                AccountingAmountColored
            };
        }
    }
}
