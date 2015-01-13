using System.Collections.Generic;

namespace JumboExcel.Formatting
{
    public class DecimalFormat : NumberFormat
    {
        private DecimalFormat(int id, string formatCode) : base(id, formatCode)
        {
        }

        /// <summary>
        /// Predefined number format <c>0.00</c>.
        /// </summary>
        public static readonly DecimalFormat FractionalTwoDecimalPlaces = new DecimalFormat(2, "0.00");

        /// <summary>
        /// Predefined number format <c>#,##0.00</c>.
        /// </summary>
        public static readonly DecimalFormat SeparatorTwoDecimalPlaces = new DecimalFormat(4, "#,##0.00");

        /// <summary>
        /// Predefined number format <c>0%</c>.
        /// </summary>
        public static readonly DecimalFormat IntegerPercents = new DecimalFormat(9, "0%");

        /// <summary>
        /// Predefined number format <c>0.00%</c>.
        /// </summary>
        public static readonly DecimalFormat PercentsTwoDecimalPlaces = new DecimalFormat(10, "0.00%");

        /// <summary>
        /// Predefined number format <c>0.00E+00</c>.
        /// </summary>
        public static readonly DecimalFormat ValueWithExponent1 = new DecimalFormat(11, "0.00E+00");

        /// <summary>
        /// Predefined number format <c>##0.0E+0</c>.
        /// </summary>
        public static readonly DecimalFormat ValueWithExponent2 = new DecimalFormat(48, "##0.0E+0");

        /// <summary>
        /// Predefined number format <c># ?/?</c>.
        /// </summary>
        public static readonly DecimalFormat FractionWithDenominator = new DecimalFormat(12, "# ?/?");

        /// <summary>
        /// Predefined number format <c># ??/??</c>.
        /// </summary>
        public static readonly DecimalFormat FractionWithDenominatorPrecise = new DecimalFormat(13, "# ??/??");

        /// <summary>
        /// Predefined number format <c>#,##0.00;(#,##0.00)</c>.
        /// </summary>
        /// <remarks></remarks>
        public static readonly DecimalFormat AccountingAmount = new DecimalFormat(39, "#,##0.00;(#,##0.00)");

        /// <summary>
        /// Predefined number format <c>#,##0.00;[Red](#,##0.00)</c>.
        /// </summary>
        public static readonly DecimalFormat AccountingAmountColored = new DecimalFormat(40, "#,##0.00;[Red](#,##0.00)");

        /// <summary>
        /// Get all decimal formats.
        /// </summary>
        public static IEnumerable<NumberFormat> GetDecimalFormats()
        {
            return new[]
            {
                FractionalTwoDecimalPlaces,
                SeparatorTwoDecimalPlaces,
                IntegerPercents,
                PercentsTwoDecimalPlaces,
                ValueWithExponent1,
                ValueWithExponent2,
                AccountingAmount,
                AccountingAmountColored
            };
        }
    }
}
