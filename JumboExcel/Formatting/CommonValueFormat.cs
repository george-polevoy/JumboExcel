using System;
using System.Collections.Generic;

namespace JumboExcel.Formatting
{
    public class CommonValueFormat : IEquatable<CommonValueFormat>
    {
        private readonly int id;
        private readonly string formatCode;

        internal int Id { get { return id; } }

        internal string FormatCode { get { return formatCode; } }

        internal CommonValueFormat(int id, string formatCode)
        {
            this.id = id;
            this.formatCode = formatCode;
        }

        /// <summary>
        /// Predefined number format <c>General</c>.
        /// </summary>
        public static readonly IntegerFormat General = new IntegerFormat(0, "General");

        /// <summary>
        /// Predefined number format <c>#,##0 ;(#,##0)</c>.
        /// </summary>
        public static readonly CommonValueFormat Undefined37 = new CommonValueFormat(37, "#,##0 ;(#,##0)");

        /// <summary>
        /// Predefined number format <c>#,##0 ;[Red](#,##0)</c>.
        /// </summary>
        public static readonly CommonValueFormat Undefined38 = new CommonValueFormat(38, "#,##0 ;[Red](#,##0)");

        /// <summary>
        /// Predefined number format <c>#,##0.00;(#,##0.00)</c>.
        /// </summary>
        public static readonly CommonValueFormat Undefined39 = new CommonValueFormat(39, "#,##0.00;(#,##0.00)");

        /// <summary>
        /// Predefined number format <c>#,##0.00;[Red](#,##0.00)</c>.
        /// </summary>
        public static readonly CommonValueFormat Undefined40 = new CommonValueFormat(40, "#,##0.00;[Red](#,##0.00)");

        /// <summary>
        /// Predefined number format <c>@</c>.
        /// </summary>
        public static readonly CommonValueFormat String = new CommonValueFormat(49, "@");

        internal static IEnumerable<CommonValueFormat> GetFormats()
        {
            return new[] {
                General,
                IntegerFormat.GeneralValue,
                DecimalFormat.FractionalTwoDecimalPlaces,
                IntegerFormat.IntegerWithSeparator,
                DecimalFormat.SeparatorTwoDecimalPlaces,
                DecimalFormat.IntegerPercents,
                DecimalFormat.PercentsTwoDecimalPlaces,
                DecimalFormat.ValueWithExponent1, DecimalFormat.FractionWithDenominator, DecimalFormat.FractionWithDenominatorPrecise,
                DateTimeFormat.DateMmDdYy,
                DateTimeFormat.DateDMmmYy,
                DateTimeFormat.DateDMmm,
                DateTimeFormat.DateMmmYy,
                DateTimeFormat.TimeAmPm,
                DateTimeFormat.TimeAmPmWithSeconds,
                DateTimeFormat.Time24,
                DateTimeFormat.Time24WithSeconds,
                DateTimeFormat.DateTime,
                Undefined37,
                Undefined38,
                Undefined39,
                Undefined40,
                DateTimeFormat.TimeMmSs,
                DateTimeFormat.TimeHMmSs,
                DateTimeFormat.TimeMmSs0,
                DecimalFormat.ValueWithExponent2,
                String,
            };
        }

        public bool Equals(CommonValueFormat other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return id == other.id;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((CommonValueFormat)obj);
        }

        public override int GetHashCode()
        {
            return id;
        }
    }
}
