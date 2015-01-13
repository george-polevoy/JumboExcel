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
        /// Predefined number format <c>@</c>.
        /// </summary>
        public static readonly CommonValueFormat String = new CommonValueFormat(49, "@");

        internal static IEnumerable<CommonValueFormat> GetFormats()
        {
            return new[] {
                NumberFormat.Default,
                IntegerFormat.General,
                DecimalFormat.FractionalTwoDecimalPlaces,
                IntegerFormat.IntegerWithSeparator,
                DecimalFormat.SeparatorTwoDecimalPlaces,
                DecimalFormat.IntegerPercents,
                DecimalFormat.PercentsTwoDecimalPlaces,
                DecimalFormat.ValueWithExponent1,
                DecimalFormat.FractionWithDenominator,
                DecimalFormat.FractionWithDenominatorPrecise,
                DateTimeFormat.DateMmDdYy,
                DateTimeFormat.DateDMmmYy,
                DateTimeFormat.DateDMmm,
                DateTimeFormat.DateMmmYy,
                DateTimeFormat.TimeAmPm,
                DateTimeFormat.TimeAmPmWithSeconds,
                DateTimeFormat.Time24,
                DateTimeFormat.Time24WithSeconds,
                DateTimeFormat.DateTime,
                IntegerFormat.AccountingAmount,
                IntegerFormat.AccountingAmountColored,
                DecimalFormat.AccountingAmount,
                DecimalFormat.AccountingAmountColored,
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
