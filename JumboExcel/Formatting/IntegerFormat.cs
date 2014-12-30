using System.Collections.Generic;

namespace JumboExcel.Formatting
{
    public class IntegerFormat : CommonValueFormat
    {
        internal IntegerFormat(int id, string formatCode) : base(id, formatCode)
        {
        }

        /// <summary>
        /// Predefined number format <c>0</c>.
        /// </summary>
        public static readonly IntegerFormat GeneralValue = new IntegerFormat(1, "0");

        /// <summary>
        /// Predefined number format <c>#,##0</c>.
        /// </summary>
        public static readonly IntegerFormat IntegerWithSeparator = new IntegerFormat(3, "#,##0");

        public static IEnumerable<IntegerFormat> GetIntegerFormsts()
        {
            return new[]
            {
                GeneralValue,
                IntegerWithSeparator
            };
        }
    }
}
