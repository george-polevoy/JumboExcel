using System.Collections.Generic;

namespace JumboExcel.Formatting
{
    /// <summary>
    /// Represents DateTime format.
    /// </summary>
    public sealed class DateTimeFormat : CommonValueFormat
    {
        public DateTimeFormat(string format) : base(-1, format)
        {
        }

        private DateTimeFormat(int id, string formatCode) : base(id, formatCode)
        {
        }

        /// <summary>
        /// Predefined number format <c>mm-dd-yy</c>.
        /// </summary>
        public static readonly DateTimeFormat DateMmDdYy = new DateTimeFormat(14, "mm-dd-yy");

        /// <summary>
        /// Predefined number format <c>d-mmm-yy</c>.
        /// </summary>
        public static readonly DateTimeFormat DateDMmmYy = new DateTimeFormat(15, "d-mmm-yy");

        /// <summary>
        /// Predefined number format <c>d-mmm</c>.
        /// </summary>
        public static readonly DateTimeFormat DateDMmm = new DateTimeFormat(16, "d-mmm");

        /// <summary>
        /// Predefined number format <c>mmm-yy</c>.
        /// </summary>
        public static readonly DateTimeFormat DateMmmYy = new DateTimeFormat(17, "mmm-yy");

        /// <summary>
        /// Predefined number format <c>h:mm AM/PM</c>.
        /// </summary>
        public static readonly DateTimeFormat TimeAmPm = new DateTimeFormat(18, "h:mm AM/PM");

        /// <summary>
        /// Predefined number format <c>h:mm:ss AM/PM</c>.
        /// </summary>
        public static readonly DateTimeFormat TimeAmPmWithSeconds = new DateTimeFormat(19, "h:mm:ss AM/PM");

        /// <summary>
        /// Predefined number format <c>h:mm</c>.
        /// </summary>
        public static readonly DateTimeFormat Time24 = new DateTimeFormat(20, "H:mm");

        /// <summary>
        /// Predefined number format <c>h:mm:ss</c>.
        /// </summary>
        public static readonly DateTimeFormat Time24WithSeconds = new DateTimeFormat(21, "H:mm:ss");

        /// <summary>
        /// Predefined number format <c>m/d/yy h:mm</c>.
        /// </summary>
        public static readonly DateTimeFormat DateTime = new DateTimeFormat(22, "m/d/yy H:mm");

        /// <summary>
        /// Predefined number format <c>mm:ss</c>.
        /// </summary>
        public static readonly DateTimeFormat TimeMmSs = new DateTimeFormat(45, "mm:ss");

        /// <summary>
        /// Predefined number format <c>[h]:mm:ss</c>.
        /// </summary>
        public static readonly DateTimeFormat TimeHMmSs = new DateTimeFormat(46, "[h]:mm:ss");

        /// <summary>
        /// Predefined number format <c>mmss.0</c>.
        /// </summary>
        public static readonly DateTimeFormat TimeMmSs0 = new DateTimeFormat(47, "mmss.0");

        public static IEnumerable<DateTimeFormat> GetDateTimeFormats()
        {
            return new[]
            {
                DateMmDdYy,
                DateDMmmYy,
                DateDMmm,
                DateMmmYy,
                TimeAmPm,
                TimeAmPmWithSeconds,
                Time24,
                Time24WithSeconds,
                DateTime,
                TimeMmSs,
                TimeHMmSs,
                TimeMmSs0
            };
        }
    }
}
