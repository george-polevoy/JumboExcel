namespace JumboExcel.Formatting
{
    public class NumberFormat : CommonValueFormat
    {
        /// <summary>
        /// Predefined number format <c>General</c>.
        /// </summary>
        internal static readonly IntegerFormat Default = new IntegerFormat(0, "");

        /// <summary>
        /// Constructor for number format.
        /// </summary>
        /// <param name="id">Id of the common number format.</param>
        /// <param name="formatCode"></param>
        internal NumberFormat(int id, string formatCode) : base(id, formatCode)
        {
        }

        public static NumberFormat FromFormatString(string formatCode)
        {
            return new NumberFormat(-1, formatCode);
        }
    }
}