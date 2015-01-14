namespace JumboExcel.Formatting
{
    /// <summary>
    /// Represents number format.
    /// </summary>
    public class NumberFormat : CommonValueFormat
    {
        /// <summary>
        /// Predefined number format <c>""</c>. Renders integer numbers as decimal, and fractional numbers as decimal with unspecified decimal places after the separator.
        /// </summary>
        internal static readonly NumberFormat Default = new NumberFormat(0, "");

        /// <summary>
        /// Constructor for number format.
        /// </summary>
        /// <param name="id">Id of the common number format.</param>
        /// <param name="formatCode"></param>
        internal NumberFormat(int id, string formatCode) : base(id, formatCode)
        {
        }

        /// <summary>
        /// Creates instance for provided Excel format.
        /// </summary>
        /// <param name="formatCode"></param>
        /// <returns></returns>
        public static NumberFormat FromFormatString(string formatCode)
        {
            return new NumberFormat(-1, formatCode);
        }
    }
}