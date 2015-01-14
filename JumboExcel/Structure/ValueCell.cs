using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents cells, holding a value.
    /// </summary>
    /// <typeparam name="TData">Type of the value.</typeparam>
    public abstract class ValueCell<TData> : CellElement where TData : struct
    {
        /// <summary>
        ///  Value.
        /// </summary>
        public TData? Value { get; private set; }

        protected ValueCell(TData? value)
        {
            Value = value;
        }
    }

    /// <summary>
    /// Represents cells, holding a number value.
    /// </summary>
    /// <typeparam name="TData"></typeparam>
    public abstract class NumberCell<TData> : ValueCell<TData> where TData : struct
    {
        /// <summary>
        /// Style.
        /// </summary>
        public NumberStyleDefinition Style { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="number">Number.</param>
        /// <param name="style">Style.</param>
        protected NumberCell(TData? number, NumberStyleDefinition style) : base(number)
        {
            Style = style;
        }

        public override string ToString()
        {
            return string.Format("Value: {0}, Style: {1}", Value, Style);
        }
    }
}
