using JumboExcel.Styling;

namespace JumboExcel.Structure
{
    public abstract class ValueCell<TData> : CellElement where TData : struct
    {
        public TData? Value { get; private set; }

        protected ValueCell(TData? value)
        {
            Value = value;
        }
    }

    public abstract class NumberCell<TData> : ValueCell<TData> where TData : struct
    {
        public override string ToString()
        {
            return string.Format("Value: {0}, Style: {1}", Value, Style);
        }

        public NumberStyleDefinition Style { get; private set; }

        protected NumberCell(TData? number, NumberStyleDefinition style) : base(number)
        {
            Style = style;
        }
    }
}
