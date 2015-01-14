using System;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents column configuration. Allows to specify widths.
    /// </summary>
    public class ColumnElement
    {
        /// <summary>
        /// Specifies the minimum column index to apply this configuration to.
        /// </summary>
        public int Min { get; private set; }

        /// <summary>
        /// Specifies the maximum column index to apply this configuration to.
        /// </summary>
        public int Max { get; private set; }

        /// <summary>
        /// Width, in excel display units to apply to the column range.
        /// </summary>
        public decimal Width { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="min">Specifies the minimum column index to apply this configuration to.</param>
        /// <param name="max">Specifies the maximum column index to apply this configuration to.</param>
        /// <param name="width">Width, in excel display units to apply to the column range.</param>
        public ColumnElement(int min, int max, decimal width)
        {
            if (min < 0)
                throw new ArgumentOutOfRangeException("min", min, "Must be > 0.");

            if (max < min)
                throw new ArgumentOutOfRangeException("max", max, "Must be > min.");

            if (width < 0)
                throw new ArgumentOutOfRangeException("width", width, "Must be > 0.");

            Min = min;
            Max = max;
            Width = width;
        }
    }
}