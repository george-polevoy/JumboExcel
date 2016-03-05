using System;

namespace JumboExcel.Structure
{
    /// <summary>
    /// Represents column configuration. Allows to specify widths.
    /// </summary>
    public sealed class ColumnConfiguration
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
        public decimal Width { get; private set; }

        /// <summary>
        /// Nesting level of grouped columns, ranging from 0 (no grouping) to 255.
        /// </summary>
        public int OutlineLevel { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="min">Specifies the minimum column index to apply this configuration to.</param>
        /// <param name="max">Specifies the maximum column index to apply this configuration to.</param>
        /// <param name="width">Width, in excel display units to apply to the column range.</param>
        /// <param name="outlineLevel">Outline level for row grouping. 0 &lt; <paramref name="outlineLevel"/> &lt; 255 </param>
        public ColumnConfiguration(int min, int max, decimal width, int outlineLevel = 0)
        {
            if (min < 0)
                throw new ArgumentOutOfRangeException("min", min, "Must be > 0.");

            if (max < min)
                throw new ArgumentOutOfRangeException("max", max, "Must be > min.");

            if (width < 0)
                throw new ArgumentOutOfRangeException("width", width, "Must be > 0.");

            if (outlineLevel < 0 || outlineLevel > 255)
                throw new ArgumentOutOfRangeException("outlineLevel", outlineLevel, "Must in range (0,255)");

            Min = min;
            Max = max;
            Width = width;
            OutlineLevel = outlineLevel;
        }
    }
}