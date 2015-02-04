using System;

namespace JumboExcel.Styling
{
    /// <summary>
    /// Represents border presence for cell sides.
    /// </summary>
    [Flags]
    public enum Border
    {
        NONE = 0,
        LEFT = 1,
        RIGHT = 1 << 1,
        TOP = 1 << 2,
        BOTTOM = 1 << 3,
        ALL = LEFT | RIGHT | TOP | BOTTOM
    }
}
