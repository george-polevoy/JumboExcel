using System;

namespace JumboExcel.Styling
{
    [Flags]
    public enum BorderDefinition
    {
        None = 0,
        Left = 1,
        Right = 1 << 1,
        Top = 1 << 2,
        Bottom = 1 << 3,
        All = Left | Right | Top | Bottom
    }
}
