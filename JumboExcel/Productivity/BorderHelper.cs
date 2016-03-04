using System;
using JumboExcel.Styling;

namespace JumboExcel.Productivity
{
    public class BorderHelper
    {
        private static Border Option(Border border, bool isOn)
        {
            return isOn ? border : Border.NONE;
        }

        public static Border BorderForRange(int r, int c, int rowSpan, int columnSpan)
        {
            if (c < 0 || c >= columnSpan || r < 0 || r >= rowSpan)
                return Border.NONE;
            return Option(Border.LEFT, c == 0) | Option(Border.RIGHT, c == columnSpan - 1) | Option(Border.TOP, r == 0) | Option(Border.BOTTOM, r == rowSpan - 1);
        }

        struct Pair<T1, T2>
        {
            public T1 a;
            public T2 b;
        }

        public static Func<int, int, T> AlternateForRange<T>(int rowSpan, int columnSpan, Func<int, int, int, int, T> rangeFunction)
        {
            var tenCases = new Pair<bool, T>[10];

            return (r, c) =>
            {
                var i = c < 0 || c >= columnSpan || r < 0 || r >= rowSpan ? 9 : (r == 0 ? 0 : r < rowSpan - 1 ? 3 : 6) + (c == 0 ? 0 : c < columnSpan - 1 ? 1 : 2);
                return tenCases[i].a ? tenCases[i].b : (tenCases[i] = new Pair<bool, T> {a = true, b = rangeFunction(r, c, rowSpan, columnSpan)}).b;
            };
        }
    }
}
