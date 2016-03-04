using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using JumboExcel.Productivity;
using JumboExcel.Structure;
using JumboExcel.Styling;
using Font = JumboExcel.Styling.Font;

namespace JumboExcel
{
    class SampleDataSources
    {
        /// <summary>
        /// Generates multiplication table similar to what you see in school math textbooks.
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<IEnumerable<CellElement>> GetMultiplicationTableCells(int columnSpan, int rowSpan)
        {
            var columns = Enumerable.Range(0, columnSpan).ToList();
            var rows = Enumerable.Range(0, rowSpan);
            var headerFont = new Font(null, 11, Color.Black, FontSlope.NORMAL, FontWeight.BOLD);

            yield return new CellElement[] { EmptyCell.Instance }.Concat(columns.Select(c => new IntegerCell(c + 1, new NumberStyle(null, headerFont, BorderHelper.BorderForRange(0, c, 1, columnSpan)))));

            var borderStyler = BorderHelper.AlternateForRange(rowSpan, columnSpan, ContentCellStyle);

            foreach (var r in rows)
            {
                var r1 = r;
                yield return
                    new CellElement[] { new IntegerCell(r + 1, new NumberStyle(null, headerFont, BorderHelper.BorderForRange(r, 0, rowSpan, 1))) }.Concat(
                        columns.Select(c => new IntegerCell((r1 + 1) * (c + 1), borderStyler(r1, c))));
            }
        }

        private static NumberStyle ContentCellStyle(int arg1, int arg2, int arg3, int arg4)
        {
            return new NumberStyle(null, null, BorderHelper.BorderForRange(arg1, arg2, arg3, arg4));
        }
    }
}