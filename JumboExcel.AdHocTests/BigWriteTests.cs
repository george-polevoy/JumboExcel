using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using JumboExcel.Formatting;
using JumboExcel.Structure;
using JumboExcel.Styling;
using NUnit.Framework;

namespace JumboExcel
{
    class BigWriteTests
    {
        [Test, Explicit]
        public void Styles()
        {
            var mediumColumns = new WorksheetParametersElement(false, false, Enumerable.Range(0, 20).Select(column => new ColumnElement(column, column, 30)));
            var columnsForFonts = new WorksheetParametersElement(false, false, new ColumnElement(0, 0, 50), new ColumnElement(1, 1, 120));

            WriteAndExecuteExcel(new[]
            {
                new WorksheetElement("Data Types", mediumColumns, GetDataTypeRows()),
                new WorksheetElement("Nullable data types styling.", mediumColumns, GetRowsForNullValues()),
                new WorksheetElement("Row Groupings", mediumColumns,
                    new RowElement(new SharedStringElement("Level 0")),
                    new RowGroupElement(
                        new RowElement(new SharedStringElement("Level 1")),
                        new RowGroupElement(
                            new RowElement(new SharedStringElement("Level 2"))),
                        new RowElement(new SharedStringElement("Level 1"))),
                    new RowElement(new SharedStringElement("Level 0")),
                    new RowGroupElement(
                        new RowElement(new SharedStringElement("Level 1")),
                        new RowElement(new SharedStringElement("Level 1")))),
                new WorksheetElement("Fonts", columnsForFonts, GetFontsRows()),
                new WorksheetElement("Colors", mediumColumns, GetColorRows()),
                new WorksheetElement("Border Styling", mediumColumns, GetBorderStylingRows()),
                new WorksheetElement("Column Widths", new WorksheetParametersElement(false, false, new ColumnElement(0, 1, 20), new ColumnElement(2,2, 60)),
                    new RowElement(
                        new SharedStringElement("Narrow"),
                        new SharedStringElement("Narrow"),
                        new SharedStringElement("Wide"),
                        new SharedStringElement("Unspec."),
                        new SharedStringElement("Messed.")),
                    new RowElement(
                        new SharedStringElement("Narrow column"),
                        new SharedStringElement("Another narrow"),
                        new SharedStringElement("This is a wide column, which displays just perfect."),
                        new SharedStringElement("This is a wide column, which has unspecified width."),
                        new SharedStringElement("This value messes with the previous.")))
            });
        }

        [Test]
        [Explicit]
        [TestCase(1000000, 1)]
        public void ReallyHuge(int rowCount, int columnCount)
        {
            WriteAndExecuteExcel(Enumerable.Range(0, 1).Select(sheet => new WorksheetElement("Sheet  " + sheet, new WorksheetParametersElement(), Enumerable.Range(0, rowCount).Select(row => new RowElement(
                new[] { new SharedStringElement("Row mod 10: " + row % 10) }
                    .Concat(Enumerable.Range(0, columnCount).SelectMany(column => new CellElement[]
                    {
                        new SharedStringElement("Column: " + column),
                        EmptyCellElement.Instance,
                        new IntegerCellElement(row*column + sheet),
                        new DecimalCellElement((decimal) column/columnCount),
                        new InlineStringElement("Row: " + row*column),
                    })))))));
        }

        [Test]
        [Explicit]
        [TestCase(1, 100000, 1)]
        [TestCase(1, 10000, 10)]
        [TestCase(1, 100000, 10)]
        [TestCase(1, 1000000, 1)]
        [TestCase(2, 1000000, 5)]
        public void PerformanceTest(int sheetCount, int rowCount, int columnCount)
        {
            var file = WriteFile(Enumerable.Range(0, sheetCount).Select(sheet => new WorksheetElement("Sheet  " + sheet, new WorksheetParametersElement(), Enumerable.Range(0, rowCount).Select(row => new RowElement(
                Enumerable.Range(0, columnCount).Select(column => new IntegerCellElement(row * column + sheet)))))));
            File.Delete(file);
        }

        private IEnumerable<RowElement> GetRowsForNullValues()
        {
            var headerStyle = new SharedStringStyleDefinition(new FontDefinition("Arial", 16, Color.White, FontSlope.Normal, FontWeight.Bold), BorderDefinition.All, Color.Teal);
            yield return new RowElement(new SharedStringElement("Value Type", headerStyle), new SharedStringElement("Missing Value", headerStyle));

            yield return MissingValueRow(new DecimalCellElement(null));
            yield return MissingValueRow(new IntegerCellElement(null));
        }

        private static RowElement MissingValueRow(CellElement element)
        {
            return new RowElement(new SharedStringElement(element.GetType().Name), element);
        }

        private IEnumerable<RowElement> GetFontsRows()
        {
            return from fontFace in new[] { null, "Calibri", "Times New Roman" }
                   from weight in new[] { FontWeight.Normal, FontWeight.Bold, }
                   from slope in new[] { FontSlope.Normal, FontSlope.Italic, }
                   from size in new[] { 7, 11, 24 }
                   select RowForFont(fontFace, slope, weight, size);
        }

        private IEnumerable<RowElement> GetColorRows()
        {
            return from foregroundColor in new[] { Color.Black, Color.Blue, Color.Brown }
                   from backgroundColor in new Color?[] { null, Color.BlanchedAlmond, Color.DarkSeaGreen, Color.Azure }
                   select RowForColor(foregroundColor, backgroundColor);
        }

        private RowElement RowForColor(Color foregroundColor, Color? backgroundColor, string comment = "")
        {
            var fontDefinition = new FontDefinition(null, 11, foregroundColor, FontSlope.Normal, FontWeight.Normal);
            var style = new SharedStringStyleDefinition(fontDefinition, BorderDefinition.None, backgroundColor);
            return new RowElement(
                new InlineStringElement(string.Format("{0} over {1}", foregroundColor.ToString(), backgroundColor == null ? "default" : backgroundColor.ToString())),
                new SharedStringElement("Quick brown fox jumps over the lazy dog. 12345676890", style), new InlineStringElement(comment));
        }

        private RowElement RowForFont(string fontFace, FontSlope slope, FontWeight weight, int size, string comment = "")
        {
            var fontDefinition = new FontDefinition(fontFace, size, Color.Black, slope, weight);
            var style = new SharedStringStyleDefinition(fontDefinition, BorderDefinition.None, null);
            return new RowElement(
                new InlineStringElement(fontDefinition.ToString()), new SharedStringElement("Quick brown fox jumps over the lazy dog. 12345676890", style), new InlineStringElement(comment));
        }

        private static IEnumerable<RowElement> GetDataTypeRows()
        {
            yield return RowForType(new InlineStringElement("It's a string."));
            yield return RowForType(new InlineStringElement(null), "With a null value.");
            yield return RowForType(new SharedStringElement("It's a shared string."));
            yield return RowForType(new SharedStringElement(null), "With a null value.");
            yield return RowForType(new SharedStringElement("123"), "Shared string, with a number-like value. Will produce a warning.");
            yield return RowForType(new InlineStringElement("123"), "Inline string, with a number-like value. Will produce a warning.");
            var rur = new NumberStyleDefinition(NumberFormat.FromFormatString("#,#0.00\"\u20BD\""), null, BorderDefinition.None, null);
            var usd = new NumberStyleDefinition(NumberFormat.FromFormatString("\"$\"#,#0.00"), null, BorderDefinition.None, null);
            var conditional = new NumberStyleDefinition(NumberFormat.FromFormatString("\"positive\"* [Green]#,##0.00;\"negative\"* [Red]-#,##0.00;\"zero\"* [Blue]#,##0.00"), null, BorderDefinition.None, null);
            yield return RowForType(new DecimalCellElement(123456.7890123m, rur), "Russian roubles.");
            yield return RowForType(new DecimalCellElement(123456.7890123m / 68m, usd), "US dollars.");
            yield return RowForType(new DecimalCellElement(123456.7890123m, conditional), "Conditional colored.");
            yield return RowForType(new DecimalCellElement(0m, conditional), "Conditional colored.");
            yield return RowForType(new DecimalCellElement(-123456.7890123m, conditional), "Conditional colored.");
            yield return RowForType(new IntegerCellElement(100L), "new IntegerCellElement(100L)");
            yield return RowForType(new DecimalCellElement(123.123m), "new DecimalCellElement(123.123m)");
            yield return RowForType(new IntegerCellElement(100000000L, new NumberStyleDefinition(default(NumberFormat), null, BorderDefinition.None)), "default(NumberFormat)");
            yield return RowForType(new IntegerCellElement(100000000L, new NumberStyleDefinition(NumberFormat.Default, null, BorderDefinition.None)), "NumberFormat.Default");
            yield return RowForType(new IntegerCellElement(100000000L, new NumberStyleDefinition(IntegerFormat.General, null, BorderDefinition.None)), "IntegerFormat.General");
            yield return RowForType(new DecimalCellElement(123.456m, new NumberStyleDefinition(default(NumberFormat), null, BorderDefinition.None)), "default(NumberFormat)");
            yield return RowForType(new DecimalCellElement(123.123m, new NumberStyleDefinition(NumberFormat.Default, null, BorderDefinition.None)), "NumberFormat.Default");
            yield return RowForType(new DecimalCellElement(123.123m, new NumberStyleDefinition(IntegerFormat.General, null, BorderDefinition.None)), "IntegerFormat.General");
            foreach (var format in IntegerFormat.GetIntegerFormats())
            {
                var numberStyleDefinition = new NumberStyleDefinition(format, null, BorderDefinition.None, null);
                foreach (var value in new long[] {-123456, -123, 0, 123, 123456})
                {
                    yield return RowForType(new IntegerCellElement(value, numberStyleDefinition), GetValueFormatComment(format, value.ToString(CultureInfo.InvariantCulture)));
                    yield return RowForType(new IntegerCellElement(value), GetNoStyleComment(value.ToString(CultureInfo.InvariantCulture)));
                }
            }
            foreach (var format in DecimalFormat.GetDecimalFormats())
            {
                var numberStyleDefinition = new NumberStyleDefinition(format, null, BorderDefinition.None, null);
                foreach (var value in new[] {-123456.123456m, -123, -0.12345m, 0, 0.12345m, 123, 123456.123456m})
                {
                    yield return RowForType(new DecimalCellElement(value, numberStyleDefinition), GetValueFormatComment(format, value.ToString(CultureInfo.InvariantCulture)));
                    yield return RowForType(new DecimalCellElement(value), GetNoStyleComment(value.ToString(CultureInfo.InvariantCulture)));
                }
            }
            var dateValue = new DateTime(2014, 12, 29, 16, 35, 56).AddMilliseconds(125);
            var customDateFormats = new[] {default(DateTimeFormat), new DateTimeFormat("m/d/yy H:mm:ss")};
            foreach (var format in DateTimeFormat.GetDateTimeFormats().Concat(customDateFormats))
            {
                yield return RowForType(new DateTimeCellElement(dateValue, new DateStyleDefinition(format, null, BorderDefinition.None, null)), GetValueFormatComment(format, dateValue.ToString("u")));
            }

            yield return RowForType(new BooleanCellElement(null), "Null boolean without style.");
            yield return RowForType(new BooleanCellElement(true), "True boolean without style.");
            yield return RowForType(new BooleanCellElement(false), "False boolean without style.");
            yield return RowForType(new BooleanCellElement(null, new BooleanStyleDefinition(null, BorderDefinition.None, Color.SkyBlue)), "Null boolean styled.");
            yield return RowForType(new BooleanCellElement(true, new BooleanStyleDefinition(null, BorderDefinition.None, Color.SkyBlue)), "True boolean styled.");
            yield return RowForType(new BooleanCellElement(false, new BooleanStyleDefinition(null, BorderDefinition.None, Color.SkyBlue)), "False boolean styled.");
        }

        private static string GetNoStyleComment(string value)
        {
            return string.Format("No format specified. Value: {0}", value);
        }

        private static string GetValueFormatComment(CommonValueFormat numberFormat, string value)
        {
            return string.Format("Format: \"{0}\", value: {1}", numberFormat != null ? numberFormat.FormatCode : null, value);
        }

        private static RowElement RowForType<T>(T element, string comment = "") where T : CellElement
        {
            return new RowElement(new InlineStringElement(typeof(T).Name), EmptyCellElement.Instance, element, EmptyCellElement.Instance, new InlineStringElement(comment));
        }

        private static IEnumerable<RowElement> GetBorderStylingRows()
        {
            Func<int, BorderDefinition> borderVariant = i =>
                ((i & 1) != 0 ? BorderDefinition.Left : BorderDefinition.None) |
                ((i & 1 << 1) != 0 ? BorderDefinition.Right : BorderDefinition.None) |
                ((i & 1 << 2) != 0 ? BorderDefinition.Top : BorderDefinition.None) |
                ((i & 1 << 3) != 0 ? BorderDefinition.Bottom : BorderDefinition.None);

            var borderVariants = Enumerable.Range(0, 16).Select(i => borderVariant(i)).ToArray();

            var borderStylingRows = new List<RowElement> { new RowElement() };
            foreach (var bv in borderVariants)
            {
                borderStylingRows.Add(new RowElement(EmptyCellElement.Instance,
                    new InlineStringElement(bv.ToString(), new StringStyleDefinition(null, bv, null))));
                AddEmptyRow(borderStylingRows);
            }
            return borderStylingRows;
        }

        private static void AddEmptyRow(List<RowElement> rows)
        {
            rows.Add(new RowElement(Enumerable.Empty<CellElement>()));
        }

        private static void WriteAndExecuteExcel(IEnumerable<WorksheetElement> worksheetElements)
        {
            var fileName = WriteFile(worksheetElements);
            Process.Start(fileName);
        }

        private static string WriteFile(IEnumerable<WorksheetElement> worksheetElements)
        {
            return WriteFileInternal(worksheetElements);
        }

        private static string WriteFileInternal(IEnumerable<WorksheetElement> worksheetElements)
        {
            var fileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var outputStream = new FileStream(fileName, FileMode.CreateNew))
            {
                var builder = new OpenXmlBuilder();
                builder.Write(
                    outputStream,
                    worksheetElements
                    );
            }
            var fileSize = new FileInfo(fileName).Length;
            Console.WriteLine("Size of the file generated: {0}", fileSize);
            Console.WriteLine(fileName);
            return fileName;
        }
    }
}
