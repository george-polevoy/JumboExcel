using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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
                        new EmptyCellElement(),
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
            yield return RowForType(new IntegerCellElement(1234567890L), "Int64 with int value.");
            yield return RowForType(new DecimalCellElement(1234567890m), "Decimal with int value.");
            yield return RowForType(new DecimalCellElement(12345.12345m), "With fractional part.");
            yield return RowForNumberStyle(12345.12345m, DecimalFormat.ValueWithExponent1, "Value with exponent.");
            yield return RowForNumberStyle(12345.12345m, DecimalFormat.FractionWithDenominator, "Fraction with denominator.");
            yield return RowForNumberStyle(12345.12345m, DecimalFormat.FractionWithDenominatorPrecise, "Fraction with denominator (precise).");
            yield return RowForNumberStyle(12345.12345m, DecimalFormat.FractionalTwoDecimalPlaces, "Two decimal places.");
            yield return RowForNumberStyle(0.12345m, DecimalFormat.PercentsTwoDecimalPlaces, "12.345%, two decimal places.");
            yield return RowForNumberStyle(1234567890.12345m, DecimalFormat.SeparatorTwoDecimalPlaces, "With separator, two decimal places.");
            yield return RowForNumberStyle(0.12345m, DecimalFormat.IntegerPercents, "12.345%, integer part.");
            yield return RowForNumberStyle(1.12345m, DecimalFormat.PercentsTwoDecimalPlaces, "112.34%.");
            yield return RowForNumberStyle(1.12345m, DecimalFormat.IntegerPercents, "Percent, large, integer part.");
            var rur = new NumberStyleDefinition("#,#0.00\"\u20BD\"", null, BorderDefinition.None, null);
            var usd = new NumberStyleDefinition("\"$\"#,#0.00", null, BorderDefinition.None, null);
            var conditional = new NumberStyleDefinition("\"positive\"* [Green]#,##0.00;\"negative\"* [Red]-#,##0.00;\"zero\"* [Blue]#,##0.00", null, BorderDefinition.None, null);
            yield return RowForType(new DecimalCellElement(123456.7890123m, rur), "Russian roubles.");
            yield return RowForType(new DecimalCellElement(123456.7890123m / 68m, usd), "US dollars.");
            yield return RowForType(new DecimalCellElement(123456.7890123m, conditional), "Conditional colored.");
            yield return RowForType(new DecimalCellElement(0m, conditional), "Conditional colored.");
            yield return RowForType(new DecimalCellElement(-123456.7890123m, conditional), "Conditional colored.");
            foreach (var dateTimeFormat in DateTimeFormat.GetDateTimeFormats())

            {
                var dateTime = new DateTime(2014, 12, 29, 16, 35, 56).AddMilliseconds(125);
                yield return RowForType(new DateCellElement(dateTime, new DateStyleDefinition(dateTimeFormat.FormatCode, null, BorderDefinition.None, null)), dateTimeFormat.FormatCode);
            }
            yield return new RowElement(new DateCellElement(DateTime.Now, new DateStyleDefinition("d-mmm-yy", null, BorderDefinition.None, null)));
        }

        private static RowElement RowForNumberStyle(decimal value, DecimalFormat commonValueFormat, string comment)
        {
            return RowForType(new DecimalCellElement(value, new NumberStyleDefinition(commonValueFormat.FormatCode, null, BorderDefinition.None, null)), comment);
        }

        private static RowElement RowForType<T>(T element, string comment = "") where T : CellElement
        {
            return new RowElement(new InlineStringElement(typeof(T).Name), new EmptyCellElement(), element, new EmptyCellElement(), new InlineStringElement(comment));
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
                borderStylingRows.Add(new RowElement(new EmptyCellElement(),
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

        private static string WriteFile(params WorksheetElement[] worksheetElements)
        {
            return WriteFileInternal(worksheetElements);
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
