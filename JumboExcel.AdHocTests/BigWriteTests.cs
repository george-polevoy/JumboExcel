using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using JumboExcel.Formatting;
using JumboExcel.Styling;
using JumboExcel.Structure;
using NUnit.Framework;
using Font = JumboExcel.Styling.Font;

namespace JumboExcel
{
    class BigWriteTests
    {
        [Test, Explicit]
        public void TestSimplestReport()
        {
            var tempFileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var file = new FileStream(tempFileName, FileMode.CreateNew))
            {
                OpenXmlBuilder.Write(file, new[] { new Worksheet("Report", new WorksheetParametersElement(), new Row(new SharedString("Hello"))) });
            }
            Process.Start(tempFileName);
        }

        [Test, Explicit]
        public void WriteWithProgress()
        {
            var progressingWorksheets = new[] { new ProgressingWorksheet<int>("Progressing", new WorksheetParametersElement(), GenerateRows) };

            var fileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var outputStream = new FileStream(fileName, FileMode.CreateNew))
            {
                foreach (var iteration in OpenXmlBuilder.WriteWithProgress(outputStream, progressingWorksheets))
                {
                    Console.WriteLine("Progress: {0}", iteration);
                }
            }
            var fileSize = new FileInfo(fileName).Length;
            Console.WriteLine("Size of the file generated: {0}", fileSize);
            Console.WriteLine(fileName);

            Process.Start(fileName);
        }

        private IEnumerable<int> GenerateRows(Action<IEnumerable<RowLevelElement>> writeElements)
        {
            yield return 0;

            writeElements(new[] { new Row(new IntegerCell(1)) });

            yield return 50;

            writeElements(new[] { new Row(new IntegerCell(2)) });

            yield return 100;
        }

        [Test, Explicit]
        public void Styles()
        {
            var mediumColumns = new WorksheetParametersElement(false, false, Enumerable.Range(0, 20).Select(column => new ColumnConfiguration(column, column, 30)));
            var columnsForFonts = new WorksheetParametersElement(false, false, new ColumnConfiguration(0, 0, 50), new ColumnConfiguration(1, 1, 120));

            WriteAndExecuteExcel(new[]
            {
                new Worksheet("Data Types", mediumColumns, GetDataTypeRows()),
                new Worksheet("Nullable data types styling.", mediumColumns, GetRowsForNullValues()),
                new Worksheet("Row Groupings", mediumColumns,
                    new Row(new SharedString("Level 1")),
                    new RowGroup(
                        new Row(new SharedString("Level 2")),
                        new RowGroup(
                            new Row(new SharedString("Level 3"))),
                        new Row(new SharedString("Level 2")),
                        new RowGroup(
                            new Row(new SharedString("Level 3"))))
                    ),
                new Worksheet("Fonts", columnsForFonts, GetFontsRows()),
                new Worksheet("Colors", mediumColumns, GetColorRows()),
                new Worksheet("Border Styling", mediumColumns, GetBorderStylingRows()),
                new Worksheet("Column Widths", new WorksheetParametersElement(false, false, new ColumnConfiguration(0, 1, 20), new ColumnConfiguration(2,2, 60)),
                    new Row(
                        new SharedString("Narrow"),
                        new SharedString("Narrow"),
                        new SharedString("Wide"),
                        new SharedString("Unspec."),
                        new SharedString("Messed.")),
                    new Row(
                        new SharedString("Narrow column"),
                        new SharedString("Another narrow"),
                        new SharedString("This is a wide column, which displays just perfect."),
                        new SharedString("This is a wide column, which has unspecified width."),
                        new SharedString("This value messes with the previous.")))
            });
        }

        [Test]
        [Explicit]
        [TestCase(1000000, 1)]
        public void ReallyHuge(int rowCount, int columnCount)
        {
            WriteAndExecuteExcel(Enumerable.Range(0, 1).Select(sheet => new Worksheet("Sheet  " + sheet, new WorksheetParametersElement(), Enumerable.Range(0, rowCount).Select(row => new Row(
                new[] { new SharedString("Row mod 10: " + row % 10) }
                    .Concat(Enumerable.Range(0, columnCount).SelectMany(column => new CellElement[]
                    {
                        new SharedString("Column: " + column),
                        EmptyCell.Instance,
                        new IntegerCell(row*column + sheet),
                        new DecimalCell((decimal) column/columnCount),
                        new InlineString("Row: " + row*column),
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
            var file = WriteFile(Enumerable.Range(0, sheetCount).Select(sheet => new Worksheet("Sheet  " + sheet, new WorksheetParametersElement(), Enumerable.Range(0, rowCount).Select(row => new Row(
                Enumerable.Range(0, columnCount).Select(column => new IntegerCell(row * column + sheet)))))));
            File.Delete(file);
        }

        private IEnumerable<Row> GetRowsForNullValues()
        {
            var headerStyle = new StringStyle(new Font("Arial", 16, Color.White, FontSlope.NORMAL, FontWeight.BOLD), Border.ALL, Color.Teal);
            yield return new Row(new SharedString("Value Type", headerStyle), new SharedString("Missing Value", headerStyle));

            yield return MissingValueRow(new DecimalCell(null));
            yield return MissingValueRow(new IntegerCell(null));
        }

        private static Row MissingValueRow(CellElement element)
        {
            return new Row(new SharedString(element.GetType().Name), element);
        }

        private IEnumerable<Row> GetFontsRows()
        {
            return from fontFace in new[] { null, "Calibri", "Times New Roman" }
                   from weight in new[] { FontWeight.NORMAL, FontWeight.BOLD, }
                   from slope in new[] { FontSlope.NORMAL, FontSlope.ITALIC, }
                   from size in new[] { 7, 11, 24 }
                   select RowForFont(fontFace, slope, weight, size);
        }

        private IEnumerable<Row> GetColorRows()
        {
            return from foregroundColor in new[] { Color.Black, Color.Blue, Color.Brown }
                   from backgroundColor in new Color?[] { null, Color.BlanchedAlmond, Color.DarkSeaGreen, Color.Azure }
                   select RowForColor(foregroundColor, backgroundColor);
        }

        private Row RowForColor(Color foregroundColor, Color? backgroundColor, string comment = "")
        {
            var fontDefinition = new Font(null, 11, foregroundColor, FontSlope.NORMAL, FontWeight.NORMAL);
            var style = new StringStyle(fontDefinition, Border.NONE, backgroundColor);
            return new Row(
                new InlineString(string.Format("{0} over {1}", foregroundColor.ToString(), backgroundColor == null ? "default" : backgroundColor.ToString())),
                new SharedString("Quick brown fox jumps over the lazy dog. 12345676890", style), new InlineString(comment));
        }

        private Row RowForFont(string fontFace, FontSlope slope, FontWeight weight, int size, string comment = "")
        {
            var fontDefinition = new Font(fontFace, size, Color.Black, slope, weight);
            var style = new StringStyle(fontDefinition, Border.NONE, null);
            return new Row(
                new InlineString(fontDefinition.ToString()), new SharedString("Quick brown fox jumps over the lazy dog. 12345676890", style), new InlineString(comment));
        }

        private static IEnumerable<Row> GetDataTypeRows()
        {
            yield return RowForType(new InlineString("It's a string."));
            yield return RowForType(new InlineString(null), "With a null value.");
            yield return RowForType(new SharedString("It's a shared string."));
            yield return RowForType(new SharedString(null), "With a null value.");
            yield return RowForType(new SharedString("123"), "Shared string, with a number-like value. Will produce a warning.");
            yield return RowForType(new InlineString("123"), "Inline string, with a number-like value. Will produce a warning.");
            var rur = new NumberStyle(NumberFormat.FromFormatString("#,#0.00\"\u20BD\""), null, Border.NONE, null);
            var usd = new NumberStyle(NumberFormat.FromFormatString("\"$\"#,#0.00"), null, Border.NONE, null);
            var conditional = new NumberStyle(NumberFormat.FromFormatString("\"positive\"* [Green]#,##0.00;\"negative\"* [Red]-#,##0.00;\"zero\"* [Blue]#,##0.00"), null, Border.NONE, null);
            yield return RowForType(new DecimalCell(123456.7890123m, rur), "Russian roubles.");
            yield return RowForType(new DecimalCell(123456.7890123m / 68m, usd), "US dollars.");
            yield return RowForType(new DecimalCell(123456.7890123m, conditional), "Conditional colored.");
            yield return RowForType(new DecimalCell(0m, conditional), "Conditional colored.");
            yield return RowForType(new DecimalCell(-123456.7890123m, conditional), "Conditional colored.");
            yield return RowForType(new IntegerCell(100L), "new IntegerCellElement(100L)");
            yield return RowForType(new DecimalCell(123.123m), "new DecimalCellElement(123.123m)");
            yield return RowForType(new IntegerCell(100000000L, new NumberStyle(default(NumberFormat), null, Border.NONE)), "default(NumberFormat)");
            yield return RowForType(new IntegerCell(100000000L, new NumberStyle(NumberFormat.Default, null, Border.NONE)), "NumberFormat.Default");
            yield return RowForType(new IntegerCell(100000000L, new NumberStyle(IntegerFormat.General, null, Border.NONE)), "IntegerFormat.General");
            yield return RowForType(new DecimalCell(123.456m, new NumberStyle(default(NumberFormat), null, Border.NONE)), "default(NumberFormat)");
            yield return RowForType(new DecimalCell(123.123m, new NumberStyle(NumberFormat.Default, null, Border.NONE)), "NumberFormat.Default");
            yield return RowForType(new DecimalCell(123.123m, new NumberStyle(IntegerFormat.General, null, Border.NONE)), "IntegerFormat.General");
            foreach (var format in IntegerFormat.GetIntegerFormats())
            {
                var numberStyleDefinition = new NumberStyle(format, null, Border.NONE, null);
                foreach (var value in new long[] {-123456, -123, 0, 123, 123456})
                {
                    yield return RowForType(new IntegerCell(value, numberStyleDefinition), GetValueFormatComment(format, value.ToString(CultureInfo.InvariantCulture)));
                    yield return RowForType(new IntegerCell(value), GetNoStyleComment(value.ToString(CultureInfo.InvariantCulture)));
                }
            }
            foreach (var format in DecimalFormat.GetDecimalFormats())
            {
                var numberStyleDefinition = new NumberStyle(format, null, Border.NONE, null);
                foreach (var value in new[] {-123456.123456m, -123, -0.12345m, 0, 0.12345m, 123, 123456.123456m})
                {
                    yield return RowForType(new DecimalCell(value, numberStyleDefinition), GetValueFormatComment(format, value.ToString(CultureInfo.InvariantCulture)));
                    yield return RowForType(new DecimalCell(value), GetNoStyleComment(value.ToString(CultureInfo.InvariantCulture)));
                }
            }
            var dateValue = new DateTime(2014, 12, 29, 16, 35, 56).AddMilliseconds(125);
            var customDateFormats = new[] {default(DateTimeFormat), new DateTimeFormat("m/d/yy H:mm:ss")};
            foreach (var format in DateTimeFormat.GetDateTimeFormats().Concat(customDateFormats))
            {
                yield return RowForType(new DateTimeCell(dateValue, new DateStyle(format, null, Border.NONE, null)), GetValueFormatComment(format, dateValue.ToString("u")));
            }

            yield return RowForType(new BooleanCell(null), "Null boolean without style.");
            yield return RowForType(new BooleanCell(true), "True boolean without style.");
            yield return RowForType(new BooleanCell(false), "False boolean without style.");
            yield return RowForType(new BooleanCell(null, new BooleanStyle(null, Border.NONE, Color.SkyBlue)), "Null boolean styled.");
            yield return RowForType(new BooleanCell(true, new BooleanStyle(null, Border.NONE, Color.SkyBlue)), "True boolean styled.");
            yield return RowForType(new BooleanCell(false, new BooleanStyle(null, Border.NONE, Color.SkyBlue)), "False boolean styled.");
        }

        private static string GetNoStyleComment(string value)
        {
            return string.Format("No format specified. Value: {0}", value);
        }

        private static string GetValueFormatComment(CommonValueFormat numberFormat, string value)
        {
            return string.Format("Format: \"{0}\", value: {1}", numberFormat != null ? numberFormat.FormatCode : null, value);
        }

        private static Row RowForType<T>(T element, string comment = "") where T : CellElement
        {
            return new Row(new InlineString(typeof(T).Name), EmptyCell.Instance, element, EmptyCell.Instance, new InlineString(comment));
        }

        private static IEnumerable<Row> GetBorderStylingRows()
        {
            Func<int, Border> borderVariant = i =>
                ((i & 1) != 0 ? Border.LEFT : Border.NONE) |
                ((i & 1 << 1) != 0 ? Border.RIGHT : Border.NONE) |
                ((i & 1 << 2) != 0 ? Border.TOP : Border.NONE) |
                ((i & 1 << 3) != 0 ? Border.BOTTOM : Border.NONE);

            var borderVariants = Enumerable.Range(0, 16).Select(i => borderVariant(i)).ToArray();

            var borderStylingRows = new List<Row> { new Row() };
            foreach (var bv in borderVariants)
            {
                borderStylingRows.Add(new Row(EmptyCell.Instance,
                    new InlineString(bv.ToString(), new StringStyle(null, bv, null))));
                AddEmptyRow(borderStylingRows);
            }
            return borderStylingRows;
        }

        private static void AddEmptyRow(List<Row> rows)
        {
            rows.Add(new Row(Enumerable.Empty<CellElement>()));
        }

        private static void WriteAndExecuteExcel(IEnumerable<Worksheet> worksheetElements)
        {
            var fileName = WriteFile(worksheetElements);
            Process.Start(fileName);
        }

        private static string WriteFile(IEnumerable<Worksheet> worksheetElements)
        {
            return WriteFileInternal(worksheetElements);
        }

        private static string WriteFileInternal(IEnumerable<Worksheet> worksheetElements)
        {
            var fileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var outputStream = new FileStream(fileName, FileMode.CreateNew))
            {
                OpenXmlBuilder.Write(
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
