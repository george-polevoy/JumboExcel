using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using JumboExcel.Formatting;
using JumboExcel.Structure;
using JumboExcel.Styling;
using NUnit.Framework;
using Font = JumboExcel.Styling.Font;

namespace JumboExcel
{
    class StyleTests
    {
        [Test, Explicit]
        public void Styles()
        {
            var mediumColumns = new WorksheetParametersElement(false, false, Enumerable.Range(0, 20).Select(column => new ColumnConfiguration(column, column, 30)), null);
            var columnsForFonts = new WorksheetParametersElement(false, false, new ColumnConfiguration(0, 0, 50), new ColumnConfiguration(1, 1, 120));

            var mergedCellStyle = new StringStyle(null, Border.ALL, Color.Beige);
            var skippedMergedCell = new InlineString(null, mergedCellStyle);
            var mergedCellPereferial = new InlineString("usual", new StringStyle(null, Border.NONE, Color.Bisque));
            TestHelper.WriteAndExecuteExcel(new[]
            {
                new Worksheet("Data Types", mediumColumns, GetDataTypeRows()),
                new Worksheet("Frozen Panes", new WorksheetParametersElement(false, false, null, new PaneFreezer(1, 1)), SampleDataSources.GetMultiplicationTableCells(40, 60).Select(r => new Row(r))),
                new Worksheet("Nullable data types styling.", mediumColumns, GetRowsForNullValues()),
                new Worksheet("Row Groupings", mediumColumns,
                    new Row(new SharedString("Level 1")),
                    new RowGroup(
                        new RowLevelElement[] {
                            new Row(new SharedString("Level 2")),
                            new RowGroup(new[]{new Row(new SharedString("Level 3"))}, false),
                            new Row(new SharedString("Level 2")),
                            new RowGroup(new[]{new Row(new SharedString("Level 3"))}, false)
                        }, false),
                    new Row(new SharedString("Level 1")),
                    new RowGroup(
                        new RowLevelElement[] {
                            new Row(new SharedString("Level 2")),
                            new RowGroup(new[]{new Row(new SharedString("Level 3"))}, true),
                            new Row(new SharedString("Level 2")),
                            new RowGroup(new[]{new Row(new SharedString("Level 3"))}, true)
                        }, true),
                    new Row(new SharedString("Level 1")),
                    new RowGroup(
                        new RowLevelElement[] {
                            new Row(new SharedString("Level 2")),
                            new RowGroup(new[]{new Row(new SharedString("Level 3"))}, false),
                            new Row(new SharedString("Level 2")),
                            new RowGroup(new[]{new Row(new SharedString("Level 3"))}, false)
                        }, false)),
                new Worksheet(
                    "Column groupings",
                    new WorksheetParametersElement(true, true,
                        new ColumnConfiguration(0, 1, 20, 2), // columns in the group of Fruit
                        new ColumnConfiguration(2,2, 30, 1),  // a column in the group of Plant
                        new ColumnConfiguration(3, 4, 20, 2), // columns in the group of Tree
                        new ColumnConfiguration(5,5, 30, 1)  // a column in the group of Plant
                        ),
                    new Row(
                        new SharedString(@"Plant\Fruit\Banana"), new SharedString(@"Plant\Fruit\Apple"), new SharedString(@"Plant\Fruit"),
                        new SharedString(@"Plant\Tree\Pine"), new SharedString(@"Plant\Tree\Oak"), new SharedString(@"Plant\Tree"), new SharedString("Plant")), 
                    new Row(new IntegerCell(1), new IntegerCell(3), new IntegerCell(4), new IntegerCell(10), new IntegerCell(30), new IntegerCell(40), new IntegerCell(44))),
                new Worksheet("Merged cells",
                    new WorksheetParametersElement(),
                    new Row(Enumerable.Repeat(mergedCellPereferial, 5)),
                    new Row(mergedCellPereferial, new SharedString("3 cols, 2 rows", mergedCellStyle), skippedMergedCell, skippedMergedCell, mergedCellPereferial),
                    new Row(mergedCellPereferial, skippedMergedCell, skippedMergedCell, new RelativeCellMerger(skippedMergedCell, 1, 2), mergedCellPereferial),
                    new Row(Enumerable.Repeat(mergedCellPereferial, 5))),
                new Worksheet("Fonts", columnsForFonts, GetFontsRows()),
                new Worksheet("Colors", mediumColumns, GetColorRows()),
                new Worksheet("Border Styling", mediumColumns, GetBorderStylingRows()),
                new Worksheet("Column Widths", new WorksheetParametersElement(false, false, new ColumnConfiguration(0, 1, 20), new ColumnConfiguration(2,2, 60)),
                    new Row(
                        new SharedString("Narrow"),
                        new SharedString("Narrow"),
                        new SharedString("Wide"),
                        new SharedString("Unspecified"),
                        new SharedString("Messed")),
                    new Row(
                        new SharedString("Narrow column"),
                        new SharedString("Another narrow"),
                        new SharedString("This is a wide column, which displays just perfect."),
                        new SharedString("This is a wide column, which has unspecified width."),
                        new SharedString("This value messes with the previous."))),
                new Worksheet("Text Rotation", new WorksheetParametersElement(false, false, new ColumnConfiguration(0,0, 20), new ColumnConfiguration(1, 1, 20)),
                    Enumerable.Range(0, 5).Select(i=>CreateRotatedRow(i * 45))),
                new Worksheet("Wrap text", new WorksheetParametersElement(false, false),
                    new Row(new SharedString("This is a wide string that should be wrapped.", CreateWrappedStringStyle()), new SharedString("This is a wide string that should be wrapped.", CreateWrappedStringStyle()), new SharedString("This is a wide string that is not wrapped.")),
                    new Row(new SharedString("This is a wide string that should not be wrapped."), new SharedString("This is a wide string that should not be wrapped."))
                    ),
                new Worksheet("This is a long worksheet name. More than 32 characters for sure.", new WorksheetParametersElement(false, false, null, null, WorksheetCompatibilityFlags.RELAX_WORKSHEET_LENGTH_CONSTRAINT))
            });
        }

        static StringStyle CreateWrappedStringStyle()
        {
            return new StringStyle(null, Border.NONE, null, new Alignment(HorizontalAlignment.LEFT, VerticalAlignment.TOP, 0, true));
        }

        static Row CreateRotatedRow(int textRotation)
        {
            return new Row(new IntegerCell(textRotation), new SharedString("Rotated text", new StringStyle(null, Border.NONE, null, new Alignment(HorizontalAlignment.LEFT, VerticalAlignment.TOP, textRotation))));
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

        private static Row RowForColor(Color foregroundColor, Color? backgroundColor, string comment = "")
        {
            var fontDefinition = new Font(null, 11, foregroundColor, FontSlope.NORMAL, FontWeight.NORMAL);
            var style = new StringStyle(fontDefinition, Border.NONE, backgroundColor);
            return new Row(
                new InlineString(string.Format("{0} over {1}", foregroundColor.ToString(), backgroundColor == null ? "default" : backgroundColor.ToString())),
                new SharedString("Quick brown fox jumps over the lazy dog. 12345676890", style), new InlineString(comment));
        }

        private static Row RowForFont(string fontFace, FontSlope slope, FontWeight weight, int size, string comment = "")
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
                foreach (var value in new long[] { -123456, -123, 0, 123, 123456 })
                {
                    yield return RowForType(new IntegerCell(value, numberStyleDefinition), GetValueFormatComment(format, value.ToString(CultureInfo.InvariantCulture)));
                    yield return RowForType(new IntegerCell(value), GetNoStyleComment(value.ToString(CultureInfo.InvariantCulture)));
                }
            }
            foreach (var format in DecimalFormat.GetDecimalFormats())
            {
                var numberStyleDefinition = new NumberStyle(format, null, Border.NONE, null);
                foreach (var value in new[] { -123456.123456m, -123, -0.12345m, 0, 0.12345m, 123, 123456.123456m })
                {
                    yield return RowForType(new DecimalCell(value, numberStyleDefinition), GetValueFormatComment(format, value.ToString(CultureInfo.InvariantCulture)));
                    yield return RowForType(new DecimalCell(value), GetNoStyleComment(value.ToString(CultureInfo.InvariantCulture)));
                }
            }
            var dateValue = new DateTime(2014, 12, 29, 16, 35, 56).AddMilliseconds(125);
            var customDateFormats = new[]
            {
                default(DateTimeFormat),
                new DateTimeFormat("m/d/yy H:mm:ss"),
                new DateTimeFormat("dd.MM.yyyy"),
                new DateTimeFormat("hh:mm:ss"),
            };
            foreach (var format in DateTimeFormat.GetDateTimeFormats().Concat(customDateFormats))
            {
                yield return RowForType(new DateTimeCell(dateValue, new DateStyle(format, null, Border.NONE, null)), GetValueFormatComment(format, dateValue.ToString("u")));
            }

            yield return TimeSpanRow("[h]:mm");

            yield return RowForType(new BooleanCell(null), "Null boolean without style.");
            yield return RowForType(new BooleanCell(true), "True boolean without style.");
            yield return RowForType(new BooleanCell(false), "False boolean without style.");
            yield return RowForType(new BooleanCell(null, new BooleanStyle(null, Border.NONE, Color.SkyBlue)), "Null boolean styled.");
            yield return RowForType(new BooleanCell(true, new BooleanStyle(null, Border.NONE, Color.SkyBlue)), "True boolean styled.");
            yield return RowForType(new BooleanCell(false, new BooleanStyle(null, Border.NONE, Color.SkyBlue)), "False boolean styled.");
        }

        private static Row TimeSpanRow(string customTimeFormatString)
        {
            var customTimeValue = DateTime.FromOADate(0).AddHours(36).AddMinutes(13);
            var customTimeFormat = new DateTimeFormat(customTimeFormatString);
            var customDateStyle = new DateStyle(customTimeFormat, null, Border.NONE, null);
            return RowForType(new DateTimeCell(customTimeValue, customDateStyle), GetValueFormatComment(customTimeFormat, customTimeValue.ToString("u")));
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
    }
}