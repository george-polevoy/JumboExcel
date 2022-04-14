using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JumboExcel.Formatting;
using JumboExcel.Styling;
using JumboExcel.Structure;
using Alignment = DocumentFormat.OpenXml.Spreadsheet.Alignment;
using Border = JumboExcel.Styling.Border;
using CellStyle = JumboExcel.Styling.CellStyle;
using Font = JumboExcel.Styling.Font;
using Tuple = System.Tuple;
using Worksheet = JumboExcel.Structure.Worksheet;

namespace JumboExcel
{
    /// <summary>
    /// Component for writing Excel documents.
    /// </summary>
    public class OpenXmlBuilder
    {
        /// <summary>
        /// Specifies the lowest safe value to use as a custom number format, not to interfere with built-in number formats.
        /// </summary>
        const int BASE_CUSTOM_FORMAT_ID = 165;

        /// <summary>
        /// Private constructor to disallow re-entrance of the methods.
        /// </summary>
        private OpenXmlBuilder() { }

        /// <summary>
        /// Default border definitions.
        /// </summary>
        private readonly DocumentFormat.OpenXml.Spreadsheet.Border[] defaultBorders = 
        {
            new DocumentFormat.OpenXml.Spreadsheet.Border(new LeftBorder(),new RightBorder(),new TopBorder(),new BottomBorder(),new DiagonalBorder()),
            new DocumentFormat.OpenXml.Spreadsheet.Border(
                new LeftBorder(new Color { Auto = true, }) { Style = BorderStyleValues.Thin },
                new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                new DiagonalBorder())
        };

        /// <summary>
        /// Default fill definitions.
        /// </summary>
        private readonly Fill[] defaultFills = { new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 }) };

        /// <summary>
        /// Default numbering format definitions.
        /// </summary>
        private readonly NumberingFormat[] defaultNumberingFormats = { new NumberingFormat { NumberFormatId = 0, FormatCode = "", } };

        /// <summary>
        /// Default cell format definitions.
        /// </summary>
        private readonly CellFormat[] defaultCellFormats = { new CellFormat { FontId = 0, FillId = 0, BorderId = 0, } };

        /// <summary>
        /// Default font definitions.
        /// </summary>
        private readonly DocumentFormat.OpenXml.Spreadsheet.Font[] defaultFonts = { new DocumentFormat.OpenXml.Spreadsheet.Font(new FontSize { Val = 11 }, new Color { Rgb = new HexBinaryValue { Value = "000000" } }, new FontName { Val = "Calibri" }) };

        /// <summary>
        /// Writes <see cref="DocumentElement"/> hierarchy to the provided <see cref="Stream"/>.
        /// This method is thread safe.
        /// </summary>
        /// <param name="outputStream">Stream to write to.</param>
        /// <param name="worksheets">Worksheets for the excel document.</param>
        public static void Write(Stream outputStream, IEnumerable<Worksheet> worksheets)
        {
            new OpenXmlBuilder().WriteInternal(outputStream, worksheets);
        }

        /// <summary>
        /// Writes <see cref="DocumentElement"/> hierarchy to the provided <see cref="Stream"/>.
        /// This method is thread safe.
        /// </summary>
        /// <param name="outputStream">Stream to write to.</param>
        /// <param name="worksheets">Worksheets for the excel document.</param>
        /// <returns>Returns an <see cref="IEnumerable{TProgress}"/> which must be enumerated to the end to complete the file generation.</returns>
        public static IEnumerable<TProgress> WriteWithProgress<TProgress>(Stream outputStream, IEnumerable<ProgressingWorksheet<TProgress>> worksheets)
        {
            return new OpenXmlBuilder().WriteWithProgressInternal(outputStream, worksheets);
        }

        /// <summary>
        /// Writes <see cref="DocumentElement"/> hierarchy to the provided <see cref="Stream"/>.
        /// This method must not be reentered.
        /// This method is not thread safe.
        /// </summary>
        /// <param name="outputStream">Stream to write to.</param>
        /// <param name="worksheets">Worksheets for the excel document.</param>
        private void WriteInternal(Stream outputStream, IEnumerable<Worksheet> worksheets)
        {
            var worksheetParts = new List<Tuple<WorksheetPart, string>>();
            using (var spreadsheetDocument = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                var sharedStrings = new SharedElementCollection<string>();
                var sharedStylesCollection = new SharedElementCollection<CellStyle>();
                var sharedStyles = new SharedCellStyleCollection(sharedStylesCollection, defaultCellFormats.Length);
                foreach (var worksheetElement in worksheets)
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetParts.Add(Tuple.Create(worksheetPart, worksheetElement.Name));
                    using (var worksheetPartWriter = OpenXmlWriter.Create(worksheetPart))
                    {
                        var elementVisitor = new OpenXmlElementVisitor(worksheetPartWriter, sharedStrings, sharedStyles);
                        worksheetElement.Accept(elementVisitor);
                    }
                }
                AddWorksheetReferences(worksheetParts, workbookPart);
                AddSharedStrings(spreadsheetDocument, sharedStrings);
                AddSharedStyles(spreadsheetDocument, sharedStylesCollection);
            }
        }

        /// <summary>
        /// Writes <see cref="DocumentElement"/> hierarchy to the provided <see cref="Stream"/>.
        /// This method must not be reentered.
        /// This method is not thread safe.
        /// </summary>
        /// <param name="outputStream">Stream to write to.</param>
        /// <param name="worksheets">Worksheets for the excel document.</param>
        /// <returns>Returns an <see cref="IEnumerable{TProgress}"/> which must be enumerated to complete the file generation.</returns>
        private IEnumerable<TProgress> WriteWithProgressInternal<TProgress>(Stream outputStream, IEnumerable<ProgressingWorksheet<TProgress>> worksheets)
        {
            var worksheetParts = new List<Tuple<WorksheetPart, string>>();
            using (var spreadsheetDocument = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                var sharedStrings = new SharedElementCollection<string>();
                var sharedStylesCollection = new SharedElementCollection<CellStyle>();
                var sharedStyles = new SharedCellStyleCollection(sharedStylesCollection, defaultCellFormats.Length);
                foreach (var worksheetElement in worksheets)
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetParts.Add(Tuple.Create(worksheetPart, worksheetElement.Name));
                    using (var worksheetPartWriter = OpenXmlWriter.Create(worksheetPart))
                    {
                        var elementVisitor = new OpenXmlElementVisitor(worksheetPartWriter, sharedStrings, sharedStyles);
                        foreach (var p in elementVisitor.VisitProgressingWorksheet(worksheetElement))
                        {
                            yield return p;
                        }
                    }
                }
                AddWorksheetReferences(worksheetParts, workbookPart);
                AddSharedStrings(spreadsheetDocument, sharedStrings);
                AddSharedStyles(spreadsheetDocument, sharedStylesCollection);
            }
        }

        /// <summary>
        /// Writes styles part to the document.
        /// </summary>
        /// <param name="spreadsheetDocument">Document.</param>
        /// <param name="sharedStyles">Styles, accumulated during worksheet parts generation.</param>
        private void AddSharedStyles(SpreadsheetDocument spreadsheetDocument, SharedElementCollection<CellStyle> sharedStyles)
        {
            if (sharedStyles.Count <= 0) return;
            var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = GenerateStyleSheet(sharedStyles);
        }

        /// <summary>
        /// Writes shared string part to the document.
        /// </summary>
        /// <param name="spreadsheetDocument">Docuement.</param>
        /// <param name="sharedStrings">Shared strings, accumulated during worksheet parts generation.</param>
        private static void AddSharedStrings(SpreadsheetDocument spreadsheetDocument, SharedElementCollection<string> sharedStrings)
        {
            if (sharedStrings.Count <= 0) return;
            var sharedStringTablePart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            var sharedStringTable = new SharedStringTable()
            {
                Count = new UInt32Value((uint)sharedStrings.Count),
                UniqueCount = new UInt32Value((uint)sharedStrings.Count)
            };
            sharedStringTablePart.SharedStringTable = sharedStringTable;
            foreach (var text in sharedStrings.DequeueAll())
            {
                sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            }
        }

        /// <summary>
        /// Establishes relations between written worksheet parts and named Sheets.
        /// </summary>
        /// <param name="worksheetParts">Named worksheet parts.</param>
        /// <param name="workbookPart">Workbook part, representing entire document.</param>
        private static void AddWorksheetReferences(ICollection<Tuple<WorksheetPart, string>> worksheetParts, WorkbookPart workbookPart)
        {
            var duplicateNames = worksheetParts.GroupBy(i => i.Item2).Select(g => new {name = g.Key, count = g.Count()}).Where(g => g.count > 1).ToList();
            if (duplicateNames.Any())
            {
                throw new InvalidOperationException("Duplicate worksheet name encountered. Checkout Exception Data") {Data = { {"Duplicates", duplicateNames}}};
            }

            if (worksheetParts.Count <= 0) return;
            using (var workbookPartWriter = OpenXmlWriter.Create(workbookPart))
            {
                workbookPartWriter.WriteElement(new Workbook(new Sheets(worksheetParts.Select((item, index) =>
                    new Sheet
                    {
                        Name = item.Item2,
                        SheetId = (uint)index + 1,
                        Id = workbookPart.GetIdOfPart(item.Item1)
                    }))));
            }
        }

        /// <summary>
        /// Converts color value to haxadecimal RGB value.
        /// </summary>
        /// <param name="color">Color.</param>
        /// <returns>Returns hexadecimal string representation of the color.</returns>
        private static string ToHex(System.Drawing.Color color)
        {
            return (color.ToArgb() & 0x00FFFFFF).ToString("X");
        }

        /// <summary>
        /// Generates stylesheet.
        /// </summary>
        /// <param name="sharedStylesCollection">Shared styles, accumulated during worksheet parts generation.</param>
        /// <returns>Returns DOM stylesheet part for the entire document.</returns>
        private Stylesheet GenerateStyleSheet(SharedElementCollection<CellStyle> sharedStylesCollection)
        {
            var commonNumberFormats = CommonValueFormat.GetFormats().ToDictionary(i => i.FormatCode);
            var customNumberFormats = new SharedElementCollection<string>();
            var fontDefinitions = new SharedElementCollection<Font>();
            var fillDefinitions = new SharedElementCollection<System.Drawing.Color>();
            var borderDefinitions = new SharedElementCollection<Border>();
            foreach (var style in sharedStylesCollection.GetAll())
            {
                if (style.Format != null)
                {
                    CommonValueFormat commonValueFormat;
                    if (!commonNumberFormats.TryGetValue(style.Format, out commonValueFormat))
                        customNumberFormats.GetOrAllocateElement(style.Format);
                }
                if (style.Font != null)
                    fontDefinitions.GetOrAllocateElement(style.Font);
                if (style.FillColor != null)
                    fillDefinitions.GetOrAllocateElement(style.FillColor.Value);
                if (style.Border != Border.NONE && style.Border != Border.ALL)
                    borderDefinitions.GetOrAllocateElement(style.Border);
            }
            var numberFormatsCount = defaultNumberingFormats.Length + customNumberFormats.Count;
            var fontDefinitionsCount = defaultFonts.Length + fontDefinitions.Count;
            var fillDefinitionsCount = defaultFills.Length + fillDefinitions.Count;
            var borderDefinitionsCount = defaultBorders.Length + borderDefinitions.Count;
            var cellFormatCount = sharedStylesCollection.Count + sharedStylesCollection.Count;
            return new Stylesheet(
                new NumberingFormats(defaultNumberingFormats.Concat(
                    customNumberFormats.GetAll().Select((formatCode, index) => new NumberingFormat {NumberFormatId = (uint) (BASE_CUSTOM_FORMAT_ID + index), FormatCode = formatCode})
                    )) {Count = (uint) numberFormatsCount},
                new Fonts(defaultFonts.Concat(
                    fontDefinitions.GetAll().Select(CreateFont))) { Count = (uint)fontDefinitionsCount },
                new Fills(defaultFills.Concat(
                    fillDefinitions.GetAll().Select(fill => new Fill(new PatternFill(
                        new ForegroundColor {Rgb = new HexBinaryValue {Value = ToHex(fill)}}
                        ) {PatternType = PatternValues.Solid}))
                    )) { Count = (uint)fillDefinitionsCount },
                new Borders(defaultBorders.Concat(
                    borderDefinitions.GetAll().Select(CreateBorder)
                    )) { Count = (uint)borderDefinitionsCount },
                new CellFormats(defaultCellFormats.Concat(
                    sharedStylesCollection.DequeueAll().Select(
                    style => CreateCellFormat(style, fontDefinitions, fillDefinitions, borderDefinitions, commonNumberFormats, customNumberFormats))
                    )) { Count = (uint)cellFormatCount }
                );
        }

        /// <summary>
        /// Generates DOM <see cref="CellFormat"/> element.
        /// </summary>
        /// <param name="style">Style definition for which the DOM element is generated.</param>
        /// <param name="fontDefinitions">Accumulated shared component.</param>
        /// <param name="fillDefinitions">Accumulated shared component.</param>
        /// <param name="borderDefinitions">Accumulated shared component.</param>
        /// <param name="commonValueFormats">Accumulated shared component.</param>
        /// <param name="customNumberFormats">Accumulated shared component.</param>
        /// <returns>Returns the new element, corresponding to the <paramref name="style"/> provided.</returns>
        private CellFormat CreateCellFormat(
            CellStyle style, SharedElementCollection<Font> fontDefinitions, SharedElementCollection<System.Drawing.Color> fillDefinitions,
            SharedElementCollection<Border> borderDefinitions, IDictionary<string, CommonValueFormat> commonValueFormats,
            SharedElementCollection<string> customNumberFormats)
        {
            var cellFormat = new CellFormat();

            cellFormat.FontId = style.Font == null ? 0 : (uint) (fontDefinitions.GetElementIndex(style.Font) + defaultFonts.Length);
            if (style.Font != null)
            {
                cellFormat.ApplyFont = true;
            }

            cellFormat.FillId = style.FillColor == null
                ? 0
                : (uint) (fillDefinitions.GetElementIndex(style.FillColor.Value) + defaultFills.Length);
            if (style.FillColor != null)
                cellFormat.ApplyFill = true;

            switch (style.Border)
            {
                case Border.NONE:
                    cellFormat.BorderId = 0;
                    break;
                case Border.ALL:
                    cellFormat.BorderId = 1;
                    cellFormat.ApplyBorder = true;
                    break;
                default:
                    cellFormat.BorderId = (uint) (borderDefinitions.GetElementIndex(style.Border) + defaultBorders.Length);
                    cellFormat.ApplyBorder = true;
                    break;
            }
            if (style.Alignment != null)
            {
                var alignment = style.Alignment;
                cellFormat.Alignment = new Alignment
                {
                    Horizontal = (HorizontalAlignmentValues)alignment.Horizontal,
                    Vertical = (VerticalAlignmentValues)alignment.Vertical,
                    TextRotation = (uint)alignment.TextRotation,
                    WrapText = alignment.WrapText
                };
                cellFormat.ApplyAlignment = true;
            }
            if (style.Format != null)
            {
                cellFormat.ApplyNumberFormat = true;
                CommonValueFormat commonValueFormat;
                if (commonValueFormats.TryGetValue(style.Format, out commonValueFormat))
                {
                    cellFormat.NumberFormatId = (uint) commonValueFormat.Id;
                }
                else
                {
                    cellFormat.NumberFormatId = (uint) (customNumberFormats.GetElementIndex(style.Format) + BASE_CUSTOM_FORMAT_ID);
                }
            }

            return cellFormat;
        }

        /// <summary>
        /// Generates <see cref="DocumentFormat.OpenXml.Spreadsheet.Border"/> DOM element for provided <see cref="Border"/>
        /// </summary>
        /// <param name="border">Border definition.</param>
        private static DocumentFormat.OpenXml.Spreadsheet.Border CreateBorder(Border border)
        {
            var elements = new List<OpenXmlElement>();
            if (border.HasFlag(Border.LEFT))
                elements.Add(new LeftBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin });
            if (border.HasFlag(Border.RIGHT))
                elements.Add(new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin });
            if (border.HasFlag(Border.TOP))
                elements.Add(new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin });
            if (border.HasFlag(Border.BOTTOM))
                elements.Add(new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin });
            elements.Add(new DiagonalBorder());
            return new DocumentFormat.OpenXml.Spreadsheet.Border(elements);
        }

        /// <summary>
        /// Generates <see cref="DocumentFormat.OpenXml.Spreadsheet.Font"/> DOM element for provided <see cref="Font"/>.
        /// </summary>
        /// <param name="font">Font definition.</param>
        private static DocumentFormat.OpenXml.Spreadsheet.Font CreateFont(Font font)
        {
            var elements = new List<OpenXmlElement>();
            if (font.FontWeight == FontWeight.BOLD)
                elements.Add(new Bold());
            if (font.FontSlope == FontSlope.ITALIC)
                elements.Add(new Italic());
            if (font.Typeface != null)
                elements.Add(new FontName { Val = font.Typeface });
            elements.Add(new FontSize { Val = (double)font.Size });
            var color = font.Color;
            elements.Add(new Color { Rgb = new HexBinaryValue { Value = ToHex(color) } });
            return new DocumentFormat.OpenXml.Spreadsheet.Font(elements);
        }
    }
}
