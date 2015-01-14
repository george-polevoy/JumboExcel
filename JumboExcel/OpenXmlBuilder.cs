using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JumboExcel.Formatting;
using JumboExcel.Structure;
using JumboExcel.Styling;
using Tuple = System.Tuple;

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
        private readonly static int BaseCustomFormatId = 165;

        /// <summary>
        /// Default border definitions.
        /// </summary>
        private readonly Border[] defaultBorders = 
        {
            new Border(new LeftBorder(),new RightBorder(),new TopBorder(),new BottomBorder(),new DiagonalBorder()),
            new Border(
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
        private readonly Font[] defaultFonts = { new Font(new FontSize { Val = 11 }, new Color { Rgb = new HexBinaryValue { Value = "000000" } }, new FontName { Val = "Calibri" }) };

        /// <summary>
        /// Writes <see cref="DocumentElement"/> hierarchy to the provided <see cref="Stream"/>.
        /// </summary>
        /// <param name="outputStream">Stream to write to.</param>
        /// <param name="worksheets">Worksheets for the excel document.</param>
        public void Write(Stream outputStream, IEnumerable<WorksheetElement> worksheets)
        {
            var worksheetParts = new List<Tuple<WorksheetPart, string>>();
            using (var spreadsheetDocument = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                var sharedStrings = new SharedElementCollection<string>();
                var sharedStylesCollection = new SharedElementCollection<CellStyleDefinition>();
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
        /// Writes styles part to the document.
        /// </summary>
        /// <param name="spreadsheetDocument">Document.</param>
        /// <param name="sharedStyles">Styles, accumulated during worksheet parts generation.</param>
        private void AddSharedStyles(SpreadsheetDocument spreadsheetDocument, SharedElementCollection<CellStyleDefinition> sharedStyles)
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
            var sharedStringTable = new SharedStringTable();
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
        private static void AddWorksheetReferences(List<Tuple<WorksheetPart, string>> worksheetParts, WorkbookPart workbookPart)
        {
            if (worksheetParts.Any())
            {
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
        private Stylesheet GenerateStyleSheet(SharedElementCollection<CellStyleDefinition> sharedStylesCollection)
        {
            var commonNumberFormats = CommonValueFormat.GetFormats().ToDictionary(i => i.FormatCode);
            var customNumberFormats = new SharedElementCollection<string>();
            var fontDefinitions = new SharedElementCollection<FontDefinition>();
            var fillDefinitions = new SharedElementCollection<System.Drawing.Color>();
            var borderDefinitions = new SharedElementCollection<BorderDefinition>();
            foreach (var style in sharedStylesCollection.GetAll())
            {
                if (style.Format != null)
                {
                    CommonValueFormat commonValueFormat;
                    if (!commonNumberFormats.TryGetValue(style.Format, out commonValueFormat))
                        customNumberFormats.GetOrAllocateElement(style.Format);
                }
                if (style.FontDefinition != null)
                    fontDefinitions.GetOrAllocateElement(style.FontDefinition);
                if (style.FillColor != null)
                    fillDefinitions.GetOrAllocateElement(style.FillColor.Value);
                if (style.BorderDefinition != BorderDefinition.None && style.BorderDefinition != BorderDefinition.All)
                    borderDefinitions.GetOrAllocateElement(style.BorderDefinition);
            }
            var numberFormatsCount = defaultNumberingFormats.Length + customNumberFormats.Count;
            return new Stylesheet(
                new NumberingFormats(defaultNumberingFormats.Concat(
                    customNumberFormats.GetAll().Select((formatCode, index) => new NumberingFormat {NumberFormatId = (uint) (BaseCustomFormatId + index), FormatCode = formatCode})
                    )) {Count = (uint) numberFormatsCount},
                new Fonts(defaultFonts.Concat(
                    fontDefinitions.GetAll().Select(CreateFont))),
                new Fills(defaultFills.Concat(
                    fillDefinitions.GetAll().Select(fill => new Fill(new PatternFill(
                        new ForegroundColor {Rgb = new HexBinaryValue {Value = ToHex(fill)}}
                        ) {PatternType = PatternValues.Solid}))
                    )),
                new Borders(defaultBorders.Concat(
                    borderDefinitions.GetAll().Select(CreateBorder)
                    )),
                new CellFormats(defaultCellFormats.Concat(
                    sharedStylesCollection.DequeueAll().Select(
                    style => CreateCellFormat(style, fontDefinitions, fillDefinitions, borderDefinitions, commonNumberFormats, customNumberFormats))
                    ))
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
            CellStyleDefinition style, SharedElementCollection<FontDefinition> fontDefinitions, SharedElementCollection<System.Drawing.Color> fillDefinitions,
            SharedElementCollection<BorderDefinition> borderDefinitions, IDictionary<string, CommonValueFormat> commonValueFormats,
            SharedElementCollection<string> customNumberFormats)
        {
            var cellFormat = new CellFormat();

            cellFormat.FontId = style.FontDefinition == null ? 0 : (uint) (fontDefinitions.GetElementIndex(style.FontDefinition) + defaultFonts.Length);
            if (style.FontDefinition != null)
            {
                cellFormat.ApplyFont = true;
            }

            cellFormat.FillId = style.FillColor == null
                ? 0
                : (uint) (fillDefinitions.GetElementIndex(style.FillColor.Value) + defaultFills.Length);
            if (style.FillColor != null)
                cellFormat.ApplyFill = true;

            switch (style.BorderDefinition)
            {
                case BorderDefinition.None:
                    cellFormat.BorderId = 0;
                    break;
                case BorderDefinition.All:
                    cellFormat.BorderId = 1;
                    cellFormat.ApplyBorder = true;
                    break;
                default:
                    cellFormat.BorderId = (uint) (borderDefinitions.GetElementIndex(style.BorderDefinition) + defaultBorders.Length);
                    cellFormat.ApplyBorder = true;
                    break;
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
                    cellFormat.NumberFormatId = (uint) (customNumberFormats.GetElementIndex(style.Format) + BaseCustomFormatId);
                }
            }

            return cellFormat;
        }

        /// <summary>
        /// Generates <see cref="Border"/> DOM element for provided <see cref="BorderDefinition"/>
        /// </summary>
        /// <param name="border">Border definition.</param>
        private static Border CreateBorder(BorderDefinition border)
        {
            var elements = new List<OpenXmlElement>();
            if (border.HasFlag(BorderDefinition.Left))
                elements.Add(new LeftBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin });
            if (border.HasFlag(BorderDefinition.Right))
                elements.Add(new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin });
            if (border.HasFlag(BorderDefinition.Top))
                elements.Add(new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin });
            if (border.HasFlag(BorderDefinition.Bottom))
                elements.Add(new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin });
            elements.Add(new DiagonalBorder());
            return new Border(elements);
        }

        /// <summary>
        /// Generates <see cref="Font"/> DOM element for provided <see cref="FontDefinition"/>.
        /// </summary>
        /// <param name="font">Font definition.</param>
        private static Font CreateFont(FontDefinition font)
        {
            var elements = new List<OpenXmlElement>();
            if (font.FontWeight == FontWeight.Bold)
                elements.Add(new Bold());
            if (font.FontSlope == FontSlope.Italic)
                elements.Add(new Italic());
            if (font.Typeface != null)
                elements.Add(new FontName { Val = font.Typeface });
            elements.Add(new FontSize { Val = (double)font.Size });
            var color = font.Color;
            elements.Add(new Color { Rgb = new HexBinaryValue { Value = ToHex(color) } });
            return new Font(elements);
        }
    }
}
