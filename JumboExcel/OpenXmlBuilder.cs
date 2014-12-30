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
    public class OpenXmlBuilder
    {
        private readonly Border[] defaultBorders = {new Border(
            new LeftBorder(),
            new RightBorder(),
            new TopBorder(),
            new BottomBorder(),
            new DiagonalBorder()),
            new Border(
                new LeftBorder(
                    new Color { Auto = true, }
                    ) { Style = BorderStyleValues.Thin },
                new RightBorder(
                    new Color { Auto = true }
                    ) { Style = BorderStyleValues.Thin },
                new TopBorder(
                    new Color { Auto = true }
                    ) { Style = BorderStyleValues.Thin },
                new BottomBorder(
                    new Color { Auto = true }
                    ) { Style = BorderStyleValues.Thin },
                new DiagonalBorder())};

        private readonly Fill[] defaultFills = { new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 }) };
        private readonly NumberingFormat[] defaultNumberingFormats = { new NumberingFormat { NumberFormatId = 0, FormatCode = "", } };
        private CellFormat[] defaultCellFormats = { new CellFormat { FontId = 0, FillId = 0, BorderId = 0, } };
        private readonly Font[] defaultFonts = new[] { new Font(new FontSize { Val = 11 }, new Color { Rgb = new HexBinaryValue { Value = "000000" } }, new FontName { Val = "Calibri" }) };

        private int CellFormatIndexCorrelation()
        {
            return defaultCellFormats.Length;
        }

        public void Write(Stream outputStream, IEnumerable<WorksheetElement> worksheets)
        {
            var worksheetParts = new List<Tuple<WorksheetPart, string>>();
            using (var spreadsheetDocument = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                var sharedStrings = new SharedElementCollection<string>();
                var sharedStylesCollection = new SharedElementCollection<CellStyleDefinition>();
                var sharedStyles = new SharedCellStyleCollection(sharedStylesCollection, CellFormatIndexCorrelation());

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

                AddWorksheetReferences(worksheetParts, workbookPart, spreadsheetDocument);

                AddSharedStrings(spreadsheetDocument, sharedStrings);

                AddSharedStyles(spreadsheetDocument, sharedStylesCollection);
            }
        }

        private void AddSharedStyles(SpreadsheetDocument spreadsheetDocument, SharedElementCollection<CellStyleDefinition> sharedStyles)
        {
            if (sharedStyles.Count <= 0) return;
            var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = GenerateStyleSheet(sharedStyles);
        }

        private void AddSharedStrings(SpreadsheetDocument spreadsheetDocument, SharedElementCollection<string> sharedStrings)
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

        private static void AddWorksheetReferences(List<Tuple<WorksheetPart, string>> worksheetParts,
            WorkbookPart workbookPart,
            SpreadsheetDocument spreadsheetDocument)
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
                            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(item.Item1)
                        }))));
                }
            }
        }

        private string ToHex(System.Drawing.Color color)
        {
            return (color.ToArgb() & 0x00FFFFFF).ToString("X");
        }

        private Stylesheet GenerateStyleSheet(SharedElementCollection<CellStyleDefinition> sharedStylesCollection)
        {
            var commonNumberFormats = CommonValueFormat.GetFormats().ToDictionary(i => i.FormatCode);

            const int baseCustomFormatId = 165;

            var customNumberFormats = new SharedElementCollection<string>();

            var fontDefinitions = new SharedElementCollection<FontDefinition>();
            var fillDefinitions = new SharedElementCollection<System.Drawing.Color>();
            var borderDefinitions = new SharedElementCollection<BorderDefinition>();

            foreach (var style in sharedStylesCollection.GetAll())
            {
                if (style.Format != null)
                    customNumberFormats.AllocateElement(style.Format);
                if (style.FontDefinition != null)
                    fontDefinitions.AllocateElement(style.FontDefinition);
                if (style.FillColor != null)
                    fillDefinitions.AllocateElement(style.FillColor.Value);
                if (style.BorderDefinition != BorderDefinition.None && style.BorderDefinition != BorderDefinition.All)
                    borderDefinitions.AllocateElement(style.BorderDefinition);
            }

            var numberFormatsCount = defaultNumberingFormats.Length + customNumberFormats.Count;

            return new Stylesheet(
                new NumberingFormats(defaultNumberingFormats.Concat(
                    customNumberFormats.GetAll().Select((formatCode, index) => new NumberingFormat {NumberFormatId = (uint) (baseCustomFormatId + index), FormatCode = formatCode})
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
                    style => CreateCellFormat(style, fontDefinitions, fillDefinitions, borderDefinitions, commonNumberFormats, customNumberFormats, baseCustomFormatId))
                    ))
                );
        }

        private CellFormat CreateCellFormat(
            CellStyleDefinition style, SharedElementCollection<FontDefinition> fontDefinitions, SharedElementCollection<System.Drawing.Color> fillDefinitions,
            SharedElementCollection<BorderDefinition> borderDefinitions, Dictionary<string, CommonValueFormat> commonNumberFormats,
            SharedElementCollection<string> customNumberFormats, int baseCustomFormatId)
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
                if (commonNumberFormats.TryGetValue(style.Format, out commonValueFormat))
                {
                    cellFormat.NumberFormatId = (uint) commonValueFormat.Id;
                }
                else
                {
                    cellFormat.NumberFormatId = (uint) (customNumberFormats.GetElementIndex(style.Format) + baseCustomFormatId);
                }
            }

            return cellFormat;
        }

        private Border CreateBorder(BorderDefinition border)
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

        private Font CreateFont(FontDefinition f)
        {
            var elements = new List<OpenXmlElement>();
            if (f.FontWeight == FontWeight.Bold)
                elements.Add(new Bold());
            if (f.FontSlope == FontSlope.Italic)
                elements.Add(new Italic());
            if (f.Typeface != null)
                elements.Add(new FontName { Val = f.Typeface });
            elements.Add(new FontSize { Val = (double)f.Size });
            var color = f.Color;
            elements.Add(new Color { Rgb = new HexBinaryValue { Value = ToHex(color) } });
            return new Font(elements);
        }
    }
}
