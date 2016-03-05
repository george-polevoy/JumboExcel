using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using JumboExcel.Formatting;
using JumboExcel.Styling;
using JumboExcel.Structure;
using Border = JumboExcel.Styling.Border;
using InlineString = JumboExcel.Structure.InlineString;
using Row = JumboExcel.Structure.Row;
using Worksheet = JumboExcel.Structure.Worksheet;

namespace JumboExcel
{
    /// <summary>
    /// Generates worksheet contents according to provided <see cref="DocumentElement"/> elements.
    /// </summary>
    class OpenXmlElementVisitor : IElementVisitor
    {
        /// <summary>
        /// Component for writing OpenXml document.
        /// </summary>
        private readonly OpenXmlWriter writer;

        /// <summary>
        /// Collection of shared strings. Accumulates strings provided in <see cref="SharedString"/> instances.
        /// </summary>
        private readonly SharedElementCollection<string> sharedStringCollection;

        /// <summary>
        /// Collection of shared cell styles.
        /// </summary>
        private readonly SharedCellStyleCollection cellStyleDefinitions;

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private readonly Cell sharedSampleCell = new Cell();

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private readonly CellValue sharedSampleCellValue = new CellValue();

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private readonly Cell sharedSampleNumberCell = new Cell {DataType = new EnumValue<CellValues>(CellValues.Number)};

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private readonly Cell sharedSampleBooleanCell = new Cell { DataType = new EnumValue<CellValues>(CellValues.Boolean) };

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private Cell sharedSampleDateTimeCell;

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private Cell sharedSampleSharedStringCell;

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private Cell sharedSampleInlineStringCell;

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private readonly DocumentFormat.OpenXml.Spreadsheet.InlineString sharedSampleInlineString = new DocumentFormat.OpenXml.Spreadsheet.InlineString();

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private readonly Text sharedSampleText = new Text();

        /// <summary>
        /// Current reusable row element.
        /// </summary>
        DocumentFormat.OpenXml.Spreadsheet.Row currentLevelSampleRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

        /// <summary>
        /// Collection of reusable elements for nested <see cref="RowGroup"/> elements.
        /// </summary>
        private readonly List<RowSlot> sampleRowOulineLevels = new List<RowSlot> { new RowSlot() };

        class RowSlot
        {
            public DocumentFormat.OpenXml.Spreadsheet.Row regularRow;
            public DocumentFormat.OpenXml.Spreadsheet.Row hiddenRow;
        }

        /// <summary>
        /// Current outline level for row grouping for writing nested <see cref="RowGroup"/> elements.
        /// </summary>
        private int outlineLevel;

        /// <summary>
        /// Current row's zero based index.
        /// </summary>
        private int rowIndex;

        /// <summary>
        /// Current column's zero based index.
        /// </summary>
        private int columnIndex;

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        Cell SharedSampleSharedStringCell
        {
            get
            {
                return sharedSampleSharedStringCell ?? (sharedSampleSharedStringCell = cellStyleDefinitions.AllocateSharedStringCell(new StringStyle(null, Border.NONE, null), CellValues.SharedString));
            }
        }

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        Cell SharedSampleDateTimeCell
        {
            get
            {
                return sharedSampleDateTimeCell ?? (sharedSampleDateTimeCell = cellStyleDefinitions.AllocateDateCell(new DateStyle(DateTimeFormat.DateMmDdYy, null, Border.NONE, null), CellValues.Number));
            }
        }

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        Cell SharedSampleInlineStringCell
        {
            get
            {
                return sharedSampleInlineStringCell ?? (sharedSampleInlineStringCell = cellStyleDefinitions.AllocateStringCell(new StringStyle(null, Border.NONE, null), CellValues.InlineString));
            }
        }

        /// <summary>
        /// List of the cell merge information entries, aggregated during visitation of the worksheet structure.
        /// </summary>
        readonly List<CellMergeInfo> mergedCells = new List<CellMergeInfo>();

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="writer">Writer for the worksheet part.</param>
        /// <param name="sharedStringCollection">Shared string collection for the workbook (shared for the entire Excel document). Items accumulated during generation are later used in shared string part generation.</param>
        /// <param name="cellStyleDefinitions">Shared cell style collection for the workbook (shared for the entire Excel document). Items accumulated in this collection are later used in stylesheet part generation.</param>
        public OpenXmlElementVisitor(OpenXmlWriter writer, SharedElementCollection<string> sharedStringCollection, SharedCellStyleCollection cellStyleDefinitions)
        {
            this.writer = writer;
            this.sharedStringCollection = sharedStringCollection;
            this.cellStyleDefinitions = cellStyleDefinitions;
        }

        public IEnumerable<TProgress> VisitProgressingWorksheet<TProgress>(params ProgressingWorksheet<TProgress>[] args)
        {
            foreach (var worksheet in args)
            {
                using (new WriterScope(writer, new DocumentFormat.OpenXml.Spreadsheet.Worksheet()))
                {
                    var worksheetParametersElement = worksheet.Parameters;
                    if (worksheetParametersElement != null)
                    {
                        WriteWorksheetParameters(worksheetParametersElement);
                    }
                    using (new WriterScope(writer, new SheetData()))
                    {
                        rowIndex = 0;
                        foreach (var p in worksheet.RowGenerator(rows =>
                        {
                            foreach (var rowLevelElement in rows)
                            {
                                rowLevelElement.Accept(this);
                                rowIndex++;
                            }
                        }))
                        {
                            yield return p;
                        }
                    }
                }
            }
        }

        public void Visit(Worksheet worksheet)
        {
            using (new WriterScope(writer, new DocumentFormat.OpenXml.Spreadsheet.Worksheet()))
            {
                var worksheetParametersElement = worksheet.Parameters;
                if (worksheetParametersElement != null)
                {
                    WriteWorksheetParameters(worksheetParametersElement);
                }

                using (new WriterScope(writer, new SheetData()))
                {
                    foreach (var rowElement in worksheet.RowsLevelElements)
                    {
                        rowElement.Accept(this);
                    }
                }

                if (mergedCells.Count > 0)
                {
                    WriteMergedCells();
                }
            }
        }

        /// <summary>
        /// Writes merged cells information.
        /// </summary>
        private void WriteMergedCells()
        {
            var ranges = mergedCells.
                GroupBy(i => i.UpperLeft).
                Select(g => new { upperLeft = g.Key, lowerRight = new { Row = g.Max(i => i.LowerRight.Row), Column = g.Max(i => i.LowerRight.Column) } }).
                Where(range => range.upperLeft.Column != range.lowerRight.Column || range.upperLeft.Row != range.lowerRight.Row).ToList();

            if (ranges.Count == 0)
                return;

            foreach (var badRange in ranges
                .Where(range => range.upperLeft.Column > range.lowerRight.Column || range.upperLeft.Row > range.lowerRight.Row || range.upperLeft.Column < 0 || range.upperLeft.Row < 0))
                throw new InvalidOperationException(string.Format("Incompatible range. upperLeft: {0}, lowerRight: {1}", badRange.upperLeft, badRange.lowerRight));

            using (new WriterScope(writer, new MergeCells()))
                foreach (var rangeSpec in ranges.Select(range => Range(CellRef(range.upperLeft.Row, range.upperLeft.Column), CellRef(range.lowerRight.Row, range.lowerRight.Column))))
                    writer.WriteElement(new MergeCell { Reference = rangeSpec });
        }

        /// <summary>
        /// Writes worksheet parameters.
        /// </summary>
        /// <param name="worksheetParametersElement">Worksheet parameters.</param>
        private void WriteWorksheetParameters(WorksheetParametersElement worksheetParametersElement)
        {
            writer.WriteElement(new SheetProperties
            {
                OutlineProperties = new OutlineProperties { SummaryBelow = worksheetParametersElement.Belo, SummaryRight = worksheetParametersElement.Right }
            });

            if (worksheetParametersElement.ColumnConfigurations != null)
            {
                writer.WriteElement(
                    new Columns(worksheetParametersElement.ColumnConfigurations.Select(c => new Column
                    {
                        CustomWidth = true,
                        Min = (uint)(c.Min + 1),
                        Max = (uint)(c.Max + 1),
                        Width = (double)c.Width,
                        OutlineLevel = (byte)c.OutlineLevel
                    })));
            }
        }

        /// <summary>
        /// Specified if the last row level element was a simple row (not a row group).
        /// </summary>
        bool lastRowLevelElementIsSimpleRow;

        public void Visit(Row rowElement)
        {
            var row = currentLevelSampleRow;
            using (new WriterScope(writer, row))
            {
                columnIndex = 0;
                foreach (var cellElement in rowElement.Cells)
                {
                    cellElement.Accept(this);
                    columnIndex++;
                }
            }
            lastRowLevelElementIsSimpleRow = true;
            rowIndex++;
        }

        public void Visit(RowGroup rowGroup)
        {
            if (!lastRowLevelElementIsSimpleRow)
                throw new InvalidOperationException("Row group must follow a simple row element at the outline level.");
            if (outlineLevel >= 255)
                throw new InvalidOperationException("Row ouline level overflow. Max row grouping level is 255.");

            var prevSampleRow = currentLevelSampleRow;
            outlineLevel++;
            var groupChildCount = 0;
            foreach (var rowElement in rowGroup.RowElements)
            {
                RowSlot slot;
                if (sampleRowOulineLevels.Count < outlineLevel + 1)
                {
                    slot = new RowSlot
                    {
                        regularRow = new DocumentFormat.OpenXml.Spreadsheet.Row { OutlineLevel = (byte)outlineLevel },
                        hiddenRow = new DocumentFormat.OpenXml.Spreadsheet.Row { OutlineLevel = (byte)outlineLevel, Hidden = outlineLevel > 0 }
                    };
                    sampleRowOulineLevels.Add(slot);
                }
                else
                {
                    slot = sampleRowOulineLevels[outlineLevel];
                }

                currentLevelSampleRow = rowGroup.Collapse ? slot.hiddenRow : slot.regularRow;

                rowElement.Accept(this);
                groupChildCount++;
            }

            if (groupChildCount < 1)
                throw new InvalidOperationException("Empty group detected.");

            lastRowLevelElementIsSimpleRow = false;
            outlineLevel--;
            currentLevelSampleRow = prevSampleRow;
        }

        public void VisitEmptyCell()
        {
            using (new WriterScope(writer, sharedSampleCell))
            {
            }
        }

        public void Visit(IntegerCell integerCell)
        {
            var sampleCell = integerCell.Style.cellStyle == null ? sharedSampleNumberCell : cellStyleDefinitions.AllocateNumberCell(integerCell.Style, CellValues.Number);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (integerCell.Value != null)
                    writer.WriteString(integerCell.Value.ToString());
            }
        }

        public void Visit(DecimalCell decimalCell)
        {
            var sampleCell = decimalCell.Style.cellStyle == null ? sharedSampleNumberCell : cellStyleDefinitions.AllocateNumberCell(decimalCell.Style, CellValues.Number);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (decimalCell.Value.HasValue)
                    writer.WriteString(decimalCell.Value.Value.ToString(CultureInfo.InvariantCulture));
            }
        }

        public void Visit(DateTimeCell dateTimeCell)
        {
            var sampleCell = dateTimeCell.Style.cellStyle == null ? SharedSampleDateTimeCell : cellStyleDefinitions.AllocateDateCell(dateTimeCell.Style, CellValues.Number);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (dateTimeCell.Value.HasValue)
                    writer.WriteString(dateTimeCell.Value.Value.ToOADate().ToString(CultureInfo.InvariantCulture));
            }
        }

        public void Visit(InlineString inlineString)
        {
            var sampleCell = inlineString.Style.cellStyle == null ? SharedSampleInlineStringCell : cellStyleDefinitions.AllocateStringCell(inlineString.Style, CellValues.InlineString);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleInlineString))
            using (new WriterScope(writer, sharedSampleText))
            {
                writer.WriteString(inlineString.Value);
            }
        }

        public void Visit(SharedString sharedString)
        {
            var sampleCell = sharedString.Style.cellStyle == null ? SharedSampleSharedStringCell : cellStyleDefinitions.AllocateSharedStringCell(sharedString.Style, CellValues.SharedString);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (sharedString.Value != null)
                    writer.WriteString(sharedStringCollection.GetOrAllocateElement(sharedString.Value).ToString());
            }
        }

        public void Visit(BooleanCell booleanCell)
        {
            var sampleCell = booleanCell.Style.cellStyle == null ? sharedSampleBooleanCell : cellStyleDefinitions.AllocateBooleanCell(booleanCell.Style, CellValues.Boolean);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (booleanCell.Value.HasValue)
                    writer.WriteString(booleanCell.Value.Value ? "1" : "0");
            }
        }

        public void Visit(CellMerger cellMerger)
        {
            cellMerger.InnerElement.Accept(this);
            mergedCells.Add(new CellMergeInfo(new CellRef(cellMerger.GetAnchorRow(rowIndex), cellMerger.GetAnchorColumn(columnIndex)), new CellRef(rowIndex, columnIndex)));
        }

        /// <summary>
        /// Gets the Excel representation of the cell range relative to the worksheet.
        /// </summary>
        /// <param name="upperLeft">Upper left cell's address in Excel format.</param>
        /// <param name="lowerRight">Lower Right cell's address in Excel format.</param>
        static string Range(string upperLeft, string lowerRight)
        {
            return upperLeft + ":" + lowerRight;
        }

        /// <summary>
        /// Gets the Excel representation of the cell address relative to the worksheet.
        /// </summary>
        /// <param name="row">Zero based row index.</param>
        /// <param name="column">Zero based column index.</param>
        /// <returns>Returns the address of the cell in the Excel format <code>ABC123</code></returns>
        static string CellRef(int row, int column)
        {
            return ExcelHelper.CellRef(row, column);
        }
    }
}
