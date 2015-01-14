using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using JumboExcel.Formatting;
using JumboExcel.Structure;
using JumboExcel.Styling;

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
        /// Collection of shared strings. Accumulates strings provided in <see cref="SharedStringElement"/> instances.
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
        private readonly InlineString sharedSampleInlineString = new InlineString();

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        private readonly Text sharedSampleText = new Text();

        /// <summary>
        /// Collection of reusable component instances for wrtiting nested <see cref="RowGroupElement"/> elements.
        /// </summary>
        private readonly List<Row> sampleRowOulineLevels = new List<Row> {new Row()};

        /// <summary>
        /// Current outline level for row grouping for writing nested <see cref="RowGroupElement"/> elements.
        /// </summary>
        private int outlineLevel;

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        Cell SharedSampleSharedStringCell
        {
            get
            {
                return sharedSampleSharedStringCell ?? (sharedSampleSharedStringCell = cellStyleDefinitions.AllocateSharedStringCell(new StringStyleDefinition(null, BorderDefinition.None, null), CellValues.SharedString));
            }
        }

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        Cell SharedSampleDateTimeCell
        {
            get
            {
                return sharedSampleDateTimeCell ?? (sharedSampleDateTimeCell = cellStyleDefinitions.AllocateDateCell(new DateStyleDefinition(DateTimeFormat.DateMmDdYy, null, BorderDefinition.None, null), CellValues.Number));
            }
        }

        /// <summary>
        /// Reusable component instance.
        /// </summary>
        Cell SharedSampleInlineStringCell
        {
            get
            {
                return sharedSampleInlineStringCell ?? (sharedSampleInlineStringCell = cellStyleDefinitions.AllocateStringCell(new StringStyleDefinition(null, BorderDefinition.None, null), CellValues.InlineString));
            }
        }

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

        public void Visit(WorksheetElement worksheetElement)
        {
            using (new WriterScope(writer, new Worksheet()))
            {
                var worksheetParametersElement = worksheetElement.Parameters;
                if (worksheetParametersElement != null)
                {
                    WriteWorksheetParameters(worksheetParametersElement);
                }

                using (new WriterScope(writer, new SheetData()))
                {
                    var rowIndex = 0;
                    foreach (var rowElement in worksheetElement.RowsLevelElements)
                    {
                        rowElement.Accept(this);
                        rowIndex++;
                    }
                }
            }
        }

        private void WriteWorksheetParameters(WorksheetParametersElement worksheetParametersElement)
        {
            writer.WriteElement(new SheetProperties
            {
                OutlineProperties = new OutlineProperties {SummaryBelow = worksheetParametersElement.Belo, SummaryRight = worksheetParametersElement.Right}
            });

            if (worksheetParametersElement.ColumnElements != null)
            {
                writer.WriteElement(
                    new Columns(worksheetParametersElement.ColumnElements.Select(c => new Column {CustomWidth = true, Min = (uint) (c.Min + 1), Max = (uint) (c.Max + 1), Width = (double) c.Width})));
            }
        }

        public void Visit(RowElement rowElement)
        {
            var row = sampleRowOulineLevels[outlineLevel];
            using (new WriterScope(writer, row))
            {
                var columnIndex = 0;
                foreach (var cellElement in rowElement.CellElements)
                {
                    cellElement.Accept(this);
                    columnIndex++;
                }
            }
        }

        public void Visit(RowGroupElement rowGroupElement)
        {
            if (outlineLevel >= 255)
                throw new InvalidOperationException("Row ouline level overflow. Max row grouping level is 255.");
            
            outlineLevel ++;
            foreach (var rowElement in rowGroupElement.RowElements)
            {
                if (sampleRowOulineLevels.Count < outlineLevel + 1)
                {
                    var sampleRow = new Row { OutlineLevel = (byte)outlineLevel};
                    sampleRowOulineLevels.Add(sampleRow);
                }
                rowElement.Accept(this);
            }
            outlineLevel--;
        }

        public void VisitEmptyCell()
        {
            using (new WriterScope(writer, sharedSampleCell))
            {
            }
        }

        public void Visit(IntegerCellElement integerCellElement)
        {
            var sampleCell = integerCellElement.Style.CellStyleDefinition == null ? sharedSampleNumberCell : cellStyleDefinitions.AllocateNumberCell(integerCellElement.Style, CellValues.Number);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (integerCellElement.Value != null)
                    writer.WriteString(integerCellElement.Value.ToString());
            }
        }
        
        public void Visit(DecimalCellElement decimalCellElement)
        {
            var sampleCell = decimalCellElement.Style.CellStyleDefinition == null ? sharedSampleNumberCell : cellStyleDefinitions.AllocateNumberCell(decimalCellElement.Style, CellValues.Number);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (decimalCellElement.Value.HasValue)
                    writer.WriteString(decimalCellElement.Value.Value.ToString(CultureInfo.InvariantCulture));
            }
        }

        public void Visit(DateTimeCellElement dateTimeCellElement)
        {
            var sampleCell = dateTimeCellElement.Style.CellStyleDefinition == null ? SharedSampleDateTimeCell : cellStyleDefinitions.AllocateDateCell(dateTimeCellElement.Style, CellValues.Number);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (dateTimeCellElement.Value.HasValue)
                    writer.WriteString(dateTimeCellElement.Value.Value.ToOADate().ToString(CultureInfo.InvariantCulture));
            }
        }

        public void Visit(InlineStringElement inlineStringElement)
        {
            var sampleCell = inlineStringElement.Style.CellStyleDefinition == null ? SharedSampleInlineStringCell : cellStyleDefinitions.AllocateStringCell(inlineStringElement.Style, CellValues.InlineString);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleInlineString))
            using (new WriterScope(writer, sharedSampleText))
            {
                writer.WriteString(inlineStringElement.Value);
            }
        }

        public void Visit(SharedStringElement sharedStringElement)
        {
            var sampleCell = sharedStringElement.Style.CellStyleDefinition == null ? SharedSampleSharedStringCell : cellStyleDefinitions.AllocateSharedStringCell(sharedStringElement.Style, CellValues.SharedString);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (sharedStringElement.Value != null)
                    writer.WriteString(sharedStringCollection.GetOrAllocateElement(sharedStringElement.Value).ToString());
            }
        }

        public void Visit(BooleanCellElement booleanCellElement)
        {
            var sampleCell = booleanCellElement.Style.CellStyleDefinition == null ? sharedSampleBooleanCell : cellStyleDefinitions.AllocateBooleanCell(booleanCellElement.Style, CellValues.Boolean);
            using (new WriterScope(writer, sampleCell))
            using (new WriterScope(writer, sharedSampleCellValue))
            {
                if (booleanCellElement.Value.HasValue)
                    writer.WriteString(booleanCellElement.Value.Value ? "1" : "0");
            }
        }
    }
}
