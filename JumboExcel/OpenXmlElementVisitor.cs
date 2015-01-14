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
    class OpenXmlElementVisitor : IElementVisitor
    {
        private readonly OpenXmlWriter writer;

        private readonly SharedElementCollection<string> sharedStringCollection;

        private readonly SharedCellStyleCollection cellStyleDefinitions;

        private readonly Cell sharedSampleCell = new Cell();

        private readonly CellValue sharedSampleCellValue = new CellValue();

        private readonly Cell sharedSampleNumberCell = new Cell {DataType = new EnumValue<CellValues>(CellValues.Number)};

        private Cell sharedSampleDateTimeCell;

        private Cell sharedSampleSharedStringCell;

        private Cell sharedSampleInlineStringCell;

        private readonly InlineString sharedSampleInlineString = new InlineString();

        private readonly Text sharedSampleText = new Text();

        private readonly List<Row> sampleRowOulineLevels = new List<Row> {new Row()};

        private int outlineLevel;
        
        Cell SharedSampleSharedStringCell
        {
            get
            {
                return sharedSampleSharedStringCell ?? (sharedSampleSharedStringCell = cellStyleDefinitions.AllocateSharedStringCell(new SharedStringStyleDefinition(null, BorderDefinition.None, null), CellValues.SharedString));
            }
        }

        Cell SharedSampleDateTimeCell
        {
            get
            {
                return sharedSampleDateTimeCell ?? (sharedSampleDateTimeCell = cellStyleDefinitions.AllocateDateCell(new DateStyleDefinition(DateTimeFormat.DateMmDdYy, null, BorderDefinition.None, null), CellValues.Number));
            }
        }

        Cell SharedSampleInlineStringCell
        {
            get
            {
                return sharedSampleInlineStringCell ?? (sharedSampleInlineStringCell = cellStyleDefinitions.AllocateStringCell(new StringStyleDefinition(null, BorderDefinition.None, null), CellValues.InlineString));
            }
        }

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
                    writer.WriteString(sharedStringCollection.AllocateElement(sharedStringElement.Value).ToString());
            }
        }
    }
}
