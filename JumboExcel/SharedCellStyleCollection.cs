using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using JumboExcel.Styling;

namespace JumboExcel
{
    class SharedCellStyleCollection
    {
        private readonly SharedElementCollection<CellStyleDefinition> cellStyleDefinitions;
        private readonly int cellFormatIndexCorrelation;

        private readonly List<Cell> allocatedSampleCells = new List<Cell>();
        
        public int Count { get { return allocatedSampleCells.Count; } }

        public SharedCellStyleCollection(SharedElementCollection<CellStyleDefinition> sharedElementCollection, int cellFormatIndexCorrelation)
        {
            cellStyleDefinitions = sharedElementCollection;
            this.cellFormatIndexCorrelation = cellFormatIndexCorrelation;
        }

        public Cell AllocateCell(CellStyleDefinition cellStyleDefinition, CellValues cellValueType)
        {
            var index = cellStyleDefinitions.AllocateElement(cellStyleDefinition);
            if (index < allocatedSampleCells.Count)
                return allocatedSampleCells[index];
            var cell = new Cell
            {
                DataType = new EnumValue<CellValues>(cellValueType),
                StyleIndex = (uint) (allocatedSampleCells.Count + cellFormatIndexCorrelation)
            };
            allocatedSampleCells.Add(cell);
            return cell;
        }
    }
}
