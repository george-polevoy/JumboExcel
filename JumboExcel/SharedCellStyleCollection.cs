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
        
        public SharedCellStyleCollection(SharedElementCollection<CellStyleDefinition> sharedElementCollection, int cellFormatIndexCorrelation)
        {
            cellStyleDefinitions = sharedElementCollection;
            this.cellFormatIndexCorrelation = cellFormatIndexCorrelation;
        }

        public Cell AllocateCell(CellStyleDefinition cellStyleDefinition, CellValues cellValueType, bool applyType = true)
        {
            var index = cellStyleDefinitions.AllocateElement(cellStyleDefinition);
            if (index < allocatedSampleCells.Count)
                return allocatedSampleCells[index];
            var cell = new Cell();
            cell.StyleIndex = (uint) (allocatedSampleCells.Count + cellFormatIndexCorrelation);
            if (applyType)
                cell.DataType = new EnumValue<CellValues>(cellValueType);
            allocatedSampleCells.Add(cell);
            return cell;
        }
    }
}
