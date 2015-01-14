using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using JumboExcel.Styling;

namespace JumboExcel
{
    /// <summary>
    /// Represents single allocated style and different cell definitions for various types.
    /// </summary>
    class CellStyleOption
    {
        public CellStyleOption(int index)
        {
            Index = index;
        }

        /// <summary>
        /// Index of allocated style in cell style collection.
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// Cached cell instance for represented style, for <see cref="NumberStyleDefinition"/>
        /// </summary>
        public Cell NumberCell { get; set; }

        /// <summary>
        /// Cached cell instance for represented style, for <see cref="DateStyleDefinition"/>
        /// </summary>
        public Cell DateCell { get; set; }

        /// <summary>
        /// Cached cell instance for represented style, for <see cref="StringStyleDefinition"/>
        /// </summary>
        public Cell StringCell { get; set; }

        public Cell SharedStringCell { get; set; }
    }

    class SharedCellStyleCollection
    {
        private readonly SharedElementCollection<CellStyleDefinition> cellStyleDefinitions;
        private readonly int cellFormatIndexCorrelation;

        private readonly List<CellStyleOption> allocatedSampleCells = new List<CellStyleOption>();
        
        public SharedCellStyleCollection(SharedElementCollection<CellStyleDefinition> sharedElementCollection, int cellFormatIndexCorrelation)
        {
            cellStyleDefinitions = sharedElementCollection;
            this.cellFormatIndexCorrelation = cellFormatIndexCorrelation;
        }

        public Cell AllocateDateCell(DateStyleDefinition cellStyleDefinition, CellValues cellValueType)
        {
            var option = AllocateCellOption(cellStyleDefinition.CellStyleDefinition);
            if (option.DateCell != null) return option.DateCell;
            var cell = CreateCell(cellValueType, option.Index);
            option.DateCell = cell;
            return cell;
        }

        public Cell AllocateNumberCell(NumberStyleDefinition cellStyleDefinition, CellValues cellValueType)
        {
            var option = AllocateCellOption(cellStyleDefinition.CellStyleDefinition);
            if (option.NumberCell != null) return option.NumberCell;
            var cell = CreateCell(cellValueType, option.Index);
            option.NumberCell = cell;
            return cell;
        }

        public Cell AllocateStringCell(StringStyleDefinition cellStyleDefinition, CellValues cellValueType)
        {
            var option = AllocateCellOption(cellStyleDefinition.CellStyleDefinition);
            if (option.StringCell != null) return option.StringCell;
            var cell = CreateCell(cellValueType, option.Index);
            option.StringCell = cell;
            return cell;
        }

        public Cell AllocateSharedStringCell(SharedStringStyleDefinition cellStyleDefinition, CellValues cellValueType)
        {
            var option = AllocateCellOption(cellStyleDefinition.CellStyleDefinition);
            if (option.SharedStringCell != null) return option.SharedStringCell;
            var cell = CreateCell(cellValueType, option.Index);
            option.SharedStringCell = cell;
            return cell;
        }

        private Cell CreateCell(CellValues cellValueType, int index)
        {
            var cell = new Cell();
            cell.StyleIndex = (uint)index;
            if (true)
                cell.DataType = new EnumValue<CellValues>(cellValueType);
            return cell;
        }

        private CellStyleOption AllocateCellOption(CellStyleDefinition cellStyleDefinition)
        {
            var index = cellStyleDefinitions.GetOrAllocateElement(cellStyleDefinition);
            if (index < allocatedSampleCells.Count)
                return allocatedSampleCells[index];
            var cellStyleOption = new CellStyleOption((allocatedSampleCells.Count + cellFormatIndexCorrelation));
            allocatedSampleCells.Add(cellStyleOption);
            return cellStyleOption;
        }
    }
}
