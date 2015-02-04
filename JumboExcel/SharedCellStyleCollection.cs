using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using JumboExcel.Styling;
using CellStyle = JumboExcel.Styling.CellStyle;

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
        /// Cached cell instance for represented style, for <see cref="NumberStyle"/>
        /// </summary>
        public Cell NumberCell { get; set; }

        /// <summary>
        /// Cached cell instance for represented style, for <see cref="DateStyle"/>
        /// </summary>
        public Cell DateCell { get; set; }

        /// <summary>
        /// Cached cell instance for represented style, for <see cref="StringStyle"/>
        /// </summary>
        public Cell StringCell { get; set; }

        /// <summary>
        /// Cached cell instance for represented style, for <see cref="StringStyle"/>
        /// </summary>
        public Cell SharedStringCell { get; set; }

        /// <summary>
        /// Cached cell instance for represented style, for <see cref="BooleanStyle"/>
        /// </summary>
        public Cell BooleanCell { get; set; }
    }

    class SharedCellStyleCollection
    {
        private readonly SharedElementCollection<CellStyle> cellStyleDefinitions;
        private readonly int cellFormatIndexCorrelation;

        private readonly List<CellStyleOption> allocatedSampleCells = new List<CellStyleOption>();
        
        public SharedCellStyleCollection(SharedElementCollection<CellStyle> sharedElementCollection, int cellFormatIndexCorrelation)
        {
            cellStyleDefinitions = sharedElementCollection;
            this.cellFormatIndexCorrelation = cellFormatIndexCorrelation;
        }

        public Cell AllocateDateCell(DateStyle cellStyle, CellValues cellValueType)
        {
            var option = AllocateCellOption(cellStyle.cellStyle);
            if (option.DateCell != null) return option.DateCell;
            var cell = CreateCell(cellValueType, option.Index);
            option.DateCell = cell;
            return cell;
        }

        public Cell AllocateNumberCell(NumberStyle cellStyle, CellValues cellValueType)
        {
            var option = AllocateCellOption(cellStyle.cellStyle);
            if (option.NumberCell != null) return option.NumberCell;
            var cell = CreateCell(cellValueType, option.Index);
            option.NumberCell = cell;
            return cell;
        }

        public Cell AllocateStringCell(StringStyle cellStyle, CellValues cellValueType)
        {
            var option = AllocateCellOption(cellStyle.cellStyle);
            if (option.StringCell != null) return option.StringCell;
            var cell = CreateCell(cellValueType, option.Index);
            option.StringCell = cell;
            return cell;
        }

        public Cell AllocateSharedStringCell(StringStyle cellStyle, CellValues cellValueType)
        {
            var option = AllocateCellOption(cellStyle.cellStyle);
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

        private CellStyleOption AllocateCellOption(CellStyle cellStyle)
        {
            var index = cellStyleDefinitions.GetOrAllocateElement(cellStyle);
            if (index < allocatedSampleCells.Count)
                return allocatedSampleCells[index];
            var cellStyleOption = new CellStyleOption((allocatedSampleCells.Count + cellFormatIndexCorrelation));
            allocatedSampleCells.Add(cellStyleOption);
            return cellStyleOption;
        }

        public Cell AllocateBooleanCell(BooleanStyle cellStyleDefinition, CellValues cellValueType)
        {
            var option = AllocateCellOption(cellStyleDefinition.cellStyle);
            if (option.SharedStringCell != null) return option.BooleanCell;
            var cell = CreateCell(cellValueType, option.Index);
            option.BooleanCell = cell;
            return cell;
        }
    }
}
