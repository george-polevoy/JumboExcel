using System.Collections.Generic;

namespace JumboExcel.Structure
{
    public abstract class RowLevelElement : DocumentElement
    {
    }

    public class RowGroupElement : RowLevelElement
    {
        private readonly IEnumerable<RowLevelElement> rowElements;

        public IEnumerable<RowLevelElement> RowElements { get { return rowElements; } }

        public RowGroupElement(IEnumerable<RowElement> rowElements)
        {
            this.rowElements = rowElements;
        }

        public RowGroupElement(params RowLevelElement[] rowElements)
        {
            this.rowElements = rowElements;
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }

    public class RowElement : RowLevelElement
    {
        private readonly IEnumerable<CellElement> cellElements;

        public IEnumerable<CellElement> CellElements
        {
            get { return cellElements; }
        }

        public RowElement(IEnumerable<CellElement> cellElements)
        {
            this.cellElements = cellElements;
        }

        public RowElement(params CellElement[] cellElements)
        {
            this.cellElements = cellElements;
        }

        public override void Accept(IElementVisitor visitor)
        {
            visitor.Visit(this);
        }
    }
}
