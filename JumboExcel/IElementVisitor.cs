using JumboExcel.Structure;

namespace JumboExcel
{
    public interface IElementVisitor
    {
        void Visit(WorksheetElement worksheetElement);
        void Visit(RowElement rowElement);
        void Visit(RowGroupElement rowGroupElement);
        void VisitEmptyCell();
        void Visit(IntegerCellElement integerCellElement);
        void Visit(DecimalCellElement decimalCellElement);
        void Visit(DateCellElement dateCellElement);
        void Visit(InlineStringElement inlineStringElement);
        void Visit(SharedStringElement sharedStringElement);
    }
}
