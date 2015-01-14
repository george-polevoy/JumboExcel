using JumboExcel.Structure;

namespace JumboExcel
{
    /// <summary>
    /// Visitor interface for hierarchical structure of <see cref="DocumentElement"/>.
    /// </summary>
    public interface IElementVisitor
    {
        void Visit(WorksheetElement worksheetElement);
        void Visit(RowElement rowElement);
        void Visit(RowGroupElement rowGroupElement);
        void VisitEmptyCell();
        void Visit(IntegerCellElement integerCellElement);
        void Visit(DecimalCellElement decimalCellElement);
        void Visit(DateTimeCellElement dateTimeCellElement);
        void Visit(InlineStringElement inlineStringElement);
        void Visit(SharedStringElement sharedStringElement);
        void Visit(BooleanCellElement booleanCellElement);
    }
}
