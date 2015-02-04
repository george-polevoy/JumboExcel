using JumboExcel.Structure;

namespace JumboExcel
{
    /// <summary>
    /// Visitor interface for hierarchical structure of <see cref="DocumentElement"/>.
    /// </summary>
    internal interface IElementVisitor
    {
        void Visit(Worksheet worksheet);
        void Visit(Row row);
        void Visit(RowGroup rowGroup);
        void VisitEmptyCell();
        void Visit(IntegerCell integerCell);
        void Visit(DecimalCell decimalCell);
        void Visit(DateTimeCell dateTimeCell);
        void Visit(InlineString inlineString);
        void Visit(SharedString sharedString);
        void Visit(BooleanCell booleanCell);
    }
}
