namespace JumboExcel.Structure
{
    public abstract class DocumentElement
    {
        public abstract void Accept(IElementVisitor visitor);
    }
}
