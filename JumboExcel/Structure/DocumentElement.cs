namespace JumboExcel.Structure
{
    /// <summary>
    /// Base class for document hierarchy elements.
    /// </summary>
    public abstract class DocumentElement
    {
        public abstract void Accept(IElementVisitor visitor);
    }
}
