namespace JumboExcel.Structure
{
    /// <summary>
    /// Base class for document hierarchy elements.
    /// </summary>
    public abstract class DocumentElement
    {
        internal abstract void Accept(IElementVisitor visitor);
    }
}
