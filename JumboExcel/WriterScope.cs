using System;
using DocumentFormat.OpenXml;

namespace JumboExcel
{
    /// <summary>
    /// Convenience helper for writing start and end element tag counterparts with <see cref="OpenXmlWriter"/>
    /// Public constructor overloads of this struct call the <see cref="OpenXmlWriter.WriteStartElement(DocumentFormat.OpenXml.OpenXmlReader)"/>
    /// and <see cref="OpenXmlWriter.WriteStartElement(DocumentFormat.OpenXml.OpenXmlReader, System.Collections.Generic.IEnumerable{DocumentFormat.OpenXml.OpenXmlAttribute})"/> methods, according to provided arguments.
    /// </summary>
    /// <example>
    /// using(new WriterScope(writer, new Workbook()))
    /// {
    ///     // Write the workbook content...
    /// }
    /// </example>
    struct WriterScope : IDisposable
    {
        private readonly OpenXmlWriter writer;

        public WriterScope(OpenXmlWriter writer, OpenXmlElement element)
        {
            this.writer = writer;
            writer.WriteStartElement(element);
        }

        public void Dispose()
        {
            writer.WriteEndElement();
        }
    }
}
