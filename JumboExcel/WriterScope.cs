using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Excel.RichData;

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

        public WriterScope(OpenXmlWriter writer, OpenXmlElement element, IEnumerable<(string key, string reference)> namespaces)
        {
            this.writer = writer;
            writer.WriteStartElement(element, Enumerable.Empty<OpenXmlAttribute>(), namespaces.Select(n => new KeyValuePair<string, string>(n.key, n.reference)));
        }

        public void Dispose()
        {
            writer.WriteEndElement();
        }
    }
}
