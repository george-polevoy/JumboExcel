using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using JumboExcel.Structure;

namespace JumboExcel
{
    static class TestHelper
    {
        internal static void WriteAndExecuteExcel(IEnumerable<Worksheet> worksheetElements)
        {
            var fileName = WriteFile(worksheetElements);
            Process.Start(fileName);
        }

        internal static string WriteFile(IEnumerable<Worksheet> worksheetElements)
        {
            var fileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var outputStream = new FileStream(fileName, FileMode.CreateNew))
            {
                OpenXmlBuilder.Write(
                    outputStream,
                    worksheetElements
                    );
            }
            var fileSize = new FileInfo(fileName).Length;
            Console.WriteLine("Size of the file generated: {0}", fileSize);
            Console.WriteLine(fileName);
            return fileName;
        }
    }
}