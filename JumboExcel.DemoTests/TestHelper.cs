using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using JumboExcel.Structure;
using NUnit.Framework;

namespace JumboExcel
{
    static class TestHelper
    {
        internal static void WriteAndExecuteExcel(IEnumerable<Worksheet> worksheetElements, string prefix = "")
        {
            var fileName = WriteFile(worksheetElements, prefix);
            Process.Start(new ProcessStartInfo(fileName) { UseShellExecute = true });
        }

        internal static string WriteFile(IEnumerable<Worksheet> worksheetElements, string prefix = "")
        {
            var fileName = Path.Combine(Path.GetTempPath(), prefix + Guid.NewGuid() + ".xlsx");
            using (var outputStream = new FileStream(fileName, FileMode.CreateNew))
            {
                OpenXmlBuilder.Write(
                    outputStream,
                    worksheetElements
                    );
            }

            TestContext.WriteLine($"cp {fileName} .");
            return fileName;
        }
    }
}