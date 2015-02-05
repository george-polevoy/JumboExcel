using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using JumboExcel.Structure;
using NUnit.Framework;

namespace JumboExcel
{
    class ProgressingWorksheetTests
    {
        [Test, Explicit]
        public void WriteWithProgress()
        {
            var progressingWorksheets = new[] { new ProgressingWorksheet<int>("Progressing", new WorksheetParametersElement(), GenerateRows) };

            var fileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var outputStream = new FileStream(fileName, FileMode.CreateNew))
            {
                foreach (var iteration in OpenXmlBuilder.WriteWithProgress(outputStream, progressingWorksheets))
                {
                    Console.WriteLine("Progress: {0}", iteration);
                }
            }
            var fileSize = new FileInfo(fileName).Length;
            Console.WriteLine("Size of the file generated: {0}", fileSize);
            Console.WriteLine(fileName);

            Process.Start(fileName);
        }

        private static IEnumerable<int> GenerateRows(Action<IEnumerable<RowLevelElement>> writeElements)
        {
            yield return 0;

            writeElements(new[] { new Row(new IntegerCell(1)) });

            yield return 50;

            writeElements(new[] { new Row(new IntegerCell(2)) });

            yield return 100;
        }
    }
}