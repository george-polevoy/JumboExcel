using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using JumboExcel.Structure;
using NUnit.Framework;

namespace JumboExcel
{
    class BigWriteTests
    {
        [Test, Explicit]
        public void TestSimplestReport()
        {
            var tempFileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var file = new FileStream(tempFileName, FileMode.CreateNew))
            {
                OpenXmlBuilder.Write(file, new[] { new Worksheet("Report", new WorksheetParametersElement(), new Row(new SharedString("Hello"))) });
            }
            Process.Start(tempFileName);
        }

        [Test]
        [Explicit]
        [TestCase(1000000, 1)]
        public void ReallyHuge(int rowCount, int columnCount)
        {
            TestHelper.WriteAndExecuteExcel(Enumerable.Range(0, 1).Select(sheet => new Worksheet("Sheet  " + sheet, new WorksheetParametersElement(), Enumerable.Range(0, rowCount).Select(row => new Row(
                new[] { new SharedString("Row mod 10: " + row % 10) }
                    .Concat(Enumerable.Range(0, columnCount).SelectMany(column => new CellElement[]
                    {
                        new SharedString("Column: " + column),
                        EmptyCell.Instance,
                        new IntegerCell(row*column + sheet),
                        new DecimalCell((decimal) column/columnCount),
                        new InlineString("Row: " + row*column),
                    })))))));
        }

        [Test]
        [Explicit]
        [TestCase(1, 100000, 1)]
        [TestCase(1, 10000, 10)]
        [TestCase(1, 100000, 10)]
        [TestCase(1, 1000000, 1)]
        [TestCase(2, 1000000, 5)]
        public void PerformanceTest(int sheetCount, int rowCount, int columnCount)
        {
            var file = TestHelper.WriteFile(Enumerable.Range(0, sheetCount).Select(sheet => new Worksheet("Sheet  " + sheet, new WorksheetParametersElement(), Enumerable.Range(0, rowCount).Select(row => new Row(
                Enumerable.Range(0, columnCount).Select(column => new IntegerCell(row * column + sheet)))))));
            File.Delete(file);
        }
    }
}
