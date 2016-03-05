using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using JumboExcel.Structure;
using NUnit.Framework;
using NUnit.Framework.Constraints;

namespace JumboExcel
{
    class WorksheetWriteTests
    {
        [Test]
        [TestCase(0, typeof(ArgumentNullException))]
        [TestCase(1, null)]
        [TestCase(31, null)]
        [TestCase(32, typeof(ArgumentOutOfRangeException))]
        public void WorksheetNameLengthIsLimited(int sheetNameLength, Type exceptionType)
        {
            Assert.That(() => new Worksheet(string.Join("", Enumerable.Repeat("a", sheetNameLength)), null),
                exceptionType == null ? (IResolveConstraint)Throws.Nothing : Throws.Exception.AssignableTo(exceptionType));
        }

        [Test]
        public void WorksheetNameIsNotNull()
        {
            Assert.That(() => new Worksheet(null, null), Throws.Exception.AssignableTo<ArgumentNullException>());
        }

        [Test]
        public void SameWorksheetNameCantBeWrittenTwice()
        {
            var ms = new MemoryStream();
            var sameSheet1 = new Worksheet("SameSheet", null);
            var sameSheet2 = new Worksheet("SameSheet", null);

            Assert.That(() => OpenXmlBuilder.Write(ms, new[] { sameSheet1, sameSheet2 }), Throws.InvalidOperationException);
        }

        [Test]
        [TestCaseSource("NegativeAnchorCasesForAbsoluteCellMerger")]
        public void AbsoluteCellMergerThrowsOnNegativeResultingAnchor(int anchorRow, int anchorColumn, int mergedRow, int mergedColumn)
        {
            var cells = new CellElement[2, 2];
            for (var row = 0; row < 2; row++)
                for (var column = 0; column < 2; column++)
                    cells[row, column] = new InlineString("*");

            cells[mergedRow, mergedColumn] = new AbsoluteCellMerger(new InlineString("+"), anchorRow, anchorColumn);

            var ms = new MemoryStream();
            var sheet = new Worksheet("SameSheet", null,
                new Row(cells[0, 0], cells[0, 1]),
                new Row(cells[1, 0], cells[1, 1])
                );

            Assert.That(() => OpenXmlBuilder.Write(ms, new[] { sheet }), Throws.InvalidOperationException);
        }

        static IEnumerable<TestCaseData> NegativeAnchorCasesForAbsoluteCellMerger()
        {
            return
                from anchorRow in new[] { -1, 0 }
                from anchorColumn in new[] { -1, 0 }
                where anchorRow < 0 || anchorColumn < 0
                from mergedRow in new[] { 0, 1 }
                from mergedColumn in new[] { 0, 1 }
                select new TestCaseData(anchorRow, anchorColumn, mergedRow, mergedColumn);
        }

        [Test]
        [TestCaseSource("NegativeAnchorCasesForRelativeCellMerger")]
        public void RelativeCellMergerThrowsOnNegativeResultingAnchor(int rowOffset, int columnOffset, int mergedRow, int mergedColumn)
        {
            var cells = new CellElement[2, 2];
            for (var row = 0; row < 2; row++)
                for (var column = 0; column < 2; column++)
                    cells[row, column] = new InlineString("*");

            cells[mergedRow, mergedColumn] = new RelativeCellMerger(new InlineString("+"), rowOffset, columnOffset);

            var ms = new MemoryStream();
            var sheet = new Worksheet("SameSheet", null,
                new Row(cells[0, 0], cells[0, 1]),
                new Row(cells[1, 0], cells[1, 1])
                );

            Assert.That(() => OpenXmlBuilder.Write(ms, new[] { sheet }), Throws.InvalidOperationException);
        }

        static IEnumerable<TestCaseData> NegativeAnchorCasesForRelativeCellMerger()
        {
            return
                from anchorRow in new[] { -1, 0 }
                from anchorColumn in new[] { -1, 0 }
                where anchorRow < 0 || anchorColumn < 0
                from mergedRow in new[] { 0, 1 }
                from mergedColumn in new[] { 0, 1 }
                select new TestCaseData(mergedRow - anchorRow, mergedColumn - anchorColumn, mergedRow, mergedColumn);
        }
    }
}