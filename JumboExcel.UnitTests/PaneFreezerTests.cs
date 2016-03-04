using System;
using System.Collections.Generic;
using System.Linq;
using JumboExcel.Structure;
using NUnit.Framework;

namespace JumboExcel
{
    public class PaneFreezerTests
    {
        [Test]
        [TestCaseSource("PaneFreezerCanConstructCases")]
        public void PaneFreezerCanHaveRowIndex(int rowIndex, int columnIndex)
        {
            var actual = new PaneFreezer(rowIndex, columnIndex);
            Assert.AreEqual(rowIndex, actual.RowIndex);
        }

        [Test]
        [TestCaseSource("PaneFreezerCanConstructCases")]
        public void PaneFreezerCanHaveColumnIndex(int rowIndex, int columnIndex)
        {
            var actual = new PaneFreezer(rowIndex, columnIndex);
            Assert.AreEqual(columnIndex, actual.ColumnIndex);
        }

        private static IEnumerable<TestCaseData> PaneFreezerCanConstructCases()
        {
            return
                from r in new[] {0, 1, 2}
                from c in new[] {0, 1, 2}
                select new TestCaseData(r, c);
        }

        [Test]
        [TestCaseSource("PaneFreezerCanNotConstructCases")]
        public void PaneFreezerCanNotConstruct(int rowIndex, int columnIndex)
        {
            Assert.That(() => new PaneFreezer(rowIndex, columnIndex), Throws.Exception.InstanceOf<ArgumentException>());
        }

        private static IEnumerable<TestCaseData> PaneFreezerCanNotConstructCases()
        {
            return
                from r in new[] { -2, -1, 0, 1, 2 }
                from c in new[] { -2, -1, 0, 1, 2 }
                where r < 0 || c < 0
                select new TestCaseData(r, c);
        }
    }
}
