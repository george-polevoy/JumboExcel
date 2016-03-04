using System.Linq;
using JumboExcel.Structure;
using NUnit.Framework;

namespace JumboExcel
{
    class FrozenPaneTests
    {
        [Test, Explicit]
        public void TestFrozenPanesByRow()
        {
            TestFreeze("Freeze by row", new PaneFreezer(1, 0));
        }

        [Test, Explicit]
        public void TestFrozenPanesByColumn()
        {
            TestFreeze("Freeze by column", new PaneFreezer(0, 1));
        }

        [Test, Explicit]
        public void TestFrozenPanesBoth()
        {
            TestFreeze("Freeze both", new PaneFreezer(1, 1));
        }

        private static void TestFreeze(string what, PaneFreezer paneFreezer)
        {
            TestHelper.WriteAndExecuteExcel(new[]
            {
                new Worksheet(what, new WorksheetParametersElement(false, false, null, paneFreezer), SampleDataSources.GetMultiplicationTableCells(40, 60).Select(r => new Row(r)))
            });
        }
    }
}