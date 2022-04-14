using System.Drawing;
using JumboExcel.Structure;
using JumboExcel.Styling;
using NUnit.Framework;

namespace JumboExcel;

public class SimpleFileTest
{
    [Test]
    public void WriteSimplestPossibleFile()
    {
        TestHelper.WriteAndExecuteExcel(new Worksheet[]
            {
                new("Sheet 1", null,
                    new Row(
                        new InlineString("Cell A1", new StringStyle(new Font(null, 20, Color.Black, FontSlope.ITALIC, FontWeight.BOLD))), new SharedString("Cell A2"))),
                new("Sheet 2", null, new Row(new SharedString("Hello")))
            }
        );
    }
}