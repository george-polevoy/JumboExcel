using JumboExcel.Structure;
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
                        new SharedString("Cell A1"), new SharedString("Cell A2")))
            }
        );
    }
}