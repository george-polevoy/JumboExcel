using JumboExcel;
using JumboExcel.Structure;
using NUnit.Framework;

public class TestFeatures
{
    [Explicit]
    [Test]
    public void TestVariousFormats()
    {
        TestHelper.WriteAndExecuteExcel(new Worksheet[]{new Worksheet("InlineString", null, new Row(new InlineString("Cell A1")))}, "InlineString_");
        TestHelper.WriteAndExecuteExcel(new Worksheet[]{new Worksheet("SharedString", null, new Row(new SharedString("Cell A1")))}, "SharedString_");
    }
}