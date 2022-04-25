using System.Drawing;
using NUnit.Framework;

namespace JumboExcel;

public class ColorTranslation
{
    [TestCase(0, 0, 0, "000000")]
    [TestCase(0, 0, 1, "000001")]
    [TestCase(0, 0, 255, "0000FF")]
    [TestCase(0, 1, 0, "000100")]
    [TestCase(0, 255, 0, "00FF00")]
    [TestCase(1, 0, 0, "010000")]
    [TestCase(255, 0, 0, "FF0000")]
    [TestCase(15, 15, 15, "0F0F0F")]
    [TestCase(255, 255, 255, "FFFFFF")]
    public void RgbToHex(int r, int g, int b, string expected)
    {
        Assert.AreEqual(expected, OpenXmlBuilder.ToHexColor(Color.FromArgb(255, r, g, b)));
    }
}